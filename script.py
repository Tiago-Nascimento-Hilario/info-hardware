import wmi  
import psutil  
import time
import os
import subprocess
from datetime import datetime
from rich.console import Console  
from rich.table import Table  

# ReportLab imports
from reportlab.lib.pagesizes import letter  
from reportlab.lib import colors  
from reportlab.platypus import SimpleDocTemplate, Paragraph, Spacer, Table as RLTable, TableStyle  
from reportlab.lib.styles import getSampleStyleSheet, ParagraphStyle  

console = Console()
c = wmi.WMI()

# ----------------------------------------------------------------------
# 1. FUNÇÕES DE COLETA DE DADOS (Retornam Dicionários/Listas)
# ----------------------------------------------------------------------

def collect_bios_info():
    data = []
    for bios in c.Win32_BIOS():
        data.append(["Fabricante da BIOS", str(bios.Manufacturer)])
        data.append(["Versão da BIOS", str(bios.SMBIOSBIOSVersion)])
        data.append(["Data de Lançamento", str(bios.ReleaseDate if hasattr(bios, 'ReleaseDate') else 'N/A')[:8]])
        data.append(["Versão SMBIOS", f"{bios.SMBIOSMajorVersion}.{bios.SMBIOSMinorVersion}"])

    try:
        cmd_secure = 'powershell "Confirm-SecureBootUEFI"'
        res = subprocess.run(cmd_secure, capture_output=True, text=True, shell=True)
        secure_boot = res.stdout.strip()
        status_sb = "Sim" if "True" in secure_boot else ("Não" if "False" in secure_boot else "Não Suportado / Legacy")
        data.append(["Secure Boot Ativo", status_sb])
    except Exception:
        data.append(["Secure Boot Ativo", "Não foi possível verificar"])

    return data

def collect_motherboard():
    data = []
    for board in c.Win32_BaseBoard():
        data.append(["Fabricante", str(board.Manufacturer)])
        data.append(["Modelo", str(board.Product)])
        data.append(["Número de Série", str(board.SerialNumber)])
    return data

def collect_cpu():
    data = []
    for cpu in c.Win32_Processor():
        data.append(["Nome", str(cpu.Name).strip() if cpu.Name else "N/A"])
        data.append(["Núcleos Físicos", str(cpu.NumberOfCores)])
        data.append(["Threads", str(cpu.NumberOfLogicalProcessors)])
        data.append(["Uso Atual", f"{psutil.cpu_percent(interval=1)}%"])
    return data

def collect_ram():
    data = []
    for mem in c.Win32_PhysicalMemory():
        cap_gb = round(int(mem.Capacity) / (1024**3), 2) if mem.Capacity else 0
        speed = f"{mem.Speed} MHz" if mem.Speed else "N/A"
        manuf = str(mem.Manufacturer) if mem.Manufacturer else "Desconhecido"
        data.append([str(mem.DeviceLocator), f"{cap_gb} GB", speed, manuf])
    return data

def collect_drives():
    data = []
    for disk in c.Win32_DiskDrive():
        size_gb = round(int(disk.Size) / (1024**3), 2) if disk.Size else 0
        data.append([str(disk.Model), str(disk.InterfaceType), f"{size_gb} GB"])
    return data

def collect_device_manager():
    devices = c.Win32_PnPEntity()
    total_devices = 0
    error_count = 0
    data = []

    for dev in devices:
        if dev.Name:
            total_devices += 1
            status_code = getattr(dev, 'ConfigManagerErrorCode', 0)
            pnp_class = str(getattr(dev, 'PNPClass', 'Outros'))
            
            if status_code != 0:
                error_count += 1
                data.append([pnp_class, str(dev.Name), "FALHA / ERRO", f"Código {status_code}"])
            else:
                data.append([pnp_class, str(dev.Name), "OK", "0 - Operacional"])

    return data, total_devices, error_count

def collect_usb_devices():
    data = []
    for usb in c.Win32_PnPEntity():
        if getattr(usb, 'PNPClass', None) == "USB" or (usb.DeviceID and "USB" in usb.DeviceID):
            data.append([str(usb.Name or "Desconhecido"), str(usb.DeviceID)])
    return data

def test_cpu_performance():
    start_time = time.time()
    _ = [i**2 for i in range(2_000_000)]
    elapsed = time.time() - start_time
    return f"{elapsed:.4f} segundos"

# ----------------------------------------------------------------------
# 2. FUNÇÕES DE IMPRESSÃO NO TERMINAL (RICH)
# ----------------------------------------------------------------------

def print_rich_table(title, headers, data, styles):
    table = Table(title=title)
    for h, s in zip(headers, styles):
        table.add_column(h, style=s)
    for row in data:
        table.add_row(*row)
    console.print(table)

# ----------------------------------------------------------------------
# 3. FUNÇÃO PARA GERAR O RELATÓRIO PDF COMPLETO (REPORTLAB)
# ----------------------------------------------------------------------

def generate_pdf_report(diagnostics_data, filename="Diagnostico_Sistema.pdf"):
    doc = SimpleDocTemplate(
        filename,
        pagesize=letter,
        rightMargin=30, leftMargin=30, topMargin=30, bottomMargin=30
    )
    styles = getSampleStyleSheet()
    
    # Estilos customizados
    title_style = ParagraphStyle(
        'TitleStyle',
        parent=styles['Heading1'],
        fontName='Helvetica-Bold',
        fontSize=18,
        textColor=colors.HexColor("#1B365D"),
        spaceAfter=6,
        alignment=1
    )
    
    subtitle_style = ParagraphStyle(
        'SubTitleStyle',
        parent=styles['Normal'],
        fontName='Helvetica-Oblique',
        fontSize=10,
        textColor=colors.gray,
        spaceAfter=15,
        alignment=1
    )

    section_style = ParagraphStyle(
        'SectionStyle',
        parent=styles['Heading2'],
        fontName='Helvetica-Bold',
        fontSize=12,
        textColor=colors.HexColor("#1B365D"),
        spaceBefore=14,
        spaceAfter=6
    )

    cell_style = ParagraphStyle(
        'CellStyle',
        parent=styles['Normal'],
        fontName='Helvetica',
        fontSize=8,
        leading=10
    )

    cell_bold_style = ParagraphStyle(
        'CellBoldStyle',
        parent=styles['Normal'],
        fontName='Helvetica-Bold',
        fontSize=8,
        leading=10,
        textColor=colors.whitesmoke
    )

    elements = []

    # Cabeçalho do Documento
    elements.append(Paragraph("Homologação - Relatório Diagnóstico de Hardware", title_style))
    elements.append(Paragraph(f"Gerado em: {datetime.now().strftime('%d/%m/%Y %H:%M:%S')}", subtitle_style))
    elements.append(Spacer(1, 10))

    # Função para criar tabelas no PDF convertendo texto longo em parágrafos embrulhados (wrap)
    def make_pdf_table(headers, rows, col_widths=None):
        formatted_data = []
        
        # Cabeçalho da tabela
        header_row = [Paragraph(f"<b>{h}</b>", cell_bold_style) for h in headers]
        formatted_data.append(header_row)

        # Conteúdo da tabela
        for row in rows:
            formatted_row = [Paragraph(str(cell), cell_style) for cell in row]
            formatted_data.append(formatted_row)

        t = RLTable(formatted_data, colWidths=col_widths, repeatRows=1)
        t.setStyle(TableStyle([
            ('BACKGROUND', (0, 0), (-1, 0), colors.HexColor("#1B365D")),
            ('BOTTOMPADDING', (0, 0), (-1, 0), 5),
            ('TOPPADDING', (0, 0), (-1, 0), 5),
            ('BACKGROUND', (0, 1), (-1, -1), colors.HexColor("#F8F9FA")),
            ('GRID', (0, 0), (-1, -1), 0.5, colors.HexColor("#D3D3D3")),
            ('VALIGN', (0, 0), (-1, -1), 'TOP'),
        ]))
        return t

    # Seções Padrão
    sections = [
        ("BIOS / UEFI", ["Propriedade", "Valor"], diagnostics_data['bios'], [180, 370]),
        ("Placa-Mãe", ["Propriedade", "Valor"], diagnostics_data['motherboard'], [180, 370]),
        ("Processador", ["Propriedade", "Valor"], diagnostics_data['cpu'], [180, 370]),
        ("Memória RAM", ["Slot / Banco", "Capacidade", "Velocidade", "Fabricante"], diagnostics_data['ram'], [120, 120, 150, 160]),
        ("Armazenamento (SSDs / HDDs)", ["Modelo", "Tipo de Interface", "Tamanho"], diagnostics_data['drives'], [250, 150, 150]),
        ("Dispositivos USB", ["Descrição", "ID do Dispositivo"], diagnostics_data['usb'], [200, 350]),
    ]

    for title, headers, data, widths in sections:
        elements.append(Paragraph(title, section_style))
        elements.append(make_pdf_table(headers, data, widths))
        elements.append(Spacer(1, 10))

    # --- SEÇÃO DO GERENCIADOR DE DISPOSITIVOS COMPLETA ---
    elements.append(Paragraph("Gerenciador de Dispositivos (Lista Completa)", section_style))
    summary_text = f"<b>Total de dispositivos listados:</b> {diagnostics_data['dev_total']} | <b>Dispositivos com Falha:</b> {diagnostics_data['dev_errors']}"
    elements.append(Paragraph(summary_text, styles['Normal']))
    elements.append(Spacer(1, 6))

    # Inclui TODOS os dispositivos coletados sem filtros
    elements.append(make_pdf_table(
        ["Categoria", "Nome do Dispositivo", "Status", "Código"], 
        diagnostics_data['device_manager'], 
        [110, 240, 90, 110]
    ))

    elements.append(Spacer(1, 10))
    elements.append(Paragraph("Desempenho da CPU", section_style))
    elements.append(Paragraph(f"<b>Tempo do Benchmark Rápido:</b> {diagnostics_data['cpu_score']}", styles['Normal']))

    # Gerar o arquivo PDF
    doc.build(elements)

# ----------------------------------------------------------------------
# 4. EXECUÇÃO PRINCIPAL
# ----------------------------------------------------------------------

if __name__ == "__main__":
    console.print("[bold blue]=== DIAGNÓSTICO DE HARDWARE E DESEMPENHO ===[/bold blue]\n")

    # Coleta todos os dados
    bios_data = collect_bios_info()
    mb_data = collect_motherboard()
    cpu_data = collect_cpu()
    ram_data = collect_ram()
    drives_data = collect_drives()
    dev_data, dev_total, dev_errors = collect_device_manager()
    usb_data = collect_usb_devices()

    # Imprime no Console
    print_rich_table("Configurações e Informações da BIOS / UEFI", ["Propriedade", "Valor"], bios_data, ["cyan", "green"])
    print_rich_table("Placa-Mãe", ["Propriedade", "Valor"], mb_data, ["cyan", "green"])
    print_rich_table("Processador", ["Propriedade", "Valor"], cpu_data, ["cyan", "green"])
    print_rich_table("Memória RAM", ["Slot / Banco", "Capacidade", "Velocidade", "Fabricante"], ram_data, ["cyan", "green", "magenta", "yellow"])
    print_rich_table("Armazenamento (SSDs / HDDs)", ["Modelo", "Tipo de Interface", "Tamanho"], drives_data, ["cyan", "green", "magenta"])
    
    console.print(f"\n[bold]Total de dispositivos analisados:[/bold] {dev_total} | [bold red]Com falha/driver ausente:[/bold red] {dev_errors}\n")
    print_rich_table("Dispositivos USB Conectados", ["Descrição", "ID do Dispositivo"], usb_data, ["cyan", "green"])

    console.print("\n[bold yellow]Executando teste de benchmark rápido de CPU...[/bold yellow]")
    cpu_score = test_cpu_performance()
    console.print(f"[bold green]Métrica de Desempenho da CPU:[/bold green] {cpu_score}\n")

    # Estrutura dados para a função do PDF
    all_diagnostics = {
        'bios': bios_data,
        'motherboard': mb_data,
        'cpu': cpu_data,
        'ram': ram_data,
        'drives': drives_data,
        'device_manager': dev_data,
        'dev_total': dev_total,
        'dev_errors': dev_errors,
        'usb': usb_data,
        'cpu_score': cpu_score
    }

    # Gerar Relatório PDF
    desktop_path = os.path.join(os.path.expanduser("~"), "Desktop")
    pdf_filename = "Diagnostico_Sistema.pdf"
    console.print(f"[bold yellow]Gerando relatório PDF: {pdf_filename}...[/bold yellow]")
    try:
        generate_pdf_report(all_diagnostics, pdf_filename)
        console.print(f"[bold green]Relatório PDF gerado com sucesso![/bold green]")
    except Exception as e:
        console.print(f"[bold red]Erro ao gerar PDF:[/bold red] {e}")

    console.print("\n[bold yellow]Pressione ENTER para sair...[/bold yellow]")
    input()