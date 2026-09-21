import wmi # type: ignore
import psutil # type: ignore
import time
import os
import subprocess
from rich.console import Console # type: ignore
from rich.table import Table # type: ignore

console = Console()
c = wmi.WMI()
 
def get_bios_info():
    table = Table(title="Configurações e Informações da BIOS / UEFI")
    table.add_column("Propriedade", style="cyan")
    table.add_column("Valor", style="green")
    
    for bios in c.Win32_BIOS():
        table.add_row("Fabricante da BIOS", str(bios.Manufacturer))
        table.add_row("Versão da BIOS", str(bios.SMBIOSBIOSVersion))
        table.add_row("Data de Lançamento", str(bios.ReleaseDate if hasattr(bios, 'ReleaseDate') else 'N/A')[:8])
        table.add_row("Versão SMBIOS", f"{bios.SMBIOSMajorVersion}.{bios.SMBIOSMinorVersion}")

    # Verificação de Boot Mode e Secure Boot via PowerShell
    try:
        cmd_secure = 'powershell "Confirm-SecureBootUEFI"'
        res = subprocess.run(cmd_secure, capture_output=True, text=True, shell=True)
        secure_boot = res.stdout.strip()
        table.add_row("Secure Boot Ativo", "Sim" if "True" in secure_boot else ("Não" if "False" in secure_boot else "Não Suportado / Legacy"))
    except Exception:
        table.add_row("Secure Boot Ativo", "Não foi possível verificar")

    return table

def get_motherboard():
    table = Table(title="Placa-Mãe")
    table.add_column("Propriedade", style="cyan")
    table.add_column("Valor", style="green")
    
    for board in c.Win32_BaseBoard():
        table.add_row("Fabricante", str(board.Manufacturer))
        table.add_row("Modelo", str(board.Product))
        table.add_row("Número de Série", str(board.SerialNumber))
    return table

def get_cpu():
    table = Table(title="Processador")
    table.add_column("Propriedade", style="cyan")
    table.add_column("Valor", style="green")
    
    for cpu in c.Win32_Processor():
        table.add_row("Nome", str(cpu.Name).strip() if cpu.Name else "N/A")
        table.add_row("Núcleos Físicos", str(cpu.NumberOfCores))
        table.add_row("Threads", str(cpu.NumberOfLogicalProcessors))
        table.add_row("Uso Atual", f"{psutil.cpu_percent(interval=1)}%")
    return table

def get_ram():
    table = Table(title="Memória RAM")
    table.add_column("Slot / Banco", style="cyan")
    table.add_column("Capacidade", style="green")
    table.add_column("Velocidade", style="magenta")
    table.add_column("Fabricante", style="yellow")
    
    for mem in c.Win32_PhysicalMemory():
        cap_gb = round(int(mem.Capacity) / (1024**3), 2) if mem.Capacity else 0
        speed = f"{mem.Speed} MHz" if mem.Speed else "N/A"
        manuf = str(mem.Manufacturer) if mem.Manufacturer else "Desconhecido"
        table.add_row(str(mem.DeviceLocator), f"{cap_gb} GB", speed, manuf)
    return table

def get_drives():
    table = Table(title="Armazenamento (SSDs / HDDs)")
    table.add_column("Modelo", style="cyan")
    table.add_column("Tipo de Interface", style="green")
    table.add_column("Tamanho", style="magenta")
    
    for disk in c.Win32_DiskDrive():
        size_gb = round(int(disk.Size) / (1024**3), 2) if disk.Size else 0
        table.add_row(str(disk.Model), str(disk.InterfaceType), f"{size_gb} GB")
    return table

def get_device_manager_status():
    table = Table(title="Gerenciador de Dispositivos (Status e Erros)")
    table.add_column("Categoria / Classe", style="cyan")
    table.add_column("Nome do Dispositivo", style="white")
    table.add_column("Status", style="bold green")
    table.add_column("Código de Erro", style="bold yellow")

    devices = c.Win32_PnPEntity()
    total_devices = 0
    error_count = 0

    for dev in devices:
        if dev.Name:
            total_devices += 1
            # ConfigStatus == 0 significa que o dispositivo está funcionando corretamente
            status_code = getattr(dev, 'ConfigManagerErrorCode', 0)
            pnp_class = str(getattr(dev, 'PNPClass', 'Outros'))
            
            if status_code != 0:
                error_count += 1
                table.add_row(
                    pnp_class,
                    str(dev.Name),
                    "[bold red]FALHA / ERRO[/bold red]",
                    f"Erro Código {status_code}"
                )
            else:
                # Opcional: Mostrar também os que estão OK (filtrados para não poluir demais)
                table.add_row(pnp_class, str(dev.Name), "OK", "0 - Operacional")

    console.print(f"[bold bold]Total de dispositivos analisados:[/bold bold] {total_devices} | [bold red]Com falha/driver ausente:[/bold red] {error_count}\n")
    return table

def get_usb_devices():
    table = Table(title="Dispositivos USB Conectados")
    table.add_column("Descrição", style="cyan")
    table.add_column("ID do Dispositivo", style="green")
    
    for usb in c.Win32_PnPEntity():
        if getattr(usb, 'PNPClass', None) == "USB" or (usb.DeviceID and "USB" in usb.DeviceID):
            table.add_row(str(usb.Name or "Desconhecido"), str(usb.DeviceID))
    return table

def test_cpu_performance():
    console.print("\n[bold yellow]Executando teste de benchmark rápido de CPU...[/bold yellow]")
    start_time = time.time()
    _ = [i**2 for i in range(2_000_000)]
    elapsed = time.time() - start_time
    return f"{elapsed:.4f} segundos (quanto menor, mais rápido)"

if __name__ == "__main__":
    console.print("[bold blue]=== DIAGNÓSTICO DE HARDWARE E DESEMPENHO ===[/bold blue]\n")
    
    console.print(get_bios_info())
    console.print(get_motherboard())
    console.print(get_cpu())
    console.print(get_ram())
    console.print(get_drives())
    console.print(get_device_manager_status())
    console.print(get_usb_devices())
    
    score = test_cpu_performance()
    console.print(f"\n[bold green]Métrica de Desempenho da CPU:[/bold green] {score}")
    console.print("\n[bold yellow]Pressione ENTER para sair...[/bold yellow]")
    input()