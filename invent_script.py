import platform
import re
import socket
import os
import subprocess
import shutil
from datetime import datetime
import sys
import threading
import tkinter as tk
from tkinter import scrolledtext

# ---------------------------------------------------------
#                        УТИЛИТЫ
# ---------------------------------------------------------
def run(cmd, shell=False, encoding=None):
    """Выполнить команду и вернуть stdout"""
    try:
        out = subprocess.check_output(
            cmd,
            shell=shell,
            stderr=subprocess.STDOUT,
            text=True,
            encoding=encoding or "cp866",
            errors="ignore"
        )
        return out.strip()
    except Exception:
        return ""


def get_script_directory():
    """Папка рядом со скриптом или EXE"""
    if getattr(sys, 'frozen', False):
        return os.path.dirname(sys.executable)
    else:
        return os.path.dirname(os.path.abspath(__file__))


# ---------------------------------------------------------
#                      NETWORK INFO
# ---------------------------------------------------------
def gather_network_info():
    res = {}
    res["collected_at"] = datetime.now().isoformat(sep=" ", timespec="seconds")
    res["hostname"] = socket.gethostname()
    res["network"] = get_network_info()
    return res


def get_network_info():
    network_info = []

    if platform.system() == "Windows":
        out = run("ipconfig /all", shell=True)
        if out:
            sections = re.split(r'\n\s*\n', out)
            current_adapter = {}

            for section in sections:
                lines = [line.strip() for line in section.split("\n") if line.strip()]
                if not lines:
                    continue

                if ':' not in lines[0]:
                    if current_adapter:
                        network_info.append(current_adapter)

                    current_adapter = {
                        'description': lines[0],
                        'physical_address': 'N/A',
                        'ipv4': [],
                        'ipv6': [],
                        'dns_servers': [],
                        'gateway': 'N/A',
                        'dhcp_enabled': 'N/A'
                    }

                for line in lines:
                    if "Physical Address" in line or "Физический адрес" in line:
                        m = re.search(r":\s*([\w-]+)", line)
                        if m:
                            current_adapter['physical_address'] = m.group(1)

                    if "IPv4" in line:
                        m = re.search(r":\s*([\d\.]+)", line)
                        if m:
                            current_adapter['ipv4'].append(m.group(1))

                    if "IPv6" in line:
                        m = re.search(r":\s*([\w:]+)", line)
                        if m:
                            current_adapter['ipv6'].append(m.group(1))

                    if "Default Gateway" in line or "Основной шлюз" in line:
                        m = re.search(r":\s*([\d\.]+)", line)
                        if m:
                            current_adapter['gateway'] = m.group(1)

                    if "DNS" in line:
                        dns = re.findall(r"\d+\.\d+\.\d+\.\d+", section)
                        current_adapter["dns_servers"] = dns

            if current_adapter:
                network_info.append(current_adapter)

    return network_info


def print_network_report(res):
    lines = []
    lines.append("=== СЕТЕВАЯ ИНФОРМАЦИЯ ===")
    lines.append("Время сбора: {}".format(res['collected_at']))
    lines.append("Имя компьютера: {}".format(res['hostname']))

    lines.append("\n--- Адаптеры ---")

    if res["network"]:
        for i, ad in enumerate(res["network"], 1):
            lines.append("Адаптер {}: {}".format(i, ad["description"]))
            lines.append("  MAC: {}".format(ad["physical_address"]))
            lines.append("  DHCP: {}".format(ad["dhcp_enabled"]))
            lines.append("  IPv4: {}".format(", ".join(ad["ipv4"]) if ad["ipv4"] else "N/A"))
            lines.append("  IPv6: {}".format(", ".join(ad["ipv6"]) if ad["ipv6"] else "N/A"))
            lines.append("  Gateway: {}".format(ad["gateway"]))
            lines.append("  DNS: {}".format(", ".join(ad["dns_servers"]) if ad["dns_servers"] else "N/A"))
            lines.append("")
    else:
        lines.append("Нет данных")

    return "\n".join(lines)


def network():
    data = gather_network_info()
    return print_network_report(data)


# ---------------------------------------------------------
#                   СИСТЕМНАЯ ИНФОРМАЦИЯ
# ---------------------------------------------------------
def gather_system_info():
    res = {}
    res["collected_at"] = datetime.now().isoformat(sep=" ", timespec="seconds")
    res["hostname"] = socket.gethostname()

    res["os"] = {
        "system": platform.system(),
        "release": platform.release(),
        "version": platform.version(),
        "pretty": platform.platform(),
    }

    return res

def get_all_users():
    """Получить всех локальных пользователей Windows"""
    users = []
    out = run("wmic useraccount get name /format:list", shell=True)
    if out:
        for line in out.splitlines():
            if line.startswith("Name="):
                name = line.split("=", 1)[1].strip()
                if name:
                    users.append(name)
    return users


def get_last_logon_users():
    """Последний вход пользователей через `net user`"""
    logons = {}
    if platform.system() == "Windows":
        users = get_all_users()
        for user in users:
            out = run(f'net user "{user}"', shell=True, encoding="cp866")
            last_logon = "N/A"
            if out:
                for line in out.splitlines():
                    if "Последний вход" in line:
                        parts = line.split("Последний вход")
                        if len(parts) > 1:
                            last_logon = parts[1].strip()
                        break
            logons[user] = last_logon
    return logons



def print_system_report(res):
    lines = []
    lines.append("=== СИСТЕМНАЯ ИНФОРМАЦИЯ ===")
    lines.append("Время сбора: {}".format(res['collected_at']))
    lines.append("Имя компьютера: {}".format(res['hostname']))

    lines.append("\n--- ОС ---")
    lines.append("Система: {}".format(res["os"]["system"]))
    lines.append("Релиз: {}".format(res["os"]["release"]))
    lines.append("Версия: {}".format(res["os"]["version"]))
    lines.append("Полное название: {}".format(res["os"]["pretty"]))
    lines.append("\n--- Последние входы пользователей ---")
    logons = get_last_logon_users()
    if logons:
        for user, last in logons.items():
            lines.append(f"  {user}: {last}")
    else:
        lines.append("  Нет данных")

    return "\n".join(lines)


def system():
    res = gather_system_info()
    return print_system_report(res)


# ---------------------------------------------------------
#                 ОБОРУДОВАНИЕ (ПОЛНЫЙ СБОР из первого кода)
# ---------------------------------------------------------

def run_hw(cmd, shell=False, encoding=None):
    """Запуск команды и возврат stdout"""
    try:
        out = subprocess.check_output(
            cmd,
            shell=shell,
            stderr=subprocess.STDOUT,
            text=True,
            encoding=encoding or "cp866",
            errors="ignore"
        )
        return out.strip()
    except Exception:
        return ""


def get_cpu_info():
    model = platform.processor() or "Неизвестно"
    cores = os.cpu_count()
    threads = os.cpu_count()
    serial = "N/A"
    
    if platform.system() == "Windows":
        out = run_hw("wmic cpu get Name /value", shell=True)
        if out:
            for line in out.splitlines():
                if "Name=" in line:
                    full_name = line.split("=", 1)[1].strip()
                    if full_name and full_name != "N/A":
                        model = full_name
        
        out = run_hw("wmic cpu get ProcessorId /value", shell=True)
        if out:
            for line in out.splitlines():
                if "ProcessorId" in line:
                    serial = line.split("=")[1].strip()
        
        out = run_hw("wmic cpu get CurrentClockSpeed /value", shell=True)
        current_speed = "N/A"
        if out:
            for line in out.splitlines():
                if "CurrentClockSpeed=" in line:
                    speed_mhz = line.split("=")[1].strip()
                    if speed_mhz and speed_mhz.isdigit():
                        current_speed = f"{int(speed_mhz) / 1000:.2f} GHz"
        
        out = run_hw("wmic cpu get MaxClockSpeed /value", shell=True)
        max_speed = "N/A"
        if out:
            for line in out.splitlines():
                if "MaxClockSpeed=" in line:
                    speed_mhz = line.split("=")[1].strip()
                    if speed_mhz and speed_mhz.isdigit():
                        max_speed = f"{int(speed_mhz) / 1000:.2f} GHz"
        
        if "@" not in model and max_speed != "N/A":
            model = f"{model} @ {max_speed}"
    
    return {
        "model": model, 
        "cores": cores, 
        "threads": threads, 
        "serial": serial,
        "current_speed": current_speed if 'current_speed' in locals() else "N/A",
        "max_speed": max_speed if 'max_speed' in locals() else "N/A"
    }


def clean_gpu_serial(pnp_id: str) -> str:
    if not pnp_id or pnp_id == "N/A":
        return "N/A"
    
    parts = pnp_id.split("\\")
    if parts:
        last = parts[-1]
        cleaned = "".join(ch for ch in last if ch.isalnum() or ch == '-')
        return cleaned if cleaned else "N/A"
    
    return "N/A"


def get_gpu_names():
    names = []
    if platform.system() == "Windows":
        out = run_hw("wmic path win32_videocontroller get Name /value", shell=True)
        if out:
            for line in out.splitlines():
                if "Name=" in line:
                    name = line.split("=", 1)[1].strip()
                    if name and name != "N/A":
                        names.append(name)
                    else:
                        names.append("Неизвестно")
    return names


def get_gpu_memory():
    memories = []
    if platform.system() == "Windows":
        out = run_hw("wmic path win32_videocontroller get AdapterRAM /value", shell=True)
        if out:
            for line in out.splitlines():
                if "AdapterRAM=" in line:
                    memory_str = line.split("=", 1)[1].strip()
                    if memory_str and memory_str != "N/A":
                        try:
                            memory_bytes = int(memory_str)
                            memory_gb = round(memory_bytes / (1024**3), 2)
                            memories.append(memory_gb)
                        except:
                            memories.append("N/A")
                    else:
                        memories.append("N/A")
    return memories


def get_gpu_drivers():
    drivers = []
    if platform.system() == "Windows":
        out = run_hw("wmic path win32_videocontroller get DriverVersion /value", shell=True)
        if out:
            for line in out.splitlines():
                if "DriverVersion=" in line:
                    driver = line.split("=", 1)[1].strip()
                    if driver and driver != "N/A":
                        drivers.append(driver)
                    else:
                        drivers.append("N/A")
    return drivers


def get_gpu_serials():
    serials = []
    if platform.system() == "Windows":
        out = run_hw("wmic path win32_videocontroller get PNPDeviceID /value", shell=True)
        if out:
            for line in out.splitlines():
                if "PNPDeviceID=" in line:
                    serial = line.split("=", 1)[1].strip()
                    if serial and serial != "N/A":
                        cleaned_serial = clean_gpu_serial(serial)
                        serials.append(cleaned_serial)
                    else:
                        serials.append("N/A")
    return serials


def get_gpu_info():
    gpus = []
    
    names = get_gpu_names()
    memories = get_gpu_memory()
    drivers = get_gpu_drivers()
    serials = get_gpu_serials()
    
    max_gpus = max(len(names), len(memories), len(drivers), len(serials))
    
    for i in range(max_gpus):
        gpu = {}
        gpu["name"] = names[i] if i < len(names) else "Неизвестно"
        gpu["memory_gb"] = memories[i] if i < len(memories) else "N/A"
        gpu["driver_version"] = drivers[i] if i < len(drivers) else "N/A"
        gpu["serial"] = serials[i] if i < len(serials) else "N/A"
        gpus.append(gpu)
    
    return gpus


def format_gpu_info(gpu):
    parts = []
    
    name = gpu.get("name", "Неизвестно")
    parts.append(name)
    
    memory_gb = gpu.get("memory_gb")
    if memory_gb and memory_gb != "N/A":
        parts.append(f"({memory_gb} ГБ)")
    
    serial = gpu.get("serial")
    if serial and serial != "N/A":
        parts.append(f"(Серийный номер: {serial})")
    
    driver = gpu.get("driver_version")
    if driver and driver != "N/A":
        parts.append(f"[Драйвер: {driver}]")
    
    return " ".join(parts)


def get_motherboard_info():
    motherboard = {}
    if platform.system() == "Windows":
        out = run_hw("wmic baseboard get Product,Manufacturer,SerialNumber /format:list", shell=True)
        if out:
            for line in out.splitlines():
                if "=" in line:
                    k, v = line.split("=", 1)
                    if k.strip() == "Product":
                        motherboard["model"] = v.strip()
                    elif k.strip() == "Manufacturer":
                        motherboard["manufacturer"] = v.strip()
                    elif k.strip() == "SerialNumber":
                        motherboard["serial"] = v.strip()
        
        if not motherboard.get("model"):
            out = run_hw("wmic computersystem get Model /format:list", shell=True)
            if out:
                for line in out.splitlines():
                    if "Model=" in line:
                        motherboard["model"] = line.split("=")[1].strip()
    
    return motherboard


def get_ram_info():
    ram = {"total": 0, "modules": []}
    if platform.system() != "Windows":
        return ram

    capacity_out = run_hw("wmic memorychip get Capacity /format:list", shell=True)
    serial_out = run_hw("wmic memorychip get SerialNumber /format:list", shell=True)
    manufacturer_out = run_hw("wmic memorychip get Manufacturer /format:list", shell=True)
    
    if (not capacity_out or "No Instance(s) Available" in capacity_out or
        not serial_out or "No Instance(s) Available" in serial_out or
        not manufacturer_out or "No Instance(s) Available" in manufacturer_out):
        return ram

    capacities = []
    lines = capacity_out.splitlines()
    for line in lines:
        line = line.strip()
        if line.startswith("Capacity="):
            try:
                capacity_value = line.split("=", 1)[1]
                capacities.append(round(int(capacity_value) / 1024**3, 2))
            except:
                capacities.append(0)
    
    serials = []
    lines = serial_out.splitlines()
    for line in lines:
        line = line.strip()
        if line.startswith("SerialNumber="):
            serial_value = line.split("=", 1)[1]
            if serial_value and serial_value != "None" and serial_value != "":
                serials.append(serial_value)
            else:
                serials.append("N/A")
    
    manufacturers = []
    lines = manufacturer_out.splitlines()
    for line in lines:
        line = line.strip()
        if line.startswith("Manufacturer="):
            manufacturer_value = line.split("=", 1)[1]
            if manufacturer_value and manufacturer_value != "None" and manufacturer_value != "":
                manufacturers.append(manufacturer_value)
            else:
                manufacturers.append("N/A")
    
    modules = []
    for i in range(max(len(capacities), len(serials), len(manufacturers))):
        module = {}
        if i < len(capacities):
            module["capacity"] = capacities[i]
        else:
            module["capacity"] = 0
            
        if i < len(serials):
            module["serial"] = serials[i]
        else:
            module["serial"] = "N/A"
            
        if i < len(manufacturers):
            module["manufacturer"] = manufacturers[i]
        else:
            module["manufacturer"] = "N/A"
            
        modules.append(module)
    
    ram["modules"] = modules
    
    total_gb = sum(m.get("capacity", 0) for m in modules)
    ram["total"] = round(total_gb, 2)
    
    return ram


def get_bios_info():
    bios = {}
    if platform.system() == "Windows":
        out = run_hw("wmic bios get SerialNumber,SMBIOSBIOSVersion /format:list", shell=True)
        if out:
            for line in out.splitlines():
                if "=" in line:
                    k, v = line.split("=", 1)
                    bios[k.strip()] = v.strip()
    return bios


def parse_disk_table(output):
    disks = []
    
    if not output or "No Instance(s) Available" in output:
        return disks
    
    lines = [line.rstrip() for line in output.splitlines() if line.strip()]
    
    if len(lines) < 2:
        return disks
    
    header = lines[0]
    data_lines = lines[1:]
    
    media_type_pos = header.find("MediaType")
    model_pos = header.find("Model")
    serial_pos = header.find("SerialNumber")
    size_pos = header.find("Size")
    
    if media_type_pos == -1:
        media_type_pos = header.find("Media Type")
    if model_pos == -1:
        model_pos = header.find("Model")
    if serial_pos == -1:
        serial_pos = header.find("Serial Number")
    if size_pos == -1:
        size_pos = header.find("Size")
    
    if media_type_pos == -1:
        media_type_pos = 0
    if model_pos == -1:
        model_pos = 20
    if serial_pos == -1:
        serial_pos = 60
    if size_pos == -1:
        size_pos = 100
    
    for line in data_lines:
        media_type = line[media_type_pos:model_pos].strip()
        model = line[model_pos:serial_pos].strip()
        serial = line[serial_pos:size_pos].strip()
        
        if size_pos < len(line):
            size_str = line[size_pos:].strip()
        else:
            size_str = "0"
        
        disk_info = {}
        disk_info["model"] = model if model else "N/A"
        
        manufacturer = "N/A"
        if model != "N/A":
            parts = model.split()
            if parts:
                manufacturer = parts[0]
        disk_info["manufacturer"] = manufacturer
        
        disk_info["serial"] = serial if serial else "N/A"
        
        try:
            size_gb = round(int(size_str) / (1024**3), 2)
        except:
            size_gb = 0
        disk_info["size_gb"] = size_gb
        
        media_type_lower = media_type.lower()
        model_lower = model.lower()
        if ("ssd" in media_type_lower or "solid" in media_type_lower or 
            "ssd" in model_lower or "nvme" in model_lower):
            disk_info["disk_type"] = "SSD"
        elif ("hdd" in media_type_lower or "hard" in media_type_lower or 
              "hdd" in model_lower or "wd" in model_lower or "seagate" in model_lower):
            disk_info["disk_type"] = "HDD"
        else:
            disk_info["disk_type"] = "N/A"
        
        disks.append(disk_info)
    
    return disks


def get_disk_info():
    if platform.system() != "Windows":
        return []
    
    commands = [
        "wmic diskdrive get Model,SerialNumber,Size,MediaType /format:table",
        "wmic diskdrive get MediaType,Model,SerialNumber,Size /format:table",
        "wmic diskdrive list brief /format:table"
    ]
    
    for cmd in commands:
        out = run_hw(cmd, shell=True)
        if out and "No Instance(s) Available" not in out:
            disks = parse_disk_table(out)
            if disks:
                return disks
    
    return []


def get_logical_disks():
    logical_disks = []
    
    if platform.system() == "Windows":
        for drive_letter in "ABCDEFGHIJKLMNOPQRSTUVWXYZ":
            drive_path = f"{drive_letter}:\\"
            if os.path.exists(drive_path):
                try:
                    total, used, free = shutil.disk_usage(drive_path)
                    total_gb = round(total / (1024**3), 2)
                    free_gb = round(free / (1024**3), 2)
                    
                    fs_type = "N/A"
                    try:
                        fs_out = run_hw(f"fsutil fsinfo volumeinfo {drive_letter}:", shell=True)
                        if "File System Name" in fs_out:
                            for line in fs_out.splitlines():
                                if "File System Name" in line:
                                    fs_type = line.split(":")[1].strip()
                                    break
                    except:
                        pass
                    
                    logical_disks.append({
                        "drive": drive_path,
                        "size_gb": total_gb,
                        "free_gb": free_gb,
                        "filesystem": fs_type
                    })
                except:
                    pass
    
    return logical_disks


def gather_hardware_info():
    res = {}
    res["collected_at"] = datetime.now().isoformat(sep=" ", timespec="seconds")
    
    res["cpu"] = get_cpu_info()
    res["gpu"] = get_gpu_info()
    res["motherboard"] = get_motherboard_info()
    res["ram"] = get_ram_info()
    res["bios"] = get_bios_info()
    res["physical_disks"] = get_disk_info()
    res["logical_disks"] = get_logical_disks()

    return res


def print_hardware_report(res):
    lines = []
    lines.append("=== ИНФОРМАЦИЯ ОБ ОБОРУДОВАНИИ ===")
    lines.append(f"Время сбора: {res['collected_at']}")

    lines.append("\n--- Материнская плата ---")
    motherboard = res["motherboard"]
    if motherboard:
        manufacturer = motherboard.get("manufacturer", "N/A")
        model = motherboard.get("model", "N/A")
        serial = motherboard.get("serial", "N/A")
        
        if manufacturer != "N/A" or model != "N/A":
            lines.append(f"  Модель: {manufacturer} {model}")
        else:
            lines.append(f"  Модель: N/A")
        lines.append(f"  Серийный номер: {serial}")
    else:
        lines.append("  Информация недоступна")
    
    cpu = res["cpu"]
    lines.append("\n--- Процессор ---")
    lines.append(f"  Модель: {cpu['model']}")
    lines.append(f"  Ядер: {cpu['cores']}")
    lines.append(f"  Потоков: {cpu['threads']}")
    lines.append(f"  Серийный номер: {cpu['serial']}")
    if cpu.get('current_speed', 'N/A') != 'N/A':
        lines.append(f"  Текущая частота: {cpu['current_speed']}")
    if cpu.get('max_speed', 'N/A') != 'N/A' and cpu.get('current_speed', 'N/A') != cpu.get('max_speed', 'N/A'):
        lines.append(f"  Максимальная частота: {cpu['max_speed']}")

    lines.append("\n--- Видеокарты ---")
    if res["gpu"]:
        for g in res["gpu"]:
            formatted_gpu = format_gpu_info(g)
            lines.append(f"  {formatted_gpu}")
    else:
        lines.append("  Информация недоступна")

    lines.append("\n--- Оперативная память ---")
    if res["ram"]["total"]:
        lines.append(f"  Всего: {res['ram']['total']} ГБ")
        for i, m in enumerate(res["ram"]["modules"], 1):
            capacity = m.get('capacity', '?')
            manufacturer = m.get('manufacturer', 'N/A')
            serial = m.get('serial', 'N/A')
            
            line_text = f"  Модуль {i}: {manufacturer} - {capacity} ГБ"
            if serial != "N/A":
                line_text += f", Серийный номер: {serial}"
            
            lines.append(line_text)
    else:
        lines.append("  Информация недоступна")

    bios = res["bios"]
    lines.append("\n--- BIOS ---")
    if bios:
        for k, v in bios.items():
            lines.append(f"  {k}: {v}")
    else:
        lines.append("  Информация недоступна")

    lines.append("\n--- Физические диски ---")
    if res["physical_disks"]:
        for i, d in enumerate(res["physical_disks"], 1):
            lines.append(f"  Диск {i}:")
            lines.append(f"    Модель: {d.get('model', 'N/A')}")
            lines.append(f"    Производитель: {d.get('manufacturer', 'N/A')}")
            lines.append(f"    Серийный номер: {d.get('serial', 'N/A')}")
            lines.append(f"    Размер: {d.get('size_gb', 0)} ГБ")
            lines.append(f"    Тип: {d.get('disk_type', 'N/A')}")
    else:
        lines.append("  Информация недоступна")

    lines.append("\n--- Логические диски ---")
    if res["logical_disks"]:
        for d in res["logical_disks"]:
            lines.append(f"  Диск {d.get('drive', 'N/A')}:")
            lines.append(f"    Размер: {d.get('size_gb', 0)} ГБ")
            lines.append(f"    Свободно: {d.get('free_gb', 0)} ГБ")
            lines.append(f"    Файловая система: {d.get('filesystem', 'N/A')}")
    else:
        lines.append("  Информация недоступна")

    return "\n".join(lines)


def hardware():
    result = gather_hardware_info()
    return print_hardware_report(result)


# ---------------------------------------------------------
#                  СОХРАНЕНИЕ ВСЕХ ОТЧЕТОВ
# ---------------------------------------------------------
def save_all_reports():
    """Собирает и сохраняет все отчеты в один файл"""
    try:
        # Собираем все отчеты
        all_reports = []
        all_reports.append("=" * 60)
        all_reports.append("ПОЛНЫЙ ОТЧЕТ О СИСТЕМЕ")
        all_reports.append("=" * 60)
        all_reports.append(f"Сформирован: {datetime.now().isoformat(sep=' ', timespec='seconds')}")
        all_reports.append("")
        
        # Добавляем системную информацию
        all_reports.append(system())
        all_reports.append("")
        all_reports.append("")
        
        # Добавляем информацию об оборудовании
        all_reports.append(hardware())
        all_reports.append("")
        all_reports.append("")
        
        # Добавляем сетевую информацию
        all_reports.append(network())
        
        # Объединяем все в одну строку
        full_report = "\n".join(all_reports)
        
        # Сохраняем в файл
        script_dir = get_script_directory()
        filepath = os.path.join(script_dir, "full_system_report.txt")
        
        with open(filepath, "w", encoding="utf-8") as f:
            f.write(full_report)
        
        # Показываем отчет в текстовом окне
        text_box.delete(1.0, tk.END)
        text_box.insert(tk.END, full_report)
        
        return f"Полный отчет сохранен в файл: {filepath}"
        
    except Exception as e:
        return f"Ошибка при сохранении полного отчета: {e}"


# ---------------------------------------------------------
#                  ВЫВОД В ТЕКСТОВОЕ ОКНО
# ---------------------------------------------------------
def run_in_thread(func):
    threading.Thread(target=lambda: execute_and_show(func), daemon=True).start()


def execute_and_show(func):
    text_box.delete(1.0, tk.END)
    text_box.insert(tk.END, "Сбор данных... Пожалуйста, подождите.\n")
    text_box.update()
    try:
        result = func()
    except Exception as e:
        result = "Ошибка: {}".format(e)
    text_box.delete(1.0, tk.END)
    text_box.insert(tk.END, result)


# ---------------------------------------------------------
#                  СОХРАНЕНИЕ ФАЙЛА
# ---------------------------------------------------------
def save_to_file():
    text = text_box.get(1.0, tk.END).strip()
    if not text:
        return

    script_dir = get_script_directory()
    path = os.path.join(script_dir, "report.txt")

    with open(path, "w", encoding="utf-8") as f:
        f.write(text)

    print("Сохранено:", path)


# ---------------------------------------------------------
#                        TKINTER UI
# ---------------------------------------------------------
root = tk.Tk()
root.title("Системная информация")
root.geometry("900x650")

frame = tk.Frame(root)
frame.pack(pady=10)

btn1 = tk.Button(frame, text="Сетевые данные", width=20, command=lambda: run_in_thread(network))
btn1.grid(row=0, column=0, padx=5)

btn2 = tk.Button(frame, text="Оборудование", width=20, command=lambda: run_in_thread(hardware))
btn2.grid(row=0, column=1, padx=5)

btn3 = tk.Button(frame, text="Система", width=20, command=lambda: run_in_thread(system))
btn3.grid(row=0, column=2, padx=5)

btn_save = tk.Button(frame, text="Сохранить отчёт", width=20, command=save_to_file)
btn_save.grid(row=0, column=3, padx=5)

# Новая кнопка для сохранения всех отчетов
btn_save_all = tk.Button(frame, text="Сохранить все отчёты", width=20, command=lambda: run_in_thread(save_all_reports))
btn_save_all.grid(row=0, column=4, padx=5)

text_box = scrolledtext.ScrolledText(root, wrap=tk.WORD, width=110, height=32)
text_box.pack(padx=10, pady=10, fill="both", expand=True)

root.mainloop()
