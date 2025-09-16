#!/usr/bin/env python3
"""
Setup script for Python Document Automation Program
Cross-platform installation helper
"""

import os
import sys
import platform
import subprocess
from pathlib import Path

def check_python_version():
    """Python sürümünü kontrol et"""
    if sys.version_info < (3, 8):
        print("❌ Python 3.8 veya daha yeni bir sürüm gerekli!")
        print(f"   Mevcut sürüm: {sys.version}")
        return False
    print(f"✅ Python sürümü uygun: {sys.version}")
    return True

def install_requirements():
    """Gereken paketleri kur"""
    print("\n📦 Gerekli paketler kuruluyor...")
    
    try:
        subprocess.check_call([sys.executable, "-m", "pip", "install", "-r", "requirements.txt"])
        print("✅ Python paketleri başarıyla kuruldu!")
        return True
    except subprocess.CalledProcessError:
        print("❌ Paket kurulumu başarısız!")
        return False

def check_libreoffice():
    """LibreOffice kurulu mu kontrol et (macOS/Linux için)"""
    system = platform.system()
    
    if system == "Windows":
        print("ℹ️  Windows: Microsoft Office COM interface kullanılacak")
        return True
    
    import shutil
    _libre = shutil.which("libreoffice") or shutil.which("soffice")
    # macOS default bundle path
    if not _libre and platform.system() == "Darwin":
        default_soffice = "/Applications/LibreOffice.app/Contents/MacOS/soffice"
        if os.path.exists(default_soffice):
            _libre = default_soffice
    if _libre:
        try:
            result = subprocess.run([_libre, "--version"], 
                                    capture_output=True, text=True, timeout=5)
            if result.returncode == 0:
                print(f"✅ LibreOffice tespit edildi ({_libre})")
                return True
        except (subprocess.TimeoutExpired, FileNotFoundError):
            pass
    print("❌ LibreOffice bulunamadı! PDF dönüştürme için kurulum yapın.")
    return False

def install_libreoffice_instructions():
    """LibreOffice kurulum talimatları"""
    system = platform.system()
    
    if system == "Darwin":  # macOS
        print("\n📋 macOS için LibreOffice kurulum:")
        print("   brew install libreoffice")
        print("   veya https://www.libreoffice.org/download/download/ adresinden indirin")
    elif system == "Linux":
        print("\n📋 Linux için LibreOffice kurulum:")
        print("   Ubuntu/Debian: sudo apt-get install libreoffice")
        print("   CentOS/RHEL: sudo yum install libreoffice")
        print("   Fedora: sudo dnf install libreoffice")

def create_folders():
    """Gerekli klasörleri oluştur"""
    folders = ["Evraklar", "yedekler"]
    
    for folder in folders:
        Path(folder).mkdir(exist_ok=True)
        print(f"📁 {folder} klasörü hazır")

def check_data_files():
    """Veri dosyalarını kontrol et"""
    required_files = [
        "veri_yapilandirma_GUNCEL.xlsx",
        "ANKARA İŞYERİ TABLOSU.xlsx", 
        "Nace Kod Listesi.xlsx"
    ]
    
    missing_files = []
    for file in required_files:
        if not os.path.exists(file):
            missing_files.append(file)
        else:
            print(f"✅ {file} bulundu")
    
    if missing_files:
        print(f"\n⚠️  Eksik dosyalar: {', '.join(missing_files)}")
        print("   Bu dosyalar olmadan program çalışmayabilir")
        return False
    
    return True

def main():
    """Ana kurulum fonksiyonu"""
    print("🚀 Python Document Automation Program Kurulumu")
    print("=" * 50)
    
    # Python sürümü kontrol
    if not check_python_version():
        return False
    
    # Paket kurulumu
    if not install_requirements():
        return False
    
    # LibreOffice kontrol
    if not check_libreoffice():
        install_libreoffice_instructions()
        print("\n⚠️  PDF dönüştürme özelliği çalışmayabilir")
    
    # Klasör oluşturma
    print("\n📁 Klasörler oluşturuluyor...")
    create_folders()
    
    # Veri dosyalarını kontrol
    print("\n📊 Veri dosyaları kontrol ediliyor...")
    check_data_files()
    
    print("\n" + "=" * 50)
    print("✅ Kurulum tamamlandı!")
    print("\nProgram çalıştırma:")
    print("   python3 EVRAKGENERATOR.py")
    print("   veya")
    print("   python3 FORMMODULU.py")
    
    return True

if __name__ == "__main__":
    success = main()
    sys.exit(0 if success else 1)
