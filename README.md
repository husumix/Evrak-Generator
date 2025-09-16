# Python Document Automation Program

Türkiye'de iş güvenliği evrakları için otomatik belge oluşturma programı.

## Özellikler

- 📝 **Firma Bilgi Formu**: SGK kodları ile otomatik veri doldurma
- 🔄 **SGK Geçmişi**: Son 3 SGK kodu otomatik öneri olarak gelir
- 📊 **RD Yöntemi Seçimi**: Matris veya Fine Kinney risk değerlendirme yöntemi
- 📄 **Akıllı Belge Filtreleme**: Seçilen RD yöntemine göre uygun dosyalar oluşur
- 📈 **Excel Hücre Güncelleme**: Yıllık Değerlendirme Raporu'nda otomatik RD yöntemi güncelleme
- 🎨 **Modern Arayüz**: Büyük pencere, renkli tasarım, kolay kullanım
- 📊 **PDF Dönüştürme**: Belgeleri PDF formatına çevirme
- 🖥️ **Cross-Platform**: Windows ve macOS/Linux desteği
- 🔄 **Yedekleme**: Otomatik yedek alma sistemi

## Sistem Gereksinimleri

### Genel
- Python 3.8 veya daha yeni
- Tkinter (Python ile birlikte gelir)

### Windows
- Microsoft Office (PDF dönüştürme için)
- pywin32 paketi

### macOS/Linux
- LibreOffice (PDF dönüştürme için)

## Kurulum

### 1. Otomatik Kurulum
```bash
python setup.py
```

### 2. Manuel Kurulum

#### Python paketlerini kur:
```bash
pip install -r requirements.txt
```

#### LibreOffice kur (macOS/Linux):
```bash
# macOS
brew install libreoffice

# Ubuntu/Debian
sudo apt-get install libreoffice

# CentOS/RHEL
sudo yum install libreoffice
```

## Dosya Yapısı

```
EVRAK GÜNCEL/
├── FORMMODULU.py                                         # Form modülü
├── EVRAKGENERATOR.py                                     # Ana generator
├── veri_yapilandirma_GUNCEL.xlsx                         # Veri şablonu
├── ANKARA İŞYERİ TABLOSU.xlsx                           # Şirket bilgileri
├── Nace Kod Listesi.xlsx                                 # NACE kodları
├── Evraklar/                                             # Şablon dosyaları
│   ├── *.docx                                           # Word şablonları
│   └── *.xlsx                                           # Excel şablonları
└── yedekler/                                             # Yedek dosyaları
```

## Kullanım

### 1. Ana Programı Çalıştır
```bash
python EVRAKGENERATOR.py
```

### 2. Form Uygulamasını Çalıştır
```bash
python FORMMODULU.py
```

### 3. Program Adımları

1. **Firma Bilgilerini Doldur**
   - SGK kodunu gir
   - Otomatik veri doldurma
   - Bilgileri kaydet

2. **Belge Oluştur**
   - Tüm belgeler veya seçili belgeler
   - PDF dönüştürme
   - Yedek alma

## Platform Farkları

### Windows
- Microsoft Office COM interface kullanır
- PDF dönüştürme için Office gerekli
- `pywin32` paketi otomatik kurulur

### macOS/Linux
- LibreOffice CLI kullanır
- PDF dönüştürme için LibreOffice gerekli
- Platform uyumlu font seçimi

## Sorun Giderme

### PDF Dönüştürme Çalışmıyor
- **Windows**: Microsoft Office kurulu mu?
- **macOS/Linux**: LibreOffice kurulu mu?

### Font Sorunları
- Program otomatik olarak platform uyumlu font seçer
- Windows: Segoe UI
- macOS: SF Pro Display
- Linux: Ubuntu

### Dosya Yolu Sorunları
- Tüm dosya yolları otomatik olarak platform uyumlu hale getirilir
- Boşluklu dosya adları desteklenir

## Geliştirici Notları

### Kod Değişiklikleri
- Platform detection eklendi
- Cross-platform file handling
- Font compatibility
- LibreOffice integration

### Güvenlik
- Dosya yolları sanitize edilir
- Güvenli subprocess kullanımı
- Error handling iyileştirildi

## Lisans

Bu program Hüseyin İLHAN tarafından geliştirilmiştir.

## Destek

Sorunlar için:
1. Log dosyalarını kontrol edin (`evrak_generator.log`, `form_debug.log`)
2. Setup scriptini tekrar çalıştırın
3. Sistem gereksinimlerini doğrulayın
