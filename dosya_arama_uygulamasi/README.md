# Dosya İçeriği Arama Motoru

Bu uygulama, belirlediğiniz bir klasörde .txt, .docx, .pdf ve .xlsx dosyalarında anahtar kelime araması yapar. Sonuçları anında listeler, önizleme sunar ve birçok kullanıcı dostu özellik içerir.

## Özellikler
- Çoklu dosya türü desteği (.txt, .docx, .pdf, .xlsx)
- Birden fazla anahtar kelimeyle arama (virgül ile ayırarak)
- Büyük/küçük harf duyarsız arama
- Hangi dosya türlerinde arama yapılacağını seçebilme
- Arama sırasında işlemi durdurabilme
- Sonuçlara çift tıklayarak dosyayı açma
- Sağ tık menüsü: Dosyayı Aç, Konumunu Aç, Yolu Kopyala
- Sonuçları TXT veya CSV olarak dışa aktarma
- Son seçilen dizini hatırlama
- Sonuçlarda içerik önizlemesi ve anahtar kelime vurgulama

## Kurulum
1. Python 3.7 veya üzeri yüklü olmalı.
2. Gerekli kütüphaneleri yükleyin:
   ```bash
   pip install PyQt5 python-docx PyPDF2 openpyxl
   ```
3. Tüm dosyaları aynı klasöre (ör: Masaüstü/dosya_arama_uygulamasi) koyun.

## Kullanım
```bash
python main.py
```

## Dosyalar
- `main.py` : Arayüz ve uygulama ana dosyası
- `file_searcher.py` : Dosya okuma ve arama yardımcı modülü
- `requirements.txt` : Gerekli Python paketleri
- `README.md` : Açıklama ve kullanım talimatları

## Notlar
- PDF ve DOCX dosyalarında bazı özel karakterler veya bozuk dosyalar okunamayabilir.
- Arama sırasında uygulama donmaz, işlemi istediğiniz an durdurabilirsiniz.
- Sonuçları kaydetmek için "Sonuçları Kaydet" butonunu kullanabilirsiniz.

Her türlü öneri ve hata bildirimi için iletişime geçebilirsiniz. 