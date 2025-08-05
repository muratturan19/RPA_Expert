# RPA Expert Automation

Bu proje, Preston simülatöründe (RPA_Expert.html) POS girişi işlemlerini otomatikleştiren bir örnek RPA uygulamasıdır. Python ve Selenium kullanılarak Excel dosyasından alınan verilerle form doldurma ve kaydetme işlemleri gerçekleştirilir.

## Kurulum

1. Python 3.10+ kurulu olmalıdır.
2. Google Chrome tarayıcısı gereklidir.
3. Bağımlılıkları yükleyin:

```bash
pip install -r requirements.txt
```

> Not: `webdriver-manager` modülü uygun ChromeDriver sürümünü otomatik indirir.

## Kullanım

1. POS verilerini `pos_data.xlsx` adlı bir Excel dosyasına girin. Dosya aşağıdaki sütunları içermelidir:
   - `Tarih` (YYYY-MM-DD)
   - `Firma`
   - `Tutar`
   - `Açıklama`
   - `Döviz` (opsiyonel)
   - `Vade Tarihi` (opsiyonel)
2. Otomasyonu çalıştırın:

```bash
python rpa_pos_entry.py --excel pos_data.xlsx
```

Windows kullanıcıları için, komut penceresinden `run.bat` dosyasını çalıştırmak yeterlidir.

## İş Akışı

1. Excel dosyasından satır satır POS verisi okunur.
2. Selenium ile yerel `RPA_Expert.html` dosyası açılır.
3. Menüde **Finans > Tahsilat > POS Girişi** yolunu izleyerek form açılır.
4. Form alanları doldurulur ve kaydedilir.
5. Her satır için süreç tekrar edilir.

## Loglama ve Hata Yönetimi

- Tüm işlemler `rpa.log` dosyasına kaydedilir.
- Herhangi bir hata meydana geldiğinde log dosyasına detayları yazılır ve işlem sıradaki kayıtla devam eder.

## Lisans

Bu proje eğitim amaçlı örnek olarak hazırlanmıştır.
