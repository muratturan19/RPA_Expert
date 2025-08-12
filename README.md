# Preston RPA System

Bu proje, Preston muhasebe yazılımı için POS hareketlerinin otomatik olarak kaydedilmesi amacıyla geliştirilmiş örnek bir RPA sistemidir. Sistem, Excel dosyalarındaki POS hareketlerini okuyarak tarihe göre gruplayan bir arka plan işlemi ve Streamlit tabanlı bir kullanıcı arayüzü içerir.

## Özellikler
- Excel dosyalarından POSH ile başlayan ve 5+ rakamla biten açıklamaları filtreleme
- Tarihe göre gruplama ve toplam tutar hesaplama
- Pytesseract tabanlı OCR ile ekran üzerindeki metinleri bulma
- EasyOCR desteği ve geliştirilmiş görüntü ön işleme ile daha yüksek doğruluk
- OpenCV ile ikon eşleştirme
- Streamlit ile gerçek zamanlı ilerleme ve log görüntüleme

## Kurulum
1. Python 3.10+ kurulu olmalıdır.
2. Bağımlılıkları yükleyin:
```bash
pip install -r requirements.txt
```
3. Tesseract'ın Türkçe dil paketi kurulmuş olmalıdır (örn. `sudo apt-get install tesseract-ocr-tur`).
4. Windows ortamında tesseract ve gerekli ekran erişim izinleri hazır olmalıdır.

## Kullanım
```bash
streamlit run preston_rpa/main.py
```
Uygulama açıldığında Excel dosyanızı yükleyip `Start RPA` butonuna basın. İşlem ilerlemesi ve loglar ekranda görüntülenecektir.

## Modüller
- `excel_processor.py`: Excel dosyasını okuyup işlem grupları oluşturur.
- `ocr_engine.py`: OCR işlemlerini gerçekleştirir.
- `image_matcher.py`: İkon tespiti için OpenCV kullanır.
- `preston_automation.py`: Preston iş akışının temel adımlarını içerir.
- `main.py`: Streamlit kullanıcı arayüzü.

## Lisans
Bu proje eğitim amaçlı hazırlanmıştır.
