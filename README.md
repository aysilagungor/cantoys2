# 🧸 CanToys — Türkçe Stok Karar Destek Sistemi

## GitHub + Streamlit Community Cloud kurulumu

1. GitHub üzerinde yeni bir depo oluşturun. **Bu klasörün içindekileri doğrudan depo kök dizinine yükleyin.** `app.py` ve `requirements.txt` aynı seviyede olmalıdır.
2. `stoklar1.xlsx` ve `stokk2.xlsx` dosyalarını **app.py ile aynı kök dizinde** tutun. GitHub dosya adları büyük/küçük harfe duyarlıdır.
3. [share.streamlit.io](https://share.streamlit.io) adresinde GitHub hesabınızla giriş yapıp **Create app / Deploy** seçin.
4. Depo ve branch seçin; **Main file path:** `app.py` yazın ve yayımlayın.
5. Gerekli kütüphaneler `requirements.txt` üzerinden otomatik yüklenir.

**Önemli:** Gerçek ERP Excel dosyaları depoya yüklenirse GitHub deposuna erişimi olan herkes bunları indirebilir. Hassas ticari veriler için herkese açık depo kullanmayın; erişim ve paylaşım ayarlarını kontrol edin. Streamlit Community Cloud uygulamasını herkese açık yayımlamak verileri de açığa çıkarabilir.

## Yerelde Windows üzerinde

Python 3.11 kurulu iken `BASLAT_WINDOWS.bat` dosyasını çalıştırın veya:

```powershell
py -m pip install -r requirements.txt
py -m streamlit run app.py
```

## İçerik

- `app.py`: Türkçe açıklamalı Streamlit uygulaması
- `stoklar1.xlsx` ve `stokk2.xlsx`: Kök dizindeki iki ERP hareket Excel dosyası
- `requirements.txt`: Cloud bağımlılıkları
- `runtime.txt`: Python sürümü
- `.streamlit/config.toml`: Tema ve sunucu ayarları

## Veri uyarısı

ERP hareketleri doğrulanmış güncel envanter değildir. Gerçek stok dosyası eklenmeden stok riski, sipariş ve transfer önerileri operasyonel karar olarak kullanılmamalıdır.

## Sorun giderme

- `File not found`: kök dizindeki Excel dosya adlarını GitHub'da kontrol edin.
- `ModuleNotFoundError`: `requirements.txt` kök dizinde mi kontrol edin, uygulamayı yeniden başlatın.
- `App failed to start`: Streamlit Cloud > Manage app > Logs bölümündeki hata mesajını inceleyin.
- Veri görünmüyor: Sol panelde entegre ERP dosyaları seçeneğini kontrol edin.
