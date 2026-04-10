# 📊 Raporlama Portali

**Doğuş Çay - Afyon Şeker Fabrikası** için geliştirilmiş masaüstü raporlama uygulaması.

## 🎯 Özellikler

### ✅ Mevcut Modüller
- **Yan Ürünler Raporu**
  - Melas satış/üretim takibi
  - Yaş Küspe (Dökme, 25 Kg, Tonluk) 
  - Kuru Küspe (50 Kg, Dökme, Peletlenmemiş)
  - Etil Alkol (Gıda, Denatüre, Kolonya, Teknik)
  - Excel export
  - Tarih aralığı filtreleme

### 🚧 Yakında Eklenecek
- Şeker Satış Raporları (A/B/C Kotası)
- Pancar Ödemeleri (Çiftçi Borç/Alacak)

## 🚀 Kurulum (Çok Kolay!)

### Adım 1: .NET 8 SDK Yükle
Eğer yüklü değilse: https://dotnet.microsoft.com/download/dotnet/8.0

### Adım 2: Projeyi Derle
`DERLE.bat` dosyasına çift tıkla. İşlem bitince `publish` klasöründe EXE oluşacak.

### Adım 3: Çalıştır
`publish\RaporlamaPortali.exe` dosyasına çift tıkla!

Tarayıcı otomatik açılacak: **http://localhost:5050**

## 📁 Nereye Koymalıyım?

`publish` klasörünün tamamını istediğin yere kopyalayabilirsin:
- `C:\Program Files\RaporlamaPortali\`
- `C:\Users\muminceylan\Desktop\RaporlamaPortali\`
- Veya başka bir yer

Masaüstüne kısayol oluşturmak için:
1. `RaporlamaPortali.exe`'ye sağ tıkla
2. "Gönder" → "Masaüstü (kısayol oluştur)"

## ⚙️ Ayarlar

`appsettings.json` dosyasında SQL Server bağlantı bilgileri var:

```json
{
  "ConnectionStrings": {
    "LogoDB": "Server=192.168.0.50\\DOGUSLGSRV;Database=DOGUSNDB;User Id=rapor2;Password=Rpr3344@@;..."
  }
}
```

## 🔒 Güvenlik

- Uygulama sadece `localhost:5050`'de çalışır
- Başka bilgisayarlardan erişilemez
- Sadece senin bilgisayarında çalışır

## ❓ Sorun Giderme

**Uygulama açılmıyor:**
- .NET 8 SDK yüklü mü kontrol et
- Antivirüs programı engelliyor olabilir

**Tarayıcı açılmıyor:**
- Manuel olarak http://localhost:5050 adresine git

**SQL Server'a bağlanamıyor:**
- Ayarlar sayfasında "Bağlantıyı Test Et" butonuna tıkla
- VPN bağlı mı kontrol et
- Şifre doğru mu kontrol et

## 📝 Sürüm Geçmişi

- **v1.0.0** - İlk sürüm (Yan Ürünler modülü)

---

*Bu uygulama Claude AI yardımıyla geliştirilmiştir.*
