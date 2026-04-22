using Dapper;
using RaporlamaPortali.Models;

namespace RaporlamaPortali.Services;

/// <summary>
/// Şeker Satış rapor servisi
/// VBA Module2 makrosundaki SekerRaporHesapla fonksiyonunun C# karşılığı
/// Anlık stok takibi ile satışları önce A Kotası, yetmezse Ticari Mal'dan düşer
/// </summary>
public class SekerSatisService
{
    private readonly DatabaseService _db;

    // Kategori indeksleri (VBA'daki gibi)
    private const int A_CUVAL = 0;
    private const int A_PAKET = 1;
    private const int B_KOTASI = 2;
    private const int C_KOTASI = 3;
    private const int TICARI_CUVAL = 4;
    private const int TICARI_PAKET = 5;

    // Değer indeksleri
    private const int DEVIR = 0;
    private const int URETIM = 1;
    private const int SATINALMA = 2;
    private const int SATIS_IADE = 3;
    private const int SATINALMA_IADE = 4;
    private const int SATIS = 5;
    private const int PROMOSYON = 6;
    private const int SARF = 7;
    private const int STOK = 8;
    private const int SATIS_TUTARI = 9;

    public SekerSatisService(DatabaseService db)
    {
        _db = db;
    }

    /// <summary>
    /// 01.09.2025 gece 00:00 itibariyle devir stokları (KG cinsinden)
    /// VBA'daki GetSekerDevirStok fonksiyonu
    /// </summary>
    private decimal GetDevirStok(int kategori)
    {
        return kategori switch
        {
            A_CUVAL => 0m,
            A_PAKET => 0m,
            B_KOTASI => 6501000m,
            C_KOTASI => 3075811m,
            TICARI_CUVAL => 792780m,
            TICARI_PAKET => 50080m,
            _ => 0m
        };
    }

    /// <summary>
    /// Malzeme koduna göre temel kategori belirle
    /// VBA'daki UrunKategorisiBelirle fonksiyonu
    /// </summary>
    private string MalzemeTemelKategori(string malzemeKodu)
    {
        if (string.IsNullOrEmpty(malzemeKodu)) return "PAKET";
        
        var kod = malzemeKodu.Trim();

        // Türk Şeker Ticari Mal kodları - HARİÇ tutulacak
        if (kod == "T.T.0.0.0" || kod == "T.S.0.0.0" || 
            kod == "T.S.9.1.03.1.1000.20" || kod == "T.S.9.1.03.1.3000.06" || 
            kod == "T.S.9.1.03.1.5000.04" || kod == "T.T.9.1.03.1.5000.04")
            return "HARIC";

        // Konya Şeker Ticari Mal - S.T.1.0.0
        if (kod == "S.T.1.0.0")
            return "TICARI_CUVAL_DIREKT";

        // A Kotası Şeker (Çuval) - S.T.0.0.0, S.T.0.0.4, S.705.00.0005
        if (kod == "S.T.0.0.0" || kod == "S.T.0.0.4" || kod == "S.705.00.0005")
            return "CUVAL";

        // B Kotası Şeker - S.T.0.0.8, S.705.00.0001
        if (kod == "S.T.0.0.8" || kod == "S.705.00.0001")
            return "B_KOTASI";

        // C Kotası Şeker - S.T.0.0.7, S.705.00.0008
        if (kod == "S.T.0.0.7" || kod == "S.705.00.0008")
            return "C_KOTASI";

        // Diğer tüm kodlar = Paketli Şeker
        return "PAKET";
    }

    /// <summary>
    /// Fiş türünü standartlaştır
    /// VBA'daki SadeSekerAnaliziYap makrosundaki Case değerleri
    /// </summary>
    private string FisTuruStandart(string fisTuru)
    {
        if (string.IsNullOrEmpty(fisTuru)) return "DIGER";
        
        var f = fisTuru.Trim();

        // VBA'daki Case değerleri (birebir eşleşme)
        if (f == "Üretimden Giriş Fişi")
            return "URETIM";
        if (f == "HAMMADDE ÇEVRİM GİRİŞİ")
            return "HAMMADDE_CEVRIM_GIRIS";
        if (f == "Satınalma İrsaliyesi")
            return "SATINALMA";
        if (f == "Toptan Satış İade İrsaliyesi")
            return "SATIS_IADE";
        if (f == "Toptan Satış İrsaliyesi")
            return "SATIS";
        if (f == "PROMS ve TEKNİK M ve SARF")
            return "PROMOSYON";
        if (f == "Sarf Fişi")
            return "SARF";
        if (f == "YEMEKHANE KULLANIMI")
            return "SARF"; // Yemekhane de sarf olarak sayılıyor
        if (f == "Ambar Fişi")
            return "DIGER";

        return "DIGER";
    }

    /// <summary>
    /// Şeker satış raporunu VBA mantığıyla hesaplar
    /// Tüm verileri tarih sırasıyla çeker ve anlık stok takibi yapar
    /// </summary>
    public async Task<List<SekerSatisOzet>> GetSekerSatisOzetAsync(DateTime baslangic, DateTime bitis)
    {
        bitis = SistemTarihi.Clamp(bitis);
        // Sonuç dizisi: 6 kategori x 10 değer (9 miktar + 1 tutar)
        var sonuc = new decimal[6, 10];
        var stoklar = new decimal[6]; // Anlık stok takibi

        // Devir stoklarını al ve anlık stokları başlat
        for (int cat = 0; cat < 6; cat++)
        {
            sonuc[cat, DEVIR] = GetDevirStok(cat);
            stoklar[cat] = sonuc[cat, DEVIR];
        }

        // Tüm verileri tarih sırasıyla çek
        // VBA R sütunu (18) = CIKIS_MIKTARI_KG, S sütunu (19) = GIRIS_MIKTAR_KG kullanıyor
        var sql = @"
            SELECT 
                TARIH,
                FIS_TURU,
                MALZEME_KODU,
                GIRIS_MIKTARI_KG = ISNULL(GIRIS_MIKTAR_KG, 0),
                CIKIS_MIKTARI_KG = ISNULL(CIKIS_MIKTARI_KG, 0),
                GIRIS_TUTARI = ISNULL(GIRIS_TUTARI, 0),
                CIKIS_TUTARI = ISNULL(CIKIS_TUTARI, 0)
            FROM INF_UT_Kısıtlı_Malzeme_Raporu_Afyon_Seker_2025 WITH(NOLOCK)
            WHERE TARIH >= @Baslangic
              AND TARIH <= @Bitis
            ORDER BY TARIH ASC";

        using var conn = _db.CreateConnection();
        var hareketler = await conn.QueryAsync<dynamic>(sql, new { Baslangic = baslangic, Bitis = bitis });

        // Her satırı sırayla işle (VBA'daki gibi)
        foreach (var hareket in hareketler)
        {
            string fisTuru = hareket.FIS_TURU?.ToString() ?? "";
            string malzemeKodu = hareket.MALZEME_KODU?.ToString() ?? "";
            decimal girisMiktar = Convert.ToDecimal(hareket.GIRIS_MIKTARI_KG ?? 0m);
            decimal cikisMiktar = Convert.ToDecimal(hareket.CIKIS_MIKTARI_KG ?? 0m);
            decimal girisTutar = Convert.ToDecimal(hareket.GIRIS_TUTARI ?? 0m);
            decimal cikisTutar = Convert.ToDecimal(hareket.CIKIS_TUTARI ?? 0m);

            var temelKat = MalzemeTemelKategori(malzemeKodu);
            var fisTuruStd = FisTuruStandart(fisTuru);

            // Hariç tutulan ürünleri atla
            if (temelKat == "HARIC")
                continue;

            // B ve C Kotası her zaman kendi kategorisinde
            if (temelKat == "B_KOTASI")
            {
                IslemKaydet(sonuc, stoklar, B_KOTASI, fisTuruStd, girisMiktar, cikisMiktar, cikisTutar);
                continue;
            }

            if (temelKat == "C_KOTASI")
            {
                IslemKaydet(sonuc, stoklar, C_KOTASI, fisTuruStd, girisMiktar, cikisMiktar, cikisTutar);
                continue;
            }

            // KONYA ŞEKER - Her zaman Ticari Mal (Çuval)
            if (temelKat == "TICARI_CUVAL_DIREKT")
            {
                IslemKaydet(sonuc, stoklar, TICARI_CUVAL, fisTuruStd, girisMiktar, cikisMiktar, cikisTutar);
                continue;
            }

            // ÇUVAL tipi işlemler
            if (temelKat == "CUVAL")
            {
                switch (fisTuruStd)
                {
                    case "URETIM":
                    case "HAMMADDE_CEVRIM_GIRIS":
                        // Üretim her zaman A Kotası'na
                        sonuc[A_CUVAL, URETIM] += girisMiktar;
                        stoklar[A_CUVAL] += girisMiktar;
                        break;

                    case "SATINALMA":
                        // Satınalma her zaman Ticari Mal'a
                        sonuc[TICARI_CUVAL, SATINALMA] += girisMiktar;
                        stoklar[TICARI_CUVAL] += girisMiktar;
                        break;

                    case "SATIS_IADE":
                        // Satıştan iade A Kotası'na
                        sonuc[A_CUVAL, SATIS_IADE] += girisMiktar;
                        stoklar[A_CUVAL] += girisMiktar;
                        break;

                    case "SATINALMA_IADE":
                        // Satınalma iadesi Ticari Mal'dan
                        sonuc[TICARI_CUVAL, SATINALMA_IADE] += Math.Abs(cikisMiktar);
                        stoklar[TICARI_CUVAL] -= Math.Abs(cikisMiktar);
                        break;

                    case "SATIS":
                    case "PROMOSYON":
                    case "SARF":
                        // Çıkış işlemleri - Önce Ticari Mal, yetmezse A Kotası
                        int cikisIdx = fisTuruStd switch
                        {
                            "SATIS" => SATIS,
                            "PROMOSYON" => PROMOSYON,
                            "SARF" => SARF,
                            _ => SATIS
                        };

                        decimal cikisToplam = Math.Abs(cikisMiktar);
                        decimal cikisTutarToplam = Math.Abs(cikisTutar);

                        if (stoklar[TICARI_CUVAL] >= cikisToplam)
                        {
                            // Ticari Mal yeterli
                            sonuc[TICARI_CUVAL, cikisIdx] += cikisToplam;
                            stoklar[TICARI_CUVAL] -= cikisToplam;
                            // Satış tutarını kaydet
                            if (fisTuruStd == "SATIS")
                                sonuc[TICARI_CUVAL, SATIS_TUTARI] += cikisTutarToplam;
                        }
                        else if (stoklar[TICARI_CUVAL] > 0)
                        {
                            // Ticari Mal kısmen yeterli - tutarı oranla böl
                            decimal ticariOran = stoklar[TICARI_CUVAL] / cikisToplam;
                            sonuc[TICARI_CUVAL, cikisIdx] += stoklar[TICARI_CUVAL];
                            sonuc[A_CUVAL, cikisIdx] += (cikisToplam - stoklar[TICARI_CUVAL]);
                            stoklar[A_CUVAL] -= (cikisToplam - stoklar[TICARI_CUVAL]);
                            // Satış tutarını oranla dağıt
                            if (fisTuruStd == "SATIS")
                            {
                                sonuc[TICARI_CUVAL, SATIS_TUTARI] += cikisTutarToplam * ticariOran;
                                sonuc[A_CUVAL, SATIS_TUTARI] += cikisTutarToplam * (1 - ticariOran);
                            }
                            stoklar[TICARI_CUVAL] = 0;
                        }
                        else
                        {
                            // Ticari Mal yok, A Kotası'ndan
                            sonuc[A_CUVAL, cikisIdx] += cikisToplam;
                            stoklar[A_CUVAL] -= cikisToplam;
                            // Satış tutarını kaydet
                            if (fisTuruStd == "SATIS")
                                sonuc[A_CUVAL, SATIS_TUTARI] += cikisTutarToplam;
                        }
                        break;
                }
                continue;
            }

            // PAKET tipi işlemler
            if (temelKat == "PAKET")
            {
                switch (fisTuruStd)
                {
                    case "URETIM":
                    case "HAMMADDE_CEVRIM_GIRIS":
                        // A Kotası (Çuval) stoğu varsa A Kotası (Paket)'e, yoksa Ticari Mal (Paket)'e
                        if (stoklar[A_CUVAL] > 0)
                        {
                            sonuc[A_PAKET, URETIM] += girisMiktar;
                            stoklar[A_PAKET] += girisMiktar;
                        }
                        else
                        {
                            sonuc[TICARI_PAKET, URETIM] += girisMiktar;
                            stoklar[TICARI_PAKET] += girisMiktar;
                        }
                        break;

                    case "SATINALMA":
                        // Satınalma her zaman Ticari Mal Paket'e
                        sonuc[TICARI_PAKET, SATINALMA] += girisMiktar;
                        stoklar[TICARI_PAKET] += girisMiktar;
                        break;

                    case "SATIS_IADE":
                        // Satıştan iade
                        if (stoklar[A_PAKET] > 0 || sonuc[A_PAKET, SATIS] > 0)
                        {
                            sonuc[A_PAKET, SATIS_IADE] += girisMiktar;
                            stoklar[A_PAKET] += girisMiktar;
                        }
                        else
                        {
                            sonuc[TICARI_PAKET, SATIS_IADE] += girisMiktar;
                            stoklar[TICARI_PAKET] += girisMiktar;
                        }
                        break;

                    case "SATINALMA_IADE":
                        // Satınalma iadesi Ticari Mal Paket'ten
                        sonuc[TICARI_PAKET, SATINALMA_IADE] += Math.Abs(cikisMiktar);
                        stoklar[TICARI_PAKET] -= Math.Abs(cikisMiktar);
                        break;

                    case "SATIS":
                    case "PROMOSYON":
                    case "SARF":
                        // Çıkış işlemleri - Önce Ticari Mal Paket, yetmezse A Kotası Paket
                        int cikisIdxP = fisTuruStd switch
                        {
                            "SATIS" => SATIS,
                            "PROMOSYON" => PROMOSYON,
                            "SARF" => SARF,
                            _ => SATIS
                        };

                        decimal cikisToplamP = Math.Abs(cikisMiktar);
                        decimal cikisTutarToplamP = Math.Abs(cikisTutar);

                        if (stoklar[TICARI_PAKET] >= cikisToplamP)
                        {
                            // Ticari Mal Paket yeterli
                            sonuc[TICARI_PAKET, cikisIdxP] += cikisToplamP;
                            stoklar[TICARI_PAKET] -= cikisToplamP;
                            // Satış tutarını kaydet
                            if (fisTuruStd == "SATIS")
                                sonuc[TICARI_PAKET, SATIS_TUTARI] += cikisTutarToplamP;
                        }
                        else if (stoklar[TICARI_PAKET] > 0)
                        {
                            // Ticari Mal Paket kısmen yeterli - tutarı oranla böl
                            decimal ticariOranP = stoklar[TICARI_PAKET] / cikisToplamP;
                            sonuc[TICARI_PAKET, cikisIdxP] += stoklar[TICARI_PAKET];
                            sonuc[A_PAKET, cikisIdxP] += (cikisToplamP - stoklar[TICARI_PAKET]);
                            stoklar[A_PAKET] -= (cikisToplamP - stoklar[TICARI_PAKET]);
                            // Satış tutarını oranla dağıt
                            if (fisTuruStd == "SATIS")
                            {
                                sonuc[TICARI_PAKET, SATIS_TUTARI] += cikisTutarToplamP * ticariOranP;
                                sonuc[A_PAKET, SATIS_TUTARI] += cikisTutarToplamP * (1 - ticariOranP);
                            }
                            stoklar[TICARI_PAKET] = 0;
                        }
                        else
                        {
                            // Ticari Mal Paket yok, A Kotası Paket'ten
                            sonuc[A_PAKET, cikisIdxP] += cikisToplamP;
                            stoklar[A_PAKET] -= cikisToplamP;
                            // Satış tutarını kaydet
                            if (fisTuruStd == "SATIS")
                                sonuc[A_PAKET, SATIS_TUTARI] += cikisTutarToplamP;
                        }
                        break;
                }
            }
        }

        // Son stokları hesapla
        for (int cat = 0; cat < 6; cat++)
        {
            sonuc[cat, STOK] = sonuc[cat, DEVIR] + sonuc[cat, URETIM] + sonuc[cat, SATINALMA] + sonuc[cat, SATIS_IADE]
                              - sonuc[cat, SATINALMA_IADE] - sonuc[cat, SATIS] - sonuc[cat, PROMOSYON] - sonuc[cat, SARF];
        }

        // Sonuçları SekerSatisOzet listesine dönüştür
        var kategoriAdlari = new[]
        {
            ("A_CUVAL", "A Kotası Şeker (Çuval)"),
            ("A_PAKET", "A Kotası Şeker (Paket)"),
            ("B_KOTASI", "B Kotası Şeker"),
            ("C_KOTASI", "C Kotası Şeker"),
            ("TICARI_CUVAL", "Ticari Mal (Çuval)"),
            ("TICARI_PAKET", "Ticari Mal (Paket)")
        };

        var liste = new List<SekerSatisOzet>();
        for (int cat = 0; cat < 6; cat++)
        {
            liste.Add(new SekerSatisOzet
            {
                Kategori = kategoriAdlari[cat].Item1,
                KategoriAdi = kategoriAdlari[cat].Item2,
                DevirStok = sonuc[cat, DEVIR],
                UretimMiktari = sonuc[cat, URETIM],
                SatinAlmaMiktari = sonuc[cat, SATINALMA],
                IadeMiktari = sonuc[cat, SATIS_IADE],
                SatinAlmaIadeMiktari = sonuc[cat, SATINALMA_IADE],
                SatisMiktari = sonuc[cat, SATIS],
                PromosyonMiktari = sonuc[cat, PROMOSYON],
                SarfMiktari = sonuc[cat, SARF],
                SatisTutari = sonuc[cat, SATIS_TUTARI]
            });
        }

        return liste;
    }

    /// <summary>
    /// İşlem kaydet yardımcı fonksiyon (B ve C Kotası için)
    /// </summary>
    private void IslemKaydet(decimal[,] sonuc, decimal[] stoklar, int katIdx, string fisTuruStd, decimal girisMiktar, decimal cikisMiktar, decimal cikisTutar = 0)
    {
        switch (fisTuruStd)
        {
            case "URETIM":
            case "HAMMADDE_CEVRIM_GIRIS":
                sonuc[katIdx, URETIM] += girisMiktar;
                stoklar[katIdx] += girisMiktar;
                break;
            case "SATINALMA":
                sonuc[katIdx, SATINALMA] += girisMiktar;
                stoklar[katIdx] += girisMiktar;
                break;
            case "SATIS_IADE":
                sonuc[katIdx, SATIS_IADE] += girisMiktar;
                stoklar[katIdx] += girisMiktar;
                break;
            case "SATINALMA_IADE":
                sonuc[katIdx, SATINALMA_IADE] += Math.Abs(cikisMiktar);
                stoklar[katIdx] -= Math.Abs(cikisMiktar);
                break;
            case "SATIS":
                sonuc[katIdx, SATIS] += Math.Abs(cikisMiktar);
                sonuc[katIdx, SATIS_TUTARI] += Math.Abs(cikisTutar);
                stoklar[katIdx] -= Math.Abs(cikisMiktar);
                break;
            case "PROMOSYON":
                sonuc[katIdx, PROMOSYON] += Math.Abs(cikisMiktar);
                stoklar[katIdx] -= Math.Abs(cikisMiktar);
                break;
            case "SARF":
                sonuc[katIdx, SARF] += Math.Abs(cikisMiktar);
                stoklar[katIdx] -= Math.Abs(cikisMiktar);
                break;
        }
    }

    /// <summary>
    /// Toplam satır hesaplar
    /// </summary>
    public SekerSatisOzet HesaplaGenelToplam(List<SekerSatisOzet> ozetler)
    {
        return new SekerSatisOzet
        {
            Kategori = "TOPLAM",
            KategoriAdi = "TOPLAM",
            DevirStok = ozetler.Sum(x => x.DevirStok),
            UretimMiktari = ozetler.Sum(x => x.UretimMiktari),
            SatinAlmaMiktari = ozetler.Sum(x => x.SatinAlmaMiktari),
            IadeMiktari = ozetler.Sum(x => x.IadeMiktari),
            SatinAlmaIadeMiktari = ozetler.Sum(x => x.SatinAlmaIadeMiktari),
            SatisMiktari = ozetler.Sum(x => x.SatisMiktari),
            PromosyonMiktari = ozetler.Sum(x => x.PromosyonMiktari),
            SarfMiktari = ozetler.Sum(x => x.SarfMiktari)
        };
    }
}
