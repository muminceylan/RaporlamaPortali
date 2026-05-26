using Dapper;
using RaporlamaPortali.Models;

namespace RaporlamaPortali.Services.Logo;

// EFatura → Logo aktarımı için orchestrator.
//   1) EFaturaService.DetayGetirAsync ile UBL'i parse et
//   2) Karşı VKN/TCKN için Logo'da cari adaylarını çıkar (LogoCariLookup)
//   3) Hizmet/Mal seçimine göre LogoFaturaModel ön-yükle (KDV maps, tevkifat 6→3, GL_CODE2 boş)
//   4) Dialog'da kullanıcı eksik alanları (MASTER_CODE, Fabrika, Ambar, MasrafMerkezi, Cari seçimi) tamamlasın
//   5) LogoUnityComService.AktarAsync ile gönder
public class LogoAktarimService
{
    private readonly EFaturaService          _efatura;
    private readonly LogoCariLookupService   _cariSvc;
    private readonly LogoAccCodesLookupService _accSvc;
    private readonly EFaturaAktarimGecmisiService _gecmis;
    private readonly DatabaseService         _db;
    private readonly ILogger<LogoAktarimService> _log;

    public LogoAktarimService(EFaturaService efatura, LogoCariLookupService cariSvc,
                              LogoAccCodesLookupService accSvc,
                              EFaturaAktarimGecmisiService gecmis,
                              DatabaseService db, ILogger<LogoAktarimService> log)
    {
        _efatura = efatura;
        _cariSvc = cariSvc;
        _accSvc  = accSvc;
        _gecmis  = gecmis;
        _db      = db;
        _log     = log;
    }

    // Fatura tarihindeki USD/TL kuru (LG_EXCHANGE_FIRMA, CRTYPE=1=USD, RATES2=TCMB Satış).
    // Logo Raporlama Dövizi (REPORTRATE) için bu kur kullanılır.
    private async Task<decimal> UsdKuruAsync(DateTime tarih, CancellationToken ct = default)
    {
        try
        {
            var sql = $@"SELECT TOP 1 RATES2 FROM LG_EXCHANGE_{_db.FirmaNo} WITH(NOLOCK)
                         WHERE CRTYPE = 1 AND EDATE <= @t AND RATES2 > 0
                         ORDER BY EDATE DESC";
            using var conn = _db.CreateConnection();
            var kur = await conn.QueryFirstOrDefaultAsync<decimal?>(
                new CommandDefinition(sql, new { t = tarih.Date }, cancellationToken: ct));
            return kur ?? 0m;
        }
        catch (Exception ex)
        {
            _log.LogWarning(ex, "USD kuru lookup başarısız (tarih {T})", tarih);
            return 0m;
        }
    }

    // UI dialog'unun her bir satır için tuttuğu state — kullanıcı buradan düzenler.
    public class AktarimSatirState
    {
        public long             EFaturaId   { get; set; }
        public EFaturaDetay     Detay       { get; set; } = new();
        public int              Type        { get; set; } = 4;   // 4=Alınan Hizmet, 1=Satın Alma

        // Cari adayları — VKN/TCKN ile bulunan tüm cariler. >1 ise dropdown çıkar.
        public List<LogoCariLookupService.CariAdayi> CariAdaylari { get; set; } = new();
        public LogoCariLookupService.CariAdayi?       SeciliCari   { get; set; }

        // Header düzeyi alanlar
        public int       Fabrika       { get; set; } = 17;
        public int       Ambar         { get; set; } = 147;
        public string    OdemeKodu     { get; set; } = "PESIN";
        public string    OzelKod       { get; set; } = "AFYON";

        // Satır düzeyi — UBL kalemlerinden 1:1 oluşur
        public List<AktarimKalemState> Kalemler { get; set; } = new();

        // Mükerrer kontrol (HazirlaAsync'te doldurulur)
        public LogoCariLookupService.FaturaMevcutSonuc MevcutKontrol { get; set; } = new();
        // Kullanıcı dialog'da fatura no'yu manuel düzenleyebilir (boş ise Detay.FaturaNo kullanılır)
        public string FaturaNoOverride { get; set; } = "";
        public string EtkinFaturaNo => string.IsNullOrWhiteSpace(FaturaNoOverride)
            ? Detay.FaturaNo
            : FaturaNoOverride.Trim();

        // Hata listesi (cari yok, GL_CODE bulunamadı vs.)
        public List<string> Uyarilar { get; set; } = new();
        public bool         GonderilebilirMi => Uyarilar.Count == 0
                                                && MevcutKontrol.Durum != LogoCariLookupService.FaturaMevcutDurumu.AyniVknMevcut
                                                && SeciliCari != null
                                                && SeciliCari.GlKodVar
                                                && Kalemler.All(k => !string.IsNullOrWhiteSpace(k.MasterCode)
                                                                  // İstisnalı (KDV %0) kalemde KDV hesabı zorunlu değil
                                                                  && (k.KdvOran == 0 || !string.IsNullOrWhiteSpace(k.KdvHesabi)));
    }

    public class AktarimKalemState
    {
        public string    MasterCode    { get; set; } = "";   // kullanıcı seçer
        public decimal   Miktar        { get; set; }
        // Birim = Logo'ya gidecek yerel birim kodu (KG, TON, GR…). Kullanıcı kart seçince
        // BirimSecenekler doldurulur, UBL'deki kod ile en uygun yerel kod default seçilir.
        public string    Birim         { get; set; } = "ADET";
        public string    UbliBirim     { get; set; } = "";   // UBL'den ham gelen kod (KGM/TNE/EA…) — referans/log
        public List<LogoBirimSetiLookupService.BirimSatir> BirimSecenekler { get; set; } = new();
        public decimal   BirimFiyat    { get; set; }
        public decimal   ToplamNet     { get; set; }
        public decimal   KdvOran       { get; set; }
        public decimal   KdvTutar      { get; set; }

        public string    MasrafMerkezi { get; set; } = "7.04";

        // Otomatik doldurulur (KDV oranı + tevkifatlı mı)
        public string    KdvHesabi          { get; set; } = "";
        public string    HizmetMalHesabi    { get; set; } = "";   // boş bırakılır — Logo kart default'unu kullanır
        public string    TevkifatKdvHesabi  { get; set; } = "";
        public string    SorumluKdvHesabi   { get; set; } = "";

        public bool      Tevkifatli    { get; set; }
        public int       TevkifatPay   { get; set; }
        public int       TevkifatPayda { get; set; }
        public decimal   TevkifatTutar { get; set; }
        public string    TevkifatKodu  { get; set; } = "";   // 3 ile başlayan (alıcı tarafı)

        public string    UbliIstisnaKodu  { get; set; } = "";  // VATEXCEPT_CODE (UBL kodu — Logo'ya bu yazılır)
        public string    UbliIstisnaSebep { get; set; } = "";  // VATEXCEPT_REASON (sebep metni)

        // Kullanıcıya gösterilen original UBL açıklaması
        public string    UbliAciklama  { get; set; } = "";
    }

    // EFatura listesinden (id, type) çiftleri için pre-fill yapar.
    public async Task<List<AktarimSatirState>> HazirlaAsync(IReadOnlyList<(long id, int type)> secimler)
    {
        var sonuc = new List<AktarimSatirState>();
        foreach (var (id, type) in secimler)
        {
            var detay = await _efatura.DetayGetirAsync(id);
            if (detay == null) continue;

            var state = new AktarimSatirState
            {
                EFaturaId = id,
                Detay     = detay,
                Type      = type,
                Fabrika   = 17,
                Ambar     = 147,
            };

            // 1) Karşı VKN ile cari adaylarını bul
            var vkn = string.IsNullOrWhiteSpace(detay.SaticiVkn) ? detay.AliciVkn : detay.SaticiVkn;
            if (!string.IsNullOrWhiteSpace(vkn))
            {
                state.CariAdaylari = await _cariSvc.BulAsync(vkn);
                if (state.CariAdaylari.Count == 1)
                    state.SeciliCari = state.CariAdaylari[0];
            }

            if (state.CariAdaylari.Count == 0)
                state.Uyarilar.Add($"VKN/TCKN {vkn} için Logo'da cari bulunamadı. Önce cari kartı açın.");
            else if (state.SeciliCari == null)
                state.Uyarilar.Add($"Bu VKN için birden fazla cari var ({state.CariAdaylari.Count} adet). Lütfen birini seçin.");
            else if (!state.SeciliCari.GlKodVar)
                state.Uyarilar.Add($"Cari {state.SeciliCari.Kod} için Muhasebe Hesap Kodu (GL_CODE) tanımlı değil. Logo'da cari kartını açın → Diğer sekmesi → Hesap Kodu doldurun.");

            // 1.5) Mükerrer / duplicate kontrolü — Logo'da bu fatura no zaten var mı?
            // NOT: Uyarilar listesine eklemiyoruz — dialog üst kısmında MevcutKontrol state'ine
            // göre kırmızı/turuncu alert gösteriliyor. Uyarilar SADECE blocker'lar için.
            // GonderilebilirMi check'i de zaten AyniVknMevcut durumunda aktarımı engelliyor.
            state.MevcutKontrol = await _cariSvc.FaturaMevcutMuAsync(detay.FaturaNo, vkn);
            if (state.MevcutKontrol.Durum == LogoCariLookupService.FaturaMevcutDurumu.BaskaVknMevcut)
            {
                // Başka VKN'de varsa K-prefix otomatik uygulansın — kullanıcı isterse dialog'dan değiştirir.
                state.FaturaNoOverride = FaturaNoYardimcisi.KEkle(detay.FaturaNo);
            }

            // 2) Her UBL kalemini AktarimKalemState'e çevir
            foreach (var k in detay.Kalemler)
            {
                var tevkifatVar = k.TevkifatTutar > 0 && !string.IsNullOrWhiteSpace(k.TevkifatKodu);
                var (pay, payda) = tevkifatVar ? TevkifatPayPaydaCevirici.Bol(k.TevkifatOran) : (0, 0);

                var ks = new AktarimKalemState
                {
                    Miktar       = k.Miktar,
                    Birim        = string.IsNullOrWhiteSpace(k.Birim) ? "ADET" : k.Birim,
                    UbliBirim    = (k.Birim ?? "").Trim(),
                    BirimFiyat   = k.BirimFiyat,
                    ToplamNet    = k.NetTutar,
                    KdvOran      = k.KdvOran,
                    KdvTutar     = k.KdvTutar,

                    Tevkifatli   = tevkifatVar,
                    TevkifatPay  = pay,
                    TevkifatPayda= payda,
                    TevkifatTutar= k.TevkifatTutar,
                    TevkifatKodu = TevkifatKoduCevirici.Cevir(k.TevkifatKodu),

                    UbliIstisnaKodu  = k.IstisnaKodu,
                    UbliIstisnaSebep = k.IstisnaSebep,
                    UbliAciklama     = string.IsNullOrWhiteSpace(k.StokAdi) ? k.Aciklama : k.StokAdi,
                };

                // KDV indirilecek hesabı (191.01.xx / 192.02.xx)
                ks.KdvHesabi = LogoHesapKoduMaps.KdvHesabiBul(k.KdvOran, tevkifatVar);
                if (tevkifatVar)
                {
                    ks.TevkifatKdvHesabi = LogoHesapKoduMaps.KdvHesabiBul(k.KdvOran, tevkifatli: true);
                    ks.SorumluKdvHesabi  = LogoHesapKoduMaps.SorumlulukHesabiBul(pay, payda);

                    if (string.IsNullOrWhiteSpace(ks.TevkifatKdvHesabi))
                        state.Uyarilar.Add($"Kalem {k.Sira}: KDV oranı %{k.KdvOran} için tevkifatlı indirilecek KDV hesap kodu maplenmemiş.");
                    if (string.IsNullOrWhiteSpace(ks.SorumluKdvHesabi))
                        state.Uyarilar.Add($"Kalem {k.Sira}: tevkifat oranı {pay}/{payda} için sorumlu KDV hesap kodu maplenmemiş.");
                }
                else
                {
                    if (string.IsNullOrWhiteSpace(ks.KdvHesabi) && k.KdvOran > 0)
                        state.Uyarilar.Add($"Kalem {k.Sira}: KDV oranı %{k.KdvOran} için indirilecek KDV hesap kodu maplenmemiş.");
                }

                state.Kalemler.Add(ks);
            }

            sonuc.Add(state);
        }
        return sonuc;
    }

    // State listesi + Logo Unity credentials ile gerçek aktarımı yapar.
    public async Task<LogoAktarimSonuc> AktarAsync(
        LogoUnityCredentials cred,
        IReadOnlyList<AktarimSatirState> stateler,
        LogoUnityComService comSvc,
        CancellationToken ct = default)
    {
        var batch = new List<(long, LogoFaturaModel)>();

        foreach (var s in stateler)
        {
            if (!s.GonderilebilirMi)
            {
                _log.LogWarning("Fatura {Id} eksik bilgi ile atlandı: {U}", s.EFaturaId, string.Join("; ", s.Uyarilar));
                continue;
            }

            var faturaTarihi = s.Detay.Tarih ?? DateTime.Today;
            var raporKuru = await UsdKuruAsync(faturaTarihi);

            var etkinNo   = s.EtkinFaturaNo;     // kullanıcı override veya K-prefix uygulanmış olabilir → FİŞ NO
            var orjinalNo = s.Detay.FaturaNo;    // her zaman gerçek e-Fatura no → BELGE NO
            var m = new LogoFaturaModel
            {
                Type        = s.Type,
                FaturaNo    = etkinNo,             // NUMBER (FICHENO) = Logo Fiş No (override varsa K'lı)
                BelgeNo     = orjinalNo,           // DOC_NUMBER (DOCODE) = Logo Belge No (her zaman orijinal e-Fatura no)
                BelgeTarihi = faturaTarihi,
                Tarih       = faturaTarihi,
                CariKod     = s.SeciliCari!.Kod,
                CariGlKod   = s.SeciliCari!.GlKod,
                Fabrika     = s.Fabrika,
                Ambar       = s.Ambar,
                OdemeKodu   = s.OdemeKodu,
                OzelKod     = s.OzelKod,
                Doviz       = string.IsNullOrWhiteSpace(s.Detay.Doviz) ? "TL" : s.Detay.Doviz,
                DovizKuru   = s.Detay.DovizKuru > 0 ? s.Detay.DovizKuru : 1m,
                EdtCurr     = DovizKodunuLogoNumaraya(s.Detay.Doviz),
                RaporlamaKuru = raporKuru,

                ToplamMatrah   = s.Detay.TaxExclusiveAmount,
                ToplamKdv      = s.Detay.TaxInclusiveAmount - s.Detay.TaxExclusiveAmount,
                ToplamNet      = s.Detay.PayableAmount,
                ToplamTevkifat = s.Kalemler.Sum(k => k.TevkifatTutar),
                Tevkifatli     = s.Kalemler.Any(k => k.Tevkifatli),
                ProfileId      = ProfileIdToLogoNumara(s.Detay.ProfileId),
                EFaturaTarihi  = faturaTarihi,
            };

            foreach (var k in s.Kalemler)
            {
                // GL_CODE1 (Mal/Hizmet hesabı) öncelik sırası:
                //   1) Kullanıcı dialog'da manuel girdi → onu kullan
                //   2) Logo kartının "Muhasebe Hesapları → Alımlar Hesabı" (LG_211_CRDACREF TYP=1)
                //   3) Yerel prefix mapping (LogoHesapKoduMaps) — Logo kartında tanımsızsa
                // Hiçbiri bulunamazsa boş; Logo'nun XML modülü kontrolü ileride hata verirse
                // kullanıcı dialog'da manuel girmek zorunda kalır.
                var malHesabi = k.HizmetMalHesabi;
                if (string.IsNullOrWhiteSpace(malHesabi))
                    malHesabi = await _accSvc.AlimHesabiAsync(k.MasterCode, s.Type);
                if (string.IsNullOrWhiteSpace(malHesabi))
                    malHesabi = LogoHesapKoduMaps.MalzemeMuhasebeHesabiBul(k.MasterCode, s.Type);

                m.Satirlar.Add(new LogoFaturaSatirModel
                {
                    // VBA referansı (komur_macros.txt:145 & hizmet_macros.txt:148):
                    //   Satın Alma (header TYPE=1) → line TYPE=0  (Mal alımı)
                    //   Alınan Hizmet (header TYPE=4) → line TYPE=4 (Hizmet)
                    // line TYPE=1 = İade faturası satırı; modüle uygun olmadığı için Logo
                    // "(1159) Fiş tip bilgisi fiş modülüne uygun değil" verir.
                    SatirTipi          = s.Type == 1 ? 0 : s.Type,
                    MasterCode         = k.MasterCode,
                    Miktar             = k.Miktar,
                    Birim              = k.Birim,
                    BirimFiyat         = k.BirimFiyat,
                    ToplamNet          = k.ToplamNet,
                    KdvOran            = k.KdvOran,
                    KdvTutar           = k.KdvTutar,
                    // Masraf Merkezi sadece Alınan Hizmet (TYPE=4) faturalarında kullanılır;
                    // Satın Alma'da (TYPE=1) Logo masraf merkezi atamasını kabul etmez.
                    MasrafMerkezi      = s.Type == 4 ? k.MasrafMerkezi : "",
                    IstisnaKodu        = k.UbliIstisnaKodu,
                    IstisnaSebep       = k.UbliIstisnaSebep,
                    KdvHesabi          = k.KdvHesabi,
                    HizmetMalHesabi    = malHesabi,
                    TevkifatKdvHesabi  = k.TevkifatKdvHesabi,
                    SorumluKdvHesabi   = k.SorumluKdvHesabi,
                    Tevkifatli         = k.Tevkifatli,
                    TevkifatPay        = k.TevkifatPay,
                    TevkifatPayda      = k.TevkifatPayda,
                    TevkifatTutar      = k.TevkifatTutar,
                    TevkifatKodu       = k.TevkifatKodu,
                });
            }

            batch.Add((s.EFaturaId, m));
        }

        if (batch.Count == 0)
        {
            return new LogoAktarimSonuc
            {
                LoginBasarili = false,
                LoginHata     = "Aktarılacak fatura yok (hepsi eksik bilgi nedeniyle atlandı)."
            };
        }

        var sonuc = await comSvc.AktarAsync(cred, batch, ct);

        // Başarılı her satır için yerel aktarım geçmişine kayıt — "E-Fatura Aktar"
        // menüsünde aynı faturanın tekrar aktarılmasını engeller.
        foreach (var f in sonuc.Faturalar.Where(x => x.Basarili))
        {
            try
            {
                int faturaTipi = stateler.FirstOrDefault(s => s.EFaturaId == f.EFaturaId)?.Type ?? 0;
                await _gecmis.KaydetAsync(new EFaturaAktarimGecmisiService.Kayit
                {
                    EFaturaId     = f.EFaturaId,
                    FaturaNo      = f.FaturaNo,
                    LogoNumber    = f.LogoNo,
                    LogoRef       = f.LogoRef,
                    FaturaTipi    = faturaTipi,
                    AktarimTarihi = DateTime.Now,
                }, ct);
            }
            catch (Exception ex)
            {
                _log.LogWarning(ex, "Aktarım geçmişi kaydedilemedi: EFaturaId={Id}", f.EFaturaId);
            }
        }
        return sonuc;
    }

    // Logo Tiger döviz numaraları — L_CURRENCYLIST.CURTYPE (DOGUSNDB firma 211'de doğrulandı):
    //   1=USD, 17=GBP (İngiliz Sterlini), 20=EUR (Euro), 2=DEM (Alman Markı)
    // Eskiden EUR=17, GBP=2 yazıyordu — YANLIŞ. EUR fatura GBP sanılıp TL'ye düşüyordu.
    private static int DovizKodunuLogoNumaraya(string? kod) => (kod ?? "TL").ToUpperInvariant() switch
    {
        ""     => 0,
        "TL"   => 0,
        "TRY"  => 0,
        "USD"  => 1,
        "EUR"  => 20,
        "GBP"  => 17,
        _      => 0,
    };

    private static int ProfileIdToLogoNumara(string? profile) => (profile ?? "").ToUpperInvariant() switch
    {
        "TEMELFATURA"   => 1,
        "TICARIFATURA"  => 2,
        "EARSIVFATURA"  => 3,
        "IHRACAT"       => 5,
        _               => 2,
    };
}
