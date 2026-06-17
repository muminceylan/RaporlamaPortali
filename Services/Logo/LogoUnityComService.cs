using System.Reflection;
using System.Runtime.InteropServices;
using RaporlamaPortali.Models;

namespace RaporlamaPortali.Services.Logo;

// Logo Tiger 3 Enterprise Unity COM API'yi late-binding ile çağırır.
// VBA makrosundaki ProcessInvoice() ile aynı alanları doldurur ve Post() ile gönderir.
//
// ÖNEMLİ:
//   1) Logo Unity COM nesneleri STA thread gerektirir. Blazor Server thread'leri MTA,
//      o yüzden her batch için ayrı bir STA thread başlatıp Login + tüm faturalar + Logout
//      bu thread üzerinde sıralı çalıştırılıyor.
//   2) Tip kütüphanesini referans etmiyoruz; "UnityObjects.UnityApplication" ProgID'sini
//      Type.GetTypeFromProgID ile alıyoruz.
//   3) .NET 8'de dynamic COM IDispatch dispatch'i RuntimeBinder ile sorunlu olduğu için
//      tüm dispatch çağrıları Type.InvokeMember ile yapılıyor (Inv/Get/Set helper'ları).
//   4) DataObjectType sabitleri — KeyNet PMHS.dll v16.02.2026 decompile'ından KESİN değerler:
//        Type=18 → doPurchInvoice = Satın Alma Faturası   ← BURADA KULLANIYORUZ
//        Type=19 → doSalesInvoice = Satış Faturası        (Toptan Satış — SabnetFaturaTransfer)
//        Type=23 → doBankAccount = Banka Hesabı (yanlış değer — V7'de bu sabit denenmişti)
//      Her iki fatura tipi de aynı LG_211_01_INVOICE tablosuna yazar (Count=531 field) ama
//      Logo MODÜL kontrolü yapar: yanlış DataObjectType → "(1159) Fiş tip bilgisi fiş modülüne
//      uygun değil". SabnetFaturaTransfer'da 14 May 2026'da bu kuralla 1159 çözüldü.
public class LogoUnityComService
{
    private readonly ILogger<LogoUnityComService> _log;

    private const int DO_PURCH_INVOICE = 18;

    // Oturum başına bir kez DataObjectType taraması yapılır — doğru sabiti bulmak için.
    private static bool _dataObjectTypeProbed;

    public LogoUnityComService(ILogger<LogoUnityComService> log)
    {
        _log = log;
    }

    public Task<LogoAktarimSonuc> AktarAsync(
        LogoUnityCredentials cred,
        IReadOnlyList<(long eFaturaId, LogoFaturaModel model)> batch,
        CancellationToken ct = default)
    {
        var tcs = new TaskCompletionSource<LogoAktarimSonuc>(TaskCreationOptions.RunContinuationsAsynchronously);

        var th = new Thread(() =>
        {
            try
            {
                var sonuc = ExecuteBatchOnStaThread(cred, batch, ct);
                tcs.TrySetResult(sonuc);
            }
            catch (Exception ex)
            {
                _log.LogError(ex, "Logo Unity batch hata");
                tcs.TrySetException(ex);
            }
        });

        th.SetApartmentState(ApartmentState.STA);
        th.IsBackground = true;
        th.Name = "LogoUnityStaThread";
        th.Start();

        return tcs.Task;
    }

    public Task<(bool ok, string message)> TestLoginAsync(LogoUnityCredentials cred)
    {
        var tcs = new TaskCompletionSource<(bool, string)>(TaskCreationOptions.RunContinuationsAsynchronously);
        var th = new Thread(() =>
        {
            object? app = null;
            try
            {
                app = CreateUnityApp();
                var ok = (bool)(Inv(app, "Login", cred.Kullanici, cred.Sifre, cred.FirmaNo) ?? false);
                if (!ok)
                {
                    tcs.TrySetResult((false, $"Login başarısız: {GetStr(app, "GetLastError")} {GetStr(app, "GetLastErrorString")}"));
                    return;
                }
                try { Inv(app, "Logout"); } catch { /* ignore */ }
                tcs.TrySetResult((true, "Logo Unity login başarılı."));
            }
            catch (Exception ex)
            {
                tcs.TrySetResult((false, ex.Message));
            }
            finally
            {
                Release(app);
            }
        });
        th.SetApartmentState(ApartmentState.STA);
        th.IsBackground = true;
        th.Start();
        return tcs.Task;
    }

    // ---------- STA thread içinden çalışan ana batch işleyici ----------

    private LogoAktarimSonuc ExecuteBatchOnStaThread(
        LogoUnityCredentials cred,
        IReadOnlyList<(long, LogoFaturaModel)> batch,
        CancellationToken ct)
    {
        var sonuc = new LogoAktarimSonuc();
        object? app = null;
        bool loggedIn = false;

        try
        {
            app = CreateUnityApp();

            bool ok;
            try
            {
                ok = (bool)(Inv(app, "Login", cred.Kullanici, cred.Sifre, cred.FirmaNo) ?? false);
            }
            catch (Exception ex)
            {
                sonuc.LoginHata = $"Login çağrısı başarısız: {ex.Message}";
                return sonuc;
            }

            if (!ok)
            {
                sonuc.LoginHata = $"Logo Login reddetti. {GetStr(app, "GetLastError")} - {GetStr(app, "GetLastErrorString")}";
                return sonuc;
            }

            loggedIn = true;
            sonuc.LoginBasarili = true;

            foreach (var (id, model) in batch)
            {
                ct.ThrowIfCancellationRequested();
                var satir = PostSingleInvoice(app!, id, model);
                sonuc.Faturalar.Add(satir);
            }
        }
        finally
        {
            if (loggedIn && app != null)
            {
                try { Inv(app, "Logout"); } catch { /* ignore */ }
            }
            Release(app);
        }

        return sonuc;
    }

    // ---------- Tek fatura post — VBA ProcessInvoice eşdeğeri ----------

    private LogoFaturaAktarimSatir PostSingleInvoice(object app, long eFaturaId, LogoFaturaModel m)
    {
        var satir = new LogoFaturaAktarimSatir
        {
            EFaturaId = eFaturaId,
            FaturaNo  = m.BelgeNo,
        };

        object? inv = null;
        try
        {
            // ESKİ V7 bulgusu: NewDataObject(23) bize satın alma faturası değil, CARİ HESAP KARTI
            // (CLCARD: LOGICALREF/CARDTYPE/IBAN/BANKREF/STOPAJ_PER...) döndürüyordu. Yani 23 sabiti
            // bu Logo Unity sürümünde doPurchInvoice değil. Doğru sayıyı keşfetmek için ilk faturada
            // 1..50 aralığını tarayıp her birinin TableName + ilk 3 field adını log'a yazıyoruz.
            if (!_dataObjectTypeProbed)
            {
                _dataObjectTypeProbed = true;
                try { DiscoverDataObjectTypes(app, eFaturaId.ToString()); }
                catch (Exception dx) { _log.LogWarning(dx, "DataObjectType keşif probe başarısız"); }
            }

            inv = Inv(app, "NewDataObject", DO_PURCH_INVOICE);
            if (inv == null)
            {
                satir.Hata = $"NewDataObject({DO_PURCH_INVOICE}) null döndü — Logo Unity bağlantısı veya yetki sorunu.";
                return satir;
            }

            // YENİ: mevcut DataObject'in TableName'ini logla — gerçekten faturayı mı yoksa cari kartı mı tutuyor görelim
            try
            {
                object? tn = IDispCall(inv, "TableName", DISPATCH_PROPERTYGET, Array.Empty<object?>());
                _log.LogInformation("LOGO inv TableName={Tn} (DO_PURCH_INVOICE={Const})", tn ?? "<null>", DO_PURCH_INVOICE);
                try
                {
                    var logFile = System.IO.Path.Combine(AppDataPaths.DataRoot, "logo_unity_diag.txt");
                    System.IO.File.AppendAllText(logFile,
                        $"---- {DateTime.Now:yyyy-MM-dd HH:mm:ss} CURRENT-DO-TABLENAME eFaturaId={eFaturaId} type={DO_PURCH_INVOICE} TableName={tn ?? "<null>"} ----\n");
                }
                catch { }
            }
            catch (Exception tnEx) { _log.LogWarning(tnEx, "TableName probe başarısız"); }

            // VBA: purcinvoices.New — return değerini de yakala (false dönerse schema yüklenmedi demektir)
            string newErr = "";
            string newRet = "?";
            try
            {
                object? newRes = Inv(inv, "New");
                newRet = newRes == null ? "null" : newRes.ToString() ?? "?";
            }
            catch (Exception nex) { newErr = $"DataObject.New(): {nex.GetType().Name}: {nex.Message}"; _log.LogWarning(nex, "DataObject.New() uyarı verdi"); }
            string newAppErr = $"{GetStr(app, "GetLastError")}/{GetStr(app, "GetLastErrorString")}";

            string invTypeInfo = $"inv={inv.GetType().FullName} IsCom={Marshal.IsComObject(inv)}";

            object? dfRaw;
            try
            {
                dfRaw = GetProp(inv, "DataFields");
            }
            catch (Exception dex)
            {
                throw new InvalidOperationException($"DataFields property çağrısı patladı. {invTypeInfo}. {newErr} | {dex.GetType().Name}: {dex.Message}", dex);
            }
            if (dfRaw == null)
                throw new InvalidOperationException($"DataFields null döndü. {invTypeInfo}. {newErr}");
            object df = dfRaw;

            ResetSetFieldDiag();

            // ---- TANI: hangi dispatch mekanizması çalışıyor? ----
            // Bu blok, FieldByName için farklı çağrı yöntemlerini deneyip ilkini bulur.
            // Sonuç _setFieldFirstError'a yazılır (boşsa hepsi başarısız, doluysa kullanılan yöntem).
            string dispatchDiag = ProbeDispatch(inv, df);

            // ---- HEADER ---- VBA pattern: inv.DataFields.FieldByName(...) freshly each time
            SetHeaderField(inv, "TYPE",         m.Type);
            if (!string.IsNullOrWhiteSpace(m.FaturaNo))
                SetHeaderField(inv, "NUMBER",   m.FaturaNo);
            SetHeaderField(inv, "DATE",         m.Tarih);
            SetHeaderField(inv, "DOC_NUMBER",   m.BelgeNo);
            SetHeaderField(inv, "DOC_DATE",     m.BelgeTarihi);
            SetHeaderField(inv, "ARP_CODE",     m.CariKod);
            if (!string.IsNullOrWhiteSpace(m.CariGlKod))
                SetHeaderField(inv, "GL_CODE",  m.CariGlKod);
            SetHeaderField(inv, "FACTORY",      m.Fabrika);
            SetHeaderField(inv, "SOURCE_WH",    m.Ambar);
            SetHeaderField(inv, "PAYMENT_CODE", m.OdemeKodu);
            SetHeaderField(inv, "TRADING_GRP",  m.OzelKod);
            SetHeaderField(inv, "POST_FLAGS",   247);

            // V18: Döviz cinsi hesapları
            //
            // Logo Tiger 3 fatura tutar alanları:
            //   TOTAL_DISCOUNTED, TOTAL_GROSS, TOTAL_VAT, TOTAL_NET   = TL (ana para birimi)
            //   TC_NET                                                  = işlem dövizi (TC = "İşlem Dövizi Cinsinden Net")
            //   RC_NET, REPORTNET                                       = raporlama dövizi (USD)
            //
            // m.ToplamNet UBL'den gelir → fatura'nın asıl döviz cinsinden (EUR fatura için EUR).
            // TL fatura için DovizKuru=1, Doviz="TL" → islemNet == tlNet, hepsi aynı değer.
            // EUR fatura için DovizKuru=52.79 → tlNet = islemNet * kur.
            bool dovizli = !string.IsNullOrWhiteSpace(m.Doviz)
                           && !string.Equals(m.Doviz, "TL", StringComparison.OrdinalIgnoreCase)
                           && !string.Equals(m.Doviz, "TRY", StringComparison.OrdinalIgnoreCase)
                           && m.DovizKuru > 0;
            decimal islemKuru = dovizli ? m.DovizKuru : 1m;
            decimal IslemToTl(decimal v) => v * islemKuru;
            decimal IslemToRc(decimal v) => m.RaporlamaKuru > 0 ? Math.Round(IslemToTl(v) / m.RaporlamaKuru, 2) : 0;

            SetHeaderField(inv, "TOTAL_DISCOUNTED", (double)IslemToTl(m.ToplamMatrah));
            SetHeaderField(inv, "TOTAL_GROSS",      (double)IslemToTl(m.ToplamMatrah));
            SetHeaderField(inv, "TOTAL_VAT",        (double)IslemToTl(m.ToplamKdv));
            SetHeaderField(inv, "TOTAL_NET",        (double)IslemToTl(m.ToplamNet));
            SetHeaderField(inv, "TC_NET",           (double)m.ToplamNet);   // işlem dövizi cinsinden
            // TOTAL_SERVICES sadece Alınan Hizmet (Type=4) için. Satın Alma XML export'larında
            // bu alan yok — set edilirse Logo "fiş tipi modüle uygun değil" verebilir.
            if (m.Type == 4)
                SetHeaderField(inv, "TOTAL_SERVICES", (double)IslemToTl(m.ToplamMatrah));

            // EDTCURR_GLOBAL_CODE = "Düzenleme Dövizi" — XML referanslarına göre:
            //   - TL fatura  → "USD" (kullanıcı raporlama dövizinde düzenlesin)
            //   - USD fatura → "USD"
            //   - EUR fatura → "EUR"
            // Yani dövizli fatura için işlem dövizi, TL için "USD" fallback.
            string edtCurrCode = dovizli ? m.Doviz : "USD";
            SetHeaderField(inv, "EDTCURR_GLOBAL_CODE", edtCurrCode);
            // CURRSEL_TOTALS / CURRSEL_DETAILS: 1=Raporlama Dövizi (USD), 2=İşlem Dövizi
            // Dövizli fatura formda işlem dövizi gösterilsin (XML: EUR fatura'da CURRSEL_TOTALS=2)
            SetHeaderField(inv, "CURRSEL_TOTALS",  dovizli ? 2 : 1);
            SetHeaderField(inv, "CURRSEL_DETAILS", dovizli ? 2 : 1);

            // Dövizli fatura için TRCURR (İşlem Dövizi) alanları
            // XML referansı (Alınan Hizmet Faturası XML si (Euro)):
            //   <TRCURR>20</TRCURR>          → Logo döviz numarası (EUR=20, USD=1)
            //   <TRCURR_GLOBAL_CODE>EUR</…>  → 3 harfli kod
            //   <TRRATE>52.7962</TRRATE>     → işlem kuru (EUR/TL)
            //   <CURR_INVOICE>1</…>          → fatura dövizli mi flag
            //   <TC_XRATE>52.7962</…>        → TC kuru = işlem kuru
            // V21 (20.05.2026): TL fatura için TRRATE/TC_XRATE set ETME — DB karşılaştırması
            // gösterdi ki çalışan TL fatura (DNZ2026000001788) TRRATE=0 ile REPORTNET dolu,
            // bizim TRRATE=1 ile REPORTNET=0. Logo TRRATE=0 (default) görünce raporlama dövizi
            // hesaplamasını yapıyor, TRRATE=1 görünce "çevirim gereksiz" sayıp 0 bırakıyor.
            if (dovizli)
            {
                SetHeaderField(inv, "TRCURR",              m.EdtCurr);
                SetHeaderField(inv, "TRCURR_GLOBAL_CODE",  m.Doviz);
                SetHeaderField(inv, "TRRATE",              (double)m.DovizKuru);
                SetHeaderField(inv, "TC_XRATE",            (double)m.DovizKuru);
                SetHeaderField(inv, "CURR_INVOICE",        1);
            }
            // TL için: hiçbir TR* alanı set ETME, default 0 kalsın

            // Raporlama Dövizi (USD) kuru + tutar.
            // V22 (20.05.2026): satış faturası DNZ2026000001788 ile yapılan testte REPORTNET
            // set EDİLMEDEN de Logo'nun otomatik hesapladığı görülmüştü. AMA Alınan Hizmet
            // faturalarında (TRCODE=4, ör. GDL2026000007098) Logo otomatik HESAPLAMIYOR →
            // REPORTNET=0 kalıyor (Liste'de "Dövizli Tutar" boş).
            // V23 (31.05.2026): manuel girilmiş TRCODE=4 faturalarda (TIA2026000374046,
            // GIB2026000040005...) REPORTNET = NETTOTAL/REPORTRATE doğrulandı. Biz de hesaplayıp
            // set ediyoruz — TRCODE'dan bağımsız tüm faturalarda güvenli.
            if (m.RaporlamaKuru > 0)
            {
                SetHeaderField(inv, "RC_XRATE",    (double)m.RaporlamaKuru);
                SetHeaderField(inv, "REPORTRATE",  (double)m.RaporlamaKuru);
                decimal reportNet = IslemToRc(m.ToplamNet);   // (m.ToplamNet TL→) / RaporlamaKuru
                if (reportNet > 0m)
                {
                    try { SetHeaderField(inv, "REPORTNET", (double)reportNet); }
                    catch (Exception ex) { _log.LogWarning(ex, "REPORTNET header set başarısız: {V}", reportNet); }
                }
            }
            else
            {
                // m.RaporlamaKuru=0 → LogoAktarimService.UsdKuruAsync kur çekemedi (LG_EXCHANGE'da
                // bu tarih için CRTYPE=1 RATES2 kaydı yok). Logo'da "Dövizli Tutar" boş kalacak.
                // Diag log'a uyarı yaz, kullanıcı görsün:
                try
                {
                    var logFile = System.IO.Path.Combine(AppDataPaths.DataRoot, "logo_unity_diag.txt");
                    System.IO.File.AppendAllText(logFile,
                        $"---- {DateTime.Now:yyyy-MM-dd HH:mm:ss} [V19-WARN] eFaturaId={eFaturaId} RaporlamaKuru=0 → RC_XRATE set edilmedi. Tarih={m.Tarih:yyyy-MM-dd} Doviz={m.Doviz}. LG_EXCHANGE_211'de bu tarihe USD kuru var mı kontrol et. ----\n");
                }
                catch { }
            }
            // ACCOUNTED_CNT (DB: ACCOUNTEDCNT) — DB'deki çalışan istisnalı faturalarda 1.
            // XML export'unda bazı kayıtlar 2 gösteriyor (entegrasyon sayacı sonradan artıyor) ama
            // ilk Post için 1 olmalı.
            SetHeaderField(inv, "ACCOUNTED_CNT",  1);
            // GRPCODE (DB: GRPCODE) — Empirik veri (LG_211_01_INVOICE sorgusu): Satın Alma istisnalı=1.
            // Default 0 ile (1159) "Fiş tipi modüle uygun değil" hatası geliyor.
            SetHeaderField(inv, "GRPCODE",        1);
            SetHeaderField(inv, "AFFECT_RISK",    1);

            SetHeaderField(inv, "PROFILE_ID",    m.ProfileId);
            SetHeaderField(inv, "ESTATUS",       12);
            SetHeaderField(inv, "EBOOK_DOCTYPE", 99);
            SetHeaderField(inv, "EXIMVAT",       0);
            SetHeaderField(inv, "EDURATION_TYPE",0);

            // V16: TIME/DOC_TIME/SHIP_TIME (Logo encoded: HH*16777216 + MM*65536 + SS*256)
            //
            // Auto-oluşan DISPATCH'in saat alanları (Düzenleme Zamanı=TIME, Sevk Zamanı=SHIP_TIME)
            // 0 olunca Logo edit-validation "e-İrsaliye için tüm bilgiler eksiksiz girilmelidir"
            // diyor ve sürücü/plaka istiyor (manuel XML'de SHIP_TIME=253769519 = 15:32:55).
            // V15'te EINVOICE_TYPE cascade'ini suçlamıştık ama V15 diag'da skip edilmesine rağmen
            // popup gelmeye devam etti — kök neden TIME = 0 imiş.
            //
            // Logo TIME encoding decode (manuel fatura veri ile doğrulandı):
            //   253769472 = 15*16777216 + 32*65536 + 55*256 = 15:32:55 (low byte=0)
            //   253769519 = aynı saat + 47 (low byte=47, muhtemelen sub-second / yarı-tick)
            //
            // Logo'nun "Sevk Zamanı > Düzenleme Zamanı" kuralı için SHIP_TIME 1 dakika ileride.
            // Header set'i Logo Post sırasında DISPATCH'e cascade ediyor (manuel form da öyle yapıyor).
            int LogoTimeEncode(int h, int mi, int s) => h * 16777216 + mi * 65536 + s * 256;
            int headerTime = LogoTimeEncode(12,  0, 0);    // 12:00:00 = 201326592
            int shipTime   = LogoTimeEncode(12,  1, 0);    // 12:01:00 = 201392128 (>TIME)
            try { SetHeaderField(inv, "TIME",      headerTime); } catch (Exception ex) { _log.LogWarning(ex, "TIME header set başarısız"); }
            try { SetHeaderField(inv, "DOC_TIME",  headerTime); } catch (Exception ex) { _log.LogWarning(ex, "DOC_TIME header set başarısız"); }
            try { SetHeaderField(inv, "SHIP_TIME", shipTime);   } catch (Exception ex) { _log.LogWarning(ex, "SHIP_TIME header set başarısız"); }

            // V15: HEADER'DA EINVOICE_TYPE SET EDİLMİYOR.
            //
            // Geçmiş: V9'da "header'da bu set edilmezse 1159 atıyor" diye not düşmüştük. O günden
            // bu yana eklediğimiz alanlar (GRPCODE=1, ACCOUNTED_CNT=1, EBOOK_DOCTYPE=99, ESTATUS=12,
            // EXIMVAT=0, EDURATION_TYPE=0, EINVOICE=1) 1159'u zaten kapatıyor.
            //
            // Mevcut sorun: Logo Unity COM Post() header'daki EINVOICE_TYPE'ı auto-oluşan DISPATCH'e
            // de cascade ediyor (STFICHE.EINVOICETYP=2 yazılıyor). Bu yüzden edit-validation
            // "e-İrsaliye için tüm bilgiler eksiksiz girilmelidir" diyor ve sürücü/plaka istiyor.
            // Manuel formda STFICHE.EINVOICETYP=0 kalıyor (cascade yok) ve validation tetiklenmiyor.
            // V13/V14: DISPATCH'e pre-Post erişip override etmeyi denedik — Logo DISPATCH'i sadece
            // Post() içinde yaratıyor (V14 diag: count=0), pre-Post override imkansız.
            //
            // V16: HEADER'A EINVOICE_TYPE SET EDİYORUZ.
            // V15'te skip ediliyordu — DISPATCH cascade ile e-İrsaliye validation tetikleniyordu.
            // Şimdi DISPATCH override (Validation sonrası, Post öncesi) zaten satır 670 civarında
            // EINVOICE_TYPE=0'a çekiyor → header'da kalan değer Logo INVOICE.EINVOICETYP'i doldurur,
            // STFICHE.EINVOICETYP=0 kalır → validation tetiklenmez.
            // İstisna=2, Tevkifat=4 (Logo standart).
            int eInvType = 0;
            if (m.Satirlar.Any(s => !string.IsNullOrWhiteSpace(s.IstisnaKodu)))
                eInvType = 2;
            else if (m.Tevkifatli || m.Satirlar.Any(s => s.Tevkifatli))
                eInvType = 4;
            string eInvFieldUsed = "";
            if (eInvType > 0)
            {
                try { SetHeaderField(inv, "EINVOICE_TYPE", eInvType); eInvFieldUsed = "EINVOICE_TYPE"; }
                catch (Exception ex) { _log.LogWarning(ex, "EINVOICE_TYPE header set başarısız: {V}", eInvType); }
            }

            // EINVOICE — DB kolonu smallint (0/1 flag), DateTime DEĞİL.
            // Logo XML export'unda <EINVOICE>1</EINVOICE> şeklinde (boolean flag).
            // DateTime gönderirsek COM sessizce reddediyor → EINVOICE=0 kalıyor
            // → header EINVOICETYP=2 (İstisna) ile çelişiyor → DB (1159).
            if (m.EFaturaTarihi.HasValue)
                SetHeaderField(inv, "EINVOICE", 1);

            // ---- SATIRLAR — TRANSACTIONS field'ına GetFieldIndex+Item ile ulaş (FieldByName null dönüyordu)
            object? txField = null;
            object? dfForTx;
            try
            {
                dfForTx = GetProp(inv, "DataFields")
                          ?? throw new InvalidOperationException("inv.DataFields null (TRANSACTIONS). " + SetFieldDiag());

                // 1) GetFieldIndex("TRANSACTIONS") + Item(idx)
                try
                {
                    object? idxObj = IDispCall(dfForTx, "GetFieldIndex",
                        (ushort)(DISPATCH_METHOD | DISPATCH_PROPERTYGET),
                        new object?[] { "TRANSACTIONS" });
                    if (idxObj != null)
                    {
                        int idx = Convert.ToInt32(idxObj);
                        if (idx >= 0)
                            txField = IDispCall(dfForTx, "Item",
                                (ushort)(DISPATCH_METHOD | DISPATCH_PROPERTYGET),
                                new object?[] { idx });
                    }
                }
                catch { /* try fallback */ }

                // 2) Fallback: FieldByName
                if (txField == null)
                    txField = Inv(dfForTx, "FieldByName", "TRANSACTIONS");
            }
            catch (Exception tex)
            {
                throw new InvalidOperationException(
                    $"FieldByName(TRANSACTIONS) patladı. {SetFieldDiag()} | {tex.GetType().Name}: {tex.Message}", tex);
            }
            if (txField == null)
            {
                string critical = QuickCriticalProbe(inv, df);
                // Gerçek alan isimlerini enumerate et — tam liste log dosyasına, kısa özet UI'ya
                int dfCount = 0;
                try
                {
                    object? c = IDispCall(df, "Count", DISPATCH_PROPERTYGET, Array.Empty<object?>());
                    if (c != null) dfCount = Convert.ToInt32(c);
                }
                catch { }
                string enumSummary = "";
                try { enumSummary = EnumerateFieldNamesAndLog(df, dfCount, eFaturaId.ToString()); }
                catch (Exception eex) { enumSummary = "[ENUM-EX:" + eex.GetType().Name + "]"; }

                string fullDiag = $"[BUILD-V19 doType={DO_PURCH_INVOICE}] {critical} | New.ret={newRet} New.err={(string.IsNullOrEmpty(newErr) ? "OK" : newErr)} AppErr={newAppErr} | {SetFieldDiag()} | {enumSummary} | TANI: " + dispatchDiag;
                // Paylaşılan AppData yoluna yaz — publish her seferinde silinmez, uzaktan da erişilebilir
                try
                {
                    var logFile = System.IO.Path.Combine(AppDataPaths.DataRoot, "logo_unity_diag.txt");
                    System.IO.File.AppendAllText(logFile,
                        $"---- {DateTime.Now:yyyy-MM-dd HH:mm:ss} eFaturaId={eFaturaId} ----\n{fullDiag}\n\n");
                }
                catch { /* logging best-effort */ }

                // UI hatası: tam diag inline — uzak makinede log dosyasına ihtiyaç olmasın
                throw new InvalidOperationException(fullDiag);
            }

            object txLines = GetProp(txField, "Lines")
                             ?? throw new InvalidOperationException("TRANSACTIONS.Lines null. " + SetFieldDiag());

            foreach (var s in m.Satirlar)
            {
                // VBA: transactions_lines.AppendLine (NOT AppendNew!)
                Inv(txLines, "AppendLine");
                int idx = Convert.ToInt32(GetProp(txLines, "Count") ?? 1) - 1;
                object line = GetLine(txLines, idx)
                              ?? throw new InvalidOperationException($"TRANSACTIONS.Lines[{idx}] null.");

                SetField(line, "TYPE",       s.SatirTipi);
                // TRCODE / IOCODE → header'dan türetilir; satırda set edilmez.
                // Üç farklı çalışan Logo XML export'unda (Satın Alma TL, Satın Alma İstisnalı,
                // Hizmet Tevkifatlı) TRANSACTION elemanında TRCODE/IOCODE yok. Set edersek
                // Logo internal validasyonu reddedebiliyor.
                SetField(line, "MASTER_CODE",s.MasterCode);
                SetField(line, "SOURCEINDEX",m.Ambar);
                SetField(line, "QUANTITY",   (double)s.Miktar);
                // V18: Dövizli fatura için PRICE/TOTAL = TL eşdeğer (Logo ana para birimi),
                //      EDT_PRICE / EXCHLINE_PRICE = işlem dövizi (EUR/USD).
                //      TL fatura için islemKuru=1 → değişmez.
                SetField(line, "PRICE",      (double)IslemToTl(s.BirimFiyat));
                SetField(line, "TOTAL",      (double)IslemToTl(s.ToplamNet));
                SetField(line, "TOTAL_NET",  (double)IslemToTl(s.ToplamNet));
                SetField(line, "VAT_BASE",   (double)IslemToTl(s.ToplamNet));
                if (dovizli)
                {
                    SetField(line, "EDT_PRICE",      (double)s.BirimFiyat);
                    SetField(line, "EXCHLINE_PRICE", (double)s.BirimFiyat);
                }
                SetField(line, "UNIT_CODE",  s.Birim);
                SetField(line, "UNIT_CONV1", 1);
                SetField(line, "UNIT_CONV2", 1);
                SetField(line, "VAT_RATE",   (double)s.KdvOran);
                SetField(line, "VAT_AMOUNT", (double)IslemToTl(s.KdvTutar));

                // İstisnalı kalem (KDV=0): Logo VATEXCEPT_CODE + VATEXCEPT_REASON dolu bekliyor.
                // Logo XML export alan adları alt çizgili — VATEXCEPTCODE (alt çizgisiz) yazınca
                // "Bu isimde bir alan bulunamadı" + (1159) hatası gelir.
                if (!string.IsNullOrWhiteSpace(s.IstisnaKodu))
                    SetField(line, "VATEXCEPT_CODE", s.IstisnaKodu);
                if (!string.IsNullOrWhiteSpace(s.IstisnaSebep))
                    SetField(line, "VATEXCEPT_REASON", s.IstisnaSebep);
                SetField(line, "BILLED",     1);
                SetField(line, "FACTORY",    m.Fabrika);
                SetField(line, "AFFECT_RISK",1);
                SetField(line, "EDT_CURR",   m.EdtCurr);

                if (!string.IsNullOrWhiteSpace(s.MasrafMerkezi))
                    SetField(line, "OHP_CODE1",  s.MasrafMerkezi);

                // GL_CODE1 = mal/stok/gider hesabı (150/153/157/740/760.xx)
                // GL_CODE2 = İndirilecek KDV (191.xx)
                // Logo XML export'unda (Satın Alma TL, Satın Alma İstisnalı, Hizmet Tevkifatlı) bu sıralama
                // sabit. Tersine çevirirsek Logo "(1159) Fiş tip bilgisi modüle uygun değil" verir
                // çünkü Satın Alma fişinde GL_CODE1=191 (KDV) kombinasyonu modül kurallarına aykırı.
                if (!string.IsNullOrWhiteSpace(s.HizmetMalHesabi))
                    SetField(line, "GL_CODE1", s.HizmetMalHesabi);
                if (!string.IsNullOrWhiteSpace(s.KdvHesabi))
                    SetField(line, "GL_CODE2", s.KdvHesabi);

                if (s.Tevkifatli)
                {
                    SetField(line, "CANDEDUCT",          1);
                    SetField(line, "ADD_TAX_EFFECT_KDV", 1);   // V19: XML referansı (Tevkifatlı) bunu set ediyor
                    SetField(line, "DEDUCTION_PART1",    s.TevkifatPay);
                    SetField(line, "DEDUCTION_PART2",    s.TevkifatPayda);
                    SetField(line, "DEDUCTION_TOT",      (double)s.TevkifatTutar);
                    if (!string.IsNullOrWhiteSpace(s.TevkifatKodu))
                        SetField(line, "DEDUCT_CODE", s.TevkifatKodu);

                    if (!string.IsNullOrWhiteSpace(s.TevkifatKdvHesabi))
                    {
                        SetField(line, "GL_CODE3", s.TevkifatKdvHesabi);
                        if (!string.IsNullOrWhiteSpace(s.MasrafMerkezi))
                            SetField(line, "OHP_CODE3", s.MasrafMerkezi);
                    }
                    if (!string.IsNullOrWhiteSpace(s.SorumluKdvHesabi))
                    {
                        SetField(line, "GL_CODE4", s.SorumluKdvHesabi);
                        if (!string.IsNullOrWhiteSpace(s.MasrafMerkezi))
                            SetField(line, "OHP_CODE4", s.MasrafMerkezi);
                    }

                    // PREACCLINES — sadece hizmet faturalarında, tevkifatlı kalemler için
                    if (m.Type == 4 && !string.IsNullOrWhiteSpace(s.MasrafMerkezi))
                    {
                        try
                        {
                            object? preField = Inv(line, "FieldByName", "PREACCLINES");
                            if (preField != null)
                            {
                                object? preLines = GetProp(preField, "Lines");
                                if (preLines != null)
                                {
                                    Inv(preLines, "AppendLine");
                                    int pidx = Convert.ToInt32(GetProp(preLines, "Count") ?? 1) - 1;
                                    object? pl = GetLine(preLines, pidx);
                                    if (pl != null)
                                    {
                                        SetField(pl, "LINENR",       1);
                                        SetField(pl, "DISTRATE",     100);
                                        SetField(pl, "DATE",         m.Tarih);
                                        SetField(pl, "PREVLINETYPE", 1);
                                        SetField(pl, "MODULNR",      1);
                                        SetField(pl, "CENTERCODE",   s.MasrafMerkezi);
                                    }
                                }
                            }
                        }
                        catch (Exception ex)
                        {
                            _log.LogWarning(ex, "PREACCLINES eklenemedi (göz ardı ediliyor)");
                        }
                    }
                }

                SetField(line, "UNIT_GLOBAL_CODE",      "ADET");
                // V18: EDTCURR = header ile aynı kural — dövizli ise işlem dövizi, TL fatura için USD
                SetField(line, "EDTCURR_GLOBAL_CODE",   edtCurrCode);
                // İşlem dövizi (line seviyesinde de) — dövizli fatura için zorunlu
                if (dovizli)
                {
                    SetField(line, "TRCURR",              m.EdtCurr);
                    SetField(line, "TRCURR_GLOBAL_CODE",  m.Doviz);
                    SetField(line, "TRRATE",              (double)m.DovizKuru);
                }
                // V21: TL satırı için TRRATE/RC_AMOUNT set ETME. Logo TRRATE=0 + REPORTRATE
                // ile otomatik hesaplıyor (DB karşılaştırması: DNZ2026000001788 = TRRATE=0 → REPORTNET dolu)
                if (m.RaporlamaKuru > 0)
                    SetField(line, "RC_XRATE", (double)m.RaporlamaKuru);
                SetField(line, "FOREIGN_TRADE_TYPE",    0);
                SetField(line, "DISTRIBUTION_TYPE_WHS", 0);
                SetField(line, "DISTRIBUTION_TYPE_FNO", 0);
                SetField(line, "FUTURE_MONTH_BEGDATE",  m.Tarih);
                SetField(line, "MONTH",                 m.Tarih.Month);
                SetField(line, "YEAR",                  m.Tarih.Year);
            }

            // ---- PAYMENT_LIST ---- VBA: purcinvoices.DataFields.FieldByName("PAYMENT_LIST").Lines (fresh chain)
            try
            {
                object? dfForPay = GetProp(inv, "DataFields");
                object? payField = dfForPay == null ? null : Inv(dfForPay, "FieldByName", "PAYMENT_LIST");
                if (payField != null)
                {
                    object? payLines = GetProp(payField, "Lines");
                    if (payLines != null)
                    {
                        Inv(payLines, "AppendLine");
                        int pyIdx = Convert.ToInt32(GetProp(payLines, "Count") ?? 1) - 1;
                        object? p = GetLine(payLines, pyIdx);
                        if (p != null)
                        {
                            SetField(p, "DATE",             m.Tarih);
                            SetField(p, "PROCDATE",         m.Tarih);
                            SetField(p, "DISCOUNT_DUEDATE", m.Tarih);
                            SetField(p, "MODULENR",         4);
                            SetField(p, "SIGN",             1);
                            SetField(p, "TRCODE",           m.Type);
                            // V18: PAYMENT.TOTAL = TL eşdeğer (header TOTAL_NET ile tutarlı)
                            SetField(p, "TOTAL",            (double)IslemToTl(m.ToplamNet));
                            SetField(p, "PAY_NO",           1);
                            SetField(p, "DISCTRDELLIST",    0);
                        }
                    }
                }
            }
            catch (Exception ex)
            {
                _log.LogWarning(ex, "PAYMENT_LIST eklenemedi (göz ardı ediliyor)");
            }

            // ---- VALIDATION (Post öncesi spesifik hataları yakala) ----
            string preValidateDiag = "";
            try
            {
                object? vRes = Inv(inv, "Validation");
                var preList = new List<string>();
                CollectValidateErrors(inv, preList);
                if (preList.Count > 0)
                    preValidateDiag = " | PreValidate: " + string.Join(" / ", preList);
                else if (vRes != null)
                    preValidateDiag = $" | Validation()={vRes}";
            }
            catch (Exception vex)
            {
                preValidateDiag = $" | Validation() patladı: {vex.Message}";
            }

            // ---- DISPATCHES EINVOICE_TYPE override (Validation sonrası, Post öncesi) ----
            // Logo Post() satın alma faturasında otomatik DISPATCH oluşturuyor ve header'daki
            // EINVOICE_TYPE'ı (istisna=2, tevkifat=4) DISPATCH'e de kopyalıyor. Bu yüzden Logo
            // edit-validasyonu DISPATCH'i "e-İrsaliye istisna kayıtlı" sayıp driver/plaka bilgisi
            // istiyor (manuel girilen faturalarda STFICHE.EINVOICETYP=0 olduğu için sorun olmuyor).
            // V13 (PAYMENT_LIST sonrası, Validation öncesi) işe yaramadı — collection o aşamada boş.
            // V14: Validation() çağrısı DISPATCH'i pre-create ediyor olabilir; sonrasında dene.
            // Diag log'a kaç satır bulunduğu yazılıyor — kesin teşhis için.
            int dispOverrideCount = 0;
            try
            {
                object? dfDisp = GetProp(inv, "DataFields");
                object? dispField = dfDisp == null ? null : Inv(dfDisp, "FieldByName", "DISPATCHES");
                if (dispField != null)
                {
                    object? dispLines = GetProp(dispField, "Lines");
                    if (dispLines != null)
                    {
                        int dispCount = Convert.ToInt32(GetProp(dispLines, "Count") ?? 0);
                        for (int di = 0; di < dispCount; di++)
                        {
                            object? dl = GetLine(dispLines, di);
                            if (dl != null)
                            {
                                SetField(dl, "EINVOICE_TYPE", 0);
                                dispOverrideCount++;
                            }
                        }
                    }
                }
            }
            catch (Exception ex)
            {
                _log.LogWarning(ex, "DISPATCH EINVOICE_TYPE override başarısız (göz ardı)");
            }
            try
            {
                var logFile = System.IO.Path.Combine(AppDataPaths.DataRoot, "logo_unity_diag.txt");
                decimal reportNetCalc = m.RaporlamaKuru > 0 ? Math.Round(m.ToplamNet / m.RaporlamaKuru, 2) : 0m;
                System.IO.File.AppendAllText(logFile,
                    $"---- {DateTime.Now:yyyy-MM-dd HH:mm:ss} [V23] eFaturaId={eFaturaId} Doviz={m.Doviz} TRCURR={m.EdtCurr} TRRATE={m.DovizKuru:N4} dövizli={dovizli} islemKuru={islemKuru:N4} RcKur={m.RaporlamaKuru:N4} EDTCURR={edtCurrCode} ToplamNet={m.ToplamNet:N2} REPORTNET={reportNetCalc:N2} (header'a SET edildi) EINV_TYPE={eInvType} ({eInvFieldUsed}) ----\n");
            }
            catch { }

            // ---- POST ----
            bool posted;
            try
            {
                posted = (bool)(Inv(inv, "Post") ?? false);
            }
            catch (Exception ex)
            {
                satir.Hata = $"Post çağrısı patladı: {ex.Message}{preValidateDiag}";
                return satir;
            }

            if (!posted)
            {
                string einvDiag = eInvType > 0
                    ? $" | EINVOICE_TYPE={eInvType} field={(string.IsNullOrEmpty(eInvFieldUsed) ? "BULUNAMADI" : eInvFieldUsed)}"
                    : "";
                CollectValidateErrors(inv, satir.ValidateErrors);
                string postValidateDiag = satir.ValidateErrors.Count > 0
                    ? " | PostValidate: " + string.Join(" / ", satir.ValidateErrors)
                    : "";
                satir.Hata = $"Post başarısız. ErrorCode={GetStr(inv, "ErrorCode")} ErrorDesc={GetStr(inv, "ErrorDesc")} DB={GetStr(inv, "DBErrorDesc")}{einvDiag}{preValidateDiag}{postValidateDiag} | {SetFieldDiag()}";

                // Header field'larını enumerate edip diag log'a yaz — bir sonraki teşhis için
                try
                {
                    object? dfRawErr = GetProp(inv, "DataFields");
                    if (dfRawErr != null)
                    {
                        int cnt = 0;
                        try { cnt = Convert.ToInt32(IDispCall(dfRawErr, "Count", DISPATCH_PROPERTYGET, Array.Empty<object?>()) ?? 0); }
                        catch { }
                        EnumerateFieldNamesAndLog(dfRawErr, cnt, $"PostFail-{eFaturaId}");
                    }
                }
                catch { /* best-effort */ }

                return satir;
            }

            satir.Basarili = true;
            satir.LogoNo   = GetFieldStr(inv, "NUMBER");
            satir.LogoRef  = GetFieldLong(inv, "INTERNAL_REFERENCE");

            // REPORTNET ikinci pass — Logo Post sırasında REPORTNET set'ini sessizce 0'lıyor
            // (V18 deneyimi). Manuel formda kullanıcı Kaydet'e bastığında dolar. Aynı obje
            // üzerinde set + Update/Save deneyelim, hangisi tutarsa o.
            if (m.RaporlamaKuru > 0)
            {
                decimal reportNet = Math.Round(m.ToplamNet / m.RaporlamaKuru, 2);
                string setMsg = "", updMsg = "";

                try { SetHeaderField(inv, "REPORTNET", (double)reportNet); setMsg = "OK"; }
                catch (Exception e1) { setMsg = "EX:" + e1.Message; }

                // Önce Update dene, başarısızsa Save, başarısızsa Post(2nd) dene.
                foreach (var method in new[] { "Update", "Save", "Post" })
                {
                    try
                    {
                        var r = Inv(inv, method);
                        bool ok = r is bool bb && bb;
                        updMsg += $" {method}={(ok ? "OK" : "false")}";
                        if (ok) break;
                    }
                    catch (Exception ex)
                    {
                        updMsg += $" {method}=EX:{ex.Message.Split('\n')[0]}";
                    }
                }

                try
                {
                    var logFile = System.IO.Path.Combine(AppDataPaths.DataRoot, "logo_unity_diag.txt");
                    System.IO.File.AppendAllText(logFile,
                        $"---- {DateTime.Now:yyyy-MM-dd HH:mm:ss} [V25-POST-UPD] eFaturaId={eFaturaId} REPORTNET={reportNet:N2} SetField={setMsg} Methods={updMsg.Trim()} LogoRef={satir.LogoRef} ----\n");
                } catch { }
            }

            return satir;
        }
        catch (Exception ex)
        {
            satir.Hata = $"Beklenmeyen hata: {ex.Message}";
            _log.LogError(ex, "PostSingleInvoice patladı (e-Fatura {Id})", eFaturaId);
            return satir;
        }
        finally
        {
            Release(inv);
        }
    }

    // ---------- COM helpers (Type.InvokeMember tabanlı, dynamic değil) ----------

    private static object CreateUnityApp()
    {
        var t = Type.GetTypeFromProgID("UnityObjects.UnityApplication", throwOnError: true)
                ?? throw new InvalidOperationException("UnityObjects.UnityApplication ProgID bulunamadı.");
        return Activator.CreateInstance(t)
                ?? throw new InvalidOperationException("UnityApplication oluşturulamadı.");
    }

    // .NET 8 Marshal.GetObjectForNativeVariant'ın VT_DISPATCH'i null'a çevirme bug'ını bypass etmek
    // için tüm COM dispatch çağrıları doğrudan IDispatch::Invoke + manuel VARIANT extraction ile yapılıyor.
    private const ushort DISPATCH_METHOD      = 1;
    private const ushort DISPATCH_PROPERTYGET = 2;
    private const ushort DISPATCH_PROPERTYPUT = 4;
    private const int    DISPID_PROPERTYPUT   = -3;

    private static object? Inv(object target, string method, params object?[] args)
        => IDispCall(target, method, (ushort)(DISPATCH_METHOD | DISPATCH_PROPERTYGET), args);

    // Asıl dispatch çağrısı.
    private static object? IDispCall(object target, string member, ushort wFlags, object?[] args)
    {
        var disp = (IDispatch)target;
        var iidNull = Guid.Empty;
        var names = new[] { member };
        var ids = new int[1];
        int hr0 = disp.GetIDsOfNames(ref iidNull, names, 1, 0, ids);
        if (hr0 < 0)
            throw Marshal.GetExceptionForHR(hr0) ?? new COMException($"GetIDsOfNames({member}) hr=0x{hr0:X8}", hr0);
        return IDispCallById(disp, ids[0], member, wFlags, args);
    }

    private static object? IDispCallById(IDispatch disp, int dispid, string memberName, ushort wFlags, object?[] args)
    {
        var iidNull = Guid.Empty;
        int cArgs = args?.Length ?? 0;
        int vsize = Marshal.SizeOf<Variant>();
        IntPtr rgvarg = IntPtr.Zero;
        IntPtr rgdispidNamed = IntPtr.Zero;
        int cNamedArgs = 0;

        if (cArgs > 0)
        {
            rgvarg = Marshal.AllocCoTaskMem(vsize * cArgs);
            for (int i = 0; i < cArgs; i++)
            {
                var v = new Variant();
                v.SetValue(args![cArgs - 1 - i]); // DISPPARAMS rgvarg LIFO
                Marshal.StructureToPtr(v, rgvarg + i * vsize, false);
            }
        }
        if (wFlags == DISPATCH_PROPERTYPUT)
        {
            rgdispidNamed = Marshal.AllocCoTaskMem(sizeof(int));
            Marshal.WriteInt32(rgdispidNamed, DISPID_PROPERTYPUT);
            cNamedArgs = 1;
        }

        try
        {
            var dp = new DISPPARAMS
            {
                rgvarg = rgvarg,
                rgdispidNamedArgs = rgdispidNamed,
                cArgs = cArgs,
                cNamedArgs = cNamedArgs
            };

            IntPtr pVarResult = Marshal.AllocCoTaskMem(vsize);
            try
            {
                Marshal.StructureToPtr(new Variant(), pVarResult, false);
                int hr = disp.Invoke(dispid, ref iidNull, 0, wFlags, ref dp, pVarResult, IntPtr.Zero, IntPtr.Zero);
                if (hr < 0)
                    throw Marshal.GetExceptionForHR(hr) ?? new COMException($"Invoke({memberName}) hr=0x{hr:X8}", hr);
                return ExtractVariantValue(pVarResult);
            }
            finally
            {
                try { NativeMethods.VariantClear(pVarResult); } catch { }
                Marshal.FreeCoTaskMem(pVarResult);
            }
        }
        finally
        {
            if (rgvarg != IntPtr.Zero)
            {
                for (int i = 0; i < cArgs; i++)
                {
                    try { NativeMethods.VariantClear(rgvarg + i * vsize); } catch { }
                }
                Marshal.FreeCoTaskMem(rgvarg);
            }
            if (rgdispidNamed != IntPtr.Zero) Marshal.FreeCoTaskMem(rgdispidNamed);
        }
    }

    // VARIANT'tan değer çıkarma. KRITIK: .NET 8'de Marshal.GetObjectForNativeVariant
    // VT_DISPATCH için null döndürüyor — bu yüzden VT_DISPATCH/VT_UNKNOWN'u manuel okuyup
    // pdispVal/punkVal IntPtr'sını GetObjectForIUnknown ile sarmalıyoruz.
    private static object? ExtractVariantValue(IntPtr pVar)
    {
        ushort vt = (ushort)Marshal.ReadInt16(pVar, 0);
        switch (vt)
        {
            case 0:  // VT_EMPTY
            case 1:  // VT_NULL
                return null;
            case 9:  // VT_DISPATCH
            case 13: // VT_UNKNOWN
                IntPtr p = Marshal.ReadIntPtr(pVar, 8);
                return p == IntPtr.Zero ? null : Marshal.GetObjectForIUnknown(p);
            case 8:  // VT_BSTR
                IntPtr pstr = Marshal.ReadIntPtr(pVar, 8);
                return pstr == IntPtr.Zero ? null : Marshal.PtrToStringBSTR(pstr);
            default:
                try { return Marshal.GetObjectForNativeVariant(pVar); }
                catch { return null; }
        }
    }

    // Lines koleksiyonundan idx'inci satırı al. VBA: lines((idx)) default member çağrısı.
    // Önce isimle "Item" dener; başarısız olursa DISPID_VALUE (0) ile default member çağırır.
    private static object? GetLine(object lines, int idx)
    {
        try
        {
            return IDispCall(lines, "Item",
                (ushort)(DISPATCH_METHOD | DISPATCH_PROPERTYGET),
                new object?[] { idx });
        }
        catch { /* try default member */ }

        // DISPID_VALUE = 0 = default member
        try
        {
            return IDispCallById((IDispatch)lines, 0, "(default)",
                (ushort)(DISPATCH_METHOD | DISPATCH_PROPERTYGET),
                new object?[] { idx });
        }
        catch { return null; }
    }

    // COM dispatch: property get (parametresiz)
    private static object? GetProp(object target, string prop)
        => IDispCall(target, prop, DISPATCH_PROPERTYGET, Array.Empty<object?>());

    // COM dispatch: property set
    private static void SetProp(object target, string prop, object value)
        => IDispCall(target, prop, DISPATCH_PROPERTYPUT, new object?[] { value });

    // VBA: purcinvoices.DataFields.FieldByName("X").Value = value
    // Diagnostic: ilk hatayı yutmuyoruz, FailReason'a not düşüyoruz.
    [ThreadStatic] private static int _setFieldOk;
    [ThreadStatic] private static string? _setFieldFirstError;
    [ThreadStatic] private static string? _setFieldFirstErrorName;

    // Header field — Logo IDataFields'da FieldByName null dönüyor (geç bağlama bug'ı?).
    // Bunun yerine SetFieldValue#4 metodunu kullanıyoruz: tek atışta isimle değer setlenir.
    // SetFieldValue başarısız olursa GetFieldIndex+Item, en son fallback FieldByName.
    private static void SetHeaderField(object inv, string name, object? value)
    {
        if (value == null) return;
        object? df = null;
        try
        {
            df = IDispCall(inv, "DataFields", DISPATCH_PROPERTYGET, Array.Empty<object?>());
            if (df == null)
            {
                if (_setFieldFirstError == null) { _setFieldFirstError = "inv.DataFields null"; _setFieldFirstErrorName = name; }
                return;
            }

            // 1) SetFieldValue(name, value) — tek atışta, FieldByName proxy'ye gerek yok
            try
            {
                IDispCall(df, "SetFieldValue",
                    (ushort)(DISPATCH_METHOD | DISPATCH_PROPERTYGET),
                    new object?[] { name, value });
                _setFieldOk++;
                return;
            }
            catch (Exception sfvEx)
            {
                if (_setFieldFirstError == null)
                {
                    _setFieldFirstError = $"SetFieldValue({name}): {sfvEx.GetType().Name}: {Truncate(sfvEx.Message, 100)}";
                    _setFieldFirstErrorName = name;
                }
            }

            // 2) Fallback: GetFieldIndex(name) → Item(idx).Value = value
            try
            {
                object? idxObj = IDispCall(df, "GetFieldIndex",
                    (ushort)(DISPATCH_METHOD | DISPATCH_PROPERTYGET),
                    new object?[] { name });
                if (idxObj != null)
                {
                    int idx = Convert.ToInt32(idxObj);
                    if (idx >= 0)
                    {
                        object? field = IDispCall(df, "Item",
                            (ushort)(DISPATCH_METHOD | DISPATCH_PROPERTYGET),
                            new object?[] { idx });
                        if (field != null)
                        {
                            IDispCall(field, "Value", DISPATCH_PROPERTYPUT, new object?[] { value });
                            _setFieldOk++;
                            return;
                        }
                    }
                }
            }
            catch { /* try next */ }

            // 3) Son çare: FieldByName(name).Value = value (Logo'da bu yol null dönüyordu)
            try
            {
                object? field = IDispCall(df, "FieldByName",
                    (ushort)(DISPATCH_METHOD | DISPATCH_PROPERTYGET),
                    new object?[] { name });
                if (field != null)
                {
                    IDispCall(field, "Value", DISPATCH_PROPERTYPUT, new object?[] { value });
                    _setFieldOk++;
                }
            }
            catch { /* swallow — diag already set */ }
        }
        catch (Exception ex)
        {
            if (_setFieldFirstError == null)
            {
                _setFieldFirstError = $"{ex.GetType().Name}: {ex.Message}";
                _setFieldFirstErrorName = name;
            }
        }
    }

    // Line/Payment field — target line/payment Line objesi. VBA: line.FieldByName(name).Value = value
    // Line objesi de muhtemelen IDataFields gibi davranır; SetFieldValue → GetFieldIndex → FieldByName fallback.
    private static void SetField(object fields, string name, object? value)
    {
        if (value == null) return;

        // 1) SetFieldValue(name, value) — tek atış
        try
        {
            IDispCall(fields, "SetFieldValue",
                (ushort)(DISPATCH_METHOD | DISPATCH_PROPERTYGET),
                new object?[] { name, value });
            _setFieldOk++;
            return;
        }
        catch { /* try next */ }

        // 2) GetFieldIndex(name) → Item(idx).Value = value
        try
        {
            object? idxObj = IDispCall(fields, "GetFieldIndex",
                (ushort)(DISPATCH_METHOD | DISPATCH_PROPERTYGET),
                new object?[] { name });
            if (idxObj != null)
            {
                int idx = Convert.ToInt32(idxObj);
                if (idx >= 0)
                {
                    object? field = IDispCall(fields, "Item",
                        (ushort)(DISPATCH_METHOD | DISPATCH_PROPERTYGET),
                        new object?[] { idx });
                    if (field != null)
                    {
                        IDispCall(field, "Value", DISPATCH_PROPERTYPUT, new object?[] { value });
                        _setFieldOk++;
                        return;
                    }
                }
            }
        }
        catch { /* try fallback */ }

        // 3) Fallback: FieldByName
        try
        {
            object? field = IDispCall(fields, "FieldByName",
                (ushort)(DISPATCH_METHOD | DISPATCH_PROPERTYGET),
                new object?[] { name });
            if (field == null)
            {
                if (_setFieldFirstError == null)
                {
                    _setFieldFirstError = $"All 3 methods failed for {name}";
                    _setFieldFirstErrorName = name;
                }
                return;
            }
            IDispCall(field, "Value", DISPATCH_PROPERTYPUT, new object?[] { value });
            _setFieldOk++;
        }
        catch (Exception ex)
        {
            if (_setFieldFirstError == null)
            {
                _setFieldFirstError = $"{ex.GetType().Name}: {ex.Message}";
                _setFieldFirstErrorName = name;
            }
        }
    }

    private static void ResetSetFieldDiag()
    {
        _setFieldOk = 0;
        _setFieldFirstError = null;
        _setFieldFirstErrorName = null;
    }

    private static string SetFieldDiag()
    {
        return $"SetField OK={_setFieldOk}" +
               (_setFieldFirstError != null
                   ? $", ilk hata [{_setFieldFirstErrorName}]: {_setFieldFirstError}"
                   : "");
    }

    // Bir field için birden çok aday adı sırayla dene; ilk başarılı adı döndür, hepsi başarısızsa "" döndür.
    // Logo'nun XML field adı (örn. EINVOICE_TYPE) COM API'de farklı olabilir (örn. EINVOICETYP).
    private static string TrySetHeaderFieldAlt(object inv, string[] candidates, object? value)
    {
        if (value == null) return "";
        foreach (var name in candidates)
        {
            int before = _setFieldOk;
            SetHeaderField(inv, name, value);
            if (_setFieldOk > before)
                return name;
        }
        return "";
    }

    // DataObjectType keşif probe — UnityApplication.NewDataObject(type) için doğru sayıyı bul.
    // 1..50 arası her değer için: NewDataObject + New() + TableName + ilk 3 field adı log'a yazılır.
    // Çıktıdan TableName='LG_211_01_INVOICE' veya '...PINV' içeren satır = satın alma faturası.
    private void DiscoverDataObjectTypes(object app, string eFaturaId)
    {
        var logFile = System.IO.Path.Combine(AppDataPaths.DataRoot, "logo_unity_diag.txt");
        var sb = new System.Text.StringBuilder();
        sb.AppendLine($"---- {DateTime.Now:yyyy-MM-dd HH:mm:ss} DATA-OBJECT-TYPE-DISCOVERY eFaturaId={eFaturaId} ----");

        for (int t = 1; t <= 50; t++)
        {
            object? probeInv = null;
            try
            {
                probeInv = Inv(app, "NewDataObject", t);
                if (probeInv == null)
                {
                    sb.AppendLine($"  Type={t,-3} : NewDataObject null");
                    continue;
                }
                try { Inv(probeInv, "New"); } catch { /* New() başarısız olabilir, schema yine de yüklenmiş olabilir */ }

                string tableName = "?";
                try
                {
                    object? tn = IDispCall(probeInv, "TableName", DISPATCH_PROPERTYGET, Array.Empty<object?>());
                    tableName = tn?.ToString() ?? "<null>";
                }
                catch (Exception ex) { tableName = "EX:" + ex.GetType().Name; }

                int fcount = 0;
                string[] firstNames = { "?", "?", "?" };
                try
                {
                    object? dfp = GetProp(probeInv, "DataFields");
                    if (dfp != null)
                    {
                        try
                        {
                            object? c = IDispCall(dfp, "Count", DISPATCH_PROPERTYGET, Array.Empty<object?>());
                            if (c != null) fcount = Convert.ToInt32(c);
                        }
                        catch { }
                        for (int i = 0; i < Math.Min(3, fcount); i++)
                        {
                            try
                            {
                                object? fld = IDispCall(dfp, "Item",
                                    (ushort)(DISPATCH_METHOD | DISPATCH_PROPERTYGET),
                                    new object?[] { i });
                                if (fld != null)
                                {
                                    try
                                    {
                                        object? nm = IDispCall(fld, "FieldName", DISPATCH_PROPERTYGET, Array.Empty<object?>());
                                        firstNames[i] = nm?.ToString() ?? "<null>";
                                    }
                                    catch { firstNames[i] = "EX"; }
                                    try { Marshal.ReleaseComObject(fld); } catch { }
                                }
                            }
                            catch { /* skip */ }
                        }
                        try { Marshal.ReleaseComObject(dfp); } catch { }
                    }
                }
                catch (Exception ex) { firstNames[0] = "DF-EX:" + ex.GetType().Name; }

                sb.AppendLine($"  Type={t,-3} : TableName={tableName,-32} Count={fcount,-3} F0={firstNames[0],-25} F1={firstNames[1],-25} F2={firstNames[2]}");
            }
            catch (Exception ex)
            {
                sb.AppendLine($"  Type={t,-3} : EX:{ex.GetType().Name}:{Truncate(ex.Message, 60)}");
            }
            finally
            {
                try { if (probeInv != null) Marshal.ReleaseComObject(probeInv); } catch { }
            }
        }

        try { System.IO.File.AppendAllText(logFile, sb.ToString() + "\n"); }
        catch { /* best-effort */ }
    }

    // Koleksiyondaki gerçek alan isimlerini dök — late-binding "TYPE" görmüyor diyorsa
    // gerçek isimler nedir? Item(0)..Item(N) iterasyonu + Name/FieldName/DBFieldName okur.
    // Tam liste log dosyasına; ilk 8 isim UI özetine.
    private static string EnumerateFieldNamesAndLog(object df, int totalCount, string eFaturaId)
    {
        int limit = Math.Min(totalCount, 600);
        var logSb = new System.Text.StringBuilder();
        var uiSb = new System.Text.StringBuilder();
        logSb.AppendLine($"---- FIELD-ENUM {DateTime.Now:yyyy-MM-dd HH:mm:ss} eFaturaId={eFaturaId} Count={totalCount} ----");

        string[] tryProps = new[] { "Name", "FieldName", "DBFieldName", "FieldNameDB", "TitleEn", "Caption", "Title" };
        string? bestProp = null;
        int uiAdded = 0;

        for (int i = 0; i < limit; i++)
        {
            object? field = null;
            try
            {
                field = IDispCall(df, "Item",
                    (ushort)(DISPATCH_METHOD | DISPATCH_PROPERTYGET),
                    new object?[] { i });
            }
            catch (Exception ex)
            {
                logSb.AppendLine($"  [{i:D3}] ITEM-EX:{ex.GetType().Name}:{Truncate(ex.Message, 80)}");
                continue;
            }
            if (field == null) { logSb.AppendLine($"  [{i:D3}] null"); continue; }

            // İlk field için tüm üyeleri dök ki hangi property gerçek ismi taşıyor görelim
            if (i == 0)
            {
                try { logSb.AppendLine("  FIELD0-MEMBERS=" + ProbeMembers(field)); }
                catch (Exception ex) { logSb.AppendLine("  FIELD0-MEMBERS-EX:" + ex.GetType().Name); }
                try { logSb.AppendLine("  FIELD0-TYPE=" + ProbeTypeName(field)); }
                catch { }
            }

            var lineSb = new System.Text.StringBuilder($"  [{i:D3}]");
            string? primaryName = null;
            foreach (var prop in tryProps)
            {
                try
                {
                    object? v = IDispCall(field, prop, DISPATCH_PROPERTYGET, Array.Empty<object?>());
                    string val = v == null ? "<null>" : (v.ToString() ?? "<null>");
                    lineSb.Append(' ').Append(prop).Append('=').Append(val);
                    if (primaryName == null && !string.IsNullOrEmpty(val) && val != "<null>")
                    {
                        primaryName = val;
                        if (bestProp == null) bestProp = prop;
                    }
                }
                catch (Exception ex) { lineSb.Append(' ').Append(prop).Append("=EX:").Append(ex.GetType().Name); }
            }
            logSb.AppendLine(lineSb.ToString());

            if (uiAdded < 12 && primaryName != null)
            {
                if (uiAdded > 0) uiSb.Append(',');
                uiSb.Append(i).Append(':').Append(primaryName);
                uiAdded++;
            }
            try { Marshal.ReleaseComObject(field); } catch { }
        }

        try
        {
            var logFile = System.IO.Path.Combine(AppDataPaths.DataRoot, "logo_unity_diag.txt");
            System.IO.File.AppendAllText(logFile, logSb.ToString() + "\n");
        }
        catch { /* best-effort */ }

        return $"[ENUM bestProp={bestProp ?? "?"}:{uiSb}]";
    }

    // Kısa, kritik 3-4 cevap — UI'de görünür yer kaybetmesin diye.
    private static string QuickCriticalProbe(object inv, object df)
    {
        string invType = "?", dfType = "?", dfCount = "?";
        try { invType = ProbeTypeName(inv); } catch { invType = "EX"; }
        try { dfType = ProbeTypeName(df); } catch { dfType = "EX"; }
        try
        {
            object? c = IDispCall(df, "Count", DISPATCH_PROPERTYGET, Array.Empty<object?>());
            dfCount = c == null ? "null" : c.ToString() ?? "?";
        }
        catch (Exception ex) { dfCount = "EX:" + ex.GetType().Name; }
        return $"inv={invType} df={dfType} df.Count={dfCount}";
    }

    // Hangi dispatch mekanizması Logo Unity COM ile çalışıyor?
    private static string ProbeDispatch(object inv, object df)
    {
        var sb = new System.Text.StringBuilder();
        var t = df.GetType();
        const string probe = "TYPE";

        // FRESH-DF: VBA pattern — cache'lenmemiş, fresh DataFields ile FieldByName
        sb.Append("[FRESH-DF:");
        try
        {
            object? freshDf = IDispCall(inv, "DataFields", DISPATCH_PROPERTYGET, Array.Empty<object?>());
            if (freshDf == null) sb.Append("null");
            else
            {
                object? r = IDispCall(freshDf, "FieldByName",
                    (ushort)(DISPATCH_METHOD | DISPATCH_PROPERTYGET),
                    new object?[] { probe });
                sb.Append(r == null ? "field=null" : "OK(" + r.GetType().Name + ")");
            }
        }
        catch (Exception ex)
        {
            sb.Append("EX " + ex.GetType().Name + ":" + Truncate(ex.Message, 100));
        }
        sb.Append("]");

        // GET-FIELD-IDX(TYPE) — GetFieldIndex#1 metodu çalışıyor mu?
        sb.Append("[GFI-TYPE:");
        try
        {
            object? r = IDispCall(df, "GetFieldIndex",
                (ushort)(DISPATCH_METHOD | DISPATCH_PROPERTYGET),
                new object?[] { probe });
            sb.Append(r == null ? "null" : r.ToString());
        }
        catch (Exception ex) { sb.Append("EX " + ex.GetType().Name + ":" + Truncate(ex.Message, 60)); }
        sb.Append("]");

        // ITEM(0) — koleksiyonda gerçek field nesnesi var mı?
        sb.Append("[ITEM0:");
        try
        {
            object? r = IDispCall(df, "Item",
                (ushort)(DISPATCH_METHOD | DISPATCH_PROPERTYGET),
                new object?[] { 0 });
            sb.Append(r == null ? "null" : "OK(" + r.GetType().Name + ")");
        }
        catch (Exception ex) { sb.Append("EX " + ex.GetType().Name + ":" + Truncate(ex.Message, 60)); }
        sb.Append("]");

        // SetFieldValue(TYPE, 99) — tek atışlık metod çalışıyor mu? (asıl ümit)
        sb.Append("[SFV-TYPE:");
        try
        {
            IDispCall(df, "SetFieldValue",
                (ushort)(DISPATCH_METHOD | DISPATCH_PROPERTYGET),
                new object?[] { probe, 99 });
            sb.Append("OK");
        }
        catch (Exception ex) { sb.Append("EX " + ex.GetType().Name + ":" + Truncate(ex.Message, 80)); }
        sb.Append("]");

        // INV-TYPE: inv (DataObject) gerçekten purchase invoice mi yoksa generic mi?
        sb.Append("[INV-TYPE:");
        try { sb.Append(ProbeTypeName(inv)); }
        catch (Exception ex) { sb.Append("EX " + ex.GetType().Name); }
        sb.Append("]");

        // DF-COUNT: DataFields koleksiyonu boş mu? .New() schema'yı doldurmadıysa Count=0 olur.
        sb.Append("[DF-COUNT:");
        try
        {
            object? c = IDispCall(df, "Count", DISPATCH_PROPERTYGET, Array.Empty<object?>());
            sb.Append(c == null ? "null" : c.ToString());
        }
        catch (Exception ex) { sb.Append("EX " + ex.GetType().Name + ":" + Truncate(ex.Message, 60)); }
        sb.Append("]");

        // INV-COUNT-OR-COMMANDS: inv üzerinde New/Load/Read gibi metodlar var mı?
        sb.Append("[INV-MEMBERS:");
        try { sb.Append(ProbeMembers(inv)); }
        catch (Exception ex) { sb.Append("EX " + ex.GetType().Name); }
        sb.Append("]");

        // GP — TargetInvocationException'ın InnerException'ı asıl COM hatasını barındırır
        sb.Append("[GP-inner:");
        try
        {
            var r = t.InvokeMember("FieldByName", BindingFlags.GetProperty, null, df, new object[] { probe });
            sb.Append(r == null ? "null" : "OK(" + r.GetType().Name + ")");
        }
        catch (Exception ex)
        {
            var real = ex.InnerException ?? ex;
            sb.Append(real.GetType().Name + " HR=0x" + (real is COMException ce ? ce.HResult.ToString("X8") : "?") + " : " + Truncate(real.Message, 120));
        }
        sb.Append("]");

        // IDispatch direkt — VARIANT'ın tipini de raporla
        sb.Append("[IDISP-vt:");
        try
        {
            int dispid;
            ushort vtOut;
            object? r = InvokeViaIDispatchVerbose(df, "FieldByName", out dispid, out vtOut, probe);
            sb.Append($"dispid={dispid},vt={vtOut},result=");
            sb.Append(r == null ? "null" : "OK(" + r.GetType().Name + ")");
        }
        catch (Exception ex)
        {
            sb.Append("EX " + ex.GetType().Name + " HR=0x" + (ex is COMException ce ? ce.HResult.ToString("X8") : "?") + " : " + Truncate(ex.Message, 120));
        }
        sb.Append("]");

        // df objesinin TypeInfo'sunu sorgula
        sb.Append("[TI:");
        try
        {
            var disp = (IDispatch)df;
            uint cnt;
            int hr0 = disp.GetTypeInfoCount(out cnt);
            sb.Append($"hr=0x{hr0:X8},count={cnt}");
        }
        catch (Exception ex) { sb.Append("EX " + ex.GetType().Name); }
        sb.Append("]");

        // df üzerinde başka bir alanı dene — belki sadece TYPE problem
        sb.Append("[TX-vt:");
        try
        {
            int dispid;
            ushort vtOut;
            object? r = InvokeViaIDispatchVerbose(df, "FieldByName", out dispid, out vtOut, "TRANSACTIONS");
            sb.Append($"dispid={dispid},vt={vtOut},result=");
            sb.Append(r == null ? "null" : "OK(" + r.GetType().Name + ")");
        }
        catch (Exception ex)
        {
            sb.Append("EX " + ex.GetType().Name + " HR=0x" + (ex is COMException ce ? ce.HResult.ToString("X8") : "?") + " : " + Truncate(ex.Message, 120));
        }
        sb.Append("]");

        // YENI: IDispCall (manuel VT_DISPATCH extraction) ile test
        sb.Append("[NEW-IDC:");
        try
        {
            object? r = IDispCall(df, "FieldByName",
                (ushort)(DISPATCH_METHOD | DISPATCH_PROPERTYGET),
                new object?[] { probe });
            sb.Append(r == null ? "null" : "OK(" + r.GetType().Name + ")");
        }
        catch (Exception ex)
        {
            sb.Append("EX " + ex.GetType().Name + ":" + Truncate(ex.Message, 100));
        }
        sb.Append("]");

        // YENI: df'nin ITypeInfo'sundan tip adı + ilk birkaç method
        sb.Append("[TYPENAME:");
        try { sb.Append(ProbeTypeName(df)); }
        catch (Exception ex) { sb.Append("EX " + ex.GetType().Name + ":" + Truncate(ex.Message, 80)); }
        sb.Append("]");

        // YENI: df'nin sahip olduğu metod/property listesi
        sb.Append("[MEMBERS:");
        try { sb.Append(ProbeMembers(df)); }
        catch (Exception ex) { sb.Append("EX " + ex.GetType().Name + ":" + Truncate(ex.Message, 80)); }
        sb.Append("]");

        // YENI: VARIANT'ın raw byte dökümü
        sb.Append("[VAR-RAW:");
        try
        {
            int dispid2; ushort vt2;
            byte[] rawBytes = ProbeRawVariant(df, "FieldByName", out dispid2, out vt2, probe);
            sb.Append("vt=" + vt2 + ",bytes=");
            for (int i = 0; i < Math.Min(rawBytes.Length, 24); i++)
            {
                sb.Append(rawBytes[i].ToString("X2"));
                if (i == 7 || i == 15) sb.Append('|');
            }
        }
        catch (Exception ex)
        {
            sb.Append("EX " + ex.GetType().Name + ":" + Truncate(ex.Message, 80));
        }
        sb.Append("]");

        return sb.ToString();
    }

    // Variant'ın ham byte içeriğini döndürür — pdispVal IntPtr'sini gözle görmek için.
    private static byte[] ProbeRawVariant(object target, string member, out int dispid, out ushort vt, params object?[] args)
    {
        dispid = 0; vt = 0;
        var disp = (IDispatch)target;
        var iidNull = Guid.Empty;
        var names = new[] { member };
        var ids = new int[1];
        int hr = disp.GetIDsOfNames(ref iidNull, names, 1, 0, ids);
        if (hr < 0) throw Marshal.GetExceptionForHR(hr) ?? new COMException($"GetIDsOfNames hr=0x{hr:X8}", hr);
        dispid = ids[0];

        int vsize = Marshal.SizeOf<Variant>();
        IntPtr rgvarg = IntPtr.Zero;
        if (args.Length > 0)
        {
            rgvarg = Marshal.AllocCoTaskMem(vsize * args.Length);
            for (int i = 0; i < args.Length; i++)
            {
                var v = new Variant();
                v.SetValue(args[args.Length - 1 - i]);
                Marshal.StructureToPtr(v, rgvarg + i * vsize, false);
            }
        }
        try
        {
            var dp = new DISPPARAMS { rgvarg = rgvarg, rgdispidNamedArgs = IntPtr.Zero, cArgs = args.Length, cNamedArgs = 0 };
            IntPtr pVarResult = Marshal.AllocCoTaskMem(vsize);
            try
            {
                Marshal.StructureToPtr(new Variant(), pVarResult, false);
                hr = disp.Invoke(ids[0], ref iidNull, 0, (ushort)(DISPATCH_METHOD | DISPATCH_PROPERTYGET), ref dp, pVarResult, IntPtr.Zero, IntPtr.Zero);
                if (hr < 0) throw Marshal.GetExceptionForHR(hr) ?? new COMException($"Invoke hr=0x{hr:X8}", hr);
                vt = (ushort)Marshal.ReadInt16(pVarResult, 0);
                byte[] bytes = new byte[vsize];
                Marshal.Copy(pVarResult, bytes, 0, vsize);
                return bytes;
            }
            finally
            {
                try { NativeMethods.VariantClear(pVarResult); } catch { }
                Marshal.FreeCoTaskMem(pVarResult);
            }
        }
        finally
        {
            if (rgvarg != IntPtr.Zero)
            {
                for (int i = 0; i < args.Length; i++)
                {
                    try { NativeMethods.VariantClear(rgvarg + i * vsize); } catch { }
                }
                Marshal.FreeCoTaskMem(rgvarg);
            }
        }
    }

    // df'nin ITypeInfo'sundan tip adı — DataFields mı yoksa başka bir şey mi olduğunu anlamak için.
    private static string ProbeTypeName(object df)
    {
        var disp = (IDispatch)df;
        if (disp.GetTypeInfoCount(out uint cnt) < 0 || cnt == 0) return "no_typeinfo";
        if (disp.GetTypeInfo(0, 0, out IntPtr pTi) < 0 || pTi == IntPtr.Zero) return "gtinfo_failed";
        try
        {
            var ti = (System.Runtime.InteropServices.ComTypes.ITypeInfo)Marshal.GetObjectForIUnknown(pTi);
            try
            {
                ti.GetDocumentation(-1, out string? name, out _, out _, out _);
                return name ?? "?";
            }
            finally { try { Marshal.ReleaseComObject(ti); } catch { } }
        }
        finally { Marshal.Release(pTi); }
    }

    // df üzerindeki TÜM method/property adı + memid — FieldByName gerçekten var mı?
    private static string ProbeMembers(object df)
    {
        var disp = (IDispatch)df;
        if (disp.GetTypeInfoCount(out uint cnt) < 0 || cnt == 0) return "no_typeinfo";
        if (disp.GetTypeInfo(0, 0, out IntPtr pTi) < 0 || pTi == IntPtr.Zero) return "no_ti";
        try
        {
            var ti = (System.Runtime.InteropServices.ComTypes.ITypeInfo)Marshal.GetObjectForIUnknown(pTi);
            try
            {
                ti.GetTypeAttr(out IntPtr pAttr);
                try
                {
                    var attr = Marshal.PtrToStructure<System.Runtime.InteropServices.ComTypes.TYPEATTR>(pAttr);
                    int total = attr.cFuncs;
                    var sb = new System.Text.StringBuilder();
                    sb.Append("funcs=").Append(total).Append(":");
                    for (int i = 0; i < total; i++)
                    {
                        ti.GetFuncDesc(i, out IntPtr pFunc);
                        try
                        {
                            var func = Marshal.PtrToStructure<System.Runtime.InteropServices.ComTypes.FUNCDESC>(pFunc);
                            string[] names = new string[1];
                            ti.GetNames(func.memid, names, 1, out _);
                            sb.Append(names[0] ?? "?").Append('#').Append(func.memid).Append(',');
                        }
                        finally { ti.ReleaseFuncDesc(pFunc); }
                    }
                    return sb.ToString();
                }
                finally { ti.ReleaseTypeAttr(pAttr); }
            }
            finally { try { Marshal.ReleaseComObject(ti); } catch { } }
        }
        finally { Marshal.Release(pTi); }
    }

    private static string Truncate(string s, int max) =>
        string.IsNullOrEmpty(s) ? "" : (s.Length <= max ? s : s.Substring(0, max));

    // Verbose: IDispatch::Invoke, dispid ve dönen VARIANT'ın vt'sini de raporlar.
    private static object? InvokeViaIDispatchVerbose(object target, string member, out int dispid, out ushort vtResult, params object?[] args)
    {
        dispid = 0;
        vtResult = 0;
        var disp = (IDispatch)target;
        var iidNull = Guid.Empty;
        var names = new[] { member };
        var ids = new int[1];
        int hr = disp.GetIDsOfNames(ref iidNull, names, 1, 0, ids);
        if (hr < 0) throw Marshal.GetExceptionForHR(hr) ?? new InvalidOperationException($"GetIDsOfNames hr=0x{hr:X8}");
        dispid = ids[0];

        IntPtr rgvarg = IntPtr.Zero;
        if (args.Length > 0)
        {
            int vsize = Marshal.SizeOf<Variant>();
            rgvarg = Marshal.AllocCoTaskMem(vsize * args.Length);
            for (int i = 0; i < args.Length; i++)
            {
                var v = new Variant();
                v.SetValue(args[args.Length - 1 - i]);
                Marshal.StructureToPtr(v, rgvarg + i * vsize, false);
            }
        }
        try
        {
            var dp = new DISPPARAMS { rgvarg = rgvarg, rgdispidNamedArgs = IntPtr.Zero, cArgs = args.Length, cNamedArgs = 0 };
            const ushort DISPATCH_METHOD = 1, DISPATCH_PROPERTYGET = 2;
            IntPtr pVarResult = Marshal.AllocCoTaskMem(Marshal.SizeOf<Variant>());
            try
            {
                Marshal.StructureToPtr(new Variant(), pVarResult, false);
                hr = disp.Invoke(ids[0], ref iidNull, 0, DISPATCH_METHOD | DISPATCH_PROPERTYGET, ref dp, pVarResult, IntPtr.Zero, IntPtr.Zero);
                if (hr < 0) throw Marshal.GetExceptionForHR(hr) ?? new InvalidOperationException($"Invoke hr=0x{hr:X8}");

                vtResult = (ushort)Marshal.ReadInt16(pVarResult, 0);
                return Marshal.GetObjectForNativeVariant(pVarResult);
            }
            finally
            {
                try { NativeMethods.VariantClear(pVarResult); } catch { }
                Marshal.FreeCoTaskMem(pVarResult);
            }
        }
        finally
        {
            if (rgvarg != IntPtr.Zero)
            {
                int vsize = Marshal.SizeOf<Variant>();
                for (int i = 0; i < args.Length; i++)
                {
                    try { NativeMethods.VariantClear(rgvarg + i * vsize); } catch { }
                }
                Marshal.FreeCoTaskMem(rgvarg);
            }
        }
    }

    // IDispatch::Invoke ile doğrudan çağrı — .NET binder'ı tamamen bypass eder.
    private static object? InvokeViaIDispatch(object target, string member, params object?[] args)
    {
        var disp = (IDispatch)target;
        var iidNull = Guid.Empty;
        var names = new[] { member };
        var ids = new int[1];
        int hr = disp.GetIDsOfNames(ref iidNull, names, 1, 0, ids);
        if (hr < 0) throw Marshal.GetExceptionForHR(hr) ?? new InvalidOperationException($"GetIDsOfNames hr=0x{hr:X8}");

        // VARIANT'ları ters sırada koy (DISPPARAMS rgvarg LIFO)
        IntPtr rgvarg = IntPtr.Zero;
        if (args.Length > 0)
        {
            int vsize = Marshal.SizeOf<Variant>();
            rgvarg = Marshal.AllocCoTaskMem(vsize * args.Length);
            for (int i = 0; i < args.Length; i++)
            {
                var v = new Variant();
                v.SetValue(args[args.Length - 1 - i]);
                Marshal.StructureToPtr(v, rgvarg + i * vsize, false);
            }
        }

        try
        {
            var dp = new DISPPARAMS
            {
                rgvarg = rgvarg,
                rgdispidNamedArgs = IntPtr.Zero,
                cArgs = args.Length,
                cNamedArgs = 0
            };

            const ushort DISPATCH_METHOD = 1;
            const ushort DISPATCH_PROPERTYGET = 2;
            object? result = null;
            IntPtr pVarResult = Marshal.AllocCoTaskMem(Marshal.SizeOf<Variant>());
            try
            {
                // Variant null'a init
                Marshal.StructureToPtr(new Variant(), pVarResult, false);

                hr = disp.Invoke(ids[0], ref iidNull, 0, DISPATCH_METHOD | DISPATCH_PROPERTYGET,
                                  ref dp, pVarResult, IntPtr.Zero, IntPtr.Zero);
                if (hr < 0) throw Marshal.GetExceptionForHR(hr) ?? new InvalidOperationException($"Invoke hr=0x{hr:X8}");

                result = Marshal.GetObjectForNativeVariant(pVarResult);
            }
            finally
            {
                try { NativeMethods.VariantClear(pVarResult); } catch { }
                Marshal.FreeCoTaskMem(pVarResult);
            }
            return result;
        }
        finally
        {
            if (rgvarg != IntPtr.Zero)
            {
                int vsize = Marshal.SizeOf<Variant>();
                for (int i = 0; i < args.Length; i++)
                {
                    try { NativeMethods.VariantClear(rgvarg + i * vsize); } catch { }
                }
                Marshal.FreeCoTaskMem(rgvarg);
            }
        }
    }

    [ComImport]
    [Guid("00020400-0000-0000-C000-000000000046")]
    [InterfaceType(ComInterfaceType.InterfaceIsIUnknown)]
    internal interface IDispatch
    {
        [PreserveSig] int GetTypeInfoCount(out uint pctinfo);
        [PreserveSig] int GetTypeInfo(uint iTInfo, uint lcid, out IntPtr typeInfo);
        [PreserveSig] int GetIDsOfNames(
            ref Guid riid,
            [MarshalAs(UnmanagedType.LPArray, ArraySubType = UnmanagedType.LPWStr)] string[] rgsNames,
            uint cNames, uint lcid,
            [Out, MarshalAs(UnmanagedType.LPArray)] int[] rgDispId);
        [PreserveSig] int Invoke(
            int dispIdMember,
            ref Guid riid,
            uint lcid,
            ushort wFlags,
            ref DISPPARAMS pDispParams,
            IntPtr pVarResult,
            IntPtr pExcepInfo,
            IntPtr puArgErr);
    }

    [StructLayout(LayoutKind.Sequential)]
    internal struct DISPPARAMS
    {
        public IntPtr rgvarg;
        public IntPtr rgdispidNamedArgs;
        public int cArgs;
        public int cNamedArgs;
    }

    [StructLayout(LayoutKind.Explicit, Size = 24)]
    internal struct Variant
    {
        [FieldOffset(0)] public ushort vt;
        [FieldOffset(8)] public IntPtr pdispVal;
        [FieldOffset(8)] public IntPtr bstrVal;
        [FieldOffset(8)] public int intVal;
        [FieldOffset(8)] public long longVal;
        [FieldOffset(8)] public double dblVal;

        public void SetValue(object? v)
        {
            if (v == null) { vt = 1; return; } // VT_NULL
            switch (v)
            {
                case string s:
                    vt = 8; // VT_BSTR
                    bstrVal = Marshal.StringToBSTR(s);
                    break;
                case int i:
                    vt = 3; // VT_I4
                    intVal = i;
                    break;
                case long l:
                    vt = 20; // VT_I8
                    longVal = l;
                    break;
                case double d:
                    vt = 5; // VT_R8
                    dblVal = d;
                    break;
                case short sh:
                    vt = 2; // VT_I2
                    intVal = sh;
                    break;
                case bool b:
                    vt = 11; // VT_BOOL
                    intVal = b ? -1 : 0;
                    break;
                case DateTime dt:
                    vt = 7; // VT_DATE
                    dblVal = dt.ToOADate();
                    break;
                default:
                    vt = 8;
                    bstrVal = Marshal.StringToBSTR(v.ToString() ?? "");
                    break;
            }
        }
    }

    internal static class NativeMethods
    {
        [System.Runtime.InteropServices.DllImport("oleaut32.dll")]
        public static extern int VariantClear(IntPtr pvarg);
    }

    // inv.DataFields.FieldByName(name).Value (string)
    private static string GetFieldStr(object inv, string name)
    {
        try
        {
            object? df = GetProp(inv, "DataFields");
            if (df == null) return "";
            object? field = Inv(df, "FieldByName", name);
            if (field == null) return "";
            object? v = GetProp(field, "Value");
            return Convert.ToString(v) ?? "";
        }
        catch { return ""; }
    }

    private static long GetFieldLong(object inv, string name)
    {
        string s = GetFieldStr(inv, name);
        return long.TryParse(s, out long v) ? v : 0L;
    }

    // app.GetLastError() veya app.ErrorCode property — string olarak güvenli oku
    private static string GetStr(object? obj, string member)
    {
        if (obj == null) return "";
        try
        {
            object? v = IDispCall(obj, member,
                (ushort)(DISPATCH_METHOD | DISPATCH_PROPERTYGET),
                Array.Empty<object?>());
            return Convert.ToString(v) ?? "";
        }
        catch { return ""; }
    }

    private static void CollectValidateErrors(object inv, List<string> hedef)
    {
        try
        {
            object? errs = GetProp(inv, "ValidateErrors");
            if (errs == null) return;
            int n = Convert.ToInt32(GetProp(errs, "Count") ?? 0);
            for (int i = 0; i < n; i++)
            {
                object? e = GetLine(errs, i);
                if (e == null) continue;
                string id  = Convert.ToString(GetProp(e, "ID"))    ?? "";
                string msg = Convert.ToString(GetProp(e, "Error")) ?? "";
                hedef.Add($"({id}) {msg}");
            }
        }
        catch { /* ignore */ }
    }

    private static void Release(object? comObj)
    {
        if (comObj == null) return;
        try
        {
            if (Marshal.IsComObject(comObj))
                Marshal.FinalReleaseComObject(comObj);
        }
        catch { /* ignore */ }
    }

    // =================================================================
    // TOPTAN SATIŞ FATURASI (doSalesInvoice = 19) — Ayni Avans → Logo
    // Referans: SabNet PMHS WinForms (AyniAvansEntegrasyonu.vb)
    // Not: KeyNet PMHS.dll enum'undan doSalesInvoice=19. SabnetFaturaTransfer
    //      14 May 2026'da 18→19 değişimiyle 1159 hatasını çözdü.
    // =================================================================

    private const int DO_SALES_INVOICE = 19;
    private const int DO_CARI_KART     = 30;  // doAccountsRP

    public Task<LogoSatisAktarimSonuc> AktarSalesAsync(
        LogoUnityCredentials cred,
        IReadOnlyList<LogoSatisFaturaModel> batch,
        CancellationToken ct = default)
    {
        var tcs = new TaskCompletionSource<LogoSatisAktarimSonuc>(TaskCreationOptions.RunContinuationsAsynchronously);
        var th = new Thread(() =>
        {
            try { tcs.TrySetResult(ExecuteSalesBatchOnStaThread(cred, batch, ct)); }
            catch (Exception ex) { _log.LogError(ex, "Logo Unity sales batch hata"); tcs.TrySetException(ex); }
        });
        th.SetApartmentState(ApartmentState.STA);
        th.IsBackground = true;
        th.Name = "LogoUnitySalesSta";
        th.Start();
        return tcs.Task;
    }

    private LogoSatisAktarimSonuc ExecuteSalesBatchOnStaThread(
        LogoUnityCredentials cred,
        IReadOnlyList<LogoSatisFaturaModel> batch,
        CancellationToken ct)
    {
        var sonuc = new LogoSatisAktarimSonuc();
        object? app = null;
        bool loggedIn = false;

        try
        {
            app = CreateUnityApp();
            bool ok;
            try { ok = (bool)(Inv(app, "Login", cred.Kullanici, cred.Sifre, cred.FirmaNo) ?? false); }
            catch (Exception ex) { sonuc.LoginHata = $"Login çağrısı hata: {ex.Message}"; return sonuc; }

            if (!ok)
            {
                sonuc.LoginHata = $"Logo Login reddetti. {GetStr(app, "GetLastError")} - {GetStr(app, "GetLastErrorString")}";
                return sonuc;
            }

            loggedIn = true;
            sonuc.LoginBasarili = true;

            foreach (var m in batch)
            {
                ct.ThrowIfCancellationRequested();
                // 1) Cari kart upsert (sessiz başarısız edebilir — Logo "zaten var" demiş olabilir)
                try { UpsertCariKart(app, m); }
                catch (Exception ex) { _log.LogWarning(ex, "Cari upsert hata: {Cari}", m.CariKod); }

                // 2) Satış faturası post
                var r = PostSingleSalesInvoice(app, m);
                sonuc.Faturalar.Add(r);
            }
        }
        finally
        {
            if (loggedIn && app != null) { try { Inv(app, "Logout"); } catch { } }
            Release(app);
        }
        return sonuc;
    }

    private void UpsertCariKart(object app, LogoSatisFaturaModel m)
    {
        object? chk = null;
        try
        {
            chk = Inv(app, "NewDataObject", DO_CARI_KART);
            if (chk == null) return;
            try { Inv(chk, "New"); } catch { /* devam */ }

            SetHeaderField(chk, "ACCOUNT_TYPE", 3);
            SetHeaderField(chk, "CODE",         m.CariKod);
            SetHeaderField(chk, "TITLE",        m.CariUnvan);
            SetHeaderField(chk, "TITLE2",       m.CariUnvan);
            SetHeaderField(chk, "AUXIL_CODE",   "SMUSTAHSIL");
            if (!string.IsNullOrWhiteSpace(m.CariTelefon)) SetHeaderField(chk, "TELEPHONE1", m.CariTelefon);
            if (!string.IsNullOrWhiteSpace(m.CariEPosta))  SetHeaderField(chk, "E_MAIL",     m.CariEPosta);
            SetHeaderField(chk, "CONTACT",      m.CariKod);
            SetHeaderField(chk, "PAYMENT_CODE", "PESIN");
            SetHeaderField(chk, "TRADING_GRP",  m.TradingGrp);
            SetHeaderField(chk, "GL_CODE",      "320.07.0000");  // SabNet referansından
            SetHeaderField(chk, "CREDIT_TYPE",  1);
            SetHeaderField(chk, "RISKFACT_CHQ",    1);
            SetHeaderField(chk, "RISKFACT_PROMNT", 1);
            SetHeaderField(chk, "PURCHBRWS",   1);
            SetHeaderField(chk, "SALESBRWS",   1);
            SetHeaderField(chk, "IMPBRWS",     1);
            SetHeaderField(chk, "EXPBRWS",     1);
            SetHeaderField(chk, "FINBRWS",     1);
            SetHeaderField(chk, "COLLATRLRISK_TYPE", 1);
            SetHeaderField(chk, "PERSCOMPANY", 1);
            SetHeaderField(chk, "TCKNO",       m.CariTcKimlik);
            SetHeaderField(chk, "USED_IN_PERIODS", 1);

            try { Inv(chk, "Post"); } catch { /* Logo "zaten var" → tolere */ }
        }
        finally { Release(chk); }
    }

    private LogoSatisFaturaResult PostSingleSalesInvoice(object app, LogoSatisFaturaModel m)
    {
        var r = new LogoSatisFaturaResult
        {
            FaturaNo = m.FaturaNo,
            FormNo   = m.FormNo,
            HesapNo  = m.HesapNo,
        };

        string diagPath = System.IO.Path.Combine(AppDataPaths.DataRoot, "logo_unity_diag.txt");
        void Diag(string s)
        {
            try { System.IO.File.AppendAllText(diagPath, $"[{DateTime.Now:HH:mm:ss}] {s}\n"); } catch { }
        }
        Diag($"==== SALES POST baslangic FaturaNo={m.FaturaNo} Tip={m.Tip} ARP={m.CariKod} MASTER={m.MasterCode} Qty={m.Miktar} Price={m.BirimFiyat} KDV={m.KdvOran} (GRPCODE=2) ====");

        object? inv = null;
        try
        {
            inv = Inv(app, "NewDataObject", DO_SALES_INVOICE);
            if (inv == null) { r.Hata = $"NewDataObject({DO_SALES_INVOICE}) null döndü"; Diag(r.Hata); return r; }
            Diag($"NewDataObject({DO_SALES_INVOICE}) OK");

            try { Inv(inv, "New"); }
            catch (Exception ex) { r.Hata = "inv.New() hata: " + ex.Message; return r; }

            // Header
            // GRPCODE: DOGUSNDB'de geçmişte basılmış TRCODE=8 faturanın GRPCODE=2 (SQL ile doğrulandı).
            // Logo Tiger 3'te bu firmaya özel eşleşme bu; 1159 "Fiş tip bilgisi fiş modülüne
            // uygun değil" hatası GRPCODE değeri TRCODE ile uyumsuz olduğunda çıkıyor.
            int grpCode = m.Tip switch
            {
                8  => 2,   // Toptan Satış (LG_211_01_INVOICE örneği: GRPCODE=2, TRCODE=8)
                1  => 1,   // Mal Alış
                2  => 2,   // Perakende Satış İade
                3  => 3,   // Toptan Satış İade
                4  => 5,   // Alınan Hizmet
                5  => 4,   // Verilen Hizmet
                6  => 2,   // Perakende Satış
                7  => 4,   // Verilen Proforma
                9  => 4,   // Verilen Vade Farkı
                10 => 5,   // Alınan Vade Farkı
                13 => 1,   // Müstahsil Makbuzu (alış grubu)
                _  => 2
            };
            SetHeaderField(inv, "GRPCODE",         grpCode);
            SetHeaderField(inv, "TYPE",            m.Tip);                  // 8 toptan satış
            SetHeaderField(inv, "NUMBER",          m.FaturaNo);
            SetHeaderField(inv, "DATE",            m.Tarih);
            SetHeaderField(inv, "DOC_DATE",        m.Tarih);
            SetHeaderField(inv, "TIME",            m.TimeKodu);
            SetHeaderField(inv, "ARP_CODE",        m.CariKod);
            SetHeaderField(inv, "POST_FLAGS",      m.PostFlags);
            SetHeaderField(inv, "VAT_RATE",        (double)m.KdvOran);
            SetHeaderField(inv, "SOURCE_WH",       m.Ambar);
            SetHeaderField(inv, "SOURCE_COST_GRP", 0);
            if (!string.IsNullOrWhiteSpace(m.Isyeri))    SetHeaderField(inv, "DIVISION",    m.Isyeri);
            if (!string.IsNullOrWhiteSpace(m.Bolum))     SetHeaderField(inv, "DEPARTMENT",  m.Bolum);
            SetHeaderField(inv, "FACTORY",     m.Fabrika);
            SetHeaderField(inv, "TRADING_GRP", m.TradingGrp);
            SetHeaderField(inv, "PAYMENT_CODE", 0);
            if (!string.IsNullOrWhiteSpace(m.ProjeKodu)) SetHeaderField(inv, "PROJECT_CODE", m.ProjeKodu);

            // Müstahsil/Satış faturası — sabit sevk adresi ve taşıyıcı kartı:
            // Tüm müstahsillerde sevk adresi "001", taşıyıcı kartı "01" kullanılıyor.
            SetHeaderField(inv, "SHIPLOC_CODE",    "001");    // Sevkiyat adresi kodu (cariye bağlı SHIPINFO)
            SetHeaderField(inv, "SHIPPING_AGENT", "01");      // Taşıyıcı (nakliyeci) cari kart kodu
            SetHeaderField(inv, "SHIPMENT_TYPE",   1);         // 1 = Kara yolu

            // Transaction (tek satır)
            object? linesObj = null;
            try
            {
                var df2 = IDispCall(inv, "DataFields", DISPATCH_PROPERTYGET, Array.Empty<object?>());
                var fb = IDispCall(df2, "FieldByName", DISPATCH_METHOD, new object?[] { "TRANSACTIONS" });
                linesObj = IDispCall(fb, "Lines", DISPATCH_PROPERTYGET, Array.Empty<object?>());
            }
            catch (Exception ex) { r.Hata = "TRANSACTIONS.Lines erişimi hata: " + ex.Message; return r; }

            if (linesObj == null) { r.Hata = "TRANSACTIONS.Lines null"; return r; }
            try { IDispCall(linesObj, "AppendLine", DISPATCH_METHOD, Array.Empty<object?>()); }
            catch (Exception ex) { r.Hata = "AppendLine hata: " + ex.Message; return r; }

            // Yeni eklenen satıra eriş — Lines.Item(Count-1)
            object? satir = null;
            try
            {
                var cnt = IDispCall(linesObj, "Count", DISPATCH_PROPERTYGET, Array.Empty<object?>());
                int count = Convert.ToInt32(cnt);
                satir = IDispCall(linesObj, "Item", DISPATCH_PROPERTYGET, new object?[] { count - 1 });
            }
            catch (Exception ex) { _log.LogWarning(ex, "Lines.Item(Count-1) erişimi hata"); }
            if (satir == null) { r.Hata = "Yeni satır referansı alınamadı"; return r; }

            SetField(satir, "TYPE",          0);                       // XML 0 (SabNet kodunda 4'tü, önce XML değerini deneyelim)
            SetField(satir, "SOURCEINDEX",   m.Ambar);
            SetField(satir, "SOURCECOSTGRP", 0);
            SetField(satir, "MASTER_CODE",   m.MasterCode);
            SetField(satir, "QUANTITY",      (double)m.Miktar);
            SetField(satir, "PRICE",         (double)m.BirimFiyat);
            SetField(satir, "UNIT_CODE",     m.Birim);
            SetField(satir, "VAT_RATE",      (double)m.KdvOran);
            SetField(satir, "EDT_CURR",      1);
            if (!string.IsNullOrWhiteSpace(m.ProjeKodu))
                SetField(satir, "PROJECT_CODE", m.ProjeKodu);

            // Post
            bool postOk = false;
            try
            {
                var pr = Inv(inv, "Post");
                postOk = pr is bool b && b;
                Diag($"Post() ret={pr ?? "null"} → postOk={postOk}");
            }
            catch (Exception ex) { r.Hata = "Post çağrısı hata: " + ex.Message; Diag(r.Hata); }

            if (!postOk)
            {
                r.Hata = (r.Hata.Length == 0 ? "Post=false" : r.Hata) + " | " + CollectInvoiceError(inv);
                CollectValidateErrors(inv, r.ValidateErrors);
                Diag("HATA: " + r.Hata);
                if (r.ValidateErrors.Count > 0)
                    Diag("ValidateErrors: " + string.Join(" || ", r.ValidateErrors));
                return r;
            }
            Diag("POST OK");

            // Başarılı — LOGICALREF al
            try
            {
                var df = IDispCall(inv, "DataFields", DISPATCH_PROPERTYGET, Array.Empty<object?>());
                if (df != null)
                {
                    var fld = IDispCall(df, "FieldByName", DISPATCH_METHOD, new object?[] { "INTERNAL_REFERENCE" });
                    if (fld != null)
                    {
                        var v = IDispCall(fld, "Value", DISPATCH_PROPERTYGET, Array.Empty<object?>());
                        if (v != null) r.LogoRef = Convert.ToInt64(v);
                    }
                }
            }
            catch { }

            r.Basarili = true;
            return r;
        }
        finally
        {
            Release(inv);
        }
    }

    private string CollectInvoiceError(object inv)
    {
        try
        {
            var code = IDispCall(inv, "ErrorCode", DISPATCH_PROPERTYGET, Array.Empty<object?>());
            var desc = IDispCall(inv, "ErrorDesc", DISPATCH_PROPERTYGET, Array.Empty<object?>());
            var dbDesc = IDispCall(inv, "DBErrorDesc", DISPATCH_PROPERTYGET, Array.Empty<object?>());
            return $"ErrorCode={code} ErrorDesc={desc} DBErrorDesc={dbDesc}";
        }
        catch { return "(ErrorCode okunamadı)"; }
    }

    // ============================================================================
    // STOK FİŞİ AKTARIMI (doMaterialSlip = 1)
    // ----------------------------------------------------------------------------
    // Üretimden Giriş Fişi (TYPE=13) ve Promosyon/Çıkış Fişi (TYPE=22) için.
    // KeyNet decompile enum'undan: doMaterial=0, doMaterialSlip=1, ...
    // XML referansı: <MATERIAL_SLIPS / SLIP / TRANSACTIONS>
    //
    // Fatura aktarımına göre çok daha basit:
    //   - Cari (ARP_CODE) yok
    //   - Fiyat / Tutar / KDV yok (stok hareketi, maliyet Logo'da otomatik)
    //   - Döviz hesapları yok (sadece TL)
    //   - GRPCODE=3 (XML'de görüldü), POST_FLAGS yok (default OK)
    // ============================================================================

    private const int DO_MATERIAL_SLIP = 1;

    public Task<StokFisiAktarimSonuc> StokFisiAktarAsync(
        LogoUnityCredentials cred,
        StokFisiModel model,
        CancellationToken ct = default)
    {
        var tcs = new TaskCompletionSource<StokFisiAktarimSonuc>(TaskCreationOptions.RunContinuationsAsynchronously);
        var th = new Thread(() =>
        {
            try
            {
                var sonuc = ExecuteStokFisiOnStaThread(cred, model, ct);
                tcs.TrySetResult(sonuc);
            }
            catch (Exception ex)
            {
                _log.LogError(ex, "Logo Unity StokFisi hata");
                tcs.TrySetException(ex);
            }
        });
        th.SetApartmentState(ApartmentState.STA);
        th.IsBackground = true;
        th.Name = "LogoUnityStaThread-StokFisi";
        th.Start();
        return tcs.Task;
    }

    private StokFisiAktarimSonuc ExecuteStokFisiOnStaThread(
        LogoUnityCredentials cred,
        StokFisiModel model,
        CancellationToken ct)
    {
        var sonuc = new StokFisiAktarimSonuc();
        object? app = null;
        bool loggedIn = false;
        var diag = new System.Text.StringBuilder();

        try
        {
            app = CreateUnityApp();

            bool ok;
            try { ok = (bool)(Inv(app, "Login", cred.Kullanici, cred.Sifre, cred.FirmaNo) ?? false); }
            catch (Exception ex) { sonuc.LoginHata = $"Login çağrısı başarısız: {ex.Message}"; return sonuc; }

            if (!ok)
            {
                sonuc.LoginHata = $"Logo Login reddetti. {GetStr(app, "GetLastError")} - {GetStr(app, "GetLastErrorString")}";
                return sonuc;
            }
            loggedIn = true;
            sonuc.LoginBasarili = true;

            ct.ThrowIfCancellationRequested();

            object? slip = null;
            try
            {
                slip = Inv(app, "NewDataObject", DO_MATERIAL_SLIP);
                if (slip == null)
                {
                    sonuc.Hata = $"NewDataObject({DO_MATERIAL_SLIP}=doMaterialSlip) null döndü.";
                    return sonuc;
                }

                try { Inv(slip, "New"); }
                catch (Exception ex) { diag.AppendLine($"New(): {ex.Message}"); }

                // Tüm header alanlarını TrySet ile sar — bir tanesi unknown olursa diğerleri devam etsin,
                // hata sebebi diag log'a yansısın.
                void TrySetH(string name, object value)
                {
                    try { SetHeaderField(slip, name, value); diag.AppendLine($"H:{name}=OK ({value})"); }
                    catch (Exception ex) { diag.AppendLine($"H:{name}=FAIL ({ex.Message.Trim()})"); }
                }

                // ---- HEADER ----
                // XML tag adları ve DB kolonu farklı olabiliyor. Logo Unity COM bazen XML tag
                // bazen DB kolon ismi bekler. İkisini de deneyelim.
                TrySetH("GRPCODE",          3);   // GROUP yerine DB kolonu
                TrySetH("GROUP",            3);   // XML alternatifi
                TrySetH("TRCODE",           model.Tip);   // TYPE yerine DB kolonu
                TrySetH("TYPE",             model.Tip);   // XML alternatifi

                if (!string.IsNullOrWhiteSpace(model.BelgeNo))
                {
                    TrySetH("FICHENO", model.BelgeNo);    // DB kolonu
                    TrySetH("NUMBER",  model.BelgeNo);    // XML alternatifi
                }

                TrySetH("DATE_",            model.Tarih);   // DB kolonu (sondaki underscore)
                TrySetH("DATE",             model.Tarih);   // XML alternatifi
                TrySetH("SOURCEINDEX",      model.Ambar);   // DB kolonu (STFICHE.SOURCEINDEX)
                TrySetH("SOURCE_WH",        model.Ambar);   // XML alternatifi
                TrySetH("SOURCEWHOUSE",     model.Ambar);   // Bir başka olası ad
                TrySetH("SOURCE_FACTORY_NR", model.Fabrika);

                int LogoTimeEncode(int h, int mi, int s) => h * 16777216 + mi * 65536 + s * 256;
                int headerTime = LogoTimeEncode(12, 0, 0);
                TrySetH("FTIME", headerTime);     // DB
                TrySetH("TIME",  headerTime);     // XML

                if (!string.IsNullOrWhiteSpace(model.Aciklama1)) TrySetH("GENEXP1", model.Aciklama1);
                if (!string.IsNullOrWhiteSpace(model.Aciklama2)) TrySetH("GENEXP2", model.Aciklama2);

                TrySetH("EBOOK_DOCTYPE", 99);
                TrySetH("CURRSEL_TOTALS", 1);

                // ---- TRANSACTIONS (satırlar) ----
                // Mevcut fatura aktarımındaki pattern: txField.Lines.AppendLine() + Item(idx) + SetField
                object? dfRaw = GetProp(slip, "DataFields");
                if (dfRaw == null)
                    throw new InvalidOperationException("DataFields null döndü.");

                object? txField = null;
                try
                {
                    int txIdx = Convert.ToInt32(Inv(dfRaw, "GetFieldIndex", "TRANSACTIONS") ?? -1);
                    if (txIdx >= 0)
                        txField = IDispCall(dfRaw, "Item", DISPATCH_PROPERTYGET, new object?[] { txIdx });
                }
                catch (Exception ex) { diag.AppendLine($"GetFieldIndex(TRANSACTIONS): {ex.Message}"); }

                if (txField == null)
                {
                    try { txField = Inv(dfRaw, "FieldByName", "TRANSACTIONS"); }
                    catch (Exception ex) { diag.AppendLine($"FieldByName(TRANSACTIONS): {ex.Message}"); }
                }

                if (txField == null)
                    throw new InvalidOperationException("TRANSACTIONS alanı bulunamadı.");
                diag.AppendLine("TRANSACTIONS field=OK");

                object? txLines = GetProp(txField, "Lines");
                if (txLines == null)
                    throw new InvalidOperationException("TRANSACTIONS.Lines null.");
                diag.AppendLine("TRANSACTIONS.Lines=OK");

                int sira = 0;
                foreach (var sat in model.Satirlar)
                {
                    ct.ThrowIfCancellationRequested();
                    sira++;

                    // VBA: transactions_lines.AppendLine
                    try { Inv(txLines, "AppendLine"); }
                    catch (Exception ex) { diag.AppendLine($"L{sira}:AppendLine FAIL ({ex.Message})"); throw; }

                    int idx = Convert.ToInt32(GetProp(txLines, "Count") ?? 1) - 1;
                    object? line = GetLine(txLines, idx);
                    if (line == null)
                        throw new InvalidOperationException($"Satır {sira}: TRANSACTIONS.Lines[{idx}] null.");
                    diag.AppendLine($"L{sira}: line index={idx} OK");

                    void TrySetS(string name, object value)
                    {
                        try { SetField(line, name, value); diag.AppendLine($"L{sira}:{name}=OK ({value})"); }
                        catch (Exception ex) { diag.AppendLine($"L{sira}:{name}=FAIL ({ex.Message.Trim()})"); }
                    }

                    TrySetS("LINE_TYPE",     0);
                    TrySetS("ITEM_CODE",     sat.MalzemeKodu);
                    TrySetS("SOURCEINDEX",   model.Ambar);
                    TrySetS("FACTORYNR",     model.Fabrika);
                    TrySetS("LINE_NUMBER",   sira);
                    TrySetS("QUANTITY",      (double)sat.Miktar);
                    if (!string.IsNullOrWhiteSpace(sat.Birim))
                        TrySetS("UNIT_CODE", sat.Birim);
                    if (!string.IsNullOrWhiteSpace(sat.Aciklama))
                        TrySetS("LINEEXP", sat.Aciklama);
                    TrySetS("EU_VAT_STATUS", 4);
                    TrySetS("EDT_CURR",      1);
                    TrySetS("UNIT_CONV1",    1);
                    TrySetS("UNIT_CONV2",    1);
                }

                // ---- POST ----
                bool posted;
                try { posted = (bool)(Inv(slip, "Post") ?? false); }
                catch (Exception ex)
                {
                    diag.AppendLine($"Post() exception: {ex.Message}");
                    sonuc.Hata = $"Post hata: {ex.Message}";
                    sonuc.DiagLog = diag.ToString();
                    return sonuc;
                }

                if (!posted)
                {
                    var errInfo = CollectInvoiceError(slip);
                    diag.AppendLine($"Post() false döndü. {errInfo}");
                    sonuc.Hata = $"Post() reddedildi. {errInfo}";
                    sonuc.DiagLog = diag.ToString();
                    return sonuc;
                }

                // Logo'nun verdiği yeni fiş numarasını oku
                try
                {
                    object? dfPost = GetProp(slip, "DataFields");
                    if (dfPost != null)
                    {
                        object? noField = FieldByName(slip, dfPost, "NUMBER");
                        if (noField != null)
                        {
                            object? noVal = GetProp(noField, "Value");
                            sonuc.LogoFisNo = noVal?.ToString();
                        }
                    }
                }
                catch (Exception ex) { diag.AppendLine($"NUMBER oku: {ex.Message}"); }

                sonuc.Basarili = true;
                sonuc.DiagLog  = diag.ToString();
                YazDosyayaDiag(model, sonuc);
                return sonuc;
            }
            finally
            {
                Release(slip);
            }
        }
        catch (OperationCanceledException) { throw; }
        catch (Exception ex)
        {
            _log.LogError(ex, "StokFisi STA hata");
            sonuc.Hata = ex.Message;
            sonuc.DiagLog = diag.ToString();
            YazDosyayaDiag(model, sonuc);
            return sonuc;
        }
        finally
        {
            if (loggedIn && app != null)
            {
                try { Inv(app, "Logout"); } catch { }
            }
            Release(app);
        }
    }

    private static void YazDosyayaDiag(StokFisiModel model, StokFisiAktarimSonuc sonuc)
    {
        try
        {
            var logFile = System.IO.Path.Combine(
                Services.AppDataPaths.DataRoot,
                "stok_fisi_diag.txt");
            var sb = new System.Text.StringBuilder();
            sb.AppendLine($"==== {DateTime.Now:yyyy-MM-dd HH:mm:ss} StokFisi Aktarım ====");
            sb.AppendLine($"Tip={model.Tip}  Tarih={model.Tarih:yyyy-MM-dd}  Ambar={model.Ambar}  Fabrika={model.Fabrika}");
            sb.AppendLine($"Satır={model.Satirlar.Count}  BelgeNo={model.BelgeNo}");
            for (int i = 0; i < model.Satirlar.Count; i++)
            {
                var s = model.Satirlar[i];
                sb.AppendLine($"  Satır#{i + 1}: {s.MalzemeKodu}  Miktar={s.Miktar} {s.Birim}");
            }
            sb.AppendLine($"LoginBasarili={sonuc.LoginBasarili}  LoginHata={sonuc.LoginHata}");
            sb.AppendLine($"Basarili={sonuc.Basarili}  Hata={sonuc.Hata}  LogoFisNo={sonuc.LogoFisNo}");
            sb.AppendLine("---- DIAG LOG ----");
            sb.AppendLine(sonuc.DiagLog);
            sb.AppendLine();
            System.IO.File.AppendAllText(logFile, sb.ToString());
        }
        catch { /* yazma hatası uygulamayı düşürmesin */ }
    }

    // FieldByName helper — DataFields koleksiyonundan alan ismi ile alan al
    private object? FieldByName(object owner, object dataFields, string name)
    {
        try { return Inv(dataFields, "FieldByName", name); }
        catch { }
        try
        {
            int idx = Convert.ToInt32(Inv(dataFields, "GetFieldIndex", name) ?? -1);
            if (idx >= 0)
            {
                object? item = IDispCall(dataFields, "Item", DISPATCH_PROPERTYGET, new object?[] { idx });
                if (item != null) return item;
            }
        }
        catch { }
        return null;
    }

    // Satır alanı set — önce SetHeaderField benzeri, başarısız olursa Item[].Value
    private void SetSatirField(object satirRow, object satirFields, string name, object value)
    {
        try { SetHeaderField(satirRow, name, value); return; }
        catch { }
        try
        {
            object? f = FieldByName(satirRow, satirFields, name);
            if (f != null) SetProp(f, "Value", value);
        }
        catch (Exception ex) { _log.LogDebug(ex, "SetSatirField({Name}) fallback başarısız", name); }
    }

    // ============================================================================
    // SATIŞ SİPARİŞİ (doSalesOrderSlip = 3) — Excel → Logo aktarımı
    //
    //   Enum sırası (KeyNet PMHS.dll v16.02.2026 decompile'ından):
    //     0=doMaterial, 1=doMaterialSlip, 2=doPurchService, 3=doSalesOrderSlip ← BU
    //
    //   Akış:
    //     1) NewDataObject(3) → slip; slip.New()
    //     2) Header alanları (TYPE/ARP_CODE/SHIPLOC_CODE/GL_CODE/DATE/SOURCE_WH/
    //        SALESMAN_CODE/PAYMENT_CODE/ORDER_STATUS/DOC_NUMBER vb.)
    //     3) Her satır: TRANSACTIONS.Lines.AppendLine() + MASTER_CODE/QUANTITY/
    //        UNIT_CODE/DUE_DATE + GetStockLinePrice(8, out price)  ← satış fiyat
    //        listesinden fiyatı Logo otomatik çeker
    //     4) slip.ApplyCampaign() — firmanın tanımlı kampanya/indirim oranlarını
    //        uygular (sağ tık → "Kampanya Uygula" ile aynı işi yapar)
    //     5) slip.Post()
    // ============================================================================

    private const int DO_SALES_ORDER = 3;

    /// <summary>
    /// Excel'deki ödeme türü metnini Logo PAYMENT_CODE'a çevirir. Sheet'e göre PEŞİN farklı:
    ///   CİPS   → "PESIN CIPS"
    ///   İÇECEK → "PESIN SC"
    /// "Vadeli" her ikisinde de → "40"
    /// Boş/tanımsız → "" (set edilmez)
    /// Büyük/küçük harf + Türkçe karakter (Ş/İ/Ç vs.) duyarsız.
    /// </summary>
    private static string OdemePlaniCozumle(string odemeMetni, string sheet)
    {
        if (string.IsNullOrWhiteSpace(odemeMetni)) return "";
        var n = odemeMetni.Trim().ToUpperInvariant()
                          .Replace("Ş", "S").Replace("İ", "I")
                          .Replace("Ç", "C").Replace("Ö", "O").Replace("Ü", "U").Replace("Ğ", "G");
        if (n == "VADELI" || n.StartsWith("VAD")) return "40";
        if (n == "PESIN" || n == "P" || n.StartsWith("PES"))
        {
            var tip = SatisSiparisExcelService.SheetTipi(sheet);
            return tip == "ICECEK" ? "PESIN SC" : "PESIN CIPS";
        }
        return "";
    }

    /// <summary>
    /// Sheet'e göre Logo Ticari İşlem Grubu kodu:
    ///   CİPS   → "G"
    ///   İÇECEK → "SC"
    ///   diğer → "G" (varsayılan)
    /// </summary>
    private static string TradingGrpCozumle(string sheet)
    {
        var tip = SatisSiparisExcelService.SheetTipi(sheet);
        if (tip == "ICECEK") return "SC";
        return "G";
    }

    public Task<SatisSiparisAktarimSonuc> SatisSiparisAktarAsync(
        LogoUnityCredentials cred,
        SatisSiparisFis fis,
        CancellationToken ct = default)
    {
        var tcs = new TaskCompletionSource<SatisSiparisAktarimSonuc>(TaskCreationOptions.RunContinuationsAsynchronously);
        var th = new Thread(() =>
        {
            try
            {
                var sonuc = ExecuteSatisSiparisOnStaThread(cred, fis, ct);
                tcs.TrySetResult(sonuc);
            }
            catch (Exception ex)
            {
                _log.LogError(ex, "Logo Unity SatışSiparişi hata");
                tcs.TrySetException(ex);
            }
        });
        th.SetApartmentState(ApartmentState.STA);
        th.IsBackground = true;
        th.Name = "LogoUnityStaThread-SatisSiparis";
        th.Start();
        return tcs.Task;
    }

    /// <summary>
    /// Tek seferde birden çok fişi aynı Login oturumunda gönderir. Login'i tekrar tekrar
    /// kurmak yerine bir kez bağlanıp tüm fişler için sırayla NewDataObject + Post yapar.
    /// </summary>
    public Task<List<SatisSiparisAktarimSonuc>> SatisSiparisBatchAktarAsync(
        LogoUnityCredentials cred,
        IReadOnlyList<SatisSiparisFis> fisler,
        CancellationToken ct = default)
    {
        var tcs = new TaskCompletionSource<List<SatisSiparisAktarimSonuc>>(TaskCreationOptions.RunContinuationsAsynchronously);
        var th = new Thread(() =>
        {
            try
            {
                var sonuclar = ExecuteSatisSiparisBatchOnStaThread(cred, fisler, ct);
                tcs.TrySetResult(sonuclar);
            }
            catch (Exception ex)
            {
                _log.LogError(ex, "Logo Unity SatışSiparişi batch hata");
                tcs.TrySetException(ex);
            }
        });
        th.SetApartmentState(ApartmentState.STA);
        th.IsBackground = true;
        th.Name = "LogoUnityStaThread-SatisSiparis-Batch";
        th.Start();
        return tcs.Task;
    }

    private SatisSiparisAktarimSonuc ExecuteSatisSiparisOnStaThread(
        LogoUnityCredentials cred, SatisSiparisFis fis, CancellationToken ct)
    {
        object? app = null;
        bool loggedIn = false;
        try
        {
            app = CreateUnityApp();
            bool ok;
            try { ok = (bool)(Inv(app, "Login", cred.Kullanici, cred.Sifre, cred.FirmaNo) ?? false); }
            catch (Exception ex) { return new SatisSiparisAktarimSonuc { Hata = "Login: " + ex.Message }; }
            if (!ok)
                return new SatisSiparisAktarimSonuc {
                    Hata = $"Login reddedildi: {GetStr(app, "GetLastError")} - {GetStr(app, "GetLastErrorString")}" };
            loggedIn = true;
            return BirFisGonder(app, fis, ct);
        }
        finally
        {
            if (loggedIn && app != null) { try { Inv(app, "Logout"); } catch { } }
            Release(app);
        }
    }

    private List<SatisSiparisAktarimSonuc> ExecuteSatisSiparisBatchOnStaThread(
        LogoUnityCredentials cred, IReadOnlyList<SatisSiparisFis> fisler, CancellationToken ct)
    {
        var sonuclar = new List<SatisSiparisAktarimSonuc>(fisler.Count);
        object? app = null;
        bool loggedIn = false;
        try
        {
            app = CreateUnityApp();
            bool ok;
            try { ok = (bool)(Inv(app, "Login", cred.Kullanici, cred.Sifre, cred.FirmaNo) ?? false); }
            catch (Exception ex)
            {
                foreach (var _ in fisler) sonuclar.Add(new SatisSiparisAktarimSonuc { Hata = "Login: " + ex.Message });
                return sonuclar;
            }
            if (!ok)
            {
                var err = $"Login reddedildi: {GetStr(app, "GetLastError")} - {GetStr(app, "GetLastErrorString")}";
                foreach (var _ in fisler) sonuclar.Add(new SatisSiparisAktarimSonuc { Hata = err });
                return sonuclar;
            }
            loggedIn = true;

            foreach (var fis in fisler)
            {
                ct.ThrowIfCancellationRequested();

                // İlk deneme: Sevkedilebilir (ORDER_STATUS=4)
                var sonuc = BirFisGonder(app, fis, ct, orderStatus: 4);

                // Risk aşılmışsa → fişi YENİ slip ile Öneri (ORDER_STATUS=1) olarak baştan dene.
                // Logo'da reddedilmiş slip üzerinde STATUS değişikliği etkili olmuyor; tertemiz yeni nesne lazım.
                if (!sonuc.Basarili && IsRiskAsildiHatasi(sonuc.Hata))
                {
                    var oneri = BirFisGonder(app, fis, ct, orderStatus: 1);
                    if (oneri.Basarili)
                    {
                        if (!string.IsNullOrWhiteSpace(oneri.LogoFisNo))
                            oneri.LogoFisNo = "Ö " + oneri.LogoFisNo;
                        sonuc = oneri;
                    }
                    else
                    {
                        sonuc.Hata = "Risk aşıldı — Öneri retry de başarısız: " + oneri.Hata;
                        sonuc.DiagLog = (sonuc.DiagLog ?? "") + Environment.NewLine
                                      + "--- ÖNERİ RETRY ---" + Environment.NewLine + (oneri.DiagLog ?? "");
                    }
                }

                sonuclar.Add(sonuc);
            }
            return sonuclar;
        }
        finally
        {
            if (loggedIn && app != null) { try { Inv(app, "Logout"); } catch { } }
            Release(app);
        }
    }

    /// <summary>
    /// Logo Post() reddediş mesajından "Cari riski aşılmıştır" hatasını yakalar.
    /// Logo hata kodu: 32436. Mesaj farklı dillerde olabilir — "risk" + "aş" anahtarları + 32436 kontrolü.
    /// </summary>
    private static bool IsRiskAsildiHatasi(string? err)
    {
        if (string.IsNullOrWhiteSpace(err)) return false;
        var n = err.ToUpperInvariant()
                   .Replace("Ş", "S").Replace("İ", "I")
                   .Replace("Ç", "C").Replace("Ö", "O").Replace("Ü", "U").Replace("Ğ", "G");
        if (n.Contains("32436")) return true;
        if (n.Contains("RISK") && (n.Contains("ASIL") || n.Contains("ASMI"))) return true;
        return false;
    }

    private SatisSiparisAktarimSonuc BirFisGonder(object app, SatisSiparisFis fis, CancellationToken ct, int orderStatus = 4)
    {
        var sonuc = new SatisSiparisAktarimSonuc { CariKodu = fis.CariKodu, AracNo = fis.AracNo };
        var diag = new System.Text.StringBuilder();
        object? slip = null;

        // Fabrika çıkışına göre İŞ YERİ / BÖLÜM / FABRİKA / AMBAR seç
        var fc = FabrikaCikisAyari.Coz(fis.FabrikaCikis);
        if (fc == null)
        {
            sonuc.Hata = $"Bilinmeyen Fabrika Çıkışı: '{fis.FabrikaCikis}' — desteklenen: "
                       + string.Join(", ", FabrikaCikisAyari.TanimliAdlar());
            return sonuc;
        }

        try
        {
            slip = Inv(app, "NewDataObject", DO_SALES_ORDER);
            if (slip == null) { sonuc.Hata = $"NewDataObject({DO_SALES_ORDER}=doSalesOrderSlip) null."; return sonuc; }
            try { Inv(slip, "New"); } catch (Exception ex) { diag.AppendLine($"New(): {ex.Message}"); }

            void TrySetH(string name, object value)
            {
                try { SetHeaderField(slip, name, value); diag.AppendLine($"H:{name}=OK ({value})"); }
                catch (Exception ex) { diag.AppendLine($"H:{name}=FAIL ({ex.Message.Trim()})"); }
            }

            int LogoTimeEncode(int h, int mi, int s) => h * 16777216 + mi * 65536 + s * 256;

            // ── HEADER ──
            TrySetH("TYPE",            7);              // 7 = öneri/sevkedilebilir sınıfı (sales order TYPE)
            TrySetH("TRCODE",          1);              // TRCODE 1 = standart sipariş
            // Sipariş tarihi = aktarımın yapıldığı GÜN (Excel'deki "Sevk Tarihi" DEĞİL).
            // Sevk/teslim tarihi satırların DUE_DATE'inde fis.TeslimTarihi olarak gider.
            var siparisTarihi = DateTime.Today;
            TrySetH("DATE",            siparisTarihi);
            TrySetH("DATE_",           siparisTarihi);
            TrySetH("TIME",            LogoTimeEncode(10, 0, 0));
            TrySetH("FTIME",           LogoTimeEncode(10, 0, 0));
            if (!string.IsNullOrWhiteSpace(fis.AracNo))
            {
                TrySetH("DOC_NUMBER", fis.AracNo);
                TrySetH("DOCODE",     fis.AracNo);
            }
            TrySetH("ARP_CODE",        fis.CariKodu);
            TrySetH("CLIENTCODE",      fis.CariKodu);
            if (!string.IsNullOrWhiteSpace(fis.SevkiyatAdresi))
            {
                TrySetH("SHIPLOC_CODE", fis.SevkiyatAdresi);
                TrySetH("SHIPADDR",     fis.SevkiyatAdresi);
            }
            if (!string.IsNullOrWhiteSpace(fis.MuhasebeHesabi))
            {
                TrySetH("GL_CODE",  fis.MuhasebeHesabi);
                TrySetH("ACCCODE",  fis.MuhasebeHesabi);
            }
            // ── İş Yeri / Bölüm / Fabrika / Ambar — fabrika çıkışına göre DİNAMİK ──
            // KeyNet decompile: Logo Unity COM doğru fabrika field adı "FACTORY" (NR yok!).
            // DIVISION = İŞ YERİ (BÖLÜM değil — masaüstüde keşfedildi).
            TrySetH("BRANCH",          fc.Branch);
            TrySetH("DIVISION",        fc.Branch);
            TrySetH("BRANCHNR",        fc.Branch);
            TrySetH("DEPARTMENT",      fc.Department);
            TrySetH("DEPNR",           fc.Department);
            TrySetH("FACTORY",         fc.FactoryNr);
            TrySetH("FACTORYNR",       fc.FactoryNr);
            TrySetH("FACTORY_NR",      fc.FactoryNr);
            TrySetH("SOURCE_WH",       fc.SourceWh);
            TrySetH("SOURCEINDEX",     fc.SourceWh);
            // 4 = Sevkedilebilir (S), 1 = Öneri (Ö). Default 4; risk aşımı retry'ında 1.
            TrySetH("ORDER_STATUS",    orderStatus);
            TrySetH("STATUS",          orderStatus);
            // Ticari İşlem Grubu — sheet'e göre: CİPS→"G", İÇECEK→"SC"
            TrySetH("TRADING_GRP",     TradingGrpCozumle(fis.Sheet));
            TrySetH("SALESMAN_CODE",   fis.SatisElemani);

            // Ödeme planı kuralı (öncelik):
            //   1) Cari kartında plan VAR → PAYMENT_CODE = cari planı kodu
            //      (Logo Unity COM cari default'unu otomatik uygulamıyor — elle yazıyoruz)
            //   2) Yoksa Excel ödeme türü, sheet'e göre:
            //        CİPS:   "Vadeli"→"40", "Peşin/P"→"PESIN CIPS"
            //        İÇECEK: "Vadeli"→"40", "Peşin/P"→"PESIN SC"
            string odemeKodu;
            if (!string.IsNullOrWhiteSpace(fis.CariOdemePlanKodu))
            {
                odemeKodu = fis.CariOdemePlanKodu.Trim();
                diag.AppendLine($"PAYMENT_CODE: cari planı = '{odemeKodu}'");
            }
            else
            {
                odemeKodu = OdemePlaniCozumle(fis.OdemeTuru, fis.Sheet);
                if (!string.IsNullOrEmpty(odemeKodu))
                    diag.AppendLine($"PAYMENT_CODE: Excel kuralı '{fis.OdemeTuru}' (sheet={fis.Sheet}) → '{odemeKodu}'");
            }
            if (!string.IsNullOrEmpty(odemeKodu))
            {
                TrySetH("PAYMENT_CODE", odemeKodu);
                TrySetH("PAYDEFCODE",   odemeKodu);
            }
            TrySetH("CURRSEL_TOTAL",   1);

            // ── TRANSACTIONS (satırlar) ──
            object? dfRaw = GetProp(slip, "DataFields") ?? throw new InvalidOperationException("DataFields null");
            object? txField = null;
            try
            {
                int txIdx = Convert.ToInt32(Inv(dfRaw, "GetFieldIndex", "TRANSACTIONS") ?? -1);
                if (txIdx >= 0) txField = IDispCall(dfRaw, "Item", DISPATCH_PROPERTYGET, new object?[] { txIdx });
            }
            catch { }
            if (txField == null) try { txField = Inv(dfRaw, "FieldByName", "TRANSACTIONS"); } catch { }
            if (txField == null) throw new InvalidOperationException("TRANSACTIONS alanı yok.");
            object? txLines = GetProp(txField, "Lines") ?? throw new InvalidOperationException("TRANSACTIONS.Lines null.");

            int sira = 0;
            foreach (var sat in fis.Satirlar)
            {
                ct.ThrowIfCancellationRequested();
                sira++;
                try { Inv(txLines, "AppendLine"); }
                catch (Exception ex) { diag.AppendLine($"L{sira}:AppendLine FAIL ({ex.Message})"); throw; }

                int idx = Convert.ToInt32(GetProp(txLines, "Count") ?? 1) - 1;
                object? line = GetLine(txLines, idx) ?? throw new InvalidOperationException($"Satır {sira} null.");

                void TrySetS(string name, object value)
                {
                    try { SetField(line, name, value); diag.AppendLine($"L{sira}:{name}=OK ({value})"); }
                    catch (Exception ex) { diag.AppendLine($"L{sira}:{name}=FAIL ({ex.Message.Trim()})"); }
                }

                TrySetS("TYPE",          0);
                TrySetS("MASTER_CODE",   sat.MalzemeKodu);
                TrySetS("ITEM_CODE",     sat.MalzemeKodu);
                TrySetS("QUANTITY",      (double)sat.KoliMiktari);
                TrySetS("AMOUNT",        (double)sat.KoliMiktari);
                TrySetS("UNIT_CODE",     "KL");
                TrySetS("UNIT_CONV1",    1);
                TrySetS("UNIT_CONV2",    1);
                TrySetS("DUE_DATE",      fis.TeslimTarihi);
                // Satır seviyesi org alanları — fabrika çıkışına göre DİNAMİK
                TrySetS("SOURCE_WH",     fc.SourceWh);
                TrySetS("SOURCEINDEX",   fc.SourceWh);
                TrySetS("DIVISION",      fc.Department);
                TrySetS("DEPARTMENT",    fc.Department);
                TrySetS("FACTORY",       fc.FactoryNr);
                TrySetS("FACTORYNR",     fc.FactoryNr);
                TrySetS("BRANCH",        fc.Branch);
                TrySetS("LINE_NUMBER",   sira);
                TrySetS("SALESMAN_CODE", fis.SatisElemani);
                TrySetS("AFFECT_RISK",   1);
                TrySetS("EDT_CURR",      1);

                // Birim fiyatı manuel set ediyoruz — PRCLIST'ten okuduğumuz "genel" fiyat
                // (cariye özel olmayan, en son tarihli). ApplyCampaign bu fiyat üzerinden
                // indirim/komisyon satırlarını üretir.
                if (sat.BirimFiyat > 0m)
                {
                    TrySetS("PRICE",      (double)sat.BirimFiyat);
                    TrySetS("PC_PRICE",   (double)sat.BirimFiyat);
                    TrySetS("EDT_PRICE",  (double)sat.BirimFiyat);
                    TrySetS("ORG_PRICE",  (double)sat.BirimFiyat);
                    TrySetS("TOTAL",      (double)(sat.BirimFiyat * sat.KoliMiktari));
                    TrySetS("TOTAL_NET",  (double)(sat.BirimFiyat * sat.KoliMiktari));
                }
                else
                {
                    diag.AppendLine($"L{sira}: BirimFiyat=0 (PRCLIST'te bulunamadı)");
                }
            }

            // ── Kampanyaları uygula ──
            try
            {
                var camp = Inv(slip, "ApplyCampaign");
                diag.AppendLine($"ApplyCampaign() ret={camp}");
            }
            catch (Exception ex) { diag.AppendLine($"ApplyCampaign FAIL ({ex.Message.Trim()})"); }

            // ── ORG ALANLARINI YENİDEN ZORLA (son söz bizim olsun) ──
            // Logo, ARP_CODE set edildikten sonra carinin default fabrikası/işyeri/ambarına göre
            // header'ı override edebiliyor (924974: FACTORYNR=26 set ettik, 0 yazıldı).
            // ApplyCampaign sonrası, Post'tan hemen önce bir kez daha set — Logo'nun override
            // mantığını atlayarak doğru değerler kalır.
            TrySetH("BRANCH",      fc.Branch);
            TrySetH("DIVISION",    fc.Branch);
            TrySetH("BRANCHNR",    fc.Branch);
            TrySetH("DEPARTMENT",  fc.Department);
            TrySetH("DEPNR",       fc.Department);
            TrySetH("FACTORY",     fc.FactoryNr);
            TrySetH("FACTORYNR",   fc.FactoryNr);
            TrySetH("FACTORY_NR",  fc.FactoryNr);
            TrySetH("SOURCE_WH",   fc.SourceWh);
            TrySetH("SOURCEINDEX", fc.SourceWh);

            // ── Post ──
            bool posted;
            try { posted = (bool)(Inv(slip, "Post") ?? false); }
            catch (Exception ex) { sonuc.Hata = "Post: " + ex.Message; sonuc.DiagLog = diag.ToString(); return sonuc; }

            if (!posted)
            {
                var err = CollectInvoiceError(slip);
                sonuc.Hata = "Post reddedildi. " + err;
                sonuc.DiagLog = diag.ToString();
                return sonuc;
            }

            // Yeni fiş numarasını oku
            try
            {
                object? dfPost = GetProp(slip, "DataFields");
                if (dfPost != null)
                {
                    object? noField = FieldByName(slip, dfPost, "NUMBER")
                                   ?? FieldByName(slip, dfPost, "FICHENO");
                    if (noField != null) sonuc.LogoFisNo = (GetProp(noField, "Value"))?.ToString();
                }
            }
            catch (Exception ex) { diag.AppendLine($"NUMBER oku: {ex.Message}"); }

            sonuc.Basarili = true;
            sonuc.DiagLog  = diag.ToString();
            return sonuc;
        }
        catch (Exception ex)
        {
            sonuc.Hata = ex.Message;
            sonuc.DiagLog = diag.ToString();
            return sonuc;
        }
        finally { Release(slip); }
    }

}

// ───────────────────────────────────────────────────────────────────────────────
// Satış Siparişi aktarım sonuç DTO'su
// ───────────────────────────────────────────────────────────────────────────────
public class SatisSiparisAktarimSonuc
{
    public string CariKodu     { get; set; } = "";
    public string AracNo       { get; set; } = "";
    public bool   Basarili     { get; set; }
    public string? Hata        { get; set; }
    public string? LogoFisNo   { get; set; }
    public string? DiagLog     { get; set; }
}
