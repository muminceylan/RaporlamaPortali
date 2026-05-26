const { Client, LocalAuth, MessageMedia } = require('whatsapp-web.js');
const qrcodeTerminal = require('qrcode-terminal');
const { exec } = require('child_process');
const path = require('path');
const fs = require('fs');
const puppeteer = require('puppeteer');
const http = require('http');

// =====================================================
// DOSYA YOLLARI
// Kod: __dirname (publish klasörü, her publish'te güncellenir)
// Veri: WHATSAPP_DATA_DIR env (publish'ten etkilenmeyen kalıcı yer)
// =====================================================

const DIZIN      = __dirname;
const VERI_DIZIN = process.env.WHATSAPP_DATA_DIR && process.env.WHATSAPP_DATA_DIR.trim().length > 0
    ? process.env.WHATSAPP_DATA_DIR
    : __dirname;

try { if (!fs.existsSync(VERI_DIZIN)) fs.mkdirSync(VERI_DIZIN, { recursive: true }); } catch (_) {}

const CONFIG_DOSYA  = path.join(VERI_DIZIN, 'whatsapp-config.json');
const DURUM_DOSYA   = path.join(VERI_DIZIN, 'whatsapp-status.json');
const LOG_DOSYA     = path.join(VERI_DIZIN, 'whatsapp-log.json');
const CIKTI_KLASORU = path.join(VERI_DIZIN, 'screenshots');

// =====================================================
// CONFIG OKUMA / DURUM YAZMA
// =====================================================

function configOku() {
    try {
        if (fs.existsSync(CONFIG_DOSYA)) {
            return JSON.parse(fs.readFileSync(CONFIG_DOSYA, 'utf8'));
        }
    } catch (e) {
        console.error('Config okuma hatasi:', e.message);
    }
    return {
        yetkiliNumaralar: [],
        tetikleyiciler: ['tüm rapor', 'tum rapor', 'tumrapor', 'tümrapor'],
        raporApiUrl: 'http://localhost:5050/api/rapor',
        excelDosyasi: ''
    };
}

function logYaz(numara, mesaj, sonuc) {
    try {
        let kayitlar = [];
        if (fs.existsSync(LOG_DOSYA)) {
            kayitlar = JSON.parse(fs.readFileSync(LOG_DOSYA, 'utf8'));
        }
        kayitlar.unshift({
            tarih: new Date().toISOString(),
            numara: numara,
            mesaj: mesaj,
            sonuc: sonuc
        });
        // Son 200 kaydı tut
        if (kayitlar.length > 200) kayitlar = kayitlar.slice(0, 200);
        fs.writeFileSync(LOG_DOSYA, JSON.stringify(kayitlar), 'utf8');
    } catch (e) {
        console.error('Log yazma hatasi:', e.message);
    }
}

let _sonDurum = '';
function durumYaz(durum, qrString) {
    try {
        _sonDurum = durum;
        fs.writeFileSync(DURUM_DOSYA, JSON.stringify({
            durum: durum,
            qrString: qrString || '',
            guncelleme: new Date().toISOString()
        }), 'utf8');
    } catch (e) {
        console.error('Durum yazma hatasi:', e.message);
    }
}

let config = configOku();

if (!fs.existsSync(CIKTI_KLASORU)) {
    fs.mkdirSync(CIKTI_KLASORU, { recursive: true });
}

// =====================================================
// ŞEKER RAPORU — KONUŞMA DURUMU
// Anahtar: numara → { adim: 'BASLANGIC' | 'BITIS', baslangic: 'YYYY-MM-DD' }
// =====================================================
const sekerKonusma = new Map();

// Türkçe ay adı → ay numarası
const AYLAR = {
    'ocak': 1, 'subat': 2, 'şubat': 2, 'mart': 3, 'nisan': 4,
    'mayis': 5, 'mayıs': 5, 'haziran': 6, 'temmuz': 7,
    'agustos': 8, 'ağustos': 8, 'eylul': 9, 'eylül': 9,
    'ekim': 10, 'kasim': 11, 'kasım': 11, 'aralik': 12, 'aralık': 12
};

/** "Eylül", "Ekim 2025" vb. → {baslangic, bitis} veya null */
function ayAdindenTarih(metin) {
    const temiz = metin.toLowerCase().replace(/[ığüşöçiI]/g, c =>
        ({'ı':'i','ğ':'g','ü':'u','ş':'s','ö':'o','ç':'c','i':'i','I':'i'}[c]||c));
    for (const [isim, no] of Object.entries(AYLAR)) {
        if (temiz.includes(isim)) {
            // Yıl var mı?
            const yilEsles = temiz.match(/\b(20\d{2})\b/);
            const yil = yilEsles ? parseInt(yilEsles[1]) : new Date().getFullYear();
            const son = new Date(yil, no, 0).getDate(); // ayın son günü
            return {
                baslangic: `${yil}-${String(no).padStart(2,'0')}-01`,
                bitis:     `${yil}-${String(no).padStart(2,'0')}-${String(son).padStart(2,'0')}`
            };
        }
    }
    return null;
}

/** "01.09.2025" veya "01/09/2025" veya "2025-09-01" → "YYYY-MM-DD" veya null */
function tarihParse(metin) {
    metin = metin.trim();
    // DD.MM.YYYY veya DD/MM/YYYY
    let m = metin.match(/^(\d{1,2})[./](\d{1,2})[./](\d{4})$/);
    if (m) return `${m[3]}-${m[2].padStart(2,'0')}-${m[1].padStart(2,'0')}`;
    // YYYY-MM-DD
    m = metin.match(/^(\d{4})-(\d{2})-(\d{2})$/);
    if (m) return metin;
    return null;
}

async function sekerRaporuGonder(message, baslangicStr, bitisStr, bulanik = false) {
    const baseUrl = (config.raporApiUrl || 'http://localhost:5050/api/rapor')
        .replace(/\/api\/.*$/, '');
    const bulanikParam = bulanik ? '&bulanik=true' : '';
    const urlAnaliz = `${baseUrl}/api/seker-analiz?baslangic=${baslangicStr}&bitis=${bitisStr}${bulanikParam}`;
    const urlRapor  = `${baseUrl}/api/seker-raporu?baslangic=${baslangicStr}&bitis=${bitisStr}${bulanikParam}`;

    console.log(`[Seker] Analiz tablosu: ${urlAnaliz}`);
    console.log(`[Seker] Baskanlık tablosu: ${urlRapor}`);

    const [htmlAnaliz, htmlRapor] = await Promise.all([
        sqlRaporuGetirRetry(urlAnaliz),
        sqlRaporuGetirRetry(urlRapor)
    ]);

    if (!htmlAnaliz && !htmlRapor) {
        await message.reply('Seker raporu alinamadi, sunucu kapali olabilir.');
        return;
    }

    // 1. resim: Ham analiz tablosu (üst tablo) — geniş viewport
    if (htmlAnaliz) {
        await htmldenPngOlusturVeGonder(message, htmlAnaliz, 'seker-analiz', 1620);
    }
    // 2. resim: Başkanlık tablosu (alt tablo)
    if (htmlRapor) {
        await htmldenPngOlusturVeGonder(message, htmlRapor, 'seker');
    }
}

// =====================================================
// WHATSAPP CLIENT
// =====================================================

const client = new Client({
    authStrategy: new LocalAuth({ dataPath: path.join(VERI_DIZIN, '.wwebjs_auth') }),
    puppeteer: {
        headless: true,
        args: [
            '--no-sandbox',
            '--disable-setuid-sandbox',
            // Kurumsal SSL inspection / MITM proxy ortamlarinda Chromium'un
            // dis HTTPS bağlantısı kurabilmesi icin sertifika kontrolünü atla
            '--ignore-certificate-errors',
            // Headless tab'ın throttle edilmesini engelle — mesaj/ready gecikmesini önler
            '--disable-background-timer-throttling',
            '--disable-backgrounding-occluded-windows',
            '--disable-renderer-backgrounding',
            '--disable-features=IsolateOrigins,site-per-process,CalculateNativeWinOcclusion'
        ]
    },
    takeoverOnConflict: true,
    takeoverTimeoutMs: 10000,
    qrMaxRetries: 5
});

// Bağlantı takılırsa kendimizi öldür, .NET tarafı yeniden başlatsın
let _hazirTimer = null;
let _qrOkundu = false;
function hazirTimerKur(saniye, sebep) {
    if (_hazirTimer) clearTimeout(_hazirTimer);
    _hazirTimer = setTimeout(() => {
        console.error(`[WhatsApp] ${saniye}sn icinde bagli olunamadi (${sebep}). Bot yeniden baslatiliyor...`);
        durumYaz('BAGLANIYOR', '');
        try { client.destroy(); } catch (_) {}
        process.exit(1);
    }, saniye * 1000);
}

client.on('qr', (qr) => {
    console.log('\n[WhatsApp] QR kodu bekleniyor...');
    qrcodeTerminal.generate(qr, { small: true });
    // Ham QR string'ini yaz, .NET tarafı image'a çevirir
    durumYaz('QR_BEKLIYOR', qr);
    _qrOkundu = true;
    // QR gösterildikten sonra 180 sn icinde 'ready' gelmezse yeniden baslat
    hazirTimerKur(180, 'QR okutma sonrasi authentication tamamlanmadi');
});

client.on('loading_screen', (percent, message) => {
    console.log(`[WhatsApp] Yukleniyor: ${percent}% ${message || ''}`);
    // BAGLI iken sync amaçlı loading_screen tetiklenebiliyor — durumu geri düşürme
    if (_sonDurum !== 'BAGLI') durumYaz('BAGLANIYOR', '');
});

client.on('authenticated', () => {
    console.log('[WhatsApp] Kimlik dogrulandi, ready bekleniyor...');
    // BAGLI iken authenticated yeniden firing edebilir (re-auth/sync).
    // Bu durumda durumu geri düşürme ve "kendini öldür" timerini kurma — yoksa döngüye girer.
    if (_sonDurum !== 'BAGLI') {
        durumYaz('BAGLANIYOR', '');
        hazirTimerKur(90, 'authenticated sonrasi ready gelmedi');
    } else {
        console.log('[WhatsApp] Zaten BAGLI iken authenticated — yok sayildi.');
    }
});

let _tumSistemHazir = false;

async function sistemWarmupYap() {
    try {
        console.log('[Warmup] Puppeteer on-baslatma...');
        await browserGetir();
        console.log('[Warmup] API ping...');
        const baseUrl = (config.raporApiUrl || 'http://localhost:5050/api/rapor').replace(/\/api\/.*$/, '');
        await new Promise((resolve) => {
            const req = http.get(baseUrl + '/', { timeout: 30000 }, (res) => {
                res.on('data', () => {});
                res.on('end', () => resolve());
            });
            req.on('error', () => resolve());
            req.on('timeout', () => { req.destroy(); resolve(); });
        });
        _tumSistemHazir = true;
        console.log('[Warmup] Sistem tamamen hazir.');
    } catch (err) {
        console.error('[Warmup] Hata:', err.message);
        _tumSistemHazir = true; // yine de kabul et, yoksa kilit kalır
    }
}

async function bagliOlarakIsaretle(kaynak) {
    if (_sonDurum === 'BAGLI') return;
    if (_hazirTimer) { clearTimeout(_hazirTimer); _hazirTimer = null; }
    config = configOku();
    console.log(`[WhatsApp] Baglandi! (kaynak: ${kaynak}) — warmup baslatiliyor...`);
    // Önce warmup'ı bitir, sonra BAGLI yaz — ilk tetikleyicide browser takılı kalmasın
    if (!_tumSistemHazir) {
        try { await sistemWarmupYap(); } catch (_) {}
    }
    console.log('[WhatsApp] Warmup tamam, BAGLI olarak isaretleniyor.');
    durumYaz('BAGLI', '');
}

client.on('ready', () => { bagliOlarakIsaretle('ready'); });

// LocalAuth ile session restore'da 'ready' bazen atlanır; change_state -> CONNECTED yedek tetikleyici
client.on('change_state', (state) => {
    console.log('[WhatsApp] state:', state);
    if (state === 'CONNECTED') bagliOlarakIsaretle('change_state');
});

// Periyodik durum sorgusu — ready/change_state hiç gelmezse 10sn'de bir kontrol et
setInterval(async () => {
    if (_sonDurum === 'BAGLI') return;
    try {
        const state = await client.getState();
        if (state === 'CONNECTED') bagliOlarakIsaretle('poll');
    } catch (_) { /* henüz hazır değil */ }
}, 10000);

// HEARTBEAT — BAGLI iken 20sn'de bir status dosyasini tazele.
// .NET watchdog'i bu zaman damgasina bakarak "kilitli BAGLI" durumlarini tespit edip restart eder.
setInterval(async () => {
    if (_sonDurum !== 'BAGLI') return;
    try {
        const state = await client.getState();
        if (state === 'CONNECTED') {
            durumYaz('BAGLI', '');
        } else {
            console.warn('[Heartbeat] state=', state, '— BAGLI degil, durum guncellemesi atlandi');
        }
    } catch (e) {
        console.warn('[Heartbeat] getState hatasi:', e.message);
    }
}, 20000);

client.on('disconnected', (reason) => {
    console.log('[WhatsApp] Baglanti kesildi:', reason);
    durumYaz('BAGLI_DEGIL', '');
});

client.on('auth_failure', (msg) => {
    console.error('[WhatsApp] Kimlik dogrulama hatasi:', msg);
    durumYaz('HATA', '');
});

client.on('message', async (message) => {
    try {
        config = configOku(); // Her mesajda taze config oku

        // Mesaj alabildiysek bağlıyız — UI'da hâlâ BAGLANIYOR görünüyorsa düzelt
        if (_sonDurum !== 'BAGLI') bagliOlarakIsaretle('message');

        // Sistem warmup tamamlanmadıysa tetikleyicileri hazırla-kabul et-bekle
        if (!_tumSistemHazir) {
            const mesajKucuk = (message.body || '').toLowerCase().trim();
            const herhangiTetikleyici =
                ['pancar rapor','pancarrapor','seker rapor','şeker rapor','sekerrapor','şekerrapor',
                 'gubre ciro','gübre ciro','gubre rapor','gübre rapor','gubreciro','gübreciro','gubrerapor','gübrerapor']
                    .some(k => mesajKucuk.includes(k)) ||
                (config.tetikleyiciler || []).some(k => mesajKucuk.includes(k.toLowerCase()));
            if (herhangiTetikleyici) {
                console.log('[Warmup] Ilk mesaj geldi, sistem hazirlanana kadar bekleniyor...');
                let bekleme = 0;
                while (!_tumSistemHazir && bekleme < 60000) {
                    await new Promise(r => setTimeout(r, 500));
                    bekleme += 500;
                }
                if (!_tumSistemHazir) {
                    await message.reply('Sistem baslatiliyor, lutfen 30 saniye sonra tekrar deneyin.');
                    return;
                }
                // Hazır olunca ek 1sn buffer
                await new Promise(r => setTimeout(r, 1000));
            }
        }

        // Bireysel mesaj: message.from = "905xxxxxxx@c.us" veya "905xxxxxxx@lid" (Meta LID)
        // Grup mesajı:    message.from = "12036xxx@g.us", message.author = "905xxxxxxx@c.us"
        // @ işaretinden sonraki her şeyi sil, sadece rakamları al
        const numaraTemizle = (raw) => raw ? raw.replace(/@\S+$/, '').replace(/\D/g, '') : null;

        // Listedeki numaralarla normalleştirilmiş karşılaştırma (son 10 hane eşleşmesi)
        const numaraEslesiyor = (numara, liste) => {
            if (!numara) return false;
            return liste.some(k => {
                const temizK = k.replace(/\D/g, '');
                return temizK === numara ||
                       temizK.slice(-10) === numara.slice(-10);
            });
        };

        let bireyselNumara = numaraTemizle(message.from);
        let grupGonderenNumara = message.author ? numaraTemizle(message.author) : null;

        let gonderenNumara = numaraEslesiyor(bireyselNumara, config.yetkiliNumaralar)
            ? bireyselNumara
            : (numaraEslesiyor(grupGonderenNumara, config.yetkiliNumaralar) ? grupGonderenNumara : null);

        // Eşleşme bulunamadıysa getContact() ile gerçek telefon numarasını sorgula
        // (Meta LID formatında from alanı gerçek numara içermez)
        if (!gonderenNumara) {
            // Olasi telefon numarasi alanlarini dene
            const numaraCandidates = [];
            try {
                const contact = await message.getContact();
                if (contact) {
                    if (contact.number) numaraCandidates.push({ kaynak: 'contact.number', val: contact.number });
                    if (contact.id && contact.id.user) numaraCandidates.push({ kaynak: 'contact.id.user', val: contact.id.user });
                    // Bazi versiyonlarda getPhoneNumber/getFormattedNumber metodu var
                    if (typeof contact.getFormattedNumber === 'function') {
                        try { const fn = await contact.getFormattedNumber(); if (fn) numaraCandidates.push({ kaynak: 'getFormattedNumber', val: fn }); } catch (_) {}
                    }
                }
            } catch (e) {
                logYaz('DEBUG', `getContact HATA: ${e.message} (from=${message.from})`, 'Yetkisiz');
            }
            // Author varsa (grup) onu da dene
            if (message.author) {
                try {
                    const aContact = await client.getContactById(message.author);
                    if (aContact && aContact.number) numaraCandidates.push({ kaynak: 'author.number', val: aContact.number });
                } catch (_) {}
            }

            // En iyi eslesmeyi bul
            for (const c of numaraCandidates) {
                const gercekNumara = (c.val || '').replace(/\D/g, '');
                if (numaraEslesiyor(gercekNumara, config.yetkiliNumaralar)) {
                    gonderenNumara = gercekNumara;
                    bireyselNumara = gercekNumara;
                    logYaz('DEBUG', `LID cozumlendi (${c.kaynak}): from=${message.from} → ${gercekNumara}`, 'Yetkili');
                    break;
                }
            }
            // Yine de eslesmediyse adaylari logla
            if (!gonderenNumara && numaraCandidates.length > 0) {
                logYaz('DEBUG', `LID adaylari eslesmedi: from=${message.from} adaylar=${JSON.stringify(numaraCandidates)}`, 'Yetkisiz');
            } else if (!gonderenNumara) {
                logYaz('DEBUG', `LID coz aday yok: from=${message.from} author=${message.author||'-'}`, 'Yetkisiz');
            }
        }

        // Yetkisiz kullanıcı: bulanik=true yaparak devam et, yetkililer için bulanik=false
        let bulanik = false;
        if (!gonderenNumara) {
            bulanik = true;
            gonderenNumara = bireyselNumara || grupGonderenNumara || 'bilinmeyen';
        }

        const mesajIcerigi = message.body.toLowerCase().trim();

        // ── Şeker Raporu konuşma durumu kontrolü (sadece yetkili kullanıcılar) ──
        if (!bulanik && sekerKonusma.has(gonderenNumara)) {
            const durum = sekerKonusma.get(gonderenNumara);

            // İptal komutu
            if (mesajIcerigi === 'iptal' || mesajIcerigi === 'vazgec') {
                sekerKonusma.delete(gonderenNumara);
                await message.reply('Seker raporu iptal edildi.');
                logYaz(gonderenNumara, message.body, 'Iptal');
                return;
            }

            if (durum.adim === 'BASLANGIC') {
                const tarih = tarihParse(message.body.trim());
                if (!tarih) {
                    await message.reply('Tarih anlasılamadı. Lütfen DD.MM.YYYY formatında girin (örn: 01.09.2025) veya "iptal" yazın.');
                    return;
                }
                sekerKonusma.set(gonderenNumara, { adim: 'BITIS', baslangic: tarih });
                await message.reply(`Baslangic: ${tarih}\nSimdi bitis tarihini girin (DD.MM.YYYY):`);
                return;
            }

            if (durum.adim === 'BITIS') {
                const tarih = tarihParse(message.body.trim());
                if (!tarih) {
                    await message.reply('Tarih anlasılamadı. Lütfen DD.MM.YYYY formatında girin (örn: 30.09.2025) veya "iptal" yazın.');
                    return;
                }
                sekerKonusma.delete(gonderenNumara);
                await message.reply('Seker raporu hazirlaniyor, lutfen bekleyin...');
                logYaz(gonderenNumara, message.body, 'Hazirlaniyor...');
                try {
                    await sekerRaporuGonder(message, durum.baslangic, tarih, false);
                    logYaz(gonderenNumara, message.body, 'Gonderildi');
                } catch (e) {
                    logYaz(gonderenNumara, message.body, 'HATA: ' + e.message);
                    await message.reply('Seker raporu gonderilirken hata: ' + e.message);
                }
                return;
            }
        }

        // Pancar raporu tetikleyicileri
        const pancarTetikleyiciler = ['pancar rapor', 'pancarrapor'];
        const pancarTetiklendi = pancarTetikleyiciler.some(k => mesajIcerigi.includes(k));

        // Şeker raporu tetikleyicileri
        const sekerTetikleyiciler = ['seker rapor', 'şeker rapor', 'sekerrapor', 'şekerrapor'];
        const sekerTetiklendi = sekerTetikleyiciler.some(k => mesajIcerigi.includes(k));

        // Gübre Ciro tetikleyicileri (ÖNCE check edilmeli ki "gübre rapor" buna düşmesin)
        const gubreCiroTetikleyiciler = ['gubre ciro', 'gübre ciro', 'gubreciro', 'gübreciro'];
        const gubreCiroTetiklendi = gubreCiroTetikleyiciler.some(k => mesajIcerigi.includes(k));

        // Gübre Stok tetikleyicileri (ciro değilse)
        const gubreStokTetikleyiciler = ['gubre rapor', 'gübre rapor', 'gubrerapor', 'gübrerapor', 'gubre raporu', 'gübre raporu'];
        const gubreStokTetiklendi = !gubreCiroTetiklendi && gubreStokTetikleyiciler.some(k => mesajIcerigi.includes(k));

        // Genel rapor tetikleyicileri
        const tetiklendi = !pancarTetiklendi && !sekerTetiklendi && !gubreCiroTetiklendi && !gubreStokTetiklendi
            && config.tetikleyiciler.some(kelime => mesajIcerigi.includes(kelime.toLowerCase()));

        if (sekerTetiklendi) {
            console.log(`\n[${new Date().toLocaleString('tr-TR')}] Seker rapor talebi: ${gonderenNumara}${bulanik?' (yetkisiz)':''}`);
            // Mesajda ay adı var mı? (örn: "Şeker Rapor Eylül")
            const ayTarih = ayAdindenTarih(mesajIcerigi);
            if (ayTarih) {
                logYaz(gonderenNumara, message.body, bulanik ? 'Hazirlaniyor (bulanik)...' : 'Hazirlaniyor...');
                await message.reply(`Seker raporu hazirlaniyor (${ayTarih.baslangic} – ${ayTarih.bitis}), lutfen bekleyin...`);
                try {
                    await sekerRaporuGonder(message, ayTarih.baslangic, ayTarih.bitis, bulanik);
                    logYaz(gonderenNumara, message.body, bulanik ? 'Gonderildi (bulanik)' : 'Gonderildi');
                } catch (e) {
                    logYaz(gonderenNumara, message.body, 'HATA: ' + e.message);
                    await message.reply('Seker raporu gonderilirken hata: ' + e.message);
                }
            } else if (!bulanik) {
                // Tarih yok ve yetkili → konuşma modunu başlat
                sekerKonusma.set(gonderenNumara, { adim: 'BASLANGIC' });
                await message.reply('Seker raporu - baslangic tarihini girin (DD.MM.YYYY):\n(veya "Seker Rapor Eylul" gibi ay adi ile de gonderebilirsiniz)\n"iptal" yazarak vazgecebilirsiniz.');
            } else {
                // Tarih yok ve yetkisiz → mevcut ay için gönder (bulanık)
                const simdi = new Date();
                const yil = simdi.getFullYear();
                const ay = String(simdi.getMonth() + 1).padStart(2, '0');
                const sonGun = String(new Date(yil, simdi.getMonth() + 1, 0).getDate()).padStart(2, '0');
                const bas = `${yil}-${ay}-01`;
                const bit = `${yil}-${ay}-${sonGun}`;
                logYaz(gonderenNumara, message.body, 'Hazirlaniyor (bulanik, mevcut ay)...');
                await message.reply('Seker raporu hazirlaniyor, lutfen bekleyin...');
                try {
                    await sekerRaporuGonder(message, bas, bit, true);
                    logYaz(gonderenNumara, message.body, 'Gonderildi (bulanik)');
                } catch (e) {
                    logYaz(gonderenNumara, message.body, 'HATA: ' + e.message);
                    await message.reply('Seker raporu gonderilirken hata: ' + e.message);
                }
            }
        } else if (pancarTetiklendi) {
            console.log(`\n[${new Date().toLocaleString('tr-TR')}] Pancar rapor talebi: ${gonderenNumara}${bulanik?' (yetkisiz)':''}`);
            logYaz(gonderenNumara, message.body, bulanik ? 'Hazirlaniyor (bulanik)...' : 'Hazirlaniyor...');
            await message.reply('Pancar raporu hazirlaniyor, lutfen bekleyin...');
            try {
                const baseUrl = (config.raporApiUrl || 'http://localhost:5050/api/rapor')
                    .replace(/\/api\/.*$/, '');
                const pancarApiUrl = `${baseUrl}/api/pancar-raporu${bulanik ? '?bulanik=true' : ''}`;
                const pancarHtml = await sqlRaporuGetirRetry(pancarApiUrl);
                if (pancarHtml) {
                    await htmldenPngOlusturVeGonder(message, pancarHtml, 'pancar');
                    logYaz(gonderenNumara, message.body, bulanik ? 'Gonderildi (bulanik)' : 'Gonderildi');
                } else {
                    await message.reply('Pancar raporu alinamadi, sunucu kapali olabilir.');
                    logYaz(gonderenNumara, message.body, 'HATA: API yanit vermedi');
                }
            } catch (raporHata) {
                logYaz(gonderenNumara, message.body, 'HATA: ' + raporHata.message);
                await message.reply('Pancar raporu gonderilirken hata: ' + raporHata.message);
            }
        } else if (gubreCiroTetiklendi) {
            console.log(`\n[${new Date().toLocaleString('tr-TR')}] Gubre Ciro talebi: ${gonderenNumara}${bulanik?' (yetkisiz)':''}`);
            logYaz(gonderenNumara, message.body, bulanik ? 'Hazirlaniyor (bulanik)...' : 'Hazirlaniyor...');
            await message.reply('Gubre ciro raporu hazirlaniyor, lutfen bekleyin...');
            try {
                const baseUrl = (config.raporApiUrl || 'http://localhost:5050/api/rapor')
                    .replace(/\/api\/.*$/, '');
                const gubreHtml = await sqlRaporuGetirRetry(`${baseUrl}/api/gubre-ciro-raporu${bulanik ? '?bulanik=true' : ''}`);
                if (gubreHtml) {
                    await htmldenPngOlusturVeGonder(message, gubreHtml, 'gubre-ciro');
                    logYaz(gonderenNumara, message.body, bulanik ? 'Gonderildi (bulanik)' : 'Gonderildi');
                } else {
                    await message.reply('Gubre ciro raporu alinamadi, sunucu kapali olabilir.');
                    logYaz(gonderenNumara, message.body, 'HATA: API yanit vermedi');
                }
            } catch (raporHata) {
                logYaz(gonderenNumara, message.body, 'HATA: ' + raporHata.message);
                await message.reply('Gubre ciro raporu gonderilirken hata: ' + raporHata.message);
            }
        } else if (gubreStokTetiklendi) {
            console.log(`\n[${new Date().toLocaleString('tr-TR')}] Gubre Stok talebi: ${gonderenNumara}${bulanik?' (yetkisiz)':''}`);
            logYaz(gonderenNumara, message.body, bulanik ? 'Hazirlaniyor (bulanik)...' : 'Hazirlaniyor...');
            await message.reply('Gubre stok raporu hazirlaniyor, lutfen bekleyin...');
            try {
                const baseUrl = (config.raporApiUrl || 'http://localhost:5050/api/rapor')
                    .replace(/\/api\/.*$/, '');
                const gubreHtml = await sqlRaporuGetirRetry(`${baseUrl}/api/gubre-stok-raporu${bulanik ? '?bulanik=true' : ''}`);
                if (gubreHtml) {
                    await htmldenPngOlusturVeGonder(message, gubreHtml, 'gubre-stok');
                    logYaz(gonderenNumara, message.body, bulanik ? 'Gonderildi (bulanik)' : 'Gonderildi');
                } else {
                    await message.reply('Gubre stok raporu alinamadi, sunucu kapali olabilir.');
                    logYaz(gonderenNumara, message.body, 'HATA: API yanit vermedi');
                }
            } catch (raporHata) {
                logYaz(gonderenNumara, message.body, 'HATA: ' + raporHata.message);
                await message.reply('Gubre stok raporu gonderilirken hata: ' + raporHata.message);
            }
        } else if (tetiklendi) {
            console.log(`\n[${new Date().toLocaleString('tr-TR')}] Rapor talebi: ${gonderenNumara}${bulanik?' (yetkisiz)':''}`);
            logYaz(gonderenNumara, message.body, bulanik ? 'Hazirlaniyor (bulanik)...' : 'Hazirlaniyor...');
            await message.reply('Rapor hazirlaniyor, lutfen bekleyin...');
            try {
                await raporOlusturVeGonder(message, bulanik);
                logYaz(gonderenNumara, message.body, bulanik ? 'Gonderildi (bulanik)' : 'Gonderildi');
            } catch (raporHata) {
                logYaz(gonderenNumara, message.body, 'HATA: ' + raporHata.message);
                throw raporHata;
            }
        } else if (!bulanik) {
            // Yetkili ama tetikleyici değil — kayıt tutma
        }
    } catch (error) {
        console.error('[WhatsApp] Mesaj hatasi:', error.message);
    }
});

// MD protocol: kendi numaranızdan kendi numaranıza gönderilen mesajlar
// bazen 'message' yerine 'message_create' event'ini tetikliyor.
// fromMe=true olanları aynı handler'a yönlendir.
client.on('message_create', async (message) => {
    try {
        // Sadece kendimizin gönderdikleri (yetkili kendi numarasından)
        if (!message.fromMe) return;
        // Mesaj boş ise (sticker vs) atla
        if (!message.body) return;
        // 'message' eventine yönlendir — aynı async handler içeriği tekrar uygulanır
        // İçeriği kopyalamak yerine event'i emit edelim
        client.emit('message', message);
    } catch (_) {}
});

// =====================================================
// ANA FONKSİYON
// =====================================================

async function raporOlusturVeGonder(message, bulanik = false) {
    config = configOku();
    const baseApiUrl = config.raporApiUrl || 'http://localhost:5050/api/rapor';
    const apiUrl = bulanik ? `${baseApiUrl}?bulanik=true` : baseApiUrl;

    console.log('SQL raporu deneniyor:', apiUrl);
    const sqlHtml = await sqlRaporuGetirRetry(apiUrl);

    if (sqlHtml) {
        console.log('SQL raporu alindi, PNG ye cevriliyor...');
        await htmldenPngOlusturVeGonder(message, sqlHtml, 'sql');
    } else {
        console.log('API kapali, Excel raporuna geciliyor...');
        const excelDosyasi = config.excelDosyasi || '';
        if (!excelDosyasi || !fs.existsSync(excelDosyasi)) {
            await message.reply('Veritabani raporu hazir degil ve Excel dosyasi bulunamadi.');
            return;
        }
        await message.reply('Veritabani raporu hazir degil, Excel raporu hazirlaniyor...');
        await excelRaporuOlusturVeGonder(message, excelDosyasi);
    }
}

// =====================================================
// SQL RAPORU
// =====================================================

function sqlRaporuGetir(apiUrl, timeoutMs = 60000) {
    return new Promise((resolve) => {
        const req = http.get(apiUrl, { timeout: timeoutMs }, (res) => {
            if (res.statusCode !== 200) {
                console.log('API hata kodu:', res.statusCode);
                resolve(null);
                return;
            }
            let data = '';
            res.setEncoding('utf8');
            res.on('data', chunk => data += chunk);
            res.on('end', () => {
                resolve(data && data.length > 500 ? data : null);
            });
        });

        req.on('error', (err) => {
            console.log('API baglanamadi:', err.message);
            resolve(null);
        });

        req.on('timeout', () => {
            req.destroy();
            console.log('API zaman asimi');
            resolve(null);
        });
    });
}

// API'yi yeniden deneme ile çağırır — ilk istek DB soğuk başladığında başarısız olabilir
async function sqlRaporuGetirRetry(apiUrl) {
    let html = await sqlRaporuGetir(apiUrl, 60000);
    if (!html) {
        console.log('İlk API denemesi başarısız, 5 saniye sonra tekrar deneniyor...');
        await new Promise(r => setTimeout(r, 5000));
        html = await sqlRaporuGetir(apiUrl, 60000);
    }
    return html;
}

// =====================================================
// HTML → PNG → WhatsApp
// =====================================================

// Kalıcı browser örneği — her istekte yeniden başlatılmaz
let _browser = null;
let _puppeteerKilit = false;

async function browserGetir() {
    if (!_browser || !_browser.isConnected()) {
        console.log('[Puppeteer] Browser baslatiliyor...');
        _browser = await puppeteer.launch({
            headless: 'new',
            timeout: 60000,
            args: ['--no-sandbox', '--disable-setuid-sandbox',
                   '--font-render-hinting=none', '--disable-font-subpixel-positioning',
                   // Anti-throttling: tab arka planda da hızlı kalsın
                   '--disable-background-timer-throttling',
                   '--disable-backgrounding-occluded-windows',
                   '--disable-renderer-backgrounding']
        });
        _browser.on('disconnected', () => {
            console.log('[Puppeteer] Browser baglantisi kesildi, bir sonraki istekte yeniden baslatilacak.');
            _browser = null;
        });
        console.log('[Puppeteer] Browser hazir.');
    }
    return _browser;
}

async function htmldenPngOlusturVeGonder(message, htmlIcerik, kaynak, viewportGenislik = 1400) {
    // Kilit bekle (max 120 sn)
    let bekleme = 0;
    while (_puppeteerKilit && bekleme < 120000) {
        await new Promise(r => setTimeout(r, 1000));
        bekleme += 1000;
    }
    if (_puppeteerKilit) {
        await message.reply('Rapor sistemi mesgul, lutfen tekrar deneyin.');
        return;
    }
    _puppeteerKilit = true;

    let page = null;
    try {
        const dosyaAdi = kaynak === 'pancar' ? 'pancar'
            : kaynak === 'seker-analiz' ? 'seker-analiz'
            : kaynak === 'seker' ? 'seker'
            : kaynak === 'gubre-ciro' ? 'gubre-ciro'
            : kaynak === 'gubre-stok' ? 'gubre-stok'
            : 'rapor';
        const htmlDosya = path.join(CIKTI_KLASORU, dosyaAdi + '.html');
        const pngDosya  = path.join(CIKTI_KLASORU, dosyaAdi + '.png');

        fs.writeFileSync(htmlDosya, htmlIcerik, 'utf8');

        const browser = await browserGetir();
        page = await browser.newPage();
        await page.setExtraHTTPHeaders({ 'Accept-Language': 'tr-TR,tr;q=0.9' });
        await page.setViewport({ width: viewportGenislik, height: 900, deviceScaleFactor: 1 });

        const fileUrl = 'file:///' + htmlDosya.replace(/\\/g, '/');
        console.log(`[PNG] page.goto: ${kaynak}`);
        await page.goto(fileUrl, { waitUntil: 'domcontentloaded', timeout: 60000 });
        await new Promise(r => setTimeout(r, 800));

        console.log(`[PNG] page.screenshot: ${kaynak}`);
        // page.screenshot timeoutsuz olabilir — Promise.race ile 60sn limit
        await Promise.race([
            page.screenshot({ path: pngDosya, fullPage: true, type: 'png' }),
            new Promise((_, reject) => setTimeout(() => reject(new Error('screenshot 60sn timeout')), 60000))
        ]);
        console.log(`[PNG] screenshot OK: ${kaynak}`);
        await page.close();
        page = null;

        const pngSize = fs.statSync(pngDosya).size;
        console.log(`PNG olusturuldu (${kaynak}), boyut: ${pngSize} byte`);

        if (pngSize > 1000) {
            const media = MessageMedia.fromFilePath(pngDosya);
            const tarih = new Date().toLocaleString('tr-TR');
            const baslik = kaynak === 'sql'
                ? `Yan Urunler + Seker Uretim-Satis-Stok Raporu\n${tarih}`
                : kaynak === 'pancar'
                    ? `Pancar Raporu\n${tarih}`
                    : kaynak === 'seker-analiz'
                        ? `Seker Kategorisi Bazli Analiz (Ham Veri)\n${tarih}`
                        : kaynak === 'seker'
                            ? `Seker Uretim-Satis-Stok Raporu\n${tarih}`
                            : kaynak === 'gubre-ciro'
                                ? `Gubre Ciro Raporu (Net, iade dusulmus)\n${tarih}`
                                : kaynak === 'gubre-stok'
                                    ? `Gubre Stok Raporu\n${tarih}`
                                    : `Tum Rapor (Excel)\n${tarih}`;
            await message.reply(media, undefined, { caption: baslik });
            console.log('Rapor gonderildi!\n');
        } else {
            throw new Error('PNG dosyasi cok kucuk: ' + pngSize);
        }
    } catch (error) {
        console.error('PNG/Gonderim hatasi:', error.message);
        logYaz('SISTEM', `PNG hata (${kaynak}): ${error.message}`, 'HATA');
        if (page) { try { await page.close(); } catch (_) {} }
        // Browser bozulmuş olabilir — bir sonraki istekte yeniden başlat
        try { if (_browser) { await _browser.close(); } } catch (_) {}
        _browser = null;
        try { await message.reply('Rapor gonderilirken hata olustu: ' + error.message); } catch (_) {}
    } finally {
        _puppeteerKilit = false;
    }
}

// =====================================================
// EXCEL RAPORU (Fallback)
// =====================================================

async function excelRaporuOlusturVeGonder(message, excelDosyasi) {
    try {
        const htmlDosya = path.join(CIKTI_KLASORU, 'rapor_excel.html');

        if (fs.existsSync(htmlDosya)) fs.unlinkSync(htmlDosya);

        const vbsIcerik = `
On Error Resume Next
Dim xlApp, xlBook
Set xlApp = CreateObject("Excel.Application")
xlApp.Visible = False
xlApp.DisplayAlerts = False
Set xlBook = xlApp.Workbooks.Open("${excelDosyasi.replace(/\\/g, '\\\\')}")
If Err.Number <> 0 Then
    WScript.Echo "HATA: " & Err.Description
    WScript.Quit 1
End If
xlApp.Run "WhatsAppRaporOlustur"
If Err.Number <> 0 Then
    WScript.Echo "HATA: Makro - " & Err.Description
    xlBook.Close False
    xlApp.Quit
    WScript.Quit 1
End If
xlBook.Close False
xlApp.Quit
Set xlBook = Nothing
Set xlApp = Nothing
WScript.Echo "BASARILI"
WScript.Quit 0
`;
        const vbsDosya = path.join(CIKTI_KLASORU, 'rapor.vbs');
        fs.writeFileSync(vbsDosya, vbsIcerik, 'ascii');

        await new Promise((resolve, reject) => {
            exec(`cscript //nologo "${vbsDosya}"`, { timeout: 180000 }, (error, stdout, stderr) => {
                console.log('VBScript:', stdout);
                if (stderr) console.log('VBScript stderr:', stderr);
                if (stdout.includes('BASARILI')) resolve(stdout);
                else if (error) reject(new Error(stdout || stderr || error.message));
                else resolve(stdout);
            });
        });

        await new Promise(r => setTimeout(r, 2000));

        if (!fs.existsSync(htmlDosya)) {
            throw new Error('HTML dosyasi olusturulamadi');
        }

        const htmlIcerik = fs.readFileSync(htmlDosya, 'utf8');
        await htmldenPngOlusturVeGonder(message, htmlIcerik, 'excel');

    } catch (error) {
        console.error('Excel rapor hatasi:', error.message);
        await message.reply('Excel raporu hazirlanamadi: ' + error.message);
    }
}

// =====================================================
// BASLAT
// =====================================================

console.log('[WhatsApp] Baslatiliyor...');
console.log('[WhatsApp] Veri dizini:', VERI_DIZIN);
durumYaz('BAGLANIYOR', '');
// initialize cagrisi ile ilk baglanti icin 120 sn sinir koy
hazirTimerKur(120, 'initialize sonrasi qr/ready gelmedi');
client.initialize();
