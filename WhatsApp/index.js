// =====================================================
// RaporlamaPortali WhatsApp Bot — Baileys
// Migration nedeni: whatsapp-web.js'in Multi-Device protokolüyle uyumsuzluğu;
// "BAGLI ama mesaj akmiyor" silent failure problemi.
// =====================================================

const { default: makeWASocket, useMultiFileAuthState, DisconnectReason,
        fetchLatestBaileysVersion, Browsers } = require('@whiskeysockets/baileys');
const pino = require('pino');
const fs = require('fs');
const path = require('path');
const http = require('http');
const puppeteer = require('puppeteer');

// =====================================================
// DOSYA YOLLARI
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
const AUTH_KLASORU  = path.join(VERI_DIZIN, '.baileys_auth');

try { if (!fs.existsSync(CIKTI_KLASORU)) fs.mkdirSync(CIKTI_KLASORU, { recursive: true }); } catch (_) {}

// =====================================================
// CONFIG / DURUM / LOG
// =====================================================

function configOku() {
    try {
        if (fs.existsSync(CONFIG_DOSYA)) return JSON.parse(fs.readFileSync(CONFIG_DOSYA, 'utf8'));
    } catch (e) { console.error('Config okuma hatasi:', e.message); }
    return {
        yetkiliNumaralar: [],
        tetikleyiciler: ['tüm rapor', 'tum rapor', 'tumrapor', 'tümrapor'],
        raporApiUrl: 'http://localhost:5050/api/rapor',
        excelDosyasi: ''
    };
}

let _sonDurum = null;
function durumYaz(durum, qrString) {
    try {
        const data = { durum, qrString: qrString || '', guncelleme: new Date().toISOString() };
        fs.writeFileSync(DURUM_DOSYA, JSON.stringify(data), 'utf8');
        _sonDurum = durum;
    } catch (e) { console.error('Durum yazma hatasi:', e.message); }
}

function logYaz(numara, mesaj, sonuc) {
    try {
        let kayitlar = [];
        if (fs.existsSync(LOG_DOSYA)) {
            try { kayitlar = JSON.parse(fs.readFileSync(LOG_DOSYA, 'utf8')); } catch (_) { kayitlar = []; }
        }
        kayitlar.unshift({ tarih: new Date().toISOString(), numara, mesaj, sonuc });
        if (kayitlar.length > 500) kayitlar = kayitlar.slice(0, 500);
        fs.writeFileSync(LOG_DOSYA, JSON.stringify(kayitlar), 'utf8');
    } catch (e) { console.error('Log yazma hatasi:', e.message); }
}

// =====================================================
// HTTP RAPOR (api/...) — eski koddan aynen
// =====================================================

function sqlRaporuGetir(url, timeoutMs) {
    timeoutMs = timeoutMs || 90000;
    return new Promise((resolve) => {
        const req = http.get(url, { timeout: timeoutMs }, (res) => {
            let veri = '';
            res.on('data', (c) => { veri += c; });
            res.on('end', () => {
                if (res.statusCode >= 200 && res.statusCode < 300 && veri.length > 100) resolve(veri);
                else { console.error(`[API] HTTP ${res.statusCode} url=${url}`); resolve(null); }
            });
        });
        req.on('error', (e) => { console.error('[API] hata:', e.message); resolve(null); });
        req.on('timeout', () => { req.destroy(); resolve(null); });
    });
}

async function sqlRaporuGetirRetry(url, maxRetry) {
    maxRetry = maxRetry || 3;
    for (let i = 0; i < maxRetry; i++) {
        const r = await sqlRaporuGetir(url);
        if (r) return r;
        if (i < maxRetry - 1) await new Promise(r => setTimeout(r, 2000));
    }
    return null;
}

// =====================================================
// PUPPETEER PNG
// =====================================================

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
                   '--disable-background-timer-throttling',
                   '--disable-backgrounding-occluded-windows',
                   '--disable-renderer-backgrounding']
        });
        _browser.on('disconnected', () => { _browser = null; });
        console.log('[Puppeteer] Browser hazir.');
    }
    return _browser;
}

async function htmldenPngOlustur(htmlIcerik, kaynak, viewportGenislik) {
    viewportGenislik = viewportGenislik || 1400;
    let bekleme = 0;
    while (_puppeteerKilit && bekleme < 120000) {
        await new Promise(r => setTimeout(r, 500));
        bekleme += 500;
    }
    if (_puppeteerKilit) throw new Error('Puppeteer kilidi 120sn icinde acilmadi');
    _puppeteerKilit = true;

    let page = null;
    try {
        const dosyaAdi = kaynak === 'pancar' ? 'pancar'
            : kaynak === 'seker-analiz' ? 'seker-analiz'
            : kaynak === 'seker' ? 'seker'
            : kaynak === 'gubre-ciro' ? 'gubre-ciro'
            : kaynak === 'gubre-stok' ? 'gubre-stok'
            : kaynak === 'cay-durum' ? 'cay-durum'
            : 'rapor';
        const htmlDosya = path.join(CIKTI_KLASORU, dosyaAdi + '.html');
        const pngDosya  = path.join(CIKTI_KLASORU, dosyaAdi + '.png');

        fs.writeFileSync(htmlDosya, htmlIcerik, 'utf8');

        const browser = await browserGetir();
        page = await browser.newPage();
        await page.setExtraHTTPHeaders({ 'Accept-Language': 'tr-TR,tr;q=0.9' });
        await page.setViewport({ width: viewportGenislik, height: 900, deviceScaleFactor: 1 });

        const fileUrl = 'file:///' + htmlDosya.replace(/\\/g, '/');
        await page.goto(fileUrl, { waitUntil: 'domcontentloaded', timeout: 60000 });
        await new Promise(r => setTimeout(r, 800));
        await Promise.race([
            page.screenshot({ path: pngDosya, fullPage: true, type: 'png' }),
            new Promise((_, reject) => setTimeout(() => reject(new Error('screenshot 60sn timeout')), 60000))
        ]);
        await page.close();
        page = null;

        const sz = fs.statSync(pngDosya).size;
        if (sz < 1000) throw new Error('PNG cok kucuk: ' + sz);
        return pngDosya;
    } catch (e) {
        if (page) { try { await page.close(); } catch (_) {} }
        try { if (_browser) await _browser.close(); } catch (_) {}
        _browser = null;
        throw e;
    } finally {
        _puppeteerKilit = false;
    }
}

function captionFor(kaynak) {
    const t = new Date().toLocaleString('tr-TR');
    if (kaynak === 'pancar')        return `Pancar Raporu\n${t}`;
    if (kaynak === 'seker-analiz')  return `Seker Kategorisi Bazli Analiz (Ham Veri)\n${t}`;
    if (kaynak === 'seker')         return `Seker Uretim-Satis-Stok Raporu\n${t}`;
    if (kaynak === 'gubre-ciro')    return `Gubre Ciro Raporu (Net, iade dusulmus)\n${t}`;
    if (kaynak === 'gubre-stok')    return `Gubre Stok Raporu\n${t}`;
    if (kaynak === 'cay-durum')     return `Cay Durum Raporu\n${t}`;
    return `Yan Urunler + Seker Uretim-Satis-Stok Raporu\n${t}`;
}

// =====================================================
// NUMARA / YETKI
// =====================================================

function numaraTemizle(raw) {
    if (!raw) return null;
    return raw.replace(/@\S+$/, '').split(':')[0].replace(/\D/g, '');
}

function numaraEslesiyor(numara, liste) {
    if (!numara) return false;
    return liste.some(k => {
        const tk = (k || '').replace(/\D/g, '');
        return tk === numara || (tk.length >= 10 && tk.slice(-10) === numara.slice(-10));
    });
}

// =====================================================
// BAILEYS BOT
// =====================================================

let _sock = null;
let _hazirZamani = 0;
let _yenidenBaglaniyor = false;
let _tumSistemHazir = false;

async function sistemWarmupYap() {
    try {
        console.log('[Warmup] Puppeteer on-baslatma...');
        await browserGetir();
        console.log('[Warmup] API ping...');
        const config = configOku();
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
        console.log('[Warmup] Sistem hazir.');
    } catch (err) {
        console.error('[Warmup] Hata:', err.message);
        _tumSistemHazir = true;
    }
}

async function botBaslat() {
    if (_yenidenBaglaniyor) return;
    _yenidenBaglaniyor = true;

    try {
        const { state, saveCreds } = await useMultiFileAuthState(AUTH_KLASORU);
        const { version } = await fetchLatestBaileysVersion();
        console.log(`[Baileys] WA Web v${version.join('.')} ile baslatiliyor`);

        _sock = makeWASocket({
            version,
            auth: state,
            logger: pino({ level: 'silent' }),
            browser: Browsers.macOS('Chrome'),
            markOnlineOnConnect: true,
            syncFullHistory: false,
            generateHighQualityLinkPreview: false,
        });

        _sock.ev.on('creds.update', saveCreds);

        _sock.ev.on('connection.update', async (update) => {
            const { connection, lastDisconnect, qr } = update;
            if (qr) {
                console.log('[Baileys] QR olusturuldu');
                durumYaz('QR_BEKLIYOR', qr);
            }
            if (connection === 'connecting') {
                console.log('[Baileys] Baglaniyor...');
                if (_sonDurum !== 'QR_BEKLIYOR') durumYaz('BAGLANIYOR', '');
            }
            if (connection === 'open') {
                console.log('[Baileys] Baglandi!');
                _hazirZamani = Date.now();
                if (!_tumSistemHazir) await sistemWarmupYap();
                durumYaz('BAGLI', '');
            }
            if (connection === 'close') {
                const reasonCode = lastDisconnect?.error?.output?.statusCode;
                const reason = lastDisconnect?.error?.message || 'unknown';
                console.log(`[Baileys] Baglanti kesildi: ${reason} (code=${reasonCode})`);
                durumYaz('BAGLI_DEGIL', '');
                _yenidenBaglaniyor = false;

                if (reasonCode === DisconnectReason.loggedOut) {
                    console.log('[Baileys] Cikis yapildi. Auth temizleniyor, yeni QR icin yeniden baslayin.');
                    try { fs.rmSync(AUTH_KLASORU, { recursive: true, force: true }); } catch (_) {}
                    setTimeout(() => botBaslat(), 3000);
                } else {
                    setTimeout(() => botBaslat(), 3000);
                }
                return;
            }
        });

        _sock.ev.on('messages.upsert', async ({ messages, type }) => {
            if (type !== 'notify') return;
            for (const m of messages) {
                try { await mesajIsle(m); } catch (e) { console.error('mesajIsle hata:', e.message); }
            }
        });

        _yenidenBaglaniyor = false;
    } catch (e) {
        console.error('[Baileys] Bot baslatma hatasi:', e.message);
        _yenidenBaglaniyor = false;
        setTimeout(() => botBaslat(), 5000);
    }
}

// =====================================================
// MESAJ ISLEME
// =====================================================

async function mesajIsle(msg) {
    if (!msg.message) return;

    // Mesaj metni cesitli yerlerde olabilir
    const body = msg.message.conversation
        || msg.message.extendedTextMessage?.text
        || msg.message.imageMessage?.caption
        || msg.message.videoMessage?.caption
        || '';
    if (!body) return;

    const fromJid = msg.key.remoteJid;
    const fromMe = !!msg.key.fromMe;
    const isGroup = fromJid?.endsWith('@g.us');
    const author = msg.key.participant; // grup mesajlarinda gercek gonderen

    // Gondereni cikar — coklu kaynak dene
    let senderPhone = null;
    let phoneKaynagi = '';

    // 1) Birincil JID (remote veya participant) — telefon formatinda mi?
    const birincilJid = isGroup ? (author || '') : (fromJid || '');
    if (birincilJid && birincilJid.endsWith('@s.whatsapp.net')) {
        senderPhone = numaraTemizle(birincilJid);
        phoneKaynagi = 'jid:s.whatsapp.net';
    }

    // 2) Mesajdaki PN ek alanlari (Baileys 7.x LID ile birlikte saklar)
    if (!senderPhone) {
        const pnAdaylar = [
            { k: 'key.participantPn', v: msg.key?.participantPn },
            { k: 'key.senderPn',      v: msg.key?.senderPn },
            { k: 'participantPn',     v: msg.participantPn },
            { k: 'senderPn',          v: msg.senderPn },
        ];
        for (const a of pnAdaylar) {
            if (a.v && /\d{6,}/.test(a.v)) {
                senderPhone = numaraTemizle(a.v);
                phoneKaynagi = a.k;
                break;
            }
        }
    }

    // 3) LID mapping uzerinden cevir
    if (!senderPhone) {
        try {
            const lidJid = isGroup ? author : fromJid;
            if (lidJid && lidJid.endsWith('@lid')) {
                const mapping = _sock?.signalRepository?.lidMapping;
                if (mapping && typeof mapping.getPNForLID === 'function') {
                    const pn = await mapping.getPNForLID(lidJid);
                    if (pn) {
                        senderPhone = numaraTemizle(pn);
                        phoneKaynagi = 'lidMapping';
                    }
                }
            }
        } catch (_) {}
    }

    // 4) onWhatsApp ile sorgula (LID'yi gercek hesaba dogrula)
    if (!senderPhone) {
        try {
            const lidJid = isGroup ? author : fromJid;
            if (lidJid && typeof _sock?.onWhatsApp === 'function') {
                const r = await _sock.onWhatsApp(lidJid);
                if (Array.isArray(r) && r.length > 0) {
                    const found = r[0];
                    if (found.jid && found.jid.endsWith('@s.whatsapp.net')) {
                        senderPhone = numaraTemizle(found.jid);
                        phoneKaynagi = 'onWhatsApp';
                    }
                }
            }
        } catch (_) {}
    }

    // 5) fromMe ise bot hesabini kullan
    if (!senderPhone && fromMe) {
        const myId = _sock?.user?.id;
        if (myId) { senderPhone = numaraTemizle(myId); phoneKaynagi = 'sock.user.id (fromMe)'; }
    }

    // 6) Yine yoksa, JID'den ham digit cikar
    if (!senderPhone) {
        senderPhone = numaraTemizle(birincilJid || fromJid || '');
        phoneKaynagi = phoneKaynagi || 'raw-jid';
    }

    const config = configOku();

    // Bot hesabinin kendi JID/LID'sini cikar (numara formuna donustur)
    const botId = _sock?.user?.id || '';
    const botLid = _sock?.user?.lid || '';
    const botIdNumara = numaraTemizle(botId);
    const botLidNumara = numaraTemizle(botLid);

    // Self-chat tespiti: sohbetin karsi tarafi bot hesabinin kendisi mi?
    // Bu durumda kullanici kendi numarasina mesaj atiyor demektir.
    const remoteNumara = numaraTemizle(fromJid);
    const selfChat = !isGroup && (
        (botIdNumara && remoteNumara === botIdNumara) ||
        (botLidNumara && remoteNumara === botLidNumara)
    );

    // KRITIK: fromMe=true (kullanicinin kendi telefonundan giden mesaj) ama self-chat degilse,
    // bu mesaj baska birine yazilmis demektir. Tetiklemeye sakin GIRMA.
    if (fromMe && !selfChat) return;

    // Yetki: self-chat'te otomatik yetkili; aksi halde gonderen numara yetkili listesinde olmali.
    const yetkili = selfChat || numaraEslesiyor(senderPhone, config.yetkiliNumaralar || []);
    const bulanik = !yetkili;
    const gonderenNumara = senderPhone || 'bilinmeyen';

    const mesajIcerigi = body.toLowerCase().trim();
    const replyTo = fromJid;

    async function reply(text) {
        try { await _sock.sendMessage(replyTo, { text }, { quoted: msg }); } catch (e) { console.error('reply hata:', e.message); }
    }
    async function replyImage(pngPath, kaynak) {
        try {
            await _sock.sendMessage(replyTo, {
                image: fs.readFileSync(pngPath),
                caption: captionFor(kaynak)
            }, { quoted: msg });
        } catch (e) { console.error('replyImage hata:', e.message); }
    }

    // Warmup henuz bitmediyse, tetikleyici varsa biraz bekle
    if (!_tumSistemHazir) {
        let b = 0;
        while (!_tumSistemHazir && b < 60000) {
            await new Promise(r => setTimeout(r, 500)); b += 500;
        }
        if (!_tumSistemHazir) { await reply('Sistem baslatiliyor, lutfen 30 saniye sonra tekrar deneyin.'); return; }
    }

    // Tetikleyiciler
    const pancarTetiklendi = ['pancar rapor', 'pancarrapor'].some(k => mesajIcerigi.includes(k));
    const sekerTetiklendi = ['seker rapor', 'şeker rapor', 'sekerrapor', 'şekerrapor'].some(k => mesajIcerigi.includes(k));
    const gubreCiroTetiklendi = ['gubre ciro', 'gübre ciro', 'gubreciro', 'gübreciro'].some(k => mesajIcerigi.includes(k));
    const gubreStokTetiklendi = !gubreCiroTetiklendi &&
        ['gubre rapor', 'gübre rapor', 'gubrerapor', 'gübrerapor', 'gubre raporu', 'gübre raporu'].some(k => mesajIcerigi.includes(k));
    const cayTetiklendi = ['cay rapor', 'çay rapor', 'cayrapor', 'çayrapor', 'cay raporu', 'çay raporu', 'cay durum', 'çay durum'].some(k => mesajIcerigi.includes(k));
    const tetiklendi = !pancarTetiklendi && !sekerTetiklendi && !gubreCiroTetiklendi && !gubreStokTetiklendi && !cayTetiklendi
        && (config.tetikleyiciler || []).some(k => mesajIcerigi.includes((k || '').toLowerCase()));

    if (!pancarTetiklendi && !sekerTetiklendi && !gubreCiroTetiklendi && !gubreStokTetiklendi && !cayTetiklendi && !tetiklendi) return;

    const baseUrl = (config.raporApiUrl || 'http://localhost:5050/api/rapor').replace(/\/api\/.*$/, '');
    const bSuf = bulanik ? '?bulanik=true' : '';

    console.log(`\n[${new Date().toLocaleString('tr-TR')}] Talep: ${gonderenNumara} - "${body}" ${bulanik?'(yetkisiz)':''}`);

    try {
        if (pancarTetiklendi) {
            logYaz(gonderenNumara, body, bulanik ? 'Hazirlaniyor (bulanik)...' : 'Hazirlaniyor...');
            await reply('Pancar raporu hazirlaniyor, lutfen bekleyin...');
            const html = await sqlRaporuGetirRetry(`${baseUrl}/api/pancar-raporu${bSuf}`);
            if (html) {
                const png = await htmldenPngOlustur(html, 'pancar');
                await replyImage(png, 'pancar');
                logYaz(gonderenNumara, body, bulanik ? 'Gonderildi (bulanik)' : 'Gonderildi');
            } else {
                await reply('Pancar raporu alinamadi.');
                logYaz(gonderenNumara, body, 'HATA: API yanit yok');
            }
        } else if (sekerTetiklendi) {
            // Şeker raporu icin tarih bazli karmasik logic var; simdilik basit: mevcut ay
            const simdi = new Date();
            const yil = simdi.getFullYear();
            const ay  = String(simdi.getMonth() + 1).padStart(2, '0');
            const sonGun = String(new Date(yil, simdi.getMonth() + 1, 0).getDate()).padStart(2, '0');
            const bas = `${yil}-${ay}-01`;
            const bit = `${yil}-${ay}-${sonGun}`;
            logYaz(gonderenNumara, body, bulanik ? 'Hazirlaniyor (bulanik, mevcut ay)...' : 'Hazirlaniyor (mevcut ay)...');
            await reply('Seker raporu hazirlaniyor (mevcut ay), lutfen bekleyin...');
            const url = `${baseUrl}/api/seker-raporu?baslangic=${bas}&bitis=${bit}${bulanik?'&bulanik=true':''}`;
            const html = await sqlRaporuGetirRetry(url);
            if (html) {
                const png = await htmldenPngOlustur(html, 'seker');
                await replyImage(png, 'seker');
                logYaz(gonderenNumara, body, bulanik ? 'Gonderildi (bulanik)' : 'Gonderildi');
            } else {
                await reply('Seker raporu alinamadi.');
                logYaz(gonderenNumara, body, 'HATA: API yanit yok');
            }
        } else if (gubreCiroTetiklendi) {
            logYaz(gonderenNumara, body, bulanik ? 'Hazirlaniyor (bulanik)...' : 'Hazirlaniyor...');
            await reply('Gubre ciro raporu hazirlaniyor, lutfen bekleyin...');
            const html = await sqlRaporuGetirRetry(`${baseUrl}/api/gubre-ciro-raporu${bSuf}`);
            if (html) {
                const png = await htmldenPngOlustur(html, 'gubre-ciro');
                await replyImage(png, 'gubre-ciro');
                logYaz(gonderenNumara, body, bulanik ? 'Gonderildi (bulanik)' : 'Gonderildi');
            } else {
                await reply('Gubre ciro raporu alinamadi.');
                logYaz(gonderenNumara, body, 'HATA: API yanit yok');
            }
        } else if (cayTetiklendi) {
            logYaz(gonderenNumara, body, bulanik ? 'Hazirlaniyor (bulanik)...' : 'Hazirlaniyor...');
            await reply('Cay durum raporu hazirlaniyor, lutfen bekleyin...');
            const html = await sqlRaporuGetirRetry(`${baseUrl}/api/cay-durum-raporu${bSuf}`);
            if (html) {
                const png = await htmldenPngOlustur(html, 'cay-durum');
                await replyImage(png, 'cay-durum');
                logYaz(gonderenNumara, body, bulanik ? 'Gonderildi (bulanik)' : 'Gonderildi');
            } else {
                await reply('Cay durum raporu alinamadi.');
                logYaz(gonderenNumara, body, 'HATA: API yanit yok');
            }
        } else if (gubreStokTetiklendi) {
            logYaz(gonderenNumara, body, bulanik ? 'Hazirlaniyor (bulanik)...' : 'Hazirlaniyor...');
            await reply('Gubre stok raporu hazirlaniyor, lutfen bekleyin...');
            const html = await sqlRaporuGetirRetry(`${baseUrl}/api/gubre-stok-raporu${bSuf}`);
            if (html) {
                const png = await htmldenPngOlustur(html, 'gubre-stok');
                await replyImage(png, 'gubre-stok');
                logYaz(gonderenNumara, body, bulanik ? 'Gonderildi (bulanik)' : 'Gonderildi');
            } else {
                await reply('Gubre stok raporu alinamadi.');
                logYaz(gonderenNumara, body, 'HATA: API yanit yok');
            }
        } else if (tetiklendi) {
            logYaz(gonderenNumara, body, bulanik ? 'Hazirlaniyor (bulanik)...' : 'Hazirlaniyor...');
            await reply('Rapor hazirlaniyor, lutfen bekleyin...');
            const html = await sqlRaporuGetirRetry(`${baseUrl}/api/rapor${bSuf}`);
            if (html) {
                const png = await htmldenPngOlustur(html, 'sql');
                await replyImage(png, 'sql');
                logYaz(gonderenNumara, body, bulanik ? 'Gonderildi (bulanik)' : 'Gonderildi');
            } else {
                await reply('Rapor alinamadi.');
                logYaz(gonderenNumara, body, 'HATA: API yanit yok');
            }
        }
    } catch (e) {
        console.error('[Trigger] HATA:', e.message);
        logYaz(gonderenNumara, body, 'HATA: ' + e.message);
        try { await reply('Rapor gonderilirken hata: ' + e.message); } catch (_) {}
    }
}

// =====================================================
// HEARTBEAT — durum dosyasini taze tut (C# watchdog icin)
// =====================================================

setInterval(() => {
    if (_sonDurum === 'BAGLI' && _sock?.user) {
        durumYaz('BAGLI', '');
    }
}, 20000);

// =====================================================
// BASLAT
// =====================================================

durumYaz('BAGLANIYOR', '');
botBaslat().catch(e => {
    console.error('[Boot] Hata:', e.message);
    durumYaz('BAGLI_DEGIL', '');
});

process.on('uncaughtException', (e) => { console.error('[UNCAUGHT]', e); });
process.on('unhandledRejection', (e) => { console.error('[UNHANDLED]', e); });
