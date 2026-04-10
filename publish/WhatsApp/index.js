const { Client, LocalAuth, MessageMedia } = require('whatsapp-web.js');
const qrcodeTerminal = require('qrcode-terminal');
const { exec } = require('child_process');
const path = require('path');
const fs = require('fs');
const puppeteer = require('puppeteer');
const http = require('http');

// =====================================================
// DOSYA YOLLARI
// =====================================================

const DIZIN         = __dirname;
const CONFIG_DOSYA  = path.join(DIZIN, 'whatsapp-config.json');
const DURUM_DOSYA   = path.join(DIZIN, 'whatsapp-status.json');
const LOG_DOSYA     = path.join(DIZIN, 'whatsapp-log.json');
const CIKTI_KLASORU = path.join(DIZIN, 'screenshots');

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

function durumYaz(durum, qrString) {
    try {
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

durumYaz('BAGLI_DEGIL', '');

// =====================================================
// WHATSAPP CLIENT
// =====================================================

const client = new Client({
    authStrategy: new LocalAuth({ dataPath: path.join(DIZIN, '.wwebjs_auth') }),
    puppeteer: {
        headless: true,
        args: ['--no-sandbox', '--disable-setuid-sandbox']
    }
});

client.on('qr', (qr) => {
    console.log('\n[WhatsApp] QR kodu bekleniyor...');
    qrcodeTerminal.generate(qr, { small: true });
    // Ham QR string'ini yaz, .NET tarafı image'a çevirir
    durumYaz('QR_BEKLIYOR', qr);
});

client.on('ready', () => {
    config = configOku();
    console.log('[WhatsApp] Baglandi!');
    durumYaz('BAGLI', '');
});

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

        // Bireysel mesaj: message.from = "905xxxxxxx@c.us"
        // Grup mesajı:    message.from = "12036xxx@g.us", message.author = "905xxxxxxx@c.us"
        const bireyselNumara = message.from.replace('@c.us', '');
        const grupGonderenNumara = message.author ? message.author.replace('@c.us', '') : null;
        const gonderenNumara = config.yetkiliNumaralar.includes(bireyselNumara)
            ? bireyselNumara
            : (grupGonderenNumara && config.yetkiliNumaralar.includes(grupGonderenNumara) ? grupGonderenNumara : null);

        if (!gonderenNumara) {
            return;
        }

        const mesajIcerigi = message.body.toLowerCase().trim();

        // Pancar raporu tetikleyicileri
        const pancarTetikleyiciler = ['pancar rapor', 'pancarrapor'];
        const pancarTetiklendi = pancarTetikleyiciler.some(k => mesajIcerigi.includes(k));

        // Genel rapor tetikleyicileri
        const tetiklendi = !pancarTetiklendi && config.tetikleyiciler.some(kelime =>
            mesajIcerigi.includes(kelime.toLowerCase())
        );

        if (pancarTetiklendi) {
            console.log(`\n[${new Date().toLocaleString('tr-TR')}] Pancar rapor talebi: ${gonderenNumara}`);
            logYaz(gonderenNumara, message.body, 'Hazirlaniyor...');
            await message.reply('Pancar raporu hazirlaniyor, lutfen bekleyin...');
            try {
                const pancarApiUrl = (config.raporApiUrl || 'http://localhost:5050/api/rapor')
                    .replace('/api/rapor', '/api/pancar-raporu');
                const pancarHtml = await sqlRaporuGetir(pancarApiUrl);
                if (pancarHtml) {
                    await htmldenPngOlusturVeGonder(message, pancarHtml, 'pancar');
                    logYaz(gonderenNumara, message.body, 'Gonderildi');
                } else {
                    await message.reply('Pancar raporu alinamadi, sunucu kapali olabilir.');
                    logYaz(gonderenNumara, message.body, 'HATA: API yanit vermedi');
                }
            } catch (raporHata) {
                logYaz(gonderenNumara, message.body, 'HATA: ' + raporHata.message);
                await message.reply('Pancar raporu gonderilirken hata: ' + raporHata.message);
            }
        } else if (tetiklendi) {
            console.log(`\n[${new Date().toLocaleString('tr-TR')}] Rapor talebi: ${gonderenNumara}`);
            logYaz(gonderenNumara, message.body, 'Hazirlaniyor...');
            await message.reply('Rapor hazirlaniyor, lutfen bekleyin...');
            try {
                await raporOlusturVeGonder(message);
                logYaz(gonderenNumara, message.body, 'Gonderildi');
            } catch (raporHata) {
                logYaz(gonderenNumara, message.body, 'HATA: ' + raporHata.message);
                throw raporHata;
            }
        } else {
            // Yetkili ama tetikleyici değil — kayıt tutma
        }
    } catch (error) {
        console.error('[WhatsApp] Mesaj hatasi:', error.message);
    }
});

// =====================================================
// ANA FONKSİYON
// =====================================================

async function raporOlusturVeGonder(message) {
    config = configOku();
    const apiUrl = config.raporApiUrl || 'http://localhost:5050/api/rapor';

    console.log('SQL raporu deneniyor:', apiUrl);
    const sqlHtml = await sqlRaporuGetir(apiUrl);

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

function sqlRaporuGetir(apiUrl) {
    return new Promise((resolve) => {
        const req = http.get(apiUrl, { timeout: 15000 }, (res) => {
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

// =====================================================
// HTML → PNG → WhatsApp
// =====================================================

// Eş zamanlı Puppeteer çakışmasını önlemek için kilit
let _puppeteerKilit = false;

async function htmldenPngOlusturVeGonder(message, htmlIcerik, kaynak) {
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

    let browser = null;
    try {
        const dosyaAdi = kaynak === 'pancar' ? 'pancar' : 'rapor';
        const htmlDosya = path.join(CIKTI_KLASORU, dosyaAdi + '.html');
        const pngDosya  = path.join(CIKTI_KLASORU, dosyaAdi + '.png');

        fs.writeFileSync(htmlDosya, htmlIcerik, 'utf8');

        browser = await puppeteer.launch({
            headless: 'new',
            args: ['--no-sandbox', '--disable-setuid-sandbox',
                   '--font-render-hinting=none', '--disable-font-subpixel-positioning']
        });

        const page = await browser.newPage();
        await page.setExtraHTTPHeaders({ 'Accept-Language': 'tr-TR,tr;q=0.9' });
        await page.setViewport({ width: 1400, height: 900, deviceScaleFactor: 1 });

        const fileUrl = 'file:///' + htmlDosya.replace(/\\/g, '/');
        await page.goto(fileUrl, { waitUntil: 'networkidle0', timeout: 60000 });
        await new Promise(r => setTimeout(r, 1500));

        await page.screenshot({ path: pngDosya, fullPage: true, type: 'png' });
        await browser.close();
        browser = null;

        const pngSize = fs.statSync(pngDosya).size;
        console.log(`PNG olusturuldu (${kaynak}), boyut: ${pngSize} byte`);

        if (pngSize > 1000) {
            const media = MessageMedia.fromFilePath(pngDosya);
            const tarih = new Date().toLocaleString('tr-TR');
            const baslik = kaynak === 'sql'
                ? `Yan Urunler + Seker Uretim-Satis-Stok Raporu\n${tarih}`
                : kaynak === 'pancar'
                    ? `Pancar Avans Raporu\n${tarih}`
                    : `Tum Rapor (Excel)\n${tarih}`;
            await message.reply(media, undefined, { caption: baslik });
            console.log('Rapor gonderildi!\n');
        } else {
            throw new Error('PNG dosyasi cok kucuk: ' + pngSize);
        }
    } catch (error) {
        console.error('PNG/Gonderim hatasi:', error.message);
        if (browser) { try { await browser.close(); } catch (_) {} }
        await message.reply('Rapor gonderilirken hata olustu: ' + error.message);
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
client.initialize();
