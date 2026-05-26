// Raporlama Portali - JavaScript Yardımcı Fonksiyonları

/**
 * Base64 encoded dosyayı indirir
 * @param {string} fileName - Dosya adı
 * @param {string} base64Content - Base64 içerik
 * @param {string} mimeType - MIME tipi
 */
function downloadFile(fileName, base64Content, mimeType) {
    const link = document.createElement('a');
    link.href = `data:${mimeType};base64,${base64Content}`;
    link.download = fileName;
    document.body.appendChild(link);
    link.click();
    document.body.removeChild(link);
}

/**
 * .NET 8 Blazor Server DotNetStreamReference'ten dosya indirir
 * @param {string} fileName
 * @param {object} contentStreamReference - Blazor DotNetStreamReference
 */
async function downloadFileFromStream(fileName, contentStreamReference) {
    const arrayBuffer = await contentStreamReference.arrayBuffer();
    const blob = new Blob([arrayBuffer]);
    const url = URL.createObjectURL(blob);
    const a = document.createElement('a');
    a.href = url;
    a.download = fileName ?? '';
    document.body.appendChild(a);
    a.click();
    a.remove();
    URL.revokeObjectURL(url);
}
window.downloadFileFromStream = downloadFileFromStream;
window.downloadFile = downloadFile;

/**
 * Tabloyu yazdırır
 * @param {string} elementId - Yazdırılacak element ID
 */
function printElement(elementId) {
    const element = document.getElementById(elementId);
    if (!element) return;
    
    const printWindow = window.open('', '_blank');
    printWindow.document.write(`
        <html>
        <head>
            <title>Rapor Yazdır</title>
            <link href="https://fonts.googleapis.com/css2?family=Roboto:wght@400;500;700&display=swap" rel="stylesheet">
            <style>
                body { font-family: 'Roboto', sans-serif; padding: 20px; }
                table { width: 100%; border-collapse: collapse; }
                th, td { border: 1px solid #ddd; padding: 8px; text-align: left; }
                th { background-color: #1B5E20; color: white; }
                tr:nth-child(even) { background-color: #f9f9f9; }
                .header { text-align: center; margin-bottom: 20px; }
                .footer { margin-top: 20px; text-align: right; font-size: 12px; color: #666; }
            </style>
        </head>
        <body>
            <div class="header">
                <h2>Doğuş Çay - Afyon Şeker Fabrikası</h2>
                <p>Rapor Tarihi: ${new Date().toLocaleDateString('tr-TR')}</p>
            </div>
            ${element.outerHTML}
            <div class="footer">
                <p>Bu rapor Raporlama Portali tarafından oluşturulmuştur.</p>
            </div>
        </body>
        </html>
    `);
    printWindow.document.close();
    printWindow.print();
}

/**
 * Clipboard'a kopyalar
 * @param {string} text - Kopyalanacak metin
 */
async function copyToClipboard(text) {
    try {
        await navigator.clipboard.writeText(text);
        return true;
    } catch (err) {
        console.error('Kopyalama hatası:', err);
        return false;
    }
}

/**
 * HTML içeriği clipboard'a hem HTML hem düz metin olarak yazar.
 * Outlook/Gmail/Word yapıştırırken tablo formatı korunur.
 * Outlook için copy event + clipboardData yöntemi en uyumlu olanı, onu önce dener.
 */
async function copyHtmlToClipboard(html, plainText) {
    plainText = plainText || html.replace(/<[^>]+>/g, '').trim();

    // YÖNTEM 1: copy event + clipboardData.setData (Outlook ile en uyumlu)
    try {
        const onCopy = (e) => {
            e.preventDefault();
            e.clipboardData.setData('text/html', html);
            e.clipboardData.setData('text/plain', plainText);
        };
        document.addEventListener('copy', onCopy);
        const ok = document.execCommand('copy');
        document.removeEventListener('copy', onCopy);
        if (ok) return true;
    } catch (e1) {
        console.warn('Yöntem 1 (copy event) başarısız:', e1);
    }

    // YÖNTEM 2: contenteditable div + execCommand('copy')
    try {
        const div = document.createElement('div');
        div.contentEditable = 'true';
        div.style.position = 'fixed';
        div.style.left = '-9999px';
        div.style.top = '0';
        div.innerHTML = html;
        document.body.appendChild(div);
        const range = document.createRange();
        range.selectNodeContents(div);
        const sel = window.getSelection();
        sel.removeAllRanges();
        sel.addRange(range);
        const ok = document.execCommand('copy');
        sel.removeAllRanges();
        document.body.removeChild(div);
        if (ok) return true;
    } catch (e2) {
        console.warn('Yöntem 2 (contenteditable) başarısız:', e2);
    }

    // YÖNTEM 3: modern Clipboard API (Outlook eski sürümleri tanımayabilir)
    try {
        const htmlBlob = new Blob([html], { type: 'text/html' });
        const textBlob = new Blob([plainText], { type: 'text/plain' });
        await navigator.clipboard.write([
            new ClipboardItem({ 'text/html': htmlBlob, 'text/plain': textBlob })
        ]);
        return true;
    } catch (e3) {
        console.error('Yöntem 3 (Clipboard API) başarısız:', e3);
        return false;
    }
}
window.copyHtmlToClipboard = copyHtmlToClipboard;
