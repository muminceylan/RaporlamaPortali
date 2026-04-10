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
