// Web Speech API ile Türkçe konuşma → metin.
// Chrome, Edge, Safari'de yerleşik. Firefox'ta yok (tarayıcı eklentisi gerekir).
// Mikrofon izni ilk kullanımda sorulur.

(function () {
    'use strict';

    let aktifTanima = null;
    let dotnetRef   = null;
    let dinleniyor  = false;

    /**
     * Tarayıcı desteğini kontrol eder. true → Web Speech API mevcut.
     */
    window.sesliKomutDesteklimi = function () {
        return !!(window.SpeechRecognition || window.webkitSpeechRecognition);
    };

    /**
     * Türkçe dinlemeyi başlatır. Bittiğinde dotnetRef.invokeMethodAsync('SesYakalandi', metin)
     * çağrılır. Hata olursa 'SesHata' çağrılır.
     *
     * @param {object} ref Blazor DotNetObjectReference
     */
    window.sesliKomutBaslat = function (ref) {
        if (dinleniyor) {
            console.warn('Zaten dinleniyor');
            return false;
        }
        if (!window.sesliKomutDesteklimi()) {
            ref.invokeMethodAsync('SesHata', 'Tarayıcı konuşma tanımayı desteklemiyor.');
            return false;
        }

        const SR = window.SpeechRecognition || window.webkitSpeechRecognition;
        const r = new SR();
        r.lang = 'tr-TR';
        r.continuous = false;       // tek bir komut yakala, dur
        r.interimResults = true;    // ara sonuçlar UI'da gösterilebilir
        r.maxAlternatives = 1;

        dotnetRef = ref;
        aktifTanima = r;
        dinleniyor = true;

        let sonMetin = '';

        r.onstart = function () {
            try { dotnetRef.invokeMethodAsync('DinlemeBasladi'); } catch (e) { console.warn(e); }
        };

        r.onresult = function (ev) {
            let interim = '';
            let final = '';
            for (let i = ev.resultIndex; i < ev.results.length; i++) {
                const sonuc = ev.results[i];
                if (sonuc.isFinal) final += sonuc[0].transcript;
                else                interim += sonuc[0].transcript;
            }
            sonMetin = final || interim;
            try { dotnetRef.invokeMethodAsync('AraSonuc', sonMetin, !!final); } catch (e) { console.warn(e); }
        };

        r.onerror = function (ev) {
            const aciklama = {
                'no-speech':       'Ses algılanmadı.',
                'audio-capture':   'Mikrofon erişilemiyor.',
                'not-allowed':     'Mikrofon izni reddedildi.',
                'aborted':         'İptal edildi.',
                'network':         'Ağ hatası.',
                'service-not-allowed': 'Tarayıcı servisi izin vermiyor.',
                'language-not-supported': 'Türkçe desteklenmiyor.',
            }[ev.error] || ev.error || 'Bilinmeyen hata';
            dinleniyor = false;
            aktifTanima = null;
            try { dotnetRef.invokeMethodAsync('SesHata', aciklama); } catch (e) { console.warn(e); }
        };

        r.onend = function () {
            dinleniyor = false;
            aktifTanima = null;
            if (sonMetin) {
                try { dotnetRef.invokeMethodAsync('SesYakalandi', sonMetin); } catch (e) { console.warn(e); }
            } else {
                try { dotnetRef.invokeMethodAsync('SesYakalandi', ''); } catch (e) { console.warn(e); }
            }
        };

        try {
            r.start();
            return true;
        } catch (e) {
            console.error('SpeechRecognition start hata:', e);
            dinleniyor = false;
            aktifTanima = null;
            try { ref.invokeMethodAsync('SesHata', 'Başlatılamadı: ' + e.message); } catch (e2) { }
            return false;
        }
    };

    /**
     * Dinlemeyi manuel durdurur (örn iptal). Mevcut ses bitince onend tetiklenir.
     */
    window.sesliKomutDurdur = function () {
        if (aktifTanima) {
            try { aktifTanima.stop(); } catch (e) { /* ignore */ }
        }
    };

    /**
     * Metni sesli okur (Türkçe). Onay metnini kullanıcıya hem yazılı hem sesli sunmak için.
     * Sesli okuma opsiyonel; başarısız olursa sessizce geçer.
     */
    window.sesliOku = function (metin) {
        if (!('speechSynthesis' in window)) return;
        try {
            const u = new SpeechSynthesisUtterance(metin);
            u.lang = 'tr-TR';
            u.rate = 1.05;
            window.speechSynthesis.cancel();
            window.speechSynthesis.speak(u);
        } catch (e) { console.warn('TTS hata:', e); }
    };
})();
