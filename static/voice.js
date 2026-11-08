/* ------------------------------------------------------------
   Shared voice guide for all kiosk pages
   - Amharic : your own recordings in static/audio/Am/*.m4a
   - English : browser AI voice (speechSynthesis)
------------------------------------------------------------ */
(function () {
    const AUDIO_BASE = '/static/audio/Am/';   // keep the capital "A" in "Am" (Vercel is case-sensitive)
    let currentAudio = null;
    let playToken = 0;                         // lets a new click cancel an older chain of recordings

    function stopAllVoice() {
        playToken++;
        if ('speechSynthesis' in window) window.speechSynthesis.cancel();
        if (currentAudio) {
            currentAudio.pause();
            currentAudio = null;
        }
    }

    // Plays one or more Amharic recordings in order, e.g. playAmharic('Welcome', 'Language').
    // Returns a Promise that resolves when they finish (or fail).
    function playAmharic(...names) {
        stopAllVoice();
        const token = playToken;

        return names.reduce((chain, name) => chain.then(() => new Promise(resolve => {
            if (token !== playToken) return resolve();      // cancelled by a newer click
            const audio = new Audio(`${AUDIO_BASE}${name}.m4a`);
            currentAudio = audio;
            audio.onended = resolve;
            audio.onerror = () => { console.warn('Missing audio file:', name); resolve(); };
            audio.play().catch(err => { console.warn('Audio blocked:', err); resolve(); });
        })), Promise.resolve());
    }

    function playEnglish(text) {
        stopAllVoice();
        if (!('speechSynthesis' in window)) {
            alert('Your browser does not support text-to-speech.');
            return;
        }
        const utterance = new SpeechSynthesisUtterance(text);
        utterance.lang = 'en-US';
        utterance.rate = 0.92;
        window.speechSynthesis.speak(utterance);
    }

    window.stopAllVoice = stopAllVoice;
    window.playAmharic = playAmharic;
    window.playEnglish = playEnglish;
})();