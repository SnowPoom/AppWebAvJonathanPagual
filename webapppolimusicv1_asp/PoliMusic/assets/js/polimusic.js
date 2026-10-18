/* =====================================================================
   PoliMusic - Interacciones de interfaz y reproductor global
   Vanilla JavaScript, sin dependencias.
   ===================================================================== */
(function () {
    'use strict';

    var PM = window.PM || {};
    PM.base = PM.base || '/';
    PM.csrf = PM.csrf || '';
    PM.playlists = PM.playlists || [];

    /* ---------------------------------------------------------------
       Utilidades
       --------------------------------------------------------------- */
    function $(sel, root) { return (root || document).querySelector(sel); }
    function $$(sel, root) { return Array.prototype.slice.call((root || document).querySelectorAll(sel)); }

    function icon(name) {
        var paths = {
            check: '<path d="m5 12.6 4.6 4.6L19 7.4"/>',
            alert: '<path d="M12 4.2 2.9 20h18.2z"/><path d="M12 9.6v4.4M12 17h.01"/>',
            info: '<circle cx="12" cy="12" r="9"/><path d="M12 11v5.2M12 7.8h.01"/>',
            plus: '<path d="M12 5.2v13.6M5.2 12h13.6"/>',
            music: '<path d="M9 17.6V6.2l10-2.1v11.4"/><circle cx="6.4" cy="18" r="2.6"/><circle cx="16.4" cy="15.9" r="2.6"/>'
        };
        return '<svg class="ico" viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="1.7" ' +
            'stroke-linecap="round" stroke-linejoin="round" aria-hidden="true">' + (paths[name] || paths.info) + '</svg>';
    }

    function fmtTime(sec) {
        if (!isFinite(sec) || sec < 0) sec = 0;
        var m = Math.floor(sec / 60);
        var s = Math.floor(sec % 60);
        return m + ':' + (s < 10 ? '0' : '') + s;
    }

    /* ---------------------------------------------------------------
       Toasts
       --------------------------------------------------------------- */
    var toastHost = $('#toasts');

    function toast(message, kind) {
        if (!toastHost) return;
        var el = document.createElement('div');
        el.className = 'toast' + (kind === 'error' ? ' toast--error' : '');
        el.innerHTML = icon(kind === 'error' ? 'alert' : 'check') + '<span></span>';
        el.querySelector('span').textContent = message;
        toastHost.appendChild(el);
        setTimeout(function () {
            el.classList.add('is-out');
            setTimeout(function () { el.remove(); }, 260);
        }, 3200);
    }

    PM.toast = toast;

    /* ---------------------------------------------------------------
       Modal generico
       --------------------------------------------------------------- */
    var modalRoot = $('#modalRoot');
    var modalTitle = $('#modalTitle');
    var modalBody = $('#modalBody');
    var modalOnClose = null;

    function openModal(opts) {
        if (!modalRoot) return;
        modalTitle.textContent = opts.title || '';
        modalBody.innerHTML = opts.html || '';
        modalRoot.hidden = false;
        document.body.style.overflow = 'hidden';
        modalOnClose = opts.onMount ? (opts.onMount(modalBody) || null) : null;
        var first = modalBody.querySelector('input,select,textarea,button');
        if (first) setTimeout(function () { first.focus(); }, 60);
    }

    function closeModal() {
        if (!modalRoot || modalRoot.hidden) return;
        modalRoot.hidden = true;
        modalBody.innerHTML = '';
        document.body.style.overflow = '';
        if (typeof modalOnClose === 'function') modalOnClose();
        modalOnClose = null;
    }

    PM.modal = { open: openModal, close: closeModal };

    if (modalRoot) {
        modalRoot.addEventListener('click', function (e) {
            if (e.target.hasAttribute('data-close')) closeModal();
        });
        document.addEventListener('keydown', function (e) {
            if (e.key === 'Escape') closeModal();
        });
    }

    /* Confirmacion personalizada */
    function confirmDialog(message, opts) {
        opts = opts || {};
        return new Promise(function (resolve) {
            var decided = false;
            openModal({
                title: opts.title || 'Confirmar acción',
                html: '<p class="muted" style="font-size:13.5px;line-height:1.6"></p>' +
                    '<div class="form-actions"><button class="btn btn--danger" data-yes type="button">' +
                    (opts.confirm || 'Eliminar') + '</button>' +
                    '<button class="btn btn--ghost" data-close type="button">Cancelar</button></div>',
                onMount: function (body) {
                    body.querySelector('p').textContent = message;
                    body.querySelector('[data-yes]').addEventListener('click', function () {
                        decided = true;
                        closeModal();
                        resolve(true);
                    });
                    return function () { if (!decided) resolve(false); };
                }
            });
        });
    }

    /* ---------------------------------------------------------------
       Enlaces de eliminacion con confirmacion
       --------------------------------------------------------------- */
    $$('[data-confirm]').forEach(function (el) {
        el.addEventListener('click', function (e) {
            e.preventDefault();
            confirmDialog(el.getAttribute('data-confirm'), {
                title: el.getAttribute('data-confirm-title') || 'Confirmar acción',
                confirm: el.getAttribute('data-confirm-ok') || 'Eliminar'
            }).then(function (ok) {
                if (!ok) return;
                if (el.tagName === 'A') {
                    window.location.href = el.getAttribute('href');
                } else if (el.form) {
                    el.form.submit();
                }
            });
        });
    });

    /* ---------------------------------------------------------------
       Navegacion lateral movil
       --------------------------------------------------------------- */
    var sidebar = $('#sidebar');
    var scrim = $('#scrim');
    var navToggle = $('#navToggle');

    function setNav(open) {
        if (!sidebar) return;
        sidebar.classList.toggle('is-open', open);
        if (scrim) scrim.hidden = !open;
        document.body.style.overflow = open ? 'hidden' : '';
    }

    if (navToggle) navToggle.addEventListener('click', function () { setNav(!sidebar.classList.contains('is-open')); });
    if (scrim) scrim.addEventListener('click', function () { setNav(false); });

    /* ---------------------------------------------------------------
       Menu de usuario y cierre de alertas
       --------------------------------------------------------------- */
    var userChip = $('#userChip');
    var userDropdown = $('#userDropdown');

    if (userChip && userDropdown) {
        userChip.addEventListener('click', function (e) {
            e.stopPropagation();
            var open = userDropdown.hidden;
            userDropdown.hidden = !open;
            userChip.setAttribute('aria-expanded', String(open));
        });
        document.addEventListener('click', function (e) {
            if (!userDropdown.hidden && !userDropdown.contains(e.target)) {
                userDropdown.hidden = true;
                userChip.setAttribute('aria-expanded', 'false');
            }
        });
    }

    $$('.alert__close').forEach(function (btn) {
        btn.addEventListener('click', function () {
            var a = btn.closest('.alert');
            if (a) a.remove();
        });
    });

    /* ---------------------------------------------------------------
       Reproductor
       --------------------------------------------------------------- */
    var audio = $('#pmAudio');
    var player = $('#player');
    if (!audio || !player) return;

    var elCover = $('#plCover');
    var elTitle = $('#plTitle');
    var elArtist = $('#plArtist');
    var elCur = $('#plCur');
    var elDur = $('#plDur');
    var elFill = $('#plFill');
    var elProgress = $('#plProgress');
    var btnPlay = $('#btnPlay');
    var btnPrev = $('#btnPrev');
    var btnNext = $('#btnNext');
    var btnShuffle = $('#btnShuffle');
    var btnRepeat = $('#btnRepeat');
    var btnMute = $('#btnMute');
    var volRange = $('#volRange');
    var volIco = $('#volIco');
    var queuePanel = $('#queuePanel');
    var queueList = $('#queueList');

    var state = {
        queue: [],
        index: -1,
        current: null,
        shuffle: localStorage.getItem('pm.shuffle') === '1',
        repeat: localStorage.getItem('pm.repeat') || 'off',
        volume: parseFloat(localStorage.getItem('pm.volume')),
        muted: false,
        errors: 0
    };

    if (!isFinite(state.volume)) state.volume = 0.8;

    /* --- Construccion de la cola a partir de las filas visibles --- */
    function buildQueue() {
        state.queue = $$('.song-row[data-song-id]').map(function (row) {
            var coverEl = row.querySelector('.song-row__cover .cover');
            return {
                id: row.getAttribute('data-song-id'),
                title: row.getAttribute('data-title') || '',
                artist: row.getAttribute('data-artist') || '',
                src: row.getAttribute('data-src') || '',
                album: row.getAttribute('data-album') || '',
                available: row.getAttribute('data-available') === '1',
                cover: coverEl ? coverEl.outerHTML : '',
                row: row
            };
        });
    }

    function findRow(id) {
        return document.querySelector('.song-row[data-song-id="' + id + '"]');
    }

    function markPlaying(id) {
        $$('.song-row').forEach(function (r) {
            r.classList.toggle('is-playing', r.getAttribute('data-song-id') === String(id));
        });
        renderQueue();
    }

    function setCover(html, label) {
        if (html) {
            elCover.innerHTML = html;
        } else {
            elCover.innerHTML = '<span class="cover cover--fb" style="--c1:#22d3ee;--c2:#34d399"><i>' +
                (label ? label.substring(0, 2).toUpperCase() : 'PM') + '</i></span>';
        }
    }

    function updatePlayUI(playing) {
        player.classList.toggle('is-playing', playing);
        btnPlay.setAttribute('aria-label', playing ? 'Pausar' : 'Reproducir');
        btnPlay.setAttribute('title', playing ? 'Pausar' : 'Reproducir');
    }

    function renderQueue() {
        if (!queueList) return;
        if (!state.queue.length) {
            queueList.innerHTML = '<div class="empty" style="padding:26px 14px"><strong>Cola vacía</strong>' +
                '<p>Abre una lista de canciones para cargar la cola.</p></div>';
            return;
        }
        queueList.innerHTML = '';
        state.queue.forEach(function (item, i) {
            var el = document.createElement('div');
            el.className = 'queue__item' + (i === state.index ? ' is-current' : '');
            el.innerHTML = (item.cover || '') + '<div><strong></strong><span></span></div>';
            el.querySelector('strong').textContent = item.title;
            el.querySelector('span').textContent = item.artist || (item.available ? 'PoliMusic' : 'Sin audio');
            el.addEventListener('click', function () { playIndex(i); });
            queueList.appendChild(el);
        });
    }

    function incrementPlays(id) {
        if (!id) return;
        fetch(PM.base + 'api/play.asp?id=' + encodeURIComponent(id), {
            method: 'POST',
            headers: { 'X-CSRF-Token': PM.csrf },
            credentials: 'same-origin'
        }).then(function (r) { return r.json(); }).then(function (data) {
            if (!data || !data.ok) return;
            var row = findRow(id);
            if (row) {
                var cell = row.querySelector('.song-row__plays');
                if (cell) cell.innerHTML = cell.innerHTML.replace(/\d+/, data.plays);
            }
        }).catch(function () { /* silencioso */ });
    }

    function playIndex(i, autoplay) {
        if (i < 0 || i >= state.queue.length) return;
        var item = state.queue[i];
        if (!item.available || !item.src) {
            toast('El archivo de audio de "' + item.title + '" no está disponible.', 'error');
            return;
        }
        state.index = i;
        state.current = item;
        audio.src = item.src;
        audio.load();
        setCover(item.cover, item.title);
        elTitle.textContent = item.title;
        elArtist.textContent = item.artist || 'Artista desconocido';
        document.title = item.title + ' \u00b7 PoliMusic';
        markPlaying(item.id);
        if (autoplay !== false) {
            var p = audio.play();
            if (p && p.catch) p.catch(function () { updatePlayUI(false); });
        }
        incrementPlays(item.id);
        updateMediaSession(item);
        saveSession(true);
    }

    function updateMediaSession(item) {
        if (!('mediaSession' in navigator)) return;
        try {
            navigator.mediaSession.metadata = new window.MediaMetadata({
                title: item.title,
                artist: item.artist || 'PoliMusic',
                album: item.album || 'PoliMusic',
                artwork: []
            });
        } catch (e) { /* navegador sin soporte */ }
    }

    function togglePlay() {
        if (!audio.src) {
            if (state.queue.length) {
                playIndex(0);
            } else {
                toast('No hay canciones en esta vista.', 'error');
            }
            return;
        }
        if (audio.paused) {
            var p = audio.play();
            if (p && p.catch) p.catch(function () { updatePlayUI(false); });
        } else {
            audio.pause();
        }
    }

    function nextIndexFrom(from, step) {
        if (!state.queue.length) return -1;
        if (state.shuffle && state.queue.length > 1) {
            var r = from;
            while (r === from) r = Math.floor(Math.random() * state.queue.length);
            return r;
        }
        var n = from + step;
        if (n >= state.queue.length) return state.repeat === 'all' ? 0 : -1;
        if (n < 0) return state.repeat === 'all' ? state.queue.length - 1 : -1;
        return n;
    }

    /* Avanza saltando las pistas sin archivo de audio */
    function goNext(auto) {
        var i = state.index, guard = 0, found = -1;
        while (guard < state.queue.length) {
            var n = nextIndexFrom(i, 1);
            if (n < 0) break;
            i = n;
            guard++;
            if (state.queue[i].available) { found = i; break; }
        }
        if (found < 0) {
            updatePlayUI(false);
            if (!auto) toast('No hay más canciones disponibles en esta lista.', 'error');
            return;
        }
        playIndex(found);
    }

    function goPrev() {
        if (audio.currentTime > 3) { audio.currentTime = 0; return; }
        var i = state.index, guard = 0, found = -1;
        while (guard < state.queue.length) {
            var n = nextIndexFrom(i, -1);
            if (n < 0) break;
            i = n;
            guard++;
            if (state.queue[i].available) { found = i; break; }
        }
        if (found < 0) { audio.currentTime = 0; return; }
        playIndex(found);
    }

    /* --- Controles --- */
    btnPlay.addEventListener('click', togglePlay);
    btnNext.addEventListener('click', function () { goNext(false); });
    btnPrev.addEventListener('click', goPrev);

    btnShuffle.addEventListener('click', function () {
        state.shuffle = !state.shuffle;
        btnShuffle.classList.toggle('is-active', state.shuffle);
        localStorage.setItem('pm.shuffle', state.shuffle ? '1' : '0');
        toast(state.shuffle ? 'Modo aleatorio activado.' : 'Modo aleatorio desactivado.');
    });

    function applyRepeat() {
        btnRepeat.classList.toggle('is-active', state.repeat !== 'off');
        btnRepeat.title = state.repeat === 'one' ? 'Repetir una' : (state.repeat === 'all' ? 'Repetir todo' : 'Repetir');
        btnRepeat.setAttribute('data-mode', state.repeat);
    }

    btnRepeat.addEventListener('click', function () {
        state.repeat = state.repeat === 'off' ? 'all' : (state.repeat === 'all' ? 'one' : 'off');
        localStorage.setItem('pm.repeat', state.repeat);
        applyRepeat();
        toast(state.repeat === 'off' ? 'Repetición desactivada.' :
            (state.repeat === 'all' ? 'Repetir toda la cola.' : 'Repetir la canción actual.'));
    });

    btnShuffle.classList.toggle('is-active', state.shuffle);
    applyRepeat();

    /* --- Volumen --- */
    function applyVolume() {
        audio.volume = state.muted ? 0 : state.volume;
        audio.muted = state.muted;
        if (volRange) {
            volRange.value = state.volume;
            volRange.style.setProperty('--vol', (state.volume * 100) + '%');
        }
        if (volIco) {
            var muted = state.muted || state.volume === 0;
            volIco.innerHTML = '<svg class="ico" viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="1.7" ' +
                'stroke-linecap="round" stroke-linejoin="round" aria-hidden="true">' +
                (muted
                    ? '<path d="M4 9.4h3.2L11.4 6v12L7.2 14.6H4z"/><path d="m15.4 9.6 4.4 4.8M19.8 9.6l-4.4 4.8"/>'
                    : '<path d="M4 9.4h3.2L11.4 6v12L7.2 14.6H4z"/><path d="M14.8 9.2a4 4 0 0 1 0 5.6"/><path d="M17.4 6.6a7.6 7.6 0 0 1 0 10.8"/>') +
                '</svg>';
        }
    }

    if (volRange) {
        volRange.addEventListener('input', function () {
            state.volume = parseFloat(volRange.value);
            state.muted = false;
            localStorage.setItem('pm.volume', state.volume);
            applyVolume();
        });
    }

    btnMute.addEventListener('click', function () {
        state.muted = !state.muted;
        applyVolume();
    });

    applyVolume();

    /* --- Progreso --- */
    function paintProgress() {
        var d = audio.duration;
        var ratio = (isFinite(d) && d > 0) ? (audio.currentTime / d) : 0;
        var pct = Math.min(100, Math.max(0, ratio * 100));
        if (elFill) elFill.style.width = pct + '%';
        var knob = elProgress ? elProgress.querySelector('.progress__knob') : null;
        if (knob) knob.style.left = pct + '%';
        if (elProgress) elProgress.setAttribute('aria-valuenow', Math.round(pct));
        if (elCur) elCur.textContent = fmtTime(audio.currentTime);
    }

    audio.addEventListener('timeupdate', function () {
        paintProgress();
        if (Math.floor(audio.currentTime) % 3 === 0) saveSession(false);
    });

    audio.addEventListener('loadedmetadata', function () {
        if (elDur) elDur.textContent = fmtTime(audio.duration);
        paintProgress();
    });

    audio.addEventListener('play', function () {
        updatePlayUI(true);
        state.errors = 0;
    });

    audio.addEventListener('pause', function () { updatePlayUI(false); });

    audio.addEventListener('ended', function () {
        if (state.repeat === 'one') {
            audio.currentTime = 0;
            audio.play();
            return;
        }
        goNext(true);
    });

    audio.addEventListener('error', function () {
        updatePlayUI(false);
        if (!audio.src) return;
        toast('No se pudo reproducir el archivo de audio.', 'error');
        state.errors += 1;
        if (state.errors < 3) setTimeout(function () { goNext(true); }, 900);
    });

    if (elProgress) {
        var dragging = false;

        function seekFromEvent(e) {
            var rect = elProgress.getBoundingClientRect();
            var x = (e.touches ? e.touches[0].clientX : e.clientX) - rect.left;
            var ratio = Math.min(1, Math.max(0, x / rect.width));
            if (isFinite(audio.duration) && audio.duration > 0) {
                audio.currentTime = ratio * audio.duration;
                paintProgress();
            }
        }

        elProgress.addEventListener('pointerdown', function (e) {
            dragging = true;
            elProgress.classList.add('is-dragging');
            elProgress.setPointerCapture && elProgress.setPointerCapture(e.pointerId);
            seekFromEvent(e);
        });

        elProgress.addEventListener('pointermove', function (e) {
            if (dragging) seekFromEvent(e);
        });

        elProgress.addEventListener('pointerup', function () {
            dragging = false;
            elProgress.classList.remove('is-dragging');
        });

        elProgress.addEventListener('keydown', function (e) {
            if (e.key === 'ArrowRight') { audio.currentTime = Math.min(audio.duration || 0, audio.currentTime + 5); e.preventDefault(); }
            if (e.key === 'ArrowLeft') { audio.currentTime = Math.max(0, audio.currentTime - 5); e.preventDefault(); }
        });
    }

    /* --- Cola --- */
    var btnQueue = $('#btnQueue');
    var btnQueueClose = $('#btnQueueClose');

    if (btnQueue && queuePanel) {
        btnQueue.addEventListener('click', function (e) {
            e.stopPropagation();
            queuePanel.hidden = !queuePanel.hidden;
            btnQueue.classList.toggle('is-active', !queuePanel.hidden);
        });
        if (btnQueueClose) btnQueueClose.addEventListener('click', function () { queuePanel.hidden = true; btnQueue.classList.remove('is-active'); });
        document.addEventListener('click', function (e) {
            if (!queuePanel.hidden && !queuePanel.contains(e.target) && e.target !== btnQueue && !btnQueue.contains(e.target)) {
                queuePanel.hidden = true;
                btnQueue.classList.remove('is-active');
            }
        });
    }

    /* --- Atajos de teclado --- */
    document.addEventListener('keydown', function (e) {
        var tag = (e.target.tagName || '').toLowerCase();
        var typing = tag === 'input' || tag === 'textarea' || tag === 'select' || e.target.isContentEditable;
        if (typing) return;
        if (e.code === 'Space') { e.preventDefault(); togglePlay(); }
        else if (e.key === 'ArrowRight' && e.shiftKey) { audio.currentTime = Math.min(audio.duration || 0, audio.currentTime + 10); }
        else if (e.key === 'ArrowLeft' && e.shiftKey) { audio.currentTime = Math.max(0, audio.currentTime - 10); }
        else if (e.key === 'n' && e.altKey) { goNext(false); }
        else if (e.key === 'p' && e.altKey) { goPrev(); }
        else if (e.key === '/') {
            var s = $('#globalSearch');
            if (s) { e.preventDefault(); s.focus(); s.select(); }
        }
    });

    /* --- Media Session --- */
    if ('mediaSession' in navigator) {
        try {
            navigator.mediaSession.setActionHandler('play', function () { audio.play(); });
            navigator.mediaSession.setActionHandler('pause', function () { audio.pause(); });
            navigator.mediaSession.setActionHandler('nexttrack', function () { goNext(false); });
            navigator.mediaSession.setActionHandler('previoustrack', goPrev);
        } catch (e) { /* sin soporte */ }
    }

    /* --- Persistencia de la sesion de escucha --- */
    var lastSave = 0;

    function saveSession(force) {
        var now = Date.now();
        if (!force && now - lastSave < 2500) return;
        lastSave = now;
        var item = state.current;
        if (!item) return;
        try {
            sessionStorage.setItem('pm.now', JSON.stringify({
                id: item.id,
                title: item.title,
                artist: item.artist,
                album: item.album,
                src: audio.currentSrc || item.src,
                cover: item.cover,
                time: audio.currentTime || 0
            }));
        } catch (e) { /* almacenamiento lleno o bloqueado */ }
    }

    function restoreSession() {
        var raw = null;
        try { raw = sessionStorage.getItem('pm.now'); } catch (e) { return; }
        if (!raw) return;
        var s;
        try { s = JSON.parse(raw); } catch (e) { return; }
        if (!s || !s.src) return;

        var inQueue = -1;
        for (var i = 0; i < state.queue.length; i++) {
            if (String(state.queue[i].id) === String(s.id)) { inQueue = i; break; }
        }
        state.current = s;
        state.index = inQueue;
        audio.src = s.src;
        setCover(s.cover, s.title);
        elTitle.textContent = s.title || 'Sin reproducción';
        elArtist.textContent = s.artist || 'Artista desconocido';
        audio.addEventListener('loadedmetadata', function () {
            if (s.time > 0 && s.time < (audio.duration || 0) - 1) {
                try { audio.currentTime = s.time; } catch (e) { /* ignorar */ }
            }
            if (elDur) elDur.textContent = fmtTime(audio.duration);
            paintProgress();
        }, { once: true });
        if (inQueue >= 0) markPlaying(s.id);
    }

    window.addEventListener('beforeunload', function () { saveSession(true); });

    /* --- Clic sobre las filas --- */
    $$('.song-row[data-song-id]').forEach(function (row) {
        row.setAttribute('tabindex', '0');
        row.setAttribute('role', 'button');

        function activate() {
            var id = row.getAttribute('data-song-id');
            var idx = -1;
            for (var i = 0; i < state.queue.length; i++) {
                if (state.queue[i].id === id) { idx = i; break; }
            }
            if (idx < 0) return;
            if (idx === state.index && !audio.paused) {
                audio.pause();
            } else {
                playIndex(idx);
            }
        }

        row.addEventListener('click', function (e) {
            if (e.target.closest('a,button,.no-play')) return;
            activate();
        });

        row.addEventListener('keydown', function (e) {
            if (e.key === 'Enter') { e.preventDefault(); activate(); }
        });
    });

    /* ---------------------------------------------------------------
       Agregar canciones a una playlist
       --------------------------------------------------------------- */
    function openPlaylistPicker(songId, songTitle) {
        if (!PM.playlists.length) {
            openModal({
                title: 'Agregar a playlist',
                html: '<p class="muted" style="font-size:13.5px;line-height:1.6;margin-bottom:16px">' +
                    'Todavía no tienes playlists creadas. Crea una para empezar a organizar tu musica.</p>' +
                    '<a class="btn btn--primary btn--block" href="' + PM.base + 'playlists.asp?nueva=1">Crear mi primera playlist</a>'
            });
            return;
        }

        var html = '<p class="muted" style="font-size:12.5px;margin-bottom:14px">Selecciona la playlist donde quieres guardar <b style="color:#e8f4ff">' +
            escapeHtml(songTitle) + '</b>.</p><div class="pick-list">';
        PM.playlists.forEach(function (pl) {
            html += '<button class="pick" type="button" data-pl="' + pl.id + '">' +
                icon('music') + '<span class="pick__meta"><strong>' + escapeHtml(pl.title) + '</strong>' +
                '<span>Playlist personal</span></span></button>';
        });
        html += '</div>';

        openModal({
            title: 'Agregar a playlist',
            html: html,
            onMount: function (body) {
                $$('.pick', body).forEach(function (btn) {
                    btn.addEventListener('click', function () {
                        addToPlaylist(btn.getAttribute('data-pl'), songId, btn);
                    });
                });
            }
        });
    }

    function escapeHtml(s) {
        return String(s || '').replace(/[&<>"']/g, function (c) {
            return { '&': '&amp;', '<': '&lt;', '>': '&gt;', '"': '&quot;', "'": '&#39;' }[c];
        });
    }

    function addToPlaylist(playlistId, songId, btn) {
        btn.disabled = true;
        var body = new FormData();
        body.append('playlist', playlistId);
        body.append('song', songId);
        body.append('csrf', PM.csrf);

        fetch(PM.base + 'api/playlist_add.asp', {
            method: 'POST',
            body: body,
            credentials: 'same-origin'
        }).then(function (r) { return r.json(); }).then(function (data) {
            if (data && data.ok) {
                toast(data.message || 'Canción agregada a la playlist.');
                closeModal();
            } else {
                toast((data && data.message) || 'No se pudo agregar la canción.', 'error');
                btn.disabled = false;
            }
        }).catch(function () {
            toast('Error de conexión con el servidor.', 'error');
            btn.disabled = false;
        });
    }

    $$('[data-add-to]').forEach(function (btn) {
        btn.addEventListener('click', function (e) {
            e.stopPropagation();
            openPlaylistPicker(btn.getAttribute('data-add-to'), btn.getAttribute('data-song-title') || '');
        });
    });

    /* Selector visual de rutas: clic en miniatura -> rellena el input */
    $$('[data-pick-grid]').forEach(function (grid) {
        var input = document.querySelector(grid.getAttribute('data-target'));
        if (!input) return;
        var buttons = $$('.thumb-pick', grid);

        function sync() {
            buttons.forEach(function (b) {
                b.classList.toggle('is-selected', b.getAttribute('data-path') === input.value);
            });
        }

        buttons.forEach(function (b) {
            b.addEventListener('click', function () {
                input.value = b.getAttribute('data-path');
                sync();
            });
        });

        input.addEventListener('input', sync);
        sync();
    });

    /* Boton "reproducir todo": arranca la primera canción disponible */
    $$('[data-play-all]').forEach(function (btn) {
        btn.addEventListener('click', function () {
            var first = -1;
            for (var i = 0; i < state.queue.length; i++) {
                if (state.queue[i].available) { first = i; break; }
            }
            if (first < 0) { toast('No hay canciones disponibles para reproducir.', 'error'); return; }
            playIndex(first);
        });
    });

    /* ---------------------------------------------------------------
       Inicio
       --------------------------------------------------------------- */
    buildQueue();
    renderQueue();
    restoreSession();

    PM.player = {
        play: function (id) {
            var idx = -1;
            for (var i = 0; i < state.queue.length; i++) {
                if (String(state.queue[i].id) === String(id)) { idx = i; break; }
            }
            if (idx >= 0) playIndex(idx);
        },
        queue: function () { return state.queue; }
    };
})();
