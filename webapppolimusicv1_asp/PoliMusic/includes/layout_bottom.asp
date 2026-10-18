<%
' =====================================================================
'  PoliMusic - Cierre de layout + reproductor global
' =====================================================================
Dim plRs, plJson
plJson = ""
If IsLoggedIn() Then
    Set plRs = DbQuery("SELECT ID_PLAYLIST, TITLE FROM TBL_PLAYLIST WHERE ID_USER = ? ORDER BY TITLE", Array(CurrentUserId()))
    Do Until plRs.EOF
        If plJson <> "" Then plJson = plJson & ","
        plJson = plJson & "{""id"":" & plRs("ID_PLAYLIST") & ",""title"":" & JsonStr(plRs("TITLE")) & "}"
        plRs.MoveNext
    Loop
    plRs.Close
    Set plRs = Nothing
End If
%>
        </main>
    </div>
</div>

<section class="player glass" id="player" aria-label="Reproductor de música">
    <div class="player__now">
        <span class="player__cover" id="plCover">
            <span class="cover cover--fb" style="--c1:#22d3ee;--c2:#34d399"><i>PM</i></span>
        </span>
        <div class="player__meta">
            <strong id="plTitle">Sin reproducción</strong>
            <span id="plArtist">Elige una canción para empezar</span>
        </div>
        <span class="eq" id="plEq" aria-hidden="true"><i></i><i></i><i></i><i></i></span>
    </div>

    <div class="player__center">
        <div class="player__controls">
            <button class="icon-btn icon-btn--ghost" id="btnShuffle" type="button" title="Aleatorio" aria-label="Aleatorio"><%=Icon("shuffle", "")%></button>
            <button class="icon-btn" id="btnPrev" type="button" title="Anterior" aria-label="Anterior"><%=Icon("prev", "")%></button>
            <button class="player__play" id="btnPlay" type="button" title="Reproducir" aria-label="Reproducir">
                <span class="player__play-ico player__play-ico--play"><%=Icon("play", "")%></span>
                <span class="player__play-ico player__play-ico--pause"><%=Icon("pause", "")%></span>
            </button>
            <button class="icon-btn" id="btnNext" type="button" title="Siguiente" aria-label="Siguiente"><%=Icon("next", "")%></button>
            <button class="icon-btn icon-btn--ghost" id="btnRepeat" type="button" title="Repetir" aria-label="Repetir"><%=Icon("repeat", "")%></button>
        </div>
        <div class="player__timeline">
            <span class="player__time" id="plCur">0:00</span>
            <div class="progress" id="plProgress" role="slider" tabindex="0" aria-label="Progreso de la canción"
                 aria-valuemin="0" aria-valuemax="100" aria-valuenow="0">
                <span class="progress__track"><span class="progress__fill" id="plFill"></span></span>
                <span class="progress__knob"></span>
            </div>
            <span class="player__time" id="plDur">0:00</span>
        </div>
    </div>

    <div class="player__right">
        <button class="icon-btn icon-btn--ghost" id="btnQueue" type="button" title="Cola de reproducción" aria-label="Cola de reproducción"><%=Icon("queue", "")%></button>
        <div class="vol">
            <button class="icon-btn icon-btn--ghost" id="btnMute" type="button" title="Silenciar" aria-label="Silenciar">
                <span id="volIco"><%=Icon("volume", "")%></span>
            </button>
            <input type="range" id="volRange" class="vol__range" min="0" max="1" step="0.01" value="0.8" aria-label="Volumen">
        </div>
    </div>

    <audio id="pmAudio" preload="metadata"></audio>
</section>

<aside class="queue glass" id="queuePanel" hidden aria-label="Cola de reproducción">
    <header class="queue__head">
        <h3>Cola de reproducción</h3>
        <button class="icon-btn icon-btn--ghost" id="btnQueueClose" type="button" aria-label="Cerrar">&times;</button>
    </header>
    <div class="queue__list" id="queueList"></div>
</aside>

<div class="toasts" id="toasts" aria-live="polite"></div>

<div class="modal" id="modalRoot" hidden>
    <div class="modal__scrim" data-close="1"></div>
    <div class="modal__box glass" role="dialog" aria-modal="true" aria-labelledby="modalTitle">
        <header class="modal__head">
            <h3 id="modalTitle">Título</h3>
            <button class="icon-btn icon-btn--ghost" data-close="1" type="button" aria-label="Cerrar">&times;</button>
        </header>
        <div class="modal__body" id="modalBody"></div>
    </div>
</div>

<script>
window.PM = {
    base: <%=JsonStr(AppBase())%>,
    csrf: <%=JsonStr(CsrfToken())%>,
    userId: <%=CurrentUserId()%>,
    userName: <%=JsonStr(Session("Usuario") & "")%>,
    isAdmin: <%=IIf(IsAdmin(), "true", "false")%>,
    playlists: [<%=plJson%>]
};
</script>
<script src="<%=AppBase()%>assets/js/polimusic.js?v=<%=APP_VERSION%>"></script>
</body>
</html>
<% DbClose() %>
