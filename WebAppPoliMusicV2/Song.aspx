<%@ Page Language="C#" AutoEventWireup="true" CodeBehind="Song.aspx.cs" Inherits="WebAppPoliMusicV2.Song" %>

<!DOCTYPE html>
<html xmlns="http://www.w3.org/1999/xhtml">
<head runat="server">
    <meta charset="utf-8" />
    <meta name="viewport" content="width=device-width, initial-scale=1" />
    <title>Canciones | PoliMusic</title>
    <style type="text/css">
        *, *::before, *::after { box-sizing: border-box; }

        html, body { margin: 0; padding: 0; }

        body {
            background: #fbfbfc;
            color: #141416;
            font-family: -apple-system, BlinkMacSystemFont, "Segoe UI", Roboto, Helvetica, Arial, sans-serif;
            font-size: 14px;
            -webkit-font-smoothing: antialiased;
            padding-bottom: 104px;
        }

        /* ---------- top bar ---------- */
        .topbar {
            position: sticky;
            top: 0;
            z-index: 10;
            background: rgba(251, 251, 252, 0.92);
            -webkit-backdrop-filter: blur(10px);
            backdrop-filter: blur(10px);
            border-bottom: 1px solid #ececef;
        }

        .topbar-inner {
            max-width: 820px;
            margin: 0 auto;
            padding: 16px 24px;
            display: flex;
            align-items: center;
            justify-content: space-between;
        }

        .brand {
            font-size: 15px;
            font-weight: 600;
            letter-spacing: -0.01em;
        }

        .brand::before {
            content: "";
            display: inline-block;
            width: 8px;
            height: 8px;
            margin-right: 9px;
            border-radius: 50%;
            background: #141416;
            vertical-align: middle;
        }

        .topbar-right {
            display: flex;
            align-items: center;
            gap: 18px;
        }

        .user-name { font-size: 13px; color: #6f6f76; }

        .logout {
            font-size: 13px;
            color: #6f6f76;
            text-decoration: none;
            padding-bottom: 1px;
            border-bottom: 1px solid transparent;
        }

        .logout:hover { color: #141416; border-bottom-color: #141416; }

        /* ---------- content ---------- */
        .content {
            max-width: 820px;
            margin: 0 auto;
            padding: 40px 24px 24px;
        }

        .page-head {
            display: flex;
            align-items: baseline;
            justify-content: space-between;
            margin-bottom: 22px;
        }

        h1 {
            margin: 0;
            font-size: 24px;
            font-weight: 600;
            letter-spacing: -0.02em;
        }

        .count { font-size: 13px; color: #8b8b93; }

        .alert {
            display: block;
            margin-bottom: 18px;
            padding: 12px 14px;
            border: 1px solid #f0cccc;
            border-radius: 8px;
            background: #fdf5f5;
            color: #a33a3a;
            font-size: 13px;
            line-height: 1.5;
        }

        /* ---------- track list ---------- */
        .track-head,
        .track {
            display: grid;
            grid-template-columns: 34px 1fr 170px 96px;
            align-items: center;
            gap: 16px;
            padding-left: 8px;
            padding-right: 8px;
        }

        .track-head {
            padding-bottom: 9px;
            font-size: 10.5px;
            text-transform: uppercase;
            letter-spacing: 0.09em;
            color: #9a9aa2;
            border-bottom: 1px solid #ececef;
        }

        .tracklist {
            list-style: none;
            margin: 0;
            padding: 0;
        }

        .track {
            padding-top: 11px;
            padding-bottom: 11px;
            border-bottom: 1px solid #ececef;
            transition: background 0.15s ease;
        }

        .track:hover { background: #f4f4f6; }

        .track-play {
            width: 32px;
            height: 32px;
            padding: 0;
            display: flex;
            align-items: center;
            justify-content: center;
            border: 1px solid #dededf;
            border-radius: 50%;
            background: #ffffff;
            color: #141416;
            cursor: pointer;
            transition: background 0.15s ease, color 0.15s ease, border-color 0.15s ease;
        }

        .track-play:hover,
        .track.is-playing .track-play {
            background: #141416;
            border-color: #141416;
            color: #ffffff;
        }

        .track-play .icon::before { content: "\25B6"; font-size: 10px; }

        .track.is-playing .track-play .icon::before { content: "\2016\2016"; font-size: 9px; letter-spacing: 1px; }

        .track-main { min-width: 0; }

        .track-title {
            display: block;
            font-size: 14px;
            font-weight: 500;
            white-space: nowrap;
            overflow: hidden;
            text-overflow: ellipsis;
        }

        .track.is-playing .track-title { font-weight: 600; }

        .track-artist {
            display: block;
            margin-top: 2px;
            font-size: 12.5px;
            color: #8b8b93;
            white-space: nowrap;
            overflow: hidden;
            text-overflow: ellipsis;
        }

        .track-album {
            font-size: 13px;
            color: #8b8b93;
            white-space: nowrap;
            overflow: hidden;
            text-overflow: ellipsis;
        }

        .track-plays,
        .col-right {
            font-size: 12.5px;
            color: #8b8b93;
            text-align: right;
            font-variant-numeric: tabular-nums;
        }

        .empty {
            padding: 56px 0;
            text-align: center;
            color: #8b8b93;
            font-size: 14px;
        }

        /* ---------- player ---------- */
        .player {
            position: fixed;
            left: 0;
            right: 0;
            bottom: 0;
            z-index: 20;
            background: rgba(255, 255, 255, 0.95);
            -webkit-backdrop-filter: blur(12px);
            backdrop-filter: blur(12px);
            border-top: 1px solid #ececef;
        }

        .player-inner {
            max-width: 820px;
            margin: 0 auto;
            padding: 12px 24px;
            display: flex;
            align-items: center;
            justify-content: space-between;
            gap: 24px;
        }

        .player-meta { min-width: 0; }

        .player-label {
            display: block;
            font-size: 10.5px;
            text-transform: uppercase;
            letter-spacing: 0.09em;
            color: #9a9aa2;
        }

        .player-title {
            display: block;
            margin-top: 3px;
            font-size: 13.5px;
            font-weight: 500;
            white-space: nowrap;
            overflow: hidden;
            text-overflow: ellipsis;
        }

        .player-title.is-idle { font-weight: 400; color: #9a9aa2; }

        audio {
            width: 320px;
            max-width: 48vw;
            height: 34px;
        }

        @media (max-width: 640px) {
            .content { padding: 28px 18px 24px; }
            .topbar-inner, .player-inner { padding-left: 18px; padding-right: 18px; }
            .track-head, .track { grid-template-columns: 34px 1fr 72px; gap: 12px; }
            .track-head .col-album, .track-album { display: none; }
            .player-inner { flex-direction: column; align-items: stretch; gap: 10px; }
            audio { width: 100%; max-width: none; }
        }
    </style>
</head>
<body>
    <form id="formSong" runat="server">

        <header class="topbar">
            <div class="topbar-inner">
                <span class="brand">PoliMusic</span>
                <div class="topbar-right">
                    <asp:Label ID="lblUserName" runat="server" CssClass="user-name"></asp:Label>
                    <a class="logout" href="Logout.aspx">Salir</a>
                </div>
            </div>
        </header>

        <main class="content">

            <div class="page-head">
                <h1>Canciones</h1>
                <asp:Label ID="lblCount" runat="server" CssClass="count"></asp:Label>
            </div>

            <asp:Label ID="lblMessage" runat="server" CssClass="alert" Visible="false"></asp:Label>

            <div class="track-head">
                <span></span>
                <span>Canción</span>
                <span class="col-album">Álbum</span>
                <span class="col-right">Reproducciones</span>
            </div>

            <ul class="tracklist">
                <asp:Repeater ID="rptSongs" runat="server">
                    <ItemTemplate>
                        <li class="track">
                            <button type="button" class="track-play"
                                data-src='<%# Attr(ResolveUrl(Convert.ToString(Eval("FilePath")))) %>'
                                data-title='<%# Attr(Eval("Name")) %>'
                                data-artist='<%# Attr(Eval("Artist")) %>'
                                title="Reproducir"
                                onclick="playSong(this); return false;">
                                <span class="icon"></span>
                            </button>
                            <div class="track-main">
                                <span class="track-title"><%# Attr(Eval("Name")) %></span>
                                <span class="track-artist"><%# Attr(Eval("Artist")) %></span>
                            </div>
                            <span class="track-album" title='<%# Attr(Eval("Album")) %>'><%# Attr(Eval("Album")) %></span>
                            <span class="track-plays" title="Reproducciones"><%# Attr(Eval("Plays")) %></span>
                        </li>
                    </ItemTemplate>
                </asp:Repeater>
            </ul>

            <asp:Panel ID="pnlEmpty" runat="server" CssClass="empty" Visible="false">
                <p>No hay canciones registradas todavía.</p>
            </asp:Panel>

        </main>

        <footer class="player">
            <div class="player-inner">
                <div class="player-meta">
                    <span class="player-label">Reproduciendo ahora</span>
                    <span id="nowPlaying" class="player-title is-idle">Ninguna canción seleccionada</span>
                </div>
                <audio id="audioPlayer" controls="controls" preload="none">
                    Tu navegador no soporta la reproducción de audio.
                </audio>
            </div>
        </footer>

    </form>

    <script type="text/javascript">
        var audioPlayer = document.getElementById('audioPlayer');
        var nowPlaying = document.getElementById('nowPlaying');
        var currentRow = null;

        function playSong(button) {
            var row = button.parentNode;
            var src = button.getAttribute('data-src');
            if (!src) { return false; }

            // The same track: pause it or resume it, never reload it.
            if (currentRow === row) {
                if (audioPlayer.paused) { audioPlayer.play(); } else { audioPlayer.pause(); }
                return false;
            }

            var title = button.getAttribute('data-title');
            var artist = button.getAttribute('data-artist');

            audioPlayer.src = src;
            audioPlayer.play();

            nowPlaying.textContent = artist ? title + ' \u2014 ' + artist : title;
            nowPlaying.className = 'player-title';

            if (currentRow) { currentRow.className = 'track'; }
            row.className = 'track is-playing';
            currentRow = row;

            return false;
        }

        audioPlayer.addEventListener('ended', function () {
            if (currentRow) { currentRow.className = 'track'; }
            currentRow = null;
            nowPlaying.textContent = 'Ninguna canción seleccionada';
            nowPlaying.className = 'player-title is-idle';
        });

        audioPlayer.addEventListener('error', function () {
            if (audioPlayer.getAttribute('src')) {
                nowPlaying.textContent = 'No se pudo reproducir el archivo. Verifica que el mp3 exista en la carpeta songFiles.';
                nowPlaying.className = 'player-title';
            }
        });
    </script>
</body>
</html>
