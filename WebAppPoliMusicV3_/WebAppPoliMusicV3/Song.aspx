<%@ Page Title="Canciones" Language="C#" MasterPageFile="~/Site.Master" AutoEventWireup="true" CodeBehind="Song.aspx.cs" Inherits="WebAppPoliMusicV3.Song" %>

<asp:Content ID="BodyContent" ContentPlaceHolderID="MainContent" runat="server">

    <style type="text/css">
        body {
            background: #f7f8fc;
        }

        .songs-page {
            max-width: 760px;
            margin: 0 auto;
            padding: 30px 16px 60px;
            color: #12161f;
            font-family: "Segoe UI", system-ui, -apple-system, "Helvetica Neue", Arial, sans-serif;
            -webkit-font-smoothing: antialiased;
        }

        /* ---------- Encabezado ---------- */

        .hero {
            margin-bottom: 24px;
        }

        .hero-eyebrow {
            margin: 0 0 8px;
            font-size: 11px;
            font-weight: 600;
            letter-spacing: .16em;
            text-transform: uppercase;
            color: #98a2b3;
        }

        .hero-title {
            margin: 0;
            font-size: 34px;
            font-weight: 600;
            line-height: 1.15;
            letter-spacing: -.025em;
            color: #0f1420;
        }

        .hero-text {
            max-width: 54ch;
            margin: 10px 0 0;
            font-size: 14.5px;
            line-height: 1.6;
            color: #6b7688;
        }

        /* ---------- Reproductor ---------- */

        .player {
            display: flex;
            flex-wrap: wrap;
            align-items: center;
            gap: 14px 16px;
            padding: 16px 18px;
            background: #ffffff;
            border: 1px solid #eceff5;
            border-radius: 18px;
            box-shadow: 0 1px 2px rgba(18, 22, 31, .04);
        }

        .player-disc {
            display: grid;
            flex: 0 0 44px;
            place-items: center;
            width: 44px;
            height: 44px;
            border-radius: 50%;
            background: linear-gradient(135deg, #eef0fe, #e3e7ff);
            color: #5b5fe0;
            font-size: 17px;
            line-height: 1;
        }

        .player.playing .player-disc {
            animation: disc-spin 9s linear infinite;
        }

        @keyframes disc-spin {
            to { transform: rotate(360deg); }
        }

        .player-info {
            display: flex;
            flex: 1 1 0%;
            flex-direction: column;
            min-width: 0;
        }

        .player-label {
            font-size: 10.5px;
            font-weight: 600;
            letter-spacing: .12em;
            text-transform: uppercase;
            color: #a3adba;
        }

        .player-title {
            overflow: hidden;
            font-size: 15px;
            font-weight: 600;
            color: #12161f;
            white-space: nowrap;
            text-overflow: ellipsis;
        }

        .player-meta {
            overflow: hidden;
            font-size: 12.5px;
            color: #8b96a5;
            white-space: nowrap;
            text-overflow: ellipsis;
        }

        .player-audio {
            flex: 1 0 100%;
            width: 100%;
            height: 36px;
            margin-top: 2px;
        }

        /* ---------- Lista ---------- */

        .panel {
            margin-top: 22px;
            background: #ffffff;
            border: 1px solid #eceff5;
            border-radius: 18px;
            overflow: hidden;
            box-shadow: 0 1px 2px rgba(18, 22, 31, .04);
        }

        .panel-head {
            display: flex;
            align-items: center;
            justify-content: space-between;
            padding: 15px 20px;
            border-bottom: 1px solid #f1f3f8;
        }

        .panel-title {
            font-size: 11px;
            font-weight: 600;
            letter-spacing: .14em;
            text-transform: uppercase;
            color: #98a2b3;
        }

        .panel-count {
            padding: 3px 10px;
            border-radius: 999px;
            background: #f2f4fb;
            font-size: 11.5px;
            font-weight: 600;
            color: #6b7688;
        }

        .tracks {
            margin: 0;
            padding: 6px;
            list-style: none;
        }

        .track {
            display: flex;
            align-items: center;
            gap: 14px;
            padding: 10px 12px;
            border-radius: 12px;
            transition: background .15s ease;
        }

        .track:hover {
            background: #fafbff;
        }

        .track-play {
            display: grid;
            flex: 0 0 38px;
            place-items: center;
            width: 38px;
            height: 38px;
            border-radius: 50%;
            background: #eef1fa;
            color: #4b52c9;
            font-size: 12px;
            line-height: 1;
            text-decoration: none;
            transition: background .15s ease, color .15s ease, transform .15s ease;
        }

        .track-play:hover,
        .track-play:focus {
            background: #5b5fe0;
            color: #ffffff;
            text-decoration: none;
            transform: scale(1.06);
        }

        .track-play:focus-visible {
            outline: 2px solid #5b5fe0;
            outline-offset: 3px;
        }

        .track-main {
            display: flex;
            flex: 1;
            flex-direction: column;
            min-width: 0;
        }

        .track-title {
            overflow: hidden;
            font-size: 14.5px;
            font-weight: 600;
            color: #12161f;
            white-space: nowrap;
            text-overflow: ellipsis;
        }

        .track-sub {
            overflow: hidden;
            font-size: 12.5px;
            color: #8b96a5;
            white-space: nowrap;
            text-overflow: ellipsis;
        }

        .track-plays {
            flex: 0 0 auto;
            font-size: 12px;
            font-weight: 600;
            color: #a3adba;
            white-space: nowrap;
            font-variant-numeric: tabular-nums;
        }

        .track-plays em {
            font-style: normal;
            font-weight: 400;
            color: #b6bfcb;
        }

        .empty-state {
            padding: 34px 20px;
            font-size: 13.5px;
            color: #8b96a5;
            text-align: center;
        }

        /* ---------- Responsive ---------- */

        @media (max-width: 575px) {
            .songs-page {
                padding-top: 22px;
            }

            .hero-title {
                font-size: 26px;
            }

            .track-plays {
                display: none;
            }
        }
    </style>

    <div class="songs-page">

        <header class="hero">
            <p class="hero-eyebrow">Tu biblioteca</p>
            <h1 class="hero-title">Hola, <asp:Label ID="lblUserName" runat="server" /></h1>
            <p class="hero-text">
                Estas son las canciones disponibles. Presiona reproducir y el contador
                de reproducciones se actualizará automáticamente.
            </p>
        </header>

        <section class="player" id="playerCard">
            <span class="player-disc" aria-hidden="true">&#9835;</span>
            <div class="player-info">
                <span class="player-label">Reproduciendo ahora</span>
                <span class="player-title" id="nowPlayingTitle">Selecciona una canción</span>
                <span class="player-meta" id="nowPlayingMeta"></span>
            </div>
            <audio id="audioPlayer" class="player-audio" controls="controls" preload="none"></audio>
        </section>

        <asp:UpdatePanel ID="upSongs" runat="server" UpdateMode="Conditional">
            <ContentTemplate>
                <div class="panel">
                    <div class="panel-head">
                        <span class="panel-title">Canciones</span>
                        <span class="panel-count"><asp:Label ID="lblTotal" runat="server" /></span>
                    </div>
                    <asp:Repeater ID="rptSongs" runat="server" OnItemCommand="rptSongs_ItemCommand" OnItemDataBound="rptSongs_ItemDataBound">
                        <HeaderTemplate>
                            <ul class="tracks">
                        </HeaderTemplate>
                        <ItemTemplate>
                            <li class="track">
                                <asp:LinkButton ID="btnPlay" runat="server" CssClass="track-play" CommandName="Play" CommandArgument='<%# Eval("ID_SONG") %>' ToolTip="Reproducir">&#9654;</asp:LinkButton>
                                <span class="track-main">
                                    <span class="track-title"><%#: Eval("SONG_NAME") %></span>
                                    <span class="track-sub"><%#: Display(Eval("ARTISTS"), "Artista desconocido") %> &#183; <%#: Display(Eval("ALBUMS"), "Álbum desconocido") %></span>
                                </span>
                                <span class="track-plays"><%#: Eval("PLAYS") %> <em><%#: PlaysLabel(Eval("PLAYS")) %></em></span>
                            </li>
                        </ItemTemplate>
                        <FooterTemplate>
                            </ul>
                        </FooterTemplate>
                    </asp:Repeater>
                    <asp:Panel ID="pnlEmpty" runat="server" CssClass="empty-state" Visible="false">
                        Todavía no hay canciones registradas en la biblioteca.
                    </asp:Panel>
                </div>
            </ContentTemplate>
        </asp:UpdatePanel>

    </div>

    <script type="text/javascript">
        function playSong(url, button) {
            var audio = document.getElementById('audioPlayer');
            if (!audio || !url) {
                return;
            }

            var track = button ? button.closest('.track') : null;
            var title = track ? track.querySelector('.track-title') : null;
            var sub = track ? track.querySelector('.track-sub') : null;

            document.getElementById('nowPlayingTitle').textContent = title ? title.textContent : 'Canción';
            document.getElementById('nowPlayingMeta').textContent = sub ? sub.textContent : '';
            document.getElementById('playerCard').className = 'player playing';

            audio.src = url;
            var promise = audio.play();
            if (promise && typeof promise.catch === 'function') {
                promise.catch(function () { /* el navegador puede pedir un clic adicional */ });
            }
        }
    </script>

</asp:Content>
