<%
' =====================================================================
'  PoliMusic - Componentes de interfaz reutilizables
' =====================================================================

' Fila de canción usada en todas las vistas (home, catálogo, álbum, playlist)
' extraActions: HTML adicional para la zona de acciones (puede ser "")
Function SongRowHtml(ByVal songId, ByVal songName, ByVal songPath, ByVal plays, _
                     ByVal artists, ByVal albumTitle, ByVal albumImage, _
                     ByVal extraActions, ByVal rowIndex)
    Dim src, available, coverLabel, cls, html, coverHtml
    src = AssetUrl(songPath)
    available = AssetExists(songPath)
    coverLabel = albumTitle
    If Trim(coverLabel & "") = "" Then coverLabel = songName
    coverHtml = Cover(albumImage, coverLabel, "")
    cls = "song-row"
    If Not available Then cls = cls & " is-unavailable"

    html = "<div class=""" & cls & """ data-song-id=""" & songId & """" & _
           " data-title=""" & HtmlAttr(songName) & """" & _
           " data-artist=""" & HtmlAttr(artists) & """" & _
           " data-album=""" & HtmlAttr(albumTitle) & """" & _
           " data-src=""" & HtmlAttr(src) & """" & _
           " data-available=""" & IIf(available, "1", "0") & """ tabindex=""0"" role=""button"">"

    html = html & "<div class=""song-row__num""><span class=""num-index"">" & rowIndex & "</span>" & _
           "<span class=""num-play"">" & Icon("play", "") & "</span></div>"

    html = html & "<div class=""song-row__cover"">" & coverHtml & "</div>"

    html = html & "<div class=""song-row__main""><div class=""song-row__title"">" & HtmlEncode(songName)
    If Not available Then html = html & " <span class=""badge-soon"">Sin audio</span>"
    html = html & "</div><div class=""song-row__artist"">" & HtmlEncode(TextOr(artists, "Artista desconocido")) & "</div></div>"

    html = html & "<div class=""song-row__album"">" & HtmlEncode(TextOr(albumTitle, "-")) & "</div>"

    html = html & "<div class=""song-row__plays"">" & Icon("headphones", "") & ToInt(plays, 0) & "</div>"

    html = html & "<div class=""song-row__actions"">" & extraActions & _
           "<button class=""icon-btn"" type=""button"" data-add-to=""" & songId & """ data-song-title=""" & _
           HtmlAttr(songName) & """ title=""Agregar a una playlist"" aria-label=""Agregar a una playlist"">" & _
           Icon("plus", "") & "</button></div>"

    html = html & "</div>"
    SongRowHtml = html
End Function

Function EmptyState(ByVal iconName, ByVal title, ByVal text, ByVal actionHtml)
    Dim html
    html = "<div class=""empty"">" & Icon(iconName, "") & "<strong>" & HtmlEncode(title) & "</strong>" & _
           "<p>" & HtmlEncode(text) & "</p>" & actionHtml & "</div>"
    EmptyState = html
End Function

' Tarjeta de álbum
Function AlbumTile(ByVal albumId, ByVal title, ByVal year, ByVal imagePath, ByVal trackCount)
    Dim html, label
    label = title & ""
    html = "<a class=""tile"" href=""" & AppBase() & "album.asp?id=" & albumId & """>" & _
           "<div class=""tile__media"">" & Cover(imagePath, label, "") & _
           "<span class=""tile__play"" aria-hidden=""true"">" & Icon("play", "") & "</span></div>" & _
           "<div class=""tile__body""><div class=""tile__title"">" & HtmlEncode(title) & "</div>" & _
           "<div class=""tile__meta"">" & IIf(ToInt(year, 0) > 0, year, "Sin año") & " &middot; " & _
           ToInt(trackCount, 0) & " " & IIf(ToInt(trackCount, 0) = 1, "canción", "canciones") & "</div></div></a>"
    AlbumTile = html
End Function

' Tarjeta de artista (circular)
Function ArtistTile(ByVal artistId, ByVal fullName, ByVal imagePath, ByVal songCount)
    Dim html
    html = "<a class=""tile tile--artist"" href=""" & AppBase() & "artist.asp?id=" & artistId & """>" & _
           "<div class=""tile__media"">" & Cover(imagePath, fullName, "cover--rounded") & _
           "<span class=""tile__play"" aria-hidden=""true"">" & Icon("play", "") & "</span></div>" & _
           "<div class=""tile__body""><div class=""tile__title"">" & HtmlEncode(fullName) & "</div>" & _
           "<div class=""tile__meta"">" & ToInt(songCount, 0) & " " & IIf(ToInt(songCount, 0) = 1, "canción", "canciones") & _
           "</div></div></a>"
    ArtistTile = html
End Function

' Tarjeta de playlist
Function PlaylistTile(ByVal playlistId, ByVal title, ByVal typeValue, ByVal photo, ByVal songCount, ByVal creationDate, ByVal ownerName)
    Dim html, typeChip
    If ToInt(typeValue, 0) = 1 Then
        typeChip = "<span class=""chip chip--private"">Privada</span>"
    Else
        typeChip = "<span class=""chip chip--public"">Pública</span>"
    End If
    html = "<a class=""tile"" href=""" & AppBase() & "playlist.asp?id=" & playlistId & """>" & _
           "<div class=""tile__media"">" & Cover(photo, title, "") & _
           "<span class=""tile__play"" aria-hidden=""true"">" & Icon("play", "") & "</span></div>" & _
           "<div class=""tile__body""><div class=""tile__title"">" & HtmlEncode(title) & "</div>" & _
           "<div class=""tile__meta"">" & ToInt(songCount, 0) & " " & IIf(ToInt(songCount, 0) = 1, "canción", "canciones") & _
           " &middot; " & HtmlEncode(ownerName) & "</div>" & _
           "<div style=""margin-top:6px"">" & typeChip & "</div></div></a>"
    PlaylistTile = html
End Function

' Consulta base de canciones con artistas, álbum y portada (una sola fuente)
Function SongBaseSql()
    SongBaseSql = "SELECT s.ID_SONG, s.SONG_NAME, s.SONG_PATH, ISNULL(s.PLAYS,0) AS PLAYS, " & _
        "ISNULL((SELECT STRING_AGG(LTRIM(RTRIM(a.ARTIST_NAME + ' ' + ISNULL(a.ARTIST_LASTNAME,''))), ', ') " & _
        "  FROM TBL_PLAYER p INNER JOIN TBL_ARTIST a ON a.ID_ARTIST = p.ID_ARTIST " & _
        "  WHERE p.ID_SONG = s.ID_SONG), '') AS ARTISTS, " & _
        "(SELECT TOP 1 al.ID_ALBUM FROM TBL_TRACKLIST t INNER JOIN TBL_ALBUM al ON al.ID_ALBUM = t.ID_ALBUM " & _
        "  WHERE t.ID_SONG = s.ID_SONG ORDER BY al.ID_ALBUM) AS ID_ALBUM, " & _
        "(SELECT TOP 1 al.TITLE_ALBUM FROM TBL_TRACKLIST t INNER JOIN TBL_ALBUM al ON al.ID_ALBUM = t.ID_ALBUM " & _
        "  WHERE t.ID_SONG = s.ID_SONG ORDER BY al.ID_ALBUM) AS ALBUM_TITLE, " & _
        "(SELECT TOP 1 al.ALBUM_IMAGE_PATH FROM TBL_TRACKLIST t INNER JOIN TBL_ALBUM al ON al.ID_ALBUM = t.ID_ALBUM " & _
        "  WHERE t.ID_SONG = s.ID_SONG ORDER BY al.ID_ALBUM) AS ALBUM_IMAGE " & _
        "FROM TBL_SONG s "
End Function
%>
