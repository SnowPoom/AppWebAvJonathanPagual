/* =====================================================================
   TBL_SONG  (esquema real de BDD_PoliMusic)
   Tabla usada por Entities/Song.cs, DAL/Song.cs y BLL/Song.cs.

   Nota: la consulta del catalogo (DAL.Song.SongBaseSelect) tambien lee
   TBL_PLAYER, TBL_ARTIST, TBL_TRACKLIST y TBL_ALBUM, por lo que esas
   tablas deben existir. Este script solo crea TBL_SONG.
   ===================================================================== */

USE [BDD_PoliMusic]
GO

IF NOT EXISTS (SELECT * FROM sys.objects WHERE object_id = OBJECT_ID(N'[dbo].[TBL_SONG]') AND type = N'U')
BEGIN
    CREATE TABLE [dbo].[TBL_SONG](
        [ID_SONG]   [int] IDENTITY(1,1) NOT NULL,
        [SONG_NAME] [varchar](50)  NOT NULL,
        [SONG_PATH] [varchar](255) NOT NULL,
        [PLAYS]     [int] NULL,
        CONSTRAINT [PK_TBL_SONG] PRIMARY KEY CLUSTERED ([ID_SONG] ASC)
    ) ON [PRIMARY]
END
GO

/* Canciones de ejemplo: el mp3 debe existir en la carpeta songFiles. */
IF NOT EXISTS (SELECT 1 FROM [dbo].[TBL_SONG] WHERE [SONG_NAME] = 'Adventure')
BEGIN
    INSERT INTO [dbo].[TBL_SONG] ([SONG_NAME], [SONG_PATH], [PLAYS])
    VALUES ('Adventure', '~/songFiles/bensound-adventure.mp3', 0)
END
GO
