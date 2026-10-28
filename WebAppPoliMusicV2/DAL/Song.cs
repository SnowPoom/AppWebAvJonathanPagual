using System;
using System.Data;
using System.Data.SqlClient;

namespace WebAppPoliMusicV2.DAL
{
    /// <summary>
    /// Data access logic for the song catalog:
    /// TBL_SONG (+ TBL_PLAYER / TBL_ARTIST / TBL_TRACKLIST / TBL_ALBUM for the joins).
    /// </summary>
    public class Song
    {
        private string strConnectionString;

        public Song(string strConnString)
        {
            strConnectionString = strConnString;
        }

        public int Save(Entities.Song song)
        {
            using (SqlConnection con = new SqlConnection(strConnectionString))
            {
                using (SqlCommand cmd = new SqlCommand())
                {
                    cmd.CommandText = "INSERT INTO [TBL_SONG]([SONG_NAME], [SONG_PATH], [PLAYS]) " +
                        "VALUES(@SONG_NAME, @SONG_PATH, @PLAYS)";
                    cmd.Parameters.AddWithValue("@SONG_NAME", song.Name);
                    cmd.Parameters.AddWithValue("@SONG_PATH", song.FilePath);
                    cmd.Parameters.AddWithValue("@PLAYS", song.Plays < 0 ? 0 : song.Plays);
                    cmd.Connection = con;
                    con.Open();
                    return cmd.ExecuteNonQuery();
                }
            }
        }

        /// <summary>
        /// Returns the whole catalog with the artists and the album of every song.
        /// </summary>
        public DataTable Read()
        {
            DataTable dtSong = new DataTable();
            using (SqlConnection con = new SqlConnection(strConnectionString))
            {
                using (SqlCommand cmd = new SqlCommand(SongBaseSelect() + " ORDER BY s.[SONG_NAME]", con))
                {
                    using (SqlDataAdapter sda = new SqlDataAdapter(cmd))
                    {
                        sda.Fill(dtSong);
                        return dtSong;
                    }
                }
            }
        }

        public DataTable GetById(int idSong)
        {
            DataTable dtSong = new DataTable();
            using (SqlConnection con = new SqlConnection(strConnectionString))
            {
                using (SqlCommand cmd = new SqlCommand(SongBaseSelect() + " WHERE s.[ID_SONG] = @ID_SONG", con))
                {
                    cmd.Parameters.AddWithValue("@ID_SONG", idSong);
                    using (SqlDataAdapter sda = new SqlDataAdapter(cmd))
                    {
                        sda.Fill(dtSong);
                        return dtSong;
                    }
                }
            }
        }

        public int Update(Entities.Song song)
        {
            using (SqlConnection con = new SqlConnection(strConnectionString))
            {
                using (SqlCommand cmd = new SqlCommand())
                {
                    cmd.CommandText = "UPDATE [TBL_SONG] SET " +
                        "[SONG_NAME] = @SONG_NAME " +
                        ", [SONG_PATH] = @SONG_PATH " +
                        ", [PLAYS] = @PLAYS " +
                        "WHERE [ID_SONG] = @ID_SONG";
                    cmd.Parameters.AddWithValue("@ID_SONG", song.ID);
                    cmd.Parameters.AddWithValue("@SONG_NAME", song.Name);
                    cmd.Parameters.AddWithValue("@SONG_PATH", song.FilePath);
                    cmd.Parameters.AddWithValue("@PLAYS", song.Plays < 0 ? 0 : song.Plays);
                    cmd.Connection = con;
                    con.Open();
                    return cmd.ExecuteNonQuery();
                }
            }
        }

        /// <summary>
        /// Deletes the song and its rows in the related tables (TBL_PLAYLIST_SONG,
        /// TBL_TRACKLIST, TBL_PLAYER) inside a transaction, so the foreign keys do not block it.
        /// </summary>
        public int Delete(int idSong)
        {
            using (SqlConnection con = new SqlConnection(strConnectionString))
            {
                con.Open();
                using (SqlTransaction tr = con.BeginTransaction())
                {
                    try
                    {
                        DeleteRelation(con, tr, "DELETE FROM [TBL_PLAYLIST_SONG] WHERE [ID_SONG] = @ID_SONG", idSong);
                        DeleteRelation(con, tr, "DELETE FROM [TBL_TRACKLIST] WHERE [ID_SONG] = @ID_SONG", idSong);
                        DeleteRelation(con, tr, "DELETE FROM [TBL_PLAYER] WHERE [ID_SONG] = @ID_SONG", idSong);
                        int rows = DeleteRelation(con, tr, "DELETE FROM [TBL_SONG] WHERE [ID_SONG] = @ID_SONG", idSong);
                        tr.Commit();
                        return rows;
                    }
                    catch (Exception)
                    {
                        tr.Rollback();
                        throw;
                    }
                }
            }
        }

        /// <summary>
        /// Returns the ID of the song with that name, or 0 if it does not exist.
        /// </summary>
        public int CheckExistSong(string name)
        {
            object returnValue;
            using (SqlConnection con = new SqlConnection(strConnectionString))
            {
                using (SqlCommand cmd = new SqlCommand())
                {
                    cmd.CommandText = "SELECT [ID_SONG] FROM [TBL_SONG] WHERE [SONG_NAME] = @SONG_NAME";
                    cmd.Parameters.AddWithValue("@SONG_NAME", name);
                    cmd.Connection = con;
                    con.Open();
                    returnValue = cmd.ExecuteScalar();
                    return returnValue == null ? 0 : Convert.ToInt32(returnValue);
                }
            }
        }

        /// <summary>
        /// Adds one play to the song counter.
        /// </summary>
        public int IncrementPlays(int idSong)
        {
            using (SqlConnection con = new SqlConnection(strConnectionString))
            {
                using (SqlCommand cmd = new SqlCommand())
                {
                    cmd.CommandText = "UPDATE [TBL_SONG] SET [PLAYS] = ISNULL([PLAYS], 0) + 1 WHERE [ID_SONG] = @ID_SONG";
                    cmd.Parameters.AddWithValue("@ID_SONG", idSong);
                    cmd.Connection = con;
                    con.Open();
                    return cmd.ExecuteNonQuery();
                }
            }
        }

        /// <summary>
        /// Base query of the catalog: song data plus the artists of TBL_PLAYER
        /// and the first album of TBL_TRACKLIST.
        /// </summary>
        private string SongBaseSelect()
        {
            return "SELECT s.[ID_SONG], s.[SONG_NAME], s.[SONG_PATH], ISNULL(s.[PLAYS], 0) AS [PLAYS], " +
                "ISNULL((SELECT STRING_AGG(LTRIM(RTRIM(a.[ARTIST_NAME] + ' ' + ISNULL(a.[ARTIST_LASTNAME], ''))), ', ') " +
                "  FROM [TBL_PLAYER] p INNER JOIN [TBL_ARTIST] a ON a.[ID_ARTIST] = p.[ID_ARTIST] " +
                "  WHERE p.[ID_SONG] = s.[ID_SONG]), '') AS [ARTISTS], " +
                "ISNULL((SELECT TOP 1 al.[TITLE_ALBUM] " +
                "  FROM [TBL_TRACKLIST] t INNER JOIN [TBL_ALBUM] al ON al.[ID_ALBUM] = t.[ID_ALBUM] " +
                "  WHERE t.[ID_SONG] = s.[ID_SONG] ORDER BY al.[ID_ALBUM]), '') AS [ALBUM] " +
                "FROM [TBL_SONG] s";
        }

        private int DeleteRelation(SqlConnection con, SqlTransaction tr, string sentence, int idSong)
        {
            using (SqlCommand cmd = new SqlCommand(sentence, con, tr))
            {
                cmd.Parameters.AddWithValue("@ID_SONG", idSong);
                return cmd.ExecuteNonQuery();
            }
        }
    }
}
