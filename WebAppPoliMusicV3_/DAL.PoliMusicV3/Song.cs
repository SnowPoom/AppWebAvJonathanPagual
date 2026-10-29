using System;
using System.Collections.Generic;
using System.Data;
using System.Data.SqlClient;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace DAL.PoliMusicV3
{
    /// <summary>
    /// Every database operation related to the [TBL_SONG] table.
    /// </summary>
    public class Song
    {
        private string strConnectionString;

        /// <summary>
        /// Base sentence used to read the songs.
        /// A song can be performed by more than one artist and can belong to more than
        /// one album, so both values are concatenated in a single column to avoid
        /// duplicated rows in the result.
        /// </summary>
        private const string SELECT_SONG =
            "SELECT s.[ID_SONG] ,s.[SONG_NAME] ,s.[SONG_PATH] ,ISNULL(s.[PLAYS], 0) AS [PLAYS] " +
            ",ISNULL(STUFF((SELECT DISTINCT ', ' + LTRIM(RTRIM(a.[ARTIST_NAME] + ' ' + ISNULL(a.[ARTIST_LASTNAME], ''))) " +
            "FROM [TBL_PLAYER] p INNER JOIN [TBL_ARTIST] a ON a.[ID_ARTIST] = p.[ID_ARTIST] " +
            "WHERE p.[ID_SONG] = s.[ID_SONG] FOR XML PATH('')), 1, 2, ''), '') AS [ARTISTS] " +
            ",ISNULL(STUFF((SELECT DISTINCT ', ' + al.[TITLE_ALBUM] " +
            "FROM [TBL_TRACKLIST] tl INNER JOIN [TBL_ALBUM] al ON al.[ID_ALBUM] = tl.[ID_ALBUM] " +
            "WHERE tl.[ID_SONG] = s.[ID_SONG] FOR XML PATH('')), 1, 2, ''), '') AS [ALBUMS] " +
            "FROM [TBL_SONG] s ";

        public Song(string strConnString)
        {
            strConnectionString = strConnString;
        }

        /// <summary>
        /// Returns every song with its artists and albums ready to be bound to a control.
        /// </summary>
        public DataTable Read()
        {
            DataTable dtSong = new DataTable();
            using (SqlConnection con = new SqlConnection(strConnectionString))
            {
                using (SqlCommand cmd = new SqlCommand(SELECT_SONG + "ORDER BY s.[SONG_NAME]", con))
                {
                    using (SqlDataAdapter sda = new SqlDataAdapter(cmd))
                    {
                        sda.Fill(dtSong);
                        return dtSong;
                    }
                }
            }
        }

        /// <summary>
        /// Returns a single song as an entity, or null when it does not exist.
        /// </summary>
        public EntityLayer.PoliMusicV3.Song GetById(int idSong)
        {
            using (SqlConnection con = new SqlConnection(strConnectionString))
            {
                using (SqlCommand cmd = new SqlCommand())
                {
                    cmd.CommandText = "SELECT [ID_SONG] ,[SONG_NAME] ,[SONG_PATH] ,ISNULL([PLAYS], 0) AS [PLAYS] " +
                        "FROM [TBL_SONG] WHERE [ID_SONG] = @ID_SONG";
                    cmd.Parameters.AddWithValue("@ID_SONG", idSong);
                    cmd.Connection = con;
                    con.Open();
                    using (SqlDataReader reader = cmd.ExecuteReader())
                    {
                        if (reader.Read())
                        {
                            return new EntityLayer.PoliMusicV3.Song(
                                Convert.ToInt32(reader["ID_SONG"]),
                                reader["SONG_NAME"].ToString(),
                                reader["SONG_PATH"].ToString(),
                                Convert.ToInt32(reader["PLAYS"]));
                        }
                    }
                }
            }
            return null;
        }

        /// <summary>
        /// Inserts a new song. [ID_SONG] is an identity column, so it is generated
        /// by the database and must not be sent.
        /// </summary>
        public int Save(EntityLayer.PoliMusicV3.Song song)
        {
            using (SqlConnection con = new SqlConnection(strConnectionString))
            {
                using (SqlCommand cmd = new SqlCommand())
                {
                    cmd.CommandText = "INSERT INTO [TBL_SONG] ([SONG_NAME], [SONG_PATH], [PLAYS]) " +
                        "VALUES (@SONG_NAME, @SONG_PATH, @PLAYS)";
                    cmd.Parameters.AddWithValue("@SONG_NAME", song.SongName);
                    cmd.Parameters.AddWithValue("@SONG_PATH", song.SongPath);
                    cmd.Parameters.AddWithValue("@PLAYS", song.Plays);
                    cmd.Connection = con;
                    con.Open();
                    return cmd.ExecuteNonQuery();
                }
            }
        }

        public int Update(EntityLayer.PoliMusicV3.Song song)
        {
            using (SqlConnection con = new SqlConnection(strConnectionString))
            {
                using (SqlCommand cmd = new SqlCommand())
                {
                    cmd.CommandText = "UPDATE [TBL_SONG] SET " +
                        "[SONG_NAME] = @SONG_NAME " +
                        ",[SONG_PATH] = @SONG_PATH " +
                        ",[PLAYS] = @PLAYS " +
                        "WHERE [ID_SONG] = @ID_SONG";
                    cmd.Parameters.AddWithValue("@ID_SONG", song.ID);
                    cmd.Parameters.AddWithValue("@SONG_NAME", song.SongName);
                    cmd.Parameters.AddWithValue("@SONG_PATH", song.SongPath);
                    cmd.Parameters.AddWithValue("@PLAYS", song.Plays);
                    cmd.Connection = con;
                    con.Open();
                    return cmd.ExecuteNonQuery();
                }
            }
        }

        /// <summary>
        /// Deletes the song and its related rows (player, tracklist and playlist song)
        /// inside a transaction, because those tables reference [TBL_SONG].
        /// </summary>
        public int Delete(int idSong)
        {
            using (SqlConnection con = new SqlConnection(strConnectionString))
            {
                con.Open();
                using (SqlTransaction transaction = con.BeginTransaction())
                {
                    try
                    {
                        string[] sentences = new string[]
                        {
                            "DELETE [TBL_PLAYLIST_SONG] WHERE [ID_SONG] = @ID_SONG",
                            "DELETE [TBL_PLAYER] WHERE [ID_SONG] = @ID_SONG",
                            "DELETE [TBL_TRACKLIST] WHERE [ID_SONG] = @ID_SONG",
                            "DELETE [TBL_SONG] WHERE [ID_SONG] = @ID_SONG"
                        };
                        int rowsAffected = 0;
                        foreach (string sentence in sentences)
                        {
                            using (SqlCommand cmd = new SqlCommand(sentence, con, transaction))
                            {
                                cmd.Parameters.AddWithValue("@ID_SONG", idSong);
                                rowsAffected = cmd.ExecuteNonQuery();
                            }
                        }
                        transaction.Commit();
                        return rowsAffected;
                    }
                    catch (Exception)
                    {
                        transaction.Rollback();
                        throw;
                    }
                }
            }
        }

        /// <summary>
        /// Adds one play to the counter of the song.
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
        /// Returns the identifier of the song with that name, or 0 when it does not exist.
        /// </summary>
        public int CheckExistSong(string songName)
        {
            object returnValue;
            using (SqlConnection con = new SqlConnection(strConnectionString))
            {
                using (SqlCommand cmd = new SqlCommand())
                {
                    cmd.CommandText = "SELECT [ID_SONG] FROM [TBL_SONG] WHERE [SONG_NAME] = @SONG_NAME";
                    cmd.Parameters.AddWithValue("@SONG_NAME", songName);
                    cmd.Connection = con;
                    con.Open();
                    returnValue = cmd.ExecuteScalar();
                    return (returnValue == null) ? 0 : Convert.ToInt32(returnValue);
                }
            }
        }
    }

}
