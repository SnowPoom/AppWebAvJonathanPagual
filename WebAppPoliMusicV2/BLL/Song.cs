using System;
using System.Collections.Generic;
using System.Data;

namespace WebAppPoliMusicV2.BLL
{
    /// <summary>
    /// Business logic for songs: it calls the DAL and applies the business rules
    /// (validations and translation of DataTable rows into Entities.Song objects).
    /// </summary>
    public class Song
    {
        // Column sizes of TBL_SONG, used by the validations.
        private const int MAX_LENGTH_SONG_NAME = 50;
        private const int MAX_LENGTH_SONG_PATH = 255;

        private string strConnectionString;

        public Song(string strConnString)
        {
            strConnectionString = strConnString;
        }

        public int Save(Entities.Song song)
        {
            string message = ValidationsRequired(song);
            if (message != string.Empty)
            {
                throw new ArgumentException(message);
            }
            message = ValidationsDuplicated(song);
            if (message != string.Empty)
            {
                throw new ArgumentException(message);
            }
            return new DAL.Song(strConnectionString).Save(song);
        }

        public int Update(Entities.Song song)
        {
            string message = ValidationsRequired(song);
            if (message != string.Empty)
            {
                throw new ArgumentException(message);
            }
            message = ValidationsDuplicated(song);
            if (message != string.Empty)
            {
                throw new ArgumentException(message);
            }
            return new DAL.Song(strConnectionString).Update(song);
        }

        public int Delete(int idSong)
        {
            if (idSong <= 0)
            {
                throw new ArgumentException("The song is required!");
            }
            return new DAL.Song(strConnectionString).Delete(idSong);
        }

        /// <summary>
        /// Returns the catalog (songs with their artists and album) as entities.
        /// </summary>
        public List<Entities.Song> GetSongs()
        {
            DataTable dtSong = new DAL.Song(strConnectionString).Read();
            return MapSongs(dtSong);
        }

        public Entities.Song GetSongById(int idSong)
        {
            DataTable dtSong = new DAL.Song(strConnectionString).GetById(idSong);
            List<Entities.Song> songs = MapSongs(dtSong);
            return songs.Count > 0 ? songs[0] : null;
        }

        /// <summary>
        /// Adds one play to the counter of the song.
        /// </summary>
        public int IncrementPlays(int idSong)
        {
            if (idSong <= 0)
            {
                return 0;
            }
            return new DAL.Song(strConnectionString).IncrementPlays(idSong);
        }

        public int CheckExistSong(string name)
        {
            return new DAL.Song(strConnectionString).CheckExistSong(name);
        }

        public string ValidationsRequired(Entities.Song song)
        {
            if (song == null)
            {
                return "The song is required!";
            }
            if (string.IsNullOrEmpty(song.Name) || song.Name.Trim() == string.Empty)
            {
                return "The song name is required!";
            }
            if (song.Name.Trim().Length > MAX_LENGTH_SONG_NAME)
            {
                return "The song name can not be longer than " + MAX_LENGTH_SONG_NAME + " characters!";
            }
            if (string.IsNullOrEmpty(song.FilePath) || song.FilePath.Trim() == string.Empty)
            {
                return "The song file is required!";
            }
            if (song.FilePath.Trim().Length > MAX_LENGTH_SONG_PATH)
            {
                return "The song path can not be longer than " + MAX_LENGTH_SONG_PATH + " characters!";
            }
            if (song.FilePath.ToLower().EndsWith(".mp3") == false)
            {
                return "Only .mp3 files are allowed!";
            }
            if (song.Plays < 0)
            {
                return "The number of plays can not be negative!";
            }
            return string.Empty;
        }

        public string ValidationsDuplicated(Entities.Song song)
        {
            int returnValue = CheckExistSong(song.Name.Trim());
            if (returnValue != 0 && returnValue != song.ID)
            {
                return "A song with that name is already registered!";
            }
            return string.Empty;
        }

        /// <summary>
        /// Translates the rows returned by the DAL into entities.
        /// </summary>
        private List<Entities.Song> MapSongs(DataTable dtSong)
        {
            List<Entities.Song> songs = new List<Entities.Song>();
            if (dtSong == null)
            {
                return songs;
            }
            foreach (DataRow row in dtSong.Rows)
            {
                Entities.Song song = new Entities.Song();
                song.ID = Convert.ToInt32(row["ID_SONG"]);
                song.Name = row["SONG_NAME"].ToString().Trim();
                song.FilePath = row["SONG_PATH"].ToString().Trim();
                song.Plays = row["PLAYS"] == DBNull.Value ? 0 : Convert.ToInt32(row["PLAYS"]);
                song.Artist = row["ARTISTS"] == DBNull.Value ? string.Empty : row["ARTISTS"].ToString().Trim();
                song.Album = row["ALBUM"] == DBNull.Value ? string.Empty : row["ALBUM"].ToString().Trim();
                songs.Add(song);
            }
            return songs;
        }
    }
}
