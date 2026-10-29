using System;
using System.Collections.Generic;
using System.Data;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace BLL.PoliMusicV3
{
    /// <summary>
    /// Business rules for the songs. It only coordinates the validations and
    /// delegates every database operation to the DAL layer.
    /// </summary>
    public class Song
    {
        private string strConnectionString;
        public Song(string strConnString)
        {
            strConnectionString = strConnString;
        }
        public int Save(EntityLayer.PoliMusicV3.Song song)
        {
            return new DAL.PoliMusicV3.Song(strConnectionString).Save(song);
        }
        public DataTable Read()
        {
            return new DAL.PoliMusicV3.Song(strConnectionString).Read();
        }
        public EntityLayer.PoliMusicV3.Song GetById(int id)
        {
            return new DAL.PoliMusicV3.Song(strConnectionString).GetById(id);
        }
        public int Update(EntityLayer.PoliMusicV3.Song song)
        {
            return new DAL.PoliMusicV3.Song(strConnectionString).Update(song);
        }
        public int Delete(int id)
        {
            return new DAL.PoliMusicV3.Song(strConnectionString).Delete(id);
        }
        public int IncrementPlays(int id)
        {
            return new DAL.PoliMusicV3.Song(strConnectionString).IncrementPlays(id);
        }
        public int CheckExistSong(string songName)
        {
            return new DAL.PoliMusicV3.Song(strConnectionString).CheckExistSong(songName);
        }
        /// <summary>
        /// Validates the required fields of the entity before saving it.
        /// </summary>
        /// <returns>An empty string when everything is fine, otherwise the error message.</returns>
        public string ValidationsSong(EntityLayer.PoliMusicV3.Song song)
        {
            if (song == null)
            {
                return "The song is required!";
            }
            if (string.IsNullOrWhiteSpace(song.SongName))
            {
                return "Song name is required!";
            }
            if (string.IsNullOrWhiteSpace(song.SongPath))
            {
                return "Song path is required!";
            }
            if (song.Plays < 0)
            {
                return "Plays cannot be a negative number!";
            }
            return string.Empty;
        }
        /// <summary>
        /// Validates the required fields plus the duplicated name before inserting.
        /// </summary>
        /// <returns>An empty string when everything is fine, otherwise the error message.</returns>
        public string ValidationsNewSong(EntityLayer.PoliMusicV3.Song song)
        {
            string messageValidations = ValidationsSong(song);
            if (messageValidations != string.Empty)
            {
                return messageValidations;
            }
            if (CheckExistSong(song.SongName.Trim()) != 0)
            {
                return "Song name already registered!";
            }
            return string.Empty;
        }
    }

}
