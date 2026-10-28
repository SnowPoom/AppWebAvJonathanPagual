using System;

namespace WebAppPoliMusicV2.Entities
{
    /// <summary>
    /// Song entity. Only attributes and their getters/setters.
    /// Maps the columns of TBL_SONG: ID_SONG, SONG_NAME, SONG_PATH, PLAYS.
    /// </summary>
    public class Song
    {
        public Song()
        {
        }

        public Song(int id, string name, string filePath, int plays)
        {
            ID = id;
            Name = name;
            FilePath = filePath;
            Plays = plays;
        }

        public int ID { get; set; }
        public string Name { get; set; }
        /// <summary>Virtual path of the mp3 file, e.g. ~/songFiles/bensound-adventure.mp3</summary>
        public string FilePath { get; set; }
        /// <summary>Number of times the song has been played (TBL_SONG.PLAYS).</summary>
        public int Plays { get; set; }

        /// <summary>
        /// Not a column of TBL_SONG: it is filled by the catalog query
        /// (TBL_PLAYER + TBL_ARTIST) and it is only read by the presentation layer.
        /// </summary>
        public string Artist { get; set; }

        /// <summary>
        /// Not a column of TBL_SONG: it is filled by the catalog query
        /// (TBL_TRACKLIST + TBL_ALBUM) and it is only read by the presentation layer.
        /// </summary>
        public string Album { get; set; }
    }
}
