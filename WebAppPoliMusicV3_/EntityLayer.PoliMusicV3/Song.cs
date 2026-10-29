using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace EntityLayer.PoliMusicV3
{
    /// <summary>
    /// Represents a row of the [TBL_SONG] table.
    /// This class only exposes properties (getters and setters), no logic at all.
    /// </summary>
    public class Song
    {
        public Song()
        {
        }
        public Song(int id, string songName, string songPath, int plays)
        {
            ID = id;
            SongName = songName;
            SongPath = songPath;
            Plays = plays;
        }
        public int ID { get; set; }
        public string SongName { get; set; }
        public string SongPath { get; set; }
        public int Plays { get; set; }
    }

}
