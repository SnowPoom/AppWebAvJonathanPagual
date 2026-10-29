using System;
using System.Configuration;
using System.Data;
using System.Web.UI;
using System.Web.UI.WebControls;

namespace WebAppPoliMusicV3
{
    public partial class Song : System.Web.UI.Page
    {
        string strConnString = ConfigurationManager.ConnectionStrings["BDD_PoliMusicConnectionString"].ConnectionString;

        protected void Page_Load(object sender, EventArgs e)
        {
            if (!Page.IsPostBack)
            {
                lblUserName.Text = Server.HtmlEncode(GetCurrentUserName());
                BindSongs();
            }
        }

        /// <summary>
        /// Reads the logged user from the session (it is stored on login) and falls back
        /// to the forms authentication name when the session is not available.
        /// </summary>
        private string GetCurrentUserName()
        {
            EntityLayer.PoliMusicV3.User user = Session[EntityLayer.PoliMusicV3.Util.Constants.USER] as EntityLayer.PoliMusicV3.User;
            if (user != null && !string.IsNullOrWhiteSpace(user.UserName))
            {
                return user.UserName;
            }
            if (Page.User != null && Page.User.Identity != null && !string.IsNullOrWhiteSpace(Page.User.Identity.Name))
            {
                return Page.User.Identity.Name;
            }
            return "invitado";
        }

        /// <summary>
        /// Fills the repeater with the songs returned by the business layer.
        /// </summary>
        private void BindSongs()
        {
            DataTable dtSongs = new BLL.PoliMusicV3.Song(strConnString).Read();
            rptSongs.DataSource = dtSongs;
            rptSongs.DataBind();
            pnlEmpty.Visible = (dtSongs == null || dtSongs.Rows.Count == 0);
            lblTotal.Text = (dtSongs == null) ? "0" : dtSongs.Rows.Count.ToString();
        }

        /// <summary>
        /// Attaches the client side playback to every play button. The audio element lives
        /// outside the update panel, so the song keeps playing while the counter is refreshed.
        /// </summary>
        protected void rptSongs_ItemDataBound(object sender, RepeaterItemEventArgs e)
        {
            if (e.Item.ItemType != ListItemType.Item && e.Item.ItemType != ListItemType.AlternatingItem)
            {
                return;
            }
            LinkButton btnPlay = (LinkButton)e.Item.FindControl("btnPlay");
            if (btnPlay == null)
            {
                return;
            }
            string songPath = Convert.ToString(DataBinder.Eval(e.Item.DataItem, "SONG_PATH"));
            if (string.IsNullOrWhiteSpace(songPath))
            {
                return;
            }
            string playUrl = ResolveUrl(songPath).Replace("\\", "\\\\").Replace("'", "\\'");
            btnPlay.OnClientClick = "playSong('" + playUrl + "', this);";
        }

        /// <summary>
        /// Counts the play of the selected song and refreshes the list asynchronously.
        /// </summary>
        protected void rptSongs_ItemCommand(object source, RepeaterCommandEventArgs e)
        {
            if (e.CommandName != "Play")
            {
                return;
            }
            int idSong;
            if (!int.TryParse(Convert.ToString(e.CommandArgument), out idSong))
            {
                return;
            }
            new BLL.PoliMusicV3.Song(strConnString).IncrementPlays(idSong);
            BindSongs();
            upSongs.Update();
        }

        /// <summary>
        /// Returns the value received, or the fallback text when it is empty.
        /// </summary>
        protected string Display(object value, string fallback)
        {
            string text = Convert.ToString(value);
            return string.IsNullOrWhiteSpace(text) ? fallback : text;
        }

        /// <summary>
        /// Returns the word "reproducción" or "reproducciones" depending on the counter.
        /// </summary>
        protected string PlaysLabel(object value)
        {
            int plays;
            int.TryParse(Convert.ToString(value), out plays);
            return (plays == 1) ? "reproducción" : "reproducciones";
        }
    }

}
