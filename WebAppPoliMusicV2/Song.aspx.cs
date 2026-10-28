using System;
using System.Collections.Generic;
using System.Configuration;
using System.Web;
using System.Web.UI;

namespace WebAppPoliMusicV2
{
    public partial class Song : System.Web.UI.Page
    {
        private string strConnString = ConfigurationManager.ConnectionStrings["BDD_PoliMusicConnectionString"].ConnectionString;

        protected void Page_Load(object sender, EventArgs e)
        {
            if (!Page.IsPostBack)
            {
                LoadCurrentUser();
                LoadSongs();
            }
        }

        /// <summary>
        /// Shows the name of the user stored in session.
        /// </summary>
        private void LoadCurrentUser()
        {
            Entities.User user = Session[Util.Constants.USER] as Entities.User;
            lblUserName.Text = (user == null) ? string.Empty : Server.HtmlEncode(user.UserName);
        }

        /// <summary>
        /// Asks the BLL for the catalog and binds it to the list.
        /// </summary>
        private void LoadSongs()
        {
            try
            {
                List<Entities.Song> songs = new BLL.Song(strConnString).GetSongs();

                rptSongs.DataSource = songs;
                rptSongs.DataBind();

                pnlEmpty.Visible = (songs.Count == 0);
                lblCount.Text = (songs.Count == 1) ? "1 canción" : songs.Count + " canciones";
            }
            catch (Exception ex)
            {
                lblMessage.Text = "No se pudieron cargar las canciones: " + Server.HtmlEncode(ex.Message);
                lblMessage.Visible = true;
                pnlEmpty.Visible = false;
                lblCount.Text = string.Empty;
            }
        }

        /// <summary>
        /// Encodes a value so it can be safely written inside an HTML attribute or as text.
        /// Single quotes are encoded too, because the data binding expressions use them
        /// as the attribute delimiter (double quotes inside <%# %> break the parser).
        /// </summary>
        protected string Attr(object value)
        {
            return HttpUtility.HtmlAttributeEncode(Convert.ToString(value)).Replace("'", "&#39;");
        }
    }
}
