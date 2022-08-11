using System;
using System.Collections.Generic;
using System.Linq;
using System.Web;
using System.Web.UI;
using System.Web.UI.WebControls;
using System.Data.SqlClient;

namespace Org
{
    public partial class Authentification : System.Web.UI.Page
    {
        SqlCommand ind_comm;
        protected void Page_Load(object sender, EventArgs e)
        {

        }

        protected void Button1_Click(object sender, EventArgs e)
        {
            try
            {
                ind_comm = new SqlCommand("select count(*) from Authentification where MATRICULE=@pseudo1 and MOT_DE_PASSE=@pass1", Connection.con);
                Connection.con.Open();
                ind_comm.Parameters.AddWithValue("@pseudo1", TextBox1.Text);
                ind_comm.Parameters.AddWithValue("@pass1", TextBox2.Text);
                int exist = (int)ind_comm.ExecuteScalar();

                if (exist == 0)
                {
                    Response.Write("<script>alert('Il ya aucun Compte parrapport a les informations Entreés');</script>");
                    TextBox1.Text = " ";
                    TextBox2.Text = " ";
                    TextBox1.Focus();

                }
                else if (exist == 1)
                {
                    Session["pseudo"] = TextBox1.Text;
                    Response.Redirect("OrgPage.aspx");
                }
            }
            catch (Exception ex)
            {
                Response.Write("<script>alert('Erreur :" + ex.Message + "');</script>");
            }
            finally
            {
                Connection.con.Close();
            }
        }
    }
}