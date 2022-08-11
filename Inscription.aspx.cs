using System;
using System.Collections.Generic;
using System.Linq;
using System.Web;
using System.Web.UI;
using System.Web.UI.WebControls;
using System.Data.SqlClient;

namespace Org
{
    public partial class Inscription : System.Web.UI.Page
    {
        SqlCommand ins_comm;
        SqlDataReader ins_dr;
        protected void Page_Load(object sender, EventArgs e)
        {

        }

        protected void Button2_Click(object sender, EventArgs e)
        {
            ins_comm = new SqlCommand("select a.NOM,a.PRENOM,s.LIB_COMPLET from AGENT a,AFFECTATION af,SOCIETE s where a.MATRICULE = af.MATRICULE and af.STE = s.STE and a.MATRICULE = @id ", Connection.con);
            ins_comm.Parameters.AddWithValue("@id", TextBox1.Text);
            try
            {
                Connection.con.Open();
                ins_dr = ins_comm.ExecuteReader();
                if (ins_dr.HasRows)
                {
                    while (ins_dr.Read())
                    {
                        TextBox2.Text = ins_dr[0].ToString();
                        TextBox3.Text = ins_dr[1].ToString();
                        TextBox4.Text = ins_dr[2].ToString();
                    }
                }
                else
                {
                    Response.Write("<script>alert('Matricule Introuvable');</script>");
                    TextBox1.Text = "";
                    TextBox2.Text = "";
                    TextBox3.Text = "";
                    TextBox4.Text = "";
                }
            }
            catch (Exception ex)
            {
                Response.Write("<script>alert('Erreur 1:" + ex.Message + "');</script>");
            }
            finally
            {
                ins_dr.Close();
                Connection.con.Close();
            }
        }

        protected void Button1_Click1(object sender, EventArgs e)
        {
            int exit=0;
            ins_comm = new SqlCommand("select count (matricule) from authentification where matricule=@id", Connection.con);
            ins_comm.Parameters.AddWithValue("@id", TextBox1.Text);
            try
            {
                Connection.con.Open();
                exit = (int)ins_comm.ExecuteScalar();
                if (exit == 1)
                {
                    Response.Write("<script>alert('Cet Utilisateur existe déja');</script>");
                }
                else
                {
                    ins_comm = new SqlCommand("", Connection.con);
                    ins_comm.Parameters.AddWithValue("@id", TextBox1.Text);
                    ins_comm.Parameters.AddWithValue("@type", DropDownList1.Text);
                    ins_comm.Parameters.AddWithValue("@mdp", TextBox5.Text);

                    ins_comm.CommandText = "insert into Authentification values(@id,@mdp,@type)";
                    try
                    {
                        ins_comm.ExecuteNonQuery();
                        Response.Write("<script>alert('L insertion effectueé avec Succeés');</script>");
                    }
                    catch (Exception ex)
                    {
                        Response.Write("<script>alert('Erreur 2:" + ex.Message + "');</script>");
                    }
                }
            }
            catch (Exception ex)
            {
                Response.Write("<script>alert('Erreur 3:" + ex.Message +"');</script>");
            }
            finally
            {
                Connection.con.Close();
            }
        }

        
    }
}