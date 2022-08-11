using System;
using System.Collections.Generic;
using System.Linq;
using System.Web;
using System.Web.UI;
using Newtonsoft.Json;
using System.Data.SqlClient;
using System.Configuration;
using System.Data;
using System.IO;
using System.Web.UI.WebControls;
using Microsoft.Office.Interop.Excel;
using DocumentFormat.OpenXml.Packaging;
using Ganss.Excel;
using ClosedXML.Excel;
using Newtonsoft.Json.Linq;
using System.Threading.Tasks;
using System.Web.Services;

namespace Org
{
    public partial class OrgPage : System.Web.UI.Page
    {
        
        SqlCommand admin_comm;
        SqlDataAdapter myCommand;
        SqlCommand cmd;
        System.Data.DataTable dt;
        string global = "select distinct a.MATRICULE as id , a.MATRIC_SUPERIEUR as pid , g.NOM + ' ' +g.prenom as name , t.INTITULE as title, c.GRADE as grade,CENTRE_GESTION from  CLASSIFICATION c, AFFECTATION a,AGENT g ,TABLE_REFERENCE t where a.MATRICULE = g.MATRICULE and CODE = c.FONCTION and c.MATRICULE = g.MATRICULE and DROIT_A_LA_PAIE=1 and FONCTION = t.CODE and TABLE_NAME = 'FONCTION'";
        string visiteur = "select distinct a.MATRICULE as id , a.MATRIC_SUPERIEUR as pid , g.NOM + ' ' +g.prenom as name , t.INTITULE as title, c.GRADE as grade,CENTRE_GESTION from SERVICE se, CLASSIFICATION c, AFFECTATION a,AGENT g ,TABLE_REFERENCE t where a.MATRICULE = g.MATRICULE and CODE= c.FONCTION and c.MATRICULE = g.MATRICULE and a.SCE = se.SCE and se.SCE = (select SCE from AFFECTATION where MATRICULE = ";
        string societe = "select distinct a.MATRICULE as id , a.MATRIC_SUPERIEUR as pid , g.NOM + ' ' +g.prenom as name , t.INTITULE as title, c.GRADE as grade,CENTRE_GESTION from  SOCIETE st, CLASSIFICATION c, AFFECTATION a,AGENT g ,TABLE_REFERENCE t where a.MATRICULE = g.MATRICULE and CODE= c.FONCTION and c.MATRICULE = g.MATRICULE and DROIT_A_LA_PAIE=1 and FONCTION = t.CODE and TABLE_NAME = 'FONCTION' and st.STE = a.STE and (a.MATRICULE = 'H00001'or a.MATRICULE = 'H00029' or st.LIB_COMPLET =";
        string direction = "select distinct a.MATRICULE as id , a.MATRIC_SUPERIEUR as pid , g.NOM + ' ' +g.prenom as name , t.INTITULE as title, c.GRADE as grade,CENTRE_GESTION from SOCIETE s, SERVICE sr, CLASSIFICATION c, AFFECTATION a,AGENT g ,TABLE_REFERENCE t where a.MATRICULE = g.MATRICULE and CODE= c.FONCTION and c.MATRICULE = g.MATRICULE and sr.SCE = a.SCE and s.STE=a.STE  and DROIT_A_LA_PAIE=1 and s.STE = sr.STE and (a.MATRICULE = 'H00001'or a.MATRICULE = 'H00029' or sr.LIB_COMPLET = ";
        string region = "with A as (select distinct * from TABLE_REFERENCE where TABLE_NAME = 'FILIERE') select distinct a.MATRICULE as id , a.MATRIC_SUPERIEUR as pid , g.NOM + ' ' +g.prenom as name , t.INTITULE as title, c.GRADE as grade, CENTRE_GESTION  from  CLASSIFICATION c, SOCIETE se ,AFFECTATION a, AGENT g, TABLE_REFERENCE t, A s where a.MATRICULE = g.MATRICULE and t.CODE= c.FONCTION and c.MATRICULE = g.MATRICULE and DROIT_A_LA_PAIE=1 and c.FILIERE = s.CODE and se.STE = a.STE and s.INTITULE = ";
        string regglob = "select distinct intitule from TABLE_REFERENCE where TABLE_NAME='FILIERE'";
        string dir_soc = "select distinct se.LIB_COMPLET from SERVICE se,SOCIETE st where se.STE = st.STE and st.LIB_COMPLET = ";

        string excel = "select distinct a.MATRICULE , a.MATRIC_SUPERIEUR , g.NOM ,g.PRENOM ,sr.LIB_COMPLET,CENTRE_GESTION ,t.INTITULE as FONCTION, c.GRADE from SOCIETE s, SERVICE sr, CLASSIFICATION c, AFFECTATION a,AGENT g ,TABLE_REFERENCE t where a.MATRICULE = g.MATRICULE and CODE= c.FONCTION and c.MATRICULE = g.MATRICULE and sr.SCE = a.SCE and s.STE=a.STE and s.STE = sr.STE and s.LIB_COMPLET = ";

        string grades = "select distinct GRADE from CLASSIFICATION c,AFFECTATION af,AGENT a where a.MATRICULE = af.MATRICULE and c.MATRICULE = a.MATRICULE and DROIT_A_LA_PAIE = 1 ";

        //string path = "C:/Users/Administrateur/Desktop/Org/Org/Scripts/org6.json";
        //string expath = "C:/Users/Administrateur/Desktop/Org/Org/Scripts/excel.json";
        //string Gpath = "C:/Users/Administrateur/Desktop/Org/Org/Scripts/all.json";
        //string Cpath= "C:/Users/Administrateur/Desktop/Org/Org/Scripts/common.json";

        string path = "C:/Temp/Websites/Menara-orgs/Scripts/org.json";
        string expath = "C:/Temp/Websites/Menara-orgs/Scripts/excel.json";
        string Gpath = "C:/Temp/Websites/Menara-orgs/Scripts/all.json";
        string Cpath = "C:/Temp/Websites/Menara-orgs/Scripts/common.json";

        //[WebMethod]
        //public static void cleardata()
        //{
        //    System.IO.File.WriteAllText("C:/Temp/Websites/Menara-orgs/Scripts/org.json", "{}");
        //}
        protected void Page_Load(object sender, EventArgs e)
        {
            string type = "";
            if (!IsPostBack)
            {
                admin_comm = new SqlCommand("select type from Authentification where MATRICULE='" + Session["pseudo"] + "'", Connection.con);
                try
                {
                    Connection.con.Open();
                    type = (string)admin_comm.ExecuteScalar();
                }
                catch (Exception ex)
                {
                    Response.Write("<script>alert('Erreur 1:" + ex.Message + "');</script>");
                }
                finally
                {
                    Connection.con.Close();
                }
            }
            if (type == "Admin")
            {
                btn_aj.Visible = true;
                btn_dir.Visible = true;
                btn_soc.Visible = true;
                string json = DataTableToJSONWithJSONNet(GetData(global));
                System.IO.File.WriteAllText(path, json);
                FillDirection("select * from SERVICE");
                ViewState["st"] = "";

            }
            else if (type == "Visiteur")
            {
                btn_aj.Visible = false;
                btn_dir.Visible = false;
                btn_soc.Visible = false;
                string json  = DataTableToJSONWithJSONNet(GetData(visiteur + "'" + Session["pseudo"].ToString() + "')"));
                System.IO.File.WriteAllText(path, json);
                FillDirection("select se.LIB_COMPLET from AFFECTATION,SERVICE se,SOCIETE st where se.STE = st.STE and AFFECTATION.STE = se.STE and MATRICULE ='" + Session["pseudo"].ToString() + "'");
            }
            else if (type == "RH")
            {
                btn_aj.Visible = false;
                btn_dir.Visible = true;
                btn_soc.Visible = true;
                string json = DataTableToJSONWithJSONNet(GetData(global));
                System.IO.File.WriteAllText(path, json);
                FillDirection("select * from SERVICE");
                ViewState["st"] = "";
            }
            else if (type == "DG")
            {
                btn_aj.Visible = false;
                btn_dir.Visible = true;
                btn_soc.Visible = false;
                string json = DataTableToJSONWithJSONNet(GetData("select a.MATRICULE as id, a.MATRIC_SUPERIEUR as pid,CENTRE_GESTION ,g.NOM + ' ' + g.prenom as name, t.INTITULE as title, c.GRADE as grade from SOCIETE st, CLASSIFICATION c, AFFECTATION a, AGENT g, TABLE_REFERENCE t where a.MATRICULE = g.MATRICULE and CODE = c.FONCTION and c.MATRICULE = g.MATRICULE and st.STE = a.STE and st.STE = (select STE from AFFECTATION where MATRICULE =" + "'" + Session["pseudo"].ToString() + "')"));
                System.IO.File.WriteAllText(path, json);
                FillDirection("select se.LIB_COMPLET from AFFECTATION,SERVICE se,SOCIETE st where se.STE = st.STE and AFFECTATION.STE = se.STE and MATRICULE ='" + Session["pseudo"].ToString() + "'");
            }
            if(type == "DG" || type == "Visiteur")
            {

            }
            //add loop for color changing
            //string all = DataTableToJSONWithJSONNet(GetData(global));
            //System.IO.File.WriteAllText(Gpath, all);
        }

        private void Excel(System.Data.DataTable t)
        {
            //Connection.con.Open();
            //cmd = new SqlCommand(query, Connection.con);



            //dt = new System.Data.DataTable();
            //dt.Load(cmd.ExecuteReader());
            //Connection.con.Close();



            using (XLWorkbook wb = new XLWorkbook())
            {
                wb.Worksheets.Add(t, "Customers");



                Response.Clear();
                Response.Buffer = true;
                Response.Charset = "";
                Response.ContentType = "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet";
                Response.AddHeader("content-disposition", "attachment;filename=List.xlsx");
                using (MemoryStream MyMemoryStream = new MemoryStream())
                {
                    wb.SaveAs(MyMemoryStream);
                    MyMemoryStream.WriteTo(Response.OutputStream);
                    Response.Flush();
                    Response.End();
                }
            }
        }

        private System.Data.DataTable Json_table()
        {
            string data = File.ReadAllText(expath);
            System.Data.DataTable dt = (System.Data.DataTable)JsonConvert.DeserializeObject(data, (typeof(System.Data.DataTable)));
            return dt;
        }

        private void FillDirection(string query)
        {
            Text1.Items.Clear();
            myCommand = new SqlDataAdapter(query, Connection.con);
            DataSet ds = new DataSet();
            myCommand.Fill(ds, "SERVICE");
            Text1.DataSource = ds;
            Text1.DataTextField = "LIB_COMPLET";
            Text1.DataValueField = "LIB_COMPLET";
            Text1.DataBind();
            Text1.Items.Insert(0, "Direction");
            ListItem I = new ListItem("Direction");
            if(!Text1.Items.Contains(I))
                Text1.Items.Insert(0, "Direction");
        }
        private void FillREGION(string query)
        {
            Text2.Items.Clear();
            myCommand = new SqlDataAdapter(query, Connection.con);
            DataSet ds = new DataSet();
            myCommand.Fill(ds, "TABLE_REFERENCE");
            Text2.DataSource = ds;
            Text2.DataTextField = "INTITULE";
            Text2.DataValueField = "INTITULE";
            Text2.DataBind();
            Text2.Items.Insert(0, "Region");
            ListItem I = new ListItem("Region");
            if (!Text2.Items.Contains(I))
                Text2.Items.Insert(0, "Region");
        }


        private void FillGrade(string query)
        {
            Text3.Items.Clear();
            myCommand = new SqlDataAdapter(query, Connection.con);
            DataSet ds = new DataSet();
            myCommand.Fill(ds, "CLASSIFICATION");
            Text3.DataSource = ds;
            Text3.DataTextField = "GRADE";
            Text3.DataValueField = "GRADE";
            Text3.DataBind();
            Text3.Items.Insert(0, "Grade");
            ListItem I = new ListItem("Grade");
            if (!Text3.Items.Contains(I))
                Text3.Items.Insert(0, "Grade");
        }



        public System.Data.DataTable GetData(string query)
        {
            Connection.con.Open();
            cmd = new SqlCommand(query, Connection.con);

            dt = new System.Data.DataTable();
            dt.Load(cmd.ExecuteReader());
            Connection.con.Close();
            return dt;
        }


        public string DataTableToJSONWithJSONNet(System.Data.DataTable table)
        {
            string JSONString = string.Empty;
            JSONString = JsonConvert.SerializeObject(table);
            return JSONString;
        }


        //protected void Dicrection(object sender, EventArgs e)
        //{
        //    string json = DataTableToJSONWithJSONNet(GetData(direction +"'" + Text1.Text + "'"));
        //    System.IO.File.WriteAllText(path, json);
        //}


        protected void MH_Click(object sender, EventArgs e)
        {
            ViewState["st"] = "MH";
            ScriptManager.RegisterStartupScript(this, GetType(), "MH", "MH();", true);
            string json = DataTableToJSONWithJSONNet(GetData(societe + "'MENARA HOLDING')"));
            System.IO.File.WriteAllText(path, json);
            FillREGION(regglob);
            FillDirection(dir_soc + "'MENARA HOLDING'");
            FillGrade(grades);
            json = DataTableToJSONWithJSONNet(GetData("with k as( select distinct a.MATRIC_SUPERIEUR as pid from  CLASSIFICATION c, AFFECTATION a,AGENT g ,TABLE_REFERENCE t where a.MATRICULE = g.MATRICULE and CODE = c.FONCTION and c.MATRICULE = g.MATRICULE and DROIT_A_LA_PAIE=1 and FONCTION = t.CODE and TABLE_NAME = 'FONCTION' and a.MATRICULE like 'H%' and MATRIC_SUPERIEUR NOT LIKE 'H%' and MATRIC_SUPERIEUR not like '') select distinct a.MATRICULE as id , a.MATRIC_SUPERIEUR as pid , g.NOM + ' ' +g.prenom as name , t.INTITULE as title, c.GRADE as grade,CENTRE_GESTION from  CLASSIFICATION c, AFFECTATION a,AGENT g ,TABLE_REFERENCE t, k where a.MATRICULE = g.MATRICULE and CODE = c.FONCTION and c.MATRICULE = g.MATRICULE and DROIT_A_LA_PAIE=1 and FONCTION = t.CODE and TABLE_NAME = 'FONCTION'and k.pid=a.MATRICULE"));
            System.IO.File.WriteAllText(Cpath, json);
        }

        protected void MP_Click(object sender, EventArgs e)
        {
            ScriptManager.RegisterStartupScript(this, GetType(), "MP", "MP();", true);
            string json = DataTableToJSONWithJSONNet(GetData(societe + "'MENARA PREFA')"));
            System.IO.File.WriteAllText(path, json);
            FillDirection(dir_soc + "'MENARA PREFA'");
            ViewState["st"] = "MP";
            FillREGION(regglob);
            FillGrade(grades);
            json = DataTableToJSONWithJSONNet(GetData("with k as( select distinct a.MATRIC_SUPERIEUR as pid from  CLASSIFICATION c, AFFECTATION a,AGENT g ,TABLE_REFERENCE t where a.MATRICULE = g.MATRICULE and CODE = c.FONCTION and c.MATRICULE = g.MATRICULE and DROIT_A_LA_PAIE=1 and FONCTION = t.CODE and TABLE_NAME = 'FONCTION' and a.MATRICULE like 'P%' and MATRIC_SUPERIEUR NOT LIKE 'P%' and MATRIC_SUPERIEUR not like '') select distinct a.MATRICULE as id , a.MATRIC_SUPERIEUR as pid , g.NOM + ' ' +g.prenom as name , t.INTITULE as title, c.GRADE as grade,CENTRE_GESTION from  CLASSIFICATION c, AFFECTATION a,AGENT g ,TABLE_REFERENCE t, k where a.MATRICULE = g.MATRICULE and CODE = c.FONCTION and c.MATRICULE = g.MATRICULE and DROIT_A_LA_PAIE=1 and FONCTION = t.CODE and TABLE_NAME = 'FONCTION'and k.pid=a.MATRICULE"));
            System.IO.File.WriteAllText(Cpath, json);
        }

        protected void CT_Click(object sender, EventArgs e)
        {
            ViewState["st"] = "CT";
            ScriptManager.RegisterStartupScript(this, GetType(), "CT", "CT();", true);
            string json = DataTableToJSONWithJSONNet(GetData(societe + "'CARRIERES ET TRANSPORT MENARA')"));
            System.IO.File.WriteAllText(path, json);
            FillDirection(dir_soc + "'CARRIERES ET TRANSPORT MENARA'");
            FillREGION(regglob);
            FillGrade(grades);
            json = DataTableToJSONWithJSONNet(GetData("with k as( select distinct a.MATRIC_SUPERIEUR as pid from  CLASSIFICATION c, AFFECTATION a,AGENT g ,TABLE_REFERENCE t where a.MATRICULE = g.MATRICULE and CODE = c.FONCTION and c.MATRICULE = g.MATRICULE and DROIT_A_LA_PAIE=1 and FONCTION = t.CODE and TABLE_NAME = 'FONCTION' and a.MATRICULE like 'T%' and MATRIC_SUPERIEUR NOT LIKE 'T%' and MATRIC_SUPERIEUR not like '') select distinct a.MATRICULE as id , a.MATRIC_SUPERIEUR as pid , g.NOM + ' ' +g.prenom as name , t.INTITULE as title, c.GRADE as grade,CENTRE_GESTION from  CLASSIFICATION c, AFFECTATION a,AGENT g ,TABLE_REFERENCE t, k where a.MATRICULE = g.MATRICULE and CODE = c.FONCTION and c.MATRICULE = g.MATRICULE and DROIT_A_LA_PAIE=1 and FONCTION = t.CODE and TABLE_NAME = 'FONCTION'and k.pid=a.MATRICULE"));
            System.IO.File.WriteAllText(Cpath, json);
        }

        protected void TC_Click(object sender, EventArgs e)
        {
            ViewState["st"] = "TCGM";
            ScriptManager.RegisterStartupScript(this, GetType(), "TC", "TC();", true);
            string json = DataTableToJSONWithJSONNet(GetData(societe + "'TRAVAUX DE CONSTRUCTION GENERALE MENARA')"));
            System.IO.File.WriteAllText(path, json);
            FillDirection(dir_soc + "'TRAVAUX DE CONSTRUCTION GENERALE MENARA'");
            FillREGION(regglob);
            json = DataTableToJSONWithJSONNet(GetData("with k as( select distinct a.MATRIC_SUPERIEUR as pid from  CLASSIFICATION c, AFFECTATION a,AGENT g ,TABLE_REFERENCE t where a.MATRICULE = g.MATRICULE and CODE = c.FONCTION and c.MATRICULE = g.MATRICULE and DROIT_A_LA_PAIE=1 and FONCTION = t.CODE and TABLE_NAME = 'FONCTION' and a.MATRICULE like 'V%' and MATRIC_SUPERIEUR NOT LIKE 'V%' and MATRIC_SUPERIEUR not like '') select distinct a.MATRICULE as id , a.MATRIC_SUPERIEUR as pid , g.NOM + ' ' +g.prenom as name , t.INTITULE as title, c.GRADE as grade,CENTRE_GESTION from  CLASSIFICATION c, AFFECTATION a,AGENT g ,TABLE_REFERENCE t, k where a.MATRICULE = g.MATRICULE and CODE = c.FONCTION and c.MATRICULE = g.MATRICULE and DROIT_A_LA_PAIE=1 and FONCTION = t.CODE and TABLE_NAME = 'FONCTION'and k.pid=a.MATRICULE"));
            System.IO.File.WriteAllText(Cpath, json);

        }

        protected void ML_Click(object sender, EventArgs e)
        {
            ScriptManager.RegisterStartupScript(this, GetType(), "ML", "ML();", true);
            string json = DataTableToJSONWithJSONNet(GetData(societe + "'MENARA LOGISTIQUE')"));
            System.IO.File.WriteAllText(path, json);
            FillDirection(dir_soc + "'MENARA LOGISTIQUE'");
            ViewState["st"] = "ML";
            FillREGION(regglob);
            FillGrade(grades);
            json = DataTableToJSONWithJSONNet(GetData("with k as( select distinct a.MATRIC_SUPERIEUR as pid from  CLASSIFICATION c, AFFECTATION a,AGENT g ,TABLE_REFERENCE t where a.MATRICULE = g.MATRICULE and CODE = c.FONCTION and c.MATRICULE = g.MATRICULE and DROIT_A_LA_PAIE=1 and FONCTION = t.CODE and TABLE_NAME = 'FONCTION' and a.MATRICULE like 'E%' and MATRIC_SUPERIEUR NOT LIKE 'E%' and MATRIC_SUPERIEUR not like '') select distinct a.MATRICULE as id , a.MATRIC_SUPERIEUR as pid , g.NOM + ' ' +g.prenom as name , t.INTITULE as title, c.GRADE as grade,CENTRE_GESTION from  CLASSIFICATION c, AFFECTATION a,AGENT g ,TABLE_REFERENCE t, k where a.MATRICULE = g.MATRICULE and CODE = c.FONCTION and c.MATRICULE = g.MATRICULE and DROIT_A_LA_PAIE=1 and FONCTION = t.CODE and TABLE_NAME = 'FONCTION'and k.pid=a.MATRICULE"));
            System.IO.File.WriteAllText(Cpath, json);
        }

        protected void MT_Click(object sender, EventArgs e)
        {
            ScriptManager.RegisterStartupScript(this, GetType(), "MTL", "MTL();", true);
            string json = DataTableToJSONWithJSONNet(GetData(societe + "'MENARA TRANSPORT ET LOGISTIQUE')"));
            System.IO.File.WriteAllText(path, json);
            FillDirection(dir_soc + "'MENARA TRANSPORT ET LOGISTIQUE'");
            ViewState["st"] = "MTL";
            FillREGION(regglob);
            FillGrade(grades);
            json = DataTableToJSONWithJSONNet(GetData("with k as( select distinct a.MATRIC_SUPERIEUR as pid from  CLASSIFICATION c, AFFECTATION a,AGENT g ,TABLE_REFERENCE t where a.MATRICULE = g.MATRICULE and CODE = c.FONCTION and c.MATRICULE = g.MATRICULE and DROIT_A_LA_PAIE=1 and FONCTION = t.CODE and TABLE_NAME = 'FONCTION' and a.MATRICULE like 'S%' and MATRIC_SUPERIEUR NOT LIKE 'S%' and MATRIC_SUPERIEUR not like '') select distinct a.MATRICULE as id , a.MATRIC_SUPERIEUR as pid , g.NOM + ' ' +g.prenom as name , t.INTITULE as title, c.GRADE as grade,CENTRE_GESTION from  CLASSIFICATION c, AFFECTATION a,AGENT g ,TABLE_REFERENCE t, k where a.MATRICULE = g.MATRICULE and CODE = c.FONCTION and c.MATRICULE = g.MATRICULE and DROIT_A_LA_PAIE=1 and FONCTION = t.CODE and TABLE_NAME = 'FONCTION'and k.pid=a.MATRICULE"));
            System.IO.File.WriteAllText(Cpath, json);
        }

        protected void IT_Click(object sender, EventArgs e)
        {
            ScriptManager.RegisterStartupScript(this, GetType(), "IT", "IT();", true);
            string json = DataTableToJSONWithJSONNet(GetData(societe + "'STE IMMOBILIERE DU TENSIFT')"));
            System.IO.File.WriteAllText(path, json);
            FillDirection(dir_soc + "'STE IMMOBILIERE DU TENSIFT'");
            ViewState["st"] = "IT";
            FillREGION(regglob);
            FillGrade(grades);
            json = DataTableToJSONWithJSONNet(GetData("with k as( select distinct a.MATRIC_SUPERIEUR as pid from  CLASSIFICATION c, AFFECTATION a,AGENT g ,TABLE_REFERENCE t where a.MATRICULE = g.MATRICULE and CODE = c.FONCTION and c.MATRICULE = g.MATRICULE and DROIT_A_LA_PAIE=1 and FONCTION = t.CODE and TABLE_NAME = 'FONCTION' and a.MATRICULE like 'I%' and MATRIC_SUPERIEUR NOT LIKE 'I%' and MATRIC_SUPERIEUR not like '') select distinct a.MATRICULE as id , a.MATRIC_SUPERIEUR as pid , g.NOM + ' ' +g.prenom as name , t.INTITULE as title, c.GRADE as grade,CENTRE_GESTION from  CLASSIFICATION c, AFFECTATION a,AGENT g ,TABLE_REFERENCE t, k where a.MATRICULE = g.MATRICULE and CODE = c.FONCTION and c.MATRICULE = g.MATRICULE and DROIT_A_LA_PAIE=1 and FONCTION = t.CODE and TABLE_NAME = 'FONCTION'and k.pid=a.MATRICULE"));
            System.IO.File.WriteAllText(Cpath, json);
        }

        protected void FZ_Click(object sender, EventArgs e)
        {
            ScriptManager.RegisterStartupScript(this, GetType(), "FZ", "FZ();", true);
            string json = DataTableToJSONWithJSONNet(GetData(societe + "'FONDATION ZAHID')"));
            System.IO.File.WriteAllText(path, json);
            FillDirection(dir_soc + "'FONDATION ZAHID'");
            ViewState["st"] = "FZ";
            FillREGION(regglob);
            FillGrade(grades);
            json = DataTableToJSONWithJSONNet(GetData("with k as( select distinct a.MATRIC_SUPERIEUR as pid from  CLASSIFICATION c, AFFECTATION a,AGENT g ,TABLE_REFERENCE t where a.MATRICULE = g.MATRICULE and CODE = c.FONCTION and c.MATRICULE = g.MATRICULE and DROIT_A_LA_PAIE=1 and FONCTION = t.CODE and TABLE_NAME = 'FONCTION' and a.MATRICULE like 'Z%' and MATRIC_SUPERIEUR NOT LIKE 'Z%' and MATRIC_SUPERIEUR not like '') select distinct a.MATRICULE as id , a.MATRIC_SUPERIEUR as pid , g.NOM + ' ' +g.prenom as name , t.INTITULE as title, c.GRADE as grade,CENTRE_GESTION from  CLASSIFICATION c, AFFECTATION a,AGENT g ,TABLE_REFERENCE t, k where a.MATRICULE = g.MATRICULE and CODE = c.FONCTION and c.MATRICULE = g.MATRICULE and DROIT_A_LA_PAIE=1 and FONCTION = t.CODE and TABLE_NAME = 'FONCTION'and k.pid=a.MATRICULE"));
            System.IO.File.WriteAllText(Cpath, json);
        }

        protected void AM_Click(object sender, EventArgs e)
        {
            ScriptManager.RegisterStartupScript(this, GetType(), "AM", "AM();", true);
            string json = DataTableToJSONWithJSONNet(GetData(societe + "'ACADÉMIE MÉNARA')"));
            System.IO.File.WriteAllText(path, json);
            FillDirection(dir_soc + "'ACADÉMIE MÉNARA'");
            ViewState["st"] = "AM";
            FillREGION(regglob);
            FillGrade(grades);
            json = DataTableToJSONWithJSONNet(GetData("with k as( select distinct a.MATRIC_SUPERIEUR as pid from  CLASSIFICATION c, AFFECTATION a,AGENT g ,TABLE_REFERENCE t where a.MATRICULE = g.MATRICULE and CODE = c.FONCTION and c.MATRICULE = g.MATRICULE and DROIT_A_LA_PAIE=1 and FONCTION = t.CODE and TABLE_NAME = 'FONCTION' and a.MATRICULE like 'D%' and MATRIC_SUPERIEUR NOT LIKE 'D%' and MATRIC_SUPERIEUR not like '') select distinct a.MATRICULE as id , a.MATRIC_SUPERIEUR as pid , g.NOM + ' ' +g.prenom as name , t.INTITULE as title, c.GRADE as grade,CENTRE_GESTION from  CLASSIFICATION c, AFFECTATION a,AGENT g ,TABLE_REFERENCE t, k where a.MATRICULE = g.MATRICULE and CODE = c.FONCTION and c.MATRICULE = g.MATRICULE and DROIT_A_LA_PAIE=1 and FONCTION = t.CODE and TABLE_NAME = 'FONCTION'and k.pid=a.MATRICULE"));
            System.IO.File.WriteAllText(Cpath, json);
        }

        protected void MG_Click(object sender, EventArgs e)
        {
            ScriptManager.RegisterStartupScript(this, GetType(), "MG", "MG();", true);
            string json = DataTableToJSONWithJSONNet(GetData(societe + "'MARRAKECH GRAND PRIX')"));
            System.IO.File.WriteAllText(path, json);
            FillDirection(dir_soc + "'MARRAKECH GRAND PRIX'");
            ViewState["st"] = "MGP";
            FillREGION(regglob);
            FillGrade(grades);
            //json = DataTableToJSONWithJSONNet(GetData("with k as( select distinct a.MATRIC_SUPERIEUR as pid from  CLASSIFICATION c, AFFECTATION a,AGENT g ,TABLE_REFERENCE t where a.MATRICULE = g.MATRICULE and CODE = c.FONCTION and c.MATRICULE = g.MATRICULE and DROIT_A_LA_PAIE=1 and FONCTION = t.CODE and TABLE_NAME = 'FONCTION' and a.MATRICULE like 'G%' and MATRIC_SUPERIEUR NOT LIKE 'G%' and MATRIC_SUPERIEUR not like '') select distinct a.MATRICULE as id , a.MATRIC_SUPERIEUR as pid , g.NOM + ' ' +g.prenom as name , t.INTITULE as title, c.GRADE as grade,CENTRE_GESTION from  CLASSIFICATION c, AFFECTATION a,AGENT g ,TABLE_REFERENCE t, k where a.MATRICULE = g.MATRICULE and CODE = c.FONCTION and c.MATRICULE = g.MATRICULE and DROIT_A_LA_PAIE=1 and FONCTION = t.CODE and TABLE_NAME = 'FONCTION'and k.pid=a.MATRICULE"));
            //System.IO.File.WriteAllText(Cpath, json);
        }

        protected void A_Click(object sender, EventArgs e)
        {
            ScriptManager.RegisterStartupScript(this, GetType(), "A", "A();", true);
            string json = DataTableToJSONWithJSONNet(GetData(societe + "'AAKAR DEVELOPPEMENT')"));
            System.IO.File.WriteAllText(path, json);
            FillDirection(dir_soc + "'AAKAR DEVELOPPEMENT'");
            ViewState["st"] = "A";
            FillREGION(regglob);
            FillGrade(grades);
            json = DataTableToJSONWithJSONNet(GetData("with k as( select distinct a.MATRIC_SUPERIEUR as pid from  CLASSIFICATION c, AFFECTATION a,AGENT g ,TABLE_REFERENCE t where a.MATRICULE = g.MATRICULE and CODE = c.FONCTION and c.MATRICULE = g.MATRICULE and DROIT_A_LA_PAIE=1 and FONCTION = t.CODE and TABLE_NAME = 'FONCTION' and a.MATRICULE like 'A%' and MATRIC_SUPERIEUR NOT LIKE 'A%' and MATRIC_SUPERIEUR not like '') select distinct a.MATRICULE as id , a.MATRIC_SUPERIEUR as pid , g.NOM + ' ' +g.prenom as name , t.INTITULE as title, c.GRADE as grade,CENTRE_GESTION from  CLASSIFICATION c, AFFECTATION a,AGENT g ,TABLE_REFERENCE t, k where a.MATRICULE = g.MATRICULE and CODE = c.FONCTION and c.MATRICULE = g.MATRICULE and DROIT_A_LA_PAIE=1 and FONCTION = t.CODE and TABLE_NAME = 'FONCTION'and k.pid=a.MATRICULE"));
            System.IO.File.WriteAllText(Cpath, json);
        }

        protected void G_Click(object sender, EventArgs e)
        {
            ScriptManager.RegisterStartupScript(this, GetType(), "G", "G();", true);
            string json = DataTableToJSONWithJSONNet(GetData(global));
            System.IO.File.WriteAllText(path, json);
            json = DataTableToJSONWithJSONNet(GetData("with k as( select distinct a.MATRIC_SUPERIEUR as pid from  CLASSIFICATION c, AFFECTATION a,AGENT g ,TABLE_REFERENCE t where a.MATRICULE = g.MATRICULE and CODE = c.FONCTION and c.MATRICULE = g.MATRICULE and DROIT_A_LA_PAIE=1 and FONCTION = t.CODE and TABLE_NAME = 'FONCTION' and a.MATRICULE like 'H%' and MATRIC_SUPERIEUR NOT LIKE 'H%' and MATRIC_SUPERIEUR not like '') select distinct a.MATRICULE as id , a.MATRIC_SUPERIEUR as pid , g.NOM + ' ' +g.prenom as name , t.INTITULE as title, c.GRADE as grade,CENTRE_GESTION from  CLASSIFICATION c, AFFECTATION a,AGENT g ,TABLE_REFERENCE t, k where a.MATRICULE = g.MATRICULE and CODE = c.FONCTION and c.MATRICULE = g.MATRICULE and DROIT_A_LA_PAIE=1 and FONCTION = t.CODE and TABLE_NAME = 'FONCTION'and k.pid=a.MATRICULE"));
            System.IO.File.WriteAllText(Cpath, json);
            FillDirection("select distinct LIB_COMPLET from SERVICE");
            FillREGION(regglob);
            FillGrade(grades);
        }


        protected void Fill(object sender, EventArgs e)
        {
            switch (ViewState["st"])
            {
                case "MH":                    
                    ScriptManager.RegisterStartupScript(this, GetType(), "MH", "MH();", true);
                    string json = DataTableToJSONWithJSONNet(GetData(direction + "'" + Text1.Text + "')" + "and s.LIB_COMPLET = 'MENARA HOLDING'"));
                    System.IO.File.WriteAllText(path, json);
                    break;
                case "MP":
                    ScriptManager.RegisterStartupScript(this, GetType(), "MP", "MP();", true);
                    string json1 = DataTableToJSONWithJSONNet(GetData(direction + "'" + Text1.Text + "')" + "and s.LIB_COMPLET = 'MENARA PREFA'"));
                    System.IO.File.WriteAllText(path, json1);
                    json1 = DataTableToJSONWithJSONNet(GetData("select distinct a.MATRICULE as id , a.MATRIC_SUPERIEUR as pid , g.NOM + ' ' +g.prenom as name , t.INTITULE as title, c.GRADE as grade,CENTRE_GESTION from  CLASSIFICATION c, AFFECTATION a,AGENT g ,TABLE_REFERENCE t where a.MATRICULE = g.MATRICULE and CODE = c.FONCTION and c.MATRICULE = g.MATRICULE and DROIT_A_LA_PAIE=1 and FONCTION = t.CODE and TABLE_NAME = 'FONCTION'and a.MATRICULE='H00001'"));
                    System.IO.File.WriteAllText(Cpath, json1);
                    break;
                case "CT":
                    ScriptManager.RegisterStartupScript(this, GetType(), "CT", "CT();", true);
                    string json2 = DataTableToJSONWithJSONNet(GetData(direction + "'" + Text1.Text + "')" + "and s.LIB_COMPLET = 'CARRIERES ET TRANSPORT MENARA'"));
                    System.IO.File.WriteAllText(path, json2);
                    System.IO.File.WriteAllText(Cpath, "");
                    break;
                case "TCGM":
                    ScriptManager.RegisterStartupScript(this, GetType(), "TC", "TC();", true);
                    string json3 = DataTableToJSONWithJSONNet(GetData(direction + "'" + Text1.Text + "')" + "and s.LIB_COMPLET = 'TRAVAUX DE CONSTRUCTION GENERALE MENARA'"));
                    System.IO.File.WriteAllText(path, json3);
                    break;
                case "ML":
                    ScriptManager.RegisterStartupScript(this, GetType(), "ML", "ML();", true);
                    string json4 = DataTableToJSONWithJSONNet(GetData(direction + "'" + Text1.Text + "')" + "and s.LIB_COMPLET = 'MENARA LOGISTIQUE'"));
                    System.IO.File.WriteAllText(path, json4);
                    break;
                case "MTL":
                    ScriptManager.RegisterStartupScript(this, GetType(), "MTL", "MTL();", true);
                    string json5 = DataTableToJSONWithJSONNet(GetData(direction + "'" + Text1.Text + "')" + "and s.LIB_COMPLET = 'MENARA TRANSPORT ET LOGISTIQUE'"));
                    System.IO.File.WriteAllText(path, json5);
                    break;
                case "IT":
                    ScriptManager.RegisterStartupScript(this, GetType(), "IT", "IT();", true);
                    string json6 = DataTableToJSONWithJSONNet(GetData(direction + "'" + Text1.Text + "')" + "and s.LIB_COMPLET = 'STE IMMOBILIERE DU TENSIFT'"));
                    System.IO.File.WriteAllText(path, json6);
                    System.IO.File.WriteAllText(Cpath, "");
                    break;
                case "FZ":
                    ScriptManager.RegisterStartupScript(this, GetType(), "FZ", "FZ();", true);
                    string json7 = DataTableToJSONWithJSONNet(GetData(direction + "'" + Text1.Text + "')" + "and s.LIB_COMPLET = 'FONDATION ZAHID'"));
                    System.IO.File.WriteAllText(path, json7);
                    break;
                case "AM":
                    ScriptManager.RegisterStartupScript(this, GetType(), "AM", "AM();", true);
                    string json8 = DataTableToJSONWithJSONNet(GetData(direction + "'" + Text1.Text + "')" + "and s.LIB_COMPLET = 'ACADÉMIE MÉNARA'"));
                    System.IO.File.WriteAllText(path, json8);
                    break;
                case "MGP":
                    ScriptManager.RegisterStartupScript(this, GetType(), "MG", "MG();", true);
                    string json9 = DataTableToJSONWithJSONNet(GetData(direction + "'" + Text1.Text + "')" + "and s.LIB_COMPLET = 'MARRAKECH GRAND PRIX'"));
                    System.IO.File.WriteAllText(path, json9);
                    break;
                case "A":
                    ScriptManager.RegisterStartupScript(this, GetType(), "A", "A();", true);
                    string json10 = DataTableToJSONWithJSONNet(GetData(direction + "'" + Text1.Text + "')" + "and s.LIB_COMPLET = 'AAKAR DEVELOPPEMENT'"));
                    System.IO.File.WriteAllText(path, json10);
                    break;
                default:
                    ScriptManager.RegisterStartupScript(this, GetType(), "G", "G();", true);
                    string json11 = DataTableToJSONWithJSONNet(GetData(direction + "'" + Text1.Text + "')" ));
                    System.IO.File.WriteAllText(path, json11);
                    break;
            }
        }
        protected void FillREG(object sender, EventArgs e)
        {
            string json;
            switch (ViewState["st"])
            {
                case "MH":
                    ScriptManager.RegisterStartupScript(this, GetType(), "MH", "MH();", true);
                     json = DataTableToJSONWithJSONNet(GetData(region +"'"+ Text2.Text + "' and se.LIB_COMPLET = 'MENARA HOLDING'"));
                    System.IO.File.WriteAllText(path, json);
                    break;
                case "MP":
                    ScriptManager.RegisterStartupScript(this, GetType(), "MP", "MP();", true);
                    json = DataTableToJSONWithJSONNet(GetData(region + "'" + Text2.Text + "' and se.LIB_COMPLET = 'MENARA PREFA'"));
                    System.IO.File.WriteAllText(path, json);
                    break;
                case "CT":
                    ScriptManager.RegisterStartupScript(this, GetType(), "CT", "CT();", true);
                     json = DataTableToJSONWithJSONNet(GetData(region + "'" + Text2.Text + "' and se.LIB_COMPLET = 'CARRIERES ET TRANSPORT MENARA'"));
                    System.IO.File.WriteAllText(path, json);
                    break;
                case "TCGM":
                    ScriptManager.RegisterStartupScript(this, GetType(), "TC", "TC();", true);
                    json = DataTableToJSONWithJSONNet(GetData(region + "'" + Text2.Text + "' and se.LIB_COMPLET = 'TRAVAUX DE CONSTRUCTION GENERALE MENARA'"));
                    System.IO.File.WriteAllText(path, json);
                    break;
                case "ML":
                    ScriptManager.RegisterStartupScript(this, GetType(), "ML", "ML();", true);
                    json = DataTableToJSONWithJSONNet(GetData(region + "'" + Text2.Text + "' and se.LIB_COMPLET = 'MENARA LOGISTIQUE'"));
                    System.IO.File.WriteAllText(path, json);
                    break;
                case "MTL":
                    ScriptManager.RegisterStartupScript(this, GetType(), "MTL", "MTL();", true);
                    json = DataTableToJSONWithJSONNet(GetData(region + "'" + Text2.Text + "' and se.LIB_COMPLET = 'MENARA TRANSPORT ET LOGISTIQUE'"));
                    System.IO.File.WriteAllText(path, json);
                    break;
                case "IT":
                    ScriptManager.RegisterStartupScript(this, GetType(), "IT", "IT();", true);
                    json = DataTableToJSONWithJSONNet(GetData(region + "'" + Text2.Text + "' and se.LIB_COMPLET = 'STE IMMOBILIERE DU TENSIFT'"));
                    System.IO.File.WriteAllText(path, json);
                    break;
                case "FZ":
                    ScriptManager.RegisterStartupScript(this, GetType(), "FZ", "FZ();", true);
                    json = DataTableToJSONWithJSONNet(GetData(region + "'" + Text2.Text + "' and se.LIB_COMPLET = 'FONDATION ZAHID'"));
                    System.IO.File.WriteAllText(path, json);
                    break;
                case "AM":
                    ScriptManager.RegisterStartupScript(this, GetType(), "AM", "AM();", true);
                    json = DataTableToJSONWithJSONNet(GetData(region + "'" + Text2.Text + "' and se.LIB_COMPLET = 'ACADÉMIE MÉNARA'"));
                    System.IO.File.WriteAllText(path, json);
                    break;
                case "MGP":
                    ScriptManager.RegisterStartupScript(this, GetType(), "MG", "MG();", true);
                    json = DataTableToJSONWithJSONNet(GetData(region + "'" + Text2.Text + "' and se.LIB_COMPLET = 'MARRAKECH GRAND PRIX'"));
                    System.IO.File.WriteAllText(path, json);
                    break;
                case "A":
                    ScriptManager.RegisterStartupScript(this, GetType(), "A", "A();", true);
                    json = DataTableToJSONWithJSONNet(GetData(region + "'" + Text2.Text + "' and se.LIB_COMPLET = 'AAKAR DEVELOPPEMENT'"));
                    System.IO.File.WriteAllText(path, json);
                    break;
                default:
                    ScriptManager.RegisterStartupScript(this, GetType(), "G", "G();", true);
                    json = DataTableToJSONWithJSONNet(GetData(region + "'" + Text2.Text + "'"));
                    System.IO.File.WriteAllText(path, json);
                    break;
            }
        }


        public bool Find(string ch1, string ch2)
        {
            string[] Alphabets = { "Z", "Y", "X", "W", "V", "U", "T", "S", "R" ,"Q" ,"P" ,"O" ,"N" ,"M" ,"L", "K" ,"J" , "I", "H", "G", "F","E","D", "C","B","A"};
            if(Array.IndexOf(Alphabets,ch1) < Array.IndexOf(Alphabets, ch2))
            {
                return true;
            }
            else if(Array.IndexOf(Alphabets, ch1) >= Array.IndexOf(Alphabets, ch2))
            {
                return false;
            }
            else
            {
                return true;
            }
        }

        protected void FillG(object sender, EventArgs e)
        {
            List<int> indexes = new List<int>();
            string json = File.ReadAllText(path);
            JArray jObj = (JArray)JsonConvert.DeserializeObject(json);
            for(int i = 0; i < jObj.Count; i++)
            {
                if (Find(Text3.Text, jObj[i]["grade"].ToString()))
                {
                    if(jObj[i]["grade"].ToString() != "HC")
                    indexes.Add(i);
                }
            }
            int j = 0;
            foreach (var item in indexes)
            {
                jObj.RemoveAt(item-j);
                j++;
            }

            json = jObj.ToString();
            System.IO.File.WriteAllText(path, json);

            switch (ViewState["st"])
            {
                case "MH":
                    ScriptManager.RegisterStartupScript(this, GetType(), "MH", "MH();", true);
                    break;
                case "MP":
                    ScriptManager.RegisterStartupScript(this, GetType(), "MP", "MP();", true);
                    break;
                case "CT":
                    ScriptManager.RegisterStartupScript(this, GetType(), "CT", "CT();", true);
                    break;
                case "TCGM":
                    ScriptManager.RegisterStartupScript(this, GetType(), "TC", "TC();", true);
                    break;
                case "ML":
                    ScriptManager.RegisterStartupScript(this, GetType(), "ML", "ML();", true);
                    break;
                case "MTL":
                    ScriptManager.RegisterStartupScript(this, GetType(), "MTL", "MTL();", true);
                    break;
                case "IT":
                    ScriptManager.RegisterStartupScript(this, GetType(), "IT", "IT();", true);
                    break;
                case "FZ":
                    ScriptManager.RegisterStartupScript(this, GetType(), "FZ", "FZ();", true);
                    break;
                case "AM":
                    ScriptManager.RegisterStartupScript(this, GetType(), "AM", "AM();", true);
                    break;
                case "MGP":
                    ScriptManager.RegisterStartupScript(this, GetType(), "MG", "MG();", true);
                    break;
                case "A":
                    ScriptManager.RegisterStartupScript(this, GetType(), "A", "A();", true);
                    break;
                default:
                    ScriptManager.RegisterStartupScript(this, GetType(), "G", "G();", true);
                    break;
            }
        }



        private void Json(string query)
        {
            string s = DataTableToJSONWithJSONNet(GetData(query));
            System.IO.File.WriteAllText(expath, s);
            string json = File.ReadAllText(expath);



            JArray jObj = (JArray)JsonConvert.DeserializeObject(json);



            for (int i = 0; i < jObj.Count; i++)
            {
                if (jObj[i]["MATRIC_SUPERIEUR"].ToString() == "")
                {
                }
                else
                {
                    for (int j = 0; j < jObj.Count; j++)
                    {
                        if (jObj[i]["MATRIC_SUPERIEUR"].ToString() == jObj[j]["MATRICULE"].ToString())
                        {
                            jObj[i]["NOM_SUP"] = jObj[j]["NOM"];
                            jObj[i]["PRENOM_SUP"] = jObj[j]["PRENOM"];
                        }
                    }
                }
            }



            for (int i = 0; i < jObj.Count; i++)
            {
                if (jObj[i]["CENTRE_GESTION"].ToString() == "")
                {
                }
                else
                {
                    for (int j = 0; j < jObj.Count; j++)
                    {
                        if (jObj[i]["CENTRE_GESTION"].ToString() == jObj[j]["MATRICULE"].ToString())
                        {
                            jObj[i]["NOM_GEST"] = jObj[j]["NOM"];
                            jObj[i]["PRENOM_GEST"] = jObj[j]["PRENOM"];
                        }
                    }
                }
            }



            var newJsonStr = jObj.ToString();
            File.WriteAllText(expath, newJsonStr);



        }

        protected void Btn_excel_Click(object sender, EventArgs e)
        {
            switch (ViewState["st"])
            {
                case "MH":
                    if(Text1.Text == "Global")
                    {
                        Json(excel + "'MENARA HOLDING'");
                        Excel(Json_table());
                    }
                    else
                    {
                        Json(excel + "'MENARA HOLDING'" + " and s.STE = a.STE and sr.LIB_COMPLET = '" + Text1.Text + "'");
                        Excel(Json_table());
                    }
                    
                    break;
                case "MP":
                    if (Text1.Text == "Global")
                    {
                        Json(excel + "'MENARA PREFA'");
                        Excel(Json_table());
                    }
                    else
                    {
                        Json(excel + "'MENARA PREFA'" + "' and sr.LIB_COMPLET = '" + Text1.Text + "'");
                        Excel(Json_table());
                    }
                    break;
                case "CT":
                    if (Text1.Text == "Global")
                    {
                        Json(excel + "'TRANSPORT'");
                        Excel(Json_table());
                    }
                    else
                    {
                        Json(excel + "'TRANSPORT'" + "' and sr.LIB_COMPLET = '" + Text1.Text + "'");
                        Excel(Json_table());
                    }
                    break;
                case "TCGM":
                    if (Text1.Text == "Global")
                    {
                        Json(excel + "'TRANSPORT'");
                        Excel(Json_table());
                    }
                    else
                    {
                        Json(excel + "'TRANSPORT'" + "' and sr.LIB_COMPLET = '" + Text1.Text + "'");
                        Excel(Json_table());
                    }
                    break;
                case "ML":
                    if (Text1.Text == "Global")
                    {
                        Json(excel + "'TRANSPORT'");
                        Excel(Json_table());
                    }
                    else
                    {
                        Json(excel + "'TRANSPORT'" + "' and sr.LIB_COMPLET = '" + Text1.Text + "'");
                        Excel(Json_table());
                    }
                    break;
                case "MTL":
                    if (Text1.Text == "Global")
                    {
                        Json(excel + "'TRANSPORT'");
                        Excel(Json_table());
                    }
                    else
                    {
                        Json(excel + "'TRANSPORT'" + "' and sr.LIB_COMPLET = '" + Text1.Text + "'");
                        Excel(Json_table());
                    }
                    break;
                case "IT":
                    if (Text1.Text == "Global")
                    {
                        Json(excel + "'TRANSPORT'");
                        Excel(Json_table());
                    }
                    else
                    {
                        Json(excel + "'TRANSPORT'" + "' and sr.LIB_COMPLET = '" + Text1.Text + "'");
                        Excel(Json_table());
                    }
                    break;
                case "FZ":
                    if (Text1.Text == "Global")
                    {
                        Json(excel + "'TRANSPORT'");
                        Excel(Json_table());
                    }
                    else
                    {
                        Json(excel + "'TRANSPORT'" + "' and sr.LIB_COMPLET = '" + Text1.Text + "'");
                        Excel(Json_table());
                    }
                    break;
                case "AM":
                    if (Text1.Text == "Global")
                    {
                        Json(excel + "'TRANSPORT'");
                        Excel(Json_table());
                    }
                    else
                    {
                        Json(excel + "'TRANSPORT'" + "' and sr.LIB_COMPLET = '" + Text1.Text + "'");
                        Excel(Json_table());
                    }
                    break;
                case "MGP":
                    if (Text1.Text == "Global")
                    {
                        Json(excel + "'TRANSPORT'");
                        Excel(Json_table());
                    }
                    else
                    {
                        Json(excel + "'TRANSPORT'" + "' and sr.LIB_COMPLET = '" + Text1.Text + "'");
                        Excel(Json_table());
                    }
                    break;
                case "A":
                    if (Text1.Text == "Global")
                    {
                        Json(excel + "'TRANSPORT'");
                        Excel(Json_table());
                    }
                    else
                    {
                        Json(excel + "'TRANSPORT'" + "' and sr.LIB_COMPLET = '" + Text1.Text + "'");
                        Excel(Json_table());
                    }
                    break;
                default:
                    Json("select a.MATRICULE, g.NOM ,g.PRENOM , a.MATRIC_SUPERIEUR ,t.INTITULE as FONCTION, c.GRADE ,CENTRE_GESTION from SOCIETE st, CLASSIFICATION c, AFFECTATION a,AGENT g ,TABLE_REFERENCE t where a.MATRICULE = g.MATRICULE and CODE = c.FONCTION and c.MATRICULE = g.MATRICULE and st.STE = a.STE");
                    Excel(Json_table());
                    break;
            }
        }

        
    }
}