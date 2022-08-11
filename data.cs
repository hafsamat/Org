using System;
using System.Collections.Generic;
using System.Linq;
using System.Web;

namespace Org
{
    public class data
    {
        private string Matricule;
        private string M_sup;
        private string Nom;
        private string Prenom;
        private string centre;
        private string dir;
        private string fonction;
        private string grade;

        public string Matricule1 { get => Matricule; set => Matricule = value; }
        public string M_sup1 { get => M_sup; set => M_sup = value; }
        public string Nom1 { get => Nom; set => Nom = value; }
        public string Prenom1 { get => Prenom; set => Prenom = value; }
        public string Centre { get => centre; set => centre = value; }
        public string Fonction { get => fonction; set => fonction = value; }
        public string Grade { get => grade; set => grade = value; }
        public string Dir { get => dir; set => dir = value; }
    }
}