using System;
using System.Collections.Generic;
using System.Linq;
using System.Web;
using System.Data.SqlClient;

namespace Org
{
    public class Connection
    {
        public static SqlConnection con = new SqlConnection(@"Data Source=192.168.1.19\AGIRH;Initial Catalog=agMenara;User ID=hmatich;Password=Hm@123");
    }
}