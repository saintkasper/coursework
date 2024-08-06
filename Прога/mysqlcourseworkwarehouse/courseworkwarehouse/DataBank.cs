using System;
using System.Collections.Generic;
using System.Data.SqlClient;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows;

namespace courseworkwarehouse
{
    public static class DataBank
    {
        public static int countGroups = 0;
        public static int cbProductsCount = 0;
        public static string user = "";
        public static string userFIO = "";
        public static string employeenum = "";
        public static Forms.FormAdmin formAdmin;
        public static Forms.FormManager formManager;
        public static string photopath = "";
        //public static SqlConnection sqlConnection = new SqlConnection(@"Data Source=DESKTOP-TL7NP8K\SQLEXPRESS;Initial Catalog=Office supplies warehouse;Integrated Security=True;");
    }
}
