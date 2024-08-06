using MySql.Data.MySqlClient;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace courseworkwarehouse
{
    internal class MySql
    {
        public static MySqlConnection connection = new MySqlConnection(@"Server = localhost;Database= courseworkwarehouse;port= 3306;User Id= root;password=251436");

        public void openConnection()
        {
            if (connection.State == System.Data.ConnectionState.Open)
            {
                connection.Open();
            }
        }

        public void closeConnection()
        {
            if (connection.State == System.Data.ConnectionState.Closed)
            {
                connection.Close();

            }
        }

        public MySqlConnection GetMySqlConnection()
        {
            return connection;
        }
    }
}
