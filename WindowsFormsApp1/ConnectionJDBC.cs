using MySql.Data.MySqlClient;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace WindowsFormsApp1
{
   
        public class ConnectionJDBC
        {
            public const string DBURL = "server=localhost;port=3306";
            public const string USER = "root";
            public const string PASSWORD = "";
            public const string DBNAME = "schimbval";

            public static MySqlConnection GetConnection()
            {
                MySqlConnection conn = new MySqlConnection(DBURL + ";database=" + DBNAME + ";uid=" + USER + ";password=" + PASSWORD);
                return conn;
            }

            public static void CloseConnection(MySqlConnection conn)
            {
                if (conn != null)
                {
                    conn.Close();
                }
            }

            public static void CloseAll(MySqlConnection conn, MySqlCommand cmd, MySqlDataReader reader)
            {
                if (reader != null)
                {
                    reader.Close();
                }
                if (cmd != null)
                {
                    cmd.Dispose();
                }
                if (conn != null)
                {
                    conn.Close();
                }
            }
        }
    }


