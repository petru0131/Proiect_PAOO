using System;
using System.Collections.Generic;
using System.Data.SqlClient;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using MySql.Data.MySqlClient;
using System.Diagnostics;
using System.Windows.Forms;

namespace WindowsFormsApp1
{
    public class Login 
    {
        public int Id { get; set; }
        public string User { get; set; }
        public string Pass { get; set; }
        public string Nume { get; set; }
        public string Rol { get; set; }
    }
    public class LoginLoad
    {
        public static Login Verif(string user, string pass)
        {
            Login utilizator = new Login();

            try
            {
                using (MySqlConnection conn = ConnectionJDBC.GetConnection())
                {
                    conn.Open();
                    
                    using (MySqlCommand cmd = new MySqlCommand("select * from util where user=@user and pass=@pass", conn))
                    {
                        cmd.Parameters.AddWithValue("@user", user);
                        cmd.Parameters.AddWithValue("@pass", pass);
                        using (MySqlDataReader rs = cmd.ExecuteReader())
                        {
                            if (rs.Read())
                            {
                                utilizator = new Login();
                                utilizator.Id = rs.GetInt32("id");
                                utilizator.Rol = rs.GetString("rol");
                            }
                        }
                    }
                }
            }
            catch (SqlException ex)
            {
                Console.WriteLine(ex.Message);
            }
            return utilizator;
        }
        public static int GetUserId(string username)
        {
            int userId = 0;
            try
            {
               
                MySqlConnection connection = ConnectionJDBC.GetConnection();
                connection.Open();
                string query = "SELECT id FROM util WHERE user = @user";
                MySqlCommand command = new MySqlCommand(query, connection);
                command.Parameters.AddWithValue("@user", username);
                MySqlDataReader reader = command.ExecuteReader();
                if (reader.HasRows)
                {
                    reader.Read();
                    userId = reader.GetInt32(0);
                }
                reader.Close();
            }
            catch (MySqlException ex)
            {
                MessageBox.Show("Error retrieving user ID: " + ex.Message);
            }
            return userId;
        }

    }
}
