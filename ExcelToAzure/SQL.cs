using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Data.SqlClient;
using Newtonsoft.Json;
using System.Windows.Forms;

namespace ExcelToAzure
{

    public static class SQL
    {
        static string ServerName = "euteg8yt58.database.windows.net";
        static string Database = "LAPRECON";
        public static void Connect(string username = "AAAzureAdminLAPRECON", string password = "wer.asc%)$#B4weAbsd234:)")
        {
            try
            {
                SqlConnectionStringBuilder builder = new SqlConnectionStringBuilder();
                builder.DataSource = ServerName;
                builder.UserID = username;
                builder.Password = password;
                builder.InitialCatalog = Database;

                using (SqlConnection connection = new SqlConnection(builder.ConnectionString))
                {
                    Console.WriteLine("\nQuery data example:");
                    Console.WriteLine("=========================================\n");

                    StringBuilder sb = new StringBuilder();
                    sb.Append("SELECT TOP 7 id ");
                    sb.Append("FROM levels;");
                    String sql = sb.ToString();
                    string totalResult = "";
                    using (SqlCommand command = new SqlCommand(sql, connection))
                    {
                        connection.Open();
                        using (SqlDataReader reader = command.ExecuteReader())
                        {
                            while (reader.Read())
                            {
                                var result = string.Format("id {0}", reader.GetInt32(0).ToString());
                                Console.WriteLine("id {0}", reader.GetInt32(0).ToString());
                                totalResult += "\n" + result;
                            }
                        }
                    }
                    MessageBox.Show(totalResult);
                }
            }
            catch (Exception e)
            {
                Console.WriteLine(e.ToString());
            }
            Console.ReadLine();
        }
    }
}
