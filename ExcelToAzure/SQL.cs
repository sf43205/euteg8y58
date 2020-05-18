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
        static string ConnectionString = "";

        private static SqlConnection Connection() => new SqlConnection(ConnectionString);
        public static bool Connect(string username = "AAAzureAdminLAPRECON", string password = "wer.asc%)$#B4weAbsd234:)")
        {
            try
            {
                var result = false;
                SqlConnectionStringBuilder builder = new SqlConnectionStringBuilder();
                builder.DataSource = ServerName;
                builder.UserID = username;
                builder.Password = password;
                builder.InitialCatalog = Database;
                ConnectionString = builder.ConnectionString;

                using (var connection = Connection())
                {
                    StringBuilder sb = new StringBuilder();
                    sb.Append("SELECT 'good';");
                    String sql = sb.ToString();
                    using (SqlCommand command = new SqlCommand(sql, connection))
                    {
                        connection.Open();
                        using (SqlDataReader reader = command.ExecuteReader())
                        {
                            while (reader.Read())
                            {
                                result = "good" == reader.GetString(0);
                            }
                        }
                    }
                }
                return result;
            }
            catch (Exception e)
            {
                Console.WriteLine(e.ToString());
                return false;
            }
        }

        internal static int InsertPhase(Phase phase)
        {
            string cmdtext = "BEGIN IF NOT EXISTS (SELECT phase FROM project_phase WHERE phase = @phase) BEGIN INSERT INTO project_phase (phase) output inserted.id values (@phase) END ELSE SELECT id FROM project_phase WHERE phase = @phase END";

            try
            {
                using (var connection = Connection())
                {
                    connection.Open();
                    using (var transaction = connection.BeginTransaction())
                    using (var command = new SqlCommand(cmdtext, connection, transaction))
                    {
                        try
                        {
                            command.Parameters.AddWithValue("phase", phase.phase);

                            phase.id = (int)command.ExecuteScalar();

                            transaction.Commit();
                        }
                        catch (Exception ex)
                        {
                            Console.WriteLine("Error executing {0}\nerror:{1}", cmdtext, ex.Message);
                            PrivateClasses.SafeInvoke(() => MessageBox.Show(ex.Message));
                            transaction.Rollback();
                        }
                    }
                    connection.Close();
                }
            }
            catch (Exception e)
            {
                Console.WriteLine("Error executing {0}\nerror:{1}", cmdtext, e.Message);
                PrivateClasses.SafeInvoke(() => MessageBox.Show(e.Message));
            }
            return phase.id;
        }

        internal static string QuerryGet(string commandtext)
        {
            string res = "[]";
            try
            {
                commandtext = commandtext.Trim(new char[] { ' ', ';' });
                commandtext += " FOR JSON AUTO;";
                using (var connection = Connection())
                using (var command = new SqlCommand(commandtext, connection))
                {
                    connection.Open();
                    using (SqlDataReader reader = command.ExecuteReader())
                    {
                        reader.Read();
                        res = reader.GetString(0) ?? "[]";
                    }
                    connection.Close();
                }
            }
            catch (Exception e)
            {
                Console.WriteLine("Error executing {0}\nerror:{1}", commandtext, e.Message);
            }
            return res;
        }

        internal static bool UpdateProject(Project project)
        {
            bool success = false;
            string cmdtext = project.id == -1 ?
                "insert into project (name, description, owner, type, duration) output inserted.id values (@name, @description, @owner, @type, @duration);"
                :
                "update project set name = @name, description = @description, owner = @owner, type = @type, duration = @duration output inserted.id where id = @id;";

            try
            {
                using (var connection = Connection())
                {
                    connection.Open();
                    using (var transaction = connection.BeginTransaction())
                    using (var command = new SqlCommand(cmdtext, connection, transaction))
                    {
                        try
                        {
                            command.Parameters.AddWithValue("id", project.id);
                            command.Parameters.AddWithValue("name", project.name);
                            command.Parameters.AddWithValue("description", project.description);
                            command.Parameters.AddWithValue("owner", project.owner);
                            command.Parameters.AddWithValue("type", project.type);
                            command.Parameters.AddWithValue("duration", project.duration);

                            project.id = (int)command.ExecuteScalar();

                            transaction.Commit();
                            success = project.id != -1;
                        }
                        catch (Exception ex)
                        {
                            Console.WriteLine("Error executing {0}\nerror:{1}", cmdtext, ex.Message);
                            PrivateClasses.SafeInvoke(() => MessageBox.Show(ex.Message));
                            transaction.Rollback();
                            success = false;
                        }
                    }
                    connection.Close();
                }
            }
            catch (Exception e)
            {
                Console.WriteLine("Error executing {0}\nerror:{1}", cmdtext, e.Message);
                PrivateClasses.SafeInvoke(() => MessageBox.Show(e.Message));
                success = false;
            }
            return success;
        }
    }
}
