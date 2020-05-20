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
        static string Database = "LAPRECON", username = "AAAzureAdminLAPRECON", password = "wer.asc%)$#B4weAbsd234:)";
        static string ConnectionString = "";
        static string DefaultUser = "HDCCOLA", DefaultPassword = "hdyuxin16";

        private static SqlConnection Connection() => new SqlConnection(ConnectionString);
        public static bool Connect(string user, string pass)
        {
            if (user != DefaultUser || pass != DefaultPassword)
            {
                return false;
            }
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
                commandtext += " FOR JSON PATH;";
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

        internal static List<Record> GetAllRecords()
        {
            var all = new List<Record>();
            string cmdtxt = "select (select r.*, " +
                            "price.unit_price as price, " +
                            "(select * from project_phase where id = r.phase_id for json auto) as phase, " +
                            "(select l.*, project.* from location l left join project on l.project_id = project.id where l.id = r.location_id for json auto) as location, " +
                            "(select t.*, level.* from template t left join levels level on t.level_id = level.id where t.id = r.template_id for json auto) as template " +
                            "from record r left join product_price price on (r.project_id = price.project_id and r.template_id = price.template_id and(r.phase_id is null or r.phase_id = price.phase_id)) " +
                            "where r.id = x.id for json path) from record x";
            QuerryRecords(cmdtxt).ForEach(row =>
            {
                try
                {
                    var result = JsonConvert.DeserializeObject<Record>(row);
                    all.Add(result);
                }
                catch (Exception e)
                {
                    Console.WriteLine("Failed to serialize record:\n{1}\ne:{0}", e.Message, row);
                }
            });
            return all;
        }

        private static List<string> QuerryRecords(string commandtext)
        {
            var count = QuerryGet("select count(id) as count from record ").ToInt();
            Form1.Bar.SafeInvoke(x =>
            {
                x.Maximum = count;
                x.Value = 0;
                x.Visible = true;
            });
            var all = new List<string>();
            try
            {
                using (var connection = Connection())
                using (var command = new SqlCommand(commandtext, connection))
                {
                    connection.Open();
                    using (SqlDataReader reader = command.ExecuteReader())
                    {
                        while (reader.Read())
                        {
                            all.Add((reader.GetString(0) ?? "{}").Replace("[", "").Replace("]", ""));
                            Form1.Bar.SafeInvoke(x => x.Value++);
                        }
                    }
                    connection.Close();
                }
            }
            catch (Exception e)
            {
                Console.WriteLine("Error executing {0}\nerror:{1}", commandtext, e.Message);
            }
            Form1.Bar.SafeInvoke(x => x.Visible = false);
            return all;
        }
    }
}
