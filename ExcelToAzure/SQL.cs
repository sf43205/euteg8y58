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
        static List<(string, string)> Users = new List<(string, string)>()
        {
            ("HDCCOLA", "hdyuxin16"),
            ("admin", "pass"),
            ("user", "0000")
        };


        private static SqlConnection Connection() => new SqlConnection(ConnectionString);
        public static bool Connect(string user, string pass)
        {
            if (!Users.Any(x => x.Item1 == user && x.Item2 == pass))
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

        internal static bool QuerryUpdate(string commandtext)
        {
            try
            {
                commandtext = commandtext.Trim();
                using (var connection = Connection())
                using (var command = new SqlCommand(commandtext, connection))
                {
                    connection.Open();
                    command.ExecuteNonQuery();
                    connection.Close();
                }
                return true;
            }
            catch (Exception e)
            {
                Console.WriteLine("Error executing {0}\nerror:{1}", commandtext, e.Message);
            }
            return false;
        }

        internal static bool UpdateProject(Project project)
        {
            bool success = false;
            string cmdtext = project.id == -1 ?
                "insert into project (name, description, owner, type, duration, gsf) output inserted.id values (@name, @description, @owner, @type, @duration, @gsf);"
                :
                "update project set name = @name, description = @description, owner = @owner, type = @type, duration = @duration, gsf = @gsf output inserted.id where id = @id;";

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
                            command.Parameters.AddWithValue("gsf", project.gsf);

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

        public static bool ImportNewData(List<Record> records)
        {
            bool success_all = false;
            string computer_that_imported = System.Windows.Forms.SystemInformation.ComputerName;
            string levelstxt = "BEGIN IF NOT EXISTS (select id from levels where UPPER(level1) = UPPER(@level1) and UPPER(name1) = UPPER(@name1) and UPPER(level2) = UPPER(@level2) and UPPER(name2) = UPPER(@name2) and UPPER(level3) = UPPER(@level3) and UPPER(name3) = UPPER(@name3) and UPPER(level4) = UPPER(@level4) and UPPER(name4) = UPPER(@name4)) " +
                               "BEGIN INSERT INTO levels(level1, name1, level2, name2, level3, name3, level4, name4) output inserted.id values(@level1, @name1, @level2, @name2, @level3, @name3, @level4, @name4) END " +
                               "ELSE SELECT id FROM levels WHERE UPPER(level1) = UPPER(@level1) and UPPER(name1) = UPPER(@name1) and UPPER(level2) = UPPER(@level2) and UPPER(name2) = UPPER(@name2) and UPPER(level3) = UPPER(@level3) and UPPER(name3) = UPPER(@name3) and UPPER(level4) = UPPER(@level4) and UPPER(name4) = UPPER(@name4) END ";
            string templatetxt = "BEGIN IF NOT EXISTS (select id from template where level_id = @level_id and UPPER(code) = UPPER(@code) and UPPER(ut) = UPPER(@ut) and UPPER(description) = UPPER(@description)) " +
                                 "BEGIN INSERT INTO template (level_id, code, description, ut) output inserted.id values (@level_id, @code, @description, @ut) END " +
                                 "ELSE SELECT id FROM template WHERE level_id = @level_id and UPPER(code) = UPPER(@code) and UPPER(ut) = UPPER(@ut) and UPPER(description) = UPPER(@description) END";
            string locationtxt = "BEGIN IF NOT EXISTS (select id from location where project_id = project_id and UPPER(code) = UPPER(@code) and UPPER(name) = UPPER(@name)) " +
                                 "BEGIN INSERT INTO location(project_id, code, name, bsf) output inserted.id values(@project_id, @code, @name, @bsf) END " +
                                 "ELSE select id from location where project_id = project_id and UPPER(code) = UPPER(@code) and UPPER(name) = UPPER(@name)) END ";
            string pricetxt = "BEGIN IF NOT EXISTS (select template_id from product_price where phase_id = @phase_id and template_id = @template_id and project_id = project_id and cast(unit_price as decimal(16,7)) = cast(@unit_price as decimal(16,7))) " + 
                              "BEGIN INSERT INTO product_price(template_id, phase_id, project_id, unit_price) values(@template_id, @phase_id, @project_id, @unit_price) END END";
            string recordtxt = "BEGIN IF NOT EXISTS (select id from record where template_id = @template_id and project_id = @project_id and location_id = @location_id and phase_id = @phase_id and cast(qty as decimal(16,7)) = cast(@qty as decimal(16,7)) and cast(total as decimal(16,7)) = cast(@total as decimal(16,7))) " +
                               "BEGIN INSERT INTO record (template_id, project_id, qty, total, comments, csi_code, trade_code, estimate_category, location_id, phase_id, computer_that_imported) output inserted.id values (@template_id, @project_id, @qty, @total, @comments, @csi_code, @trade_code, @estimate_category, @location_id, @phase_id, @computer_that_imported + ' INSERTED') END " +
                               "ELSE UPDATE record set comments = @comments, time_recorded = getdate(), csi_code = @csi_code, trade_code = @trade_code, estimate_category = @estimate_category, computer_that_imported = computer_that_imported + ' ' + @computer_that_imported + ' UPDATED' output inserted.id where id = (select top 1 id from record where template_id = @template_id and project_id = @project_id and location_id = @location_id and phase_id = @phase_id and cast(qty as decimal(16, 7)) = cast(@qty as decimal(16, 7)) and cast(total as decimal(16, 7)) = cast(@total as decimal(16, 7))) END";
            Form1.Bar.SafeInvoke(x =>
            {
                x.Maximum = records.Count();
                x.Value = 0;
                x.Visible = true;
            });
            using (var connection = Connection())
            {
                connection.Open();
                using (var transaction = connection.BeginTransaction())
                {
                    try
                    {
                        records.ForEach(record =>
                        {
                            using (var levelscmd = new SqlCommand(levelstxt, connection, transaction))
                            using (var templatecmd = new SqlCommand(templatetxt, connection, transaction))
                            using (var locationcmd = new SqlCommand(locationtxt, connection, transaction))
                            using (var pricecmd = new SqlCommand(pricetxt, connection, transaction))
                            using (var recordcmd = new SqlCommand(recordtxt, connection, transaction))
                            {
                                try
                                {
                                    levelscmd.Parameters.AddWithValue("level1", record.template.level.level1);
                                    levelscmd.Parameters.AddWithValue("name1", record.template.level.name1);
                                    levelscmd.Parameters.AddWithValue("level2", record.template.level.level1);
                                    levelscmd.Parameters.AddWithValue("name2", record.template.level.name1);
                                    levelscmd.Parameters.AddWithValue("level3", record.template.level.level1);
                                    levelscmd.Parameters.AddWithValue("name3", record.template.level.name1);
                                    levelscmd.Parameters.AddWithValue("level4", record.template.level.level1);
                                    levelscmd.Parameters.AddWithValue("name4", record.template.level.name1);

                                    record.template.level.id = (int)levelscmd.ExecuteScalar();
                                    Console.WriteLine("Success levelscmd");
                                    levelscmd.Dispose();

                                    templatecmd.Parameters.AddWithValue("level_id", record.template.level.id);
                                    templatecmd.Parameters.AddWithValue("code", record.template.code);
                                    templatecmd.Parameters.AddWithValue("description", record.template.description);
                                    templatecmd.Parameters.AddWithValue("ut", record.template.ut);

                                    record.template.id = (int)templatecmd.ExecuteScalar();
                                    Console.WriteLine("Success templatecmd");
                                    templatecmd.Dispose();

                                    locationcmd.Parameters.AddWithValue("code", record.location.code);
                                    locationcmd.Parameters.AddWithValue("name", record.location.name);
                                    locationcmd.Parameters.AddWithValue("project_id", record.location.project.id);
                                    locationcmd.Parameters.AddWithValue("bsf", record.location.bsf);

                                    record.location.id = (int)locationcmd.ExecuteScalar();
                                    Console.WriteLine("Success locationcmd");
                                    locationcmd.Dispose();

                                    pricecmd.Parameters.AddWithValue("project_id", record.location.project.id);
                                    pricecmd.Parameters.AddWithValue("template_id", record.template.id);
                                    pricecmd.Parameters.AddWithValue("phase_id", record.phase.id);
                                    pricecmd.Parameters.AddWithValue("unit_price", record.price);

                                    pricecmd.ExecuteNonQuery();
                                    Console.WriteLine("Success pricecmd");
                                    pricecmd.Dispose();

                                    recordcmd.Parameters.AddWithValue("template_id", record.template.id);
                                    recordcmd.Parameters.AddWithValue("project_id", record.location.project.id);
                                    recordcmd.Parameters.AddWithValue("phase_id", record.phase.id);
                                    recordcmd.Parameters.AddWithValue("location_id", record.location.id);
                                    recordcmd.Parameters.AddWithValue("qty", record.qty);
                                    recordcmd.Parameters.AddWithValue("total", record.total);
                                    recordcmd.Parameters.AddWithValue("comments", record.comments);
                                    recordcmd.Parameters.AddWithValue("csi_code", record.csi_code);
                                    recordcmd.Parameters.AddWithValue("trade_code", record.trade_code);
                                    recordcmd.Parameters.AddWithValue("estimate_category", record.estimate_category);
                                    recordcmd.Parameters.AddWithValue("computer_that_imported", computer_that_imported);

                                    record.id = (int)recordcmd.ExecuteScalar();
                                    Console.WriteLine("Success recordcmd");
                                    recordcmd.Dispose();

                                    success_all = record.id != -1;
                                    Form1.Bar.SafeInvoke(x => x.Value++);
                                }
                                catch (Exception ex)
                                {
                                    Console.WriteLine("Error executing {0}\nerror:{1}", JsonConvert.SerializeObject(record), ex.Message);
                                    //PrivateClasses.SafeInvoke(() => MessageBox.Show(ex.Message));
                                    success_all = false;
                                }
                            }
                        });
                        transaction.Commit();
                    }
                    catch (Exception e)
                    {
                        Console.WriteLine("Error executing {0}\nerror:{1}", "ImportNewData", e.Message);
                        PrivateClasses.SafeInvoke(() => MessageBox.Show(Form1.Bar.Value.ToString() + " were successfully imported", "Some records were not imported"));
                        transaction.Rollback();
                        connection.Close();
                        return false;
                    }
                }
                connection.Close();
            }
            Form1.Bar.SafeInvoke(x => x.Visible = false);
            return success_all;
        }
    }
}
