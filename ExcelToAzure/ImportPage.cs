using Newtonsoft.Json;
using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Runtime.Remoting.Channels;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;

namespace ExcelToAzure
{
    public partial class ImportPage : Form
    {
        static List<Project> Data = new List<Project>();
        public ImportPage()
        {
            InitializeComponent();
            Dock = DockStyle.Fill;
            //Table.AutoScroll = true;
            AutoScroll = true;
            this.VisibleChanged += (s, e) => Reload();
        }

        protected override void OnShown(EventArgs e)
        {
            base.OnShown(e);
        }

        static bool fetching = false;
        private void Reload()
        {
            if (fetching) return;
            fetching = true;
            if (Row.DefaultRow == null)
                Row.DefaultRow = Table.Controls[0].Clone();
            Project.GetAll((res) =>
            {
                Data = res;
                Table.Controls.Clear();
                Table.RowStyles.Clear();
                Table.RowCount = Data.Count();
                Data.ForEach(x =>
                {
                    var row = new Row(x);
                    RowStyle style = new RowStyle()
                    {
                        SizeType = SizeType.Absolute,
                        Height = 48
                    };
                    Table.Controls.Add(row.control);
                    row.Clicked += (s, e) => PhaseSelection.Open(s as Project);
                    row.RightClicked += (s, e) => NewProject.Open(x, (saved) =>
                    {
                        if (saved)
                            Reload();
                    });
                    Table.RowStyles.Add(style);
                });
                Table.Controls.Add(new Panel());
                Table.Refresh();
                fetching = false;
            });
        }

        private class Row
        {
            public static Control DefaultRow;
            Label id, name, description, type;
            public Control control;
            private Project data;
            public EventHandler Clicked, RightClicked;

            public Row(Project p)
            {
                control = DefaultRow.Clone();
                control.Size = DefaultRow.Size;
                var all = control.All();
                id = all.FindLink(id, "id");
                name = all.FindLink(name, "name");
                description = all.FindLink(description, "description");
                type = all.FindLink(type, "type");
                project = p;
                //control.Show();
                all.ForEach(x => x.MouseClick += (s, e) =>
                {
                    if (e.Button == MouseButtons.Left)
                        Clicked?.Invoke(project, e);
                    else if (e.Button == MouseButtons.Right)
                        RightClicked?.Invoke(project, e);
                });
            }

            public Project project
            {
                get => data.New();
                set
                {
                    data = value.New();
                    id.Text = data.id.ToString();
                    name.Text = data.name;
                    description.Text = data.description;
                    type.Text = data.type;
                }
            }
                
        }

        private void CreateNewCompany_Click(object sender, EventArgs e)
        {
            NewProject.Open(new Project(), (saved) =>
            {
                if (saved)
                    Reload();
            });
        }
    }
}
