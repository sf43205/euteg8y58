using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Security;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;

namespace ExcelToAzure
{
    public partial class ColumnSelection : Form
    {
        static ColumnSelection control;
        List<Header> Headers = new List<Header>();
        List<Header> PairedHeaders = new List<Header>();
        List<string[]> Rows = new List<string[]>();
        Project project;
        Phase phase;
        int page = 0;

        public ColumnSelection()
        {
            InitializeComponent();
            this.TopLevel = false;
            this.AutoScroll = true;
        }

        private void PhaseSelection_Load(object sender, EventArgs e)
        {
            control = this;
            control.TopLevel = false;
            AutoScroll = true;
        }

        public static void Open(List<string[]> rows, Project p, Phase ph)
        {
            if (!rows.Any())
            {
                MessageBox.Show("No data found");
                return;
            }
            if (control == null)
                control = new ColumnSelection();
            Form1.Navigate(control);
            control.SetData(rows, p, ph);
        }

        void SetData(List<string[]> rows, Project p, Phase ph)
        {
            page = 0;
            Headers.Clear();
            PairedHeaders = (new Header[properties.Length]).ToList();
            for (int i = 0; i < rows.First().Length; i++)
            {
                Headers.Add(new Header() { index = i, value = rows.First()[i] });
            }
            Headers = Headers.OrderBy(x => x.index).ToList();
            project = p.New();
            phase = ph.New();
            Rows = rows.ToList();
            Rows.RemoveAt(0);
            SetListBox();
        }

        void SetListBox()
        {
            ListBox.Items.Clear();
            Headers.FindAll(x => !PairedHeaders.Select(p => p.index).Contains(x.index)).ForEach(item => ListBox.Items.Add(item.index.ToString() + " " + item.value));
            title.Text = string.Format("LINK EXCEL COLUMNS TO DATA {0}/{1}", (page + 1).ToString(), properties.Length.ToString());
            txtDataProperty.Text = properties[page];
            txtColumnName.Text = PairedHeaders[page].value;
        }

        private void cancel_Click(object sender, EventArgs e)
        {
            if (page > 0)
            {
                page--;
                SetListBox();
            }
            else if(MessageBox.Show("Data will be lost", "Go back to Project selection?", MessageBoxButtons.YesNoCancel) == DialogResult.Yes)
                    Form1.Navigate(Form1.ImportPage);
        }

        private void create_Click(object sender, EventArgs e)
        {
            page++;
            if (page >= properties.Length)
            {
                Finish();
                page = properties.Length;
            }
            else
                SetListBox();
        }

        private void Finish()
        {
            if (PairedHeaders.Any(h => !(h.value ?? "").Any()))
            {
                MessageBox.Show("Go back and make sure all values are paired", "Some properties are missing a column data!");
                return;
            }
            var text = "These are the pairings (Excell --> DataBase):\n" + string.Join("\n", PairedHeaders.Select(x => string.Format("{0} --> {1}", x.value, properties[PairedHeaders.IndexOf(x)])));
            var result = MessageBox.Show(text, "Is everything correct?", MessageBoxButtons.YesNoCancel, MessageBoxIcon.Warning, MessageBoxDefaultButton.Button3);
            if (result == DialogResult.Yes)
            {
                InsertDataToDB(PrepData());
                Form1.Navigate(Form1.ImportPage);
                //Xls.ShowDataInNewApp(PrepData());
            }
        }

        private async void InsertDataToDB(List<Record> allrecs)
        {
            var success = false;
            while (Form1.Bar.Visible)
            {
                await Task.Delay(1000);
            }
            await Task.Run(() => success = SQL.ImportNewData(allrecs));
            if (success)
                MessageBox.Show("Successfully imported data");
            else
                MessageBox.Show("Failed to import all the data");
        }

        List<Record> PrepData()
        {
            List<Record> records = new List<Record>();
            for (int iRow = 0; iRow < Rows.Count(); iRow++)
            {
                records.Add(new Record());
                records[iRow].location.project = project.New();
                records[iRow].phase = phase.New();
                records[iRow].location.code = Rows[iRow][PairedHeaders[0].index];
                records[iRow].location.name = Rows[iRow][PairedHeaders[1].index];
                records[iRow].location.bsf = 0;
                records[iRow].template.level.level1 = Rows[iRow][PairedHeaders[2].index];
                records[iRow].template.level.name1 = Rows[iRow][PairedHeaders[3].index];
                records[iRow].template.level.level2 = Rows[iRow][PairedHeaders[4].index];
                records[iRow].template.level.name2 = Rows[iRow][PairedHeaders[5].index];
                records[iRow].template.level.level3 = Rows[iRow][PairedHeaders[6].index];
                records[iRow].template.level.name3 = Rows[iRow][PairedHeaders[7].index];
                records[iRow].template.level.level4 = Rows[iRow][PairedHeaders[8].index];
                records[iRow].template.level.name4 = Rows[iRow][PairedHeaders[9].index];
                records[iRow].template.code = Rows[iRow][PairedHeaders[10].index];
                records[iRow].template.description = Rows[iRow][PairedHeaders[11].index];
                records[iRow].qty = Rows[iRow][PairedHeaders[12].index].ToDecimal();
                records[iRow].template.ut = Rows[iRow][PairedHeaders[13].index];
                records[iRow].price = Rows[iRow][PairedHeaders[14].index].ToDecimal();
                records[iRow].total = Rows[iRow][PairedHeaders[15].index].ToDecimal();
                records[iRow].comments = Rows[iRow][PairedHeaders[16].index];
                records[iRow].csi_code = Rows[iRow][PairedHeaders[17].index];
                records[iRow].trade_code = Rows[iRow][PairedHeaders[18].index];
                records[iRow].estimate_category = Rows[iRow][PairedHeaders[19].index];
            }
            return records;
        }

        private void ListBox_SelectedIndexChanged(object sender, EventArgs e)
        {
            var Text = ListBox.SelectedItem as string;
            this.Flash(140, 4, txtColumnName, txtDataProperty);
            PairedHeaders[page] = Headers.Find(x => (x.index.ToString() + " " + x.value) == Text);
            SetListBox();
        }

        private void txtColumnName_Click(object sender, EventArgs e)
        {
            PairedHeaders[page] = new Header() { index = -1, value = "" };
            SetListBox();
        }

        struct Header
        {
            public string value;
            public int index;
        }

        string[] properties = new string[]
        {
            "location.code",
            "location.name",
            //"location.bsf",
            "level1",
            "name1",
            "level2",
            "name2",
            "level3",
            "name3",
            "level4",
            "name4",
            "template.code",
            "description",
            "qty",
            "ut",
            "price",
            "total",
            "comments",
            "csi_code",
            "trade_code",
            "estimate_category"
        };
    }
}
