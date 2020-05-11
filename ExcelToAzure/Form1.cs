using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;
using Microsoft.Office.Interop.Excel;
using Excel = Microsoft.Office.Interop.Excel;
using System.Reflection;
using System.Security;
using System.IO;

namespace ExcelToAzure
{
    public partial class Form1 : Form
    {
        Form Login, Data, Import;
        static Control Dash;

        private OpenFileDialog openFileDialog1;
        public Form1()
        {
            InitializeComponent();
            Dash = Dashboard;
            openFileDialog1 = new OpenFileDialog()
            {
                FileName = "Select a Excel file",
                Filter = "Text files (*.xls)|*.xls",
                Title = "Open Excel file"
            };

            Login = new LoginPage();
        }

        protected override void OnLoad(EventArgs e)
        {
            base.OnLoad(e);
            Navigate(Login);
        }
        private void btnLogin_Click(object sender, EventArgs e)
        {
            if (Login == null)
                Login = new LoginPage();
            Navigate(Login);
        }

        private void btnImport_Click(object sender, EventArgs e)
        {
            if (Import == null)
                Import = new ImportPage();
            Navigate(Import);
        }

        private void btnData_Click(object sender, EventArgs e)
        {
            Dashboard.BackColor = Color.White;
            Xls.ShowDataInNewApp();
        }

        private static void Navigate (Form form)
        {
            form.TopLevel = false;
            Dash.Controls.Clear();
            Dash.Controls.Add(form);
            form.Show();
            form.Dock = DockStyle.Fill;
        }
    }
}
