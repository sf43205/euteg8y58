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
        public static Form1 Main;
        public static Control Dash, Login, Data, Import;

        public Form1()
        {
            InitializeComponent();
            Main = this;
            Dash = Dashboard;
            CheckShow();
            Login = GetActiveControl(new LoginPage());
        }
        public void CheckShow()
        {
            bool visible = LoginPage.LoggedIn;
            ImportMenu.Visible = visible;
            DataMenu.Visible = visible;
        }
        Control GetActiveControl(Form f)
        {
            f.TopLevel = false;
            //f.Show();
            return f;
        }

        protected override void OnLoad(EventArgs e)
        {
            base.OnLoad(e);
            Navigate(Login);
        }
        private void btnLogin_Click(object sender, EventArgs e)
        {
            if (Login == null)
                Login = GetActiveControl(new LoginPage());
            Navigate(Login);
        }

        private void btnImport_Click(object sender, EventArgs e)
        {
            if (Import == null)
                Import = GetActiveControl(new ImportPage());
            Navigate(Import);
            
        }

        private void btnData_Click(object sender, EventArgs e)
        {
            Dashboard.BackColor = Color.White;
            //Xls.ShowDataInNewApp();
        }

        public static void Navigate (Control cx)
        {
            var c = LoginPage.LoggedIn ? cx : Login;
            Dash.SafeInvoke(x =>
            {
                //form.TopLevel = false;
                //form.AutoSize = true;
                x.Controls.Clear();
                x.Controls.Add(c);
                c.Dock = DockStyle.Fill;
                c.Show();
            });
        }
    }
}
