using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;

namespace ExcelToAzure
{
    public partial class LoginPage : Form
    {
        public static bool LoggedIn = false;
        public LoginPage()
        {
            InitializeComponent();
            Dock = DockStyle.Fill;
        }

        private void btnLogin_Click(object sender, EventArgs e)
        {
            if (btnLogin.Text == "LOGIN")
            {
                btnLogin.Text = "ATTEMPTING TO LOGIN...";
                btnLogin.BackColor = Color.PaleTurquoise;
                //await Task.Delay(2000);
                //MessageBox.Show("Failed to LOGIN");
                if (SQL.Connect())
                {
                    LoggedIn = true;
                    Form1.Main.CheckShow();
                    MessageBox.Show("Success!", "Login");
                }
                else
                {
                    LoggedIn = false;
                    Form1.Main.CheckShow();
                    MessageBox.Show("Failed!", "Login");
                }
                btnLogin.Text = "LOGIN";
                btnLogin.BackColor = Color.Teal;
            }
        }

        private void LoginPage_Resize(object sender, EventArgs e)
        {
            //Flow.Dock = DockStyle.Fill;
        }

    }
}
