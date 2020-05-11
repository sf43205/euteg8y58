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
                SQL.Connect();

                btnLogin.Text = "LOGIN";
                btnLogin.BackColor = Color.Teal;
            }
        }
    }
}
