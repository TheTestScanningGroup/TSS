using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;
using TestScaningSystem.BusinessLayer;

namespace TestScaningSystem.PresentationLayer
{
    public partial class Login : Form
    {
        public Login()
        {
            InitializeComponent();
        }

        private void btnCancel_Click(object sender, EventArgs e)
        {
            Environment.Exit(0);
        }

        private void btnLogin_Click(object sender, EventArgs e)
        {
            if (txtUsername.Text != null && txtPassword.Text != null)
            {
                string Username = txtUsername.Text;
                string Password = txtPassword.Text;
                DataHandeler DH = new DataHandeler();
                if (DH.Login(Username, Password))
                {
                    GenerateTests GT = new GenerateTests();
                    GT.Show();
                    this.Hide();
                }
            }
            else
            {
                MessageBox.Show("Please enter user details","Incorect User Detaills",MessageBoxButtons.RetryCancel,MessageBoxIcon.Error);
            }
        }
    }
}
