using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Text;
using System.Windows.Forms;

namespace myStatisticalAnalysis
{
    public partial class Login : Form
    {
        public Login()
        {
            InitializeComponent();
            tbPassword.UseSystemPasswordChar = true;
        }

        private void btLogin_Click(object sender, EventArgs e)
        {
            if (tbPassword.Text == "infac")
            {
                DialogResult = DialogResult.OK;
                this.Close();

            }
            else
            {
                MessageBox.Show("密码输入错误！");
            }
        }

        private void btCancel_Click(object sender, EventArgs e)
        {

            DialogResult = DialogResult.Cancel;
            this.Close();
        }

        private void LoginShown(object sender, EventArgs e)
        {
            //  if (this.Tag != null)
            //  {
            //if (this.Tag != null&&0 == (int)this.Tag)
            /// tbPassword.Text = "";
            //   }
            tbPassword.Text = "";
        }

        private void logcheck(object sender, EventArgs e)
        {
            btLogin_Click( sender,  e);
        }
    }
}