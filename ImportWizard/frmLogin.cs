using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Text;
using System.Windows.Forms;

using DevComponents.DotNetBar;
using Ferrero.BLL;
using Ray.Encrypt;

namespace Ferrero
{
    public partial class frmLogin : Office2007Form 
    {
        public frmLogin()
        {
            InitializeComponent();
        }

        public string UserName
        {
            get { return textBox1.Text; }
            set { textBox1.Text = value; }
        }

        public string Password
        {
            get { return textBox2.Text; }
            set { textBox2.Text = value; }
        }

        private void frmLogin_Load(object sender, EventArgs e)
        {
            //this.textBox1 .Text = encrypt.EncryptPassword("kingdee", 12);
            //2015-03-09 use administrator/kingdee login
            UserName = "administrator";
            Password = "kingdee";
            //this.Width = 380;
            //this.Height = 260;

        }

        private void button1_Click(object sender, EventArgs e)
        {
            AccountService bAccountService = new AccountService();
            PassService bPassService = new PassService();
            if (bAccountService.UserLogin("Account", UserName, bPassService.EncryptPassword(Password, 12)))
            ///if (1 == 1)
            {
                this.DialogResult = System.Windows.Forms.DialogResult.OK;
                this.Close();
            }
            else
            {
                this.DialogResult = System.Windows.Forms.DialogResult.None;
                MessageBox.Show("账号或密码错误，请重新输入!");  
            }
        }

        private void button2_Click(object sender, EventArgs e)
        {
            this.Close();
        }

    }
}
