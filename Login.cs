using Spire.Xls;
using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Diagnostics.Eventing.Reader;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;
using System.Xml.Linq;
using static System.Windows.Forms.VisualStyles.VisualStyleElement;

namespace Cabaluna
{
    public partial class Login : Form
    {
        public Login()
        {
            
            InitializeComponent();
        }
        FormBase f1 = new FormBase();
        Form2 f2 = new Form2();
        Dashboard db = new Dashboard();
        private void btnLogin_Click(object sender, EventArgs e)
        {

            Workbook book = new Workbook();
            book.LoadFromFile(@"C:\Users\ACT-STUDENT\Desktop\Book.xlsx");
            Worksheet sheet = book.Worksheets[0] ;
            //Worksheet sheet1 = book.Worksheets[1];

            int row = sheet.Rows.Length;
            //int rows = sheet1.Rows.Length + 1;
            bool log = false;


            for (int i = 2; i <= row; i++)
            {

                if (sheet.Range[i,11].Value == txtUsername.Text && sheet.Range[i,12].Value == txtPassword.Text)
                {
                    MessageBox.Show("Successfully logged in.", "Logged in.", MessageBoxButtons.OK, MessageBoxIcon.Information);
                    Mylogs ml = new Mylogs();
                    ml.insertLogs(txtUsername.Text, "Logged in.");

                    db.lblName.Text = sheet.Range[i, 2].Value;
                    db.lblDate.Text = sheet.Range[i,15].Value;
                    log = true;
                    break;
                }
               else
               {
                    log = false;
               }
              
            }
            if (log == true)
            {

                db.Show();
                this.Hide();
            }
            else
            {
                MessageBox.Show("No account found.", "Error.", MessageBoxButtons.OK, MessageBoxIcon.Information);
            }         
        }

        private void chkShowPassword_CheckedChanged(object sender, EventArgs e)
        {

           if (!chkShowPassword.Checked == true)
           {
               txtPassword.UseSystemPasswordChar = false;
           }
            else
            {
              txtPassword.UseSystemPasswordChar = true;
            }
        }

        private void btnSignUp_Click(object sender, EventArgs e)
        {
            if (txtUsername.Text == "" && txtPassword.Text == "")
            {
                MessageBox.Show("Please input fields.", "Invalid.", MessageBoxButtons.OK, MessageBoxIcon.Error);
            }

            else if(txtUsername.Text != "" && txtPassword.Text == "")
            {
                MessageBox.Show("Please input password.", "Invalid.", MessageBoxButtons.OK, MessageBoxIcon.Error);
            }
            else if (txtUsername.Text == "" && txtPassword.Text != "")
            {
                MessageBox.Show("Please input username.", "Invalid.", MessageBoxButtons.OK, MessageBoxIcon.Error);
            }
            else
            {
                MessageBox.Show("Account have been successfully created." , "Success." , MessageBoxButtons.OK , MessageBoxIcon.Information);
                f1.txtUsername.Text = txtUsername.Text;
                f1.txtPassword.Text = txtPassword.Text;
                f1.Show();
                this.Hide();
            }
        }
    }
}
