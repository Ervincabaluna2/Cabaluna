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

namespace Cabaluna
{
    public partial class Dashboard : Form
    {
        Logs logs = new Logs();   
        Mylogs ml = new Mylogs();
        Workbook book = new Workbook();
        Form2 f2 = new Form2();
        FormBase f1 = new FormBase();


        public Dashboard()
        {
            InitializeComponent();
            lblCountBasketball.Text = showCount(4, "Basketball").ToString();
            lblCountVolley.Text = showCount(4, "Volleyball").ToString();
            lblCountSoccer.Text = showCount(4, "Soccer").ToString();
            lblCountBSIT.Text = showCount(13,"BSIT").ToString();
            lblCountBSTM.Text = showCount(13, "BSTM").ToString();
            lblCountBSHM.Text = showCount(13, "BSHM").ToString();
            lblCountMale.Text = showCount(3, "Male").ToString();
            lblCountFemale.Text = showCount(3, "Female").ToString();
            lblCountWhite.Text = showCount(5,"White").ToString();
            lblCountBlack.Text = showCount(5,"Black").ToString();
            lblCountActive.Text = showCount(14, "1").ToString();
            lblCountInactive.Text = showCount(14, "0").ToString();

        }

        //private void btnActive_Click(object sender, EventArgs e)
        //{
        //    Worksheet sheet = book.Worksheets[0];
        //    DataTable datatable = sheet.ExportDataTable();
        //    active.showStudent("1");
        //    active.Show();
        //}

        private void btnLogout_Click(object sender, EventArgs e)
        {
            MessageBox.Show("Are you sure you want to logout?" , "Logout." , MessageBoxButtons.OK , MessageBoxIcon.Information);
            Login login = new Login();
            login.Show();
            this.Hide();
        }

        public int showCount(int c,string val)
        {

            book.LoadFromFile(@"C:\Users\ACT-STUDENT\Desktop\Book.xlsx");
            Worksheet sheet = book.Worksheets[0];

            int row = sheet.Rows.Length;
            int counter = 0;

            for(int i = 2; i<= row; i++)
            {
                if (sheet.Range[i,c].Value.Trim() == val.Trim())
                {
                    counter++;
                }
            }
            return counter;
        }

        private void btnInactive_Click(object sender, EventArgs e)
        {
            InactiveStatus inactive = new InactiveStatus();
            inactive.Show();
        }

        private void btnActive_Click(object sender, EventArgs e)
        {
            ActiveStatus active = new ActiveStatus();
            active.Show();
        }

        private void button1_Click(object sender, EventArgs e)
        {
            f1.Show();
        }

        private void btnLogs_Click(object sender, EventArgs e)
        {
            logs.Show();
        }
    }
}
