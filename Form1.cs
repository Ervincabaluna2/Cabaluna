using Spire.Xls;
using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;

namespace Cabaluna
{
    
    public partial class FormBase : Form
    {

        Form2 f2 = new Form2();
        string[] student = new string[100];
        int i = 0;
        Mylogs ml = new Mylogs();
        
        public FormBase()
        {
            InitializeComponent();
        }

        public void btnAdd_Click(object sender, EventArgs e)
        {
            string data = "";           
            string data2 = "";
            string data3 = "";
            string data4 = "";
            string data5 = "";
            string data6 = "";
            string data7 = "";
            string data8 = "";
            string data9 = "";
            string data10 = "";
            string data11 = "";
            string data12 = "";

            data += txtName.Text + " ";
            
            if (radMale.Checked)
            {
                data2 += radMale.Text + " " ;
            }
            if (radFemale.Checked)
            {
                data2 += radFemale.Text + " ";
            }
            if (chkBasketball.Checked)
            { 
                data3 += chkBasketball.Text + " ";
            }
            if (chkVolleyball.Checked)
            {
                data3 += chkVolleyball.Text + " ";
            }
            if (chkSoccer.Checked)
            {
                data3 += chkSoccer.Text + " ";
            }
            data4 += cmbFaveColor.Text + " " ;
            data5 += txtSaying.Text + " ";
            data6 += txtAddress.Text + "";
            data7 += txtEmail.Text + "";
            data8 += txtAge.Text;
            data9 += txtUsername.Text + "";
            data10 += txtPassword.Text + "";
            data11 += dtpDate.Value.ToString("yyyy-MM-dd") + " ";
            data12 += cmbCourse.Text + "";


            student[i] =  data + data2 + data3 + data4 + data5 + data6 + data7 + data8 + data9 + data10 + data11 + data12;
            
            i++;
            txtName.Clear();
            radMale.Checked = false;
            radFemale.Checked = false;
            chkBasketball.Checked = false;
            chkVolleyball.Checked = false;
            chkSoccer.Checked = false;
            cmbFaveColor.Text = null;
            cmbCourse.Text = null;
            txtSaying.Clear();
            txtAddress.Clear();
            txtAge.Clear();
            txtEmail.Clear();
            txtPassword.Clear();
            txtUsername.Clear();
            dtpDate.Value = DateTime.Now;
            //f2.insert(i,data,data2,data3,data4,data5,data6,data7,data8,data9,data10,data11);

            Workbook book = new Workbook();
            book.LoadFromFile(@"C:\Users\ACT-STUDENT\Desktop\Book.xlsx");
            Worksheet sheet = book.Worksheets[0];
            int row = sheet.Rows.Length + 1;

            sheet.Range[row, 1].Value = row.ToString();
            sheet.Range[row, 2].Value = data;
            sheet.Range[row, 3].Value = data2;
            sheet.Range[row, 4].Value = data3;
            sheet.Range[row, 5].Value = data4;
            sheet.Range[row, 6].Value = data5;
            sheet.Range[row, 7].Value = data6;
            sheet.Range[row, 8].Value = data7;
            sheet.Range[row, 9].Value = data11;
            sheet.Range[row, 10].Value = data8;
            sheet.Range[row, 11].Value = data9;
            sheet.Range[row, 12].Value = data10;
            sheet.Range[row, 13].Value = data12;
            sheet.Range[row, 14].Value = "1";


            Mylogs ml = new Mylogs();
            ml.insertLogs(txtUsername.Text, "Added new account.");
            book.SaveToFile(@"C:\Users\ACT-STUDENT\Desktop\Book.xlsx", ExcelVersion.Version2016);
            DataTable datatable = sheet.ExportDataTable();
            f2.dgv.DataSource = datatable;
        }

        public void btnDisplay_Click(object sender, EventArgs e)
        {
            //string val = "";
            //for(int x = 0; x < student.Length; x++)
            // {
            //    val += "[" + x + "] = " + student[x] + " \n";                
            //}
            //MessageBox.Show(val);
            //
            f2.Show();
        }

        public void txtName_KeyPress(object sender, KeyPressEventArgs e)
        {
            if (!char.IsControl(e.KeyChar) && !char.IsLetter(e.KeyChar) && e.KeyChar != ' ')
            {
                e.Handled = true;
            }
        }

        private void btnUpdate_Click(object sender, EventArgs e)
        {
            string data = "";
            string data2 = "";
            string data3 = "";
            string data4 = "";
            string data5 = "";
            string data6 = "";
            string data7 = "";
            string data8 = "";
            string data9 = "";
            string data10 = "";
            string data11 = "";
            string data12 = "";

            data += txtName.Text + " ";

            if (radMale.Checked)
            {
                data2 += radMale.Text + " ";
            }
            if (radFemale.Checked)
            {
                data2 += radFemale.Text + " ";
            }
            if (chkBasketball.Checked)
            {
                data3 += chkBasketball.Text + " ";
            }
            if (chkVolleyball.Checked)
            {
                data3 += chkVolleyball.Text + " ";
            }
            if (chkSoccer.Checked)
            {
                data3 += chkSoccer.Text + " ";
            }
            data4 += cmbFaveColor.Text + " ";
            data5 += txtSaying.Text + " ";
            data6 += txtAddress.Text + "";
            data7 += txtEmail.Text + "";
            data8 += txtAge.Text;
            data9 += txtUsername.Text + "";
            data10 += txtPassword.Text + "";
            data11 += dtpDate.Value.ToString("yyyy-MM-dd") + " ";
            data12 += cmbCourse.Text + "";


            student[i] = data + data2 + data3 + data4 + data5 + data6 + data7 + data8 + data9 + data10 + data11 + data12;

            i++;
            txtName.Clear();
            radMale.Checked = false;
            radFemale.Checked = false;
            chkBasketball.Checked = false;
            chkVolleyball.Checked = false;
            chkSoccer.Checked = false;
            cmbFaveColor.Text = null;
            cmbCourse.Text = null;
            txtSaying.Clear();
            txtAddress.Clear();
            txtAge.Clear();
            txtEmail.Clear();
            txtPassword.Clear();
            txtUsername.Clear();
            dtpDate.Value = DateTime.Now;
            //f2.insert(i,data,data2,data3,data4,data5,data6,data7,data8,data9,data10,data11);          

            Workbook book = new Workbook();
            book.LoadFromFile(@"C:\Users\ACT-STUDENT\Desktop\Book.xlsx");
            Worksheet sheet = book.Worksheets[0];
            int row = Convert.ToInt32(lbliD.Text) + 2;

            sheet.Range[row, 1].Value = row.ToString();
            sheet.Range[row, 2].Value = data;
            sheet.Range[row, 3].Value = data2;
            sheet.Range[row, 4].Value = data3;
            sheet.Range[row, 5].Value = data4;
            sheet.Range[row, 6].Value = data5;
            sheet.Range[row, 7].Value = data6;
            sheet.Range[row, 8].Value = data7;
            sheet.Range[row, 9].Value = data11;
            sheet.Range[row, 10].Value = data8;
            sheet.Range[row, 11].Value = data9;
            sheet.Range[row, 12].Value = data10;
            sheet.Range[row, 13].Value = data12;
            sheet.Range[row, 14].Value = "1";

            Mylogs ml = new Mylogs();
            ml.insertLogs(txtUsername.Text, "Added new account.");

            book.SaveToFile(@"C:\Users\ACT-STUDENT\Desktop\Book.xlsx", ExcelVersion.Version2016);
            DataTable datatable = sheet.ExportDataTable();
            f2.dgv.DataSource = datatable;
        }

        private void dtpDate_ValueChanged(object sender, EventArgs e)
        {
            string[] date = dtpDate.Value.ToString().Split(' ');
            string[] d = date[0].Split('/');
            txtAge.Text = (2025 - Convert.ToInt32(d[2])).ToString();
        }

        public string checkEmpty()
        {
            string errors = "";

            foreach(Control c in Controls)
            {
                if(c is TextBox)
                {
                    if(c.Text == "")
                    {
                        errors += c.Name + "is empty";
                    }
                }
            }
            return errors;

        }
    }
}
