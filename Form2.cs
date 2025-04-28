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
    public partial class Form2 : Form
    {
        ActiveStatus active = new ActiveStatus();
        Mylogs ml = new Mylogs();
        Workbook book = new Workbook();
        public Form2()
        {
            InitializeComponent();
            // LoadExcelFile();
            //showStudent("1");
        }

        public void LoadExcelFile()
        {
            
            book.LoadFromFile(@"C:\Users\ACT-STUDENT\Desktop\Book.xlsx");
            Worksheet sheet = book.Worksheets[0];
            DataTable datatable = sheet.ExportDataTable();
            dgv.DataSource = datatable;
        }
       
        public void insert (int ID, string data, string data2,string data3, string data4, string data5, string data6, string data7, string data8, string data9, string data10, string data11)
        {
            
            int i = dgv.Rows.Add();
            var row = dgv.Rows[i];
            if (dgv.Columns.Count > 11)
            {
                dgv.Rows[i].Cells[0].Value = ID;
                dgv.Rows[i].Cells[1].Value = data;
                dgv.Rows[i].Cells[2].Value = data2;
                dgv.Rows[i].Cells[3].Value = data3;
                dgv.Rows[i].Cells[4].Value = data4;
                dgv.Rows[i].Cells[5].Value = data5;
                dgv.Rows[i].Cells[6].Value = data6;
                dgv.Rows[i].Cells[7].Value = data7;
                dgv.Rows[i].Cells[8].Value = data8;
                dgv.Rows[i].Cells[9].Value = data9;
                dgv.Rows[i].Cells[10].Value = data10;
                dgv.Rows[i].Cells[11].Value = data11;
            }
            
        }
        public void update(int ID, string data, string data2, string data3, string data4, string data5)
        {

         
            dgv.Rows[ID].Cells[1].Value = data;
            dgv.Rows[ID].Cells[2].Value = data2;
            dgv.Rows[ID].Cells[3].Value = data3;
            dgv.Rows[ID].Cells[4].Value = data4;
            dgv.Rows[ID].Cells[5].Value = data5;
        }

        public void btnDelete_Click(object sender, EventArgs e)
        {
            if (MessageBox.Show("Are you sure you want to delete this data?", "Delete.", MessageBoxButtons.YesNo, MessageBoxIcon.Question) == DialogResult.Yes)
            {
             foreach(DataGridViewRow row in this.dgv.SelectedRows)
                {
                    dgv.Rows.Remove(row);
                }   
            }
            else
            {
                return;
            }

        }

        public void btnClear_Click(object sender, EventArgs e)
        {
           if (MessageBox.Show("Are you sure you want to clear this data?" ,"Exit.", MessageBoxButtons.YesNo , MessageBoxIcon.Question) == DialogResult.Yes)
            {             
                    dgv.Rows.Clear();      
            }
            else
            {
                return;
            }
            
        }

        public void dgv_CellMouseDoubleClick(object sender, DataGridViewCellMouseEventArgs e)
        {
            //FormBase f1 = (FormBase)Application.OpenForms["FormBase"];
            FormBase f1 = new FormBase();
            int r = dgv.CurrentCell.RowIndex;
            dgv.Rows[r].Cells[0].Value.ToString();
            f1.lbliD.Text = r.ToString();

            f1.txtName.Text = dgv.Rows[r].Cells[1].Value.ToString();

            string gender = dgv.Rows[r].Cells[2].Value.ToString();
            if(gender.Trim() == "Male")
            {
                f1.radMale.Checked = true;
            }
            else
            {
                f1.radFemale.Checked = true;
            }
            string hobbies = dgv.Rows[r].Cells[3].Value.ToString();
            string[] h = hobbies.Split(' ');

            foreach (string ho in h)
            {

                if (ho.Trim().Equals ("Basketball"))
                {
                    f1.chkBasketball.Checked = true;
                }
                if (ho.Trim().Equals("Volleyball"))
                {
                    f1.chkVolleyball.Checked = true;
                }
                if (ho.Trim().Equals("Soccer"))          
                {
                    f1.chkSoccer.Checked = true;
                }
            }

            string FavoriteColor = dgv.Rows[r].Cells[4].Value.ToString();
            f1.cmbFaveColor.SelectedItem = FavoriteColor.Trim();
            f1.txtSaying.Text = dgv.Rows[r].Cells[5].Value.ToString();
            f1.txtAddress.Text = dgv.Rows[r].Cells[6].Value.ToString();
            f1.txtEmail.Text = dgv.Rows[r].Cells[7].Value.ToString();
            f1.txtAge.Text = dgv.Rows[r].Cells[9].Value.ToString();
            f1.txtUsername.Text = dgv.Rows[r].Cells[10].Value.ToString();
            f1.txtPassword.Text = dgv.Rows[r].Cells[11].Value.ToString();
            string dateValue = dgv.Rows[r].Cells[8].Value.ToString();
            DateTime parsedDate;

            if (DateTime.TryParse(dateValue, out parsedDate))
            {
                f1.dtpDate.Value = parsedDate;
            }
            else
            {
                MessageBox.Show("Date format is invalid. Please check the data.", "Error", MessageBoxButtons.OK, MessageBoxIcon.Error);
            }
            f1.lbliD.Text = r.ToString();

            f1.btnAdd.Visible = false;
            f1.btnUpdate.Visible = true;
            f1.Show();
            this.Hide();

        }

        private void btnSearch_Click(object sender, EventArgs e)
        {
           
            dgv.ClearSelection();
            

            foreach (DataGridViewRow r in dgv.Rows)
            {
                if (r.Cells[0].Value.ToString().Equals(txtSearch.Text))
                {
                    Mylogs ml = new Mylogs();

                    r.Selected = true;
                    break;
                }
            }
        }
       

        private void btnAdd_Click(object sender, EventArgs e)
        {
            FormBase fb = new FormBase();         
            fb.Show();
            this.Hide();
        }
        //public void showStudent(string status)
        //{
        //    book.LoadFromFile(@"C:\Users\ACT-STUDENT\Desktop\Book.xlsx");
        //    Worksheet sheet = book.Worksheets[0];
        //    DataTable dt = sheet.ExportDataTable();

        //    // Clear existing rows in the DataGridView
        //    //active.dgvActive.Rows.Clear();

        //    // Filter the DataTable based on the status
        //    DataRow[] filteredRows = dt.Select($"Status = '{status}'");

        //    // Insert filtered rows into the DataGridView
        //    foreach (DataRow r in filteredRows)
        //    {
        //        int rowIndex = active.dgvActive.Rows.Add();
        //        DataGridViewRow row = active.dgvActive.Rows[rowIndex];
        //        row.Cells[0].Value = r[0].ToString(); // ID
        //        row.Cells[1].Value = r[1].ToString(); // Name
        //        row.Cells[2].Value = r[2].ToString(); // Gender
        //        row.Cells[3].Value = r[3].ToString(); // Hobbies
        //        row.Cells[4].Value = r[4].ToString(); // Favorite Color
        //        row.Cells[5].Value = r[5].ToString(); // Saying
        //        row.Cells[6].Value = r[6].ToString(); // Address
        //        row.Cells[7].Value = r[7].ToString(); // Email
        //        row.Cells[8].Value = r[8].ToString(); // Date
        //        row.Cells[9].Value = r[9].ToString(); // Age
        //        row.Cells[10].Value = r[10].ToString(); // Username
        //        row.Cells[11].Value = r[11].ToString(); // Password

        //    }
        //}

        private void btnClose_Click(object sender, EventArgs e)
        {
            this.Hide();
        }

       
    }
}
