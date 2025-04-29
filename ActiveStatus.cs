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
    public partial class ActiveStatus : Form
    {
        //InactiveStatus inactive = new InactiveStatus();
        Workbook book = new Workbook();
        Dashboard db = new Dashboard();
        
        public ActiveStatus()
        {
            InitializeComponent();
            LoadActiveData();
        }

        
        public void LoadActiveData()
        {
            Workbook book = new Workbook();
            book.LoadFromFile(@"C:\Users\ACT-STUDENT\source\repos\Cabaluna\Book.xlsx");
            Worksheet sheet = book.Worksheets[0];
            DataTable dt = sheet.ExportDataTable();

            DataRow[] activeRows = dt.Select("Status = '1'");
            DataTable activeTable = activeRows.CopyToDataTable();
            dgvActive.DataSource = activeTable;
        }

        private void btnDelete_Click(object sender, EventArgs e)
        {
            DialogResult Yes = MessageBox.Show("Are you sure you want to delete the selected info?", "Notice", MessageBoxButtons.YesNo);

            if (Yes == DialogResult.Yes)
            {
                Workbook book = new Workbook();
                book.LoadFromFile(@"C:\Users\ACT-STUDENT\source\repos\Cabaluna\Book.xlsx");
                Worksheet sheet = book.Worksheets[0];
                int row = dgvActive.CurrentCell.RowIndex + 2;

                sheet.Range[row, 14].Value = "0";

                book.SaveToFile(@"C:\Users\ACT-STUDENT\source\repos\Cabaluna\Book.xlsx", ExcelVersion.Version2016);
            }
        }

        //private void UpdateStatusInExcel(int id, int newStatus)
        //{
        //    Workbook book = new Workbook();
        //    book.LoadFromFile(@"C:\Users\ACT-STUDENT\source\repos\Cabaluna\Book.xlsx");
        //    Worksheet sheet = book.Worksheets[0];

        //    // Find the row with the matching ID and update the status
        //    for (int i = 1; i <= sheet.LastRow; i++) // Assuming the first row is the header
        //    {
        //        if (Convert.ToInt32(sheet[i, 0].Value) == id) // Assuming the ID is in the first column
        //        {
        //            sheet[i, 1].Value = newStatus; // Assuming the status is in the second column
        //            break;
        //        }
        //    }

        //    // Save the changes back to the Excel file
        //    book.SaveToFile(@"C:\Users\ACT-STUDENT\source\repos\Cabaluna\Book.xlsx", ExcelVersion.Version2013);
        //}
        public int showCount(int c, string val)
        {

            book.LoadFromFile(@"C:\Users\ACT-STUDENT\source\repos\Cabaluna\Book.xlsx");
            Worksheet sheet = book.Worksheets[0];

            int row = sheet.Rows.Length;
            int counter = 0;

            for (int i = 2; i <= row; i++)
            {
                if (sheet.Range[i, c].Value.Trim() == val.Trim())
                {
                    counter++;
                }
            }
            return counter;
        }
    }
}
