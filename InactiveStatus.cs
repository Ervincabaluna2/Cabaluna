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
    public partial class InactiveStatus : Form
    {
        Workbook book = new Workbook();

        public InactiveStatus()
        {
            InitializeComponent();
            LoadInactiveData();
        }

        private void InactiveStatus_Load(object sender, EventArgs e)
        {

        }
        public void LoadInactiveData()
        {
            book.LoadFromFile(@"C:\Users\ACT-STUDENT\source\repos\Cabaluna\Book.xlsx");
            Worksheet sheet = book.Worksheets[0];
            DataTable dt = sheet.ExportDataTable();

            DataRow[] inactiveRows = dt.Select("Status = '0'");
            DataTable inactiveTable = inactiveRows.CopyToDataTable();
            dgvInactive.DataSource = inactiveTable;
        }

        public void AddInactiveRecord(DataGridViewRow row)
        {
            // Create a new DataRow and add it to the DataTable bound to your DataGridView
            DataTable inactiveTable = (DataTable)dgvInactive.DataSource; // Assuming dgvInactive is your DataGridView

            DataRow newRow = inactiveTable.NewRow();
            for (int i = 0; i < row.Cells.Count; i++)
            {
                newRow[i] = row.Cells[i].Value;
            }
            inactiveTable.Rows.Add(newRow);
        }

        private void btnDelete_Click(object sender, EventArgs e)
        {
            DialogResult Yes = MessageBox.Show("Are you sure you want to delete the selected info?", "Notice", MessageBoxButtons.YesNo);

            if (Yes == DialogResult.Yes)
            {
                Workbook book = new Workbook();
                book.LoadFromFile(@"C:\Users\ACT-STUDENT\source\repos\Cabaluna\Book.xlsx");
                Worksheet sheet = book.Worksheets[0];
                int row = dgvInactive.CurrentCell.RowIndex + 2;

                sheet.Range[row, 14].Value = "1";

                book.SaveToFile(@"C:\Users\ACT-STUDENT\source\repos\Cabaluna\Book.xlsx", ExcelVersion.Version2016);
            }
        }
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
