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
            Workbook book = new Workbook();
            book.LoadFromFile(@"C:\Users\ACT-STUDENT\Desktop\Book.xlsx");
            Worksheet sheet = book.Worksheets[0];
            DataTable dt = sheet.ExportDataTable();

            DataRow[] inactiveRows = dt.Select("Status = '0'");
            DataTable inactiveTable = inactiveRows.CopyToDataTable();
            dgvInactive.DataSource = inactiveTable;
        }

       
    }


}
