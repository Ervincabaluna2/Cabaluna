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
        Workbook book = new Workbook();
        
        public ActiveStatus()
        {
            InitializeComponent();
            LoadActiveData();
        }

        
        public void LoadActiveData()
        {
            Workbook book = new Workbook();
            book.LoadFromFile(@"C:\Users\ACT-STUDENT\Desktop\Book.xlsx");
            Worksheet sheet = book.Worksheets[0];
            DataTable dt = sheet.ExportDataTable();

            DataRow[] activeRows = dt.Select("Status = '1'");
            DataTable activeTable = activeRows.CopyToDataTable();
            dgvActive.DataSource = activeTable;
        }
    }
}
