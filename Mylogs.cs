using Spire.Xls;
using System;
using System.Collections.Generic;
using System.Data;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;

namespace Cabaluna
{
    internal class Mylogs
    {
        Workbook book = new Workbook();
        
        public void insertLogs (string user, string message)
        {
            book.LoadFromFile(@"C:\Users\ACT-STUDENT\Desktop\Book.xlsx");
            Worksheet sheet = book.Worksheets[1];
            int row = sheet.Rows.Length + 1;
            sheet.Range[row,1].Value = user;
            sheet.Range[row, 2].Value = message;
            sheet.Range[row, 3].Value = DateTime.Now.ToString("MM/dd/yyyy");
            sheet.Range[row, 4].Value = DateTime.Now.ToString("hh:mm:ss");
            book.SaveToFile(@"C:\Users\ACT-STUDENT\Desktop\Book.xlsx");
        }

        public void showLogs(DataGridView dgv)
        {
            book.LoadFromFile(@"C:\Users\ACT-STUDENT\Desktop\Book.xlsx");
            Worksheet sheet = book.Worksheets[1];
            DataTable datatable = sheet.ExportDataTable();
            dgv.DataSource = datatable;
        }
    }
}
