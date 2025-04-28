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
    public partial class Logs : Form
    {
        public Logs()
        {
            InitializeComponent();
            LoadLogsData();
        }

        private void LoadLogsData()
        {
            Workbook book = new Workbook();
            book.LoadFromFile(@"C:\Users\ACT-STUDENT\Desktop\Book.xlsx");
            Worksheet sheet = book.Worksheets[1];

            DataTable dt = new DataTable();

            dt.Columns.Add("Name");
            dt.Columns.Add("Saying");
            dt.Columns.Add("Date");
            dt.Columns.Add("Time");


            for (int rowIndex = 2; rowIndex <= sheet.LastRow; rowIndex++)
            {
                DataRow row = dt.NewRow();
                row["Name"] = sheet[rowIndex, 1].Value;
                row["Saying"] = sheet[rowIndex, 2].Value;
                row["Date"] = sheet[rowIndex, 3].Value;
                row["Time"] = sheet[rowIndex, 4].Value;
                dt.Rows.Add(row);
            }

            dgvLogs.DataSource = dt;
        }
    }
}

