using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;

namespace WinformExcelSqliteExample
{
    public partial class Form1 : Form
    {
        public Form1()
        {
            InitializeComponent();
        }

        private void fillFromExcel_Click(object sender, EventArgs e)
        {
            var dataTable = ExcelHandler.ImportExceltoDatatable("data.xlsx");
            dataGridView1.DataSource = dataTable;
        }
    }
}
