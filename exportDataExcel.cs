using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;
using Microsoft.Office.Interop.Excel;

namespace exportExcel
{
    public partial class Form1 : Form
    {
        public Form1()
        {
            InitializeComponent();
        }

        private void Form1_Load(object sender, EventArgs e)
        {
            // TODO: данная строка кода позволяет загрузить данные в таблицу "анализПоведенияDataSet.Товар". При необходимости она может быть перемещена или удалена.
            this.товарTableAdapter.Fill(this.анализПоведенияDataSet.Товар);

        }

        private void button1_Click(object sender, EventArgs e)
        {
            ExportToExcel(товарDataGridView);
        }
	//Disable in datagridview: edit, delete, add!

        private void ExportToExcel(DataGridView grid)
        {
            if (grid.Rows.Count > 0)
            {
                Microsoft.Office.Interop.Excel.Application excel = new Microsoft.Office.Interop.Excel.Application();
                excel.Application.Workbooks.Add(Type.Missing);

                for (int i = 1; i < grid.Columns.Count + 1; i++)
                {
                    excel.Cells[1, i] = grid.Columns[i - 1].HeaderText;
                }

                for (int i = 0; i < grid.Rows.Count; i++)
                {
                    for (int j = 0; j < grid.Columns.Count; j++)
                    {
                        excel.Cells[i + 2, j + 1] = grid.Rows[i].Cells[j].Value.ToString();
                    }
                }

                excel.Columns.AutoFit();
                excel.Visible = true;
            }
            else
            {
                MessageBox.Show("No data to export!");
            }
        }
    }
}