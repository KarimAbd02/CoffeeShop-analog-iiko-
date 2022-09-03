using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;
using System.Data.OleDb;
using Excel = Microsoft.Office.Interop.Excel;
using System.IO;

namespace CO
{
    public partial class Form2 : Form
    {
        public Form2()
        {
            InitializeComponent();
        }

        private void button5_Click(object sender, EventArgs e)
        {
            this.Hide();
        }

        private void Form2_Load(object sender, EventArgs e)
        {
            
        }

        private void button7_Click(object sender, EventArgs e)
        {
            if (radioButton1.Checked)
            {
                string connectionString = "Provider=Microsoft.Jet.OLEDB.4.0;Data Source=|DataDirectory|\\Coffeeorange.mdb";
                OleDbConnection connection = new OleDbConnection(connectionString);
                connection.Open();
                OleDbDataAdapter adap;
                DataSet ds;
                adap = new OleDbDataAdapter("SELECT Напиток.Название,  Заказы.Количество, Заказы.Счет, Заказы.IDПользователя FROM Напиток INNER JOIN Заказы ON Напиток.IDНапитка = Заказы.IDНапитка WHERE Заказы.[Дата продажи] ='" + DateTime.Now.ToShortDateString() + "'", connectionString);
                ds = new System.Data.DataSet();
                adap.Fill(ds, "Products");
                dataGridView1.DataSource = ds.Tables[0];
                dataGridView1.Columns[0].HeaderText = "Название напитка";
                dataGridView1.Columns[1].HeaderText = "Количество";
                dataGridView1.Columns[2].HeaderText = "Счет";
                connection.Close();

                dataGridView1.ClearSelection();
                dataGridView1.Columns[0].Width = 195;

                Excel.Application ex;                       // объявляем переменную для приложения
                ex = new Excel.Application();               // связываем переменную с Excel
                ex.Visible = true;                          // отображаем на экране
                Excel.Workbook ex_book = ex.Workbooks.Add(); // создаем в приложении новую рабочую книгу



                ex.Worksheets[1].Range(ex.Worksheets[1].Cells[1, 1], ex.Worksheets[1].Cells[1, 6]).Merge();
                ex.Worksheets[1].Range(ex.Worksheets[1].Cells[2, 1], ex.Worksheets[1].Cells[2, 6]).Merge();
                ex.Worksheets[1].Cells[2, 1] = "Продажи за " + DateTime.Now.ToShortDateString();
                // заполняем автоматически таблицу в Excel данными из таблицы DataGridView1
                int i = 0;
                int j = 0;
                for (i = 0; i < dataGridView1.Rows.Count; i++)
                {
                    for (j = 0; j < dataGridView1.Columns.Count; j++)
                    {
                        if (i == 0)
                            ex.Worksheets[1].Cells[i + 3, j + 1] = dataGridView1.Columns[j].HeaderText.ToString();
                        else
                            ex.Worksheets[1].Cells[i + 3, j + 1].Value = dataGridView1.Rows[i - 1].Cells[j].Value.ToString();
                    }
                }
                int h = 4;
                int Sum = 0;
                while (h < (i + 3))
                {
                    Sum += Convert.ToInt32(ex.Worksheets[1].Cells[h, 3].Text);
                    h++;
                }
                ex.Worksheets[1].Cells[i + 4, 2] = "Итого ";
                ex.Worksheets[1].Cells[i + 4, 3] = Sum;

                Excel.Range xl_range = ex.Worksheets[1].Range(ex.Worksheets[1].Cells[2, 1],
                    ex.Worksheets[1].Cells[dataGridView1.Rows.Count + 1, dataGridView1.Columns.Count]); // выделение заполненной таблицы в Excel
                xl_range.EntireColumn.AutoFit();      // автоширина и автовысота
                xl_range.EntireRow.AutoFit();
            }
            
        }
    }
}

