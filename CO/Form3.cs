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

namespace CO
{
    public partial class Form3 : Form
    {
        public Form3()
        {
            InitializeComponent();
        }

        private void button11_Click(object sender, EventArgs e)
        {
            OleDbConnection con = new OleDbConnection("Provider=Microsoft.Jet.OLEDB.4.0;Data Source=|DataDirectory|\\Coffeeorange.mdb");
            con.Open();
            string queryString = "insert into  Заказы (IDПользователя, IDНапитка, [Дата продажи], Количество, Счет) values (" + Class1.ID + ",9,'" + DateTime.Now.ToShortDateString() + "',1,110)";
            OleDbCommand command = new OleDbCommand(queryString, con);
            command.ExecuteNonQuery();
            con.Close();
            MessageBox.Show("Добавлено");

            con.Open();
            string queryString1 = "SELECT [IDПродукта 1],[Количество 1],[IDПродукта 2],[Количество 2],[IDПродукта 3],[Количество 3] FROM Напиток WHERE IDНапитка=9";
            OleDbCommand command1 = new OleDbCommand(queryString1, con);
            OleDbDataReader dataReader1 = command1.ExecuteReader();
            string id1 = "", id2 = "", id3 = "";
            int k1 = 0, k2 = 0, k3 = 0;
            while (dataReader1.Read())
            {
                id1 = Convert.ToString(dataReader1[0]);
                k1 = Convert.ToInt32(dataReader1[1]);
                id2 = Convert.ToString(dataReader1[2]);
                k2 = Convert.ToInt32(dataReader1[3]);
                id3 = Convert.ToString(dataReader1[4]);
                k3 = Convert.ToInt32(dataReader1[5]);
            }
            con.Close();
            dataReader1.Close();
            Products(id1, k1);
            Products(id2, k2);
            Products(id3, k3);
        }

        private void button1_Click(object sender, EventArgs e)
        {
            OleDbConnection con = new OleDbConnection("Provider=Microsoft.Jet.OLEDB.4.0;Data Source=|DataDirectory|\\Coffeeorange.mdb");
            con.Open();
            string queryString = "insert into  Заказы (IDПользователя, IDНапитка, [Дата продажи], Количество, Счет) values (" + Class1.ID + ",2,'" 
                + DateTime.Now.ToShortDateString() + "',1,100)";
            OleDbCommand command = new OleDbCommand(queryString, con);
            command.ExecuteNonQuery();
            con.Close();
            MessageBox.Show("Добавлено");

            con.Open();
            string queryString1 = "SELECT [IDПродукта 1],[Количество 1],[IDПродукта 2],[Количество 2],[IDПродукта 3],[Количество 3] FROM Напиток WHERE IDНапитка=2";
            OleDbCommand command1 = new OleDbCommand(queryString1, con);
            OleDbDataReader dataReader1 = command1.ExecuteReader();
            string id1 = "", id2 = "", id3 = "";
            int k1 = 0, k2 = 0, k3 = 0;
            while (dataReader1.Read())
            {
                id1 = Convert.ToString(dataReader1[0]);
                k1 = Convert.ToInt32(dataReader1[1]);
                id2 = Convert.ToString(dataReader1[2]);
                k2 = Convert.ToInt32(dataReader1[3]);
                id3 = Convert.ToString(dataReader1[4]);
                k3 = Convert.ToInt32(dataReader1[5]);
            }
            con.Close();
            dataReader1.Close();
            Products(id1, k1);
            Products(id2, k2);
            Products(id3, k3);
        }

        private void Form3_Load(object sender, EventArgs e)
        {
            label2.Text = Class1.fio;
        }

        private void button2_Click(object sender, EventArgs e)
        {
            OleDbConnection con = new OleDbConnection("Provider=Microsoft.Jet.OLEDB.4.0;Data Source=|DataDirectory|\\Coffeeorange.mdb");
            con.Open();
            string queryString = "insert into  Заказы (IDПользователя, IDНапитка, [Дата продажи], Количество, Счет) values (" + Class1.ID + ",7,'" + DateTime.Now.ToShortDateString() + "',1,120)";
            OleDbCommand command = new OleDbCommand(queryString, con);
            command.ExecuteNonQuery();
            con.Close();
            MessageBox.Show("Добавлено");

            con.Open();
            string queryString1 = "SELECT [IDПродукта 1],[Количество 1],[IDПродукта 2],[Количество 2],[IDПродукта 3],[Количество 3] FROM Напиток WHERE IDНапитка=7";
            OleDbCommand command1 = new OleDbCommand(queryString1, con);
            OleDbDataReader dataReader1 = command1.ExecuteReader();
            string id1 = "", id2 = "", id3 = "";
            int k1 = 0, k2 = 0, k3 = 0;
            while (dataReader1.Read())
            {
                id1 = Convert.ToString(dataReader1[0]);
                k1 = Convert.ToInt32(dataReader1[1]);
                id2 = Convert.ToString(dataReader1[2]);
                k2 = Convert.ToInt32(dataReader1[3]);
                id3 = Convert.ToString(dataReader1[4]);
                k3 = Convert.ToInt32(dataReader1[5]);
            }
            con.Close();
            dataReader1.Close();
            Products(id1, k1);
            Products(id2, k2);
            Products(id3, k3);
        }

        private void button3_Click(object sender, EventArgs e)
        {
            OleDbConnection con = new OleDbConnection("Provider=Microsoft.Jet.OLEDB.4.0;Data Source=|DataDirectory|\\Coffeeorange.mdb");
            con.Open();
            string queryString = "insert into  Заказы (IDПользователя, IDНапитка, [Дата продажи], Количество, Счет) values (" + Class1.ID + ",10,'" + DateTime.Now.ToShortDateString() + "',1,140)";
            OleDbCommand command = new OleDbCommand(queryString, con);
            command.ExecuteNonQuery();
            con.Close();
            MessageBox.Show("Добавлено");

            con.Open();
            string queryString1 = "SELECT [IDПродукта 1],[Количество 1],[IDПродукта 2],[Количество 2],[IDПродукта 3],[Количество 3] FROM Напиток WHERE IDНапитка=10";
            OleDbCommand command1 = new OleDbCommand(queryString1, con);
            OleDbDataReader dataReader1 = command1.ExecuteReader();
            string id1 = "", id2 = "", id3 = "";
            int k1 = 0, k2 = 0, k3 = 0;
            while (dataReader1.Read())
            {
                id1 = Convert.ToString(dataReader1[0]);
                k1 = Convert.ToInt32(dataReader1[1]);
                id2 = Convert.ToString(dataReader1[2]);
                k2 = Convert.ToInt32(dataReader1[3]);
                id3 = Convert.ToString(dataReader1[4]);
                k3 = Convert.ToInt32(dataReader1[5]);
            }
            con.Close();
            dataReader1.Close();
            Products(id1, k1);
            Products(id2, k2);
            Products(id3, k3);
        }

        private void button13_Click(object sender, EventArgs e)
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
                ex.Worksheets[1].Cells[1, 1] = "Сотрудник - " + label2.Text;
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

        private void button5_Click(object sender, EventArgs e)
        {
            this.Hide();
            frmMain ss = new frmMain();
            ss.Show();
        }

        private void button4_Click(object sender, EventArgs e)
        {
            OleDbConnection con = new OleDbConnection("Provider=Microsoft.Jet.OLEDB.4.0;Data Source=|DataDirectory|\\Coffeeorange.mdb");
            con.Open();
            string queryString = "insert into  Заказы (IDПользователя, IDНапитка, [Дата продажи], Количество, Счет) values (" + Class1.ID + ",3,'" + DateTime.Now.ToShortDateString() + "',1,100)";
            OleDbCommand command = new OleDbCommand(queryString, con);
            command.ExecuteNonQuery();
            con.Close();
            MessageBox.Show("Добавлено");

            con.Open();
            string queryString1 = "SELECT [IDПродукта 1],[Количество 1],[IDПродукта 2],[Количество 2],[IDПродукта 3],[Количество 3] FROM Напиток WHERE IDНапитка=3";
            OleDbCommand command1 = new OleDbCommand(queryString1, con);
            OleDbDataReader dataReader1 = command1.ExecuteReader();
            string id1 = "", id2 = "", id3 = "";
            int k1 = 0, k2 = 0, k3 = 0;
            while (dataReader1.Read())
            {
                id1 = Convert.ToString(dataReader1[0]);
                k1 = Convert.ToInt32(dataReader1[1]);
                id2 = Convert.ToString(dataReader1[2]);
                k2 = Convert.ToInt32(dataReader1[3]);
                id3 = Convert.ToString(dataReader1[4]);
                k3 = Convert.ToInt32(dataReader1[5]);
            }
            con.Close();
            dataReader1.Close();
            Products(id1, k1);
            Products(id2, k2);
            Products(id3, k3);
        }

        private void button6_Click(object sender, EventArgs e)
        {
            OleDbConnection con = new OleDbConnection("Provider=Microsoft.Jet.OLEDB.4.0;Data Source=|DataDirectory|\\Coffeeorange.mdb");
            con.Open();
            string queryString = "insert into  Заказы (IDПользователя, IDНапитка, [Дата продажи], Количество, Счет) values (" + Class1.ID + ",8,'" + DateTime.Now.ToShortDateString() + "',1,120)";
            OleDbCommand command = new OleDbCommand(queryString, con);
            command.ExecuteNonQuery();
            con.Close();
            MessageBox.Show("Добавлено");

            con.Open();
            string queryString1 = "SELECT [IDПродукта 1],[Количество 1],[IDПродукта 2],[Количество 2],[IDПродукта 3],[Количество 3] FROM Напиток WHERE IDНапитка=8";
            OleDbCommand command1 = new OleDbCommand(queryString1, con);
            OleDbDataReader dataReader1 = command1.ExecuteReader();
            string id1 = "", id2 = "", id3 = "";
            int k1 = 0, k2 = 0, k3 = 0;
            while (dataReader1.Read())
            {
                id1 = Convert.ToString(dataReader1[0]);
                k1 = Convert.ToInt32(dataReader1[1]);
                id2 = Convert.ToString(dataReader1[2]);
                k2 = Convert.ToInt32(dataReader1[3]);
                id3 = Convert.ToString(dataReader1[4]);
                k3 = Convert.ToInt32(dataReader1[5]);
            }
            con.Close();
            dataReader1.Close();
            Products(id1, k1);
            Products(id2, k2);
            Products(id3, k3);
        }

        private void button7_Click(object sender, EventArgs e)
        {
            OleDbConnection con = new OleDbConnection("Provider=Microsoft.Jet.OLEDB.4.0;Data Source=|DataDirectory|\\Coffeeorange.mdb");
            con.Open();
            string queryString = "insert into  Заказы (IDПользователя, IDНапитка, [Дата продажи], Количество, Счет) values (" + Class1.ID + ",11,'" + DateTime.Now.ToShortDateString() + "',1,140)";
            OleDbCommand command = new OleDbCommand(queryString, con);
            command.ExecuteNonQuery();
            con.Close();
            MessageBox.Show("Добавлено");

            con.Open();
            string queryString1 = "SELECT [IDПродукта 1],[Количество 1],[IDПродукта 2],[Количество 2],[IDПродукта 3],[Количество 3] FROM Напиток WHERE IDНапитка=11";
            OleDbCommand command1 = new OleDbCommand(queryString1, con);
            OleDbDataReader dataReader1 = command1.ExecuteReader();
            string id1 = "", id2 = "", id3 = "";
            int k1 = 0, k2 = 0, k3 = 0;
            while (dataReader1.Read())
            {
                id1 = Convert.ToString(dataReader1[0]);
                k1 = Convert.ToInt32(dataReader1[1]);
                id2 = Convert.ToString(dataReader1[2]);
                k2 = Convert.ToInt32(dataReader1[3]);
                id3 = Convert.ToString(dataReader1[4]);
                k3 = Convert.ToInt32(dataReader1[5]);
            }
            con.Close();
            dataReader1.Close();
            Products(id1, k1);
            Products(id2, k2);
            Products(id3, k3);
        }

        private void button8_Click(object sender, EventArgs e)
        {
            OleDbConnection con = new OleDbConnection("Provider=Microsoft.Jet.OLEDB.4.0;Data Source=|DataDirectory|\\Coffeeorange.mdb");
            con.Open();
            string queryString = "insert into  Заказы (IDПользователя, IDНапитка, [Дата продажи], Количество, Счет) values (" + Class1.ID + ",4,'" + DateTime.Now.ToShortDateString() + "',1,80)";
            OleDbCommand command = new OleDbCommand(queryString, con);
            command.ExecuteNonQuery();
            con.Close();
            MessageBox.Show("Добавлено");

            con.Open();
            string queryString1 = "SELECT [IDПродукта 1],[Количество 1],[IDПродукта 2],[Количество 2],[IDПродукта 3],[Количество 3] FROM Напиток WHERE IDНапитка=4";
            OleDbCommand command1 = new OleDbCommand(queryString1, con);
            OleDbDataReader dataReader1 = command1.ExecuteReader();
            string id1 = "", id2 = "", id3 = "";
            int k1 = 0, k2 = 0, k3 = 0;
            while (dataReader1.Read())
            {
                id1 = Convert.ToString(dataReader1[0]);
                k1 = Convert.ToInt32(dataReader1[1]);
                id2 = Convert.ToString(dataReader1[2]);
                k2 = Convert.ToInt32(dataReader1[3]);
                id3 = Convert.ToString(dataReader1[4]);
                k3 = Convert.ToInt32(dataReader1[5]);
            }
            con.Close();
            dataReader1.Close();
            Products(id1, k1);
            Products(id2, k2);
            Products(id3, k3);
        }

        private void button9_Click(object sender, EventArgs e)
        {
            OleDbConnection con = new OleDbConnection("Provider=Microsoft.Jet.OLEDB.4.0;Data Source=|DataDirectory|\\Coffeeorange.mdb");
            con.Open();
            string queryString = "insert into  Заказы (IDПользователя, IDНапитка, [Дата продажи], Количество, Счет) values (" + Class1.ID + ",5,'" + DateTime.Now.ToShortDateString() + "',1,100)";
            OleDbCommand command = new OleDbCommand(queryString, con);
            command.ExecuteNonQuery();
            con.Close();
            MessageBox.Show("Добавлено");

            con.Open();
            string queryString1 = "SELECT [IDПродукта 1],[Количество 1],[IDПродукта 2],[Количество 2],[IDПродукта 3],[Количество 3] FROM Напиток WHERE IDНапитка=5";
            OleDbCommand command1 = new OleDbCommand(queryString1, con);
            OleDbDataReader dataReader1 = command1.ExecuteReader();
            string id1 = "", id2 = "", id3 = "";
            int k1 = 0, k2 = 0, k3 = 0;
            while (dataReader1.Read())
            {
                id1 = Convert.ToString(dataReader1[0]);
                k1 = Convert.ToInt32(dataReader1[1]);
                id2 = Convert.ToString(dataReader1[2]);
                k2 = Convert.ToInt32(dataReader1[3]);
                id3 = Convert.ToString(dataReader1[4]);
                k3 = Convert.ToInt32(dataReader1[5]);
            }
            con.Close();
            dataReader1.Close();
            Products(id1, k1);
            Products(id2, k2);
            Products(id3, k3);
        }

        private void button10_Click(object sender, EventArgs e)
        {
            OleDbConnection con = new OleDbConnection("Provider=Microsoft.Jet.OLEDB.4.0;Data Source=|DataDirectory|\\Coffeeorange.mdb");
            con.Open();
            string queryString = "insert into  Заказы (IDПользователя, IDНапитка, [Дата продажи], Количество, Счет) values (" + Class1.ID + ",6,'" + DateTime.Now.ToShortDateString() + "',1,100)";
            OleDbCommand command = new OleDbCommand(queryString, con);
            command.ExecuteNonQuery();
            con.Close();
            MessageBox.Show("Добавлено");

            con.Open();
            string queryString1 = "SELECT [IDПродукта 1],[Количество 1],[IDПродукта 2],[Количество 2],[IDПродукта 3],[Количество 3] FROM Напиток WHERE IDНапитка=6";
            OleDbCommand command1 = new OleDbCommand(queryString1, con);
            OleDbDataReader dataReader1 = command1.ExecuteReader();
            string id1 = "", id2 = "", id3 = "";
            int k1 = 0, k2 = 0, k3 = 0;
            while (dataReader1.Read())
            {
                id1 = Convert.ToString(dataReader1[0]);
                k1 = Convert.ToInt32(dataReader1[1]);
                id2 = Convert.ToString(dataReader1[2]);
                k2 = Convert.ToInt32(dataReader1[3]);
                id3 = Convert.ToString(dataReader1[4]);
                k3 = Convert.ToInt32(dataReader1[5]);
            }
            con.Close();
            dataReader1.Close();
            Products(id1, k1);
            Products(id2, k2);
            Products(id3, k3);
        }

        private void button12_Click(object sender, EventArgs e)
        {
            OleDbConnection con = new OleDbConnection("Provider=Microsoft.Jet.OLEDB.4.0;Data Source=|DataDirectory|\\Coffeeorange.mdb");
            con.Open();
            string queryString = "insert into  Заказы (IDПользователя, IDНапитка, [Дата продажи], Количество, Счет) values (" + Class1.ID + ",12,'" + DateTime.Now.ToShortDateString() + "',1,120)";
            OleDbCommand command = new OleDbCommand(queryString, con);
            command.ExecuteNonQuery();
            con.Close();
            MessageBox.Show("Добавлено");

            con.Open();
            string queryString1 = "SELECT " +
                "[IDПродукта 1],[Количество 1],[IDПродукта 2],[Количество 2],[IDПродукта 3],[Количество 3] FROM Напиток WHERE IDНапитка=12";
            OleDbCommand command1 = new OleDbCommand(queryString1, con);
            OleDbDataReader dataReader1 = command1.ExecuteReader();
            string id1 = "", id2 = "", id3 = "";
            int k1 = 0, k2 = 0, k3 = 0;
            while (dataReader1.Read())
            {
                id1 = Convert.ToString(dataReader1[0]);
                k1 = Convert.ToInt32(dataReader1[1]);
                id2 = Convert.ToString(dataReader1[2]);
                k2 = Convert.ToInt32(dataReader1[3]);
                id3 = Convert.ToString(dataReader1[4]);
                k3 = Convert.ToInt32(dataReader1[5]);
            }
            con.Close();
            dataReader1.Close();
            Products(id1, k1);
            Products(id2, k2);
            Products(id3, k3);
        }

        public void Products(string id1, int k1)
        {
            OleDbConnection con = new OleDbConnection("Provider=Microsoft.Jet.OLEDB.4.0;Data Source=|DataDirectory|\\Coffeeorange.mdb");
            con.Open();
            string queryString1 = "SELECT Количество FROM Продукты WHERE IDПродукта=" + id1;
            OleDbCommand command1 = new OleDbCommand(queryString1, con);
            OleDbDataReader dataReader1 = command1.ExecuteReader();
            int Sumk = 0;
            while (dataReader1.Read())
            {
                Sumk = Convert.ToInt32(dataReader1[0]);
            }
            con.Close();
            dataReader1.Close();

            int K = Sumk - k1;
            con.Open();
            string queryString = "UPDATE Продукты SET [Количество] =" + K + " WHERE IDПродукта =" + id1;
            OleDbCommand command = new OleDbCommand(queryString, con);
            command.ExecuteNonQuery();
            con.Close();
            // MessageBox.Show("Успешно изменен!");



        }
    }
}
