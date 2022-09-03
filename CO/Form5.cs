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
using System.IO;

namespace CO
{
    public partial class Form5 : Form
    {
        public Form5()
        {
            InitializeComponent();
        }
        public static string BD = "Provider=Microsoft.Jet.OLEDB.4.0;Data Source=" + Path.Combine(Application.StartupPath, "Coffeeorange.mdb;");
        private OleDbConnection con;

        private void Form5_Load(object sender, EventArgs e)
        {
            int ScreenWidth = Screen.PrimaryScreen.Bounds.Width;
            int ScreenHeight = Screen.PrimaryScreen.Bounds.Height;
            this.Location = new Point((ScreenWidth / 2) - (this.Width / 2),
                (ScreenHeight / 2) - (this.Height / 2));
            this.ControlBox = false;


            string put = "SELECT * FROM Продукты";
            OleDbDataAdapter dataAdapter = new OleDbDataAdapter(put, BD);
            // создаем объект DataSet
            DataSet ds = new DataSet();
            // заполняем таблицу Order  
            // данными из базы данных
            dataAdapter.Fill(ds, "Продукты");
            dataGridView1.DataSource = ds.Tables[0].DefaultView;
        }

        private void button5_Click(object sender, EventArgs e)
        {
            this.Hide();
            frmMain ss = new frmMain();
            ss.Show();
        }

        private void button8_Click(object sender, EventArgs e)
        {
            try
            {
                con = new OleDbConnection(BD);
                con.Open();
                string queryString = "UPDATE Продукты SET [Количество] ='" + textBox3.Text + "' WHERE IDПродукта =" + comboBox1.Text;
                OleDbCommand command = new OleDbCommand(queryString, con);
                command.ExecuteNonQuery();
                con.Close();
                MessageBox.Show("Успешно изменен!");

                con.Open();
                string put = "SELECT * FROM Продукты";
                OleDbDataAdapter dataAdapter = new OleDbDataAdapter(put, BD);
                DataSet ds = new DataSet();
                dataAdapter.Fill(ds, "Продукты");
                dataGridView1.DataSource = ds.Tables[0].DefaultView;
                this.Refresh();

            textBox3.Text = "";
            textBox3.Clear();
                comboBox1.SelectedIndex = -1;
            }
            catch
            {
                MessageBox.Show("Выберите корректный IDПродукта, либо введите количество продукта!");
            }

        }

        private void button3_Click(object sender, EventArgs e)
        {
            try
            {
                con = new OleDbConnection(BD);
            con.Open();
            string queryString = "UPDATE Продукты SET [Количество] ='" + textBox2.Text + "' WHERE IDПродукта =" + comboBox1.Text;
            OleDbCommand command = new OleDbCommand(queryString, con);
            command.ExecuteNonQuery();
            con.Close();
            MessageBox.Show("Успешно изменен!");

            con.Open();
            string put = "SELECT * FROM Продукты";
            OleDbDataAdapter dataAdapter = new OleDbDataAdapter(put, BD);
            DataSet ds = new DataSet();
            dataAdapter.Fill(ds, "Продукты");
            dataGridView1.DataSource = ds.Tables[0].DefaultView;
            this.Refresh();
            textBox2.Text = "";
            textBox2.Clear();
                comboBox1.SelectedIndex = -1;
            }
                catch
                {
                    MessageBox.Show("Выберите корректныйIDПродукта, либо введите количество продукта!");
                }
            }

        private void button6_Click(object sender, EventArgs e)
        {
            try 
            { 
            con = new OleDbConnection(BD);
            con.Open();
            string queryString = "UPDATE Продукты SET [Количество] ='" + textBox1.Text + "' WHERE IDПродукта =" + comboBox1.Text;
            OleDbCommand command = new OleDbCommand(queryString, con);
            command.ExecuteNonQuery();
            con.Close();
            MessageBox.Show("Успешно изменен!");

            con.Open();
            string put = "SELECT * FROM Продукты";
            OleDbDataAdapter dataAdapter = new OleDbDataAdapter(put, BD);
            DataSet ds = new DataSet();
            dataAdapter.Fill(ds, "Продукты");
            dataGridView1.DataSource = ds.Tables[0].DefaultView;
            this.Refresh();
            textBox1.Text = "";
            textBox1.Clear();
                comboBox1.SelectedIndex = -1;
            }
            catch
            {
                MessageBox.Show("Выберите корректный IDПродукта, либо введите количество продукта!");
            }

        }

        private void button9_Click(object sender, EventArgs e)
        {
            textBox1.Text = "";
            textBox1.Clear();
            textBox2.Text = "";
            textBox2.Clear();
            textBox3.Text = "";
            textBox3.Clear();
            comboBox1.SelectedIndex = -1;
        }

        private void textBox3_KeyPress(object sender, KeyPressEventArgs e)
        {
            char number = e.KeyChar;

            if (!Char.IsDigit(number) && number != 8) // цифры и клавиша BackSpace
            {
                e.Handled = true;
            }
        }

        private void textBox2_KeyPress(object sender, KeyPressEventArgs e)
        {
            char number = e.KeyChar;

            if (!Char.IsDigit(number) && number != 8) // цифры и клавиша BackSpace
            {
                e.Handled = true;
            }
        }

        private void textBox1_KeyPress(object sender, KeyPressEventArgs e)
        {
            char number = e.KeyChar;

            if (!Char.IsDigit(number) && number != 8) // цифры и клавиша BackSpace
            {
                e.Handled = true;
            }
        }

        private void button4_Click(object sender, EventArgs e)
        {
            try
            {
            con = new OleDbConnection(BD);
            con.Open();


            con = new OleDbConnection(BD);
            con.Open();
            string queryString = "UPDATE Продукты SET [Количество] ='" + 0 + "' WHERE IDПродукта =" + comboBox1.Text;
            OleDbCommand command = new OleDbCommand(queryString, con);
            command.ExecuteNonQuery();
            con.Close();
            MessageBox.Show("Успешно удалено!");


            string put = "SELECT * FROM Продукты";
            OleDbDataAdapter dataAdapter = new OleDbDataAdapter(put, BD);
            // создаем объект DataSet
            DataSet ds = new DataSet();
            // заполняем таблицу Order  
            // данными из базы данных
            dataAdapter.Fill(ds, "Продукты");
            dataGridView1.DataSource = ds.Tables[0].DefaultView;
                comboBox1.SelectedIndex = -1;
            }
             catch
            {
            MessageBox.Show("Выберите IDПродукта!");
            }
        }

        private void button2_Click(object sender, EventArgs e)
        {

        }

        private void label5_Click(object sender, EventArgs e)
        {

        }

        private void button7_Click(object sender, EventArgs e)
        {
            MessageBox.Show("Для изменения выберите IDПродукта, который необходимо изменить. Далее введите в текстовом поле надлежайщей ячейки новое значение и нажмите на кнопку 'Изменить' \n" +
                "Для удаления выберите IDПродукта, который необходимо удалить. Далее нажмите на кнопку 'Удалить Значение' ");
        }

        private void comboBox1_SelectedIndexChanged(object sender, EventArgs e)
        {

        }

        private void dataGridView1_CellContentClick(object sender, DataGridViewCellEventArgs e)
        {

        }
    }
}
