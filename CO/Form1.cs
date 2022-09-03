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
    public partial class Form1 : Form
    {
        public Form1()
        {
            InitializeComponent();
        }

        private void label2_Click(object sender, EventArgs e)
        {

        }

        private void Form1_Load(object sender, EventArgs e)
        {
        }

        private void textBox1_Enter(object sender, EventArgs e)
        {

            if (textBox1.Text == "Логин")
            {
                textBox1.Text = "";

                textBox1.ForeColor = Color.Black;
            }
        }

        private void textBox1_Leave(object sender, EventArgs e)
        {
            if (textBox1.Text == "")
            {
                textBox1.Text = "Логин";

                textBox1.ForeColor = Color.Silver;
            }
        }

        private void textBox2_Enter(object sender, EventArgs e)
        {

            if (textBox2.Text == "Пароль")
            {
                textBox2.Text = "";

                textBox2.ForeColor = Color.Black;
                textBox2.UseSystemPasswordChar = true;
            }
        }

        private void textBox2_Leave(object sender, EventArgs e)
        {
            if (textBox2.Text == "")
            {
                textBox2.Text = "Пароль";

                textBox2.ForeColor = Color.Silver;
            }
        }

        private void button5_Click(object sender, EventArgs e)
        {
            this.Hide();
            frmMain ss = new frmMain();
            ss.Show();
        }

        private void button1_Click(object sender, EventArgs e)
        {
            OleDbConnection con = new OleDbConnection(@"Provider=Microsoft.Jet.OLEDB.4.0;Data Source=|DataDirectory|\\Coffeeorange.mdb");
            OleDbDataAdapter dataAdapter = new OleDbDataAdapter("Select Логин, ФИО,[IDПользователя]  From Пользователь where Логин ='" + textBox1.Text + "' and Пароль ='" + textBox2.Text + "'", con);
            DataTable dt = new DataTable();
            dataAdapter.Fill(dt);
            // Проверяем, что количество строк из БД больше нуля
            if (dt.Rows.Count > 0)
            {
                // Нужный Вам ID
                string ID = dt.Rows[0][0].ToString();
                if (ID != "Admin") 
                {
                    string fio = dt.Rows[0][1].ToString();
                    string id = dt.Rows[0][2].ToString();
                    Class1.fio = fio;
                    Class1.ID = id;
                    this.Hide();
                    Form3 ss = new Form3();
                    ss.Show();
                }
            }
            else
            {
                MessageBox.Show("Неправильно введённые Логин или пароль");
            }
        }

        private void button9_Click(object sender, EventArgs e)
        {
            textBox1.Clear();
            textBox2.Clear();
        }
    }
}
