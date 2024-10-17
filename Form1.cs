using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Data.OleDb;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Text.RegularExpressions; // Добавляем пространство имен для работы с регулярными выражениями
using System.Threading.Tasks;
using System.Windows.Forms;

namespace m6
{
    public partial class Form1 : Form
    {
        public Form1()
        {
            InitializeComponent();
        }

        private void выходToolStripMenuItem_Click(object sender, EventArgs e)
        {
            Application.Exit();
        }

        private void btn_Install_Click(object sender, EventArgs e)
        {
            // Создаем соединение
            string connectionString = "provider=Microsoft.Jet.OLEDB.4.0;Data Source=Database.mdb";
            OleDbConnection dbConnection = new OleDbConnection(connectionString); // создаем соединение

            // Выполняем запрос к БД
            dbConnection.Open();
            string query = "SELECT * FROM table1";
            OleDbCommand dbCommand = new OleDbCommand(query, dbConnection);
            OleDbDataReader dbReader = dbCommand.ExecuteReader();

            // Проверка данных
            if (dbReader.HasRows == false)
            {
                MessageBox.Show("Данные не найдены", "Ошибка!");
            }
            else
            {
                while (dbReader.Read())
                {
                    dataGridView1.Rows.Add(dbReader["id"], dbReader["description"], dbReader["d"], dbReader["state"]);
                }
            }

            // Закрываем соединение
            dbReader.Close();
            dbConnection.Close();
        }

        private bool IsDateValid(string date)
        {
            // Проверяем дату с помощью регулярного выражения в формате "dd.mm.yyyy"
            Regex regex = new Regex(@"^\d{2}\.\d{2}\.\d{4}$");
            return regex.IsMatch(date);
        }

        private void btn_Add_Click(object sender, EventArgs e)
        {
            // Проверка количества выбранных строк
            if (dataGridView1.SelectedRows.Count != 1)
            {
                MessageBox.Show("Выберите одну строку!", "Внимание!");
                return;
            }

            // Заполним выбранную строку
            int index = dataGridView1.SelectedRows[0].Index;

            // Проверка на данные в таблице
            if (dataGridView1.Rows[index].Cells[0].Value == null ||
                dataGridView1.Rows[index].Cells[1].Value == null ||
                dataGridView1.Rows[index].Cells[2].Value == null ||
                dataGridView1.Rows[index].Cells[3].Value == null)
            {
                MessageBox.Show("Не все данные введены!", "Внимание!");
                return;
            }

            // Считываем данные
            string id = dataGridView1.Rows[index].Cells[0].Value.ToString();
            string description = dataGridView1.Rows[index].Cells[1].Value.ToString();
            string d = dataGridView1.Rows[index].Cells[2].Value.ToString();
            string state = dataGridView1.Rows[index].Cells[3].Value.ToString();

            // Проверка формата даты
            if (!IsDateValid(d))
            {
                MessageBox.Show("Дата должна быть введена в формате dd.mm.yyyy", "Ошибка!");
                return;
            }

            // Создаем соединение
            string connectionString = "provider=Microsoft.Jet.OLEDB.4.0;Data Source=Database.mdb";
            OleDbConnection dbConnection = new OleDbConnection(connectionString); // создаем соединение

            // Выполняем запрос к БД
            dbConnection.Open();
            string query = "INSERT INTO table1 VALUES (" + id + ", '" + description + "', '" + d + "', '" + state + "')";
            OleDbCommand dbCommand = new OleDbCommand(query, dbConnection);

            // Выполняем запрос
            if (dbCommand.ExecuteNonQuery() != 1)
                MessageBox.Show("Ошибка выполнения запроса!", "Ошибка!");
            else
                MessageBox.Show("Данные добавлены", "Внимание!");

            // Закрываем соединение с БД
            dbConnection.Close();
        }

        private void btn_Update_Click(object sender, EventArgs e)
        {
            // Проверка количества выбранных строк
            if (dataGridView1.SelectedRows.Count != 1)
            {
                MessageBox.Show("Выберите одну строку!", "Внимание!");
                return;
            }

            // Заполним выбранную строку
            int index = dataGridView1.SelectedRows[0].Index;

            // Проверка на данные в таблице
            if (dataGridView1.Rows[index].Cells[0].Value == null ||
                dataGridView1.Rows[index].Cells[1].Value == null ||
                dataGridView1.Rows[index].Cells[2].Value == null ||
                dataGridView1.Rows[index].Cells[3].Value == null)
            {
                MessageBox.Show("Не все данные введены!", "Внимание!");
                return;
            }

            // Считываем данные
            string id = dataGridView1.Rows[index].Cells[0].Value.ToString();
            string description = dataGridView1.Rows[index].Cells[1].Value.ToString();
            string d = dataGridView1.Rows[index].Cells[2].Value.ToString();
            string state = dataGridView1.Rows[index].Cells[3].Value.ToString();

            // Проверка формата даты
            if (!IsDateValid(d))
            {
                MessageBox.Show("Дата должна быть введена в формате dd.mm.yyyy", "Ошибка!");
                return;
            }

            // Создаем соединение
            string connectionString = "provider=Microsoft.Jet.OLEDB.4.0;Data Source=Database.mdb";
            OleDbConnection dbConnection = new OleDbConnection(connectionString); // создаем соединение

            // Выполняем запрос к БД
            dbConnection.Open();
            string query = "UPDATE table1 SET Description ='" + description + "', D = '" + d + "', State = '" + state + "' WHERE id = " + id;
            OleDbCommand dbCommand = new OleDbCommand(query, dbConnection);

            // Выполняем запрос
            if (dbCommand.ExecuteNonQuery() != 1)
                MessageBox.Show("Ошибка выполнения запроса!", "Ошибка!");
            else
            {
                MessageBox.Show("Данные изменены", "Внимание!");
            }

            // Закрываем соединение с БД
            dbConnection.Close();
        }

        private void btn_Clear_Click(object sender, EventArgs e)
        {
            //проверка количества выбранных строк
            if (dataGridView1.SelectedRows.Count != 1)
            {
                MessageBox.Show("Выберите одну строку!", "Внимение!");
                return;
            }

            //Заполним выбранную строку
            int index = dataGridView1.SelectedRows[0].Index;

            //Проверка на данные в таблице
            if (dataGridView1.Rows[index].Cells[0].Value == null)
            {
                MessageBox.Show("Не все данные введены!", "Внимение!");
                return;
            }

            //Считываем данные
            string id = dataGridView1.Rows[index].Cells[0].Value.ToString();
            
            //Создаем соединение
            string connectionString = "provider=Microsoft.Jet.OLEDB.4.0;Data Source=Database.mdb";//строка соединения
            OleDbConnection dbConnection = new OleDbConnection(connectionString); //создаем соединение

            //Выполняем запрос к БД
            dbConnection.Open();
            string query = "DELETE FROM table1 WHERE id =" + id;
            OleDbCommand dbCommand = new OleDbCommand(query, dbConnection);

            //Выполняем запрос
            if (dbCommand.ExecuteNonQuery() != 1)
                MessageBox.Show("Ошибка выполнения запроса!", "Ошибка!");
            else
            {
                MessageBox.Show("Данные удалены", "Внимание!");
                //Удаляем данные из таблицы 
                dataGridView1.Rows.RemoveAt(index);

            }

            //Закрываем соединение с БД
            dbConnection.Close();
        }
    }
}
