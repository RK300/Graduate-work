using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Configuration;
using System.Data;
using System.Data.SqlClient;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;
using techSupport.edit_form;

namespace techSupport.forms
{
    public partial class worker_form : Form
    {
        private SqlConnection sqlConnection = null;

        public worker_form()
        {
            sqlConnection = new SqlConnection(ConfigurationManager.ConnectionStrings["db"].ConnectionString);
            sqlConnection.Open();
            InitializeComponent();
            RefreshTable();
        }

        private void RefreshTable()
        {
            string query = "SELECT Worker.id, Worker.surname AS [Фамилия], Worker.name AS [Имя], Worker.patronymic AS [Отчество], Position.name AS [Должность], Worker.email AS [Электронная почта], Worker.phone AS [Номер телефона], Worker.login AS [Логин] FROM Worker, Position WHERE Worker.position = Position.id";
            var connectionString = ConfigurationManager.ConnectionStrings["db"].ConnectionString;
            using (SqlDataAdapter adapter = new SqlDataAdapter(query, connectionString))
            {
                DataTable dataTable = new DataTable();
                adapter.Fill(dataTable);
                dataGridView1.DataSource = dataTable;
            }
            dataGridView1.Columns[0].Visible = false;
        }

        private void rjButton1_Click(object sender, EventArgs e)
        {
            if (new worker_edit().ShowDialog() == DialogResult.OK)
            {
                RefreshTable();
                MessageBox.Show("Запись успешно добавлена!", "Успех!");
            }
        }

        private void rjButton2_Click(object sender, EventArgs e)
        {
            worker_edit F2 = new worker_edit();
            F2.Text = "Редактирование сотрудника";
            F2.IDCHANGE = dataGridView1.Rows[dataGridView1.CurrentCell.RowIndex].Cells[0].Value.ToString();
            if (F2.ShowDialog() == DialogResult.OK)
            {
                RefreshTable();
                MessageBox.Show("Запись успешно изменена!", "Успех!");
            }
        }

        private void rjButton3_Click(object sender, EventArgs e)
        {
            if (MessageBox.Show("Вы действительно хотите удалить запись?", "Подтверждение удаления", MessageBoxButtons.YesNo, MessageBoxIcon.Warning) == DialogResult.Yes)
            {
                try
                {
                    SqlCommand command = new SqlCommand($"DELETE FROM Worker WHERE id = {dataGridView1.Rows[dataGridView1.CurrentCell.RowIndex].Cells[0].Value}", sqlConnection);
                    command.ExecuteNonQuery();
                    MessageBox.Show("Запись успешно удалена!", "Успех!");
                    RefreshTable();
                }
                catch
                {
                    MessageBox.Show("Ошибка, попробуйте еще раз!", "Внимание", MessageBoxButtons.OK, MessageBoxIcon.Error);
                }
            }
        }

        private void rjButton5_Click(object sender, EventArgs e)
        {
            if (!String.IsNullOrWhiteSpace(textBox1.Text))
            {
                string query = $"SELECT Worker.id, Worker.surname AS [Фамилия], Worker.name AS [Имя], Worker.patronymic AS [Отчество], Position.name AS [Должность], Worker.email AS [Электронная почта], Worker.phone AS [Номер телефона], Worker.login AS [Логин], Worker.password AS [Пароль] FROM Worker, Position WHERE Worker.position = Position.id AND Worker.surname + ' ' + Worker.name + ' ' + Worker.patronymic + ' ' + Worker.email + ' ' + Worker.phone + ' ' + Worker.login + ' ' + Worker.password LIKE '%{textBox1.Text}%'";
                var connectionString = ConfigurationManager.ConnectionStrings["db"].ConnectionString;
                using (SqlDataAdapter adapter = new SqlDataAdapter(query, connectionString))
                {
                    DataTable dataTable = new DataTable();
                    adapter.Fill(dataTable);
                    dataGridView1.DataSource = dataTable;
                }
            }
            else
            {
                RefreshTable();
            }
        }
    }
}