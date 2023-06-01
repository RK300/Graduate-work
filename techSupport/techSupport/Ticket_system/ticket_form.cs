using Microsoft.Office.Interop.Word;
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
using System.Windows.Media;
using techSupport.Analytics;
using techSupport.edit_form;
using techSupport.edit_forms_two;
using techSupport.view_form;
using Color = System.Drawing.Color;

namespace techSupport.Ticket_system
{
    public partial class ticket_form : Form
    {
        private SqlConnection sqlConnection = null;

        private int UserId;

        private void setUserId(int id) { UserId = id; }

        public ticket_form(int id)
        {
            setUserId(id);
            sqlConnection = new SqlConnection(ConfigurationManager.ConnectionStrings["db"].ConnectionString);
            sqlConnection.Open();
            InitializeComponent();
            RefreshTable();
        }

        private void RefreshTable()
        {
            //string query = "SELECT Ticket.id, (Clients.surname + ' ' + Clients.name + ' ' + Clients.patronymic) AS [Клиент], Products.name AS [Продукт], (Worker.surname + ' ' + Worker.name + ' ' + Worker.patronymic) AS [Принявший сотрудник], Type.name AS [Тип проблемы], Ticket.header AS [Заголовок], Ticket.description AS [Описание], Ticket.priority AS [Приоритет], Ticket.status AS [Статус], Ticket.application_data AS [Дата подачи], Ticket.completion_data AS [Дата закрытия], Ticket.actual_data AS [Фактическая дата закрытия] FROM Ticket, Clients, Products, Worker, Type WHERE Ticket.client = Clients.id AND Ticket.product = Products.id AND Ticket.worker = Worker.id AND Ticket.type = Type.id";
            //string query = "SELECT Ticket.id, (Clients.surname + ' ' + Clients.name + ' ' + Clients.patronymic) AS [Клиент], Products.name AS [Продукт], (Worker.surname + ' ' + Worker.name + ' ' + Worker.patronymic) AS [Принявший сотрудник], Type.name AS [Тип проблемы], Ticket.header AS [Заголовок], Ticket.priority AS [Приоритет], Ticket.status AS [Статус], Ticket.application_data AS [Дата подачи], Ticket.completion_data AS [Дата закрытия], Ticket.actual_data AS [Фактическая дата закрытия] FROM Ticket, Clients, Products, Worker, Type WHERE Ticket.client = Clients.id AND Ticket.product = Products.id AND Ticket.worker = Worker.id AND Ticket.type = Type.id";
            //string query = "SELECT Ticket.id, (Clients.surname + ' ' + Clients.name + ' ' + Clients.patronymic) AS [Клиент], Products.name AS [Продукт], Type.name AS [Тип проблемы], Ticket.header AS [Заголовок], Ticket.priority AS [Приоритет], Ticket.status AS [Статус], Ticket.application_data AS [Дата подачи], Ticket.completion_data AS [Дата закрытия], Ticket.actual_data AS [Фактическая дата закрытия] FROM Ticket, Clients, Products, Worker, Type WHERE Ticket.client = Clients.id AND Ticket.product = Products.id AND Ticket.worker = Worker.id AND Ticket.type = Type.id";
            string query = "SELECT Ticket.id, Clients.CompanyName AS [Клиент], Products.name AS [Продукт], Type.name AS [Тип проблемы], Ticket.header AS [Заголовок], Ticket.priority AS [Приоритет], Ticket.status AS [Статус], Ticket.application_data AS [Дата подачи], Ticket.completion_data AS [Дата закрытия] FROM Ticket, Clients, Products, Worker, Type WHERE Ticket.client = Clients.id AND Ticket.product = Products.id AND Ticket.worker = Worker.id AND Ticket.type = Type.id";
            var connectionString = ConfigurationManager.ConnectionStrings["db"].ConnectionString;
            using (SqlDataAdapter adapter = new SqlDataAdapter(query, connectionString))
            {
                System.Data.DataTable dataTable = new System.Data.DataTable();
                adapter.Fill(dataTable);
                dataGridView1.DataSource = dataTable;
            }

            dataGridView1.Columns[0].Visible = false;

            for (int i = 0; i < this.dataGridView1.Rows.Count - 1; i++)
            {
                if (this.dataGridView1.Rows[i].Cells["Статус"].Value.ToString() == "Открыт")
                {
                    for (int j = 0; j < dataGridView1.Rows[i].Cells.Count; j++)
                    {
                        this.dataGridView1.Rows[i].Cells[j].Style.BackColor = UIColors.SpecGreen;
                    }
                }
                if (this.dataGridView1.Rows[i].Cells["Статус"].Value.ToString() == "В работе")
                {
                    for (int j = 0; j < dataGridView1.Rows[i].Cells.Count; j++)
                    {
                        this.dataGridView1.Rows[i].Cells[j].Style.BackColor = UIColors.SpecYellow;
                    }
                }
                if (this.dataGridView1.Rows[i].Cells["Статус"].Value.ToString() == "Закрыт")
                {
                    for (int j = 0; j < dataGridView1.Rows[i].Cells.Count; j++)
                    {
                        this.dataGridView1.Rows[i].Cells[j].Style.BackColor = UIColors.SpecRed;
                    }
                }
            }
        }

        private void rjButton1_Click(object sender, EventArgs e)
        {
            if (new ticket_edit(UserId).ShowDialog() == DialogResult.OK)
            {
                RefreshTable();
                rjRadioButton1.Checked = true;
                MessageBox.Show("Запись успешно добавлена!", "Успех!");

                if ((dataGridView1.Rows.Count - 2) < 0)
                {
                    dataGridView1.CurrentCell = dataGridView1.Rows[0].Cells[1];
                }
                else
                {
                    dataGridView1.CurrentCell = dataGridView1.Rows[dataGridView1.Rows.Count - 2].Cells[1];
                    SubRefresh();
                }
            }
        }

        private void rjButton2_Click(object sender, EventArgs e)
        {
            ticket_edit F2 = new ticket_edit(UserId);
            F2.Text = "Редактирование тикета";
            F2.IDCHANGE = dataGridView1.Rows[dataGridView1.CurrentCell.RowIndex].Cells[0].Value.ToString();
            var m_curentCell = dataGridView1.CurrentCell.RowIndex;
            if (F2.ShowDialog() == DialogResult.OK)
            {
                RefreshTable();
                rjRadioButton1.Checked = true;
                MessageBox.Show("Запись успешно изменена!", "Успех!");
                if ((dataGridView1.Rows.Count - 1) < 0)
                {
                    dataGridView1.CurrentCell = dataGridView1.Rows[0].Cells[1];
                }
                else
                {
                    dataGridView1.CurrentCell = dataGridView1.Rows[m_curentCell].Cells[1];
                    SubRefresh();
                }
            }
        }

        private void rjButton3_Click(object sender, EventArgs e)
        {
            if (MessageBox.Show("Вы действительно хотите удалить запись?", "Подтверждение удаления", MessageBoxButtons.YesNo, MessageBoxIcon.Warning) == DialogResult.Yes)
            {
                try
                {
                    SqlCommand command = new SqlCommand($"DELETE FROM Ticket WHERE id = {dataGridView1.Rows[dataGridView1.CurrentCell.RowIndex].Cells[0].Value}", sqlConnection);
                    command.ExecuteNonQuery();
                    MessageBox.Show("Запись успешно удалена!", "Успех!");
                    RefreshTable();
                    rjRadioButton1.Checked = true;
                }
                catch
                {
                    MessageBox.Show("Ошибка, попробуйте еще раз!", "Внимание", MessageBoxButtons.OK, MessageBoxIcon.Error);
                }
            }
        }

        private void rjButton5_Click(object sender, EventArgs e)
        {
            if ( rjRadioButton1.Checked ) 
            {
                if (!String.IsNullOrWhiteSpace(textBox1.Text))
                {
                    //string query = $"SELECT Ticket.id, (Clients.surname + ' ' + Clients.name + ' ' + Clients.patronymic) AS [Клиент], Products.name AS [Продукт], (Worker.surname + ' ' + Worker.name + ' ' + Worker.patronymic) AS [Принявший сотрудник], Type.name AS [Тип проблемы], Ticket.header AS [Заголовок], Ticket.description AS [Описание], Ticket.priority AS [Приоритет], Ticket.status AS [Статус], Ticket.application_data AS [Дата подачи], Ticket.completion_data AS [Дата закрытия], Ticket.actual_data AS [Фактическая дата закрытия] FROM Ticket, Clients, Products, Worker, Type WHERE Ticket.client = Clients.id AND Ticket.product = Products.id AND Ticket.worker = Worker.id AND Ticket.type = Type.id AND Ticket.header + ' ' + Ticket.description + ' ' + Ticket.priority + ' ' + Ticket.status LIKE '%{textBox1.Text}%'";
                    //string query = $"SELECT Ticket.id, (Clients.surname + ' ' + Clients.name + ' ' + Clients.patronymic) AS [Клиент], Products.name AS [Продукт], (Worker.surname + ' ' + Worker.name + ' ' + Worker.patronymic) AS [Принявший сотрудник], Type.name AS [Тип проблемы], Ticket.header AS [Заголовок], Ticket.priority AS [Приоритет], Ticket.status AS [Статус], Ticket.application_data AS [Дата подачи], Ticket.completion_data AS [Дата закрытия], Ticket.actual_data AS [Фактическая дата закрытия] FROM Ticket, Clients, Products, Worker, Type WHERE Ticket.client = Clients.id AND Ticket.product = Products.id AND Ticket.worker = Worker.id AND Ticket.type = Type.id AND Ticket.header + ' ' + Ticket.description + ' ' + Ticket.priority + ' ' + Ticket.status LIKE '%{textBox1.Text}%'";
                    string query = $"SELECT Ticket.id, Clients.CompanyName AS [Клиент], Products.name AS [Продукт], (Worker.surname + ' ' + Worker.name + ' ' + Worker.patronymic) AS [Принявший сотрудник], Type.name AS [Тип проблемы], Ticket.header AS [Заголовок], Ticket.priority AS [Приоритет], Ticket.status AS [Статус], Ticket.application_data AS [Дата подачи], Ticket.completion_data AS [Дата закрытия] FROM Ticket, Clients, Products, Worker, Type WHERE Ticket.client = Clients.id AND Ticket.product = Products.id AND Ticket.worker = Worker.id AND Ticket.type = Type.id AND Clients.CompanyName + ' ' + Products.name + ' ' + Type.name + ' ' + Ticket.header + ' ' + Ticket.description + ' ' + Ticket.priority + ' ' + Ticket.status LIKE '%{textBox1.Text}%'";
                    var connectionString = ConfigurationManager.ConnectionStrings["db"].ConnectionString;
                    using (SqlDataAdapter adapter = new SqlDataAdapter(query, connectionString))
                    {
                        System.Data.DataTable dataTable = new System.Data.DataTable();
                        adapter.Fill(dataTable);
                        dataGridView1.DataSource = dataTable;
                    }

                    for (int i = 0; i < this.dataGridView1.Rows.Count - 1; i++)
                    {
                        if (this.dataGridView1.Rows[i].Cells["Статус"].Value.ToString() == "Открыт")
                        {
                            for (int j = 0; j < dataGridView1.Rows[i].Cells.Count; j++)
                            {
                                this.dataGridView1.Rows[i].Cells[j].Style.BackColor = UIColors.SpecGreen;
                            }
                        }
                        if (this.dataGridView1.Rows[i].Cells["Статус"].Value.ToString() == "В работе")
                        {
                            for (int j = 0; j < dataGridView1.Rows[i].Cells.Count; j++)
                            {
                                this.dataGridView1.Rows[i].Cells[j].Style.BackColor = UIColors.SpecYellow;
                            }
                        }
                        if (this.dataGridView1.Rows[i].Cells["Статус"].Value.ToString() == "Закрыт")
                        {
                            for (int j = 0; j < dataGridView1.Rows[i].Cells.Count; j++)
                            {
                                this.dataGridView1.Rows[i].Cells[j].Style.BackColor = UIColors.SpecRed;
                            }
                        }
                    }
                }
                else
                {
                    RefreshTable();
                    rjRadioButton1.Checked = true;
                }
            }

            if (rjRadioButton4.Checked)
            {
                if (!String.IsNullOrWhiteSpace(textBox1.Text))
                {
                    //string query = $"SELECT Ticket.id, (Clients.surname + ' ' + Clients.name + ' ' + Clients.patronymic) AS [Клиент], Products.name AS [Продукт], (Worker.surname + ' ' + Worker.name + ' ' + Worker.patronymic) AS [Принявший сотрудник], Type.name AS [Тип проблемы], Ticket.header AS [Заголовок], Ticket.description AS [Описание], Ticket.priority AS [Приоритет], Ticket.status AS [Статус], Ticket.application_data AS [Дата подачи], Ticket.completion_data AS [Дата закрытия], Ticket.actual_data AS [Фактическая дата закрытия] FROM Ticket, Clients, Products, Worker, Type WHERE Ticket.status != 'Закрыт' AND Ticket.client = Clients.id AND Ticket.product = Products.id AND Ticket.worker = Worker.id AND Ticket.type = Type.id AND Ticket.header + ' ' + Ticket.description + ' ' + Ticket.priority + ' ' + Ticket.status LIKE '%{textBox1.Text}%'";
                    //string query = $"SELECT Ticket.id, (Clients.surname + ' ' + Clients.name + ' ' + Clients.patronymic) AS [Клиент], Products.name AS [Продукт], (Worker.surname + ' ' + Worker.name + ' ' + Worker.patronymic) AS [Принявший сотрудник], Type.name AS [Тип проблемы], Ticket.header AS [Заголовок], Ticket.priority AS [Приоритет], Ticket.status AS [Статус], Ticket.application_data AS [Дата подачи], Ticket.completion_data AS [Дата закрытия], Ticket.actual_data AS [Фактическая дата закрытия] FROM Ticket, Clients, Products, Worker, Type WHERE Ticket.status != 'Закрыт' AND Ticket.client = Clients.id AND Ticket.product = Products.id AND Ticket.worker = Worker.id AND Ticket.type = Type.id AND Ticket.header + ' ' + Ticket.description + ' ' + Ticket.priority + ' ' + Ticket.status LIKE '%{textBox1.Text}%'";
                    string query = $"SELECT Ticket.id, Clients.CompanyName AS [Клиент], Products.name AS [Продукт], (Worker.surname + ' ' + Worker.name + ' ' + Worker.patronymic) AS [Принявший сотрудник], Type.name AS [Тип проблемы], Ticket.header AS [Заголовок], Ticket.priority AS [Приоритет], Ticket.status AS [Статус], Ticket.application_data AS [Дата подачи], Ticket.completion_data AS [Дата закрытия] FROM Ticket, Clients, Products, Worker, Type WHERE Ticket.status != 'Закрыт' AND Ticket.client = Clients.id AND Ticket.product = Products.id AND Ticket.worker = Worker.id AND Ticket.type = Type.id AND Clients.CompanyName + ' ' + Products.name + ' ' + Type.name + ' ' + Ticket.header + ' ' + Ticket.description + ' ' + Ticket.priority + ' ' + Ticket.status LIKE '%{textBox1.Text}%'";
                    var connectionString = ConfigurationManager.ConnectionStrings["db"].ConnectionString;
                    using (SqlDataAdapter adapter = new SqlDataAdapter(query, connectionString))
                    {
                        System.Data.DataTable dataTable = new System.Data.DataTable();
                        adapter.Fill(dataTable);
                        dataGridView1.DataSource = dataTable;
                    }

                    for (int i = 0; i < this.dataGridView1.Rows.Count - 1; i++)
                    {
                        if (this.dataGridView1.Rows[i].Cells["Статус"].Value.ToString() == "Открыт")
                        {
                            for (int j = 0; j < dataGridView1.Rows[i].Cells.Count; j++)
                            {
                                this.dataGridView1.Rows[i].Cells[j].Style.BackColor = UIColors.SpecGreen;
                            }
                        }
                        if (this.dataGridView1.Rows[i].Cells["Статус"].Value.ToString() == "В работе")
                        {
                            for (int j = 0; j < dataGridView1.Rows[i].Cells.Count; j++)
                            {
                                this.dataGridView1.Rows[i].Cells[j].Style.BackColor = UIColors.SpecYellow;
                            }
                        }
                        if (this.dataGridView1.Rows[i].Cells["Статус"].Value.ToString() == "Закрыт")
                        {
                            for (int j = 0; j < dataGridView1.Rows[i].Cells.Count; j++)
                            {
                                this.dataGridView1.Rows[i].Cells[j].Style.BackColor = UIColors.SpecRed;
                            }
                        }
                    }
                }
                else
                {
                    RefreshTable4();
                }
            }

            if (rjRadioButton5.Checked)
            {
                if (!String.IsNullOrWhiteSpace(textBox1.Text))
                {
                    //string query = $"SELECT Ticket.id, (Clients.surname + ' ' + Clients.name + ' ' + Clients.patronymic) AS [Клиент], Products.name AS [Продукт], (Worker.surname + ' ' + Worker.name + ' ' + Worker.patronymic) AS [Принявший сотрудник], Type.name AS [Тип проблемы], Ticket.header AS [Заголовок], Ticket.description AS [Описание], Ticket.priority AS [Приоритет], Ticket.status AS [Статус], Ticket.application_data AS [Дата подачи], Ticket.completion_data AS [Дата закрытия], Ticket.actual_data AS [Фактическая дата закрытия] FROM Ticket, Clients, Products, Worker, Type WHERE Ticket.status = 'Закрыт' AND Ticket.client = Clients.id AND Ticket.product = Products.id AND Ticket.worker = Worker.id AND Ticket.type = Type.id AND Ticket.header + ' ' + Ticket.description + ' ' + Ticket.priority + ' ' + Ticket.status LIKE '%{textBox1.Text}%'";
                    string query = $"SELECT Ticket.id, Clients.CompanyName AS [Клиент], Products.name AS [Продукт], (Worker.surname + ' ' + Worker.name + ' ' + Worker.patronymic) AS [Принявший сотрудник], Type.name AS [Тип проблемы], Ticket.header AS [Заголовок], Ticket.priority AS [Приоритет], Ticket.status AS [Статус], Ticket.application_data AS [Дата подачи], Ticket.completion_data AS [Дата закрытия], Ticket.actual_data AS [Фактическая дата закрытия] FROM Ticket, Clients, Products, Worker, Type WHERE Ticket.status = 'Закрыт' AND Ticket.client = Clients.id AND Ticket.product = Products.id AND Ticket.worker = Worker.id AND Ticket.type = Type.id AND Clients.CompanyName + ' ' + Products.name + ' ' + Type.name + ' ' + Ticket.header + ' ' + Ticket.description + ' ' + Ticket.priority + ' ' + Ticket.status LIKE '%{textBox1.Text}%'";
                    var connectionString = ConfigurationManager.ConnectionStrings["db"].ConnectionString;
                    using (SqlDataAdapter adapter = new SqlDataAdapter(query, connectionString))
                    {
                        System.Data.DataTable dataTable = new System.Data.DataTable();
                        adapter.Fill(dataTable);
                        dataGridView1.DataSource = dataTable;
                    }

                    for (int i = 0; i < this.dataGridView1.Rows.Count - 1; i++)
                    {
                        if (this.dataGridView1.Rows[i].Cells["Статус"].Value.ToString() == "Закрыт")
                        {
                            for (int j = 0; j < dataGridView1.Rows[i].Cells.Count; j++)
                            {
                                this.dataGridView1.Rows[i].Cells[j].Style.BackColor = UIColors.SpecRed;
                            }
                        }
                    }
                }
                else
                {
                    RefreshTable5();
                }
            }

            if (rjRadioButton2.Checked)
            {
                if (!String.IsNullOrWhiteSpace(textBox1.Text))
                {
                    //string query = $"SELECT Ticket.id, (Clients.surname + ' ' + Clients.name + ' ' + Clients.patronymic) AS [Клиент], Products.name AS [Продукт], (Worker.surname + ' ' + Worker.name + ' ' + Worker.patronymic) AS [Принявший сотрудник], Type.name AS [Тип проблемы], Ticket.header AS [Заголовок], Ticket.description AS [Описание], Ticket.priority AS [Приоритет], Ticket.status AS [Статус], Ticket.application_data AS [Дата подачи], Ticket.completion_data AS [Дата закрытия], Ticket.actual_data AS [Фактическая дата закрытия] FROM Ticket, Clients, Products, Worker, Type WHERE Ticket.status != 'Закрыт' AND Ticket.client = Clients.id AND Ticket.product = Products.id AND Ticket.worker = Worker.id AND Ticket.type = Type.id AND Ticket.header + ' ' + Ticket.description + ' ' + Ticket.priority + ' ' + Ticket.status LIKE '%{textBox1.Text}%'";
                    //string query = $"SELECT Ticket.id, (Clients.surname + ' ' + Clients.name + ' ' + Clients.patronymic) AS [Клиент], Products.name AS [Продукт], (Worker.surname + ' ' + Worker.name + ' ' + Worker.patronymic) AS [Принявший сотрудник], Type.name AS [Тип проблемы], Ticket.header AS [Заголовок], Ticket.priority AS [Приоритет], Ticket.status AS [Статус], Ticket.application_data AS [Дата подачи], Ticket.completion_data AS [Дата закрытия], Ticket.actual_data AS [Фактическая дата закрытия] FROM Ticket, Clients, Products, Worker, Type WHERE Ticket.status != 'Закрыт' AND Ticket.client = Clients.id AND Ticket.product = Products.id AND Ticket.worker = Worker.id AND Ticket.type = Type.id AND Ticket.header + ' ' + Ticket.description + ' ' + Ticket.priority + ' ' + Ticket.status LIKE '%{textBox1.Text}%'";
                    //string query = $"SELECT Ticket.id, (Clients.surname + ' ' + Clients.name + ' ' + Clients.patronymic) AS [Клиент], Products.name AS [Продукт], Type.name AS [Тип проблемы], Ticket.header AS [Заголовок], Ticket.priority AS [Приоритет], Ticket.status AS [Статус], Ticket.application_data AS [Дата подачи], Ticket.completion_data AS [Дата закрытия], Ticket.actual_data AS [Фактическая дата закрытия] FROM Ticket, Clients, Products, Worker, Type WHERE Ticket.status != 'Закрыт' AND Ticket.client = Clients.id AND Ticket.product = Products.id AND Ticket.worker = Worker.id AND Ticket.type = Type.id AND Ticket.header + ' ' + Ticket.description + ' ' + Ticket.priority + ' ' + Ticket.status LIKE '%{textBox1.Text}%'";
                    string query = $"SELECT Ticket.id, Clients.CompanyName AS [Клиент], Products.name AS [Продукт], Type.name AS [Тип проблемы], Ticket.header AS [Заголовок], Ticket.priority AS [Приоритет], Ticket.status AS [Статус], Ticket.application_data AS [Дата подачи], Ticket.completion_data AS [Дата закрытия] FROM Ticket, Clients, Products, Worker, Type WHERE Ticket.status != 'Закрыт' AND Ticket.client = Clients.id AND Ticket.product = Products.id AND Ticket.worker = Worker.id AND Ticket.type = Type.id AND Clients.CompanyName + ' ' + Products.name + ' ' + Type.name + ' ' + Ticket.header + ' ' + Ticket.description + ' ' + Ticket.priority + ' ' + Ticket.status LIKE '%{textBox1.Text}%'";
                    var connectionString = ConfigurationManager.ConnectionStrings["db"].ConnectionString;
                    using (SqlDataAdapter adapter = new SqlDataAdapter(query, connectionString))
                    {
                        System.Data.DataTable dataTable = new System.Data.DataTable();
                        adapter.Fill(dataTable);
                        dataGridView1.DataSource = dataTable;
                    }

                    for (int i = 0; i < this.dataGridView1.Rows.Count - 1; i++)
                    {
                        if (this.dataGridView1.Rows[i].Cells["Приоритет"].Value.ToString() == "Низкий")
                        {
                            for (int j = 0; j < dataGridView1.Rows[i].Cells.Count; j++)
                            {
                                this.dataGridView1.Rows[i].Cells[j].Style.BackColor = UIColors.PriorityLow;
                            }
                        }
                        if (this.dataGridView1.Rows[i].Cells["Приоритет"].Value.ToString() == "Средний")
                        {
                            for (int j = 0; j < dataGridView1.Rows[i].Cells.Count; j++)
                            {
                                this.dataGridView1.Rows[i].Cells[j].Style.BackColor = UIColors.PriorityMedium;
                            }
                        }
                        if (this.dataGridView1.Rows[i].Cells["Приоритет"].Value.ToString() == "Высокий")
                        {
                            for (int j = 0; j < dataGridView1.Rows[i].Cells.Count; j++)
                            {
                                this.dataGridView1.Rows[i].Cells[j].Style.BackColor = UIColors.PriorityHight;
                            }
                        }
                    }
                }
                else
                {
                    RefreshTable2();
                }
            }

            if (rjRadioButton3.Checked)
            {
                if (!String.IsNullOrWhiteSpace(textBox1.Text))
                {
                    //string query = $"SELECT Ticket.id, (Clients.surname + ' ' + Clients.name + ' ' + Clients.patronymic) AS [Клиент], Products.name AS [Продукт], (Worker.surname + ' ' + Worker.name + ' ' + Worker.patronymic) AS [Принявший сотрудник], Type.name AS [Тип проблемы], Ticket.header AS [Заголовок], Ticket.description AS [Описание], Ticket.priority AS [Приоритет], Ticket.status AS [Статус], Ticket.application_data AS [Дата подачи], Ticket.completion_data AS [Дата закрытия], Ticket.actual_data AS [Фактическая дата закрытия] FROM Ticket, Clients, Products, Worker, Type WHERE Ticket.actual_data > Ticket.completion_data AND Ticket.client = Clients.id AND Ticket.product = Products.id AND Ticket.worker = Worker.id AND Ticket.type = Type.id AND Ticket.header + ' ' + Ticket.description + ' ' + Ticket.priority + ' ' + Ticket.status LIKE '%{textBox1.Text}%'";
                    //string query = $"SELECT Ticket.id, (Clients.surname + ' ' + Clients.name + ' ' + Clients.patronymic) AS [Клиент], Products.name AS [Продукт], (Worker.surname + ' ' + Worker.name + ' ' + Worker.patronymic) AS [Принявший сотрудник], Type.name AS [Тип проблемы], Ticket.header AS [Заголовок], Ticket.priority AS [Приоритет], Ticket.status AS [Статус], Ticket.application_data AS [Дата подачи], Ticket.completion_data AS [Дата закрытия], Ticket.actual_data AS [Фактическая дата закрытия] FROM Ticket, Clients, Products, Worker, Type WHERE Ticket.actual_data > Ticket.completion_data AND Ticket.client = Clients.id AND Ticket.product = Products.id AND Ticket.worker = Worker.id AND Ticket.type = Type.id AND Ticket.header + ' ' + Ticket.description + ' ' + Ticket.priority + ' ' + Ticket.status LIKE '%{textBox1.Text}%'";
                    string query = $"SELECT Ticket.id, Clients.CompanyName AS [Клиент], Products.name AS [Продукт], Type.name AS [Тип проблемы], Ticket.header AS [Заголовок], Ticket.priority AS [Приоритет], Ticket.status AS [Статус], Ticket.application_data AS [Дата подачи], Ticket.completion_data AS [Дата закрытия], Ticket.actual_data AS [Фактическая дата закрытия] FROM Ticket, Clients, Products, Worker, Type WHERE Ticket.completion_data < CAST('{DateTime.Now.ToString()}' AS DATETIME) AND Ticket.status != 'Закрыт' AND Ticket.client = Clients.id AND Ticket.product = Products.id AND Ticket.worker = Worker.id AND Ticket.type = Type.id AND Clients.CompanyName + ' ' + Products.name + ' ' + Type.name + ' ' + Ticket.header + ' ' + Ticket.description + ' ' + Ticket.priority + ' ' + Ticket.status LIKE '%{textBox1.Text}%'";
                    var connectionString = ConfigurationManager.ConnectionStrings["db"].ConnectionString;
                    using (SqlDataAdapter adapter = new SqlDataAdapter(query, connectionString))
                    {
                        System.Data.DataTable dataTable = new System.Data.DataTable();
                        adapter.Fill(dataTable);
                        dataGridView1.DataSource = dataTable;
                    }

                    for (int i = 0; i < this.dataGridView1.Rows.Count - 1; i++)
                    {
                        for (int j = 0; j < dataGridView1.Rows[i].Cells.Count; j++)
                        {
                            this.dataGridView1.Rows[i].Cells[j].Style.BackColor = UIColors.RedOverdue;
                            this.dataGridView1.Rows[i].Cells[j].Style.ForeColor = Color.White;
                        }
                    }
                }
                else
                {
                    RefreshTable3();
                }
            }

            if (rjRadioButton6.Checked)
            {
                if (!String.IsNullOrWhiteSpace(textBox1.Text))
                {
                    string query = $"SELECT Ticket.id, Clients.CompanyName AS [Клиент], Products.name AS [Продукт], Type.name AS [Тип проблемы], Ticket.header AS [Заголовок], Ticket.priority AS [Приоритет], Ticket.status AS [Статус], Ticket.application_data AS [Дата подачи], Ticket.completion_data AS [Дата закрытия] FROM Ticket, Clients, Products, Worker, Type WHERE Ticket.status != 'Закрыт' AND Ticket.client = Clients.id AND Ticket.product = Products.id AND Ticket.worker = Worker.id AND Ticket.type = Type.id AND Clients.CompanyName + ' ' + Products.name + ' ' + Type.name + ' ' + Ticket.header + ' ' + Ticket.description + ' ' + Ticket.priority + ' ' + Ticket.status LIKE '%{textBox1.Text}%' ORDER BY Ticket.completion_data ASC";
                    var connectionString = ConfigurationManager.ConnectionStrings["db"].ConnectionString;
                    using (SqlDataAdapter adapter = new SqlDataAdapter(query, connectionString))
                    {
                        System.Data.DataTable dataTable = new System.Data.DataTable();
                        adapter.Fill(dataTable);
                        dataGridView1.DataSource = dataTable;
                    }

                    for (int i = 0; i < this.dataGridView1.Rows.Count - 1; i++)
                    {
                        if (this.dataGridView1.Rows[i].Cells["Статус"].Value.ToString() == "Открыт")
                        {
                            for (int j = 0; j < dataGridView1.Rows[i].Cells.Count; j++)
                            {
                                this.dataGridView1.Rows[i].Cells[j].Style.BackColor = UIColors.SpecGreen;
                            }
                        }
                        if (this.dataGridView1.Rows[i].Cells["Статус"].Value.ToString() == "В работе")
                        {
                            for (int j = 0; j < dataGridView1.Rows[i].Cells.Count; j++)
                            {
                                this.dataGridView1.Rows[i].Cells[j].Style.BackColor = UIColors.SpecYellow;
                            }
                        }
                        if (this.dataGridView1.Rows[i].Cells["Статус"].Value.ToString() == "Закрыт")
                        {
                            for (int j = 0; j < dataGridView1.Rows[i].Cells.Count; j++)
                            {
                                this.dataGridView1.Rows[i].Cells[j].Style.BackColor = UIColors.SpecRed;
                            }
                        }
                    }
                }
                else
                {
                    RefreshTable6();
                }
            }
        }

        private void rjButton4_Click(object sender, EventArgs e)
        {
            int id = (int)dataGridView1.Rows[dataGridView1.CurrentCell.RowIndex].Cells[0].Value;
            ticket_view c = new ticket_view(id);
            c.Show();
        }

        private void dataGridView1_SelectionChanged(object sender, EventArgs e)
        {
            if ( dataGridView1.CurrentCell != null ) 
            {
                string query2 = $"SELECT Charge.id, (Worker.surname + ' ' + Worker.name + ' ' + Worker.patronymic) AS [Сотрудник], Charge.charge AS [Изменения], Charge.update_date AS [Дата изменения] FROM Worker, Charge, Ticket WHERE Charge.worker = Worker.id AND Charge.ticket = Ticket.id AND Ticket.id = '{dataGridView1.Rows[dataGridView1.CurrentCell.RowIndex].Cells[0].Value.ToString()}'";
                var connectionString2 = ConfigurationManager.ConnectionStrings["db"].ConnectionString;
                using (SqlDataAdapter adapter = new SqlDataAdapter(query2, connectionString2))
                {
                    System.Data.DataTable dataTable = new System.Data.DataTable();
                    adapter.Fill(dataTable);
                    dataGridView2.DataSource = dataTable;
                }
                dataGridView2.Columns[0].Visible = false;

                string query = $"SELECT ImageFiles.id, File2Tiket.id, ImageFiles.image AS [Изображение] FROM ImageFiles, File2Tiket WHERE ImageFiles.id = File2Tiket.photo AND File2Tiket.ticket = '{dataGridView1.Rows[dataGridView1.CurrentCell.RowIndex].Cells[0].Value.ToString()}'";
                using (SqlDataAdapter adapter = new SqlDataAdapter(query, connectionString2))
                {
                    System.Data.DataTable dataTable = new System.Data.DataTable();
                    adapter.Fill(dataTable);
                    dataGridView3.DataSource = dataTable;
                }
                dataGridView3.Columns[0].Visible = false;
                dataGridView3.Columns[1].Visible = false;
            }
        }

        private void rjButton8_Click(object sender, EventArgs e)
        {
            if (new charge_edit((int)dataGridView1.Rows[dataGridView1.CurrentCell.RowIndex].Cells[0].Value).ShowDialog() == DialogResult.OK)
            {
                var m_currentIndex = dataGridView1.CurrentCell.RowIndex;
                RefreshTable();
                rjRadioButton1.Checked = true;
                MessageBox.Show("Запись успешно добавлена!", "Успех!");

                if ((dataGridView1.Rows.Count - 1) < 0)
                {
                    dataGridView1.CurrentCell = dataGridView1.Rows[0].Cells[1];
                }
                else
                {
                    dataGridView1.CurrentCell = dataGridView1.Rows[m_currentIndex].Cells[1];
                    SubRefresh();

                    if ((dataGridView2.Rows.Count - 1) < 0)
                    {
                        dataGridView2.CurrentCell = dataGridView2.Rows[0].Cells[1];
                    }
                    else
                    {
                        dataGridView2.CurrentCell = dataGridView2.Rows[dataGridView2.Rows.Count - 1].Cells[1];
                    }
                }
            }
        }

        private void rjButton7_Click(object sender, EventArgs e)
        {
            charge_edit F2 = new charge_edit((int)dataGridView1.Rows[dataGridView1.CurrentCell.RowIndex].Cells[0].Value);
            F2.Text = "Редактирование изменения";
            F2.IDCHANGE = dataGridView2.Rows[dataGridView2.CurrentCell.RowIndex].Cells[0].Value.ToString();
            if (F2.ShowDialog() == DialogResult.OK)
            {
                var m_currentIndex = dataGridView1.CurrentCell.RowIndex;
                var s_currentIndex = dataGridView2.CurrentCell.RowIndex;
                RefreshTable();
                rjRadioButton1.Checked = true;
                MessageBox.Show("Запись успешно изменена!", "Успех!");

                if ((dataGridView1.Rows.Count - 1) < 0)
                {
                    dataGridView1.CurrentCell = dataGridView1.Rows[0].Cells[1];
                }
                else
                {
                    dataGridView1.CurrentCell = dataGridView1.Rows[m_currentIndex].Cells[1];
                    SubRefresh();

                    if ((dataGridView2.Rows.Count - 1) < 0)
                    {
                        dataGridView2.CurrentCell = dataGridView2.Rows[0].Cells[1];
                    }
                    else
                    {
                        dataGridView2.CurrentCell = dataGridView2.Rows[s_currentIndex].Cells[1];
                    }
                }
            }
        }

        private void rjButton6_Click(object sender, EventArgs e)
        {
            if (MessageBox.Show("Вы действительно хотите удалить запись?", "Подтверждение удаления", MessageBoxButtons.YesNo, MessageBoxIcon.Warning) == DialogResult.Yes)
            {
                try
                {
                    SqlCommand command = new SqlCommand($"DELETE FROM Charge WHERE id = {dataGridView2.Rows[dataGridView2.CurrentCell.RowIndex].Cells[0].Value}", sqlConnection);
                    command.ExecuteNonQuery();
                    MessageBox.Show("Запись успешно удалена!", "Успех!");
                    RefreshTable();
                    rjRadioButton1.Checked = true;
                }
                catch
                {
                    MessageBox.Show("Ошибка, попробуйте еще раз!", "Внимание", MessageBoxButtons.OK, MessageBoxIcon.Error);
                }
            }
        }

        private void rjRadioButton1_CheckedChanged(object sender, EventArgs e)
        {
            if ( rjRadioButton1.Checked == true ) 
            {
                RefreshTable();

                label6.Visible = false; label7.Visible = false; label8.Visible = false; label9.Visible = false;
                pictureBox6.Visible = false; pictureBox5.Visible = false;

                label5.Visible = false;
                label4.Visible = true; label3.Visible = true; label2.Visible = true; label1.Visible = true;
                pictureBox4.Visible = true; pictureBox3.Visible = true; pictureBox2.Visible = true; pictureBox1.Visible = true;

                label4.Text = "Выбранная запись";
                pictureBox4.BackColor = SystemColors.Highlight;

                label3.Text = "Статус - \"Открыт\"";
                pictureBox3.BackColor = UIColors.SpecGreen;

                label1.Text = "Статус - \"В работе\"";
                pictureBox1.BackColor = UIColors.SpecYellow;

                label2.Text = "Статус - \"Закрыт\"";
                pictureBox2.BackColor = UIColors.SpecRed;
            }
        }

        private void RefreshTable4()
        {
            //string query = "SELECT Ticket.id, (Clients.surname + ' ' + Clients.name + ' ' + Clients.patronymic) AS [Клиент], Products.name AS [Продукт], (Worker.surname + ' ' + Worker.name + ' ' + Worker.patronymic) AS [Принявший сотрудник], Type.name AS [Тип проблемы], Ticket.header AS [Заголовок], Ticket.description AS [Описание], Ticket.priority AS [Приоритет], Ticket.status AS [Статус], Ticket.application_data AS [Дата подачи], Ticket.completion_data AS [Дата закрытия], Ticket.actual_data AS [Фактическая дата закрытия] FROM Ticket, Clients, Products, Worker, Type WHERE Ticket.status != 'Закрыт' AND Ticket.client = Clients.id AND Ticket.product = Products.id AND Ticket.worker = Worker.id AND Ticket.type = Type.id";
            //string query = "SELECT Ticket.id, (Clients.surname + ' ' + Clients.name + ' ' + Clients.patronymic) AS [Клиент], Products.name AS [Продукт], (Worker.surname + ' ' + Worker.name + ' ' + Worker.patronymic) AS [Принявший сотрудник], Type.name AS [Тип проблемы], Ticket.header AS [Заголовок], Ticket.priority AS [Приоритет], Ticket.status AS [Статус], Ticket.application_data AS [Дата подачи], Ticket.completion_data AS [Дата закрытия], Ticket.actual_data AS [Фактическая дата закрытия] FROM Ticket, Clients, Products, Worker, Type WHERE Ticket.status != 'Закрыт' AND Ticket.client = Clients.id AND Ticket.product = Products.id AND Ticket.worker = Worker.id AND Ticket.type = Type.id";
            //string query = "SELECT Ticket.id, (Clients.surname + ' ' + Clients.name + ' ' + Clients.patronymic) AS [Клиент], Products.name AS [Продукт], Type.name AS [Тип проблемы], Ticket.header AS [Заголовок], Ticket.priority AS [Приоритет], Ticket.status AS [Статус], Ticket.application_data AS [Дата подачи], Ticket.completion_data AS [Дата закрытия], Ticket.actual_data AS [Фактическая дата закрытия] FROM Ticket, Clients, Products, Worker, Type WHERE Ticket.status != 'Закрыт' AND Ticket.client = Clients.id AND Ticket.product = Products.id AND Ticket.worker = Worker.id AND Ticket.type = Type.id";
            string query = "SELECT Ticket.id, Clients.CompanyName AS [Клиент], Products.name AS [Продукт], Type.name AS [Тип проблемы], Ticket.header AS [Заголовок], Ticket.priority AS [Приоритет], Ticket.status AS [Статус], Ticket.application_data AS [Дата подачи], Ticket.completion_data AS [Дата закрытия] FROM Ticket, Clients, Products, Worker, Type WHERE Ticket.status != 'Закрыт' AND Ticket.client = Clients.id AND Ticket.product = Products.id AND Ticket.worker = Worker.id AND Ticket.type = Type.id";
            var connectionString = ConfigurationManager.ConnectionStrings["db"].ConnectionString;
            using (SqlDataAdapter adapter = new SqlDataAdapter(query, connectionString))
            {
                System.Data.DataTable dataTable = new System.Data.DataTable();
                adapter.Fill(dataTable);
                dataGridView1.DataSource = dataTable;
            }
            dataGridView1.Columns[0].Visible = false;

            for (int i = 0; i < this.dataGridView1.Rows.Count - 1; i++)
            {
                if (this.dataGridView1.Rows[i].Cells["Статус"].Value.ToString() == "Открыт")
                {
                    for (int j = 0; j < dataGridView1.Rows[i].Cells.Count; j++)
                    {
                        this.dataGridView1.Rows[i].Cells[j].Style.BackColor = UIColors.SpecGreen;
                    }
                }
                if (this.dataGridView1.Rows[i].Cells["Статус"].Value.ToString() == "В работе")
                {
                    for (int j = 0; j < dataGridView1.Rows[i].Cells.Count; j++)
                    {
                        this.dataGridView1.Rows[i].Cells[j].Style.BackColor = UIColors.SpecYellow;
                    }
                }
                if (this.dataGridView1.Rows[i].Cells["Статус"].Value.ToString() == "Закрыт")
                {
                    for (int j = 0; j < dataGridView1.Rows[i].Cells.Count; j++)
                    {
                        this.dataGridView1.Rows[i].Cells[j].Style.BackColor = UIColors.SpecRed;
                    }
                }
            }
        }

        private void rjRadioButton4_CheckedChanged(object sender, EventArgs e)
        {
            if ( rjRadioButton4.Checked ) 
            {
                RefreshTable4();

                label6.Visible = false; label7.Visible = false; label8.Visible = false; label9.Visible = false;
                pictureBox6.Visible = false; pictureBox5.Visible = false;

                label5.Visible = false;
                label4.Visible = true; label3.Visible = false; label2.Visible = true; label1.Visible = true;
                pictureBox4.Visible = true; pictureBox3.Visible = false; pictureBox2.Visible = true; pictureBox1.Visible = true;

                label4.Text = "Выбранная запись";
                pictureBox4.BackColor = SystemColors.Highlight;

                label2.Text = "Статус - \"Открыт\"";
                pictureBox2.BackColor = UIColors.SpecGreen;

                label1.Text = "Статус - \"В работе\"";
                pictureBox1.BackColor = UIColors.SpecYellow;
            }
        }

        private void RefreshTable5()
        {
            //string query = "SELECT Ticket.id, (Clients.surname + ' ' + Clients.name + ' ' + Clients.patronymic) AS [Клиент], Products.name AS [Продукт], (Worker.surname + ' ' + Worker.name + ' ' + Worker.patronymic) AS [Принявший сотрудник], Type.name AS [Тип проблемы], Ticket.header AS [Заголовок], Ticket.description AS [Описание], Ticket.priority AS [Приоритет], Ticket.status AS [Статус], Ticket.application_data AS [Дата подачи], Ticket.completion_data AS [Дата закрытия], Ticket.actual_data AS [Фактическая дата закрытия] FROM Ticket, Clients, Products, Worker, Type WHERE Ticket.status = 'Закрыт' AND Ticket.client = Clients.id AND Ticket.product = Products.id AND Ticket.worker = Worker.id AND Ticket.type = Type.id";
            //string query = "SELECT Ticket.id, (Clients.surname + ' ' + Clients.name + ' ' + Clients.patronymic) AS [Клиент], Products.name AS [Продукт], (Worker.surname + ' ' + Worker.name + ' ' + Worker.patronymic) AS [Принявший сотрудник], Type.name AS [Тип проблемы], Ticket.header AS [Заголовок], Ticket.priority AS [Приоритет], Ticket.status AS [Статус], Ticket.application_data AS [Дата подачи], Ticket.completion_data AS [Дата закрытия], Ticket.actual_data AS [Фактическая дата закрытия] FROM Ticket, Clients, Products, Worker, Type WHERE Ticket.status = 'Закрыт' AND Ticket.client = Clients.id AND Ticket.product = Products.id AND Ticket.worker = Worker.id AND Ticket.type = Type.id";
            string query = "SELECT Ticket.id, Clients.CompanyName AS [Клиент], Products.name AS [Продукт], Type.name AS [Тип проблемы], Ticket.header AS [Заголовок], Ticket.priority AS [Приоритет], Ticket.status AS [Статус], Ticket.application_data AS [Дата подачи], Ticket.completion_data AS [Дата закрытия], Ticket.actual_data AS [Фактическая дата закрытия] FROM Ticket, Clients, Products, Worker, Type WHERE Ticket.status = 'Закрыт' AND Ticket.client = Clients.id AND Ticket.product = Products.id AND Ticket.worker = Worker.id AND Ticket.type = Type.id";
            var connectionString = ConfigurationManager.ConnectionStrings["db"].ConnectionString;
            using (SqlDataAdapter adapter = new SqlDataAdapter(query, connectionString))
            {
                System.Data.DataTable dataTable = new System.Data.DataTable();
                adapter.Fill(dataTable);
                dataGridView1.DataSource = dataTable;
            }
            dataGridView1.Columns[0].Visible = false;

            for (int i = 0; i < this.dataGridView1.Rows.Count - 1; i++)
            {
                if (this.dataGridView1.Rows[i].Cells["Статус"].Value.ToString() == "Закрыт")
                {
                    for (int j = 0; j < dataGridView1.Rows[i].Cells.Count; j++)
                    {
                        this.dataGridView1.Rows[i].Cells[j].Style.BackColor = UIColors.SpecRed;
                    }
                }
            }
        }

        private void rjRadioButton5_CheckedChanged(object sender, EventArgs e)
        {
            if (rjRadioButton5.Checked)
            {
                RefreshTable5();

                label6.Visible = false; label7.Visible = false; label8.Visible = false; label9.Visible = false;
                pictureBox6.Visible = false; pictureBox5.Visible = false;

                label5.Visible = false;
                label4.Visible = false; label3.Visible = false; label2.Visible = true; label1.Visible = true;
                pictureBox4.Visible = false; pictureBox3.Visible = false; pictureBox2.Visible = true; pictureBox1.Visible = true;

                label1.Text = "Выбранная запись";
                pictureBox1.BackColor = SystemColors.Highlight;

                label2.Text = "Статус - \"Закрыт\"";
                pictureBox2.BackColor = UIColors.SpecRed;
            }
        }

        private struct UIColors
        {
            //Статус - Открыт
            public readonly static Color SpecGreen = Color.FromArgb(39, 174, 96);
            //Статус - В работе
            public readonly static Color SpecYellow = Color.FromArgb(241, 196, 15);
            //Статус - Закрыт
            public readonly static Color SpecRed = Color.FromArgb(231, 76, 60);
            //Просроченная запись
            public readonly static Color RedOverdue = Color.FromArgb(235, 47, 6);
            //Приоритет - Низкий
            public readonly static Color PriorityLow = Color.FromArgb(9, 184, 62);
            //Приоритет - Средний
            public readonly static Color PriorityMedium = Color.FromArgb(255, 252, 0);
            //Приоритет - Высокий
            public readonly static Color PriorityHight = Color.FromArgb(223, 32, 41);
        }

        private void RefreshTable2()
        {
            //string query = "SELECT Ticket.id, (Clients.surname + ' ' + Clients.name + ' ' + Clients.patronymic) AS [Клиент], Products.name AS [Продукт], (Worker.surname + ' ' + Worker.name + ' ' + Worker.patronymic) AS [Принявший сотрудник], Type.name AS [Тип проблемы], Ticket.header AS [Заголовок], Ticket.description AS [Описание], Ticket.priority AS [Приоритет], Ticket.status AS [Статус], Ticket.application_data AS [Дата подачи], Ticket.completion_data AS [Дата закрытия], Ticket.actual_data AS [Фактическая дата закрытия] FROM Ticket, Clients, Products, Worker, Type WHERE Ticket.status != 'Закрыт' AND Ticket.client = Clients.id AND Ticket.product = Products.id AND Ticket.worker = Worker.id AND Ticket.type = Type.id";
            //string query = "SELECT Ticket.id, (Clients.surname + ' ' + Clients.name + ' ' + Clients.patronymic) AS [Клиент], Products.name AS [Продукт], (Worker.surname + ' ' + Worker.name + ' ' + Worker.patronymic) AS [Принявший сотрудник], Type.name AS [Тип проблемы], Ticket.header AS [Заголовок], Ticket.priority AS [Приоритет], Ticket.status AS [Статус], Ticket.application_data AS [Дата подачи], Ticket.completion_data AS [Дата закрытия], Ticket.actual_data AS [Фактическая дата закрытия] FROM Ticket, Clients, Products, Worker, Type WHERE Ticket.status != 'Закрыт' AND Ticket.client = Clients.id AND Ticket.product = Products.id AND Ticket.worker = Worker.id AND Ticket.type = Type.id";
            //string query = "SELECT Ticket.id, (Clients.surname + ' ' + Clients.name + ' ' + Clients.patronymic) AS [Клиент], Products.name AS [Продукт], Type.name AS [Тип проблемы], Ticket.header AS [Заголовок], Ticket.priority AS [Приоритет], Ticket.status AS [Статус], Ticket.application_data AS [Дата подачи], Ticket.completion_data AS [Дата закрытия], Ticket.actual_data AS [Фактическая дата закрытия] FROM Ticket, Clients, Products, Worker, Type WHERE Ticket.status != 'Закрыт' AND Ticket.client = Clients.id AND Ticket.product = Products.id AND Ticket.worker = Worker.id AND Ticket.type = Type.id";
            string query = "SELECT Ticket.id, Clients.CompanyName AS [Клиент], Products.name AS [Продукт], Type.name AS [Тип проблемы], Ticket.header AS [Заголовок], Ticket.priority AS [Приоритет], Ticket.status AS [Статус], Ticket.application_data AS [Дата подачи], Ticket.completion_data AS [Дата закрытия] FROM Ticket, Clients, Products, Worker, Type WHERE Ticket.status != 'Закрыт' AND Ticket.client = Clients.id AND Ticket.product = Products.id AND Ticket.worker = Worker.id AND Ticket.type = Type.id";
            var connectionString = ConfigurationManager.ConnectionStrings["db"].ConnectionString;
            using (SqlDataAdapter adapter = new SqlDataAdapter(query, connectionString))
            {
                System.Data.DataTable dataTable = new System.Data.DataTable();
                adapter.Fill(dataTable);
                dataGridView1.DataSource = dataTable;
            }

            for (int i = 0; i < this.dataGridView1.Rows.Count - 1; i++)
            {
                if (this.dataGridView1.Rows[i].Cells["Приоритет"].Value.ToString() == "Низкий")
                {
                    for (int j = 0; j < dataGridView1.Rows[i].Cells.Count; j++) 
                    {
                        this.dataGridView1.Rows[i].Cells[j].Style.BackColor = UIColors.PriorityLow;
                    }
                }
                if (this.dataGridView1.Rows[i].Cells["Приоритет"].Value.ToString() == "Средний")
                {
                    for (int j = 0; j < dataGridView1.Rows[i].Cells.Count; j++)
                    {
                        this.dataGridView1.Rows[i].Cells[j].Style.BackColor = UIColors.PriorityMedium;
                    }
                }
                if (this.dataGridView1.Rows[i].Cells["Приоритет"].Value.ToString() == "Высокий")
                {
                    for (int j = 0; j < dataGridView1.Rows[i].Cells.Count; j++)
                    {
                        this.dataGridView1.Rows[i].Cells[j].Style.BackColor = UIColors.PriorityHight;
                    }
                }
            }

            dataGridView1.Columns[0].Visible = false;
        }

        private void rjRadioButton2_CheckedChanged(object sender, EventArgs e)
        {
            if (rjRadioButton2.Checked)
            {
                RefreshTable2();

                label6.Visible = true; label7.Visible = true; label8.Visible = true; label9.Visible = true;
                pictureBox6.Visible = true; pictureBox5.Visible = true;

                label5.Visible = false;
                label4.Visible = false; label3.Visible = false; label2.Visible = false; label1.Visible = false;
                pictureBox4.Visible = false; pictureBox3.Visible = false;

                pictureBox2.Visible = true; pictureBox1.Visible = true;

                label9.Text = "Выбранная запись";
                pictureBox6.BackColor = SystemColors.Highlight;

                label6.Text = "Приоритет - \"Низкий\"";
                pictureBox5.BackColor = UIColors.PriorityLow;

                label7.Text = "Приоритет - \"Средний\"";
                pictureBox1.BackColor = UIColors.PriorityMedium;

                label8.Text = "Приоритет - \"Высокий\"";
                pictureBox2.BackColor = UIColors.PriorityHight;
            }
        }

        private void RefreshTable3()
        {
            //string query = "SELECT Ticket.id, (Clients.surname + ' ' + Clients.name + ' ' + Clients.patronymic) AS [Клиент], Products.name AS [Продукт], (Worker.surname + ' ' + Worker.name + ' ' + Worker.patronymic) AS [Принявший сотрудник], Type.name AS [Тип проблемы], Ticket.header AS [Заголовок], Ticket.description AS [Описание], Ticket.priority AS [Приоритет], Ticket.status AS [Статус], Ticket.application_data AS [Дата подачи], Ticket.completion_data AS [Дата закрытия], Ticket.actual_data AS [Фактическая дата закрытия] FROM Ticket, Clients, Products, Worker, Type WHERE Ticket.actual_data > Ticket.completion_data AND Ticket.client = Clients.id AND Ticket.product = Products.id AND Ticket.worker = Worker.id AND Ticket.type = Type.id";
            //string query = "SELECT Ticket.id, (Clients.surname + ' ' + Clients.name + ' ' + Clients.patronymic) AS [Клиент], Products.name AS [Продукт], (Worker.surname + ' ' + Worker.name + ' ' + Worker.patronymic) AS [Принявший сотрудник], Type.name AS [Тип проблемы], Ticket.header AS [Заголовок], Ticket.priority AS [Приоритет], Ticket.status AS [Статус], Ticket.application_data AS [Дата подачи], Ticket.completion_data AS [Дата закрытия], Ticket.actual_data AS [Фактическая дата закрытия] FROM Ticket, Clients, Products, Worker, Type WHERE Ticket.actual_data > Ticket.completion_data AND Ticket.client = Clients.id AND Ticket.product = Products.id AND Ticket.worker = Worker.id AND Ticket.type = Type.id";
            string query = $"SELECT Ticket.id, Clients.CompanyName AS [Клиент], Products.name AS [Продукт], Type.name AS [Тип проблемы], Ticket.header AS [Заголовок], Ticket.priority AS [Приоритет], Ticket.status AS [Статус], Ticket.application_data AS [Дата подачи], Ticket.completion_data AS [Дата закрытия] FROM Ticket, Clients, Products, Worker, Type WHERE Ticket.completion_data < CAST('{DateTime.Now.ToString()}' AS DATETIME) AND Ticket.status != 'Закрыт' AND Ticket.client = Clients.id AND Ticket.product = Products.id AND Ticket.worker = Worker.id AND Ticket.type = Type.id";
            var connectionString = ConfigurationManager.ConnectionStrings["db"].ConnectionString;
            using (SqlDataAdapter adapter = new SqlDataAdapter(query, connectionString))
            {
                System.Data.DataTable dataTable = new System.Data.DataTable();
                adapter.Fill(dataTable);
                dataGridView1.DataSource = dataTable;
            }

            for (int i = 0; i < this.dataGridView1.Rows.Count - 1; i++)
            {
                for (int j = 0; j < dataGridView1.Rows[i].Cells.Count; j++)
                {
                    this.dataGridView1.Rows[i].Cells[j].Style.BackColor = UIColors.RedOverdue;
                    this.dataGridView1.Rows[i].Cells[j].Style.ForeColor = Color.White;
                }
            }

            dataGridView1.Columns[0].Visible = false;
        }

        private void rjRadioButton3_CheckedChanged(object sender, EventArgs e)
        {
            if (rjRadioButton3.Checked)
            {
                RefreshTable3();

                label6.Visible = false; label7.Visible = false; label8.Visible = false; label9.Visible = false;
                pictureBox6.Visible = false; pictureBox5.Visible = false;

                label4.Visible = false; label3.Visible = false; label2.Visible = false; label1.Visible = true;
                pictureBox4.Visible = false; pictureBox3.Visible = false; pictureBox2.Visible = true; pictureBox1.Visible = true;

                label1.Text = "Выбранная запись";
                pictureBox1.BackColor = SystemColors.Highlight;

                label5.Visible = true;
                label5.Text = "Просроченная запись";
                pictureBox2.BackColor = UIColors.RedOverdue;
            }
        }

        private void ticket_form_Load(object sender, EventArgs e)
        {
            rjRadioButton1.Checked = true;
            if (rjRadioButton1.Checked == true)
            {
                RefreshTable();
            }
        }

        private void dataGridView1_Sorted(object sender, EventArgs e)
        {
            if (rjRadioButton1.Checked == true)
            {
                for (int i = 0; i < this.dataGridView1.Rows.Count - 1; i++)
                {
                    if (this.dataGridView1.Rows[i].Cells["Статус"].Value.ToString() == "Открыт")
                    {
                        for (int j = 0; j < dataGridView1.Rows[i].Cells.Count; j++)
                        {
                            this.dataGridView1.Rows[i].Cells[j].Style.BackColor = UIColors.SpecGreen;
                        }
                    }
                    if (this.dataGridView1.Rows[i].Cells["Статус"].Value.ToString() == "В работе")
                    {
                        for (int j = 0; j < dataGridView1.Rows[i].Cells.Count; j++)
                        {
                            this.dataGridView1.Rows[i].Cells[j].Style.BackColor = UIColors.SpecYellow;
                        }
                    }
                    if (this.dataGridView1.Rows[i].Cells["Статус"].Value.ToString() == "Закрыт")
                    {
                        for (int j = 0; j < dataGridView1.Rows[i].Cells.Count; j++)
                        {
                            this.dataGridView1.Rows[i].Cells[j].Style.BackColor = UIColors.SpecRed;
                        }
                    }
                }
            }

            if (rjRadioButton4.Checked == true)
            {
                for (int i = 0; i < this.dataGridView1.Rows.Count - 1; i++)
                {
                    if (this.dataGridView1.Rows[i].Cells["Статус"].Value.ToString() == "Открыт")
                    {
                        for (int j = 0; j < dataGridView1.Rows[i].Cells.Count; j++)
                        {
                            this.dataGridView1.Rows[i].Cells[j].Style.BackColor = UIColors.SpecGreen;
                        }
                    }
                    if (this.dataGridView1.Rows[i].Cells["Статус"].Value.ToString() == "В работе")
                    {
                        for (int j = 0; j < dataGridView1.Rows[i].Cells.Count; j++)
                        {
                            this.dataGridView1.Rows[i].Cells[j].Style.BackColor = UIColors.SpecYellow;
                        }
                    }
                    if (this.dataGridView1.Rows[i].Cells["Статус"].Value.ToString() == "Закрыт")
                    {
                        for (int j = 0; j < dataGridView1.Rows[i].Cells.Count; j++)
                        {
                            this.dataGridView1.Rows[i].Cells[j].Style.BackColor = UIColors.SpecRed;
                        }
                    }
                }
            }

            if (rjRadioButton5.Checked == true)
            {
                for (int i = 0; i < this.dataGridView1.Rows.Count - 1; i++)
                {
                    if (this.dataGridView1.Rows[i].Cells["Статус"].Value.ToString() == "Закрыт")
                    {
                        for (int j = 0; j < dataGridView1.Rows[i].Cells.Count; j++)
                        {
                            this.dataGridView1.Rows[i].Cells[j].Style.BackColor = UIColors.SpecRed;
                        }
                    }
                }
            }

            if (rjRadioButton2.Checked == true)
            {
                for (int i = 0; i < this.dataGridView1.Rows.Count - 1; i++)
                {
                    if (this.dataGridView1.Rows[i].Cells["Приоритет"].Value.ToString() == "Низкий")
                    {
                        for (int j = 0; j < dataGridView1.Rows[i].Cells.Count; j++)
                        {
                            this.dataGridView1.Rows[i].Cells[j].Style.BackColor = UIColors.PriorityLow;
                        }
                    }
                    if (this.dataGridView1.Rows[i].Cells["Приоритет"].Value.ToString() == "Средний")
                    {
                        for (int j = 0; j < dataGridView1.Rows[i].Cells.Count; j++)
                        {
                            this.dataGridView1.Rows[i].Cells[j].Style.BackColor = UIColors.PriorityMedium;
                        }
                    }
                    if (this.dataGridView1.Rows[i].Cells["Приоритет"].Value.ToString() == "Высокий")
                    {
                        for (int j = 0; j < dataGridView1.Rows[i].Cells.Count; j++)
                        {
                            this.dataGridView1.Rows[i].Cells[j].Style.BackColor = UIColors.PriorityHight;
                        }
                    }
                }
            }

            if (rjRadioButton3.Checked == true)
            {
                for (int i = 0; i < this.dataGridView1.Rows.Count - 1; i++)
                {
                    for (int j = 0; j < dataGridView1.Rows[i].Cells.Count; j++)
                    {
                        this.dataGridView1.Rows[i].Cells[j].Style.BackColor = UIColors.RedOverdue;
                        this.dataGridView1.Rows[i].Cells[j].Style.ForeColor = Color.White;
                    }
                }
            }

            if (rjRadioButton6.Checked == true) 
            {
                for (int i = 0; i < this.dataGridView1.Rows.Count - 1; i++)
                {
                    if (this.dataGridView1.Rows[i].Cells["Статус"].Value.ToString() == "Открыт")
                    {
                        for (int j = 0; j < dataGridView1.Rows[i].Cells.Count; j++)
                        {
                            this.dataGridView1.Rows[i].Cells[j].Style.BackColor = UIColors.SpecGreen;
                        }
                    }
                    if (this.dataGridView1.Rows[i].Cells["Статус"].Value.ToString() == "В работе")
                    {
                        for (int j = 0; j < dataGridView1.Rows[i].Cells.Count; j++)
                        {
                            this.dataGridView1.Rows[i].Cells[j].Style.BackColor = UIColors.SpecYellow;
                        }
                    }
                    if (this.dataGridView1.Rows[i].Cells["Статус"].Value.ToString() == "Закрыт")
                    {
                        for (int j = 0; j < dataGridView1.Rows[i].Cells.Count; j++)
                        {
                            this.dataGridView1.Rows[i].Cells[j].Style.BackColor = UIColors.SpecRed;
                        }
                    }
                }
            }
        }

        private void rjButton12_Click(object sender, EventArgs e)
        {
            if (dataGridView3.CurrentCell != null ) 
            {
                int id = (int)dataGridView3.Rows[dataGridView3.CurrentCell.RowIndex].Cells[0].Value;
                photo_view c = new photo_view(id);
                c.Show();
            }
        }

        private void rjButton9_Click(object sender, EventArgs e)
        {
            if (MessageBox.Show("Вы действительно хотите удалить запись?", "Подтверждение удаления", MessageBoxButtons.YesNo, MessageBoxIcon.Warning) == DialogResult.Yes)
            {
                int File2Ticket_id = (int)dataGridView3.Rows[dataGridView3.CurrentCell.RowIndex].Cells[1].Value;
                int ImageFile_id = (int)dataGridView3.Rows[dataGridView3.CurrentCell.RowIndex].Cells[0].Value;
                try
                {
                    SqlCommand command = new SqlCommand($"DELETE FROM File2Tiket WHERE id = {File2Ticket_id}", sqlConnection);
                    command.ExecuteNonQuery();
                }
                catch
                {
                    MessageBox.Show("Ошибка, попробуйте еще раз!", "Внимание", MessageBoxButtons.OK, MessageBoxIcon.Error);
                }
                try
                {
                    SqlCommand command2 = new SqlCommand($"DELETE FROM ImageFiles WHERE id = {ImageFile_id}", sqlConnection);
                    command2.ExecuteNonQuery();

                    MessageBox.Show("Запись успешно удалена!", "Успех!");
                    RefreshTable();
                    rjRadioButton1.Checked = true;
                }
                catch
                {
                    MessageBox.Show("Ошибка, попробуйте еще раз!", "Внимание", MessageBoxButtons.OK, MessageBoxIcon.Error);
                }

            }
        }

        private void rjButton11_Click(object sender, EventArgs e)
        {
            if (new file_edit((int)dataGridView1.Rows[dataGridView1.CurrentCell.RowIndex].Cells[0].Value).ShowDialog() == DialogResult.OK)
            {
                var m_currentIndex = dataGridView1.CurrentCell.RowIndex;
                RefreshTable();
                rjRadioButton1.Checked = true;
                MessageBox.Show("Продукт успешно добавлен!", "Успех!");

                if ((dataGridView1.Rows.Count - 1) < 0)
                {
                    dataGridView1.CurrentCell = dataGridView1.Rows[0].Cells[1];
                }
                else
                {
                    dataGridView1.CurrentCell = dataGridView1.Rows[m_currentIndex].Cells[1];
                    SubRefresh();

                    if ((dataGridView3.Rows.Count - 1) < 0)
                    {
                        dataGridView3.CurrentCell = dataGridView3.Rows[0].Cells[2];
                    }
                    else
                    {
                        dataGridView3.CurrentCell = dataGridView3.Rows[dataGridView3.Rows.Count - 1].Cells[2];
                    }
                }
            }
        }

        private void rjButton10_Click(object sender, EventArgs e)
        {
            file_edit F2 = new file_edit((int)dataGridView1.Rows[dataGridView1.CurrentCell.RowIndex].Cells[0].Value);
            F2.Text = "Редактирование изображения";
            F2.IDCHANGE = dataGridView3.Rows[dataGridView3.CurrentCell.RowIndex].Cells[0].Value.ToString();
            if (F2.ShowDialog() == DialogResult.OK)
            {
                var m_currentIndex = dataGridView1.CurrentCell.RowIndex;
                var s_currentIndex = dataGridView3.CurrentCell.RowIndex;
                RefreshTable();
                rjRadioButton1.Checked = true;
                MessageBox.Show("Запись успешно изменена!", "Успех!");

                if ((dataGridView1.Rows.Count - 1) < 0)
                {
                    dataGridView1.CurrentCell = dataGridView1.Rows[0].Cells[1];
                }
                else
                {
                    dataGridView1.CurrentCell = dataGridView1.Rows[m_currentIndex].Cells[1];
                    SubRefresh();

                    if ((dataGridView3.Rows.Count - 1) < 0)
                    {
                        dataGridView3.CurrentCell = dataGridView3.Rows[0].Cells[2];
                    }
                    else
                    {
                        dataGridView3.CurrentCell = dataGridView3.Rows[s_currentIndex].Cells[2];
                    }
                }
            }
        }

        private void RefreshTable6()
        {
            string query = "SELECT Ticket.id, Clients.CompanyName AS [Клиент], Products.name AS [Продукт], Type.name AS [Тип проблемы], Ticket.header AS [Заголовок], Ticket.priority AS [Приоритет], Ticket.status AS [Статус], Ticket.application_data AS [Дата подачи], Ticket.completion_data AS [Дата закрытия] FROM Ticket, Clients, Products, Worker, Type WHERE Ticket.status != 'Закрыт' AND Ticket.client = Clients.id AND Ticket.product = Products.id AND Ticket.worker = Worker.id AND Ticket.type = Type.id ORDER BY Ticket.completion_data ASC";
            var connectionString = ConfigurationManager.ConnectionStrings["db"].ConnectionString;
            using (SqlDataAdapter adapter = new SqlDataAdapter(query, connectionString))
            {
                System.Data.DataTable dataTable = new System.Data.DataTable();
                adapter.Fill(dataTable);
                dataGridView1.DataSource = dataTable;
            }
            dataGridView1.Columns[0].Visible = false;

            for (int i = 0; i < this.dataGridView1.Rows.Count - 1; i++)
            {
                if (this.dataGridView1.Rows[i].Cells["Статус"].Value.ToString() == "Открыт")
                {
                    for (int j = 0; j < dataGridView1.Rows[i].Cells.Count; j++)
                    {
                        this.dataGridView1.Rows[i].Cells[j].Style.BackColor = UIColors.SpecGreen;
                    }
                }
                if (this.dataGridView1.Rows[i].Cells["Статус"].Value.ToString() == "В работе")
                {
                    for (int j = 0; j < dataGridView1.Rows[i].Cells.Count; j++)
                    {
                        this.dataGridView1.Rows[i].Cells[j].Style.BackColor = UIColors.SpecYellow;
                    }
                }
                if (this.dataGridView1.Rows[i].Cells["Статус"].Value.ToString() == "Закрыт")
                {
                    for (int j = 0; j < dataGridView1.Rows[i].Cells.Count; j++)
                    {
                        this.dataGridView1.Rows[i].Cells[j].Style.BackColor = UIColors.SpecRed;
                    }
                }
            }
        }

        private void rjRadioButton6_CheckedChanged(object sender, EventArgs e)
        {
            if (rjRadioButton6.Checked)
            {
                RefreshTable6();
                
                label6.Visible = false; label7.Visible = false; label8.Visible = false; label9.Visible = false;
                pictureBox6.Visible = false; pictureBox5.Visible = false;
                
                label5.Visible = false;
                label4.Visible = true; label3.Visible = false; label2.Visible = true; label1.Visible = true;
                pictureBox4.Visible = true; pictureBox3.Visible = false; pictureBox2.Visible = true; pictureBox1.Visible = true;
                
                label4.Text = "Выбранная запись";
                pictureBox4.BackColor = SystemColors.Highlight;
                
                label2.Text = "Статус - \"Открыт\"";
                pictureBox2.BackColor = UIColors.SpecGreen;
                
                label1.Text = "Статус - \"В работе\"";
                pictureBox1.BackColor = UIColors.SpecYellow;
            }
        }

        private void SubRefresh() 
        {
            if (dataGridView1.CurrentCell != null)
            {
                string query2 = $"SELECT Charge.id, (Worker.surname + ' ' + Worker.name + ' ' + Worker.patronymic) AS [Сотрудник], Charge.charge AS [Изменения], Charge.update_date AS [Дата изменения] FROM Worker, Charge, Ticket WHERE Charge.worker = Worker.id AND Charge.ticket = Ticket.id AND Ticket.id = '{dataGridView1.Rows[dataGridView1.CurrentCell.RowIndex].Cells[0].Value.ToString()}'";
                var connectionString2 = ConfigurationManager.ConnectionStrings["db"].ConnectionString;
                using (SqlDataAdapter adapter = new SqlDataAdapter(query2, connectionString2))
                {
                    System.Data.DataTable dataTable = new System.Data.DataTable();
                    adapter.Fill(dataTable);
                    dataGridView2.DataSource = dataTable;
                }
                dataGridView2.Columns[0].Visible = false;

                string query = $"SELECT ImageFiles.id, File2Tiket.id, ImageFiles.image AS [Изображение] FROM ImageFiles, File2Tiket WHERE ImageFiles.id = File2Tiket.photo AND File2Tiket.ticket = '{dataGridView1.Rows[dataGridView1.CurrentCell.RowIndex].Cells[0].Value.ToString()}'";
                using (SqlDataAdapter adapter = new SqlDataAdapter(query, connectionString2))
                {
                    System.Data.DataTable dataTable = new System.Data.DataTable();
                    adapter.Fill(dataTable);
                    dataGridView3.DataSource = dataTable;
                }
                dataGridView3.Columns[0].Visible = false;
                dataGridView3.Columns[1].Visible = false;
            }
        }

        private void rjButton13_Click(object sender, EventArgs e)
        {
            Analytic analytic = new Analytic();
            analytic.ShowDialog();
        }
    }
}