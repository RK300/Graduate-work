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
using DataTable = System.Data.DataTable;

namespace techSupport.view_form
{
    public partial class ticket_view : Form
    {
        private SqlConnection sqlConnection = null;

        private int m_id;

        private void SetId(int id) { m_id = id; }

        public ticket_view(int id)
        {
            sqlConnection = new SqlConnection(ConfigurationManager.ConnectionStrings["db"].ConnectionString);
            sqlConnection.Open();
            SetId(id);
            InitializeComponent();
            setName();
            setText();
            setInfo();
            setPhoto();
            setHistory();
        }

        private void setHistory()
        {
            string query = $"SELECT (Worker.surname + ' ' + Worker.name + ' ' + Worker.patronymic) AS [Сотрудник], Charge.charge AS [Изменения], Charge.update_date AS [Дата изменения] FROM Worker, Charge, Ticket WHERE Charge.worker = Worker.id AND Charge.ticket = Ticket.id AND Ticket.id = '{m_id}'";
            var connectionString = ConfigurationManager.ConnectionStrings["db"].ConnectionString;
            using (SqlDataAdapter adapter = new SqlDataAdapter(query, connectionString))
            {
                DataTable dataTable = new DataTable();
                adapter.Fill(dataTable);
                dataGridView2.DataSource = dataTable;
            }
        }

        private void setPhoto() 
        {
            string query = $"SELECT ImageFiles.id, ImageFiles.image FROM ImageFiles, File2Tiket WHERE ImageFiles.id = File2Tiket.photo AND File2Tiket.ticket = '{m_id}'";
            var connectionString = ConfigurationManager.ConnectionStrings["db"].ConnectionString;
            using (SqlDataAdapter adapter = new SqlDataAdapter(query, connectionString))
            {
                DataTable dataTable = new DataTable();
                adapter.Fill(dataTable);
                dataGridView1.DataSource = dataTable;
            }
            dataGridView1.Columns[0].Visible = false;
        }

        private void setInfo()
        {
            string query = $"SELECT (Clients.CompanyName + ' | ' + Clients.surname + ' ' + Clients.name + ' ' + Clients.patronymic), Products.name, (Worker.surname + ' ' + Worker.name + ' ' + Worker.patronymic), Type.name, Ticket.application_data, Ticket.completion_data, Ticket.actual_data, Ticket.status, Ticket.priority FROM Ticket, Clients, Products, Worker, Type WHERE Ticket.client = Clients.id AND Ticket.product = Products.id AND Ticket.worker = Worker.id AND Ticket.type = Type.id AND Ticket.id = '{m_id}'";
            SqlDataAdapter adapter = new SqlDataAdapter();
            SqlCommand command = new SqlCommand(query, sqlConnection);
            DataTable tb = new DataTable();
            adapter.SelectCommand = command;
            adapter.Fill(tb);
            label4.Text = "Клиент: " + tb.Rows[0][0].ToString();
            label3.Text = "Продукт: " + tb.Rows[0][1].ToString();
            label1.Text = "Принявший сотрудник: " + tb.Rows[0][2].ToString();
            label2.Text = "Тип: " + tb.Rows[0][3].ToString();

            DateTime dateTime = DateTime.Now;
            dateTime = (DateTime)tb.Rows[0][4];
            label5.Text = "Дата подачи: " + dateTime.ToString("D");

            dateTime = (DateTime)tb.Rows[0][5];
            label6.Text = "Дата сдачи: " + dateTime.ToString("D");

            label13.Text = "Статус: " + tb.Rows[0][7].ToString();
            if (tb.Rows[0][7].ToString() == "Закрыт" ) 
            {
                dateTime = (DateTime)tb.Rows[0][6];
                label7.Text = "Фактическая дата сдачи: " + dateTime.ToString("D");
            }
            else
            {
                label7.Visible = false;
            }
            label12.Text = "Приоритет: " + tb.Rows[0][8].ToString();
        }

        private void setName()
        {
            string query = $"SELECT ('#' + CAST(Ticket.id AS nvarchar) + ' | ' + Clients.CompanyName) FROM Ticket, Clients WHERE Ticket.client = Clients.id AND Ticket.id = '{m_id}'";
            SqlDataAdapter adapter = new SqlDataAdapter();
            SqlCommand command = new SqlCommand(query, sqlConnection);
            DataTable tb = new DataTable();
            adapter.SelectCommand = command;
            adapter.Fill(tb);
            this.Text = "Просмотр тикета:: " + tb.Rows[0][0].ToString();
        }

        private void setText()
        {
            string query = $"SELECT header, description FROM Ticket WHERE id = '{m_id}'";
            SqlDataAdapter adapter = new SqlDataAdapter();
            SqlCommand command = new SqlCommand(query, sqlConnection);
            DataTable tb = new DataTable();
            adapter.SelectCommand = command;
            adapter.Fill(tb);
            richTextBox2.Text = tb.Rows[0][0].ToString();
            richTextBox1.Text = tb.Rows[0][1].ToString();
        }

        private void rjButton4_Click(object sender, EventArgs e)
        {
            int id = (int)dataGridView1.Rows[dataGridView1.CurrentCell.RowIndex].Cells[0].Value;
            photo_view c = new photo_view(id);
            c.Show();
        }
    }
}
