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
using techSupport.forms;
using techSupport.new_forms;

namespace techSupport.Ticket_system
{
    public partial class ticket_edit : Form
    {
        private SqlConnection sqlConnection = null;

        private int UserId;

        private void setUserId(int id) { UserId = id; }

        public ticket_edit(int id)
        {
            setUserId(id);
            sqlConnection = new SqlConnection(ConfigurationManager.ConnectionStrings["db"].ConnectionString);
            sqlConnection.Open();
            InitializeComponent();
            combobox(comboBox1, "SELECT id, (surname + ' ' + name + ' ' + patronymic) AS [FIO] FROM [Clients]", "FIO", "id");
            combobox(comboBox3, "SELECT id, (surname + ' ' + name + ' ' + patronymic) AS [FIO] FROM [Worker]", "FIO", "id");
            combobox(comboBox4, "SELECT id, name FROM [Type]", "name", "id");
            comboBox3.Enabled = false;
            comboBox3.SelectedValue = UserId;
            comboBox6.Enabled = false;
            comboBox6.Text = "Открыт";
            dateTimePicker1.Value = DateTime.Now;
            dateTimePicker2.Value = DateTime.Now;
            dateTimePicker3.Value = DateTime.Now;
        }

        private bool isChange = false;
        private string idChange;

        public string IDCHANGE
        {
            set
            {
                idChange = value;
                isChange = true;
                FillBoxes(value);
            }
        }

        private void combobox(ComboBox c, string query, string displaymember, string valuemember)
        {
            var connectionString = ConfigurationManager.ConnectionStrings["db"].ConnectionString;
            using (SqlDataAdapter adapter = new SqlDataAdapter(query, connectionString))
            {
                System.Data.DataTable datatable = new System.Data.DataTable();
                adapter.Fill(datatable);
                c.DataSource = datatable;
                c.DisplayMember = displaymember;
                c.ValueMember = valuemember;
                c.SelectedIndex = -1;
            }
        }

        private void FillBoxes(string id)
        {
            comboBox3.Enabled = true;
            comboBox3.SelectedValue = UserId;
            comboBox6.Enabled = true;

            string query = $"SELECT * FROM Ticket WHERE id = {id}";
            var connectionString = ConfigurationManager.ConnectionStrings["db"].ConnectionString;
            using (SqlDataAdapter adapter = new SqlDataAdapter(query, connectionString))
            {
                DataTable dataTable = new DataTable();
                adapter.Fill(dataTable);

                comboBox1.SelectedValue = (int)dataTable.Rows[0][1];
                comboBox2.SelectedValue = (int)dataTable.Rows[0][2];
                comboBox3.SelectedValue = (int)dataTable.Rows[0][3];
                comboBox4.SelectedValue = (int)dataTable.Rows[0][4];

                richTextBox1.Text = dataTable.Rows[0][5].ToString();
                richTextBox2.Text = dataTable.Rows[0][6].ToString();
                comboBox5.Text = dataTable.Rows[0][7].ToString();
                comboBox6.Text = dataTable.Rows[0][8].ToString();

                dateTimePicker1.Value = (DateTime)dataTable.Rows[0][9];
                dateTimePicker2.Value = (DateTime)dataTable.Rows[0][10];
                dateTimePicker3.Value = (DateTime)dataTable.Rows[0][11];
            }
        }

        private void button2_Click(object sender, EventArgs e)
        {
            if (String.IsNullOrWhiteSpace(richTextBox1.Text) || String.IsNullOrWhiteSpace(richTextBox2.Text) || String.IsNullOrWhiteSpace(comboBox5.Text) || String.IsNullOrWhiteSpace(comboBox1.Text) || String.IsNullOrWhiteSpace(comboBox1.Text) || String.IsNullOrWhiteSpace(comboBox2.Text) || String.IsNullOrWhiteSpace(comboBox3.Text) || String.IsNullOrWhiteSpace(comboBox4.Text) || String.IsNullOrWhiteSpace(comboBox6.Text))
                MessageBox.Show("Необходимо заполнить все данные!", "Ошибка!");
            else
            {
                if (!isChange)
                {
                    string query = "INSERT INTO Ticket (client, product, worker, type, header, description, priority, status, application_data, completion_data, actual_data)" +
                    "VALUES(@client, @product, @worker, @type, @header, @description, @priority, @status, @application_data, @completion_data, @actual_data)";
                    using (SqlCommand command = new SqlCommand(query, sqlConnection))
                    {
                        command.Parameters.AddWithValue("@client", comboBox1.SelectedValue);
                        command.Parameters.AddWithValue("@product", comboBox2.SelectedValue);
                        command.Parameters.AddWithValue("@worker", UserId);
                        command.Parameters.AddWithValue("@type", comboBox4.SelectedValue);
                        command.Parameters.AddWithValue("@header", richTextBox1.Text);
                        command.Parameters.AddWithValue("@description", richTextBox2.Text);
                        command.Parameters.AddWithValue("@priority", comboBox5.Text);
                        command.Parameters.AddWithValue("@status", comboBox6.Text);
                        command.Parameters.AddWithValue("@application_data", dateTimePicker1.Value);
                        command.Parameters.AddWithValue("@completion_data", dateTimePicker2.Value);
                        command.Parameters.AddWithValue("@actual_data", dateTimePicker3.Value);
                        command.ExecuteNonQuery();
                        this.DialogResult = DialogResult.OK;
                    }
                }
                else
                {
                    string query = $"UPDATE Ticket SET client = @client, product = @product, worker = @worker, type = @type, header = @header, description = @description, priority = @priority, status = @status, application_data = @application_data, completion_data = @completion_data, actual_data = @actual_data WHERE id = {idChange}";
                    using (SqlCommand command = new SqlCommand(query, sqlConnection))
                    {
                        command.Parameters.AddWithValue("@client", comboBox1.SelectedValue);
                        command.Parameters.AddWithValue("@product", comboBox2.SelectedValue);
                        command.Parameters.AddWithValue("@worker", comboBox3.SelectedValue);
                        command.Parameters.AddWithValue("@type", comboBox4.SelectedValue);
                        command.Parameters.AddWithValue("@header", richTextBox1.Text);
                        command.Parameters.AddWithValue("@description", richTextBox2.Text);
                        command.Parameters.AddWithValue("@priority", comboBox5.Text);
                        command.Parameters.AddWithValue("@status", comboBox6.Text);
                        command.Parameters.AddWithValue("@application_data", dateTimePicker1.Value);
                        command.Parameters.AddWithValue("@completion_data", dateTimePicker2.Value);
                        command.Parameters.AddWithValue("@actual_data", dateTimePicker3.Value);
                        command.ExecuteNonQuery();
                        this.DialogResult = DialogResult.OK;
                    }
                }
            }
        }

        private void comboBox1_SelectedIndexChanged(object sender, EventArgs e)
        {
            comboBox2.DataSource = null;
            if (comboBox1.SelectedIndex != 0) 
            {
                if (comboBox1.SelectedValue != null) 
                {
                    int m_id = (int)comboBox1.SelectedValue;
                    combobox(comboBox2, $"SELECT Products.id AS [n1], Products.name AS [n2] FROM [Products], [User2Product], [Clients] WHERE Products.id = User2Product.product AND Clients.id = User2Product.client AND User2Product.client = {m_id}", "n2", "n1");
                }
            }
        }

        private void iconPictureBox1_Click(object sender, EventArgs e)
        {
            if (new client_edit().ShowDialog() == DialogResult.OK)
            {
                comboBox1.DataSource = null;
                combobox(comboBox1, "SELECT id, (surname + ' ' + name + ' ' + patronymic) AS [FIO] FROM [Clients]", "FIO", "id");
                MessageBox.Show("Запись успешно добавлена!", "Успех!");
            }
        }

        private void iconPictureBox3_Click(object sender, EventArgs e)
        {
            if (new worker_edit().ShowDialog() == DialogResult.OK)
            {
                comboBox3.DataSource = null;
                combobox(comboBox3, "SELECT id, (surname + ' ' + name + ' ' + patronymic) AS [FIO] FROM [Worker]", "FIO", "id");
                MessageBox.Show("Запись успешно добавлена!", "Успех!");
            }
        }

        private void iconPictureBox4_Click(object sender, EventArgs e)
        {
            if (new type_edit().ShowDialog() == DialogResult.OK)
            {
                comboBox4.DataSource = null;
                combobox(comboBox4, "SELECT id, name FROM [Type]", "name", "id");
                MessageBox.Show("Запись успешно добавлена!", "Успех!");
            }
        }

        private void iconPictureBox2_Click(object sender, EventArgs e)
        {
            if (new products_form().ShowDialog() == DialogResult.OK)
            {
                comboBox2.DataSource = null;
                if (comboBox1.SelectedIndex != 0)
                {
                    if (comboBox1.SelectedValue != null)
                    {
                        int m_id = (int)comboBox1.SelectedValue;
                        combobox(comboBox2, $"SELECT Products.id AS [n1], Products.name AS [n2] FROM [Products], [User2Product], [Clients] WHERE Products.id = User2Product.product AND Clients.id = User2Product.client AND User2Product.client = {m_id}", "n2", "n1");
                    }
                }
                MessageBox.Show("Запись успешно добавлена!", "Успех!");
            }
        }
    }
}