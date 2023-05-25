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

namespace techSupport.edit_form
{
    public partial class client2prod_edit : Form
    {
        private SqlConnection sqlConnection = null;

        public client2prod_edit()
        {
            sqlConnection = new SqlConnection(ConfigurationManager.ConnectionStrings["db"].ConnectionString);
            sqlConnection.Open();
            InitializeComponent();
            combobox(comboBox1, "SELECT id, ('#' + CAST(id AS nvarchar) + ' | ' + surname + ' ' + name + ' ' + patronymic) AS [FIO] FROM [Clients]", "FIO", "id");
            combobox(comboBox2, "SELECT id, name FROM [Products]", "name", "id");
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
            string query = $"SELECT * FROM User2Product WHERE id = {id}";
            var connectionString = ConfigurationManager.ConnectionStrings["db"].ConnectionString;
            using (SqlDataAdapter adapter = new SqlDataAdapter(query, connectionString))
            {
                DataTable dataTable = new DataTable();
                adapter.Fill(dataTable);
                comboBox1.SelectedValue = (int)dataTable.Rows[0][1];
                comboBox2.SelectedValue = (int)dataTable.Rows[0][2];
            }
        }

        private void button2_Click(object sender, EventArgs e)
        {
            if (String.IsNullOrWhiteSpace(comboBox2.Text) ||
                String.IsNullOrWhiteSpace(comboBox1.Text))
                MessageBox.Show("Необходимо заполнить все данные!", "Ошибка!");
            else
            {
                if (!isChange)
                {
                    string query = "INSERT INTO User2Product (client, product)" +
                    "VALUES(@client, @product)";
                    using (SqlCommand command = new SqlCommand(query, sqlConnection))
                    {
                        command.Parameters.AddWithValue("@client", comboBox1.SelectedValue);
                        command.Parameters.AddWithValue("@product", comboBox2.SelectedValue);
                        command.ExecuteNonQuery();
                        this.DialogResult = DialogResult.OK;
                    }
                }
                else
                {
                    string query = $"UPDATE User2Product SET client = @client, product = @product WHERE id = {idChange}";
                    using (SqlCommand command = new SqlCommand(query, sqlConnection))
                    {
                        command.Parameters.AddWithValue("@client", comboBox1.SelectedValue);
                        command.Parameters.AddWithValue("@product", comboBox2.SelectedValue);
                        command.ExecuteNonQuery();
                        this.DialogResult = DialogResult.OK;
                    }
                }
            }
        }

        private void iconPictureBox2_Click(object sender, EventArgs e)
        {
            if (new client_edit().ShowDialog() == DialogResult.OK)
            {
                comboBox1.DataSource = null;
                combobox(comboBox1, "SELECT id, ('#' + CAST(id AS nvarchar) + ' | ' + surname + ' ' + name + ' ' + patronymic) AS [FIO] FROM [Clients]", "FIO", "id");
                MessageBox.Show("Запись успешно добавлена!", "Успех!");
            }
        }

        private void iconPictureBox1_Click(object sender, EventArgs e)
        {
            if (new products_edit().ShowDialog() == DialogResult.OK)
            {
                comboBox2.DataSource = null;
                combobox(comboBox2, "SELECT id, name FROM [Products]", "name", "id");
                MessageBox.Show("Запись успешно добавлена!", "Успех!");
            }
        }
    }
}