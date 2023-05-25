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
using techSupport.new_forms;

namespace techSupport.forms
{
    public partial class products_form : Form
    {
        private SqlConnection sqlConnection = null;

        public products_form()
        {
            sqlConnection = new SqlConnection(ConfigurationManager.ConnectionStrings["db"].ConnectionString);
            sqlConnection.Open();
            InitializeComponent();
            combobox(comboBox1, "SELECT id, name FROM [Company]", "name", "id");
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
            string query = $"SELECT * FROM Products WHERE id = {id}";
            var connectionString = ConfigurationManager.ConnectionStrings["db"].ConnectionString;
            using (SqlDataAdapter adapter = new SqlDataAdapter(query, connectionString))
            {
                DataTable dataTable = new DataTable();
                adapter.Fill(dataTable);

                login_textBox.Text = dataTable.Rows[0][1].ToString();
                richTextBox1.Text = dataTable.Rows[0][2].ToString();
                comboBox6.Text = dataTable.Rows[0][3].ToString();
                comboBox1.SelectedValue = (int)dataTable.Rows[0][4];
            }
        }

        private void button2_Click(object sender, EventArgs e)
        {
            if (String.IsNullOrWhiteSpace(login_textBox.Text) ||
                String.IsNullOrWhiteSpace(richTextBox1.Text) ||
                String.IsNullOrWhiteSpace(comboBox6.Text) ||
                String.IsNullOrWhiteSpace(comboBox1.Text))
                MessageBox.Show("Необходимо заполнить все данные!", "Ошибка!");
            else
            {
                if (!isChange)
                {
                    string query = "INSERT INTO Products (name, description, category, company)" +
                    "VALUES(@name, @description, @category, @company)";
                    using (SqlCommand command = new SqlCommand(query, sqlConnection))
                    {
                        command.Parameters.AddWithValue("@name", login_textBox.Text);
                        command.Parameters.AddWithValue("@description", richTextBox1.Text);
                        command.Parameters.AddWithValue("@category", comboBox6.Text);
                        command.Parameters.AddWithValue("@company", comboBox1.SelectedValue);
                        command.ExecuteNonQuery();
                        this.DialogResult = DialogResult.OK;
                    }
                }
                else
                {
                    string query = $"UPDATE Products SET name = @name, description = @description, category = @category, company = @company WHERE id = {idChange}";
                    using (SqlCommand command = new SqlCommand(query, sqlConnection))
                    {
                        command.Parameters.AddWithValue("@name", login_textBox.Text);
                        command.Parameters.AddWithValue("@description", richTextBox1.Text);
                        command.Parameters.AddWithValue("@category", comboBox6.Text);
                        command.Parameters.AddWithValue("@company", comboBox1.SelectedValue);
                        command.ExecuteNonQuery();
                        this.DialogResult = DialogResult.OK;
                    }
                }
            }
        }

        private void iconPictureBox1_Click(object sender, EventArgs e)
        {
            if (new company_edit().ShowDialog() == DialogResult.OK)
            {
                comboBox1.DataSource = null;
                combobox(comboBox1, "SELECT id, name FROM [Company]", "name", "id");
                MessageBox.Show("Запись успешно добавлена!", "Успех!");
            }
        }
    }
}