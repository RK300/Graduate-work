using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Configuration;
using System.Data;
using System.Data.SqlClient;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Text.RegularExpressions;
using System.Threading.Tasks;
using System.Windows.Forms;

namespace techSupport.edit_form
{
    public partial class company_edit : Form
    {
        private SqlConnection sqlConnection = null;

        public company_edit()
        {
            sqlConnection = new SqlConnection(ConfigurationManager.ConnectionStrings["db"].ConnectionString);
            sqlConnection.Open();
            InitializeComponent();
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

        private void FillBoxes(string id)
        {
            string query = $"SELECT * FROM Company WHERE id = {id}";
            var connectionString = ConfigurationManager.ConnectionStrings["db"].ConnectionString;
            using (SqlDataAdapter adapter = new SqlDataAdapter(query, connectionString))
            {
                DataTable dataTable = new DataTable();
                adapter.Fill(dataTable);
                login_textBox.Text = dataTable.Rows[0][1].ToString();
                textBox2.Text = dataTable.Rows[0][2].ToString();
                maskedTextBox3.Text = dataTable.Rows[0][3].ToString();
                maskedTextBox4.Text = dataTable.Rows[0][4].ToString();
                maskedTextBox2.Text = dataTable.Rows[0][5].ToString();
                maskedTextBox1.Text = dataTable.Rows[0][6].ToString();
            }
        }

        private void button2_Click(object sender, EventArgs e)
        {
            if (String.IsNullOrWhiteSpace(login_textBox.Text) ||
                String.IsNullOrWhiteSpace(maskedTextBox3.Text) ||
                String.IsNullOrWhiteSpace(maskedTextBox4.Text) ||
                String.IsNullOrWhiteSpace(maskedTextBox2.Text) ||
                String.IsNullOrWhiteSpace(maskedTextBox1.Text) ||
                String.IsNullOrWhiteSpace(textBox2.Text))
                MessageBox.Show("Необходимо заполнить все данные!", "Ошибка!");
            else
            {
                if (!isChange)
                {
                    string query = "INSERT INTO Company (name, phone, email, Checking, BIK, YNP)" +
                    "VALUES(@name, @phone, @email, @Checking, @BIK, @YNP)";
                    using (SqlCommand command = new SqlCommand(query, sqlConnection))
                    {
                        command.Parameters.AddWithValue("@name", login_textBox.Text);
                        command.Parameters.AddWithValue("@phone", maskedTextBox3.Text);
                        command.Parameters.AddWithValue("@email", textBox2.Text);
                        command.Parameters.AddWithValue("@Checking", maskedTextBox4.Text);
                        command.Parameters.AddWithValue("@BIK", maskedTextBox2.Text);
                        command.Parameters.AddWithValue("@YNP", maskedTextBox1.Text);
                        command.ExecuteNonQuery();
                        this.DialogResult = DialogResult.OK;
                    }
                }
                else
                {
                    string query = $"UPDATE Company SET name = @name, phone = @phone, email = @email, Checking = @Checking, BIK = @BIK, YNP = @YNP WHERE id = {idChange}";
                    using (SqlCommand command = new SqlCommand(query, sqlConnection))
                    {
                        command.Parameters.AddWithValue("@name", login_textBox.Text);
                        command.Parameters.AddWithValue("@phone", maskedTextBox3.Text);
                        command.Parameters.AddWithValue("@email", textBox2.Text);
                        command.Parameters.AddWithValue("@Checking", maskedTextBox4.Text);
                        command.Parameters.AddWithValue("@BIK", maskedTextBox2.Text);
                        command.Parameters.AddWithValue("@YNP", maskedTextBox1.Text);
                        command.ExecuteNonQuery();
                        this.DialogResult = DialogResult.OK;
                    }
                }
            }
        }

        private void textBox2_Validating(object sender, CancelEventArgs e)
        {
            if (!ValidEmailAddress(textBox2.Text))
            {
                e.Cancel = true;
                textBox2.Select(0, textBox2.Text.Length);
                MessageBox.Show("Введите корректный e-mail!", "Ошибка!");
            }
        }

        private bool ValidEmailAddress(string email)
        {
            string regexPattern = @"^[a-zA-Z0-9._%+-]+@[a-zA-Z0-9.-]+\.[a-zA-Z]{2,4}$";
            return Regex.IsMatch(email, regexPattern);
        }
    }
}