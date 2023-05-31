using md5_sql_hash;
using System;
using System.Collections;
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
using techSupport.new_forms;

namespace techSupport.edit_form
{
    public partial class worker_edit : Form
    {
        private SqlConnection sqlConnection = null;

        public worker_edit()
        {
            sqlConnection = new SqlConnection(ConfigurationManager.ConnectionStrings["db"].ConnectionString);
            sqlConnection.Open();
            InitializeComponent();
            combobox(comboBox1, "SELECT id, name FROM [Position]", "name", "id");
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
            string query = $"SELECT * FROM Worker WHERE id = {id}";
            var connectionString = ConfigurationManager.ConnectionStrings["db"].ConnectionString;
            using (SqlDataAdapter adapter = new SqlDataAdapter(query, connectionString))
            {
                DataTable dataTable = new DataTable();
                adapter.Fill(dataTable);
                login_textBox.Text = dataTable.Rows[0][1].ToString();
                textBox1.Text = dataTable.Rows[0][2].ToString();
                textBox2.Text = dataTable.Rows[0][3].ToString();
                comboBox1.SelectedValue = (int)dataTable.Rows[0][4];
                textBox3.Text = dataTable.Rows[0][5].ToString();
                maskedTextBox1.Text = dataTable.Rows[0][6].ToString();
                textBox5.Text = dataTable.Rows[0][7].ToString();
                textBox6.Text = dataTable.Rows[0][8].ToString();
            }
        }

        private void button2_Click(object sender, EventArgs e)
        {
            if (String.IsNullOrWhiteSpace(login_textBox.Text) ||
                String.IsNullOrWhiteSpace(textBox1.Text) ||
                String.IsNullOrWhiteSpace(textBox2.Text) ||
                String.IsNullOrWhiteSpace(maskedTextBox1.Text) ||
                String.IsNullOrWhiteSpace(comboBox1.Text) ||
                String.IsNullOrWhiteSpace(textBox3.Text))
                MessageBox.Show("Необходимо заполнить все данные!", "Ошибка!");
            else
            {
                if (!isChange)
                {
                    string query = "INSERT INTO Worker (surname, name, patronymic, position, email, phone, login, password)" +
                    "VALUES(@surname, @name, @patronymic, @position, @email, @phone, @login, @password)";
                    using (SqlCommand command = new SqlCommand(query, sqlConnection))
                    {
                        command.Parameters.AddWithValue("@surname", login_textBox.Text);
                        command.Parameters.AddWithValue("@name", textBox1.Text);
                        command.Parameters.AddWithValue("@patronymic", textBox2.Text);
                        command.Parameters.AddWithValue("@position", comboBox1.SelectedValue);
                        command.Parameters.AddWithValue("@email", textBox3.Text);
                        command.Parameters.AddWithValue("@phone", maskedTextBox1.Text);
                        command.Parameters.AddWithValue("@login", textBox5.Text);
                        command.Parameters.AddWithValue("@password", md5.hashPassword(textBox6.Text));
                        command.ExecuteNonQuery();
                        this.DialogResult = DialogResult.OK;
                    }
                }
                else
                {
                    string m_pass;
                    string query_1 = $"SELECT password FROM Worker WHERE id = {idChange}";
                    var connectionString = ConfigurationManager.ConnectionStrings["db"].ConnectionString;
                    using (SqlDataAdapter adapter = new SqlDataAdapter(query_1, connectionString))
                    {
                        DataTable dataTable = new DataTable();
                        adapter.Fill(dataTable);
                        m_pass = dataTable.Rows[0][0].ToString();
                    }

                    string query = $"UPDATE Worker SET surname = @surname, name = @name, patronymic = @patronymic, position = @position, email = @email, phone = @phone, login = @login, password = @password WHERE id = {idChange}";
                    using (SqlCommand command = new SqlCommand(query, sqlConnection))
                    {
                        command.Parameters.AddWithValue("@surname", login_textBox.Text);
                        command.Parameters.AddWithValue("@name", textBox1.Text);
                        command.Parameters.AddWithValue("@patronymic", textBox2.Text);
                        command.Parameters.AddWithValue("@position", comboBox1.SelectedValue);
                        command.Parameters.AddWithValue("@email", textBox3.Text);
                        command.Parameters.AddWithValue("@phone", maskedTextBox1.Text);
                        command.Parameters.AddWithValue("@login", textBox5.Text);
                        if ( m_pass == textBox6.Text )
                        {
                            command.Parameters.AddWithValue("@password", textBox6.Text);
                        }
                        else 
                        {
                            command.Parameters.AddWithValue("@password", md5.hashPassword(textBox6.Text));
                        }
                        command.ExecuteNonQuery();
                        this.DialogResult = DialogResult.OK;
                    }
                }
            }
        }

        private void iconPictureBox2_Click(object sender, EventArgs e)
        {
            if (new position_edit().ShowDialog() == DialogResult.OK)
            {
                comboBox1.DataSource = null;
                combobox(comboBox1, "SELECT id, name FROM [Position]", "name", "id");
                MessageBox.Show("Запись успешно добавлена!", "Успех!");
            }
        }

        private void textBox3_Validating(object sender, CancelEventArgs e)
        {
            if (!ValidEmailAddress(textBox3.Text))
            {
                e.Cancel = true;
                textBox3.Select(0, textBox3.Text.Length);
                MessageBox.Show("Введите корректный e-mail!", "Ошибка!");
            }
        }

        private bool ValidEmailAddress(string email)
        {
            string regexPattern = @"^[a-zA-Z0-9._%+-]+@[a-zA-Z0-9.-]+\.[a-zA-Z]{2,4}$";
            return Regex.IsMatch(email, regexPattern);
        }

        private void comboBox1_KeyPress(object sender, KeyPressEventArgs e)
        {
            ((ComboBox)(sender)).DroppedDown = true;
            if ((char.IsControl(e.KeyChar)))
                return;
            string Str = ((ComboBox)(sender)).Text.Substring(0, ((ComboBox)(sender)).SelectionStart) + e.KeyChar;
            int Index = ((ComboBox)(sender)).FindStringExact(Str);
            if (Index == -1)
                Index = ((ComboBox)(sender)).FindString(Str);
            ((ComboBox)sender).SelectedIndex = Index;
            ((ComboBox)(sender)).SelectionStart = Str.Length;
            ((ComboBox)(sender)).SelectionLength = ((ComboBox)(sender)).Text.Length - ((ComboBox)(sender)).SelectionStart;
            e.Handled = true;
        }
    }
}