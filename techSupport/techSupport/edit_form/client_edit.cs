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
using techSupport.new_forms;

namespace techSupport.edit_form
{
    public partial class client_edit : Form
    {
        private SqlConnection sqlConnection = null;

        public client_edit()
        {
            sqlConnection = new SqlConnection(ConfigurationManager.ConnectionStrings["db"].ConnectionString);
            sqlConnection.Open();
            InitializeComponent();
            combobox(comboBox1, "SELECT id, name FROM [Locality]", "name", "id");
            combobox(comboBox2, "SELECT id, name FROM [Bank]", "name", "id");
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
            string query = $"SELECT * FROM Clients WHERE id = {id}";
            var connectionString = ConfigurationManager.ConnectionStrings["db"].ConnectionString;
            using (SqlDataAdapter adapter = new SqlDataAdapter(query, connectionString))
            {
                DataTable dataTable = new DataTable();
                adapter.Fill(dataTable);
                textBox8.Text = dataTable.Rows[0][1].ToString();
                login_textBox.Text = dataTable.Rows[0][2].ToString();
                textBox1.Text = dataTable.Rows[0][3].ToString();
                textBox2.Text = dataTable.Rows[0][4].ToString();
                comboBox1.SelectedValue = (int)dataTable.Rows[0][5];
                comboBox2.SelectedValue = (int)dataTable.Rows[0][6];
                textBox3.Text = dataTable.Rows[0][7].ToString();
                maskedTextBox1.Text = dataTable.Rows[0][8].ToString();
                maskedTextBox4.Text = dataTable.Rows[0][9].ToString();
                maskedTextBox2.Text = dataTable.Rows[0][10].ToString();
                maskedTextBox3.Text = dataTable.Rows[0][11].ToString();
                textBox9.Text = dataTable.Rows[0][12].ToString();
                textBox10.Text = dataTable.Rows[0][13].ToString();
                textBox11.Text = dataTable.Rows[0][14].ToString();
            }
        }

        private void button2_Click(object sender, EventArgs e)
        {
            if (String.IsNullOrWhiteSpace(login_textBox.Text) ||
                String.IsNullOrWhiteSpace(textBox1.Text) ||
                String.IsNullOrWhiteSpace(textBox2.Text) ||
                String.IsNullOrWhiteSpace(maskedTextBox1.Text) ||
                String.IsNullOrWhiteSpace(maskedTextBox2.Text) ||
                String.IsNullOrWhiteSpace(maskedTextBox3.Text) ||
                String.IsNullOrWhiteSpace(maskedTextBox4.Text) ||
                String.IsNullOrWhiteSpace(textBox8.Text) ||
                String.IsNullOrWhiteSpace(textBox9.Text) ||
                String.IsNullOrWhiteSpace(textBox11.Text) ||
                String.IsNullOrWhiteSpace(textBox10.Text) ||
                String.IsNullOrWhiteSpace(comboBox1.Text) ||
                String.IsNullOrWhiteSpace(comboBox2.Text) ||
                String.IsNullOrWhiteSpace(textBox3.Text))
                MessageBox.Show("Необходимо заполнить все данные!", "Ошибка!");
            else
            {
                if (!isChange)
                {
                    string query = "INSERT INTO Clients (CompanyName, surname, name, patronymic, city, bank, email, phone, Checking, BIK, YNP, street, house, corpus)" +
                    "VALUES(@CompanyName, @surname, @name, @patronymic, @city, @bank, @email, @phone, @Checking, @BIK, @YNP, @street, @house, @corpus)";
                    using (SqlCommand command = new SqlCommand(query, sqlConnection))
                    {
                        command.Parameters.AddWithValue("@surname", login_textBox.Text);
                        command.Parameters.AddWithValue("@name", textBox1.Text);
                        command.Parameters.AddWithValue("@patronymic", textBox2.Text);
                        command.Parameters.AddWithValue("@city", comboBox1.SelectedValue);
                        command.Parameters.AddWithValue("@bank", comboBox2.SelectedValue);
                        command.Parameters.AddWithValue("@email", textBox3.Text);
                        command.Parameters.AddWithValue("@CompanyName", textBox8.Text);
                        command.Parameters.AddWithValue("@phone", maskedTextBox1.Text);
                        command.Parameters.AddWithValue("@Checking", maskedTextBox4.Text);
                        command.Parameters.AddWithValue("@BIK", maskedTextBox2.Text);
                        command.Parameters.AddWithValue("@YNP", maskedTextBox3.Text);
                        command.Parameters.AddWithValue("@street", textBox9.Text);
                        command.Parameters.AddWithValue("@house", textBox10.Text);
                        command.Parameters.AddWithValue("@corpus", textBox11.Text);
                        command.ExecuteNonQuery();
                        this.DialogResult = DialogResult.OK;
                    }
                }
                else
                {
                    string query = $"UPDATE Clients SET CompanyName = @CompanyName, surname = @surname, name = @name, patronymic = @patronymic, city = @city, bank = @bank, email = @email, phone = @phone, Checking = @Checking, BIK = @BIK, YNP = @YNP, street = @street, house = @house, corpus = @corpus WHERE id = {idChange}";
                    using (SqlCommand command = new SqlCommand(query, sqlConnection))
                    {
                        command.Parameters.AddWithValue("@surname", login_textBox.Text);
                        command.Parameters.AddWithValue("@name", textBox1.Text);
                        command.Parameters.AddWithValue("@patronymic", textBox2.Text);
                        command.Parameters.AddWithValue("@city", comboBox1.SelectedValue);
                        command.Parameters.AddWithValue("@bank", comboBox2.SelectedValue);
                        command.Parameters.AddWithValue("@email", textBox3.Text);
                        command.Parameters.AddWithValue("@CompanyName", textBox8.Text);
                        command.Parameters.AddWithValue("@phone", maskedTextBox1.Text);
                        command.Parameters.AddWithValue("@Checking", maskedTextBox4.Text);
                        command.Parameters.AddWithValue("@BIK", maskedTextBox2.Text);
                        command.Parameters.AddWithValue("@YNP", maskedTextBox3.Text);
                        command.Parameters.AddWithValue("@street", textBox9.Text);
                        command.Parameters.AddWithValue("@house", textBox10.Text);
                        command.Parameters.AddWithValue("@corpus", textBox11.Text);
                        command.ExecuteNonQuery();
                        this.DialogResult = DialogResult.OK;
                    }
                }
            }
        }

        private bool ValidEmailAddress(string email)
        {
            string regexPattern = @"^[a-zA-Z0-9._%+-]+@[a-zA-Z0-9.-]+\.[a-zA-Z]{2,4}$";
            return Regex.IsMatch(email, regexPattern);
        }

        private void iconPictureBox1_Click(object sender, EventArgs e)
        {
            if (new bank_edit().ShowDialog() == DialogResult.OK)
            {
                comboBox2.DataSource = null;
                combobox(comboBox2, "SELECT id, name FROM [Bank]", "name", "id");
                MessageBox.Show("Запись успешно добавлена!", "Успех!");
            }
        }

        private void iconPictureBox2_Click(object sender, EventArgs e)
        {
            if (new locality_edit().ShowDialog() == DialogResult.OK)
            {
                comboBox1.DataSource = null;
                combobox(comboBox1, "SELECT id, name FROM [Locality]", "name", "id");
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

        private void comboBox2_KeyPress(object sender, KeyPressEventArgs e)
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