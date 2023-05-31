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
using System.Windows.Controls;
using System.Windows.Forms;
using techSupport.edit_form;
using techSupport.forms;
using ComboBox = System.Windows.Forms.ComboBox;

namespace techSupport.new_forms
{
    public partial class dogovor_edit : Form
    {
        private SqlConnection sqlConnection = null;

        public dogovor_edit()
        {
            sqlConnection = new SqlConnection(ConfigurationManager.ConnectionStrings["db"].ConnectionString);
            sqlConnection.Open();
            InitializeComponent();
            combobox(comboBox2, "SELECT id, (CompanyName + ' | ' + surname + ' ' + name + ' ' + patronymic) AS [FIO] FROM [Clients]", "FIO", "id");

            if (isChange == true)
            {
                textBox2.Enabled = true;
            }
            else
            {
                textBox2.Enabled = false;
                SetNumber();
            }
        }

        private void SetNumber()
        {
            string query = "SELECT MAX(nomer) FROM Treaty";
            var connectionString = ConfigurationManager.ConnectionStrings["db"].ConnectionString;
            using (SqlDataAdapter adapter = new SqlDataAdapter(query, connectionString))
            {
                System.Data.DataTable dataTable = new System.Data.DataTable();
                adapter.Fill(dataTable);
                int m_Number = (int)dataTable.Rows[0][0] + 1;
                textBox2.Text = m_Number.ToString();
            }
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

        private void combobox(System.Windows.Forms.ComboBox c, string query, string displaymember, string valuemember)
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
            string query = $"SELECT * FROM Treaty WHERE id = {id}";
            var connectionString = ConfigurationManager.ConnectionStrings["db"].ConnectionString;
            using (SqlDataAdapter adapter = new SqlDataAdapter(query, connectionString))
            {
                DataTable dataTable = new DataTable();
                adapter.Fill(dataTable);
                comboBox2.SelectedValue = (int)dataTable.Rows[0][1];
                textBox2.Text = dataTable.Rows[0][2].ToString();
                comboBox1.SelectedValue = (int)dataTable.Rows[0][3];
                dateTimePicker1.Value = (DateTime)dataTable.Rows[0][4];
                dateTimePicker2.Value = (DateTime)dataTable.Rows[0][5];
                dateTimePicker3.Value = (DateTime)dataTable.Rows[0][6];
            }
            textBox2.Enabled = true;
        }

        private void button2_Click_1(object sender, EventArgs e)
        {
            if (String.IsNullOrWhiteSpace(comboBox2.Text) || String.IsNullOrWhiteSpace(dateTimePicker1.Text) || String.IsNullOrWhiteSpace(dateTimePicker2.Text) || String.IsNullOrWhiteSpace(dateTimePicker3.Text))
                MessageBox.Show("Необходимо заполнить все данные!", "Ошибка!");
            else
            {
                if (!isChange)
                {
                    string query = "INSERT INTO Treaty (client, product, dateСonclusion, dataFrom, dateTo, nomer)" +
                    "VALUES(@client, @product, @dateСonclusion, @dataFrom, @dateTo, @nomer)";
                    using (SqlCommand command = new SqlCommand(query, sqlConnection))
                    {
                        command.Parameters.AddWithValue("@client", comboBox2.SelectedValue);
                        command.Parameters.AddWithValue("@product", comboBox1.SelectedValue);
                        command.Parameters.AddWithValue("@dateСonclusion", dateTimePicker1.Value);
                        command.Parameters.AddWithValue("@dataFrom", dateTimePicker2.Value);
                        command.Parameters.AddWithValue("@dateTo", dateTimePicker3.Value);
                        command.Parameters.AddWithValue("@nomer", textBox2.Text);
                        command.ExecuteNonQuery();
                        this.DialogResult = DialogResult.OK;
                    }
                }
                else
                {
                    string query = $"UPDATE Treaty SET client = @client, product = @product, dateСonclusion = @dateСonclusion, dataFrom = @dataFrom, dateTo = @dateTo, nomer = @nomer WHERE id = {idChange}";
                    using (SqlCommand command = new SqlCommand(query, sqlConnection))
                    {
                        command.Parameters.AddWithValue("@client", comboBox2.SelectedValue);
                        command.Parameters.AddWithValue("@product", comboBox1.SelectedValue);
                        command.Parameters.AddWithValue("@dateСonclusion", dateTimePicker1.Value);
                        command.Parameters.AddWithValue("@dataFrom", dateTimePicker2.Value);
                        command.Parameters.AddWithValue("@dateTo", dateTimePicker3.Value);
                        command.Parameters.AddWithValue("@nomer", textBox2.Text);
                        command.ExecuteNonQuery();
                        this.DialogResult = DialogResult.OK;
                    }
                }
            }
        }

        private void comboBox2_SelectedIndexChanged(object sender, EventArgs e)
        {
            comboBox1.DataSource = null;
            if (comboBox2.SelectedIndex != 0)
            {
                if (comboBox2.SelectedValue != null)
                {
                    int m_id = (int)comboBox2.SelectedValue;
                    combobox(comboBox1, $"SELECT Products.id AS [n1], Products.name AS [n2] FROM [Products], [User2Product], [Clients] WHERE Products.id = User2Product.product AND Clients.id = User2Product.client AND User2Product.client = {m_id}", "n2", "n1");
                }
            }
        }

        private void iconPictureBox2_Click(object sender, EventArgs e)
        {
            if (new client_edit().ShowDialog() == DialogResult.OK)
            {
                comboBox2.DataSource = null;
                combobox(comboBox2, "SELECT id, (surname + ' ' + name + ' ' + patronymic) AS [FIO] FROM [Clients]", "FIO", "id");
                MessageBox.Show("Запись успешно добавлена!", "Успех!");
            }
        }

        private void iconPictureBox1_Click(object sender, EventArgs e)
        {
            if (new products_form().ShowDialog() == DialogResult.OK)
            {
                comboBox1.DataSource = null;
                if (comboBox2.SelectedIndex != 0)
                {
                    if (comboBox2.SelectedValue != null)
                    {
                        int m_id = (int)comboBox2.SelectedValue;
                        combobox(comboBox1, $"SELECT Products.id AS [n1], Products.name AS [n2] FROM [Products], [User2Product], [Clients] WHERE Products.id = User2Product.product AND Clients.id = User2Product.client AND User2Product.client = {m_id}", "n2", "n1");
                    }
                }
                MessageBox.Show("Запись успешно добавлена!", "Успех!");
            }
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