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

namespace techSupport.new_forms
{
    public partial class bank_edit : Form
    {
        private SqlConnection sqlConnection = null;

        public bank_edit()
        {
            sqlConnection = new SqlConnection(ConfigurationManager.ConnectionStrings["db"].ConnectionString);
            sqlConnection.Open();
            InitializeComponent();
            combobox(comboBox2, "SELECT id, name FROM [Locality]", "name", "id");
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
            string query = $"SELECT * FROM Bank WHERE id = {id}";
            var connectionString = ConfigurationManager.ConnectionStrings["db"].ConnectionString;
            using (SqlDataAdapter adapter = new SqlDataAdapter(query, connectionString))
            {
                DataTable dataTable = new DataTable();
                adapter.Fill(dataTable);
                comboBox2.SelectedValue = (int)dataTable.Rows[0][1];
                textBox1.Text = dataTable.Rows[0][2].ToString();
                maskedTextBox2.Text = dataTable.Rows[0][3].ToString();
                login_textBox.Text = dataTable.Rows[0][4].ToString();
                textBox5.Text = dataTable.Rows[0][5].ToString();
                textBox4.Text = dataTable.Rows[0][6].ToString();
            }
        }

        private void button2_Click(object sender, EventArgs e)
        {
            if (String.IsNullOrWhiteSpace(login_textBox.Text) ||
                String.IsNullOrWhiteSpace(textBox1.Text) ||
                String.IsNullOrWhiteSpace(maskedTextBox2.Text) ||
                String.IsNullOrWhiteSpace(textBox5.Text) ||
                String.IsNullOrWhiteSpace(textBox4.Text) ||
                String.IsNullOrWhiteSpace(comboBox2.Text))
                MessageBox.Show("Необходимо заполнить все данные!", "Ошибка!");
            else
            {
                if (!isChange)
                {
                    string query = "INSERT INTO Bank (location, name, BIK, street, house, corpse)" +
                    "VALUES(@location, @name, @BIK, @street, @house, @corpse)";
                    using (SqlCommand command = new SqlCommand(query, sqlConnection))
                    {
                        command.Parameters.AddWithValue("@location", comboBox2.SelectedValue);
                        command.Parameters.AddWithValue("@name", textBox1.Text);
                        command.Parameters.AddWithValue("@BIK", maskedTextBox2.Text);
                        command.Parameters.AddWithValue("@street", login_textBox.Text);
                        command.Parameters.AddWithValue("@house", textBox5.Text);
                        command.Parameters.AddWithValue("@corpse", textBox4.Text);
                        command.ExecuteNonQuery();
                        this.DialogResult = DialogResult.OK;
                    }
                }
                else
                {
                    string query = $"UPDATE Bank SET location = @location, name = @name, BIK = @BIK, street = @street, house = @house, corpse = @corpse WHERE id = {idChange}";
                    using (SqlCommand command = new SqlCommand(query, sqlConnection))
                    {
                        command.Parameters.AddWithValue("@location", comboBox2.SelectedValue);
                        command.Parameters.AddWithValue("@name", textBox1.Text);
                        command.Parameters.AddWithValue("@BIK", maskedTextBox2.Text);
                        command.Parameters.AddWithValue("@street", login_textBox.Text);
                        command.Parameters.AddWithValue("@house", textBox5.Text);
                        command.Parameters.AddWithValue("@corpse", textBox4.Text);
                        command.ExecuteNonQuery();
                        this.DialogResult = DialogResult.OK;
                    }
                }
            }
        }

        private void iconPictureBox2_Click(object sender, EventArgs e)
        {
            if (new locality_edit().ShowDialog() == DialogResult.OK)
            {
                comboBox2.DataSource = null;
                combobox(comboBox2, "SELECT id, name FROM [Locality]", "name", "id");
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
    }
}
