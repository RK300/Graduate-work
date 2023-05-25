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
    public partial class position_edit : Form
    {
        private SqlConnection sqlConnection = null;

        public position_edit()
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
            string query = $"SELECT * FROM Position WHERE id = {id}";
            var connectionString = ConfigurationManager.ConnectionStrings["db"].ConnectionString;
            using (SqlDataAdapter adapter = new SqlDataAdapter(query, connectionString))
            {
                DataTable dataTable = new DataTable();
                adapter.Fill(dataTable);
                login_textBox.Text = dataTable.Rows[0][1].ToString();
            }
        }

        private void button2_Click(object sender, EventArgs e)
        {
            if (String.IsNullOrWhiteSpace(login_textBox.Text))
                MessageBox.Show("Необходимо заполнить все данные!", "Ошибка!");
            else
            {
                if (!isChange)
                {
                    string query = "INSERT INTO Position (name)" +
                    "VALUES(@name)";
                    using (SqlCommand command = new SqlCommand(query, sqlConnection))
                    {
                        command.Parameters.AddWithValue("@name", login_textBox.Text);
                        command.ExecuteNonQuery();
                        this.DialogResult = DialogResult.OK;
                    }
                }
                else
                {
                    string query = $"UPDATE Position SET name = @name WHERE id = {idChange}";
                    using (SqlCommand command = new SqlCommand(query, sqlConnection))
                    {
                        command.Parameters.AddWithValue("@name", login_textBox.Text);
                        command.ExecuteNonQuery();
                        this.DialogResult = DialogResult.OK;
                    }
                }
            }
        }
    }
}