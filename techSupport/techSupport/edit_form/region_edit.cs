﻿using System;
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
using techSupport.new_forms;

namespace techSupport.edit_form
{
    public partial class region_edit : Form
    {
        private SqlConnection sqlConnection = null;

        public region_edit()
        {
            sqlConnection = new SqlConnection(ConfigurationManager.ConnectionStrings["db"].ConnectionString);
            sqlConnection.Open();
            InitializeComponent();
            combobox(comboBox1, "SELECT id, name FROM [Country]", "name", "id");
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
            string query = $"SELECT * FROM Region WHERE id = {id}";
            var connectionString = ConfigurationManager.ConnectionStrings["db"].ConnectionString;
            using (SqlDataAdapter adapter = new SqlDataAdapter(query, connectionString))
            {
                DataTable dataTable = new DataTable();
                adapter.Fill(dataTable);
                comboBox1.SelectedValue = (int)dataTable.Rows[0][1];
                login_textBox.Text = dataTable.Rows[0][2].ToString();
            }
        }

        private void button2_Click(object sender, EventArgs e)
        {
            if (String.IsNullOrWhiteSpace(login_textBox.Text) ||
                String.IsNullOrWhiteSpace(comboBox1.Text))
                MessageBox.Show("Необходимо заполнить все данные!", "Ошибка!");
            else
            {
                if (!isChange)
                {
                    string query = "INSERT INTO Region (country, name)" +
                    "VALUES(@country, @name)";
                    using (SqlCommand command = new SqlCommand(query, sqlConnection))
                    {
                        command.Parameters.AddWithValue("@country", comboBox1.SelectedValue);
                        command.Parameters.AddWithValue("@name", login_textBox.Text);
                        command.ExecuteNonQuery();
                        this.DialogResult = DialogResult.OK;
                    }
                }
                else
                {
                    string query = $"UPDATE Region SET country = @country, name = @name WHERE id = {idChange}";
                    using (SqlCommand command = new SqlCommand(query, sqlConnection))
                    {
                        command.Parameters.AddWithValue("@country", comboBox1.SelectedValue);
                        command.Parameters.AddWithValue("@name", login_textBox.Text);
                        command.ExecuteNonQuery();
                        this.DialogResult = DialogResult.OK;
                    }
                }
            }
        }

        private void iconPictureBox2_Click(object sender, EventArgs e)
        {
            if (new country_edit().ShowDialog() == DialogResult.OK)
            {
                comboBox1.DataSource = null;
                combobox(comboBox1, "SELECT id, name FROM [Country]", "name", "id");
                MessageBox.Show("Запись успешно добавлена!", "Успех!");
            }
        }
    }
}