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

namespace techSupport.Ticket_system
{
    public partial class type_form : Form
    {
        private SqlConnection sqlConnection = null;

        public type_form()
        {
            sqlConnection = new SqlConnection(ConfigurationManager.ConnectionStrings["db"].ConnectionString);
            sqlConnection.Open();
            InitializeComponent();
            RefreshTable();
        }

        private void RefreshTable()
        {
            string query = "SELECT id AS [Номер], name AS [Название] FROM Type";
            var connectionString = ConfigurationManager.ConnectionStrings["db"].ConnectionString;
            using (SqlDataAdapter adapter = new SqlDataAdapter(query, connectionString))
            {
                DataTable dataTable = new DataTable();
                adapter.Fill(dataTable);
                dataGridView1.DataSource = dataTable;
            }
            dataGridView1.Columns[0].Visible = false;
        }

        private void rjButton1_Click(object sender, EventArgs e)
        {
            if (new type_edit().ShowDialog() == DialogResult.OK)
            {
                RefreshTable();
                MessageBox.Show("Запись успешно добавлена!", "Успех!");
            }
        }

        private void rjButton2_Click(object sender, EventArgs e)
        {
            type_edit F2 = new type_edit();
            F2.Text = "Редактирование типа";
            F2.IDCHANGE = dataGridView1.Rows[dataGridView1.CurrentCell.RowIndex].Cells[0].Value.ToString();
            if (F2.ShowDialog() == DialogResult.OK)
            {
                RefreshTable();
                MessageBox.Show("Запись успешно изменена!", "Успех!");
            }
        }

        private void rjButton3_Click(object sender, EventArgs e)
        {
            if (MessageBox.Show("Вы действительно хотите удалить запись?", "Подтверждение удаления", MessageBoxButtons.YesNo, MessageBoxIcon.Warning) == DialogResult.Yes)
            {
                try
                {
                    SqlCommand command = new SqlCommand($"DELETE FROM Type WHERE id = {dataGridView1.Rows[dataGridView1.CurrentCell.RowIndex].Cells[0].Value}", sqlConnection);
                    command.ExecuteNonQuery();
                    MessageBox.Show("Запись успешно удалена!", "Успех!");
                    RefreshTable();
                }
                catch
                {
                    MessageBox.Show("Ошибка, попробуйте еще раз!", "Внимание", MessageBoxButtons.OK, MessageBoxIcon.Error);
                }
            }
        }
    }
}
