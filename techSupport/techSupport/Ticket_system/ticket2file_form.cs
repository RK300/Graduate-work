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

namespace techSupport.Ticket_system
{
    public partial class ticket2file_form : Form
    {
        private SqlConnection sqlConnection = null;

        private int tid;

        public ticket2file_form(int m_id)
        {
            sqlConnection = new SqlConnection(ConfigurationManager.ConnectionStrings["db"].ConnectionString);
            sqlConnection.Open();
            InitializeComponent();
            RefreshTable();
            SetId(m_id);
        }

        private void SetId(int id) { tid = id; }

        private void RefreshTable()
        {
            string query = "SELECT File2Tiket.id, ('#' + CAST(Ticket.id AS nvarchar) + ' | ' + Clients.surname + ' ' + Clients.name + ' ' + Clients.patronymic + ' | ' + CAST(Ticket.application_data AS nvarchar)) AS [Тикет], ('#' + CAST(ImageFiles.id AS nvarchar) + ' | ' + CAST(ImageFiles.date_upload AS nvarchar)) AS [Файл] FROM File2Tiket , Ticket, ImageFiles, Clients WHERE File2Tiket.ticket = Ticket.id AND File2Tiket.photo = ImageFiles.id AND Ticket.client = Clients.id";
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
            if (new file2ticket_edit(tid).ShowDialog() == DialogResult.OK)
            {
                RefreshTable();
                MessageBox.Show("Запись успешно добавлена!", "Успех!");
            }
        }

        private void rjButton2_Click(object sender, EventArgs e)
        {
            file2ticket_edit F2 = new file2ticket_edit(tid);
            F2.Text = "Редактирование присвоения";
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
                    SqlCommand command = new SqlCommand($"DELETE FROM File2Tiket WHERE id = {dataGridView1.Rows[dataGridView1.CurrentCell.RowIndex].Cells[0].Value}", sqlConnection);
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