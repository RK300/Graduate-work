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
using techSupport.new_forms;

namespace techSupport.Ticket_system
{
    public partial class file2ticket_edit : Form
    {
        private SqlConnection sqlConnection = null;

        private int tid;

        public file2ticket_edit(int m_id)
        {
            sqlConnection = new SqlConnection(ConfigurationManager.ConnectionStrings["db"].ConnectionString);
            sqlConnection.Open();
            InitializeComponent();
            combobox(comboBox1, "SELECT Ticket.id, ('#' + CAST(Ticket.id AS nvarchar) + ' | ' + Clients.surname + ' ' + Clients.name + ' ' + Clients.patronymic + ' | ' + CAST(Ticket.application_data AS nvarchar)) AS [TICKET] FROM Ticket, Clients", "TICKET", "id");
            combobox(comboBox2, "SELECT id, ('#' + CAST(id AS nvarchar) + ' | ' + CAST(date_upload AS nvarchar)) as [IMG] FROM [ImageFiles]", "IMG", "id");
            SetId(m_id);
        }

        private void SetId(int m_id) { tid = m_id; }

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
            string query = $"SELECT * FROM File2Tiket WHERE id = {id}";
            var connectionString = ConfigurationManager.ConnectionStrings["db"].ConnectionString;
            using (SqlDataAdapter adapter = new SqlDataAdapter(query, connectionString))
            {
                DataTable dataTable = new DataTable();
                adapter.Fill(dataTable);
                comboBox1.SelectedValue = (int)dataTable.Rows[0][1];
                comboBox2.SelectedValue = (int)dataTable.Rows[0][2];
            }
        }

        private void button2_Click(object sender, EventArgs e)
        {
            if (String.IsNullOrWhiteSpace(comboBox2.Text) ||
                String.IsNullOrWhiteSpace(comboBox1.Text))
                MessageBox.Show("Необходимо заполнить все данные!", "Ошибка!");
            else
            {
                if (!isChange)
                {
                    string query = "INSERT INTO File2Tiket (ticket, photo)" +
                    "VALUES(@ticket, @photo)";
                    using (SqlCommand command = new SqlCommand(query, sqlConnection))
                    {
                        command.Parameters.AddWithValue("@ticket", comboBox1.SelectedValue);
                        command.Parameters.AddWithValue("@photo", comboBox2.SelectedValue);
                        command.ExecuteNonQuery();
                        this.DialogResult = DialogResult.OK;
                    }
                }
                else
                {
                    string query = $"UPDATE File2Tiket SET ticket = @ticket, photo = @photo WHERE id = {idChange}";
                    using (SqlCommand command = new SqlCommand(query, sqlConnection))
                    {
                        command.Parameters.AddWithValue("@ticket", comboBox1.SelectedValue);
                        command.Parameters.AddWithValue("@photo", comboBox2.SelectedValue);
                        command.ExecuteNonQuery();
                        this.DialogResult = DialogResult.OK;
                    }
                }
            }
        }

        private void iconPictureBox2_Click(object sender, EventArgs e)
        {
            if (new ticket_edit(tid).ShowDialog() == DialogResult.OK)
            {
                comboBox1.DataSource = null;
                combobox(comboBox1, "SELECT Ticket.id, ('#' + CAST(Ticket.id AS nvarchar) + ' | ' + Clients.surname + ' ' + Clients.name + ' ' + Clients.patronymic + ' | ' + CAST(Ticket.application_data AS nvarchar)) AS [TICKET] FROM Ticket, Clients", "TICKET", "id");
                MessageBox.Show("Запись успешно добавлена!", "Успех!");
            }
        }

        private void iconPictureBox1_Click(object sender, EventArgs e)
        {
            /*if (new file_edit().ShowDialog() == DialogResult.OK)
            {
                comboBox2.DataSource = null;
                combobox(comboBox2, "SELECT id, ('#' + CAST(id AS nvarchar) + ' | ' + CAST(date_upload AS nvarchar)) as [IMG] FROM [ImageFiles]", "IMG", "id");
                MessageBox.Show("Запись успешно добавлена!", "Успех!");
            }*/
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