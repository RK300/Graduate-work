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
using System.Windows.Input;
using techSupport.edit_form;
using techSupport.new_forms;

namespace techSupport.edit_forms_two
{
    public partial class charge_edit : Form
    {
        private SqlConnection sqlConnection = null;
        private int m_tid;

        public charge_edit(int ticket_id)
        {
            sqlConnection = new SqlConnection(ConfigurationManager.ConnectionStrings["db"].ConnectionString);
            sqlConnection.Open();
            InitializeComponent();
            combobox(comboBox1, "SELECT id, (surname + ' ' + name + ' ' + patronymic) AS [FIO] FROM [Worker]", "FIO", "id");
            dateTimePicker1.Value = DateTime.Now;
            SetId(ticket_id);
        }

        private void SetId(int tid) { m_tid =  tid; }

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
            string query = $"SELECT * FROM Charge WHERE id = {id}";
            var connectionString = ConfigurationManager.ConnectionStrings["db"].ConnectionString;
            using (SqlDataAdapter adapter = new SqlDataAdapter(query, connectionString))
            {
                DataTable dataTable = new DataTable();
                adapter.Fill(dataTable);
                comboBox1.SelectedValue = (int)dataTable.Rows[0][2];
                richTextBox1.Text = dataTable.Rows[0][3].ToString();
                dateTimePicker1.Value = (DateTime)dataTable.Rows[0][4];
            }

            string query_check = $"SELECT status FROM Ticket WHERE id = {m_tid}";
            using (SqlDataAdapter adapter = new SqlDataAdapter(query_check, connectionString))
            {
                DataTable dataTable = new DataTable();
                adapter.Fill(dataTable);

                string test_value = dataTable.Rows[0][0].ToString();
                if ( test_value == "Закрыт" ) 
                {
                    checkBox1.Checked = true;
                }
                else 
                {
                    checkBox1.Checked = false;
                }
            }
        }

        private void button2_Click(object sender, EventArgs e)
        {
            if (String.IsNullOrWhiteSpace(comboBox1.Text))
                MessageBox.Show("Необходимо заполнить все данные!", "Ошибка!");
            else
            {
                if (!isChange)
                {
                    string query = "INSERT INTO Charge (ticket, worker, charge, update_date)" +
                    "VALUES(@ticket, @worker, @charge, @update_date)";
                    using (SqlCommand command = new SqlCommand(query, sqlConnection))
                    {
                        command.Parameters.AddWithValue("@ticket", m_tid.ToString());
                        command.Parameters.AddWithValue("@worker", comboBox1.SelectedValue);
                        command.Parameters.AddWithValue("@charge", richTextBox1.Text);
                        command.Parameters.AddWithValue("@update_date", dateTimePicker1.Value);
                        command.ExecuteNonQuery();
                        this.DialogResult = DialogResult.OK;
                    }

                    int m_Count = 0;
                    string stat = "";

                    string query2 = $"SELECT COUNT(Charge.id) FROM Ticket, Charge WHERE Charge.ticket = Ticket.id AND Ticket.id = {m_tid}";
                    var connectionString = ConfigurationManager.ConnectionStrings["db"].ConnectionString;
                    using (SqlDataAdapter adapter = new SqlDataAdapter(query2, connectionString))
                    {
                        DataTable dataTable = new DataTable();
                        adapter.Fill(dataTable);
                        m_Count = (int)dataTable.Rows[0][0];
                    }

                    string query3 = $"SELECT Ticket.status FROM Ticket, Charge WHERE Charge.ticket = Ticket.id AND Ticket.id = {m_tid}";
                    using (SqlDataAdapter adapter = new SqlDataAdapter(query3, connectionString))
                    {
                        DataTable dataTable = new DataTable();
                        adapter.Fill(dataTable);
                        stat = dataTable.Rows[0][0].ToString();
                    }

                    if ( m_Count >= 1 ) 
                    {
                        if (stat == "Открыт") 
                        {
                            string query_update = $"UPDATE Ticket SET status = @status WHERE id = {m_tid}";
                            using (SqlCommand command = new SqlCommand(query_update, sqlConnection))
                            {
                                command.Parameters.AddWithValue("@status", "В работе");
                                command.ExecuteNonQuery();
                                this.DialogResult = DialogResult.OK;
                            }
                        }
                    }

                    if (checkBox1.Checked) 
                    {
                        string query_close = $"UPDATE Ticket SET status = @status, actual_data = @actual_data WHERE id = {m_tid}";
                        using (SqlCommand command = new SqlCommand(query_close, sqlConnection))
                        {
                            command.Parameters.AddWithValue("@status", "Закрыт");
                            command.Parameters.AddWithValue("@actual_data", dateTimePicker1.Value);
                            command.ExecuteNonQuery();
                            this.DialogResult = DialogResult.OK;
                        }
                    }

                }
                else
                {
                    string query = $"UPDATE Charge SET ticket = @ticket, worker = @worker, charge = @charge, update_date = @update_date WHERE id = {idChange}";
                    using (SqlCommand command = new SqlCommand(query, sqlConnection))
                    {
                        command.Parameters.AddWithValue("@ticket", m_tid.ToString());
                        command.Parameters.AddWithValue("@worker", comboBox1.SelectedValue);
                        command.Parameters.AddWithValue("@charge", richTextBox1.Text);
                        command.Parameters.AddWithValue("@update_date", dateTimePicker1.Value);
                        command.ExecuteNonQuery();
                        this.DialogResult = DialogResult.OK;
                    }

                    if (checkBox1.Checked)
                    {
                        string query_close = $"UPDATE Ticket SET status = @status, actual_data = @actual_data WHERE id = {m_tid}";
                        using (SqlCommand command = new SqlCommand(query_close, sqlConnection))
                        {
                            command.Parameters.AddWithValue("@status", "Закрыт");
                            command.Parameters.AddWithValue("@actual_data", DateTime.Now);
                            command.ExecuteNonQuery();
                            this.DialogResult = DialogResult.OK;
                        }
                    }
                }
            }
        }

        private void iconPictureBox2_Click(object sender, EventArgs e)
        {
            if (new worker_edit().ShowDialog() == DialogResult.OK)
            {
                comboBox1.DataSource = null;
                combobox(comboBox1, "SELECT id, (surname + ' ' + name + ' ' + patronymic) AS [FIO] FROM [Worker]", "FIO", "id");
                MessageBox.Show("Запись успешно добавлена!", "Успех!");
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
    }
}