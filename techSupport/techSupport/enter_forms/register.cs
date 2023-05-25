using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Configuration;
using System.Data;
using System.Data.SqlClient;
using System.Drawing;
using System.Linq;
using System.Runtime.InteropServices;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;

namespace techSupport.enter_forms
{
    public partial class register : Form
    {
        private SqlConnection sqlConnection = null;

        public register()
        {
            sqlConnection = new SqlConnection(ConfigurationManager.ConnectionStrings["db"].ConnectionString);
            sqlConnection.Open();
            InitializeComponent();
            combobox(comboBox1, "SELECT id, name FROM [Position]", "name", "id");
        }

        private void button1_Click(object sender, EventArgs e)
        {
            auth auth = new auth();
            auth.Show();
            this.Close();
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

        private void iconPictureBox1_Click(object sender, EventArgs e)
        {
            Application.Exit();
        }

        private void iconPictureBox2_Click(object sender, EventArgs e)
        {
            this.WindowState = FormWindowState.Minimized;
        }

        [DllImport("user32.DLL", EntryPoint = "ReleaseCapture")]
        private extern static void ReleaseCapture();
        [DllImport("user32.DLL", EntryPoint = "SendMessage")]
        private extern static void SendMessage(System.IntPtr hWnd, int wMsg, int wParam, int lParam);

        private void panel1_MouseDown(object sender, MouseEventArgs e)
        {
            ReleaseCapture();
            SendMessage(this.Handle, 0x112, 0xf012, 0);
        }

        private void button2_Click(object sender, EventArgs e)
        {
            if (String.IsNullOrWhiteSpace(comboBox1.Text) || String.IsNullOrWhiteSpace(login_textBox.Text) || String.IsNullOrWhiteSpace(textBox1.Text) || String.IsNullOrWhiteSpace(textBox2.Text) || String.IsNullOrWhiteSpace(textBox3.Text) || String.IsNullOrWhiteSpace(textBox4.Text) || String.IsNullOrWhiteSpace(textBox6.Text) || String.IsNullOrWhiteSpace(textBox7.Text)) 
            {
                MessageBox.Show("Необходимо заполнить все данные!", "Ошибка!");
            }
            else 
            {
                string query = "INSERT INTO Worker (surname, name, patronymic, position, email, phone, login, password)" +
                "VALUES(@surname, @name, @patronymic, @position, @email, @phone, @login, @password)";
                using (SqlCommand command = new SqlCommand(query, sqlConnection))
                {
                    command.Parameters.AddWithValue("@surname", login_textBox.Text);
                    command.Parameters.AddWithValue("@name", textBox1.Text);
                    command.Parameters.AddWithValue("@patronymic", textBox2.Text);
                    command.Parameters.AddWithValue("@position", comboBox1.SelectedValue);
                    command.Parameters.AddWithValue("@email", textBox4.Text);
                    command.Parameters.AddWithValue("@phone", textBox3.Text);
                    command.Parameters.AddWithValue("@login", textBox7.Text);
                    command.Parameters.AddWithValue("@password", textBox6.Text);
                    command.ExecuteNonQuery();
                }
                MessageBox.Show("Вы успешно зарегистрировались!", "Успех!");
                auth auth = new auth();
                auth.Show();
                this.Hide();
            }
        }
    }
}
