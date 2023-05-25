using System;
using System.Configuration;
using System.Data;
using System.Data.SqlClient;
using System.Runtime.InteropServices;
using System.Windows.Forms;
using techSupport.mainmenu;
using md5_sql_hash;

namespace techSupport
{
    public partial class auth : Form
    {
        private SqlConnection sqlConnection = null;

        public auth()
        {
            sqlConnection = new SqlConnection(ConfigurationManager.ConnectionStrings["db"].ConnectionString);
            sqlConnection.Open();
            InitializeComponent();
        }

        private void button1_Click(object sender, EventArgs e)
        {
            /*register r = new register();
            r.Show();
            this.Hide();*/
            Application.Exit();
        }

        private void iconPictureBox1_Click(object sender, EventArgs e)
        {
            Application.Exit();
        }

        private void iconPictureBox2_Click(object sender, EventArgs e)
        {
            this.WindowState = FormWindowState.Minimized;
        }
        
        private void button2_Click(object sender, EventArgs e)
        {
            string m_loginUser, m_loginPassword;
            if ( String.IsNullOrWhiteSpace(login_textBox.Text) || String.IsNullOrWhiteSpace(password_textBox.Text )) 
            {
                MessageBox.Show("Необходимо заполнить все данные!", "Ошибка!");
            }
            else 
            { 
                m_loginUser = login_textBox.Text;
                m_loginPassword = password_textBox.Text;
                var pass = md5.hashPassword(m_loginPassword);

                DataTable tb = new DataTable();

                string query = $"SELECT id, login, password FROM Worker WHERE login = '{m_loginUser}' AND password = '{pass}'";
                SqlDataAdapter adapter = new SqlDataAdapter();
                SqlCommand command = new SqlCommand(query, sqlConnection);
                adapter.SelectCommand = command;
                adapter.Fill(tb);

                if (tb.Rows.Count == 1) 
                {
                    mainmenu_form t = new mainmenu_form((int)tb.Rows[0][0]);
                    t.Show();
                    this.Hide();
                    MessageBox.Show("Вы успешно вошли!", "Успех!");
                }
                else 
                {
                    MessageBox.Show("Такого аккаунта не существует! Или произошла ошибка!", "Ошибка!");
                }
            }
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
    }
}
