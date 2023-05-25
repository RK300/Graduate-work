using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Configuration;
using System.Data;
using System.Data.SqlClient;
using System.Drawing;
using System.Drawing.Imaging;
using System.IO;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Controls;
using System.Windows.Forms;

namespace techSupport.Ticket_system
{
    public partial class file_edit : Form
    {
        private SqlConnection sqlConnection = null;
        private int m_tid;

        public file_edit(int tid)
        {
            sqlConnection = new SqlConnection(ConfigurationManager.ConnectionStrings["db"].ConnectionString);
            sqlConnection.Open();
            InitializeComponent();
            SetId(tid);
        }

        private void SetId(int tid) { m_tid = tid; }

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
            string query = $"SELECT * FROM ImageFiles WHERE id = {id}";
            var connectionString = ConfigurationManager.ConnectionStrings["db"].ConnectionString;
            using (SqlDataAdapter adapter = new SqlDataAdapter(query, connectionString))
            {
                DataTable dataTable = new DataTable();
                adapter.Fill(dataTable);
                var imageData = (byte[])dataTable.Rows[0][1];
                using (var memoryStream = new MemoryStream(imageData))
                {
                    pictureBox1.Image = System.Drawing.Image.FromStream(memoryStream);
                }
            }
        }

        private void button2_Click(object sender, EventArgs e)
        {
            if ( pictureBox1.Image == null )
                MessageBox.Show("Необходимо заполнить все данные!", "Ошибка!");
            else
            {
                if (!isChange)
                {
                    DateTime DateNow = DateTime.Now;
                    string query = "INSERT INTO ImageFiles (image, date_upload)" +
                    "VALUES(@image, @date_upload)";
                    using (SqlCommand command = new SqlCommand(query, sqlConnection))
                    {
                        var upload_image = new Bitmap(pictureBox1.Image);
                        using (var memoryStream = new MemoryStream())
                        {
                            upload_image.Save(memoryStream, ImageFormat.Jpeg);
                            memoryStream.Position = 0;

                            var sqlParameter = new SqlParameter("@image", SqlDbType.VarBinary, (int)memoryStream.Length)
                            {
                                Value = memoryStream.ToArray()
                            };
                            command.Parameters.Add(sqlParameter);
                            command.Parameters.AddWithValue("@date_upload", DateNow);
                        }
                        command.ExecuteNonQuery();
                    }
                    
                    int m_photoID;

                    string select_query = $"SELECT * FROM ImageFiles ORDER BY ID DESC";
                    var connectionString = ConfigurationManager.ConnectionStrings["db"].ConnectionString;
                    using (SqlDataAdapter adapter = new SqlDataAdapter(select_query, connectionString))
                    {
                        DataTable dataTable = new DataTable();
                        adapter.Fill(dataTable);
                        m_photoID = (int)dataTable.Rows[0][0];
                    }

                    string query2 = "INSERT INTO File2Tiket (ticket, photo)" +
                    "VALUES(@ticket, @photo)";
                    using (SqlCommand command = new SqlCommand(query2, sqlConnection))
                    {
                        command.Parameters.AddWithValue("@ticket", m_tid);
                        command.Parameters.AddWithValue("@photo", m_photoID);
                        command.ExecuteNonQuery();
                        this.DialogResult = DialogResult.OK;
                    }
                }
                else
                {
                    string query = $"UPDATE ImageFiles SET image = @image, date_upload = @date_upload WHERE id = {idChange}";
                    using (SqlCommand command = new SqlCommand(query, sqlConnection))
                    {
                        var upload_image = new Bitmap(pictureBox1.Image);
                        using (var memoryStream = new MemoryStream())
                        {
                            upload_image.Save(memoryStream, ImageFormat.Jpeg);
                            memoryStream.Position = 0;

                            var sqlParameter = new SqlParameter("@image", SqlDbType.VarBinary, (int)memoryStream.Length)
                            {
                                Value = memoryStream.ToArray()
                            };
                            command.Parameters.Add(sqlParameter);
                            command.Parameters.AddWithValue("@date_upload", DateTime.Now);
                        }
                        command.ExecuteNonQuery();
                        this.DialogResult = DialogResult.OK;
                    }
                }
            }
        }

        private void button1_Click(object sender, EventArgs e)
        {
            openFileDialog1.Filter = "Image files (*.bmp, *.jpg, *.gif, *.png)|*.bmp;*.jpg;*.gif;*.png";

            if (openFileDialog1.ShowDialog() == DialogResult.Cancel)
                return;

            string path = openFileDialog1.FileName;
            pictureBox1.Image = System.Drawing.Image.FromFile(path);
        }
    }
}