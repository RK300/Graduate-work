using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Configuration;
using System.Data;
using System.Data.SqlClient;
using System.Drawing;
using System.IO;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;

namespace techSupport.view_form
{
    public partial class photo_view : Form
    {
        private SqlConnection sqlConnection = null;

        private int m_Id;

        private void setId(int id) { m_Id = id; }

        public photo_view(int id)
        {
            sqlConnection = new SqlConnection(ConfigurationManager.ConnectionStrings["db"].ConnectionString);
            sqlConnection.Open();
            setId(id);
            InitializeComponent();
            loadPhoto();
        }

        private void loadPhoto() 
        {
            string query = $"SELECT * FROM ImageFiles WHERE id = {m_Id}";
            var connectionString = ConfigurationManager.ConnectionStrings["db"].ConnectionString;
            using (SqlDataAdapter adapter = new SqlDataAdapter(query, connectionString))
            {
                DataTable dataTable = new DataTable();
                adapter.Fill(dataTable);
                var imageData = (byte[])dataTable.Rows[0][1];
                using (var memoryStream = new MemoryStream(imageData))
                {
                    pictureBox1.Image = Image.FromStream(memoryStream);
                }
            }
        }
    }
}
