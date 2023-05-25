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

namespace techSupport.view_form
{
    public partial class client_view : Form
    {
        private SqlConnection sqlConnection = null;

        private int m_id;

        public client_view(int id)
        {
            sqlConnection = new SqlConnection(ConfigurationManager.ConnectionStrings["db"].ConnectionString);
            sqlConnection.Open();
            InitializeComponent();
            setId(id);
            setName();
            setInfo();
            setProduct();
            setBank();
        }

        private void setId(int id) { m_id = id; }

        private void setProduct() 
        {
            string query = $"SELECT Products.id, Products.name AS [Название продукта] FROM Products, User2Product, Clients WHERE User2Product.client = Clients.id AND User2Product.product = Products.id AND Clients.id = '{m_id}'";
            var connectionString = ConfigurationManager.ConnectionStrings["db"].ConnectionString;
            using (SqlDataAdapter adapter = new SqlDataAdapter(query, connectionString))
            {
                DataTable dataTable = new DataTable();
                adapter.Fill(dataTable);
                dataGridView1.DataSource = dataTable;
            }
            dataGridView1.Columns[0].Visible = false;
        }

        private void setInfo() 
        {
            string query = $"SELECT (Clients.surname + ' ' + Clients.name + ' ' + Clients.patronymic), Clients.email, (Locality.name + ', ' + District.name + ' район, ' + Region.name + ' область, ' + Country.name + '.'), Clients.phone, Clients.Checking, Clients.BIK, Clients.YNP FROM Clients, Locality, District, Region, Country WHERE Clients.city = Locality.id AND Locality.district = District.id AND District.region = Region.id AND Region.country = Country.id AND Clients.id = '{m_id}'";
            SqlDataAdapter adapter = new SqlDataAdapter();
            SqlCommand command = new SqlCommand(query, sqlConnection);
            DataTable tb = new DataTable();
            adapter.SelectCommand = command;
            adapter.Fill(tb);
            label4.Text = "Представитель: " + tb.Rows[0][0].ToString();
            label3.Text = "Адрес: " + tb.Rows[0][2].ToString();
            label1.Text = "Электронная почта: " + tb.Rows[0][1].ToString();
            label6.Text = "Номер телефона: " + tb.Rows[0][3].ToString();
            label7.Text = "Расчётный счёт: " + tb.Rows[0][4].ToString();
            label8.Text = "БИК: " + tb.Rows[0][5].ToString();
            label9.Text = "УНП: " + tb.Rows[0][6].ToString();
        }

        private void setBank()
        {
            string query = $"SELECT Bank.name, (Locality.name + ', ул.' + Bank.street + ', д.' + Bank.house + ', к.' + Bank.corpse) FROM Bank, Clients, Locality WHERE Bank.id = Clients.bank AND Bank.location = Locality.id AND Clients.id = '{m_id}'";
            SqlDataAdapter adapter = new SqlDataAdapter();
            SqlCommand command = new SqlCommand(query, sqlConnection);
            DataTable tb = new DataTable();
            adapter.SelectCommand = command;
            adapter.Fill(tb);

            label5.Text = "Банк: " + tb.Rows[0][0].ToString() + ", " + tb.Rows[0][1].ToString();
        }

        private void setName() 
        {
            string query = $"SELECT CompanyName FROM Clients WHERE id = '{m_id}'";
            SqlDataAdapter adapter = new SqlDataAdapter();
            SqlCommand command = new SqlCommand(query, sqlConnection);
            DataTable tb = new DataTable();
            adapter.SelectCommand = command;
            adapter.Fill(tb);
            this.Text = "Просмотр клиента: " + tb.Rows[0][0].ToString();
            label10.Text = "Название компании: " + tb.Rows[0][0].ToString();
        }

        private void dataGridView1_SelectionChanged(object sender, EventArgs e)
        {
            if (dataGridView1.CurrentCell != null)
            {
                int client_id = m_id;
                int prod_id = (int)dataGridView1.Rows[dataGridView1.CurrentCell.RowIndex].Cells[0].Value;

                string query = $"SELECT Treaty.id, ('Договор' + ' №' + convert(nvarchar(max), Treaty.nomer, 0) + ' ' + '|' + ' ' + convert(nvarchar(max), Treaty.dateСonclusion, 0) + ' ' + '(' + convert(nvarchar(max), Treaty.dataFrom, 0) + ' - ' + convert(nvarchar(max), Treaty.dateTo, 0) + ')') AS [Договор] FROM Treaty, Clients, Products, User2Product WHERE Treaty.client = Clients.id AND Treaty.product = Products.id AND User2Product.client = Treaty.client AND User2Product.product = Treaty.product AND Clients.id = {client_id} AND Products.id = {prod_id}";
                var connectionString = ConfigurationManager.ConnectionStrings["db"].ConnectionString;
                using (SqlDataAdapter adapter = new SqlDataAdapter(query, connectionString))
                {
                    DataTable dataTable = new DataTable();
                    adapter.Fill(dataTable);
                    dataGridView2.DataSource = dataTable;
                }
                dataGridView2.Columns[0].Visible = false;
            }
        }
    }
}
