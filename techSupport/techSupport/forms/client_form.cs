using Newtonsoft.Json;
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
using techSupport.edit_form;
using techSupport.edit_forms_two;
using techSupport.new_forms;
using techSupport.view_form;
using SettingsClass;
using WordHelperClass;

namespace techSupport.forms
{
    public partial class client_form : Form
    {
        private SqlConnection sqlConnection = null;

        public client_form()
        {
            sqlConnection = new SqlConnection(ConfigurationManager.ConnectionStrings["db"].ConnectionString);
            sqlConnection.Open();
            InitializeComponent();
            RefreshTable();
        }

        private void RefreshTable()
        {
            //string query = "SELECT Clients.id, Clients.CompanyName AS [Название компании], Clients.surname AS [Фамилия], Clients.name AS [Имя], Clients.patronymic AS [Отчество], Bank.Name AS [Банк], Locality.name AS [Населенный пункт], Clients.email AS [Электронная почта], phone AS [Номер телефона], Checking AS [Расчетный счёт], Clients.BIK AS [БИК], YNP AS [УНП], Clients.street AS [Улица], Clients.house AS [Дом], Clients.corpus AS [Корпус] FROM Clients, Locality, Bank WHERE Clients.city = Locality.id AND Clients.bank = Bank.id";
            //string query = "SELECT Clients.id, Clients.CompanyName AS [Название компании], Clients.surname AS [Фамилия], Clients.name AS [Имя], Clients.patronymic AS [Отчество], Bank.Name AS [Банк], Locality.name AS [Населенный пункт], Clients.email AS [Электронная почта], phone AS [Номер телефона], Checking AS [Расчетный счёт], Clients.BIK AS [БИК], YNP AS [УНП] FROM Clients, Locality, Bank WHERE Clients.city = Locality.id AND Clients.bank = Bank.id";
            string query = "SELECT Clients.id, Clients.CompanyName AS [Название компании], Clients.surname AS [Фамилия], Clients.name AS [Имя], Clients.patronymic AS [Отчество], Bank.Name AS [Банк], Locality.name AS [Населенный пункт], Clients.email AS [Электронная почта], phone AS [Номер телефона] FROM Clients, Locality, Bank WHERE Clients.city = Locality.id AND Clients.bank = Bank.id";
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
            if (new client_edit().ShowDialog() == DialogResult.OK)
            {
                RefreshTable();
                MessageBox.Show("Запись успешно добавлена!", "Успех!");

                if ((dataGridView1.Rows.Count - 1) < 0)
                {
                    dataGridView1.CurrentCell = dataGridView1.Rows[0].Cells[1];
                }
                else
                {
                    dataGridView1.CurrentCell = dataGridView1.Rows[dataGridView1.Rows.Count - 1].Cells[1];
                    SubRefresh();
                }
            }
        }

        private void rjButton2_Click(object sender, EventArgs e)
        {
            client_edit F2 = new client_edit();
            F2.Text = "Редактирование клиента";
            F2.IDCHANGE = dataGridView1.Rows[dataGridView1.CurrentCell.RowIndex].Cells[0].Value.ToString();
            var m_curentCell = dataGridView1.CurrentCell.RowIndex;

            if (F2.ShowDialog() == DialogResult.OK)
            {
                RefreshTable();
                MessageBox.Show("Запись успешно изменена!", "Успех!");
                if ((dataGridView1.Rows.Count - 1) < 0)
                {
                    dataGridView1.CurrentCell = dataGridView1.Rows[0].Cells[1];
                }
                else
                {
                    dataGridView1.CurrentCell = dataGridView1.Rows[m_curentCell].Cells[1];
                    SubRefresh();
                }
            }
        }

        private void rjButton3_Click(object sender, EventArgs e)
        {
            if (MessageBox.Show("Вы действительно хотите удалить запись?", "Подтверждение удаления", MessageBoxButtons.YesNo, MessageBoxIcon.Warning) == DialogResult.Yes)
            {
                try
                {
                    SqlCommand command = new SqlCommand($"DELETE FROM Clients WHERE id = {dataGridView1.Rows[dataGridView1.CurrentCell.RowIndex].Cells[0].Value}", sqlConnection);
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

        private void rjButton5_Click(object sender, EventArgs e)
        {
            if (!String.IsNullOrWhiteSpace(textBox1.Text))
            {
                //string query = $"SELECT Clients.id, Clients.CompanyName AS [Название компании], Clients.surname AS [Фамилия], Clients.name AS [Имя], Clients.patronymic AS [Отчество], Bank.Name AS [Банк], Locality.name AS [Населенный пункт], Clients.email AS [Электронная почта], phone AS [Номер телефона], Checking AS [Расчетный счёт], Clients.BIK AS [БИК], YNP AS [УНП], Clients.street AS [Улица], Clients.house AS [Дом], Clients.corpus AS [Корпус] FROM Clients, Locality, Bank WHERE Clients.bank = Bank.id AND Clients.city = Locality.id AND Clients.surname + ' ' + Clients.name + ' ' + Clients.patronymic + ' ' + Clients.email + ' ' + Clients.phone + ' ' + Clients.Checking + ' ' + Clients.BIK + ' ' + Clients.YNP + ' ' + Clients.CompanyName + ' ' + Bank.name + ' ' + Clients.corpus + ' ' + Clients.house + ' ' + Clients.street LIKE '%{textBox1.Text}%'";
                //string query = $"SELECT Clients.id, Clients.CompanyName AS [Название компании], Clients.surname AS [Фамилия], Clients.name AS [Имя], Clients.patronymic AS [Отчество], Bank.Name AS [Банк], Locality.name AS [Населенный пункт], Clients.email AS [Электронная почта], phone AS [Номер телефона], Checking AS [Расчетный счёт], Clients.BIK AS [БИК], YNP AS [УНП] FROM Clients, Locality, Bank WHERE Clients.bank = Bank.id AND Clients.city = Locality.id AND Clients.surname + ' ' + Clients.name + ' ' + Clients.patronymic + ' ' + Clients.email + ' ' + Clients.phone + ' ' + Clients.Checking + ' ' + Clients.BIK + ' ' + Clients.YNP + ' ' + Clients.CompanyName + ' ' + Bank.name + ' ' + Clients.corpus + ' ' + Clients.house + ' ' + Clients.street LIKE '%{textBox1.Text}%'";
                string query = $"SELECT Clients.id, Clients.CompanyName AS [Название компании], Clients.surname AS [Фамилия], Clients.name AS [Имя], Clients.patronymic AS [Отчество], Bank.Name AS [Банк], Locality.name AS [Населенный пункт], Clients.email AS [Электронная почта], phone AS [Номер телефона] FROM Clients, Locality, Bank WHERE Clients.bank = Bank.id AND Clients.city = Locality.id AND Clients.surname + ' ' + Clients.name + ' ' + Clients.patronymic + ' ' + Clients.email + ' ' + Clients.phone + ' ' + Clients.Checking + ' ' + Clients.BIK + ' ' + Clients.YNP + ' ' + Clients.CompanyName + ' ' + Bank.name + ' ' + Clients.corpus + ' ' + Clients.house + ' ' + Clients.street LIKE '%{textBox1.Text}%'";
                var connectionString = ConfigurationManager.ConnectionStrings["db"].ConnectionString;
                using (SqlDataAdapter adapter = new SqlDataAdapter(query, connectionString))
                {
                    DataTable dataTable = new DataTable();
                    adapter.Fill(dataTable);
                    dataGridView1.DataSource = dataTable;
                }
            }
            else
            {
                RefreshTable();
            }
        }

        private void rjButton4_Click(object sender, EventArgs e)
        {
            int id = (int)dataGridView1.Rows[dataGridView1.CurrentCell.RowIndex].Cells[0].Value;
            client_view c = new client_view(id);
            c.Show();
        }

        private void dataGridView1_SelectionChanged(object sender, EventArgs e)
        {
            string query = $"SELECT User2Product.id, Products.name AS [Название продукта] FROM Products, User2Product, Clients WHERE User2Product.client = Clients.id AND User2Product.product = Products.id AND Clients.id = '{dataGridView1.Rows[dataGridView1.CurrentCell.RowIndex].Cells[0].Value.ToString()}'";
            var connectionString = ConfigurationManager.ConnectionStrings["db"].ConnectionString;
            using (SqlDataAdapter adapter = new SqlDataAdapter(query, connectionString))
            {
                DataTable dataTable = new DataTable();
                adapter.Fill(dataTable);
                dataGridView2.DataSource = dataTable;
            }
            dataGridView2.Columns[0].Visible = false;
        }

        private void rjButton6_Click(object sender, EventArgs e)
        {
            if (MessageBox.Show("Вы действительно хотите удалить запись?", "Подтверждение удаления", MessageBoxButtons.YesNo, MessageBoxIcon.Warning) == DialogResult.Yes)
            {
                SqlCommand command = new SqlCommand($"DELETE FROM User2Product WHERE id = {dataGridView2.Rows[dataGridView2.CurrentCell.RowIndex].Cells[0].Value}", sqlConnection);
                try 
                { 
                    command.ExecuteNonQuery(); 
                }
                catch 
                {
                    MessageBox.Show("Ошибка, попробуйте еще раз!", "Внимание", MessageBoxButtons.OK, MessageBoxIcon.Error);
                    return;
                }
                MessageBox.Show("Запись успешно удалена!", "Успех!");
                RefreshTable();
            }
        }

        private void rjButton8_Click(object sender, EventArgs e)
        {
            if (new product_edit_client((int)dataGridView1.Rows[dataGridView1.CurrentCell.RowIndex].Cells[0].Value).ShowDialog() == DialogResult.OK)
            {
                var m_currentIndex = dataGridView1.CurrentCell.RowIndex;
                RefreshTable();
                MessageBox.Show("Запись успешно добавлена!", "Успех!");

                if ((dataGridView1.Rows.Count - 1) < 0)
                {
                    dataGridView1.CurrentCell = dataGridView1.Rows[0].Cells[1];
                }
                else
                {
                    dataGridView1.CurrentCell = dataGridView1.Rows[m_currentIndex].Cells[1];
                    SubRefresh();

                    if ((dataGridView2.Rows.Count - 1) < 0)
                    {
                        dataGridView2.CurrentCell = dataGridView2.Rows[0].Cells[1];
                    }
                    else
                    {
                        dataGridView2.CurrentCell = dataGridView2.Rows[dataGridView2.Rows.Count - 1].Cells[1];
                        DogovorRefresh();
                    }
                }
            }
        }

        private void rjButton7_Click(object sender, EventArgs e)
        {
            product_edit_client F2 = new product_edit_client((int)dataGridView1.Rows[dataGridView1.CurrentCell.RowIndex].Cells[0].Value);
            F2.Text = "Редактирование продукта";
            F2.IDCHANGE = dataGridView2.Rows[dataGridView2.CurrentCell.RowIndex].Cells[0].Value.ToString();
            if (F2.ShowDialog() == DialogResult.OK)
            {
                var m_currentIndex = dataGridView1.CurrentCell.RowIndex;
                var s_currentIndex = dataGridView2.CurrentCell.RowIndex;
                RefreshTable();
                MessageBox.Show("Запись успешно изменена!", "Успех!");

                if ((dataGridView1.Rows.Count - 1) < 0)
                {
                    dataGridView1.CurrentCell = dataGridView1.Rows[0].Cells[1];
                }
                else
                {
                    dataGridView1.CurrentCell = dataGridView1.Rows[m_currentIndex].Cells[1];
                    SubRefresh();

                    if ((dataGridView2.Rows.Count - 1) < 0)
                    {
                        dataGridView2.CurrentCell = dataGridView2.Rows[0].Cells[1];
                    }
                    else
                    {
                        dataGridView2.CurrentCell = dataGridView2.Rows[s_currentIndex].Cells[1];
                        DogovorRefresh();
                    }
                }
            }
        }

        private void dataGridView2_SelectionChanged(object sender, EventArgs e)
        {
            int client_id = (int)dataGridView1.Rows[dataGridView1.CurrentCell.RowIndex].Cells[0].Value;

            int user2prod_id;

            if ( dataGridView2.CurrentCell != null ) 
            {
                user2prod_id = (int)dataGridView2.Rows[dataGridView2.CurrentCell.RowIndex].Cells[0].Value;
                int prod_id;

                string query2 = $"SELECT * FROM User2Product WHERE id = {user2prod_id}";
                var connectionString = ConfigurationManager.ConnectionStrings["db"].ConnectionString;
                using (SqlDataAdapter adapter = new SqlDataAdapter(query2, connectionString))
                {
                    DataTable dataTable = new DataTable();
                    adapter.Fill(dataTable);

                    if (dataTable.Rows.Count == 0)
                    {
                        return;
                    }

                    prod_id = (int)dataTable.Rows[0][2];
                }

                //string query = $"SELECT Treaty.id, ('Договор' + ' №' + convert(nvarchar(max), Treaty.nomer, 0) + ' ' + '|' + ' ' + convert(nvarchar(max), Treaty.dateСonclusion, 0) + ' ' + '(' + convert(nvarchar(max), Treaty.dataFrom, 0) + ' - ' + convert(nvarchar(max), Treaty.dateTo, 0) + ')') AS [Договор] FROM Treaty, Clients, Products, User2Product WHERE Treaty.client = Clients.id AND Treaty.product = Products.id AND User2Product.client = Treaty.client AND User2Product.product = Treaty.product AND Clients.id = {client_id} AND Products.id = {prod_id}";
                string query = $"SELECT Treaty.id, ('Договор' + ' №' + convert(nvarchar(max), Treaty.nomer, 0) + ' ' + '|' + ' ' + convert(nvarchar(max), Treaty.dateСonclusion, 0) + ' ' + '(' + convert(nvarchar(max), Treaty.dataFrom, 0) + ' - ' + convert(nvarchar(max), Treaty.dateTo, 0) + ')') AS [Договор] FROM Treaty, Clients, Products, User2Product WHERE Treaty.client = Clients.id AND Treaty.product = Products.id AND User2Product.client = Treaty.client AND User2Product.product = Treaty.product AND Clients.id = {client_id} AND Products.id = {prod_id} ORDER BY Treaty.id DESC";
                using (SqlDataAdapter adapter = new SqlDataAdapter(query, connectionString))
                {
                    DataTable dataTable = new DataTable();
                    adapter.Fill(dataTable);
                    dataGridView3.DataSource = dataTable;
                }
                dataGridView3.Columns[0].Visible = false;
            }
        }

        private void rjButton11_Click(object sender, EventArgs e)
        {
            int client_id = (int)dataGridView1.Rows[dataGridView1.CurrentCell.RowIndex].Cells[0].Value;
            int user2prod_id = (int)dataGridView2.Rows[dataGridView2.CurrentCell.RowIndex].Cells[0].Value;
            int prod_id;

            string query2 = $"SELECT * FROM User2Product WHERE id = {user2prod_id}";
            var connectionString = ConfigurationManager.ConnectionStrings["db"].ConnectionString;
            using (SqlDataAdapter adapter = new SqlDataAdapter(query2, connectionString))
            {
                DataTable dataTable = new DataTable();
                adapter.Fill(dataTable);
                prod_id = (int)dataTable.Rows[0][2];
            }

            if (new dogovor2_edit(client_id, prod_id).ShowDialog() == DialogResult.OK)
            {
                var m_currentIndex = dataGridView1.CurrentCell.RowIndex;
                var s_currentIndex = dataGridView2.CurrentCell.RowIndex;
                RefreshTable();
                MessageBox.Show("Продукт успешно добавлен!", "Успех!");
                if ((dataGridView1.Rows.Count - 1) < 0)
                {
                    dataGridView1.CurrentCell = dataGridView1.Rows[0].Cells[1];
                }
                else
                {
                    dataGridView1.CurrentCell = dataGridView1.Rows[m_currentIndex].Cells[1];
                    SubRefresh();

                    if ((dataGridView2.Rows.Count - 1) < 0)
                    {
                        dataGridView2.CurrentCell = dataGridView2.Rows[0].Cells[1];
                    }
                    else
                    {
                        dataGridView2.CurrentCell = dataGridView2.Rows[dataGridView2.Rows.Count - 1].Cells[1];
                        DogovorRefresh();

                        if ((dataGridView3.Rows.Count - 1) < 0)
                        {
                            dataGridView3.CurrentCell = dataGridView3.Rows[0].Cells[1];
                        }
                        else
                        {
                            dataGridView3.CurrentCell = dataGridView3.Rows[0].Cells[1];
                        }
                    }
                }
            }
        }

        private void rjButton9_Click(object sender, EventArgs e)
        {
            if (MessageBox.Show("Вы действительно хотите удалить запись?", "Подтверждение удаления", MessageBoxButtons.YesNo, MessageBoxIcon.Warning) == DialogResult.Yes)
            {
                try
                {
                    SqlCommand command = new SqlCommand($"DELETE FROM Treaty WHERE id = {dataGridView3.Rows[dataGridView3.CurrentCell.RowIndex].Cells[0].Value}", sqlConnection);
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

        private void rjButton10_Click(object sender, EventArgs e)
        {
            int client_id = (int)dataGridView1.Rows[dataGridView1.CurrentCell.RowIndex].Cells[0].Value;
            int user2prod_id = (int)dataGridView2.Rows[dataGridView2.CurrentCell.RowIndex].Cells[0].Value;
            int prod_id;

            string query2 = $"SELECT * FROM User2Product WHERE id = {user2prod_id}";
            var connectionString = ConfigurationManager.ConnectionStrings["db"].ConnectionString;
            using (SqlDataAdapter adapter = new SqlDataAdapter(query2, connectionString))
            {
                DataTable dataTable = new DataTable();
                adapter.Fill(dataTable);
                prod_id = (int)dataTable.Rows[0][2];
            }

            dogovor2_edit F2 = new dogovor2_edit(client_id, prod_id);
            F2.Text = "Редактирование договора";
            F2.IDCHANGE = dataGridView3.Rows[dataGridView3.CurrentCell.RowIndex].Cells[0].Value.ToString();
            if (F2.ShowDialog() == DialogResult.OK)
            {
                var m_currentIndex = dataGridView1.CurrentCell.RowIndex;
                var s_currentIndex = dataGridView2.CurrentCell.RowIndex;
                var x_currentIndex = dataGridView3.CurrentCell.RowIndex;
                RefreshTable();
                MessageBox.Show("Запись успешно изменена!", "Успех!");
                if ((dataGridView1.Rows.Count - 1) < 0)
                {
                    dataGridView1.CurrentCell = dataGridView1.Rows[0].Cells[1];
                }
                else
                {
                    dataGridView1.CurrentCell = dataGridView1.Rows[m_currentIndex].Cells[1];
                    SubRefresh();

                    if ((dataGridView2.Rows.Count - 1) < 0)
                    {
                        dataGridView2.CurrentCell = dataGridView2.Rows[0].Cells[1];
                    }
                    else
                    {
                        dataGridView2.CurrentCell = dataGridView2.Rows[dataGridView2.Rows.Count - 1].Cells[1];
                        DogovorRefresh();

                        if ((dataGridView3.Rows.Count - 1) < 0)
                        {
                            dataGridView3.CurrentCell = dataGridView3.Rows[0].Cells[1];
                        }
                        else
                        {
                            dataGridView3.CurrentCell = dataGridView3.Rows[x_currentIndex].Cells[1];
                        }
                    }
                }
            }
        }

        private void dataGridView1_SelectionChanged_1(object sender, EventArgs e)
        {
            if ( dataGridView1.CurrentCell != null ) 
            {
                string query = $"SELECT User2Product.id, Products.name AS [Название продукта] FROM Products, User2Product, Clients WHERE User2Product.client = Clients.id AND User2Product.product = Products.id AND Clients.id = '{dataGridView1.Rows[dataGridView1.CurrentCell.RowIndex].Cells[0].Value.ToString()}'";
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

        private void rjButton12_Click(object sender, EventArgs e)
        {
            var settings = File.Exists("settings.json") ? JsonConvert.DeserializeObject<Settings>(File.ReadAllText("settings.json")) : new Settings
            {
                FIO = "Иванов Иван Иванович",
                NazvComp = "УП 'ИВЦ-ТЕСТ'",
                RascSchet = "BY12BAPB25125674500100000000",
                Bank = "ОАО «Белагропромбанк» г.Молодечно BAPBBY2X",
                Adress = "г. Молодечно",
                YNP = "600021518",
                OKPO = "05552555"
            };

            File.WriteAllText("settings.json", JsonConvert.SerializeObject(settings));

            int v_id = (int)dataGridView3.Rows[dataGridView3.CurrentCell.RowIndex].Cells[0].Value;

            string ProductName;
            string c_FIO;
            string c_NazvComp, c_RascSchet, c_Bank, c_BIK, c_Adress, c_UNP;
            string Nomer, Data_zak, DataFrom, DataTo, Price;

            string query = $"SELECT (Clients.surname + ' ' + Clients.name + ' ' + Clients.patronymic) AS [FIO], Clients.CompanyName, Clients.Checking, Bank.name, Clients.BIK, (Locality.name + ', ' + Clients.street + ', ' + Clients.house + ', ' + Clients.corpus), Clients.YNP FROM Clients, Bank, Locality, Treaty WHERE Treaty.client = Clients.id AND Clients.city = Locality.id AND Clients.bank = Bank.id AND Treaty.id = {v_id}";
            var connectionString = ConfigurationManager.ConnectionStrings["db"].ConnectionString;
            using (SqlDataAdapter adapter = new SqlDataAdapter(query, connectionString))
            {
                DataTable dataTable = new DataTable();
                adapter.Fill(dataTable);
                c_FIO = dataTable.Rows[0][0].ToString();
                c_NazvComp = dataTable.Rows[0][1].ToString();
                c_RascSchet = dataTable.Rows[0][2].ToString();
                c_Bank = dataTable.Rows[0][3].ToString();
                c_BIK = dataTable.Rows[0][4].ToString();
                c_Adress = dataTable.Rows[0][5].ToString();
                c_UNP = dataTable.Rows[0][6].ToString();
            }

            string query2 = $"SELECT * FROM Treaty WHERE Treaty.id = {v_id}";
            using (SqlDataAdapter adapter = new SqlDataAdapter(query2, connectionString))
            {
                DataTable dataTable = new DataTable();
                adapter.Fill(dataTable);

                Nomer = dataTable.Rows[0][2].ToString();

                DateTime dateTime = DateTime.Now;
                dateTime = (DateTime)dataTable.Rows[0][4];
                Data_zak = dateTime.ToString("D");
                dateTime = (DateTime)dataTable.Rows[0][5];
                DataFrom = dateTime.ToString("D");
                dateTime = (DateTime)dataTable.Rows[0][6];
                DataTo = dateTime.ToString("D");

                //Data_zak = dataTable.Rows[0][3].ToString();
                //DataFrom = dataTable.Rows[0][4].ToString();
                //DataTo = dataTable.Rows[0][5].ToString();

            }

            string query3 = $"SELECT Products.name FROM Treaty, Products WHERE Treaty.product = Products.id AND Treaty.id = {v_id}";
            using (SqlDataAdapter adapter = new SqlDataAdapter(query3, connectionString))
            {
                DataTable dataTable = new DataTable();
                adapter.Fill(dataTable);

                ProductName = dataTable.Rows[0][0].ToString();
            }

            string c_Initsal = GetInitials(c_FIO);
            string p_Initsal = GetInitials(settings.FIO);

            var helper = new WordHelper("blank.doc");

            var items = new Dictionary<string, string>
            {
                { "<CITY>", GetCityFromAddress(settings.Adress) },
                { "<INIT_CLIENT>", c_Initsal },
                { "<INIT_PROD>", p_Initsal },
                { "<NAZV_PO>", ProductName },
                { "<FIO_PROD_2>", p_Initsal },
                { "<NUM>", Nomer },
                { "<DATA_ZAK>", Data_zak },
                { "<NAME_CORP_PROD>", settings.NazvComp },
                { "<NAME_COMP_PROD>", settings.NazvComp },
                { "<FIO_CLIENT>", c_FIO },
                { "<DATE_FROM>",  DataFrom},
                { "<DATE_TO>",  DataTo},
                { "<RASH_CHET>", settings.RascSchet },
                { "<BANK_PROD>",  settings.Bank },
                { "<ADRESS_PROD>", settings.Adress },
                { "<YNP>", settings.YNP },
                { "<OKPO>", settings.OKPO },
                { "<NAZV_PROD>", settings.NazvComp },
                { "<FIO_PROD>", settings.FIO },
                { "<NAZV_CLIENT>", c_NazvComp },
                { "<CLIENT_RASCHET>", c_RascSchet },
                { "<CLIENT_BANK>", c_Bank },
                { "<CLIENT_BIK>", c_BIK },
                { "<CLIENT_ADRESS>", c_Adress },
                { "<CLIENT_UNP>", c_UNP },
            };

            helper.Process(items);
        }

        public static string GetCityFromAddress(string address)
        {
            string[] addressParts = address.Split(',');
            string city = addressParts[0].Trim();

            if (city.StartsWith("г."))
            {
                return city;
            }
            else
            {
                return "Ошибка, не правильно составлены настройки вывода!";
            }
        }

        public string GetInitials(string name)
        {
            var parts = name.Split();
            if (parts.Length < 2)
            {
                return name;
            }
            var firstInitial = parts[1][0];
            var lastInitial = parts.Length > 2 ? parts[2][0].ToString() : "";
            return $"{parts[0]} {firstInitial}. {lastInitial}.".Trim();
        }

        private void rjButton13_Click(object sender, EventArgs e)
        {
            settings_edit F2 = new settings_edit();
            if (F2.ShowDialog() == DialogResult.OK)
            {
                MessageBox.Show("Настройки успешно изменены!", "Успех!");
            }
        }

        private void SubRefresh() 
        {
            if (dataGridView1.CurrentCell != null)
            {
                string query = $"SELECT User2Product.id, Products.name AS [Название продукта] FROM Products, User2Product, Clients WHERE User2Product.client = Clients.id AND User2Product.product = Products.id AND Clients.id = '{dataGridView1.Rows[dataGridView1.CurrentCell.RowIndex].Cells[0].Value.ToString()}'";
                var connectionString = ConfigurationManager.ConnectionStrings["db"].ConnectionString;
                using (SqlDataAdapter adapter = new SqlDataAdapter(query, connectionString))
                {
                    DataTable dataTable = new DataTable();
                    adapter.Fill(dataTable);
                    dataGridView2.DataSource = dataTable;
                }
                dataGridView2.Columns[0].Visible = false;
                DogovorRefresh();
            }
        }

        private void DogovorRefresh()
        {
            int client_id = (int)dataGridView1.Rows[dataGridView1.CurrentCell.RowIndex].Cells[0].Value;

            int user2prod_id;

            if (dataGridView2.CurrentCell != null)
            {
                user2prod_id = (int)dataGridView2.Rows[dataGridView2.CurrentCell.RowIndex].Cells[0].Value;
                int prod_id;

                string query2 = $"SELECT * FROM User2Product WHERE id = {user2prod_id}";
                var connectionString = ConfigurationManager.ConnectionStrings["db"].ConnectionString;
                using (SqlDataAdapter adapter = new SqlDataAdapter(query2, connectionString))
                {
                    DataTable dataTable = new DataTable();
                    adapter.Fill(dataTable);
                    prod_id = (int)dataTable.Rows[0][2];
                }

                //string query = $"SELECT Treaty.id, ('Договор' + ' №' + convert(nvarchar(max), Treaty.nomer, 0) + ' ' + '|' + ' ' + convert(nvarchar(max), Treaty.dateСonclusion, 0) + ' ' + '(' + convert(nvarchar(max), Treaty.dataFrom, 0) + ' - ' + convert(nvarchar(max), Treaty.dateTo, 0) + ')') AS [Договор] FROM Treaty, Clients, Products, User2Product WHERE Treaty.client = Clients.id AND Treaty.product = Products.id AND User2Product.client = Treaty.client AND User2Product.product = Treaty.product AND Clients.id = {client_id} AND Products.id = {prod_id}";
                string query = $"SELECT Treaty.id, ('Договор' + ' №' + convert(nvarchar(max), Treaty.nomer, 0) + ' ' + '|' + ' ' + convert(nvarchar(max), Treaty.dateСonclusion, 0) + ' ' + '(' + convert(nvarchar(max), Treaty.dataFrom, 0) + ' - ' + convert(nvarchar(max), Treaty.dateTo, 0) + ')') AS [Договор] FROM Treaty, Clients, Products, User2Product WHERE Treaty.client = Clients.id AND Treaty.product = Products.id AND User2Product.client = Treaty.client AND User2Product.product = Treaty.product AND Clients.id = {client_id} AND Products.id = {prod_id} ORDER BY Treaty.id DESC";
                using (SqlDataAdapter adapter = new SqlDataAdapter(query, connectionString))
                {
                    DataTable dataTable = new DataTable();
                    adapter.Fill(dataTable);
                    dataGridView3.DataSource = dataTable;
                }
                dataGridView3.Columns[0].Visible = false;
            }
        }
    }
}
