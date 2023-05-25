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
using Newtonsoft.Json;
using techSupport.edit_form;
using SettingsClass;
using static System.Windows.Forms.VisualStyles.VisualStyleElement.Rebar;
using WordHelperClass;
using System.Text.RegularExpressions;

namespace techSupport.new_forms
{
    public partial class dogovor_form : Form
    {
        private SqlConnection sqlConnection = null;

        public dogovor_form()
        {
            sqlConnection = new SqlConnection(ConfigurationManager.ConnectionStrings["db"].ConnectionString);
            sqlConnection.Open();
            InitializeComponent();
            RefreshTable();
        }

        private void RefreshTable()
        {
            string query = "SELECT Treaty.id, Treaty.nomer AS [Номер], (Clients.surname + ' ' + Clients.name + ' ' + Clients.patronymic) AS [Клиент], Products.name AS [Продукт], Treaty.dateСonclusion AS [Дата формирования], Treaty.dataFrom AS [Дата действия с], Treaty.dateTo AS [Дата действия по] FROM Treaty, Clients, Products, User2Product WHERE Treaty.client = Clients.id AND Treaty.product = Products.id AND User2Product.client = Clients.id AND User2Product.product = Products.id";
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
            if (new dogovor_edit().ShowDialog() == DialogResult.OK)
            {
                RefreshTable();
                MessageBox.Show("Запись успешно добавлена!", "Успех!");
            }
        }

        private void rjButton2_Click(object sender, EventArgs e)
        {
            dogovor_edit F2 = new dogovor_edit();
            F2.Text = "Редактирование договора";
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
                    SqlCommand command = new SqlCommand($"DELETE FROM Treaty WHERE id = {dataGridView1.Rows[dataGridView1.CurrentCell.RowIndex].Cells[0].Value}", sqlConnection);
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
                string query = $"SELECT Treaty.id, Treaty.nomer AS [Номер], (Clients.surname + ' ' + Clients.name + ' ' + Clients.patronymic) AS [Клиент], Products.name AS [Продукт], Treaty.dateСonclusion AS [Дата формирования], Treaty.dataFrom AS [Дата действия с], Treaty.dateTo AS [Дата действия по] FROM Treaty, Clients, Products, User2Product WHERE Treaty.client = Clients.id AND Treaty.product = Products.id AND User2Product.client = Clients.id AND User2Product.product = Products.id AND Clients.surname + ' ' + Clients.name + ' ' + Clients.patronymic + ' ' + Products.name + ' ' + convert(nvarchar(max), Treaty.nomer, 0) + ' ' + convert(nvarchar(max), Treaty.dateTo, 0) + ' ' + convert(nvarchar(max), Treaty.dateСonclusion, 0) LIKE '%{textBox1.Text}%'";
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

            int v_id = (int)dataGridView1.Rows[dataGridView1.CurrentCell.RowIndex].Cells[0].Value;

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

        private static string ExctraxtIni(string s)
        {
            var inits = Regex.Match(s, @"(\w+)\s+(\w+)\s+(\w+)").Groups;
            return string.Format("{0} {1}. {2}.", inits[1], inits[2].Value[0], inits[3].Value[0]);
        }

        private void rjButton6_Click(object sender, EventArgs e)
        {
            settings_edit F2 = new settings_edit();
            if (F2.ShowDialog() == DialogResult.OK)
            {
                MessageBox.Show("Настройки успешно изменены!", "Успех!");
            }
        }
    }
}
