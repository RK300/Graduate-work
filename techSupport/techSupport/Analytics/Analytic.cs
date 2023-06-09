using ClosedXML.Excel;
using DocumentFormat.OpenXml.Office2010.Excel;
using Newtonsoft.Json;
using SettingsClass;
using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Configuration;
using System.Data;
using System.Data.SqlClient;
using System.Drawing;
using System.IO;
using System.Linq;
using System.Runtime.InteropServices.ComTypes;
using System.Text;
using System.Text.RegularExpressions;
using System.Threading.Tasks;
using System.Windows.Forms;

namespace techSupport.Analytics
{
    public partial class Analytic : Form
    {
        private SqlConnection sqlConnection = null;
        private int WorkerId;

        public Analytic(int id)
        {
            SetId(id);
            sqlConnection = new SqlConnection(ConfigurationManager.ConnectionStrings["db"].ConnectionString);
            sqlConnection.Open();
            InitializeComponent();
            combobox(comboBox1, "SELECT id, (CompanyName + ' | ' + surname + ' ' + name + ' ' + patronymic) AS [FIO] FROM [Clients]", "FIO", "id");
            combobox(comboBox2, "SELECT Products.id, Products.name FROM [Products]", "name", "id");
            combobox(comboBox3, "SELECT id, (surname + ' ' + name + ' ' + patronymic) AS [FIO] FROM [Worker]", "FIO", "id");
        }

        private string GetWorkerName() 
        {
            string query = $"SELECT (surname + ' ' + name + ' ' + patronymic) FROM Worker WHERE id = {WorkerId}";
            var connectionString = ConfigurationManager.ConnectionStrings["db"].ConnectionString;
            using (SqlDataAdapter adapter = new SqlDataAdapter(query, connectionString))
            {
                DataTable dataTable = new DataTable();
                adapter.Fill(dataTable);
                return dataTable.Rows[0][0].ToString();
            }
        }

        private string GetMyCompany()
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

            return settings.NazvComp;
        }

        private void SetId(int id) { WorkerId = id; }

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

        private void comboBox3_KeyPress(object sender, KeyPressEventArgs e)
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

        private void button1_Click(object sender, EventArgs e)
        {
            using (var workbook = new XLWorkbook())
            {
                var worksheet = workbook.Worksheets.Add("Тикеты за период");
                string filename = "Тикеты за период (" + dateTimePicker1.Value.ToString("D") + " - " + dateTimePicker2.Value.ToString("D") + ").xlsx";

                var col1 = worksheet.Column("A");
                col1.Width = col1.Width * 2.8;

                var col2 = worksheet.Column("B");
                col2.Width = col2.Width * 2.8;

                var col11 = worksheet.Column("C");
                col11.Width = col11.Width * 3;

                var col3 = worksheet.Column("D");
                col3.Width = col3.Width * 2.5;

                var col4 = worksheet.Column("E");
                col4.Width = col4.Width * 4.2;

                var col5 = worksheet.Column("F");
                col5.Width = col5.Width * 4.6;

                var col6 = worksheet.Column("G");
                col6.Width = col6.Width * 2;

                var col7 = worksheet.Column("H");
                col7.Width = col7.Width * 2;

                var col8 = worksheet.Column("I");
                col8.Width = col8.Width * 2.2;

                var col9 = worksheet.Column("J");
                col9.Width = col9.Width * 2.2;

                var col10 = worksheet.Column("K");
                col10.Width = col10.Width * 2.5;

                //ШАПКА

                worksheet.Cell("A1").Value = "Отчёт";
                worksheet.Cell("A1").Style.Alignment.Horizontal = XLAlignmentHorizontalValues.Center;
                worksheet.Cell("A1").Style.Alignment.Vertical = XLAlignmentVerticalValues.Center;
                worksheet.Cell("A1").Style.Font.Bold = true;
                worksheet.Cell("A1").Style.Font.FontName = "Times New Roman";
                worksheet.Cell("A1").Style.Font.FontSize = 16;
                worksheet.Range("A1:K1").Merge();

                worksheet.Cell("A2").Value = "Тикеты за период";
                worksheet.Cell("A2").Style.Alignment.Horizontal = XLAlignmentHorizontalValues.Center;
                worksheet.Cell("A2").Style.Alignment.Vertical = XLAlignmentVerticalValues.Center;
                worksheet.Cell("A2").Style.Font.Bold = true;
                worksheet.Cell("A2").Style.Font.FontName = "Times New Roman";
                worksheet.Cell("A2").Style.Font.FontSize = 16;
                worksheet.Range("A2:K2").Merge();

                string Period = dateTimePicker1.Value.ToString("d") + " - " + dateTimePicker2.Value.ToString("d");

                worksheet.Cell("A3").Value = Period;
                worksheet.Cell("A3").Style.Alignment.Horizontal = XLAlignmentHorizontalValues.Center;
                worksheet.Cell("A3").Style.Alignment.Vertical = XLAlignmentVerticalValues.Center;
                worksheet.Cell("A3").Style.Font.Bold = true;
                worksheet.Cell("A3").Style.Font.FontName = "Times New Roman";
                worksheet.Cell("A3").Style.Font.FontSize = 16;
                worksheet.Range("A3:K3").Merge();

                string Form = "Дата формировани: " + DateTime.Now.ToString();

                worksheet.Cell("I4").Value = Form;
                worksheet.Cell("I4").Style.Alignment.Horizontal = XLAlignmentHorizontalValues.Right;
                worksheet.Cell("I4").Style.Alignment.Vertical = XLAlignmentVerticalValues.Center;
                worksheet.Cell("I4").Style.Font.Bold = true;
                worksheet.Cell("I4").Style.Font.FontName = "Times New Roman";
                worksheet.Cell("I4").Style.Font.FontSize = 16;
                worksheet.Range("I4:K4").Merge();

                worksheet.Cell("A4").Value = "Сформировал: " + GetInitials(GetWorkerName());
                worksheet.Cell("A4").Style.Alignment.Horizontal = XLAlignmentHorizontalValues.Left;
                worksheet.Cell("A4").Style.Alignment.Vertical = XLAlignmentVerticalValues.Center;
                worksheet.Cell("A4").Style.Font.Bold = true;
                worksheet.Cell("A4").Style.Font.FontName = "Times New Roman";
                worksheet.Cell("A4").Style.Font.FontSize = 16;
                worksheet.Range("A4:C4").Merge();

                worksheet.Cell("E4").Value = GetMyCompany();
                worksheet.Cell("E4").Style.Alignment.Horizontal = XLAlignmentHorizontalValues.Center;
                worksheet.Cell("E4").Style.Alignment.Vertical = XLAlignmentVerticalValues.Center;
                worksheet.Cell("E4").Style.Font.Bold = true;
                worksheet.Cell("E4").Style.Font.FontName = "Times New Roman";
                worksheet.Cell("E4").Style.Font.FontSize = 16;
                worksheet.Range("E4:F4").Merge();

                worksheet.Row(1).Height = worksheet.Row(1).Height * 1.5;
                worksheet.Row(2).Height = worksheet.Row(2).Height * 1.5;
                worksheet.Row(3).Height = worksheet.Row(3).Height * 1.5;

                //worksheet.Rows().AdjustToContents();
                //worksheet.Columns().AdjustToContents();

                //ТАБЛИЦА

                worksheet.Cell("A5").Value = "Клиент";
                worksheet.Cell("A5").Style.Alignment.Horizontal = XLAlignmentHorizontalValues.Center;
                worksheet.Cell("A5").Style.Alignment.Vertical = XLAlignmentVerticalValues.Center;
                worksheet.Cell("A5").Style.Font.Bold = true;
                worksheet.Cell("A5").Style.Font.FontName = "Times New Roman";
                worksheet.Cell("A5").Style.Font.FontSize = 14;
                worksheet.Cell("A5").Style.Alignment.WrapText = true;
                worksheet.Range("A5:A7").Style.Border.OutsideBorder = XLBorderStyleValues.Thin;
                worksheet.Range("A5:A7").Merge();

                worksheet.Cell("B5").Value = "Продукт";
                worksheet.Cell("B5").Style.Alignment.Horizontal = XLAlignmentHorizontalValues.Center;
                worksheet.Cell("B5").Style.Alignment.Vertical = XLAlignmentVerticalValues.Center;
                worksheet.Cell("B5").Style.Font.Bold = true;
                worksheet.Cell("B5").Style.Font.FontName = "Times New Roman";
                worksheet.Cell("B5").Style.Font.FontSize = 14;
                worksheet.Cell("B5").Style.Alignment.WrapText = true;
                worksheet.Range("B5:B7").Style.Border.OutsideBorder = XLBorderStyleValues.Thin;
                worksheet.Range("B5:B7").Merge();

                worksheet.Cell("C5").Value = "Принявший сотрудник";
                worksheet.Cell("C5").Style.Alignment.Horizontal = XLAlignmentHorizontalValues.Center;
                worksheet.Cell("C5").Style.Alignment.Vertical = XLAlignmentVerticalValues.Center;
                worksheet.Cell("C5").Style.Font.Bold = true;
                worksheet.Cell("C5").Style.Font.FontName = "Times New Roman";
                worksheet.Cell("C5").Style.Font.FontSize = 14;
                worksheet.Cell("C5").Style.Alignment.WrapText = true;
                worksheet.Range("C5:C7").Style.Border.OutsideBorder = XLBorderStyleValues.Thin;
                worksheet.Range("C5:C7").Merge();

                worksheet.Cell("D5").Value = "Тип проблемы";
                worksheet.Cell("D5").Style.Alignment.Horizontal = XLAlignmentHorizontalValues.Center;
                worksheet.Cell("D5").Style.Alignment.Vertical = XLAlignmentVerticalValues.Center;
                worksheet.Cell("D5").Style.Font.Bold = true;
                worksheet.Cell("D5").Style.Font.FontName = "Times New Roman";
                worksheet.Cell("D5").Style.Font.FontSize = 14;
                worksheet.Cell("D5").Style.Alignment.WrapText = true;
                worksheet.Range("D5:D7").Style.Border.OutsideBorder = XLBorderStyleValues.Thin;
                worksheet.Range("D5:D7").Merge();

                worksheet.Cell("E5").Value = "Заголовок";
                worksheet.Cell("E5").Style.Alignment.Horizontal = XLAlignmentHorizontalValues.Center;
                worksheet.Cell("E5").Style.Alignment.Vertical = XLAlignmentVerticalValues.Center;
                worksheet.Cell("E5").Style.Font.Bold = true;
                worksheet.Cell("E5").Style.Font.FontName = "Times New Roman";
                worksheet.Cell("E5").Style.Font.FontSize = 14;
                worksheet.Cell("E5").Style.Alignment.WrapText = true;
                worksheet.Range("E5:E7").Style.Border.OutsideBorder = XLBorderStyleValues.Thin;
                worksheet.Range("E5:E7").Merge();

                worksheet.Cell("F5").Value = "Описание";
                worksheet.Cell("F5").Style.Alignment.Horizontal = XLAlignmentHorizontalValues.Center;
                worksheet.Cell("F5").Style.Alignment.Vertical = XLAlignmentVerticalValues.Center;
                worksheet.Cell("F5").Style.Font.Bold = true;
                worksheet.Cell("F5").Style.Font.FontName = "Times New Roman";
                worksheet.Cell("F5").Style.Font.FontSize = 14;
                worksheet.Cell("F5").Style.Alignment.WrapText = true;
                worksheet.Range("F5:F7").Style.Border.OutsideBorder = XLBorderStyleValues.Thin;
                worksheet.Range("F5:F7").Merge();

                worksheet.Cell("G5").Value = "Приоритет";
                worksheet.Cell("G5").Style.Alignment.Horizontal = XLAlignmentHorizontalValues.Center;
                worksheet.Cell("G5").Style.Alignment.Vertical = XLAlignmentVerticalValues.Center;
                worksheet.Cell("G5").Style.Font.Bold = true;
                worksheet.Cell("G5").Style.Font.FontName = "Times New Roman";
                worksheet.Cell("G5").Style.Font.FontSize = 14;
                worksheet.Cell("G5").Style.Alignment.WrapText = true;
                worksheet.Range("G5:G7").Style.Border.OutsideBorder = XLBorderStyleValues.Thin;
                worksheet.Range("G5:G7").Merge();

                worksheet.Cell("H5").Value = "Статус";
                worksheet.Cell("H5").Style.Alignment.Horizontal = XLAlignmentHorizontalValues.Center;
                worksheet.Cell("H5").Style.Alignment.Vertical = XLAlignmentVerticalValues.Center;
                worksheet.Cell("H5").Style.Font.Bold = true;
                worksheet.Cell("H5").Style.Font.FontName = "Times New Roman";
                worksheet.Cell("H5").Style.Font.FontSize = 14;
                worksheet.Cell("H5").Style.Alignment.WrapText = true;
                worksheet.Range("H5:H7").Style.Border.OutsideBorder = XLBorderStyleValues.Thin;
                worksheet.Range("H5:H7").Merge();

                worksheet.Cell("I5").Value = "Дата подачи";
                worksheet.Cell("I5").Style.Alignment.Horizontal = XLAlignmentHorizontalValues.Center;
                worksheet.Cell("I5").Style.Alignment.Vertical = XLAlignmentVerticalValues.Center;
                worksheet.Cell("I5").Style.Font.Bold = true;
                worksheet.Cell("I5").Style.Font.FontName = "Times New Roman";
                worksheet.Cell("I5").Style.Font.FontSize = 14;
                worksheet.Cell("I5").Style.Alignment.WrapText = true;
                worksheet.Range("I5:I7").Style.Border.OutsideBorder = XLBorderStyleValues.Thin;
                worksheet.Range("I5:I7").Merge();

                worksheet.Cell("J5").Value = "Дата закрытия";
                worksheet.Cell("J5").Style.Alignment.Horizontal = XLAlignmentHorizontalValues.Center;
                worksheet.Cell("J5").Style.Alignment.Vertical = XLAlignmentVerticalValues.Center;
                worksheet.Cell("J5").Style.Font.Bold = true;
                worksheet.Cell("J5").Style.Font.FontName = "Times New Roman";
                worksheet.Cell("J5").Style.Font.FontSize = 14;
                worksheet.Cell("J5").Style.Alignment.WrapText = true;
                worksheet.Range("J5:J7").Style.Border.OutsideBorder = XLBorderStyleValues.Thin;
                worksheet.Range("J5:J7").Merge();

                worksheet.Cell("K5").Value = "Фактическая дата закрытия";
                worksheet.Cell("K5").Style.Alignment.Horizontal = XLAlignmentHorizontalValues.Center;
                worksheet.Cell("K5").Style.Alignment.Vertical = XLAlignmentVerticalValues.Center;
                worksheet.Cell("K5").Style.Font.Bold = true;
                worksheet.Cell("K5").Style.Font.FontName = "Times New Roman";
                worksheet.Cell("K5").Style.Font.FontSize = 14;
                worksheet.Cell("K5").Style.Alignment.WrapText = true;
                worksheet.Range("K5:K7").Style.Border.OutsideBorder = XLBorderStyleValues.Thin;
                worksheet.Range("K5:K7").Merge();

                worksheet.Row(7).Height = worksheet.Row(7).Height * 2;

                //ДАННЫЕ

                //string query = $"SELECT Ticket.id, Clients.CompanyName AS [Клиент], Products.name AS [Продукт], (Worker.surname + ' ' + Worker.name + ' ' + Worker.patronymic) AS [Принявший сотрудник], Type.name AS [Тип проблемы], Ticket.header AS [Заголовок], Ticket.description AS [Описание], Ticket.priority AS [Приоритет], Ticket.status AS [Статус], Ticket.application_data AS [Дата подачи], Ticket.completion_data AS [Дата закрытия], Ticket.actual_data AS [Фактическая дата закрытия] FROM Ticket, Clients, Products, Worker, Type WHERE Ticket.client = Clients.id AND Ticket.product = Products.id AND Ticket.worker = Worker.id AND Ticket.type = Type.id AND Ticket.application_data >= {dateTimePicker1.Value.ToString()} AND Ticket.application_data <= {dateTimePicker2.Value.ToString()}";
                //string query = $"SELECT Clients.CompanyName AS [Клиент], Products.name AS [Продукт], (Worker.surname + ' ' + Worker.name + ' ' + Worker.patronymic) AS [Принявший сотрудник], Type.name AS [Тип проблемы], Ticket.header AS [Заголовок], Ticket.description AS [Описание], Ticket.priority AS [Приоритет], Ticket.status AS [Статус], Ticket.application_data AS [Дата подачи], Ticket.completion_data AS [Дата закрытия], Ticket.actual_data AS [Фактическая дата закрытия] FROM Ticket, Clients, Products, Worker, Type WHERE Ticket.client = Clients.id AND Ticket.product = Products.id AND Ticket.worker = Worker.id AND Ticket.type = Type.id AND Ticket.application_data >= CAST('{dateTimePicker1.Value.ToString("d")}' AS DATETIME) AND Ticket.application_data <= CAST('{dateTimePicker2.Value.ToString("d")}' AS DATETIME)";
                string query = $"SELECT Clients.CompanyName AS [Клиент], Products.name AS [Продукт], (Worker.surname + ' ' + Worker.name + ' ' + Worker.patronymic) AS [Принявший сотрудник], Type.name AS [Тип проблемы], Ticket.header AS [Заголовок], Ticket.description AS [Описание], Ticket.priority AS [Приоритет], Ticket.status AS [Статус], Ticket.application_data AS [Дата подачи], Ticket.completion_data AS [Дата закрытия], Ticket.actual_data AS [Фактическая дата закрытия] FROM Ticket, Clients, Products, Worker, Type WHERE Ticket.client = Clients.id AND Ticket.product = Products.id AND Ticket.worker = Worker.id AND Ticket.type = Type.id AND Ticket.application_data >= CAST('{dateTimePicker1.Value.ToString()}' AS DATETIME) AND Ticket.application_data <= CAST('{dateTimePicker2.Value.ToString()}' AS DATETIME) ORDER BY Ticket.application_data ASC";
                var connectionString = ConfigurationManager.ConnectionStrings["db"].ConnectionString;
                using (SqlDataAdapter adapter = new SqlDataAdapter(query, connectionString))
                {
                    DataTable dataTable = new DataTable();
                    adapter.Fill(dataTable);

                    if ( dataTable.Rows.Count == 0 ) 
                    {
                        MessageBox.Show("Нет тикетов за этот временной промежуток!", "Ошибка", MessageBoxButtons.OK, MessageBoxIcon.Error);
                        return;
                    }

                    DateTime fix = DateTime.Now;

                    for ( int i = 0, b = 8, x = dataTable.Rows.Count; i < x; i++ ) 
                    {
                        string n = "A" + b.ToString();

                        worksheet.Cell(n).Value = dataTable.Rows[i][0].ToString();
                        worksheet.Cell(n).Style.Alignment.Horizontal = XLAlignmentHorizontalValues.Center;
                        worksheet.Cell(n).Style.Alignment.Vertical = XLAlignmentVerticalValues.Center;
                        worksheet.Cell(n).Style.Font.FontName = "Times New Roman";
                        worksheet.Cell(n).Style.Font.FontSize = 14;
                        worksheet.Cell(n).Style.Alignment.WrapText = true;
                        worksheet.Cell(n).Style.Border.OutsideBorder = XLBorderStyleValues.Thin;

                        n = "B" + b.ToString();

                        worksheet.Cell(n).Value = dataTable.Rows[i][1].ToString();
                        worksheet.Cell(n).Style.Alignment.Horizontal = XLAlignmentHorizontalValues.Center;
                        worksheet.Cell(n).Style.Alignment.Vertical = XLAlignmentVerticalValues.Center;
                        worksheet.Cell(n).Style.Font.FontName = "Times New Roman";
                        worksheet.Cell(n).Style.Font.FontSize = 14;
                        worksheet.Cell(n).Style.Alignment.WrapText = true;
                        worksheet.Cell(n).Style.Border.OutsideBorder = XLBorderStyleValues.Thin;

                        n = "C" + b.ToString();

                        worksheet.Cell(n).Value = dataTable.Rows[i][2].ToString();
                        worksheet.Cell(n).Style.Alignment.Horizontal = XLAlignmentHorizontalValues.Center;
                        worksheet.Cell(n).Style.Alignment.Vertical = XLAlignmentVerticalValues.Center;
                        worksheet.Cell(n).Style.Font.FontName = "Times New Roman";
                        worksheet.Cell(n).Style.Font.FontSize = 14;
                        worksheet.Cell(n).Style.Alignment.WrapText = true;
                        worksheet.Cell(n).Style.Border.OutsideBorder = XLBorderStyleValues.Thin;

                        n = "D" + b.ToString();

                        worksheet.Cell(n).Value = dataTable.Rows[i][3].ToString();
                        worksheet.Cell(n).Style.Alignment.Horizontal = XLAlignmentHorizontalValues.Center;
                        worksheet.Cell(n).Style.Alignment.Vertical = XLAlignmentVerticalValues.Center;
                        worksheet.Cell(n).Style.Font.FontName = "Times New Roman";
                        worksheet.Cell(n).Style.Font.FontSize = 14;
                        worksheet.Cell(n).Style.Alignment.WrapText = true;
                        worksheet.Cell(n).Style.Border.OutsideBorder = XLBorderStyleValues.Thin;

                        n = "E" + b.ToString();

                        worksheet.Cell(n).Value = dataTable.Rows[i][4].ToString();
                        worksheet.Cell(n).Style.Alignment.Horizontal = XLAlignmentHorizontalValues.Center;
                        worksheet.Cell(n).Style.Alignment.Vertical = XLAlignmentVerticalValues.Center;
                        worksheet.Cell(n).Style.Font.FontName = "Times New Roman";
                        worksheet.Cell(n).Style.Font.FontSize = 14;
                        worksheet.Cell(n).Style.Alignment.WrapText = true;
                        worksheet.Cell(n).Style.Border.OutsideBorder = XLBorderStyleValues.Thin;

                        n = "F" + b.ToString();

                        worksheet.Cell(n).Value = dataTable.Rows[i][5].ToString();
                        worksheet.Cell(n).Style.Alignment.Horizontal = XLAlignmentHorizontalValues.Center;
                        worksheet.Cell(n).Style.Alignment.Vertical = XLAlignmentVerticalValues.Center;
                        worksheet.Cell(n).Style.Font.FontName = "Times New Roman";
                        worksheet.Cell(n).Style.Font.FontSize = 14;
                        worksheet.Cell(n).Style.Alignment.WrapText = true;
                        worksheet.Cell(n).Style.Border.OutsideBorder = XLBorderStyleValues.Thin;

                        n = "G" + b.ToString();

                        worksheet.Cell(n).Value = dataTable.Rows[i][6].ToString();
                        worksheet.Cell(n).Style.Alignment.Horizontal = XLAlignmentHorizontalValues.Center;
                        worksheet.Cell(n).Style.Alignment.Vertical = XLAlignmentVerticalValues.Center;
                        worksheet.Cell(n).Style.Font.FontName = "Times New Roman";
                        worksheet.Cell(n).Style.Font.FontSize = 14;
                        worksheet.Cell(n).Style.Alignment.WrapText = true;
                        worksheet.Cell(n).Style.Border.OutsideBorder = XLBorderStyleValues.Thin;

                        n = "H" + b.ToString();

                        //Статус
                        worksheet.Cell(n).Value = dataTable.Rows[i][7].ToString();
                        worksheet.Cell(n).Style.Alignment.Horizontal = XLAlignmentHorizontalValues.Center;
                        worksheet.Cell(n).Style.Alignment.Vertical = XLAlignmentVerticalValues.Center;
                        worksheet.Cell(n).Style.Font.FontName = "Times New Roman";
                        worksheet.Cell(n).Style.Font.FontSize = 14;
                        worksheet.Cell(n).Style.Alignment.WrapText = true;
                        worksheet.Cell(n).Style.Border.OutsideBorder = XLBorderStyleValues.Thin;

                        n = "I" + b.ToString();

                        //Даты
                        fix = (DateTime)dataTable.Rows[i][8];
                        worksheet.Cell(n).Value = fix.ToString("d");
                        worksheet.Cell(n).Style.Alignment.Horizontal = XLAlignmentHorizontalValues.Center;
                        worksheet.Cell(n).Style.Alignment.Vertical = XLAlignmentVerticalValues.Center;
                        worksheet.Cell(n).Style.Font.FontName = "Times New Roman";
                        worksheet.Cell(n).Style.Font.FontSize = 14;
                        worksheet.Cell(n).Style.Alignment.WrapText = true;
                        worksheet.Cell(n).Style.Border.OutsideBorder = XLBorderStyleValues.Thin;

                        n = "J" + b.ToString();

                        fix = (DateTime)dataTable.Rows[i][9];
                        worksheet.Cell(n).Value = fix.ToString("d");
                        worksheet.Cell(n).Style.Alignment.Horizontal = XLAlignmentHorizontalValues.Center;
                        worksheet.Cell(n).Style.Alignment.Vertical = XLAlignmentVerticalValues.Center;
                        worksheet.Cell(n).Style.Font.FontName = "Times New Roman";
                        worksheet.Cell(n).Style.Font.FontSize = 14;
                        worksheet.Cell(n).Style.Alignment.WrapText = true;
                        worksheet.Cell(n).Style.Border.OutsideBorder = XLBorderStyleValues.Thin;

                        n = "K" + b.ToString();

                        //Фактическая
                        if ( dataTable.Rows[i][7].ToString() == "Закрыт" ) 
                        {
                            fix = (DateTime)dataTable.Rows[i][10];
                            worksheet.Cell(n).Value = fix.ToString("d");
                        }
                        else
                        {
                            worksheet.Cell(n).Value = "Нет";
                        }
                        worksheet.Cell(n).Style.Alignment.Horizontal = XLAlignmentHorizontalValues.Center;
                        worksheet.Cell(n).Style.Alignment.Vertical = XLAlignmentVerticalValues.Center;
                        worksheet.Cell(n).Style.Font.FontName = "Times New Roman";
                        worksheet.Cell(n).Style.Font.FontSize = 14;
                        worksheet.Cell(n).Style.Alignment.WrapText = true;
                        worksheet.Cell(n).Style.Border.OutsideBorder = XLBorderStyleValues.Thin;

                        b++;
                    }
                }

                //worksheet.Column(1).AdjustToContents();
                //worksheet.Rows(1, 1000).AdjustToContents();

                //СОХРАНЕНИЕ

                //worksheet.Columns().AdjustToContents();

                saveFileDialog1.Filter = "Microsoft Excel Files (*.xlsx)|*.xlsx";
                saveFileDialog1.FileName = filename;
                if (saveFileDialog1.ShowDialog() == DialogResult.OK)
                {
                    try
                    {
                        string filename2 = saveFileDialog1.FileName;
                        workbook.SaveAs(filename2);
                        System.Diagnostics.Process.Start(filename2);
                        MessageBox.Show("Отчет сформирован!", "Успех!", MessageBoxButtons.OK, MessageBoxIcon.Asterisk);
                    }
                    catch
                    {
                        MessageBox.Show("Файл уже существует или произошла другая ошибка!", "Ошибка", MessageBoxButtons.OK, MessageBoxIcon.Error);
                    }
                }
                else
                {
                    MessageBox.Show("Вы отменили сохранение!", "Ошибка", MessageBoxButtons.OK, MessageBoxIcon.Error);
                    return;
                }
            }
        }

        private void button2_Click(object sender, EventArgs e)
        {
            if (comboBox1.SelectedValue == null)
                return;

            using (var workbook = new XLWorkbook())
            {
                string CompanyName;
                string query2 = $"SELECT CompanyName FROM Clients WHERE id = {comboBox1.SelectedValue}";
                var connectionString = ConfigurationManager.ConnectionStrings["db"].ConnectionString;
                using (SqlDataAdapter adapter = new SqlDataAdapter(query2, connectionString))
                {
                    DataTable dataTable = new DataTable();
                    adapter.Fill(dataTable);

                    CompanyName = dataTable.Rows[0][0].ToString();
                }

                var worksheet = workbook.Worksheets.Add("Тикеты клиента");
                string filename = "Тикеты " + CompanyName + " на " + DateTime.Now.ToString("d") + ".xlsx";

                var col1 = worksheet.Column("A");
                col1.Width = col1.Width * 2.8;

                var col2 = worksheet.Column("B");
                col2.Width = col2.Width * 2.8;

                var col11 = worksheet.Column("C");
                col11.Width = col11.Width * 3;

                var col3 = worksheet.Column("D");
                col3.Width = col3.Width * 2.5;

                var col4 = worksheet.Column("E");
                col4.Width = col4.Width * 4.2;

                var col5 = worksheet.Column("F");
                col5.Width = col5.Width * 4.6;

                var col6 = worksheet.Column("G");
                col6.Width = col6.Width * 2;

                var col7 = worksheet.Column("H");
                col7.Width = col7.Width * 2;

                var col8 = worksheet.Column("I");
                col8.Width = col8.Width * 2.2;

                var col9 = worksheet.Column("J");
                col9.Width = col9.Width * 2.2;

                var col10 = worksheet.Column("K");
                col10.Width = col10.Width * 2.5;

                //ШАПКА

                worksheet.Cell("A1").Value = "Отчёт";
                worksheet.Cell("A1").Style.Alignment.Horizontal = XLAlignmentHorizontalValues.Center;
                worksheet.Cell("A1").Style.Alignment.Vertical = XLAlignmentVerticalValues.Center;
                worksheet.Cell("A1").Style.Font.Bold = true;
                worksheet.Cell("A1").Style.Font.FontName = "Times New Roman";
                worksheet.Cell("A1").Style.Font.FontSize = 16;
                worksheet.Range("A1:K1").Merge();

                string Shapka = "Все тикеты клиента - " + CompanyName;

                worksheet.Cell("A2").Value = Shapka;
                worksheet.Cell("A2").Style.Alignment.Horizontal = XLAlignmentHorizontalValues.Center;
                worksheet.Cell("A2").Style.Alignment.Vertical = XLAlignmentVerticalValues.Center;
                worksheet.Cell("A2").Style.Font.Bold = true;
                worksheet.Cell("A2").Style.Font.FontName = "Times New Roman";
                worksheet.Cell("A2").Style.Font.FontSize = 16;
                worksheet.Range("A2:K2").Merge();

                string Period = "на " + DateTime.Now.ToString("d");

                worksheet.Cell("A3").Value = Period;
                worksheet.Cell("A3").Style.Alignment.Horizontal = XLAlignmentHorizontalValues.Center;
                worksheet.Cell("A3").Style.Alignment.Vertical = XLAlignmentVerticalValues.Center;
                worksheet.Cell("A3").Style.Font.Bold = true;
                worksheet.Cell("A3").Style.Font.FontName = "Times New Roman";
                worksheet.Cell("A3").Style.Font.FontSize = 16;
                worksheet.Range("A3:K3").Merge();

                string Form = "Дата формировани: " + DateTime.Now.ToString();

                worksheet.Cell("I4").Value = Form;
                worksheet.Cell("I4").Style.Alignment.Horizontal = XLAlignmentHorizontalValues.Right;
                worksheet.Cell("I4").Style.Alignment.Vertical = XLAlignmentVerticalValues.Center;
                worksheet.Cell("I4").Style.Font.Bold = true;
                worksheet.Cell("I4").Style.Font.FontName = "Times New Roman";
                worksheet.Cell("I4").Style.Font.FontSize = 16;
                worksheet.Range("I4:K4").Merge();

                worksheet.Row(1).Height = worksheet.Row(1).Height * 1.5;
                worksheet.Row(2).Height = worksheet.Row(2).Height * 1.5;
                worksheet.Row(3).Height = worksheet.Row(3).Height * 1.5;

                worksheet.Cell("A4").Value = "Сформировал: " + GetInitials(GetWorkerName());
                worksheet.Cell("A4").Style.Alignment.Horizontal = XLAlignmentHorizontalValues.Left;
                worksheet.Cell("A4").Style.Alignment.Vertical = XLAlignmentVerticalValues.Center;
                worksheet.Cell("A4").Style.Font.Bold = true;
                worksheet.Cell("A4").Style.Font.FontName = "Times New Roman";
                worksheet.Cell("A4").Style.Font.FontSize = 16;
                worksheet.Range("A4:C4").Merge();

                worksheet.Cell("E4").Value = GetMyCompany();
                worksheet.Cell("E4").Style.Alignment.Horizontal = XLAlignmentHorizontalValues.Center;
                worksheet.Cell("E4").Style.Alignment.Vertical = XLAlignmentVerticalValues.Center;
                worksheet.Cell("E4").Style.Font.Bold = true;
                worksheet.Cell("E4").Style.Font.FontName = "Times New Roman";
                worksheet.Cell("E4").Style.Font.FontSize = 16;
                worksheet.Range("E4:F4").Merge();

                //worksheet.Rows().AdjustToContents();
                //worksheet.Columns().AdjustToContents();

                //ТАБЛИЦА

                worksheet.Cell("A5").Value = "Клиент";
                worksheet.Cell("A5").Style.Alignment.Horizontal = XLAlignmentHorizontalValues.Center;
                worksheet.Cell("A5").Style.Alignment.Vertical = XLAlignmentVerticalValues.Center;
                worksheet.Cell("A5").Style.Font.Bold = true;
                worksheet.Cell("A5").Style.Font.FontName = "Times New Roman";
                worksheet.Cell("A5").Style.Font.FontSize = 14;
                worksheet.Cell("A5").Style.Alignment.WrapText = true;
                worksheet.Range("A5:A7").Style.Border.OutsideBorder = XLBorderStyleValues.Thin;
                worksheet.Range("A5:A7").Merge();

                worksheet.Cell("B5").Value = "Продукт";
                worksheet.Cell("B5").Style.Alignment.Horizontal = XLAlignmentHorizontalValues.Center;
                worksheet.Cell("B5").Style.Alignment.Vertical = XLAlignmentVerticalValues.Center;
                worksheet.Cell("B5").Style.Font.Bold = true;
                worksheet.Cell("B5").Style.Font.FontName = "Times New Roman";
                worksheet.Cell("B5").Style.Font.FontSize = 14;
                worksheet.Cell("B5").Style.Alignment.WrapText = true;
                worksheet.Range("B5:B7").Style.Border.OutsideBorder = XLBorderStyleValues.Thin;
                worksheet.Range("B5:B7").Merge();

                worksheet.Cell("C5").Value = "Принявший сотрудник";
                worksheet.Cell("C5").Style.Alignment.Horizontal = XLAlignmentHorizontalValues.Center;
                worksheet.Cell("C5").Style.Alignment.Vertical = XLAlignmentVerticalValues.Center;
                worksheet.Cell("C5").Style.Font.Bold = true;
                worksheet.Cell("C5").Style.Font.FontName = "Times New Roman";
                worksheet.Cell("C5").Style.Font.FontSize = 14;
                worksheet.Cell("C5").Style.Alignment.WrapText = true;
                worksheet.Range("C5:C7").Style.Border.OutsideBorder = XLBorderStyleValues.Thin;
                worksheet.Range("C5:C7").Merge();

                worksheet.Cell("D5").Value = "Тип проблемы";
                worksheet.Cell("D5").Style.Alignment.Horizontal = XLAlignmentHorizontalValues.Center;
                worksheet.Cell("D5").Style.Alignment.Vertical = XLAlignmentVerticalValues.Center;
                worksheet.Cell("D5").Style.Font.Bold = true;
                worksheet.Cell("D5").Style.Font.FontName = "Times New Roman";
                worksheet.Cell("D5").Style.Font.FontSize = 14;
                worksheet.Cell("D5").Style.Alignment.WrapText = true;
                worksheet.Range("D5:D7").Style.Border.OutsideBorder = XLBorderStyleValues.Thin;
                worksheet.Range("D5:D7").Merge();

                worksheet.Cell("E5").Value = "Заголовок";
                worksheet.Cell("E5").Style.Alignment.Horizontal = XLAlignmentHorizontalValues.Center;
                worksheet.Cell("E5").Style.Alignment.Vertical = XLAlignmentVerticalValues.Center;
                worksheet.Cell("E5").Style.Font.Bold = true;
                worksheet.Cell("E5").Style.Font.FontName = "Times New Roman";
                worksheet.Cell("E5").Style.Font.FontSize = 14;
                worksheet.Cell("E5").Style.Alignment.WrapText = true;
                worksheet.Range("E5:E7").Style.Border.OutsideBorder = XLBorderStyleValues.Thin;
                worksheet.Range("E5:E7").Merge();

                worksheet.Cell("F5").Value = "Описание";
                worksheet.Cell("F5").Style.Alignment.Horizontal = XLAlignmentHorizontalValues.Center;
                worksheet.Cell("F5").Style.Alignment.Vertical = XLAlignmentVerticalValues.Center;
                worksheet.Cell("F5").Style.Font.Bold = true;
                worksheet.Cell("F5").Style.Font.FontName = "Times New Roman";
                worksheet.Cell("F5").Style.Font.FontSize = 14;
                worksheet.Cell("F5").Style.Alignment.WrapText = true;
                worksheet.Range("F5:F7").Style.Border.OutsideBorder = XLBorderStyleValues.Thin;
                worksheet.Range("F5:F7").Merge();

                worksheet.Cell("G5").Value = "Приоритет";
                worksheet.Cell("G5").Style.Alignment.Horizontal = XLAlignmentHorizontalValues.Center;
                worksheet.Cell("G5").Style.Alignment.Vertical = XLAlignmentVerticalValues.Center;
                worksheet.Cell("G5").Style.Font.Bold = true;
                worksheet.Cell("G5").Style.Font.FontName = "Times New Roman";
                worksheet.Cell("G5").Style.Font.FontSize = 14;
                worksheet.Cell("G5").Style.Alignment.WrapText = true;
                worksheet.Range("G5:G7").Style.Border.OutsideBorder = XLBorderStyleValues.Thin;
                worksheet.Range("G5:G7").Merge();

                worksheet.Cell("H5").Value = "Статус";
                worksheet.Cell("H5").Style.Alignment.Horizontal = XLAlignmentHorizontalValues.Center;
                worksheet.Cell("H5").Style.Alignment.Vertical = XLAlignmentVerticalValues.Center;
                worksheet.Cell("H5").Style.Font.Bold = true;
                worksheet.Cell("H5").Style.Font.FontName = "Times New Roman";
                worksheet.Cell("H5").Style.Font.FontSize = 14;
                worksheet.Cell("H5").Style.Alignment.WrapText = true;
                worksheet.Range("H5:H7").Style.Border.OutsideBorder = XLBorderStyleValues.Thin;
                worksheet.Range("H5:H7").Merge();

                worksheet.Cell("I5").Value = "Дата подачи";
                worksheet.Cell("I5").Style.Alignment.Horizontal = XLAlignmentHorizontalValues.Center;
                worksheet.Cell("I5").Style.Alignment.Vertical = XLAlignmentVerticalValues.Center;
                worksheet.Cell("I5").Style.Font.Bold = true;
                worksheet.Cell("I5").Style.Font.FontName = "Times New Roman";
                worksheet.Cell("I5").Style.Font.FontSize = 14;
                worksheet.Cell("I5").Style.Alignment.WrapText = true;
                worksheet.Range("I5:I7").Style.Border.OutsideBorder = XLBorderStyleValues.Thin;
                worksheet.Range("I5:I7").Merge();

                worksheet.Cell("J5").Value = "Дата закрытия";
                worksheet.Cell("J5").Style.Alignment.Horizontal = XLAlignmentHorizontalValues.Center;
                worksheet.Cell("J5").Style.Alignment.Vertical = XLAlignmentVerticalValues.Center;
                worksheet.Cell("J5").Style.Font.Bold = true;
                worksheet.Cell("J5").Style.Font.FontName = "Times New Roman";
                worksheet.Cell("J5").Style.Font.FontSize = 14;
                worksheet.Cell("J5").Style.Alignment.WrapText = true;
                worksheet.Range("J5:J7").Style.Border.OutsideBorder = XLBorderStyleValues.Thin;
                worksheet.Range("J5:J7").Merge();

                worksheet.Cell("K5").Value = "Фактическая дата закрытия";
                worksheet.Cell("K5").Style.Alignment.Horizontal = XLAlignmentHorizontalValues.Center;
                worksheet.Cell("K5").Style.Alignment.Vertical = XLAlignmentVerticalValues.Center;
                worksheet.Cell("K5").Style.Font.Bold = true;
                worksheet.Cell("K5").Style.Font.FontName = "Times New Roman";
                worksheet.Cell("K5").Style.Font.FontSize = 14;
                worksheet.Cell("K5").Style.Alignment.WrapText = true;
                worksheet.Range("K5:K7").Style.Border.OutsideBorder = XLBorderStyleValues.Thin;
                worksheet.Range("K5:K7").Merge();

                worksheet.Row(7).Height = worksheet.Row(7).Height * 2;

                //ДАННЫЕ

                //string query = $"SELECT Ticket.id, Clients.CompanyName AS [Клиент], Products.name AS [Продукт], (Worker.surname + ' ' + Worker.name + ' ' + Worker.patronymic) AS [Принявший сотрудник], Type.name AS [Тип проблемы], Ticket.header AS [Заголовок], Ticket.description AS [Описание], Ticket.priority AS [Приоритет], Ticket.status AS [Статус], Ticket.application_data AS [Дата подачи], Ticket.completion_data AS [Дата закрытия], Ticket.actual_data AS [Фактическая дата закрытия] FROM Ticket, Clients, Products, Worker, Type WHERE Ticket.client = Clients.id AND Ticket.product = Products.id AND Ticket.worker = Worker.id AND Ticket.type = Type.id AND Ticket.application_data >= {dateTimePicker1.Value.ToString()} AND Ticket.application_data <= {dateTimePicker2.Value.ToString()}";
                //string query = $"SELECT Clients.CompanyName AS [Клиент], Products.name AS [Продукт], (Worker.surname + ' ' + Worker.name + ' ' + Worker.patronymic) AS [Принявший сотрудник], Type.name AS [Тип проблемы], Ticket.header AS [Заголовок], Ticket.description AS [Описание], Ticket.priority AS [Приоритет], Ticket.status AS [Статус], Ticket.application_data AS [Дата подачи], Ticket.completion_data AS [Дата закрытия], Ticket.actual_data AS [Фактическая дата закрытия] FROM Ticket, Clients, Products, Worker, Type WHERE Ticket.client = Clients.id AND Ticket.product = Products.id AND Ticket.worker = Worker.id AND Ticket.type = Type.id AND Ticket.application_data >= CAST('{dateTimePicker1.Value.ToString("d")}' AS DATETIME) AND Ticket.application_data <= CAST('{dateTimePicker2.Value.ToString("d")}' AS DATETIME)";
                string query = $"SELECT Clients.CompanyName AS [Клиент], Products.name AS [Продукт], (Worker.surname + ' ' + Worker.name + ' ' + Worker.patronymic) AS [Принявший сотрудник], Type.name AS [Тип проблемы], Ticket.header AS [Заголовок], Ticket.description AS [Описание], Ticket.priority AS [Приоритет], Ticket.status AS [Статус], Ticket.application_data AS [Дата подачи], Ticket.completion_data AS [Дата закрытия], Ticket.actual_data AS [Фактическая дата закрытия] FROM Ticket, Clients, Products, Worker, Type WHERE Ticket.client = Clients.id AND Ticket.product = Products.id AND Ticket.worker = Worker.id AND Ticket.type = Type.id AND Ticket.client = {comboBox1.SelectedValue} ORDER BY Ticket.application_data ASC";
                using (SqlDataAdapter adapter = new SqlDataAdapter(query, connectionString))
                {
                    DataTable dataTable = new DataTable();
                    adapter.Fill(dataTable);

                    if (dataTable.Rows.Count == 0)
                    {
                        MessageBox.Show("Нет тикетов этого пользователя!", "Ошибка", MessageBoxButtons.OK, MessageBoxIcon.Error);
                        return;
                    }

                    DateTime fix = DateTime.Now;

                    for (int i = 0, b = 8, x = dataTable.Rows.Count; i < x; i++)
                    {
                        string n = "A" + b.ToString();

                        worksheet.Cell(n).Value = dataTable.Rows[i][0].ToString();
                        worksheet.Cell(n).Style.Alignment.Horizontal = XLAlignmentHorizontalValues.Center;
                        worksheet.Cell(n).Style.Alignment.Vertical = XLAlignmentVerticalValues.Center;
                        worksheet.Cell(n).Style.Font.FontName = "Times New Roman";
                        worksheet.Cell(n).Style.Font.FontSize = 14;
                        worksheet.Cell(n).Style.Alignment.WrapText = true;
                        worksheet.Cell(n).Style.Border.OutsideBorder = XLBorderStyleValues.Thin;

                        n = "B" + b.ToString();

                        worksheet.Cell(n).Value = dataTable.Rows[i][1].ToString();
                        worksheet.Cell(n).Style.Alignment.Horizontal = XLAlignmentHorizontalValues.Center;
                        worksheet.Cell(n).Style.Alignment.Vertical = XLAlignmentVerticalValues.Center;
                        worksheet.Cell(n).Style.Font.FontName = "Times New Roman";
                        worksheet.Cell(n).Style.Font.FontSize = 14;
                        worksheet.Cell(n).Style.Alignment.WrapText = true;
                        worksheet.Cell(n).Style.Border.OutsideBorder = XLBorderStyleValues.Thin;

                        n = "C" + b.ToString();

                        worksheet.Cell(n).Value = dataTable.Rows[i][2].ToString();
                        worksheet.Cell(n).Style.Alignment.Horizontal = XLAlignmentHorizontalValues.Center;
                        worksheet.Cell(n).Style.Alignment.Vertical = XLAlignmentVerticalValues.Center;
                        worksheet.Cell(n).Style.Font.FontName = "Times New Roman";
                        worksheet.Cell(n).Style.Font.FontSize = 14;
                        worksheet.Cell(n).Style.Alignment.WrapText = true;
                        worksheet.Cell(n).Style.Border.OutsideBorder = XLBorderStyleValues.Thin;

                        n = "D" + b.ToString();

                        worksheet.Cell(n).Value = dataTable.Rows[i][3].ToString();
                        worksheet.Cell(n).Style.Alignment.Horizontal = XLAlignmentHorizontalValues.Center;
                        worksheet.Cell(n).Style.Alignment.Vertical = XLAlignmentVerticalValues.Center;
                        worksheet.Cell(n).Style.Font.FontName = "Times New Roman";
                        worksheet.Cell(n).Style.Font.FontSize = 14;
                        worksheet.Cell(n).Style.Alignment.WrapText = true;
                        worksheet.Cell(n).Style.Border.OutsideBorder = XLBorderStyleValues.Thin;

                        n = "E" + b.ToString();

                        worksheet.Cell(n).Value = dataTable.Rows[i][4].ToString();
                        worksheet.Cell(n).Style.Alignment.Horizontal = XLAlignmentHorizontalValues.Center;
                        worksheet.Cell(n).Style.Alignment.Vertical = XLAlignmentVerticalValues.Center;
                        worksheet.Cell(n).Style.Font.FontName = "Times New Roman";
                        worksheet.Cell(n).Style.Font.FontSize = 14;
                        worksheet.Cell(n).Style.Alignment.WrapText = true;
                        worksheet.Cell(n).Style.Border.OutsideBorder = XLBorderStyleValues.Thin;

                        n = "F" + b.ToString();

                        worksheet.Cell(n).Value = dataTable.Rows[i][5].ToString();
                        worksheet.Cell(n).Style.Alignment.Horizontal = XLAlignmentHorizontalValues.Center;
                        worksheet.Cell(n).Style.Alignment.Vertical = XLAlignmentVerticalValues.Center;
                        worksheet.Cell(n).Style.Font.FontName = "Times New Roman";
                        worksheet.Cell(n).Style.Font.FontSize = 14;
                        worksheet.Cell(n).Style.Alignment.WrapText = true;
                        worksheet.Cell(n).Style.Border.OutsideBorder = XLBorderStyleValues.Thin;

                        n = "G" + b.ToString();

                        worksheet.Cell(n).Value = dataTable.Rows[i][6].ToString();
                        worksheet.Cell(n).Style.Alignment.Horizontal = XLAlignmentHorizontalValues.Center;
                        worksheet.Cell(n).Style.Alignment.Vertical = XLAlignmentVerticalValues.Center;
                        worksheet.Cell(n).Style.Font.FontName = "Times New Roman";
                        worksheet.Cell(n).Style.Font.FontSize = 14;
                        worksheet.Cell(n).Style.Alignment.WrapText = true;
                        worksheet.Cell(n).Style.Border.OutsideBorder = XLBorderStyleValues.Thin;

                        n = "H" + b.ToString();

                        //Статус
                        worksheet.Cell(n).Value = dataTable.Rows[i][7].ToString();
                        worksheet.Cell(n).Style.Alignment.Horizontal = XLAlignmentHorizontalValues.Center;
                        worksheet.Cell(n).Style.Alignment.Vertical = XLAlignmentVerticalValues.Center;
                        worksheet.Cell(n).Style.Font.FontName = "Times New Roman";
                        worksheet.Cell(n).Style.Font.FontSize = 14;
                        worksheet.Cell(n).Style.Alignment.WrapText = true;
                        worksheet.Cell(n).Style.Border.OutsideBorder = XLBorderStyleValues.Thin;

                        n = "I" + b.ToString();

                        //Даты
                        fix = (DateTime)dataTable.Rows[i][8];
                        worksheet.Cell(n).Value = fix.ToString("d");
                        worksheet.Cell(n).Style.Alignment.Horizontal = XLAlignmentHorizontalValues.Center;
                        worksheet.Cell(n).Style.Alignment.Vertical = XLAlignmentVerticalValues.Center;
                        worksheet.Cell(n).Style.Font.FontName = "Times New Roman";
                        worksheet.Cell(n).Style.Font.FontSize = 14;
                        worksheet.Cell(n).Style.Alignment.WrapText = true;
                        worksheet.Cell(n).Style.Border.OutsideBorder = XLBorderStyleValues.Thin;

                        n = "J" + b.ToString();

                        fix = (DateTime)dataTable.Rows[i][9];
                        worksheet.Cell(n).Value = fix.ToString("d");
                        worksheet.Cell(n).Style.Alignment.Horizontal = XLAlignmentHorizontalValues.Center;
                        worksheet.Cell(n).Style.Alignment.Vertical = XLAlignmentVerticalValues.Center;
                        worksheet.Cell(n).Style.Font.FontName = "Times New Roman";
                        worksheet.Cell(n).Style.Font.FontSize = 14;
                        worksheet.Cell(n).Style.Alignment.WrapText = true;
                        worksheet.Cell(n).Style.Border.OutsideBorder = XLBorderStyleValues.Thin;

                        n = "K" + b.ToString();

                        //Фактическая
                        if (dataTable.Rows[i][7].ToString() == "Закрыт")
                        {
                            fix = (DateTime)dataTable.Rows[i][10];
                            worksheet.Cell(n).Value = fix.ToString("d");
                        }
                        else
                        {
                            worksheet.Cell(n).Value = "Нет";
                        }
                        worksheet.Cell(n).Style.Alignment.Horizontal = XLAlignmentHorizontalValues.Center;
                        worksheet.Cell(n).Style.Alignment.Vertical = XLAlignmentVerticalValues.Center;
                        worksheet.Cell(n).Style.Font.FontName = "Times New Roman";
                        worksheet.Cell(n).Style.Font.FontSize = 14;
                        worksheet.Cell(n).Style.Alignment.WrapText = true;
                        worksheet.Cell(n).Style.Border.OutsideBorder = XLBorderStyleValues.Thin;

                        b++;
                    }
                }

                //worksheet.Column(1).AdjustToContents();
                //worksheet.Rows(1, 1000).AdjustToContents();

                //СОХРАНЕНИЕ

                //worksheet.Columns().AdjustToContents();

                saveFileDialog1.Filter = "Microsoft Excel Files (*.xlsx)|*.xlsx";
                saveFileDialog1.FileName = filename;
                if (saveFileDialog1.ShowDialog() == DialogResult.OK)
                {
                    try
                    {
                        string filename2 = saveFileDialog1.FileName;
                        workbook.SaveAs(filename2);
                        System.Diagnostics.Process.Start(filename2);
                        MessageBox.Show("Отчет сформирован!", "Успех!", MessageBoxButtons.OK, MessageBoxIcon.Asterisk);
                    }
                    catch
                    {
                        MessageBox.Show("Файл уже существует или произошла другая ошибка!", "Ошибка", MessageBoxButtons.OK, MessageBoxIcon.Error);
                    }
                }
                else
                {
                    MessageBox.Show("Вы отменили сохранение!", "Ошибка", MessageBoxButtons.OK, MessageBoxIcon.Error);
                    return;
                }
            }
        }

        private void button3_Click(object sender, EventArgs e)
        {
            if (comboBox2.SelectedValue == null)
                return;

            using (var workbook = new XLWorkbook())
            {
                string CompanyName;
                string query2 = $"SELECT name FROM Products WHERE id = {comboBox2.SelectedValue}";
                var connectionString = ConfigurationManager.ConnectionStrings["db"].ConnectionString;
                using (SqlDataAdapter adapter = new SqlDataAdapter(query2, connectionString))
                {
                    DataTable dataTable = new DataTable();
                    adapter.Fill(dataTable);

                    CompanyName = dataTable.Rows[0][0].ToString();
                }

                var worksheet = workbook.Worksheets.Add("Тикеты по ПО");
                string filename = "Тикеты по " + CompanyName + " за период (" + dateTimePicker3.Value.ToString("D") + " - " + dateTimePicker4.Value.ToString("D") + ").xlsx";

                StringBuilder sb = new StringBuilder(filename);
                filename = FixInvalidChars(sb, ' ');

                filename = System.Text.RegularExpressions.Regex.Replace(filename, @"\s+", " ");

                var col1 = worksheet.Column("A");
                col1.Width = col1.Width * 2.8;

                var col2 = worksheet.Column("B");
                col2.Width = col2.Width * 2.8;

                var col11 = worksheet.Column("C");
                col11.Width = col11.Width * 3;

                var col3 = worksheet.Column("D");
                col3.Width = col3.Width * 2.5;

                var col4 = worksheet.Column("E");
                col4.Width = col4.Width * 4.2;

                var col5 = worksheet.Column("F");
                col5.Width = col5.Width * 4.6;

                var col6 = worksheet.Column("G");
                col6.Width = col6.Width * 2;

                var col7 = worksheet.Column("H");
                col7.Width = col7.Width * 2;

                var col8 = worksheet.Column("I");
                col8.Width = col8.Width * 2.2;

                var col9 = worksheet.Column("J");
                col9.Width = col9.Width * 2.2;

                var col10 = worksheet.Column("K");
                col10.Width = col10.Width * 2.5;

                //ШАПКА

                worksheet.Cell("A1").Value = "Отчёт";
                worksheet.Cell("A1").Style.Alignment.Horizontal = XLAlignmentHorizontalValues.Center;
                worksheet.Cell("A1").Style.Alignment.Vertical = XLAlignmentVerticalValues.Center;
                worksheet.Cell("A1").Style.Font.Bold = true;
                worksheet.Cell("A1").Style.Font.FontName = "Times New Roman";
                worksheet.Cell("A1").Style.Font.FontSize = 16;
                worksheet.Range("A1:K1").Merge();

                string Shapka = "Тикеты по " + CompanyName + " за период";

                worksheet.Cell("A2").Value = Shapka;
                worksheet.Cell("A2").Style.Alignment.Horizontal = XLAlignmentHorizontalValues.Center;
                worksheet.Cell("A2").Style.Alignment.Vertical = XLAlignmentVerticalValues.Center;
                worksheet.Cell("A2").Style.Font.Bold = true;
                worksheet.Cell("A2").Style.Font.FontName = "Times New Roman";
                worksheet.Cell("A2").Style.Font.FontSize = 16;
                worksheet.Range("A2:K2").Merge();

                string Period = dateTimePicker3.Value.ToString("d") + " - " + dateTimePicker4.Value.ToString("d");

                worksheet.Cell("A3").Value = Period;
                worksheet.Cell("A3").Style.Alignment.Horizontal = XLAlignmentHorizontalValues.Center;
                worksheet.Cell("A3").Style.Alignment.Vertical = XLAlignmentVerticalValues.Center;
                worksheet.Cell("A3").Style.Font.Bold = true;
                worksheet.Cell("A3").Style.Font.FontName = "Times New Roman";
                worksheet.Cell("A3").Style.Font.FontSize = 16;
                worksheet.Range("A3:K3").Merge();

                string Form = "Дата формировани: " + DateTime.Now.ToString();

                worksheet.Cell("I4").Value = Form;
                worksheet.Cell("I4").Style.Alignment.Horizontal = XLAlignmentHorizontalValues.Right;
                worksheet.Cell("I4").Style.Alignment.Vertical = XLAlignmentVerticalValues.Center;
                worksheet.Cell("I4").Style.Font.Bold = true;
                worksheet.Cell("I4").Style.Font.FontName = "Times New Roman";
                worksheet.Cell("I4").Style.Font.FontSize = 16;
                worksheet.Range("I4:K4").Merge();

                worksheet.Row(1).Height = worksheet.Row(1).Height * 1.5;
                worksheet.Row(2).Height = worksheet.Row(2).Height * 1.5;
                worksheet.Row(3).Height = worksheet.Row(3).Height * 1.5;

                worksheet.Cell("A4").Value = "Сформировал: " + GetInitials(GetWorkerName());
                worksheet.Cell("A4").Style.Alignment.Horizontal = XLAlignmentHorizontalValues.Left;
                worksheet.Cell("A4").Style.Alignment.Vertical = XLAlignmentVerticalValues.Center;
                worksheet.Cell("A4").Style.Font.Bold = true;
                worksheet.Cell("A4").Style.Font.FontName = "Times New Roman";
                worksheet.Cell("A4").Style.Font.FontSize = 16;
                worksheet.Range("A4:C4").Merge();

                worksheet.Cell("E4").Value = GetMyCompany();
                worksheet.Cell("E4").Style.Alignment.Horizontal = XLAlignmentHorizontalValues.Center;
                worksheet.Cell("E4").Style.Alignment.Vertical = XLAlignmentVerticalValues.Center;
                worksheet.Cell("E4").Style.Font.Bold = true;
                worksheet.Cell("E4").Style.Font.FontName = "Times New Roman";
                worksheet.Cell("E4").Style.Font.FontSize = 16;
                worksheet.Range("E4:F4").Merge();

                //worksheet.Rows().AdjustToContents();
                //worksheet.Columns().AdjustToContents();

                //ТАБЛИЦА

                worksheet.Cell("A5").Value = "Клиент";
                worksheet.Cell("A5").Style.Alignment.Horizontal = XLAlignmentHorizontalValues.Center;
                worksheet.Cell("A5").Style.Alignment.Vertical = XLAlignmentVerticalValues.Center;
                worksheet.Cell("A5").Style.Font.Bold = true;
                worksheet.Cell("A5").Style.Font.FontName = "Times New Roman";
                worksheet.Cell("A5").Style.Font.FontSize = 14;
                worksheet.Cell("A5").Style.Alignment.WrapText = true;
                worksheet.Range("A5:A7").Style.Border.OutsideBorder = XLBorderStyleValues.Thin;
                worksheet.Range("A5:A7").Merge();

                worksheet.Cell("B5").Value = "Продукт";
                worksheet.Cell("B5").Style.Alignment.Horizontal = XLAlignmentHorizontalValues.Center;
                worksheet.Cell("B5").Style.Alignment.Vertical = XLAlignmentVerticalValues.Center;
                worksheet.Cell("B5").Style.Font.Bold = true;
                worksheet.Cell("B5").Style.Font.FontName = "Times New Roman";
                worksheet.Cell("B5").Style.Font.FontSize = 14;
                worksheet.Cell("B5").Style.Alignment.WrapText = true;
                worksheet.Range("B5:B7").Style.Border.OutsideBorder = XLBorderStyleValues.Thin;
                worksheet.Range("B5:B7").Merge();

                worksheet.Cell("C5").Value = "Принявший сотрудник";
                worksheet.Cell("C5").Style.Alignment.Horizontal = XLAlignmentHorizontalValues.Center;
                worksheet.Cell("C5").Style.Alignment.Vertical = XLAlignmentVerticalValues.Center;
                worksheet.Cell("C5").Style.Font.Bold = true;
                worksheet.Cell("C5").Style.Font.FontName = "Times New Roman";
                worksheet.Cell("C5").Style.Font.FontSize = 14;
                worksheet.Cell("C5").Style.Alignment.WrapText = true;
                worksheet.Range("C5:C7").Style.Border.OutsideBorder = XLBorderStyleValues.Thin;
                worksheet.Range("C5:C7").Merge();

                worksheet.Cell("D5").Value = "Тип проблемы";
                worksheet.Cell("D5").Style.Alignment.Horizontal = XLAlignmentHorizontalValues.Center;
                worksheet.Cell("D5").Style.Alignment.Vertical = XLAlignmentVerticalValues.Center;
                worksheet.Cell("D5").Style.Font.Bold = true;
                worksheet.Cell("D5").Style.Font.FontName = "Times New Roman";
                worksheet.Cell("D5").Style.Font.FontSize = 14;
                worksheet.Cell("D5").Style.Alignment.WrapText = true;
                worksheet.Range("D5:D7").Style.Border.OutsideBorder = XLBorderStyleValues.Thin;
                worksheet.Range("D5:D7").Merge();

                worksheet.Cell("E5").Value = "Заголовок";
                worksheet.Cell("E5").Style.Alignment.Horizontal = XLAlignmentHorizontalValues.Center;
                worksheet.Cell("E5").Style.Alignment.Vertical = XLAlignmentVerticalValues.Center;
                worksheet.Cell("E5").Style.Font.Bold = true;
                worksheet.Cell("E5").Style.Font.FontName = "Times New Roman";
                worksheet.Cell("E5").Style.Font.FontSize = 14;
                worksheet.Cell("E5").Style.Alignment.WrapText = true;
                worksheet.Range("E5:E7").Style.Border.OutsideBorder = XLBorderStyleValues.Thin;
                worksheet.Range("E5:E7").Merge();

                worksheet.Cell("F5").Value = "Описание";
                worksheet.Cell("F5").Style.Alignment.Horizontal = XLAlignmentHorizontalValues.Center;
                worksheet.Cell("F5").Style.Alignment.Vertical = XLAlignmentVerticalValues.Center;
                worksheet.Cell("F5").Style.Font.Bold = true;
                worksheet.Cell("F5").Style.Font.FontName = "Times New Roman";
                worksheet.Cell("F5").Style.Font.FontSize = 14;
                worksheet.Cell("F5").Style.Alignment.WrapText = true;
                worksheet.Range("F5:F7").Style.Border.OutsideBorder = XLBorderStyleValues.Thin;
                worksheet.Range("F5:F7").Merge();

                worksheet.Cell("G5").Value = "Приоритет";
                worksheet.Cell("G5").Style.Alignment.Horizontal = XLAlignmentHorizontalValues.Center;
                worksheet.Cell("G5").Style.Alignment.Vertical = XLAlignmentVerticalValues.Center;
                worksheet.Cell("G5").Style.Font.Bold = true;
                worksheet.Cell("G5").Style.Font.FontName = "Times New Roman";
                worksheet.Cell("G5").Style.Font.FontSize = 14;
                worksheet.Cell("G5").Style.Alignment.WrapText = true;
                worksheet.Range("G5:G7").Style.Border.OutsideBorder = XLBorderStyleValues.Thin;
                worksheet.Range("G5:G7").Merge();

                worksheet.Cell("H5").Value = "Статус";
                worksheet.Cell("H5").Style.Alignment.Horizontal = XLAlignmentHorizontalValues.Center;
                worksheet.Cell("H5").Style.Alignment.Vertical = XLAlignmentVerticalValues.Center;
                worksheet.Cell("H5").Style.Font.Bold = true;
                worksheet.Cell("H5").Style.Font.FontName = "Times New Roman";
                worksheet.Cell("H5").Style.Font.FontSize = 14;
                worksheet.Cell("H5").Style.Alignment.WrapText = true;
                worksheet.Range("H5:H7").Style.Border.OutsideBorder = XLBorderStyleValues.Thin;
                worksheet.Range("H5:H7").Merge();

                worksheet.Cell("I5").Value = "Дата подачи";
                worksheet.Cell("I5").Style.Alignment.Horizontal = XLAlignmentHorizontalValues.Center;
                worksheet.Cell("I5").Style.Alignment.Vertical = XLAlignmentVerticalValues.Center;
                worksheet.Cell("I5").Style.Font.Bold = true;
                worksheet.Cell("I5").Style.Font.FontName = "Times New Roman";
                worksheet.Cell("I5").Style.Font.FontSize = 14;
                worksheet.Cell("I5").Style.Alignment.WrapText = true;
                worksheet.Range("I5:I7").Style.Border.OutsideBorder = XLBorderStyleValues.Thin;
                worksheet.Range("I5:I7").Merge();

                worksheet.Cell("J5").Value = "Дата закрытия";
                worksheet.Cell("J5").Style.Alignment.Horizontal = XLAlignmentHorizontalValues.Center;
                worksheet.Cell("J5").Style.Alignment.Vertical = XLAlignmentVerticalValues.Center;
                worksheet.Cell("J5").Style.Font.Bold = true;
                worksheet.Cell("J5").Style.Font.FontName = "Times New Roman";
                worksheet.Cell("J5").Style.Font.FontSize = 14;
                worksheet.Cell("J5").Style.Alignment.WrapText = true;
                worksheet.Range("J5:J7").Style.Border.OutsideBorder = XLBorderStyleValues.Thin;
                worksheet.Range("J5:J7").Merge();

                worksheet.Cell("K5").Value = "Фактическая дата закрытия";
                worksheet.Cell("K5").Style.Alignment.Horizontal = XLAlignmentHorizontalValues.Center;
                worksheet.Cell("K5").Style.Alignment.Vertical = XLAlignmentVerticalValues.Center;
                worksheet.Cell("K5").Style.Font.Bold = true;
                worksheet.Cell("K5").Style.Font.FontName = "Times New Roman";
                worksheet.Cell("K5").Style.Font.FontSize = 14;
                worksheet.Cell("K5").Style.Alignment.WrapText = true;
                worksheet.Range("K5:K7").Style.Border.OutsideBorder = XLBorderStyleValues.Thin;
                worksheet.Range("K5:K7").Merge();

                worksheet.Row(7).Height = worksheet.Row(7).Height * 2;

                //ДАННЫЕ

                //string query = $"SELECT Ticket.id, Clients.CompanyName AS [Клиент], Products.name AS [Продукт], (Worker.surname + ' ' + Worker.name + ' ' + Worker.patronymic) AS [Принявший сотрудник], Type.name AS [Тип проблемы], Ticket.header AS [Заголовок], Ticket.description AS [Описание], Ticket.priority AS [Приоритет], Ticket.status AS [Статус], Ticket.application_data AS [Дата подачи], Ticket.completion_data AS [Дата закрытия], Ticket.actual_data AS [Фактическая дата закрытия] FROM Ticket, Clients, Products, Worker, Type WHERE Ticket.client = Clients.id AND Ticket.product = Products.id AND Ticket.worker = Worker.id AND Ticket.type = Type.id AND Ticket.application_data >= {dateTimePicker1.Value.ToString()} AND Ticket.application_data <= {dateTimePicker2.Value.ToString()}";
                //string query = $"SELECT Clients.CompanyName AS [Клиент], Products.name AS [Продукт], (Worker.surname + ' ' + Worker.name + ' ' + Worker.patronymic) AS [Принявший сотрудник], Type.name AS [Тип проблемы], Ticket.header AS [Заголовок], Ticket.description AS [Описание], Ticket.priority AS [Приоритет], Ticket.status AS [Статус], Ticket.application_data AS [Дата подачи], Ticket.completion_data AS [Дата закрытия], Ticket.actual_data AS [Фактическая дата закрытия] FROM Ticket, Clients, Products, Worker, Type WHERE Ticket.client = Clients.id AND Ticket.product = Products.id AND Ticket.worker = Worker.id AND Ticket.type = Type.id AND Ticket.application_data >= CAST('{dateTimePicker1.Value.ToString("d")}' AS DATETIME) AND Ticket.application_data <= CAST('{dateTimePicker2.Value.ToString("d")}' AS DATETIME)";
                string query = $"SELECT Clients.CompanyName AS [Клиент], Products.name AS [Продукт], (Worker.surname + ' ' + Worker.name + ' ' + Worker.patronymic) AS [Принявший сотрудник], Type.name AS [Тип проблемы], Ticket.header AS [Заголовок], Ticket.description AS [Описание], Ticket.priority AS [Приоритет], Ticket.status AS [Статус], Ticket.application_data AS [Дата подачи], Ticket.completion_data AS [Дата закрытия], Ticket.actual_data AS [Фактическая дата закрытия] FROM Ticket, Clients, Products, Worker, Type WHERE Ticket.client = Clients.id AND Ticket.product = Products.id AND Ticket.worker = Worker.id AND Ticket.type = Type.id AND Ticket.application_data >= CAST('{dateTimePicker3.Value.ToString()}' AS DATETIME) AND Ticket.application_data <= CAST('{dateTimePicker4.Value.ToString()}' AS DATETIME) AND Ticket.product = {comboBox2.SelectedValue} ORDER BY Ticket.application_data ASC";
                using (SqlDataAdapter adapter = new SqlDataAdapter(query, connectionString))
                {
                    DataTable dataTable = new DataTable();
                    adapter.Fill(dataTable);

                    if (dataTable.Rows.Count == 0)
                    {
                        MessageBox.Show("Нет тикетов по этому программному продукту!", "Ошибка", MessageBoxButtons.OK, MessageBoxIcon.Error);
                        return;
                    }

                    DateTime fix = DateTime.Now;

                    for (int i = 0, b = 8, x = dataTable.Rows.Count; i < x; i++)
                    {
                        string n = "A" + b.ToString();

                        worksheet.Cell(n).Value = dataTable.Rows[i][0].ToString();
                        worksheet.Cell(n).Style.Alignment.Horizontal = XLAlignmentHorizontalValues.Center;
                        worksheet.Cell(n).Style.Alignment.Vertical = XLAlignmentVerticalValues.Center;
                        worksheet.Cell(n).Style.Font.FontName = "Times New Roman";
                        worksheet.Cell(n).Style.Font.FontSize = 14;
                        worksheet.Cell(n).Style.Alignment.WrapText = true;
                        worksheet.Cell(n).Style.Border.OutsideBorder = XLBorderStyleValues.Thin;

                        n = "B" + b.ToString();

                        worksheet.Cell(n).Value = dataTable.Rows[i][1].ToString();
                        worksheet.Cell(n).Style.Alignment.Horizontal = XLAlignmentHorizontalValues.Center;
                        worksheet.Cell(n).Style.Alignment.Vertical = XLAlignmentVerticalValues.Center;
                        worksheet.Cell(n).Style.Font.FontName = "Times New Roman";
                        worksheet.Cell(n).Style.Font.FontSize = 14;
                        worksheet.Cell(n).Style.Alignment.WrapText = true;
                        worksheet.Cell(n).Style.Border.OutsideBorder = XLBorderStyleValues.Thin;

                        n = "C" + b.ToString();

                        worksheet.Cell(n).Value = dataTable.Rows[i][2].ToString();
                        worksheet.Cell(n).Style.Alignment.Horizontal = XLAlignmentHorizontalValues.Center;
                        worksheet.Cell(n).Style.Alignment.Vertical = XLAlignmentVerticalValues.Center;
                        worksheet.Cell(n).Style.Font.FontName = "Times New Roman";
                        worksheet.Cell(n).Style.Font.FontSize = 14;
                        worksheet.Cell(n).Style.Alignment.WrapText = true;
                        worksheet.Cell(n).Style.Border.OutsideBorder = XLBorderStyleValues.Thin;

                        n = "D" + b.ToString();

                        worksheet.Cell(n).Value = dataTable.Rows[i][3].ToString();
                        worksheet.Cell(n).Style.Alignment.Horizontal = XLAlignmentHorizontalValues.Center;
                        worksheet.Cell(n).Style.Alignment.Vertical = XLAlignmentVerticalValues.Center;
                        worksheet.Cell(n).Style.Font.FontName = "Times New Roman";
                        worksheet.Cell(n).Style.Font.FontSize = 14;
                        worksheet.Cell(n).Style.Alignment.WrapText = true;
                        worksheet.Cell(n).Style.Border.OutsideBorder = XLBorderStyleValues.Thin;

                        n = "E" + b.ToString();

                        worksheet.Cell(n).Value = dataTable.Rows[i][4].ToString();
                        worksheet.Cell(n).Style.Alignment.Horizontal = XLAlignmentHorizontalValues.Center;
                        worksheet.Cell(n).Style.Alignment.Vertical = XLAlignmentVerticalValues.Center;
                        worksheet.Cell(n).Style.Font.FontName = "Times New Roman";
                        worksheet.Cell(n).Style.Font.FontSize = 14;
                        worksheet.Cell(n).Style.Alignment.WrapText = true;
                        worksheet.Cell(n).Style.Border.OutsideBorder = XLBorderStyleValues.Thin;

                        n = "F" + b.ToString();

                        worksheet.Cell(n).Value = dataTable.Rows[i][5].ToString();
                        worksheet.Cell(n).Style.Alignment.Horizontal = XLAlignmentHorizontalValues.Center;
                        worksheet.Cell(n).Style.Alignment.Vertical = XLAlignmentVerticalValues.Center;
                        worksheet.Cell(n).Style.Font.FontName = "Times New Roman";
                        worksheet.Cell(n).Style.Font.FontSize = 14;
                        worksheet.Cell(n).Style.Alignment.WrapText = true;
                        worksheet.Cell(n).Style.Border.OutsideBorder = XLBorderStyleValues.Thin;

                        n = "G" + b.ToString();

                        worksheet.Cell(n).Value = dataTable.Rows[i][6].ToString();
                        worksheet.Cell(n).Style.Alignment.Horizontal = XLAlignmentHorizontalValues.Center;
                        worksheet.Cell(n).Style.Alignment.Vertical = XLAlignmentVerticalValues.Center;
                        worksheet.Cell(n).Style.Font.FontName = "Times New Roman";
                        worksheet.Cell(n).Style.Font.FontSize = 14;
                        worksheet.Cell(n).Style.Alignment.WrapText = true;
                        worksheet.Cell(n).Style.Border.OutsideBorder = XLBorderStyleValues.Thin;

                        n = "H" + b.ToString();

                        //Статус
                        worksheet.Cell(n).Value = dataTable.Rows[i][7].ToString();
                        worksheet.Cell(n).Style.Alignment.Horizontal = XLAlignmentHorizontalValues.Center;
                        worksheet.Cell(n).Style.Alignment.Vertical = XLAlignmentVerticalValues.Center;
                        worksheet.Cell(n).Style.Font.FontName = "Times New Roman";
                        worksheet.Cell(n).Style.Font.FontSize = 14;
                        worksheet.Cell(n).Style.Alignment.WrapText = true;
                        worksheet.Cell(n).Style.Border.OutsideBorder = XLBorderStyleValues.Thin;

                        n = "I" + b.ToString();

                        //Даты
                        fix = (DateTime)dataTable.Rows[i][8];
                        worksheet.Cell(n).Value = fix.ToString("d");
                        worksheet.Cell(n).Style.Alignment.Horizontal = XLAlignmentHorizontalValues.Center;
                        worksheet.Cell(n).Style.Alignment.Vertical = XLAlignmentVerticalValues.Center;
                        worksheet.Cell(n).Style.Font.FontName = "Times New Roman";
                        worksheet.Cell(n).Style.Font.FontSize = 14;
                        worksheet.Cell(n).Style.Alignment.WrapText = true;
                        worksheet.Cell(n).Style.Border.OutsideBorder = XLBorderStyleValues.Thin;

                        n = "J" + b.ToString();

                        fix = (DateTime)dataTable.Rows[i][9];
                        worksheet.Cell(n).Value = fix.ToString("d");
                        worksheet.Cell(n).Style.Alignment.Horizontal = XLAlignmentHorizontalValues.Center;
                        worksheet.Cell(n).Style.Alignment.Vertical = XLAlignmentVerticalValues.Center;
                        worksheet.Cell(n).Style.Font.FontName = "Times New Roman";
                        worksheet.Cell(n).Style.Font.FontSize = 14;
                        worksheet.Cell(n).Style.Alignment.WrapText = true;
                        worksheet.Cell(n).Style.Border.OutsideBorder = XLBorderStyleValues.Thin;

                        n = "K" + b.ToString();

                        //Фактическая
                        if (dataTable.Rows[i][7].ToString() == "Закрыт")
                        {
                            fix = (DateTime)dataTable.Rows[i][10];
                            worksheet.Cell(n).Value = fix.ToString("d");
                        }
                        else
                        {
                            worksheet.Cell(n).Value = "Нет";
                        }
                        worksheet.Cell(n).Style.Alignment.Horizontal = XLAlignmentHorizontalValues.Center;
                        worksheet.Cell(n).Style.Alignment.Vertical = XLAlignmentVerticalValues.Center;
                        worksheet.Cell(n).Style.Font.FontName = "Times New Roman";
                        worksheet.Cell(n).Style.Font.FontSize = 14;
                        worksheet.Cell(n).Style.Alignment.WrapText = true;
                        worksheet.Cell(n).Style.Border.OutsideBorder = XLBorderStyleValues.Thin;

                        b++;
                    }
                }

                //worksheet.Column(1).AdjustToContents();
                //worksheet.Rows(1, 1000).AdjustToContents();

                //СОХРАНЕНИЕ

                //worksheet.Columns().AdjustToContents();

                saveFileDialog1.Filter = "Microsoft Excel Files (*.xlsx)|*.xlsx";
                saveFileDialog1.FileName = filename;
                if (saveFileDialog1.ShowDialog() == DialogResult.OK)
                {
                    try
                    {
                        string filename2 = saveFileDialog1.FileName;
                        workbook.SaveAs(filename2);
                        System.Diagnostics.Process.Start(filename2);
                        MessageBox.Show("Отчет сформирован!", "Успех!", MessageBoxButtons.OK, MessageBoxIcon.Asterisk);
                    }
                    catch
                    {
                        MessageBox.Show("Файл уже существует или произошла другая ошибка!", "Ошибка", MessageBoxButtons.OK, MessageBoxIcon.Error);
                    }
                }
                else
                {
                    MessageBox.Show("Вы отменили сохранение!", "Ошибка", MessageBoxButtons.OK, MessageBoxIcon.Error);
                    return;
                }
            }
        }
        public static string FixInvalidChars(StringBuilder text, char replaceTo)
        {
            if (text == null)
            {
                throw new ArgumentNullException("text");
            }
            if (text.Length <= 0)
            {
                return text.ToString();
            }

            foreach (char badChar in Path.GetInvalidPathChars())
            {
                text.Replace(badChar, replaceTo);
            }

            foreach (char badChar in new[] { '?', '\\', '/', ':', '"', '*', '>', '<', '|' })
            {
                text.Replace(badChar, replaceTo);
            }
            return text.ToString();
        }

        private void button4_Click(object sender, EventArgs e)
        {
            if (comboBox3.SelectedValue == null)
                return;

            using (var workbook = new XLWorkbook())
            {
                string CompanyName;
                string query2 = $"SELECT (Worker.surname + ' ' + Worker.name + ' ' + Worker.patronymic) FROM Worker WHERE id = {comboBox3.SelectedValue}";
                var connectionString = ConfigurationManager.ConnectionStrings["db"].ConnectionString;
                using (SqlDataAdapter adapter = new SqlDataAdapter(query2, connectionString))
                {
                    DataTable dataTable = new DataTable();
                    adapter.Fill(dataTable);

                    CompanyName = dataTable.Rows[0][0].ToString();
                }

                string initsal = GetInitials(CompanyName);

                var worksheet = workbook.Worksheets.Add("Тикеты сотрудника");
                string filename = "Тикеты принятые " + initsal + " за период (" + dateTimePicker3.Value.ToString("D") + " - " + dateTimePicker4.Value.ToString("D") + ").xlsx";

                StringBuilder sb = new StringBuilder(filename);
                filename = FixInvalidChars(sb, ' ');

                filename = System.Text.RegularExpressions.Regex.Replace(filename, @"\s+", " ");

                var col1 = worksheet.Column("A");
                col1.Width = col1.Width * 2.8;

                var col2 = worksheet.Column("B");
                col2.Width = col2.Width * 2.8;

                var col11 = worksheet.Column("C");
                col11.Width = col11.Width * 3;

                var col3 = worksheet.Column("D");
                col3.Width = col3.Width * 2.5;

                var col4 = worksheet.Column("E");
                col4.Width = col4.Width * 4.2;

                var col5 = worksheet.Column("F");
                col5.Width = col5.Width * 4.6;

                var col6 = worksheet.Column("G");
                col6.Width = col6.Width * 2;

                var col7 = worksheet.Column("H");
                col7.Width = col7.Width * 2;

                var col8 = worksheet.Column("I");
                col8.Width = col8.Width * 2.2;

                var col9 = worksheet.Column("J");
                col9.Width = col9.Width * 2.2;

                var col10 = worksheet.Column("K");
                col10.Width = col10.Width * 2.5;

                //ШАПКА

                worksheet.Cell("A1").Value = "Отчёт";
                worksheet.Cell("A1").Style.Alignment.Horizontal = XLAlignmentHorizontalValues.Center;
                worksheet.Cell("A1").Style.Alignment.Vertical = XLAlignmentVerticalValues.Center;
                worksheet.Cell("A1").Style.Font.Bold = true;
                worksheet.Cell("A1").Style.Font.FontName = "Times New Roman";
                worksheet.Cell("A1").Style.Font.FontSize = 16;
                worksheet.Range("A1:K1").Merge();

                string Shapka = "Тикеты принятые " + initsal + " за период";

                worksheet.Cell("A2").Value = Shapka;
                worksheet.Cell("A2").Style.Alignment.Horizontal = XLAlignmentHorizontalValues.Center;
                worksheet.Cell("A2").Style.Alignment.Vertical = XLAlignmentVerticalValues.Center;
                worksheet.Cell("A2").Style.Font.Bold = true;
                worksheet.Cell("A2").Style.Font.FontName = "Times New Roman";
                worksheet.Cell("A2").Style.Font.FontSize = 16;
                worksheet.Range("A2:K2").Merge();

                string Period = dateTimePicker5.Value.ToString("d") + " - " + dateTimePicker6.Value.ToString("d");

                worksheet.Cell("A3").Value = Period;
                worksheet.Cell("A3").Style.Alignment.Horizontal = XLAlignmentHorizontalValues.Center;
                worksheet.Cell("A3").Style.Alignment.Vertical = XLAlignmentVerticalValues.Center;
                worksheet.Cell("A3").Style.Font.Bold = true;
                worksheet.Cell("A3").Style.Font.FontName = "Times New Roman";
                worksheet.Cell("A3").Style.Font.FontSize = 16;
                worksheet.Range("A3:K3").Merge();

                string Form = "Дата формировани: " + DateTime.Now.ToString();

                worksheet.Cell("I4").Value = Form;
                worksheet.Cell("I4").Style.Alignment.Horizontal = XLAlignmentHorizontalValues.Right;
                worksheet.Cell("I4").Style.Alignment.Vertical = XLAlignmentVerticalValues.Center;
                worksheet.Cell("I4").Style.Font.Bold = true;
                worksheet.Cell("I4").Style.Font.FontName = "Times New Roman";
                worksheet.Cell("I4").Style.Font.FontSize = 16;
                worksheet.Range("I4:K4").Merge();

                worksheet.Row(1).Height = worksheet.Row(1).Height * 1.5;
                worksheet.Row(2).Height = worksheet.Row(2).Height * 1.5;
                worksheet.Row(3).Height = worksheet.Row(3).Height * 1.5;

                worksheet.Cell("A4").Value = "Сформировал: " + GetInitials(GetWorkerName());
                worksheet.Cell("A4").Style.Alignment.Horizontal = XLAlignmentHorizontalValues.Left;
                worksheet.Cell("A4").Style.Alignment.Vertical = XLAlignmentVerticalValues.Center;
                worksheet.Cell("A4").Style.Font.Bold = true;
                worksheet.Cell("A4").Style.Font.FontName = "Times New Roman";
                worksheet.Cell("A4").Style.Font.FontSize = 16;
                worksheet.Range("A4:C4").Merge();

                worksheet.Cell("E4").Value = GetMyCompany();
                worksheet.Cell("E4").Style.Alignment.Horizontal = XLAlignmentHorizontalValues.Center;
                worksheet.Cell("E4").Style.Alignment.Vertical = XLAlignmentVerticalValues.Center;
                worksheet.Cell("E4").Style.Font.Bold = true;
                worksheet.Cell("E4").Style.Font.FontName = "Times New Roman";
                worksheet.Cell("E4").Style.Font.FontSize = 16;
                worksheet.Range("E4:F4").Merge();

                //worksheet.Rows().AdjustToContents();
                //worksheet.Columns().AdjustToContents();

                //ТАБЛИЦА

                worksheet.Cell("A5").Value = "Клиент";
                worksheet.Cell("A5").Style.Alignment.Horizontal = XLAlignmentHorizontalValues.Center;
                worksheet.Cell("A5").Style.Alignment.Vertical = XLAlignmentVerticalValues.Center;
                worksheet.Cell("A5").Style.Font.Bold = true;
                worksheet.Cell("A5").Style.Font.FontName = "Times New Roman";
                worksheet.Cell("A5").Style.Font.FontSize = 14;
                worksheet.Cell("A5").Style.Alignment.WrapText = true;
                worksheet.Range("A5:A7").Style.Border.OutsideBorder = XLBorderStyleValues.Thin;
                worksheet.Range("A5:A7").Merge();

                worksheet.Cell("B5").Value = "Продукт";
                worksheet.Cell("B5").Style.Alignment.Horizontal = XLAlignmentHorizontalValues.Center;
                worksheet.Cell("B5").Style.Alignment.Vertical = XLAlignmentVerticalValues.Center;
                worksheet.Cell("B5").Style.Font.Bold = true;
                worksheet.Cell("B5").Style.Font.FontName = "Times New Roman";
                worksheet.Cell("B5").Style.Font.FontSize = 14;
                worksheet.Cell("B5").Style.Alignment.WrapText = true;
                worksheet.Range("B5:B7").Style.Border.OutsideBorder = XLBorderStyleValues.Thin;
                worksheet.Range("B5:B7").Merge();

                worksheet.Cell("C5").Value = "Принявший сотрудник";
                worksheet.Cell("C5").Style.Alignment.Horizontal = XLAlignmentHorizontalValues.Center;
                worksheet.Cell("C5").Style.Alignment.Vertical = XLAlignmentVerticalValues.Center;
                worksheet.Cell("C5").Style.Font.Bold = true;
                worksheet.Cell("C5").Style.Font.FontName = "Times New Roman";
                worksheet.Cell("C5").Style.Font.FontSize = 14;
                worksheet.Cell("C5").Style.Alignment.WrapText = true;
                worksheet.Range("C5:C7").Style.Border.OutsideBorder = XLBorderStyleValues.Thin;
                worksheet.Range("C5:C7").Merge();

                worksheet.Cell("D5").Value = "Тип проблемы";
                worksheet.Cell("D5").Style.Alignment.Horizontal = XLAlignmentHorizontalValues.Center;
                worksheet.Cell("D5").Style.Alignment.Vertical = XLAlignmentVerticalValues.Center;
                worksheet.Cell("D5").Style.Font.Bold = true;
                worksheet.Cell("D5").Style.Font.FontName = "Times New Roman";
                worksheet.Cell("D5").Style.Font.FontSize = 14;
                worksheet.Cell("D5").Style.Alignment.WrapText = true;
                worksheet.Range("D5:D7").Style.Border.OutsideBorder = XLBorderStyleValues.Thin;
                worksheet.Range("D5:D7").Merge();

                worksheet.Cell("E5").Value = "Заголовок";
                worksheet.Cell("E5").Style.Alignment.Horizontal = XLAlignmentHorizontalValues.Center;
                worksheet.Cell("E5").Style.Alignment.Vertical = XLAlignmentVerticalValues.Center;
                worksheet.Cell("E5").Style.Font.Bold = true;
                worksheet.Cell("E5").Style.Font.FontName = "Times New Roman";
                worksheet.Cell("E5").Style.Font.FontSize = 14;
                worksheet.Cell("E5").Style.Alignment.WrapText = true;
                worksheet.Range("E5:E7").Style.Border.OutsideBorder = XLBorderStyleValues.Thin;
                worksheet.Range("E5:E7").Merge();

                worksheet.Cell("F5").Value = "Описание";
                worksheet.Cell("F5").Style.Alignment.Horizontal = XLAlignmentHorizontalValues.Center;
                worksheet.Cell("F5").Style.Alignment.Vertical = XLAlignmentVerticalValues.Center;
                worksheet.Cell("F5").Style.Font.Bold = true;
                worksheet.Cell("F5").Style.Font.FontName = "Times New Roman";
                worksheet.Cell("F5").Style.Font.FontSize = 14;
                worksheet.Cell("F5").Style.Alignment.WrapText = true;
                worksheet.Range("F5:F7").Style.Border.OutsideBorder = XLBorderStyleValues.Thin;
                worksheet.Range("F5:F7").Merge();

                worksheet.Cell("G5").Value = "Приоритет";
                worksheet.Cell("G5").Style.Alignment.Horizontal = XLAlignmentHorizontalValues.Center;
                worksheet.Cell("G5").Style.Alignment.Vertical = XLAlignmentVerticalValues.Center;
                worksheet.Cell("G5").Style.Font.Bold = true;
                worksheet.Cell("G5").Style.Font.FontName = "Times New Roman";
                worksheet.Cell("G5").Style.Font.FontSize = 14;
                worksheet.Cell("G5").Style.Alignment.WrapText = true;
                worksheet.Range("G5:G7").Style.Border.OutsideBorder = XLBorderStyleValues.Thin;
                worksheet.Range("G5:G7").Merge();

                worksheet.Cell("H5").Value = "Статус";
                worksheet.Cell("H5").Style.Alignment.Horizontal = XLAlignmentHorizontalValues.Center;
                worksheet.Cell("H5").Style.Alignment.Vertical = XLAlignmentVerticalValues.Center;
                worksheet.Cell("H5").Style.Font.Bold = true;
                worksheet.Cell("H5").Style.Font.FontName = "Times New Roman";
                worksheet.Cell("H5").Style.Font.FontSize = 14;
                worksheet.Cell("H5").Style.Alignment.WrapText = true;
                worksheet.Range("H5:H7").Style.Border.OutsideBorder = XLBorderStyleValues.Thin;
                worksheet.Range("H5:H7").Merge();

                worksheet.Cell("I5").Value = "Дата подачи";
                worksheet.Cell("I5").Style.Alignment.Horizontal = XLAlignmentHorizontalValues.Center;
                worksheet.Cell("I5").Style.Alignment.Vertical = XLAlignmentVerticalValues.Center;
                worksheet.Cell("I5").Style.Font.Bold = true;
                worksheet.Cell("I5").Style.Font.FontName = "Times New Roman";
                worksheet.Cell("I5").Style.Font.FontSize = 14;
                worksheet.Cell("I5").Style.Alignment.WrapText = true;
                worksheet.Range("I5:I7").Style.Border.OutsideBorder = XLBorderStyleValues.Thin;
                worksheet.Range("I5:I7").Merge();

                worksheet.Cell("J5").Value = "Дата закрытия";
                worksheet.Cell("J5").Style.Alignment.Horizontal = XLAlignmentHorizontalValues.Center;
                worksheet.Cell("J5").Style.Alignment.Vertical = XLAlignmentVerticalValues.Center;
                worksheet.Cell("J5").Style.Font.Bold = true;
                worksheet.Cell("J5").Style.Font.FontName = "Times New Roman";
                worksheet.Cell("J5").Style.Font.FontSize = 14;
                worksheet.Cell("J5").Style.Alignment.WrapText = true;
                worksheet.Range("J5:J7").Style.Border.OutsideBorder = XLBorderStyleValues.Thin;
                worksheet.Range("J5:J7").Merge();

                worksheet.Cell("K5").Value = "Фактическая дата закрытия";
                worksheet.Cell("K5").Style.Alignment.Horizontal = XLAlignmentHorizontalValues.Center;
                worksheet.Cell("K5").Style.Alignment.Vertical = XLAlignmentVerticalValues.Center;
                worksheet.Cell("K5").Style.Font.Bold = true;
                worksheet.Cell("K5").Style.Font.FontName = "Times New Roman";
                worksheet.Cell("K5").Style.Font.FontSize = 14;
                worksheet.Cell("K5").Style.Alignment.WrapText = true;
                worksheet.Range("K5:K7").Style.Border.OutsideBorder = XLBorderStyleValues.Thin;
                worksheet.Range("K5:K7").Merge();

                worksheet.Row(7).Height = worksheet.Row(7).Height * 2;

                //ДАННЫЕ

                //string query = $"SELECT Ticket.id, Clients.CompanyName AS [Клиент], Products.name AS [Продукт], (Worker.surname + ' ' + Worker.name + ' ' + Worker.patronymic) AS [Принявший сотрудник], Type.name AS [Тип проблемы], Ticket.header AS [Заголовок], Ticket.description AS [Описание], Ticket.priority AS [Приоритет], Ticket.status AS [Статус], Ticket.application_data AS [Дата подачи], Ticket.completion_data AS [Дата закрытия], Ticket.actual_data AS [Фактическая дата закрытия] FROM Ticket, Clients, Products, Worker, Type WHERE Ticket.client = Clients.id AND Ticket.product = Products.id AND Ticket.worker = Worker.id AND Ticket.type = Type.id AND Ticket.application_data >= {dateTimePicker1.Value.ToString()} AND Ticket.application_data <= {dateTimePicker2.Value.ToString()}";
                //string query = $"SELECT Clients.CompanyName AS [Клиент], Products.name AS [Продукт], (Worker.surname + ' ' + Worker.name + ' ' + Worker.patronymic) AS [Принявший сотрудник], Type.name AS [Тип проблемы], Ticket.header AS [Заголовок], Ticket.description AS [Описание], Ticket.priority AS [Приоритет], Ticket.status AS [Статус], Ticket.application_data AS [Дата подачи], Ticket.completion_data AS [Дата закрытия], Ticket.actual_data AS [Фактическая дата закрытия] FROM Ticket, Clients, Products, Worker, Type WHERE Ticket.client = Clients.id AND Ticket.product = Products.id AND Ticket.worker = Worker.id AND Ticket.type = Type.id AND Ticket.application_data >= CAST('{dateTimePicker1.Value.ToString("d")}' AS DATETIME) AND Ticket.application_data <= CAST('{dateTimePicker2.Value.ToString("d")}' AS DATETIME)";
                string query = $"SELECT Clients.CompanyName AS [Клиент], Products.name AS [Продукт], (Worker.surname + ' ' + Worker.name + ' ' + Worker.patronymic) AS [Принявший сотрудник], Type.name AS [Тип проблемы], Ticket.header AS [Заголовок], Ticket.description AS [Описание], Ticket.priority AS [Приоритет], Ticket.status AS [Статус], Ticket.application_data AS [Дата подачи], Ticket.completion_data AS [Дата закрытия], Ticket.actual_data AS [Фактическая дата закрытия] FROM Ticket, Clients, Products, Worker, Type WHERE Ticket.client = Clients.id AND Ticket.product = Products.id AND Ticket.worker = Worker.id AND Ticket.type = Type.id AND Ticket.application_data >= CAST('{dateTimePicker5.Value.ToString()}' AS DATETIME) AND Ticket.application_data <= CAST('{dateTimePicker6.Value.ToString()}' AS DATETIME) AND Ticket.worker = {comboBox3.SelectedValue} ORDER BY Ticket.application_data ASC";
                using (SqlDataAdapter adapter = new SqlDataAdapter(query, connectionString))
                {
                    DataTable dataTable = new DataTable();
                    adapter.Fill(dataTable);

                    if (dataTable.Rows.Count == 0)
                    {
                        MessageBox.Show("Нет тикетов принятых этим сотрудником!", "Ошибка", MessageBoxButtons.OK, MessageBoxIcon.Error);
                        return;
                    }

                    DateTime fix = DateTime.Now;

                    for (int i = 0, b = 8, x = dataTable.Rows.Count; i < x; i++)
                    {
                        string n = "A" + b.ToString();

                        worksheet.Cell(n).Value = dataTable.Rows[i][0].ToString();
                        worksheet.Cell(n).Style.Alignment.Horizontal = XLAlignmentHorizontalValues.Center;
                        worksheet.Cell(n).Style.Alignment.Vertical = XLAlignmentVerticalValues.Center;
                        worksheet.Cell(n).Style.Font.FontName = "Times New Roman";
                        worksheet.Cell(n).Style.Font.FontSize = 14;
                        worksheet.Cell(n).Style.Alignment.WrapText = true;
                        worksheet.Cell(n).Style.Border.OutsideBorder = XLBorderStyleValues.Thin;

                        n = "B" + b.ToString();

                        worksheet.Cell(n).Value = dataTable.Rows[i][1].ToString();
                        worksheet.Cell(n).Style.Alignment.Horizontal = XLAlignmentHorizontalValues.Center;
                        worksheet.Cell(n).Style.Alignment.Vertical = XLAlignmentVerticalValues.Center;
                        worksheet.Cell(n).Style.Font.FontName = "Times New Roman";
                        worksheet.Cell(n).Style.Font.FontSize = 14;
                        worksheet.Cell(n).Style.Alignment.WrapText = true;
                        worksheet.Cell(n).Style.Border.OutsideBorder = XLBorderStyleValues.Thin;

                        n = "C" + b.ToString();

                        worksheet.Cell(n).Value = dataTable.Rows[i][2].ToString();
                        worksheet.Cell(n).Style.Alignment.Horizontal = XLAlignmentHorizontalValues.Center;
                        worksheet.Cell(n).Style.Alignment.Vertical = XLAlignmentVerticalValues.Center;
                        worksheet.Cell(n).Style.Font.FontName = "Times New Roman";
                        worksheet.Cell(n).Style.Font.FontSize = 14;
                        worksheet.Cell(n).Style.Alignment.WrapText = true;
                        worksheet.Cell(n).Style.Border.OutsideBorder = XLBorderStyleValues.Thin;

                        n = "D" + b.ToString();

                        worksheet.Cell(n).Value = dataTable.Rows[i][3].ToString();
                        worksheet.Cell(n).Style.Alignment.Horizontal = XLAlignmentHorizontalValues.Center;
                        worksheet.Cell(n).Style.Alignment.Vertical = XLAlignmentVerticalValues.Center;
                        worksheet.Cell(n).Style.Font.FontName = "Times New Roman";
                        worksheet.Cell(n).Style.Font.FontSize = 14;
                        worksheet.Cell(n).Style.Alignment.WrapText = true;
                        worksheet.Cell(n).Style.Border.OutsideBorder = XLBorderStyleValues.Thin;

                        n = "E" + b.ToString();

                        worksheet.Cell(n).Value = dataTable.Rows[i][4].ToString();
                        worksheet.Cell(n).Style.Alignment.Horizontal = XLAlignmentHorizontalValues.Center;
                        worksheet.Cell(n).Style.Alignment.Vertical = XLAlignmentVerticalValues.Center;
                        worksheet.Cell(n).Style.Font.FontName = "Times New Roman";
                        worksheet.Cell(n).Style.Font.FontSize = 14;
                        worksheet.Cell(n).Style.Alignment.WrapText = true;
                        worksheet.Cell(n).Style.Border.OutsideBorder = XLBorderStyleValues.Thin;

                        n = "F" + b.ToString();

                        worksheet.Cell(n).Value = dataTable.Rows[i][5].ToString();
                        worksheet.Cell(n).Style.Alignment.Horizontal = XLAlignmentHorizontalValues.Center;
                        worksheet.Cell(n).Style.Alignment.Vertical = XLAlignmentVerticalValues.Center;
                        worksheet.Cell(n).Style.Font.FontName = "Times New Roman";
                        worksheet.Cell(n).Style.Font.FontSize = 14;
                        worksheet.Cell(n).Style.Alignment.WrapText = true;
                        worksheet.Cell(n).Style.Border.OutsideBorder = XLBorderStyleValues.Thin;

                        n = "G" + b.ToString();

                        worksheet.Cell(n).Value = dataTable.Rows[i][6].ToString();
                        worksheet.Cell(n).Style.Alignment.Horizontal = XLAlignmentHorizontalValues.Center;
                        worksheet.Cell(n).Style.Alignment.Vertical = XLAlignmentVerticalValues.Center;
                        worksheet.Cell(n).Style.Font.FontName = "Times New Roman";
                        worksheet.Cell(n).Style.Font.FontSize = 14;
                        worksheet.Cell(n).Style.Alignment.WrapText = true;
                        worksheet.Cell(n).Style.Border.OutsideBorder = XLBorderStyleValues.Thin;

                        n = "H" + b.ToString();

                        //Статус
                        worksheet.Cell(n).Value = dataTable.Rows[i][7].ToString();
                        worksheet.Cell(n).Style.Alignment.Horizontal = XLAlignmentHorizontalValues.Center;
                        worksheet.Cell(n).Style.Alignment.Vertical = XLAlignmentVerticalValues.Center;
                        worksheet.Cell(n).Style.Font.FontName = "Times New Roman";
                        worksheet.Cell(n).Style.Font.FontSize = 14;
                        worksheet.Cell(n).Style.Alignment.WrapText = true;
                        worksheet.Cell(n).Style.Border.OutsideBorder = XLBorderStyleValues.Thin;

                        n = "I" + b.ToString();

                        //Даты
                        fix = (DateTime)dataTable.Rows[i][8];
                        worksheet.Cell(n).Value = fix.ToString("d");
                        worksheet.Cell(n).Style.Alignment.Horizontal = XLAlignmentHorizontalValues.Center;
                        worksheet.Cell(n).Style.Alignment.Vertical = XLAlignmentVerticalValues.Center;
                        worksheet.Cell(n).Style.Font.FontName = "Times New Roman";
                        worksheet.Cell(n).Style.Font.FontSize = 14;
                        worksheet.Cell(n).Style.Alignment.WrapText = true;
                        worksheet.Cell(n).Style.Border.OutsideBorder = XLBorderStyleValues.Thin;

                        n = "J" + b.ToString();

                        fix = (DateTime)dataTable.Rows[i][9];
                        worksheet.Cell(n).Value = fix.ToString("d");
                        worksheet.Cell(n).Style.Alignment.Horizontal = XLAlignmentHorizontalValues.Center;
                        worksheet.Cell(n).Style.Alignment.Vertical = XLAlignmentVerticalValues.Center;
                        worksheet.Cell(n).Style.Font.FontName = "Times New Roman";
                        worksheet.Cell(n).Style.Font.FontSize = 14;
                        worksheet.Cell(n).Style.Alignment.WrapText = true;
                        worksheet.Cell(n).Style.Border.OutsideBorder = XLBorderStyleValues.Thin;

                        n = "K" + b.ToString();

                        //Фактическая
                        if (dataTable.Rows[i][7].ToString() == "Закрыт")
                        {
                            fix = (DateTime)dataTable.Rows[i][10];
                            worksheet.Cell(n).Value = fix.ToString("d");
                        }
                        else
                        {
                            worksheet.Cell(n).Value = "Нет";
                        }
                        worksheet.Cell(n).Style.Alignment.Horizontal = XLAlignmentHorizontalValues.Center;
                        worksheet.Cell(n).Style.Alignment.Vertical = XLAlignmentVerticalValues.Center;
                        worksheet.Cell(n).Style.Font.FontName = "Times New Roman";
                        worksheet.Cell(n).Style.Font.FontSize = 14;
                        worksheet.Cell(n).Style.Alignment.WrapText = true;
                        worksheet.Cell(n).Style.Border.OutsideBorder = XLBorderStyleValues.Thin;

                        b++;
                    }
                }

                //worksheet.Column(1).AdjustToContents();
                //worksheet.Rows(1, 1000).AdjustToContents();

                //СОХРАНЕНИЕ

                //worksheet.Columns().AdjustToContents();

                saveFileDialog1.Filter = "Microsoft Excel Files (*.xlsx)|*.xlsx";
                saveFileDialog1.FileName = filename;
                if (saveFileDialog1.ShowDialog() == DialogResult.OK)
                {
                    try
                    {
                        string filename2 = saveFileDialog1.FileName;
                        workbook.SaveAs(filename2);
                        System.Diagnostics.Process.Start(filename2);
                        MessageBox.Show("Отчет сформирован!", "Успех!", MessageBoxButtons.OK, MessageBoxIcon.Asterisk);
                    }
                    catch
                    {
                        MessageBox.Show("Файл уже существует или произошла другая ошибка!", "Ошибка", MessageBoxButtons.OK, MessageBoxIcon.Error);
                    }
                }
                else
                {
                    MessageBox.Show("Вы отменили сохранение!", "Ошибка", MessageBoxButtons.OK, MessageBoxIcon.Error);
                    return;
                }
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
    }
}
