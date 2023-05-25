using Newtonsoft.Json;
using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.IO;
using System.Linq;
using System.Text;
using SettingsClass;
using System.Threading.Tasks;
using System.Windows.Forms;

namespace techSupport.new_forms
{
    public partial class settings_edit : Form
    {
        public settings_edit()
        {
            InitializeComponent();
            LoadBox();
        }

        private void LoadBox() 
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

            textBox7.Text = settings.FIO;
            textBox1.Text = settings.NazvComp;
            textBox2.Text = settings.RascSchet;
            textBox3.Text = settings.Bank;
            textBox5.Text = settings.YNP;
            textBox4.Text = settings.OKPO;
            textBox6.Text = settings.Adress;

            File.WriteAllText("settings.json", JsonConvert.SerializeObject(settings));
        }

        private void button2_Click(object sender, EventArgs e)
        {
            var settings = File.Exists("settings.json") ? JsonConvert.DeserializeObject<Settings>(File.ReadAllText("settings.json")) : null;

            settings.FIO = textBox7.Text;
            settings.NazvComp = textBox1.Text;
            settings.RascSchet = textBox2.Text;
            settings.Bank = textBox3.Text;
            settings.YNP = textBox5.Text;
            settings.OKPO = textBox4.Text;
            settings.Adress = textBox6.Text;

            File.WriteAllText("settings.json", JsonConvert.SerializeObject(settings));
            this.DialogResult = DialogResult.OK;
        }
    }
}
