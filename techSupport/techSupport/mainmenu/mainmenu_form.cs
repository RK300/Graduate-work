using FontAwesome.Sharp;
using RJCodeAdvance.RJControls;
using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Runtime.InteropServices;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Controls;
using System.Windows.Forms;
using System.Windows.Media;
using techSupport.edit_form;
using techSupport.forms;
using techSupport.Ticket_system;
using techSupport.enter_forms;
using Color = System.Drawing.Color;
using Control = System.Windows.Forms.Control;
using Panel = System.Windows.Forms.Panel;
using static System.Windows.Forms.VisualStyles.VisualStyleElement;
using System.Configuration;
using System.Data.SqlClient;
using techSupport.new_forms;
using System.Security.Cryptography;

namespace techSupport.mainmenu
{
    public partial class mainmenu_form : Form
    {
        private IconButton currentBtn;
        private System.Windows.Forms.Panel leftBorderBtn;
        private Form currentChildForm;

        public int userId;

        private void setUserId(int m_id) { userId = m_id; }

        public mainmenu_form(int m_id)
        {
            InitializeComponent();
            setUserId(m_id);
            leftBorderBtn = new Panel();
            leftBorderBtn.Size = new Size(7, 60);
            panelMenu.Controls.Add(leftBorderBtn);
            SetName(m_id);
        }
        
        private void SetName(int id) 
        {
            string query = $"SELECT id, surname, name, patronymic FROM Worker WHERE id = {id}";
            var connectionString = ConfigurationManager.ConnectionStrings["db"].ConnectionString;
            using (SqlDataAdapter adapter = new SqlDataAdapter(query, connectionString))
            {
                DataTable dataTable = new DataTable();
                adapter.Fill(dataTable);
                string Text = dataTable.Rows[0][1].ToString() + ' ' + dataTable.Rows[0][2].ToString() + ' ' + dataTable.Rows[0][3].ToString();
                label5.Text = Text;
            }
        }

        //structs
        private struct RgbColors 
        {
            public readonly static Color color1 = Color.FromArgb(172, 126, 241);
            public readonly static Color color2 = Color.FromArgb(249, 118, 176);
            public readonly static Color color3 = Color.FromArgb(253, 138, 114);
            public readonly static Color color4 = Color.FromArgb(95, 77, 221);
            public readonly static Color color5 = Color.FromArgb(231, 76, 60);
        }


        //methods
        private void OpenChildForm(Form childForm, string Text) 
        {
            if (currentChildForm != null)
            {
                // open only form
                currentChildForm.Close();
            }
            currentChildForm = childForm;
            childForm.TopLevel = false;
            childForm.FormBorderStyle = FormBorderStyle.None;
            childForm.Dock = DockStyle.Fill;
            panelDesktop.Controls.Add(childForm);
            panelDesktop.Tag = childForm;
            childForm.BringToFront();
            childForm.Show();
            label3.Text = Text;
        }

        private void ActivatedButton(object senderBtn, Color color) 
        {
            if (senderBtn != null) 
            {
                DisableButton();

                currentBtn = (IconButton)senderBtn;
                currentBtn.BackColor = Color.FromArgb(37, 36, 81);
                currentBtn.ForeColor = color;
                currentBtn.TextAlign = ContentAlignment.MiddleCenter;
                currentBtn.IconColor = color;
                currentBtn.TextImageRelation = TextImageRelation.TextBeforeImage;
                currentBtn.ImageAlign = ContentAlignment.MiddleRight;
                leftBorderBtn.BackColor = color;
                leftBorderBtn.Location = new Point(0, currentBtn.Location.Y);
                leftBorderBtn.Visible = true;
                leftBorderBtn.BringToFront();
                iconPictureBox3.IconChar = currentBtn.IconChar;
                iconPictureBox3.IconColor = color;
            }
        }

        private void DisableButton() 
        { 
            if (currentBtn != null) 
            {
                currentBtn.BackColor = Color.FromArgb(42, 42, 42);
                currentBtn.ForeColor = Color.Gainsboro;
                currentBtn.TextAlign = ContentAlignment.MiddleLeft;
                currentBtn.IconColor = Color.Gainsboro;
                currentBtn.TextImageRelation = TextImageRelation.ImageBeforeText;
                currentBtn.ImageAlign = ContentAlignment.MiddleLeft;
            }
        }

        private void Open_DropdownMenu(RJDropdownMenu dropdownMenu, object sender) 
        {
            Control controls = (Control)sender;
            dropdownMenu.VisibleChanged += new EventHandler((sender2, ev) => DropdownMenu_VisibleChanged(sender2, ev, controls));
            dropdownMenu.Show(controls, controls.Width, 0);
        }

        private void DropdownMenu_VisibleChanged(object sender, EventArgs e, Control ctrl) 
        { 
            RJDropdownMenu dropdownMenu = (RJDropdownMenu)sender;
            if  ( !DesignMode ) 
            { 
                if (dropdownMenu.Visible) 
                {
                    ctrl.BackColor = Color.FromArgb(48, 48, 48);
                }
                else 
                {
                    ctrl.BackColor = Color.FromArgb(42, 42, 42);
                }
            }
        }

        private void iconButton1_Click(object sender, EventArgs e)
        {
            ActivatedButton(sender, RgbColors.color1);
            OpenChildForm(new client_form(), "Клиенты");
        }

        private void iconButton2_Click(object sender, EventArgs e)
        {
            ActivatedButton(sender, RgbColors.color5);
            Open_DropdownMenu(rjDropdownMenu4, sender);
        }

        private void iconButton3_Click(object sender, EventArgs e)
        {
            ActivatedButton(sender, RgbColors.color3);
            OpenChildForm(new products_edit(), "Продукты");
        }

        private void pictureBox1_Click(object sender, EventArgs e)
        {
            currentChildForm.Close();
            Reset();
        }

        private void Reset()
        {
            DisableButton();
            leftBorderBtn.Visible = true;
            iconPictureBox3.IconChar = IconChar.Home;
            iconPictureBox3.IconColor = Color.White;
            label3.Text = "Главная";
        }

        [DllImport("user32.DLL", EntryPoint = "ReleaseCapture")]
        private extern static void ReleaseCapture();
        [DllImport("user32.DLL", EntryPoint = "SendMessage")]
        private extern static void SendMessage(System.IntPtr hWnd, int wMsg, int wParam, int lParam);

        private void panelTitleBar_MouseDown(object sender, MouseEventArgs e)
        {
            ReleaseCapture();
            SendMessage(this.Handle, 0x112, 0xf012, 0);
        }

        private void Panel1_MouseDown(
            object sender, MouseEventArgs e)
        {
            ReleaseCapture();
            SendMessage(this.Handle, 0x112, 0xf012, 0);
        }

        private void iconPictureBox1_Click(object sender, EventArgs e)
        {
            Application.Exit();
        }

        private void iconPictureBox2_Click(object sender, EventArgs e)
        {
            this.WindowState = FormWindowState.Minimized;
        }

        private void iconButton4_Click(object sender, EventArgs e)
        {
            ActivatedButton(sender, RgbColors.color2);
            Open_DropdownMenu(rjDropdownMenu1, sender);
        }

        private void населенныйПунктToolStripMenuItem_Click(object sender, EventArgs e)
        {
            OpenChildForm(new locality_form(), "Населенные пункты");
        }

        private void районToolStripMenuItem_Click(object sender, EventArgs e)
        {
            OpenChildForm(new district_form(), "Районы");
        }

        private void областьToolStripMenuItem_Click(object sender, EventArgs e)
        {
            OpenChildForm(new region_form(), "Область");
        }

        private void странаToolStripMenuItem_Click(object sender, EventArgs e)
        {
            OpenChildForm(new country_form(), "Страны");
        }

        private void компанииToolStripMenuItem_Click(object sender, EventArgs e)
        {
            OpenChildForm(new company_form(), "Компании");
        }

        private void продуктыToolStripMenuItem_Click(object sender, EventArgs e)
        {
            OpenChildForm(new products_edit(), "Продукты");
        }

        private void сотрудникиToolStripMenuItem_Click(object sender, EventArgs e)
        {
            OpenChildForm(new worker_form(), "Сотрудники");
        }

        private void должностиToolStripMenuItem_Click(object sender, EventArgs e)
        {
            OpenChildForm(new position_form(), "Должности");
        }

        private void iconButton5_Click(object sender, EventArgs e)
        {
            auth auth = new auth();
            auth.Show();
            this.Close();
        }

        private void тикетыToolStripMenuItem_Click(object sender, EventArgs e)
        {
            OpenChildForm(new ticket_form(userId), "Тикеты");
        }

        private void файлыToolStripMenuItem_Click(object sender, EventArgs e)
        {
            OpenChildForm(new file_form(), "Файлы");
        }

        private void типПроблемыToolStripMenuItem_Click(object sender, EventArgs e)
        {
            OpenChildForm(new type_form(), "Тип");
        }

        private void привязатьФайлКТикетуToolStripMenuItem_Click(object sender, EventArgs e)
        {

        }

        private void iconButton6_Click(object sender, EventArgs e)
        {
            ActivatedButton(sender, RgbColors.color2);
            OpenChildForm(new dogovor_form(), "Договора");
        }

        private void toolStripMenuItem1_Click(object sender, EventArgs e)
        {
            OpenChildForm(new client_form(), "Клиенты");
        }

        private void toolStripMenuItem2_Click(object sender, EventArgs e)
        {
            OpenChildForm(new client2prod_form(), "Присвоение продуктов");
        }

        private void iconButton7_Click(object sender, EventArgs e)
        {
            ActivatedButton(sender, RgbColors.color2);
            OpenChildForm(new bank_form(), "Банки");
        }

        private void iconButton8_Click(object sender, EventArgs e)
        {
            ActivatedButton(sender, RgbColors.color3);
            OpenChildForm(new company_form(), "Компании");
        }

        private void iconButton9_Click(object sender, EventArgs e)
        {
            ActivatedButton(sender, RgbColors.color4);
            Open_DropdownMenu(rjDropdownMenu3, sender);
        }
    }
}