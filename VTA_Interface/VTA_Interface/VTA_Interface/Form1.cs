using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;

namespace VTA_Interface
{
    public partial class Form1 : Form
    {

        private Button currentButton;
        private Random random;
        private int temIndex;
        private Form activateform;



        public Form1()
        {
            InitializeComponent();
            random = new Random();
            closechildform.Visible = false;
        }


        private Color selectThemeColor()
        {
            int index = random.Next(Themecolor.colorList.Count);
            while (temIndex == index)
            {
               index = random.Next(Themecolor.colorList.Count);
            }
            temIndex = index;
            string color = Themecolor.colorList[index];
            return ColorTranslator.FromHtml(color);
            }

        private void activatebutton(object btnsender)
        {
            if (btnsender != null)
            {
                if (currentButton != (Button)btnsender)
                {
                    disablebutton();
                    Color color = selectThemeColor();
                    currentButton = (Button)btnsender;
                    currentButton.BackColor = color;
                    currentButton.ForeColor = Color.White;
                    currentButton.Font = new System.Drawing.Font("Microsoft Sans Serif", 10F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
                    closechildform.Visible = true;
                }
            }

        }

        private void disablebutton()
        {
            foreach (Control previousbtn in menu.Controls)
            {
                if (previousbtn.GetType() == typeof(Button))
                {
                    previousbtn.BackColor = Color.FromArgb(64, 64, 64);
                    previousbtn.ForeColor = Color.Gainsboro;
                    previousbtn.Font = new System.Drawing.Font("Microsoft Sans Serif", 9F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
                }
            }
        }

        private void openchildform(Form childForm, object btnSender)
        {
            if (activateform != null)
            {
                activateform.Close();
            }
            activatebutton(btnSender);
            activateform = childForm;
            childForm.TopLevel = false;
            childForm.FormBorderStyle = FormBorderStyle.None;
            childForm.Dock = DockStyle.Fill;
            this.desktop.Controls.Add(childForm);
            this.desktop.Tag = childForm;
            childForm.BringToFront();
            childForm.Show();
            lbtitle.Text = childForm.Text;

            

        }

        private void pay_Click(object sender, EventArgs e)
        {
            openchildform(new Forms.Payments(), sender);
        }

        private void regi_Click(object sender, EventArgs e)
        {
            openchildform(new Forms.Registration(), sender);
        }

        private void mod_Click(object sender, EventArgs e)
        {
            openchildform(new Forms.Modules_Marks(), sender);
        }

        private void ojt_Click(object sender, EventArgs e)
        {
            openchildform(new Forms.OJT_Details(), sender);
        }

        private void exam_Click(object sender, EventArgs e)
        {
            openchildform(new Forms.Exam_Details(), sender);
        }

        private void enq_Click(object sender, EventArgs e)
        {
            openchildform(new Forms.Student_Enqueries(), sender);
        }

        private void Form1_Load(object sender, EventArgs e)
        {

            pictureBox3.Parent = pictureBox2;
            pictureBox3.BackColor = Color.Transparent;
           // label2.Parent = pictureBox2;
           // label2.BackColor = Color.Transparent;



        }

        private void close_Click(object sender, EventArgs e)
        {
            Application.Exit();
        }

        private void button1_Click(object sender, EventArgs e)
        {
            openchildform(new Forms.Drop_Outs(), sender);
        }

        private void button2_Click(object sender, EventArgs e)
        {
            if (activateform != null)
                activateform.Close();
            Reset();

            
        }

        private void Reset()
        {
            disablebutton();
            lbtitle.Text = "HOME";
            currentButton = null;
            closechildform.Visible = false;
            

        }

        private void button2_Click_1(object sender, EventArgs e)
        {
            openchildform(new Forms.Attendance(), sender);
        }

        private void desktop_Paint(object sender, PaintEventArgs e)
        {

        }

        private void accounts_Click(object sender, EventArgs e)
        {
            openchildform(new Forms.Accounts(), sender);
        }

        private void pictureBox1_Click(object sender, EventArgs e)
        {
            this.Close();
            new Loading().Show();
        }

        private void title_Paint(object sender, PaintEventArgs e)
        {

        }

        private void label1_Click(object sender, EventArgs e)
        {

        }

        private void button3_Click(object sender, EventArgs e)
        {
          
        }

        private void label2_Click(object sender, EventArgs e)
        {

        }

        private void timer1_Tick(object sender, EventArgs e)
        {
            string time = DateTime.Now.ToString("dddd MM/dd/yyyy            hh:mm:ss  tt");
            //label2.Text = time;
        }

        private void pictureBox2_Click(object sender, EventArgs e)
        {

        }

        private void pictureBox2_Click_1(object sender, EventArgs e)
        {

        }
    }
}
