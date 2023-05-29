using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;
using MySql.Data.MySqlClient;

namespace VTA_Interface
{
    public partial class Loading : Form
    {
        public Loading()
        {
            InitializeComponent();
            ch.Checked = false;
            con = new MySqlConnection("server = localhost ; user = root ; pwd=; database = kasthury");
            textBox1.Focus();
            textBox2.UseSystemPasswordChar = true;
        }

        MySqlConnection con;


        private void label4_Click(object sender, EventArgs e)
        {
            
        }

        private void Loading_Load(object sender, EventArgs e)
        {
            label4.Parent = pictureBox1;
            label4.BackColor = Color.Transparent;
            label2.Parent = pictureBox1;
            label2.BackColor = Color.Transparent;
            label3.Parent = pictureBox1;
            label3.BackColor = Color.Transparent;
        }

        private void label6_Click(object sender, EventArgs e)
        {
            Application.Exit();
        }

        private void ch_OnChange(object sender, EventArgs e)
        {

        }

        private void bunifuCheckbox1_OnChange(object sender, EventArgs e)
        {
            if (ch.Checked == true)
            {
                textBox2.UseSystemPasswordChar = false;
            }

            if (ch.Checked == false)
            {
                textBox2.UseSystemPasswordChar = true;
            }
        }

        private void pictureBox6_Click(object sender, EventArgs e)
        {
            
           

            try
            {
                con.Open();
                string sql = "select * from acc where user = '"+ textBox1.Text+"' and pass = '"+textBox2.Text+"'";
                MySqlDataAdapter da = new MySqlDataAdapter(sql, con);
                DataTable dt = new DataTable();
                da.Fill(dt);

                if (dt.Rows.Count == 0) 
                {
                    MessageBox.Show("Sorry Invalid Access...!" , "Error" , MessageBoxButtons.OK , MessageBoxIcon.Warning);
                    return ;
                }

                    else 
                    {
                        this.Hide();
                          new Form1().Show();


                    }
                if (textBox1.Text == "admin" && textBox2.Text == "root")
                {
                    this.Hide();
                    new Form1().Show();
                }

                }
               
           
            

            catch (Exception e1)
            {
                MessageBox.Show(e1.Message);
            }

            finally
            {
                con.Close();

            }
        }

        private void textBox1_TextChanged(object sender, EventArgs e)
        {

        }

        private void textBox1_KeyDown(object sender, KeyEventArgs e)
        {
            if (e.KeyCode == Keys.Enter)
            {
                textBox2.Focus();
            }
        }

        private void textBox2_KeyDown(object sender, KeyEventArgs e)
        {
            if (e.KeyCode == Keys.Enter)
            {

                try
                {
                    con.Open();
                    string sql = "select * from acc where user = '" + textBox1.Text + "' and pass = '" + textBox2.Text + "'";
                    MySqlDataAdapter da = new MySqlDataAdapter(sql, con);
                    DataTable dt = new DataTable();
                    da.Fill(dt);

                    if (dt.Rows.Count == 0)
                    {
                        MessageBox.Show("Sorry Invalid Access...!", "Error", MessageBoxButtons.OK, MessageBoxIcon.Warning);
                        return;
                    }

                    else
                    {
                        this.Hide();
                        new Form1().Show();


                    }

                }




                catch (Exception e1)
                {
                    MessageBox.Show(e1.Message);
                }

                finally
                {
                    con.Close();

                }
            }
        }

        private void pictureBox1_Click(object sender, EventArgs e)
        {

        }
    }
}
