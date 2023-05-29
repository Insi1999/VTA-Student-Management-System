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

namespace VTA_Interface.Forms
{
    public partial class Accounts : Form
    {
        public Accounts()
        {
            InitializeComponent();
            con = new MySqlConnection("server=localhost ; user=root ; pwd=; database=kasthury");
            datashow();
        }

        MySqlConnection con;

        void datashow()
    {
        try
        {
            string sql = "select * from acc";
            MySqlDataAdapter da = new MySqlDataAdapter(sql, con);
            DataTable dt = new DataTable();
            da.Fill(dt);
            dataGridView1.DataSource = dt;
        }
        catch (Exception e1)
        {
            MessageBox.Show(e1.Message);
        }

    }

        void clear ()
        {
            uname.Clear();
            pass.Clear();
            name.Clear();
            des.Clear(); 
        }

        private void Accounts_Load(object sender, EventArgs e)
        {

        }

        private void button1_Click(object sender, EventArgs e)
        {
            if (uname.Text == "" || pass.Text == "" || name.Text == "" || des.Text == "")
            {
                MessageBox.Show("Please Fill Full Details", "Error", MessageBoxButtons.OK, MessageBoxIcon.Error);
            }

            else
            {
                try
                {
                    con.Open();
                    string sql = "insert into acc (user,pass,nm,des) values ('" + uname.Text + "' , '" + pass.Text + "' , '" + name.Text + "' , '" + des.Text + "')";
                    MySqlCommand cmd = new MySqlCommand(sql, con);
                    cmd.ExecuteNonQuery();
                    datashow();
                    MessageBox.Show("Successfully Added");
                    clear();
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

        private void dataGridView1_RowHeaderMouseClick(object sender, DataGridViewCellMouseEventArgs e)
        {
            uname.Text = dataGridView1.Rows[e.RowIndex].Cells[1].Value.ToString();
            pass.Text = dataGridView1.Rows[e.RowIndex].Cells[2].Value.ToString();
            name.Text = dataGridView1.Rows[e.RowIndex].Cells[3].Value.ToString();
            des.Text = dataGridView1.Rows[e.RowIndex].Cells[4].Value.ToString();

        }

        private void button3_Click(object sender, EventArgs e)
        {
            try
            {
                con.Open();
                string sql = "delete from acc where user = '"+uname.Text+"'";
                MySqlCommand cmd = new MySqlCommand(sql, con);
                cmd.ExecuteNonQuery();
                datashow();
                clear();
                MessageBox.Show("Successfully Deleted");

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

        private void button4_Click(object sender, EventArgs e)
        {
            clear();

        }

        private void button2_Click(object sender, EventArgs e)
        {
            try
            {
                con.Open();
                string sql = "update acc set nm = '" + name.Text + "' , pass = '" + pass.Text + "' , des = '" + des.Text + "'  where user = '"+uname.Text+"'";
                MySqlCommand cmd = new MySqlCommand(sql, con);
                cmd.ExecuteNonQuery();
                datashow();
            }

            catch (Exception e1)
            {
                MessageBox.Show(e1.Message);
            }

            finally
            {
                con.Close();
                clear();
            }
            
        }

        private void dataGridView1_CellContentClick(object sender, DataGridViewCellEventArgs e)
        {

        }




    }
}
