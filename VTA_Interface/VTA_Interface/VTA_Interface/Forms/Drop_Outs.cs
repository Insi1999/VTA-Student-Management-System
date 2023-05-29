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
    public partial class Drop_Outs : Form
    {
        public Drop_Outs()
        {
            InitializeComponent();
            con = new MySqlConnection("server=localhost;user=root;pwd=;database=kasthury");
            cb.SelectedIndex = 0;
            cyear.SelectedIndex = 0;
            datashow();
            cb.SelectedIndex = 0;
        }

        MySqlConnection con ;

          
        string b;

        void lishow()
        {

            if (r1.Checked)
            {
                b = "1st";
            }

            if (r2.Checked)
            {
                b = "2nd";
            }

            try
            {

                con.Open();
                string sql = "Select  * from stdetail where years = '" + cyear.SelectedItem.ToString() + "' AND batch = '" + b + "'; ";
                MySqlCommand cmd = new MySqlCommand(sql, con);
                DataTable dt = new DataTable();
                MySqlDataReader dr = cmd.ExecuteReader();


                while (dr.Read())
                {

                    lb.Items.Add(dr.GetString("stname"));
                }
            }

            catch (Exception ex)
            {
                MessageBox.Show(ex.Message);
            }
            finally
            {
                con.Close();
            }
        }


        void clear()
        {
            txtmis.Text = "";
            txtname.Text = "";
            txtad.Text = "";
            textBox5.Text = "";
            datashow();
  
        }

        private void Drop_Outs_Load(object sender, EventArgs e)
        {
            
            
        }


        void datashow()
        {
            try
            {
                string sql = "select * from dropout";
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

        private void txtdes_TextChanged(object sender, EventArgs e)
        {

        }

        private void txtdes_KeyPress(object sender, KeyPressEventArgs e)
        {
            const char Delete = (char)8;
            e.Handled = !char.IsDigit(e.KeyChar) && e.KeyChar!= Delete;
        }

        private void lb_SelectedIndexChanged(object sender, EventArgs e)
        {
            try
            {
            con.Open();
            string sql1 = "select * from stdetail where stname='" + lb.Text + "';";
            MySqlCommand cmd = new MySqlCommand(sql1, con);
            MySqlDataReader dr;
           
                dr = cmd.ExecuteReader();

                if (dr.Read())
                {
                    string mis = dr.GetString("mis"); ;
                    string name = dr.GetString("stname");

                    txtname.Text = name;
                    txtmis.Text = mis;
                }

            }
            catch (Exception ex)
            {
                MessageBox.Show(ex.Message);
            }
            finally
            {

                con.Close();
            }
            
        }

        private void r1_CheckedChanged(object sender, EventArgs e)
        {
            lb.Items.Clear();
            lishow();
        }

        private void r2_CheckedChanged(object sender, EventArgs e)
        {
            lb.Items.Clear();
            lishow();
        }

        private void cyear_SelectedIndexChanged(object sender, EventArgs e)
        {
            lb.Items.Clear();
            lishow();
        }

        private void button1_Click(object sender, EventArgs e)
        {
          
            try
            {

                con.Open();
                string sql = "insert into dropout (mis , stname ,batch,year, lastatt , des) values ('" + txtmis.Text + "' , '" + txtname.Text + "' ,'" + b + "','" + cyear.SelectedItem.ToString() + "', '" + txtad.Text + "' , '" + cb.SelectedItem.ToString() + "');update stdetail set sts='Dropout' where mis='" + txtmis.Text + "'";
                MySqlCommand cmd = new MySqlCommand(sql, con);
                cmd.ExecuteNonQuery();
                MessageBox.Show("Successfully Added");
               datashow();

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

        private void dataGridView1_RowHeaderMouseClick(object sender, DataGridViewCellMouseEventArgs e)
        {
            try
            {
                txtmis.Text = dataGridView1.Rows[e.RowIndex].Cells[0].Value.ToString();
                txtname.Text = dataGridView1.Rows[e.RowIndex].Cells[1].Value.ToString();
                txtad.Text = dataGridView1.Rows[e.RowIndex].Cells[4].Value.ToString();
                cb.Text = dataGridView1.Rows[e.RowIndex].Cells[5].Value.ToString();
            }

            catch (Exception e1)
            {
                MessageBox.Show(e1.Message);
            }

            finally
            {

            }
           

        }

        private void button3_Click(object sender, EventArgs e)
        {
            try
            {
                con.Open();
                string sql = "delete from dropout where mis = '" + txtmis.Text + "'";
                MySqlCommand cmd = new MySqlCommand(sql, con);
                cmd.ExecuteNonQuery();
                MessageBox.Show("Successfully Deleted...!");
               datashow();
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

        private void button2_Click(object sender, EventArgs e)
        {
            try
            {
                con.Open();
                string sql = "update dropout set lastatt = '"+txtad.Text+"' ,des = '"+cb.SelectedItem.ToString()+"' where mis = '"+txtmis.Text+"' ";
                MySqlCommand cmd = new MySqlCommand(sql, con);
                cmd.ExecuteNonQuery();
                MessageBox.Show("Successfully Updated");
                datashow();
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

        private void button5_Click(object sender, EventArgs e)
        {
            clear();
        }

        private void textBox5_TextChanged(object sender, EventArgs e)
        {
            if (cbatch.Text != "")
            {
                try
                {
                    con.Open();
                    string sql = "select * from dropout  inner join stdetail on dropout.mis=stdetail.mis  where (stdetail.batch ='" + cbatch.SelectedItem.ToString() + "')  and (dropout.mis like  '%" + textBox5.Text + "%' or dropout.stname like  '%" + textBox5.Text + "%' or  stdetail.years like  '%" + textBox5.Text + "%' or dropout.des like  '%" + textBox5.Text + "%'or stdetail.years like '%" + txtsearch.Text + "%') ";
                    MySqlDataAdapter da = new MySqlDataAdapter(sql, con);
                    DataTable dt = new DataTable();
                    da.Fill(dt);
                    dataGridView1.DataSource = dt;
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

            else
            {

                try
                {
                    con.Open();
                    string sql = "select * from dropout  inner join stdetail on dropout.mis=stdetail.mis  where  (dropout.mis like  '%" + textBox5.Text + "%' or dropout.stname like  '%" + textBox5.Text + "%' or  stdetail.years like  '%" + textBox5.Text + "%' or dropout.des like  '%" + textBox5.Text + "%' or stdetail.years like '%" + txtsearch.Text + "%') ";
                    MySqlDataAdapter da = new MySqlDataAdapter(sql, con);
                    DataTable dt = new DataTable();
                    da.Fill(dt);
                    dataGridView1.DataSource = dt;
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

        private void txtsearch_TextChanged(object sender, EventArgs e)
        {
            lb.Items.Clear();

            try
            {

                con.Open();
                string sql = "select stname from stdetail where stname like '%" + txtsearch.Text + "%'";
                MySqlCommand cmd = new MySqlCommand(sql, con);

                MySqlDataReader dr = cmd.ExecuteReader();
                while (dr.Read())
                {
                    lb.Items.Add(dr.GetString("stname"));
                }
                if (txtsearch.Text == "")
                {
                    lb.Items.Clear();
                }

            }
            catch (Exception ex)
            {
                MessageBox.Show(ex.Message);
            }
            finally
            {
                con.Close();
            }
            if (txtsearch.Text == "" && cyear.Text != "")
            {
                lishow();
            }
        }

        private void button3_Click_1(object sender, EventArgs e)
        {
            try
            {
                Microsoft.Office.Interop.Excel.Application xcelapp = new Microsoft.Office.Interop.Excel.Application();
                xcelapp.Application.Workbooks.Add(Type.Missing);
                xcelapp.Visible = true;
                xcelapp.Columns.AutoFit();

                // Storing header part in Excel   
                for (int i = 1; i < dataGridView1.Columns.Count + 1; i++)
                {
                    xcelapp.Cells[1, i] = dataGridView1.Columns[i - 1].HeaderText;
                }

                // Storing Each row and column value to excel sheet   
                for (int i = 0; i < dataGridView1.Rows.Count; i++)
                {
                    for (int j = 0; j < dataGridView1.Columns.Count; j++)
                    {
                        if (j == 2 || j == 5)
                        {
                            xcelapp.Cells[i + 2, j + 1] = "'" + dataGridView1.Rows[i].Cells[j].Value.ToString();
                        }
                        else
                        {
                            xcelapp.Cells[i + 2, j + 1] = dataGridView1.Rows[i].Cells[j].Value.ToString();
                        }
                    }
                }

            }

            catch (Exception ex)
            {

            }
        }
    }
}
