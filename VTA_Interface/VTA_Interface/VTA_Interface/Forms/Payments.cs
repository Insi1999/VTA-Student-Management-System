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
using System.IO;

namespace VTA_Interface.Forms
{
    public partial class Payments : Form
    {
        Registration reg = new Registration();
        MySqlConnection con;
        string b;
        double tot;




        public Payments()
        {
            InitializeComponent();
            dbconnection();
            datashowp();
            cyear.SelectedItem = "2023";

        }
      void  dbconnection()
        {
              con = new MySqlConnection("server=localhost;user=root;pwd=;database=kasthury");
        }
        void lishow()
        {
            try
            {
                if (r1.Checked)
                {
                    b = "1st";
                }

                if (r2.Checked)
                {
                    b = "2nd";
                }

                con.Open();
                string sql = "Select  * from stdetail where years = '"+cyear.SelectedItem.ToString()+"' AND batch = '"+b+"'; ";
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
            finally {

                con.Close();
            }


        }
       void datashowp()
        {
            try
            {
               con.Open();
                string sql="select * from payments";
                MySqlDataAdapter da = new MySqlDataAdapter(sql, con);

                DataTable dt = new DataTable();
                da.Fill(dt);
                dataGridView1.DataSource = dt;

            }
            catch(Exception ex)
            {
                MessageBox.Show(ex + "Error");
            }
            finally
            {
               con.Close();
            }
        }
        private void Payments_Load(object sender, EventArgs e)
        {
           
        }

        private void button4_Click(object sender, EventArgs e)
        {
            lb.Items.Clear();
            lishow();
        }

        private void lbnames_SelectedIndexChanged(object sender, EventArgs e)
        {
            string sql1 = "select * from stdetail where stname='" + lb.Text + "';";
            MySqlCommand cmd = new MySqlCommand(sql1, con);
            MySqlDataReader dr;
            try
            {
                con.Open();


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
            finally {

                con.Close();
            }
        }

        private void comboBox1_SelectedIndexChanged(object sender, EventArgs e)
        {
            if (r1.Checked == true || r2.Checked == true)
            {
                lb.Items.Clear();
                lishow();
            }
        }


        private void button1_Click(object sender, EventArgs e)
        {
            try
            {
                con.Open();

                string sql = "update payments set ddfee='" + double.Parse(txtddf.Text) + "',cbtfee='" + double.Parse(txtcbtf.Text) + "',examfee='" + double.Parse(txtef.Text) + "',regfee='" + double.Parse(rf.Text) + "',total='" + tot + "',des='" + txtdes.Text + "' where mis='"+txtmis.Text+"';";
                MySqlCommand cmd = new MySqlCommand(sql, con);
                cmd.ExecuteNonQuery();
                MessageBox.Show("Saved Successfully");
               
                
            }
            catch (Exception ex)
            {
                MessageBox.Show(ex.Message);
            }
            finally
            {
               
                con.Close();

            }
            datashowp();
            clearall();
        }

        private void txtddf_TextChanged(object sender, EventArgs e)
        {

            try
            {

                if (txtddf.Text == "")
                {
                    txtddf.Text = "0";
                    txtddf.SelectionLength = txtddf.Text.Length;
                }
                tot = double.Parse(txtcbtf.Text) + double.Parse(txtddf.Text) + double.Parse(txtef.Text) + double.Parse(rf.Text);
                txttf.Text = "" + tot;
            }
            catch (Exception ex)
            {
                MessageBox.Show(ex.Message);
            }


          
            
        }

        private void txtcbtf_TextChanged(object sender, EventArgs e)
        {
            if (txtcbtf.Text == "")
            {
                txtcbtf.Text = "0";
                txtcbtf.SelectionLength = txtcbtf.Text.Length;
            }
            tot = double.Parse(txtcbtf.Text) + double.Parse(txtddf.Text) + double.Parse(txtef.Text) + double.Parse(rf.Text);
            txttf.Text = "" + tot;
        }

        private void txtef_TextChanged(object sender, EventArgs e)
        {
            if (txtef.Text == "")
            {
                txtef.Text = "0";
                txtef.SelectionLength = txtef.Text.Length;
            }
            tot = double.Parse(txtcbtf.Text) + double.Parse(txtddf.Text) + double.Parse(txtef.Text) + double.Parse(rf.Text);
            txttf.Text = "" + tot;
        }

        private void txttf_TextChanged(object sender, EventArgs e)
        {
            
            if (txttf.Text == "")
            {
                txttf.Text = "0";
            }
            tot = double.Parse(txtcbtf.Text) + double.Parse(txtddf.Text) + double.Parse(txtef.Text) + double.Parse(rf.Text);
            txttf.Text = "" + tot;
            txtdes.Enabled = true;
            if (txttf.Text == "0")
            {
                txtdes.Enabled = false;
                txtdes.Clear();
            }
        }
      void  clearall()
        {

            txtcbtf.Text="0";
            txtddf.Text = "0";
            txtef.Text = "0";
            tot = 0;
            txttf.Text = "";
            txtmis.Text = "";
            txtname.Clear();
            txtdes.Clear();
            datashowp();
            txtserpay.Clear();
            rf.Text = "0";
        }
        private void button5_Click(object sender, EventArgs e)
        {

            r1.Checked = false;
            r2.Checked = false;
            txtcbtf.Text = "0";
            txtddf.Text = "0";
            txtef.Text = "0";
            tot = 0;
            txttf.Text = "";
            txtmis.Text = "";
            txtname.Clear();
            lb.Items.Clear();
            txtdes.Clear();
            datashowp();
            txtserpay.Clear();
            rf.Clear();

        }

        private void button2_Click(object sender, EventArgs e)
        {

            try
            {
                con.Open();
                tot = double.Parse(txtcbtf.Text) + double.Parse(txtddf.Text) + double.Parse(txtef.Text);
                 txttf.Text = "" + tot;
                 string sql = "update payments set ddfee='" + double.Parse(txtddf.Text) + "',cbtfee='" + double.Parse(txtcbtf.Text) + "',regfee='" + double.Parse(rf.Text) + "',examfee='" + double.Parse(txtef.Text) + "',total='" + tot + "',des='" + txtdes.Text + "' where mis='" + txtmis.Text + "';";
                MySqlCommand cmd = new MySqlCommand(sql, con);
                cmd.ExecuteNonQuery();
                MessageBox.Show("Updated Successfully");
                
                
            }
            catch (Exception ex)
            {
                MessageBox.Show(ex + "Error");
            }
            finally
            {
               
                con.Close();

            }
            clearall();
            datashowp();


        }
        string pid;
        private void dataGridView1_RowHeaderMouseClick(object sender, DataGridViewCellMouseEventArgs e)
        {
            txtmis.Text = dataGridView1.Rows[e.RowIndex].Cells[1].Value.ToString();
            txtname.Text = dataGridView1.Rows[e.RowIndex].Cells[2].Value.ToString();
            txtddf.Text = dataGridView1.Rows[e.RowIndex].Cells[3].Value.ToString();
            txtcbtf.Text = dataGridView1.Rows[e.RowIndex].Cells[4].Value.ToString();
            txtef.Text = dataGridView1.Rows[e.RowIndex].Cells[5].Value.ToString();
            rf.Text = dataGridView1.Rows[e.RowIndex].Cells[6].Value.ToString();
            txttf.Text = dataGridView1.Rows[e.RowIndex].Cells[7].Value.ToString();
            txtdes.Text = dataGridView1.Rows[e.RowIndex].Cells[8].Value.ToString();
        }

        private void button3_Click(object sender, EventArgs e)
        {
            try
            {
                con.Open();
                string sql = "delete from payments where mis='" +txtmis.Text + "';";
                MySqlCommand cmd = new MySqlCommand(sql, con);
                cmd.ExecuteNonQuery();
                MessageBox.Show("Successfully Deleted the Record");
                clearall();
            }
            catch (Exception ex)
            {
                MessageBox.Show(ex + "Error");
            }
            finally
            {
                datashowp();
                con.Close();

            }
        }

        private void txtserpay_TextChanged(object sender, EventArgs e)
        {
            if (cbatch.Text != "")
            {
                try
                {
                    con.Open();
                    string sql = "select * from payments  inner join stdetail on payments.mis=stdetail.mis  where (stdetail.batch ='" + cbatch.SelectedItem.ToString() + "')  and   ( payments.mis like '%" + txtserpay.Text + "%'or payments.stname like '%" + txtserpay.Text + "%' or payments.pid like '%" + txtserpay.Text + "%' or stdetail.years like '%" + txtserpay.Text + "%') ;";
                    MySqlCommand cmd = new MySqlCommand(sql, con);
                    MySqlDataAdapter da = new MySqlDataAdapter(sql, con);
                    DataTable dt = new DataTable();
                    da.Fill(dt);
                    dataGridView1.DataSource = dt;
                    cmd.ExecuteNonQuery();

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
            else
            {

                try
                {
                    con.Open();
                    string sql = "select * from payments  inner join stdetail on payments.mis=stdetail.mis  where   ( payments.mis like '%" + txtserpay.Text + "%'or payments.stname like '%" + txtserpay.Text + "%' or payments.pid like '%" + txtserpay.Text + "%' or stdetail.years like '%" + txtserpay.Text + "%') ;";
                    MySqlCommand cmd = new MySqlCommand(sql, con);
                    MySqlDataAdapter da = new MySqlDataAdapter(sql, con);
                    DataTable dt = new DataTable();
                    da.Fill(dt);
                    dataGridView1.DataSource = dt;
                    cmd.ExecuteNonQuery();

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

        private void textBox1_TextChanged(object sender, EventArgs e)
        {
            try
            {
                if (rf.Text == "")
                {
                    rf.Text = "0";
                    rf.SelectionLength = rf.Text.Length;
                }
                tot = double.Parse(txtcbtf.Text) + double.Parse(txtddf.Text) + double.Parse(txtef.Text) + double.Parse(rf.Text);
                txttf.Text = "" + tot;
            }

            catch (Exception ex)
            {
                MessageBox.Show(ex.Message);
                
            }
        }

        private void dataGridView1_CellContentClick(object sender, DataGridViewCellEventArgs e)
        {

        }

        private void txtddf_MouseClick(object sender, MouseEventArgs e)
        {
            txtddf.SelectionLength = txtddf.Text.Length;
        }

        private void txtddf_KeyPress(object sender, KeyPressEventArgs e)
        {
            const char Delete = (char)8;
            const char dot = (char)46;
            e.Handled = !char.IsDigit(e.KeyChar) && e.KeyChar != Delete && e.KeyChar != dot;
        }

        private void txtcbtf_KeyPress(object sender, KeyPressEventArgs e)
        {
            const char Delete = (char)8;
            const char dot = (char)46;
            e.Handled = !char.IsDigit(e.KeyChar) && e.KeyChar != Delete && e.KeyChar != dot;
        }

        private void rf_KeyPress(object sender, KeyPressEventArgs e)
        {
            const char Delete = (char)8;
            const char dot = (char)46;
            e.Handled = !char.IsDigit(e.KeyChar) && e.KeyChar != Delete && e.KeyChar != dot;
        }

        private void txtef_KeyPress(object sender, KeyPressEventArgs e)
        {
            const char Delete = (char)8;
            const char dot = (char)46;
            e.Handled = !char.IsDigit(e.KeyChar) && e.KeyChar != Delete && e.KeyChar != dot;
        }

        private void txtddf_Enter(object sender, EventArgs e)
        {
           
        }

        private void txtddf_Leave(object sender, EventArgs e)
        {

        }

        private void txtddf_Click(object sender, EventArgs e)
        {
            txtddf.SelectionLength = txtddf.TextLength;
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

        private void cbatch_SelectedIndexChanged(object sender, EventArgs e)
        {

        }

        private void button3_Click_1(object sender, EventArgs e)
        {
            try
            {
                Microsoft.Office.Interop.Excel.Application xcelapp = new Microsoft.Office.Interop.Excel.Application();
                xcelapp.Application.Workbooks.Add(Type.Missing); 
                xcelapp.Visible = true;
                xcelapp.Columns.AutoFit();
  
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
