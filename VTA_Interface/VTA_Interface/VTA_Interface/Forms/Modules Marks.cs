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
    public partial class Modules_Marks : Form
    {
        public Modules_Marks()
        {
            InitializeComponent();
             con = new MySqlConnection("server=localhost;user=root;pwd=;database=kasthury");
             datashow();
             mname.SelectedIndex = 0;
        }
        MySqlConnection con;
        string b;
            string prac,theo;
            Boolean x;
           string practical, theory;
           void txtmarksdisable()
           {
               if (lb.Text == "")
               {
                   txttheo.Text = "";
                   txtprac.Text = "";
                   txtname.Text = "";
                   txtmis.Text = "";
                   txttheo.Enabled = false;
                   txtprac.Enabled = false;
               }

           }
        
            void markshow()
            {
                moduleverify();

                MySqlDataReader dr;

                try
                {
                    string sql = "select * from modules where mis='" + txtmis.Text + "' and "+theo+" AND "+prac+" is not null;";
                    MySqlCommand cmd = new MySqlCommand(sql, con);
                    con.Open();
                    dr = cmd.ExecuteReader();
                    while (dr.Read())
                    {
                        // txtmis.Text = dr.GetString("mis");
                        // txtname.Text = dr.GetString("stname");
                       
                        theory = dr.GetString(theo);
                        practical = dr.GetString(prac);
                        txtprac.Text = practical.ToString();
                        txttheo.Text = theory.ToString();

                        
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

            void moduleverify()
            {
                if (mname.SelectedIndex == 0) { prac = "m1prac"; theo = "m1theo"; }
                if (mname.SelectedIndex == 1) { prac = "m2prac"; theo = "m2theo"; }
                if (mname.SelectedIndex == 2) { prac = "m3prac"; theo = "m3theo"; }
                if (mname.SelectedIndex == 3) { prac = "m4prac"; theo = "m4theo"; }
                if (mname.SelectedIndex == 4) { prac = "m5prac"; theo = "m5theo"; }
                if (mname.SelectedIndex == 5) { prac = "m6prac"; theo = "m6theo"; }
                if (mname.SelectedIndex == 6) { prac = "m7prac"; theo = "m7theo"; }
                if (mname.SelectedIndex == 7) { prac = "m8prac"; theo = "m8theo"; }
                if (mname.SelectedIndex == 8) { prac = "m9prac"; theo = "m9theo"; }
                if (mname.SelectedIndex == 9) { prac = "m10prac"; theo = "m10theo"; }
                if (mname.SelectedIndex == 10) { prac = "m11prac"; theo = "m11theo"; }

            }
             
     public  void  lishow(){
       
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
              string sql = "Select  * from stdetail where years = '" + year.SelectedItem.ToString() + "' AND batch = '" + b + "'; ";
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
        
        private void button1_Click(object sender, EventArgs e)
        {
              try
                {
                        con.Open();
                        string sql1 = "update modules set " + prac + "='" + txtprac.Text + "'," + theo + "='" + txttheo.Text + "' where mis='" + txtmis.Text + "';";
                        MySqlCommand cmd = new MySqlCommand(sql1, con);
                        cmd.ExecuteNonQuery();
                        MessageBox.Show("Successfully Entered");
                }
                catch (Exception ex)
                {
                   
                    MessageBox.Show(ex.Message);
                }
                finally
                {
                    con.Close();
                }

              datashow();
              markshow();
              if (txttheo.Text != "" || txtprac.Text != "")
              {
                  txttheo.Enabled = false;
                  txtprac.Enabled = false;
                  BTNEDIT.Enabled = true;
              }
              if (txttheo.Text == "" || txtprac.Text == "")
              {
                  txttheo.Enabled = true;
                  txtprac.Enabled = true;
                  BTNEDIT.Enabled = false;
              }
        }

        private void comboBox1_SelectedIndexChanged(object sender, EventArgs e)
        {
            lmname.Text = mname.Text;
            txttheo.Enabled = true;
            txtprac.Enabled = true;

            
                txttheo.Clear();
                txtprac.Clear();

                markshow();
                if (txttheo.Text != "" || txtprac.Text != "")
                {
                    txttheo.Enabled = false;
                    txtprac.Enabled = false;
                    BTNEDIT.Enabled = true;
                }
                if (txttheo.Text == "" || txtprac.Text == "")
                {
                    txttheo.Enabled = true;
                    txtprac.Enabled = true;
                    BTNEDIT.Enabled = false;
                }
               
        }

        void datashow()
    {
        string sql = "select * from modules;";
        MySqlDataAdapter da = new MySqlDataAdapter(sql , con);
            DataTable dt = new DataTable() ;
            da.Fill(dt);
            dataGridView1.DataSource = dt;
    }


        private void Modules_Marks_Load(object sender, EventArgs e)
        {
            year.SelectedIndex = 0;
            BTNEDIT.Enabled = false;
            btnok.Visible = false;
            btnclose.Visible = false;
        }

        private void r1_CheckedChanged(object sender, EventArgs e)
        {
            lb.Items.Clear();
            lishow();
            txtmarksdisable();
        }

        private void r2_CheckedChanged(object sender, EventArgs e)
        {
            lb.Items.Clear();
            lishow();
            txtmarksdisable();
        }

        private void r1_Click(object sender, EventArgs e)
        {
            lb.Items.Clear();
            lishow();
            txtmarksdisable();
        }

        private void lb_SelectedIndexChanged(object sender, EventArgs e)
        {
            txttheo.Clear();
            txtprac.Clear();
           
                string sql1 = "select * from stdetail where stname='" + lb.Text + "';";
                MySqlCommand cmd = new MySqlCommand(sql1, con);
                MySqlDataReader dr, dr2;
                try
                {
                    con.Open();
                    dr = cmd.ExecuteReader();
                    while (dr.Read())
                    {
                        txtmis.Text = dr.GetString("mis");
                        txtname.Text = dr.GetString("stname");
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

                if (mname.Text != "")
                {
                    markshow();
                }

                if (txttheo.Text != "" || txtprac.Text != "")
                {
                    txttheo.Enabled = false;
                    txtprac.Enabled = false;
                    BTNEDIT.Enabled = true;
                }
                if (txttheo.Text == "" || txtprac.Text == "")
                {
                    txttheo.Enabled = true;
                    txtprac.Enabled = true;
                    BTNEDIT.Enabled = false;
                }
                


        }

        private void txtmis_TextChanged(object sender, EventArgs e)
        {
           
        }

        private void textBox5_Click(object sender, EventArgs e)
        {

        }

        private void txtsearch_DoubleClick(object sender, EventArgs e)
        {
           
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
            if (txtsearch.Text == "" && year.Text != "")
            {
                lishow();
            }
           
                
            
        }

        private void btnserchclose_Click(object sender, EventArgs e)
        {
            
        }

        private void lb_Click(object sender, EventArgs e)
        {
            
        }

        private void lb_MouseClick(object sender, MouseEventArgs e)
        {
            txtmarksdisable();
            if (txttheo.Text != "" || txtprac.Text != "")
            {
                txttheo.Enabled = false;
                txtprac.Enabled = false;
            }
        }

        private void modules_SelectedIndexChanged(object sender, EventArgs e)
        {
            if (r1.Checked == true || r2.Checked == true)
            {
                lb.Items.Clear();
                lishow();
            }
            txtmarksdisable();
        }

        private void groupBox1_Enter(object sender, EventArgs e)
        {
            
        }

        private void button2_Click(object sender, EventArgs e)
        {
            btnok.Visible = true;
            btnclose.Visible = true;
            txttheo.Enabled = true;
            txtprac.Enabled = true;
           
        }

        private void checkBox1_CheckedChanged(object sender, EventArgs e)
        {

        }

        private void dataGridView1_CellContentClick(object sender, DataGridViewCellEventArgs e)
        {

        }

        private void dataGridView1_RowHeaderMouseClick(object sender, DataGridViewCellMouseEventArgs e)
        {


            txtmis.Text = dataGridView1.Rows[e.RowIndex].Cells[0].Value.ToString();
            txtname.Text = dataGridView1.Rows[e.RowIndex].Cells[1].Value.ToString();
            txttheo.Enabled = true;
            txtprac.Enabled = true;
        }

        private void textBox10_TextChanged(object sender, EventArgs e)
        {
            if (cbatch.Text != "")
            {
                try
                {
                    con.Open();
                    string sql = "select * from modules inner join stdetail on modules.mis=stdetail.mis  where (stdetail.batch ='" + cbatch.SelectedItem.ToString() + "')  and   (modules.mis like '%" + textBox10.Text + "%' or modules.stname like '%" + textBox10.Text + "%' or stdetail.years like '%"+textBox10.Text+"%') ;";
                    MySqlDataAdapter da = new MySqlDataAdapter(sql, con);
                    DataTable dt = new DataTable();
                    da.Fill(dt);
                    dataGridView1.DataSource = dt;


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
            else {
                try
                {
                    con.Open();
                    string sql = "select * from modules inner join stdetail on modules.mis=stdetail.mis  where   modules.mis like '%" + textBox10.Text + "%' or modules.stname like '%" + textBox10.Text + "%' or stdetail.years like'%" + textBox10.Text + "%';";
                    MySqlDataAdapter da = new MySqlDataAdapter(sql, con);
                    DataTable dt = new DataTable();
                    da.Fill(dt);
                    dataGridView1.DataSource = dt;


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

        private void txttheo_KeyPress(object sender, KeyPressEventArgs e)
        {
            const char Delete = (char)8;
            const char dot = (char)46;
            e.Handled = !char.IsDigit(e.KeyChar) && e.KeyChar != Delete && e.KeyChar != dot;
        }

        private void txtprac_KeyPress(object sender, KeyPressEventArgs e)
        {
            const char Delete = (char)8;
            const char dot = (char)46;
            e.Handled = !char.IsDigit(e.KeyChar) && e.KeyChar != Delete && e.KeyChar != dot;
        }

        private void button3_Click(object sender, EventArgs e)
        {
            try
            {
                con.Open();
                string sql = "delete from modules where mis = '"+txtmis.Text+"'";
                MySqlCommand cmd = new MySqlCommand(sql, con);
                cmd.ExecuteNonQuery();
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

        private void button4_Click(object sender, EventArgs e)
        {
            clear();
        }

        void clear()
        {
            txtmis.Clear();
            txtname.Clear();
            txtprac.Clear();
            txttheo.Clear();
            textBox10.Clear();
            txttheo.Enabled =false;
            txtprac.Enabled = false;

        }

        private void btnok_Click(object sender, EventArgs e)
        {
            try
            {
                con.Open();
                string sql1 = "update modules set " + prac + "='" + txtprac.Text + "'," + theo + "='" + txttheo.Text + "' where mis='" + txtmis.Text + "';";
                MySqlCommand cmd = new MySqlCommand(sql1, con);
                cmd.ExecuteNonQuery();
                MessageBox.Show("Successfully Updated");
            }
            catch (Exception ex)
            {

                MessageBox.Show(ex.Message);
            }
            finally
            {
                con.Close();
            }
            txtprac.Enabled = false;
            txttheo.Enabled = false;
            btnclose.Visible = false;
            btnok.Visible = false;
            datashow();
            markshow();
        }

        private void btnclose_Click(object sender, EventArgs e)
        {
            txtprac.Enabled = false;
            txttheo.Enabled = false;
            BTNEDIT.Enabled = false;
            btnok.Visible = false;
            btnclose.Visible = false;

            if (txttheo.Text != "" || txtprac.Text != "")
            {
                txttheo.Enabled = false;
                txtprac.Enabled = false;
                BTNEDIT.Enabled = true;
            }
            if (txttheo.Text == "" || txtprac.Text == "")
            {
                txttheo.Enabled = true;
                txtprac.Enabled = true;
                BTNEDIT.Enabled = false;
            }
        }

        private void txttheo_TextChanged(object sender, EventArgs e)
        {

        }

        private void button2_Click_1(object sender, EventArgs e)
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
