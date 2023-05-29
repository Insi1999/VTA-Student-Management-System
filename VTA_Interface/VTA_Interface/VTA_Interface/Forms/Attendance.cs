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
    public partial class Attendance : Form
    {
        public Attendance()
        {
            InitializeComponent();
            connect();
            cyear.SelectedItem = "2023";
            cmbm.SelectedItem = "January";
            datashowa();
            cyear.SelectedItem = "2023";
            txtpercent.Enabled = false;
        }

        MySqlConnection con;
        string db, b, amonth,clmonth;
        double x,y,monthtot,classtot;


        void connect()
        {

            db = "server=localhost;user=root;pwd=;database=kasthury";
            con = new MySqlConnection(db);

        }

        void monthverify()
        {
            if (cmbm.SelectedIndex == 0) { clmonth = "cjanuary"; amonth = "january"; }
            if (cmbm.SelectedIndex == 1) { clmonth = "cFebruary"; amonth = "february"; }
            if (cmbm.SelectedIndex == 2) { clmonth = "cmarch"; amonth = "march"; }
            if (cmbm.SelectedIndex == 3) { clmonth = "capril"; amonth = "april"; }
            if (cmbm.SelectedIndex == 4) { clmonth = "cmay"; amonth = "may"; }
            if (cmbm.SelectedIndex == 5) { clmonth = "cjune"; amonth = "june"; }
            if (cmbm.SelectedIndex == 6) { clmonth = "cjuly"; amonth = "july"; }
            if (cmbm.SelectedIndex == 7) { clmonth = "caugust"; amonth = "august"; }
            if (cmbm.SelectedIndex == 8) { clmonth = "cseptember"; amonth = "september"; }
            if (cmbm.SelectedIndex == 9) { clmonth = "coctober"; amonth = "october"; }
            if (cmbm.SelectedIndex == 10) { clmonth = "cnovember"; amonth = "november"; }
            if (cmbm.SelectedIndex == 11) { clmonth = "cdecember"; amonth = "december"; }



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
                string sql = "Select  * from stdetail where years = '" + cyear.SelectedItem.ToString() + "' AND batch = '" + b + "'; ";
                MySqlCommand cmd = new MySqlCommand(sql, con);
                MySqlDataReader dr = cmd.ExecuteReader();

                while (dr.Read())
                {
                    lb.Items.Add(dr.GetString("stname"));
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


        private void r1_CheckedChanged(object sender, EventArgs e)
        {
                lb.Items.Clear();
                lishow();
        }
        double classdays = 0,per;
        void gettotal()
        {
            try
            {
                string sql1 = "select  * from attand where mis ='" + txtmis.Text + "' and totaldays and totalclassdays is not null  ";
                MySqlCommand cmd2 = new MySqlCommand(sql1, con);
                MySqlDataReader dr2 = cmd2.ExecuteReader();

                while (dr2.Read())
                {
                    
                   
                        y = dr2.GetDouble("totaldays");
                        classdays = dr2.GetDouble("totalclassdays");
                        per = dr2.GetDouble("percentage");
                        tdays.Text = "" + y;
                        txttotclass.Text = "" + classdays;
                        txtpercent.Text = "" + per;

                }

                dr2.Close();
            }
            catch (Exception ex)
            {
                MessageBox.Show(ex.Message);
            }
            finally
            {

            }
        }
        void showdays()
        {

            try
            {
                con.Open();

                string sql = "Select mis , stname from stdetail where stname = '" + lb.Text + "';";
                MySqlCommand cmd = new MySqlCommand(sql, con);
                MySqlDataReader dr = cmd.ExecuteReader();

                while (dr.Read())
                {
                    txtmis.Text = dr.GetString("mis");
                    txtstname.Text = dr.GetString("stname");
                }

                dr.Close();
               
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
        private void lb_SelectedIndexChanged(object sender, EventArgs e)
        {
            tdays.Text = "";
            txttotclass.Text = "";
            txtpercent.Text = "";
            txtmt.Text = "";

            try
            {
                con.Open();
                
                string sql = "Select mis , stname from stdetail where stname = '" + lb.Text + "';";
                MySqlCommand cmd = new MySqlCommand(sql, con);
                MySqlDataReader dr = cmd.ExecuteReader();

                while (dr.Read())
                {
                    txtmis.Text = dr.GetString("mis");
                    txtstname.Text = dr.GetString("stname");
                }

                dr.Close();
                gettotal();
                con.Close();
                getmonthandclasstot();
            }

            catch (Exception ex)
            {
                MessageBox.Show(ex.Message);
            }
            finally
            {
                
            }
           
        }

        private void cyear_SelectedIndexChanged(object sender, EventArgs e)
        {
            lb.Items.Clear();
            lishow();
        }

        private void r2_CheckedChanged(object sender, EventArgs e)
        {
                lb.Items.Clear();
                lishow();
        }
        void datashowa()
        {
            try
            {
                con.Open();
                string sql = "select * from attand;";
                MySqlDataAdapter da = new MySqlDataAdapter(sql, con);
                DataTable dt = new DataTable();
                da.Fill(dt);
                dataGridView1.DataSource = dt;

            }
            catch(Exception ex)
            {
                MessageBox.Show(ex.Message);
            }
            con.Close();
        }

        private void Attendance_Load(object sender, EventArgs e)
        {
            btnok.Visible = false; btnedit.Enabled = false;
            btnclosesearch.Visible = false;


        }

        double percentage;
        private void button1_Click(object sender, EventArgs e)
        {
           
            try
            {
               
                con.Open();
                monthverify();
                gettotal();
                
               
                x = y + double.Parse(txtmt.Text);
                classdays = classdays + double.Parse(txttc.Text);
                percentage = x / classdays * 100;
                string sql1 = "update attand set " + amonth + "='" + double.Parse( txtmt.Text) + "'," + clmonth + "='" + double.Parse(txttc.Text )+ "',totaldays='"+x+"',totalclassdays='"+classdays+"',percentage='"+percentage+"' where mis ='"+ txtmis.Text + "';";
                MySqlCommand cmd = new MySqlCommand(sql1, con);
                cmd.ExecuteNonQuery();
                MessageBox.Show("Successfully Entered");
               
                txtmt.Enabled = false;
                txttc.Enabled = false;
                gettotal();
            }
            catch (Exception ex)
            {

                MessageBox.Show(ex.Message);
            }
            finally
            {
                con.Close();
            }
            getmonthandclasstot();
            datashowa();
            
        }

        private void dataGridView1_CellContentClick(object sender, DataGridViewCellEventArgs e)
        {

        }
        void getmonthandclasstot()
        {
            try
            {
                con.Open();

                if (txtmt.Text == "" || txttc.Text == "")
                {
                    txtmt.Enabled = true;
                    txttc.Enabled = true;
                    btnedit.Enabled = false;
                }

                monthverify();
   
                string sql = "select * from attand where mis='" + txtmis.Text + "' and " + amonth + " and " + clmonth + " is not null ;";
                MySqlCommand cmd = new MySqlCommand(sql, con);

                MySqlDataReader dr = cmd.ExecuteReader();

                while (dr.Read())
                {
                    monthtot = dr.GetDouble(amonth);
                    classtot = dr.GetDouble(clmonth);
                    txtmt.Text = "" + monthtot;
                    txttc.Text = "" + classtot;


                    if (txtmt.Text != "" || txttc.Text != "")
                    {
                        txtmt.Enabled = false;
                        txttc.Enabled = false;
                        btnedit.Enabled = true;
                    }

                }
                dr.Close();
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

        private void cmbm_SelectedIndexChanged(object sender, EventArgs e)
        {
            txtmt.Clear();
            txttc.Clear();
            getmonthandclasstot();
        }

        private void btnedit_Click(object sender, EventArgs e)
        {
            txtmt.Enabled = true;
            txttc.Enabled = true;
            btnok.Visible = true;
            btnclosesearch.Visible = true;
        }
        double d,g,j,m;
        private void btnok_Click(object sender, EventArgs e)
        {
            monthverify();
            con.Open();
            
            gettotal();
            y = y - monthtot;
            classdays = classdays - classtot;
            string sql1 = "update attand set totaldays='" + y + "',totalclassdays='"+classdays+"',percentage=0 where mis ='" + txtmis.Text + "' ;";
            MySqlCommand cmd = new MySqlCommand(sql1, con);
            cmd.ExecuteNonQuery();
            gettotal();
            con.Close();
            monthverify();
            con.Open();
            g = y + double.Parse(txtmt.Text);
            m = classdays + double.Parse(txttc.Text);
            percentage = g / m * 100;
            string sql2 = "update attand set " + amonth + "='" + double.Parse(txtmt.Text) + "'," + clmonth + "='" + double.Parse(txttc.Text) + "',totaldays='" + g + "',totalclassdays='" + m + "',percentage='"+percentage+"' where mis ='" + txtmis.Text + "' ;";
            MySqlCommand cmd2 = new MySqlCommand(sql2, con);
            cmd2.ExecuteNonQuery();
            MessageBox.Show("Successfully Updated");
            con.Close();

            txtmt.Enabled = false;
            txttc.Enabled = false;
            btnok.Visible = false;
            btnclosesearch.Visible = false;
            con.Open();
            gettotal();
            con.Close();
            
            
            getmonthandclasstot();
            datashowa();
        }

        private void txtmt_KeyPress(object sender, KeyPressEventArgs e)
        {
            const char Delete = (char)8;
            const char dot = (char)46;
            e.Handled = !char.IsDigit(e.KeyChar) && e.KeyChar != Delete&&e.KeyChar!=dot;

        }

        private void txttc_KeyPress(object sender, KeyPressEventArgs e)
        {
            const char Delete = (char)8;
            const char dot = (char)46;
            e.Handled = !char.IsDigit(e.KeyChar) && e.KeyChar != Delete && e.KeyChar != dot;
        }

        private void txttc_TextChanged(object sender, EventArgs e)
        {

        }

        private void button2_Click(object sender, EventArgs e)
        {
            txtmt.Enabled = false;
            txttc.Enabled = false;
            btnok.Visible = false;
            btnclosesearch.Visible = false;
        }

        private void txtsearch_TextChanged(object sender, EventArgs e)
        {
            if (cbatch.Text != "")
            {
                try
                {
                    con.Open();
                    string sql = "select * from attand inner join stdetail on attand.mis=stdetail.mis  where (stdetail.batch ='" + cbatch.SelectedItem.ToString() + "')  and (attand.mis like '%" + txtsearch.Text + "%' or attand.stname like '%" + txtsearch.Text + "%' or stdetail.years like '%" + txtsearch.Text + "%') ;";
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
            else
            {

                try
                {
                    con.Open();
                    string sql = "select * from attand inner join stdetail on attand.mis=stdetail.mis  where  (attand.mis like '%" + txtsearch.Text + "%' or attand.stname like '%" + txtsearch.Text + "%'  or stdetail.years like '%" + txtsearch.Text + "%') ;";
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

        private void txtsearch2_TextChanged(object sender, EventArgs e)
        {
            lb.Items.Clear();

            try
            {

                con.Open();
                string sql = "select stname from stdetail where stname like '%" + txtsearch2.Text + "%'";
                MySqlCommand cmd = new MySqlCommand(sql, con);

                MySqlDataReader dr = cmd.ExecuteReader();
                while (dr.Read())
                {
                    lb.Items.Add(dr.GetString("stname"));



                }
                if (txtsearch2.Text == "")
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
            if (txtsearch2.Text == "" && cyear.Text != "")
            {
                lishow();
            }
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
    }
}
