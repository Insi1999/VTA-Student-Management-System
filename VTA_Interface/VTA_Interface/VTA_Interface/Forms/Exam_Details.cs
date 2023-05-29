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
    public partial class Exam_Details : Form
    {
        public Exam_Details()
        {
            InitializeComponent();
            connect();
            datashow();
            cyear.SelectedItem = "2023";
        }
        string b;

        MySqlConnection con;
        string db = "server=localhost;user=root;pwd=;database=kasthury";

        void clear()
        {
            txtmis.Text = "";
            txtindex.Text = "";
            txtname.Text = "";
            txtnic.Text = "";
            r1.Checked = false;
            r2.Checked = false;
            rprint.Checked = false;
            rissue.Checked = false;
            rn2.Checked = false;
            rn3.Checked = false;
            rn4.Checked = false;
            rtfail.Checked = false;
            rtpass.Checked = false;
            rpfail.Checked = false;
            rppass.Checked = false;
        }

        void connect()
        {
            con = new MySqlConnection(db);
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



        void datashow()
        {
            con.Open();
            try
            {
                string sql = "select * from final";
                MySqlDataAdapter da = new MySqlDataAdapter(sql, con);
                DataTable dt = new DataTable();
                da.Fill(dt);
                dataGridView1.DataSource = dt;
            }

            catch (Exception e1)
            {

                MessageBox.Show(e1.Message);

            }

            con.Close();
        }

        private void dataGridView1_CellContentClick(object sender, DataGridViewCellEventArgs e)
        {

        }

        private void Exam_Details_Load(object sender, EventArgs e)
        {

        }

        private void lb_SelectedIndexChanged(object sender, EventArgs e)
        {
            string sql1 = "select * from stdetail where stname='" + lb.SelectedItem.ToString()+ "';";
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
                    txtnic.Text = dr.GetString("nic");

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

        private void cyear_SelectedIndexChanged(object sender, EventArgs e)
        {
            lb.Items.Clear();
            lishow();
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

        string certsts , theory , practical , nvq;
        

        void rbtnsvalid()
        {

            if (rprint.Checked) { certsts = rprint.Text; }
            if (rissue.Checked) { certsts = rissue.Text; }

            if (rtpass.Checked) { theory = rtpass.Text;}
            if (rtfail.Checked) { theory = rtfail.Text; }

            if (rppass.Checked) { practical = rppass.Text; }
            if (rpfail.Checked) { practical = rpfail.Text; }

            if (rn2.Checked) { nvq = rn2.Text;}
            if (rn3.Checked) { nvq = rn3.Text; }
            if (rn4.Checked) { nvq = rn4.Text; }
           
        }

        private void button1_Click(object sender, EventArgs e)
        {
            try
            {
                con.Open();
                rbtnsvalid();

                string sql = "update final set indexno = '" + txtindex.Text + "',currentdetail = '" + cbs.SelectedItem.ToString() + "',certificate = '" + certsts + "',theory= '" + theory + "',practical= '" + practical + "',nvq='" + nvq + "' where mis = '" + txtmis.Text + "';";
                MySqlCommand cmd = new MySqlCommand(sql, con);
                cmd.ExecuteNonQuery();
            }

            catch (Exception e1)
            {
                MessageBox.Show(e1.Message);
            }

            finally
            {
                con.Close();
            }
            clear();
            datashow();
        }

        private void textBox9_TextChanged(object sender, EventArgs e)
        {
            if (cbatch.Text != "")
            {
                try
                {
                    con.Open();
                    string sql = "select * from final inner join stdetail on final.mis=stdetail.mis  where (stdetail.batch ='" + cbatch.SelectedItem.ToString() + "')  and  (final.mis like '%" + textBox9.Text + "%' or final.stname like '%" + textBox9.Text + "%' or final.indexno like '%" + textBox9.Text + "%' or final.currentdetail like '%" + textBox9.Text + "%' or stdetail.nic like '%" + textBox9.Text + "%'or final.theory like '%" + textBox9.Text + "%' or final.practical like '%" + textBox9.Text + "%' or final.nvq like '%" + textBox9.Text + "%' or stdetail.years like '%" + textBox9.Text + "%');";
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
                    string sql = "select * from final inner join stdetail on final.mis=stdetail.mis  where   (final.mis like '%" + textBox9.Text + "%' or final.stname like '%" + textBox9.Text + "%' or final.indexno like '%" + textBox9.Text + "%' or final.currentdetail like '%" + textBox9.Text + "%' or stdetail.nic like '%" + textBox9.Text + "%'or final.theory like '%" + textBox9.Text + "%' or final.practical like '%" + textBox9.Text + "%' or final.nvq like '%" + textBox9.Text + "%' or stdetail.years like '%" + textBox9.Text + "%');";
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

        private void dataGridView1_RowHeaderMouseClick(object sender, DataGridViewCellMouseEventArgs e)
        {
            string cp, th, pr, nvq;

            txtmis.Text = dataGridView1.Rows[e.RowIndex].Cells[0].Value.ToString();
            txtnic.Text = dataGridView1.Rows[e.RowIndex].Cells[1].Value.ToString();
            txtname.Text = dataGridView1.Rows[e.RowIndex].Cells[2].Value.ToString();
            txtindex.Text = dataGridView1.Rows[e.RowIndex].Cells[3].Value.ToString();
            cbs.Text = dataGridView1.Rows[e.RowIndex].Cells[4].Value.ToString();
            cp = dataGridView1.Rows[e.RowIndex].Cells[5].Value.ToString();
            th = dataGridView1.Rows[e.RowIndex].Cells[6].Value.ToString();
            pr = dataGridView1.Rows[e.RowIndex].Cells[7].Value.ToString();
            nvq= dataGridView1.Rows[e.RowIndex].Cells[8].Value.ToString();

            if (cp == rprint.Text) { rprint.Checked = true; }
            if (cp == rissue.Text) { rissue.Checked = true; }
            if (th == rtpass.Text) { rtpass.Checked = true; }
            if (th == rtfail.Text) { rtfail.Checked = true; }
            if (pr == rppass.Text) { rppass.Checked = true; }
            if (pr == rpfail.Text) { rpfail.Checked = true; }
            if (nvq == rn2.Text)   { rn2.Checked    = true; }
            if (nvq == rn3.Text)   { rn3.Checked    = true; }
            if (nvq == rn4.Text)   { rn4.Checked    = true; }  
        }

        private void button3_Click(object sender, EventArgs e)
        {
            try
            {
                con.Open();

                string sql = "delete from final where mis = '"+txtmis.Text+"'";
                MySqlCommand cmd = new MySqlCommand(sql, con);
                cmd.ExecuteNonQuery();
                MessageBox.Show("Successfully deleted");
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
            clear();
        }

        private void button2_Click(object sender, EventArgs e)
        {
            try
            {
                con.Open();
                rbtnsvalid();
                string sql = "update final set indexno = '"+txtindex.Text+"',currentdetail = '"+cbs.SelectedItem.ToString()+"' ,certificate= '"+certsts+"' , theory = '"+theory+"' , practical = '"+practical+"' , nvq = '"+nvq+"' where mis = '"+txtmis.Text+"' ";
                MySqlCommand cmd = new MySqlCommand(sql, con);
                cmd.ExecuteNonQuery();
                
                MessageBox.Show("Successfully Updated..");
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
            datashow();
        }

        private void button4_Click(object sender, EventArgs e)
        {
            clear();
            cbs.Text = "";
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
