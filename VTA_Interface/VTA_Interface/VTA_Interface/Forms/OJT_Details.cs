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
    public partial class OJT_Details : Form
    {
        public OJT_Details()
        {
            InitializeComponent();
            du.SelectedIndex = 0;
            cyear.SelectedItem = "2023";
            con = new MySqlConnection("server=localhost;user=root;pwd=;database=kasthury");
            datashow();
        }

        MySqlConnection con;
        string salissue, comtype;
        string b;
        void datashow()

        {

            try
            {
                string sql = "select * from ojt";
                MySqlDataAdapter da = new MySqlDataAdapter(sql , con);
                DataTable dt = new DataTable();
                da.Fill(dt);
                dataGridView1.DataSource = dt;
            }

            catch (Exception e1)
            {
                MessageBox.Show(e1.Message);
            }

        }


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
        private void radioButton6_CheckedChanged(object sender, EventArgs e)
        {
            lb.Items.Clear();
            lishow();
        }

        private void listBox1_SelectedIndexChanged(object sender, EventArgs e)
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
            finally
            {

                con.Close();
            }
        }

        private void OJT_Details_Load(object sender, EventArgs e)
        {
            
        }

        private void comboBox1_SelectedIndexChanged(object sender, EventArgs e)
        {

            if (r1.Checked == true || r2.Checked == true)
            {
                lb.Items.Clear();
                lishow();
            }
        }

        private void radioButton7_CheckedChanged(object sender, EventArgs e)
        {
            lb.Items.Clear();
            lishow();
        }

        private void button1_Click(object sender, EventArgs e)
        {
            if (txtojt.Text == "" || txtdepcon.Text == "" || txtcorname.Text == "" || txtcorcontact.Text == "")
            {
                MessageBox.Show("Please Enter Valuss Before Save", "Alert");
            }
            else
            {
                if (ryes.Checked) { salissue = "Yes"; }
                if (rno.Checked) { salissue = "NO"; }
                if (rgov.Checked) { comtype = "Government"; }
                if (rpvt.Checked) { comtype = "Private "; }
                if (rsgov.Checked) { comtype = "Semi Government "; }

                try
                {
                    con.Open();
                    string sql = "update ojt set duration='" + du.SelectedItem.ToString() + "', ojtplace='" + txtojt.Text + "',depcontact='" + txtdepcon.Text + "',corname='"+txtcorname.Text+"',corcontact='" + txtcorcontact.Text + "',salissue='" + salissue + "',companytype='" + comtype + "' where mis='"+txtmis.Text+"'";
                    MySqlCommand cmd = new MySqlCommand(sql, con);
                    cmd.ExecuteNonQuery();
                    MessageBox.Show("Successfully Added");
                }

                catch (Exception ex)
                {
                    MessageBox.Show(ex.Message);
                }
                finally { con.Close(); }
                datashow();
            }
        }


        private void button2_Click(object sender, EventArgs e)
        {

            if (ryes.Checked) { salissue = "Yes"; }
            if (rno.Checked) { salissue = "NO"; }
            if (rgov.Checked) { comtype = "Government"; }
            if (rpvt.Checked) { comtype = "Private "; }
            if (rsgov.Checked) { comtype = "Semi Government "; }

            try
            {
                con.Open();
                string sql = "update ojt set duration='" + du.SelectedItem.ToString() +"', ojtplace='" + txtojt.Text + "',depcontact='" + txtdepcon.Text + "',corname='" + txtcorname.Text + "',corcontact='" + txtcorcontact.Text + "',salissue='" + salissue + "',companytype='" + comtype + "' where mis='" + txtmis.Text + "'";
                MySqlCommand cmd = new MySqlCommand(sql, con);
                cmd.ExecuteNonQuery();
                MessageBox.Show("Successfully Updated");
            }

            catch (Exception ex)
            {
                MessageBox.Show(ex.Message);
            }
            finally { con.Close(); }
            datashow(); 
        }

        private void textBox9_TextChanged(object sender, EventArgs e)
        {
            if (cbatch.Text != "")
            {
                try
                {
                    con.Open();
                    string sql = "select * from ojt   inner join stdetail on ojt.mis=stdetail.mis  where (stdetail.batch ='" + cbatch.SelectedItem.ToString() + "')  and (ojt.mis like '%" + textBox9.Text + "%'or ojt.stname like '%" + textBox9.Text + "%' or stdetail.years like '%" + textBox9.Text + "%')";
                    MySqlCommand cmd = new MySqlCommand(sql, con);
                    MySqlDataAdapter da = new MySqlDataAdapter(sql, con);
                    DataTable dt = new DataTable();
                    da.Fill(dt);
                    dataGridView1.DataSource = dt;
                    cmd.ExecuteNonQuery();

                }

                catch (Exception e1)
                {

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
                    string sql = "select * from ojt   inner join stdetail on ojt.mis=stdetail.mis  where  (ojt.mis like '%" + textBox9.Text + "%'or ojt.stname like '%" + textBox9.Text + "%' or stdetail.years like '%" + textBox9.Text + "%')";
                    MySqlCommand cmd = new MySqlCommand(sql, con);
                    MySqlDataAdapter da = new MySqlDataAdapter(sql, con);
                    DataTable dt = new DataTable();
                    da.Fill(dt);
                    dataGridView1.DataSource = dt;
                    cmd.ExecuteNonQuery();

                }

                catch (Exception e1)
                {

                }

                finally
                {
                    con.Close();
                }



            }
        }

        private void txtdepcon_KeyPress(object sender, KeyPressEventArgs e)
        {
            const char Delete = (char)8;
            e.Handled = !char.IsDigit(e.KeyChar) && e.KeyChar != Delete;
        }

        private void txtcorcontact_KeyPress(object sender, KeyPressEventArgs e)
        {
            const char Delete = (char)8;
            e.Handled = !char.IsDigit(e.KeyChar) && e.KeyChar != Delete;
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

        private void dataGridView1_RowHeaderMouseClick(object sender, DataGridViewCellMouseEventArgs e)
        {
            txtmis.Text = dataGridView1.Rows[e.RowIndex].Cells[0].Value.ToString();
            txtname.Text = dataGridView1.Rows[e.RowIndex].Cells[1].Value.ToString();
            du.Text = dataGridView1.Rows[e.RowIndex].Cells[2].Value.ToString();
            txtojt.Text = dataGridView1.Rows[e.RowIndex].Cells[3].Value.ToString();
            txtdepcon.Text = dataGridView1.Rows[e.RowIndex].Cells[4].Value.ToString();
            txtcorname.Text = dataGridView1.Rows[e.RowIndex].Cells[5].Value.ToString();
            txtcorcontact.Text= dataGridView1.Rows[e.RowIndex].Cells[6].Value.ToString();
            if (dataGridView1.Rows[e.RowIndex].Cells[7].Value.ToString() == "Yes")
            {
                ryes.Checked = true;

            }
            else { rno.Checked = true; }

            if (dataGridView1.Rows[e.RowIndex].Cells[8].Value.ToString() == "Government")
            {
                rgov.Checked = true;

            }
            else if (dataGridView1.Rows[e.RowIndex].Cells[8].Value.ToString() == "Private ")
            {
                rpvt.Checked = true;

            }
            else
            {

                rsgov.Checked = true;

            }
            

        }

        private void txtsearch_TextChanged_1(object sender, EventArgs e)
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

        private void button3_Click(object sender, EventArgs e)
        {
            con.Open();
            string sql = "update ojt set duration='',ojtplace='',depcontact='',corname='',corcontact='',salissue='',companytype='' where mis='"+txtmis.Text+"' ";
            MySqlCommand cmd = new MySqlCommand(sql,con);
            cmd.ExecuteNonQuery();
            con.Close();
            datashow();
        }

        private void du_SelectedIndexChanged(object sender, EventArgs e)
        {

        }

        private void txtojt_KeyPress(object sender, KeyPressEventArgs e)
        {
            
            
                const char Delete = (char)8;
                const char space = (char)32;
                e.Handled = !char.IsLetter(e.KeyChar) && e.KeyChar != Delete && e.KeyChar != space;
        }

        private void txtcorname_TextChanged(object sender, EventArgs e)
        {
           
        }

        private void txtcorname_KeyPress(object sender, KeyPressEventArgs e)
        {
                const char Delete = (char)8;
                const char space = (char)32;
                e.Handled = !char.IsLetter(e.KeyChar) && e.KeyChar != Delete&&e.KeyChar!=space;
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

