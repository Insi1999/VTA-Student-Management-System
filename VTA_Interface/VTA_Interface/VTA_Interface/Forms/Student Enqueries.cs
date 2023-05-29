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
    public partial class Student_Enqueries : Form
    {
        public Student_Enqueries()
        {
            InitializeComponent();
            connect();
            dataGridView1.Visible = false;
        }

        MySqlConnection con;
        string db;
        void connect()
        {

            db = "server=localhost;user=root;pwd=;database=kasthury";
            con = new MySqlConnection(db);

        }

        private void textBox1_TextChanged(object sender, EventArgs e)
        {   
            dataGridView1.Visible = true;
            if (txtsearch.Text == "") { dataGridView1.Visible = false; }

            if (cbatch.Text != "")
            {
                try
                {
                    con.Open();
                    string sql = "select  stdetail.stname as STUDENT_NAME,stdetail.mis as MIS_NO,stdetail.years as YEAR,stdetail.batch as BATCH,attand.totaldays as STUDENT_TOTAL_ATTANDANCE,attand.totalclassdays as TOTAL_CLASS_DAYS,attand.percentage as ATTANDANCE,payments.total as TOTAL_PAY,ojt.ojtplace as OJT_PLACE,final.theory as THEORY,final.practical as PRACTICAL,final.nvq as NVQ_LEVEL,final.indexno as INDEX_NO,stdetail.sts as DROPOUT from stdetail inner join attand on stdetail.stname=attand.stname inner join payments on stdetail.mis=payments.mis inner join ojt on stdetail.mis=ojt.mis inner join final on stdetail.mis=final.mis where (stdetail.batch ='" + cbatch.SelectedItem.ToString() + "' ) and (stdetail.stname like'%" + txtsearch.Text + "%' or stdetail.nic like'%" + txtsearch.Text + "%' or stdetail.mis like'%" + txtsearch.Text + "%'  or stdetail.years like'%" + txtsearch.Text + "%') ";
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
                    string sql = "select stdetail.stname as STUDENT_NAME,stdetail.mis as MIS_NO,stdetail.years as YEAR,stdetail.batch as BATCH,attand.totaldays as STUDENT_TOTAL_ATTANDANCE,attand.totalclassdays as TOTAL_CLASS_DAYS,attand.percentage as ATTANDANCE,payments.total as TOTAL_PAY,ojt.ojtplace as OJT_PLACE,final.theory as THEORY,final.practical as PRACTICAL,final.nvq as NVQ_LEVEL,final.indexno as INDEX_NO,stdetail.sts as DROPOUT from stdetail inner join attand on stdetail.mis=attand.mis inner join payments on stdetail.mis=payments.mis inner join ojt on stdetail.mis=ojt.mis inner join final on stdetail.mis=final.mis where  (stdetail.stname like'%" + txtsearch.Text + "%' or stdetail.nic like'%" + txtsearch.Text + "%' or stdetail.mis like'%" + txtsearch.Text + "%'  or stdetail.years like'%" + txtsearch.Text + "%') ";
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

        private void copytocb()
        {
            try
            {
                  
                dataGridView1.SelectAll();
                DataObject obj = dataGridView1.GetClipboardContent();

                if (obj != null)
                    Clipboard.SetDataObject(obj);
                else
                    MessageBox.Show("data is empty");

            }


            catch (Exception ex)
            {

                MessageBox.Show(ex.Message);
            }


        }

        private void button1_Click(object sender, EventArgs e)
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

        private void comboBox1_SelectedIndexChanged(object sender, EventArgs e)
        {

        }
    }
    }

