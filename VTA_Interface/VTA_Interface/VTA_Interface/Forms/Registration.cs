using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;
using System.IO;
using MySql.Data.MySqlClient;
using System.Drawing.Imaging;


namespace VTA_Interface.Forms
{
    public partial class Registration : Form
    {
        MySqlConnection con;
       
       
        string db = "server=localhost;user=root;pwd=;database=kasthury";
        string b, g, ub, ug, doc, d, imgtext, y, dm, dimis;
        string saipar = "/00";
        void rbtns()
        {
            if (r1st.Checked) { b = r1st.Text; dm = "1"; }
            if (rsecond.Checked) { b = rsecond.Text; dm = "2"; }
            if (rmale.Checked) { g = rmale.Text; }
            if (rfemale.Checked) { g = rfemale.Text; }
            if (checkBox1.Checked) { doc = "Submitted"; }
            if (checkBox1.Checked == false) { doc = "Not Submitted"; }
        }

        public Registration()
        {
            InitializeComponent();

            dbconnect();
            cyear.SelectedItem = "2023";
        }

        public void dbconnect()
        {
            con = new MySqlConnection(db);
        }

        private void radioButton4_CheckedChanged(object sender, EventArgs e)
        {

            if (r1st.Checked == true || rsecond.Checked == true)
            {
                txtstname.Enabled = true;
                txtstname.Focus();
            }
         
            
        }

        private void radioButton3_CheckedChanged(object sender, EventArgs e)
        {
            if (r1st.Checked == true || rsecond.Checked == true)
            {
                txtstname.Enabled = true;
                txtstname.Focus();
            }
        }

        private void radioButton1_CheckedChanged(object sender, EventArgs e)
        {
            checkBox1.Enabled = true;
            button1.Enabled = true;
            button2.Enabled = true;
         
            button4.Enabled = true;
        }

        private void radioButton2_CheckedChanged(object sender, EventArgs e)
        {
            checkBox1.Enabled = true;
        }

        private void btnpicture_Click(object sender, EventArgs e)
        {
            
            OpenFileDialog opf= new OpenFileDialog();
            opf.Filter = "Choose Image(*.jpg;*.png;*.gif;)|*.jpg;*.png;*.gif";
            if (opf.ShowDialog() == DialogResult.OK)
            {
                imgtext = opf.FileName;
                pic.Image = new Bitmap(opf.FileName);
                pic.ImageLocation = opf.FileName;
                pic.Image = Image.FromFile(opf.FileName);
            }
        }

        private void button1_Click(object sender, EventArgs e)
        {
            rbtns();
           
            try
            {
                con.Open();
                string year = cyear.SelectedItem.ToString();
                MemoryStream ms = new MemoryStream();
               

                Image imge = pic.Image;
                byte[] img;
                ImageConverter im = new ImageConverter();
                img = (byte[])im.ConvertTo(imge, typeof(byte[]));
               
                if(checkBox1.Checked) {doc = "Submitted";}
                if(checkBox1.Checked == false) {doc = "Not Submitted";}
             
                string snum = txtmis.Text;
                dnum = "JM/" + y + "/ICTT4/" + dm + "/00" + snum;
                txtmis1.Text = "" + dnum;

                
                string sql1 = "insert into stdetail (mis ,stname , nic , years , batch , gender , dob , address , contact , science , maths , english , stpic , doc) values('" + txtmis1.Text + "','" + txtstname.Text + "','" + txtnic.Text + "','" + year + "','" + b + "','" + g + "' , '" + datepicker.Text + "','" + txtaddress.Text + "' , '" + txtcontact.Text + "' , '" + txtscience.Text + "' , '" + txtmath.Text + "' , '" + txtenglish.Text + "','" + Path.GetFileName(pic.ImageLocation) + "' , '" + doc + "'  );insert into ojt (mis,stname) values('" + txtmis1.Text + "','" + txtstname.Text + "');insert into final (mis,stname,nic) values('" + txtmis1.Text + "','" + txtstname.Text + "','"+txtnic.Text+"');";
                 string sql2 = "insert into attand (mis,stname) values('" + txtmis1.Text + "','" + txtstname.Text + "')";
                 string sql3 = "insert into modules (mis,stname) values('" + txtmis1.Text + "','" + txtstname.Text + "')";
                 string sql4 = "insert into payments (mis,stname) values('" + txtmis1.Text + "','" + txtstname.Text + "')";

                MySqlCommand cmd1 = new MySqlCommand(sql1, con);
                cmd1.ExecuteNonQuery();
                MySqlCommand cmd2 = new MySqlCommand(sql2, con);
                cmd2.ExecuteNonQuery();
                MySqlCommand cmd3 = new MySqlCommand(sql3, con);
                cmd3.ExecuteNonQuery();
                MySqlCommand cmd4 = new MySqlCommand(sql4, con);
                cmd4.ExecuteNonQuery();
                MessageBox.Show("Registration Successful");
                clearall();
                con.Close();
                try
                {
                    File.Copy(imgtext, Application.StartupPath + @"/images/" + Path.GetFileName(pic.ImageLocation));
                }
                catch (Exception ex)
                {
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

            datashow();


        }
        void datashow()
        {
            try
            { 
                con.Open();
                string sql = "select * from stdetail where stpic is not null;";
                MySqlCommand cmd = new MySqlCommand(sql, con);
                MySqlDataAdapter da = new MySqlDataAdapter();
                da.SelectCommand = cmd;
                DataTable dt = new DataTable();
                da.Fill(dt);
                dt.Columns.Add("pictures", Type.GetType("System.Byte[]"));
                dataGridView1.DataSource = dt;
                foreach (DataRow row in dt.Rows)
                {
                    row["pictures"] = File.ReadAllBytes(Application.StartupPath + @"/images/" + Path.GetFileName(row["stpic"].ToString()));
                }
              
            }

            catch (Exception e1)
            {
               
            }

            finally
            {
                con.Close();
            }


        }


        private void Registration_Load(object sender, EventArgs e)
        {
            datashow();
        }

        private void dataGridView1_CellContentClick(object sender, DataGridViewCellEventArgs e)
        {
            try{

           
            string dmis;
         
            txtstname.Text = dataGridView1.Rows[e.RowIndex].Cells[1].Value.ToString();
            txtnic.Text = dataGridView1.Rows[e.RowIndex].Cells[2].Value.ToString();
            cyear.Text = dataGridView1.Rows[e.RowIndex].Cells[3].Value.ToString();
            ub = dataGridView1.Rows[e.RowIndex].Cells[4].Value.ToString();
            ug = dataGridView1.Rows[e.RowIndex].Cells[5].Value.ToString();
            datepicker.Text = dataGridView1.Rows[e.RowIndex].Cells[6].Value.ToString();
            txtaddress.Text = dataGridView1.Rows[e.RowIndex].Cells[7].Value.ToString();
            txtcontact.Text = dataGridView1.Rows[e.RowIndex].Cells[8].Value.ToString();
            txtscience.Text = dataGridView1.Rows[e.RowIndex].Cells[9].Value.ToString();
            txtmath.Text = dataGridView1.Rows[e.RowIndex].Cells[10].Value.ToString();
            txtenglish.Text = dataGridView1.Rows[e.RowIndex].Cells[11].Value.ToString();
          

            MemoryStream ms = new MemoryStream((byte[])dataGridView1.Rows[e.RowIndex].Cells[15].Value);
            pic.Image = Image.FromStream(ms);                                                                        
        
            if (ub == "1st") { r1st.Checked = true; }
            if (ub == "2nd") { rsecond.Checked = true; }
            if (ug == "Male") { rmale.Checked = true; }
            if (ug == "Female") { rfemale.Checked = true; }
            
        }
               catch (Exception ex1)

{
    MessageBox.Show(ex1.Message);
}
   
        }

        string year;
        private void cyear_SelectedIndexChanged(object sender, EventArgs e)
        {
            year = cyear.SelectedItem.ToString();
            y = cyear.Text;
            y = y.Substring(2, y.Length - 2);      
        }
    
 

        private void button3_Click(object sender, EventArgs e)
        {
            try
            {
                con.Open();
                string sql = "delete from stdetail where mis='" + txtmis1.Text + "'; delete from ojt where mis='" + txtmis1.Text + "'; delete from attand where mis='" + txtmis1.Text + "'; delete from final where mis='" + txtmis1.Text + "';delete from modules where mis='" + txtmis1.Text + "'; delete from payments where mis='" + txtmis1.Text + "'; delete from dropout where mis='" + txtmis1.Text + "'; ";      //insi
                MySqlCommand cmd = new MySqlCommand(sql, con);
                cmd.ExecuteNonQuery();
                datashow();
                MessageBox.Show("Successfully Deleted");
            }
            catch (Exception ex)
            {
                MessageBox.Show(ex + "Error");
            }
            finally
            {
                con.Close();
                clearall();
            }
        }
        void clearall()
        {
            rmale.Checked = false;
            rfemale.Checked = false;
            r1st.Checked = false;
            rsecond.Checked = false;
            checkBox1.Checked = false;
            txtmis.Clear();
            txtstname.Clear();
            txtnic.Clear();
            txtaddress.Clear();
            txtcontact.Clear();
            txtscience.Clear();
            txtmath.Clear();
            txtenglish.Clear();
            txtmis1.Clear();
            txtmis.Enabled = true;
            cyear.SelectedItem = "2023";
            txtstname.Enabled = false;
            txtnic.Enabled = false;
            txtmis.Enabled = false;
            txtaddress.Enabled = false;
            txtcontact.Enabled = false;
            txtscience.Enabled = false;
            txtmath.Enabled = false;
            txtenglish.Enabled = false;
            rmale.Enabled = false;
            rfemale.Enabled = false;
            checkBox1.Enabled = false;
            button1.Enabled = false;
            button2.Enabled = false;
            pic.Image = null;
        }

        private void button2_Click(object sender, EventArgs e)
        {
            rbtns();
            

            try
            {
                con.Open();
                string year = cyear.SelectedItem.ToString();
             

                string sql = "update stdetail,attand,ojt,final,modules,payments  set stdetail.stpic='" + Path.GetFileName(pic.ImageLocation) + "',  payments.stname='" + txtstname.Text + "', modules.stname='" + txtstname.Text + "',final.stname='" + txtstname.Text + "', ojt.stname='" + txtstname.Text + "', attand.stname='" + txtstname.Text + "',stdetail.stname  = '" + txtstname.Text + "', stdetail.nic = '" + txtnic.Text + "', stdetail. address= '" + txtaddress.Text + "', stdetail. contact= '" + txtcontact.Text + "' , stdetail.science= '" + txtscience.Text + "' ,stdetail. maths= '" + txtmath.Text + "' ,stdetail. english= '" + txtenglish.Text + "',stdetail.batch = '" + b + "',stdetail.years = '" + year + "',stdetail.gender = '" + g + "',stdetail.dob= '" + datepicker.Text + "',stdetail.doc='" + doc + "'   where stdetail.mis='" + txtmis1.Text + "' and attand.mis='" + txtmis1.Text + "' and ojt.mis='" + txtmis1.Text + "' and final.mis='" + txtmis1.Text + "' and modules.mis='" + txtmis1.Text + "' and payments.mis='" + txtmis1.Text + "';";  //years , batch , gender , dob ,
                MySqlCommand cmd = new MySqlCommand(sql, con);
                cmd.ExecuteNonQuery();
                MessageBox.Show("Update Successful");      
               

                

            }
            catch (Exception ex)
            {
                MessageBox.Show(ex + "Error");
            }
            finally
            {
                clearall();
               
            }
            try
            {

                File.Copy(imgtext, Application.StartupPath + @"/images/" + Path.GetFileName(pic.ImageLocation));
                datashow();
            }

            catch(Exception ex){

            }
            
            datashow();
            con.Close();
        }

        private void txtserreg_TextChanged(object sender, EventArgs e)
        {
            if (r1st.Checked) { b = r1st.Text; }
            if (rsecond.Checked) { b = rsecond.Text; }
            if (rmale.Checked) { g = rmale.Text; }
            if (rfemale.Checked) { g = rfemale.Text; }

            if (cbatch.Text != "")
            {
                try
                {
                    con.Open();
                    string sql = "select mis ,stname , nic , years , batch , gender , dob , address , contact , science , maths , english,doc from stdetail where (stdetail.batch ='" + cbatch.SelectedItem.ToString() + "')  and (mis like '%" + txtserreg.Text + "%'or stname like '%" + txtserreg.Text + "%' or nic like '%" + txtserreg.Text + "%' or contact like '%" + txtserreg.Text + "%' or years like '%" + txtserreg.Text + "%' or gender like '%" + txtserreg.Text + "%' or years like '%" + txtserreg.Text + "%' or address like '%" + txtserreg.Text + "%');";
                  
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
                    string sql = "select mis ,stname , nic , years , batch , gender , dob , address , contact , science , maths , english,doc from stdetail where  mis like '%" + txtserreg.Text + "%'or stname like '%" + txtserreg.Text + "%' or nic like '%" + txtserreg.Text + "%' or contact like '%" + txtserreg.Text + "%' or years like '%" + txtserreg.Text + "%' or gender like '%" + txtserreg.Text + "%' or years like '%" + txtserreg.Text + "%' or address like '%" + txtserreg.Text + "%';";
                    
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

        private void button4_Click(object sender, EventArgs e)
        {
            clearall();

        }

        private void dataGridView1_RowHeaderMouseClick(object sender, DataGridViewCellMouseEventArgs e)
        {
           


            txtmis1.Text = dataGridView1.Rows[e.RowIndex].Cells[0].Value.ToString();
            txtstname.Text = dataGridView1.Rows[e.RowIndex].Cells[1].Value.ToString();
            txtnic.Text = dataGridView1.Rows[e.RowIndex].Cells[2].Value.ToString();
            cyear.Text = dataGridView1.Rows[e.RowIndex].Cells[3].Value.ToString();
            ub = dataGridView1.Rows[e.RowIndex].Cells[4].Value.ToString();
            ug = dataGridView1.Rows[e.RowIndex].Cells[5].Value.ToString();
            datepicker.Text = dataGridView1.Rows[e.RowIndex].Cells[6].Value.ToString();
            txtaddress.Text = dataGridView1.Rows[e.RowIndex].Cells[7].Value.ToString();
            txtcontact.Text = dataGridView1.Rows[e.RowIndex].Cells[8].Value.ToString();
            txtscience.Text = dataGridView1.Rows[e.RowIndex].Cells[9].Value.ToString();
            txtmath.Text = dataGridView1.Rows[e.RowIndex].Cells[10].Value.ToString();
            txtenglish.Text = dataGridView1.Rows[e.RowIndex].Cells[11].Value.ToString();
            d = dataGridView1.Rows[e.RowIndex].Cells[12].Value.ToString();
            try
            {
                MemoryStream ms = new MemoryStream((byte[])dataGridView1.Rows[e.RowIndex].Cells[15].Value);
                pic.Image = Image.FromStream(ms);
            }
            catch(Exception ex){

            }
       

            if (ub == "1st") { r1st.Checked = true; }
            if (ub == "2nd") { rsecond.Checked = true; }
            if (ug == "Male") { rmale.Checked = true; }
            if (ug == "Female") { rfemale.Checked = true; }
            if (d == "Submitted") { checkBox1.Checked = true; }
            if (d == "Not Submitted") { checkBox1.Checked = false; }
        }
        string snum;
        string dnum;
       
        private void txtmis_TextChanged(object sender, EventArgs e)
        {
          
            if (r1st.Checked) { b = r1st.Text; dm = "1"; }
            if (rsecond.Checked) { b = rsecond.Text; dm = "2"; }

            if (txtmis.Text == "")
            {
                txtmis.Text = "0";
            }

         
            string snum = txtmis.Text;
      
 
           dnum = "JM/" + y + "/ICTT4/" + dm + saipar + snum;
           txtmis1.Text = dnum;

           if (txtmis.Text == "0")
           {
               txtmis1.Clear();
           }

            
        }

        private void txtnic_TextChanged(object sender, EventArgs e)
        {
            txtaddress.Enabled = true;
            if (txtnic.Text == "")
            {
                txtaddress.Enabled = false;
            }
        }

        private void txtnic_KeyPress(object sender, KeyPressEventArgs e)
        {
            if (txtnic.Text == " ")
            {
                txtnic.Clear();
            }
            else
            {


                const char Delete = (char)8;
                const char v = (char)86;
                const char shift = (char)16;
         
                e.Handled = !char.IsDigit(e.KeyChar) && e.KeyChar != Delete && e.KeyChar != v && e.KeyChar != shift ;
      
            }
        }

        private void txtcontact_KeyPress(object sender, KeyPressEventArgs e)
        {
            const char Delete = (char)8;
            e.Handled = !char.IsDigit(e.KeyChar) && e.KeyChar != Delete;
        }

        private void txtstname_TextChanged(object sender, EventArgs e)
        {
            if (txtstname.Text == "" || txtstname.Text == " " || txtstname.Text == "  " || txtstname.Text == "   " || txtstname.Text == "    ")
            {
                txtmis.Enabled = false;
                
            }
            else
            {
                txtmis.Enabled = true;
               
                datepicker.Enabled = true;
              


            }
        }

        private void txtmis1_TextChanged(object sender, EventArgs e)
        {
            if (txtmis1.Text == "") 
            {
                txtnic.Enabled = false;

            }

            else
            {
                txtnic.Enabled = true;
            }
        }

        private void txtstname_KeyPress(object sender, KeyPressEventArgs e)
        {
            if (txtstname.Text==" ")
            {
                txtstname.Clear();
                
                
            }
            else
            {
                const char space = (char)8;
                const char space1=(char)32;
                const char dot = (char)46;
               
                e.Handled = !char.IsLetter(e.KeyChar)&&e.KeyChar!=space&&e.KeyChar!=space1&&e.KeyChar!=dot;
              

               
            }
        }

        private void txtscience_KeyPress(object sender, KeyPressEventArgs e)
        {
            if (txtscience.Text == " ")
            {
                txtscience.Clear();
                txtscience.Focus();

            }
     
        }

        private void txtscience_TextChanged(object sender, EventArgs e)
        {
            if (txtscience.Text == "A" || txtscience.Text == "B" || txtscience.Text == "C" || txtscience.Text == "W" || txtscience.Text == "F" || txtscience.Text == "")
            {
                txtmath.Enabled = true;
                if (txtscience.Text == "")
                {
                    txtmath.Enabled = false;
                }
            }
            else
            {
                MessageBox.Show("Invalid Input or Format", "Alert");
                txtscience.Clear();
            }
        }

        private void txtenglish_TextChanged(object sender, EventArgs e)
        {
            rmale.Enabled = true;
            rfemale.Enabled = true;
            if (txtenglish.Text == "A" || txtenglish.Text == "B" || txtenglish.Text == "C" || txtenglish.Text == "W" || txtenglish.Text == "F" || txtenglish.Text == "")
            {
                rmale.Enabled = true;
                rfemale.Enabled = true;
                if (txtenglish.Text == "")
                {
                    rmale.Enabled = false;
                    rfemale.Enabled = false;
                }
            }
            else
            {
                MessageBox.Show("Invalid Input or Format", "Alert");
                txtenglish.Clear();
            }
        }

        private void txtmath_TextChanged(object sender, EventArgs e)
        {
            txtenglish.Enabled = true;
            if (txtmath.Text == "A" || txtmath.Text == "B" || txtmath.Text == "C" || txtmath.Text == "W" || txtmath.Text == "F" || txtmath.Text == "")
            {
                txtenglish.Enabled = true;
                if (txtmath.Text == "")
                {
                    txtenglish.Enabled = false;
                }
            }
            else
            {
                MessageBox.Show("Invalid Input or Format", "Alert");
                txtmath.Clear();
            }
            
        }

        private void txtaddress_TextChanged(object sender, EventArgs e)
        {
            txtcontact.Enabled = true;
            if (txtaddress.Text == "")
            {
                txtcontact.Enabled = false;
            }
        }

        private void txtcontact_TextChanged(object sender, EventArgs e)
        {
            txtscience.Enabled = true;
            if (txtcontact.Text == "")
            {
                txtscience.Enabled = true;
            }

        }

        private void txtmis_KeyPress(object sender, KeyPressEventArgs e)
        {
            const char Delete = (char)8;
            e.Handled = !char.IsDigit(e.KeyChar) && e.KeyChar != Delete;
        }

        private void button5_Click(object sender, EventArgs e)
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

        private void button6_Click(object sender, EventArgs e)
        {
            datashow();
        }
    }
}
