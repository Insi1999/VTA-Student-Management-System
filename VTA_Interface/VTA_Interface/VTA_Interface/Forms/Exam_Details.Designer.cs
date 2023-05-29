namespace VTA_Interface.Forms
{
    partial class Exam_Details
    {
        /// <summary>
        /// Required designer variable.
        /// </summary>
        private System.ComponentModel.IContainer components = null;

        /// <summary>
        /// Clean up any resources being used.
        /// </summary>
        /// <param name="disposing">true if managed resources should be disposed; otherwise, false.</param>
        protected override void Dispose(bool disposing)
        {
            if (disposing && (components != null))
            {
                components.Dispose();
            }
            base.Dispose(disposing);
        }

        #region Windows Form Designer generated code

        /// <summary>
        /// Required method for Designer support - do not modify
        /// the contents of this method with the code editor.
        /// </summary>
        private void InitializeComponent()
        {
            this.button2 = new System.Windows.Forms.Button();
            this.button1 = new System.Windows.Forms.Button();
            this.textBox9 = new System.Windows.Forms.TextBox();
            this.label13 = new System.Windows.Forms.Label();
            this.groupBox2 = new System.Windows.Forms.GroupBox();
            this.rtfail = new System.Windows.Forms.RadioButton();
            this.rtpass = new System.Windows.Forms.RadioButton();
            this.dataGridView1 = new System.Windows.Forms.DataGridView();
            this.label9 = new System.Windows.Forms.Label();
            this.label7 = new System.Windows.Forms.Label();
            this.label3 = new System.Windows.Forms.Label();
            this.label2 = new System.Windows.Forms.Label();
            this.txtindex = new System.Windows.Forms.TextBox();
            this.txtnic = new System.Windows.Forms.TextBox();
            this.txtname = new System.Windows.Forms.TextBox();
            this.txtmis = new System.Windows.Forms.TextBox();
            this.groupBox3 = new System.Windows.Forms.GroupBox();
            this.rpfail = new System.Windows.Forms.RadioButton();
            this.rppass = new System.Windows.Forms.RadioButton();
            this.groupBox1 = new System.Windows.Forms.GroupBox();
            this.rissue = new System.Windows.Forms.RadioButton();
            this.rprint = new System.Windows.Forms.RadioButton();
            this.groupBox4 = new System.Windows.Forms.GroupBox();
            this.rn4 = new System.Windows.Forms.RadioButton();
            this.rn3 = new System.Windows.Forms.RadioButton();
            this.rn2 = new System.Windows.Forms.RadioButton();
            this.lb = new System.Windows.Forms.ListBox();
            this.groupBox5 = new System.Windows.Forms.GroupBox();
            this.r2 = new System.Windows.Forms.RadioButton();
            this.r1 = new System.Windows.Forms.RadioButton();
            this.cyear = new System.Windows.Forms.ComboBox();
            this.label10 = new System.Windows.Forms.Label();
            this.cbs = new System.Windows.Forms.ComboBox();
            this.label1 = new System.Windows.Forms.Label();
            this.button4 = new System.Windows.Forms.Button();
            this.txtsearch = new System.Windows.Forms.TextBox();
            this.cbatch = new System.Windows.Forms.ComboBox();
            this.button3 = new System.Windows.Forms.Button();
            this.label11 = new System.Windows.Forms.Label();
            this.label4 = new System.Windows.Forms.Label();
            this.groupBox2.SuspendLayout();
            ((System.ComponentModel.ISupportInitialize)(this.dataGridView1)).BeginInit();
            this.groupBox3.SuspendLayout();
            this.groupBox1.SuspendLayout();
            this.groupBox4.SuspendLayout();
            this.groupBox5.SuspendLayout();
            this.SuspendLayout();
            // 
            // button2
            // 
            this.button2.Location = new System.Drawing.Point(293, 320);
            this.button2.Name = "button2";
            this.button2.Size = new System.Drawing.Size(56, 29);
            this.button2.TabIndex = 54;
            this.button2.Text = "Update";
            this.button2.UseVisualStyleBackColor = true;
            this.button2.Click += new System.EventHandler(this.button2_Click);
            // 
            // button1
            // 
            this.button1.Location = new System.Drawing.Point(212, 320);
            this.button1.Name = "button1";
            this.button1.Size = new System.Drawing.Size(68, 29);
            this.button1.TabIndex = 53;
            this.button1.Text = "Add";
            this.button1.UseVisualStyleBackColor = true;
            this.button1.Click += new System.EventHandler(this.button1_Click);
            // 
            // textBox9
            // 
            this.textBox9.Location = new System.Drawing.Point(698, 331);
            this.textBox9.Name = "textBox9";
            this.textBox9.Size = new System.Drawing.Size(154, 20);
            this.textBox9.TabIndex = 52;
            this.textBox9.TextChanged += new System.EventHandler(this.textBox9_TextChanged);
            // 
            // label13
            // 
            this.label13.AutoSize = true;
            this.label13.Location = new System.Drawing.Point(651, 334);
            this.label13.Name = "label13";
            this.label13.Size = new System.Drawing.Size(41, 13);
            this.label13.TabIndex = 51;
            this.label13.Text = "Search";
            // 
            // groupBox2
            // 
            this.groupBox2.Controls.Add(this.rtfail);
            this.groupBox2.Controls.Add(this.rtpass);
            this.groupBox2.Location = new System.Drawing.Point(581, 179);
            this.groupBox2.Name = "groupBox2";
            this.groupBox2.Size = new System.Drawing.Size(109, 107);
            this.groupBox2.TabIndex = 50;
            this.groupBox2.TabStop = false;
            this.groupBox2.Text = "Theory";
            // 
            // rtfail
            // 
            this.rtfail.AutoSize = true;
            this.rtfail.Location = new System.Drawing.Point(28, 69);
            this.rtfail.Name = "rtfail";
            this.rtfail.Size = new System.Drawing.Size(41, 17);
            this.rtfail.TabIndex = 25;
            this.rtfail.TabStop = true;
            this.rtfail.Text = "Fail";
            this.rtfail.UseVisualStyleBackColor = true;
            // 
            // rtpass
            // 
            this.rtpass.AutoSize = true;
            this.rtpass.Location = new System.Drawing.Point(28, 31);
            this.rtpass.Name = "rtpass";
            this.rtpass.Size = new System.Drawing.Size(48, 17);
            this.rtpass.TabIndex = 24;
            this.rtpass.TabStop = true;
            this.rtpass.Text = "Pass";
            this.rtpass.UseVisualStyleBackColor = true;
            // 
            // dataGridView1
            // 
            this.dataGridView1.ColumnHeadersHeightSizeMode = System.Windows.Forms.DataGridViewColumnHeadersHeightSizeMode.AutoSize;
            this.dataGridView1.Dock = System.Windows.Forms.DockStyle.Bottom;
            this.dataGridView1.Location = new System.Drawing.Point(0, 375);
            this.dataGridView1.Name = "dataGridView1";
            this.dataGridView1.Size = new System.Drawing.Size(884, 237);
            this.dataGridView1.TabIndex = 49;
            this.dataGridView1.CellContentClick += new System.Windows.Forms.DataGridViewCellEventHandler(this.dataGridView1_CellContentClick);
            this.dataGridView1.RowHeaderMouseClick += new System.Windows.Forms.DataGridViewCellMouseEventHandler(this.dataGridView1_RowHeaderMouseClick);
            // 
            // label9
            // 
            this.label9.AutoSize = true;
            this.label9.Location = new System.Drawing.Point(227, 161);
            this.label9.Name = "label9";
            this.label9.Size = new System.Drawing.Size(53, 13);
            this.label9.TabIndex = 46;
            this.label9.Text = "Index_No";
            // 
            // label7
            // 
            this.label7.AutoSize = true;
            this.label7.Location = new System.Drawing.Point(226, 67);
            this.label7.Name = "label7";
            this.label7.Size = new System.Drawing.Size(45, 13);
            this.label7.TabIndex = 45;
            this.label7.Text = "NIC_No";
            // 
            // label3
            // 
            this.label3.AutoSize = true;
            this.label3.Location = new System.Drawing.Point(226, 111);
            this.label3.Name = "label3";
            this.label3.Size = new System.Drawing.Size(51, 13);
            this.label3.TabIndex = 41;
            this.label3.Text = "St_Name";
            // 
            // label2
            // 
            this.label2.AutoSize = true;
            this.label2.Location = new System.Drawing.Point(226, 32);
            this.label2.Name = "label2";
            this.label2.Size = new System.Drawing.Size(46, 13);
            this.label2.TabIndex = 40;
            this.label2.Text = "MIS_No";
            // 
            // txtindex
            // 
            this.txtindex.Location = new System.Drawing.Point(351, 154);
            this.txtindex.Name = "txtindex";
            this.txtindex.Size = new System.Drawing.Size(160, 20);
            this.txtindex.TabIndex = 38;
            // 
            // txtnic
            // 
            this.txtnic.Enabled = false;
            this.txtnic.Location = new System.Drawing.Point(351, 60);
            this.txtnic.Name = "txtnic";
            this.txtnic.Size = new System.Drawing.Size(160, 20);
            this.txtnic.TabIndex = 37;
            // 
            // txtname
            // 
            this.txtname.Enabled = false;
            this.txtname.Location = new System.Drawing.Point(350, 104);
            this.txtname.Name = "txtname";
            this.txtname.Size = new System.Drawing.Size(160, 20);
            this.txtname.TabIndex = 33;
            // 
            // txtmis
            // 
            this.txtmis.Enabled = false;
            this.txtmis.Location = new System.Drawing.Point(351, 25);
            this.txtmis.Name = "txtmis";
            this.txtmis.Size = new System.Drawing.Size(160, 20);
            this.txtmis.TabIndex = 32;
            // 
            // groupBox3
            // 
            this.groupBox3.Controls.Add(this.rpfail);
            this.groupBox3.Controls.Add(this.rppass);
            this.groupBox3.Location = new System.Drawing.Point(743, 179);
            this.groupBox3.Name = "groupBox3";
            this.groupBox3.Size = new System.Drawing.Size(109, 107);
            this.groupBox3.TabIndex = 56;
            this.groupBox3.TabStop = false;
            this.groupBox3.Text = "Practical";
            // 
            // rpfail
            // 
            this.rpfail.AutoSize = true;
            this.rpfail.Location = new System.Drawing.Point(28, 69);
            this.rpfail.Name = "rpfail";
            this.rpfail.Size = new System.Drawing.Size(41, 17);
            this.rpfail.TabIndex = 25;
            this.rpfail.TabStop = true;
            this.rpfail.Text = "Fail";
            this.rpfail.UseVisualStyleBackColor = true;
            // 
            // rppass
            // 
            this.rppass.AutoSize = true;
            this.rppass.Location = new System.Drawing.Point(28, 31);
            this.rppass.Name = "rppass";
            this.rppass.Size = new System.Drawing.Size(48, 17);
            this.rppass.TabIndex = 24;
            this.rppass.TabStop = true;
            this.rppass.Text = "Pass";
            this.rppass.UseVisualStyleBackColor = true;
            // 
            // groupBox1
            // 
            this.groupBox1.Controls.Add(this.rissue);
            this.groupBox1.Controls.Add(this.rprint);
            this.groupBox1.Location = new System.Drawing.Point(542, 35);
            this.groupBox1.Name = "groupBox1";
            this.groupBox1.Size = new System.Drawing.Size(146, 115);
            this.groupBox1.TabIndex = 57;
            this.groupBox1.TabStop = false;
            this.groupBox1.Text = "Certificate_Status";
            // 
            // rissue
            // 
            this.rissue.AutoSize = true;
            this.rissue.Location = new System.Drawing.Point(28, 69);
            this.rissue.Name = "rissue";
            this.rissue.Size = new System.Drawing.Size(56, 17);
            this.rissue.TabIndex = 25;
            this.rissue.TabStop = true;
            this.rissue.Text = "Issued";
            this.rissue.UseVisualStyleBackColor = true;
            // 
            // rprint
            // 
            this.rprint.AutoSize = true;
            this.rprint.Location = new System.Drawing.Point(28, 31);
            this.rprint.Name = "rprint";
            this.rprint.Size = new System.Drawing.Size(104, 17);
            this.rprint.TabIndex = 24;
            this.rprint.TabStop = true;
            this.rprint.Text = "Printing_Process";
            this.rprint.UseVisualStyleBackColor = true;
            // 
            // groupBox4
            // 
            this.groupBox4.Controls.Add(this.rn4);
            this.groupBox4.Controls.Add(this.rn3);
            this.groupBox4.Controls.Add(this.rn2);
            this.groupBox4.Location = new System.Drawing.Point(743, 35);
            this.groupBox4.Name = "groupBox4";
            this.groupBox4.Size = new System.Drawing.Size(109, 134);
            this.groupBox4.TabIndex = 58;
            this.groupBox4.TabStop = false;
            this.groupBox4.Text = "NVQ_Level";
            // 
            // rn4
            // 
            this.rn4.AutoSize = true;
            this.rn4.Location = new System.Drawing.Point(28, 102);
            this.rn4.Name = "rn4";
            this.rn4.Size = new System.Drawing.Size(63, 17);
            this.rn4.TabIndex = 26;
            this.rn4.TabStop = true;
            this.rn4.Text = "NVQ - 4";
            this.rn4.UseVisualStyleBackColor = true;
            // 
            // rn3
            // 
            this.rn3.AutoSize = true;
            this.rn3.Location = new System.Drawing.Point(28, 69);
            this.rn3.Name = "rn3";
            this.rn3.Size = new System.Drawing.Size(63, 17);
            this.rn3.TabIndex = 25;
            this.rn3.TabStop = true;
            this.rn3.Text = "NVQ - 3";
            this.rn3.UseVisualStyleBackColor = true;
            // 
            // rn2
            // 
            this.rn2.AutoSize = true;
            this.rn2.Location = new System.Drawing.Point(28, 31);
            this.rn2.Name = "rn2";
            this.rn2.Size = new System.Drawing.Size(63, 17);
            this.rn2.TabIndex = 24;
            this.rn2.TabStop = true;
            this.rn2.Text = "NVQ - 2";
            this.rn2.UseVisualStyleBackColor = true;
            // 
            // lb
            // 
            this.lb.FormattingEnabled = true;
            this.lb.Location = new System.Drawing.Point(15, 111);
            this.lb.Name = "lb";
            this.lb.Size = new System.Drawing.Size(189, 186);
            this.lb.TabIndex = 59;
            this.lb.SelectedIndexChanged += new System.EventHandler(this.lb_SelectedIndexChanged);
            // 
            // groupBox5
            // 
            this.groupBox5.Controls.Add(this.r2);
            this.groupBox5.Controls.Add(this.r1);
            this.groupBox5.Location = new System.Drawing.Point(107, 25);
            this.groupBox5.Name = "groupBox5";
            this.groupBox5.Size = new System.Drawing.Size(97, 78);
            this.groupBox5.TabIndex = 60;
            this.groupBox5.TabStop = false;
            this.groupBox5.Text = "Batch";
            // 
            // r2
            // 
            this.r2.AutoSize = true;
            this.r2.Location = new System.Drawing.Point(15, 55);
            this.r2.Name = "r2";
            this.r2.Size = new System.Drawing.Size(43, 17);
            this.r2.TabIndex = 1;
            this.r2.TabStop = true;
            this.r2.Text = "2nd";
            this.r2.UseVisualStyleBackColor = true;
            this.r2.CheckedChanged += new System.EventHandler(this.r2_CheckedChanged);
            // 
            // r1
            // 
            this.r1.AutoSize = true;
            this.r1.Location = new System.Drawing.Point(15, 19);
            this.r1.Name = "r1";
            this.r1.Size = new System.Drawing.Size(39, 17);
            this.r1.TabIndex = 0;
            this.r1.TabStop = true;
            this.r1.Text = "1st";
            this.r1.UseVisualStyleBackColor = true;
            this.r1.CheckedChanged += new System.EventHandler(this.r1_CheckedChanged);
            // 
            // cyear
            // 
            this.cyear.FormattingEnabled = true;
            this.cyear.Items.AddRange(new object[] {
            "2023",
            "2024",
            "2025",
            "2026",
            "2027",
            "2028",
            "2029",
            "2030",
            "2031",
            "2032",
            "2033",
            "2034",
            "2035",
            "2036",
            "2037",
            "2038",
            "2039",
            "2040"});
            this.cyear.Location = new System.Drawing.Point(12, 56);
            this.cyear.Name = "cyear";
            this.cyear.Size = new System.Drawing.Size(68, 21);
            this.cyear.TabIndex = 61;
            this.cyear.SelectedIndexChanged += new System.EventHandler(this.cyear_SelectedIndexChanged);
            // 
            // label10
            // 
            this.label10.AutoSize = true;
            this.label10.Location = new System.Drawing.Point(233, 214);
            this.label10.Name = "label10";
            this.label10.Size = new System.Drawing.Size(93, 13);
            this.label10.TabIndex = 47;
            this.label10.Text = "St_Current_Status";
            // 
            // cbs
            // 
            this.cbs.FormattingEnabled = true;
            this.cbs.Items.AddRange(new object[] {
            "Studies",
            "Work",
            "Other"});
            this.cbs.Location = new System.Drawing.Point(351, 211);
            this.cbs.Name = "cbs";
            this.cbs.Size = new System.Drawing.Size(121, 21);
            this.cbs.TabIndex = 62;
            // 
            // label1
            // 
            this.label1.AutoSize = true;
            this.label1.Location = new System.Drawing.Point(33, 35);
            this.label1.Name = "label1";
            this.label1.Size = new System.Drawing.Size(29, 13);
            this.label1.TabIndex = 40;
            this.label1.Text = "Year";
            // 
            // button4
            // 
            this.button4.Location = new System.Drawing.Point(364, 320);
            this.button4.Name = "button4";
            this.button4.Size = new System.Drawing.Size(63, 29);
            this.button4.TabIndex = 63;
            this.button4.Text = "Clear";
            this.button4.UseVisualStyleBackColor = true;
            this.button4.Click += new System.EventHandler(this.button4_Click);
            // 
            // txtsearch
            // 
            this.txtsearch.Location = new System.Drawing.Point(27, 329);
            this.txtsearch.Name = "txtsearch";
            this.txtsearch.Size = new System.Drawing.Size(154, 20);
            this.txtsearch.TabIndex = 64;
            this.txtsearch.TextChanged += new System.EventHandler(this.txtsearch_TextChanged);
            // 
            // cbatch
            // 
            this.cbatch.FormattingEnabled = true;
            this.cbatch.Items.AddRange(new object[] {
            "1st",
            "2nd"});
            this.cbatch.Location = new System.Drawing.Point(581, 331);
            this.cbatch.Name = "cbatch";
            this.cbatch.Size = new System.Drawing.Size(56, 21);
            this.cbatch.TabIndex = 79;
            // 
            // button3
            // 
            this.button3.Image = global::VTA_Interface.Properties.Resources.excel;
            this.button3.ImageAlign = System.Drawing.ContentAlignment.MiddleLeft;
            this.button3.Location = new System.Drawing.Point(454, 317);
            this.button3.Name = "button3";
            this.button3.Size = new System.Drawing.Size(97, 34);
            this.button3.TabIndex = 99;
            this.button3.Text = "EXPORT";
            this.button3.TextAlign = System.Drawing.ContentAlignment.MiddleRight;
            this.button3.UseVisualStyleBackColor = true;
            this.button3.Click += new System.EventHandler(this.button3_Click_1);
            // 
            // label11
            // 
            this.label11.AutoSize = true;
            this.label11.Location = new System.Drawing.Point(592, 313);
            this.label11.Name = "label11";
            this.label11.Size = new System.Drawing.Size(35, 13);
            this.label11.TabIndex = 100;
            this.label11.Text = "Batch";
            // 
            // label4
            // 
            this.label4.AutoSize = true;
            this.label4.Location = new System.Drawing.Point(86, 306);
            this.label4.Name = "label4";
            this.label4.Size = new System.Drawing.Size(41, 13);
            this.label4.TabIndex = 101;
            this.label4.Text = "Search";
            // 
            // Exam_Details
            // 
            this.AutoScaleDimensions = new System.Drawing.SizeF(6F, 13F);
            this.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font;
            this.ClientSize = new System.Drawing.Size(884, 612);
            this.Controls.Add(this.label4);
            this.Controls.Add(this.label11);
            this.Controls.Add(this.button3);
            this.Controls.Add(this.cbatch);
            this.Controls.Add(this.txtsearch);
            this.Controls.Add(this.button4);
            this.Controls.Add(this.cbs);
            this.Controls.Add(this.cyear);
            this.Controls.Add(this.groupBox5);
            this.Controls.Add(this.lb);
            this.Controls.Add(this.groupBox4);
            this.Controls.Add(this.groupBox1);
            this.Controls.Add(this.groupBox3);
            this.Controls.Add(this.button2);
            this.Controls.Add(this.button1);
            this.Controls.Add(this.textBox9);
            this.Controls.Add(this.label13);
            this.Controls.Add(this.groupBox2);
            this.Controls.Add(this.dataGridView1);
            this.Controls.Add(this.label10);
            this.Controls.Add(this.label9);
            this.Controls.Add(this.label7);
            this.Controls.Add(this.label3);
            this.Controls.Add(this.label1);
            this.Controls.Add(this.label2);
            this.Controls.Add(this.txtindex);
            this.Controls.Add(this.txtnic);
            this.Controls.Add(this.txtname);
            this.Controls.Add(this.txtmis);
            this.Name = "Exam_Details";
            this.Text = "Exam_Details";
            this.Load += new System.EventHandler(this.Exam_Details_Load);
            this.groupBox2.ResumeLayout(false);
            this.groupBox2.PerformLayout();
            ((System.ComponentModel.ISupportInitialize)(this.dataGridView1)).EndInit();
            this.groupBox3.ResumeLayout(false);
            this.groupBox3.PerformLayout();
            this.groupBox1.ResumeLayout(false);
            this.groupBox1.PerformLayout();
            this.groupBox4.ResumeLayout(false);
            this.groupBox4.PerformLayout();
            this.groupBox5.ResumeLayout(false);
            this.groupBox5.PerformLayout();
            this.ResumeLayout(false);
            this.PerformLayout();

        }

        #endregion

        private System.Windows.Forms.Button button2;
        private System.Windows.Forms.Button button1;
        private System.Windows.Forms.TextBox textBox9;
        private System.Windows.Forms.Label label13;
        private System.Windows.Forms.GroupBox groupBox2;
        private System.Windows.Forms.DataGridView dataGridView1;
        private System.Windows.Forms.Label label9;
        private System.Windows.Forms.Label label7;
        private System.Windows.Forms.Label label3;
        private System.Windows.Forms.Label label2;
        private System.Windows.Forms.TextBox txtindex;
        private System.Windows.Forms.TextBox txtnic;
        private System.Windows.Forms.TextBox txtname;
        private System.Windows.Forms.TextBox txtmis;
        private System.Windows.Forms.RadioButton rtfail;
        private System.Windows.Forms.RadioButton rtpass;
        private System.Windows.Forms.GroupBox groupBox3;
        private System.Windows.Forms.RadioButton rpfail;
        private System.Windows.Forms.RadioButton rppass;
        private System.Windows.Forms.GroupBox groupBox1;
        private System.Windows.Forms.RadioButton rissue;
        private System.Windows.Forms.RadioButton rprint;
        private System.Windows.Forms.GroupBox groupBox4;
        private System.Windows.Forms.RadioButton rn4;
        private System.Windows.Forms.RadioButton rn3;
        private System.Windows.Forms.RadioButton rn2;
        private System.Windows.Forms.ListBox lb;
        private System.Windows.Forms.GroupBox groupBox5;
        private System.Windows.Forms.RadioButton r2;
        private System.Windows.Forms.RadioButton r1;
        private System.Windows.Forms.ComboBox cyear;
        private System.Windows.Forms.Label label10;
        private System.Windows.Forms.ComboBox cbs;
        private System.Windows.Forms.Label label1;
        private System.Windows.Forms.Button button4;
        private System.Windows.Forms.TextBox txtsearch;
        private System.Windows.Forms.ComboBox cbatch;
        private System.Windows.Forms.Button button3;
        private System.Windows.Forms.Label label11;
        private System.Windows.Forms.Label label4;
    }
}