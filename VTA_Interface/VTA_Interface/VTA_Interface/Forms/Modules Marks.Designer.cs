namespace VTA_Interface.Forms
{
    partial class Modules_Marks
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
            this.BTNEDIT = new System.Windows.Forms.Button();
            this.button1 = new System.Windows.Forms.Button();
            this.groupBox1 = new System.Windows.Forms.GroupBox();
            this.btnclose = new System.Windows.Forms.Button();
            this.btnok = new System.Windows.Forms.Button();
            this.lmname = new System.Windows.Forms.Label();
            this.label4 = new System.Windows.Forms.Label();
            this.label3 = new System.Windows.Forms.Label();
            this.txtprac = new System.Windows.Forms.TextBox();
            this.txttheo = new System.Windows.Forms.TextBox();
            this.textBox10 = new System.Windows.Forms.TextBox();
            this.label6 = new System.Windows.Forms.Label();
            this.label2 = new System.Windows.Forms.Label();
            this.label1 = new System.Windows.Forms.Label();
            this.txtname = new System.Windows.Forms.TextBox();
            this.txtmis = new System.Windows.Forms.TextBox();
            this.mname = new System.Windows.Forms.ComboBox();
            this.dataGridView1 = new System.Windows.Forms.DataGridView();
            this.label5 = new System.Windows.Forms.Label();
            this.txtsearch = new System.Windows.Forms.TextBox();
            this.lb = new System.Windows.Forms.ListBox();
            this.year = new System.Windows.Forms.ComboBox();
            this.r1 = new System.Windows.Forms.RadioButton();
            this.r2 = new System.Windows.Forms.RadioButton();
            this.button4 = new System.Windows.Forms.Button();
            this.label7 = new System.Windows.Forms.Label();
            this.cbatch = new System.Windows.Forms.ComboBox();
            this.button3 = new System.Windows.Forms.Button();
            this.label11 = new System.Windows.Forms.Label();
            this.label8 = new System.Windows.Forms.Label();
            this.groupBox1.SuspendLayout();
            ((System.ComponentModel.ISupportInitialize)(this.dataGridView1)).BeginInit();
            this.SuspendLayout();
            // 
            // BTNEDIT
            // 
            this.BTNEDIT.Location = new System.Drawing.Point(341, 338);
            this.BTNEDIT.Name = "BTNEDIT";
            this.BTNEDIT.Size = new System.Drawing.Size(63, 27);
            this.BTNEDIT.TabIndex = 64;
            this.BTNEDIT.Text = "Edit";
            this.BTNEDIT.UseVisualStyleBackColor = true;
            this.BTNEDIT.Click += new System.EventHandler(this.button2_Click);
            // 
            // button1
            // 
            this.button1.Location = new System.Drawing.Point(267, 338);
            this.button1.Name = "button1";
            this.button1.Size = new System.Drawing.Size(63, 27);
            this.button1.TabIndex = 63;
            this.button1.Text = "Add";
            this.button1.UseVisualStyleBackColor = true;
            this.button1.Click += new System.EventHandler(this.button1_Click);
            // 
            // groupBox1
            // 
            this.groupBox1.Controls.Add(this.btnclose);
            this.groupBox1.Controls.Add(this.btnok);
            this.groupBox1.Controls.Add(this.lmname);
            this.groupBox1.Controls.Add(this.label4);
            this.groupBox1.Controls.Add(this.label3);
            this.groupBox1.Controls.Add(this.txtprac);
            this.groupBox1.Controls.Add(this.txttheo);
            this.groupBox1.Location = new System.Drawing.Point(384, 97);
            this.groupBox1.Name = "groupBox1";
            this.groupBox1.Size = new System.Drawing.Size(466, 151);
            this.groupBox1.TabIndex = 62;
            this.groupBox1.TabStop = false;
            this.groupBox1.Text = "Marks";
            this.groupBox1.Enter += new System.EventHandler(this.groupBox1_Enter);
            // 
            // btnclose
            // 
            this.btnclose.Location = new System.Drawing.Point(346, 67);
            this.btnclose.Name = "btnclose";
            this.btnclose.Size = new System.Drawing.Size(38, 25);
            this.btnclose.TabIndex = 77;
            this.btnclose.Text = "X";
            this.btnclose.UseVisualStyleBackColor = true;
            this.btnclose.Click += new System.EventHandler(this.btnclose_Click);
            // 
            // btnok
            // 
            this.btnok.Location = new System.Drawing.Point(302, 67);
            this.btnok.Name = "btnok";
            this.btnok.Size = new System.Drawing.Size(38, 25);
            this.btnok.TabIndex = 76;
            this.btnok.Text = "OK";
            this.btnok.UseVisualStyleBackColor = true;
            this.btnok.Click += new System.EventHandler(this.btnok_Click);
            // 
            // lmname
            // 
            this.lmname.AutoSize = true;
            this.lmname.Location = new System.Drawing.Point(21, 25);
            this.lmname.Name = "lmname";
            this.lmname.Size = new System.Drawing.Size(0, 13);
            this.lmname.TabIndex = 35;
            // 
            // label4
            // 
            this.label4.AutoSize = true;
            this.label4.Location = new System.Drawing.Point(169, 85);
            this.label4.Name = "label4";
            this.label4.Size = new System.Drawing.Size(48, 13);
            this.label4.TabIndex = 34;
            this.label4.Text = "Practical";
            // 
            // label3
            // 
            this.label3.AutoSize = true;
            this.label3.Location = new System.Drawing.Point(166, 48);
            this.label3.Name = "label3";
            this.label3.Size = new System.Drawing.Size(40, 13);
            this.label3.TabIndex = 33;
            this.label3.Text = "Theory";
            // 
            // txtprac
            // 
            this.txtprac.Location = new System.Drawing.Point(238, 84);
            this.txtprac.Name = "txtprac";
            this.txtprac.Size = new System.Drawing.Size(58, 20);
            this.txtprac.TabIndex = 27;
            this.txtprac.KeyPress += new System.Windows.Forms.KeyPressEventHandler(this.txtprac_KeyPress);
            // 
            // txttheo
            // 
            this.txttheo.Location = new System.Drawing.Point(238, 43);
            this.txttheo.Name = "txttheo";
            this.txttheo.Size = new System.Drawing.Size(58, 20);
            this.txttheo.TabIndex = 26;
            this.txttheo.TextChanged += new System.EventHandler(this.txttheo_TextChanged);
            this.txttheo.KeyPress += new System.Windows.Forms.KeyPressEventHandler(this.txttheo_KeyPress);
            // 
            // textBox10
            // 
            this.textBox10.Location = new System.Drawing.Point(625, 346);
            this.textBox10.Name = "textBox10";
            this.textBox10.Size = new System.Drawing.Size(170, 20);
            this.textBox10.TabIndex = 60;
            this.textBox10.TextChanged += new System.EventHandler(this.textBox10_TextChanged);
            // 
            // label6
            // 
            this.label6.AutoSize = true;
            this.label6.Location = new System.Drawing.Point(383, 74);
            this.label6.Name = "label6";
            this.label6.Size = new System.Drawing.Size(73, 13);
            this.label6.TabIndex = 54;
            this.label6.Text = "Module Name";
            // 
            // label2
            // 
            this.label2.AutoSize = true;
            this.label2.Location = new System.Drawing.Point(27, 346);
            this.label2.Name = "label2";
            this.label2.Size = new System.Drawing.Size(51, 13);
            this.label2.TabIndex = 53;
            this.label2.Text = "St_Name";
            // 
            // label1
            // 
            this.label1.AutoSize = true;
            this.label1.Location = new System.Drawing.Point(27, 309);
            this.label1.Name = "label1";
            this.label1.Size = new System.Drawing.Size(46, 13);
            this.label1.TabIndex = 52;
            this.label1.Text = "MIS_No";
            // 
            // txtname
            // 
            this.txtname.Enabled = false;
            this.txtname.Location = new System.Drawing.Point(92, 342);
            this.txtname.Name = "txtname";
            this.txtname.Size = new System.Drawing.Size(148, 20);
            this.txtname.TabIndex = 49;
            // 
            // txtmis
            // 
            this.txtmis.Enabled = false;
            this.txtmis.Location = new System.Drawing.Point(92, 305);
            this.txtmis.Name = "txtmis";
            this.txtmis.Size = new System.Drawing.Size(148, 20);
            this.txtmis.TabIndex = 48;
            this.txtmis.TextChanged += new System.EventHandler(this.txtmis_TextChanged);
            // 
            // mname
            // 
            this.mname.FormattingEnabled = true;
            this.mname.Items.AddRange(new object[] {
            "Maintain Files & Folders (M1)",
            "Perform Word Processing (M2)",
            "Prepare Spread Sheet (M3)",
            "Prepare Presentations (M4)",
            "Maintain Databases (M5)",
            "Perform Internet and Electronic Mail Operations (M6)",
            "Install and Configure Operating System (M7)",
            "Conduct Installation And Troubleshooting of Network From Client PC (M8)",
            "Develop Graphics for Web / Print Media (M9)",
            "Design and Develop Static Web Pages (M10)",
            "Develop Basic Software Solutions (M11)"});
            this.mname.Location = new System.Drawing.Point(462, 70);
            this.mname.Name = "mname";
            this.mname.Size = new System.Drawing.Size(388, 21);
            this.mname.TabIndex = 66;
            this.mname.SelectedIndexChanged += new System.EventHandler(this.comboBox1_SelectedIndexChanged);
            // 
            // dataGridView1
            // 
            this.dataGridView1.ColumnHeadersHeightSizeMode = System.Windows.Forms.DataGridViewColumnHeadersHeightSizeMode.AutoSize;
            this.dataGridView1.Dock = System.Windows.Forms.DockStyle.Bottom;
            this.dataGridView1.Location = new System.Drawing.Point(0, 387);
            this.dataGridView1.Name = "dataGridView1";
            this.dataGridView1.Size = new System.Drawing.Size(884, 225);
            this.dataGridView1.TabIndex = 67;
            this.dataGridView1.CellContentClick += new System.Windows.Forms.DataGridViewCellEventHandler(this.dataGridView1_CellContentClick);
            this.dataGridView1.RowHeaderMouseClick += new System.Windows.Forms.DataGridViewCellMouseEventHandler(this.dataGridView1_RowHeaderMouseClick);
            // 
            // label5
            // 
            this.label5.AutoSize = true;
            this.label5.Location = new System.Drawing.Point(108, 279);
            this.label5.Name = "label5";
            this.label5.Size = new System.Drawing.Size(41, 13);
            this.label5.TabIndex = 68;
            this.label5.Text = "Search";
            // 
            // txtsearch
            // 
            this.txtsearch.Location = new System.Drawing.Point(28, 256);
            this.txtsearch.Name = "txtsearch";
            this.txtsearch.Size = new System.Drawing.Size(211, 20);
            this.txtsearch.TabIndex = 69;
            this.txtsearch.Click += new System.EventHandler(this.textBox5_Click);
            this.txtsearch.TextChanged += new System.EventHandler(this.txtsearch_TextChanged);
            this.txtsearch.DoubleClick += new System.EventHandler(this.txtsearch_DoubleClick);
            // 
            // lb
            // 
            this.lb.FormattingEnabled = true;
            this.lb.Location = new System.Drawing.Point(28, 88);
            this.lb.Name = "lb";
            this.lb.Size = new System.Drawing.Size(212, 160);
            this.lb.TabIndex = 71;
            this.lb.Click += new System.EventHandler(this.lb_Click);
            this.lb.MouseClick += new System.Windows.Forms.MouseEventHandler(this.lb_MouseClick);
            this.lb.SelectedIndexChanged += new System.EventHandler(this.lb_SelectedIndexChanged);
            // 
            // year
            // 
            this.year.FormattingEnabled = true;
            this.year.Items.AddRange(new object[] {
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
            this.year.Location = new System.Drawing.Point(28, 52);
            this.year.Name = "year";
            this.year.Size = new System.Drawing.Size(121, 21);
            this.year.TabIndex = 72;
            this.year.SelectedIndexChanged += new System.EventHandler(this.modules_SelectedIndexChanged);
            // 
            // r1
            // 
            this.r1.AutoSize = true;
            this.r1.Location = new System.Drawing.Point(165, 54);
            this.r1.Name = "r1";
            this.r1.Size = new System.Drawing.Size(70, 17);
            this.r1.TabIndex = 73;
            this.r1.TabStop = true;
            this.r1.Text = "1st Batch";
            this.r1.UseVisualStyleBackColor = true;
            this.r1.CheckedChanged += new System.EventHandler(this.r1_CheckedChanged);
            this.r1.Click += new System.EventHandler(this.r1_Click);
            // 
            // r2
            // 
            this.r2.AutoSize = true;
            this.r2.Location = new System.Drawing.Point(234, 54);
            this.r2.Name = "r2";
            this.r2.Size = new System.Drawing.Size(74, 17);
            this.r2.TabIndex = 74;
            this.r2.TabStop = true;
            this.r2.Text = "2nd Batch";
            this.r2.UseVisualStyleBackColor = true;
            this.r2.CheckedChanged += new System.EventHandler(this.r2_CheckedChanged);
            // 
            // button4
            // 
            this.button4.Location = new System.Drawing.Point(414, 338);
            this.button4.Name = "button4";
            this.button4.Size = new System.Drawing.Size(68, 27);
            this.button4.TabIndex = 75;
            this.button4.Text = "Clear";
            this.button4.UseVisualStyleBackColor = true;
            this.button4.Click += new System.EventHandler(this.button4_Click);
            // 
            // label7
            // 
            this.label7.AutoSize = true;
            this.label7.Location = new System.Drawing.Point(573, 349);
            this.label7.Name = "label7";
            this.label7.Size = new System.Drawing.Size(41, 13);
            this.label7.TabIndex = 34;
            this.label7.Text = "Search";
            // 
            // cbatch
            // 
            this.cbatch.FormattingEnabled = true;
            this.cbatch.Items.AddRange(new object[] {
            "1st",
            "2nd"});
            this.cbatch.Location = new System.Drawing.Point(801, 346);
            this.cbatch.Name = "cbatch";
            this.cbatch.Size = new System.Drawing.Size(56, 21);
            this.cbatch.TabIndex = 76;
            // 
            // button3
            // 
            this.button3.Image = global::VTA_Interface.Properties.Resources.excel;
            this.button3.ImageAlign = System.Drawing.ContentAlignment.MiddleLeft;
            this.button3.Location = new System.Drawing.Point(662, 298);
            this.button3.Name = "button3";
            this.button3.Size = new System.Drawing.Size(97, 33);
            this.button3.TabIndex = 79;
            this.button3.Text = "EXPORT";
            this.button3.TextAlign = System.Drawing.ContentAlignment.MiddleRight;
            this.button3.UseVisualStyleBackColor = true;
            this.button3.Click += new System.EventHandler(this.button3_Click_1);
            // 
            // label11
            // 
            this.label11.AutoSize = true;
            this.label11.Location = new System.Drawing.Point(812, 329);
            this.label11.Name = "label11";
            this.label11.Size = new System.Drawing.Size(35, 13);
            this.label11.TabIndex = 99;
            this.label11.Text = "Batch";
            // 
            // label8
            // 
            this.label8.AutoSize = true;
            this.label8.Location = new System.Drawing.Point(77, 29);
            this.label8.Name = "label8";
            this.label8.Size = new System.Drawing.Size(29, 13);
            this.label8.TabIndex = 100;
            this.label8.Text = "Year";
            // 
            // Modules_Marks
            // 
            this.AutoScaleDimensions = new System.Drawing.SizeF(6F, 13F);
            this.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font;
            this.ClientSize = new System.Drawing.Size(884, 612);
            this.Controls.Add(this.label8);
            this.Controls.Add(this.label11);
            this.Controls.Add(this.button3);
            this.Controls.Add(this.cbatch);
            this.Controls.Add(this.button4);
            this.Controls.Add(this.label7);
            this.Controls.Add(this.textBox10);
            this.Controls.Add(this.r2);
            this.Controls.Add(this.txtsearch);
            this.Controls.Add(this.r1);
            this.Controls.Add(this.label5);
            this.Controls.Add(this.year);
            this.Controls.Add(this.lb);
            this.Controls.Add(this.dataGridView1);
            this.Controls.Add(this.mname);
            this.Controls.Add(this.BTNEDIT);
            this.Controls.Add(this.button1);
            this.Controls.Add(this.groupBox1);
            this.Controls.Add(this.label6);
            this.Controls.Add(this.label2);
            this.Controls.Add(this.label1);
            this.Controls.Add(this.txtname);
            this.Controls.Add(this.txtmis);
            this.Name = "Modules_Marks";
            this.Text = "Modules_Marks";
            this.Load += new System.EventHandler(this.Modules_Marks_Load);
            this.groupBox1.ResumeLayout(false);
            this.groupBox1.PerformLayout();
            ((System.ComponentModel.ISupportInitialize)(this.dataGridView1)).EndInit();
            this.ResumeLayout(false);
            this.PerformLayout();

        }

        #endregion

        private System.Windows.Forms.Button BTNEDIT;
        private System.Windows.Forms.Button button1;
        private System.Windows.Forms.GroupBox groupBox1;
        private System.Windows.Forms.Label label4;
        private System.Windows.Forms.Label label3;
        private System.Windows.Forms.TextBox txtprac;
        private System.Windows.Forms.TextBox txttheo;
        private System.Windows.Forms.TextBox textBox10;
        private System.Windows.Forms.Label label6;
        private System.Windows.Forms.Label label2;
        private System.Windows.Forms.Label label1;
        private System.Windows.Forms.TextBox txtname;
        private System.Windows.Forms.TextBox txtmis;
        private System.Windows.Forms.ComboBox mname;
        private System.Windows.Forms.DataGridView dataGridView1;
        private System.Windows.Forms.Label label5;
        private System.Windows.Forms.TextBox txtsearch;
        private System.Windows.Forms.ListBox lb;
        private System.Windows.Forms.ComboBox year;
        private System.Windows.Forms.RadioButton r1;
        private System.Windows.Forms.RadioButton r2;
        private System.Windows.Forms.Label lmname;
        private System.Windows.Forms.Button button4;
        private System.Windows.Forms.Label label7;
        private System.Windows.Forms.Button btnclose;
        private System.Windows.Forms.Button btnok;
        private System.Windows.Forms.ComboBox cbatch;
        private System.Windows.Forms.Button button3;
        private System.Windows.Forms.Label label11;
        private System.Windows.Forms.Label label8;
    }
}