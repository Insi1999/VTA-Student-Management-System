namespace VTA_Interface.Forms
{
    partial class Registration
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
            this.txtmis = new System.Windows.Forms.TextBox();
            this.txtstname = new System.Windows.Forms.TextBox();
            this.cyear = new System.Windows.Forms.ComboBox();
            this.datepicker = new System.Windows.Forms.DateTimePicker();
            this.txtnic = new System.Windows.Forms.TextBox();
            this.txtaddress = new System.Windows.Forms.TextBox();
            this.txtcontact = new System.Windows.Forms.TextBox();
            this.label2 = new System.Windows.Forms.Label();
            this.label3 = new System.Windows.Forms.Label();
            this.label4 = new System.Windows.Forms.Label();
            this.label6 = new System.Windows.Forms.Label();
            this.label7 = new System.Windows.Forms.Label();
            this.label8 = new System.Windows.Forms.Label();
            this.label9 = new System.Windows.Forms.Label();
            this.label10 = new System.Windows.Forms.Label();
            this.label11 = new System.Windows.Forms.Label();
            this.label12 = new System.Windows.Forms.Label();
            this.groupBox1 = new System.Windows.Forms.GroupBox();
            this.txtmath = new System.Windows.Forms.TextBox();
            this.txtscience = new System.Windows.Forms.TextBox();
            this.txtenglish = new System.Windows.Forms.TextBox();
            this.dataGridView1 = new System.Windows.Forms.DataGridView();
            this.rmale = new System.Windows.Forms.RadioButton();
            this.rfemale = new System.Windows.Forms.RadioButton();
            this.groupBox2 = new System.Windows.Forms.GroupBox();
            this.label13 = new System.Windows.Forms.Label();
            this.txtserreg = new System.Windows.Forms.TextBox();
            this.button1 = new System.Windows.Forms.Button();
            this.button2 = new System.Windows.Forms.Button();
            this.rsecond = new System.Windows.Forms.RadioButton();
            this.r1st = new System.Windows.Forms.RadioButton();
            this.checkBox1 = new System.Windows.Forms.CheckBox();
            this.groupBox3 = new System.Windows.Forms.GroupBox();
            this.button4 = new System.Windows.Forms.Button();
            this.label1 = new System.Windows.Forms.Label();
            this.txtmis1 = new System.Windows.Forms.TextBox();
            this.btnpicture = new System.Windows.Forms.Button();
            this.pic = new System.Windows.Forms.PictureBox();
            this.button5 = new System.Windows.Forms.Button();
            this.cbatch = new System.Windows.Forms.ComboBox();
            this.label5 = new System.Windows.Forms.Label();
            this.groupBox1.SuspendLayout();
            ((System.ComponentModel.ISupportInitialize)(this.dataGridView1)).BeginInit();
            this.groupBox2.SuspendLayout();
            this.groupBox3.SuspendLayout();
            ((System.ComponentModel.ISupportInitialize)(this.pic)).BeginInit();
            this.SuspendLayout();
            // 
            // txtmis
            // 
            this.txtmis.Enabled = false;
            this.txtmis.Location = new System.Drawing.Point(298, 92);
            this.txtmis.MaxLength = 2;
            this.txtmis.Name = "txtmis";
            this.txtmis.Size = new System.Drawing.Size(37, 20);
            this.txtmis.TabIndex = 0;
            this.txtmis.Text = "0";
            this.txtmis.TextChanged += new System.EventHandler(this.txtmis_TextChanged);
            this.txtmis.KeyPress += new System.Windows.Forms.KeyPressEventHandler(this.txtmis_KeyPress);
            // 
            // txtstname
            // 
            this.txtstname.Enabled = false;
            this.txtstname.Location = new System.Drawing.Point(127, 54);
            this.txtstname.Name = "txtstname";
            this.txtstname.Size = new System.Drawing.Size(160, 20);
            this.txtstname.TabIndex = 1;
            this.txtstname.TextChanged += new System.EventHandler(this.txtstname_TextChanged);
            this.txtstname.KeyPress += new System.Windows.Forms.KeyPressEventHandler(this.txtstname_KeyPress);
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
            this.cyear.Location = new System.Drawing.Point(127, 19);
            this.cyear.Name = "cyear";
            this.cyear.Size = new System.Drawing.Size(77, 21);
            this.cyear.TabIndex = 2;
            this.cyear.SelectedIndexChanged += new System.EventHandler(this.cyear_SelectedIndexChanged);
            // 
            // datepicker
            // 
            this.datepicker.CustomFormat = "yyyy-MM-dd";
            this.datepicker.Enabled = false;
            this.datepicker.Format = System.Windows.Forms.DateTimePickerFormat.Custom;
            this.datepicker.Location = new System.Drawing.Point(127, 185);
            this.datepicker.Name = "datepicker";
            this.datepicker.Size = new System.Drawing.Size(160, 20);
            this.datepicker.TabIndex = 4;
            // 
            // txtnic
            // 
            this.txtnic.Enabled = false;
            this.txtnic.Location = new System.Drawing.Point(127, 140);
            this.txtnic.MaxLength = 12;
            this.txtnic.Name = "txtnic";
            this.txtnic.Size = new System.Drawing.Size(160, 20);
            this.txtnic.TabIndex = 5;
            this.txtnic.TextChanged += new System.EventHandler(this.txtnic_TextChanged);
            this.txtnic.KeyPress += new System.Windows.Forms.KeyPressEventHandler(this.txtnic_KeyPress);
            // 
            // txtaddress
            // 
            this.txtaddress.Enabled = false;
            this.txtaddress.Location = new System.Drawing.Point(127, 223);
            this.txtaddress.Name = "txtaddress";
            this.txtaddress.Size = new System.Drawing.Size(160, 20);
            this.txtaddress.TabIndex = 7;
            this.txtaddress.TextChanged += new System.EventHandler(this.txtaddress_TextChanged);
            // 
            // txtcontact
            // 
            this.txtcontact.Enabled = false;
            this.txtcontact.Location = new System.Drawing.Point(127, 263);
            this.txtcontact.MaxLength = 10;
            this.txtcontact.Name = "txtcontact";
            this.txtcontact.Size = new System.Drawing.Size(160, 20);
            this.txtcontact.TabIndex = 8;
            this.txtcontact.TextChanged += new System.EventHandler(this.txtcontact_TextChanged);
            this.txtcontact.KeyPress += new System.Windows.Forms.KeyPressEventHandler(this.txtcontact_KeyPress);
            // 
            // label2
            // 
            this.label2.AutoSize = true;
            this.label2.Location = new System.Drawing.Point(33, 98);
            this.label2.Name = "label2";
            this.label2.Size = new System.Drawing.Size(46, 13);
            this.label2.TabIndex = 11;
            this.label2.Text = "MIS_No";
            // 
            // label3
            // 
            this.label3.AutoSize = true;
            this.label3.Location = new System.Drawing.Point(33, 56);
            this.label3.Name = "label3";
            this.label3.Size = new System.Drawing.Size(51, 13);
            this.label3.TabIndex = 12;
            this.label3.Text = "St_Name";
            // 
            // label4
            // 
            this.label4.AutoSize = true;
            this.label4.Location = new System.Drawing.Point(33, 23);
            this.label4.Name = "label4";
            this.label4.Size = new System.Drawing.Size(29, 13);
            this.label4.TabIndex = 13;
            this.label4.Text = "Year";
            // 
            // label6
            // 
            this.label6.AutoSize = true;
            this.label6.Location = new System.Drawing.Point(33, 189);
            this.label6.Name = "label6";
            this.label6.Size = new System.Drawing.Size(72, 13);
            this.label6.TabIndex = 15;
            this.label6.Text = "Date_of_Birth";
            // 
            // label7
            // 
            this.label7.AutoSize = true;
            this.label7.Location = new System.Drawing.Point(33, 144);
            this.label7.Name = "label7";
            this.label7.Size = new System.Drawing.Size(45, 13);
            this.label7.TabIndex = 16;
            this.label7.Text = "NIC_No";
            // 
            // label8
            // 
            this.label8.AutoSize = true;
            this.label8.Location = new System.Drawing.Point(54, 119);
            this.label8.Name = "label8";
            this.label8.Size = new System.Drawing.Size(41, 13);
            this.label8.TabIndex = 17;
            this.label8.Text = "English";
            // 
            // label9
            // 
            this.label9.AutoSize = true;
            this.label9.Location = new System.Drawing.Point(35, 226);
            this.label9.Name = "label9";
            this.label9.Size = new System.Drawing.Size(45, 13);
            this.label9.TabIndex = 18;
            this.label9.Text = "Address";
            // 
            // label10
            // 
            this.label10.AutoSize = true;
            this.label10.Location = new System.Drawing.Point(35, 269);
            this.label10.Name = "label10";
            this.label10.Size = new System.Drawing.Size(64, 13);
            this.label10.TabIndex = 19;
            this.label10.Text = "Contact_No";
            // 
            // label11
            // 
            this.label11.AutoSize = true;
            this.label11.Location = new System.Drawing.Point(56, 34);
            this.label11.Name = "label11";
            this.label11.Size = new System.Drawing.Size(46, 13);
            this.label11.TabIndex = 20;
            this.label11.Text = "Science";
            // 
            // label12
            // 
            this.label12.AutoSize = true;
            this.label12.Location = new System.Drawing.Point(56, 76);
            this.label12.Name = "label12";
            this.label12.Size = new System.Drawing.Size(36, 13);
            this.label12.TabIndex = 21;
            this.label12.Text = "Maths";
            // 
            // groupBox1
            // 
            this.groupBox1.Controls.Add(this.label12);
            this.groupBox1.Controls.Add(this.label11);
            this.groupBox1.Controls.Add(this.label8);
            this.groupBox1.Controls.Add(this.txtmath);
            this.groupBox1.Controls.Add(this.txtscience);
            this.groupBox1.Controls.Add(this.txtenglish);
            this.groupBox1.Location = new System.Drawing.Point(634, 23);
            this.groupBox1.Name = "groupBox1";
            this.groupBox1.Size = new System.Drawing.Size(226, 160);
            this.groupBox1.TabIndex = 22;
            this.groupBox1.TabStop = false;
            this.groupBox1.Text = "O/L Results";
            // 
            // txtmath
            // 
            this.txtmath.Enabled = false;
            this.txtmath.Location = new System.Drawing.Point(121, 73);
            this.txtmath.MaxLength = 1;
            this.txtmath.Name = "txtmath";
            this.txtmath.Size = new System.Drawing.Size(52, 20);
            this.txtmath.TabIndex = 10;
            this.txtmath.TextChanged += new System.EventHandler(this.txtmath_TextChanged);
            // 
            // txtscience
            // 
            this.txtscience.Enabled = false;
            this.txtscience.Location = new System.Drawing.Point(121, 30);
            this.txtscience.MaxLength = 1;
            this.txtscience.Name = "txtscience";
            this.txtscience.Size = new System.Drawing.Size(52, 20);
            this.txtscience.TabIndex = 9;
            this.txtscience.TextChanged += new System.EventHandler(this.txtscience_TextChanged);
            this.txtscience.KeyPress += new System.Windows.Forms.KeyPressEventHandler(this.txtscience_KeyPress);
            // 
            // txtenglish
            // 
            this.txtenglish.Enabled = false;
            this.txtenglish.Location = new System.Drawing.Point(121, 116);
            this.txtenglish.MaxLength = 1;
            this.txtenglish.Name = "txtenglish";
            this.txtenglish.Size = new System.Drawing.Size(52, 20);
            this.txtenglish.TabIndex = 6;
            this.txtenglish.TextChanged += new System.EventHandler(this.txtenglish_TextChanged);
            // 
            // dataGridView1
            // 
            this.dataGridView1.AllowUserToAddRows = false;
            this.dataGridView1.AllowUserToDeleteRows = false;
            this.dataGridView1.ColumnHeadersHeightSizeMode = System.Windows.Forms.DataGridViewColumnHeadersHeightSizeMode.AutoSize;
            this.dataGridView1.Dock = System.Windows.Forms.DockStyle.Bottom;
            this.dataGridView1.Location = new System.Drawing.Point(0, 384);
            this.dataGridView1.Name = "dataGridView1";
            this.dataGridView1.ReadOnly = true;
            this.dataGridView1.RowTemplate.Height = 100;
            this.dataGridView1.Size = new System.Drawing.Size(884, 228);
            this.dataGridView1.TabIndex = 23;
            this.dataGridView1.CellContentClick += new System.Windows.Forms.DataGridViewCellEventHandler(this.dataGridView1_CellContentClick);
            this.dataGridView1.RowHeaderMouseClick += new System.Windows.Forms.DataGridViewCellMouseEventHandler(this.dataGridView1_RowHeaderMouseClick);
            // 
            // rmale
            // 
            this.rmale.AutoSize = true;
            this.rmale.Enabled = false;
            this.rmale.Location = new System.Drawing.Point(28, 31);
            this.rmale.Name = "rmale";
            this.rmale.Size = new System.Drawing.Size(48, 17);
            this.rmale.TabIndex = 24;
            this.rmale.TabStop = true;
            this.rmale.Text = "Male";
            this.rmale.UseVisualStyleBackColor = true;
            this.rmale.CheckedChanged += new System.EventHandler(this.radioButton1_CheckedChanged);
            // 
            // rfemale
            // 
            this.rfemale.AutoSize = true;
            this.rfemale.Enabled = false;
            this.rfemale.Location = new System.Drawing.Point(28, 69);
            this.rfemale.Name = "rfemale";
            this.rfemale.Size = new System.Drawing.Size(59, 17);
            this.rfemale.TabIndex = 25;
            this.rfemale.TabStop = true;
            this.rfemale.Text = "Female";
            this.rfemale.UseVisualStyleBackColor = true;
            this.rfemale.CheckedChanged += new System.EventHandler(this.radioButton2_CheckedChanged);
            // 
            // groupBox2
            // 
            this.groupBox2.Controls.Add(this.rfemale);
            this.groupBox2.Controls.Add(this.rmale);
            this.groupBox2.Location = new System.Drawing.Point(634, 201);
            this.groupBox2.Name = "groupBox2";
            this.groupBox2.Size = new System.Drawing.Size(109, 111);
            this.groupBox2.TabIndex = 26;
            this.groupBox2.TabStop = false;
            this.groupBox2.Text = "Gender";
            // 
            // label13
            // 
            this.label13.AutoSize = true;
            this.label13.Location = new System.Drawing.Point(617, 354);
            this.label13.Name = "label13";
            this.label13.Size = new System.Drawing.Size(41, 13);
            this.label13.TabIndex = 27;
            this.label13.Text = "Search";
            // 
            // txtserreg
            // 
            this.txtserreg.Location = new System.Drawing.Point(664, 351);
            this.txtserreg.Name = "txtserreg";
            this.txtserreg.Size = new System.Drawing.Size(154, 20);
            this.txtserreg.TabIndex = 28;
            this.txtserreg.TextChanged += new System.EventHandler(this.txtserreg_TextChanged);
            // 
            // button1
            // 
            this.button1.Enabled = false;
            this.button1.Location = new System.Drawing.Point(36, 338);
            this.button1.Name = "button1";
            this.button1.Size = new System.Drawing.Size(55, 29);
            this.button1.TabIndex = 29;
            this.button1.Text = "Add";
            this.button1.UseVisualStyleBackColor = true;
            this.button1.Click += new System.EventHandler(this.button1_Click);
            // 
            // button2
            // 
            this.button2.Enabled = false;
            this.button2.Location = new System.Drawing.Point(113, 338);
            this.button2.Name = "button2";
            this.button2.Size = new System.Drawing.Size(55, 29);
            this.button2.TabIndex = 30;
            this.button2.Text = "Update";
            this.button2.UseVisualStyleBackColor = true;
            this.button2.Click += new System.EventHandler(this.button2_Click);
            // 
            // rsecond
            // 
            this.rsecond.AutoSize = true;
            this.rsecond.Location = new System.Drawing.Point(341, 19);
            this.rsecond.Name = "rsecond";
            this.rsecond.Size = new System.Drawing.Size(74, 17);
            this.rsecond.TabIndex = 25;
            this.rsecond.TabStop = true;
            this.rsecond.Text = "2nd Batch";
            this.rsecond.UseVisualStyleBackColor = true;
            this.rsecond.CheckedChanged += new System.EventHandler(this.radioButton3_CheckedChanged);
            // 
            // r1st
            // 
            this.r1st.AutoSize = true;
            this.r1st.Location = new System.Drawing.Point(250, 19);
            this.r1st.Name = "r1st";
            this.r1st.Size = new System.Drawing.Size(70, 17);
            this.r1st.TabIndex = 24;
            this.r1st.TabStop = true;
            this.r1st.Text = "1st Batch";
            this.r1st.UseVisualStyleBackColor = true;
            this.r1st.CheckedChanged += new System.EventHandler(this.radioButton4_CheckedChanged);
            // 
            // checkBox1
            // 
            this.checkBox1.AutoSize = true;
            this.checkBox1.Enabled = false;
            this.checkBox1.Location = new System.Drawing.Point(29, 54);
            this.checkBox1.Name = "checkBox1";
            this.checkBox1.Size = new System.Drawing.Size(44, 17);
            this.checkBox1.TabIndex = 34;
            this.checkBox1.Text = "Yes";
            this.checkBox1.UseVisualStyleBackColor = true;
            // 
            // groupBox3
            // 
            this.groupBox3.Controls.Add(this.checkBox1);
            this.groupBox3.Location = new System.Drawing.Point(761, 201);
            this.groupBox3.Name = "groupBox3";
            this.groupBox3.Size = new System.Drawing.Size(99, 111);
            this.groupBox3.TabIndex = 35;
            this.groupBox3.TabStop = false;
            this.groupBox3.Text = "Documents Submitted";
            // 
            // button4
            // 
            this.button4.Location = new System.Drawing.Point(188, 338);
            this.button4.Name = "button4";
            this.button4.Size = new System.Drawing.Size(55, 29);
            this.button4.TabIndex = 36;
            this.button4.Text = "Clear";
            this.button4.UseVisualStyleBackColor = true;
            this.button4.Click += new System.EventHandler(this.button4_Click);
            // 
            // label1
            // 
            this.label1.AutoSize = true;
            this.label1.Location = new System.Drawing.Point(124, 94);
            this.label1.Name = "label1";
            this.label1.Size = new System.Drawing.Size(0, 13);
            this.label1.TabIndex = 37;
            // 
            // txtmis1
            // 
            this.txtmis1.Enabled = false;
            this.txtmis1.Location = new System.Drawing.Point(127, 92);
            this.txtmis1.Name = "txtmis1";
            this.txtmis1.Size = new System.Drawing.Size(160, 20);
            this.txtmis1.TabIndex = 38;
            this.txtmis1.TextChanged += new System.EventHandler(this.txtmis1_TextChanged);
            // 
            // btnpicture
            // 
            this.btnpicture.Location = new System.Drawing.Point(479, 15);
            this.btnpicture.Name = "btnpicture";
            this.btnpicture.Size = new System.Drawing.Size(50, 28);
            this.btnpicture.TabIndex = 32;
            this.btnpicture.Text = "Picture";
            this.btnpicture.UseVisualStyleBackColor = true;
            this.btnpicture.Click += new System.EventHandler(this.btnpicture_Click);
            // 
            // pic
            // 
            this.pic.BackgroundImageLayout = System.Windows.Forms.ImageLayout.Stretch;
            this.pic.BorderStyle = System.Windows.Forms.BorderStyle.FixedSingle;
            this.pic.Location = new System.Drawing.Point(406, 49);
            this.pic.Name = "pic";
            this.pic.Size = new System.Drawing.Size(192, 200);
            this.pic.SizeMode = System.Windows.Forms.PictureBoxSizeMode.StretchImage;
            this.pic.TabIndex = 33;
            this.pic.TabStop = false;
            // 
            // button5
            // 
            this.button5.Image = global::VTA_Interface.Properties.Resources.excel;
            this.button5.ImageAlign = System.Drawing.ContentAlignment.MiddleLeft;
            this.button5.Location = new System.Drawing.Point(271, 338);
            this.button5.Name = "button5";
            this.button5.Size = new System.Drawing.Size(88, 29);
            this.button5.TabIndex = 56;
            this.button5.Text = "EXPORT";
            this.button5.TextAlign = System.Drawing.ContentAlignment.MiddleRight;
            this.button5.UseVisualStyleBackColor = true;
            this.button5.Click += new System.EventHandler(this.button5_Click);
            // 
            // cbatch
            // 
            this.cbatch.FormattingEnabled = true;
            this.cbatch.Items.AddRange(new object[] {
            "1st",
            "2nd"});
            this.cbatch.Location = new System.Drawing.Point(824, 351);
            this.cbatch.Name = "cbatch";
            this.cbatch.Size = new System.Drawing.Size(56, 21);
            this.cbatch.TabIndex = 77;
            // 
            // label5
            // 
            this.label5.AutoSize = true;
            this.label5.Location = new System.Drawing.Point(831, 335);
            this.label5.Name = "label5";
            this.label5.Size = new System.Drawing.Size(35, 13);
            this.label5.TabIndex = 78;
            this.label5.Text = "Batch";
            // 
            // Registration
            // 
            this.AutoScaleDimensions = new System.Drawing.SizeF(6F, 13F);
            this.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font;
            this.ClientSize = new System.Drawing.Size(884, 612);
            this.Controls.Add(this.label5);
            this.Controls.Add(this.cbatch);
            this.Controls.Add(this.button5);
            this.Controls.Add(this.txtmis1);
            this.Controls.Add(this.label1);
            this.Controls.Add(this.button4);
            this.Controls.Add(this.groupBox3);
            this.Controls.Add(this.pic);
            this.Controls.Add(this.btnpicture);
            this.Controls.Add(this.rsecond);
            this.Controls.Add(this.r1st);
            this.Controls.Add(this.button2);
            this.Controls.Add(this.button1);
            this.Controls.Add(this.txtserreg);
            this.Controls.Add(this.label13);
            this.Controls.Add(this.groupBox2);
            this.Controls.Add(this.dataGridView1);
            this.Controls.Add(this.groupBox1);
            this.Controls.Add(this.label10);
            this.Controls.Add(this.label9);
            this.Controls.Add(this.label7);
            this.Controls.Add(this.label6);
            this.Controls.Add(this.label4);
            this.Controls.Add(this.label3);
            this.Controls.Add(this.label2);
            this.Controls.Add(this.txtcontact);
            this.Controls.Add(this.txtaddress);
            this.Controls.Add(this.txtnic);
            this.Controls.Add(this.datepicker);
            this.Controls.Add(this.cyear);
            this.Controls.Add(this.txtstname);
            this.Controls.Add(this.txtmis);
            this.Name = "Registration";
            this.Text = "Registration";
            this.Load += new System.EventHandler(this.Registration_Load);
            this.groupBox1.ResumeLayout(false);
            this.groupBox1.PerformLayout();
            ((System.ComponentModel.ISupportInitialize)(this.dataGridView1)).EndInit();
            this.groupBox2.ResumeLayout(false);
            this.groupBox2.PerformLayout();
            this.groupBox3.ResumeLayout(false);
            this.groupBox3.PerformLayout();
            ((System.ComponentModel.ISupportInitialize)(this.pic)).EndInit();
            this.ResumeLayout(false);
            this.PerformLayout();

        }

        #endregion

        private System.Windows.Forms.TextBox txtmis;
        private System.Windows.Forms.TextBox txtstname;
        private System.Windows.Forms.DateTimePicker datepicker;
        private System.Windows.Forms.TextBox txtnic;
        private System.Windows.Forms.TextBox txtaddress;
        private System.Windows.Forms.TextBox txtcontact;
        private System.Windows.Forms.Label label2;
        private System.Windows.Forms.Label label3;
        private System.Windows.Forms.Label label4;
        private System.Windows.Forms.Label label6;
        private System.Windows.Forms.Label label7;
        private System.Windows.Forms.Label label8;
        private System.Windows.Forms.Label label9;
        private System.Windows.Forms.Label label10;
        private System.Windows.Forms.Label label11;
        private System.Windows.Forms.Label label12;
        private System.Windows.Forms.GroupBox groupBox1;
        private System.Windows.Forms.DataGridView dataGridView1;
        private System.Windows.Forms.RadioButton rmale;
        private System.Windows.Forms.RadioButton rfemale;
        private System.Windows.Forms.GroupBox groupBox2;
        private System.Windows.Forms.Label label13;
        private System.Windows.Forms.TextBox txtserreg;
        private System.Windows.Forms.Button button1;
        private System.Windows.Forms.Button button2;
        private System.Windows.Forms.TextBox txtmath;
        private System.Windows.Forms.TextBox txtscience;
        private System.Windows.Forms.TextBox txtenglish;
        private System.Windows.Forms.RadioButton rsecond;
        private System.Windows.Forms.RadioButton r1st;
        public System.Windows.Forms.ComboBox cyear;
        private System.Windows.Forms.CheckBox checkBox1;
        private System.Windows.Forms.GroupBox groupBox3;
        private System.Windows.Forms.Button button4;
        private System.Windows.Forms.Label label1;
        private System.Windows.Forms.TextBox txtmis1;
        private System.Windows.Forms.Button btnpicture;
        private System.Windows.Forms.PictureBox pic;
        private System.Windows.Forms.Button button5;
        private System.Windows.Forms.ComboBox cbatch;
        private System.Windows.Forms.Label label5;
    }
}