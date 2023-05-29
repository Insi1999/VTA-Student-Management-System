namespace VTA_Interface.Forms
{
    partial class Payments
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
            this.components = new System.ComponentModel.Container();
            this.button2 = new System.Windows.Forms.Button();
            this.button1 = new System.Windows.Forms.Button();
            this.groupBox1 = new System.Windows.Forms.GroupBox();
            this.txtcbtf = new System.Windows.Forms.TextBox();
            this.label5 = new System.Windows.Forms.Label();
            this.label3 = new System.Windows.Forms.Label();
            this.txtddf = new System.Windows.Forms.TextBox();
            this.label10 = new System.Windows.Forms.Label();
            this.txtserpay = new System.Windows.Forms.TextBox();
            this.txtdes = new System.Windows.Forms.TextBox();
            this.txttf = new System.Windows.Forms.TextBox();
            this.label9 = new System.Windows.Forms.Label();
            this.label8 = new System.Windows.Forms.Label();
            this.label7 = new System.Windows.Forms.Label();
            this.label2 = new System.Windows.Forms.Label();
            this.label1 = new System.Windows.Forms.Label();
            this.txtef = new System.Windows.Forms.TextBox();
            this.txtname = new System.Windows.Forms.TextBox();
            this.txtmis = new System.Windows.Forms.TextBox();
            this.dataGridView1 = new System.Windows.Forms.DataGridView();
            this.Column1 = new System.Windows.Forms.DataGridViewTextBoxColumn();
            this.Column2 = new System.Windows.Forms.DataGridViewTextBoxColumn();
            this.Column3 = new System.Windows.Forms.DataGridViewTextBoxColumn();
            this.Column4 = new System.Windows.Forms.DataGridViewTextBoxColumn();
            this.Column5 = new System.Windows.Forms.DataGridViewTextBoxColumn();
            this.EXAM_FEE = new System.Windows.Forms.DataGridViewTextBoxColumn();
            this.Column6 = new System.Windows.Forms.DataGridViewTextBoxColumn();
            this.Column7 = new System.Windows.Forms.DataGridViewTextBoxColumn();
            this.Column8 = new System.Windows.Forms.DataGridViewTextBoxColumn();
            this.lb = new System.Windows.Forms.ListBox();
            this.cyear = new System.Windows.Forms.ComboBox();
            this.label4 = new System.Windows.Forms.Label();
            this.r1 = new System.Windows.Forms.RadioButton();
            this.r2 = new System.Windows.Forms.RadioButton();
            this.button5 = new System.Windows.Forms.Button();
            this.kasthuryDataSet = new VTA_Interface.kasthuryDataSet();
            this.kasthuryDataSetBindingSource = new System.Windows.Forms.BindingSource(this.components);
            this.label6 = new System.Windows.Forms.Label();
            this.rf = new System.Windows.Forms.TextBox();
            this.txtsearch = new System.Windows.Forms.TextBox();
            this.cbatch = new System.Windows.Forms.ComboBox();
            this.button3 = new System.Windows.Forms.Button();
            this.label11 = new System.Windows.Forms.Label();
            this.label12 = new System.Windows.Forms.Label();
            this.groupBox1.SuspendLayout();
            ((System.ComponentModel.ISupportInitialize)(this.dataGridView1)).BeginInit();
            ((System.ComponentModel.ISupportInitialize)(this.kasthuryDataSet)).BeginInit();
            ((System.ComponentModel.ISupportInitialize)(this.kasthuryDataSetBindingSource)).BeginInit();
            this.SuspendLayout();
            // 
            // button2
            // 
            this.button2.Location = new System.Drawing.Point(705, 107);
            this.button2.Name = "button2";
            this.button2.Size = new System.Drawing.Size(102, 48);
            this.button2.TabIndex = 64;
            this.button2.Text = "Update";
            this.button2.UseVisualStyleBackColor = true;
            this.button2.Click += new System.EventHandler(this.button2_Click);
            // 
            // button1
            // 
            this.button1.Location = new System.Drawing.Point(705, 27);
            this.button1.Name = "button1";
            this.button1.Size = new System.Drawing.Size(102, 48);
            this.button1.TabIndex = 63;
            this.button1.Text = "Add";
            this.button1.UseVisualStyleBackColor = true;
            this.button1.Click += new System.EventHandler(this.button1_Click);
            // 
            // groupBox1
            // 
            this.groupBox1.Controls.Add(this.txtcbtf);
            this.groupBox1.Controls.Add(this.label5);
            this.groupBox1.Controls.Add(this.label3);
            this.groupBox1.Controls.Add(this.txtddf);
            this.groupBox1.Location = new System.Drawing.Point(459, 20);
            this.groupBox1.Name = "groupBox1";
            this.groupBox1.Size = new System.Drawing.Size(205, 111);
            this.groupBox1.TabIndex = 62;
            this.groupBox1.TabStop = false;
            this.groupBox1.Text = " Diary Fees";
            // 
            // txtcbtf
            // 
            this.txtcbtf.Location = new System.Drawing.Point(102, 73);
            this.txtcbtf.Name = "txtcbtf";
            this.txtcbtf.Size = new System.Drawing.Size(68, 20);
            this.txtcbtf.TabIndex = 35;
            this.txtcbtf.Text = "0";
            this.txtcbtf.TextChanged += new System.EventHandler(this.txtcbtf_TextChanged);
            this.txtcbtf.KeyPress += new System.Windows.Forms.KeyPressEventHandler(this.txtcbtf_KeyPress);
            // 
            // label5
            // 
            this.label5.AutoSize = true;
            this.label5.Location = new System.Drawing.Point(13, 73);
            this.label5.Name = "label5";
            this.label5.Size = new System.Drawing.Size(55, 13);
            this.label5.TabIndex = 34;
            this.label5.Text = "CBT Diary";
            // 
            // label3
            // 
            this.label3.AutoSize = true;
            this.label3.Location = new System.Drawing.Point(11, 31);
            this.label3.Name = "label3";
            this.label3.Size = new System.Drawing.Size(57, 13);
            this.label3.TabIndex = 33;
            this.label3.Text = "Daily Diary";
            // 
            // txtddf
            // 
            this.txtddf.Location = new System.Drawing.Point(102, 28);
            this.txtddf.Name = "txtddf";
            this.txtddf.Size = new System.Drawing.Size(68, 20);
            this.txtddf.TabIndex = 26;
            this.txtddf.Text = "0";
            this.txtddf.Click += new System.EventHandler(this.txtddf_Click);
            this.txtddf.MouseClick += new System.Windows.Forms.MouseEventHandler(this.txtddf_MouseClick);
            this.txtddf.TextChanged += new System.EventHandler(this.txtddf_TextChanged);
            this.txtddf.Enter += new System.EventHandler(this.txtddf_Enter);
            this.txtddf.KeyPress += new System.Windows.Forms.KeyPressEventHandler(this.txtddf_KeyPress);
            this.txtddf.Leave += new System.EventHandler(this.txtddf_Leave);
            // 
            // label10
            // 
            this.label10.AutoSize = true;
            this.label10.Location = new System.Drawing.Point(473, 343);
            this.label10.Name = "label10";
            this.label10.Size = new System.Drawing.Size(41, 13);
            this.label10.TabIndex = 61;
            this.label10.Text = "Search";
            // 
            // txtserpay
            // 
            this.txtserpay.Location = new System.Drawing.Point(637, 340);
            this.txtserpay.Name = "txtserpay";
            this.txtserpay.Size = new System.Drawing.Size(170, 20);
            this.txtserpay.TabIndex = 60;
            this.txtserpay.TextChanged += new System.EventHandler(this.txtserpay_TextChanged);
            // 
            // txtdes
            // 
            this.txtdes.Enabled = false;
            this.txtdes.Location = new System.Drawing.Point(561, 276);
            this.txtdes.Name = "txtdes";
            this.txtdes.Size = new System.Drawing.Size(69, 20);
            this.txtdes.TabIndex = 59;
            // 
            // txttf
            // 
            this.txttf.Enabled = false;
            this.txttf.Location = new System.Drawing.Point(561, 224);
            this.txttf.Name = "txttf";
            this.txttf.Size = new System.Drawing.Size(69, 20);
            this.txttf.TabIndex = 58;
            this.txttf.Text = "0";
            this.txttf.TextChanged += new System.EventHandler(this.txttf_TextChanged);
            // 
            // label9
            // 
            this.label9.AutoSize = true;
            this.label9.Location = new System.Drawing.Point(473, 279);
            this.label9.Name = "label9";
            this.label9.Size = new System.Drawing.Size(60, 13);
            this.label9.TabIndex = 57;
            this.label9.Text = "Describtion";
            // 
            // label8
            // 
            this.label8.AutoSize = true;
            this.label8.Location = new System.Drawing.Point(473, 224);
            this.label8.Name = "label8";
            this.label8.Size = new System.Drawing.Size(31, 13);
            this.label8.TabIndex = 56;
            this.label8.Text = "Total";
            // 
            // label7
            // 
            this.label7.AutoSize = true;
            this.label7.Location = new System.Drawing.Point(473, 186);
            this.label7.Name = "label7";
            this.label7.Size = new System.Drawing.Size(57, 13);
            this.label7.TabIndex = 55;
            this.label7.Text = "Exam_Fee";
            // 
            // label2
            // 
            this.label2.AutoSize = true;
            this.label2.Location = new System.Drawing.Point(42, 353);
            this.label2.Name = "label2";
            this.label2.Size = new System.Drawing.Size(51, 13);
            this.label2.TabIndex = 53;
            this.label2.Text = "St_Name";
            // 
            // label1
            // 
            this.label1.AutoSize = true;
            this.label1.Location = new System.Drawing.Point(42, 312);
            this.label1.Name = "label1";
            this.label1.Size = new System.Drawing.Size(46, 13);
            this.label1.TabIndex = 52;
            this.label1.Text = "MIS_No";
            // 
            // txtef
            // 
            this.txtef.Location = new System.Drawing.Point(561, 183);
            this.txtef.Name = "txtef";
            this.txtef.Size = new System.Drawing.Size(69, 20);
            this.txtef.TabIndex = 51;
            this.txtef.Text = "0";
            this.txtef.TextChanged += new System.EventHandler(this.txtef_TextChanged);
            this.txtef.KeyPress += new System.Windows.Forms.KeyPressEventHandler(this.txtef_KeyPress);
            // 
            // txtname
            // 
            this.txtname.Enabled = false;
            this.txtname.Location = new System.Drawing.Point(122, 350);
            this.txtname.Name = "txtname";
            this.txtname.Size = new System.Drawing.Size(170, 20);
            this.txtname.TabIndex = 49;
            // 
            // txtmis
            // 
            this.txtmis.Enabled = false;
            this.txtmis.Location = new System.Drawing.Point(122, 309);
            this.txtmis.Name = "txtmis";
            this.txtmis.Size = new System.Drawing.Size(170, 20);
            this.txtmis.TabIndex = 48;
            // 
            // dataGridView1
            // 
            this.dataGridView1.ColumnHeadersHeightSizeMode = System.Windows.Forms.DataGridViewColumnHeadersHeightSizeMode.AutoSize;
            this.dataGridView1.Columns.AddRange(new System.Windows.Forms.DataGridViewColumn[] {
            this.Column1,
            this.Column2,
            this.Column3,
            this.Column4,
            this.Column5,
            this.EXAM_FEE,
            this.Column6,
            this.Column7,
            this.Column8});
            this.dataGridView1.Dock = System.Windows.Forms.DockStyle.Bottom;
            this.dataGridView1.Location = new System.Drawing.Point(0, 390);
            this.dataGridView1.Name = "dataGridView1";
            this.dataGridView1.RowHeadersWidth = 50;
            this.dataGridView1.Size = new System.Drawing.Size(884, 222);
            this.dataGridView1.TabIndex = 66;
            this.dataGridView1.CellContentClick += new System.Windows.Forms.DataGridViewCellEventHandler(this.dataGridView1_CellContentClick);
            this.dataGridView1.RowHeaderMouseClick += new System.Windows.Forms.DataGridViewCellMouseEventHandler(this.dataGridView1_RowHeaderMouseClick);
            // 
            // Column1
            // 
            this.Column1.DataPropertyName = "pid";
            this.Column1.HeaderText = "ID";
            this.Column1.Name = "Column1";
            this.Column1.Width = 20;
            // 
            // Column2
            // 
            this.Column2.DataPropertyName = "mis";
            this.Column2.HeaderText = "MIS";
            this.Column2.Name = "Column2";
            this.Column2.Width = 120;
            // 
            // Column3
            // 
            this.Column3.DataPropertyName = "stname";
            this.Column3.HeaderText = "NAME";
            this.Column3.Name = "Column3";
            // 
            // Column4
            // 
            this.Column4.DataPropertyName = "ddfee";
            this.Column4.HeaderText = "DAILY_DIARY_FEE";
            this.Column4.Name = "Column4";
            this.Column4.Width = 110;
            // 
            // Column5
            // 
            this.Column5.DataPropertyName = "cbtfee";
            this.Column5.HeaderText = "CBT_DIARY_FEE";
            this.Column5.Name = "Column5";
            // 
            // EXAM_FEE
            // 
            this.EXAM_FEE.DataPropertyName = "examfee";
            this.EXAM_FEE.HeaderText = "EXAM_FEE";
            this.EXAM_FEE.Name = "EXAM_FEE";
            // 
            // Column6
            // 
            this.Column6.DataPropertyName = "regfee";
            this.Column6.HeaderText = "REGISTRATION_FEE";
            this.Column6.Name = "Column6";
            this.Column6.Width = 120;
            // 
            // Column7
            // 
            this.Column7.DataPropertyName = "total";
            this.Column7.HeaderText = "TOTAL";
            this.Column7.Name = "Column7";
            this.Column7.Width = 70;
            // 
            // Column8
            // 
            this.Column8.DataPropertyName = "des";
            this.Column8.HeaderText = "DESCRIPTION";
            this.Column8.Name = "Column8";
            // 
            // lb
            // 
            this.lb.Font = new System.Drawing.Font("Microsoft Sans Serif", 9.75F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.lb.FormattingEnabled = true;
            this.lb.ItemHeight = 16;
            this.lb.Location = new System.Drawing.Point(52, 100);
            this.lb.Name = "lb";
            this.lb.Size = new System.Drawing.Size(187, 196);
            this.lb.TabIndex = 67;
            this.lb.SelectedIndexChanged += new System.EventHandler(this.lbnames_SelectedIndexChanged);
            // 
            // cyear
            // 
            this.cyear.FormattingEnabled = true;
            this.cyear.ImeMode = System.Windows.Forms.ImeMode.NoControl;
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
            this.cyear.Location = new System.Drawing.Point(102, 18);
            this.cyear.Name = "cyear";
            this.cyear.Size = new System.Drawing.Size(72, 21);
            this.cyear.TabIndex = 69;
            this.cyear.SelectedIndexChanged += new System.EventHandler(this.comboBox1_SelectedIndexChanged);
            // 
            // label4
            // 
            this.label4.AutoSize = true;
            this.label4.Location = new System.Drawing.Point(49, 18);
            this.label4.Name = "label4";
            this.label4.Size = new System.Drawing.Size(29, 13);
            this.label4.TabIndex = 70;
            this.label4.Text = "Year";
            // 
            // r1
            // 
            this.r1.AutoSize = true;
            this.r1.Location = new System.Drawing.Point(204, 22);
            this.r1.Name = "r1";
            this.r1.Size = new System.Drawing.Size(70, 17);
            this.r1.TabIndex = 71;
            this.r1.TabStop = true;
            this.r1.Text = "1st Batch";
            this.r1.UseVisualStyleBackColor = true;
            this.r1.CheckedChanged += new System.EventHandler(this.r1_CheckedChanged);
            // 
            // r2
            // 
            this.r2.AutoSize = true;
            this.r2.Location = new System.Drawing.Point(280, 22);
            this.r2.Name = "r2";
            this.r2.Size = new System.Drawing.Size(74, 17);
            this.r2.TabIndex = 72;
            this.r2.TabStop = true;
            this.r2.Text = "2nd Batch";
            this.r2.UseVisualStyleBackColor = true;
            this.r2.CheckedChanged += new System.EventHandler(this.r2_CheckedChanged);
            // 
            // button5
            // 
            this.button5.Location = new System.Drawing.Point(706, 186);
            this.button5.Name = "button5";
            this.button5.Size = new System.Drawing.Size(101, 43);
            this.button5.TabIndex = 73;
            this.button5.Text = "Clear All";
            this.button5.UseVisualStyleBackColor = true;
            this.button5.Click += new System.EventHandler(this.button5_Click);
            // 
            // kasthuryDataSet
            // 
            this.kasthuryDataSet.DataSetName = "kasthuryDataSet";
            this.kasthuryDataSet.SchemaSerializationMode = System.Data.SchemaSerializationMode.IncludeSchema;
            // 
            // kasthuryDataSetBindingSource
            // 
            this.kasthuryDataSetBindingSource.DataSource = this.kasthuryDataSet;
            this.kasthuryDataSetBindingSource.Position = 0;
            // 
            // label6
            // 
            this.label6.AutoSize = true;
            this.label6.Location = new System.Drawing.Point(472, 156);
            this.label6.Name = "label6";
            this.label6.Size = new System.Drawing.Size(87, 13);
            this.label6.TabIndex = 74;
            this.label6.Text = "Registration_Fee";
            // 
            // rf
            // 
            this.rf.Location = new System.Drawing.Point(561, 153);
            this.rf.Name = "rf";
            this.rf.Size = new System.Drawing.Size(69, 20);
            this.rf.TabIndex = 75;
            this.rf.Text = "0";
            this.rf.TextChanged += new System.EventHandler(this.textBox1_TextChanged);
            this.rf.KeyPress += new System.Windows.Forms.KeyPressEventHandler(this.rf_KeyPress);
            // 
            // txtsearch
            // 
            this.txtsearch.Location = new System.Drawing.Point(52, 74);
            this.txtsearch.Name = "txtsearch";
            this.txtsearch.Size = new System.Drawing.Size(187, 20);
            this.txtsearch.TabIndex = 76;
            this.txtsearch.TextChanged += new System.EventHandler(this.txtsearch_TextChanged);
            // 
            // cbatch
            // 
            this.cbatch.FormattingEnabled = true;
            this.cbatch.Items.AddRange(new object[] {
            "1st",
            "2nd"});
            this.cbatch.Location = new System.Drawing.Point(561, 340);
            this.cbatch.Name = "cbatch";
            this.cbatch.Size = new System.Drawing.Size(68, 21);
            this.cbatch.TabIndex = 78;
            this.cbatch.SelectedIndexChanged += new System.EventHandler(this.cbatch_SelectedIndexChanged);
            // 
            // button3
            // 
            this.button3.Image = global::VTA_Interface.Properties.Resources.excel;
            this.button3.ImageAlign = System.Drawing.ContentAlignment.MiddleLeft;
            this.button3.Location = new System.Drawing.Point(705, 258);
            this.button3.Name = "button3";
            this.button3.Size = new System.Drawing.Size(102, 34);
            this.button3.TabIndex = 100;
            this.button3.Text = "EXPORT";
            this.button3.TextAlign = System.Drawing.ContentAlignment.MiddleRight;
            this.button3.UseVisualStyleBackColor = true;
            this.button3.Click += new System.EventHandler(this.button3_Click_1);
            // 
            // label11
            // 
            this.label11.AutoSize = true;
            this.label11.Location = new System.Drawing.Point(119, 55);
            this.label11.Name = "label11";
            this.label11.Size = new System.Drawing.Size(41, 13);
            this.label11.TabIndex = 101;
            this.label11.Text = "Search";
            // 
            // label12
            // 
            this.label12.AutoSize = true;
            this.label12.Location = new System.Drawing.Point(575, 324);
            this.label12.Name = "label12";
            this.label12.Size = new System.Drawing.Size(35, 13);
            this.label12.TabIndex = 102;
            this.label12.Text = "Batch";
            // 
            // Payments
            // 
            this.AutoScaleDimensions = new System.Drawing.SizeF(6F, 13F);
            this.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font;
            this.ClientSize = new System.Drawing.Size(884, 612);
            this.Controls.Add(this.label12);
            this.Controls.Add(this.label11);
            this.Controls.Add(this.button3);
            this.Controls.Add(this.cbatch);
            this.Controls.Add(this.txtsearch);
            this.Controls.Add(this.rf);
            this.Controls.Add(this.label6);
            this.Controls.Add(this.button5);
            this.Controls.Add(this.r2);
            this.Controls.Add(this.r1);
            this.Controls.Add(this.label4);
            this.Controls.Add(this.cyear);
            this.Controls.Add(this.lb);
            this.Controls.Add(this.dataGridView1);
            this.Controls.Add(this.button2);
            this.Controls.Add(this.button1);
            this.Controls.Add(this.groupBox1);
            this.Controls.Add(this.label10);
            this.Controls.Add(this.txtserpay);
            this.Controls.Add(this.txtdes);
            this.Controls.Add(this.txttf);
            this.Controls.Add(this.label9);
            this.Controls.Add(this.label8);
            this.Controls.Add(this.label7);
            this.Controls.Add(this.label2);
            this.Controls.Add(this.label1);
            this.Controls.Add(this.txtef);
            this.Controls.Add(this.txtname);
            this.Controls.Add(this.txtmis);
            this.Name = "Payments";
            this.Text = "Payments";
            this.Load += new System.EventHandler(this.Payments_Load);
            this.groupBox1.ResumeLayout(false);
            this.groupBox1.PerformLayout();
            ((System.ComponentModel.ISupportInitialize)(this.dataGridView1)).EndInit();
            ((System.ComponentModel.ISupportInitialize)(this.kasthuryDataSet)).EndInit();
            ((System.ComponentModel.ISupportInitialize)(this.kasthuryDataSetBindingSource)).EndInit();
            this.ResumeLayout(false);
            this.PerformLayout();

        }

        #endregion

        private System.Windows.Forms.Button button2;
        private System.Windows.Forms.Button button1;
        private System.Windows.Forms.GroupBox groupBox1;
        private System.Windows.Forms.Label label3;
        private System.Windows.Forms.TextBox txtddf;
        private System.Windows.Forms.Label label10;
        private System.Windows.Forms.TextBox txtserpay;
        private System.Windows.Forms.TextBox txtdes;
        private System.Windows.Forms.TextBox txttf;
        private System.Windows.Forms.Label label9;
        private System.Windows.Forms.Label label8;
        private System.Windows.Forms.Label label7;
        private System.Windows.Forms.Label label2;
        private System.Windows.Forms.Label label1;
        private System.Windows.Forms.TextBox txtef;
        private System.Windows.Forms.TextBox txtname;
        private System.Windows.Forms.TextBox txtmis;
        private System.Windows.Forms.DataGridView dataGridView1;
        private System.Windows.Forms.ListBox lb;
        private System.Windows.Forms.ComboBox cyear;
        private System.Windows.Forms.Label label4;
        private System.Windows.Forms.RadioButton r1;
        private System.Windows.Forms.RadioButton r2;
        private System.Windows.Forms.TextBox txtcbtf;
        private System.Windows.Forms.Label label5;
        private System.Windows.Forms.Button button5;
        private System.Windows.Forms.BindingSource kasthuryDataSetBindingSource;
        private kasthuryDataSet kasthuryDataSet;
        private System.Windows.Forms.Label label6;
        private System.Windows.Forms.TextBox rf;
        private System.Windows.Forms.DataGridViewTextBoxColumn Column1;
        private System.Windows.Forms.DataGridViewTextBoxColumn Column2;
        private System.Windows.Forms.DataGridViewTextBoxColumn Column3;
        private System.Windows.Forms.DataGridViewTextBoxColumn Column4;
        private System.Windows.Forms.DataGridViewTextBoxColumn Column5;
        private System.Windows.Forms.DataGridViewTextBoxColumn EXAM_FEE;
        private System.Windows.Forms.DataGridViewTextBoxColumn Column6;
        private System.Windows.Forms.DataGridViewTextBoxColumn Column7;
        private System.Windows.Forms.DataGridViewTextBoxColumn Column8;
        private System.Windows.Forms.TextBox txtsearch;
        private System.Windows.Forms.ComboBox cbatch;
        private System.Windows.Forms.Button button3;
        private System.Windows.Forms.Label label11;
        private System.Windows.Forms.Label label12;

    }
}