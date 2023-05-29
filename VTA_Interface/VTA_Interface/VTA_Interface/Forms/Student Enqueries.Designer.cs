namespace VTA_Interface.Forms
{
    partial class Student_Enqueries
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
            this.dataGridView1 = new System.Windows.Forms.DataGridView();
            this.label1 = new System.Windows.Forms.Label();
            this.txtsearch = new System.Windows.Forms.TextBox();
            this.lstname = new System.Windows.Forms.Label();
            this.button1 = new System.Windows.Forms.Button();
            this.cbatch = new System.Windows.Forms.ComboBox();
            this.label2 = new System.Windows.Forms.Label();
            ((System.ComponentModel.ISupportInitialize)(this.dataGridView1)).BeginInit();
            this.SuspendLayout();
            // 
            // dataGridView1
            // 
            this.dataGridView1.ColumnHeadersHeightSizeMode = System.Windows.Forms.DataGridViewColumnHeadersHeightSizeMode.AutoSize;
            this.dataGridView1.Dock = System.Windows.Forms.DockStyle.Bottom;
            this.dataGridView1.Location = new System.Drawing.Point(0, 384);
            this.dataGridView1.Name = "dataGridView1";
            this.dataGridView1.RowHeadersWidth = 61;
            this.dataGridView1.Size = new System.Drawing.Size(884, 228);
            this.dataGridView1.TabIndex = 50;
            // 
            // label1
            // 
            this.label1.AutoSize = true;
            this.label1.Font = new System.Drawing.Font("Tahoma", 27.75F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.label1.Location = new System.Drawing.Point(362, 29);
            this.label1.Name = "label1";
            this.label1.Size = new System.Drawing.Size(155, 45);
            this.label1.TabIndex = 51;
            this.label1.Text = "REPORT";
            // 
            // txtsearch
            // 
            this.txtsearch.Location = new System.Drawing.Point(378, 247);
            this.txtsearch.Name = "txtsearch";
            this.txtsearch.Size = new System.Drawing.Size(122, 20);
            this.txtsearch.TabIndex = 52;
            this.txtsearch.TextChanged += new System.EventHandler(this.textBox1_TextChanged);
            // 
            // lstname
            // 
            this.lstname.AutoSize = true;
            this.lstname.Font = new System.Drawing.Font("Microsoft Sans Serif", 14.25F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.lstname.Location = new System.Drawing.Point(411, 100);
            this.lstname.Name = "lstname";
            this.lstname.Size = new System.Drawing.Size(57, 24);
            this.lstname.TabIndex = 54;
            this.lstname.Text = "Batch";
            // 
            // button1
            // 
            this.button1.Image = global::VTA_Interface.Properties.Resources.excel;
            this.button1.ImageAlign = System.Drawing.ContentAlignment.MiddleLeft;
            this.button1.Location = new System.Drawing.Point(391, 293);
            this.button1.Name = "button1";
            this.button1.Size = new System.Drawing.Size(97, 54);
            this.button1.TabIndex = 55;
            this.button1.Text = "EXPORT";
            this.button1.TextAlign = System.Drawing.ContentAlignment.MiddleRight;
            this.button1.UseVisualStyleBackColor = true;
            this.button1.Click += new System.EventHandler(this.button1_Click);
            // 
            // cbatch
            // 
            this.cbatch.FormattingEnabled = true;
            this.cbatch.Items.AddRange(new object[] {
            "1st",
            "2nd"});
            this.cbatch.Location = new System.Drawing.Point(379, 150);
            this.cbatch.Name = "cbatch";
            this.cbatch.Size = new System.Drawing.Size(121, 21);
            this.cbatch.TabIndex = 57;
            this.cbatch.SelectedIndexChanged += new System.EventHandler(this.comboBox1_SelectedIndexChanged);
            // 
            // label2
            // 
            this.label2.AutoSize = true;
            this.label2.Font = new System.Drawing.Font("Microsoft Sans Serif", 14.25F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.label2.Location = new System.Drawing.Point(404, 197);
            this.label2.Name = "label2";
            this.label2.Size = new System.Drawing.Size(70, 24);
            this.label2.TabIndex = 58;
            this.label2.Text = "Search";
            // 
            // Student_Enqueries
            // 
            this.AutoScaleDimensions = new System.Drawing.SizeF(6F, 13F);
            this.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font;
            this.ClientSize = new System.Drawing.Size(884, 612);
            this.Controls.Add(this.label2);
            this.Controls.Add(this.cbatch);
            this.Controls.Add(this.button1);
            this.Controls.Add(this.lstname);
            this.Controls.Add(this.txtsearch);
            this.Controls.Add(this.label1);
            this.Controls.Add(this.dataGridView1);
            this.Name = "Student_Enqueries";
            this.Text = "Student_Enqueries";
            ((System.ComponentModel.ISupportInitialize)(this.dataGridView1)).EndInit();
            this.ResumeLayout(false);
            this.PerformLayout();

        }

        #endregion

        private System.Windows.Forms.DataGridView dataGridView1;
        private System.Windows.Forms.Label label1;
        private System.Windows.Forms.TextBox txtsearch;
        private System.Windows.Forms.Label lstname;
        private System.Windows.Forms.Button button1;
        private System.Windows.Forms.ComboBox cbatch;
        private System.Windows.Forms.Label label2;
    }
}