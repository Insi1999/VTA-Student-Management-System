namespace VTA_Interface
{
    partial class Form1
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
            System.ComponentModel.ComponentResourceManager resources = new System.ComponentModel.ComponentResourceManager(typeof(Form1));
            this.menu = new System.Windows.Forms.Panel();
            this.accounts = new System.Windows.Forms.Button();
            this.button2 = new System.Windows.Forms.Button();
            this.button1 = new System.Windows.Forms.Button();
            this.enq = new System.Windows.Forms.Button();
            this.exam = new System.Windows.Forms.Button();
            this.ojt = new System.Windows.Forms.Button();
            this.mod = new System.Windows.Forms.Button();
            this.pay = new System.Windows.Forms.Button();
            this.regi = new System.Windows.Forms.Button();
            this.logo = new System.Windows.Forms.Panel();
            this.label1 = new System.Windows.Forms.Label();
            this.title = new System.Windows.Forms.Panel();
            this.pictureBox1 = new System.Windows.Forms.PictureBox();
            this.closechildform = new System.Windows.Forms.Button();
            this.lbtitle = new System.Windows.Forms.Label();
            this.desktop = new System.Windows.Forms.Panel();
            this.pictureBox3 = new System.Windows.Forms.PictureBox();
            this.label2 = new System.Windows.Forms.Label();
            this.pictureBox2 = new System.Windows.Forms.PictureBox();
            this.timer1 = new System.Windows.Forms.Timer(this.components);
            this.bunifuElipse1 = new ns1.BunifuElipse(this.components);
            this.menu.SuspendLayout();
            this.logo.SuspendLayout();
            this.title.SuspendLayout();
            ((System.ComponentModel.ISupportInitialize)(this.pictureBox1)).BeginInit();
            this.desktop.SuspendLayout();
            ((System.ComponentModel.ISupportInitialize)(this.pictureBox3)).BeginInit();
            ((System.ComponentModel.ISupportInitialize)(this.pictureBox2)).BeginInit();
            this.SuspendLayout();
            // 
            // menu
            // 
            this.menu.BackColor = System.Drawing.Color.FromArgb(((int)(((byte)(64)))), ((int)(((byte)(64)))), ((int)(((byte)(64)))));
            this.menu.Controls.Add(this.accounts);
            this.menu.Controls.Add(this.button2);
            this.menu.Controls.Add(this.button1);
            this.menu.Controls.Add(this.enq);
            this.menu.Controls.Add(this.exam);
            this.menu.Controls.Add(this.ojt);
            this.menu.Controls.Add(this.mod);
            this.menu.Controls.Add(this.pay);
            this.menu.Controls.Add(this.regi);
            this.menu.Controls.Add(this.logo);
            this.menu.Dock = System.Windows.Forms.DockStyle.Left;
            this.menu.Location = new System.Drawing.Point(0, 0);
            this.menu.Name = "menu";
            this.menu.Size = new System.Drawing.Size(200, 688);
            this.menu.TabIndex = 0;
            // 
            // accounts
            // 
            this.accounts.Dock = System.Windows.Forms.DockStyle.Bottom;
            this.accounts.FlatAppearance.BorderSize = 0;
            this.accounts.FlatStyle = System.Windows.Forms.FlatStyle.Flat;
            this.accounts.Font = new System.Drawing.Font("Microsoft Sans Serif", 9F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.accounts.ForeColor = System.Drawing.Color.White;
            this.accounts.Image = global::VTA_Interface.Properties.Resources.user;
            this.accounts.ImageAlign = System.Drawing.ContentAlignment.MiddleLeft;
            this.accounts.Location = new System.Drawing.Point(0, 628);
            this.accounts.Name = "accounts";
            this.accounts.Padding = new System.Windows.Forms.Padding(12, 0, 0, 0);
            this.accounts.Size = new System.Drawing.Size(200, 60);
            this.accounts.TabIndex = 9;
            this.accounts.Text = "Accounts";
            this.accounts.UseVisualStyleBackColor = true;
            this.accounts.Click += new System.EventHandler(this.accounts_Click);
            // 
            // button2
            // 
            this.button2.Dock = System.Windows.Forms.DockStyle.Top;
            this.button2.FlatAppearance.BorderSize = 0;
            this.button2.FlatStyle = System.Windows.Forms.FlatStyle.Flat;
            this.button2.Font = new System.Drawing.Font("Microsoft Sans Serif", 9F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.button2.ForeColor = System.Drawing.Color.White;
            this.button2.Image = global::VTA_Interface.Properties.Resources.atte;
            this.button2.ImageAlign = System.Drawing.ContentAlignment.MiddleLeft;
            this.button2.Location = new System.Drawing.Point(0, 492);
            this.button2.Name = "button2";
            this.button2.Padding = new System.Windows.Forms.Padding(12, 0, 0, 0);
            this.button2.Size = new System.Drawing.Size(200, 60);
            this.button2.TabIndex = 8;
            this.button2.Text = "Attendance";
            this.button2.UseVisualStyleBackColor = true;
            this.button2.Click += new System.EventHandler(this.button2_Click_1);
            // 
            // button1
            // 
            this.button1.Dock = System.Windows.Forms.DockStyle.Top;
            this.button1.FlatAppearance.BorderSize = 0;
            this.button1.FlatStyle = System.Windows.Forms.FlatStyle.Flat;
            this.button1.Font = new System.Drawing.Font("Microsoft Sans Serif", 9F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.button1.ForeColor = System.Drawing.Color.White;
            this.button1.Image = global::VTA_Interface.Properties.Resources.drop_out__2_;
            this.button1.ImageAlign = System.Drawing.ContentAlignment.MiddleLeft;
            this.button1.Location = new System.Drawing.Point(0, 432);
            this.button1.Name = "button1";
            this.button1.Padding = new System.Windows.Forms.Padding(12, 0, 0, 0);
            this.button1.Size = new System.Drawing.Size(200, 60);
            this.button1.TabIndex = 7;
            this.button1.Text = "Drop Outs";
            this.button1.UseVisualStyleBackColor = true;
            this.button1.Click += new System.EventHandler(this.button1_Click);
            // 
            // enq
            // 
            this.enq.Dock = System.Windows.Forms.DockStyle.Top;
            this.enq.FlatAppearance.BorderSize = 0;
            this.enq.FlatStyle = System.Windows.Forms.FlatStyle.Flat;
            this.enq.Font = new System.Drawing.Font("Microsoft Sans Serif", 9F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.enq.ForeColor = System.Drawing.Color.White;
            this.enq.Image = global::VTA_Interface.Properties.Resources.query;
            this.enq.ImageAlign = System.Drawing.ContentAlignment.MiddleLeft;
            this.enq.Location = new System.Drawing.Point(0, 372);
            this.enq.Name = "enq";
            this.enq.Padding = new System.Windows.Forms.Padding(12, 0, 0, 0);
            this.enq.Size = new System.Drawing.Size(200, 60);
            this.enq.TabIndex = 6;
            this.enq.Text = "Student Inqueries";
            this.enq.UseVisualStyleBackColor = true;
            this.enq.Click += new System.EventHandler(this.enq_Click);
            // 
            // exam
            // 
            this.exam.Dock = System.Windows.Forms.DockStyle.Top;
            this.exam.FlatAppearance.BorderSize = 0;
            this.exam.FlatStyle = System.Windows.Forms.FlatStyle.Flat;
            this.exam.Font = new System.Drawing.Font("Microsoft Sans Serif", 9F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.exam.ForeColor = System.Drawing.Color.White;
            this.exam.Image = global::VTA_Interface.Properties.Resources.exe;
            this.exam.ImageAlign = System.Drawing.ContentAlignment.MiddleLeft;
            this.exam.Location = new System.Drawing.Point(0, 312);
            this.exam.Name = "exam";
            this.exam.Padding = new System.Windows.Forms.Padding(12, 0, 0, 0);
            this.exam.Size = new System.Drawing.Size(200, 60);
            this.exam.TabIndex = 5;
            this.exam.Text = "Final Assesment";
            this.exam.UseVisualStyleBackColor = true;
            this.exam.Click += new System.EventHandler(this.exam_Click);
            // 
            // ojt
            // 
            this.ojt.Dock = System.Windows.Forms.DockStyle.Top;
            this.ojt.FlatAppearance.BorderSize = 0;
            this.ojt.FlatStyle = System.Windows.Forms.FlatStyle.Flat;
            this.ojt.Font = new System.Drawing.Font("Microsoft Sans Serif", 9F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.ojt.ForeColor = System.Drawing.Color.White;
            this.ojt.Image = global::VTA_Interface.Properties.Resources.brief;
            this.ojt.ImageAlign = System.Drawing.ContentAlignment.MiddleLeft;
            this.ojt.Location = new System.Drawing.Point(0, 252);
            this.ojt.Name = "ojt";
            this.ojt.Padding = new System.Windows.Forms.Padding(12, 0, 0, 0);
            this.ojt.Size = new System.Drawing.Size(200, 60);
            this.ojt.TabIndex = 4;
            this.ojt.Text = "OJT Details";
            this.ojt.UseVisualStyleBackColor = true;
            this.ojt.Click += new System.EventHandler(this.ojt_Click);
            // 
            // mod
            // 
            this.mod.Dock = System.Windows.Forms.DockStyle.Top;
            this.mod.FlatAppearance.BorderSize = 0;
            this.mod.FlatStyle = System.Windows.Forms.FlatStyle.Flat;
            this.mod.Font = new System.Drawing.Font("Microsoft Sans Serif", 9F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.mod.ForeColor = System.Drawing.Color.White;
            this.mod.Image = global::VTA_Interface.Properties.Resources.mod;
            this.mod.ImageAlign = System.Drawing.ContentAlignment.MiddleLeft;
            this.mod.Location = new System.Drawing.Point(0, 192);
            this.mod.Name = "mod";
            this.mod.Padding = new System.Windows.Forms.Padding(12, 0, 0, 0);
            this.mod.Size = new System.Drawing.Size(200, 60);
            this.mod.TabIndex = 3;
            this.mod.Text = "Moduels Marks";
            this.mod.UseVisualStyleBackColor = true;
            this.mod.Click += new System.EventHandler(this.mod_Click);
            // 
            // pay
            // 
            this.pay.Dock = System.Windows.Forms.DockStyle.Top;
            this.pay.FlatAppearance.BorderSize = 0;
            this.pay.FlatStyle = System.Windows.Forms.FlatStyle.Flat;
            this.pay.Font = new System.Drawing.Font("Microsoft Sans Serif", 9F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.pay.ForeColor = System.Drawing.Color.White;
            this.pay.Image = global::VTA_Interface.Properties.Resources.wallet_icon;
            this.pay.ImageAlign = System.Drawing.ContentAlignment.MiddleLeft;
            this.pay.Location = new System.Drawing.Point(0, 132);
            this.pay.Name = "pay";
            this.pay.Padding = new System.Windows.Forms.Padding(12, 0, 0, 0);
            this.pay.Size = new System.Drawing.Size(200, 60);
            this.pay.TabIndex = 2;
            this.pay.Text = "Payments";
            this.pay.UseVisualStyleBackColor = true;
            this.pay.Click += new System.EventHandler(this.pay_Click);
            // 
            // regi
            // 
            this.regi.Dock = System.Windows.Forms.DockStyle.Top;
            this.regi.FlatAppearance.BorderSize = 0;
            this.regi.FlatStyle = System.Windows.Forms.FlatStyle.Flat;
            this.regi.Font = new System.Drawing.Font("Microsoft Sans Serif", 9F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.regi.ForeColor = System.Drawing.Color.White;
            this.regi.Image = global::VTA_Interface.Properties.Resources.regi1;
            this.regi.ImageAlign = System.Drawing.ContentAlignment.MiddleLeft;
            this.regi.Location = new System.Drawing.Point(0, 72);
            this.regi.Name = "regi";
            this.regi.Padding = new System.Windows.Forms.Padding(12, 0, 0, 0);
            this.regi.Size = new System.Drawing.Size(200, 60);
            this.regi.TabIndex = 1;
            this.regi.Text = "Registration";
            this.regi.UseVisualStyleBackColor = true;
            this.regi.Click += new System.EventHandler(this.regi_Click);
            // 
            // logo
            // 
            this.logo.BackColor = System.Drawing.Color.Maroon;
            this.logo.Controls.Add(this.label1);
            this.logo.Dock = System.Windows.Forms.DockStyle.Top;
            this.logo.Location = new System.Drawing.Point(0, 0);
            this.logo.Name = "logo";
            this.logo.Size = new System.Drawing.Size(200, 72);
            this.logo.TabIndex = 0;
            // 
            // label1
            // 
            this.label1.AutoSize = true;
            this.label1.Font = new System.Drawing.Font("Microsoft Sans Serif", 14.25F, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.label1.ForeColor = System.Drawing.Color.White;
            this.label1.Location = new System.Drawing.Point(67, 24);
            this.label1.Name = "label1";
            this.label1.Size = new System.Drawing.Size(63, 24);
            this.label1.TabIndex = 0;
            this.label1.Text = "Menu";
            this.label1.TextAlign = System.Drawing.ContentAlignment.MiddleCenter;
            this.label1.Click += new System.EventHandler(this.label1_Click);
            // 
            // title
            // 
            this.title.BackColor = System.Drawing.Color.Brown;
            this.title.Controls.Add(this.pictureBox1);
            this.title.Controls.Add(this.closechildform);
            this.title.Controls.Add(this.lbtitle);
            this.title.Dock = System.Windows.Forms.DockStyle.Top;
            this.title.Location = new System.Drawing.Point(200, 0);
            this.title.Name = "title";
            this.title.Size = new System.Drawing.Size(889, 72);
            this.title.TabIndex = 1;
            this.title.Paint += new System.Windows.Forms.PaintEventHandler(this.title_Paint);
            // 
            // pictureBox1
            // 
            this.pictureBox1.Image = global::VTA_Interface.Properties.Resources.shutdown__1_;
            this.pictureBox1.Location = new System.Drawing.Point(850, 12);
            this.pictureBox1.Name = "pictureBox1";
            this.pictureBox1.Size = new System.Drawing.Size(27, 25);
            this.pictureBox1.TabIndex = 0;
            this.pictureBox1.TabStop = false;
            this.pictureBox1.Click += new System.EventHandler(this.pictureBox1_Click);
            // 
            // closechildform
            // 
            this.closechildform.Dock = System.Windows.Forms.DockStyle.Left;
            this.closechildform.FlatAppearance.BorderSize = 0;
            this.closechildform.FlatStyle = System.Windows.Forms.FlatStyle.Flat;
            this.closechildform.Font = new System.Drawing.Font("Microsoft Sans Serif", 14.25F, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.closechildform.ForeColor = System.Drawing.Color.White;
            this.closechildform.Location = new System.Drawing.Point(0, 0);
            this.closechildform.Name = "closechildform";
            this.closechildform.Size = new System.Drawing.Size(75, 72);
            this.closechildform.TabIndex = 2;
            this.closechildform.Text = "X";
            this.closechildform.UseVisualStyleBackColor = true;
            this.closechildform.Click += new System.EventHandler(this.button2_Click);
            // 
            // lbtitle
            // 
            this.lbtitle.Anchor = System.Windows.Forms.AnchorStyles.None;
            this.lbtitle.AutoSize = true;
            this.lbtitle.Font = new System.Drawing.Font("Microsoft Sans Serif", 20.25F, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.lbtitle.ForeColor = System.Drawing.Color.White;
            this.lbtitle.Location = new System.Drawing.Point(422, 24);
            this.lbtitle.Name = "lbtitle";
            this.lbtitle.Size = new System.Drawing.Size(99, 31);
            this.lbtitle.TabIndex = 0;
            this.lbtitle.Text = "HOME";
            // 
            // desktop
            // 
            this.desktop.Controls.Add(this.pictureBox3);
            this.desktop.Controls.Add(this.label2);
            this.desktop.Controls.Add(this.pictureBox2);
            this.desktop.Dock = System.Windows.Forms.DockStyle.Fill;
            this.desktop.Location = new System.Drawing.Point(200, 72);
            this.desktop.Name = "desktop";
            this.desktop.Size = new System.Drawing.Size(889, 616);
            this.desktop.TabIndex = 2;
            this.desktop.Paint += new System.Windows.Forms.PaintEventHandler(this.desktop_Paint);
            // 
            // pictureBox3
            // 
            this.pictureBox3.Image = global::VTA_Interface.Properties.Resources.newLogo1;
            this.pictureBox3.Location = new System.Drawing.Point(358, 69);
            this.pictureBox3.Name = "pictureBox3";
            this.pictureBox3.Size = new System.Drawing.Size(183, 69);
            this.pictureBox3.SizeMode = System.Windows.Forms.PictureBoxSizeMode.StretchImage;
            this.pictureBox3.TabIndex = 2;
            this.pictureBox3.TabStop = false;
            // 
            // label2
            // 
            this.label2.AutoSize = true;
            this.label2.BackColor = System.Drawing.Color.FromArgb(((int)(((byte)(192)))), ((int)(((byte)(64)))), ((int)(((byte)(0)))));
            this.label2.Font = new System.Drawing.Font("Tahoma", 20.25F, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.label2.ForeColor = System.Drawing.Color.FromArgb(((int)(((byte)(157)))), ((int)(((byte)(27)))), ((int)(((byte)(69)))));
            this.label2.Location = new System.Drawing.Point(223, 69);
            this.label2.Name = "label2";
            this.label2.Size = new System.Drawing.Size(82, 33);
            this.label2.TabIndex = 0;
            this.label2.Text = "Time";
            this.label2.Visible = false;
            this.label2.Click += new System.EventHandler(this.label2_Click);
            // 
            // pictureBox2
            // 
            this.pictureBox2.Image = ((System.Drawing.Image)(resources.GetObject("pictureBox2.Image")));
            this.pictureBox2.Location = new System.Drawing.Point(0, 0);
            this.pictureBox2.Name = "pictureBox2";
            this.pictureBox2.Size = new System.Drawing.Size(889, 616);
            this.pictureBox2.SizeMode = System.Windows.Forms.PictureBoxSizeMode.StretchImage;
            this.pictureBox2.TabIndex = 3;
            this.pictureBox2.TabStop = false;
            this.pictureBox2.Click += new System.EventHandler(this.pictureBox2_Click_1);
            // 
            // timer1
            // 
            this.timer1.Enabled = true;
            this.timer1.Tick += new System.EventHandler(this.timer1_Tick);
            // 
            // bunifuElipse1
            // 
            this.bunifuElipse1.ElipseRadius = 10;
            this.bunifuElipse1.TargetControl = this;
            // 
            // Form1
            // 
            this.AutoScaleDimensions = new System.Drawing.SizeF(6F, 13F);
            this.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font;
            this.ClientSize = new System.Drawing.Size(1089, 688);
            this.Controls.Add(this.desktop);
            this.Controls.Add(this.title);
            this.Controls.Add(this.menu);
            this.FormBorderStyle = System.Windows.Forms.FormBorderStyle.None;
            this.Name = "Form1";
            this.StartPosition = System.Windows.Forms.FormStartPosition.CenterScreen;
            this.Text = "Form1";
            this.Load += new System.EventHandler(this.Form1_Load);
            this.menu.ResumeLayout(false);
            this.logo.ResumeLayout(false);
            this.logo.PerformLayout();
            this.title.ResumeLayout(false);
            this.title.PerformLayout();
            ((System.ComponentModel.ISupportInitialize)(this.pictureBox1)).EndInit();
            this.desktop.ResumeLayout(false);
            this.desktop.PerformLayout();
            ((System.ComponentModel.ISupportInitialize)(this.pictureBox3)).EndInit();
            ((System.ComponentModel.ISupportInitialize)(this.pictureBox2)).EndInit();
            this.ResumeLayout(false);

        }

        #endregion

        private System.Windows.Forms.Panel menu;
        private System.Windows.Forms.Panel logo;
        private System.Windows.Forms.Button regi;
        private System.Windows.Forms.Button enq;
        private System.Windows.Forms.Button exam;
        private System.Windows.Forms.Button ojt;
        private System.Windows.Forms.Button mod;
        private System.Windows.Forms.Button pay;
        private System.Windows.Forms.Panel title;
        private System.Windows.Forms.Label label1;
        private System.Windows.Forms.Label lbtitle;
        private System.Windows.Forms.Panel desktop;
        private System.Windows.Forms.Button button1;
        private System.Windows.Forms.Button closechildform;
        private System.Windows.Forms.Button button2;
        private System.Windows.Forms.Button accounts;
        private System.Windows.Forms.PictureBox pictureBox1;
        private System.Windows.Forms.Timer timer1;
        private ns1.BunifuElipse bunifuElipse1;
        private System.Windows.Forms.PictureBox pictureBox3;
        private System.Windows.Forms.PictureBox pictureBox2;
        private System.Windows.Forms.Label label2;
    }
}

