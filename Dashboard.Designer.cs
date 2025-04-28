namespace Cabaluna
{
    partial class Dashboard
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
            this.pictureBox1 = new System.Windows.Forms.PictureBox();
            this.panel1 = new System.Windows.Forms.Panel();
            this.btnLogout = new System.Windows.Forms.Button();
            this.btnLogs = new System.Windows.Forms.Button();
            this.btnInactive = new System.Windows.Forms.Button();
            this.btnActive = new System.Windows.Forms.Button();
            this.label4 = new System.Windows.Forms.Label();
            this.label3 = new System.Windows.Forms.Label();
            this.label2 = new System.Windows.Forms.Label();
            this.lblName = new System.Windows.Forms.Label();
            this.panel2 = new System.Windows.Forms.Panel();
            this.lblCountActive = new System.Windows.Forms.Label();
            this.label5 = new System.Windows.Forms.Label();
            this.panel3 = new System.Windows.Forms.Panel();
            this.lblCountInactive = new System.Windows.Forms.Label();
            this.label8 = new System.Windows.Forms.Label();
            this.panel4 = new System.Windows.Forms.Panel();
            this.lblCountFemale = new System.Windows.Forms.Label();
            this.label18 = new System.Windows.Forms.Label();
            this.label17 = new System.Windows.Forms.Label();
            this.lblCountMale = new System.Windows.Forms.Label();
            this.label10 = new System.Windows.Forms.Label();
            this.panel5 = new System.Windows.Forms.Panel();
            this.lblCountBSHM = new System.Windows.Forms.Label();
            this.lblCountBSTM = new System.Windows.Forms.Label();
            this.label30 = new System.Windows.Forms.Label();
            this.label29 = new System.Windows.Forms.Label();
            this.label28 = new System.Windows.Forms.Label();
            this.lblCountBSIT = new System.Windows.Forms.Label();
            this.label16 = new System.Windows.Forms.Label();
            this.panel6 = new System.Windows.Forms.Panel();
            this.lblCountSoccer = new System.Windows.Forms.Label();
            this.lblCountVolley = new System.Windows.Forms.Label();
            this.label25 = new System.Windows.Forms.Label();
            this.label24 = new System.Windows.Forms.Label();
            this.label23 = new System.Windows.Forms.Label();
            this.lblCountBasketball = new System.Windows.Forms.Label();
            this.label14 = new System.Windows.Forms.Label();
            this.panel7 = new System.Windows.Forms.Panel();
            this.lblCountBlack = new System.Windows.Forms.Label();
            this.label21 = new System.Windows.Forms.Label();
            this.label20 = new System.Windows.Forms.Label();
            this.lblCountWhite = new System.Windows.Forms.Label();
            this.label12 = new System.Windows.Forms.Label();
            this.button1 = new System.Windows.Forms.Button();
            this.lblDate = new System.Windows.Forms.Label();
            ((System.ComponentModel.ISupportInitialize)(this.pictureBox1)).BeginInit();
            this.panel1.SuspendLayout();
            this.panel2.SuspendLayout();
            this.panel3.SuspendLayout();
            this.panel4.SuspendLayout();
            this.panel5.SuspendLayout();
            this.panel6.SuspendLayout();
            this.panel7.SuspendLayout();
            this.SuspendLayout();
            // 
            // pictureBox1
            // 
            this.pictureBox1.BackColor = System.Drawing.SystemColors.Control;
            this.pictureBox1.Location = new System.Drawing.Point(45, 28);
            this.pictureBox1.Name = "pictureBox1";
            this.pictureBox1.Size = new System.Drawing.Size(100, 79);
            this.pictureBox1.TabIndex = 0;
            this.pictureBox1.TabStop = false;
            // 
            // panel1
            // 
            this.panel1.BackColor = System.Drawing.SystemColors.ActiveCaptionText;
            this.panel1.Controls.Add(this.lblDate);
            this.panel1.Controls.Add(this.btnLogout);
            this.panel1.Controls.Add(this.btnLogs);
            this.panel1.Controls.Add(this.btnInactive);
            this.panel1.Controls.Add(this.btnActive);
            this.panel1.Controls.Add(this.label4);
            this.panel1.Controls.Add(this.label3);
            this.panel1.Controls.Add(this.label2);
            this.panel1.Controls.Add(this.lblName);
            this.panel1.Controls.Add(this.pictureBox1);
            this.panel1.Location = new System.Drawing.Point(0, -2);
            this.panel1.Name = "panel1";
            this.panel1.Size = new System.Drawing.Size(198, 584);
            this.panel1.TabIndex = 1;
            // 
            // btnLogout
            // 
            this.btnLogout.Location = new System.Drawing.Point(45, 470);
            this.btnLogout.Name = "btnLogout";
            this.btnLogout.Size = new System.Drawing.Size(103, 57);
            this.btnLogout.TabIndex = 8;
            this.btnLogout.Text = "Logout";
            this.btnLogout.UseVisualStyleBackColor = true;
            this.btnLogout.Click += new System.EventHandler(this.btnLogout_Click);
            // 
            // btnLogs
            // 
            this.btnLogs.Location = new System.Drawing.Point(45, 388);
            this.btnLogs.Name = "btnLogs";
            this.btnLogs.Size = new System.Drawing.Size(103, 57);
            this.btnLogs.TabIndex = 7;
            this.btnLogs.Text = "Logs";
            this.btnLogs.UseVisualStyleBackColor = true;
            this.btnLogs.Click += new System.EventHandler(this.btnLogs_Click);
            // 
            // btnInactive
            // 
            this.btnInactive.Location = new System.Drawing.Point(45, 309);
            this.btnInactive.Name = "btnInactive";
            this.btnInactive.Size = new System.Drawing.Size(103, 57);
            this.btnInactive.TabIndex = 6;
            this.btnInactive.Text = "Inactive Students";
            this.btnInactive.UseVisualStyleBackColor = true;
            this.btnInactive.Click += new System.EventHandler(this.btnInactive_Click);
            // 
            // btnActive
            // 
            this.btnActive.Location = new System.Drawing.Point(45, 228);
            this.btnActive.Name = "btnActive";
            this.btnActive.Size = new System.Drawing.Size(103, 57);
            this.btnActive.TabIndex = 2;
            this.btnActive.Text = "Active Students";
            this.btnActive.UseVisualStyleBackColor = true;
            this.btnActive.Click += new System.EventHandler(this.btnActive_Click);
            // 
            // label4
            // 
            this.label4.AutoSize = true;
            this.label4.BackColor = System.Drawing.SystemColors.ActiveCaptionText;
            this.label4.ForeColor = System.Drawing.SystemColors.Control;
            this.label4.Location = new System.Drawing.Point(-3, 190);
            this.label4.Name = "label4";
            this.label4.Size = new System.Drawing.Size(205, 13);
            this.label4.TabIndex = 5;
            this.label4.Text = "_________________________________";
            // 
            // label3
            // 
            this.label3.AutoSize = true;
            this.label3.BackColor = System.Drawing.SystemColors.Control;
            this.label3.Font = new System.Drawing.Font("Microsoft Sans Serif", 12F, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.label3.Location = new System.Drawing.Point(-4, 191);
            this.label3.Name = "label3";
            this.label3.Size = new System.Drawing.Size(0, 20);
            this.label3.TabIndex = 4;
            // 
            // label2
            // 
            this.label2.AutoSize = true;
            this.label2.BackColor = System.Drawing.SystemColors.Control;
            this.label2.Font = new System.Drawing.Font("Microsoft Sans Serif", 15.75F, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.label2.Location = new System.Drawing.Point(31, 165);
            this.label2.Name = "label2";
            this.label2.Size = new System.Drawing.Size(126, 25);
            this.label2.TabIndex = 3;
            this.label2.Text = "Dashboard";
            // 
            // lblName
            // 
            this.lblName.AutoSize = true;
            this.lblName.BackColor = System.Drawing.SystemColors.Control;
            this.lblName.Location = new System.Drawing.Point(34, 123);
            this.lblName.Name = "lblName";
            this.lblName.Size = new System.Drawing.Size(0, 13);
            this.lblName.TabIndex = 2;
            // 
            // panel2
            // 
            this.panel2.BackColor = System.Drawing.SystemColors.ActiveCaption;
            this.panel2.Controls.Add(this.lblCountActive);
            this.panel2.Controls.Add(this.label5);
            this.panel2.Location = new System.Drawing.Point(235, 55);
            this.panel2.Name = "panel2";
            this.panel2.Size = new System.Drawing.Size(155, 146);
            this.panel2.TabIndex = 2;
            // 
            // lblCountActive
            // 
            this.lblCountActive.AutoSize = true;
            this.lblCountActive.BackColor = System.Drawing.Color.Transparent;
            this.lblCountActive.ForeColor = System.Drawing.Color.Green;
            this.lblCountActive.Location = new System.Drawing.Point(61, 66);
            this.lblCountActive.Name = "lblCountActive";
            this.lblCountActive.Size = new System.Drawing.Size(25, 13);
            this.lblCountActive.TabIndex = 10;
            this.lblCountActive.Text = "100";
            // 
            // label5
            // 
            this.label5.AutoSize = true;
            this.label5.BackColor = System.Drawing.Color.Transparent;
            this.label5.Font = new System.Drawing.Font("Microsoft Sans Serif", 9.75F, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.label5.Location = new System.Drawing.Point(21, 15);
            this.label5.Name = "label5";
            this.label5.Size = new System.Drawing.Size(114, 16);
            this.label5.TabIndex = 9;
            this.label5.Text = "Active Students";
            // 
            // panel3
            // 
            this.panel3.BackColor = System.Drawing.SystemColors.ActiveCaption;
            this.panel3.Controls.Add(this.lblCountInactive);
            this.panel3.Controls.Add(this.label8);
            this.panel3.Location = new System.Drawing.Point(423, 55);
            this.panel3.Name = "panel3";
            this.panel3.Size = new System.Drawing.Size(155, 146);
            this.panel3.TabIndex = 3;
            // 
            // lblCountInactive
            // 
            this.lblCountInactive.AutoSize = true;
            this.lblCountInactive.BackColor = System.Drawing.Color.Transparent;
            this.lblCountInactive.ForeColor = System.Drawing.Color.Red;
            this.lblCountInactive.Location = new System.Drawing.Point(65, 66);
            this.lblCountInactive.Name = "lblCountInactive";
            this.lblCountInactive.Size = new System.Drawing.Size(19, 13);
            this.lblCountInactive.TabIndex = 12;
            this.lblCountInactive.Text = "20";
            // 
            // label8
            // 
            this.label8.AutoSize = true;
            this.label8.BackColor = System.Drawing.Color.Transparent;
            this.label8.Font = new System.Drawing.Font("Microsoft Sans Serif", 9.75F, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.label8.Location = new System.Drawing.Point(17, 15);
            this.label8.Name = "label8";
            this.label8.Size = new System.Drawing.Size(125, 16);
            this.label8.TabIndex = 11;
            this.label8.Text = "Inactive Students";
            // 
            // panel4
            // 
            this.panel4.BackColor = System.Drawing.SystemColors.ActiveCaption;
            this.panel4.Controls.Add(this.lblCountFemale);
            this.panel4.Controls.Add(this.label18);
            this.panel4.Controls.Add(this.label17);
            this.panel4.Controls.Add(this.lblCountMale);
            this.panel4.Controls.Add(this.label10);
            this.panel4.Location = new System.Drawing.Point(612, 55);
            this.panel4.Name = "panel4";
            this.panel4.Size = new System.Drawing.Size(155, 146);
            this.panel4.TabIndex = 3;
            // 
            // lblCountFemale
            // 
            this.lblCountFemale.AutoSize = true;
            this.lblCountFemale.BackColor = System.Drawing.Color.Transparent;
            this.lblCountFemale.Location = new System.Drawing.Point(104, 86);
            this.lblCountFemale.Name = "lblCountFemale";
            this.lblCountFemale.Size = new System.Drawing.Size(19, 13);
            this.lblCountFemale.TabIndex = 17;
            this.lblCountFemale.Text = "20";
            // 
            // label18
            // 
            this.label18.AutoSize = true;
            this.label18.BackColor = System.Drawing.Color.Transparent;
            this.label18.Location = new System.Drawing.Point(94, 57);
            this.label18.Name = "label18";
            this.label18.Size = new System.Drawing.Size(41, 13);
            this.label18.TabIndex = 16;
            this.label18.Text = "Female";
            // 
            // label17
            // 
            this.label17.AutoSize = true;
            this.label17.BackColor = System.Drawing.Color.Transparent;
            this.label17.Location = new System.Drawing.Point(23, 57);
            this.label17.Name = "label17";
            this.label17.Size = new System.Drawing.Size(30, 13);
            this.label17.TabIndex = 15;
            this.label17.Text = "Male";
            // 
            // lblCountMale
            // 
            this.lblCountMale.AutoSize = true;
            this.lblCountMale.BackColor = System.Drawing.Color.Transparent;
            this.lblCountMale.Location = new System.Drawing.Point(27, 86);
            this.lblCountMale.Name = "lblCountMale";
            this.lblCountMale.Size = new System.Drawing.Size(19, 13);
            this.lblCountMale.TabIndex = 14;
            this.lblCountMale.Text = "20";
            // 
            // label10
            // 
            this.label10.AutoSize = true;
            this.label10.BackColor = System.Drawing.Color.Transparent;
            this.label10.Font = new System.Drawing.Font("Microsoft Sans Serif", 9.75F, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.label10.Location = new System.Drawing.Point(50, 15);
            this.label10.Name = "label10";
            this.label10.Size = new System.Drawing.Size(58, 16);
            this.label10.TabIndex = 13;
            this.label10.Text = "Gender";
            // 
            // panel5
            // 
            this.panel5.BackColor = System.Drawing.SystemColors.ActiveCaption;
            this.panel5.Controls.Add(this.lblCountBSHM);
            this.panel5.Controls.Add(this.lblCountBSTM);
            this.panel5.Controls.Add(this.label30);
            this.panel5.Controls.Add(this.label29);
            this.panel5.Controls.Add(this.label28);
            this.panel5.Controls.Add(this.lblCountBSIT);
            this.panel5.Controls.Add(this.label16);
            this.panel5.Location = new System.Drawing.Point(612, 297);
            this.panel5.Name = "panel5";
            this.panel5.Size = new System.Drawing.Size(155, 146);
            this.panel5.TabIndex = 6;
            // 
            // lblCountBSHM
            // 
            this.lblCountBSHM.AutoSize = true;
            this.lblCountBSHM.BackColor = System.Drawing.Color.Transparent;
            this.lblCountBSHM.Location = new System.Drawing.Point(116, 111);
            this.lblCountBSHM.Name = "lblCountBSHM";
            this.lblCountBSHM.Size = new System.Drawing.Size(19, 13);
            this.lblCountBSHM.TabIndex = 25;
            this.lblCountBSHM.Text = "20";
            // 
            // lblCountBSTM
            // 
            this.lblCountBSTM.AutoSize = true;
            this.lblCountBSTM.BackColor = System.Drawing.Color.Transparent;
            this.lblCountBSTM.Location = new System.Drawing.Point(116, 78);
            this.lblCountBSTM.Name = "lblCountBSTM";
            this.lblCountBSTM.Size = new System.Drawing.Size(19, 13);
            this.lblCountBSTM.TabIndex = 24;
            this.lblCountBSTM.Text = "20";
            // 
            // label30
            // 
            this.label30.AutoSize = true;
            this.label30.BackColor = System.Drawing.Color.Transparent;
            this.label30.Location = new System.Drawing.Point(22, 111);
            this.label30.Name = "label30";
            this.label30.Size = new System.Drawing.Size(38, 13);
            this.label30.TabIndex = 23;
            this.label30.Text = "BSHM";
            // 
            // label29
            // 
            this.label29.AutoSize = true;
            this.label29.BackColor = System.Drawing.Color.Transparent;
            this.label29.Location = new System.Drawing.Point(22, 78);
            this.label29.Name = "label29";
            this.label29.Size = new System.Drawing.Size(37, 13);
            this.label29.TabIndex = 22;
            this.label29.Text = "BSTM";
            // 
            // label28
            // 
            this.label28.AutoSize = true;
            this.label28.BackColor = System.Drawing.Color.Transparent;
            this.label28.Location = new System.Drawing.Point(23, 45);
            this.label28.Name = "label28";
            this.label28.Size = new System.Drawing.Size(31, 13);
            this.label28.TabIndex = 21;
            this.label28.Text = "BSIT";
            // 
            // lblCountBSIT
            // 
            this.lblCountBSIT.AutoSize = true;
            this.lblCountBSIT.BackColor = System.Drawing.Color.Transparent;
            this.lblCountBSIT.Location = new System.Drawing.Point(116, 45);
            this.lblCountBSIT.Name = "lblCountBSIT";
            this.lblCountBSIT.Size = new System.Drawing.Size(19, 13);
            this.lblCountBSIT.TabIndex = 20;
            this.lblCountBSIT.Text = "20";
            // 
            // label16
            // 
            this.label16.AutoSize = true;
            this.label16.BackColor = System.Drawing.Color.Transparent;
            this.label16.Font = new System.Drawing.Font("Microsoft Sans Serif", 9.75F, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.label16.Location = new System.Drawing.Point(54, 13);
            this.label16.Name = "label16";
            this.label16.Size = new System.Drawing.Size(56, 16);
            this.label16.TabIndex = 19;
            this.label16.Text = "Course";
            // 
            // panel6
            // 
            this.panel6.BackColor = System.Drawing.SystemColors.ActiveCaption;
            this.panel6.Controls.Add(this.lblCountSoccer);
            this.panel6.Controls.Add(this.lblCountVolley);
            this.panel6.Controls.Add(this.label25);
            this.panel6.Controls.Add(this.label24);
            this.panel6.Controls.Add(this.label23);
            this.panel6.Controls.Add(this.lblCountBasketball);
            this.panel6.Controls.Add(this.label14);
            this.panel6.Location = new System.Drawing.Point(423, 297);
            this.panel6.Name = "panel6";
            this.panel6.Size = new System.Drawing.Size(155, 146);
            this.panel6.TabIndex = 5;
            // 
            // lblCountSoccer
            // 
            this.lblCountSoccer.AutoSize = true;
            this.lblCountSoccer.BackColor = System.Drawing.Color.Transparent;
            this.lblCountSoccer.Location = new System.Drawing.Point(116, 111);
            this.lblCountSoccer.Name = "lblCountSoccer";
            this.lblCountSoccer.Size = new System.Drawing.Size(19, 13);
            this.lblCountSoccer.TabIndex = 23;
            this.lblCountSoccer.Text = "20";
            // 
            // lblCountVolley
            // 
            this.lblCountVolley.AutoSize = true;
            this.lblCountVolley.BackColor = System.Drawing.Color.Transparent;
            this.lblCountVolley.Location = new System.Drawing.Point(116, 78);
            this.lblCountVolley.Name = "lblCountVolley";
            this.lblCountVolley.Size = new System.Drawing.Size(19, 13);
            this.lblCountVolley.TabIndex = 22;
            this.lblCountVolley.Text = "20";
            // 
            // label25
            // 
            this.label25.AutoSize = true;
            this.label25.BackColor = System.Drawing.Color.Transparent;
            this.label25.Location = new System.Drawing.Point(22, 111);
            this.label25.Name = "label25";
            this.label25.Size = new System.Drawing.Size(41, 13);
            this.label25.TabIndex = 21;
            this.label25.Text = "Soccer";
            // 
            // label24
            // 
            this.label24.AutoSize = true;
            this.label24.BackColor = System.Drawing.Color.Transparent;
            this.label24.Location = new System.Drawing.Point(22, 78);
            this.label24.Name = "label24";
            this.label24.Size = new System.Drawing.Size(51, 13);
            this.label24.TabIndex = 20;
            this.label24.Text = "Volleyball";
            // 
            // label23
            // 
            this.label23.AutoSize = true;
            this.label23.BackColor = System.Drawing.Color.Transparent;
            this.label23.Location = new System.Drawing.Point(22, 45);
            this.label23.Name = "label23";
            this.label23.Size = new System.Drawing.Size(56, 13);
            this.label23.TabIndex = 19;
            this.label23.Text = "Basketball";
            // 
            // lblCountBasketball
            // 
            this.lblCountBasketball.AutoSize = true;
            this.lblCountBasketball.BackColor = System.Drawing.Color.Transparent;
            this.lblCountBasketball.Location = new System.Drawing.Point(116, 45);
            this.lblCountBasketball.Name = "lblCountBasketball";
            this.lblCountBasketball.Size = new System.Drawing.Size(19, 13);
            this.lblCountBasketball.TabIndex = 18;
            this.lblCountBasketball.Text = "20";
            // 
            // label14
            // 
            this.label14.AutoSize = true;
            this.label14.BackColor = System.Drawing.Color.Transparent;
            this.label14.Font = new System.Drawing.Font("Microsoft Sans Serif", 9.75F, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.label14.Location = new System.Drawing.Point(49, 13);
            this.label14.Name = "label14";
            this.label14.Size = new System.Drawing.Size(66, 16);
            this.label14.TabIndex = 17;
            this.label14.Text = "Hobbies";
            // 
            // panel7
            // 
            this.panel7.BackColor = System.Drawing.SystemColors.ActiveCaption;
            this.panel7.Controls.Add(this.lblCountBlack);
            this.panel7.Controls.Add(this.label21);
            this.panel7.Controls.Add(this.label20);
            this.panel7.Controls.Add(this.lblCountWhite);
            this.panel7.Controls.Add(this.label12);
            this.panel7.Location = new System.Drawing.Point(235, 297);
            this.panel7.Name = "panel7";
            this.panel7.Size = new System.Drawing.Size(155, 146);
            this.panel7.TabIndex = 4;
            // 
            // lblCountBlack
            // 
            this.lblCountBlack.AutoSize = true;
            this.lblCountBlack.BackColor = System.Drawing.Color.Transparent;
            this.lblCountBlack.Location = new System.Drawing.Point(99, 89);
            this.lblCountBlack.Name = "lblCountBlack";
            this.lblCountBlack.Size = new System.Drawing.Size(19, 13);
            this.lblCountBlack.TabIndex = 19;
            this.lblCountBlack.Text = "20";
            // 
            // label21
            // 
            this.label21.AutoSize = true;
            this.label21.BackColor = System.Drawing.Color.Transparent;
            this.label21.Location = new System.Drawing.Point(28, 89);
            this.label21.Name = "label21";
            this.label21.Size = new System.Drawing.Size(34, 13);
            this.label21.TabIndex = 18;
            this.label21.Text = "Black";
            // 
            // label20
            // 
            this.label20.AutoSize = true;
            this.label20.BackColor = System.Drawing.Color.Transparent;
            this.label20.Location = new System.Drawing.Point(28, 54);
            this.label20.Name = "label20";
            this.label20.Size = new System.Drawing.Size(35, 13);
            this.label20.TabIndex = 17;
            this.label20.Text = "White";
            // 
            // lblCountWhite
            // 
            this.lblCountWhite.AutoSize = true;
            this.lblCountWhite.BackColor = System.Drawing.Color.Transparent;
            this.lblCountWhite.Location = new System.Drawing.Point(99, 54);
            this.lblCountWhite.Name = "lblCountWhite";
            this.lblCountWhite.Size = new System.Drawing.Size(19, 13);
            this.lblCountWhite.TabIndex = 16;
            this.lblCountWhite.Text = "20";
            // 
            // label12
            // 
            this.label12.AutoSize = true;
            this.label12.BackColor = System.Drawing.Color.Transparent;
            this.label12.Font = new System.Drawing.Font("Microsoft Sans Serif", 9.75F, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.label12.Location = new System.Drawing.Point(55, 17);
            this.label12.Name = "label12";
            this.label12.Size = new System.Drawing.Size(44, 16);
            this.label12.TabIndex = 15;
            this.label12.Text = "Color";
            // 
            // button1
            // 
            this.button1.Location = new System.Drawing.Point(612, 490);
            this.button1.Name = "button1";
            this.button1.Size = new System.Drawing.Size(155, 35);
            this.button1.TabIndex = 9;
            this.button1.Text = "Add New Student";
            this.button1.UseVisualStyleBackColor = true;
            this.button1.Click += new System.EventHandler(this.button1_Click);
            // 
            // lblDate
            // 
            this.lblDate.AutoSize = true;
            this.lblDate.BackColor = System.Drawing.Color.White;
            this.lblDate.ForeColor = System.Drawing.Color.Black;
            this.lblDate.Location = new System.Drawing.Point(81, 143);
            this.lblDate.Name = "lblDate";
            this.lblDate.Size = new System.Drawing.Size(0, 13);
            this.lblDate.TabIndex = 11;
            // 
            // Dashboard
            // 
            this.AutoScaleDimensions = new System.Drawing.SizeF(6F, 13F);
            this.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font;
            this.BackColor = System.Drawing.Color.DarkSlateGray;
            this.ClientSize = new System.Drawing.Size(800, 580);
            this.Controls.Add(this.button1);
            this.Controls.Add(this.panel5);
            this.Controls.Add(this.panel6);
            this.Controls.Add(this.panel7);
            this.Controls.Add(this.panel4);
            this.Controls.Add(this.panel3);
            this.Controls.Add(this.panel2);
            this.Controls.Add(this.panel1);
            this.Name = "Dashboard";
            this.StartPosition = System.Windows.Forms.FormStartPosition.CenterScreen;
            this.Text = "Dashboard";
            ((System.ComponentModel.ISupportInitialize)(this.pictureBox1)).EndInit();
            this.panel1.ResumeLayout(false);
            this.panel1.PerformLayout();
            this.panel2.ResumeLayout(false);
            this.panel2.PerformLayout();
            this.panel3.ResumeLayout(false);
            this.panel3.PerformLayout();
            this.panel4.ResumeLayout(false);
            this.panel4.PerformLayout();
            this.panel5.ResumeLayout(false);
            this.panel5.PerformLayout();
            this.panel6.ResumeLayout(false);
            this.panel6.PerformLayout();
            this.panel7.ResumeLayout(false);
            this.panel7.PerformLayout();
            this.ResumeLayout(false);

        }

        #endregion

        private System.Windows.Forms.PictureBox pictureBox1;
        private System.Windows.Forms.Panel panel1;
        private System.Windows.Forms.Label label2;
        private System.Windows.Forms.Label label4;
        private System.Windows.Forms.Label label3;
        private System.Windows.Forms.Button btnLogout;
        private System.Windows.Forms.Button btnLogs;
        private System.Windows.Forms.Button btnInactive;
        private System.Windows.Forms.Button btnActive;
        private System.Windows.Forms.Panel panel2;
        private System.Windows.Forms.Label label5;
        private System.Windows.Forms.Panel panel3;
        private System.Windows.Forms.Panel panel4;
        private System.Windows.Forms.Panel panel5;
        private System.Windows.Forms.Panel panel6;
        private System.Windows.Forms.Panel panel7;
        private System.Windows.Forms.Label lblCountActive;
        private System.Windows.Forms.Label lblCountInactive;
        private System.Windows.Forms.Label label8;
        private System.Windows.Forms.Label lblCountMale;
        private System.Windows.Forms.Label label10;
        private System.Windows.Forms.Label lblCountBSIT;
        private System.Windows.Forms.Label label16;
        private System.Windows.Forms.Label lblCountBasketball;
        private System.Windows.Forms.Label label14;
        private System.Windows.Forms.Label lblCountWhite;
        private System.Windows.Forms.Label label12;
        private System.Windows.Forms.Label lblCountFemale;
        private System.Windows.Forms.Label label18;
        private System.Windows.Forms.Label label17;
        private System.Windows.Forms.Label label25;
        private System.Windows.Forms.Label label24;
        private System.Windows.Forms.Label label23;
        private System.Windows.Forms.Label lblCountBlack;
        private System.Windows.Forms.Label label21;
        private System.Windows.Forms.Label label20;
        private System.Windows.Forms.Label lblCountBSHM;
        private System.Windows.Forms.Label lblCountBSTM;
        private System.Windows.Forms.Label label30;
        private System.Windows.Forms.Label label29;
        private System.Windows.Forms.Label label28;
        private System.Windows.Forms.Label lblCountSoccer;
        private System.Windows.Forms.Label lblCountVolley;
        public System.Windows.Forms.Label lblName;
        private System.Windows.Forms.Button button1;
        public System.Windows.Forms.Label lblDate;
    }
}