namespace Cabaluna
{
    partial class FormBase
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
            this.lblName = new System.Windows.Forms.Label();
            this.lblGender = new System.Windows.Forms.Label();
            this.lblHobbies = new System.Windows.Forms.Label();
            this.lblFavoriteColor = new System.Windows.Forms.Label();
            this.lblSaying = new System.Windows.Forms.Label();
            this.txtName = new System.Windows.Forms.TextBox();
            this.radMale = new System.Windows.Forms.RadioButton();
            this.radFemale = new System.Windows.Forms.RadioButton();
            this.chkBasketball = new System.Windows.Forms.CheckBox();
            this.chkVolleyball = new System.Windows.Forms.CheckBox();
            this.chkSoccer = new System.Windows.Forms.CheckBox();
            this.cmbFaveColor = new System.Windows.Forms.ComboBox();
            this.txtSaying = new System.Windows.Forms.TextBox();
            this.btnAdd = new System.Windows.Forms.Button();
            this.btnDisplay = new System.Windows.Forms.Button();
            this.btnClear = new System.Windows.Forms.Button();
            this.btnUpdate = new System.Windows.Forms.Button();
            this.lbliD = new System.Windows.Forms.Label();
            this.txtAddress = new System.Windows.Forms.TextBox();
            this.txtEmail = new System.Windows.Forms.TextBox();
            this.lblAddress = new System.Windows.Forms.Label();
            this.lblEmail = new System.Windows.Forms.Label();
            this.label1 = new System.Windows.Forms.Label();
            this.dtpDate = new System.Windows.Forms.DateTimePicker();
            this.lblAge = new System.Windows.Forms.Label();
            this.txtAge = new System.Windows.Forms.TextBox();
            this.lblUsername = new System.Windows.Forms.Label();
            this.lblPassword = new System.Windows.Forms.Label();
            this.txtUsername = new System.Windows.Forms.TextBox();
            this.txtPassword = new System.Windows.Forms.TextBox();
            this.lblIdentification = new System.Windows.Forms.Label();
            this.lblCourse = new System.Windows.Forms.Label();
            this.cmbCourse = new System.Windows.Forms.ComboBox();
            this.SuspendLayout();
            // 
            // lblName
            // 
            this.lblName.AutoSize = true;
            this.lblName.Font = new System.Drawing.Font("Microsoft Sans Serif", 12F, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.lblName.Location = new System.Drawing.Point(30, 32);
            this.lblName.Name = "lblName";
            this.lblName.Size = new System.Drawing.Size(60, 20);
            this.lblName.TabIndex = 0;
            this.lblName.Text = "Name:";
            // 
            // lblGender
            // 
            this.lblGender.AutoSize = true;
            this.lblGender.Font = new System.Drawing.Font("Microsoft Sans Serif", 12F, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.lblGender.Location = new System.Drawing.Point(30, 86);
            this.lblGender.Name = "lblGender";
            this.lblGender.Size = new System.Drawing.Size(74, 20);
            this.lblGender.TabIndex = 1;
            this.lblGender.Text = "Gender:";
            // 
            // lblHobbies
            // 
            this.lblHobbies.AutoSize = true;
            this.lblHobbies.Font = new System.Drawing.Font("Microsoft Sans Serif", 12F, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.lblHobbies.Location = new System.Drawing.Point(30, 132);
            this.lblHobbies.Name = "lblHobbies";
            this.lblHobbies.Size = new System.Drawing.Size(80, 20);
            this.lblHobbies.TabIndex = 2;
            this.lblHobbies.Text = "Hobbies:";
            // 
            // lblFavoriteColor
            // 
            this.lblFavoriteColor.AutoSize = true;
            this.lblFavoriteColor.Font = new System.Drawing.Font("Microsoft Sans Serif", 12F, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.lblFavoriteColor.Location = new System.Drawing.Point(346, 232);
            this.lblFavoriteColor.Name = "lblFavoriteColor";
            this.lblFavoriteColor.Size = new System.Drawing.Size(126, 20);
            this.lblFavoriteColor.TabIndex = 3;
            this.lblFavoriteColor.Text = "Favorite Color:";
            // 
            // lblSaying
            // 
            this.lblSaying.AutoSize = true;
            this.lblSaying.Font = new System.Drawing.Font("Microsoft Sans Serif", 12F, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.lblSaying.Location = new System.Drawing.Point(346, 275);
            this.lblSaying.Name = "lblSaying";
            this.lblSaying.Size = new System.Drawing.Size(68, 20);
            this.lblSaying.TabIndex = 4;
            this.lblSaying.Text = "Saying:";
            // 
            // txtName
            // 
            this.txtName.Location = new System.Drawing.Point(120, 34);
            this.txtName.Name = "txtName";
            this.txtName.Size = new System.Drawing.Size(180, 20);
            this.txtName.TabIndex = 5;
            this.txtName.KeyPress += new System.Windows.Forms.KeyPressEventHandler(this.txtName_KeyPress);
            // 
            // radMale
            // 
            this.radMale.AutoSize = true;
            this.radMale.Location = new System.Drawing.Point(134, 89);
            this.radMale.Name = "radMale";
            this.radMale.Size = new System.Drawing.Size(48, 17);
            this.radMale.TabIndex = 6;
            this.radMale.TabStop = true;
            this.radMale.Text = "Male";
            this.radMale.UseVisualStyleBackColor = true;
            // 
            // radFemale
            // 
            this.radFemale.AutoSize = true;
            this.radFemale.Location = new System.Drawing.Point(241, 89);
            this.radFemale.Name = "radFemale";
            this.radFemale.Size = new System.Drawing.Size(59, 17);
            this.radFemale.TabIndex = 7;
            this.radFemale.TabStop = true;
            this.radFemale.Text = "Female";
            this.radFemale.UseVisualStyleBackColor = true;
            // 
            // chkBasketball
            // 
            this.chkBasketball.AutoSize = true;
            this.chkBasketball.Location = new System.Drawing.Point(120, 136);
            this.chkBasketball.Name = "chkBasketball";
            this.chkBasketball.Size = new System.Drawing.Size(75, 17);
            this.chkBasketball.TabIndex = 8;
            this.chkBasketball.Text = "Basketball";
            this.chkBasketball.UseVisualStyleBackColor = true;
            // 
            // chkVolleyball
            // 
            this.chkVolleyball.AutoSize = true;
            this.chkVolleyball.Location = new System.Drawing.Point(120, 171);
            this.chkVolleyball.Name = "chkVolleyball";
            this.chkVolleyball.Size = new System.Drawing.Size(70, 17);
            this.chkVolleyball.TabIndex = 9;
            this.chkVolleyball.Text = "Volleyball";
            this.chkVolleyball.UseVisualStyleBackColor = true;
            // 
            // chkSoccer
            // 
            this.chkSoccer.AutoSize = true;
            this.chkSoccer.Location = new System.Drawing.Point(120, 206);
            this.chkSoccer.Name = "chkSoccer";
            this.chkSoccer.Size = new System.Drawing.Size(60, 17);
            this.chkSoccer.TabIndex = 10;
            this.chkSoccer.Text = "Soccer";
            this.chkSoccer.UseVisualStyleBackColor = true;
            // 
            // cmbFaveColor
            // 
            this.cmbFaveColor.DropDownStyle = System.Windows.Forms.ComboBoxStyle.DropDownList;
            this.cmbFaveColor.FormattingEnabled = true;
            this.cmbFaveColor.Items.AddRange(new object[] {
            "Black",
            "White"});
            this.cmbFaveColor.Location = new System.Drawing.Point(505, 231);
            this.cmbFaveColor.Name = "cmbFaveColor";
            this.cmbFaveColor.Size = new System.Drawing.Size(121, 21);
            this.cmbFaveColor.TabIndex = 11;
            // 
            // txtSaying
            // 
            this.txtSaying.Location = new System.Drawing.Point(446, 275);
            this.txtSaying.Multiline = true;
            this.txtSaying.Name = "txtSaying";
            this.txtSaying.Size = new System.Drawing.Size(180, 127);
            this.txtSaying.TabIndex = 12;
            // 
            // btnAdd
            // 
            this.btnAdd.Location = new System.Drawing.Point(350, 408);
            this.btnAdd.Name = "btnAdd";
            this.btnAdd.Size = new System.Drawing.Size(75, 23);
            this.btnAdd.TabIndex = 13;
            this.btnAdd.Text = "Add";
            this.btnAdd.UseVisualStyleBackColor = true;
            this.btnAdd.Click += new System.EventHandler(this.btnAdd_Click);
            // 
            // btnDisplay
            // 
            this.btnDisplay.Location = new System.Drawing.Point(448, 408);
            this.btnDisplay.Name = "btnDisplay";
            this.btnDisplay.Size = new System.Drawing.Size(75, 23);
            this.btnDisplay.TabIndex = 14;
            this.btnDisplay.Text = "Display";
            this.btnDisplay.UseVisualStyleBackColor = true;
            this.btnDisplay.Click += new System.EventHandler(this.btnDisplay_Click);
            // 
            // btnClear
            // 
            this.btnClear.Location = new System.Drawing.Point(551, 408);
            this.btnClear.Name = "btnClear";
            this.btnClear.Size = new System.Drawing.Size(75, 23);
            this.btnClear.TabIndex = 16;
            this.btnClear.Text = "Clear";
            this.btnClear.UseVisualStyleBackColor = true;
            // 
            // btnUpdate
            // 
            this.btnUpdate.Location = new System.Drawing.Point(350, 408);
            this.btnUpdate.Name = "btnUpdate";
            this.btnUpdate.Size = new System.Drawing.Size(75, 23);
            this.btnUpdate.TabIndex = 17;
            this.btnUpdate.Text = "Update";
            this.btnUpdate.UseVisualStyleBackColor = true;
            this.btnUpdate.Visible = false;
            this.btnUpdate.Click += new System.EventHandler(this.btnUpdate_Click);
            // 
            // lbliD
            // 
            this.lbliD.AutoSize = true;
            this.lbliD.Font = new System.Drawing.Font("Microsoft Sans Serif", 12F, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.lbliD.Location = new System.Drawing.Point(43, 128);
            this.lbliD.Name = "lbliD";
            this.lbliD.Size = new System.Drawing.Size(0, 20);
            this.lbliD.TabIndex = 18;
            this.lbliD.Visible = false;
            // 
            // txtAddress
            // 
            this.txtAddress.Location = new System.Drawing.Point(120, 280);
            this.txtAddress.Name = "txtAddress";
            this.txtAddress.Size = new System.Drawing.Size(180, 20);
            this.txtAddress.TabIndex = 19;
            // 
            // txtEmail
            // 
            this.txtEmail.Location = new System.Drawing.Point(120, 341);
            this.txtEmail.Name = "txtEmail";
            this.txtEmail.Size = new System.Drawing.Size(180, 20);
            this.txtEmail.TabIndex = 20;
            // 
            // lblAddress
            // 
            this.lblAddress.AutoSize = true;
            this.lblAddress.Font = new System.Drawing.Font("Microsoft Sans Serif", 12F, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.lblAddress.Location = new System.Drawing.Point(30, 280);
            this.lblAddress.Name = "lblAddress";
            this.lblAddress.Size = new System.Drawing.Size(80, 20);
            this.lblAddress.TabIndex = 21;
            this.lblAddress.Text = "Address:";
            // 
            // lblEmail
            // 
            this.lblEmail.AutoSize = true;
            this.lblEmail.Font = new System.Drawing.Font("Microsoft Sans Serif", 12F, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.lblEmail.Location = new System.Drawing.Point(30, 339);
            this.lblEmail.Name = "lblEmail";
            this.lblEmail.Size = new System.Drawing.Size(58, 20);
            this.lblEmail.TabIndex = 22;
            this.lblEmail.Text = "Email:";
            // 
            // label1
            // 
            this.label1.AutoSize = true;
            this.label1.Font = new System.Drawing.Font("Microsoft Sans Serif", 12F, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.label1.Location = new System.Drawing.Point(30, 411);
            this.label1.Name = "label1";
            this.label1.Size = new System.Drawing.Size(80, 20);
            this.label1.TabIndex = 23;
            this.label1.Text = "Birthday:";
            // 
            // dtpDate
            // 
            this.dtpDate.Location = new System.Drawing.Point(120, 411);
            this.dtpDate.Name = "dtpDate";
            this.dtpDate.Size = new System.Drawing.Size(200, 20);
            this.dtpDate.TabIndex = 24;
            this.dtpDate.ValueChanged += new System.EventHandler(this.dtpDate_ValueChanged);
            // 
            // lblAge
            // 
            this.lblAge.AutoSize = true;
            this.lblAge.Font = new System.Drawing.Font("Microsoft Sans Serif", 12F, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.lblAge.Location = new System.Drawing.Point(346, 27);
            this.lblAge.Name = "lblAge";
            this.lblAge.Size = new System.Drawing.Size(46, 20);
            this.lblAge.TabIndex = 25;
            this.lblAge.Text = "Age:";
            // 
            // txtAge
            // 
            this.txtAge.Location = new System.Drawing.Point(446, 29);
            this.txtAge.Name = "txtAge";
            this.txtAge.ReadOnly = true;
            this.txtAge.Size = new System.Drawing.Size(180, 20);
            this.txtAge.TabIndex = 26;
            // 
            // lblUsername
            // 
            this.lblUsername.AutoSize = true;
            this.lblUsername.Font = new System.Drawing.Font("Microsoft Sans Serif", 12F, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.lblUsername.Location = new System.Drawing.Point(346, 104);
            this.lblUsername.Name = "lblUsername";
            this.lblUsername.Size = new System.Drawing.Size(96, 20);
            this.lblUsername.TabIndex = 27;
            this.lblUsername.Text = "Username:";
            // 
            // lblPassword
            // 
            this.lblPassword.AutoSize = true;
            this.lblPassword.Font = new System.Drawing.Font("Microsoft Sans Serif", 12F, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.lblPassword.Location = new System.Drawing.Point(346, 149);
            this.lblPassword.Name = "lblPassword";
            this.lblPassword.Size = new System.Drawing.Size(91, 20);
            this.lblPassword.TabIndex = 28;
            this.lblPassword.Text = "Password:";
            // 
            // txtUsername
            // 
            this.txtUsername.Location = new System.Drawing.Point(446, 106);
            this.txtUsername.Name = "txtUsername";
            this.txtUsername.ReadOnly = true;
            this.txtUsername.Size = new System.Drawing.Size(180, 20);
            this.txtUsername.TabIndex = 29;
            // 
            // txtPassword
            // 
            this.txtPassword.Location = new System.Drawing.Point(446, 149);
            this.txtPassword.Name = "txtPassword";
            this.txtPassword.ReadOnly = true;
            this.txtPassword.Size = new System.Drawing.Size(180, 20);
            this.txtPassword.TabIndex = 30;
            // 
            // lblIdentification
            // 
            this.lblIdentification.AutoSize = true;
            this.lblIdentification.Font = new System.Drawing.Font("Microsoft Sans Serif", 12F, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.lblIdentification.Location = new System.Drawing.Point(28, 229);
            this.lblIdentification.Name = "lblIdentification";
            this.lblIdentification.Size = new System.Drawing.Size(0, 20);
            this.lblIdentification.TabIndex = 31;
            // 
            // lblCourse
            // 
            this.lblCourse.AutoSize = true;
            this.lblCourse.Font = new System.Drawing.Font("Microsoft Sans Serif", 12F, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.lblCourse.Location = new System.Drawing.Point(346, 193);
            this.lblCourse.Name = "lblCourse";
            this.lblCourse.Size = new System.Drawing.Size(71, 20);
            this.lblCourse.TabIndex = 32;
            this.lblCourse.Text = "Course:";
            // 
            // cmbCourse
            // 
            this.cmbCourse.DropDownStyle = System.Windows.Forms.ComboBoxStyle.DropDownList;
            this.cmbCourse.FormattingEnabled = true;
            this.cmbCourse.Items.AddRange(new object[] {
            "BSIT",
            "BSHM",
            "BSCPE",
            "BSTM",
            "BSN",
            "BSED",
            "BEED"});
            this.cmbCourse.Location = new System.Drawing.Point(505, 195);
            this.cmbCourse.Name = "cmbCourse";
            this.cmbCourse.Size = new System.Drawing.Size(121, 21);
            this.cmbCourse.TabIndex = 33;
            // 
            // FormBase
            // 
            this.AutoScaleDimensions = new System.Drawing.SizeF(6F, 13F);
            this.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font;
            this.ClientSize = new System.Drawing.Size(675, 481);
            this.Controls.Add(this.cmbCourse);
            this.Controls.Add(this.lblCourse);
            this.Controls.Add(this.lblIdentification);
            this.Controls.Add(this.txtPassword);
            this.Controls.Add(this.txtUsername);
            this.Controls.Add(this.lblPassword);
            this.Controls.Add(this.lblUsername);
            this.Controls.Add(this.txtAge);
            this.Controls.Add(this.lblAge);
            this.Controls.Add(this.dtpDate);
            this.Controls.Add(this.label1);
            this.Controls.Add(this.lblEmail);
            this.Controls.Add(this.lblAddress);
            this.Controls.Add(this.txtEmail);
            this.Controls.Add(this.txtAddress);
            this.Controls.Add(this.lbliD);
            this.Controls.Add(this.btnClear);
            this.Controls.Add(this.btnDisplay);
            this.Controls.Add(this.btnAdd);
            this.Controls.Add(this.txtSaying);
            this.Controls.Add(this.cmbFaveColor);
            this.Controls.Add(this.chkSoccer);
            this.Controls.Add(this.chkVolleyball);
            this.Controls.Add(this.chkBasketball);
            this.Controls.Add(this.radFemale);
            this.Controls.Add(this.radMale);
            this.Controls.Add(this.txtName);
            this.Controls.Add(this.lblSaying);
            this.Controls.Add(this.lblFavoriteColor);
            this.Controls.Add(this.lblHobbies);
            this.Controls.Add(this.lblGender);
            this.Controls.Add(this.lblName);
            this.Controls.Add(this.btnUpdate);
            this.Name = "FormBase";
            this.StartPosition = System.Windows.Forms.FormStartPosition.CenterScreen;
            this.Text = "FormBase";
            this.ResumeLayout(false);
            this.PerformLayout();

        }

        #endregion

        private System.Windows.Forms.Label lblName;
        private System.Windows.Forms.Label lblGender;
        private System.Windows.Forms.Label lblHobbies;
        private System.Windows.Forms.Label lblFavoriteColor;
        private System.Windows.Forms.Label lblSaying;
        public System.Windows.Forms.TextBox txtName;
        public System.Windows.Forms.RadioButton radMale;
        public System.Windows.Forms.RadioButton radFemale;
        public System.Windows.Forms.CheckBox chkBasketball;
        public System.Windows.Forms.CheckBox chkVolleyball;
        public System.Windows.Forms.CheckBox chkSoccer;
        public System.Windows.Forms.ComboBox cmbFaveColor;
        public System.Windows.Forms.TextBox txtSaying;
        public System.Windows.Forms.Button btnAdd;
        public System.Windows.Forms.Button btnDisplay;
        public System.Windows.Forms.Button btnClear;
        public System.Windows.Forms.Button btnUpdate;
        public System.Windows.Forms.Label lbliD;
        public System.Windows.Forms.TextBox txtAddress;
        public System.Windows.Forms.TextBox txtEmail;
        private System.Windows.Forms.Label lblAddress;
        private System.Windows.Forms.Label lblEmail;
        private System.Windows.Forms.Label label1;
        private System.Windows.Forms.Label lblAge;
        public System.Windows.Forms.TextBox txtAge;
        private System.Windows.Forms.Label lblUsername;
        private System.Windows.Forms.Label lblPassword;
        public System.Windows.Forms.TextBox txtUsername;
        public System.Windows.Forms.TextBox txtPassword;
        public System.Windows.Forms.DateTimePicker dtpDate;
        private System.Windows.Forms.Label lblIdentification;
        private System.Windows.Forms.Label lblCourse;
        public System.Windows.Forms.ComboBox cmbCourse;
    }
}

