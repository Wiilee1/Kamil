namespace probanr2
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
            this.cmdSelect = new System.Windows.Forms.Button();
            this.openFileDialog1 = new System.Windows.Forms.OpenFileDialog();
            this.listView1 = new System.Windows.Forms.ListView();
            this.Name1 = ((System.Windows.Forms.ColumnHeader)(new System.Windows.Forms.ColumnHeader()));
            this.Address = ((System.Windows.Forms.ColumnHeader)(new System.Windows.Forms.ColumnHeader()));
            this.Email = ((System.Windows.Forms.ColumnHeader)(new System.Windows.Forms.ColumnHeader()));
            this.County = ((System.Windows.Forms.ColumnHeader)(new System.Windows.Forms.ColumnHeader()));
            this.Phone = ((System.Windows.Forms.ColumnHeader)(new System.Windows.Forms.ColumnHeader()));
            this.Reference = ((System.Windows.Forms.ColumnHeader)(new System.Windows.Forms.ColumnHeader()));
            this.exitButton = new System.Windows.Forms.Button();
            this.saveFileDialog1 = new System.Windows.Forms.SaveFileDialog();
            this.cmbEmp = new System.Windows.Forms.ComboBox();
            this.loadCustomers_button = new System.Windows.Forms.Button();
            this.saveCustomer_button = new System.Windows.Forms.Button();
            this.SuspendLayout();
            // 
            // cmdSelect
            // 
            this.cmdSelect.Location = new System.Drawing.Point(1548, 43);
            this.cmdSelect.Margin = new System.Windows.Forms.Padding(4, 5, 4, 5);
            this.cmdSelect.Name = "cmdSelect";
            this.cmdSelect.Size = new System.Drawing.Size(408, 35);
            this.cmdSelect.TabIndex = 0;
            this.cmdSelect.Text = "Open Excel File";
            this.cmdSelect.UseVisualStyleBackColor = true;
            this.cmdSelect.Click += new System.EventHandler(this.button1_Click);
            // 
            // openFileDialog1
            // 
            this.openFileDialog1.FileName = "openFileDialog1";
            this.openFileDialog1.FileOk += new System.ComponentModel.CancelEventHandler(this.openFileDialog1_FileOk);
            // 
            // listView1
            // 
            this.listView1.Columns.AddRange(new System.Windows.Forms.ColumnHeader[] {
            this.Name1,
            this.Address,
            this.Email,
            this.County,
            this.Phone,
            this.Reference});
            this.listView1.HideSelection = false;
            this.listView1.Location = new System.Drawing.Point(108, 43);
            this.listView1.Margin = new System.Windows.Forms.Padding(4, 5, 4, 5);
            this.listView1.Name = "listView1";
            this.listView1.Size = new System.Drawing.Size(1416, 947);
            this.listView1.TabIndex = 1;
            this.listView1.UseCompatibleStateImageBehavior = false;
            this.listView1.View = System.Windows.Forms.View.Details;
            // 
            // Name1
            // 
            this.Name1.Text = "Name";
            this.Name1.Width = 99;
            // 
            // Address
            // 
            this.Address.Text = "Address";
            this.Address.Width = 92;
            // 
            // Email
            // 
            this.Email.Text = "Email";
            this.Email.Width = 111;
            // 
            // County
            // 
            this.County.Text = "County";
            this.County.Width = 104;
            // 
            // Phone
            // 
            this.Phone.Text = "Phone";
            this.Phone.Width = 120;
            // 
            // Reference
            // 
            this.Reference.Text = "Reference";
            this.Reference.Width = 104;
            // 
            // exitButton
            // 
            this.exitButton.Cursor = System.Windows.Forms.Cursors.IBeam;
            this.exitButton.Location = new System.Drawing.Point(1674, 1014);
            this.exitButton.Margin = new System.Windows.Forms.Padding(4, 5, 4, 5);
            this.exitButton.Name = "exitButton";
            this.exitButton.Size = new System.Drawing.Size(216, 86);
            this.exitButton.TabIndex = 2;
            this.exitButton.Text = "Exit";
            this.exitButton.UseVisualStyleBackColor = true;
            this.exitButton.Click += new System.EventHandler(this.exitButton_Click);
            // 
            // saveFileDialog1
            // 
            this.saveFileDialog1.FileOk += new System.ComponentModel.CancelEventHandler(this.saveFileDialog1_FileOk);
            // 
            // cmbEmp
            // 
            this.cmbEmp.FormattingEnabled = true;
            this.cmbEmp.Location = new System.Drawing.Point(1569, 234);
            this.cmbEmp.Margin = new System.Windows.Forms.Padding(4, 5, 4, 5);
            this.cmbEmp.Name = "cmbEmp";
            this.cmbEmp.Size = new System.Drawing.Size(366, 28);
            this.cmbEmp.TabIndex = 3;
            // 
            // loadCustomers_button
            // 
            this.loadCustomers_button.Location = new System.Drawing.Point(1548, 425);
            this.loadCustomers_button.Name = "loadCustomers_button";
            this.loadCustomers_button.Size = new System.Drawing.Size(200, 106);
            this.loadCustomers_button.TabIndex = 4;
            this.loadCustomers_button.Text = "Load File";
            this.loadCustomers_button.UseVisualStyleBackColor = true;
            this.loadCustomers_button.Click += new System.EventHandler(this.loadCustomers_button_Click);
            // 
            // saveCustomer_button
            // 
            this.saveCustomer_button.Location = new System.Drawing.Point(1548, 652);
            this.saveCustomer_button.Name = "saveCustomer_button";
            this.saveCustomer_button.Size = new System.Drawing.Size(200, 95);
            this.saveCustomer_button.TabIndex = 5;
            this.saveCustomer_button.Text = "Save File";
            this.saveCustomer_button.UseVisualStyleBackColor = true;
            this.saveCustomer_button.Click += new System.EventHandler(this.saveCustomer_button_Click);
            // 
            // Form1
            // 
            this.AutoScaleDimensions = new System.Drawing.SizeF(9F, 20F);
            this.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font;
            this.ClientSize = new System.Drawing.Size(1988, 1151);
            this.Controls.Add(this.saveCustomer_button);
            this.Controls.Add(this.loadCustomers_button);
            this.Controls.Add(this.cmbEmp);
            this.Controls.Add(this.exitButton);
            this.Controls.Add(this.listView1);
            this.Controls.Add(this.cmdSelect);
            this.FormBorderStyle = System.Windows.Forms.FormBorderStyle.None;
            this.Margin = new System.Windows.Forms.Padding(4, 5, 4, 5);
            this.Name = "Form1";
            this.Text = "Form1";
            this.Load += new System.EventHandler(this.Form1_Load);
            this.ResumeLayout(false);

        }

        #endregion

        private System.Windows.Forms.Button cmdSelect;
        private System.Windows.Forms.OpenFileDialog openFileDialog1;
        private System.Windows.Forms.ListView listView1;
        private System.Windows.Forms.ColumnHeader Name1;
        private System.Windows.Forms.ColumnHeader Address;
        private System.Windows.Forms.ColumnHeader Email;
        private System.Windows.Forms.ColumnHeader County;
        private System.Windows.Forms.ColumnHeader Phone;
        private System.Windows.Forms.ColumnHeader Reference;
        private System.Windows.Forms.Button exitButton;
        private System.Windows.Forms.SaveFileDialog saveFileDialog1;
        private System.Windows.Forms.ComboBox cmbEmp;
        private System.Windows.Forms.Button loadCustomers_button;
        private System.Windows.Forms.Button saveCustomer_button;
    }
}

