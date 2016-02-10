namespace PM_Folder_Generator
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
            this.label1 = new System.Windows.Forms.Label();
            this.label2 = new System.Windows.Forms.Label();
            this.label3 = new System.Windows.Forms.Label();
            this.Q2COrderNumberTextBox = new System.Windows.Forms.TextBox();
            this.projectNameTextBox = new System.Windows.Forms.TextBox();
            this.createFoldersButton = new System.Windows.Forms.Button();
            this.SuspendLayout();
            // 
            // label1
            // 
            this.label1.AutoSize = true;
            this.label1.Font = new System.Drawing.Font("Arial", 12F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.label1.Location = new System.Drawing.Point(56, 44);
            this.label1.Name = "label1";
            this.label1.Size = new System.Drawing.Size(275, 18);
            this.label1.TabIndex = 0;
            this.label1.Text = "Project Management Folder Generator";
            // 
            // label2
            // 
            this.label2.AutoSize = true;
            this.label2.Font = new System.Drawing.Font("Arial", 12F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.label2.Location = new System.Drawing.Point(22, 116);
            this.label2.Name = "label2";
            this.label2.Size = new System.Drawing.Size(148, 18);
            this.label2.TabIndex = 1;
            this.label2.Text = "Q2C Order Number:";
            // 
            // label3
            // 
            this.label3.AutoSize = true;
            this.label3.Font = new System.Drawing.Font("Arial", 12F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.label3.Location = new System.Drawing.Point(22, 178);
            this.label3.Name = "label3";
            this.label3.Size = new System.Drawing.Size(107, 18);
            this.label3.TabIndex = 2;
            this.label3.Text = "Project Name:";
            // 
            // Q2COrderNumberTextBox
            // 
            this.Q2COrderNumberTextBox.Location = new System.Drawing.Point(203, 114);
            this.Q2COrderNumberTextBox.Name = "Q2COrderNumberTextBox";
            this.Q2COrderNumberTextBox.Size = new System.Drawing.Size(128, 20);
            this.Q2COrderNumberTextBox.TabIndex = 0;
            this.Q2COrderNumberTextBox.TextChanged += new System.EventHandler(this.EnterQ2COrderNumber);
            // 
            // projectNameTextBox
            // 
            this.projectNameTextBox.Location = new System.Drawing.Point(203, 176);
            this.projectNameTextBox.Name = "projectNameTextBox";
            this.projectNameTextBox.Size = new System.Drawing.Size(128, 20);
            this.projectNameTextBox.TabIndex = 1;
            this.projectNameTextBox.TextChanged += new System.EventHandler(this.EnterProjectName);
            // 
            // createFoldersButton
            // 
            this.createFoldersButton.Font = new System.Drawing.Font("Arial", 10F, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.createFoldersButton.Location = new System.Drawing.Point(152, 262);
            this.createFoldersButton.Name = "createFoldersButton";
            this.createFoldersButton.Size = new System.Drawing.Size(75, 40);
            this.createFoldersButton.TabIndex = 2;
            this.createFoldersButton.Text = "CREATE";
            this.createFoldersButton.UseVisualStyleBackColor = true;
            this.createFoldersButton.Click += new System.EventHandler(this.createFoldersButton_Click);
            // 
            // Form1
            // 
            this.AutoScaleDimensions = new System.Drawing.SizeF(6F, 13F);
            this.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font;
            this.ClientSize = new System.Drawing.Size(384, 361);
            this.Controls.Add(this.createFoldersButton);
            this.Controls.Add(this.projectNameTextBox);
            this.Controls.Add(this.Q2COrderNumberTextBox);
            this.Controls.Add(this.label3);
            this.Controls.Add(this.label2);
            this.Controls.Add(this.label1);
            this.FormBorderStyle = System.Windows.Forms.FormBorderStyle.Fixed3D;
            this.MaximizeBox = false;
            this.Name = "Form1";
            this.Text = "PM Folder Generator";
            this.ResumeLayout(false);
            this.PerformLayout();

        }

        #endregion

        private System.Windows.Forms.Label label1;
        private System.Windows.Forms.Label label2;
        private System.Windows.Forms.Label label3;
        private System.Windows.Forms.TextBox Q2COrderNumberTextBox;
        private System.Windows.Forms.TextBox projectNameTextBox;
        private System.Windows.Forms.Button createFoldersButton;
    }
}

