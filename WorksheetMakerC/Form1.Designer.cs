namespace WorksheetMakerC
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
            System.ComponentModel.ComponentResourceManager resources = new System.ComponentModel.ComponentResourceManager(typeof(Form1));
            this.button1 = new System.Windows.Forms.Button();
            this.btnSelFile = new System.Windows.Forms.Button();
            this.textBoxSrcPath = new System.Windows.Forms.TextBox();
            this.openFileDialog = new System.Windows.Forms.OpenFileDialog();
            this.folderBrowserDialog = new System.Windows.Forms.FolderBrowserDialog();
            this.button2 = new System.Windows.Forms.Button();
            this.textBoxTempPath = new System.Windows.Forms.TextBox();
            this.btnFolderPath = new System.Windows.Forms.Button();
            this.textBoxFolderPath = new System.Windows.Forms.TextBox();
            this.SuspendLayout();
            // 
            // button1
            // 
            this.button1.BackColor = System.Drawing.SystemColors.ActiveCaption;
            this.button1.Location = new System.Drawing.Point(90, 254);
            this.button1.Name = "button1";
            this.button1.Size = new System.Drawing.Size(138, 42);
            this.button1.TabIndex = 0;
            this.button1.Text = "Start!";
            this.button1.UseVisualStyleBackColor = false;
            this.button1.Click += new System.EventHandler(this.button1_Click);
            // 
            // btnSelFile
            // 
            this.btnSelFile.Location = new System.Drawing.Point(90, 88);
            this.btnSelFile.Name = "btnSelFile";
            this.btnSelFile.Size = new System.Drawing.Size(186, 47);
            this.btnSelFile.TabIndex = 1;
            this.btnSelFile.Text = "Select Source File";
            this.btnSelFile.UseVisualStyleBackColor = true;
            this.btnSelFile.Click += new System.EventHandler(this.btnSelFile_Click);
            // 
            // textBoxSrcPath
            // 
            this.textBoxSrcPath.Location = new System.Drawing.Point(330, 115);
            this.textBoxSrcPath.Name = "textBoxSrcPath";
            this.textBoxSrcPath.Size = new System.Drawing.Size(290, 20);
            this.textBoxSrcPath.TabIndex = 2;
            // 
            // button2
            // 
            this.button2.Location = new System.Drawing.Point(90, 356);
            this.button2.Name = "button2";
            this.button2.Size = new System.Drawing.Size(186, 47);
            this.button2.TabIndex = 4;
            this.button2.Text = "Select Template File";
            this.button2.UseVisualStyleBackColor = true;
            this.button2.Visible = false;
            this.button2.Click += new System.EventHandler(this.button2_Click);
            // 
            // textBoxTempPath
            // 
            this.textBoxTempPath.Location = new System.Drawing.Point(330, 383);
            this.textBoxTempPath.Name = "textBoxTempPath";
            this.textBoxTempPath.Size = new System.Drawing.Size(290, 20);
            this.textBoxTempPath.TabIndex = 5;
            this.textBoxTempPath.Visible = false;
            // 
            // btnFolderPath
            // 
            this.btnFolderPath.Location = new System.Drawing.Point(90, 162);
            this.btnFolderPath.Name = "btnFolderPath";
            this.btnFolderPath.Size = new System.Drawing.Size(186, 47);
            this.btnFolderPath.TabIndex = 6;
            this.btnFolderPath.Text = "Select Folder To Save Files";
            this.btnFolderPath.UseVisualStyleBackColor = true;
            this.btnFolderPath.Click += new System.EventHandler(this.btnFolderPath_Click);
            // 
            // textBoxFolderPath
            // 
            this.textBoxFolderPath.Location = new System.Drawing.Point(330, 189);
            this.textBoxFolderPath.Name = "textBoxFolderPath";
            this.textBoxFolderPath.Size = new System.Drawing.Size(290, 20);
            this.textBoxFolderPath.TabIndex = 7;
            // 
            // Form1
            // 
            this.AutoScaleDimensions = new System.Drawing.SizeF(6F, 13F);
            this.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font;
            this.ClientSize = new System.Drawing.Size(800, 450);
            this.Controls.Add(this.textBoxFolderPath);
            this.Controls.Add(this.btnFolderPath);
            this.Controls.Add(this.textBoxTempPath);
            this.Controls.Add(this.button2);
            this.Controls.Add(this.textBoxSrcPath);
            this.Controls.Add(this.btnSelFile);
            this.Controls.Add(this.button1);
            this.Icon = ((System.Drawing.Icon)(resources.GetObject("$this.Icon")));
            this.Name = "Form1";
            this.Text = "WorksheetMaker";
            this.ResumeLayout(false);
            this.PerformLayout();

        }

        #endregion

        private System.Windows.Forms.Button button1;
        private System.Windows.Forms.Button btnSelFile;
        private System.Windows.Forms.TextBox textBoxSrcPath;
        private System.Windows.Forms.OpenFileDialog openFileDialog;
        private System.Windows.Forms.FolderBrowserDialog folderBrowserDialog;
        private System.Windows.Forms.Button button2;
        private System.Windows.Forms.TextBox textBoxTempPath;
        private System.Windows.Forms.Button btnFolderPath;
        private System.Windows.Forms.TextBox textBoxFolderPath;
    }
}

