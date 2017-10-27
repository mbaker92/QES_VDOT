namespace NewVdotProcessing
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
            this.textBox1 = new System.Windows.Forms.TextBox();
            this.label1 = new System.Windows.Forms.Label();
            this.label2 = new System.Windows.Forms.Label();
            this.folderBrowserDialog1 = new System.Windows.Forms.FolderBrowserDialog();
            this.Browser = new System.Windows.Forms.Button();
            this.Start = new System.Windows.Forms.Button();
            this.label3 = new System.Windows.Forms.Label();
            this.label4 = new System.Windows.Forms.Label();
            this.InstructButton = new System.Windows.Forms.Button();
            this.button2 = new System.Windows.Forms.Button();
            this.SuspendLayout();
            // 
            // textBox1
            // 
            this.textBox1.BackColor = System.Drawing.Color.White;
            this.textBox1.Location = new System.Drawing.Point(42, 68);
            this.textBox1.Name = "textBox1";
            this.textBox1.Size = new System.Drawing.Size(399, 20);
            this.textBox1.TabIndex = 0;
            this.textBox1.Text = "NewDatabase";
            // 
            // label1
            // 
            this.label1.AutoSize = true;
            this.label1.Location = new System.Drawing.Point(12, 42);
            this.label1.Name = "label1";
            this.label1.Size = new System.Drawing.Size(149, 13);
            this.label1.TabIndex = 1;
            this.label1.Text = "Enter Name of New Database";
            // 
            // label2
            // 
            this.label2.AutoSize = true;
            this.label2.Location = new System.Drawing.Point(12, 104);
            this.label2.Name = "label2";
            this.label2.Size = new System.Drawing.Size(113, 13);
            this.label2.TabIndex = 2;
            this.label2.Text = "Select Folder Location";
            // 
            // folderBrowserDialog1
            // 
            this.folderBrowserDialog1.Description = "Choose Folder Location of New Database";
            this.folderBrowserDialog1.SelectedPath = "\\\\QES-RENO\\Projects\\VDOT";
            // 
            // Browser
            // 
            this.Browser.Location = new System.Drawing.Point(366, 99);
            this.Browser.Name = "Browser";
            this.Browser.Size = new System.Drawing.Size(75, 23);
            this.Browser.TabIndex = 3;
            this.Browser.Text = "Browse";
            this.Browser.UseVisualStyleBackColor = true;
            this.Browser.Click += new System.EventHandler(this.Browser_Click);
            // 
            // Start
            // 
            this.Start.Location = new System.Drawing.Point(215, 153);
            this.Start.Name = "Start";
            this.Start.Size = new System.Drawing.Size(75, 23);
            this.Start.TabIndex = 4;
            this.Start.Text = "START";
            this.Start.UseVisualStyleBackColor = true;
            this.Start.Click += new System.EventHandler(this.Start_Click);
            // 
            // label3
            // 
            this.label3.AutoSize = true;
            this.label3.Location = new System.Drawing.Point(12, 129);
            this.label3.Name = "label3";
            this.label3.Size = new System.Drawing.Size(0, 13);
            this.label3.TabIndex = 5;
            // 
            // label4
            // 
            this.label4.AutoSize = true;
            this.label4.Font = new System.Drawing.Font("Microsoft Sans Serif", 8.25F, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.label4.Location = new System.Drawing.Point(12, 18);
            this.label4.Name = "label4";
            this.label4.Size = new System.Drawing.Size(214, 13);
            this.label4.TabIndex = 6;
            this.label4.Text = "Read Instructions Before Processing";
            // 
            // InstructButton
            // 
            this.InstructButton.Location = new System.Drawing.Point(311, 13);
            this.InstructButton.Name = "InstructButton";
            this.InstructButton.Size = new System.Drawing.Size(130, 23);
            this.InstructButton.TabIndex = 7;
            this.InstructButton.Text = "View Instructions";
            this.InstructButton.UseVisualStyleBackColor = true;
            this.InstructButton.Click += new System.EventHandler(this.InstructButton_Click);
            // 
            // button2
            // 
            this.button2.Location = new System.Drawing.Point(243, 153);
            this.button2.Name = "button2";
            this.button2.Size = new System.Drawing.Size(118, 23);
            this.button2.TabIndex = 8;
            this.button2.Text = "High Level Review";
            this.button2.UseVisualStyleBackColor = true;
            this.button2.Visible = false;
            this.button2.Click += new System.EventHandler(this.button2_Click);
            // 
            // Form1
            // 
            this.AutoScaleDimensions = new System.Drawing.SizeF(6F, 13F);
            this.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font;
            this.BackColor = System.Drawing.SystemColors.Control;
            this.ClientSize = new System.Drawing.Size(497, 188);
            this.Controls.Add(this.button2);
            this.Controls.Add(this.InstructButton);
            this.Controls.Add(this.label4);
            this.Controls.Add(this.label3);
            this.Controls.Add(this.Start);
            this.Controls.Add(this.Browser);
            this.Controls.Add(this.label2);
            this.Controls.Add(this.label1);
            this.Controls.Add(this.textBox1);
            this.Name = "Form1";
            this.Text = "VDOT Processing";
            this.ResumeLayout(false);
            this.PerformLayout();

        }

        #endregion

        private System.Windows.Forms.TextBox textBox1;
        private System.Windows.Forms.Label label1;
        private System.Windows.Forms.Label label2;
        private System.Windows.Forms.FolderBrowserDialog folderBrowserDialog1;
        private System.Windows.Forms.Button Browser;
        private System.Windows.Forms.Button Start;
        private System.Windows.Forms.Label label3;
        private System.Windows.Forms.Label label4;
        private System.Windows.Forms.Button InstructButton;
        private System.Windows.Forms.Button button2;
    }
}

