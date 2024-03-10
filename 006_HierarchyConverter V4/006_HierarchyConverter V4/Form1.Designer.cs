namespace _006_HierarchyConverter_V4
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
            this.progressBar = new System.Windows.Forms.ProgressBar();
            this.InputBtn = new System.Windows.Forms.Button();
            this.CnvrtBtn = new System.Windows.Forms.Button();
            this.openFileDialog1 = new System.Windows.Forms.OpenFileDialog();
            this.textBox2 = new System.Windows.Forms.TextBox();
            this.OutputBtn = new System.Windows.Forms.Button();
            this.label1 = new System.Windows.Forms.Label();
            this.textBox3 = new System.Windows.Forms.TextBox();
            this.openFileDialog2 = new System.Windows.Forms.OpenFileDialog();
            this.VerificationBtn = new System.Windows.Forms.Button();
            this.ExitBtn = new System.Windows.Forms.Button();
            this.RestartBtn = new System.Windows.Forms.Button();
            this.button1 = new System.Windows.Forms.Button();
            this.textBox4 = new System.Windows.Forms.TextBox();
            this.SuspendLayout();
            // 
            // textBox1
            // 
            this.textBox1.BackColor = System.Drawing.Color.AliceBlue;
            this.textBox1.Cursor = System.Windows.Forms.Cursors.No;
            this.textBox1.Font = new System.Drawing.Font("Times New Roman", 9.5F, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.textBox1.HideSelection = false;
            this.textBox1.Location = new System.Drawing.Point(359, 26);
            this.textBox1.Margin = new System.Windows.Forms.Padding(4, 3, 4, 3);
            this.textBox1.Multiline = true;
            this.textBox1.Name = "textBox1";
            this.textBox1.ReadOnly = true;
            this.textBox1.Size = new System.Drawing.Size(366, 32);
            this.textBox1.TabIndex = 0;
            this.textBox1.TabStop = false;
            // 
            // progressBar
            // 
            this.progressBar.BackColor = System.Drawing.Color.Azure;
            this.progressBar.Cursor = System.Windows.Forms.Cursors.No;
            this.progressBar.Location = new System.Drawing.Point(71, 337);
            this.progressBar.Margin = new System.Windows.Forms.Padding(4, 3, 4, 3);
            this.progressBar.Name = "progressBar";
            this.progressBar.Size = new System.Drawing.Size(654, 27);
            this.progressBar.Style = System.Windows.Forms.ProgressBarStyle.Continuous;
            this.progressBar.TabIndex = 1;
            // 
            // InputBtn
            // 
            this.InputBtn.BackColor = System.Drawing.SystemColors.GradientActiveCaption;
            this.InputBtn.Cursor = System.Windows.Forms.Cursors.Hand;
            this.InputBtn.DialogResult = System.Windows.Forms.DialogResult.Cancel;
            this.InputBtn.FlatAppearance.BorderColor = System.Drawing.Color.Black;
            this.InputBtn.FlatAppearance.MouseOverBackColor = System.Drawing.Color.White;
            this.InputBtn.Font = new System.Drawing.Font("Times New Roman", 9F, System.Drawing.FontStyle.Bold);
            this.InputBtn.Location = new System.Drawing.Point(71, 32);
            this.InputBtn.Margin = new System.Windows.Forms.Padding(4, 3, 4, 3);
            this.InputBtn.Name = "InputBtn";
            this.InputBtn.Size = new System.Drawing.Size(216, 35);
            this.InputBtn.TabIndex = 2;
            this.InputBtn.Text = "Select Input File";
            this.InputBtn.UseVisualStyleBackColor = false;
            this.InputBtn.Click += new System.EventHandler(this.InputBtn_Click);
            // 
            // CnvrtBtn
            // 
            this.CnvrtBtn.BackColor = System.Drawing.SystemColors.GradientActiveCaption;
            this.CnvrtBtn.Cursor = System.Windows.Forms.Cursors.Hand;
            this.CnvrtBtn.DialogResult = System.Windows.Forms.DialogResult.Cancel;
            this.CnvrtBtn.Font = new System.Drawing.Font("Times New Roman", 9F, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.CnvrtBtn.Location = new System.Drawing.Point(274, 399);
            this.CnvrtBtn.Margin = new System.Windows.Forms.Padding(4, 3, 4, 3);
            this.CnvrtBtn.Name = "CnvrtBtn";
            this.CnvrtBtn.Size = new System.Drawing.Size(233, 44);
            this.CnvrtBtn.TabIndex = 3;
            this.CnvrtBtn.Text = "Convert";
            this.CnvrtBtn.UseVisualStyleBackColor = false;
            this.CnvrtBtn.Click += new System.EventHandler(this.CnvrtBtn_Click);
            // 
            // openFileDialog1
            // 
            this.openFileDialog1.FileName = "openFileDialog1";
            // 
            // textBox2
            // 
            this.textBox2.BackColor = System.Drawing.Color.AliceBlue;
            this.textBox2.Cursor = System.Windows.Forms.Cursors.No;
            this.textBox2.Font = new System.Drawing.Font("Times New Roman", 9.75F, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.textBox2.Location = new System.Drawing.Point(359, 240);
            this.textBox2.Margin = new System.Windows.Forms.Padding(2);
            this.textBox2.Multiline = true;
            this.textBox2.Name = "textBox2";
            this.textBox2.ReadOnly = true;
            this.textBox2.Size = new System.Drawing.Size(366, 34);
            this.textBox2.TabIndex = 4;
            // 
            // OutputBtn
            // 
            this.OutputBtn.BackColor = System.Drawing.SystemColors.GradientActiveCaption;
            this.OutputBtn.Cursor = System.Windows.Forms.Cursors.Hand;
            this.OutputBtn.Font = new System.Drawing.Font("Times New Roman", 9F, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.OutputBtn.Location = new System.Drawing.Point(71, 240);
            this.OutputBtn.Margin = new System.Windows.Forms.Padding(2);
            this.OutputBtn.Name = "OutputBtn";
            this.OutputBtn.Size = new System.Drawing.Size(216, 34);
            this.OutputBtn.TabIndex = 5;
            this.OutputBtn.Text = "Select Output Path";
            this.OutputBtn.UseVisualStyleBackColor = false;
            this.OutputBtn.Click += new System.EventHandler(this.OutputBtn_Click);
            // 
            // label1
            // 
            this.label1.AutoSize = true;
            this.label1.BackColor = System.Drawing.SystemColors.GradientInactiveCaption;
            this.label1.Cursor = System.Windows.Forms.Cursors.Arrow;
            this.label1.Font = new System.Drawing.Font("Times New Roman", 11F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.label1.ForeColor = System.Drawing.SystemColors.ControlText;
            this.label1.Location = new System.Drawing.Point(68, 367);
            this.label1.Margin = new System.Windows.Forms.Padding(2, 0, 2, 0);
            this.label1.Name = "label1";
            this.label1.Size = new System.Drawing.Size(35, 17);
            this.label1.TabIndex = 6;
            this.label1.Text = "label";
            // 
            // textBox3
            // 
            this.textBox3.BackColor = System.Drawing.Color.AliceBlue;
            this.textBox3.Cursor = System.Windows.Forms.Cursors.No;
            this.textBox3.Font = new System.Drawing.Font("Times New Roman", 9.75F, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.textBox3.Location = new System.Drawing.Point(359, 100);
            this.textBox3.Margin = new System.Windows.Forms.Padding(2);
            this.textBox3.Multiline = true;
            this.textBox3.Name = "textBox3";
            this.textBox3.ReadOnly = true;
            this.textBox3.Size = new System.Drawing.Size(366, 34);
            this.textBox3.TabIndex = 7;
            // 
            // openFileDialog2
            // 
            this.openFileDialog2.FileName = "openFileDialog2";
            // 
            // VerificationBtn
            // 
            this.VerificationBtn.BackColor = System.Drawing.SystemColors.GradientActiveCaption;
            this.VerificationBtn.Cursor = System.Windows.Forms.Cursors.Hand;
            this.VerificationBtn.Font = new System.Drawing.Font("Times New Roman", 9F, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.VerificationBtn.Location = new System.Drawing.Point(71, 100);
            this.VerificationBtn.Margin = new System.Windows.Forms.Padding(2);
            this.VerificationBtn.Name = "VerificationBtn";
            this.VerificationBtn.Size = new System.Drawing.Size(216, 34);
            this.VerificationBtn.TabIndex = 8;
            this.VerificationBtn.Text = "Select Verification File";
            this.VerificationBtn.UseVisualStyleBackColor = false;
            this.VerificationBtn.Click += new System.EventHandler(this.VerificationBtn_Click);
            // 
            // ExitBtn
            // 
            this.ExitBtn.BackColor = System.Drawing.Color.Red;
            this.ExitBtn.Cursor = System.Windows.Forms.Cursors.Hand;
            this.ExitBtn.Font = new System.Drawing.Font("Times New Roman", 9F, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.ExitBtn.ForeColor = System.Drawing.SystemColors.ButtonHighlight;
            this.ExitBtn.Location = new System.Drawing.Point(666, 494);
            this.ExitBtn.Margin = new System.Windows.Forms.Padding(4, 3, 4, 3);
            this.ExitBtn.Name = "ExitBtn";
            this.ExitBtn.Size = new System.Drawing.Size(88, 27);
            this.ExitBtn.TabIndex = 9;
            this.ExitBtn.Text = "Exit";
            this.ExitBtn.UseVisualStyleBackColor = false;
            this.ExitBtn.Click += new System.EventHandler(this.ExitBtn_Click);
            // 
            // RestartBtn
            // 
            this.RestartBtn.BackColor = System.Drawing.Color.Green;
            this.RestartBtn.Cursor = System.Windows.Forms.Cursors.Hand;
            this.RestartBtn.Font = new System.Drawing.Font("Times New Roman", 9F, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.RestartBtn.Location = new System.Drawing.Point(37, 504);
            this.RestartBtn.Margin = new System.Windows.Forms.Padding(4, 3, 4, 3);
            this.RestartBtn.Name = "RestartBtn";
            this.RestartBtn.Size = new System.Drawing.Size(88, 27);
            this.RestartBtn.TabIndex = 10;
            this.RestartBtn.Text = "Reset";
            this.RestartBtn.UseVisualStyleBackColor = false;
            this.RestartBtn.Click += new System.EventHandler(this.RestartBtn_Click);
            // 
            // button1
            // 
            this.button1.BackColor = System.Drawing.SystemColors.GradientActiveCaption;
            this.button1.Font = new System.Drawing.Font("Times New Roman", 9F, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.button1.Location = new System.Drawing.Point(71, 171);
            this.button1.Name = "button1";
            this.button1.Size = new System.Drawing.Size(216, 35);
            this.button1.TabIndex = 11;
            this.button1.Text = "Select Maximo Job File";
            this.button1.UseVisualStyleBackColor = false;
            this.button1.Click += new System.EventHandler(this.MaximoJob_Click);
            // 
            // textBox4
            // 
            this.textBox4.BackColor = System.Drawing.Color.GhostWhite;
            this.textBox4.Location = new System.Drawing.Point(359, 171);
            this.textBox4.Multiline = true;
            this.textBox4.Name = "textBox4";
            this.textBox4.Size = new System.Drawing.Size(366, 35);
            this.textBox4.TabIndex = 12;
            this.textBox4.TextChanged += new System.EventHandler(this.textBox4_TextChanged);
            // 
            // Form1
            // 
            this.AutoScaleDimensions = new System.Drawing.SizeF(7F, 15F);
            this.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font;
            this.BackColor = System.Drawing.SystemColors.GradientInactiveCaption;
            this.BackgroundImageLayout = System.Windows.Forms.ImageLayout.Center;
            this.ClientSize = new System.Drawing.Size(830, 603);
            this.Controls.Add(this.textBox4);
            this.Controls.Add(this.button1);
            this.Controls.Add(this.RestartBtn);
            this.Controls.Add(this.ExitBtn);
            this.Controls.Add(this.VerificationBtn);
            this.Controls.Add(this.textBox3);
            this.Controls.Add(this.label1);
            this.Controls.Add(this.OutputBtn);
            this.Controls.Add(this.textBox2);
            this.Controls.Add(this.CnvrtBtn);
            this.Controls.Add(this.InputBtn);
            this.Controls.Add(this.progressBar);
            this.Controls.Add(this.textBox1);
            this.DoubleBuffered = true;
            this.Font = new System.Drawing.Font("Times New Roman", 9F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.FormBorderStyle = System.Windows.Forms.FormBorderStyle.Fixed3D;
            this.Margin = new System.Windows.Forms.Padding(4, 3, 4, 3);
            this.MaximizeBox = false;
            this.Name = "Form1";
            this.StartPosition = System.Windows.Forms.FormStartPosition.CenterScreen;
            this.Text = "Hierarchy Conversion";
            this.Load += new System.EventHandler(this.Form1_Load);
            this.ResumeLayout(false);
            this.PerformLayout();

        }

        #endregion

        private System.Windows.Forms.TextBox textBox1;
        private System.Windows.Forms.Button InputBtn;
        private System.Windows.Forms.Button CnvrtBtn;
        private System.Windows.Forms.OpenFileDialog openFileDialog1;
        public System.Windows.Forms.ProgressBar progressBar;
        private System.Windows.Forms.TextBox textBox2;
        private System.Windows.Forms.Button OutputBtn;
        private System.Windows.Forms.Label label1;
        private System.Windows.Forms.TextBox textBox3;
        private System.Windows.Forms.OpenFileDialog openFileDialog2;
        private System.Windows.Forms.Button VerificationBtn;
        private System.Windows.Forms.Button ExitBtn;
        private System.Windows.Forms.Button RestartBtn;
        private System.Windows.Forms.Button button1;
        private System.Windows.Forms.TextBox textBox4;
    }
}

