namespace SQT
{
    partial class MainMenu
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
            System.ComponentModel.ComponentResourceManager resources = new System.ComponentModel.ComponentResourceManager(typeof(MainMenu));
            this.btSQAT = new System.Windows.Forms.Button();
            this.lbTitleText = new System.Windows.Forms.Label();
            this.label1 = new System.Windows.Forms.Label();
            this.pictureBox1 = new System.Windows.Forms.PictureBox();
            this.btPinDiff = new System.Windows.Forms.Button();
            this.progressBar1 = new System.Windows.Forms.ProgressBar();
            this.btnLoadOldQuote = new System.Windows.Forms.Button();
            ((System.ComponentModel.ISupportInitialize)(this.pictureBox1)).BeginInit();
            this.SuspendLayout();
            // 
            // btSQAT
            // 
            this.btSQAT.FlatAppearance.MouseDownBackColor = System.Drawing.Color.Cyan;
            this.btSQAT.FlatAppearance.MouseOverBackColor = System.Drawing.Color.FromArgb(((int)(((byte)(192)))), ((int)(((byte)(255)))), ((int)(((byte)(255)))));
            this.btSQAT.Font = new System.Drawing.Font("Microsoft Sans Serif", 18F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.btSQAT.Location = new System.Drawing.Point(129, 634);
            this.btSQAT.Name = "btSQAT";
            this.btSQAT.Size = new System.Drawing.Size(394, 62);
            this.btSQAT.TabIndex = 46;
            this.btSQAT.TabStop = false;
            this.btSQAT.Text = "Admin Settings";
            this.btSQAT.UseVisualStyleBackColor = true;
            this.btSQAT.Click += new System.EventHandler(this.btSQAT_Click);
            // 
            // lbTitleText
            // 
            this.lbTitleText.AutoSize = true;
            this.lbTitleText.Font = new System.Drawing.Font("Microsoft Sans Serif", 24F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.lbTitleText.ForeColor = System.Drawing.Color.Black;
            this.lbTitleText.Location = new System.Drawing.Point(134, 303);
            this.lbTitleText.Name = "lbTitleText";
            this.lbTitleText.Size = new System.Drawing.Size(374, 55);
            this.lbTitleText.TabIndex = 47;
            this.lbTitleText.Text = "Quote Calcuator";
            this.lbTitleText.TextAlign = System.Drawing.ContentAlignment.MiddleCenter;
            // 
            // label1
            // 
            this.label1.AutoSize = true;
            this.label1.Font = new System.Drawing.Font("Microsoft Sans Serif", 9F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.label1.ForeColor = System.Drawing.Color.Black;
            this.label1.Location = new System.Drawing.Point(218, 698);
            this.label1.Name = "label1";
            this.label1.Size = new System.Drawing.Size(211, 22);
            this.label1.TabIndex = 54;
            this.label1.Text = "© Access Elevators 2023";
            // 
            // pictureBox1
            // 
            this.pictureBox1.BackgroundImage = ((System.Drawing.Image)(resources.GetObject("pictureBox1.BackgroundImage")));
            this.pictureBox1.BackgroundImageLayout = System.Windows.Forms.ImageLayout.Stretch;
            this.pictureBox1.InitialImage = global::SQT.Properties.Resources.New_Project;
            this.pictureBox1.Location = new System.Drawing.Point(129, 3);
            this.pictureBox1.Margin = new System.Windows.Forms.Padding(4, 5, 4, 5);
            this.pictureBox1.Name = "pictureBox1";
            this.pictureBox1.Size = new System.Drawing.Size(394, 302);
            this.pictureBox1.TabIndex = 53;
            this.pictureBox1.TabStop = false;
            this.pictureBox1.Click += new System.EventHandler(this.pictureBox1_Click);
            // 
            // btPinDiff
            // 
            this.btPinDiff.BackgroundImageLayout = System.Windows.Forms.ImageLayout.Stretch;
            this.btPinDiff.FlatAppearance.MouseDownBackColor = System.Drawing.Color.Cyan;
            this.btPinDiff.FlatAppearance.MouseOverBackColor = System.Drawing.Color.FromArgb(((int)(((byte)(192)))), ((int)(((byte)(255)))), ((int)(((byte)(255)))));
            this.btPinDiff.Font = new System.Drawing.Font("Microsoft Sans Serif", 24F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.btPinDiff.Location = new System.Drawing.Point(129, 408);
            this.btPinDiff.Name = "btPinDiff";
            this.btPinDiff.Size = new System.Drawing.Size(394, 132);
            this.btPinDiff.TabIndex = 52;
            this.btPinDiff.TabStop = false;
            this.btPinDiff.Text = "Generate Quote";
            this.btPinDiff.UseVisualStyleBackColor = true;
            this.btPinDiff.Click += new System.EventHandler(this.btPinDiff_Click);
            // 
            // progressBar1
            // 
            this.progressBar1.Location = new System.Drawing.Point(129, 365);
            this.progressBar1.Margin = new System.Windows.Forms.Padding(4, 5, 4, 5);
            this.progressBar1.Name = "progressBar1";
            this.progressBar1.Size = new System.Drawing.Size(394, 35);
            this.progressBar1.Style = System.Windows.Forms.ProgressBarStyle.Continuous;
            this.progressBar1.TabIndex = 55;
            this.progressBar1.Visible = false;
            this.progressBar1.Click += new System.EventHandler(this.progressBar1_Click);
            // 
            // btnLoadOldQuote
            // 
            this.btnLoadOldQuote.FlatAppearance.MouseDownBackColor = System.Drawing.Color.Cyan;
            this.btnLoadOldQuote.FlatAppearance.MouseOverBackColor = System.Drawing.Color.FromArgb(((int)(((byte)(192)))), ((int)(((byte)(255)))), ((int)(((byte)(255)))));
            this.btnLoadOldQuote.Font = new System.Drawing.Font("Microsoft Sans Serif", 24F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.btnLoadOldQuote.Location = new System.Drawing.Point(129, 546);
            this.btnLoadOldQuote.Name = "btnLoadOldQuote";
            this.btnLoadOldQuote.Size = new System.Drawing.Size(394, 82);
            this.btnLoadOldQuote.TabIndex = 56;
            this.btnLoadOldQuote.TabStop = false;
            this.btnLoadOldQuote.Text = "Load Quote";
            this.btnLoadOldQuote.UseVisualStyleBackColor = true;
            this.btnLoadOldQuote.Click += new System.EventHandler(this.btnLoadOldQuote_Click);
            // 
            // MainMenu
            // 
            this.AutoScaleDimensions = new System.Drawing.SizeF(9F, 20F);
            this.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font;
            this.BackColor = System.Drawing.Color.SlateGray;
            this.ClientSize = new System.Drawing.Size(652, 728);
            this.Controls.Add(this.btnLoadOldQuote);
            this.Controls.Add(this.progressBar1);
            this.Controls.Add(this.label1);
            this.Controls.Add(this.pictureBox1);
            this.Controls.Add(this.btPinDiff);
            this.Controls.Add(this.lbTitleText);
            this.Controls.Add(this.btSQAT);
            this.Icon = ((System.Drawing.Icon)(resources.GetObject("$this.Icon")));
            this.Margin = new System.Windows.Forms.Padding(4, 5, 4, 5);
            this.MaximizeBox = false;
            this.MinimizeBox = false;
            this.Name = "MainMenu";
            this.StartPosition = System.Windows.Forms.FormStartPosition.CenterScreen;
            this.Text = "SQAT";
            this.Load += new System.EventHandler(this.MainMenu_Load);
            ((System.ComponentModel.ISupportInitialize)(this.pictureBox1)).EndInit();
            this.ResumeLayout(false);
            this.PerformLayout();

        }

        #endregion

        private System.Windows.Forms.Button btSQAT;
        private System.Windows.Forms.Button btPinDiff;
        private System.Windows.Forms.Label lbTitleText;
        private System.Windows.Forms.PictureBox pictureBox1;
        private System.Windows.Forms.Label label1;
        private System.Windows.Forms.ProgressBar progressBar1;
        private System.Windows.Forms.Button btnLoadOldQuote;
    }
}