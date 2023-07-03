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
            this.panelShipping = new System.Windows.Forms.Panel();
            this.btnLoadSelectedITem = new System.Windows.Forms.Button();
            this.extraCostsClose = new System.Windows.Forms.Button();
            this.tbLoadSearch = new System.Windows.Forms.TextBox();
            this.listBox1 = new System.Windows.Forms.ListBox();
            ((System.ComponentModel.ISupportInitialize)(this.pictureBox1)).BeginInit();
            this.panelShipping.SuspendLayout();
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
            this.btSQAT.Text = "Settings";
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
            this.pictureBox1.InitialImage = null;
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
            // panelShipping
            // 
            this.panelShipping.BackColor = System.Drawing.Color.LightSteelBlue;
            this.panelShipping.Controls.Add(this.btnLoadSelectedITem);
            this.panelShipping.Controls.Add(this.extraCostsClose);
            this.panelShipping.Controls.Add(this.tbLoadSearch);
            this.panelShipping.Controls.Add(this.listBox1);
            this.panelShipping.Location = new System.Drawing.Point(571, 14);
            this.panelShipping.Margin = new System.Windows.Forms.Padding(4, 5, 4, 5);
            this.panelShipping.Name = "panelShipping";
            this.panelShipping.Size = new System.Drawing.Size(737, 706);
            this.panelShipping.TabIndex = 191;
            // 
            // btnLoadSelectedITem
            // 
            this.btnLoadSelectedITem.Location = new System.Drawing.Point(599, 50);
            this.btnLoadSelectedITem.Name = "btnLoadSelectedITem";
            this.btnLoadSelectedITem.Size = new System.Drawing.Size(116, 47);
            this.btnLoadSelectedITem.TabIndex = 442;
            this.btnLoadSelectedITem.Text = "LOAD";
            this.btnLoadSelectedITem.UseVisualStyleBackColor = true;
            this.btnLoadSelectedITem.Click += new System.EventHandler(this.btnLoadSelectedITem_Click);
            // 
            // extraCostsClose
            // 
            this.extraCostsClose.BackColor = System.Drawing.Color.FromArgb(((int)(((byte)(255)))), ((int)(((byte)(128)))), ((int)(((byte)(128)))));
            this.extraCostsClose.Font = new System.Drawing.Font("Microsoft Sans Serif", 12F, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.extraCostsClose.Location = new System.Drawing.Point(3, 3);
            this.extraCostsClose.Name = "extraCostsClose";
            this.extraCostsClose.Size = new System.Drawing.Size(36, 41);
            this.extraCostsClose.TabIndex = 441;
            this.extraCostsClose.Text = "X";
            this.extraCostsClose.UseVisualStyleBackColor = false;
            this.extraCostsClose.Click += new System.EventHandler(this.extraCostsClose_Click);
            // 
            // tbLoadSearch
            // 
            this.tbLoadSearch.Font = new System.Drawing.Font("Calibri", 16F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.tbLoadSearch.Location = new System.Drawing.Point(36, 50);
            this.tbLoadSearch.Name = "tbLoadSearch";
            this.tbLoadSearch.Size = new System.Drawing.Size(557, 47);
            this.tbLoadSearch.TabIndex = 1;
            this.tbLoadSearch.TextChanged += new System.EventHandler(this.tbLoadSearch_TextChanged);
            // 
            // listBox1
            // 
            this.listBox1.Font = new System.Drawing.Font("Calibri", 12F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.listBox1.FormattingEnabled = true;
            this.listBox1.ItemHeight = 29;
            this.listBox1.Items.AddRange(new object[] {
            "test1",
            "test2",
            "test3",
            "test4",
            "test5"});
            this.listBox1.Location = new System.Drawing.Point(36, 113);
            this.listBox1.Name = "listBox1";
            this.listBox1.Size = new System.Drawing.Size(679, 555);
            this.listBox1.TabIndex = 0;
            this.listBox1.SelectedIndexChanged += new System.EventHandler(this.listBox1_SelectedIndexChanged);
            // 
            // MainMenu
            // 
            this.AutoScaleDimensions = new System.Drawing.SizeF(9F, 20F);
            this.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font;
            this.BackColor = System.Drawing.Color.SlateGray;
            this.ClientSize = new System.Drawing.Size(1358, 728);
            this.Controls.Add(this.panelShipping);
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
            this.panelShipping.ResumeLayout(false);
            this.panelShipping.PerformLayout();
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
        private System.Windows.Forms.Panel panelShipping;
        private System.Windows.Forms.ListBox listBox1;
        private System.Windows.Forms.TextBox tbLoadSearch;
        private System.Windows.Forms.Button extraCostsClose;
        private System.Windows.Forms.Button btnLoadSelectedITem;
    }
}