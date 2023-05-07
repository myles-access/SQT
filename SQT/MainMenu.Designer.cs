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
            this.label3 = new System.Windows.Forms.Label();
            this.btPinSingle = new System.Windows.Forms.Button();
            this.btPinMulti = new System.Windows.Forms.Button();
            this.btPinDiff = new System.Windows.Forms.Button();
            this.SuspendLayout();
            // 
            // label3
            // 
            this.label3.AutoSize = true;
            this.label3.Font = new System.Drawing.Font("Microsoft Sans Serif", 24F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.label3.ForeColor = System.Drawing.Color.Black;
            this.label3.Location = new System.Drawing.Point(64, 71);
            this.label3.Name = "label3";
            this.label3.Size = new System.Drawing.Size(262, 55);
            this.label3.TabIndex = 47;
            this.label3.Text = "Main Menu";
            // 
            // btPinSingle
            // 
            this.btPinSingle.Font = new System.Drawing.Font("Microsoft Sans Serif", 24F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.btPinSingle.Location = new System.Drawing.Point(16, 160);
            this.btPinSingle.Name = "btPinSingle";
            this.btPinSingle.Size = new System.Drawing.Size(368, 131);
            this.btPinSingle.TabIndex = 48;
            this.btPinSingle.TabStop = false;
            this.btPinSingle.Text = "Single Lift";
            this.btPinSingle.UseVisualStyleBackColor = true;
            this.btPinSingle.Click += new System.EventHandler(this.button1_Click);
            // 
            // btPinMulti
            // 
            this.btPinMulti.Font = new System.Drawing.Font("Microsoft Sans Serif", 24F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.btPinMulti.Location = new System.Drawing.Point(16, 297);
            this.btPinMulti.Name = "btPinMulti";
            this.btPinMulti.Size = new System.Drawing.Size(368, 131);
            this.btPinMulti.TabIndex = 50;
            this.btPinMulti.TabStop = false;
            this.btPinMulti.Text = "Multiple Lifts";
            this.btPinMulti.UseVisualStyleBackColor = true;
            this.btPinMulti.Click += new System.EventHandler(this.btPinMulti_Click);
            // 
            // btPinDiff
            // 
            this.btPinDiff.Font = new System.Drawing.Font("Microsoft Sans Serif", 24F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.btPinDiff.Location = new System.Drawing.Point(16, 434);
            this.btPinDiff.Name = "btPinDiff";
            this.btPinDiff.Size = new System.Drawing.Size(368, 131);
            this.btPinDiff.TabIndex = 52;
            this.btPinDiff.TabStop = false;
            this.btPinDiff.Text = "Different Lifts";
            this.btPinDiff.UseVisualStyleBackColor = true;
            this.btPinDiff.Click += new System.EventHandler(this.btPinDiff_Click);
            // 
            // MainMenu
            // 
            this.AutoScaleDimensions = new System.Drawing.SizeF(9F, 20F);
            this.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font;
            this.BackColor = System.Drawing.Color.SlateGray;
            this.ClientSize = new System.Drawing.Size(395, 584);
            this.Controls.Add(this.btPinDiff);
            this.Controls.Add(this.btPinMulti);
            this.Controls.Add(this.btPinSingle);
            this.Controls.Add(this.label3);
            this.Margin = new System.Windows.Forms.Padding(4, 5, 4, 5);
            this.Name = "MainMenu";
            this.StartPosition = System.Windows.Forms.FormStartPosition.CenterScreen;
            this.Text = "SQAT";
            this.Load += new System.EventHandler(this.MainMenu_Load);
            this.ResumeLayout(false);
            this.PerformLayout();

        }

        #endregion
        private System.Windows.Forms.Label label3;
        private System.Windows.Forms.Button btPinSingle;
        private System.Windows.Forms.Button btPinMulti;
        private System.Windows.Forms.Button btPinDiff;
    }
}