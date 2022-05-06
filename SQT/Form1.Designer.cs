namespace SQT
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
            this.tBAddress = new System.Windows.Forms.TextBox();
            this.tBQuoteNumber = new System.Windows.Forms.TextBox();
            this.label2 = new System.Windows.Forms.Label();
            this.SuspendLayout();
            // 
            // label1
            // 
            this.label1.AutoSize = true;
            this.label1.Font = new System.Drawing.Font("Microsoft Sans Serif", 16F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.label1.Location = new System.Drawing.Point(12, 9);
            this.label1.Name = "label1";
            this.label1.Size = new System.Drawing.Size(135, 37);
            this.label1.TabIndex = 0;
            this.label1.Text = "Address";
            // 
            // tBAddress
            // 
            this.tBAddress.Font = new System.Drawing.Font("Microsoft Sans Serif", 16F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.tBAddress.Location = new System.Drawing.Point(153, 6);
            this.tBAddress.Name = "tBAddress";
            this.tBAddress.Size = new System.Drawing.Size(806, 44);
            this.tBAddress.TabIndex = 1;
            this.tBAddress.Text = "385-856 Pennant Hills Road, Pennant Hills";
            this.tBAddress.TextAlign = System.Windows.Forms.HorizontalAlignment.Center;
            // 
            // tBQuoteNumber
            // 
            this.tBQuoteNumber.Font = new System.Drawing.Font("Microsoft Sans Serif", 16F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.tBQuoteNumber.Location = new System.Drawing.Point(153, 56);
            this.tBQuoteNumber.Name = "tBQuoteNumber";
            this.tBQuoteNumber.Size = new System.Drawing.Size(246, 44);
            this.tBQuoteNumber.TabIndex = 3;
            this.tBQuoteNumber.Text = "Qu225-0005";
            this.tBQuoteNumber.TextAlign = System.Windows.Forms.HorizontalAlignment.Center;
            // 
            // label2
            // 
            this.label2.AutoSize = true;
            this.label2.Font = new System.Drawing.Font("Microsoft Sans Serif", 16F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.label2.Location = new System.Drawing.Point(12, 59);
            this.label2.Name = "label2";
            this.label2.Size = new System.Drawing.Size(131, 37);
            this.label2.TabIndex = 2;
            this.label2.Text = "Quote #";
            // 
            // Form1
            // 
            this.AutoScaleDimensions = new System.Drawing.SizeF(9F, 20F);
            this.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font;
            this.ClientSize = new System.Drawing.Size(1742, 1136);
            this.Controls.Add(this.tBQuoteNumber);
            this.Controls.Add(this.label2);
            this.Controls.Add(this.tBAddress);
            this.Controls.Add(this.label1);
            this.Name = "Form1";
            this.Text = "Form1";
            this.Load += new System.EventHandler(this.Form1_Load);
            this.ResumeLayout(false);
            this.PerformLayout();

        }

        #endregion

        private System.Windows.Forms.Label label1;
        private System.Windows.Forms.TextBox tBAddress;
        private System.Windows.Forms.TextBox tBQuoteNumber;
        private System.Windows.Forms.Label label2;
    }
}

