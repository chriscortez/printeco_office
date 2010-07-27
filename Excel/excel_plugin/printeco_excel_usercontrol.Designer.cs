namespace excel_plugin
{
    partial class printeco_excel_usercontrol
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

        #region Component Designer generated code

        /// <summary> 
        /// Required method for Designer support - do not modify 
        /// the contents of this method with the code editor.
        /// </summary>
        private void InitializeComponent()
        {
            this.button1 = new System.Windows.Forms.Button();
            this.label1 = new System.Windows.Forms.Label();
            this.origNumPagesLbl = new System.Windows.Forms.Label();
            this.label2 = new System.Windows.Forms.Label();
            this.finalNumPagesLbl = new System.Windows.Forms.Label();
            this.SuspendLayout();
            // 
            // button1
            // 
            this.button1.Location = new System.Drawing.Point(24, 168);
            this.button1.Name = "button1";
            this.button1.Size = new System.Drawing.Size(75, 23);
            this.button1.TabIndex = 0;
            this.button1.Text = "button1";
            this.button1.UseVisualStyleBackColor = true;
            this.button1.Click += new System.EventHandler(this.button1_Click);
            // 
            // label1
            // 
            this.label1.AutoSize = true;
            this.label1.Dock = System.Windows.Forms.DockStyle.Left;
            this.label1.Location = new System.Drawing.Point(0, 0);
            this.label1.Name = "label1";
            this.label1.Size = new System.Drawing.Size(99, 13);
            this.label1.TabIndex = 1;
            this.label1.Text = "Original # of pages:";
            // 
            // origNumPagesLbl
            // 
            this.origNumPagesLbl.AutoSize = true;
            this.origNumPagesLbl.Dock = System.Windows.Forms.DockStyle.Right;
            this.origNumPagesLbl.Location = new System.Drawing.Point(150, 0);
            this.origNumPagesLbl.Name = "origNumPagesLbl";
            this.origNumPagesLbl.Size = new System.Drawing.Size(0, 13);
            this.origNumPagesLbl.TabIndex = 2;
            // 
            // label2
            // 
            this.label2.Location = new System.Drawing.Point(0, 42);
            this.label2.Name = "label2";
            this.label2.Size = new System.Drawing.Size(100, 23);
            this.label2.TabIndex = 3;
            this.label2.Text = "Final # of pages:";
            // 
            // finalNumPagesLbl
            // 
            this.finalNumPagesLbl.Anchor = System.Windows.Forms.AnchorStyles.Right;
            this.finalNumPagesLbl.AutoSize = true;
            this.finalNumPagesLbl.Location = new System.Drawing.Point(147, 42);
            this.finalNumPagesLbl.Name = "finalNumPagesLbl";
            this.finalNumPagesLbl.Size = new System.Drawing.Size(0, 13);
            this.finalNumPagesLbl.TabIndex = 4;
            // 
            // printeco_excel_usercontrol
            // 
            this.AutoScaleDimensions = new System.Drawing.SizeF(6F, 13F);
            this.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font;
            this.Controls.Add(this.finalNumPagesLbl);
            this.Controls.Add(this.label2);
            this.Controls.Add(this.origNumPagesLbl);
            this.Controls.Add(this.label1);
            this.Controls.Add(this.button1);
            this.Name = "printeco_excel_usercontrol";
            this.Size = new System.Drawing.Size(150, 500);
            this.ResumeLayout(false);
            this.PerformLayout();

        }

        #endregion

        private System.Windows.Forms.Button button1;
        private System.Windows.Forms.Label label1;
        private System.Windows.Forms.Label origNumPagesLbl;
        private System.Windows.Forms.Label label2;
        private System.Windows.Forms.Label finalNumPagesLbl;
    }
}
