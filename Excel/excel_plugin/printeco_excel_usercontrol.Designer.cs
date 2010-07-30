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
            this.printBtn = new System.Windows.Forms.Button();
            this.origNumPagesLbl = new System.Windows.Forms.Label();
            this.label2 = new System.Windows.Forms.Label();
            this.finalNumPagesLbl = new System.Windows.Forms.Label();
            this.label1 = new System.Windows.Forms.Label();
            this.label3 = new System.Windows.Forms.Label();
            this.savedPagesLbl = new System.Windows.Forms.Label();
            this.printPreviewBtn = new System.Windows.Forms.Button();
            this.printerList = new System.Windows.Forms.ComboBox();
            this.numCopiesUpDwn = new System.Windows.Forms.NumericUpDown();
            ((System.ComponentModel.ISupportInitialize)(this.numCopiesUpDwn)).BeginInit();
            this.SuspendLayout();
            // 
            // printBtn
            // 
            this.printBtn.BackColor = System.Drawing.Color.Transparent;
            this.printBtn.Image = global::excel_plugin.Properties.Resources.print_button;
            this.printBtn.Location = new System.Drawing.Point(3, 75);
            this.printBtn.Name = "printBtn";
            this.printBtn.Size = new System.Drawing.Size(84, 71);
            this.printBtn.TabIndex = 0;
            this.printBtn.UseVisualStyleBackColor = false;
            this.printBtn.Click += new System.EventHandler(this.printBtn_Click);
            // 
            // origNumPagesLbl
            // 
            this.origNumPagesLbl.AutoSize = true;
            this.origNumPagesLbl.Location = new System.Drawing.Point(121, 4);
            this.origNumPagesLbl.Name = "origNumPagesLbl";
            this.origNumPagesLbl.Size = new System.Drawing.Size(0, 13);
            this.origNumPagesLbl.TabIndex = 2;
            // 
            // label2
            // 
            this.label2.Location = new System.Drawing.Point(0, 33);
            this.label2.Name = "label2";
            this.label2.Size = new System.Drawing.Size(100, 23);
            this.label2.TabIndex = 3;
            this.label2.Text = "Final # of pages:";
            // 
            // finalNumPagesLbl
            // 
            this.finalNumPagesLbl.Anchor = ((System.Windows.Forms.AnchorStyles)(((System.Windows.Forms.AnchorStyles.Top | System.Windows.Forms.AnchorStyles.Bottom)
                        | System.Windows.Forms.AnchorStyles.Left)));
            this.finalNumPagesLbl.AutoSize = true;
            this.finalNumPagesLbl.Location = new System.Drawing.Point(121, 33);
            this.finalNumPagesLbl.Name = "finalNumPagesLbl";
            this.finalNumPagesLbl.Size = new System.Drawing.Size(0, 13);
            this.finalNumPagesLbl.TabIndex = 4;
            // 
            // label1
            // 
            this.label1.Enabled = false;
            this.label1.Location = new System.Drawing.Point(0, 4);
            this.label1.Name = "label1";
            this.label1.Size = new System.Drawing.Size(115, 29);
            this.label1.TabIndex = 5;
            this.label1.Text = "Original # of pages:";
            // 
            // label3
            // 
            this.label3.AutoSize = true;
            this.label3.Location = new System.Drawing.Point(0, 59);
            this.label3.Name = "label3";
            this.label3.Size = new System.Drawing.Size(74, 13);
            this.label3.TabIndex = 6;
            this.label3.Text = "Pages Saved:";
            // 
            // savedPagesLbl
            // 
            this.savedPagesLbl.AutoSize = true;
            this.savedPagesLbl.Location = new System.Drawing.Point(121, 59);
            this.savedPagesLbl.Name = "savedPagesLbl";
            this.savedPagesLbl.Size = new System.Drawing.Size(0, 13);
            this.savedPagesLbl.TabIndex = 7;
            // 
            // printPreviewBtn
            // 
            this.printPreviewBtn.Location = new System.Drawing.Point(3, 209);
            this.printPreviewBtn.Name = "printPreviewBtn";
            this.printPreviewBtn.Size = new System.Drawing.Size(84, 23);
            this.printPreviewBtn.TabIndex = 8;
            this.printPreviewBtn.Text = "print preview";
            this.printPreviewBtn.UseVisualStyleBackColor = true;
            this.printPreviewBtn.Click += new System.EventHandler(this.printPreviewBtn_Click);
            // 
            // printerList
            // 
            this.printerList.FormattingEnabled = true;
            this.printerList.Location = new System.Drawing.Point(3, 152);
            this.printerList.Name = "printerList";
            this.printerList.Size = new System.Drawing.Size(176, 21);
            this.printerList.TabIndex = 9;
            // 
            // numCopiesUpDwn
            // 
            this.numCopiesUpDwn.Font = new System.Drawing.Font("Microsoft Sans Serif", 8.5F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.numCopiesUpDwn.Location = new System.Drawing.Point(124, 102);
            this.numCopiesUpDwn.Name = "numCopiesUpDwn";
            this.numCopiesUpDwn.Size = new System.Drawing.Size(54, 20);
            this.numCopiesUpDwn.TabIndex = 10;
            this.numCopiesUpDwn.TextAlign = System.Windows.Forms.HorizontalAlignment.Center;
            this.numCopiesUpDwn.Value = new decimal(new int[] {
            1,
            0,
            0,
            0});
            // 
            // printeco_excel_usercontrol
            // 
            this.AutoScaleDimensions = new System.Drawing.SizeF(6F, 13F);
            this.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font;
            this.BackColor = System.Drawing.Color.LightGray;
            this.Controls.Add(this.numCopiesUpDwn);
            this.Controls.Add(this.printerList);
            this.Controls.Add(this.printPreviewBtn);
            this.Controls.Add(this.savedPagesLbl);
            this.Controls.Add(this.label3);
            this.Controls.Add(this.label1);
            this.Controls.Add(this.finalNumPagesLbl);
            this.Controls.Add(this.label2);
            this.Controls.Add(this.origNumPagesLbl);
            this.Controls.Add(this.printBtn);
            this.Name = "printeco_excel_usercontrol";
            this.Size = new System.Drawing.Size(218, 500);
            ((System.ComponentModel.ISupportInitialize)(this.numCopiesUpDwn)).EndInit();
            this.ResumeLayout(false);
            this.PerformLayout();

        }

        #endregion

        private System.Windows.Forms.Button printBtn;
        private System.Windows.Forms.Label origNumPagesLbl;
        private System.Windows.Forms.Label label2;
        private System.Windows.Forms.Label finalNumPagesLbl;
        private System.Windows.Forms.Label label1;
        private System.Windows.Forms.Label label3;
        private System.Windows.Forms.Label savedPagesLbl;
        private System.Windows.Forms.Button printPreviewBtn;
        private System.Windows.Forms.ComboBox printerList;
        private System.Windows.Forms.NumericUpDown numCopiesUpDwn;
    }
}
