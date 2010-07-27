using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Drawing;
using System.Data;
using System.Linq;
using System.Text;
using System.Windows.Forms;
using Office = Microsoft.Office.Core;
using Excel = Microsoft.Office.Interop.Excel;

namespace excel_plugin
{
    public partial class printeco_excel_usercontrol : UserControl
    {
        public printeco_excel_usercontrol()
        {
            InitializeComponent();
            
        }

        private void button1_Click(object sender, EventArgs e)
        {
        }


        public void updatePageCounts(int initialNumPages, int finalNumPages)
        {
            origNumPagesLbl.Text = initialNumPages.ToString();
            finalNumPagesLbl.Text = finalNumPages.ToString();

        }


    }
}
