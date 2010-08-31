using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Windows.Forms;
using System.Drawing.Printing;

namespace printeco_exceladdin
{
    public partial class ecoPrintDialog : Form
    {
        public ecoPrintDialog()
        {
            InitializeComponent();

            setupPrinterList();
        }



        // Creates a list of printers for the user to select
        private void setupPrinterList()
        {
            // Set up printer list
            PrintDocument prtdoc = new PrintDocument();
            string strDefaultPrinter = prtdoc.PrinterSettings.PrinterName;
            foreach (String strPrinter in PrinterSettings.InstalledPrinters)
            {
                printerList.Items.Add(strPrinter);
                if (strPrinter == strDefaultPrinter)
                {
                    printerList.SelectedIndex = printerList.Items.IndexOf(strPrinter);
                }
            }
        }

        // Handle event when user clicks "Print" in print dialog box
        private void dialogPrintBtn_Click(object sender, EventArgs e)
        {

            // Get the name of the currently selected printer
            int selectedIndex = printerList.SelectedIndex;
            printeco_exceladdin.AddinModule.printerToUse = printerList.Items[selectedIndex].ToString();

            // Tell add-in module that user has clicked "print"
            printeco_exceladdin.AddinModule.printThis = true;
            this.Close();
        }




    }
}
