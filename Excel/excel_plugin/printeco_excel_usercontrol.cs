using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Drawing;
using System.Data;
using System.Linq;
using System.Text;
using System.Windows.Forms;
using System.Drawing.Printing;
using Office = Microsoft.Office.Core;
using Excel = Microsoft.Office.Interop.Excel;

namespace excel_plugin
{
    public partial class printeco_excel_usercontrol : UserControl
    {
        public printeco_excel_usercontrol()
        {
            InitializeComponent();

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

        // Update page count labels on task pane
        public void updatePageCounts(int initialNumPages, int finalNumPages)
        {
            origNumPagesLbl.Text = initialNumPages.ToString();
            finalNumPagesLbl.Text = finalNumPages.ToString();

            int pagesSaved = initialNumPages - finalNumPages;
            savedPagesLbl.Text = pagesSaved.ToString();
        }


        // Print the document
        private void printBtn_Click(object sender, EventArgs e)
        {
            
            Excel._Worksheet activeSheet = ((Excel._Worksheet)Globals.ThisAddIn.Application.ActiveSheet);
            
            // Get the name of the currently selected printer
            int selectedIndex = printerList.SelectedIndex;
            String printerToUse = printerList.Items[selectedIndex].ToString();

            // Get number of copies to print
            int numCopies = (int)numCopiesUpDwn.Value;

            // Uncomment to show print dialog box...
            // Globals.ThisAddIn.Application.Dialogs[Excel.XlBuiltInDialog.xlDialogPrint].Show();

            // ...or to print directly
            try
            {
                activeSheet.PrintOutEx(System.Type.Missing, System.Type.Missing, numCopies, false, printerToUse, System.Type.Missing, System.Type.Missing, System.Type.Missing, System.Type.Missing);
            }
            catch (Exception ex)
            {
                MessageBox.Show(ex.Message);
            }


            // used for reference
//            void PrintOut(
//    Object From,
//    Object To,
//    Object Copies,
//    Object Preview,
//    Object ActivePrinter,
//    Object PrintToFile,
//    Object Collate,
//    Object PrToFileName
//)
        }

        // Show a print preview of the document
        private void printPreviewBtn_Click(object sender, EventArgs e)
        {
            Globals.ThisAddIn.Application.ActiveSheet.PrintPreview();
        }




    }
}
