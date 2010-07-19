using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Windows.Forms;
using System.Xml.Linq;
using Excel = Microsoft.Office.Interop.Excel;
using Office = Microsoft.Office.Core;
using Microsoft.Office.Tools.Excel;

namespace excel_plugin
{
    public partial class ThisAddIn
    {
        // Declare the menu variable at the class level.
        private Office.CommandBarButton menuCommand;
        private string menuTag = "A unique tag";

        #region Create Menu Command
        // If the menu already exists, remove it.
        private void CheckIfMenuBarExists()
        {
            try
            {
                Office.CommandBarPopup foundMenu = (Office.CommandBarPopup)
                    this.Application.CommandBars.ActiveMenuBar.FindControl(
                    Office.MsoControlType.msoControlPopup, System.Type.Missing, menuTag, true, true);

                if (foundMenu != null)
                {
                    foundMenu.Delete(true);
                }
            }
            catch (Exception ex)
            {
                MessageBox.Show(ex.Message);
            }
        }


        // Create the menu, if it does not exist.
        private void AddMenuBar()
        {
            try
            {
                Office.CommandBarPopup cmdBarControl = null;
                Office.CommandBar menubar = (Office.CommandBar)Application.CommandBars.ActiveMenuBar;
                int controlCount = menubar.Controls.Count;
                string menuCaption = "PrintEco";

                // Add the menu.
                cmdBarControl = (Office.CommandBarPopup)menubar.Controls.Add(
                    Office.MsoControlType.msoControlPopup, missing, missing, controlCount, true);

                if (cmdBarControl != null)
                {
                    cmdBarControl.Caption = menuCaption;

                    // Add the menu command.
                    menuCommand = (Office.CommandBarButton)cmdBarControl.Controls.Add(
                        Office.MsoControlType.msoControlButton, missing, missing, missing, true);

                    menuCommand.Caption = "PrintEcotize";
                    menuCommand.Tag = "PrintEcotize";
                    menuCommand.FaceId = 65;

                    menuCommand.Click += new Microsoft.Office.Core._CommandBarButtonEvents_ClickEventHandler(
                        menuCommand_Click);
                }
            }
            catch (Exception e)
            {
                MessageBox.Show(e.Message);
            }
        }

        #endregion

        // Format page when the menu is clicked.
        private void menuCommand_Click(Microsoft.Office.Core.CommandBarButton Ctrl, ref bool CancelDefault)
        {
            Excel.Workbook activeWorkbook = ((Excel.Workbook)Application.ActiveWorkbook);
            Excel.Worksheet activeWorksheet = ((Excel.Worksheet)Application.ActiveSheet);


            int initialNumPages = 0;
            // TODO: verify if all worksheets should be optimized
            // Do this if only the active worksheet needs to be be optimized
            initialNumPages = activeWorksheet.PageSetup.Pages.Count;
            MessageBox.Show("Before optimization: " + initialNumPages + " pages");

            // Manually set margins to "narrow" setting in Excel
            activeWorksheet.PageSetup.LeftMargin = Application.InchesToPoints(.25);
            activeWorksheet.PageSetup.TopMargin = Application.InchesToPoints(.75);
            activeWorksheet.PageSetup.RightMargin = Application.InchesToPoints(.25);
            activeWorksheet.PageSetup.BottomMargin = Application.InchesToPoints(.75);

            // Set zoom level to 90% if above 90
            if (activeWorksheet.PageSetup.Zoom > 90)
            {
                activeWorksheet.PageSetup.Zoom = 90;
            }


            // Compute which orientation (landscape or portrait) results in fewer pages 
            activeWorksheet.PageSetup.Orientation = Excel.XlPageOrientation.xlLandscape;
            int finalNumPages = activeWorksheet.PageSetup.Pages.Count;
            activeWorksheet.PageSetup.Orientation = Excel.XlPageOrientation.xlPortrait;
            if (activeWorksheet.PageSetup.Pages.Count > finalNumPages)
                activeWorksheet.PageSetup.Orientation = Excel.XlPageOrientation.xlLandscape;


            // Do this if all worksheets need to be optimized
            /** foreach (Excel.Worksheet sheet in allSheets)
             {
                 numPages += sheet.PageSetup.Pages.Count;
                 sheet.PageSetup.LeftMargin = Application.InchesToPoints(.25);
                 sheet.PageSetup.TopMargin = Application.InchesToPoints(.75);
                 sheet.PageSetup.RightMargin = Application.InchesToPoints(.25);
                 sheet.PageSetup.BottomMargin = Application.InchesToPoints(.75);
                 sheet.PageSetup.Orientation = Excel.XlPageOrientation.xlLandscape;
                 sheet.PageSetup.Zoom = 90;
                 //sheet.PrintPreview();
             }
             **/

            finalNumPages = activeWorksheet.PageSetup.Pages.Count;
            MessageBox.Show("After optimization: " + finalNumPages + " pages");
            activeWorkbook.PrintPreview();
        }

        private void ThisAddIn_Startup(object sender, System.EventArgs e)
        {
            CheckIfMenuBarExists();
            AddMenuBar();
        }

        private void ThisAddIn_Shutdown(object sender, System.EventArgs e)
        {
        }
        #region VSTO generated code

        /// <summary>
        /// Required method for Designer support - do not modify
        /// the contents of this method with the code editor.
        /// </summary>
        private void InternalStartup()
        {
            this.Startup += new System.EventHandler(ThisAddIn_Startup);
            this.Shutdown += new System.EventHandler(ThisAddIn_Shutdown);
        }
        
        #endregion
    }
}
