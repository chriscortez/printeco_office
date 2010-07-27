using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Windows.Forms;
using System.Xml.Linq;
using MySql.Data.MySqlClient;
using Excel = Microsoft.Office.Interop.Excel;
using Office = Microsoft.Office.Core;
using Microsoft.Office.Tools.Excel;

namespace excel_plugin
{
    public partial class ThisAddIn
    {
        // Declare user control for custom task pane
        private printeco_excel_usercontrol myUserControl;
        private Microsoft.Office.Tools.CustomTaskPane myCustomTaskPane;

        // Declare the menu variable at the class level.
        private Office.CommandBarButton menuCommand;
        private string menuTag = "A unique tag";

        public int initialNumPages;
        public int finalNumPages;


        private void ThisAddIn_Startup(object sender, System.EventArgs e)
        {
            myUserControl = new printeco_excel_usercontrol();
            myCustomTaskPane = this.CustomTaskPanes.Add(myUserControl, "My Task Pane");
            myCustomTaskPane.DockPosition = Office.MsoCTPDockPosition.msoCTPDockPositionLeft;

            CheckIfMenuBarExists();
            AddMenuBar();
            // addFileCommand();
        }

        private void ThisAddIn_Shutdown(object sender, System.EventArgs e)
        {
        }



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

        // Add "Printeco..." to file commands
        private void addFileCommand()
        {
            Office.CommandBarControl cmdCtrl = 
                Application.CommandBars.FindControl(System.Type.Missing, 4, System.Type.Missing, System.Type.Missing);



            Application.CommandBars.DisableCustomize = false;
            string boolean = (Application.CommandBars.DisableCustomize) ? "disabled" : "enabled";
            MessageBox.Show(boolean);
            Office.CommandBar menuBar = Application.CommandBars["File"];
            if(menuBar != null)
                MessageBox.Show(menuBar.accName);
            Application.CommandBars.DisableCustomize = false;


            Office.CommandBarControl cmdBarControl =
                menuBar.Controls.Add(Office.MsoControlType.msoControlPopup, System.Type.Missing, System.Type.Missing, 2, true);

            cmdBarControl.Move(Application.CommandBars["File"], 3);
            cmdBarControl.Caption = "PrintEco";



            MessageBox.Show(cmdCtrl.accName);
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

            initialNumPages = 0;
            
            // Do this if only the active worksheet needs to be be optimized
            initialNumPages = activeWorksheet.PageSetup.Pages.Count;
            MessageBox.Show("Before optimization: " + initialNumPages + " pages");

            // Begin optimization...
            setMargins();
            setZoomLevel();
            computeOrientation();

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
            //activeWorkbook.PrintPreview();

            //updateDatabase(initialNumPages, finalNumPages);

            myCustomTaskPane.Visible = true;
            myUserControl.updatePageCounts(initialNumPages, finalNumPages);
        }


        private void updateDatabase(int initialNumPages, int finalNumpages)
        {
            //string MyConString = "SERVER=http://db2480.perfora.net;" +
            //      "DATABASE=db333199449;" +
            //      "UID=dbo333199449;" +
            //      "PASSWORD=Napra888;";

           //  Local host works fine
            string MyConString = "SERVER=localhost;" +
                "DATABASE=printeco_data;" +
                "UID=root;" +
                "PASSWORD=;";

            MySqlConnection connection = new MySqlConnection(MyConString);
            MySqlCommand command = connection.CreateCommand();
            int pagesSaved = initialNumPages - finalNumpages;
            command.CommandText = "INSERT INTO job (workstation_id , initial_pages , pages_saved  , company_id) VALUES ('" + System.Environment.MachineName + "'," + initialNumPages + ", " +
                                                    pagesSaved + ",1)";
            try
            {
                connection.Open();
            }
            catch (MySqlException mySqlError)
            {
                MessageBox.Show(mySqlError.Message);
            }
            command.ExecuteNonQuery();
            connection.Close();
        }



        // Manually set margins of page
        private void setMargins()
        {
            Excel.Worksheet activeWorksheet = ((Excel.Worksheet)Application.ActiveSheet);

            // These values correspond with the "narrow" in Excel
            activeWorksheet.PageSetup.LeftMargin = Application.InchesToPoints(.25);
            activeWorksheet.PageSetup.TopMargin = Application.InchesToPoints(.75);
            activeWorksheet.PageSetup.RightMargin = Application.InchesToPoints(.25);
            activeWorksheet.PageSetup.BottomMargin = Application.InchesToPoints(.75);
        }

        // Manually set zoom level
        private void setZoomLevel()
        {
            Excel.Worksheet activeWorksheet = ((Excel.Worksheet)Application.ActiveSheet);

            // Set zoom level to 90% if above 90, leave alone otherwise
            if (activeWorksheet.PageSetup.Zoom > 90)
            {
                activeWorksheet.PageSetup.Zoom = 90;
            }

        }

        // Compute which orientation (landscape or portrait) results in fewer pages
        // and set accordingly
        private void computeOrientation()
        {

            Excel.Worksheet activeWorksheet = ((Excel.Worksheet)Application.ActiveSheet);

            activeWorksheet.PageSetup.Orientation = Excel.XlPageOrientation.xlLandscape;
            int numPages = activeWorksheet.PageSetup.Pages.Count;
            activeWorksheet.PageSetup.Orientation = Excel.XlPageOrientation.xlPortrait;
            if (activeWorksheet.PageSetup.Pages.Count > numPages)
                activeWorksheet.PageSetup.Orientation = Excel.XlPageOrientation.xlLandscape;
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
