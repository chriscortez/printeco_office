using System;
using System.Runtime.InteropServices;
using System.ComponentModel;
using System.Windows.Forms;
using Excel = Microsoft.Office.Interop.Excel;

namespace printeco_exceladdin
{
    /// <summary>
    ///   Add-in Express Add-in Module
    /// </summary>
    [GuidAttribute("686F9CBF-3BFC-42FB-9E65-F2851A8367E9"), ProgId("printeco_exceladdin.AddinModule")]
    public class AddinModule : AddinExpress.MSO.ADXAddinModule
    {

        public AddinModule()
        {
            InitializeComponent();
            // Please add any initialization code to the AddinInitialize event handler

            PrintEcoButton.OnClick += new AddinExpress.MSO.ADXRibbonOnAction_EventHandler(PrintEcoButton_OnClick);
        }




        // My variables
        private double origLeftMargin, origTopMargin, origRightMargin, origBotMargin;
        private Excel.XlPageOrientation origOrientation;
        private int origZoom;
        private bool isOptimized = false;
        private int initialNumPages, finalNumPages;

        // AddinExpress components
        private AddinExpress.MSO.ADXRibbonOfficeMenu adxRibbonOfficeMenu1;
        private AddinExpress.MSO.ADXRibbonButton PrintEcoButton;


        void PrintEcoButton_OnClick(object sender, AddinExpress.MSO.IRibbonControl control, bool pressed)
        {
            Excel.Worksheet activeWorksheet = ((Excel.Worksheet)ExcelApp.ActiveSheet);

            getOrigSettings();
            optimizeDocument();

            // Reset print preview ribbon to display optimized results
            //printPreviewTab.ribbon.Invalidate();

            // First show user a print preview, then show custom task pane when preview is closed
            activeWorksheet.PrintPreview(true);
        }


        // Store original page settings
        private void getOrigSettings()
        {
            Excel.Worksheet activeWorksheet = ((Excel.Worksheet)ExcelApp.ActiveSheet);

            // Get original margin settings
            origLeftMargin = activeWorksheet.PageSetup.LeftMargin;
            origTopMargin = activeWorksheet.PageSetup.TopMargin;
            origRightMargin = activeWorksheet.PageSetup.RightMargin;
            origBotMargin = activeWorksheet.PageSetup.BottomMargin;

            // Get original orientation
            origOrientation = activeWorksheet.PageSetup.Orientation;


            // Get original font

            // Get original view information
            origZoom = (int)activeWorksheet.PageSetup.Zoom;
        }

        private void optimizeDocument()
        {
            if (!isOptimized)
            {
                Excel.Worksheet activeWorksheet = ((Excel.Worksheet)ExcelApp.ActiveSheet);
                initialNumPages = 0;

                // Do this if only the active worksheet needs to be be optimized
                initialNumPages = activeWorksheet.PageSetup.Pages.Count;

                // Begin optimization...
                setMargins();
                setZoomLevel();
                computeOrientation();

                finalNumPages = activeWorksheet.PageSetup.Pages.Count;
                MessageBox.Show("Final num pages = " + finalNumPages);
                isOptimized = true;
            }
        }

        // Manually set margins of page
        private void setMargins()
        {
            Excel.Worksheet activeWorksheet = ((Excel.Worksheet)ExcelApp.ActiveSheet);

            // These values correspond with the "narrow" in Excel
            activeWorksheet.PageSetup.LeftMargin = ExcelApp.InchesToPoints(.25);
            activeWorksheet.PageSetup.TopMargin = ExcelApp.InchesToPoints(.75);
            activeWorksheet.PageSetup.RightMargin = ExcelApp.InchesToPoints(.25);
            activeWorksheet.PageSetup.BottomMargin = ExcelApp.InchesToPoints(.75);
        }

        // Manually set zoom level
        private void setZoomLevel()
        {
            Excel.Worksheet activeWorksheet = ((Excel.Worksheet)ExcelApp.ActiveSheet);

            // Set zoom level to 90% if above 90, leave alone otherwise
            if ((int)activeWorksheet.PageSetup.Zoom > 90)
            {
                activeWorksheet.PageSetup.Zoom = 90;
            }
        }

        // Compute which orientation (landscape or portrait) results in fewer pages
        // and set accordingly
        private void computeOrientation()
        {
            Excel.Worksheet activeWorksheet = ((Excel.Worksheet)ExcelApp.ActiveSheet);

            activeWorksheet.PageSetup.Orientation = Excel.XlPageOrientation.xlLandscape;
            int numPages = activeWorksheet.PageSetup.Pages.Count;
            activeWorksheet.PageSetup.Orientation = Excel.XlPageOrientation.xlPortrait;
            if (activeWorksheet.PageSetup.Pages.Count > numPages)
                activeWorksheet.PageSetup.Orientation = Excel.XlPageOrientation.xlLandscape;
        }
 


        #region Component Designer generated code
        /// <summary>
        /// Required by designer
        /// </summary>
        private System.ComponentModel.IContainer components;
 
        /// <summary>
        /// Required by designer support - do not modify
        /// the following method
        /// </summary>
        private void InitializeComponent()
        {
            this.components = new System.ComponentModel.Container();
            this.adxRibbonOfficeMenu1 = new AddinExpress.MSO.ADXRibbonOfficeMenu(this.components);
            this.PrintEcoButton = new AddinExpress.MSO.ADXRibbonButton(this.components);

            // 
            // adxRibbonOfficeMenu1
            // 
            this.adxRibbonOfficeMenu1.Controls.Add(this.PrintEcoButton);
            this.adxRibbonOfficeMenu1.Ribbons = AddinExpress.MSO.ADXRibbons.msrExcelWorkbook;
            // 
            // PrintEcoButton
            // 
            this.PrintEcoButton.Caption = "PrintEco";
            this.PrintEcoButton.Id = "adxRibbonButton_de8117a3c8d74d5b9214d749c5468885";
            this.PrintEcoButton.ImageTransparentColor = System.Drawing.Color.Transparent;
            this.PrintEcoButton.InsertBeforeIdMso = "FilePrintMenu";
            this.PrintEcoButton.Ribbons = AddinExpress.MSO.ADXRibbons.msrExcelWorkbook;
            // 
            // AddinModule
            // 
            this.AddinName = "printeco_exceladdin";
            this.SupportedApps = AddinExpress.MSO.ADXOfficeHostApp.ohaExcel;

        }
        #endregion
 
        #region Add-in Express automatic code
 
        // Required by Add-in Express - do not modify
        // the methods within this region
 
        public override System.ComponentModel.IContainer GetContainer()
        {
            if (components == null)
                components = new System.ComponentModel.Container();
            return components;
        }
 
        [ComRegisterFunctionAttribute]
        public static void AddinRegister(Type t)
        {
            AddinExpress.MSO.ADXAddinModule.ADXRegister(t);
        }
 
        [ComUnregisterFunctionAttribute]
        public static void AddinUnregister(Type t)
        {
            AddinExpress.MSO.ADXAddinModule.ADXUnregister(t);
        }
 
        public override void UninstallControls()
        {
            base.UninstallControls();
        }

        #endregion

        public Excel._Application ExcelApp
        {
            get
            {
                return (HostApplication as Excel._Application);
            }
        }
    }
}

