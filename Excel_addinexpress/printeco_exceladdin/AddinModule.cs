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
            myPrintDialog = new ecoPrintDialog();

            PrintEcoButton.OnClick += new AddinExpress.MSO.ADXRibbonOnAction_EventHandler(PrintEcoButton_OnClick);
            toggleOriginalBtn.OnClick += new AddinExpress.MSO.ADXRibbonOnAction_EventHandler(toggleOriginalBtn_OnClick);
            closeBtn.OnClick += new AddinExpress.MSO.ADXRibbonOnAction_EventHandler(closeBtn_OnClick);
            printBtn.OnClick += new AddinExpress.MSO.ADXRibbonOnAction_EventHandler(printBtn_OnClick);
         //   pageRemovalBtn.OnClick += new AddinExpress.MSO.ADXRibbonOnAction_EventHandler(pageRemovalBtn_OnClick);
        }

        void pageRemovalBtn_OnClick(object sender, AddinExpress.MSO.IRibbonControl control, bool pressed)
        {
            switcher = "custom remove";
            SendKeys.Send("{ESC}");
        }

        private void customPageRemoval()
        {
            Excel.Worksheet activeWorksheet = ((Excel.Worksheet)ExcelApp.ActiveSheet);

            // Change view to 'page layout' view to allow user to customize
            ExcelApp.ActiveWindow.View = Excel.XlWindowView.xlPageLayoutView;


            Excel.HPageBreaks HPageBreaks = activeWorksheet.HPageBreaks;
            Excel.VPageBreaks VPageBreaks = activeWorksheet.VPageBreaks;
            Excel.Range endBoundary;

            // Test selection
            Excel.Range beginBoundary = HPageBreaks[1].Location;
            endBoundary = HPageBreaks[2].Location;
            Excel.Range theRange = activeWorksheet.get_Range(beginBoundary, endBoundary);

            Excel.Range theOtherRange = activeWorksheet.get_Range("A1", VPageBreaks[1].Location);

            Excel.Range theRealRange = ExcelApp.Intersect(theRange, theOtherRange);



            theRealRange.BorderAround(System.Type.Missing, Excel.XlBorderWeight.xlThick, Excel.XlColorIndex.xlColorIndexAutomatic, Excel.XlRgbColor.rgbDarkRed);




            //for(int i = 1; i < pageBreaks.Count; i++)
            //{
            //    endBoundary = pageBreaks[i].Location;

            //    // Create range by selecting all cells within boundaries
            //    activeWorksheet.get_Range(beginBoundary, endBoundary).Select();
            //}


            // Disable cell selection
            //activeWorksheet.EnableSelection = Excel.XlEnableSelection.xlNoSelection;

            // Return Excel to normal settings
            activeWorksheet.EnableSelection = Excel.XlEnableSelection.xlUnlockedCells;
        }

        // Handle event when user clicks on "Print" button in print preview
        void printBtn_OnClick(object sender, AddinExpress.MSO.IRibbonControl control, bool pressed)
        {
            myPrintDialog.ShowDialog();

            // Print document if user wants to
            if (printThis)
            {
               
                switcher = "print";
                printThis = false;

                // Exit print preview
                SendKeys.Send("{ESC}");
            }
            else
            {
                switcher = "";
            }

        }

        // Send document to printer
        private void printDocument()
        {

            Excel.Worksheet activeWorksheet = ((Excel.Worksheet)ExcelApp.ActiveSheet);
            
            // Get number of copies to print
            int numCopies = 1;

            // Print document
            try
            {
                activeWorksheet.PrintOutEx(System.Type.Missing, System.Type.Missing, System.Type.Missing, false, printerToUse, System.Type.Missing, System.Type.Missing, System.Type.Missing, System.Type.Missing);

                if (finalNumPages != 0)
                {
                    // For each copy of job, send to database
                    for (int i = 0; i < numCopies; i++)
                    {
                        // Update database
                    }
                }
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

        // My variables
        private double origLeftMargin, origTopMargin, origRightMargin, origBotMargin;
        private Excel.XlPageOrientation origOrientation;
        private int origZoom;
        private bool isOptimized = false;
        private int initialNumPages, finalNumPages;
        private string switcher;
        private ecoPrintDialog myPrintDialog;
        private Excel.Range[] rangesToRemove;
        public static bool printThis;
        public static string printerToUse;

        // AddinExpress components
        private AddinExpress.MSO.ADXRibbonOfficeMenu adxRibbonOfficeMenu1;
        private AddinExpress.MSO.ADXRibbonTab adxRibbonTab1;
        private AddinExpress.MSO.ADXRibbonGroup adxRibbonGroup1;
        private AddinExpress.MSO.ADXRibbonGroup adxRibbonGroup2;
        private AddinExpress.MSO.ADXRibbonGroup adxRibbonGroup3;
        private AddinExpress.MSO.ADXRibbonGroup ecoRibbonGroup;
        private AddinExpress.MSO.ADXRibbonButton toggleOriginalBtn;
        private AddinExpress.MSO.ADXRibbonButton closeBtn;
        private AddinExpress.MSO.ADXRibbonBox adxRibbonBox1;
        private AddinExpress.MSO.ADXRibbonBox adxRibbonBox2;
        private ImageList imageList1;
        private AddinExpress.MSO.ADXRibbonBox adxRibbonBox3;
        private AddinExpress.MSO.ADXRibbonButton printBtn;
        private AddinExpress.MSO.ADXRibbonBox adxRibbonBox4;
        private AddinExpress.MSO.ADXRibbonButton pageSetupBtn;
        private AddinExpress.MSO.ADXRibbonBox adxRibbonBox5;
        private AddinExpress.MSO.ADXRibbonButton pageRemovalBtn;
        private AddinExpress.MSO.ADXRibbonBox adxRibbonBox6;
        private AddinExpress.MSO.ADXRibbonLabel adxRibbonLabel1;
        private AddinExpress.MSO.ADXRibbonLabel numSavedLbl;
        private AddinExpress.MSO.ADXRibbonButton PrintEcoButton;

        // This event occurs when user clicks "PrintEco" in File menu
        void PrintEcoButton_OnClick(object sender, AddinExpress.MSO.IRibbonControl control, bool pressed)
        {
            Excel.Worksheet activeWorksheet = ((Excel.Worksheet)ExcelApp.ActiveSheet);

            getOrigSettings();
            optimizeDocument();

            updatePageLbl();

            // Reset print preview ribbon to display optimized results
            setupPrintPreviewRibbon();

            // Show user the custom print preview
            bool done = false;
            switcher = "";
            do
            {
                activeWorksheet.PrintPreview(true);
                switch (switcher)
                {
                    case "convert":
                        if (isOptimized)
                        {
                            resetOrigSettings();
                            updatePageLbl();
                        }
                        else
                        {
                            optimizeDocument();
                            updatePageLbl();
                        }
                        break;
                    case "finish":
                        resetOrigSettings();
                        done = true;
                        break;
                    case "print":
                        printDocument();
                        if(printThis)
                            done = true;
                        break;
                    case "custom remove":
                        customPageRemoval();
                        done = true;
                        break;
                }


            } while (done != true);
            resetPrintPreviewRibbon();

        }


        private void updatePageLbl()
        {
            if (isOptimized)
            {
                numSavedLbl.Caption = (initialNumPages - finalNumPages).ToString() + " pages.";
                numSavedLbl.ShowCaption = true;
            }
            else
            {
                numSavedLbl.Caption = "0 pages. Optimize document to maximize savings!";
            }
        }

        // Handle event when user toggles original and optimized document
        void toggleOriginalBtn_OnClick(object sender, AddinExpress.MSO.IRibbonControl control, bool pressed)
        {
            switcher = "convert";
            SendKeys.Send("{ESC}");
        }


        // Handle event when user closes print preview
        void closeBtn_OnClick(object sender, AddinExpress.MSO.IRibbonControl control, bool pressed)
        {
            switcher = "finish";
            SendKeys.Send("{ESC}");
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

        // Main optimization algorithm
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
                isOptimized = true;
            }
        }

        // Hide built-in controls and display custom controls in print preview
        private void setupPrintPreviewRibbon()
        {
            this.adxRibbonGroup1.Visible = false;
            this.adxRibbonGroup2.Visible = false;
            this.adxRibbonGroup3.Visible = false;
            this.ecoRibbonGroup.Visible = true;
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


        // Before closing or upon user's request, restore original settings
        public void resetOrigSettings()
        {
            if (isOptimized)
            {
                Excel.Worksheet activeWorksheet = ((Excel.Worksheet)ExcelApp.ActiveSheet);

                // Reset margins
                activeWorksheet.PageSetup.LeftMargin = origLeftMargin;
                activeWorksheet.PageSetup.TopMargin = origTopMargin;
                activeWorksheet.PageSetup.RightMargin = origRightMargin;
                activeWorksheet.PageSetup.BottomMargin = origBotMargin;


                // Reset orientation
                activeWorksheet.PageSetup.Orientation = origOrientation;

                // Reset font


                // Reset zoom
                activeWorksheet.PageSetup.Zoom = origZoom;

                // Reset page counts
                initialNumPages = 0;
                finalNumPages = 0;

                isOptimized = false;
            }
        }

        // Return print preview to its original settings
        private void resetPrintPreviewRibbon()
        {
            this.adxRibbonGroup1.Visible = true;
            this.adxRibbonGroup2.Visible = true;
            this.adxRibbonGroup3.Visible = true;
            this.ecoRibbonGroup.Visible = false;
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
            System.ComponentModel.ComponentResourceManager resources = new System.ComponentModel.ComponentResourceManager(typeof(AddinModule));
            this.adxRibbonOfficeMenu1 = new AddinExpress.MSO.ADXRibbonOfficeMenu(this.components);
            this.PrintEcoButton = new AddinExpress.MSO.ADXRibbonButton(this.components);
            this.adxRibbonTab1 = new AddinExpress.MSO.ADXRibbonTab(this.components);
            this.adxRibbonGroup1 = new AddinExpress.MSO.ADXRibbonGroup(this.components);
            this.adxRibbonGroup2 = new AddinExpress.MSO.ADXRibbonGroup(this.components);
            this.adxRibbonGroup3 = new AddinExpress.MSO.ADXRibbonGroup(this.components);
            this.ecoRibbonGroup = new AddinExpress.MSO.ADXRibbonGroup(this.components);
            this.adxRibbonBox2 = new AddinExpress.MSO.ADXRibbonBox(this.components);
            this.closeBtn = new AddinExpress.MSO.ADXRibbonButton(this.components);
            this.adxRibbonBox3 = new AddinExpress.MSO.ADXRibbonBox(this.components);
            this.printBtn = new AddinExpress.MSO.ADXRibbonButton(this.components);
            this.imageList1 = new System.Windows.Forms.ImageList(this.components);
            this.adxRibbonBox4 = new AddinExpress.MSO.ADXRibbonBox(this.components);
            this.pageSetupBtn = new AddinExpress.MSO.ADXRibbonButton(this.components);
            this.adxRibbonBox5 = new AddinExpress.MSO.ADXRibbonBox(this.components);
            this.pageRemovalBtn = new AddinExpress.MSO.ADXRibbonButton(this.components);
            this.adxRibbonBox1 = new AddinExpress.MSO.ADXRibbonBox(this.components);
            this.toggleOriginalBtn = new AddinExpress.MSO.ADXRibbonButton(this.components);
            this.adxRibbonLabel1 = new AddinExpress.MSO.ADXRibbonLabel(this.components);
            this.adxRibbonBox6 = new AddinExpress.MSO.ADXRibbonBox(this.components);
            this.numSavedLbl = new AddinExpress.MSO.ADXRibbonLabel(this.components);
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
            // adxRibbonTab1
            // 
            this.adxRibbonTab1.Caption = "PrintEco";
            this.adxRibbonTab1.Controls.Add(this.adxRibbonGroup1);
            this.adxRibbonTab1.Controls.Add(this.adxRibbonGroup2);
            this.adxRibbonTab1.Controls.Add(this.adxRibbonGroup3);
            this.adxRibbonTab1.Controls.Add(this.ecoRibbonGroup);
            this.adxRibbonTab1.Id = "adxRibbonTab_0156a0f713aa42f58bc151b7052e7cb4";
            this.adxRibbonTab1.IdMso = "TabPrintPreview";
            this.adxRibbonTab1.Ribbons = AddinExpress.MSO.ADXRibbons.msrExcelWorkbook;
            // 
            // adxRibbonGroup1
            // 
            this.adxRibbonGroup1.Caption = "adxRibbonGroup1";
            this.adxRibbonGroup1.Id = "adxRibbonGroup_d24f81399ec241aa903488ae69a42061";
            this.adxRibbonGroup1.IdMso = "GroupPrintPreviewPrint";
            this.adxRibbonGroup1.ImageTransparentColor = System.Drawing.Color.Transparent;
            this.adxRibbonGroup1.Ribbons = AddinExpress.MSO.ADXRibbons.msrExcelWorkbook;
            // 
            // adxRibbonGroup2
            // 
            this.adxRibbonGroup2.Caption = "adxRibbonGroup2";
            this.adxRibbonGroup2.Id = "adxRibbonGroup_67d28edc32c549c892f59f1403d12d34";
            this.adxRibbonGroup2.IdMso = "GroupPrintPreviewZoom";
            this.adxRibbonGroup2.ImageTransparentColor = System.Drawing.Color.Transparent;
            this.adxRibbonGroup2.Ribbons = AddinExpress.MSO.ADXRibbons.msrExcelWorkbook;
            // 
            // adxRibbonGroup3
            // 
            this.adxRibbonGroup3.Caption = "adxRibbonGroup3";
            this.adxRibbonGroup3.Id = "adxRibbonGroup_c19543e281b44145901e76be953d6b7e";
            this.adxRibbonGroup3.IdMso = "GroupPrintPreviewPreview";
            this.adxRibbonGroup3.ImageTransparentColor = System.Drawing.Color.Transparent;
            this.adxRibbonGroup3.Ribbons = AddinExpress.MSO.ADXRibbons.msrExcelWorkbook;
            // 
            // ecoRibbonGroup
            // 
            this.ecoRibbonGroup.Caption = "PrintEco";
            this.ecoRibbonGroup.Controls.Add(this.adxRibbonBox6);
            this.ecoRibbonGroup.Controls.Add(this.adxRibbonBox1);
            this.ecoRibbonGroup.Controls.Add(this.adxRibbonBox3);
            this.ecoRibbonGroup.Controls.Add(this.adxRibbonBox4);
            this.ecoRibbonGroup.Controls.Add(this.adxRibbonBox5);
            this.ecoRibbonGroup.Controls.Add(this.adxRibbonBox2);
            this.ecoRibbonGroup.Id = "adxRibbonGroup_2c4bf74981c144eb8bfada8799e5386d";
            this.ecoRibbonGroup.ImageTransparentColor = System.Drawing.Color.Transparent;
            this.ecoRibbonGroup.Ribbons = AddinExpress.MSO.ADXRibbons.msrExcelWorkbook;
            // 
            // adxRibbonBox2
            // 
            this.adxRibbonBox2.BoxStyle = AddinExpress.MSO.ADXRibbonXBoxStyle.Vertical;
            this.adxRibbonBox2.Controls.Add(this.closeBtn);
            this.adxRibbonBox2.Id = "adxRibbonBox_0a7c59916a9245c3affb16afa81cd4c9";
            this.adxRibbonBox2.Ribbons = AddinExpress.MSO.ADXRibbons.msrExcelWorkbook;
            // 
            // closeBtn
            // 
            this.closeBtn.Caption = "Close Print Preview";
            this.closeBtn.Id = "adxRibbonButton_c507fc84ce944b17aff49109d181d782";
            this.closeBtn.ImageMso = "PrintPreviewClose";
            this.closeBtn.ImageTransparentColor = System.Drawing.Color.Transparent;
            this.closeBtn.Ribbons = AddinExpress.MSO.ADXRibbons.msrExcelWorkbook;
            this.closeBtn.Size = AddinExpress.MSO.ADXRibbonXControlSize.Large;
            // 
            // adxRibbonBox3
            // 
            this.adxRibbonBox3.Controls.Add(this.printBtn);
            this.adxRibbonBox3.Id = "adxRibbonBox_cd94a61109a64c8184393f11f0b87b26";
            this.adxRibbonBox3.Ribbons = AddinExpress.MSO.ADXRibbons.msrExcelWorkbook;
            // 
            // printBtn
            // 
            this.printBtn.Caption = "Print This Shit";
            this.printBtn.Id = "adxRibbonButton_a1c09b5ef3ab4f58b946872ab0a2cc6e";
            this.printBtn.Image = 0;
            this.printBtn.ImageList = this.imageList1;
            this.printBtn.ImageTransparentColor = System.Drawing.Color.Transparent;
            this.printBtn.Ribbons = AddinExpress.MSO.ADXRibbons.msrExcelWorkbook;
            this.printBtn.Size = AddinExpress.MSO.ADXRibbonXControlSize.Large;
            // 
            // imageList1
            // 
            this.imageList1.ImageStream = ((System.Windows.Forms.ImageListStreamer)(resources.GetObject("imageList1.ImageStream")));
            this.imageList1.TransparentColor = System.Drawing.Color.Transparent;
            this.imageList1.Images.SetKeyName(0, "AWESOME_FACE!!!.png");
            this.imageList1.Images.SetKeyName(1, "printecoPreviewLogo.png");
            // 
            // adxRibbonBox4
            // 
            this.adxRibbonBox4.Controls.Add(this.pageSetupBtn);
            this.adxRibbonBox4.Id = "adxRibbonBox_f63d70b212eb44eca0687e27528aa2aa";
            this.adxRibbonBox4.Ribbons = AddinExpress.MSO.ADXRibbons.msrExcelWorkbook;
            // 
            // pageSetupBtn
            // 
            this.pageSetupBtn.Caption = "Page Setup, Yo";
            this.pageSetupBtn.Id = "adxRibbonButton_6d691bb772ae48f08ee8df5415347f18";
            this.pageSetupBtn.IdMso = "PageSetupPageDialog";
            this.pageSetupBtn.Image = 0;
            this.pageSetupBtn.ImageList = this.imageList1;
            this.pageSetupBtn.ImageTransparentColor = System.Drawing.Color.Transparent;
            this.pageSetupBtn.Ribbons = AddinExpress.MSO.ADXRibbons.msrExcelWorkbook;
            this.pageSetupBtn.Size = AddinExpress.MSO.ADXRibbonXControlSize.Large;
            // 
            // adxRibbonBox5
            // 
            this.adxRibbonBox5.Controls.Add(this.pageRemovalBtn);
            this.adxRibbonBox5.Id = "adxRibbonBox_c90b2e1dbc7e40cea666d00bd16acbee";
            this.adxRibbonBox5.Ribbons = AddinExpress.MSO.ADXRibbons.msrExcelWorkbook;
            // 
            // pageRemovalBtn
            // 
            this.pageRemovalBtn.Caption = "Custom Page Removal";
            this.pageRemovalBtn.Id = "adxRibbonButton_897b34daca11404ebd4dcc6601e6c834";
            this.pageRemovalBtn.Image = 0;
            this.pageRemovalBtn.ImageList = this.imageList1;
            this.pageRemovalBtn.ImageTransparentColor = System.Drawing.Color.Transparent;
            this.pageRemovalBtn.Ribbons = AddinExpress.MSO.ADXRibbons.msrExcelWorkbook;
            this.pageRemovalBtn.Size = AddinExpress.MSO.ADXRibbonXControlSize.Large;
            // 
            // adxRibbonBox1
            // 
            this.adxRibbonBox1.Controls.Add(this.toggleOriginalBtn);
            this.adxRibbonBox1.Id = "adxRibbonBox_f451430ca97a49c88e3a6937eb1d5a98";
            this.adxRibbonBox1.Ribbons = AddinExpress.MSO.ADXRibbons.msrExcelWorkbook;
            // 
            // toggleOriginalBtn
            // 
            this.toggleOriginalBtn.Caption = "Toggle Original";
            this.toggleOriginalBtn.Id = "adxRibbonButton_87f9e0d12ffc406a856efc82fc7b6c69";
            this.toggleOriginalBtn.Image = 0;
            this.toggleOriginalBtn.ImageList = this.imageList1;
            this.toggleOriginalBtn.ImageTransparentColor = System.Drawing.Color.Transparent;
            this.toggleOriginalBtn.Ribbons = AddinExpress.MSO.ADXRibbons.msrExcelWorkbook;
            this.toggleOriginalBtn.Size = AddinExpress.MSO.ADXRibbonXControlSize.Large;
            this.toggleOriginalBtn.ToggleButton = true;
            // 
            // adxRibbonLabel1
            // 
            this.adxRibbonLabel1.Caption = "You saved:";
            this.adxRibbonLabel1.Id = "adxRibbonLabel_fd4be98b5f674b4e9a370e57044afd77";
            this.adxRibbonLabel1.Ribbons = AddinExpress.MSO.ADXRibbons.msrExcelWorkbook;
            // 
            // adxRibbonBox6
            // 
            this.adxRibbonBox6.BoxStyle = AddinExpress.MSO.ADXRibbonXBoxStyle.Vertical;
            this.adxRibbonBox6.Controls.Add(this.adxRibbonLabel1);
            this.adxRibbonBox6.Controls.Add(this.numSavedLbl);
            this.adxRibbonBox6.Id = "adxRibbonBox_c42f6c4fa79c4570b3ae5691519e5c17";
            this.adxRibbonBox6.Ribbons = AddinExpress.MSO.ADXRibbons.msrExcelWorkbook;
            // 
            // numSavedLbl
            // 
            this.numSavedLbl.Id = "adxRibbonLabel_933dd710d7104829b9b79698bdafeead";
            this.numSavedLbl.Ribbons = AddinExpress.MSO.ADXRibbons.msrExcelWorkbook;
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

