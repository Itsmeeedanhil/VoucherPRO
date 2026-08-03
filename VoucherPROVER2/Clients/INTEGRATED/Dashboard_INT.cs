using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Drawing.Printing;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;
using CrystalDecisions.Shared;
using CrystalDecisions.CrystalReports.Engine;
using CrystalDecisions.Windows.Forms;
using CrystalDecisions.ReportAppServer;
using static VoucherPROVER2.Clients.INT.Dataclass_INT;
using System.IO;
using System.Data.OleDb;
using VoucherPROVER2.Clients.INT;
using VoucherPROVER2.Clients.INTEGRATED;


namespace VoucherPROVER2.Clients.INT
{
    public partial class Dashboard_INT : Form
    {
        public Dashboard_INT()
        {
            InitializeComponent();

            accessToDatabase = new AccessToDatabase_INT();

            this.CreateHandle();
        }

        private PrintDocument printDocument;
        private PrintPreviewControl printPreviewControl;
        private CrystalReportViewer reportViewer;
        private AccessToDatabase_INT accessToDatabase;


        FlowLayoutPanel panel_Company;

        ComboBox comboBox_Forms;
        ComboBox comboBox_Company;

        Label label_SeriesNumberText;
        Label label_SignatoryRRStatus;
        private Label label_VoucherType;
        private ComboBox comboBox_VoucherType;

        TextBox textBox_SeriesNumber;
        TextBox textBox_ReceivedByRR;
        TextBox textBox_CheckedByRR;
        ComboBox comboBox_Currency;
        Label label_CurrencyText;

        FlowLayoutPanel panel_PayeeOverride;
        TextBox textBox_PayeeOverride;

        Panel panel_Main;
        Panel panel_Main_CR;

        FlowLayoutPanel panel_Printing;
        FlowLayoutPanel panel_SeriesNumber;
        FlowLayoutPanel panel_Signatory;
        FlowLayoutPanel panel_RRSignatory;
        FlowLayoutPanel panel_RefNumber;
        FlowLayoutPanel panel_RefNumberCrystalReport;

        List<CheckTable> cheque = new List<CheckTable>();
        List<CheckTableGrid> checkivp = new List<CheckTableGrid>();
        List<BillTable> bills = new List<BillTable>();
        List<CheckTableExpensesAndItems> checks = new List<CheckTableExpensesAndItems>();
        List<ItemReciept> receipts = new List<ItemReciept>();
        List<BillTable> apvData = new List<BillTable>();
        List<CheckTableExpensesAndItems> cvData = new List<CheckTableExpensesAndItems>();
        List<JournalGridItem> journal = new List<JournalGridItem>();

        static int sideBarWidth = 250;
        int seriesNumber = 1;

        //private const int itemsPerPage = 16;
        private int itemCounter;
        private int pageCounter;

        Font font_Label = new Font("Microsoft Sans Serif", 9);

        public FlowLayoutPanel Panel_SBPayeeOverride()
        {
            panel_PayeeOverride = new FlowLayoutPanel
            {
                Dock = DockStyle.Top,
                Height = 61,
                Width = sideBarWidth - 10,
                BackColor = Color.LightGray,
                Padding = new Padding(5, 2, 5, 5),
                BorderStyle = BorderStyle.FixedSingle,
                Visible = false // Default hidden
            };

            Label label_Text = new Label
            {
                Parent = panel_PayeeOverride,
                Width = sideBarWidth - 10,
                Text = "PAYEE :",
                TextAlign = ContentAlignment.MiddleLeft,
                Font = font_Label,
            };

            textBox_PayeeOverride = new TextBox
            {
                Parent = panel_PayeeOverride,
                Width = sideBarWidth - 28,
                Font = font_Label,
            };

            return panel_PayeeOverride;
        }


        public FlowLayoutPanel Panel_SBCompany()
        {
            panel_Company = new FlowLayoutPanel
            {
                Dock = DockStyle.Top,
                Height = 150,
                Width = sideBarWidth - 10,
                BackColor = Color.LightGray,
                Padding = new Padding(5, 2, 5, 5),
                BorderStyle = BorderStyle.FixedSingle,
                Visible = (GlobalVariables.client == "INT")
            };

            Label label_CompanyText = new Label
            {
                Parent = panel_Company,
                Width = sideBarWidth - 10,
                Text = "SELECT COMPANY:",
                TextAlign = ContentAlignment.MiddleLeft,
                Font = font_Label,
            };

            comboBox_Company = new ComboBox
            {
                Parent = panel_Company,
                Width = sideBarWidth - 28,
                DropDownStyle = ComboBoxStyle.DropDownList,
                Font = font_Label,
            };

            comboBox_Company.Items.AddRange(new string[]
            {
                "INTEGRATED CONTRACTOR & Plumbing Works, Inc.",
            });

            if (comboBox_Company.Items.Count > 0)
            {   
                comboBox_Company.SelectedIndex = 0;
            }

            comboBox_Company.SelectedIndexChanged += (sender, e) =>
            {
                string formType = "";
                if (comboBox_Forms.SelectedIndex == 1) formType = "CV";
                else if (comboBox_Forms.SelectedIndex == 3) formType = "JV";

                if (formType != "")
                {
                    string selectedCompany = comboBox_Company.SelectedItem.ToString();
                    seriesNumber = accessToDatabase.GetSeriesNumberFromDatabase(formType, selectedCompany);
                    UpdateSeriesNumberINT(formType);
                }
            };

            // =========================================================================
            // VOUCHER TYPE DROPDOWN (ONLY SHOWS FOR JOURNAL VOUCHER)
            // =========================================================================
            label_VoucherType = new Label // <-- Removed "Label" keyword here (uses class field)
            {
                Parent = panel_Company,
                Width = sideBarWidth - 10,
                Text = "SELECT VOUCHER TITLE:",
                TextAlign = ContentAlignment.MiddleLeft,
                Font = font_Label,
                Margin = new Padding(0, 5, 0, 0),
                Visible = (comboBox_Forms != null && comboBox_Forms.SelectedIndex == 3) // Initial check
            };

            comboBox_VoucherType = new ComboBox // <-- Removed "ComboBox" keyword here (uses class field)
            {
                Parent = panel_Company,
                Width = sideBarWidth - 28,
                DropDownStyle = ComboBoxStyle.DropDownList,
                Font = font_Label,
                Visible = (comboBox_Forms != null && comboBox_Forms.SelectedIndex == 3) // Initial check
            };

            comboBox_VoucherType.Items.AddRange(new string[]
            {
                "JOURNAL ENTRY VOUCHER",
                "EMPLOYEE SUPPLIES VOUCHER"
            });
            comboBox_VoucherType.SelectedIndex = 0;
            // =========================================================================

            label_CurrencyText = new Label
            {
                Parent = panel_Company,
                Width = sideBarWidth - 10,
                Text = "SELECT CURRENCY:",
                TextAlign = ContentAlignment.MiddleLeft,
                Font = font_Label,
                Margin = new Padding(0, 5, 0, 0)
            };

            comboBox_Currency = new ComboBox
            {
                Parent = panel_Company,
                Width = sideBarWidth - 28,
                DropDownStyle = ComboBoxStyle.DropDownList,
                Font = font_Label,
            };

            comboBox_Currency.Items.AddRange(new string[] { "Peso (₱)", "Dollar ($)" });
            comboBox_Currency.SelectedIndex = 0;

            return panel_Company;
        }

        public Panel ContainerPanel()
        {
            Panel panel_Container = new Panel
            {
                Dock = DockStyle.Fill,
            };

            Panel panel_Title = TitlePanel();
            panel_Main = MainPanel();
            panel_Main_CR = MainPanel_CR();
            Panel panel_SideBar = SideBarPanel();

            panel_SideBar.Parent = panel_Container;
            panel_Title.Parent = panel_Container;
            panel_Main.Parent = panel_Container;
            panel_Main_CR.Parent = panel_Container;

            return panel_Container;
        }

        public Panel TitlePanel()
        {
            Panel panel_Title = new Panel
            {
                Dock = DockStyle.Top,
                Padding = new Padding(5),
                Height = 50,
                BackColor = Color.FromArgb(51, 183, 240),
            };

            Label labelTop = new Label
            {
                Parent = panel_Title,
                Font = new Font("Microsoft Sans Serif", 12, FontStyle.Regular),
                Dock = DockStyle.Fill,
                //Text = "QUICKBOOKS SALES INVOICE",
                Text = "V o u c h e r P r o",
                TextAlign = ContentAlignment.MiddleRight,
                ForeColor = Color.White,
            };

            return panel_Title;
        }

        public Panel MainPanel()
        {
            Panel panel_Main = new Panel
            {
                BackColor = Color.LightGray,
                Dock = DockStyle.Fill,
                Padding = new Padding(sideBarWidth, 50, 0, 0),
                //Height = 300,
            };

            printPreviewControl = new PrintPreviewControl
            {
                Parent = panel_Main,
                Dock = DockStyle.Fill,
                Zoom = 1,
                Visible = false,
            };

            return panel_Main;
        }

        public Panel MainPanel_CR()
        {
            Panel panel_Main_CR = new Panel
            {
                BackColor = Color.LightGray,
                Dock = DockStyle.Fill,
                Padding = new Padding(sideBarWidth, 50, 0, 0),
            };

            reportViewer = new CrystalReportViewer
            {
                Parent = panel_Main_CR,
                Dock = DockStyle.Fill,
                ShowCopyButton = false,
                ShowPrintButton = true,
                ShowExportButton = false,
                ShowRefreshButton = false,
                ShowGroupTreeButton = false,
                ShowTextSearchButton = false,
                ShowParameterPanelButton = false,
                ToolPanelView = ToolPanelViewType.None
            };

            foreach (Control control in reportViewer.Controls)
            {
                if (control is System.Windows.Forms.ToolStrip toolStrip)
                {
                    foreach (ToolStripItem item in toolStrip.Items)
                    {
                        // FIX: Check if ToolTipText is not null before checking Contains
                        if (string.IsNullOrEmpty(item.ToolTipText) || !item.ToolTipText.Contains("Print"))
                        {
                            continue;
                        }

                        // If we get here, we found the Print button
                        item.Click += (s, e) =>
                        {
                            // Check if we are in IVP mode
                            /*if (GlobalVariables.client == "IVP")
                            {
                                try
                                {
                                    string formType = "";
                                    if (comboBox_Forms.SelectedIndex == 1) formType = "CV";
                                    else if (comboBox_Forms.SelectedIndex == 3) formType = "JV";
                                    else if (comboBox_Forms.SelectedIndex == 4) formType = "APV";

                                    string selectedCompany = comboBox_Company.SelectedItem?.ToString();

                                    if (formType != "" && !string.IsNullOrEmpty(selectedCompany))
                                    {
                                        // 1. Increment in memory
                                        seriesNumber++;

                                        // 2. Update Database
                                        accessToDatabase.UpdateManualSeriesNumber(formType, seriesNumber, selectedCompany);

                                        // 3. Update Sidebar UI safely
                                        this.BeginInvoke((MethodInvoker)delegate
                                        {
                                            UpdateSeriesNumberIVP(formType);
                                        });
                                    }
                                }
                                catch (Exception ex)
                                {
                                    MessageBox.Show($"Error updating series number: {ex.Message}");
                                }
                            }*/ // FOR MANUAL ENTRY

                            if (GlobalVariables.client == "INT")
                            {
                                string formType = "";
                                if (comboBox_Forms.SelectedIndex == 1) formType = "CV";
                                else if (comboBox_Forms.SelectedIndex == 3) formType = "JV";
                                else if (comboBox_Forms.SelectedIndex == 4) formType = "APV";

                                string selectedCompany = comboBox_Company.SelectedItem?.ToString();

                                if (!string.IsNullOrEmpty(formType) && !string.IsNullOrEmpty(selectedCompany))
                                {
                                    // 1. Increment the number currently in memory
                                    seriesNumber++;

                                    // 2. Save the NEW number to the database automatically
                                    accessToDatabase.UpdateManualSeriesNumber(formType, seriesNumber, selectedCompany);

                                    // 3. Update the TextBox display to show the NEXT number (CV-00002)
                                    this.BeginInvoke((MethodInvoker)delegate
                                    {
                                        UpdateSeriesNumberINT(formType);
                                    });
                                }
                            }

                        };
                    }
                }
            }

            return panel_Main_CR;
        }

        private Panel SideBarPanel()
        {
            FlowLayoutPanel panel_SideBar = new FlowLayoutPanel
            {
                Dock = DockStyle.Left,
                Width = sideBarWidth,
                Padding = new Padding(2),
                //BackColor = Color.Green,
                BackColor = Color.FromArgb(9, 102, 176)
            };

            // - FORMS --------------------------------------------------
            FlowLayoutPanel panels_Forms = Panel_SBForms();
            panels_Forms.Parent = panel_SideBar;

            // - SERIES NUMBER ------------------------------------------
            panel_SeriesNumber = Panel_SBSeriesNumber();
            panel_SeriesNumber.Parent = panel_SideBar;
            panel_SeriesNumber.Visible = false;

            if (GlobalVariables.client == "INT")
            {
                FlowLayoutPanel panel_Company = Panel_SBCompany();
                panel_Company.Parent = panel_SideBar;

                // --- ADD THIS BLOCK ---
                FlowLayoutPanel panel_Payee = Panel_SBPayeeOverride();
                panel_Payee.Parent = panel_SideBar;
                // ----------------------
            }

            // - REF NUMBER ---------------------------------------------
            panel_RefNumber = Panel_SBRefNumber();
            panel_RefNumberCrystalReport = Panel_SBRefNumber_CR();
            panel_RefNumber.Parent = panel_SideBar;
            panel_RefNumberCrystalReport.Parent = panel_SideBar;
            panel_RefNumber.Visible = false;
            panel_RefNumberCrystalReport.Visible = false;

            // - SIGNATORY ----------------------------------------------
            panel_Signatory = Panel_SBSignatory();
            panel_Signatory.Parent = panel_SideBar;
            panel_Signatory.Visible = false;

            // - RR SIGNATORY -------------------------------------------
            if (GlobalVariables.client == "LEADS")
            {
                panel_RRSignatory = Panel_SBRRSignatory();
                panel_RRSignatory.Parent = panel_SideBar;
                panel_RRSignatory.Visible = false;
            }

            // - PRINTING -----------------------------------------------
            FlowLayoutPanel panel_Printing = Panel_SBPrinting();
            panel_Printing.Parent = panel_SideBar;

            // ----------------------------------------------------------

            return panel_SideBar;
        }

        public FlowLayoutPanel Panel_SBForms()
        {
            FlowLayoutPanel panel_Forms = new FlowLayoutPanel
            {
                //Parent = panel_SideBar,
                Dock = DockStyle.Top,
                Height = 61,
                Width = sideBarWidth - 10,
                BackColor = Color.LightGray,
                Padding = new Padding(5, 2, 5, 5),
                BorderStyle = BorderStyle.FixedSingle,
            };

            Label label_FormText = new Label
            {
                Parent = panel_Forms,
                Width = sideBarWidth - 10,
                Text = "SELECT FORM:",
                TextAlign = ContentAlignment.MiddleLeft,
                Font = font_Label,
            };

            comboBox_Forms = new ComboBox
            {
                Parent = panel_Forms,
                Width = sideBarWidth - 28,
                DropDownStyle = ComboBoxStyle.DropDownList,
                Font = font_Label,
            };
            if (GlobalVariables.client == "INT")
            {
                comboBox_Forms.Items.AddRange(new string[]
            {
                "",
                "Online Voucher (Write Checks)",
                "Check",
                "Journal Voucher (General Journal)",
                "Check Voucher (Enter Bills)"

            });
                comboBox_Forms.SelectedIndex = 0;
                comboBox_Forms.SelectedIndexChanged += ComboBox_Forms_SelectedIndexChanged;
            }

            return panel_Forms;
        }

        public FlowLayoutPanel Panel_SBSeriesNumber()
        {
            FlowLayoutPanel panel_SeriesNumber = new FlowLayoutPanel
            {
                //Parent = panel_SideBar,
                Dock = DockStyle.Top,
                Height = 62,
                Width = sideBarWidth - 10,
                BackColor = Color.LightGray,
                Padding = new Padding(5, 2, 5, 5),
                BorderStyle = BorderStyle.FixedSingle,
                Visible = false,
            };

            label_SeriesNumberText = new Label
            {
                Parent = panel_SeriesNumber,
                Width = sideBarWidth - 30,
                Text = "Current Series Number:",
                TextAlign = ContentAlignment.MiddleLeft,
                Font = font_Label,
            };

            textBox_SeriesNumber = new TextBox
            {
                Parent = panel_SeriesNumber,
                Width = 156,
                Font = new Font("Microsoft Sans Serif", 10),
            };
            //textBox_SeriesNumber.TextChanged += TextBox_SeriesNumber_TextChanged;
            //textBox_SeriesNumber.Leave += TextBox_SeriesNumber_Leave;

            Button button_Decrement = new Button
            {
                Parent = panel_SeriesNumber,
                Height = 28,
                Width = 28,
                Text = "-",
                TextAlign = ContentAlignment.MiddleCenter,
                Margin = new Padding(0, 1, 0, 0),
                BackColor = Color.Transparent,
            };
            button_Decrement.Click += (sender, e) =>
            {
                if (GlobalVariables.client == "INT")
                {
                    seriesNumber--;
                    string prefix = "";
                    if (comboBox_Forms.SelectedIndex == 1) prefix = "CV";
                    else if (comboBox_Forms.SelectedIndex == 3) prefix = "JV";
                    else if (comboBox_Forms.SelectedIndex == 4) prefix = "APV";

                    UpdateSeriesNumberINT(prefix); // Use the new 5-digit formatter
                }




            };

            Button button_Increment = new Button
            {
                Parent = panel_SeriesNumber,
                Height = 28,
                Width = 28,
                Text = "+",
                TextAlign = ContentAlignment.MiddleCenter,
                Margin = new Padding(3, 1, 3, 0),
                BackColor = Color.Transparent,
            };
            button_Increment.Click += (sender, e) =>
            {
                if (GlobalVariables.client == "INT")
                {
                    seriesNumber++;
                    string prefix = "";
                    if (comboBox_Forms.SelectedIndex == 1) prefix = "CV";
                    else if (comboBox_Forms.SelectedIndex == 3) prefix = "JV";
                    else if (comboBox_Forms.SelectedIndex == 4) prefix = "APV";
                    UpdateSeriesNumberINT(prefix); // Use the new 5-digit formatter
                }
            };

            return panel_SeriesNumber;
        }


        public FlowLayoutPanel Panel_SBRefNumber_CR()
        {
            FlowLayoutPanel panel_RefNumber_CR = new FlowLayoutPanel
            {
                //Parent = panel_SideBar,
                Dock = DockStyle.Top,
                Height = 90,
                Width = sideBarWidth - 10,
                BackColor = Color.LightGray,
                Padding = new Padding(5, 2, 5, 5),
                BorderStyle = BorderStyle.FixedSingle,
                //Visible = false
            };

            Label label_RefNumberText = new Label
            {
                Parent = panel_RefNumber_CR,
                Width = sideBarWidth - 30,
                Text = "ENTER REFERENCE NUMBER: ",
                TextAlign = ContentAlignment.MiddleLeft,
                Font = font_Label,
            };

            TextBox textBox_ReferenceNumber_CR = new TextBox
            {
                Parent = panel_RefNumber_CR,
                Width = sideBarWidth - 30, // 190
                Font = font_Label,
            };

            Button button_SearchRefNum_CR = new Button
            {
                Parent = panel_RefNumber_CR,
                Height = 26,
                Width = sideBarWidth - 30,
                Text = "SEARCH",
                BackColor = Color.Transparent,
            };
            button_SearchRefNum_CR.Click += (sender, e) =>
            {
                if (comboBox_Forms.SelectedIndex == 0)
                {
                    MessageBox.Show("Please select a form.", "Notice", MessageBoxButtons.OK);
                }
                else if (comboBox_Forms.SelectedIndex != 0 && textBox_ReferenceNumber_CR.Text != "")
                {
                    if (GlobalVariables.client == "INT")
                    {
                        // -------------------------------------------------------------
                        // OPTION 1: CHECK VOUCHER
                        // -------------------------------------------------------------
                        if (comboBox_Forms.SelectedIndex == 1)
                        {
                            bool cvDataExists = false;
                            try
                            {
                                CRCV_INT cRCV_INT = new CRCV_INT();
                                string databasePath = Path.Combine(Application.StartupPath, "CheckDatabase.accdb");
                                SetDatabaseLocation(cRCV_INT, databasePath);

                                AccessQueries_INT accessQueries = new AccessQueries_INT();
                                string refNumberCR = textBox_ReferenceNumber_CR.Text;

                                cvData = accessQueries.GetCheckExpensesAndItemsData_INT(refNumberCR);

                                if (cvData != null && cvData.Count > 0)
                                {
                                    cvDataExists = true;

                                    TextObject textObject_CVPayee = cRCV_INT.ReportDefinition.ReportObjects["TextCVPayee"] as TextObject;
                                    TextObject textObject_CVDatenow = cRCV_INT.ReportDefinition.ReportObjects["TextCVDatenow"] as TextObject;
                                    TextObject textObject_CVAddress = cRCV_INT.ReportDefinition.ReportObjects["TextCVAddress"] as TextObject;


                                    TextObject textObject_PreparedBy = cRCV_INT.ReportDefinition.ReportObjects["TextPreparedBy"] as TextObject;
                                    TextObject textObject_PreparedByPos = cRCV_INT.ReportDefinition.ReportObjects["TextPreparedByPosition"] as TextObject;
                                    TextObject textObject_CheckedBy = cRCV_INT.ReportDefinition.ReportObjects["TextCheckedBy"] as TextObject;
                                    TextObject textObject_CheckedByPos = cRCV_INT.ReportDefinition.ReportObjects["TextCheckedByPosition"] as TextObject;
                                    TextObject textObject_ApprovedBy = cRCV_INT.ReportDefinition.ReportObjects["TextApprovedBy"] as TextObject;
                                    TextObject textObject_ApprovedByPos = cRCV_INT.ReportDefinition.ReportObjects["TextApprovedByPosition"] as TextObject;

                                    TextObject textObject_CVBankAccount = cRCV_INT.ReportDefinition.ReportObjects["TextCVBankAccount"] as TextObject;
                                    TextObject textObject_CVRefNumber = cRCV_INT.ReportDefinition.ReportObjects["TextCVRefNumber"] as TextObject;
                                    TextObject textObject_CVCheckDate = cRCV_INT.ReportDefinition.ReportObjects["TextCVCheckDate"] as TextObject;
                                    TextObject textObject_CVAmountinWords = cRCV_INT.ReportDefinition.ReportObjects["TextCVAmountInWords"] as TextObject;
                                    TextObject textObject_CVTotal = cRCV_INT.ReportDefinition.ReportObjects["TextCVAmount"] as TextObject;


                                    AccessToDatabase_INT accessToDatabase = new AccessToDatabase_INT();
                                    var signatories = accessToDatabase.RetrieveAllSignatoryData();

                                    
                                    double amount = cvData[0].TotalAmount;
                                    string amountInWords = AccessToDatabase_INT.AmountToWordsConverter.Convert(amount);
                                    string rawBank = cvData[0].BankAccount ?? "";

                                    // 2. Extract only the part after the ':'
                                    string bank = rawBank.Contains(":")
                                        ? rawBank.Split(':').Last().Trim()
                                        : rawBank;

                                    textObject_CVRefNumber.Text = refNumberCR;

                                    var b = cvData[0];

                                    // Line 1: Combine Addr1, Addr2, Addr3, Addr4 into one string separated by commas
                                    string streetLine = string.Join(", ", new[] {
                                                 b.AddressBlockAddr1,
                                                 b.AddressBlockAddr2,
                                                 b.AddressBlockAddr3,
                                                 b.AddressBlockAddr4
                                             }.Where(s => !string.IsNullOrWhiteSpace(s)));

                                    // Line 2: City (Add State/Zip here if you have them in your BillTable)
                                    string cityLine = string.Join(" ", new[] {
                                                 b.AddressCity,
                                             }.Where(s => !string.IsNullOrWhiteSpace(s)));

                                    // Final: Join the two lines with a single NewLine
                                    string fullAddress = string.Join(Environment.NewLine, new[] { streetLine, cityLine }.Where(s => !string.IsNullOrWhiteSpace(s)));

                                    textObject_CVCheckDate.Text = cvData[0].DueDate.ToString("MMMM dd, yyyy");
                                    textObject_CVPayee.Text = cvData[0].PayeeFullName;
                                    textObject_CVAddress.Text = fullAddress;
                                    Console.WriteLine($"Payee Address: {fullAddress}"); // Debugging line
                                    textObject_CVDatenow.Text = DateTime.Now.ToString("MMMM dd, yyyy");
                                    textObject_CVTotal.Text = cvData[0].TotalAmount.ToString("N2");


                                    

                                    textObject_PreparedBy.Text = signatories.PreparedByName;
                                    textObject_PreparedByPos.Text = signatories.PreparedByPosition;
                                    textObject_CheckedBy.Text = signatories.ReviewedByName;
                                    textObject_CheckedByPos.Text = signatories.ReviewedByPosition;
                                    textObject_ApprovedBy.Text = signatories.ApprovedByName;
                                    textObject_ApprovedByPos.Text = signatories.ApprovedByPosition;
                                    //textObject_ReceivedBy.Text = signatories.ReceivedByName;
                                    //textObject_ReceivedByPos.Text = signatories.ReceivedByPosition;

                                    textObject_CVAmountinWords.Text = "          " + amountInWords;
                                    textObject_CVBankAccount.Text = bank;

                                    SubreportObject subreportObject = cRCV_INT.ReportDefinition.ReportObjects["SubreportCVDetailsIVP"] as SubreportObject;
                                    if (subreportObject != null)
                                    {
                                        ReportDocument subReportDocument = cRCV_INT.OpenSubreport(subreportObject.SubreportName);

                                        InsertDataToCheckVoucherCompiledINT(refNumberCR, cvData);
                                    }
                                    SubreportObject subreportObjectcredit = cRCV_INT.ReportDefinition.ReportObjects["SubreportCVDetailsINTCredit"] as SubreportObject;
                                    if (subreportObjectcredit != null)
                                    {
                                        ReportDocument subReportDocumentcredit = cRCV_INT.OpenSubreport(subreportObjectcredit.SubreportName);

                                        InsertDataToCheckVoucherCompiledINT(refNumberCR, cvData);
                                    }

                                    SubreportObject subreportObject2 = cRCV_INT.ReportDefinition.ReportObjects["SubreportCVDetailsINT"] as SubreportObject;
                                    if (subreportObject2 != null)
                                    {
                                        ReportDocument subReportDocument2 = cRCV_INT.OpenSubreport(subreportObject2.SubreportName);
                                        TextObject textObject_Remarks = subReportDocument2.ReportDefinition.ReportObjects["TextRemarks"] as TextObject;
                                        TextObject textObject_CVSubTotal = subReportDocument2.ReportDefinition.ReportObjects["TextCVSubTotalAmount"] as TextObject;



                                        textObject_Remarks.Text = SafeTruncate(cvData[0].Memo, 500);
                                        textObject_CVSubTotal.Text = cvData[0].TotalAmount.ToString("N2");

                                        InsertDataToCheckVoucherCompiledINT(refNumberCR, cvData);
                                    }

                                    cRCV_INT.SetParameterValue("ReferenceNumber", refNumberCR);

                                    panel_Printing.Visible = false;
                                    panel_Signatory.Visible = true;
                                    panel_Main.Visible = false;
                                    panel_Main_CR.Visible = true;

                                    reportViewer.ReportSource = cRCV_INT;
                                    reportViewer.RefreshReport();
                                }
                            }
                            catch (Exception ex)
                            {
                                MessageBox.Show($"IVP CV ERROR:\n{ex.Message}", "Error", MessageBoxButtons.OK, MessageBoxIcon.Error);
                            }

                            if (!cvDataExists)
                            {
                                string refNumberCR = textBox_ReferenceNumber_CR.Text;
                                GenerateBillPaymentReport_INT(refNumberCR);
                            }
                        }

                        else if (comboBox_Forms.SelectedIndex == 3)
                        {
                            CRJV_INT cRJV_INT = new CRJV_INT();
                            string databasePath = Path.Combine(Application.StartupPath, "CheckDatabase.accdb");
                            SetDatabaseLocation(cRJV_INT, databasePath);

                            AccessQueries_INT accessQueries = new AccessQueries_INT();
                            string refNumberCR = textBox_ReferenceNumber_CR.Text;

                            journal = accessQueries.GetJournalEntryForGrid(refNumberCR);

                            if (journal != null && journal.Count > 0)
                            {
                                // =========================================================================
                                // READ VOUCHER TYPE SELECTION
                                // =========================================================================
                                string selectedVoucherTitle = comboBox_VoucherType?.SelectedItem?.ToString() ?? "JOURNAL ENTRY VOUCHER";
                                string voucherNoLabel = (selectedVoucherTitle == "EMPLOYEE SUPPLIES VOUCHER") ? "E.S.V. No.:" : "J.V. No.:";

                                // Pass title to report
                                if (cRJV_INT.ReportDefinition.ReportObjects["TextReportTitle"] is TextObject textObject_ReportTitle)
                                {
                                    textObject_ReportTitle.Text = selectedVoucherTitle;
                                }

                                // Pass label to report
                                if (cRJV_INT.ReportDefinition.ReportObjects["TextVoucherNoLabel"] is TextObject textObject_VoucherNoLabel)
                                {
                                    textObject_VoucherNoLabel.Text = voucherNoLabel;
                                }
                                // =========================================================================

                                // Rest of your existing code below...
                                string Memo = "                         " + journal[journal.Count - 1].Memo;

                                var journalLineType = journal[0].GetType();
                                while (journal.Count < 10)
                                {
                                    var emptyLine = Activator.CreateInstance(journalLineType);
                                    try { journalLineType.GetProperty("AccountNumber")?.SetValue(emptyLine, ""); } catch { }
                                    try { journalLineType.GetProperty("Particulars")?.SetValue(emptyLine, ""); } catch { }
                                    try { journalLineType.GetProperty("Debit")?.SetValue(emptyLine, 0.0); } catch { }
                                    try { journalLineType.GetProperty("Credit")?.SetValue(emptyLine, 0.0); } catch { }
                                    try { journalLineType.GetProperty("Memo")?.SetValue(emptyLine, ""); } catch { }

                                    journal.Add((dynamic)emptyLine);
                                }


                                TextObject textObject_JVCheckDate = cRJV_INT.ReportDefinition.ReportObjects["TextJVCheckDate"] as TextObject;
                                TextObject textObject_JVRefnumber = cRJV_INT.ReportDefinition.ReportObjects["TextJVRefnumber"] as TextObject;
                                

                                TextObject textObject_CompanyName = cRJV_INT.ReportDefinition.ReportObjects["TextCompanyName"] as TextObject;
                                if (textObject_CompanyName != null && comboBox_Company != null && comboBox_Company.SelectedItem != null)
                                {
                                    textObject_CompanyName.Text = comboBox_Company.SelectedItem.ToString();
                                }

                                TextObject textObject_PreparedBy = cRJV_INT.ReportDefinition.ReportObjects["TextPreparedBy"] as TextObject;
                                //TextObject textObject_PreparedByPos = cRJV_INT.ReportDefinition.ReportObjects["TextPreparedByPosition"] as TextObject;
                                TextObject textObject_CheckedBy = cRJV_INT.ReportDefinition.ReportObjects["TextCheckedBy"] as TextObject;
                                //TextObject textObject_CheckedByPos = cRJV_INT.ReportDefinition.ReportObjects["TextCheckedByPosition"] as TextObject;
                                TextObject textObject_ApprovedBy = cRJV_INT.ReportDefinition.ReportObjects["TextApprovedBy"] as TextObject;
                                //TextObject textObject_ApprovedByPos = cRJV_INT.ReportDefinition.ReportObjects["TextApprovedByPosition"] as TextObject;

                                if (textObject_JVCheckDate != null) textObject_JVCheckDate.Text = journal[0].Date.ToString("MMMM dd, yyyy");

                                string Refnumber = refNumberCR;
                                double debitTotalAmount = 0;
                                double creditTotalAmount = 0;

                                foreach (var line in journal)
                                {
                                    debitTotalAmount += line.Debit;
                                    creditTotalAmount += line.Credit;
                                }

                                if (textObject_JVRefnumber != null) textObject_JVRefnumber.Text = refNumberCR;

                                AccessToDatabase_INT accessToDatabase = new AccessToDatabase_INT();
                                var signatories = accessToDatabase.RetrieveAllSignatoryData();

                                textObject_PreparedBy.Text = signatories.PreparedByName;
                                //textObject_PreparedByPos.Text = signatories.PreparedByPosition;
                                textObject_CheckedBy.Text = signatories.ReviewedByName;
                                //textObject_CheckedByPos.Text = signatories.ReviewedByPosition;
                                textObject_ApprovedBy.Text = signatories.ApprovedByName;
                                //textObject_ApprovedByPos.Text = signatories.ApprovedByPosition;

                                // 4. Handle Subreport
                                SubreportObject subreportObject = cRJV_INT.ReportDefinition.ReportObjects["SubreportJVDetailsIVP"] as SubreportObject;
                                if (subreportObject != null)
                                {
                                    // Open the subreport document
                                    ReportDocument subReportDocument = cRJV_INT.OpenSubreport(subreportObject.SubreportName);

                                }

                                InsertDataToJournalCompiled(refNumberCR, journal);

                                // 6. Final Report Settings
                                cRJV_INT.SetParameterValue("ReferenceNumber", refNumberCR);

                                panel_Printing.Visible = false;
                                panel_Signatory.Visible = true;
                                panel_Main.Visible = false;
                                panel_Main_CR.Visible = true;

                                reportViewer.ReportSource = cRJV_INT;
                                reportViewer.RefreshReport();
                            }
                            else
                            {
                                MessageBox.Show("No Journal Entry found for this Reference Number.");
                            }
                        }
                        else if (comboBox_Forms.SelectedIndex == 4) // APV
                        {
                            string refNumberCR = textBox_ReferenceNumber_CR.Text;
                            // You can reuse GenerateBillPaymentReport_IVP or create a specific APV one:
                            GenerateAPVReport_INT(refNumberCR);
                        }
                    }

                }
                else
                {
                    MessageBox.Show("Please enter a reference number.", "Notice", MessageBoxButtons.OK);
                }
            };

            return panel_RefNumber_CR;
        }

        private bool GenerateAPVReport_INT(string refNumberCR)
        {
            try
            {
                CRAPV_INTBILL cRAPV_INTBILL = new CRAPV_INTBILL();
                string databasePathBILL = Path.Combine(Application.StartupPath, "CheckDatabase.accdb");
                SetDatabaseLocation(cRAPV_INTBILL, databasePathBILL);

                AccessQueries_INT accessQueries = new AccessQueries_INT();
                List<BillTable> bills = accessQueries.GetBillData_INT_DirectBill(refNumberCR);

                if (bills == null || bills.Count == 0)
                    return false;

                TextObject textObject_CVBILLAmountInWords = null;
                TextObject textObject_CVBILLCheckDate = null;
                TextObject textObject_CVBILLDate = null;
                TextObject textObject_CVBILLPayee = null;
                TextObject textObject_CVBILLAddress = null;
                TextObject textObject_CVBILLTotalAmount = null;
                TextObject textObject_CVBILLBankAccount = null;
                TextObject textObject_CVBILLRefNumber = null;
                TextObject textObject_PreparedBy = null;
                TextObject textObject_PreparedByPos = null;
                TextObject textObject_CheckedBy = null;
                TextObject textObject_CheckedByPos = null;
                TextObject textObject_ApprovedBy = null;
                TextObject textObject_ApprovedByPos = null;

                try
                {
                    textObject_CVBILLAddress = cRAPV_INTBILL.ReportDefinition.ReportObjects["TextCVBILLAddress"] as TextObject;
                    textObject_CVBILLPayee = cRAPV_INTBILL.ReportDefinition.ReportObjects["TextCVBILLPayee"] as TextObject;
                    textObject_CVBILLDate = cRAPV_INTBILL.ReportDefinition.ReportObjects["TextCVBILLDate"] as TextObject;
                    textObject_CVBILLAmountInWords = cRAPV_INTBILL.ReportDefinition.ReportObjects["TextCVBILLAmountInWords"] as TextObject;
                    textObject_CVBILLCheckDate = cRAPV_INTBILL.ReportDefinition.ReportObjects["TextCVBILLCheckDate"] as TextObject;
                    textObject_CVBILLRefNumber = cRAPV_INTBILL.ReportDefinition.ReportObjects["TextCVBILLRefNumber"] as TextObject;
                    textObject_CVBILLTotalAmount = cRAPV_INTBILL.ReportDefinition.ReportObjects["TextCVBILLTotalAmount"] as TextObject;
                    textObject_CVBILLBankAccount = cRAPV_INTBILL.ReportDefinition.ReportObjects["TextCVBILLBankAccount"] as TextObject;

                    textObject_PreparedBy = cRAPV_INTBILL.ReportDefinition.ReportObjects["TextPreparedBy"] as TextObject;
                    textObject_PreparedByPos = cRAPV_INTBILL.ReportDefinition.ReportObjects["TextPreparedByPosition"] as TextObject;
                    textObject_CheckedBy = cRAPV_INTBILL.ReportDefinition.ReportObjects["TextCheckedBy"] as TextObject;
                    textObject_CheckedByPos = cRAPV_INTBILL.ReportDefinition.ReportObjects["TextCheckedByPosition"] as TextObject;
                    textObject_ApprovedBy = cRAPV_INTBILL.ReportDefinition.ReportObjects["TextApprovedBy"] as TextObject;
                    textObject_ApprovedByPos = cRAPV_INTBILL.ReportDefinition.ReportObjects["TextApprovedByPosition"] as TextObject;

                    AccessToDatabase_INT accessToDatabase = new AccessToDatabase_INT();

                    var (PreparedByName, PreparedByPosition,
                         ReviewedByName, ReviewedByPosition,
                         RecommendingApprovalName, RecommendingApprovalPosition,
                         ApprovedByName, ApprovedByPosition,
                         ReceivedByName, ReceivedByPosition) = accessToDatabase.RetrieveAllSignatoryData();

                    if (textObject_PreparedBy != null) textObject_PreparedBy.Text = PreparedByName;
                    if (textObject_PreparedByPos != null) textObject_PreparedByPos.Text = PreparedByPosition;
                    if (textObject_CheckedBy != null) textObject_CheckedBy.Text = ReviewedByName;
                    if (textObject_CheckedByPos != null) textObject_CheckedByPos.Text = ReviewedByPosition;
                    if (textObject_ApprovedBy != null) textObject_ApprovedBy.Text = ApprovedByName;
                    if (textObject_ApprovedByPos != null) textObject_ApprovedByPos.Text = ApprovedByPosition;
                }
                catch
                {
                    throw;
                }

                // =========================================================================
                // CONSOLIDATED TOTALS ACROSS ALL BILLS
                // =========================================================================
                double totalVoucherAmount = bills.Sum(bill =>
                {
                    double lineSum = bill.ItemDetails?.Sum(d => d.ItemLineAmount > 0 ? d.ItemLineAmount : d.ExpenseLineAmount) ?? 0;
                    return lineSum > 0 ? lineSum : (bill.AmountDue > 0 ? bill.AmountDue : bill.Amount);
                });

                string amountInWords = "          " + AccessToDatabase_INT.AmountToWordsConverter.Convert(totalVoucherAmount);

                string bankaccount = (bills[0].BankAccount ?? "").Contains(":")
                                    ? (bills[0].BankAccount ?? "").Split(':').Last().Trim()
                                    : (bills[0].BankAccount ?? "");

                var b = bills[0]; // Now 'b' won't conflict with the lambda above!

                string streetLine = string.Join(", ", new[] {
            b.VendorAddressAddr1,
            b.VendorAddressAddr2,
            b.VendorAddressAddr3,
            b.VendorAddressAddr4
                }.Where(s => !string.IsNullOrWhiteSpace(s)));

                string cityLine = string.Join(" ", new[] {
            b.VendorAddressCity,
                }.Where(s => !string.IsNullOrWhiteSpace(s)));

                string fullAddress = string.Join(Environment.NewLine, new[] { streetLine, cityLine }.Where(s => !string.IsNullOrWhiteSpace(s)));

                if (textObject_CVBILLRefNumber != null) textObject_CVBILLRefNumber.Text = refNumberCR;
                if (textObject_CVBILLAddress != null) textObject_CVBILLAddress.Text = fullAddress;
                if (textObject_CVBILLDate != null) textObject_CVBILLDate.Text = DateTime.Now.ToString("MMMM dd, yyyy");
                if (textObject_CVBILLAmountInWords != null) textObject_CVBILLAmountInWords.Text = amountInWords;
                if (textObject_CVBILLCheckDate != null) textObject_CVBILLCheckDate.Text = bills[0].DueDate.ToString("MMMM dd, yyyy");
                if (textObject_CVBILLPayee != null) textObject_CVBILLPayee.Text = bills[0].PayeeFullName ?? "";
                if (textObject_CVBILLTotalAmount != null) textObject_CVBILLTotalAmount.Text = totalVoucherAmount.ToString("N2");
                if (textObject_CVBILLBankAccount != null) textObject_CVBILLBankAccount.Text = bankaccount;

                // Subreport 1: Details / Remarks Summary
                SubreportObject subreportObject = cRAPV_INTBILL.ReportDefinition.ReportObjects["SubreportCVDetailsINT"] as SubreportObject;
                if (subreportObject != null)
                {
                    ReportDocument subReportDocument = cRAPV_INTBILL.OpenSubreport(subreportObject.SubreportName);

                    TextObject textObject_BILLSubRemarks = subReportDocument.ReportDefinition.ReportObjects["TextBILLRemarks"] as TextObject;
                    TextObject textObject_BILLSubAmountPayable = subReportDocument.ReportDefinition.ReportObjects["TextBILLSubAmountPayable"] as TextObject;

                   

                    if (textObject_BILLSubRemarks != null) textObject_BILLSubRemarks.Text = bills[0].Memo;
                    if (textObject_BILLSubAmountPayable != null) textObject_BILLSubAmountPayable.Text = totalVoucherAmount.ToString("N2");
                }

                // Subreport 2: IVP (Debit Details)
                SubreportObject subreportObjectIVP = cRAPV_INTBILL.ReportDefinition.ReportObjects["SubreportCVDetailsIVP"] as SubreportObject;
                if (subreportObjectIVP != null)
                {
                    cRAPV_INTBILL.OpenSubreport(subreportObjectIVP.SubreportName);
                }

                // Subreport 3: INTCredit (Credit Details)
                SubreportObject subreportObjectINTCredit = cRAPV_INTBILL.ReportDefinition.ReportObjects["SubreportCVDetailsINTCredit"] as SubreportObject;
                if (subreportObjectINTCredit != null)
                {
                    cRAPV_INTBILL.OpenSubreport(subreportObjectINTCredit.SubreportName);
                }

                // Populate MS Access Staging Database
                InsertDataToBillCompiled(refNumberCR, bills);

                cRAPV_INTBILL.SetParameterValue("ReferenceNumber", refNumberCR);

                panel_Printing.Visible = false;
                panel_Signatory.Visible = true;
                panel_Main.Visible = false;
                panel_Main_CR.Visible = true;

                reportViewer.ReportSource = cRAPV_INTBILL;
                reportViewer.RefreshReport();

                return true;
            }
            catch (Exception ex)
            {
                MessageBox.Show($"KAYAK ERROR HEHEHE:\n{ex.Message}", "Error", MessageBoxButtons.OK, MessageBoxIcon.Error);
                return false;
            }
        }

        private bool GenerateBillPaymentReport_INT(string refNumberCR)
        {
            try
            {
                CRCV_INTBILL cRCV_INTBILL = new CRCV_INTBILL();
                string databasePathBILL = Path.Combine(Application.StartupPath, "CheckDatabase.accdb");
                SetDatabaseLocation(cRCV_INTBILL, databasePathBILL);

                AccessQueries_INT accessQueries = new AccessQueries_INT();
                List<BillTable> bills = accessQueries.GetBillData_INT(refNumberCR);

                if (bills == null || bills.Count == 0)
                    return false;

                TextObject textObject_CVBILLAmountInWords = null;
                TextObject textObject_CVBILLCheckDate = null;
                TextObject textObject_CVBILLBankAccount = null;
                TextObject textObject_CVBILLAmount = null;
                TextObject textObject_CVBILLPayee = null;
                TextObject textObject_CVBILLRefnumber = null;
                TextObject textObject_CVBILLAddress = null;
                TextObject textObject_CVBILLDate = null;
                TextObject textObject_PreparedBy = null;
                TextObject textObject_PreparedByPos = null;
                TextObject textObject_CheckedBy = null;
                TextObject textObject_CheckedByPos = null;
                TextObject textObject_ApprovedBy = null;
                TextObject textObject_ApprovedByPos = null;

                try
                {
                    textObject_CVBILLPayee = cRCV_INTBILL.ReportDefinition.ReportObjects["TextCVBILLPayee"] as TextObject;
                    textObject_CVBILLAddress = cRCV_INTBILL.ReportDefinition.ReportObjects["TextCVBILLAddress"] as TextObject;
                    textObject_CVBILLDate = cRCV_INTBILL.ReportDefinition.ReportObjects["TextCVBILLDate"] as TextObject;

                    textObject_CVBILLAmountInWords = cRCV_INTBILL.ReportDefinition.ReportObjects["TextCVBILLAmountinWords"] as TextObject;
                    textObject_CVBILLCheckDate = cRCV_INTBILL.ReportDefinition.ReportObjects["TextCVBILLCheckDate"] as TextObject;
                    textObject_CVBILLRefnumber = cRCV_INTBILL.ReportDefinition.ReportObjects["TextCVRefNumber"] as TextObject;
                    textObject_CVBILLBankAccount = cRCV_INTBILL.ReportDefinition.ReportObjects["TextCVBILLBankAccount"] as TextObject;
                    textObject_CVBILLAmount = cRCV_INTBILL.ReportDefinition.ReportObjects["TextCVBILLLAmount"] as TextObject;

                    textObject_PreparedBy = cRCV_INTBILL.ReportDefinition.ReportObjects["TextPreparedBy"] as TextObject;
                    textObject_PreparedByPos = cRCV_INTBILL.ReportDefinition.ReportObjects["TextPreparedByPosition"] as TextObject;
                    textObject_CheckedBy = cRCV_INTBILL.ReportDefinition.ReportObjects["TextCheckedBy"] as TextObject;
                    textObject_CheckedByPos = cRCV_INTBILL.ReportDefinition.ReportObjects["TextCheckedByPosition"] as TextObject;
                    textObject_ApprovedBy = cRCV_INTBILL.ReportDefinition.ReportObjects["TextApprovedBy"] as TextObject;
                    textObject_ApprovedByPos = cRCV_INTBILL.ReportDefinition.ReportObjects["TextApprovedByPosition"] as TextObject;

                    AccessToDatabase_INT accessToDatabase = new AccessToDatabase_INT();

                    var (PreparedByName, PreparedByPosition,
                         ReviewedByName, ReviewedByPosition,
                         RecommendingApprovalName, RecommendingApprovalPosition,
                         ApprovedByName, ApprovedByPosition,
                         ReceivedByName, ReceivedByPosition) = accessToDatabase.RetrieveAllSignatoryData();

                    if (textObject_PreparedBy != null) textObject_PreparedBy.Text = PreparedByName;
                    if (textObject_PreparedByPos != null) textObject_PreparedByPos.Text = PreparedByPosition;
                    if (textObject_CheckedBy != null) textObject_CheckedBy.Text = ReviewedByName;
                    if (textObject_CheckedByPos != null) textObject_CheckedByPos.Text = ReviewedByPosition;
                    if (textObject_ApprovedBy != null) textObject_ApprovedBy.Text = ApprovedByName;
                    if (textObject_ApprovedByPos != null) textObject_ApprovedByPos.Text = ApprovedByPosition;
                }
                catch
                {
                    throw;
                }

                // =========================================================================
                // 1. CALCULATE INDIVIDUAL BILL AMOUNTS AND SUMMARY REMARKS
                // =========================================================================
                var billSummaryList = bills
                    .Where(x => !string.IsNullOrWhiteSpace(x.RefNumber) || !string.IsNullOrWhiteSpace(x.AppliedRefNumber))
                    .GroupBy(x => !string.IsNullOrWhiteSpace(x.AppliedRefNumber) ? x.AppliedRefNumber.Trim() : x.RefNumber.Trim())
                    .Select(g =>
                    {
                        var firstBill = g.First();
                        // Determine paid amount per bill without defaulting to total check amount
                        double paidAmount = firstBill.AppliedAmount > 0
                            ? firstBill.AppliedAmount
                            : (firstBill.Amount > 0 ? firstBill.Amount : firstBill.AmountDue);

                        return new
                        {
                            RefNumber = g.Key,
                            Amount = paidAmount
                        };
                    })
                    .ToList();

                // Build remarks string: SI#4414 - 240.00, etc.
                string billRemarksText = string.Join(Environment.NewLine,
                    billSummaryList.Select(b => $"SI#{b.RefNumber} - {b.Amount:N2}"));

                // Append main Memo if present
                string mainMemo = bills[0].BillMemo ?? bills[0].Memo;
                if (!string.IsNullOrWhiteSpace(mainMemo))
                {
                    billRemarksText = string.IsNullOrWhiteSpace(billRemarksText)
                        ? mainMemo
                        : $"{mainMemo}{Environment.NewLine}{billRemarksText}";
                }

                // Real total payout (sum of actual bill payments: 240 + 432 + 1000 = 1,672.00)
                double realCheckTotal = billSummaryList.Sum(x => x.Amount);
                if (realCheckTotal == 0) realCheckTotal = bills[0].Amount;

                string amountInWords = "          " + AccessToDatabase_INT.AmountToWordsConverter.Convert(realCheckTotal);
                string refumber2 = refNumberCR.Contains("/") ? refNumberCR.Split('/').Last() : refNumberCR;
                string bankaccount = (bills[0].BankAccount ?? "").Contains(":")
                                    ? (bills[0].BankAccount ?? "").Split(':').Last().Trim()
                                    : (bills[0].BankAccount ?? "");

                var bObj = bills[0];

                // Line 1: Combine Addr1 and Addr2
                string streetLine = string.Join(", ", new[] {
            bObj.Address,
            bObj.Address2,
        }.Where(s => !string.IsNullOrWhiteSpace(s)));

                // Line 2: City
                string cityLine = string.Join(" ", new[] {
            bObj.VendorAddressCity,
        }.Where(s => !string.IsNullOrWhiteSpace(s)));

                string fullAddress = string.Join(Environment.NewLine, new[] { streetLine, cityLine }.Where(s => !string.IsNullOrWhiteSpace(s)));

                if (textObject_CVBILLDate != null) textObject_CVBILLDate.Text = DateTime.Now.ToString("MMMM dd, yyyy");
                if (textObject_CVBILLAddress != null) textObject_CVBILLAddress.Text = fullAddress;
                if (textObject_CVBILLRefnumber != null) textObject_CVBILLRefnumber.Text = refumber2;
                if (textObject_CVBILLAmountInWords != null) textObject_CVBILLAmountInWords.Text = amountInWords;
                if (textObject_CVBILLAmount != null) textObject_CVBILLAmount.Text = realCheckTotal.ToString("N2");
                if (textObject_CVBILLCheckDate != null) textObject_CVBILLCheckDate.Text = bills[0].DueDate.ToString("MMMM dd, yyyy");
                if (textObject_CVBILLPayee != null) textObject_CVBILLPayee.Text = bills[0].PayeeFullName ?? "";
                if (textObject_CVBILLBankAccount != null) textObject_CVBILLBankAccount.Text = bankaccount;

                // Subreport 1: Debit Details
                SubreportObject subreportObjectIVP = cRCV_INTBILL.ReportDefinition.ReportObjects["SubreportCVDetailsIVP"] as SubreportObject;
                if (subreportObjectIVP != null)
                {
                    cRCV_INTBILL.OpenSubreport(subreportObjectIVP.SubreportName);
                }

                // Subreport 2: Credit Details
                SubreportObject subreportObjectINTCredit = cRCV_INTBILL.ReportDefinition.ReportObjects["SubreportCVDetailsINTCredit"] as SubreportObject;
                if (subreportObjectINTCredit != null)
                {
                    cRCV_INTBILL.OpenSubreport(subreportObjectINTCredit.SubreportName);
                }

                // Subreport 3: Remarks and Subtotal
                SubreportObject subreportObjectINT = cRCV_INTBILL.ReportDefinition.ReportObjects["SubreportCVDetailsINT"] as SubreportObject;
                if (subreportObjectINT != null)
                {
                    ReportDocument subReportDocumentINT = cRCV_INTBILL.OpenSubreport(subreportObjectINT.SubreportName);

                    TextObject textObject_BILLSubRemarks = subReportDocumentINT.ReportDefinition.ReportObjects["TextBILLRemarks"] as TextObject;
                    

                    if (textObject_BILLSubRemarks != null)
                        textObject_BILLSubRemarks.Text = SafeTruncate(billRemarksText, 500);

                }

                // Populate staging database table
                InsertDataToBillCompiled(refNumberCR, bills);

                cRCV_INTBILL.SetParameterValue("ReferenceNumber", refNumberCR);

                panel_Printing.Visible = false;
                panel_Signatory.Visible = true;
                panel_Main.Visible = false;
                panel_Main_CR.Visible = true;

                reportViewer.ReportSource = cRCV_INTBILL;
                reportViewer.RefreshReport();

                return true;
            }
            catch (Exception ex)
            {
                MessageBox.Show($"KAYAK ERROR HEHEHE:\n{ex.Message}", "Error", MessageBoxButtons.OK, MessageBoxIcon.Error);
                return false;
            }
        }

        public static void InsertDataToCheckVoucherCompiledINT(string refNumber, List<CheckTableExpensesAndItems> checkData)
        {
            string connectionString = AccessToDatabase_INT.GetAccessConnectionString();
            double debitTotalAmount = 0;
            double creditTotalAmount = 0;

            // Helper function to safely truncate strings to DB limit
            string SafeTruncate(string value, int maxLength)
            {
                if (string.IsNullOrEmpty(value)) return "";
                return value.Length <= maxLength ? value : value.Substring(0, maxLength);
            }

            using (OleDbConnection connection = new OleDbConnection(connectionString))
            {
                connection.Open();

                // 1. Clear old staging data
                string deleteQuery = "DELETE FROM CheckVoucherCompiled";
                using (OleDbCommand deleteCommand = new OleDbCommand(deleteQuery, connection))
                {
                    try
                    {
                        deleteCommand.ExecuteNonQuery();
                    }
                    catch (Exception ex)
                    {
                        MessageBox.Show($"Error deleting data: {ex.Message}", "Error", MessageBoxButtons.OK, MessageBoxIcon.Error);
                        return;
                    }
                }

                // 2. Prepare Insert Query
                string insertQuery = @"
                INSERT INTO CheckVoucherCompiled 
                (RefNumber, [Particulars], [Class], [Debit], [Credit], [Memo], [CustomerJob]) 
                VALUES 
                (@RefNumber, @Particulars, @Class, @Debit, @Credit, @Memo, @CustomerJob)";

                // =========================================================================
                // AGGREGATION / GROUPING LOGIC (DEBITS)
                // =========================================================================

                // Group Positive Item Debits
                var groupedItemDebits = checkData
                    .Where(x => !string.IsNullOrEmpty(x.Item) && x.ItemAmount > 0)
                    .GroupBy(x => new {
                        Name = x.Item.Trim(),
                        Class = x.ItemClass ?? "",
                        Memo = x.ExpensesMemo ?? "",
                        CustomerJob = x.ExpensesCustomerJob ?? ""
                    })
                    .Select(g => new {
                        Particulars = g.Key.Name,
                        Class = g.Key.Class,
                        Memo = g.Key.Memo,
                        CustomerJob = g.Key.CustomerJob,
                        TotalAmount = g.Sum(x => x.ItemAmount)
                    });

                // Group Positive Expense Debits
                var groupedExpenseDebits = checkData
                    .Where(x => !string.IsNullOrEmpty(x.Account) && x.ExpensesAmount > 0)
                    .GroupBy(x => new {
                        Name = x.Account.Trim(),
                        Class = x.ExpenseClass ?? "",
                        Memo = x.ExpensesMemo ?? "",
                        CustomerJob = x.ExpensesCustomerJob ?? ""
                    })
                    .Select(g => new {
                        Particulars = g.Key.Name,
                        Class = g.Key.Class,
                        Memo = g.Key.Memo,
                        CustomerJob = g.Key.CustomerJob,
                        TotalAmount = g.Sum(x => x.ExpensesAmount)
                    });

                // Execute Debit Inserts
                foreach (var entry in groupedItemDebits.Concat(groupedExpenseDebits))
                {
                    try
                    {
                        debitTotalAmount += entry.TotalAmount;

                        using (OleDbCommand command = new OleDbCommand(insertQuery, connection))
                        {
                            command.Parameters.AddWithValue("@RefNumber", refNumber);
                            command.Parameters.AddWithValue("@Particulars", SafeTruncate(entry.Particulars, 255));
                            command.Parameters.AddWithValue("@Class", string.IsNullOrEmpty(entry.Class) ? (object)DBNull.Value : entry.Class);
                            command.Parameters.AddWithValue("@Debit", entry.TotalAmount.ToString("N2"));
                            command.Parameters.AddWithValue("@Credit", "");
                            command.Parameters.AddWithValue("@Memo", SafeTruncate(entry.Memo, 255));
                            command.Parameters.AddWithValue("@CustomerJob", SafeTruncate(entry.CustomerJob, 255));
                            command.ExecuteNonQuery();
                        }
                    }
                    catch (Exception ex)
                    {
                        Console.WriteLine($"Error processing aggregated Debit line: {ex.Message}");
                    }
                }

                // =========================================================================
                // AGGREGATION / GROUPING LOGIC (CREDITS - Amounts < 0)
                // =========================================================================

                // Group Negative Item Credits (e.g. discounts/returns)
                var groupedItemCredits = checkData
                    .Where(x => !string.IsNullOrEmpty(x.Item) && x.ItemAmount < 0)
                    .GroupBy(x => new {
                        Name = x.Item.Trim(),
                        Class = x.ItemClass ?? "",
                        Memo = x.ExpensesMemo ?? "",
                        CustomerJob = x.ExpensesCustomerJob ?? ""
                    })
                    .Select(g => new {
                        Particulars = g.Key.Name,
                        Class = g.Key.Class,
                        Memo = g.Key.Memo,
                        CustomerJob = g.Key.CustomerJob,
                        TotalAmount = Math.Abs(g.Sum(x => x.ItemAmount))
                    });

                // Group Negative Expense Credits (e.g. Tax Withholdings / EWT)
                var groupedExpenseCredits = checkData
                    .Where(x => !string.IsNullOrEmpty(x.Account) && x.ExpensesAmount < 0)
                    .GroupBy(x => new {
                        Name = x.Account.Trim(),
                        Class = x.ExpenseClass ?? "",
                        Memo = x.ExpensesMemo ?? "",
                        CustomerJob = x.ExpensesCustomerJob ?? ""
                    })
                    .Select(g => new {
                        Particulars = g.Key.Name,
                        Class = g.Key.Class,
                        Memo = g.Key.Memo,
                        CustomerJob = g.Key.CustomerJob,
                        TotalAmount = Math.Abs(g.Sum(x => x.ExpensesAmount))
                    });

                // Execute Credit Line Inserts
                foreach (var entry in groupedItemCredits.Concat(groupedExpenseCredits))
                {
                    try
                    {
                        creditTotalAmount += entry.TotalAmount;

                        using (OleDbCommand command = new OleDbCommand(insertQuery, connection))
                        {
                            command.Parameters.AddWithValue("@RefNumber", refNumber);
                            command.Parameters.AddWithValue("@Particulars", SafeTruncate(entry.Particulars, 255));
                            command.Parameters.AddWithValue("@Class", string.IsNullOrEmpty(entry.Class) ? (object)DBNull.Value : entry.Class);
                            command.Parameters.AddWithValue("@Debit", "");
                            command.Parameters.AddWithValue("@Credit", entry.TotalAmount.ToString("N2"));
                            command.Parameters.AddWithValue("@Memo", SafeTruncate(entry.Memo, 255));
                            command.Parameters.AddWithValue("@CustomerJob", SafeTruncate(entry.CustomerJob, 255));
                            command.ExecuteNonQuery();
                        }
                    }
                    catch (Exception ex)
                    {
                        Console.WriteLine($"Error processing aggregated Credit line: {ex.Message}");
                    }
                }

                // =========================================================================
                // PASS 3: MAIN BANK ACCOUNT BALANCING CREDIT ENTRY
                // =========================================================================
                if (checkData != null && checkData.Count > 0)
                {
                    try
                    {
                        var mainCheck = checkData[0];
                        double finalCheckCredit = mainCheck.TotalAmount;
                        creditTotalAmount += finalCheckCredit;

                        using (OleDbCommand command = new OleDbCommand(insertQuery, connection))
                        {
                            command.Parameters.AddWithValue("@RefNumber", refNumber);

                            string bankName = mainCheck.BankAccount.Contains(":")
                                ? mainCheck.BankAccount.Split(':').Last().Trim()
                                : mainCheck.BankAccount;

                            command.Parameters.AddWithValue("@Particulars", SafeTruncate(bankName, 255));
                            command.Parameters.AddWithValue("@Class", (object)DBNull.Value);
                            command.Parameters.AddWithValue("@Debit", "");
                            command.Parameters.AddWithValue("@Credit", finalCheckCredit.ToString("N2"));
                            command.Parameters.AddWithValue("@Memo", SafeTruncate(mainCheck.Memo, 255));
                            command.Parameters.AddWithValue("@CustomerJob", (object)DBNull.Value);

                            command.ExecuteNonQuery();
                        }
                    }
                    catch (Exception ex)
                    {
                        Console.WriteLine($"Error processing Main Bank Credit entry: {ex.Message}");
                    }
                }

                // =========================================================================
                // PASS 4: DESCRIPTION-ONLY FALLBACKS
                // =========================================================================
                var descriptionOnlyEntries = checkData
                    .Where(x => string.IsNullOrEmpty(x.Item) && string.IsNullOrEmpty(x.Account) && !string.IsNullOrEmpty(x.ItemDescription))
                    .GroupBy(x => new {
                        Description = x.ItemDescription.Trim(),
                        Memo = x.ExpensesMemo ?? "",
                        CustomerJob = x.ExpensesCustomerJob ?? ""
                    })
                    .Select(g => g.Key);

                foreach (var desc in descriptionOnlyEntries)
                {
                    try
                    {
                        using (OleDbCommand command = new OleDbCommand(insertQuery, connection))
                        {
                            command.Parameters.AddWithValue("@RefNumber", refNumber);
                            command.Parameters.AddWithValue("@Particulars", SafeTruncate(desc.Description, 255));
                            command.Parameters.AddWithValue("@Class", (object)DBNull.Value);
                            command.Parameters.AddWithValue("@Debit", "");
                            command.Parameters.AddWithValue("@Credit", "");
                            command.Parameters.AddWithValue("@Memo", SafeTruncate(desc.Memo, 255));
                            command.Parameters.AddWithValue("@CustomerJob", SafeTruncate(desc.CustomerJob, 255));
                            command.ExecuteNonQuery();
                        }
                    }
                    catch { }
                }

                connection.Close();
            }

            Console.WriteLine($"Total Debit: {debitTotalAmount:F2}, Total Credit: {creditTotalAmount:F2}");
        }

        public static void InsertDataToJournalCompiled(string refNumber, List<JournalGridItem> journalData)
        {
            string connectionString = AccessToDatabase_INT.GetAccessConnectionString();

            double debitTotalAmount = 0;
            double creditTotalAmount = 0;

            // Local helper function to safely truncate text to database limits
            string SafeTruncate(string value, int maxLength)
            {
                if (string.IsNullOrEmpty(value)) return "";
                return value.Length <= maxLength ? value : value.Substring(0, maxLength);
            }

            using (OleDbConnection connection = new OleDbConnection(connectionString))
            {
                connection.Open();

                // 1. Clear old data
                string deleteQuery = "DELETE FROM JV_Compiled";
                using (OleDbCommand deleteCommand = new OleDbCommand(deleteQuery, connection))
                {
                    try
                    {
                        deleteCommand.ExecuteNonQuery();
                    }
                    catch (Exception ex)
                    {
                        MessageBox.Show($"Error deleting data: {ex.Message}");
                        return;
                    }
                }

                // 2. Prepare Insert Query (Added [AccountNumber] field)
                string insertQuery = @"
                        INSERT INTO JV_Compiled 
                        (RefNumber, [AccountNumber], [Particulars], [Class], [Name], [Debit], [Credit], [Memo]) 
                        VALUES 
                        (@RefNumber, @AccountNumber, @Particulars, @Class, @Name, @Debit, @Credit, @Memo)";

                foreach (var line in journalData)
                {
                    try
                    {
                        // MAPPING VARIABLES (With Safe Truncation to prevent DB overflow)
                        string accountNumber = SafeTruncate(line.AccountNumber, 50);
                        string particulars = SafeTruncate(line.AccountName, 500);
                        string className = line.Class;
                        string nameValue = SafeTruncate(line.Name, 255);
                        string memoValue = SafeTruncate(line.Memo, 1000);

                        string debitStr = "";
                        string creditStr = "";

                        // ---------------------------------------------------------
                        // SEPARATE DEBIT / CREDIT LOGIC
                        // ---------------------------------------------------------
                        if (line.Debit != 0)
                        {
                            debitTotalAmount += line.Debit;
                            debitStr = line.Debit.ToString("N2");
                        }
                        else if (line.Credit != 0)
                        {
                            creditTotalAmount += line.Credit;
                            creditStr = line.Credit.ToString("N2");
                        }

                        // EXECUTE INSERT
                        using (OleDbCommand command = new OleDbCommand(insertQuery, connection))
                        {
                            // OleDb relies strictly on POSITIONAL parameter matching.
                            // The order below EXACTLY matches the INSERT statement above.
                            command.Parameters.AddWithValue("@RefNumber", refNumber);
                            command.Parameters.AddWithValue("@AccountNumber", string.IsNullOrEmpty(accountNumber) ? (object)DBNull.Value : accountNumber);
                            command.Parameters.AddWithValue("@Particulars", particulars);

                            // Handle Class nulls
                            command.Parameters.AddWithValue("@Class", string.IsNullOrEmpty(className) ? (object)DBNull.Value : className);

                            // Name Parameter
                            command.Parameters.AddWithValue("@Name", string.IsNullOrEmpty(nameValue) ? (object)DBNull.Value : nameValue);

                            // Insert separated Debit and Credit strings
                            command.Parameters.AddWithValue("@Debit", debitStr);
                            command.Parameters.AddWithValue("@Credit", creditStr);

                            command.Parameters.AddWithValue("@Memo", memoValue);

                            command.ExecuteNonQuery();
                        }
                    }
                    catch (Exception ex)
                    {
                        MessageBox.Show($"Error processing journal line: {ex.Message}", "Error", MessageBoxButtons.OK, MessageBoxIcon.Error);
                    }
                }

                connection.Close();
            }

            // Console Log for verification
            Console.WriteLine($"Processed. Total Debit: {debitTotalAmount:F2}, Total Credit: {creditTotalAmount:F2}");
        }

        public static void InsertDataToBillCompiled(string refNumber, List<BillTable> bills)
        {
            string connectionString = AccessToDatabase_INT.GetAccessConnectionString();
            double debitTotalAmount = 0;
            double creditTotalAmount = 0;

            string SafeTruncate(string value, int maxLength)
            {
                if (string.IsNullOrEmpty(value)) return "";
                return value.Length <= maxLength ? value : value.Substring(0, maxLength);
            }

            using (OleDbConnection connection = new OleDbConnection(connectionString))
            {
                connection.Open();

                // 1. Clear old staging data
                string deleteQuery = "DELETE FROM Bill_Compiled";
                using (OleDbCommand deleteCommand = new OleDbCommand(deleteQuery, connection))
                {
                    try { deleteCommand.ExecuteNonQuery(); }
                    catch (Exception ex)
                    {
                        MessageBox.Show($"Error deleting data: {ex.Message}", "Error", MessageBoxButtons.OK, MessageBoxIcon.Error);
                        return;
                    }
                }

                string insertQuery = @"
        INSERT INTO Bill_Compiled 
        (RefNumber, [Particulars], [Class], [Debit], [Credit], [Memo], [CustomerJob]) 
        VALUES 
        (@RefNumber, @Particulars, @Class, @Debit, @Credit, @Memo, @CustomerJob)";

                var allDetails = bills
                    .Where(b => b.ItemDetails != null)
                    .SelectMany(b => b.ItemDetails.Select(d => new { Bill = b, Detail = d }));

                // =========================================================================
                // PASS 1: POSITIVE DEBIT EXPENSE / ITEM LINES (> 0)
                // =========================================================================
                var groupedItemDebits = allDetails
                    .Where(x => !string.IsNullOrEmpty(x.Detail.ItemLineItemRefFullName) && x.Detail.ItemLineAmount > 0)
                    .GroupBy(x => x.Detail.ItemLineItemRefFullName.Trim())
                    .Select(g => new {
                        Particulars = g.Key,
                        Memo = g.First().Detail.ItemLineMemo ?? "",
                        TotalAmount = g.Sum(x => x.Detail.ItemLineAmount)
                    });

                var groupedExpenseDebits = allDetails
                    .Where(x => !string.IsNullOrEmpty(x.Detail.ExpenseLineItemRefFullName) && x.Detail.ExpenseLineAmount > 0)
                    .GroupBy(x => x.Detail.ExpenseLineItemRefFullName.Trim())
                    .Select(g => new {
                        Particulars = g.Key,
                        Memo = g.First().Detail.ExpenseLineMemo ?? "",
                        TotalAmount = g.Sum(x => x.Detail.ExpenseLineAmount)
                    });

                foreach (var entry in groupedItemDebits.Concat(groupedExpenseDebits))
                {
                    debitTotalAmount += entry.TotalAmount;
                    using (OleDbCommand command = new OleDbCommand(insertQuery, connection))
                    {
                        command.Parameters.AddWithValue("@RefNumber", refNumber ?? (object)DBNull.Value);
                        command.Parameters.AddWithValue("@Particulars", SafeTruncate(entry.Particulars, 255));
                        command.Parameters.AddWithValue("@Class", (object)DBNull.Value);
                        command.Parameters.AddWithValue("@Debit", entry.TotalAmount.ToString("N2"));
                        command.Parameters.AddWithValue("@Credit", "");
                        command.Parameters.AddWithValue("@Memo", SafeTruncate(entry.Memo, 255));
                        command.Parameters.AddWithValue("@CustomerJob", (object)DBNull.Value);
                        command.ExecuteNonQuery();
                    }
                }

                // =========================================================================
                // PASS 2: NEGATIVE EXPENSE/ITEM LINES -> INSERT AS CREDITS (< 0)
                // (This catches the -2,420.00 Withholding Tax line from the Bill!)
                // =========================================================================
                var groupedExpenseCredits = allDetails
                    .Where(x => !string.IsNullOrEmpty(x.Detail.ExpenseLineItemRefFullName) && x.Detail.ExpenseLineAmount < 0)
                    .GroupBy(x => {
                        string rawAcc = x.Detail.ExpenseLineItemRefFullName.Trim();
                        return rawAcc.Contains(":") ? rawAcc.Split(':').Last().Trim() : rawAcc;
                    })
                    .Select(g => new {
                        Particulars = g.Key,
                        Memo = g.First().Detail.ExpenseLineMemo ?? "",
                        TotalCreditAmount = Math.Abs(g.Sum(x => x.Detail.ExpenseLineAmount))
                    });

                double totalNegativeCredits = 0;

                foreach (var entry in groupedExpenseCredits)
                {
                    totalNegativeCredits += entry.TotalCreditAmount;
                    creditTotalAmount += entry.TotalCreditAmount;

                    using (OleDbCommand command = new OleDbCommand(insertQuery, connection))
                    {
                        command.Parameters.AddWithValue("@RefNumber", refNumber ?? (object)DBNull.Value);
                        command.Parameters.AddWithValue("@Particulars", SafeTruncate(entry.Particulars, 255));
                        command.Parameters.AddWithValue("@Class", (object)DBNull.Value);
                        command.Parameters.AddWithValue("@Debit", "");
                        command.Parameters.AddWithValue("@Credit", entry.TotalCreditAmount.ToString("N2"));
                        command.Parameters.AddWithValue("@Memo", SafeTruncate(entry.Memo, 255));
                        command.Parameters.AddWithValue("@CustomerJob", (object)DBNull.Value);
                        command.ExecuteNonQuery();
                    }
                }

                // Catch payment discounts if any exist
                var groupedDiscounts = bills
                    .Where(b => b.AppliedToTxnDiscountAmount > 0 && !string.IsNullOrEmpty(b.AppliedToTxnDiscountAccountRefFullName))
                    .GroupBy(b => {
                        string rawAcc = b.AppliedToTxnDiscountAccountRefFullName;
                        return rawAcc.Contains(":") ? rawAcc.Split(':').Last().Trim() : rawAcc.Trim();
                    })
                    .Select(g => new {
                        AccountName = g.Key,
                        TotalDiscount = g.Sum(b => Math.Abs(b.AppliedToTxnDiscountAmount))
                    });

                foreach (var disc in groupedDiscounts)
                {
                    totalNegativeCredits += disc.TotalDiscount;
                    creditTotalAmount += disc.TotalDiscount;

                    using (OleDbCommand command = new OleDbCommand(insertQuery, connection))
                    {
                        command.Parameters.AddWithValue("@RefNumber", refNumber ?? (object)DBNull.Value);
                        command.Parameters.AddWithValue("@Particulars", SafeTruncate(disc.AccountName, 255));
                        command.Parameters.AddWithValue("@Class", (object)DBNull.Value);
                        command.Parameters.AddWithValue("@Debit", "");
                        command.Parameters.AddWithValue("@Credit", disc.TotalDiscount.ToString("N2"));
                        command.Parameters.AddWithValue("@Memo", SafeTruncate("Withholding Tax Applied", 255));
                        command.Parameters.AddWithValue("@CustomerJob", (object)DBNull.Value);
                        command.ExecuteNonQuery();
                    }
                }

                // =========================================================================
                // PASS 3: NET ACCOUNTS PAYABLE CREDIT ENTRY
                // =========================================================================
                if (bills != null && bills.Count > 0)
                {
                    try
                    {
                        double netPaymentCredit = debitTotalAmount - totalNegativeCredits;
                        creditTotalAmount += netPaymentCredit;

                        var mainBill = bills[0];
                        string apAccount = !string.IsNullOrEmpty(mainBill.APAccountRefFullName)
                                            ? mainBill.APAccountRefFullName
                                            : (!string.IsNullOrEmpty(mainBill.AccountName) ? mainBill.AccountName : "Accounts Payable");

                        if (apAccount.Contains(":"))
                        {
                            apAccount = apAccount.Split(':').Last().Trim();
                        }

                        using (OleDbCommand command = new OleDbCommand(insertQuery, connection))
                        {
                            command.Parameters.AddWithValue("@RefNumber", refNumber ?? (object)DBNull.Value);
                            command.Parameters.AddWithValue("@Particulars", SafeTruncate(apAccount, 255));
                            command.Parameters.AddWithValue("@Class", (object)DBNull.Value);
                            command.Parameters.AddWithValue("@Debit", "");
                            command.Parameters.AddWithValue("@Credit", netPaymentCredit.ToString("N2"));
                            command.Parameters.AddWithValue("@Memo", SafeTruncate(mainBill.BillMemo ?? mainBill.Memo, 255));
                            command.Parameters.AddWithValue("@CustomerJob", (object)DBNull.Value);
                            command.ExecuteNonQuery();
                        }
                    }
                    catch (Exception ex)
                    {
                        Console.WriteLine($"Error processing AP Credit entry: {ex.Message}");
                    }
                }

                connection.Close();
            }
        }

        private FlowLayoutPanel Panel_SBRefNumber()
        {
            FlowLayoutPanel panel_RefNumber = new FlowLayoutPanel
            {
                //Parent = panel_SideBar,
                Dock = DockStyle.Top,
                Height = 90,
                Width = sideBarWidth - 10,
                BackColor = Color.LightGray,
                Padding = new Padding(5, 2, 5, 5),
                BorderStyle = BorderStyle.FixedSingle,
                //Visible = false
            };

            Label label_RefNumberText = new Label
            {
                Parent = panel_RefNumber,
                Width = sideBarWidth - 30,
                Text = "ENTER REFERENCE NUMBER:",
                TextAlign = ContentAlignment.MiddleLeft,
                Font = font_Label,
            };

            TextBox textBox_ReferenceNumber = new TextBox
            {
                Parent = panel_RefNumber,
                Width = sideBarWidth - 30, // 190
                Font = font_Label,
            };

            Button button_SearchRefNum = new Button
            {
                Parent = panel_RefNumber,
                Height = 26,
                Width = sideBarWidth - 30,
                Text = "SEARCH",
                BackColor = Color.Transparent,
            };

            button_SearchRefNum.Click += (sender, e) =>
            {
                if (comboBox_Forms.SelectedIndex == 0)
                {
                    MessageBox.Show("Please select a form.", "Notice", MessageBoxButtons.OK);
                }
                else if (comboBox_Forms.SelectedIndex != 0 && textBox_ReferenceNumber.Text != "")
                {
                    string refNumber = textBox_ReferenceNumber.Text;
                    AccessQueries_INT queries = new AccessQueries_INT();

                    cheque = new List<CheckTable>();
                    bills = new List<BillTable>();
                    checks = new List<CheckTableExpensesAndItems>();
                    receipts = new List<ItemReciept>();
                    apvData = new List<BillTable>();
                    checkivp = new List<CheckTableGrid>();

                    object data = null;
                    
                    if (GlobalVariables.client == "INT")
                    {
                        if (comboBox_Forms.SelectedIndex == 2) // Check
                        {
                            checkivp = queries.GetCheckDataINT(refNumber);
                            data = checkivp;
                        }
                    }

                    //if (checks.Count > 0 || bills.Count > 0 || receipts.Count > 0)
                    if (data is System.Collections.ICollection colletion && colletion.Count > 0)
                    {
                        if (GlobalVariables.client == "INT")
                        {
                            Layouts_INT layouts_INT = new Layouts_INT();
                            System.Drawing.Printing.PaperSize paperSize = new System.Drawing.Printing.PaperSize("Custom", 850, 1100);
                            printDocument = new PrintDocument();
                            printDocument.DefaultPageSettings.PaperSize = paperSize;
                            printDocument.PrinterSettings.DefaultPageSettings.PaperSize = paperSize;

                            int selectedIndex = comboBox_Forms.SelectedIndex;
                            string seriesNumber = textBox_SeriesNumber.Text;

                            // Capture the override text
                            string payeeOverride = textBox_PayeeOverride.Text;

                            itemCounter = 0;
                            pageCounter = 1;
                            printPreviewControl.StartPage = 0;

                            printDocument.PrintPage += (s, ev) =>
                            {
                                // Pass payeeOverride to the layout function
                                layouts_INT.PrintPage_INT(s, ev, selectedIndex, seriesNumber, data, payeeOverride);
                            };
                        }

                        printPreviewControl.Document = printDocument;
                        printPreviewControl.Visible = true;
                        panel_Printing.Visible = true;
                    }
                    else
                    {
                        MessageBox.Show("No data found for the provided reference number.", "Error", MessageBoxButtons.OK, MessageBoxIcon.Error);
                    }

                }
                else
                {
                    MessageBox.Show("Please enter a reference number.", "Notice", MessageBoxButtons.OK);
                }
            };
            return panel_RefNumber;
        }

        public FlowLayoutPanel Panel_SBSignatory()
        {
            FlowLayoutPanel panel_Signatory = new FlowLayoutPanel
            {
                //Parent = groupBox_Signatory,
                //Parent = panel_SideBar,
                Dock = DockStyle.Top,
                Height = 141,
                Width = sideBarWidth - 10,
                //BackColor = Color.Transparent,
                BackColor = Color.LightGray,
                Padding = new Padding(5, 2, 5, 0),
                BorderStyle = BorderStyle.FixedSingle,
            };

            Label label_SignatoryText = new Label
            {
                Parent = panel_Signatory,
                Width = sideBarWidth - 30,
                Text = "SIGNATORY",
                TextAlign = ContentAlignment.MiddleCenter,
                //Font = new Font("Microsoft Sans Serif", 8, FontStyle.Bold),
                Font = font_Label,
            };

            ComboBox comboBox_Signatory = new ComboBox
            {
                Parent = panel_Signatory,
                Width = sideBarWidth - 28,
                DropDownStyle = ComboBoxStyle.DropDownList,
                Font = font_Label,
            };

            if (GlobalVariables.client == "INT")
            {
                comboBox_Signatory.Items.AddRange(new string[]
                {
                    "Select Signatory Option",
                    "Prepared By:",
                    "Certified Corrected By:",
                    "Approved By:",
                    "Received Payment By:",
                });
            }


            else
            {
                comboBox_Signatory.Items.AddRange(new string[]
                {
                    "Select Signatory Option",
                    "Prepared By:",
                    "Checked By:",
                    "Approved By:",
                    "Noted By:",
                });
            }

            comboBox_Signatory.SelectedIndex = 0;

            Label label_SignatoryName = new Label
            {
                Parent = panel_Signatory,
                Width = 48,
                Text = "Name:",
                TextAlign = ContentAlignment.MiddleLeft,
                Font = new Font("Microsoft Sans Serif", 8),
            };

            TextBox textBox_SignatoryName = new TextBox
            {
                Parent = panel_Signatory,
                Width = 165, // 250
                Font = new Font("Microsoft Sans Serif", 8),
            };

            Label label_SignatoryPosition = new Label
            {
                Parent = panel_Signatory,
                Width = 48,
                Text = "Position:",
                TextAlign = ContentAlignment.MiddleLeft,
                Font = new Font("Microsoft Sans Serif", 8),
            };

            TextBox textBox_SignatoryPosition = new TextBox
            {
                Parent = panel_Signatory,
                Width = 165, // 250
                Font = new Font("Microsoft Sans Serif", 8),
            };

            Button button_SaveSignatory = new Button
            {
                Parent = panel_Signatory,
                Height = 25,
                Width = 100,
                Text = "SAVE",
                TextAlign = ContentAlignment.MiddleCenter,
                Font = new Font("Microsoft Sans Serif", 8),
                BackColor = Color.Transparent,
            };

            Label label_SignatoryStatus = new Label
            {
                Parent = panel_Signatory,
                Height = 22,
                Width = 110,
                //Text = "Saved!",
                TextAlign = ContentAlignment.MiddleCenter,
                Font = new Font("Microsoft Sans Serif", 8),
                Margin = new Padding(0, 3, 0, 0),
            };

            button_SaveSignatory.Click += (sender, e) =>
            {
                if (comboBox_Signatory.SelectedIndex == 0)
                {
                    MessageBox.Show("Please selecet an option");
                }
                else
                {
                    string signatoryName = textBox_SignatoryName.Text;
                    string signatoryPosition = textBox_SignatoryPosition.Text;

                    int choice = comboBox_Signatory.SelectedIndex;

                    accessToDatabase.SaveSignatoryData(choice, signatoryName, signatoryPosition);
                    label_SignatoryStatus.Text = "Saved";
                }
            };

            comboBox_Signatory.SelectedIndexChanged += (sender, e) =>
            {
                if (comboBox_Signatory.SelectedIndex == 0)
                {
                    textBox_SignatoryName.Text = "";
                    textBox_SignatoryPosition.Text = "";
                }
                else
                {
                    label_SignatoryStatus.Text = "";
                    int choice = comboBox_Signatory.SelectedIndex;
                    var signatoryData = accessToDatabase.RetrieveSignatoryData(choice);

                    textBox_SignatoryName.Text = signatoryData.Name;
                    textBox_SignatoryPosition.Text = signatoryData.Position;
                }
            };

            return panel_Signatory;
        }

        private FlowLayoutPanel Panel_SBRRSignatory()
        {
            FlowLayoutPanel panel_RRSignatory = new FlowLayoutPanel
            {
                //Parent = groupBox_Signatory,
                //Parent = panel_SideBar,
                Dock = DockStyle.Top,
                Height = 106,
                Width = sideBarWidth - 10,
                //BackColor = Color.Transparent,
                BackColor = Color.LightGray,
                Padding = new Padding(5, 2, 5, 0),
                BorderStyle = BorderStyle.FixedSingle,
                //Visible = false
            };

            Label panel_Title = new Label
            {
                Parent = panel_RRSignatory,
                Dock = DockStyle.Top,
                Text = "SIGNATORY (RR)",
                Width = sideBarWidth - 30,
                //BackColor = Color.SandyBrown,
                TextAlign = ContentAlignment.MiddleCenter,
            };

            Label label_ReceivedBy = new Label
            {
                Parent = panel_RRSignatory,
                Dock = DockStyle.Top,
                Text = "Received By:",
                TextAlign = ContentAlignment.MiddleLeft,
                Width = 71,
                //BackColor = Color.ForestGreen,
            };

            textBox_ReceivedByRR = new TextBox
            {
                Parent = panel_RRSignatory,
                Dock = DockStyle.Top,
                Width = 145,
                Margin = new Padding(0, 2, 0, 0),
            };

            Label label_CheckedBy = new Label
            {
                Parent = panel_RRSignatory,
                Dock = DockStyle.Top,
                Text = "Checked By:",
                TextAlign = ContentAlignment.MiddleLeft,
                Width = 71,
                //BackColor = Color.ForestGreen,
            };

            textBox_CheckedByRR = new TextBox
            {
                Parent = panel_RRSignatory,
                Dock = DockStyle.Top,
                Width = 145,
                Margin = new Padding(0, 2, 0, 0),
            };

            Button button_SaveRRSignatory = new Button
            {
                Parent = panel_RRSignatory,
                Height = 25,
                Width = 100,
                Text = "SAVE",
                TextAlign = ContentAlignment.MiddleCenter,
                Font = new Font("Microsoft Sans Serif", 8),
                BackColor = Color.Transparent,
            };

            label_SignatoryRRStatus = new Label
            {
                Parent = panel_RRSignatory,
                Height = 22,
                Width = 110,
                //Text = "Saved!",
                TextAlign = ContentAlignment.MiddleCenter,
                Font = new Font("Microsoft Sans Serif", 8),
                Margin = new Padding(0, 3, 0, 0),
            };

            button_SaveRRSignatory.Click += (sender, e) =>
            {
                string signatoryName = textBox_ReceivedByRR.Text;
                string signatoryPosition = textBox_CheckedByRR.Text;

                //int choice = comboBox_Signatory.SelectedIndex;

                accessToDatabase.SaveSignatoryRRData(signatoryName, signatoryPosition);
                label_SignatoryRRStatus.Text = "Saved";
            };

            return panel_RRSignatory;
        }

        private FlowLayoutPanel Panel_SBPrinting()
        {
            panel_Printing = new FlowLayoutPanel
            {
                //Parent = panel_SideBar,
                Dock = DockStyle.Top,
                Height = 110,
                Width = sideBarWidth - 10,
                BackColor = Color.LightGray,
                Padding = new Padding(5),
                BorderStyle = BorderStyle.FixedSingle,
                Visible = false,
            };

            Button button_ZoomOut = new Button
            {
                Parent = panel_Printing,
                Text = "Zoom Out",
                Height = 28,
                Width = 108,
                BackColor = Color.Transparent,
            };
            button_ZoomOut.Click += (sender, e) =>
            {
                if (printPreviewControl.Zoom >= 0.1)
                {
                    printPreviewControl.Zoom -= 0.1;
                }
            };

            Button button_ZoomIn = new Button
            {
                Parent = panel_Printing,
                Text = "Zoom In",
                Height = 28,
                Width = 108,
                BackColor = Color.Transparent,
            };
            button_ZoomIn.Click += (sender, e) =>
            {
                printPreviewControl.Zoom += 0.1;
            };

            Button button_PreviousPage = new Button
            {
                Parent = panel_Printing,
                Text = "Previous Page",
                Height = 28,
                Width = 108,
                BackColor = Color.Transparent,
            };
            button_PreviousPage.Click += (sender, e) =>
            {
                if (printPreviewControl.StartPage > 0)
                {
                    printPreviewControl.StartPage--;
                }
            };

            Button button_NextPage = new Button
            {
                Parent = panel_Printing,
                Text = "Next Page",
                Height = 28,
                Width = 108,
                BackColor = Color.Transparent,
            };
            button_NextPage.Click += (sender, e) =>
            {
                if (printPreviewControl.StartPage < pageCounter - 1)
                {
                    printPreviewControl.StartPage++;
                }
            };

            Button button_Print = new Button
            {
                Parent = panel_Printing,
                Text = "Print",
                Height = 28,
                Width = 222,
                BackColor = Color.Transparent,
            };
            button_Print.Click += (sender, e) =>
            {
                try
                {
                    // Reset counters for new print job
                    itemCounter = 0;
                    pageCounter = 1;

                    if (comboBox_Forms.SelectedIndex == 4) // APV
                    {
                        int totalItemDetails = apvData.Sum(apvData => apvData.ItemDetails.Count);

                        int totalPages = (int)Math.Ceiling((double)totalItemDetails / GlobalVariables.itemsPerPageAPV);
                        Console.WriteLine($"Print: APV Data Count: {totalItemDetails}, Total Pages: {totalPages}");
                        printDocument.PrinterSettings.MaximumPage = totalPages;
                    }

                    // Update preview control to start at the first page
                    printPreviewControl.StartPage = 0;

                    PrintDialog printDialog = new PrintDialog
                    {
                        Document = printDocument,
                    };

                    if (printDialog.ShowDialog() == DialogResult.OK)
                    {
                        GlobalVariables.includeImage = false;
                        printDialog.Document.Print();

                        // Hide preview after printing
                        printPreviewControl.Visible = false;
                        printPreviewControl.Zoom = 1;
                        panel_Printing.Visible = false;


                        if (GlobalVariables.client == "LEADS")
                        {
                            string columnName = comboBox_Forms.SelectedIndex == 2 ? "CVSeries" : "APVSeries";
                            accessToDatabase.IncrementSeriesNumberInDatabase(columnName); // Increment for next print

                            seriesNumber = accessToDatabase.GetSeriesNumberFromDatabase(columnName);
                            UpdateSeriesNumber(comboBox_Forms.SelectedIndex == 2 ? "CV" : "APV");
                        }
                        else if (GlobalVariables.client == "KAYAK")
                        {
                            /* string columnName = comboBox_Forms.SelectedIndex == 1 ? "CVSeries" : "APVSeries";
                             accessToDatabase.IncrementSeriesNumberInDatabase(columnName); // Increment for next print

                             seriesNumber = accessToDatabase.GetSeriesNumberFromDatabase(columnName);
                             UpdateSeriesNumber(comboBox_Forms.SelectedIndex == 1 ? "CV" : "APV");*/
                        }
                        else if (GlobalVariables.client == "CPI")
                        {
                            string columnName = comboBox_Forms.SelectedIndex == 1 ? "CVSeries" : "APVSeries";
                            accessToDatabase.IncrementSeriesNumberInDatabase(columnName); // Increment for next print

                            seriesNumber = accessToDatabase.GetSeriesNumberFromDatabase(columnName);
                            UpdateSeriesNumber(comboBox_Forms.SelectedIndex == 1 ? "CV" : "APV");
                        }

                        else if (GlobalVariables.client == "INT")
                        {
                            // 1. Determine Form Type
                            string formType = "";
                            if (comboBox_Forms.SelectedIndex == 1) formType = "CV";
                            else if (comboBox_Forms.SelectedIndex == 3) formType = "JV";
                            else if (comboBox_Forms.SelectedIndex == 4) formType = "APV";

                            if (formType != "")
                            {
                                // 2. Get Selected Company
                                string selectedCompany = comboBox_Company.SelectedItem?.ToString();

                                if (!string.IsNullOrEmpty(selectedCompany))
                                {
                                    // 3. Increment the number in memory
                                    seriesNumber++;

                                    // 4. Save the new number to the Database using the specific company column
                                    // Note: accessToDatabase.UpdateManualSeriesNumber handles the column mapping logic you wrote earlier
                                    accessToDatabase.UpdateManualSeriesNumber(formType, seriesNumber, selectedCompany);

                                    // 5. Update the UI with the new format (e.g., CV00002)
                                    UpdateSeriesNumberINT(formType);
                                }
                            }
                        }

                    }
                }
                catch (Exception ex)
                {
                    MessageBox.Show($"An error occurred while printing: {ex.Message}", "Error", MessageBoxButtons.OK, MessageBoxIcon.Error);
                }
                GlobalVariables.includeImage = true;
            };

            return panel_Printing;
        }

        private void ComboBox_Forms_SelectedIndexChanged(object sender, EventArgs e)
        {
            if (reportViewer != null)
            {
                reportViewer.ReportSource = null;
                reportViewer.Refresh();
            }

            seriesNumber = 0;
            textBox_SeriesNumber.Text = "";

            if (GlobalVariables.client == "INT")
            {
                // 1. Hide Voucher Type controls by default at the start of every change
                SetVoucherTypeVisibility(false);

                // 2. Control panel_Company visibility (Include Online Voucher index if needed)
                // Adjust these indices to match your combo box order:
                // (e.g., 1 = CV, 3 = JV, 4 = APV, 5 = Online Voucher)
                if (comboBox_Forms.SelectedIndex == 1 || comboBox_Forms.SelectedIndex == 3 ||
                    comboBox_Forms.SelectedIndex == 4 || comboBox_Forms.SelectedIndex == 5)
                {
                    panel_Company.Visible = true;
                }
                else
                {
                    panel_Company.Visible = false;
                }

                if (panel_PayeeOverride != null) panel_PayeeOverride.Visible = false;

                string prefix = "";
                if (comboBox_Forms.SelectedIndex == 1) prefix = "CV";
                else if (comboBox_Forms.SelectedIndex == 4) prefix = "APV";
                else if (comboBox_Forms.SelectedIndex == 3) prefix = "JV";

                if (prefix != "")
                {
                    string selectedCompany = comboBox_Company.SelectedItem?.ToString();
                    if (!string.IsNullOrEmpty(selectedCompany))
                    {
                        seriesNumber = accessToDatabase.GetSeriesNumberFromDatabase(prefix, selectedCompany);
                        UpdateSeriesNumberINT(prefix);
                    }
                    else
                    {
                        textBox_SeriesNumber.Text = $"{prefix}-00000";
                    }
                }

                switch (comboBox_Forms.SelectedIndex)
                {
                    case 1: // Check Voucher
                        prefix = "CV";
                        panel_SeriesNumber.Visible = false;
                        panel_RefNumber.Visible = false;
                        panel_RefNumberCrystalReport.Visible = true;
                        panel_Signatory.Visible = true;
                        label_SeriesNumberText.Text = "Current Series Number: CV";

                        if (label_CurrencyText != null) label_CurrencyText.Visible = true;
                        if (comboBox_Currency != null) comboBox_Currency.Visible = true;

                        panel_Main.Visible = false;
                        panel_Main_CR.Visible = true;
                        break;

                    case 2: // Check
                        panel_SeriesNumber.Visible = false;
                        panel_RefNumber.Visible = true;
                        panel_RefNumberCrystalReport.Visible = false;
                        panel_Signatory.Visible = false;
                        if (panel_PayeeOverride != null) panel_PayeeOverride.Visible = true;

                        panel_Main.Visible = true;
                        panel_Main_CR.Visible = false;
                        break;

                    case 3: // Journal Voucher
                        prefix = "JV";
                        panel_RefNumber.Visible = false;
                        panel_RefNumberCrystalReport.Visible = true;
                        panel_Signatory.Visible = true;
                        panel_SeriesNumber.Visible = false;

                        panel_Main.Visible = false;
                        panel_Main_CR.Visible = true;

                        label_SeriesNumberText.Text = "Current Series Number: JV";

                        if (label_CurrencyText != null) label_CurrencyText.Visible = false;
                        if (comboBox_Currency != null) comboBox_Currency.Visible = false;

                        // ONLY SHOW VOUCHER TITLE FOR JOURNAL VOUCHER
                        SetVoucherTypeVisibility(true);
                        break;

                    case 4: // Accounts Payable Voucher
                        prefix = "APV";
                        panel_SeriesNumber.Visible = false;
                        panel_RefNumber.Visible = false;
                        panel_RefNumberCrystalReport.Visible = true;
                        panel_Signatory.Visible = true;

                        label_SeriesNumberText.Text = "Current Series Number: APV";

                        if (label_CurrencyText != null) label_CurrencyText.Visible = false;
                        if (comboBox_Currency != null) comboBox_Currency.Visible = false;

                        panel_Main.Visible = false;
                        panel_Main_CR.Visible = true;
                        break;

                    case 5: // Online Voucher (Change number to match your actual Index)
                        panel_SeriesNumber.Visible = false;
                        panel_RefNumber.Visible = false;
                        panel_RefNumberCrystalReport.Visible = true;
                        panel_Signatory.Visible = true;

                        if (label_CurrencyText != null) label_CurrencyText.Visible = false;
                        if (comboBox_Currency != null) comboBox_Currency.Visible = false;

                        panel_Main.Visible = false;
                        panel_Main_CR.Visible = true;
                        break;

                    default:
                        panel_RefNumber.Visible = false;
                        panel_RefNumberCrystalReport.Visible = false;
                        panel_Signatory.Visible = false;
                        panel_SeriesNumber.Visible = false;
                        panel_Main.Visible = false;
                        panel_Main_CR.Visible = false;
                        panel_Company.Visible = false;
                        return;
                }
            }
        }

        private void SetVoucherTypeVisibility(bool isVisible)
        {
            if (label_VoucherType != null) label_VoucherType.Visible = isVisible;
            if (comboBox_VoucherType != null) comboBox_VoucherType.Visible = isVisible;

            // Adjust company panel height dynamically based on whether Voucher Type is visible
            if (panel_Company != null)
            {
                panel_Company.Height = isVisible ? 120 : 61;
            }
        }

        private void SetDatabaseLocation(ReportDocument reportDocument, string databasePath)
        {
            // Iterate through each table in the report
            foreach (Table table in reportDocument.Database.Tables)
            {
                TableLogOnInfo tableLogOnInfo = table.LogOnInfo;

                // Update the connection information
                tableLogOnInfo.ConnectionInfo.ServerName = databasePath;
                tableLogOnInfo.ConnectionInfo.DatabaseName = ""; //or databasePath
                tableLogOnInfo.ConnectionInfo.UserID = ""; // Leave blank for Access
                tableLogOnInfo.ConnectionInfo.Password = ""; // Leave blank for Access

                // Apply the updated information to the table
                table.ApplyLogOnInfo(tableLogOnInfo);
            }

            // Update subreports if any
            foreach (Section section in reportDocument.ReportDefinition.Sections)
            {
                foreach (ReportObject reportObject in section.ReportObjects)
                {
                    if (reportObject.Kind == ReportObjectKind.SubreportObject)
                    {
                        SubreportObject subreportObject = (SubreportObject)reportObject;
                        ReportDocument subreportDocument = subreportObject.OpenSubreport(subreportObject.SubreportName);
                        SetDatabaseLocation(subreportDocument, databasePath);
                    }
                }
            }
        }

        private void UpdateSeriesNumber(string prefix)
        {
            textBox_SeriesNumber.Text = $"{prefix}{seriesNumber:000}"; // Formats seriesNumber as a 3-digit number
        }

        private void UpdateSeriesNumberINT(string formPrefix)
        {
            // Ensure accessToDatabase is initialized
            if (accessToDatabase == null) accessToDatabase = new AccessToDatabase_INT();

            // Format with a hyphen and 5 digits (e.g., CV-00001)
            textBox_SeriesNumber.Text = $"{seriesNumber:00000}";
        }

        private string SafeTruncate(string value, int maxLength)
        {
            if (string.IsNullOrEmpty(value)) return value;
            return value.Length > maxLength ? value.Substring(0, maxLength) : value;
        }
    }
}
