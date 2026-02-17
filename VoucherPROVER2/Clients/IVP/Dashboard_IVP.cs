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
using static VoucherPROVER2.Clients.IVP.Dataclass_IVP;
using System.IO;
using System.Data.OleDb;


namespace VoucherPROVER2.Clients.IVP
{
    public partial class Dashboard_IVP : Form
    {
        public Dashboard_IVP()
        {
            InitializeComponent();

            accessToDatabase = new AccessToDatabase_IVP();
        }

        private PrintDocument printDocument;
        private PrintPreviewControl printPreviewControl;
        private CrystalReportViewer reportViewer;
        private AccessToDatabase_IVP accessToDatabase;


        FlowLayoutPanel panel_Company;

        ComboBox comboBox_Forms;
        ComboBox comboBox_Company;

        Label label_SeriesNumberText;
        Label label_SignatoryRRStatus;

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
                Height = 120,
                Width = sideBarWidth - 10,
                BackColor = Color.LightGray,
                Padding = new Padding(5, 2, 5, 5),
                BorderStyle = BorderStyle.FixedSingle,
                // Only visible if client is IVP
                Visible = (GlobalVariables.client == "IVP")
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

            // ADD YOUR COMPANY NAMES HERE
            comboBox_Company.Items.AddRange(new string[]
            {

                // ---------------- IVP COMPANIES ----------------
                "North Luzon",
                "South Luzon",
                "Visayas",
                "Mindanao",
                "Metro Manila",
                "Iberica Verheilen Pharmaceuticals Group.",
                "Verheilen Iberica HealthCare Company Inc.",
                "My Health Shield NutriPharm Inc.",
                "Central Luzon",

            });

            // Set default selection
            if (comboBox_Company.Items.Count > 0)
            {
                comboBox_Company.SelectedIndex = 0;
            }

            comboBox_Company.SelectedIndexChanged += (sender, e) =>
            {
                // Only update if we are on CV (Index 1) or JV (Index 3)
                string formType = "";
                if (comboBox_Forms.SelectedIndex == 1) formType = "CV";
                else if (comboBox_Forms.SelectedIndex == 3) formType = "JV";

                if (formType != "")
                {
                    string selectedCompany = comboBox_Company.SelectedItem.ToString();

                    // Fetch the correct number for this specific company
                    seriesNumber = accessToDatabase.GetSeriesNumberFromDatabase(formType, selectedCompany);

                    // Update the textbox text
                    UpdateSeriesNumberIVP(formType);
                }
            };

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

                            if (GlobalVariables.client == "IVP")
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
                                        UpdateSeriesNumberIVP(formType);
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

            if (GlobalVariables.client == "IVP")
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
            if (GlobalVariables.client == "IVP")
            {
                comboBox_Forms.Items.AddRange(new string[]
            {
                "",
                "Check Voucher",
                "Check",
                "Journal Voucher",
                "Accounts Payable Voucher",

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
                if (GlobalVariables.client == "IVP")
                {
                    seriesNumber--;
                    string prefix = "";
                    if (comboBox_Forms.SelectedIndex == 1) prefix = "CV";
                    else if (comboBox_Forms.SelectedIndex == 3) prefix = "JV";
                    else if (comboBox_Forms.SelectedIndex == 4) prefix = "APV";

                    UpdateSeriesNumberIVP(prefix); // Use the new 5-digit formatter
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
                if (GlobalVariables.client == "IVP")
                {
                    seriesNumber++;
                    string prefix = "";
                    if (comboBox_Forms.SelectedIndex == 1) prefix = "CV";
                    else if (comboBox_Forms.SelectedIndex == 3) prefix = "JV";
                    else if (comboBox_Forms.SelectedIndex == 4) prefix = "APV";
                    UpdateSeriesNumberIVP(prefix); // Use the new 5-digit formatter
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
                    if (GlobalVariables.client == "IVP")
                    {
                        // -------------------------------------------------------------
                        // OPTION 1: CHECK VOUCHER
                        // -------------------------------------------------------------
                        if (comboBox_Forms.SelectedIndex == 1)
                        {
                            bool cvDataExists = false;
                            try
                            {
                                CRCV_IVP cRCV_IVP = new CRCV_IVP();
                                string databasePath = Path.Combine(Application.StartupPath, "CheckDatabase.accdb");
                                SetDatabaseLocation(cRCV_IVP, databasePath);

                                AccessQueries_IVP accessQueries = new AccessQueries_IVP();
                                string refNumberCR = textBox_ReferenceNumber_CR.Text;

                                cvData = accessQueries.GetCheckExpensesAndItemsData_IVP(refNumberCR);

                                if (cvData != null && cvData.Count > 0)
                                {
                                    cvDataExists = true;

                                    TextObject textObject_CVRefNumber = cRCV_IVP.ReportDefinition.ReportObjects["TextCVRefNumber"] as TextObject;
                                    //TextObject textObject_CVAmountInWords = cRCV_IVP.ReportDefinition.ReportObjects["TextCVAmountInWords"] as TextObject;
                                    TextObject textObject_CVCheckDate = cRCV_IVP.ReportDefinition.ReportObjects["TextCVCheckDate"] as TextObject;
                                    TextObject textObject_CVPayee = cRCV_IVP.ReportDefinition.ReportObjects["TextCVPayee"] as TextObject;
                                    //TextObject textObject_CVTotalAmount = cRCV_IVP.ReportDefinition.ReportObjects["TextCVTotalAmount"] as TextObject;
                                    TextObject textObject_CVTotalDebitAmount = cRCV_IVP.ReportDefinition.ReportObjects["TextCVTotalDebitAmount"] as TextObject;
                                    TextObject textObject_CVTotalCreditAmount = cRCV_IVP.ReportDefinition.ReportObjects["TextCVTotalCreditAmount"] as TextObject;

                                    TextObject textObject_CompanyName = cRCV_IVP.ReportDefinition.ReportObjects["TextCompanyName"] as TextObject;
                                    if (textObject_CompanyName != null && comboBox_Company != null && comboBox_Company.SelectedItem != null)
                                    {
                                        textObject_CompanyName.Text = comboBox_Company.SelectedItem.ToString();
                                    }

                                    TextObject textObject_PreparedBy = cRCV_IVP.ReportDefinition.ReportObjects["TextPreparedBy"] as TextObject;
                                    TextObject textObject_PreparedByPos = cRCV_IVP.ReportDefinition.ReportObjects["TextPreparedByPosition"] as TextObject;
                                    TextObject textObject_CheckedBy = cRCV_IVP.ReportDefinition.ReportObjects["TextCheckedBy"] as TextObject;
                                    TextObject textObject_CheckedByPos = cRCV_IVP.ReportDefinition.ReportObjects["TextCheckedByPosition"] as TextObject;
                                    TextObject textObject_ApprovedBy = cRCV_IVP.ReportDefinition.ReportObjects["TextApprovedBy"] as TextObject;
                                    TextObject textObject_ApprovedByPos = cRCV_IVP.ReportDefinition.ReportObjects["TextApprovedByPosition"] as TextObject;
                                    TextObject textObject_ReceivedBy = cRCV_IVP.ReportDefinition.ReportObjects["TextReceivedBy"] as TextObject;
                                    TextObject textObject_ReceivedByPos = cRCV_IVP.ReportDefinition.ReportObjects["TextReceivedByPosition"] as TextObject;

                                    AccessToDatabase_IVP accessToDatabase = new AccessToDatabase_IVP();
                                    var signatories = accessToDatabase.RetrieveAllSignatoryData();

                                    double amount = cvData[0].TotalAmount;
                                    string amountInWords = AccessToDatabase_IVP.AmountToWordsConverter.Convert(amount);

                                    textObject_CVRefNumber.Text = textBox_SeriesNumber.Text;
                                    //textObject_CVAmountInWords.Text = amountInWords;
                                    textObject_CVCheckDate.Text = DateTime.Now.ToString("MMMM dd, yyyy");
                                    textObject_CVPayee.Text = cvData[0].PayeeFullName;
                                    //textObject_CVTotalAmount.Text = cvData[0].TotalAmount.ToString("N2");

                                    textObject_PreparedBy.Text = signatories.PreparedByName;
                                    textObject_PreparedByPos.Text = signatories.PreparedByPosition;
                                    textObject_CheckedBy.Text = signatories.ReviewedByName;
                                    textObject_CheckedByPos.Text = signatories.ReviewedByPosition;
                                    textObject_ApprovedBy.Text = signatories.ApprovedByName;
                                    textObject_ApprovedByPos.Text = signatories.ApprovedByPosition;
                                    textObject_ReceivedBy.Text = signatories.ReceivedByName;
                                    textObject_ReceivedByPos.Text = signatories.ReceivedByPosition;

                                    double debitTotalAmount = 0;
                                    double creditTotalAmount = 0;

                                    foreach (var data in cvData)
                                    {
                                        try
                                        {
                                            double itemAmount = data.ItemAmount;
                                            if (itemAmount > 0) debitTotalAmount += itemAmount;
                                            else if (itemAmount < 0) creditTotalAmount += Math.Abs(itemAmount);

                                            if (!string.IsNullOrEmpty(data.Account))
                                            {
                                                double expenseAmount = data.ExpensesAmount;
                                                if (expenseAmount > 0) debitTotalAmount += expenseAmount;
                                                else if (expenseAmount < 0) creditTotalAmount += Math.Abs(expenseAmount);
                                            }
                                        }
                                        catch (Exception ex) { MessageBox.Show($"Error computing totals: {ex.Message}"); }
                                    }

                                    textObject_CVTotalDebitAmount.Text = debitTotalAmount.ToString("N2");
                                    textObject_CVTotalCreditAmount.Text = debitTotalAmount.ToString("N2");

                                    SubreportObject subreportObject = cRCV_IVP.ReportDefinition.ReportObjects["SubreportCVDetailsIVP"] as SubreportObject;
                                    if (subreportObject != null)
                                    {
                                        ReportDocument subReportDocument = cRCV_IVP.OpenSubreport(subreportObject.SubreportName);
                                        TextObject textObject_Remarks = subReportDocument.ReportDefinition.ReportObjects["TextRemarks"] as TextObject;
                                        TextObject textObject_CVSubTotal = subReportDocument.ReportDefinition.ReportObjects["TextCVSubTotalAmount"] as TextObject;
                                        TextObject textObject_CVSubCheckNumber = subReportDocument.ReportDefinition.ReportObjects["TextCVSubCheckNumber"] as TextObject;
                                        TextObject textObject_CVSubCheckDate = subReportDocument.ReportDefinition.ReportObjects["TextCVSubCheckDate"] as TextObject;
                                        TextObject textObject_SubAccountPayable = subReportDocument.ReportDefinition.ReportObjects["TextSubAccountPayable"] as TextObject;
                                        TextObject textObject_SubAmountPayable = subReportDocument.ReportDefinition.ReportObjects["TextSubAmountPayable"] as TextObject;

                                        TextObject textObject_PaidSign = subReportDocument.ReportDefinition.ReportObjects["TextPaidSign"] as TextObject;
                                        if (textObject_PaidSign != null)
                                        {
                                            // Index 0 is Peso, Index 1 is Dollar
                                            textObject_PaidSign.Text = comboBox_Currency.SelectedIndex == 1 ? "$" : "₱";
                                        }

                                        textObject_Remarks.Text = cvData[0].Memo;
                                        textObject_CVSubTotal.Text = cvData[0].TotalAmount.ToString("N2");
                                        textObject_CVSubCheckNumber.Text = cvData[0].RefNumber;
                                        textObject_CVSubCheckDate.Text = cvData[0].DueDate.ToString("MMMM dd, yyyy");
                                        textObject_SubAccountPayable.Text = cvData[0].BankAccount;
                                        textObject_SubAmountPayable.Text = debitTotalAmount.ToString("N2");

                                        InsertDataToCheckVoucherCompiledIVP(refNumberCR, cvData);
                                    }

                                    cRCV_IVP.SetParameterValue("ReferenceNumber", refNumberCR);

                                    panel_Printing.Visible = false;
                                    panel_Signatory.Visible = true;
                                    panel_Main.Visible = false;
                                    panel_Main_CR.Visible = true;

                                    reportViewer.ReportSource = cRCV_IVP;
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
                                GenerateBillPaymentReport_IVP(refNumberCR);
                            }
                        }

                        else if (comboBox_Forms.SelectedIndex == 3)
                        {
                            CRJV_IVP cRJV_IVP = new CRJV_IVP();
                            string databasePath = Path.Combine(Application.StartupPath, "CheckDatabase.accdb");
                            SetDatabaseLocation(cRJV_IVP, databasePath);

                            AccessQueries_IVP accessQueries = new AccessQueries_IVP();
                            string refNumberCR = textBox_ReferenceNumber_CR.Text;

                            // 1. Get the correct data
                            journal = accessQueries.GetJournalEntryForGrid(refNumberCR);

                            if (journal != null && journal.Count > 0)
                            {
                                // 2. Set Header Text Objects
                                TextObject textObject_JVRefNumber = cRJV_IVP.ReportDefinition.ReportObjects["TextJVRefNumber"] as TextObject;
                                TextObject textObject_JVCheckDate = cRJV_IVP.ReportDefinition.ReportObjects["TextJVCheckDate"] as TextObject;
                                TextObject textObject_JVTransactDate = cRJV_IVP.ReportDefinition.ReportObjects["TextJVTransactDate"] as TextObject;
                                TextObject textObject_JVTotalDebitAmount = cRJV_IVP.ReportDefinition.ReportObjects["TextJVTotalDebitAmount"] as TextObject;
                                TextObject textObject_JVTotalCreditAmount = cRJV_IVP.ReportDefinition.ReportObjects["TextJVTotalCreditAmount"] as TextObject;

                                TextObject textObject_CompanyName = cRJV_IVP.ReportDefinition.ReportObjects["TextCompanyName"] as TextObject;
                                if (textObject_CompanyName != null && comboBox_Company != null && comboBox_Company.SelectedItem != null)
                                {
                                    textObject_CompanyName.Text = comboBox_Company.SelectedItem.ToString();
                                }


                                TextObject textObject_PreparedBy = cRJV_IVP.ReportDefinition.ReportObjects["TextPreparedBy"] as TextObject;
                                TextObject textObject_PreparedByPos = cRJV_IVP.ReportDefinition.ReportObjects["TextPreparedByPosition"] as TextObject;
                                TextObject textObject_CheckedBy = cRJV_IVP.ReportDefinition.ReportObjects["TextCheckedBy"] as TextObject;
                                TextObject textObject_CheckedByPos = cRJV_IVP.ReportDefinition.ReportObjects["TextCheckedByPosition"] as TextObject;
                                TextObject textObject_ApprovedBy = cRJV_IVP.ReportDefinition.ReportObjects["TextApprovedBy"] as TextObject;
                                TextObject textObject_ApprovedByPos = cRJV_IVP.ReportDefinition.ReportObjects["TextApprovedByPosition"] as TextObject;

                                if (textObject_JVRefNumber != null) textObject_JVRefNumber.Text = textBox_SeriesNumber.Text;
                                if (textObject_JVCheckDate != null) textObject_JVCheckDate.Text = DateTime.Now.ToString("MMMM dd, yyyy");
                                if (textObject_JVTransactDate != null) textObject_JVTransactDate.Text = journal[0].Date.ToString("MMMM dd, yyyy");

                                double debitTotalAmount = 0;
                                double creditTotalAmount = 0;

                                foreach (var line in journal)
                                {
                                    debitTotalAmount += line.Debit;
                                    creditTotalAmount += line.Credit;
                                }
                                if (textObject_JVTotalDebitAmount != null) textObject_JVTotalDebitAmount.Text = debitTotalAmount.ToString("N2");
                                if (textObject_JVTotalCreditAmount != null) textObject_JVTotalCreditAmount.Text = creditTotalAmount.ToString("N2");


                                AccessToDatabase_IVP accessToDatabase = new AccessToDatabase_IVP();
                                var signatories = accessToDatabase.RetrieveAllSignatoryData();


                                textObject_PreparedBy.Text = signatories.PreparedByName;
                                textObject_PreparedByPos.Text = signatories.PreparedByPosition;
                                textObject_CheckedBy.Text = signatories.ReviewedByName;
                                textObject_CheckedByPos.Text = signatories.ReviewedByPosition;
                                textObject_ApprovedBy.Text = signatories.ApprovedByName;
                                textObject_ApprovedByPos.Text = signatories.ApprovedByPosition;

                                // 4. Handle Subreport
                                SubreportObject subreportObject = cRJV_IVP.ReportDefinition.ReportObjects["SubreportJVDetailsIVP"] as SubreportObject;
                                if (subreportObject != null)
                                {
                                    ReportDocument subReportDocument = cRJV_IVP.OpenSubreport(subreportObject.SubreportName);

                                    TextObject textObject_SubAccountPayable = subReportDocument.ReportDefinition.ReportObjects["TextJVSUBAccountsPayable"] as TextObject;
                                    TextObject textObject_SubAmountPayable = subReportDocument.ReportDefinition.ReportObjects["TextJVSUBAmountPayable"] as TextObject;


                                    if (textObject_SubAccountPayable != null) textObject_SubAccountPayable.Text = journal[0].AccountName;

                                    if (textObject_SubAmountPayable != null) textObject_SubAmountPayable.Text = debitTotalAmount.ToString("N2");
                                }

                                InsertDataToJournalCompiled(refNumberCR, journal);

                                // 6. Final Report Settings
                                cRJV_IVP.SetParameterValue("ReferenceNumber", refNumberCR);

                                panel_Printing.Visible = false;
                                panel_Signatory.Visible = true;
                                panel_Main.Visible = false;
                                panel_Main_CR.Visible = true;

                                reportViewer.ReportSource = cRJV_IVP;
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
                            GenerateAPVReport_IVP(refNumberCR);
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

        private bool GenerateAPVReport_IVP(string refNumberCR)
        {
            try
            {
                CRAPV_IVPBILL cRAPV_IVPBILL = new CRAPV_IVPBILL();
                string databasePathBILL = Path.Combine(Application.StartupPath, "CheckDatabase.accdb");
                SetDatabaseLocation(cRAPV_IVPBILL, databasePathBILL);

                AccessQueries_IVP accessQueries = new AccessQueries_IVP();
                List<BillTable> bills = accessQueries.GetBillData_IVP_DirectBill(refNumberCR);

                if (bills == null || bills.Count == 0)
                    return false;

                TextObject textObject_CVBILLCheckNumber = null;
                //TextObject textObject_CVBILLAmountInWords = null;
                TextObject textObject_CVBILLCheckDate = null;
                TextObject textObject_CVBILLPayee = null;
                TextObject textObject_CVBILLAddress = null;
                TextObject textObject_CVBILLTIN = null;
                TextObject textObject_CVBILLCurrency = null;
                TextObject textObject_CVBILLCurrate = null;
                //TextObject textObject_CVBILLTotalAmount = null;
                TextObject textObject_CVBILLTotalDebitAmount = null;
                TextObject textObject_CVBILLTotalCreditAmount = null;
                TextObject textObject_PreparedBy = null;
                TextObject textObject_PreparedByPos = null;
                TextObject textObject_CheckedBy = null;
                TextObject textObject_CheckedByPos = null;
                TextObject textObject_ApprovedBy = null;
                TextObject textObject_ApprovedByPos = null;
                TextObject textObject_ReceivedBy = null;
                TextObject textObject_ReceivedByPos = null;

                try
                {
                    textObject_CVBILLCheckNumber = cRAPV_IVPBILL.ReportDefinition.ReportObjects["TextCVBILLSeriesnumber"] as TextObject;
                    textObject_CVBILLAddress = cRAPV_IVPBILL.ReportDefinition.ReportObjects["TextCVBILLAddress"] as TextObject;
                    textObject_CVBILLTIN = cRAPV_IVPBILL.ReportDefinition.ReportObjects["TextCVBILLTIN"] as TextObject;
                    textObject_CVBILLCurrency = cRAPV_IVPBILL.ReportDefinition.ReportObjects["TextCVBILLCurrency"] as TextObject;
                    textObject_CVBILLCurrate = cRAPV_IVPBILL.ReportDefinition.ReportObjects["TextCVBILLCurrate"] as TextObject;
                    //textObject_CVBILLAmountInWords = cRCV_IVPBILL.ReportDefinition.ReportObjects["TextCVBILLAmountInWords"] as TextObject;
                    textObject_CVBILLCheckDate = cRAPV_IVPBILL.ReportDefinition.ReportObjects["TextCVBILLCheckDate"] as TextObject;
                    textObject_CVBILLPayee = cRAPV_IVPBILL.ReportDefinition.ReportObjects["TextCVBILLPayee"] as TextObject;
                    //textObject_CVBILLTotalAmount = cRCV_IVPBILL.ReportDefinition.ReportObjects["TextCVBILLTotalAmount"] as TextObject;
                    textObject_CVBILLTotalDebitAmount = cRAPV_IVPBILL.ReportDefinition.ReportObjects["TextCVBILLTotalDebitAmount"] as TextObject;
                    textObject_CVBILLTotalCreditAmount = cRAPV_IVPBILL.ReportDefinition.ReportObjects["TextCVBILLTotalCreditAmount"] as TextObject;

                    TextObject textObject_CompanyName = cRAPV_IVPBILL.ReportDefinition.ReportObjects["TextCompanyName"] as TextObject;
                    if (textObject_CompanyName != null && comboBox_Company != null && comboBox_Company.SelectedItem != null)
                    {
                        textObject_CompanyName.Text = comboBox_Company.SelectedItem.ToString();
                    }




                    textObject_PreparedBy = cRAPV_IVPBILL.ReportDefinition.ReportObjects["TextPreparedBy"] as TextObject;
                    textObject_PreparedByPos = cRAPV_IVPBILL.ReportDefinition.ReportObjects["TextPreparedByPosition"] as TextObject;
                    textObject_CheckedBy = cRAPV_IVPBILL.ReportDefinition.ReportObjects["TextCheckedBy"] as TextObject;
                    textObject_CheckedByPos = cRAPV_IVPBILL.ReportDefinition.ReportObjects["TextCheckedByPosition"] as TextObject;
                    textObject_ApprovedBy = cRAPV_IVPBILL.ReportDefinition.ReportObjects["TextApprovedBy"] as TextObject;
                    textObject_ApprovedByPos = cRAPV_IVPBILL.ReportDefinition.ReportObjects["TextApprovedByPosition"] as TextObject;
                    textObject_ReceivedBy = cRAPV_IVPBILL.ReportDefinition.ReportObjects["TextReceivedBy"] as TextObject;
                    textObject_ReceivedByPos = cRAPV_IVPBILL.ReportDefinition.ReportObjects["TextReceivedByPosition"] as TextObject;

                    AccessToDatabase_IVP accessToDatabase = new AccessToDatabase_IVP();

                    var (PreparedByName, PreparedByPosition,
                       ReviewedByName, ReviewedByPosition,
                       RecommendingApprovalName, RecommendingApprovalPosition,
                       ApprovedByName, ApprovedByPosition,
                       ReceivedByName, ReceivedByPosition) = accessToDatabase.RetrieveAllSignatoryData();


                    double debitTotalAmount = 0;
                    double creditTotalAmount = 0;

                    textObject_PreparedBy.Text = PreparedByName;
                    textObject_PreparedByPos.Text = PreparedByPosition;
                    textObject_CheckedBy.Text = ReviewedByName;
                    textObject_CheckedByPos.Text = ReviewedByPosition;
                    textObject_ApprovedBy.Text = ApprovedByName;
                    textObject_ApprovedByPos.Text = ApprovedByPosition;
                    textObject_ReceivedBy.Text = ReceivedByName;
                    textObject_ReceivedByPos.Text = ReceivedByPosition;

                    foreach (var bill in bills) // 'bills' is List<BillTable>
                    {
                        foreach (var item in bill.ItemDetails)
                        {
                            try
                            {
                                // Handle ItemLineAmount
                                if (item.ItemLineAmount != 0)
                                {
                                    if (item.ItemLineAmount > 0)
                                        debitTotalAmount += item.ItemLineAmount;
                                    else
                                        creditTotalAmount += Math.Abs(item.ItemLineAmount);
                                }

                                // Handle ExpenseLineAmount
                                if (item.ExpenseLineAmount != 0)
                                {
                                    if (item.ExpenseLineAmount > 0)
                                        debitTotalAmount += item.ExpenseLineAmount;
                                    else
                                        creditTotalAmount += Math.Abs(item.ExpenseLineAmount);
                                }
                            }
                            catch (Exception ex)
                            {
                                MessageBox.Show($"Error processing item detail: {ex.Message}", "Error", MessageBoxButtons.OK, MessageBoxIcon.Error);
                            }
                        }
                    }

                    textObject_CVBILLTotalDebitAmount.Text = debitTotalAmount.ToString("N2");
                    textObject_CVBILLTotalCreditAmount.Text = debitTotalAmount.ToString("N2");

                }
                catch
                {
                    throw;
                }


                double amount = bills[0].AmountDue;
                string amountInWords = AccessToDatabase_IVP.AmountToWordsConverter.Convert(amount);
                var b = bills[0];

                // Line 1: Combine Addr1, Addr2, Addr3, Addr4 into one string separated by commas
                string streetLine = string.Join(", ", new[] {
                     b.VendorAddressAddr1,
                     b.VendorAddressAddr2,
                     b.VendorAddressAddr3,
                     b.VendorAddressAddr4
                 }.Where(s => !string.IsNullOrWhiteSpace(s)));

                // Line 2: City (Add State/Zip here if you have them in your BillTable)
                string cityLine = string.Join(" ", new[] {
                     b.VendorAddressCity,
                 }.Where(s => !string.IsNullOrWhiteSpace(s)));

                // Final: Join the two lines with a single NewLine
                string fullAddress = string.Join(Environment.NewLine, new[] { streetLine, cityLine }.Where(s => !string.IsNullOrWhiteSpace(s)));

                if (textObject_CVBILLCheckNumber != null) textObject_CVBILLCheckNumber.Text = textBox_SeriesNumber.Text;
                if (textObject_CVBILLAddress != null) textObject_CVBILLAddress.Text = fullAddress;
                //if (textObject_CVBILLAmountInWords != null) textObject_CVBILLAmountInWords.Text = amountInWords;
                if (textObject_CVBILLCheckDate != null) textObject_CVBILLCheckDate.Text = DateTime.Now.ToString("MMMM dd, yyyy");
                if (textObject_CVBILLPayee != null) textObject_CVBILLPayee.Text = bills[0].PayeeFullName ?? "";
                //if (textObject_CVBILLTotalAmount != null) textObject_CVBILLTotalAmount.Text = bills[0].AmountDue.ToString("N2");

                if (textObject_CVBILLTIN != null) textObject_CVBILLTIN.Text = bills[0].Tin ?? "";
                if (textObject_CVBILLCurrency != null) textObject_CVBILLCurrency.Text = bills[0].Currency ?? "";
                if (textObject_CVBILLCurrate != null) textObject_CVBILLCurrate.Text = bills[0].Exchangerate.ToString("N2");

                SubreportObject subreportObject = null;
                try
                {
                    subreportObject = cRAPV_IVPBILL.ReportDefinition.ReportObjects["SubreportCVBILLDetailsIVP"] as SubreportObject;
                }
                catch
                {
                    throw;
                }

                if (subreportObject != null)
                {
                    ReportDocument subReportDocument = null;
                    try
                    {
                        subReportDocument = cRAPV_IVPBILL.OpenSubreport(subreportObject.SubreportName);
                    }
                    catch
                    {
                        throw;
                    }

                    try
                    {
                        TextObject textObject_BILLSubRemarks = subReportDocument.ReportDefinition.ReportObjects["TextBILLRemarks"] as TextObject;
                        TextObject textObject_BILLCVSubCheckDate = subReportDocument.ReportDefinition.ReportObjects["TextCVBILLSubCheckDate"] as TextObject;
                        TextObject textObject_BILLCVSubTotal = subReportDocument.ReportDefinition.ReportObjects["TextCVBILLSUBTotalAmount"] as TextObject;
                        TextObject textObject_BILLCVSubCheckNumber = subReportDocument.ReportDefinition.ReportObjects["TextCVBILLSubCheckNumber"] as TextObject;
                        TextObject textObject_BILLCVTERMS = subReportDocument.ReportDefinition.ReportObjects["TextCVBILLSubTerms"] as TextObject;
                        //TextObject textObject_BILLSubAccountPayable = subReportDocument.ReportDefinition.ReportObjects["TextBILLSubAccountPayable"] as TextObject;
                        TextObject textObject_BILLSubAmountPayable = subReportDocument.ReportDefinition.ReportObjects["TextBILLSubAmountPayable"] as TextObject;
                        TextObject textObject_PaidSign = subReportDocument.ReportDefinition.ReportObjects["TextPaidSign"] as TextObject;
                        if (textObject_PaidSign != null)
                        {
                            textObject_PaidSign.Text = comboBox_Currency.SelectedIndex == 1 ? "$" : "₱";
                        }

                        if (textObject_BILLSubRemarks != null) textObject_BILLSubRemarks.Text = bills[0].Memo ?? "";
                        if (textObject_BILLCVTERMS != null) textObject_BILLCVTERMS.Text = bills[0].TermsRefFullName ?? "";
                        if (textObject_BILLCVSubCheckDate != null) textObject_BILLCVSubCheckDate.Text = bills[0].DueDate.ToString("MMMM dd, yyyy");
                        if (textObject_BILLCVSubTotal != null) textObject_BILLCVSubTotal.Text = bills[0].AmountDue.ToString("N2");
                        if (textObject_BILLCVSubCheckNumber != null) textObject_BILLCVSubCheckNumber.Text = bills[0].RefNumber ?? "";
                        //if (textObject_BILLCVSubCheckNumber != null) textObject_BILLSubAccountPayable.Text = bills[0].BankAccount ?? "";
                        if (textObject_BILLSubAmountPayable != null)
                        {
                            // Sums the AmountDue of all items in the bills list
                            double totalAmountDue = bills.Sum(a => a.AmountDue);
                            textObject_BILLSubAmountPayable.Text = totalAmountDue.ToString("N2");
                        }

                        InsertDataToBillCompiled(refNumberCR, bills);
                    }
                    catch
                    {
                        throw;
                    }
                }

                cRAPV_IVPBILL.SetParameterValue("ReferenceNumber", refNumberCR);

                panel_Printing.Visible = false;
                panel_Signatory.Visible = true;
                panel_Main.Visible = false;
                panel_Main_CR.Visible = true;

                reportViewer.ReportSource = cRAPV_IVPBILL;
                reportViewer.RefreshReport();

                return true;
            }
            catch (Exception ex)
            {
                MessageBox.Show($"KAYAK ERROR HEHEHE:\n{ex.Message}", "Error", MessageBoxButtons.OK, MessageBoxIcon.Error);
                return false;
            }
        }

        private bool GenerateBillPaymentReport_IVP(string refNumberCR)
        {
            try
            {
                CRCV_IVPBILL cRCV_IVPBILL = new CRCV_IVPBILL();
                string databasePathBILL = Path.Combine(Application.StartupPath, "CheckDatabase.accdb");
                SetDatabaseLocation(cRCV_IVPBILL, databasePathBILL);

                AccessQueries_IVP accessQueries = new AccessQueries_IVP();
                List<BillTable> bills = accessQueries.GetBillData_IVP(refNumberCR);

                if (bills == null || bills.Count == 0)
                    return false;

                TextObject textObject_CVBILLCheckNumber = null;
                //TextObject textObject_CVBILLAmountInWords = null;
                TextObject textObject_CVBILLCheckDate = null;
                TextObject textObject_CVBILLPayee = null;
                //TextObject textObject_CVBILLTotalAmount = null;
                TextObject textObject_CVBILLTotalDebitAmount = null;
                TextObject textObject_CVBILLTotalCreditAmount = null;
                TextObject textObject_PreparedBy = null;
                TextObject textObject_PreparedByPos = null;
                TextObject textObject_CheckedBy = null;
                TextObject textObject_CheckedByPos = null;
                TextObject textObject_ApprovedBy = null;
                TextObject textObject_ApprovedByPos = null;
                TextObject textObject_ReceivedBy = null;
                TextObject textObject_ReceivedByPos = null;

                try
                {
                    textObject_CVBILLCheckNumber = cRCV_IVPBILL.ReportDefinition.ReportObjects["TextCVBILLSeriesnumber"] as TextObject;
                    //textObject_CVBILLAmountInWords = cRCV_IVPBILL.ReportDefinition.ReportObjects["TextCVBILLAmountInWords"] as TextObject;
                    textObject_CVBILLCheckDate = cRCV_IVPBILL.ReportDefinition.ReportObjects["TextCVBILLCheckDate"] as TextObject;
                    textObject_CVBILLPayee = cRCV_IVPBILL.ReportDefinition.ReportObjects["TextCVBILLPayee"] as TextObject;
                    //textObject_CVBILLTotalAmount = cRCV_IVPBILL.ReportDefinition.ReportObjects["TextCVBILLTotalAmount"] as TextObject;
                    textObject_CVBILLTotalDebitAmount = cRCV_IVPBILL.ReportDefinition.ReportObjects["TextCVBILLTotalDebitAmount"] as TextObject;
                    textObject_CVBILLTotalCreditAmount = cRCV_IVPBILL.ReportDefinition.ReportObjects["TextCVBILLTotalCreditAmount"] as TextObject;

                    TextObject textObject_CompanyName = cRCV_IVPBILL.ReportDefinition.ReportObjects["TextCompanyName"] as TextObject;
                    if (textObject_CompanyName != null && comboBox_Company != null && comboBox_Company.SelectedItem != null)
                    {
                        textObject_CompanyName.Text = comboBox_Company.SelectedItem.ToString();
                    }


                    textObject_PreparedBy = cRCV_IVPBILL.ReportDefinition.ReportObjects["TextPreparedBy"] as TextObject;
                    textObject_PreparedByPos = cRCV_IVPBILL.ReportDefinition.ReportObjects["TextPreparedByPosition"] as TextObject;
                    textObject_CheckedBy = cRCV_IVPBILL.ReportDefinition.ReportObjects["TextCheckedBy"] as TextObject;
                    textObject_CheckedByPos = cRCV_IVPBILL.ReportDefinition.ReportObjects["TextCheckedByPosition"] as TextObject;
                    textObject_ApprovedBy = cRCV_IVPBILL.ReportDefinition.ReportObjects["TextApprovedBy"] as TextObject;
                    textObject_ApprovedByPos = cRCV_IVPBILL.ReportDefinition.ReportObjects["TextApprovedByPosition"] as TextObject;
                    textObject_ReceivedBy = cRCV_IVPBILL.ReportDefinition.ReportObjects["TextReceivedBy"] as TextObject;
                    textObject_ReceivedByPos = cRCV_IVPBILL.ReportDefinition.ReportObjects["TextReceivedByPosition"] as TextObject;

                    AccessToDatabase_IVP accessToDatabase = new AccessToDatabase_IVP();

                    var (PreparedByName, PreparedByPosition,
                       ReviewedByName, ReviewedByPosition,
                       RecommendingApprovalName, RecommendingApprovalPosition,
                       ApprovedByName, ApprovedByPosition,
                       ReceivedByName, ReceivedByPosition) = accessToDatabase.RetrieveAllSignatoryData();


                    double debitTotalAmount = 0;
                    double creditTotalAmount = 0;

                    textObject_PreparedBy.Text = PreparedByName;
                    textObject_PreparedByPos.Text = PreparedByPosition;
                    textObject_CheckedBy.Text = ReviewedByName;
                    textObject_CheckedByPos.Text = ReviewedByPosition;
                    textObject_ApprovedBy.Text = ApprovedByName;
                    textObject_ApprovedByPos.Text = ApprovedByPosition;
                    textObject_ReceivedBy.Text = ReceivedByName;
                    textObject_ReceivedByPos.Text = ReceivedByPosition;

                    foreach (var bill in bills) // 'bills' is List<BillTable>
                    {
                        foreach (var item in bill.ItemDetails)
                        {
                            try
                            {
                                // Handle ItemLineAmount
                                if (item.ItemLineAmount != 0)
                                {
                                    if (item.ItemLineAmount > 0)
                                        debitTotalAmount += item.ItemLineAmount;
                                    else
                                        creditTotalAmount += Math.Abs(item.ItemLineAmount);
                                }

                                // Handle ExpenseLineAmount
                                if (item.ExpenseLineAmount != 0)
                                {
                                    if (item.ExpenseLineAmount > 0)
                                        debitTotalAmount += item.ExpenseLineAmount;
                                    else
                                        creditTotalAmount += Math.Abs(item.ExpenseLineAmount);
                                }
                            }
                            catch (Exception ex)
                            {
                                MessageBox.Show($"Error processing item detail: {ex.Message}", "Error", MessageBoxButtons.OK, MessageBoxIcon.Error);
                            }
                        }
                    }

                    textObject_CVBILLTotalDebitAmount.Text = debitTotalAmount.ToString("N2");
                    textObject_CVBILLTotalCreditAmount.Text = debitTotalAmount.ToString("N2");

                }
                catch
                {
                    throw;
                }


                double amount = bills[0].AmountDue;
                string amountInWords = AccessToDatabase_IVP.AmountToWordsConverter.Convert(amount);

                if (textObject_CVBILLCheckNumber != null) textObject_CVBILLCheckNumber.Text = textBox_SeriesNumber.Text;
                //if (textObject_CVBILLAmountInWords != null) textObject_CVBILLAmountInWords.Text = amountInWords;
                if (textObject_CVBILLCheckDate != null) textObject_CVBILLCheckDate.Text = DateTime.Now.ToString("MMMM dd, yyyy");
                if (textObject_CVBILLPayee != null) textObject_CVBILLPayee.Text = bills[0].PayeeFullName ?? "";
                //if (textObject_CVBILLTotalAmount != null) textObject_CVBILLTotalAmount.Text = bills[0].AmountDue.ToString("N2");

                SubreportObject subreportObject = null;
                try
                {
                    subreportObject = cRCV_IVPBILL.ReportDefinition.ReportObjects["SubreportCVBILLDetailsIVP"] as SubreportObject;
                }
                catch
                {
                    throw;
                }

                if (subreportObject != null)
                {
                    ReportDocument subReportDocument = null;
                    try
                    {
                        subReportDocument = cRCV_IVPBILL.OpenSubreport(subreportObject.SubreportName);
                    }
                    catch
                    {
                        throw;
                    }

                    try
                    {
                        TextObject textObject_BILLSubRemarks = subReportDocument.ReportDefinition.ReportObjects["TextBILLRemarks"] as TextObject;
                        TextObject textObject_BILLCVSubCheckDate = subReportDocument.ReportDefinition.ReportObjects["TextCVBILLSubCheckDate"] as TextObject;
                        TextObject textObject_BILLCVSubTotal = subReportDocument.ReportDefinition.ReportObjects["TextCVBILLSUBTotalAmount"] as TextObject;
                        TextObject textObject_BILLCVSubCheckNumber = subReportDocument.ReportDefinition.ReportObjects["TextCVBILLSubCheckNumber"] as TextObject;
                        TextObject textObject_BILLSubAccountPayable = subReportDocument.ReportDefinition.ReportObjects["TextBILLSubAccountPayable"] as TextObject;
                        TextObject textObject_BILLSubAmountPayable = subReportDocument.ReportDefinition.ReportObjects["TextBILLSubAmountPayable"] as TextObject;
                        TextObject textObject_PaidSign = subReportDocument.ReportDefinition.ReportObjects["TextPaidSign"] as TextObject;
                        if (textObject_PaidSign != null)
                        {
                            textObject_PaidSign.Text = comboBox_Currency.SelectedIndex == 1 ? "$" : "₱";
                        }

                        if (textObject_BILLSubRemarks != null) textObject_BILLSubRemarks.Text = bills[0].BillMemo ?? "";
                        if (textObject_BILLCVSubCheckDate != null) textObject_BILLCVSubCheckDate.Text = bills[0].DueDate.ToString("MMMM dd, yyyy");
                        if (textObject_BILLCVSubTotal != null) textObject_BILLCVSubTotal.Text = bills[0].Amount.ToString("N2");
                        if (textObject_BILLCVSubCheckNumber != null) textObject_BILLCVSubCheckNumber.Text = bills[0].RefNumber ?? "";
                        if (textObject_BILLCVSubCheckNumber != null) textObject_BILLSubAccountPayable.Text = bills[0].BankAccount ?? "";
                        if (textObject_BILLSubAmountPayable != null)
                        {
                            // Sums the AmountDue of all items in the bills list
                            double totalAmountDue = bills.Sum(b => b.AmountDue);
                            textObject_BILLSubAmountPayable.Text = totalAmountDue.ToString("N2");
                        }

                        InsertDataToBillCompiled(refNumberCR, bills);
                    }
                    catch
                    {
                        throw;
                    }
                }

                cRCV_IVPBILL.SetParameterValue("ReferenceNumber", refNumberCR);

                panel_Printing.Visible = false;
                panel_Signatory.Visible = true;
                panel_Main.Visible = false;
                panel_Main_CR.Visible = true;

                reportViewer.ReportSource = cRCV_IVPBILL;
                reportViewer.RefreshReport();

                return true;
            }
            catch (Exception ex)
            {
                MessageBox.Show($"KAYAK ERROR HEHEHE:\n{ex.Message}", "Error", MessageBoxButtons.OK, MessageBoxIcon.Error);
                return false;
            }
        }

        public static void InsertDataToCheckVoucherCompiledIVP(string refNumber, List<CheckTableExpensesAndItems> checkData)
        {
            string connectionString = AccessToDatabase_IVP.GetAccessConnectionString();
            double debitTotalAmount = 0;
            double creditTotalAmount = 0;

            using (OleDbConnection connection = new OleDbConnection(connectionString))
            {
                connection.Open();

                // 1. Clear old data
                string deleteQuery = "DELETE FROM CheckVoucherCompiled";
                using (OleDbCommand deleteCommand = new OleDbCommand(deleteQuery, connection))
                {
                    try
                    {
                        deleteCommand.ExecuteNonQuery();
                        Console.WriteLine("Old data has been deleted from CheckVoucherCompiled.");
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

                foreach (var check in checkData)
                {
                    try
                    {
                        // COMMON FIELDS
                        string memoValue = string.IsNullOrEmpty(check.ExpensesMemo) ? "" : check.ExpensesMemo;
                        string customerJob = string.IsNullOrEmpty(check.ExpensesCustomerJob) ? "" : check.ExpensesCustomerJob;

                        // ---------------------------------------------------------
                        // FIXED SECTION: INSERT ITEM ENTRY
                        // Changed 'check.ItemName' to 'check.Item'
                        // ---------------------------------------------------------
                        if (!string.IsNullOrEmpty(check.Item))
                        {
                            string itemName = check.Item; // FIXED: use check.Item
                            string itemClass = check.ItemClass;
                            double itemAmount = check.ItemAmount;

                            string debit = itemAmount > 0 ? itemAmount.ToString("N2") : "";
                            string credit = itemAmount < 0 ? Math.Abs(itemAmount).ToString("N2") : "";

                            if (itemAmount > 0) debitTotalAmount += itemAmount;
                            else if (itemAmount < 0) creditTotalAmount += Math.Abs(itemAmount);

                            using (OleDbCommand command = new OleDbCommand(insertQuery, connection))
                            {
                                command.Parameters.AddWithValue("@RefNumber", refNumber);
                                command.Parameters.AddWithValue("@Particulars", itemName);
                                command.Parameters.AddWithValue("@Class", string.IsNullOrEmpty(itemClass) ? (object)DBNull.Value : itemClass);
                                command.Parameters.AddWithValue("@Debit", debit);
                                command.Parameters.AddWithValue("@Credit", credit);
                                command.Parameters.AddWithValue("@Memo", memoValue);
                                command.Parameters.AddWithValue("@CustomerJob", customerJob);

                                command.ExecuteNonQuery();
                            }
                        }

                        // ---------------------------------------------------------
                        // FIXED SECTION: INSERT EXPENSE ENTRY
                        // Changed 'check.AccountNameCheck' to 'check.Account'
                        // ---------------------------------------------------------
                        if (!string.IsNullOrEmpty(check.Account))
                        {
                            // Note: check.AccountNumber might also be empty if you didn't assign it in the retrieval function.
                            // If check.AccountNumber is null, this line might look like " - Utilities". 
                            // You might want to just use check.Account if you don't have numbers.
                            string expenseName = check.Account;

                            // If you have account numbers populated, use this format instead:
                            // string expenseName = check.AccountNumber + " - " + check.Account;

                            string expenseClass = check.ExpenseClass; // Ensure this property name matches too (ExpenseClass vs AccountClassCheck)
                            double expenseAmount = check.ExpensesAmount;

                            string debit = expenseAmount > 0 ? expenseAmount.ToString("N2") : "";
                            string credit = expenseAmount < 0 ? Math.Abs(expenseAmount).ToString("N2") : "";

                            if (expenseAmount > 0) debitTotalAmount += expenseAmount;
                            else if (expenseAmount < 0) creditTotalAmount += Math.Abs(expenseAmount);

                            using (OleDbCommand command = new OleDbCommand(insertQuery, connection))
                            {
                                command.Parameters.AddWithValue("@RefNumber", refNumber);
                                command.Parameters.AddWithValue("@Particulars", expenseName);
                                command.Parameters.AddWithValue("@Class", string.IsNullOrEmpty(expenseClass) ? (object)DBNull.Value : expenseClass);
                                command.Parameters.AddWithValue("@Debit", debit);
                                command.Parameters.AddWithValue("@Credit", credit);
                                command.Parameters.AddWithValue("@Memo", memoValue);
                                command.Parameters.AddWithValue("@CustomerJob", customerJob);

                                command.ExecuteNonQuery();
                            }
                        }
                        if (string.IsNullOrEmpty(check.Item) && string.IsNullOrEmpty(check.Account) && !string.IsNullOrEmpty(check.ItemDescription))
                        {
                            using (OleDbCommand command = new OleDbCommand(insertQuery, connection))
                            {
                                command.Parameters.AddWithValue("@RefNumber", refNumber);
                                // We use the Description field here since Item/Account are blank
                                command.Parameters.AddWithValue("@Particulars", check.ItemDescription);
                                command.Parameters.AddWithValue("@Class", (object)DBNull.Value);
                                command.Parameters.AddWithValue("@Debit", "");
                                command.Parameters.AddWithValue("@Credit", "");
                                command.Parameters.AddWithValue("@Memo", memoValue);
                                command.Parameters.AddWithValue("@CustomerJob", customerJob);

                                command.ExecuteNonQuery();
                            }
                        }
                    }
                    catch (Exception ex)
                    {
                        MessageBox.Show($"Error processing check data: {ex.Message}", "Error", MessageBoxButtons.OK, MessageBoxIcon.Error);
                    }
                }

                connection.Close();
            }

            Console.WriteLine($"Total Debit: {debitTotalAmount:F2}, Total Credit: {creditTotalAmount:F2}");
        }

        public static void InsertDataToJournalCompiled(string refNumber, List<JournalGridItem> journalData)
        {
            string connectionString = AccessToDatabase_IVP.GetAccessConnectionString();

            double debitTotalAmount = 0;
            double creditTotalAmount = 0;

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

                // 2. Prepare Insert Query
                // ADDED: [Name] column and @Name placeholder
                string insertQuery = @"
                            INSERT INTO JV_Compiled 
                            (RefNumber, [Particulars], [Class], [Name], [Debit], [Credit], [Memo]) 
                            VALUES 
                            (@RefNumber, @Particulars, @Class, @Name, @Debit, @Credit, @Memo)";

                foreach (var line in journalData)
                {
                    try
                    {
                        // MAPPING VARIABLES
                        string particulars = string.IsNullOrEmpty(line.AccountName) ? "" : line.AccountName;
                        string className = line.Class;
                        // ADDED: Name mapping (Ensuring it handles nulls)
                        string nameValue = string.IsNullOrEmpty(line.Name) ? "" : line.Name;
                        string memoValue = string.IsNullOrEmpty(line.Memo) ? "" : line.Memo;

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
                            // IMPORTANT: The order of these parameters MUST match the order in the SQL string above
                            command.Parameters.AddWithValue("@RefNumber", refNumber);
                            command.Parameters.AddWithValue("@Particulars", particulars);

                            // Handle Class nulls
                            command.Parameters.AddWithValue("@Class", string.IsNullOrEmpty(className) ? (object)DBNull.Value : className);

                            // ADDED: Name Parameter
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
            string connectionString = AccessToDatabase_IVP.GetAccessConnectionString();
            double debitTotalAmount = 0;
            double creditTotalAmount = 0;

            using (OleDbConnection connection = new OleDbConnection(connectionString))
            {
                try
                {
                    connection.Open();

                    // 1. CLEAR OLD DATA
                    string deleteQuery = "DELETE FROM Bill_Compiled";
                    using (OleDbCommand deleteCommand = new OleDbCommand(deleteQuery, connection))
                    {
                        deleteCommand.ExecuteNonQuery();
                    }

                    // 2. PREPARE INSERT QUERY (Note: Order is Ref, Part, Class, Memo, Job, Debit, Credit)
                    string insertQuery = @"INSERT INTO Bill_Compiled 
                           (RefNumber, Particulars, [Class], [Memo], [CustomerJob], Debit, Credit) 
                           VALUES (?, ?, ?, ?, ?, ?, ?)";

                    foreach (var bill in bills)
                    {
                        foreach (var detail in bill.ItemDetails)
                        {
                            string particulars = "";
                            string classVal = "";
                            string memo = "";
                            string customerJob = "";
                            double amount = 0;

                            // Determine if this is an Item Line or an Expense Line to prevent double counting
                            if (!string.IsNullOrEmpty(detail.ItemLineItemRefFullName))
                            {
                                particulars = detail.ItemLineItemRefFullName;
                                classVal = detail.ItemLineClassRefFullName ?? "";
                                memo = detail.ItemLineMemo ?? "";
                                customerJob = detail.ItemLineCustomerJob ?? "";
                                amount = detail.ItemLineAmount;
                            }
                            else if (!string.IsNullOrEmpty(detail.ExpenseLineItemRefFullName))
                            {
                                // Use bill level AccountNumber if available
                                particulars = (!string.IsNullOrEmpty(bill.AccountNumber) ? bill.AccountNumber + " - " : "") + detail.ExpenseLineItemRefFullName;
                                classVal = detail.ExpenseLineClassRefFullName ?? "";
                                memo = detail.ExpenseLineMemo ?? "";
                                customerJob = detail.ExpenseLineCustomerJob ?? "";
                                amount = detail.ExpenseLineAmount;
                            }
                            else
                            {
                                // Skip empty lines
                                continue;
                            }

                            // Calculate Debit/Credit
                            string debitStr = amount > 0 ? amount.ToString("N2") : "";
                            string creditStr = amount < 0 ? Math.Abs(amount).ToString("N2") : "";

                            if (amount > 0) debitTotalAmount += amount;
                            else if (amount < 0) creditTotalAmount += Math.Abs(amount);

                            // 3. EXECUTE INSERT
                            using (OleDbCommand command = new OleDbCommand(insertQuery, connection))
                            {
                                // OleDb requires exact positional order
                                command.Parameters.Add("?", OleDbType.VarWChar).Value = refNumber ?? (object)DBNull.Value;
                                command.Parameters.Add("?", OleDbType.VarWChar).Value = particulars ?? "";
                                command.Parameters.Add("?", string.IsNullOrWhiteSpace(classVal) ? (object)DBNull.Value : classVal);
                                command.Parameters.Add("?", memo ?? (object)DBNull.Value);
                                command.Parameters.Add("?", customerJob ?? (object)DBNull.Value);
                                command.Parameters.Add("?", debitStr);
                                command.Parameters.Add("?", creditStr);

                                command.ExecuteNonQuery();
                            }
                        }
                    }
                    connection.Close();
                }
                catch (Exception ex)
                {
                    MessageBox.Show($"Error: {ex.Message}");
                }
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
                    AccessQueries_IVP queries = new AccessQueries_IVP();

                    cheque = new List<CheckTable>();
                    bills = new List<BillTable>();
                    checks = new List<CheckTableExpensesAndItems>();
                    receipts = new List<ItemReciept>();
                    apvData = new List<BillTable>();
                    checkivp = new List<CheckTableGrid>();

                    object data = null;
                    
                    if (GlobalVariables.client == "IVP")
                    {
                        if (comboBox_Forms.SelectedIndex == 2) // Check
                        {
                            checkivp = queries.GetCheckDataIVP(refNumber);
                            data = checkivp;
                        }
                    }

                    //if (checks.Count > 0 || bills.Count > 0 || receipts.Count > 0)
                    if (data is System.Collections.ICollection colletion && colletion.Count > 0)
                    {
                        if (GlobalVariables.client == "IVP")
                        {
                            Layouts_IVP layouts_IVP = new Layouts_IVP();
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
                                layouts_IVP.PrintPage_IVP(s, ev, selectedIndex, seriesNumber, data, payeeOverride);
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

            if (GlobalVariables.client == "IVP")
            {
                comboBox_Signatory.Items.AddRange(new string[]
                {
                    "Select Signatory Option",
                    "Prepared By:",
                    "Checked By:",
                    "Approved By:",
                    "Released By:",
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

                    if (comboBox_Forms.SelectedIndex == 3) // APV
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

                        else if (GlobalVariables.client == "IVP")
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
                                    UpdateSeriesNumberIVP(formType);
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

            if (GlobalVariables.client == "IVP")
            {
                // Logic for visibility of company panel
                /*if (GlobalVariables.client == "IVP" && comboBox_Forms.SelectedItem?.ToString() == "Check")
                {
                    MessageBox.Show("The 'Check' form is not accessible.", "Access Restricted", MessageBoxButtons.OK, MessageBoxIcon.Warning);

                    // Silently reset to blank (index 0) so they can't stay on this tab
                    comboBox_Forms.SelectedIndex = 0;
                    return;
                }*/


                if (comboBox_Forms.SelectedIndex == 1 || comboBox_Forms.SelectedIndex == 3 || comboBox_Forms.SelectedIndex == 4) // CV or JV
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
                else if (comboBox_Forms.SelectedIndex == 3) prefix = "JV";
                else if (comboBox_Forms.SelectedIndex == 4) prefix = "APV";

                if (prefix != "")
                {
                    string selectedCompany = comboBox_Company.SelectedItem?.ToString();
                    if (!string.IsNullOrEmpty(selectedCompany))
                    {
                        seriesNumber = accessToDatabase.GetSeriesNumberFromDatabase(prefix, selectedCompany);
                        UpdateSeriesNumberIVP(prefix);
                    }
                    else
                    {
                        // If no company selected, show prefix only or leave empty
                        textBox_SeriesNumber.Text = $"{prefix}-00000";
                    }
                }

                switch (comboBox_Forms.SelectedIndex)
                {
                    case 1: // Check Voucher
                        prefix = "CV";
                        panel_SeriesNumber.Visible = true;
                        panel_RefNumber.Visible = false;
                        panel_RefNumberCrystalReport.Visible = true;
                        panel_Signatory.Visible = true;
                        label_SeriesNumberText.Text = "Current Series Number: CV";

                        if (label_CurrencyText != null) label_CurrencyText.Visible = true;
                        if (comboBox_Currency != null) comboBox_Currency.Visible = true;
                        panel_Company.Height = 120;

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
                        panel_SeriesNumber.Visible = true;

                        panel_Main.Visible = false;
                        panel_Main_CR.Visible = true;

                        label_SeriesNumberText.Text = "Current Series Number: JV";

                        if (label_CurrencyText != null) label_CurrencyText.Visible = false;
                        if (comboBox_Currency != null) comboBox_Currency.Visible = false;
                        panel_Company.Height = 61;
                        break;
                    case 4: // Accounts Payable Voucher
                        prefix = "APV";
                        panel_SeriesNumber.Visible = true;
                        panel_RefNumber.Visible = false;
                        panel_RefNumberCrystalReport.Visible = true;
                        panel_Signatory.Visible = true;

                        label_SeriesNumberText.Text = "Current Series Number: APV";

                        // Layout settings
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



            else
            {
                if (comboBox_Forms.SelectedIndex == 1) // CV
                {
                    panel_SeriesNumber.Visible = true;
                    label_SeriesNumberText.Text = "Current Series Number: CV";
                    textBox_SeriesNumber.Text = "CV" + seriesNumber;
                }
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

        private void TextBox_SeriesNumber_TextChanged(object sender, EventArgs e)
        {
            if (GlobalVariables.client == "IVP")
            {
                if (!string.IsNullOrEmpty(textBox_SeriesNumber.Text))
                {
                    string formPrefix = "";
                    if (comboBox_Forms.SelectedIndex == 1) formPrefix = "CV";
                    else if (comboBox_Forms.SelectedIndex == 3) formPrefix = "JV";

                    if (!string.IsNullOrEmpty(formPrefix))
                    {
                        // Clean the input: Just remove "CV" or "JV"
                        // The user sees "CV00001", stripping "CV" leaves "00001"
                        string cleanInput = textBox_SeriesNumber.Text
                            .Replace(formPrefix, "")
                            .Trim();

                        if (int.TryParse(cleanInput, out int adjustedSeries))
                        {
                            seriesNumber = adjustedSeries;
                        }
                    }
                }
            }

        }
        private void TextBox_SeriesNumber_Leave(object sender, EventArgs e)
        {
            if (GlobalVariables.client == "IVP")
            {
                string formType = "";
                if (comboBox_Forms.SelectedIndex == 1) formType = "CV";
                else if (comboBox_Forms.SelectedIndex == 3) formType = "JV";
                else if (comboBox_Forms.SelectedIndex == 4) formType = "APV";

                if (!string.IsNullOrEmpty(formType) && comboBox_Company.SelectedItem != null)
                {
                    // Pass the Company Name so Database knows which column (e.g. NL_CV) to update
                    accessToDatabase.UpdateManualSeriesNumber(formType, seriesNumber, comboBox_Company.SelectedItem.ToString());
                }
            }
        }

        private void UpdateSeriesNumber(string prefix)
        {
            textBox_SeriesNumber.Text = $"{prefix}{seriesNumber:000}"; // Formats seriesNumber as a 3-digit number
        }
        private void RefreshSeriesNumber(string columnName)
        {
            seriesNumber = accessToDatabase.GetSeriesNumberFromDatabase(columnName);
            string prefix = comboBox_Forms.SelectedIndex == 2 ? "CV" : "APV";
            textBox_SeriesNumber.Text = $"{prefix}{seriesNumber:000}";
        }

        // 1. HELPER: Maps the full company name to the short code (e.g., "North Luzon" -> "NL")
        private string GetCompanyCode(string fullCompanyName)
        {
            if (string.IsNullOrEmpty(fullCompanyName)) return "";

            switch (fullCompanyName)
            {
                case "North Luzon": return "NL";
                case "South Luzon": return "SL";
                case "Visayas": return "VIS";
                case "Mindanao": return "MIN";
                case "Metro Manila": return "MM";
                case "Iberica Verheilen Pharmaceuticals Group.": return "IVP";
                case "Verheilen Iberica HealthCare Company Inc.": return "VIHC";
                case "My Health Shield NutriPharm Inc.": return "MHS";
                case "Central Luzon": return "CL";
                //----- GIC COMPANY------
                case "Greenfloor Innovations Corporation": return "GIC";
                default: return "";
            }
        }

        // 2. HELPER: Specific update logic for IVP (Format: CODE-CV00001)

        /* private void UpdateSeriesNumberIVP(string formPrefix)
         {
             // Get the company code
             string selectedCompany = comboBox_Company.SelectedItem?.ToString();
             string companyCode = GetCompanyCode(selectedCompany);

             // Format with 5 digits (00000)
             if (!string.IsNullOrEmpty(companyCode))
             {
                 textBox_SeriesNumber.Text = $"{formPrefix}{seriesNumber:00000}";
             }
             else
             {
                 // Fallback if no company is selected
                 textBox_SeriesNumber.Text = $"{formPrefix}{seriesNumber:00000}";
             }
         }*/ //FOR MANUAL SERIES NUMBER ENTRY

        private void UpdateSeriesNumberIVP(string formPrefix)
        {
            // Ensure accessToDatabase is initialized
            if (accessToDatabase == null) accessToDatabase = new AccessToDatabase_IVP();

            // Format with a hyphen and 5 digits (e.g., CV-00001)
            textBox_SeriesNumber.Text = $"{formPrefix}-{seriesNumber:00000}";
        }
    }
}
