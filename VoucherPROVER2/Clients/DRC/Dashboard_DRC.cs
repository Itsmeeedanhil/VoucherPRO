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
using static VoucherPROVER2.Clients.DRC.Dataclass_DRC;
using System.IO;
using System.Data.OleDb;
using VoucherPROVER2.Clients.ENA;


namespace VoucherPROVER2.Clients.DRC
{
    public partial class Dashboard_DRC : Form
    {
        public Dashboard_DRC()
        {
            InitializeComponent();

            accessToDatabase = new AccessToDatabase_DRC();
        }

        private PrintDocument printDocument;
        private PrintPreviewControl printPreviewControl;
        private CrystalReportViewer reportViewer;
        private AccessToDatabase_DRC accessToDatabase;


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
                Visible = (GlobalVariables.client == "DRC")
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
                "DASMARINAS RENAL CARE CENTER INC.",

            });

            // Set default selection
            if (comboBox_Company.Items.Count > 0)
            {
                comboBox_Company.SelectedIndex = 0;
            }

            comboBox_Company.SelectedIndexChanged += (sender, e) =>
            {
                string formType = "";
                if (comboBox_Forms.SelectedIndex == 1) formType = "CV";
                else if (comboBox_Forms.SelectedIndex == 3) formType = "JV";
                else if (comboBox_Forms.SelectedIndex == 4) formType = "APV";
                else if (comboBox_Forms.SelectedIndex == 5) formType = "IR";

                if (formType != "")
                {
                    string selectedCompany = comboBox_Company.SelectedItem.ToString();
                    seriesNumber = accessToDatabase.GetSeriesNumberFromDatabase(formType, selectedCompany);
                    UpdateSeriesNumberDRC(formType);
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
                        if (string.IsNullOrEmpty(item.ToolTipText) || !item.ToolTipText.Contains("Print"))
                        {
                            continue;
                        }

                        item.Click += (s, e) =>
                        {
                            if (GlobalVariables.client == "DRC")
                            {
                                string formType = "";
                                if (comboBox_Forms.SelectedIndex == 1) formType = "CV";
                                else if (comboBox_Forms.SelectedIndex == 3) formType = "JV";
                                else if (comboBox_Forms.SelectedIndex == 4) formType = "APV";
                                else if (comboBox_Forms.SelectedIndex == 5) formType = "IR";

                                string selectedCompany = comboBox_Company.SelectedItem?.ToString();

                                if (!string.IsNullOrEmpty(formType) && !string.IsNullOrEmpty(selectedCompany))
                                {
                                    seriesNumber++;
                                    accessToDatabase.UpdateManualSeriesNumber(formType, seriesNumber, selectedCompany);

                                    if (this.IsHandleCreated && !this.IsDisposed)
                                    {
                                        this.BeginInvoke((MethodInvoker)delegate
                                        {
                                            UpdateSeriesNumberDRC(formType);
                                        });
                                    }
                                    else
                                    {
                                        UpdateSeriesNumberDRC(formType);
                                    }
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

            if (GlobalVariables.client == "DRC")
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
            if (GlobalVariables.client == "DRC")
            {
                comboBox_Forms.Items.AddRange(new string[]
            {
                "",
                "Check Voucher",
                "Check",
                "Journal Voucher",
                "Accounts Payable Voucher",
                "Receiving Report",

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
                if (GlobalVariables.client == "DRC")
                {
                    seriesNumber--;
                    string prefix = "";
                    if (comboBox_Forms.SelectedIndex == 1) prefix = "CV";
                    else if (comboBox_Forms.SelectedIndex == 3) prefix = "JV";
                    else if (comboBox_Forms.SelectedIndex == 4) prefix = "APV";
                    else if (comboBox_Forms.SelectedIndex == 5) prefix = "IR";

                    UpdateSeriesNumberDRC(prefix);
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
                if (GlobalVariables.client == "DRC")
                {
                    seriesNumber++;
                    string prefix = "";
                    if (comboBox_Forms.SelectedIndex == 1) prefix = "CV";
                    else if (comboBox_Forms.SelectedIndex == 3) prefix = "JV";
                    else if (comboBox_Forms.SelectedIndex == 4) prefix = "APV";
                    else if (comboBox_Forms.SelectedIndex == 5) prefix = "IR";

                    UpdateSeriesNumberDRC(prefix);
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
                    if (GlobalVariables.client == "DRC")
                    {
                        // -------------------------------------------------------------
                        // OPTION 1: CHECK VOUCHER
                        // -------------------------------------------------------------
                        if (comboBox_Forms.SelectedIndex == 1)
                        {
                            bool cvDataExists = false;
                            try
                            {
                                CRCV_DRC cRCV_DRC = new CRCV_DRC();
                                string databasePath = Path.Combine(Application.StartupPath, "CheckDatabase.accdb");
                                SetDatabaseLocation(cRCV_DRC, databasePath);

                                AccessQueries_DRC accessQueries = new AccessQueries_DRC();
                                string refNumberCR = textBox_ReferenceNumber_CR.Text;

                                cvData = accessQueries.GetCheckExpensesAndItemsData_DRC(refNumberCR);

                                if (cvData != null && cvData.Count > 0)
                                {
                                    cvDataExists = true;

                                    TextObject textObject_CVRefNumber = cRCV_DRC.ReportDefinition.ReportObjects["TextCVRefNumber"] as TextObject;
                                    TextObject textObject_CVDateTime = cRCV_DRC.ReportDefinition.ReportObjects["TextCVDateTime"] as TextObject;
                                    TextObject textObject_CVPayee = cRCV_DRC.ReportDefinition.ReportObjects["TextCVPayee"] as TextObject;
                                    TextObject textObject_CVAddress = cRCV_DRC.ReportDefinition.ReportObjects["TextCVAddress"] as TextObject;
                                    TextObject textObject_CVTotalDebitAmount = cRCV_DRC.ReportDefinition.ReportObjects["TextCVTotalDebitAmount"] as TextObject;
                                    TextObject textObject_CVTotalCreditAmount = cRCV_DRC.ReportDefinition.ReportObjects["TextCVTotalCreditAmount"] as TextObject;

                                    TextObject textObject_CompanyName = cRCV_DRC.ReportDefinition.ReportObjects["TextCompanyName"] as TextObject;
                                    if (textObject_CompanyName != null && comboBox_Company != null && comboBox_Company.SelectedItem != null)
                                    {
                                        textObject_CompanyName.Text = comboBox_Company.SelectedItem.ToString();
                                    }

                                    TextObject textObject_PreparedBy = cRCV_DRC.ReportDefinition.ReportObjects["TextPreparedBy"] as TextObject;
                                    TextObject textObject_PreparedByPos = cRCV_DRC.ReportDefinition.ReportObjects["TextPreparedByPosition"] as TextObject;
                                    TextObject textObject_CheckedBy = cRCV_DRC.ReportDefinition.ReportObjects["TextCheckedBy"] as TextObject;
                                    TextObject textObject_CheckedByPos = cRCV_DRC.ReportDefinition.ReportObjects["TextCheckedByPosition"] as TextObject;
                                    TextObject textObject_ApprovedBy = cRCV_DRC.ReportDefinition.ReportObjects["TextApprovedBy"] as TextObject;
                                    TextObject textObject_ApprovedByPos = cRCV_DRC.ReportDefinition.ReportObjects["TextApprovedByPosition"] as TextObject;
                                    TextObject textObject_ReceivedBy = cRCV_DRC.ReportDefinition.ReportObjects["TextReceivedBy"] as TextObject;
                                    TextObject textObject_ReceivedByPos = cRCV_DRC.ReportDefinition.ReportObjects["TextReceivedByPosition"] as TextObject;


                                    TextObject textObject_CVCheckNumber = cRCV_DRC.ReportDefinition.ReportObjects["TextCVCheckNum"] as TextObject;
                                    TextObject textObject_CVCheckBank = cRCV_DRC.ReportDefinition.ReportObjects["TextCVCheckBank"] as TextObject;
                                    TextObject textObject_CVCheckDate = cRCV_DRC.ReportDefinition.ReportObjects["TextCVCheckDate"] as TextObject;
                                    TextObject textObject_CVDuePayment = cRCV_DRC.ReportDefinition.ReportObjects["TextCVDuePayment"] as TextObject;

                                    AccessToDatabase_DRC accessToDatabase = new AccessToDatabase_DRC();
                                    var signatories = accessToDatabase.RetrieveAllSignatoryData();

                                    string rawBank = cvData[0].BankAccount ?? "";

                                    string bank = rawBank.Contains(":")
                                        ? rawBank.Split(':').Last().Trim()
                                        : rawBank;

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


                                    double amount = cvData[0].TotalAmount;
                                    string amountInWords = AccessToDatabase_DRC.AmountToWordsConverter.Convert(amount);

                                    textObject_CVRefNumber.Text = textBox_SeriesNumber.Text;
                                    textObject_CVDateTime.Text = DateTime.Now.ToString("MMMM dd, yyyy");
                                    textObject_CVPayee.Text = cvData[0].PayeeFullName;
                                    textObject_CVAddress.Text = fullAddress;

                                    textObject_CVCheckNumber.Text = cvData[0].RefNumber;
                                    textObject_CVCheckBank.Text = bank;
                                    textObject_CVCheckDate.Text = cvData[0].DueDate.ToString("MMMM dd, yyyy");
                                    textObject_CVDuePayment.Text = amount.ToString();
                                    
                                    


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

                                    textObject_CVTotalDebitAmount.Text = $"PHP {debitTotalAmount:N2}";
                                    textObject_CVTotalCreditAmount.Text = $"PHP {debitTotalAmount:N2}";


                                    SubreportObject subreportObject = cRCV_DRC.ReportDefinition.ReportObjects["SubreportCVDetailsIVP"] as SubreportObject;
                                    if (subreportObject != null)
                                    {
                                        ReportDocument subReportDocument = cRCV_DRC.OpenSubreport(subreportObject.SubreportName);
                                        TextObject textObject_Remarks = subReportDocument.ReportDefinition.ReportObjects["TextRemarks"] as TextObject;
                                        TextObject textObject_SubAccountPayable = subReportDocument.ReportDefinition.ReportObjects["TextSubAccountPayable"] as TextObject;
                                        TextObject textObject_SubAmountPayable = subReportDocument.ReportDefinition.ReportObjects["TextSubAmountPayable"] as TextObject;
                                        TextObject textObject_SubAccountCode = subReportDocument.ReportDefinition.ReportObjects["TextSubAccountCode"] as TextObject;


                                        string subbank = cvData[0].BankAccount ?? "";
                                        string accountcode = cvData[0].AccountNumber ?? "";

                                        string subfinalbank = subbank.Contains(":")
                                            ? rawBank.Split(':').Last().Trim()
                                            : rawBank;

                                        textObject_Remarks.Text = cvData[0].Memo;
                                        textObject_SubAccountPayable.Text = subfinalbank;
                                        textObject_SubAmountPayable.Text = debitTotalAmount.ToString("N2");
                                        textObject_SubAccountCode.Text = accountcode;

                                        Console.WriteLine($"Subreport Bank: {subfinalbank}, Account Code: {accountcode}");


                                        InsertDataToCheckVoucherCompiledDRC(refNumberCR, cvData);
                                    }

                                    cRCV_DRC.SetParameterValue("ReferenceNumber", refNumberCR);

                                    panel_Printing.Visible = false;
                                    panel_Signatory.Visible = true;
                                    panel_Main.Visible = false;
                                    panel_Main_CR.Visible = true;

                                    reportViewer.ReportSource = cRCV_DRC;
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
                                GenerateBillPaymentReport_DRC(refNumberCR);
                            }
                        }

                        else if (comboBox_Forms.SelectedIndex == 3)
                        {
                            CRJV_DRC cRJV_DRC = new CRJV_DRC();
                            string databasePath = Path.Combine(Application.StartupPath, "CheckDatabase.accdb");
                            SetDatabaseLocation(cRJV_DRC, databasePath);

                            AccessQueries_DRC accessQueries = new AccessQueries_DRC();
                            string refNumberCR = textBox_ReferenceNumber_CR.Text;

                            // 1. Get the correct data
                            journal = accessQueries.GetJournalEntryForGrid(refNumberCR);

                            if (journal != null && journal.Count > 0)
                            {
                                // 2. Set Header Text Objects
                                TextObject textObject_JVRefNumber = cRJV_DRC.ReportDefinition.ReportObjects["TextJVRefNumber"] as TextObject;
                                TextObject textObject_JVCheckDate = cRJV_DRC.ReportDefinition.ReportObjects["TextJVCheckDate"] as TextObject;
                                TextObject textObject_JVTransactDate = cRJV_DRC.ReportDefinition.ReportObjects["TextJVTransactDate"] as TextObject;
                                TextObject textObject_JVTotalDebitAmount = cRJV_DRC.ReportDefinition.ReportObjects["TextJVTotalDebitAmount"] as TextObject;
                                TextObject textObject_JVTotalCreditAmount = cRJV_DRC.ReportDefinition.ReportObjects["TextJVTotalCreditAmount"] as TextObject;

                                TextObject textObject_CompanyName = cRJV_DRC.ReportDefinition.ReportObjects["TextCompanyName"] as TextObject;
                                if (textObject_CompanyName != null && comboBox_Company != null && comboBox_Company.SelectedItem != null)
                                {
                                    textObject_CompanyName.Text = comboBox_Company.SelectedItem.ToString();
                                }


                                TextObject textObject_PreparedBy = cRJV_DRC.ReportDefinition.ReportObjects["TextPreparedBy"] as TextObject;
                                TextObject textObject_PreparedByPos = cRJV_DRC.ReportDefinition.ReportObjects["TextPreparedByPosition"] as TextObject;
                                TextObject textObject_CheckedBy = cRJV_DRC.ReportDefinition.ReportObjects["TextCheckedBy"] as TextObject;
                                TextObject textObject_CheckedByPos = cRJV_DRC.ReportDefinition.ReportObjects["TextCheckedByPosition"] as TextObject;
                                TextObject textObject_ApprovedBy = cRJV_DRC.ReportDefinition.ReportObjects["TextApprovedBy"] as TextObject;
                                TextObject textObject_ApprovedByPos = cRJV_DRC.ReportDefinition.ReportObjects["TextApprovedByPosition"] as TextObject;

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
                                if (textObject_JVTotalDebitAmount != null)
                                    textObject_JVTotalDebitAmount.Text = $"PHP {debitTotalAmount:N2}";

                                if (textObject_JVTotalCreditAmount != null)
                                    textObject_JVTotalCreditAmount.Text = $"PHP {creditTotalAmount:N2}";


                                AccessToDatabase_DRC accessToDatabase = new AccessToDatabase_DRC();
                                var signatories = accessToDatabase.RetrieveAllSignatoryData();


                                textObject_PreparedBy.Text = signatories.PreparedByName;
                                textObject_PreparedByPos.Text = signatories.PreparedByPosition;
                                textObject_CheckedBy.Text = signatories.ReviewedByName;
                                textObject_CheckedByPos.Text = signatories.ReviewedByPosition;
                                textObject_ApprovedBy.Text = signatories.ApprovedByName;
                                textObject_ApprovedByPos.Text = signatories.ApprovedByPosition;

                                // 4. Handle Subreport
                                SubreportObject subreportObject = cRJV_DRC.ReportDefinition.ReportObjects["SubreportJVDetailsIVP"] as SubreportObject;
                                if (subreportObject != null)
                                {
                                    ReportDocument subReportDocument = cRJV_DRC.OpenSubreport(subreportObject.SubreportName);

                                    TextObject textObject_SubAccountPayable = subReportDocument.ReportDefinition.ReportObjects["TextJVSUBAccountsPayable"] as TextObject;
                                    TextObject textObject_SubAmountPayable = subReportDocument.ReportDefinition.ReportObjects["TextJVSUBAmountPayable"] as TextObject;


                                    if (textObject_SubAccountPayable != null) textObject_SubAccountPayable.Text = journal[0].AccountName;

                                    if (textObject_SubAmountPayable != null)
                                        textObject_SubAmountPayable.Text = debitTotalAmount.ToString("N2");
                                }

                                InsertDataToJournalCompiled(refNumberCR, journal);

                                // 6. Final Report Settings
                                cRJV_DRC.SetParameterValue("ReferenceNumber", refNumberCR);

                                panel_Printing.Visible = false;
                                panel_Signatory.Visible = true;
                                panel_Main.Visible = false;
                                panel_Main_CR.Visible = true;

                                reportViewer.ReportSource = cRJV_DRC;
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
                            GenerateAPVReport_DRC(refNumberCR);
                        }
                        // 5. ITEM RECEIPT (IR) - NEW MODULE ENTRY
                        else if (comboBox_Forms.SelectedIndex == 5)
                        {
                            string refNumberCR = textBox_ReferenceNumber_CR.Text;
                            GenerateItemReceiptReport_DRC(refNumberCR);
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

        private bool GenerateAPVReport_DRC(string refNumberCR)
        {
            try
            {
                CRAPV_DRCBILL cRAPV_DRCBILL = new CRAPV_DRCBILL();
                string databasePathBILL = Path.Combine(Application.StartupPath, "CheckDatabase.accdb");
                SetDatabaseLocation(cRAPV_DRCBILL, databasePathBILL);

                AccessQueries_DRC accessQueries = new AccessQueries_DRC();
                List<BillTable> bills = accessQueries.GetBillData_DRC_DirectBill(refNumberCR);

                if (bills == null || bills.Count == 0)
                    return false;

                TextObject textObject_CVBILLCheckNumber = null;
                TextObject textObject_CVBILLCheckDate = null;
                TextObject textObject_CVBILLPayee = null;
                TextObject textObject_CVBILLTerms = null;
                TextObject textObject_CVBILLAddress = null;
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
                TextObject textObject_CVBILLBank = null;
                TextObject textObject_CVBILLNumber = null;
                TextObject textObject_CVBILLDate = null;
                TextObject textObject_CVBILLDue = null;

                try
                {
                    textObject_CVBILLCheckNumber = cRAPV_DRCBILL.ReportDefinition.ReportObjects["TextCVBILLSeriesnumber"] as TextObject;
                    textObject_CVBILLCheckDate = cRAPV_DRCBILL.ReportDefinition.ReportObjects["TextCVBILLCheckDate"] as TextObject;
                    textObject_CVBILLPayee = cRAPV_DRCBILL.ReportDefinition.ReportObjects["TextCVBILLPayee"] as TextObject;
                    textObject_CVBILLAddress = cRAPV_DRCBILL.ReportDefinition.ReportObjects["TextCVBILLAddress"] as TextObject;
                    textObject_CVBILLTerms = cRAPV_DRCBILL.ReportDefinition.ReportObjects["TextCVBILLTerms"] as TextObject;
                    textObject_CVBILLTotalDebitAmount = cRAPV_DRCBILL.ReportDefinition.ReportObjects["TextCVBILLTotalDebitAmount"] as TextObject;
                    textObject_CVBILLTotalCreditAmount = cRAPV_DRCBILL.ReportDefinition.ReportObjects["TextCVBILLTotalCreditAmount"] as TextObject;

                    TextObject textObject_CompanyName = cRAPV_DRCBILL.ReportDefinition.ReportObjects["TextCompanyName"] as TextObject;
                    if (textObject_CompanyName != null && comboBox_Company != null && comboBox_Company.SelectedItem != null)
                    {
                        textObject_CompanyName.Text = comboBox_Company.SelectedItem.ToString();
                    }


                    textObject_CVBILLBank = cRAPV_DRCBILL.ReportDefinition.ReportObjects["TextCVBILLBank"] as TextObject;
                    textObject_CVBILLNumber = cRAPV_DRCBILL.ReportDefinition.ReportObjects["TextCVBILLNumber"] as TextObject;
                    textObject_CVBILLDate = cRAPV_DRCBILL.ReportDefinition.ReportObjects["TextCVBILLDate"] as TextObject;
                    textObject_CVBILLDue = cRAPV_DRCBILL.ReportDefinition.ReportObjects["TextCVBILLDue"] as TextObject;



                    textObject_PreparedBy = cRAPV_DRCBILL.ReportDefinition.ReportObjects["TextPreparedBy"] as TextObject;
                    textObject_PreparedByPos = cRAPV_DRCBILL.ReportDefinition.ReportObjects["TextPreparedByPosition"] as TextObject;
                    textObject_CheckedBy = cRAPV_DRCBILL.ReportDefinition.ReportObjects["TextCheckedBy"] as TextObject;
                    textObject_CheckedByPos = cRAPV_DRCBILL.ReportDefinition.ReportObjects["TextCheckedByPosition"] as TextObject;
                    textObject_ApprovedBy = cRAPV_DRCBILL.ReportDefinition.ReportObjects["TextApprovedBy"] as TextObject;
                    textObject_ApprovedByPos = cRAPV_DRCBILL.ReportDefinition.ReportObjects["TextApprovedByPosition"] as TextObject;
                    textObject_ReceivedBy = cRAPV_DRCBILL.ReportDefinition.ReportObjects["TextReceivedBy"] as TextObject;
                    textObject_ReceivedByPos = cRAPV_DRCBILL.ReportDefinition.ReportObjects["TextReceivedByPosition"] as TextObject;

                    AccessToDatabase_DRC accessToDatabase = new AccessToDatabase_DRC();

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

                    if (textObject_CVBILLTotalDebitAmount != null)
                        textObject_CVBILLTotalDebitAmount.Text = $"PHP {debitTotalAmount:N2}";

                    if (textObject_CVBILLTotalCreditAmount != null)
                        textObject_CVBILLTotalCreditAmount.Text = $"PHP {debitTotalAmount:N2}";

                }
                catch
                {
                    throw;
                }


                double amount = bills[0].AmountDue;
                string amountInWords = AccessToDatabase_DRC.AmountToWordsConverter.Convert(amount);

                string rawBank = bills[0].BankAccount ?? "";

                string bank = rawBank.Contains(":")
                    ? rawBank.Split(':').Last().Trim()
                    : rawBank;

                var c = bills[0];

                // Line 1: Combine Addr1, Addr2, Addr3, Addr4 into one string separated by commas
                string streetLine = string.Join(", ", new[] {
                                                 c.VendorAddressAddr1,
                                                 c.VendorAddressAddr2,
                                                 c.VendorAddressAddr3,
                                                 c.VendorAddressAddr4
                                             }.Where(s => !string.IsNullOrWhiteSpace(s)));

                // Line 2: City (Add State/Zip here if you have them in your BillTable)
                string cityLine = string.Join(" ", new[] {
                                                 c.VendorAddressCity,
                                             }.Where(s => !string.IsNullOrWhiteSpace(s)));

                // Final: Join the two lines with a single NewLine
                string fullAddress = string.Join(Environment.NewLine, new[] { streetLine, cityLine }.Where(s => !string.IsNullOrWhiteSpace(s)));

                if (textObject_CVBILLCheckNumber != null) textObject_CVBILLCheckNumber.Text = textBox_SeriesNumber.Text;
                if (textObject_CVBILLAddress != null) textObject_CVBILLAddress.Text = fullAddress;
                if (textObject_CVBILLCheckDate != null) textObject_CVBILLCheckDate.Text = DateTime.Now.ToString("MMMM dd, yyyy");
                if (textObject_CVBILLPayee != null) textObject_CVBILLPayee.Text = bills[0].PayeeFullName ?? "";
                if (textObject_CVBILLTerms != null) textObject_CVBILLTerms.Text = bills[0].TermsRefFullName ?? "";


                if (textObject_CVBILLBank != null) textObject_CVBILLBank.Text = bank;
                if (textObject_CVBILLNumber != null) textObject_CVBILLNumber.Text = bills[0].RefNumber ?? "";
                if (textObject_CVBILLDate != null) textObject_CVBILLDate.Text = bills[0].DueDate.ToString("MMMM dd, yyyy") ?? "";
                if (textObject_CVBILLDue != null)
                    textObject_CVBILLDue.Text = amount.ToString("N2");

                SubreportObject subreportObject = null;
                try
                {
                    subreportObject = cRAPV_DRCBILL.ReportDefinition.ReportObjects["SubreportCVBILLDetailsIVP"] as SubreportObject;
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
                        subReportDocument = cRAPV_DRCBILL.OpenSubreport(subreportObject.SubreportName);
                    }
                    catch
                    {
                        throw;
                    }

                    try
                    {
                        TextObject textObject_BILLSubRemarks = subReportDocument.ReportDefinition.ReportObjects["TextBILLRemarks"] as TextObject;
                        TextObject textObject_BILLSubAccountPayable = subReportDocument.ReportDefinition.ReportObjects["TextBILLSubAccountPayable"] as TextObject;
                        TextObject textObject_BILLSubAmountPayable = subReportDocument.ReportDefinition.ReportObjects["TextBILLSubAmountPayable"] as TextObject;
                        TextObject textObject_BILLSubAccountCode = subReportDocument.ReportDefinition.ReportObjects["TextBILLSubAccountCode"] as TextObject;


                        if (textObject_BILLSubRemarks != null) textObject_BILLSubRemarks.Text = bills[0].Memo ?? "";
                        if (textObject_BILLSubAccountPayable != null) textObject_BILLSubAccountPayable.Text = bills[0].APAccountRefFullName ?? "";
                        if (textObject_BILLSubAccountCode != null) textObject_BILLSubAccountCode.Text = bills[0].AccountNumber ?? "";
                        if (textObject_BILLSubAmountPayable != null)
                        {
                            // Sums the AmountDue of all items in the bills list
                            double totalAmountDue = bills.Sum(b => b.AmountDue);
                            textObject_BILLSubAmountPayable.Text = totalAmountDue.ToString("N2");
                        }

                        InsertDataToBillAPVCompiled(refNumberCR, bills);
                    }
                    catch
                    {
                        throw;
                    }
                }

                cRAPV_DRCBILL.SetParameterValue("ReferenceNumber", refNumberCR);

                panel_Printing.Visible = false;
                panel_Signatory.Visible = true;
                panel_Main.Visible = false;
                panel_Main_CR.Visible = true;

                reportViewer.ReportSource = cRAPV_DRCBILL;
                reportViewer.RefreshReport();

                return true;
            }
            catch (Exception ex)
            {
                MessageBox.Show($"KAYAK ERROR HEHEHE:\n{ex.Message}", "Error", MessageBoxButtons.OK, MessageBoxIcon.Error);
                return false;
            }
        }

        private bool GenerateItemReceiptReport_DRC(string refNumberCR)
        {
            try
            {
                CRIR_DRC cRIR_DRC = new CRIR_DRC();
                string databasePathIR = Path.Combine(Application.StartupPath, "CheckDatabase.accdb");
                SetDatabaseLocation(cRIR_DRC, databasePathIR);

                AccessQueries_DRC accessQueries = new AccessQueries_DRC();
                receipts = accessQueries.GetItemReceiptData_DRC(refNumberCR);

                if (receipts == null || receipts.Count == 0)
                    return false;

                TextObject textObject_IRSeriesNumber = null;
                TextObject textObject_IRDate = null;
                TextObject textObject_IRVendor = null;
                TextObject textObject_IRAddress = null;
                TextObject textObject_IRTotalDebitAmount = null;
                TextObject textObject_IRTotalCreditAmount = null;
                TextObject textObject_IRBank = null;
                TextObject textObject_IRRefnumber = null;
                TextObject textObject_IRCheckDate = null;
                TextObject textObject_IRDueAmount = null;

                TextObject textObject_CompanyName = null;
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
                    textObject_IRSeriesNumber = cRIR_DRC.ReportDefinition.ReportObjects["TextIRSeriesNumber"] as TextObject;
                    textObject_IRDate = cRIR_DRC.ReportDefinition.ReportObjects["TextIRDate"] as TextObject;
                    textObject_IRVendor = cRIR_DRC.ReportDefinition.ReportObjects["TextIRVendor"] as TextObject;
                    textObject_IRAddress = cRIR_DRC.ReportDefinition.ReportObjects["TextIRAddress"] as TextObject;
                    textObject_IRTotalDebitAmount = cRIR_DRC.ReportDefinition.ReportObjects["TextIRTotalDebitAmount"] as TextObject;
                    textObject_IRTotalCreditAmount = cRIR_DRC.ReportDefinition.ReportObjects["TextIRTotalCreditAmount"] as TextObject;
                    textObject_IRBank = cRIR_DRC.ReportDefinition.ReportObjects["TextIRBank"] as TextObject;
                    textObject_IRRefnumber = cRIR_DRC.ReportDefinition.ReportObjects["TextIRRefnumber"] as TextObject;
                    textObject_IRCheckDate = cRIR_DRC.ReportDefinition.ReportObjects["TextIRCheckDate"] as TextObject;
                    textObject_IRDueAmount = cRIR_DRC.ReportDefinition.ReportObjects["TextIRDueAmoount"] as TextObject; 
                    textObject_CompanyName = cRIR_DRC.ReportDefinition.ReportObjects["TextCompanyName"] as TextObject;
                    if (textObject_CompanyName != null && comboBox_Company != null && comboBox_Company.SelectedItem != null)
                    {
                        textObject_CompanyName.Text = comboBox_Company.SelectedItem.ToString();
                    }

                    textObject_PreparedBy = cRIR_DRC.ReportDefinition.ReportObjects["TextPreparedBy"] as TextObject;
                    textObject_PreparedByPos = cRIR_DRC.ReportDefinition.ReportObjects["TextPreparedByPosition"] as TextObject;
                    textObject_CheckedBy = cRIR_DRC.ReportDefinition.ReportObjects["TextCheckedBy"] as TextObject;
                    textObject_CheckedByPos = cRIR_DRC.ReportDefinition.ReportObjects["TextCheckedByPosition"] as TextObject;
                    textObject_ApprovedBy = cRIR_DRC.ReportDefinition.ReportObjects["TextApprovedBy"] as TextObject;
                    textObject_ApprovedByPos = cRIR_DRC.ReportDefinition.ReportObjects["TextApprovedByPosition"] as TextObject;
                    textObject_ReceivedBy = cRIR_DRC.ReportDefinition.ReportObjects["TextReceivedBy"] as TextObject;
                    textObject_ReceivedByPos = cRIR_DRC.ReportDefinition.ReportObjects["TextReceivedByPosition"] as TextObject;

                    AccessToDatabase_DRC accessToDatabase = new AccessToDatabase_DRC();

                    var (PreparedByName, PreparedByPosition,
                         ReviewedByName, ReviewedByPosition,
                         RecommendingApprovalName, RecommendingApprovalPosition,
                         ApprovedByName, ApprovedByPosition,
                         ReceivedByName, ReceivedByPosition) = accessToDatabase.RetrieveAllSignatoryData();


                    double debitTotalAmount = 0;
                    double creditTotalAmount = 0;

                    foreach (var receipt in receipts) // 'receipts' is List<ItemReciept>
                    {
                        try
                        {
                            double lineAmount = 0;

                            // Extract amount based on line type
                            if (receipt.ReceiptItemType == ReceiptItemType.ReceiptItem)
                            {
                                lineAmount = receipt.ItemAmount;
                            }
                            else if (receipt.ReceiptItemType == ReceiptItemType.RecieptExpense)
                            {
                                lineAmount = receipt.ExpensesAmount;
                            }

                            // Accumulate Debit and Credit
                            if (lineAmount != 0)
                            {
                                if (lineAmount > 0)
                                {
                                    debitTotalAmount += lineAmount;
                                }
                                else
                                {
                                    creditTotalAmount += Math.Abs(lineAmount);
                                }
                            }
                        }
                        catch (Exception ex)
                        {
                            MessageBox.Show($"Error processing item receipt detail: {ex.Message}", "Error", MessageBoxButtons.OK, MessageBoxIcon.Error);
                        }
                    }


                    textObject_PreparedBy.Text = PreparedByName;
                    textObject_PreparedByPos.Text = PreparedByPosition;
                    textObject_CheckedBy.Text = ReviewedByName;
                    textObject_CheckedByPos.Text = ReviewedByPosition;
                    textObject_ApprovedBy.Text = ApprovedByName;
                    textObject_ApprovedByPos.Text = ApprovedByPosition;
                    textObject_ReceivedBy.Text = ReceivedByName;
                    textObject_ReceivedByPos.Text = ReceivedByPosition;


                    if (textObject_IRTotalDebitAmount != null)
                        textObject_IRTotalDebitAmount.Text = $"PHP {debitTotalAmount:N2}";

                    if (textObject_IRTotalCreditAmount != null)
                        textObject_IRTotalCreditAmount.Text = $"PHP {debitTotalAmount:N2}";

                    var firstReceipt = receipts.FirstOrDefault();
                    if (firstReceipt != null)
                    {
                        // Format Multi-line Address
                        string streetLine = string.Join(", ", new[] {
                            firstReceipt.Addr1,
                            firstReceipt.Addr2,
                            firstReceipt.Addr3,
                            firstReceipt.Addr4
                        }.Where(s => !string.IsNullOrWhiteSpace(s)));

                        string cityLine = firstReceipt.AddrCity ?? "";
                        string fullAddress = string.Join(Environment.NewLine, new[] { streetLine, cityLine }.Where(s => !string.IsNullOrWhiteSpace(s)));

                        // Clean Bank / AP Account Name (Strip parent hierarchy if colon exists)
                        string rawBank = firstReceipt.BankAccount ?? "";
                        string bank = rawBank.Contains(":") ? rawBank.Split(':').Last().Trim() : rawBank;

                        // Assign Text Objects
                        if (textObject_IRSeriesNumber != null) textObject_IRSeriesNumber.Text = textBox_SeriesNumber.Text;
                        if (textObject_IRDate != null) textObject_IRDate.Text = firstReceipt.TxnDate.ToString("MMMM dd, yyyy");
                        if (textObject_IRCheckDate != null) textObject_IRCheckDate.Text = DateTime.Now.ToString("MMMM dd, yyyy");
                        if (textObject_IRVendor != null) textObject_IRVendor.Text = firstReceipt.VendorFullName ?? "";
                        if (textObject_IRAddress != null) textObject_IRAddress.Text = fullAddress;
                        if (textObject_IRBank != null) textObject_IRBank.Text = bank;
                        if (textObject_IRRefnumber != null) textObject_IRRefnumber.Text = firstReceipt.RefNumber ?? "";
                        if (textObject_IRDueAmount != null) textObject_IRDueAmount.Text = firstReceipt.TotalAmount.ToString("N2");
                    }
                }
                catch
                {
                    throw;
                }

                // Handle Subreport
                SubreportObject subreportObject = null;
                try
                {
                    subreportObject = cRIR_DRC.ReportDefinition.ReportObjects["SubreportIRDetails"] as SubreportObject;
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
                        subReportDocument = cRIR_DRC.OpenSubreport(subreportObject.SubreportName);
                    }
                    catch
                    {
                        throw;
                    }

                    try
                    {
                        TextObject textObject_IRSubRemarks = subReportDocument.ReportDefinition.ReportObjects["TextIRRemarks"] as TextObject;
                        TextObject textObject_IRSubAccountPayable = subReportDocument.ReportDefinition.ReportObjects["TextIRSubAccountPayable"] as TextObject;
                        TextObject textObject_IRSubAmountPayable = subReportDocument.ReportDefinition.ReportObjects["TextIRSubAmountPayable"] as TextObject;
                        TextObject textObject_IRSubAccountCode = subReportDocument.ReportDefinition.ReportObjects["TextIRSubAccountCode"] as TextObject;

                        if (receipts.Count > 0)
                        {
                            if (textObject_IRSubRemarks != null) textObject_IRSubRemarks.Text = receipts[0].Memo ?? "";
                            if (textObject_IRSubAccountPayable != null) textObject_IRSubAccountPayable.Text = receipts[0].BankAccount ?? "";
                            if (textObject_IRSubAccountCode != null) textObject_IRSubAccountCode.Text = receipts[0].AccountNumber ?? "";
                            if (textObject_IRSubAmountPayable != null) textObject_IRSubAmountPayable.Text = receipts[0].TotalAmount.ToString("N2");
                        }

                        // Populate database compiled table before rendering subreport
                        InsertDataToItemReceiptCompiled(refNumberCR, receipts);
                    }
                    catch
                    {
                        throw;
                    }
                }
                else
                {
                    // Fallback: execute insertion even if subreport object isn't isolated by name
                    InsertDataToItemReceiptCompiled(refNumberCR, receipts);
                }

                cRIR_DRC.SetParameterValue("ReferenceNumber", refNumberCR);

                panel_Printing.Visible = false;
                panel_Signatory.Visible = true;
                panel_Main.Visible = false;
                panel_Main_CR.Visible = true;

                reportViewer.ReportSource = cRIR_DRC;
                reportViewer.RefreshReport();

                return true;
            }
            catch (Exception ex)
            {
                MessageBox.Show($"KAYAK ERROR HEHEHE:\n{ex.Message}", "Error", MessageBoxButtons.OK, MessageBoxIcon.Error);
                return false;
            }
        }

        private bool GenerateBillPaymentReport_DRC(string refNumberCR)
        {
            try
            {
                CRCV_DRCBILL cRCV_DRCBILL = new CRCV_DRCBILL();
                string databasePathBILL = Path.Combine(Application.StartupPath, "CheckDatabase.accdb");
                SetDatabaseLocation(cRCV_DRCBILL, databasePathBILL);

                AccessQueries_DRC accessQueries = new AccessQueries_DRC();
                List<BillTable> bills = accessQueries.GetBillData_DRC(refNumberCR);

                if (bills == null || bills.Count == 0)
                    return false;

                TextObject textObject_CVBILLCheckNumber = null;
                TextObject textObject_CVBILLCheckDate = null;
                TextObject textObject_CVBILLPayee = null;
                TextObject textObject_CVBILLAddress = null;
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
                TextObject textObject_CVBILLBank = null;
                TextObject textObject_CVBILLNumber = null;
                TextObject textObject_CVBILLDate = null;
                TextObject textObject_CVBILLDue = null;

                try
                {
                    textObject_CVBILLCheckNumber = cRCV_DRCBILL.ReportDefinition.ReportObjects["TextCVBILLSeriesnumber"] as TextObject;
                    textObject_CVBILLCheckDate = cRCV_DRCBILL.ReportDefinition.ReportObjects["TextCVBILLCheckDate"] as TextObject;
                    textObject_CVBILLPayee = cRCV_DRCBILL.ReportDefinition.ReportObjects["TextCVBILLPayee"] as TextObject;
                    textObject_CVBILLAddress = cRCV_DRCBILL.ReportDefinition.ReportObjects["TextCVBILLAddress"] as TextObject;
                    textObject_CVBILLTotalDebitAmount = cRCV_DRCBILL.ReportDefinition.ReportObjects["TextCVBILLTotalDebitAmount"] as TextObject;
                    textObject_CVBILLTotalCreditAmount = cRCV_DRCBILL.ReportDefinition.ReportObjects["TextCVBILLTotalCreditAmount"] as TextObject;

                    TextObject textObject_CompanyName = cRCV_DRCBILL.ReportDefinition.ReportObjects["TextCompanyName"] as TextObject;
                    if (textObject_CompanyName != null && comboBox_Company != null && comboBox_Company.SelectedItem != null)
                    {
                        textObject_CompanyName.Text = comboBox_Company.SelectedItem.ToString();
                    }


                    textObject_CVBILLBank = cRCV_DRCBILL.ReportDefinition.ReportObjects["TextCVBILLBank"] as TextObject;
                    textObject_CVBILLNumber = cRCV_DRCBILL.ReportDefinition.ReportObjects["TextCVBILLNumber"] as TextObject;
                    textObject_CVBILLDate = cRCV_DRCBILL.ReportDefinition.ReportObjects["TextCVBILLDate"] as TextObject;
                    textObject_CVBILLDue = cRCV_DRCBILL.ReportDefinition.ReportObjects["TextCVBILLDue"] as TextObject;



                    textObject_PreparedBy = cRCV_DRCBILL.ReportDefinition.ReportObjects["TextPreparedBy"] as TextObject;
                    textObject_PreparedByPos = cRCV_DRCBILL.ReportDefinition.ReportObjects["TextPreparedByPosition"] as TextObject;
                    textObject_CheckedBy = cRCV_DRCBILL.ReportDefinition.ReportObjects["TextCheckedBy"] as TextObject;
                    textObject_CheckedByPos = cRCV_DRCBILL.ReportDefinition.ReportObjects["TextCheckedByPosition"] as TextObject;
                    textObject_ApprovedBy = cRCV_DRCBILL.ReportDefinition.ReportObjects["TextApprovedBy"] as TextObject;
                    textObject_ApprovedByPos = cRCV_DRCBILL.ReportDefinition.ReportObjects["TextApprovedByPosition"] as TextObject;
                    textObject_ReceivedBy = cRCV_DRCBILL.ReportDefinition.ReportObjects["TextReceivedBy"] as TextObject;
                    textObject_ReceivedByPos = cRCV_DRCBILL.ReportDefinition.ReportObjects["TextReceivedByPosition"] as TextObject;

                    AccessToDatabase_DRC accessToDatabase = new AccessToDatabase_DRC();

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

                    if (textObject_CVBILLTotalDebitAmount != null)
                        textObject_CVBILLTotalDebitAmount.Text = $"PHP {debitTotalAmount:N2}";

                    if (textObject_CVBILLTotalCreditAmount != null)
                        textObject_CVBILLTotalCreditAmount.Text = $"PHP {debitTotalAmount:N2}";

                }
                catch
                {
                    throw;
                }


                double amount = bills[0].AmountDue;
                string amountInWords = AccessToDatabase_DRC.AmountToWordsConverter.Convert(amount);

                string rawBank = bills[0].BankAccount ?? "";

                string bank = rawBank.Contains(":")
                    ? rawBank.Split(':').Last().Trim()
                    : rawBank;

                var c = bills[0];

                // Line 1: Combine Addr1, Addr2, Addr3, Addr4 into one string separated by commas
                string streetLine = string.Join(", ", new[] {
                                                 c.VendorAddressAddr1,
                                                 c.VendorAddressAddr2,
                                                 c.VendorAddressAddr3,
                                                 c.VendorAddressAddr4
                                             }.Where(s => !string.IsNullOrWhiteSpace(s)));

                // Line 2: City (Add State/Zip here if you have them in your BillTable)
                string cityLine = string.Join(" ", new[] {
                                                 c.VendorAddressCity,
                                             }.Where(s => !string.IsNullOrWhiteSpace(s)));

                // Final: Join the two lines with a single NewLine
                string fullAddress = string.Join(Environment.NewLine, new[] { streetLine, cityLine }.Where(s => !string.IsNullOrWhiteSpace(s)));

                if (textObject_CVBILLCheckNumber != null) textObject_CVBILLCheckNumber.Text = textBox_SeriesNumber.Text;
                if (textObject_CVBILLAddress != null) textObject_CVBILLAddress.Text = fullAddress;
                if (textObject_CVBILLCheckDate != null) textObject_CVBILLCheckDate.Text = DateTime.Now.ToString("MMMM dd, yyyy");
                if (textObject_CVBILLPayee != null) textObject_CVBILLPayee.Text = bills[0].PayeeFullName ?? "";
                
                
                if (textObject_CVBILLBank != null) textObject_CVBILLBank.Text = bank;
                if (textObject_CVBILLNumber != null) textObject_CVBILLNumber.Text = bills[0].RefNumber ?? "";
                if (textObject_CVBILLDate != null) textObject_CVBILLDate.Text = bills[0].DueDate.ToString("MMMM dd, yyyy") ?? "";
                if (textObject_CVBILLDue != null)
                    textObject_CVBILLDue.Text = amount.ToString("N2");

                SubreportObject subreportObject = null;
                try
                {
                    subreportObject = cRCV_DRCBILL.ReportDefinition.ReportObjects["SubreportCVBILLDetailsIVP"] as SubreportObject;
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
                        subReportDocument = cRCV_DRCBILL.OpenSubreport(subreportObject.SubreportName);
                    }
                    catch
                    {
                        throw;
                    }

                    try
                    {
                        TextObject textObject_BILLSubRemarks = subReportDocument.ReportDefinition.ReportObjects["TextBILLRemarks"] as TextObject;
                        TextObject textObject_BILLSubAccountPayable = subReportDocument.ReportDefinition.ReportObjects["TextBILLSubAccountPayable"] as TextObject;
                        TextObject textObject_BILLSubAmountPayable = subReportDocument.ReportDefinition.ReportObjects["TextBILLSubAmountPayable"] as TextObject;
                        TextObject textObject_BILLSubAccountCode = subReportDocument.ReportDefinition.ReportObjects["TextBILLSubAccountCode"] as TextObject;
                        

                        if (textObject_BILLSubRemarks != null) textObject_BILLSubRemarks.Text = bills[0].BillMemo ?? "";
                        if (textObject_BILLSubAccountPayable != null) textObject_BILLSubAccountPayable.Text = bills[0].BankAccount ?? "";
                        if (textObject_BILLSubAccountCode != null) textObject_BILLSubAccountCode.Text = bills[0].AccountNumber ?? "";
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

                cRCV_DRCBILL.SetParameterValue("ReferenceNumber", refNumberCR);

                panel_Printing.Visible = false;
                panel_Signatory.Visible = true;
                panel_Main.Visible = false;
                panel_Main_CR.Visible = true;

                reportViewer.ReportSource = cRCV_DRCBILL;
                reportViewer.RefreshReport();

                return true;
            }
            catch (Exception ex)
            {
                MessageBox.Show($"KAYAK ERROR HEHEHE:\n{ex.Message}", "Error", MessageBoxButtons.OK, MessageBoxIcon.Error);
                return false;
            }
        }


        public static void InsertDataToItemReceiptCompiled(string refNumber, List<ItemReciept> itemReceipts)
        {
            string connectionString = AccessToDatabase_DRC.GetAccessConnectionString();
            double debitTotalAmount = 0;
            double creditTotalAmount = 0;

            using (OleDbConnection connection = new OleDbConnection(connectionString))
            {
                try
                {
                    connection.Open();

                    // 1. Clear old data
                    string deleteQuery = "DELETE FROM IR_Compiled";
                    using (OleDbCommand deleteCommand = new OleDbCommand(deleteQuery, connection))
                    {
                        deleteCommand.ExecuteNonQuery();
                    }

                    // 2. Prepare Insert Query (Replaced Amount with Debit & Credit)
                    // Order: RefNumber, AccountNumber, Item, Description, Quantity, Cost, Debit, Credit
                    string insertQuery = @"INSERT INTO IR_Compiled 
                (RefNumber, [AccountNumber], [Item], [Description], [Quantity], [Cost], [Debit], [Credit]) 
                VALUES (?, ?, ?, ?, ?, ?, ?, ?)";

                    foreach (var detail in itemReceipts)
                    {
                        string accountNumber = "";
                        string particulars = "";
                        string description = "";
                        double quantity = 0;
                        double cost = 0;
                        double amount = 0;

                        // Handle Item Lines vs Expense Lines
                        if (detail.ReceiptItemType == ReceiptItemType.ReceiptItem)
                        {
                            accountNumber = ""; // Items usually do not have an AccountNumber
                            particulars = detail.Item ?? "";
                            description = detail.ItemDescription ?? "";
                            quantity = detail.ItemQuantity;
                            cost = detail.ItemCost;
                            amount = detail.ItemAmount;
                        }
                        else if (detail.ReceiptItemType == ReceiptItemType.RecieptExpense)
                        {
                            accountNumber = detail.AccountNumber ?? ""; // Extract account number populated from QB
                            particulars = detail.Account ?? "";
                            description = detail.ExpensesMemo ?? "";
                            quantity = 0;
                            cost = 0;
                            amount = detail.ExpensesAmount;
                        }

                        // Clean particulars: Extract string after colon if sub-account/sub-item format exists (e.g. "Parent:Child")
                        if (!string.IsNullOrEmpty(particulars) && particulars.Contains(":"))
                        {
                            particulars = particulars.Substring(particulars.LastIndexOf(':') + 1).Trim();
                        }

                        // Calculate Debit / Credit strings and totals
                        string debitStr = amount > 0 ? amount.ToString("N2") : "";
                        string creditStr = amount < 0 ? Math.Abs(amount).ToString("N2") : "";

                        if (amount > 0) debitTotalAmount += amount;
                        else if (amount < 0) creditTotalAmount += Math.Abs(amount);

                        // 3. Execute Insert Command
                        using (OleDbCommand command = new OleDbCommand(insertQuery, connection))
                        {
                            // OleDb parameter ordering MUST match the SQL query order exactly
                            command.Parameters.Add("?", OleDbType.VarWChar).Value = string.IsNullOrEmpty(refNumber) ? (object)DBNull.Value : refNumber;

                            // ACCOUNT NUMBER PARAMETER
                            command.Parameters.Add("?", OleDbType.VarWChar).Value = string.IsNullOrWhiteSpace(accountNumber)
                                ? (object)DBNull.Value
                                : accountNumber;

                            command.Parameters.Add("?", OleDbType.VarWChar).Value = particulars;
                            command.Parameters.Add("?", OleDbType.VarWChar).Value = description;
                            command.Parameters.Add("?", OleDbType.Double).Value = quantity;
                            command.Parameters.Add("?", OleDbType.Double).Value = cost;
                            command.Parameters.Add("?", OleDbType.VarWChar).Value = debitStr;
                            command.Parameters.Add("?", OleDbType.VarWChar).Value = creditStr;

                            command.ExecuteNonQuery();
                        }
                    }

                    connection.Close();
                }
                catch (Exception ex)
                {
                    MessageBox.Show($"Error compiling Item Receipt data: {ex.Message}", "Error", MessageBoxButtons.OK, MessageBoxIcon.Error);
                }
            }
        }


        public static void InsertDataToCheckVoucherCompiledDRC(string refNumber, List<CheckTableExpensesAndItems> checkData)
        {
            string connectionString = AccessToDatabase_DRC.GetAccessConnectionString();
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

                // 2. Prepare Insert Query (Includes AccountNumber column)
                string insertQuery = @"
                        INSERT INTO CheckVoucherCompiled 
                        (RefNumber, [AccountNumber], [Particulars], [Class], [Debit], [Credit], [Memo], [CustomerJob]) 
                        VALUES 
                        (@RefNumber, @AccountNumber, @Particulars, @Class, @Debit, @Credit, @Memo, @CustomerJob)";

                foreach (var check in checkData)
                {
                    try
                    {
                        // COMMON FIELDS
                        string memoValue = string.IsNullOrEmpty(check.ExpensesMemo) ? "" : check.ExpensesMemo;
                        string customerJob = string.IsNullOrEmpty(check.ExpensesCustomerJob) ? "" : check.ExpensesCustomerJob;

                        // ---------------------------------------------------------
                        // INSERT ITEM ENTRY
                        // ---------------------------------------------------------
                        if (!string.IsNullOrEmpty(check.Item))
                        {
                            string itemName = check.Item;
                            string itemClass = check.ItemClass;
                            double itemAmount = check.ItemAmount;

                            string debit = itemAmount > 0 ? itemAmount.ToString("N2") : "";
                            string credit = itemAmount < 0 ? Math.Abs(itemAmount).ToString("N2") : "";

                            if (itemAmount > 0) debitTotalAmount += itemAmount;
                            else if (itemAmount < 0) creditTotalAmount += Math.Abs(itemAmount);

                            // --- DEBUG LOG ---
                            Console.WriteLine($"[ITEM ENTRY] Item: '{itemName}' | AccountNumber is NULL for Items");

                            using (OleDbCommand command = new OleDbCommand(insertQuery, connection))
                            {
                                command.Parameters.AddWithValue("@RefNumber", refNumber);
                                command.Parameters.AddWithValue("@AccountNumber", DBNull.Value); // Items usually don't have an AccountNumber
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
                        // INSERT EXPENSE ENTRY
                        // ---------------------------------------------------------
                        if (!string.IsNullOrEmpty(check.Account))
                        {
                            string accountNumber = check.AccountNumber;
                            string expenseName = check.Account;
                            string expenseClass = check.ExpenseClass;
                            double expenseAmount = check.ExpensesAmount;

                            string debit = expenseAmount > 0 ? expenseAmount.ToString("N2") : "";
                            string credit = expenseAmount < 0 ? Math.Abs(expenseAmount).ToString("N2") : "";

                            if (expenseAmount > 0) debitTotalAmount += expenseAmount;
                            else if (expenseAmount < 0) creditTotalAmount += Math.Abs(expenseAmount);

                            // --- DEBUG LOG ---
                            Console.WriteLine($"[EXPENSE ENTRY] Account: '{expenseName}' | AccountNumber Value: '{(string.IsNullOrEmpty(accountNumber) ? "<EMPTY/NULL>" : accountNumber)}'");

                            using (OleDbCommand command = new OleDbCommand(insertQuery, connection))
                            {
                                command.Parameters.AddWithValue("@RefNumber", refNumber);
                                command.Parameters.AddWithValue("@AccountNumber", string.IsNullOrEmpty(accountNumber) ? (object)DBNull.Value : accountNumber);
                                command.Parameters.AddWithValue("@Particulars", expenseName);
                                command.Parameters.AddWithValue("@Class", string.IsNullOrEmpty(expenseClass) ? (object)DBNull.Value : expenseClass);
                                command.Parameters.AddWithValue("@Debit", debit);
                                command.Parameters.AddWithValue("@Credit", credit);
                                command.Parameters.AddWithValue("@Memo", memoValue);
                                command.Parameters.AddWithValue("@CustomerJob", customerJob);

                                command.ExecuteNonQuery();
                            }
                        }

                        // ---------------------------------------------------------
                        // INSERT DESCRIPTION ONLY ENTRY
                        // ---------------------------------------------------------
                        if (string.IsNullOrEmpty(check.Item) && string.IsNullOrEmpty(check.Account) && !string.IsNullOrEmpty(check.ItemDescription))
                        {
                            // --- DEBUG LOG ---
                            Console.WriteLine($"[DESCRIPTION ENTRY] Desc: '{check.ItemDescription}' | AccountNumber is NULL");

                            using (OleDbCommand command = new OleDbCommand(insertQuery, connection))
                            {
                                command.Parameters.AddWithValue("@RefNumber", refNumber);
                                command.Parameters.AddWithValue("@AccountNumber", DBNull.Value);
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
            string connectionString = AccessToDatabase_ENA.GetAccessConnectionString();

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

                // 2. Prepare Insert Query
                string insertQuery = @"
                    INSERT INTO JV_Compiled 
                    (RefNumber, [Particulars], [Class], [Name], [Debit], [Credit], [Memo]) 
                    VALUES 
                    (@RefNumber, @Particulars, @Class, @Name, @Debit, @Credit, @Memo)";

                foreach (var line in journalData)
                {
                    try
                    {
                        // MAPPING VARIABLES (With Safe Truncation to prevent DB overflow)
                        string particulars = SafeTruncate(line.AccountName, 255);
                        string className = line.Class;
                        string nameValue = SafeTruncate(line.Name, 255);
                        string memoValue = SafeTruncate(line.Memo, 255); // Change 255 to 500 if the database column is 'Long Text/Memo'

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
            string connectionString = AccessToDatabase_DRC.GetAccessConnectionString();
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

                    // 2. PREPARE INSERT QUERY (Includes AccountNumber column)
                    // Order: RefNumber, AccountNumber, Particulars, Class, Memo, CustomerJob, Debit, Credit
                    string insertQuery = @"INSERT INTO Bill_Compiled 
                   (RefNumber, [AccountNumber], Particulars, [Class], [Memo], [CustomerJob], Debit, Credit) 
                   VALUES (?, ?, ?, ?, ?, ?, ?, ?)";

                    foreach (var bill in bills)
                    {
                        foreach (var detail in bill.ItemDetails)
                        {
                            string accountNumber = "";
                            string rawParticulars = "";
                            string classVal = "";
                            string memo = "";
                            string customerJob = "";
                            double amount = 0;

                            // Determine if this is an Item Line or an Expense Line
                            if (!string.IsNullOrEmpty(detail.ItemLineItemRefFullName))
                            {
                                accountNumber = "";
                                rawParticulars = detail.ItemLineItemRefFullName;
                                classVal = detail.ItemLineClassRefFullName ?? "";
                                memo = detail.ItemLineMemo ?? "";
                                customerJob = detail.ItemLineCustomerJob ?? "";
                                amount = detail.ItemLineAmount;
                            }
                            else if (!string.IsNullOrEmpty(detail.ExpenseLineItemRefFullName))
                            {
                                accountNumber = detail.ExpenseLineAccountNumber ?? bill.AccountNumber ?? "";
                                rawParticulars = detail.ExpenseLineItemRefFullName;
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

                            // ---------------------------------------------------------
                            // EXTRACT DATA AFTER THE COLON (":") FOR PARTICULARS
                            // ---------------------------------------------------------
                            string particulars = rawParticulars;
                            if (!string.IsNullOrEmpty(particulars) && particulars.Contains(":"))
                            {
                                // Extracts everything after the last colon and trims any leading/trailing spaces
                                particulars = particulars.Substring(particulars.LastIndexOf(':') + 1).Trim();
                            }

                            // Calculate Debit/Credit
                            string debitStr = amount > 0 ? amount.ToString("N2") : "";
                            string creditStr = amount < 0 ? Math.Abs(amount).ToString("N2") : "";

                            if (amount > 0) debitTotalAmount += amount;
                            else if (amount < 0) creditTotalAmount += Math.Abs(amount);

                            // 3. EXECUTE INSERT
                            using (OleDbCommand command = new OleDbCommand(insertQuery, connection))
                            {
                                // OleDb requires EXACT positional order as listed in the SQL query
                                command.Parameters.Add("?", OleDbType.VarWChar).Value = refNumber ?? (object)DBNull.Value;

                                // ACCOUNT NUMBER PARAMETER
                                command.Parameters.Add("?", OleDbType.VarWChar).Value = string.IsNullOrWhiteSpace(accountNumber)
                                    ? (object)DBNull.Value
                                    : accountNumber;

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

        public static void InsertDataToBillAPVCompiled(string refNumber, List<BillTable> bills)
        {
            string connectionString = AccessToDatabase_DRC.GetAccessConnectionString();

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

                    // 2. PREPARE INSERT QUERY
                    string insertQuery = @"INSERT INTO Bill_Compiled 
            (RefNumber, [AccountNumber], Particulars, [Class], [Memo], [CustomerJob], Debit, Credit) 
            VALUES (?, ?, ?, ?, ?, ?, ?, ?)";

                    // 3. FLATTEN ALL ITEM & EXPENSE LINES
                    var allLines = bills.SelectMany(bill => bill.ItemDetails.Select(detail =>
                    {
                        string accountNumber = "";
                        string rawParticulars = "";
                        string classVal = "";
                        string memo = "";
                        string customerJob = "";
                        double amount = 0;

                        if (!string.IsNullOrEmpty(detail.ItemLineItemRefFullName))
                        {
                            accountNumber = "";
                            // Fallback to ItemLineItemRefFullName if AssetAccount is empty
                            rawParticulars = !string.IsNullOrWhiteSpace(detail.ItemLineAssetAccountRefFullName)
                                ? detail.ItemLineAssetAccountRefFullName
                                : detail.ItemLineItemRefFullName;

                            classVal = detail.ItemLineClassRefFullName ?? "";
                            memo = detail.ItemLineMemo ?? "";
                            customerJob = detail.ItemLineCustomerJob ?? "";
                            amount = detail.ItemLineAmount;
                        }
                        else if (!string.IsNullOrEmpty(detail.ExpenseLineItemRefFullName))
                        {
                            accountNumber = detail.ExpenseLineAccountNumber ?? bill.AccountNumber ?? "";
                            rawParticulars = detail.ExpenseLineItemRefFullName;
                            classVal = detail.ExpenseLineClassRefFullName ?? "";
                            memo = detail.ExpenseLineMemo ?? "";
                            customerJob = detail.ExpenseLineCustomerJob ?? "";
                            amount = detail.ExpenseLineAmount;
                        }
                        else
                        {
                            return null; // Ignore empty lines
                        }

                        // Extract name after colon (e.g., "Inventories:Food" -> "Food")
                        string particulars = rawParticulars;
                        if (!string.IsNullOrEmpty(particulars) && particulars.Contains(":"))
                        {
                            particulars = particulars.Substring(particulars.LastIndexOf(':') + 1).Trim();
                        }

                        return new
                        {
                            AccountNumber = accountNumber,
                            Particulars = particulars,
                            Class = classVal,
                            Memo = memo,
                            CustomerJob = customerJob,
                            Amount = amount
                        };
                    }))
                    .Where(x => x != null && !string.IsNullOrEmpty(x.Particulars));

                    // 4. GROUP & CONSOLIDATE BY PARTICULAR & ACCOUNT NUMBER
                    var consolidatedLines = allLines
                        .GroupBy(x => new { x.Particulars, x.AccountNumber })
                        .Select(g => new
                        {
                            Particulars = g.Key.Particulars,
                            AccountNumber = g.Key.AccountNumber,
                            Class = g.First().Class,
                            Memo = g.First().Memo,
                            CustomerJob = g.First().CustomerJob,
                            TotalAmount = g.Sum(x => x.Amount) // Summed consolidated amount
                        });

                    // 5. EXECUTE INSERT FOR CONSOLIDATED ROWS
                    foreach (var item in consolidatedLines)
                    {
                        string debitStr = item.TotalAmount > 0 ? item.TotalAmount.ToString("N2") : "";
                        string creditStr = item.TotalAmount < 0 ? Math.Abs(item.TotalAmount).ToString("N2") : "";

                        using (OleDbCommand command = new OleDbCommand(insertQuery, connection))
                        {
                            command.Parameters.Add("?", OleDbType.VarWChar).Value = refNumber ?? (object)DBNull.Value;

                            command.Parameters.Add("?", OleDbType.VarWChar).Value = string.IsNullOrWhiteSpace(item.AccountNumber)
                                ? (object)DBNull.Value
                                : item.AccountNumber;

                            command.Parameters.Add("?", OleDbType.VarWChar).Value = item.Particulars ?? "";
                            command.Parameters.Add("?", string.IsNullOrWhiteSpace(item.Class) ? (object)DBNull.Value : item.Class);
                            command.Parameters.Add("?", item.Memo ?? (object)DBNull.Value);
                            command.Parameters.Add("?", item.CustomerJob ?? (object)DBNull.Value);
                            command.Parameters.Add("?", debitStr);
                            command.Parameters.Add("?", creditStr);

                            command.ExecuteNonQuery();
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
                    AccessQueries_DRC queries = new AccessQueries_DRC();

                    cheque = new List<CheckTable>();
                    bills = new List<BillTable>();
                    checks = new List<CheckTableExpensesAndItems>();
                    receipts = new List<ItemReciept>();
                    apvData = new List<BillTable>();
                    checkivp = new List<CheckTableGrid>();

                    object data = null;
                    
                    if (GlobalVariables.client == "DRC")
                    {
                        if (comboBox_Forms.SelectedIndex == 2) // Check
                        {
                            checkivp = queries.GetCheckDataDRC(refNumber);
                            data = checkivp;
                        }
                    }

                    //if (checks.Count > 0 || bills.Count > 0 || receipts.Count > 0)
                    if (data is System.Collections.ICollection colletion && colletion.Count > 0)
                    {
                        if (GlobalVariables.client == "DRC")
                        {
                            Layouts_DRC layouts_DRC = new Layouts_DRC();
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
                                layouts_DRC.PrintPage_DRC(s, ev, selectedIndex, seriesNumber, data, payeeOverride);
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

            if (GlobalVariables.client == "DRC")
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
                    itemCounter = 0;
                    pageCounter = 1;

                    if (comboBox_Forms.SelectedIndex == 3)
                    {
                        int totalItemDetails = apvData.Sum(apvData => apvData.ItemDetails.Count);
                        int totalPages = (int)Math.Ceiling((double)totalItemDetails / GlobalVariables.itemsPerPageAPV);
                        printDocument.PrinterSettings.MaximumPage = totalPages;
                    }

                    printPreviewControl.StartPage = 0;

                    PrintDialog printDialog = new PrintDialog
                    {
                        Document = printDocument,
                    };

                    if (printDialog.ShowDialog() == DialogResult.OK)
                    {
                        GlobalVariables.includeImage = false;
                        printDialog.Document.Print();

                        printPreviewControl.Visible = false;
                        printPreviewControl.Zoom = 1;
                        panel_Printing.Visible = false;

                        if (GlobalVariables.client == "DRC")
                        {
                            string formType = "";
                            if (comboBox_Forms.SelectedIndex == 1) formType = "CV";
                            else if (comboBox_Forms.SelectedIndex == 3) formType = "JV";
                            else if (comboBox_Forms.SelectedIndex == 4) formType = "APV";
                            else if (comboBox_Forms.SelectedIndex == 5) formType = "IR";

                            if (formType != "")
                            {
                                string selectedCompany = comboBox_Company.SelectedItem?.ToString();

                                if (!string.IsNullOrEmpty(selectedCompany))
                                {
                                    seriesNumber++;
                                    accessToDatabase.UpdateManualSeriesNumber(formType, seriesNumber, selectedCompany);
                                    UpdateSeriesNumberDRC(formType);
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

            if (GlobalVariables.client == "DRC")
            {
                if (comboBox_Forms.SelectedIndex == 1 || comboBox_Forms.SelectedIndex == 3 || comboBox_Forms.SelectedIndex == 4 || comboBox_Forms.SelectedIndex == 5)
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
                else if (comboBox_Forms.SelectedIndex == 5) prefix = "IR";

                if (prefix != "")
                {
                    string selectedCompany = comboBox_Company.SelectedItem?.ToString();
                    if (!string.IsNullOrEmpty(selectedCompany))
                    {
                        seriesNumber = accessToDatabase.GetSeriesNumberFromDatabase(prefix, selectedCompany);
                        UpdateSeriesNumberDRC(prefix);
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

                        panel_Main.Visible = false;
                        panel_Main_CR.Visible = true;
                        break;

                    case 5: // Item Receipt
                        prefix = "IR";
                        panel_SeriesNumber.Visible = true;
                        panel_RefNumber.Visible = false;
                        panel_RefNumberCrystalReport.Visible = true;
                        panel_Signatory.Visible = true;

                        label_SeriesNumberText.Text = "Current Series Number: IR";

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
                if (comboBox_Forms.SelectedIndex == 1)
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
            if (GlobalVariables.client == "DRC")
            {
                if (!string.IsNullOrEmpty(textBox_SeriesNumber.Text))
                {
                    string formPrefix = "";
                    if (comboBox_Forms.SelectedIndex == 1) formPrefix = "CV";
                    else if (comboBox_Forms.SelectedIndex == 3) formPrefix = "JV";
                    else if (comboBox_Forms.SelectedIndex == 4) formPrefix = "APV";
                    else if (comboBox_Forms.SelectedIndex == 5) formPrefix = "IR";

                    if (!string.IsNullOrEmpty(formPrefix))
                    {
                        string cleanInput = textBox_SeriesNumber.Text
                            .Replace(formPrefix, "")
                            .Replace("-", "")
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
            if (GlobalVariables.client == "DRC")
            {
                string formType = "";
                if (comboBox_Forms.SelectedIndex == 1) formType = "CV";
                else if (comboBox_Forms.SelectedIndex == 3) formType = "JV";
                else if (comboBox_Forms.SelectedIndex == 4) formType = "APV";
                else if (comboBox_Forms.SelectedIndex == 5) formType = "IR";

                if (!string.IsNullOrEmpty(formType) && comboBox_Company.SelectedItem != null)
                {
                    accessToDatabase.UpdateManualSeriesNumber(formType, seriesNumber, comboBox_Company.SelectedItem.ToString());
                }
            }
        }

        private void UpdateSeriesNumber(string prefix)
        {
            textBox_SeriesNumber.Text = $"{prefix}{seriesNumber:000}";
        }

        private void RefreshSeriesNumber(string columnName)
        {
            seriesNumber = accessToDatabase.GetSeriesNumberFromDatabase(columnName);
            string prefix = comboBox_Forms.SelectedIndex == 2 ? "CV" : "APV";
            textBox_SeriesNumber.Text = $"{prefix}{seriesNumber:000}";
        }

        private string GetCompanyCode(string fullCompanyName)
        {
            if (string.IsNullOrEmpty(fullCompanyName)) return "";

            switch (fullCompanyName)
            {
                case "DASMARINAS RENAL CARE CENTER INC.": return "DRC";
                default: return "";
            }
        }

        private void UpdateSeriesNumberDRC(string formPrefix)
        {
            if (accessToDatabase == null) accessToDatabase = new AccessToDatabase_DRC();
            textBox_SeriesNumber.Text = $"{formPrefix}-{seriesNumber:00000}";
        }
    }
}
