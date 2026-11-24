using CrystalDecisions.CrystalReports.Engine;
using CrystalDecisions.ReportAppServer;
using CrystalDecisions.Shared;
using CrystalDecisions.Windows.Forms;
using QBFC16Lib;
using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Data.OleDb;
using System.Drawing;
using System.Drawing.Printing;
using System.IO;
using System.Linq;
using System.Text;
using System.Text.RegularExpressions;
using System.Threading.Tasks;
using System.Windows.Forms;
using VoucherPROVER2.Clients.INT;
using VoucherPROVER2.Clients.INTEGRATED;
using static VoucherPROVER2.Clients.INT.Dataclass_INT;


namespace VoucherPROVER2.Clients.INT
{
    public partial class Dashboard_INT : Form
    {
        public Dashboard_INT()
        {
            InitializeComponent();

            accessToDatabase = new AccessToDatabase_INT();

            // 1. Reset database signatory data on open
            accessToDatabase.ResetSignatoryData();

            // 2. Pre-warm SAP Crystal Reports engine in background so first search is instant
            Task.Run(() =>
            {
                try
                {
                    using (var dummy = new CRCV_INT())
                    {
                        var _ = dummy.ReportDefinition;
                    }
                }
                catch { }
            });

            // 3. Reset database signatory data & dispose QB session on form close
            this.FormClosing += Dashboard_INT_FormClosing;

            this.CreateHandle();
        }

        private void Dashboard_INT_FormClosing(object sender, FormClosingEventArgs e)
        {
            try
            {
                _accessQueries?.Dispose();

                if (accessToDatabase != null)
                {
                    accessToDatabase.ResetSignatoryData();
                }
            }
            catch
            {
                // Suppress errors during application exit
            }
        }

        private PrintDocument printDocument;
        private PrintPreviewControl printPreviewControl;
        private CrystalReportViewer reportViewer;
        private AccessToDatabase_INT accessToDatabase;
        private AccessQueries_INT _accessQueries = new AccessQueries_INT();


        FlowLayoutPanel panel_Company;

        ComboBox comboBox_Forms;
        ComboBox comboBox_Company;

        Label label_SeriesNumberText;
        Label label_SignatoryRRStatus;
        private Label label_VoucherType;

        ComboBox comboBox_Signatory;
        TextBox textBox_SignatoryName;
        TextBox textBox_SignatoryPosition;
        Label label_SignatoryStatus;

        private ComboBox comboBox_VoucherType;

        private Label label_APAccount;
        private ComboBox comboBox_APAccount;

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
                Visible = false
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
                Visible = (GlobalVariables.client == "INT")
            };

            Label label_CompanyText = new Label
            {
                Parent = panel_Company,
                Width = sideBarWidth - 10,
                Text = "COMPANY:",
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
            // AP ACCOUNT SELECTION (Vouchers Payable / Notes Payable)
            // =========================================================================
            label_APAccount = new Label
            {
                Parent = panel_Company,
                Width = sideBarWidth - 10,
                Text = "A/P ACCOUNT:",
                TextAlign = ContentAlignment.MiddleLeft,
                Font = font_Label,
                Margin = new Padding(0, 5, 0, 0),
                Visible = (comboBox_Forms != null && (comboBox_Forms.SelectedIndex == 1 || comboBox_Forms.SelectedIndex == 4))
            };

            comboBox_APAccount = new ComboBox
            {
                Parent = panel_Company,
                Width = sideBarWidth - 28,
                DropDownStyle = ComboBoxStyle.DropDownList,
                Font = font_Label,
                Visible = (comboBox_Forms != null && (comboBox_Forms.SelectedIndex == 1 || comboBox_Forms.SelectedIndex == 4))
            };

            comboBox_APAccount.Items.AddRange(new string[]
            {
                "2300 - Vouchers Payable",
                "2100 - Notes Payable"
            });
            comboBox_APAccount.SelectedIndex = 0;

            // =========================================================================
            // VOUCHER TYPE DROPDOWN (ONLY SHOWS FOR JOURNAL VOUCHER)
            // =========================================================================
            label_VoucherType = new Label
            {
                Parent = panel_Company,
                Width = sideBarWidth - 10,
                Text = "SELECT VOUCHER TITLE:",
                TextAlign = ContentAlignment.MiddleLeft,
                Font = font_Label,
                Margin = new Padding(0, 5, 0, 0),
                Visible = (comboBox_Forms != null && comboBox_Forms.SelectedIndex == 3)
            };

            comboBox_VoucherType = new ComboBox
            {
                Parent = panel_Company,
                Width = sideBarWidth - 28,
                DropDownStyle = ComboBoxStyle.DropDownList,
                Font = font_Label,
                Visible = (comboBox_Forms != null && comboBox_Forms.SelectedIndex == 3)
            };

            comboBox_VoucherType.Items.AddRange(new string[]
            {
                "JOURNAL VOUCHER",
                "EMPLOYEES SUPPLIES VOUCHER"
            });
            comboBox_VoucherType.SelectedIndex = 0;

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
                            if (GlobalVariables.client == "INT")
                            {
                                string formType = "";
                                if (comboBox_Forms.SelectedIndex == 1) formType = "CV";
                                else if (comboBox_Forms.SelectedIndex == 3) formType = "JV";
                                else if (comboBox_Forms.SelectedIndex == 4) formType = "APV";

                                string selectedCompany = comboBox_Company.SelectedItem?.ToString();

                                if (!string.IsNullOrEmpty(formType) && !string.IsNullOrEmpty(selectedCompany))
                                {
                                    seriesNumber++;
                                    accessToDatabase.UpdateManualSeriesNumber(formType, seriesNumber, selectedCompany);
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
                BackColor = Color.FromArgb(9, 102, 176)
            };

            FlowLayoutPanel panels_Forms = Panel_SBForms();
            panels_Forms.Parent = panel_SideBar;

            panel_SeriesNumber = Panel_SBSeriesNumber();
            panel_SeriesNumber.Parent = panel_SideBar;
            panel_SeriesNumber.Visible = false;

            if (GlobalVariables.client == "INT")
            {
                FlowLayoutPanel panel_Company = Panel_SBCompany();
                panel_Company.Parent = panel_SideBar;

                FlowLayoutPanel panel_Payee = Panel_SBPayeeOverride();
                panel_Payee.Parent = panel_SideBar;
            }

            panel_RefNumber = Panel_SBRefNumber();
            panel_RefNumberCrystalReport = Panel_SBRefNumber_CR();
            panel_RefNumber.Parent = panel_SideBar;
            panel_RefNumberCrystalReport.Parent = panel_SideBar;
            panel_RefNumber.Visible = false;
            panel_RefNumberCrystalReport.Visible = false;

            panel_Signatory = Panel_SBSignatory();
            panel_Signatory.Parent = panel_SideBar;
            panel_Signatory.Visible = false;

            if (GlobalVariables.client == "LEADS")
            {
                panel_RRSignatory = Panel_SBRRSignatory();
                panel_RRSignatory.Parent = panel_SideBar;
                panel_RRSignatory.Visible = false;
            }

            FlowLayoutPanel panel_Printing = Panel_SBPrinting();
            panel_Printing.Parent = panel_SideBar;

            return panel_SideBar;
        }

        public FlowLayoutPanel Panel_SBForms()
        {
            FlowLayoutPanel panel_Forms = new FlowLayoutPanel
            {
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
                    "Check Voucher / Bills Payment",
                    "Check",
                    "Journal Voucher",
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

                    UpdateSeriesNumberINT(prefix);
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
                    UpdateSeriesNumberINT(prefix);
                }
            };

            return panel_SeriesNumber;
        }


        public FlowLayoutPanel Panel_SBRefNumber_CR()
        {
            FlowLayoutPanel panel_RefNumber_CR = new FlowLayoutPanel
            {
                Dock = DockStyle.Top,
                Height = 90,
                Width = sideBarWidth - 10,
                BackColor = Color.LightGray,
                Padding = new Padding(5, 2, 5, 5),
                BorderStyle = BorderStyle.FixedSingle,
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
                Width = sideBarWidth - 30,
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
                    return;
                }

                if (string.IsNullOrWhiteSpace(textBox_ReferenceNumber_CR.Text))
                {
                    MessageBox.Show("Please enter a reference number.", "Notice", MessageBoxButtons.OK);
                    return;
                }

                Cursor.Current = Cursors.WaitCursor;
                button_SearchRefNum_CR.Enabled = false;

                try
                {
                    if (GlobalVariables.client == "INT")
                    {
                        string refNumberCR = textBox_ReferenceNumber_CR.Text.Trim();

                        // -------------------------------------------------------------
                        // OPTION 1: CHECK VOUCHER
                        // -------------------------------------------------------------
                        if (comboBox_Forms.SelectedIndex == 1)
                        {
                            bool cvDataExists = false;
                            bool foundCheckMismatch = false;

                            try
                            {
                                cvData = _accessQueries.GetCheckExpensesAndItemsData_INT(refNumberCR, out foundCheckMismatch);

                                if (cvData != null && cvData.Count > 0)
                                {
                                    cvDataExists = true;

                                    // Insert staging data ONCE before initializing report
                                    InsertDataToCheckVoucherCompiledINT(refNumberCR, cvData);

                                    CRCV_INT cRCV_INT = new CRCV_INT();
                                    string databasePath = Path.Combine(Application.StartupPath, "CheckDatabase.accdb");
                                    SetDatabaseLocation(cRCV_INT, databasePath);

                                    var signatories = accessToDatabase.RetrieveAllSignatoryData("CV");

                                    double amount = cvData[0].TotalAmount;
                                    string amountInWords = AccessToDatabase_INT.AmountToWordsConverter.Convert(amount);
                                    string rawBank = cvData[0].BankAccount ?? "";
                                    string bank = rawBank.Contains(":")
                                        ? rawBank.Split(':').Last().Trim()
                                        : rawBank;

                                    var b = cvData[0];

                                    string streetLine = string.Join(", ", new[] {
                                        b.AddressBlockAddr1,
                                        b.AddressBlockAddr2,
                                        b.AddressBlockAddr3,
                                        b.AddressBlockAddr4
                                    }.Where(s => !string.IsNullOrWhiteSpace(s)));

                                    string cityLine = string.Join(" ", new[] {
                                        b.AddressCity,
                                    }.Where(s => !string.IsNullOrWhiteSpace(s)));

                                    string fullAddress = string.Join(Environment.NewLine, new[] { streetLine, cityLine }.Where(s => !string.IsNullOrWhiteSpace(s)));

                                    (cRCV_INT.ReportDefinition.ReportObjects["TextCVRefNumber"] as TextObject).Text = cvData[0].RefNumber;
                                    (cRCV_INT.ReportDefinition.ReportObjects["TextCVCheckDate"] as TextObject).Text = cvData[0].DueDate.ToString("MMMM dd, yyyy");
                                    (cRCV_INT.ReportDefinition.ReportObjects["TextCVPayee"] as TextObject).Text = cvData[0].PayeeFullName;
                                    (cRCV_INT.ReportDefinition.ReportObjects["TextCVAddress"] as TextObject).Text = fullAddress;
                                    (cRCV_INT.ReportDefinition.ReportObjects["TextCVDatenow"] as TextObject).Text = cvData[0].DueDate.ToString("MMMM dd, yyyy");
                                    (cRCV_INT.ReportDefinition.ReportObjects["TextCVAmount"] as TextObject).Text = cvData[0].TotalAmount.ToString("N2");

                                    (cRCV_INT.ReportDefinition.ReportObjects["TextPreparedBy"] as TextObject).Text = signatories.PreparedByName;
                                    (cRCV_INT.ReportDefinition.ReportObjects["TextPreparedByPosition"] as TextObject).Text = signatories.PreparedByPosition;
                                    (cRCV_INT.ReportDefinition.ReportObjects["TextCheckedBy"] as TextObject).Text = signatories.ReviewedByName;
                                    (cRCV_INT.ReportDefinition.ReportObjects["TextCheckedByPosition"] as TextObject).Text = signatories.ReviewedByPosition;
                                    (cRCV_INT.ReportDefinition.ReportObjects["TextApprovedBy"] as TextObject).Text = signatories.ApprovedByName;
                                    (cRCV_INT.ReportDefinition.ReportObjects["TextApprovedByPosition"] as TextObject).Text = signatories.ApprovedByPosition;

                                    (cRCV_INT.ReportDefinition.ReportObjects["TextCVAmountInWords"] as TextObject).Text = "          " + amountInWords;
                                    (cRCV_INT.ReportDefinition.ReportObjects["TextCVBankAccount"] as TextObject).Text = bank;

                                    if (cRCV_INT.ReportDefinition.ReportObjects["SubreportCVDetailsINT"] is SubreportObject subreportObject2)
                                    {
                                        ReportDocument subReportDocument2 = cRCV_INT.OpenSubreport(subreportObject2.SubreportName);
                                        (subReportDocument2.ReportDefinition.ReportObjects["TextRemarks"] as TextObject).Text = SafeTruncate(cvData[0].Memo, 500);
                                        (subReportDocument2.ReportDefinition.ReportObjects["TextCVSubTotalAmount"] as TextObject).Text = cvData[0].TotalAmount.ToString("N2");
                                    }

                                    cRCV_INT.SetParameterValue("ReferenceNumber", refNumberCR);

                                    panel_Printing.Visible = false;
                                    panel_Signatory.Visible = true;
                                    panel_Main.Visible = false;
                                    panel_Main_CR.Visible = true;

                                    reportViewer.ReportSource = cRCV_INT;
                                }
                            }
                            catch (Exception ex)
                            {
                                MessageBox.Show($"IVP CV ERROR:\n{ex.Message}", "Error", MessageBoxButtons.OK, MessageBoxIcon.Error);
                            }

                            if (!cvDataExists && !foundCheckMismatch)
                            {
                                GenerateBillPaymentReport_INT(refNumberCR);
                            }
                        }
                        // -------------------------------------------------------------
                        // OPTION 3: JOURNAL VOUCHER
                        // -------------------------------------------------------------
                        else if (comboBox_Forms.SelectedIndex == 3)
                        {
                            string selectedVoucherTitle = comboBox_VoucherType?.SelectedItem?.ToString()?.Trim() ?? "JOURNAL VOUCHER";

                            if (!ValidateVoucherReferenceNumber(selectedVoucherTitle, refNumberCR, out string validationError))
                            {
                                MessageBox.Show(validationError, "Voucher Type Mismatch", MessageBoxButtons.OK, MessageBoxIcon.Warning);
                                return;
                            }

                            journal = _accessQueries.GetJournalEntryForGrid(refNumberCR);

                            if (journal != null && journal.Count > 0)
                            {
                                InsertDataToJournalCompiled(refNumberCR, journal);

                                CRJV_INT cRJV_INT = new CRJV_INT();
                                string databasePath = Path.Combine(Application.StartupPath, "CheckDatabase.accdb");
                                SetDatabaseLocation(cRJV_INT, databasePath);

                                bool isESV = selectedVoucherTitle.StartsWith("EMPLOYEE", StringComparison.OrdinalIgnoreCase);
                                string voucherNoLabel = isESV ? "E.S.V. No.:" : "J.V. No.:";

                                if (cRJV_INT.ReportDefinition.ReportObjects["TextReportTitle"] is TextObject textObject_ReportTitle)
                                {
                                    textObject_ReportTitle.Text = selectedVoucherTitle;
                                }

                                if (cRJV_INT.ReportDefinition.ReportObjects["TextVoucherNoLabel"] is TextObject textObject_VoucherNoLabel)
                                {
                                    textObject_VoucherNoLabel.Text = voucherNoLabel;
                                }

                                TextObject textObject_JVCheckDate = cRJV_INT.ReportDefinition.ReportObjects["TextJVCheckDate"] as TextObject;
                                TextObject textObject_JVRefnumber = cRJV_INT.ReportDefinition.ReportObjects["TextJVRefnumber"] as TextObject;
                                TextObject textObject_JVMemo = cRJV_INT.ReportDefinition.ReportObjects["TextMemo"] as TextObject;

                                TextObject textObject_CompanyName = cRJV_INT.ReportDefinition.ReportObjects["TextCompanyName"] as TextObject;
                                if (textObject_CompanyName != null && comboBox_Company?.SelectedItem != null)
                                {
                                    textObject_CompanyName.Text = comboBox_Company.SelectedItem.ToString();
                                }

                                TextObject textObject_PreparedBy = cRJV_INT.ReportDefinition.ReportObjects["TextPreparedBy"] as TextObject;
                                TextObject textObject_CheckedBy = cRJV_INT.ReportDefinition.ReportObjects["TextCheckedBy"] as TextObject;
                                TextObject textObject_ApprovedBy = cRJV_INT.ReportDefinition.ReportObjects["TextApprovedBy"] as TextObject;

                                if (textObject_JVCheckDate != null)
                                    textObject_JVCheckDate.Text = journal[0].Date.ToString("MMMM dd, yyyy");

                                string firstMemo = journal
                                    ?.Select(x => x?.Memo)
                                    ?.FirstOrDefault(m => !string.IsNullOrWhiteSpace(m)) ?? "";

                                string displayMemo = firstMemo.Contains('.')
                                        ? firstMemo.Split(new[] { '.' }, 2)[1].Trim()
                                        : firstMemo.Trim();

                                if (textObject_JVRefnumber != null)
                                    textObject_JVRefnumber.Text = refNumberCR;

                                if (textObject_JVMemo != null)
                                    textObject_JVMemo.Text = displayMemo;

                                var signatories = accessToDatabase.RetrieveAllSignatoryData("JV");

                                if (textObject_PreparedBy != null) textObject_PreparedBy.Text = signatories.PreparedByName;
                                if (textObject_CheckedBy != null) textObject_CheckedBy.Text = signatories.ReviewedByName;
                                if (textObject_ApprovedBy != null) textObject_ApprovedBy.Text = signatories.ApprovedByName;

                                double totalDebit = journal.Sum(x => x.Debit);
                                double totalCredit = journal.Sum(x => x.Credit);

                                cRJV_INT.SetParameterValue("TotalDebit", totalDebit > 0 ? totalDebit.ToString("N2") : "");
                                cRJV_INT.SetParameterValue("TotalCredit", totalCredit > 0 ? totalCredit.ToString("N2") : "");
                                cRJV_INT.SetParameterValue("ReferenceNumber", refNumberCR);

                                panel_Printing.Visible = false;
                                panel_Signatory.Visible = true;
                                panel_Main.Visible = false;
                                panel_Main_CR.Visible = true;

                                reportViewer.ReportSource = cRJV_INT;
                            }
                            else
                            {
                                MessageBox.Show("No Journal Entry found for this Reference Number.");
                            }
                        }
                        // -------------------------------------------------------------
                        // OPTION 4: APV
                        // -------------------------------------------------------------
                        else if (comboBox_Forms.SelectedIndex == 4)
                        {
                            GenerateAPVReport_INT(refNumberCR);
                        }
                    }
                }
                finally
                {
                    button_SearchRefNum_CR.Enabled = true;
                    Cursor.Current = Cursors.Default;
                }
            };

            return panel_RefNumber_CR;
        }

        private string WrapLongWords(string input, int maxWordLength = 30)
        {
            if (string.IsNullOrWhiteSpace(input))
                return string.Empty;

            string[] words = input.Split(new[] { ' ' }, StringSplitOptions.None);

            for (int i = 0; i < words.Length; i++)
            {
                if (words[i].Length > maxWordLength)
                {
                    var chunks = new List<string>();
                    for (int j = 0; j < words[i].Length; j += maxWordLength)
                    {
                        chunks.Add(words[i].Substring(j, Math.Min(maxWordLength, words[i].Length - j)));
                    }
                    words[i] = string.Join(" ", chunks);
                }
            }

            return string.Join(" ", words);
        }


        private bool GenerateAPVReport_INT(string refNumberCR)
        {
            try
            {
                List<BillTable> bills = _accessQueries.GetBillData_INT_DirectBill(refNumberCR);

                if (bills == null || bills.Count == 0)
                    return false;

                string selectedAP = comboBox_APAccount?.SelectedItem?.ToString();
                if (string.IsNullOrWhiteSpace(selectedAP))
                {
                    selectedAP = comboBox_APAccount?.Text;
                }
                if (string.IsNullOrWhiteSpace(selectedAP))
                {
                    selectedAP = "Vouchers Payable";
                }

                // Populate MS Access Staging Database BEFORE loading report
                InsertDataToAPVBillCompiled(refNumberCR, bills, selectedAP);

                CRAPV_INTBILL cRAPV_INTBILL = new CRAPV_INTBILL();
                string databasePathBILL = Path.Combine(Application.StartupPath, "CheckDatabase.accdb");
                SetDatabaseLocation(cRAPV_INTBILL, databasePathBILL);

                var signatories = accessToDatabase.RetrieveAllSignatoryData("APV");

                var textObject_PreparedBy = cRAPV_INTBILL.ReportDefinition.ReportObjects["TextPreparedBy"] as TextObject;
                var textObject_PreparedByPos = cRAPV_INTBILL.ReportDefinition.ReportObjects["TextPreparedByPosition"] as TextObject;
                var textObject_CheckedBy = cRAPV_INTBILL.ReportDefinition.ReportObjects["TextCheckedBy"] as TextObject;
                var textObject_CheckedByPos = cRAPV_INTBILL.ReportDefinition.ReportObjects["TextCheckedByPosition"] as TextObject;
                var textObject_ApprovedBy = cRAPV_INTBILL.ReportDefinition.ReportObjects["TextApprovedBy"] as TextObject;
                var textObject_ApprovedByPos = cRAPV_INTBILL.ReportDefinition.ReportObjects["TextApprovedByPosition"] as TextObject;

                if (textObject_PreparedBy != null) textObject_PreparedBy.Text = signatories.PreparedByName;
                if (textObject_PreparedByPos != null) textObject_PreparedByPos.Text = signatories.PreparedByPosition;
                if (textObject_CheckedBy != null) textObject_CheckedBy.Text = signatories.ReviewedByName;
                if (textObject_CheckedByPos != null) textObject_CheckedByPos.Text = signatories.ReviewedByPosition;
                if (textObject_ApprovedBy != null) textObject_ApprovedBy.Text = signatories.ApprovedByName;
                if (textObject_ApprovedByPos != null) textObject_ApprovedByPos.Text = signatories.ApprovedByPosition;

                double totalVoucherAmount = bills.Sum(bill =>
                {
                    double lineSum = bill.ItemDetails?.Sum(d => d.ItemLineAmount > 0 ? d.ItemLineAmount : (d.ExpenseLineAmount > 0 ? d.ExpenseLineAmount : 0)) ?? 0;
                    return lineSum > 0 ? lineSum : (bill.AmountDue > 0 ? bill.AmountDue : bill.Amount);
                });

                string amountInWords = "          " + AccessToDatabase_INT.AmountToWordsConverter.Convert(totalVoucherAmount);

                string rawBank = bills.FirstOrDefault(x => !string.IsNullOrWhiteSpace(x.BankAccount))?.BankAccount ?? "";
                string bankaccount = rawBank.Contains(":") ? rawBank.Split(':').Last().Trim() : rawBank;

                var primaryBill = bills.FirstOrDefault(x => !string.IsNullOrWhiteSpace(x.VendorAddressAddr1) || !string.IsNullOrWhiteSpace(x.VendorAddressCity)) ?? bills[0];

                string streetLine = string.Join(", ", new[] {
                    primaryBill.VendorAddressAddr1,
                    primaryBill.VendorAddressAddr2,
                    primaryBill.VendorAddressAddr3,
                    primaryBill.VendorAddressAddr4
                }.Where(s => !string.IsNullOrWhiteSpace(s)));

                string cityLine = string.Join(" ", new[] {
                    primaryBill.VendorAddressCity,
                }.Where(s => !string.IsNullOrWhiteSpace(s)));

                string fullAddress = string.Join(Environment.NewLine, new[] { streetLine, cityLine }.Where(s => !string.IsNullOrWhiteSpace(s)));

                string payeeNames = string.Join(", ", bills.Select(x => x.PayeeFullName).Where(p => !string.IsNullOrWhiteSpace(p)).Distinct());
                DateTime latestDueDate = bills.Max(x => x.DueDate);

                var textObject_CVBILLAddress = cRAPV_INTBILL.ReportDefinition.ReportObjects["TextCVBILLAddress"] as TextObject;
                var textObject_CVBILLPayee = cRAPV_INTBILL.ReportDefinition.ReportObjects["TextCVBILLPayee"] as TextObject;
                var textObject_CVBILLDate = cRAPV_INTBILL.ReportDefinition.ReportObjects["TextCVBILLDate"] as TextObject;
                var textObject_CVBILLAmountInWords = cRAPV_INTBILL.ReportDefinition.ReportObjects["TextCVBILLAmountInWords"] as TextObject;
                var textObject_CVBILLCheckDate = cRAPV_INTBILL.ReportDefinition.ReportObjects["TextCVBILLCheckDate"] as TextObject;
                var textObject_CVBILLRefNumber = cRAPV_INTBILL.ReportDefinition.ReportObjects["TextCVBILLRefNumber"] as TextObject;
                var textObject_CVBILLTotalAmount = cRAPV_INTBILL.ReportDefinition.ReportObjects["TextCVBILLTotalAmount"] as TextObject;
                var textObject_CVBILLBankAccount = cRAPV_INTBILL.ReportDefinition.ReportObjects["TextCVBILLBankAccount"] as TextObject;

                if (textObject_CVBILLRefNumber != null) textObject_CVBILLRefNumber.Text = refNumberCR;
                if (textObject_CVBILLAddress != null) textObject_CVBILLAddress.Text = fullAddress;
                if (textObject_CVBILLDate != null) textObject_CVBILLDate.Text = DateTime.Now.ToString("MMMM dd, yyyy");
                if (textObject_CVBILLAmountInWords != null) textObject_CVBILLAmountInWords.Text = amountInWords;
                if (textObject_CVBILLCheckDate != null) textObject_CVBILLCheckDate.Text = latestDueDate.ToString("MMMM dd, yyyy");
                if (textObject_CVBILLPayee != null) textObject_CVBILLPayee.Text = payeeNames;
                if (textObject_CVBILLTotalAmount != null) textObject_CVBILLTotalAmount.Text = totalVoucherAmount.ToString("N2");
                if (textObject_CVBILLBankAccount != null) textObject_CVBILLBankAccount.Text = bankaccount;

                if (cRAPV_INTBILL.ReportDefinition.ReportObjects["SubreportCVDetailsINT"] is SubreportObject subreportObject)
                {
                    ReportDocument subReportDocument = cRAPV_INTBILL.OpenSubreport(subreportObject.SubreportName);

                    TextObject textObject_BILLSubRemarks = subReportDocument.ReportDefinition.ReportObjects["TextBILLRemarks"] as TextObject;
                    TextObject textObject_BILLSubAmountPayable = subReportDocument.ReportDefinition.ReportObjects["TextBILLSubAmountPayable"] as TextObject;

                    var memoLines = new List<string>();
                    var amountLines = new List<string>();

                    if (bills.Count > 1)
                    {
                        memoLines.Add("Payment for the following:");
                        amountLines.Add("");
                    }

                    foreach (var b in bills)
                    {
                        string currentMemo = !string.IsNullOrWhiteSpace(b.Memo)
                            ? b.Memo.Trim()
                            : (!string.IsNullOrWhiteSpace(b.BillMemo) ? b.BillMemo.Trim() : "");

                        if (string.IsNullOrWhiteSpace(currentMemo))
                        {
                            currentMemo = b.RefNumber ?? "";
                        }

                        double currentBillAmount = b.ItemDetails?.Sum(d =>
                            d.ItemLineAmount > 0 ? d.ItemLineAmount : (d.ExpenseLineAmount > 0 ? d.ExpenseLineAmount : 0)
                        ) ?? 0;

                        if (currentBillAmount <= 0)
                        {
                            currentBillAmount = b.AmountDue > 0 ? b.AmountDue : b.Amount;
                        }

                        memoLines.Add(currentMemo);
                        amountLines.Add(currentBillAmount.ToString("N2"));
                    }

                    if (textObject_BILLSubRemarks != null)
                        textObject_BILLSubRemarks.Text = string.Join("\r\n", memoLines);

                    if (textObject_BILLSubAmountPayable != null)
                        textObject_BILLSubAmountPayable.Text = string.Join("\r\n", amountLines);
                }

                cRAPV_INTBILL.SetParameterValue("ReferenceNumber", refNumberCR);

                panel_Printing.Visible = false;
                panel_Signatory.Visible = true;
                panel_Main.Visible = false;
                panel_Main_CR.Visible = true;

                reportViewer.ReportSource = cRAPV_INTBILL;

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
                List<BillTable> bills = _accessQueries.GetBillData_INT(refNumberCR);

                if (bills == null || bills.Count == 0)
                    return false;

                string selectedAP = comboBox_APAccount?.SelectedItem?.ToString();
                if (string.IsNullOrWhiteSpace(selectedAP))
                {
                    selectedAP = comboBox_APAccount?.Text;
                }
                if (string.IsNullOrWhiteSpace(selectedAP))
                {
                    selectedAP = !string.IsNullOrWhiteSpace(bills[0].APAccountRefFullName)
                        ? bills[0].APAccountRefFullName
                        : "Vouchers Payable";
                }

                // Populate MS Access Staging Database BEFORE loading report
                var (totalDebit, totalCredit) = InsertDataToBillCompiled(refNumberCR, bills, selectedAP);

                CRCV_INTBILL cRCV_INTBILL = new CRCV_INTBILL();
                string databasePathBILL = Path.Combine(Application.StartupPath, "CheckDatabase.accdb");
                SetDatabaseLocation(cRCV_INTBILL, databasePathBILL);

                var signatories = accessToDatabase.RetrieveAllSignatoryData("CV");

                var textObject_PreparedBy = cRCV_INTBILL.ReportDefinition.ReportObjects["TextPreparedBy"] as TextObject;
                var textObject_PreparedByPos = cRCV_INTBILL.ReportDefinition.ReportObjects["TextPreparedByPosition"] as TextObject;
                var textObject_CheckedBy = cRCV_INTBILL.ReportDefinition.ReportObjects["TextCheckedBy"] as TextObject;
                var textObject_CheckedByPos = cRCV_INTBILL.ReportDefinition.ReportObjects["TextCheckedByPosition"] as TextObject;
                var textObject_ApprovedBy = cRCV_INTBILL.ReportDefinition.ReportObjects["TextApprovedBy"] as TextObject;
                var textObject_ApprovedByPos = cRCV_INTBILL.ReportDefinition.ReportObjects["TextApprovedByPosition"] as TextObject;

                if (textObject_PreparedBy != null) textObject_PreparedBy.Text = signatories.PreparedByName;
                if (textObject_PreparedByPos != null) textObject_PreparedByPos.Text = signatories.PreparedByPosition;
                if (textObject_CheckedBy != null) textObject_CheckedBy.Text = signatories.ReviewedByName;
                if (textObject_CheckedByPos != null) textObject_CheckedByPos.Text = signatories.ReviewedByPosition;
                if (textObject_ApprovedBy != null) textObject_ApprovedBy.Text = signatories.ApprovedByName;
                if (textObject_ApprovedByPos != null) textObject_ApprovedByPos.Text = signatories.ApprovedByPosition;

                var billSummaryList = bills
                    .Select(b =>
                    {
                        string memoText = !string.IsNullOrWhiteSpace(b.BillMemo)
                            ? b.BillMemo.Trim()
                            : (!string.IsNullOrWhiteSpace(b.Memo)
                                ? b.Memo.Trim()
                                : (!string.IsNullOrWhiteSpace(b.AppliedRefNumber) ? $"SI#{b.AppliedRefNumber.Trim()}" : ""));

                        double paidAmount = b.AppliedAmount > 0
                            ? b.AppliedAmount
                            : (b.Amount > 0 ? b.Amount : b.AmountDue);

                        return new
                        {
                            Memo = memoText,
                            Amount = paidAmount
                        };
                    })
                    .ToList();

                var memoLines = new List<string>();
                var amountLines = new List<string>();

                if (billSummaryList.Count > 1)
                {
                    memoLines.Add("Payment for the following:");
                    amountLines.Add("\u00A0");
                }

                const int maxMemoCharsPerLine = 39;

                foreach (var item in billSummaryList)
                {
                    string rawMemo = (item.Memo ?? "").Trim();
                    var wrappedMemoParts = new List<string>();

                    string[] lines = rawMemo.Replace("\r\n", "\n").Split('\n');

                    foreach (var line in lines)
                    {
                        string[] words = line.Split(new[] { ' ' }, StringSplitOptions.RemoveEmptyEntries);
                        string currentLine = "";

                        foreach (var rawWord in words)
                        {
                            string word = rawWord;

                            while (word.Length > maxMemoCharsPerLine)
                            {
                                if (!string.IsNullOrEmpty(currentLine))
                                {
                                    wrappedMemoParts.Add(currentLine);
                                    currentLine = "";
                                }
                                wrappedMemoParts.Add(word.Substring(0, maxMemoCharsPerLine));
                                word = word.Substring(maxMemoCharsPerLine);
                            }

                            if (string.IsNullOrEmpty(currentLine))
                            {
                                currentLine = word;
                            }
                            else if ((currentLine.Length + 1 + word.Length) <= maxMemoCharsPerLine)
                            {
                                currentLine += " " + word;
                            }
                            else
                            {
                                wrappedMemoParts.Add(currentLine);
                                currentLine = word;
                            }
                        }

                        if (!string.IsNullOrEmpty(currentLine))
                        {
                            wrappedMemoParts.Add(currentLine);
                        }
                    }

                    if (wrappedMemoParts.Count == 0)
                    {
                        wrappedMemoParts.Add("");
                    }

                    for (int i = 0; i < wrappedMemoParts.Count; i++)
                    {
                        memoLines.Add(wrappedMemoParts[i]);
                        amountLines.Add(i == 0 ? item.Amount.ToString("N2") : "\u00A0");
                    }
                }

                string billRemarksMemoText = string.Join("\r\n", memoLines);
                string billRemarksAmountText = string.Join("\r\n", amountLines);

                double realCheckTotal = bills[0].TotalCheckAmount > 0 ? bills[0].TotalCheckAmount : billSummaryList.Sum(x => x.Amount);
                string rawWords = AccessToDatabase_INT.AmountToWordsConverter.Convert(realCheckTotal);

                string fullText = "          " + rawWords;
                int maxLine1Chars = 74;

                string line1 = fullText;
                string line2 = "";

                if (fullText.Length > maxLine1Chars)
                {
                    int breakIndex = fullText.LastIndexOf(' ', maxLine1Chars);
                    if (breakIndex != -1)
                    {
                        line1 = fullText.Substring(0, breakIndex).TrimEnd();
                        line2 = fullText.Substring(breakIndex + 1).TrimStart();
                    }
                }

                string refumber2 = bills[0].RefNumber.Contains("/") ? bills[0].RefNumber.Split('/').Last() : bills[0].RefNumber;
                string rawBankAccount = (bills[0].BankAccount ?? "").Contains(":")
                    ? (bills[0].BankAccount ?? "").Split(':').Last().Trim()
                    : (bills[0].BankAccount ?? "");

                string bankaccount = Regex.Replace(rawBankAccount, @"^\d+\s*[-:·.]*\s*", "").Trim();

                var bObj = bills[0];

                string fullAddress = string.Join(", ", new[] {
                    bObj.Address,
                    bObj.Address2,
                    bObj.VendorAddressCity
                }.Where(s => !string.IsNullOrWhiteSpace(s)).Select(s => s.Trim().TrimEnd(',')));

                var textObject_CVBILLPayee = cRCV_INTBILL.ReportDefinition.ReportObjects["TextCVBILLPayee"] as TextObject;
                var textObject_CVBILLAddress = cRCV_INTBILL.ReportDefinition.ReportObjects["TextCVBILLAddress"] as TextObject;
                var textObject_CVBILLDate = cRCV_INTBILL.ReportDefinition.ReportObjects["TextCVBILLDate"] as TextObject;
                var textObject_Line1 = cRCV_INTBILL.ReportDefinition.ReportObjects["TextAmountInWordsLine1"] as TextObject;
                var textObject_Line2 = cRCV_INTBILL.ReportDefinition.ReportObjects["TextAmountInWordsLine2"] as TextObject;
                var textObject_CVBILLCheckDate = cRCV_INTBILL.ReportDefinition.ReportObjects["TextCVBILLCheckDate"] as TextObject;
                var textObject_CVBILLRefnumber = cRCV_INTBILL.ReportDefinition.ReportObjects["TextCVRefNumber"] as TextObject;
                var textObject_CVBILLBankAccount = cRCV_INTBILL.ReportDefinition.ReportObjects["TextCVBILLBankAccount"] as TextObject;
                var textObject_CVBILLAmount = cRCV_INTBILL.ReportDefinition.ReportObjects["TextCVBILLLAmount"] as TextObject;
                var textObject_TotalDebit = cRCV_INTBILL.ReportDefinition.ReportObjects["TextCVBILLLAmountDebitdetails"] as TextObject;
                var textObject_TotalCredit = cRCV_INTBILL.ReportDefinition.ReportObjects["TextCVBILLLAmountCreditdetails"] as TextObject;
                var TextObject_CVBILLAmountdetails = cRCV_INTBILL.ReportDefinition.ReportObjects["TextCVBILLLAmountdetails"] as TextObject;

                if (textObject_CVBILLDate != null) textObject_CVBILLDate.Text = bills[0].DateCreated.ToString("MMMM dd, yyyy");
                if (textObject_CVBILLAddress != null) textObject_CVBILLAddress.Text = fullAddress;
                if (textObject_CVBILLRefnumber != null) textObject_CVBILLRefnumber.Text = refumber2;
                if (textObject_Line1 != null) textObject_Line1.Text = line1;
                if (textObject_Line2 != null) textObject_Line2.Text = line2;
                if (textObject_CVBILLAmount != null) textObject_CVBILLAmount.Text = realCheckTotal.ToString("N2");
                if (textObject_CVBILLCheckDate != null) textObject_CVBILLCheckDate.Text = bills[0].BillCheckDate.ToString("MMMM dd, yyyy");
                if (textObject_CVBILLPayee != null) textObject_CVBILLPayee.Text = bills[0].PayeeFullName ?? "";
                if (textObject_CVBILLBankAccount != null) textObject_CVBILLBankAccount.Text = bankaccount;
                if (TextObject_CVBILLAmountdetails != null) TextObject_CVBILLAmountdetails.Text = realCheckTotal.ToString("N2");

                if (textObject_TotalDebit != null)
                    textObject_TotalDebit.Text = totalDebit.ToString("N2");

                if (textObject_TotalCredit != null)
                    textObject_TotalCredit.Text = totalCredit.ToString("N2");

                if (cRCV_INTBILL.ReportDefinition.ReportObjects["SubreportCVDetailsINT"] is SubreportObject subreportObjectINT)
                {
                    ReportDocument subReportDocumentINT = cRCV_INTBILL.OpenSubreport(subreportObjectINT.SubreportName);

                    TextObject textObject_BILLSubRemarks = subReportDocumentINT.ReportDefinition.ReportObjects["TextBILLRemarks"] as TextObject;
                    TextObject textObject_BILLSubRemarksAmount = subReportDocumentINT.ReportDefinition.ReportObjects["TextBILLRemarksAmount"] as TextObject;

                    if (textObject_BILLSubRemarks != null)
                        textObject_BILLSubRemarks.Text = billRemarksMemoText;

                    if (textObject_BILLSubRemarksAmount != null)
                        textObject_BILLSubRemarksAmount.Text = billRemarksAmountText;
                }

                cRCV_INTBILL.SetParameterValue("ReferenceNumber", refNumberCR);

                panel_Printing.Visible = false;
                panel_Signatory.Visible = true;
                panel_Main.Visible = false;
                panel_Main_CR.Visible = true;

                reportViewer.ReportSource = cRCV_INTBILL;

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

            string SafeTruncate(string value, int maxLength)
            {
                if (string.IsNullOrEmpty(value)) return "";
                return value.Length <= maxLength ? value : value.Substring(0, maxLength);
            }

            using (OleDbConnection connection = new OleDbConnection(connectionString))
            {
                connection.Open();
                using (OleDbTransaction tx = connection.BeginTransaction())
                {
                    try
                    {
                        using (OleDbCommand deleteCommand = new OleDbCommand("DELETE FROM CheckVoucherCompiled", connection, tx))
                        {
                            deleteCommand.ExecuteNonQuery();
                        }

                        string insertQuery = @"
                        INSERT INTO CheckVoucherCompiled 
                        (RefNumber, [Particulars], [Class], [Debit], [Credit], [Memo], [CustomerJob]) 
                        VALUES 
                        (@RefNumber, @Particulars, @Class, @Debit, @Credit, @Memo, @CustomerJob)";

                        using (OleDbCommand command = new OleDbCommand(insertQuery, connection, tx))
                        {
                            command.Parameters.Add("@RefNumber", OleDbType.VarWChar, 50);
                            command.Parameters.Add("@Particulars", OleDbType.VarWChar, 255);
                            command.Parameters.Add("@Class", OleDbType.VarWChar, 100);
                            command.Parameters.Add("@Debit", OleDbType.VarWChar, 50);
                            command.Parameters.Add("@Credit", OleDbType.VarWChar, 50);
                            command.Parameters.Add("@Memo", OleDbType.VarWChar, 255);
                            command.Parameters.Add("@CustomerJob", OleDbType.VarWChar, 255);

                            var groupedItemDebits = checkData
                                .Where(x => !string.IsNullOrEmpty(x.Item) && x.ItemAmount > 0)
                                .GroupBy(x => new
                                {
                                    Name = x.Item.Trim(),
                                    Class = x.ItemClass ?? "",
                                    Memo = x.ExpensesMemo ?? "",
                                    CustomerJob = x.ExpensesCustomerJob ?? ""
                                })
                                .Select(g => new
                                {
                                    Particulars = g.Key.Name,
                                    Class = g.Key.Class,
                                    Memo = g.Key.Memo,
                                    CustomerJob = g.Key.CustomerJob,
                                    TotalAmount = g.Sum(x => x.ItemAmount)
                                });

                            var groupedExpenseDebits = checkData
                                .Where(x => !string.IsNullOrEmpty(x.Account) && x.ExpensesAmount > 0)
                                .GroupBy(x => new
                                {
                                    Name = x.Account.Trim(),
                                    Class = x.ExpenseClass ?? "",
                                    Memo = x.ExpensesMemo ?? "",
                                    CustomerJob = x.ExpensesCustomerJob ?? ""
                                })
                                .Select(g => new
                                {
                                    Particulars = g.Key.Name,
                                    Class = g.Key.Class,
                                    Memo = g.Key.Memo,
                                    CustomerJob = g.Key.CustomerJob,
                                    TotalAmount = g.Sum(x => x.ExpensesAmount)
                                });

                            foreach (var entry in groupedItemDebits.Concat(groupedExpenseDebits))
                            {
                                debitTotalAmount += entry.TotalAmount;
                                command.Parameters["@RefNumber"].Value = refNumber;
                                command.Parameters["@Particulars"].Value = SafeTruncate(entry.Particulars, 255);
                                command.Parameters["@Class"].Value = string.IsNullOrEmpty(entry.Class) ? (object)DBNull.Value : entry.Class;
                                command.Parameters["@Debit"].Value = entry.TotalAmount.ToString("N2");
                                command.Parameters["@Credit"].Value = "";
                                command.Parameters["@Memo"].Value = SafeTruncate(entry.Memo, 255);
                                command.Parameters["@CustomerJob"].Value = SafeTruncate(entry.CustomerJob, 255);
                                command.ExecuteNonQuery();
                            }

                            var groupedItemCredits = checkData
                                .Where(x => !string.IsNullOrEmpty(x.Item) && x.ItemAmount < 0)
                                .GroupBy(x => new
                                {
                                    Name = x.Item.Trim(),
                                    Class = x.ItemClass ?? "",
                                    Memo = x.ExpensesMemo ?? "",
                                    CustomerJob = x.ExpensesCustomerJob ?? ""
                                })
                                .Select(g => new
                                {
                                    Particulars = g.Key.Name,
                                    Class = g.Key.Class,
                                    Memo = g.Key.Memo,
                                    CustomerJob = g.Key.CustomerJob,
                                    TotalAmount = Math.Abs(g.Sum(x => x.ItemAmount))
                                });

                            var groupedExpenseCredits = checkData
                                .Where(x => !string.IsNullOrEmpty(x.Account) && x.ExpensesAmount < 0)
                                .GroupBy(x => new
                                {
                                    Name = x.Account.Trim(),
                                    Class = x.ExpenseClass ?? "",
                                    Memo = x.ExpensesMemo ?? "",
                                    CustomerJob = x.ExpensesCustomerJob ?? ""
                                })
                                .Select(g => new
                                {
                                    Particulars = g.Key.Name,
                                    Class = g.Key.Class,
                                    Memo = g.Key.Memo,
                                    CustomerJob = g.Key.CustomerJob,
                                    TotalAmount = Math.Abs(g.Sum(x => x.ExpensesAmount))
                                });

                            foreach (var entry in groupedItemCredits.Concat(groupedExpenseCredits))
                            {
                                creditTotalAmount += entry.TotalAmount;
                                command.Parameters["@RefNumber"].Value = refNumber;
                                command.Parameters["@Particulars"].Value = SafeTruncate(entry.Particulars, 255);
                                command.Parameters["@Class"].Value = string.IsNullOrEmpty(entry.Class) ? (object)DBNull.Value : entry.Class;
                                command.Parameters["@Debit"].Value = "";
                                command.Parameters["@Credit"].Value = entry.TotalAmount.ToString("N2");
                                command.Parameters["@Memo"].Value = SafeTruncate(entry.Memo, 255);
                                command.Parameters["@CustomerJob"].Value = SafeTruncate(entry.CustomerJob, 255);
                                command.ExecuteNonQuery();
                            }

                            if (checkData != null && checkData.Count > 0)
                            {
                                var mainCheck = checkData[0];
                                double finalCheckCredit = mainCheck.TotalAmount;
                                creditTotalAmount += finalCheckCredit;

                                string bankName = mainCheck.BankAccount.Contains(":")
                                    ? mainCheck.BankAccount.Split(':').Last().Trim()
                                    : mainCheck.BankAccount;

                                command.Parameters["@RefNumber"].Value = refNumber;
                                command.Parameters["@Particulars"].Value = SafeTruncate(bankName, 255);
                                command.Parameters["@Class"].Value = DBNull.Value;
                                command.Parameters["@Debit"].Value = "";
                                command.Parameters["@Credit"].Value = finalCheckCredit.ToString("N2");
                                command.Parameters["@Memo"].Value = SafeTruncate(mainCheck.Memo, 255);
                                command.Parameters["@CustomerJob"].Value = DBNull.Value;
                                command.ExecuteNonQuery();
                            }

                            var descriptionOnlyEntries = checkData
                                .Where(x => string.IsNullOrEmpty(x.Item) && string.IsNullOrEmpty(x.Account) && !string.IsNullOrEmpty(x.ItemDescription))
                                .GroupBy(x => new
                                {
                                    Description = x.ItemDescription.Trim(),
                                    Memo = x.ExpensesMemo ?? "",
                                    CustomerJob = x.ExpensesCustomerJob ?? ""
                                })
                                .Select(g => g.Key);

                            foreach (var desc in descriptionOnlyEntries)
                            {
                                command.Parameters["@RefNumber"].Value = refNumber;
                                command.Parameters["@Particulars"].Value = SafeTruncate(desc.Description, 255);
                                command.Parameters["@Class"].Value = DBNull.Value;
                                command.Parameters["@Debit"].Value = "";
                                command.Parameters["@Credit"].Value = "";
                                command.Parameters["@Memo"].Value = SafeTruncate(desc.Memo, 255);
                                command.Parameters["@CustomerJob"].Value = SafeTruncate(desc.CustomerJob, 255);
                                command.ExecuteNonQuery();
                            }
                        }

                        tx.Commit();
                    }
                    catch (Exception ex)
                    {
                        tx.Rollback();
                        MessageBox.Show($"Error saving check voucher data: {ex.Message}", "Error", MessageBoxButtons.OK, MessageBoxIcon.Error);
                    }
                }
            }
        }

        public static void InsertDataToJournalCompiled(string refNumber, List<JournalGridItem> journalData)
        {
            string connectionString = AccessToDatabase_INT.GetAccessConnectionString();

            string SafeTruncate(string value, int maxLength)
            {
                if (string.IsNullOrEmpty(value)) return "";
                return value.Length <= maxLength ? value : value.Substring(0, maxLength);
            }

            using (OleDbConnection connection = new OleDbConnection(connectionString))
            {
                connection.Open();
                using (OleDbTransaction tx = connection.BeginTransaction())
                {
                    try
                    {
                        using (OleDbCommand deleteCommand = new OleDbCommand("DELETE FROM JV_Compiled", connection, tx))
                        {
                            deleteCommand.ExecuteNonQuery();
                        }

                        string insertQuery = @"INSERT INTO JV_Compiled 
                                       (RefNumber, [AccountNumber], [Particulars], [Class], [Name], [Debit], [Credit], [Memo]) 
                                       VALUES 
                                       (@RefNumber, @AccountNumber, @Particulars, @Class, @Name, @Debit, @Credit, @Memo)";

                        using (OleDbCommand command = new OleDbCommand(insertQuery, connection, tx))
                        {
                            command.Parameters.Add("@RefNumber", OleDbType.VarWChar, 50);
                            command.Parameters.Add("@AccountNumber", OleDbType.VarWChar, 50);
                            command.Parameters.Add("@Particulars", OleDbType.VarWChar, 255);
                            command.Parameters.Add("@Class", OleDbType.VarWChar, 100);
                            command.Parameters.Add("@Name", OleDbType.VarWChar, 255);
                            command.Parameters.Add("@Debit", OleDbType.VarWChar, 50);
                            command.Parameters.Add("@Credit", OleDbType.VarWChar, 50);
                            command.Parameters.Add("@Memo", OleDbType.VarWChar, 255);

                            foreach (var line in journalData)
                            {
                                string accountNumber = SafeTruncate(line.AccountNumber, 50);
                                string particulars = SafeTruncate(line.AccountName, 255);
                                string className = line.Class;
                                string nameValue = SafeTruncate(line.Name, 255);
                                string memoValue = SafeTruncate(line.Memo, 255);

                                string debitStr = line.Debit != 0 ? line.Debit.ToString("N2") : "";
                                string creditStr = (line.Debit == 0 && line.Credit != 0) ? line.Credit.ToString("N2") : "";

                                command.Parameters["@RefNumber"].Value = refNumber;
                                command.Parameters["@AccountNumber"].Value = string.IsNullOrEmpty(accountNumber) ? (object)DBNull.Value : accountNumber;
                                command.Parameters["@Particulars"].Value = particulars;
                                command.Parameters["@Class"].Value = string.IsNullOrEmpty(className) ? (object)DBNull.Value : className;
                                command.Parameters["@Name"].Value = string.IsNullOrEmpty(nameValue) ? (object)DBNull.Value : nameValue;
                                command.Parameters["@Debit"].Value = debitStr;
                                command.Parameters["@Credit"].Value = creditStr;
                                command.Parameters["@Memo"].Value = memoValue;
                                command.ExecuteNonQuery();
                            }
                        }

                        tx.Commit();
                    }
                    catch (Exception ex)
                    {
                        tx.Rollback();
                        MessageBox.Show($"Error processing journal lines: {ex.Message}", "Error", MessageBoxButtons.OK, MessageBoxIcon.Error);
                    }
                }
            }
        }

        public static (double TotalDebit, double TotalCredit) InsertDataToBillCompiled(string refNumber, List<BillTable> bills, string selectedAPAccount = "Vouchers Payable")
        {
            string connectionString = AccessToDatabase_INT.GetAccessConnectionString();
            double debitTotalAmount = 0;
            double creditTotalAmount = 0;

            string SafeTruncate(string value, int maxLength)
            {
                if (string.IsNullOrEmpty(value)) return "";
                return value.Length <= maxLength ? value : value.Substring(0, maxLength);
            }

            string GetCleanAccountName(string raw)
            {
                if (string.IsNullOrWhiteSpace(raw)) return "";
                var match = Regex.Match(raw.Trim(), @"^[\d]+(?:[-\.][\dA-Za-z]+)*\s*[-:.]?\s*(.*)$");
                if (match.Success && !string.IsNullOrWhiteSpace(match.Groups[1].Value))
                {
                    return match.Groups[1].Value.Trim();
                }
                return raw.Trim();
            }

            string FormatParticularSpacing(string raw)
            {
                if (string.IsNullOrWhiteSpace(raw)) return "";

                var match = Regex.Match(raw.Trim(), @"^([\d]+(?:[-\.][\dA-Za-z]+)*)\s*[-:.]\s*(.*)$");
                if (match.Success)
                {
                    string code = match.Groups[1].Value.Trim();
                    string name = match.Groups[2].Value.Trim();

                    int spacesCount;
                    if (code.Length >= 7) spacesCount = 1;
                    else if (code.Length == 6) spacesCount = 3;
                    else if (code.Length == 5) spacesCount = 4;
                    else if (code.Length == 4) spacesCount = 6;
                    else spacesCount = 3;

                    return $"{code}{new string(' ', spacesCount)}{name}";
                }

                return raw.Trim();
            }

            using (OleDbConnection connection = new OleDbConnection(connectionString))
            {
                connection.Open();
                using (OleDbTransaction tx = connection.BeginTransaction())
                {
                    try
                    {
                        using (OleDbCommand deleteCommand = new OleDbCommand("DELETE FROM Bill_Compiled", connection, tx))
                        {
                            deleteCommand.ExecuteNonQuery();
                        }

                        string insertQuery = @"INSERT INTO Bill_Compiled (RefNumber, [Particulars], [Class], [Debit], [Credit], [Memo], [CustomerJob]) 
                                               VALUES (@RefNumber, @Particulars, @Class, @Debit, @Credit, @Memo, @CustomerJob)";

                        var allDetails = bills
                            .Where(b => b.ItemDetails != null)
                            .SelectMany(b => b.ItemDetails.Select(d => new { Bill = b, Detail = d }))
                            .ToList();

                        var groupedDebits = allDetails
                            .Where(x => !string.IsNullOrEmpty(x.Detail.ItemLineItemRefFullName) && x.Detail.ItemLineAmount > 0)
                            .GroupBy(x => x.Detail.ItemLineItemRefFullName.Trim())
                            .Select(g => new
                            {
                                Particulars = g.Key,
                                Memo = string.Join("; ", g.Select(x => x.Detail.ItemLineMemo).Where(m => !string.IsNullOrWhiteSpace(m)).Distinct()),
                                TotalAmount = g.Sum(x => x.Detail.ItemLineAmount)
                            })
                            .OrderBy(x => GetCleanAccountName(x.Particulars), StringComparer.OrdinalIgnoreCase)
                            .ToList();

                        using (OleDbCommand command = new OleDbCommand(insertQuery, connection, tx))
                        {
                            command.Parameters.Add("@RefNumber", OleDbType.VarWChar, 50);
                            command.Parameters.Add("@Particulars", OleDbType.VarWChar, 255);
                            command.Parameters.Add("@Class", OleDbType.VarWChar, 100);
                            command.Parameters.Add("@Debit", OleDbType.VarWChar, 50);
                            command.Parameters.Add("@Credit", OleDbType.VarWChar, 50);
                            command.Parameters.Add("@Memo", OleDbType.VarWChar, 255);
                            command.Parameters.Add("@CustomerJob", OleDbType.VarWChar, 255);

                            foreach (var entry in groupedDebits)
                            {
                                debitTotalAmount += entry.TotalAmount;
                                command.Parameters["@RefNumber"].Value = (object)refNumber ?? DBNull.Value;
                                command.Parameters["@Particulars"].Value = SafeTruncate(FormatParticularSpacing(entry.Particulars), 255);
                                command.Parameters["@Class"].Value = DBNull.Value;
                                command.Parameters["@Debit"].Value = entry.TotalAmount.ToString("N2");
                                command.Parameters["@Credit"].Value = "";
                                command.Parameters["@Memo"].Value = SafeTruncate(entry.Memo, 255);
                                command.Parameters["@CustomerJob"].Value = DBNull.Value;
                                command.ExecuteNonQuery();
                            }

                            var creditEntries = new List<(string Particulars, double Amount, string Memo)>();
                            double totalDeductedCredits = 0;

                            var groupedExpenseCredits = allDetails
                                .Where(x => !string.IsNullOrEmpty(x.Detail.ItemLineItemRefFullName) && x.Detail.ItemLineAmount < 0)
                                .GroupBy(x => x.Detail.ItemLineItemRefFullName.Trim())
                                .Select(g => new
                                {
                                    Particulars = g.Key,
                                    Memo = string.Join("; ", g.Select(x => x.Detail.ItemLineMemo).Where(m => !string.IsNullOrWhiteSpace(m)).Distinct()),
                                    TotalCreditAmount = Math.Abs(g.Sum(x => x.Detail.ItemLineAmount))
                                });

                            foreach (var exp in groupedExpenseCredits)
                            {
                                totalDeductedCredits += exp.TotalCreditAmount;
                                creditEntries.Add((exp.Particulars, exp.TotalCreditAmount, exp.Memo));
                            }

                            var groupedDiscounts = bills
                                .Where(b => b.AppliedToTxnDiscountAmount > 0 && !string.IsNullOrEmpty(b.AppliedToTxnDiscountAccountRefFullName))
                                .GroupBy(b => b.AppliedToTxnDiscountAccountRefFullName.Trim())
                                .Select(g => new
                                {
                                    AccountName = g.Key,
                                    TotalDiscount = g.Sum(b => Math.Abs(b.AppliedToTxnDiscountAmount))
                                });

                            foreach (var disc in groupedDiscounts)
                            {
                                totalDeductedCredits += disc.TotalDiscount;
                                creditEntries.Add((disc.AccountName, disc.TotalDiscount, "Withholding Tax Applied"));
                            }

                            var groupedBillCredits = bills
                                .Where(b => b.AppliedBillCredits != null && b.AppliedBillCredits.Count > 0)
                                .SelectMany(b => b.AppliedBillCredits)
                                .GroupBy(c => c.AccountRefFullName.Trim())
                                .Select(g => new
                                {
                                    AccountName = g.Key,
                                    TotalCreditApplied = g.Sum(c => c.AppliedAmount),
                                    Memo = string.Join("; ", g.Select(c => $"Credit #{c.CreditRefNumber} {c.Memo}".Trim()).Distinct())
                                });

                            foreach (var cred in groupedBillCredits)
                            {
                                totalDeductedCredits += cred.TotalCreditApplied;
                                creditEntries.Add((cred.AccountName, cred.TotalCreditApplied, cred.Memo));
                            }

                            if (bills != null && bills.Count > 0)
                            {
                                double netPaymentCredit = debitTotalAmount - totalDeductedCredits;
                                string consolidatedMemo = string.Join(" | ", bills
                                    .Select(b => !string.IsNullOrWhiteSpace(b.BillMemo) ? b.BillMemo.Trim() : b.Memo?.Trim())
                                    .Where(m => !string.IsNullOrWhiteSpace(m))
                                    .Distinct());

                                string apAccount = string.IsNullOrWhiteSpace(selectedAPAccount) ? "Vouchers Payable" : selectedAPAccount;

                                if (netPaymentCredit > 0)
                                {
                                    creditEntries.Add((apAccount, netPaymentCredit, consolidatedMemo));
                                }
                            }

                            var sortedCredits = creditEntries
                                .OrderBy(c => GetCleanAccountName(c.Particulars), StringComparer.OrdinalIgnoreCase)
                                .ToList();

                            foreach (var credit in sortedCredits)
                            {
                                creditTotalAmount += credit.Amount;
                                command.Parameters["@RefNumber"].Value = (object)refNumber ?? DBNull.Value;
                                command.Parameters["@Particulars"].Value = SafeTruncate(FormatParticularSpacing(credit.Particulars), 255);
                                command.Parameters["@Class"].Value = DBNull.Value;
                                command.Parameters["@Debit"].Value = "";
                                command.Parameters["@Credit"].Value = credit.Amount.ToString("N2");
                                command.Parameters["@Memo"].Value = SafeTruncate(credit.Memo, 255);
                                command.Parameters["@CustomerJob"].Value = DBNull.Value;
                                command.ExecuteNonQuery();
                            }
                        }

                        tx.Commit();
                    }
                    catch (Exception ex)
                    {
                        tx.Rollback();
                        MessageBox.Show($"Error saving bill data: {ex.Message}", "Error", MessageBoxButtons.OK, MessageBoxIcon.Error);
                        return (0, 0);
                    }
                }
            }

            return (debitTotalAmount, creditTotalAmount);
        }

        public static void InsertDataToAPVBillCompiled(string refNumber, List<BillTable> bills, string selectedAPAccount = "Vouchers Payable")
        {
            string connectionString = AccessToDatabase_INT.GetAccessConnectionString();
            double debitTotalAmount = 0;

            string SafeTruncate(string value, int maxLength)
            {
                if (string.IsNullOrEmpty(value)) return "";
                return value.Length <= maxLength ? value : value.Substring(0, maxLength);
            }

            using (OleDbConnection connection = new OleDbConnection(connectionString))
            {
                connection.Open();
                using (OleDbTransaction tx = connection.BeginTransaction())
                {
                    try
                    {
                        using (OleDbCommand deleteCommand = new OleDbCommand("DELETE FROM Bill_Compiled", connection, tx))
                        {
                            deleteCommand.ExecuteNonQuery();
                        }

                        string insertQuery = @"
                        INSERT INTO Bill_Compiled 
                        (RefNumber, [Particulars], [Class], [Debit], [Credit], [Memo], [CustomerJob]) 
                        VALUES 
                        (@RefNumber, @Particulars, @Class, @Debit, @Credit, @Memo, @CustomerJob)";

                        var allDetails = bills
                            .Where(b => b.ItemDetails != null)
                            .SelectMany(b => b.ItemDetails.Select(d => new { Bill = b, Detail = d }))
                            .ToList();

                        using (OleDbCommand command = new OleDbCommand(insertQuery, connection, tx))
                        {
                            command.Parameters.Add("@RefNumber", OleDbType.VarWChar, 50);
                            command.Parameters.Add("@Particulars", OleDbType.VarWChar, 255);
                            command.Parameters.Add("@Class", OleDbType.VarWChar, 100);
                            command.Parameters.Add("@Debit", OleDbType.VarWChar, 50);
                            command.Parameters.Add("@Credit", OleDbType.VarWChar, 50);
                            command.Parameters.Add("@Memo", OleDbType.VarWChar, 255);
                            command.Parameters.Add("@CustomerJob", OleDbType.VarWChar, 255);

                            var groupedItemDebits = allDetails
                                .Where(x => !string.IsNullOrEmpty(x.Detail.ItemLineItemRefFullName) && x.Detail.ItemLineAmount > 0)
                                .GroupBy(x => x.Detail.ItemLineItemRefFullName.Trim())
                                .Select(g => new
                                {
                                    Particulars = g.Key,
                                    Memo = string.Join("; ", g.Select(x => x.Detail.ItemLineMemo).Where(m => !string.IsNullOrWhiteSpace(m)).Distinct()),
                                    TotalAmount = g.Sum(x => x.Detail.ItemLineAmount)
                                });

                            var groupedExpenseDebits = allDetails
                                .Where(x => !string.IsNullOrEmpty(x.Detail.ExpenseLineItemRefFullName) && x.Detail.ExpenseLineAmount > 0)
                                .GroupBy(x => x.Detail.ExpenseLineItemRefFullName.Trim())
                                .Select(g => new
                                {
                                    Particulars = g.Key,
                                    Memo = string.Join("; ", g.Select(x => x.Detail.ExpenseLineMemo).Where(m => !string.IsNullOrWhiteSpace(m)).Distinct()),
                                    TotalAmount = g.Sum(x => x.Detail.ExpenseLineAmount)
                                });

                            foreach (var entry in groupedItemDebits.Concat(groupedExpenseDebits))
                            {
                                debitTotalAmount += entry.TotalAmount;
                                command.Parameters["@RefNumber"].Value = (object)refNumber ?? DBNull.Value;
                                command.Parameters["@Particulars"].Value = SafeTruncate(entry.Particulars, 255);
                                command.Parameters["@Class"].Value = DBNull.Value;
                                command.Parameters["@Debit"].Value = entry.TotalAmount.ToString("N2");
                                command.Parameters["@Credit"].Value = "";
                                command.Parameters["@Memo"].Value = SafeTruncate(entry.Memo, 255);
                                command.Parameters["@CustomerJob"].Value = DBNull.Value;
                                command.ExecuteNonQuery();
                            }

                            var groupedExpenseCredits = allDetails
                                .Where(x => !string.IsNullOrEmpty(x.Detail.ExpenseLineItemRefFullName) && x.Detail.ExpenseLineAmount < 0)
                                .GroupBy(x =>
                                {
                                    string rawAcc = x.Detail.ExpenseLineItemRefFullName.Trim();
                                    return rawAcc.Contains(":") ? rawAcc.Split(':').Last().Trim() : rawAcc;
                                })
                                .Select(g => new
                                {
                                    Particulars = g.Key,
                                    Memo = string.Join("; ", g.Select(x => x.Detail.ExpenseLineMemo).Where(m => !string.IsNullOrWhiteSpace(m)).Distinct()),
                                    TotalCreditAmount = Math.Abs(g.Sum(x => x.Detail.ExpenseLineAmount))
                                });

                            double totalNegativeCredits = 0;

                            foreach (var entry in groupedExpenseCredits)
                            {
                                totalNegativeCredits += entry.TotalCreditAmount;
                                command.Parameters["@RefNumber"].Value = (object)refNumber ?? DBNull.Value;
                                command.Parameters["@Particulars"].Value = SafeTruncate(entry.Particulars, 255);
                                command.Parameters["@Class"].Value = DBNull.Value;
                                command.Parameters["@Debit"].Value = "";
                                command.Parameters["@Credit"].Value = entry.TotalCreditAmount.ToString("N2");
                                command.Parameters["@Memo"].Value = SafeTruncate(entry.Memo, 255);
                                command.Parameters["@CustomerJob"].Value = DBNull.Value;
                                command.ExecuteNonQuery();
                            }

                            var groupedDiscounts = bills
                                .Where(b => b.AppliedToTxnDiscountAmount > 0 && !string.IsNullOrEmpty(b.AppliedToTxnDiscountAccountRefFullName))
                                .GroupBy(b =>
                                {
                                    string rawAcc = b.AppliedToTxnDiscountAccountRefFullName;
                                    return rawAcc.Contains(":") ? rawAcc.Split(':').Last().Trim() : rawAcc.Trim();
                                })
                                .Select(g => new
                                {
                                    AccountName = g.Key,
                                    Memos = string.Join("; ", bills.Select(x => x.Memo ?? x.BillMemo).Where(m => !string.IsNullOrWhiteSpace(m)).Distinct()),
                                    TotalDiscount = g.Sum(b => Math.Abs(b.AppliedToTxnDiscountAmount))
                                });

                            foreach (var disc in groupedDiscounts)
                            {
                                totalNegativeCredits += disc.TotalDiscount;
                                string discMemo = !string.IsNullOrWhiteSpace(disc.Memos) ? disc.Memos : "Withholding Tax Applied";

                                command.Parameters["@RefNumber"].Value = (object)refNumber ?? DBNull.Value;
                                command.Parameters["@Particulars"].Value = SafeTruncate(disc.AccountName, 255);
                                command.Parameters["@Class"].Value = DBNull.Value;
                                command.Parameters["@Debit"].Value = "";
                                command.Parameters["@Credit"].Value = disc.TotalDiscount.ToString("N2");
                                command.Parameters["@Memo"].Value = SafeTruncate(discMemo, 255);
                                command.Parameters["@CustomerJob"].Value = DBNull.Value;
                                command.ExecuteNonQuery();
                            }

                            if (bills != null && bills.Count > 0)
                            {
                                double netPaymentCredit = debitTotalAmount - totalNegativeCredits;
                                string allHeaderMemos = string.Join(" | ", bills
                                    .Select(b => !string.IsNullOrWhiteSpace(b.BillMemo) ? b.BillMemo.Trim() : b.Memo?.Trim())
                                    .Where(m => !string.IsNullOrWhiteSpace(m))
                                    .Distinct());

                                string apAccount = string.IsNullOrWhiteSpace(selectedAPAccount) ? "Vouchers Payable" : selectedAPAccount;

                                command.Parameters["@RefNumber"].Value = (object)refNumber ?? DBNull.Value;
                                command.Parameters["@Particulars"].Value = SafeTruncate(apAccount, 255);
                                command.Parameters["@Class"].Value = DBNull.Value;
                                command.Parameters["@Debit"].Value = "";
                                command.Parameters["@Credit"].Value = netPaymentCredit.ToString("N2");
                                command.Parameters["@Memo"].Value = SafeTruncate(allHeaderMemos, 255);
                                command.Parameters["@CustomerJob"].Value = DBNull.Value;
                                command.ExecuteNonQuery();
                            }
                        }

                        tx.Commit();
                    }
                    catch (Exception ex)
                    {
                        tx.Rollback();
                        MessageBox.Show($"Error saving APV data: {ex.Message}", "Error", MessageBoxButtons.OK, MessageBoxIcon.Error);
                    }
                }
            }
        }



        private FlowLayoutPanel Panel_SBRefNumber()
        {
            FlowLayoutPanel panel_RefNumber = new FlowLayoutPanel
            {
                Dock = DockStyle.Top,
                Height = 90,
                Width = sideBarWidth - 10,
                BackColor = Color.LightGray,
                Padding = new Padding(5, 2, 5, 5),
                BorderStyle = BorderStyle.FixedSingle,
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
                Width = sideBarWidth - 30,
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
                    Cursor.Current = Cursors.WaitCursor;
                    button_SearchRefNum.Enabled = false;

                    try
                    {
                        string refNumber = textBox_ReferenceNumber.Text;

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
                                checkivp = _accessQueries.GetCheckDataINT(refNumber);
                                data = checkivp;
                            }
                        }

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
                                string payeeOverride = textBox_PayeeOverride.Text;

                                itemCounter = 0;
                                pageCounter = 1;
                                printPreviewControl.StartPage = 0;

                                printDocument.PrintPage += (s, ev) =>
                                {
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
                    finally
                    {
                        button_SearchRefNum.Enabled = true;
                        Cursor.Current = Cursors.Default;
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
                Dock = DockStyle.Top,
                Height = 141,
                Width = sideBarWidth - 10,
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
                Font = font_Label,
            };

            comboBox_Signatory = new ComboBox
            {
                Parent = panel_Signatory,
                Width = sideBarWidth - 28,
                DropDownStyle = ComboBoxStyle.DropDownList,
                Font = font_Label,
            };

            UpdateSignatoryDropdownOptions();

            Label label_SignatoryName = new Label
            {
                Parent = panel_Signatory,
                Width = 48,
                Text = "Name:",
                TextAlign = ContentAlignment.MiddleLeft,
                Font = new Font("Microsoft Sans Serif", 8),
            };

            textBox_SignatoryName = new TextBox
            {
                Parent = panel_Signatory,
                Width = 165,
                Font = new Font("Microsoft Sans Serif", 8),
            };
            textBox_SignatoryName.Leave += (sender, e) =>
            {
                textBox_SignatoryName.Text = AccessToDatabase_INT.CapitalizeFirstLetterOfWords(textBox_SignatoryName.Text);
            };

            Label label_SignatoryPosition = new Label
            {
                Parent = panel_Signatory,
                Width = 48,
                Text = "Position:",
                TextAlign = ContentAlignment.MiddleLeft,
                Font = new Font("Microsoft Sans Serif", 8),
            };

            textBox_SignatoryPosition = new TextBox
            {
                Parent = panel_Signatory,
                Width = 165,
                Font = new Font("Microsoft Sans Serif", 8),
            };
            textBox_SignatoryPosition.Leave += (sender, e) =>
            {
                textBox_SignatoryPosition.Text = AccessToDatabase_INT.CapitalizeFirstLetterOfWords(textBox_SignatoryPosition.Text);
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

            label_SignatoryStatus = new Label
            {
                Parent = panel_Signatory,
                Height = 22,
                Width = 110,
                TextAlign = ContentAlignment.MiddleCenter,
                Font = new Font("Microsoft Sans Serif", 8),
                Margin = new Padding(0, 3, 0, 0),
            };

            button_SaveSignatory.Click += (sender, e) =>
            {
                if (comboBox_Signatory.SelectedIndex == 0)
                {
                    MessageBox.Show("Please select an option", "Notice", MessageBoxButtons.OK, MessageBoxIcon.Information);
                }
                else
                {
                    string currentFormType = GetCurrentFormTypeKey();

                    string signatoryName = AccessToDatabase_INT.CapitalizeFirstLetterOfWords(textBox_SignatoryName.Text.Trim());
                    string signatoryPosition = AccessToDatabase_INT.CapitalizeFirstLetterOfWords(textBox_SignatoryPosition.Text.Trim());

                    textBox_SignatoryName.Text = signatoryName;
                    textBox_SignatoryPosition.Text = signatoryPosition;

                    int choice = comboBox_Signatory.SelectedIndex;

                    accessToDatabase.SaveSignatoryData(currentFormType, choice, signatoryName, signatoryPosition);
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
                    string currentFormType = GetCurrentFormTypeKey();
                    int choice = comboBox_Signatory.SelectedIndex;

                    var signatoryData = accessToDatabase.RetrieveSignatoryData(currentFormType, choice);

                    textBox_SignatoryName.Text = signatoryData.Name;
                    textBox_SignatoryPosition.Text = signatoryData.Position;
                }
            };

            return panel_Signatory;
        }

        private void UpdateSignatoryDropdownOptions()
        {
            if (comboBox_Signatory == null) return;

            comboBox_Signatory.Items.Clear();

            if (GlobalVariables.client == "INT")
            {
                string thirdOption = (comboBox_Forms != null && comboBox_Forms.SelectedIndex == 3) ? "Approved By:" : "A/P:";

                comboBox_Signatory.Items.AddRange(new string[]
                {
                    "Select Signatory Option",
                    "Prepared By:",
                    "Checked By:",
                    thirdOption
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
                    "Noted By:"
                });
            }

            comboBox_Signatory.SelectedIndex = 0;
        }

        private FlowLayoutPanel Panel_SBRRSignatory()
        {
            FlowLayoutPanel panel_RRSignatory = new FlowLayoutPanel
            {
                Dock = DockStyle.Top,
                Height = 106,
                Width = sideBarWidth - 10,
                BackColor = Color.LightGray,
                Padding = new Padding(5, 2, 5, 0),
                BorderStyle = BorderStyle.FixedSingle,
            };

            Label panel_Title = new Label
            {
                Parent = panel_RRSignatory,
                Dock = DockStyle.Top,
                Text = "SIGNATORY (RR)",
                Width = sideBarWidth - 30,
                TextAlign = ContentAlignment.MiddleCenter,
            };

            Label label_ReceivedBy = new Label
            {
                Parent = panel_RRSignatory,
                Dock = DockStyle.Top,
                Text = "Received By:",
                TextAlign = ContentAlignment.MiddleLeft,
                Width = 71,
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
                TextAlign = ContentAlignment.MiddleCenter,
                Font = new Font("Microsoft Sans Serif", 8),
                Margin = new Padding(0, 3, 0, 0),
            };

            button_SaveRRSignatory.Click += (sender, e) =>
            {
                string signatoryName = textBox_ReceivedByRR.Text;
                string signatoryPosition = textBox_CheckedByRR.Text;

                accessToDatabase.SaveSignatoryRRData(signatoryName, signatoryPosition);
                label_SignatoryRRStatus.Text = "Saved";
            };

            return panel_RRSignatory;
        }

        private FlowLayoutPanel Panel_SBPrinting()
        {
            panel_Printing = new FlowLayoutPanel
            {
                Dock = DockStyle.Top,
                Height = 50,
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
                Visible = false,
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
                Visible = false,
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
                Visible = false,
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
                Visible = false,
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

                    if (comboBox_Forms.SelectedIndex == 4)
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

                        if (GlobalVariables.client == "LEADS")
                        {
                            string columnName = comboBox_Forms.SelectedIndex == 2 ? "CVSeries" : "APVSeries";
                            accessToDatabase.IncrementSeriesNumberInDatabase(columnName);
                            seriesNumber = accessToDatabase.GetSeriesNumberFromDatabase(columnName);
                            UpdateSeriesNumber(comboBox_Forms.SelectedIndex == 2 ? "CV" : "APV");
                        }
                        else if (GlobalVariables.client == "CPI")
                        {
                            string columnName = comboBox_Forms.SelectedIndex == 1 ? "CVSeries" : "APVSeries";
                            accessToDatabase.IncrementSeriesNumberInDatabase(columnName);
                            seriesNumber = accessToDatabase.GetSeriesNumberFromDatabase(columnName);
                            UpdateSeriesNumber(comboBox_Forms.SelectedIndex == 1 ? "CV" : "APV");
                        }
                        else if (GlobalVariables.client == "INT")
                        {
                            string formType = "";
                            if (comboBox_Forms.SelectedIndex == 1) formType = "CV";
                            else if (comboBox_Forms.SelectedIndex == 3) formType = "JV";
                            else if (comboBox_Forms.SelectedIndex == 4) formType = "APV";

                            if (formType != "")
                            {
                                string selectedCompany = comboBox_Company.SelectedItem?.ToString();
                                if (!string.IsNullOrEmpty(selectedCompany))
                                {
                                    seriesNumber++;
                                    accessToDatabase.UpdateManualSeriesNumber(formType, seriesNumber, selectedCompany);
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

        private string GetCurrentFormTypeKey()
        {
            switch (comboBox_Forms?.SelectedIndex)
            {
                case 1: return "CV";
                case 2: return "CHECK";
                case 3: return "JV";
                case 4: return "APV";
                default: return "DEFAULT";
            }
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

            UpdateSignatoryDropdownOptions();

            if (textBox_SignatoryName != null) textBox_SignatoryName.Text = "";
            if (textBox_SignatoryPosition != null) textBox_SignatoryPosition.Text = "";
            if (label_SignatoryStatus != null) label_SignatoryStatus.Text = "";

            if (GlobalVariables.client == "INT")
            {
                SetVoucherTypeVisibility(false);
                SetAPAccountVisibility(false);

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
                    case 1: // Check Voucher (CV)
                        prefix = "CV";
                        panel_SeriesNumber.Visible = false;
                        panel_RefNumber.Visible = false;
                        panel_RefNumberCrystalReport.Visible = true;
                        panel_Signatory.Visible = true;
                        label_SeriesNumberText.Text = "Current Series Number: CV";

                        if (label_CurrencyText != null) label_CurrencyText.Visible = true;
                        if (comboBox_Currency != null) comboBox_Currency.Visible = true;

                        SetAPAccountVisibility(true);

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

                    case 3: // Journal Voucher (JV)
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

                        SetVoucherTypeVisibility(true);
                        break;

                    case 4: // Accounts Payable Voucher (APV)
                        prefix = "APV";
                        panel_SeriesNumber.Visible = false;
                        panel_RefNumber.Visible = false;
                        panel_RefNumberCrystalReport.Visible = true;
                        panel_Signatory.Visible = true;

                        label_SeriesNumberText.Text = "Current Series Number: APV";

                        if (label_CurrencyText != null) label_CurrencyText.Visible = false;
                        if (comboBox_Currency != null) comboBox_Currency.Visible = false;

                        SetAPAccountVisibility(true);

                        panel_Main.Visible = false;
                        panel_Main_CR.Visible = true;
                        break;

                    case 5: // Online Voucher
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

        private void SetAPAccountVisibility(bool visible)
        {
            if (label_APAccount != null) label_APAccount.Visible = visible;
            if (comboBox_APAccount != null) comboBox_APAccount.Visible = visible;
        }

        private void SetVoucherTypeVisibility(bool visible)
        {
            if (label_VoucherType != null) label_VoucherType.Visible = visible;
            if (comboBox_VoucherType != null) comboBox_VoucherType.Visible = visible;
        }

        private void SetDatabaseLocation(ReportDocument reportDocument, string databasePath)
        {
            foreach (Table table in reportDocument.Database.Tables)
            {
                TableLogOnInfo tableLogOnInfo = table.LogOnInfo;
                tableLogOnInfo.ConnectionInfo.ServerName = databasePath;
                tableLogOnInfo.ConnectionInfo.DatabaseName = "";
                tableLogOnInfo.ConnectionInfo.UserID = "";
                tableLogOnInfo.ConnectionInfo.Password = "";
                table.ApplyLogOnInfo(tableLogOnInfo);
            }

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
            textBox_SeriesNumber.Text = $"{prefix}{seriesNumber:000}";
        }

        private void UpdateSeriesNumberINT(string formPrefix)
        {
            if (accessToDatabase == null) accessToDatabase = new AccessToDatabase_INT();
            textBox_SeriesNumber.Text = $"{seriesNumber:00000}";
        }

        private string SafeTruncate(string value, int maxLength)
        {
            if (string.IsNullOrEmpty(value)) return value;
            return value.Length > maxLength ? value.Substring(0, maxLength) : value;
        }

        private bool ValidateVoucherReferenceNumber(string voucherTitle, string refNumber, out string errorMessage)
        {
            errorMessage = "";
            string trimmedRef = refNumber.Trim();

            bool isESV = voucherTitle.StartsWith("EMPLOYEE", StringComparison.OrdinalIgnoreCase);

            if (isESV)
            {
                if (!Regex.IsMatch(trimmedRef, @"^\d{3}-\d{2}$"))
                {
                    errorMessage = "Invalid Reference Number for Employees Supplies Voucher.\nExpected format: 123-26 (xxx-xx)";
                    return false;
                }
            }
            else
            {
                bool isOldFormat = Regex.IsMatch(trimmedRef, @"^\d{2}\.\d{2,3}\.\d{2,3}$");
                bool isNewFormat = Regex.IsMatch(trimmedRef, @"^\d{6}$");

                if (!isOldFormat && !isNewFormat)
                {
                    errorMessage = "Invalid Reference Number for Journal Voucher.\nExpected format:\n- Dot format: 10.06.26 / 10.126.26 / 10.126.261 (xx.xx.xx, xx.xxx.xx, or xx.xxx.xxx)\n- New: 000001 (6 digits)";
                    return false;
                }
            }

            return true;
        }
    }
}