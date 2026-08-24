using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Drawing.Drawing2D;
using System.Drawing.Printing;
using System.IO;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;
using System.Data.OleDb;
using CrystalDecisions.Shared;
using CrystalDecisions.CrystalReports.Engine;
using CrystalDecisions.Windows.Forms;
using static VoucherPROVER2.Clients.INT.Dataclass_INT;
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

        // UI Controls
        private FlowLayoutPanel panel_Company;
        private ComboBox comboBox_Forms;
        private ComboBox comboBox_Company;

        private Label label_SeriesNumberText;
        private Label label_SignatoryRRStatus;
        private Label label_VoucherType;
        private ComboBox comboBox_VoucherType;

        private Label label_APAccount;
        private ComboBox comboBox_APAccount;

        private TextBox textBox_SeriesNumber;
        private TextBox textBox_ReceivedByRR;
        private TextBox textBox_CheckedByRR;

        private FlowLayoutPanel panel_PayeeOverride;
        private TextBox textBox_PayeeOverride;

        private Panel panel_Main;
        private Panel panel_Main_CR;

        private FlowLayoutPanel panel_Printing;
        private FlowLayoutPanel panel_SeriesNumber;
        private FlowLayoutPanel panel_Signatory;
        private FlowLayoutPanel panel_RRSignatory;
        private FlowLayoutPanel panel_RefNumber;
        private FlowLayoutPanel panel_RefNumberCrystalReport;

        // Data caches
        private List<CheckTable> cheque = new List<CheckTable>();
        private List<CheckTableGrid> checkivp = new List<CheckTableGrid>();
        private List<BillTable> bills = new List<BillTable>();
        private List<CheckTableExpensesAndItems> checks = new List<CheckTableExpensesAndItems>();
        private List<ItemReciept> receipts = new List<ItemReciept>();
        private List<BillTable> apvData = new List<BillTable>();
        private List<CheckTableExpensesAndItems> cvData = new List<CheckTableExpensesAndItems>();
        private List<JournalGridItem> journal = new List<JournalGridItem>();

        private const int sideBarWidth = 270;
        private int seriesNumber = 1;
        private int itemCounter;
        private int pageCounter;

        // =========================================================================
        // DYNAMIC THEME SYSTEM (DARK / LIGHT MODE)
        // =========================================================================
        private bool isDarkMode = true;
        private FlowLayoutPanel panel_SideBar;
        private Button button_ThemeToggle;

        // Dark Palette
        private static readonly Color DarkSidebarBg      = Color.FromArgb(15, 23, 42);    // Midnight Navy (#0F172A)
        private static readonly Color DarkCardBg         = Color.FromArgb(30, 41, 59);    // Slate Card (#1E293B)
        private static readonly Color DarkCardBorder     = Color.FromArgb(51, 65, 85);    // Slate Border (#334155)
        private static readonly Color DarkTextPrimary    = Color.FromArgb(248, 250, 252); // White Text (#F8FAFC)
        private static readonly Color DarkTextMuted      = Color.FromArgb(148, 163, 184); // Muted Slate (#94A3B8)
        private static readonly Color DarkInputBg        = Color.White;
        private static readonly Color DarkInputText      = Color.FromArgb(15, 23, 42);
        private static readonly Color DarkSecondaryBtn   = Color.FromArgb(51, 65, 85);    // Slate Button (#334155)
        private static readonly Color DarkSecondaryHover = Color.FromArgb(71, 85, 105);   // Slate Hover (#475569)
        private static readonly Color DarkSecondaryText  = Color.FromArgb(248, 250, 252); // Button Text

        // Light Palette
        private static readonly Color LightSidebarBg     = Color.FromArgb(241, 245, 249); // Soft Slate 100 (#F1F5F9)
        private static readonly Color LightCardBg        = Color.FromArgb(255, 255, 255); // Pure White (#FFFFFF)
        private static readonly Color LightCardBorder    = Color.FromArgb(203, 213, 225); // Slate 300 (#CBD5E1)
        private static readonly Color LightTextPrimary   = Color.FromArgb(15, 23, 42);    // Slate 900 (#0F172A)
        private static readonly Color LightTextMuted     = Color.FromArgb(71, 85, 105);   // Slate 600 (#475569)
        private static readonly Color LightInputBg       = Color.White;
        private static readonly Color LightInputText     = Color.FromArgb(15, 23, 42);
        private static readonly Color LightSecondaryBtn  = Color.FromArgb(226, 232, 240); // Slate 200 (#E2E8F0)
        private static readonly Color LightSecondaryHover= Color.FromArgb(203, 213, 225); // Slate 300 (#CBD5E1)
        private static readonly Color LightSecondaryText = Color.FromArgb(30, 41, 59);    // Slate 800 (#1E293B)

        // Accent Colors (Shared across themes)
        private static readonly Color ColorPrimaryBtn    = Color.FromArgb(37, 99, 235);   // Vibrant Blue (#2563EB)
        private static readonly Color ColorPrimaryHover  = Color.FromArgb(29, 78, 216);   // Hover Blue (#1D4ED8)
        private static readonly Color ColorSuccessBtn    = Color.FromArgb(16, 185, 129);  // Emerald (#10B981)
        private static readonly Color ColorSuccessHover  = Color.FromArgb(5, 150, 105);   // Emerald Hover (#059669)

        // Current Active Theme Properties
        private Color CurrentSidebarBg      => isDarkMode ? DarkSidebarBg : LightSidebarBg;
        private Color CurrentCardBg         => isDarkMode ? DarkCardBg : LightCardBg;
        private Color CurrentCardBorder     => isDarkMode ? DarkCardBorder : LightCardBorder;
        private Color CurrentTextPrimary    => isDarkMode ? DarkTextPrimary : LightTextPrimary;
        private Color CurrentTextMuted      => isDarkMode ? DarkTextMuted : LightTextMuted;
        private Color CurrentInputBg        => isDarkMode ? DarkInputBg : LightInputBg;
        private Color CurrentInputText      => isDarkMode ? DarkInputText : LightInputText;
        private Color CurrentSecondaryBtn   => isDarkMode ? DarkSecondaryBtn : LightSecondaryBtn;
        private Color CurrentSecondaryHover => isDarkMode ? DarkSecondaryHover : LightSecondaryHover;
        private Color CurrentSecondaryText  => isDarkMode ? DarkSecondaryText : LightSecondaryText;
        private Color CurrentMainBg         => isDarkMode ? Color.FromArgb(15, 23, 42) : Color.FromArgb(241, 245, 249);

        // Aliases for unified theme access
        private Color ColorTextMuted      => CurrentTextMuted;
        private Color ColorTextPrimary    => CurrentTextPrimary;
        private Color ColorMainBg         => CurrentMainBg;
        private Color ColorTitleBg        => isDarkMode ? Color.FromArgb(15, 23, 42) : Color.FromArgb(241, 245, 249);
        private Color ColorCardBorder     => CurrentCardBorder;
        private Color ColorCardBg         => CurrentCardBg;
        private Color ColorSidebarBg      => CurrentSidebarBg;
        private Color ColorInputBg        => CurrentInputBg;
        private Color ColorInputText      => CurrentInputText;
        private Color ColorSecondaryBtn   => CurrentSecondaryBtn;
        private Color ColorSecondaryHover => CurrentSecondaryHover;

        private static readonly Font FontCardHeader = new Font("Segoe UI", 8.5f, FontStyle.Bold);
        private static readonly Font FontInputLabel = new Font("Segoe UI", 8.5f, FontStyle.Regular);
        private static readonly Font FontInput      = new Font("Segoe UI", 9.5f, FontStyle.Regular);
        private static readonly Font FontButton     = new Font("Segoe UI", 9f, FontStyle.Bold);

        // =========================================================================
        // MODERN CONTROL FACTORY HELPERS
        // =========================================================================
        private FlowLayoutPanel CreateCardPanel(string headerTitle, int width)
        {
            var card = new FlowLayoutPanel
            {
                Width = width,
                AutoSize = true,
                AutoSizeMode = AutoSizeMode.GrowAndShrink,
                FlowDirection = FlowDirection.TopDown,
                WrapContents = false,
                BackColor = CurrentCardBg,
                Padding = new Padding(12, 10, 12, 12),
                Margin = new Padding(0, 0, 0, 10),
                Tag = "Card"
            };

            card.Paint += (s, e) =>
            {
                using (Pen borderPen = new Pen(CurrentCardBorder, 1))
                {
                    e.Graphics.DrawRectangle(borderPen, 0, 0, card.Width - 1, card.Height - 1);
                }
            };

            if (!string.IsNullOrEmpty(headerTitle))
            {
                var lblHeader = new Label
                {
                    Text = headerTitle,
                    Font = FontCardHeader,
                    ForeColor = CurrentTextMuted,
                    Width = width - 24,
                    Height = 20,
                    TextAlign = ContentAlignment.MiddleLeft,
                    Margin = new Padding(0, 0, 0, 6),
                    Tag = "Header"
                };
                card.Controls.Add(lblHeader);
            }

            return card;
        }

        private Button CreateModernButton(string text, Color bg, Color hoverBg, Color textColor, int width, int height, Font font = null, string tag = null)
        {
            Button btn = new Button
            {
                Text = text,
                Width = width,
                Height = height,
                BackColor = bg,
                ForeColor = textColor,
                Font = font ?? FontButton,
                FlatStyle = FlatStyle.Flat,
                Cursor = Cursors.Hand,
                TextAlign = ContentAlignment.MiddleCenter,
                Margin = new Padding(0, 3, 0, 3),
                Tag = tag
            };
            btn.FlatAppearance.BorderSize = 0;
            btn.FlatAppearance.MouseOverBackColor = hoverBg;
            btn.FlatAppearance.MouseDownBackColor = ControlPaint.Dark(hoverBg, 0.1f);
            return btn;
        }

        private TextBox CreateModernTextBox(int width)
        {
            return new TextBox
            {
                Width = width,
                Font = FontInput,
                BackColor = CurrentInputBg,
                ForeColor = CurrentInputText,
                BorderStyle = BorderStyle.FixedSingle,
                Margin = new Padding(0, 2, 0, 6),
                Tag = "Input"
            };
        }

        private ComboBox CreateModernComboBox(int width)
        {
            return new ComboBox
            {
                Width = width,
                Font = FontInput,
                DropDownStyle = ComboBoxStyle.DropDownList,
                BackColor = CurrentInputBg,
                ForeColor = CurrentInputText,
                FlatStyle = FlatStyle.Flat,
                Margin = new Padding(0, 2, 0, 6),
                Tag = "Input"
            };
        }

        private void ApplyTheme()
        {
            if (panel_SideBar != null)
            {
                panel_SideBar.BackColor = CurrentSidebarBg;
                panel_SideBar.Invalidate();
            }

            if (button_ThemeToggle != null)
            {
                button_ThemeToggle.Text = isDarkMode ? "🌙  DARK MODE  (Click for Light ☀️)" : "☀️  LIGHT MODE  (Click for Dark 🌙)";
                button_ThemeToggle.BackColor = CurrentSecondaryBtn;
                button_ThemeToggle.ForeColor = CurrentSecondaryText;
                button_ThemeToggle.FlatAppearance.MouseOverBackColor = CurrentSecondaryHover;
            }

            if (panel_Main != null) panel_Main.BackColor = CurrentMainBg;
            if (panel_Main_CR != null) panel_Main_CR.BackColor = CurrentMainBg;
            if (printPreviewControl != null) printPreviewControl.BackColor = CurrentMainBg;

            UpdateControlTreeTheme(panel_SideBar);
        }

        private void UpdateControlTreeTheme(Control parent)
        {
            if (parent == null) return;

            foreach (Control c in parent.Controls)
            {
                if (c is FlowLayoutPanel flp && (string)flp.Tag == "Card")
                {
                    flp.BackColor = CurrentCardBg;
                    flp.Invalidate();
                }
                else if (c is Label lbl)
                {
                    if ((string)lbl.Tag == "Header" || (string)lbl.Tag == "Muted")
                    {
                        lbl.ForeColor = CurrentTextMuted;
                    }
                    else if ((string)lbl.Tag == "StatusSuccess")
                    {
                        lbl.ForeColor = ColorSuccessBtn;
                    }
                    else
                    {
                        lbl.ForeColor = CurrentTextPrimary;
                    }
                }
                else if (c is TextBox txt)
                {
                    txt.BackColor = CurrentInputBg;
                    txt.ForeColor = CurrentInputText;
                }
                else if (c is ComboBox cmb)
                {
                    cmb.BackColor = CurrentInputBg;
                    cmb.ForeColor = CurrentInputText;
                }
                else if (c is Button btn && btn != button_ThemeToggle)
                {
                    if ((string)btn.Tag == "Secondary")
                    {
                        btn.BackColor = CurrentSecondaryBtn;
                        btn.ForeColor = CurrentSecondaryText;
                        btn.FlatAppearance.MouseOverBackColor = CurrentSecondaryHover;
                    }
                }

                if (c.HasChildren)
                {
                    UpdateControlTreeTheme(c);
                }
            }
        }

        // =========================================================================
        // SIDEBAR CARD BUILDERS
        // =========================================================================
        public FlowLayoutPanel Panel_SBThemeToggle()
        {
            int cardInnerWidth = sideBarWidth - 44;
            FlowLayoutPanel panel_Theme = CreateCardPanel("🎨  THEME PREFERENCE", sideBarWidth - 24);

            button_ThemeToggle = CreateModernButton(
                isDarkMode ? "🌙  DARK MODE  (Click for Light ☀️)" : "☀️  LIGHT MODE  (Click for Dark 🌙)",
                CurrentSecondaryBtn,
                CurrentSecondaryHover,
                CurrentSecondaryText,
                cardInnerWidth,
                32,
                new Font("Segoe UI", 9f, FontStyle.Bold),
                "ThemeToggle"
            );
            button_ThemeToggle.Click += (sender, e) =>
            {
                isDarkMode = !isDarkMode;
                ApplyTheme();
            };

            panel_Theme.Controls.Add(button_ThemeToggle);
            return panel_Theme;
        }

        public FlowLayoutPanel Panel_SBPayeeOverride()
        {
            int cardInnerWidth = sideBarWidth - 44;
            panel_PayeeOverride = CreateCardPanel("👤  PAYEE OVERRIDE", sideBarWidth - 24);
            panel_PayeeOverride.Visible = false;

            Label label_Text = new Label
            {
                Text = "Override Payee Name:",
                Font = FontInputLabel,
                ForeColor = ColorTextMuted,
                Width = cardInnerWidth,
                Height = 18,
                TextAlign = ContentAlignment.MiddleLeft,
                Margin = new Padding(0, 0, 0, 3)
            };
            panel_PayeeOverride.Controls.Add(label_Text);

            textBox_PayeeOverride = CreateModernTextBox(cardInnerWidth);
            panel_PayeeOverride.Controls.Add(textBox_PayeeOverride);

            return panel_PayeeOverride;
        }

        public FlowLayoutPanel Panel_SBCompany()
        {
            int cardInnerWidth = sideBarWidth - 44;
            panel_Company = CreateCardPanel("🏢  COMPANY & ACCOUNTS", sideBarWidth - 24);
            panel_Company.Visible = (GlobalVariables.client == "INT");

            Label label_CompanyText = new Label
            {
                Text = "Company:",
                Font = FontInputLabel,
                ForeColor = ColorTextMuted,
                Width = cardInnerWidth,
                Height = 18,
                TextAlign = ContentAlignment.MiddleLeft,
                Margin = new Padding(0, 0, 0, 3)
            };
            panel_Company.Controls.Add(label_CompanyText);

            comboBox_Company = CreateModernComboBox(cardInnerWidth);
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
                else if (comboBox_Forms.SelectedIndex == 4) formType = "APV";

                if (formType != "")
                {
                    string selectedCompany = comboBox_Company.SelectedItem?.ToString() ?? "";
                    seriesNumber = accessToDatabase.GetSeriesNumberFromDatabase(formType, selectedCompany);
                    UpdateSeriesNumberINT(formType);
                }
            };
            panel_Company.Controls.Add(comboBox_Company);

            // AP Account Dropdown
            label_APAccount = new Label
            {
                Text = "💳  A/P Account:",
                Font = FontInputLabel,
                ForeColor = ColorTextMuted,
                Width = cardInnerWidth,
                Height = 18,
                TextAlign = ContentAlignment.MiddleLeft,
                Margin = new Padding(0, 6, 0, 3),
                Visible = false
            };
            panel_Company.Controls.Add(label_APAccount);

            comboBox_APAccount = CreateModernComboBox(cardInnerWidth);
            comboBox_APAccount.Items.AddRange(new string[]
            {
                "Vouchers Payable",
                "Notes Payable"
            });
            comboBox_APAccount.SelectedIndex = 0;
            comboBox_APAccount.Visible = false;
            panel_Company.Controls.Add(comboBox_APAccount);

            // Voucher Title Dropdown (Journal Voucher)
            label_VoucherType = new Label
            {
                Text = "📝  Voucher Title:",
                Font = FontInputLabel,
                ForeColor = ColorTextMuted,
                Width = cardInnerWidth,
                Height = 18,
                TextAlign = ContentAlignment.MiddleLeft,
                Margin = new Padding(0, 6, 0, 3),
                Visible = false
            };
            panel_Company.Controls.Add(label_VoucherType);

            comboBox_VoucherType = CreateModernComboBox(cardInnerWidth);
            comboBox_VoucherType.Items.AddRange(new string[]
            {
                "JOURNAL VOUCHER",
                "EMPLOYEES SUPPLIES VOUCHER"
            });
            comboBox_VoucherType.SelectedIndex = 0;
            comboBox_VoucherType.Visible = false;
            panel_Company.Controls.Add(comboBox_VoucherType);

            return panel_Company;
        }

        public Panel ContainerPanel()
        {
            Panel panel_Container = new Panel
            {
                Dock = DockStyle.Fill,
                BackColor = ColorMainBg
            };

            Panel panel_Title = TitlePanel();
            panel_Main = MainPanel();
            panel_Main_CR = MainPanel_CR();
            Panel panel_SideBar = SideBarPanel();

            panel_Container.Controls.Add(panel_Main);
            panel_Container.Controls.Add(panel_Main_CR);
            panel_Container.Controls.Add(panel_SideBar);
            panel_Container.Controls.Add(panel_Title);

            return panel_Container;
        }

        public Panel TitlePanel()
        {
            Panel panel_Title = new Panel
            {
                Dock = DockStyle.Top,
                Height = 54,
                BackColor = ColorTitleBg,
                Padding = new Padding(16, 0, 20, 0)
            };

            panel_Title.Paint += (s, e) =>
            {
                using (Pen p = new Pen(ColorCardBorder, 1))
                {
                    e.Graphics.DrawLine(p, 0, panel_Title.Height - 1, panel_Title.Width, panel_Title.Height - 1);
                }
            };

            Label labelLogo = new Label
            {
                Parent = panel_Title,
                Text = "VOUCHERPRO",
                Font = new Font("Segoe UI", 13f, FontStyle.Bold),
                ForeColor = Color.White,
                AutoSize = true,
                Location = new Point(16, 14)
            };

            Label labelTag = new Label
            {
                Parent = panel_Title,
                Text = "INTEGRATED",
                Font = new Font("Segoe UI", 8f, FontStyle.Bold),
                ForeColor = Color.FromArgb(59, 130, 246),
                BackColor = Color.FromArgb(30, 41, 59),
                Padding = new Padding(6, 2, 6, 2),
                AutoSize = true,
                Location = new Point(160, 16)
            };

            Label labelCompanyHeader = new Label
            {
                Parent = panel_Title,
                Text = "INTEGRATED CONTRACTOR & PLUMBING WORKS, INC.",
                Font = new Font("Segoe UI", 9f, FontStyle.Regular),
                ForeColor = Color.FromArgb(148, 163, 184),
                Dock = DockStyle.Right,
                TextAlign = ContentAlignment.MiddleRight,
                AutoSize = false,
                Width = 460
            };

            return panel_Title;
        }

        public Panel MainPanel()
        {
            Panel panel_Main_Local = new Panel
            {
                BackColor = ColorMainBg,
                Dock = DockStyle.Fill
            };

            printPreviewControl = new PrintPreviewControl
            {
                Parent = panel_Main_Local,
                Dock = DockStyle.Fill,
                Zoom = 1,
                Visible = false,
                BackColor = Color.FromArgb(226, 232, 240)
            };

            return panel_Main_Local;
        }

        public Panel MainPanel_CR()
        {
            Panel panel_Main_CR_Local = new Panel
            {
                BackColor = ColorMainBg,
                Dock = DockStyle.Fill
            };

            reportViewer = new CrystalReportViewer
            {
                Parent = panel_Main_CR_Local,
                Dock = DockStyle.Fill,
                ShowCopyButton = false,
                ShowPrintButton = true,
                ShowExportButton = false,
                ShowRefreshButton = false,
                ShowGroupTreeButton = false,
                ShowTextSearchButton = false,
                ShowParameterPanelButton = false,
                ToolPanelView = ToolPanelViewType.None,
                BorderStyle = BorderStyle.None
            };

            foreach (Control control in reportViewer.Controls)
            {
                if (control is ToolStrip toolStrip)
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

            return panel_Main_CR_Local;
        }

        private Panel SideBarPanel()
        {
            panel_SideBar = new FlowLayoutPanel
            {
                Dock = DockStyle.Left,
                Width = sideBarWidth,
                AutoScroll = true,
                WrapContents = false,
                FlowDirection = FlowDirection.TopDown,
                BackColor = CurrentSidebarBg,
                Padding = new Padding(12, 12, 12, 20)
            };

            panel_SideBar.Paint += (s, e) =>
            {
                using (Pen p = new Pen(CurrentCardBorder, 1))
                {
                    e.Graphics.DrawLine(p, panel_SideBar.Width - 1, 0, panel_SideBar.Width - 1, panel_SideBar.Height);
                }
            };

            // 0. THEME TOGGLE CARD
            FlowLayoutPanel panel_Theme = Panel_SBThemeToggle();
            panel_SideBar.Controls.Add(panel_Theme);

            // 1. FORMS CARD
            FlowLayoutPanel panels_Forms = Panel_SBForms();
            panel_SideBar.Controls.Add(panels_Forms);

            // 2. COMPANY & ACCOUNTS CARD
            if (GlobalVariables.client == "INT")
            {
                panel_Company = Panel_SBCompany();
                panel_SideBar.Controls.Add(panel_Company);

                panel_PayeeOverride = Panel_SBPayeeOverride();
                panel_SideBar.Controls.Add(panel_PayeeOverride);
            }

            // 3. SERIES NUMBER CARD
            panel_SeriesNumber = Panel_SBSeriesNumber();
            panel_SideBar.Controls.Add(panel_SeriesNumber);
            panel_SeriesNumber.Visible = false;

            // 4. REF NUMBER CARDS
            panel_RefNumber = Panel_SBRefNumber();
            panel_RefNumberCrystalReport = Panel_SBRefNumber_CR();
            panel_SideBar.Controls.Add(panel_RefNumber);
            panel_SideBar.Controls.Add(panel_RefNumberCrystalReport);
            panel_RefNumber.Visible = false;
            panel_RefNumberCrystalReport.Visible = false;

            // 5. SIGNATORY CARD
            panel_Signatory = Panel_SBSignatory();
            panel_SideBar.Controls.Add(panel_Signatory);
            panel_Signatory.Visible = false;

            // 6. RR SIGNATORY (LEADS ONLY)
            if (GlobalVariables.client == "LEADS")
            {
                panel_RRSignatory = Panel_SBRRSignatory();
                panel_SideBar.Controls.Add(panel_RRSignatory);
                panel_RRSignatory.Visible = false;
            }

            // 7. PRINTING CONTROLS CARD
            panel_Printing = Panel_SBPrinting();
            panel_SideBar.Controls.Add(panel_Printing);
            panel_Printing.Visible = false;

            return panel_SideBar;
        }

        public FlowLayoutPanel Panel_SBForms()
        {
            int cardInnerWidth = sideBarWidth - 44;
            FlowLayoutPanel panel_Forms = CreateCardPanel("📄  FORM SELECTION", sideBarWidth - 24);

            Label label_FormText = new Label
            {
                Text = "Select Form Type:",
                Font = FontInputLabel,
                ForeColor = ColorTextMuted,
                Width = cardInnerWidth,
                Height = 18,
                TextAlign = ContentAlignment.MiddleLeft,
                Margin = new Padding(0, 0, 0, 3)
            };
            panel_Forms.Controls.Add(label_FormText);

            comboBox_Forms = CreateModernComboBox(cardInnerWidth);
            if (GlobalVariables.client == "INT")
            {
                comboBox_Forms.Items.AddRange(new string[]
                {
                    "-- Select Form --",
                    "Check Voucher / Bills Payment",
                    "Check",
                    "Journal Voucher",
                    "Check Voucher / Enter Bills"
                });
                comboBox_Forms.SelectedIndex = 0;
                comboBox_Forms.SelectedIndexChanged += ComboBox_Forms_SelectedIndexChanged;
            }
            panel_Forms.Controls.Add(comboBox_Forms);

            return panel_Forms;
        }

        public FlowLayoutPanel Panel_SBSeriesNumber()
        {
            int cardInnerWidth = sideBarWidth - 44;
            panel_SeriesNumber = CreateCardPanel("🔢  SERIES NUMBER", sideBarWidth - 24);
            panel_SeriesNumber.Visible = false;

            label_SeriesNumberText = new Label
            {
                Text = "Current Series Number:",
                Font = FontInputLabel,
                ForeColor = ColorTextMuted,
                Width = cardInnerWidth,
                Height = 18,
                TextAlign = ContentAlignment.MiddleLeft,
                Margin = new Padding(0, 0, 0, 4)
            };
            panel_SeriesNumber.Controls.Add(label_SeriesNumberText);

            FlowLayoutPanel rowPanel = new FlowLayoutPanel
            {
                Width = cardInnerWidth,
                Height = 32,
                FlowDirection = FlowDirection.LeftToRight,
                WrapContents = false,
                Margin = new Padding(0),
                Padding = new Padding(0),
                BackColor = Color.Transparent
            };

            textBox_SeriesNumber = new TextBox
            {
                Width = cardInnerWidth - 72,
                Height = 28,
                Font = new Font("Segoe UI", 10.5f, FontStyle.Bold),
                BackColor = ColorInputBg,
                ForeColor = ColorInputText,
                BorderStyle = BorderStyle.FixedSingle,
                TextAlign = HorizontalAlignment.Center,
                Margin = new Padding(0, 0, 4, 0)
            };
            rowPanel.Controls.Add(textBox_SeriesNumber);

            Button button_Decrement = CreateModernButton("-", CurrentSecondaryBtn, CurrentSecondaryHover, CurrentSecondaryText, 32, 28, new Font("Segoe UI", 11f, FontStyle.Bold), "Secondary");
            button_Decrement.Margin = new Padding(0, 0, 4, 0);
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
            rowPanel.Controls.Add(button_Decrement);

            Button button_Increment = CreateModernButton("+", CurrentSecondaryBtn, CurrentSecondaryHover, CurrentSecondaryText, 32, 28, new Font("Segoe UI", 11f, FontStyle.Bold), "Secondary");
            button_Increment.Margin = new Padding(0);
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
            rowPanel.Controls.Add(button_Increment);

            panel_SeriesNumber.Controls.Add(rowPanel);
            return panel_SeriesNumber;
        }

        private static T GetReportObjectSafe<T>(ReportDocument reportDoc, string objectName) where T : ReportObject
        {
            if (reportDoc == null || string.IsNullOrEmpty(objectName)) return null;
            try
            {
                return reportDoc.ReportDefinition.ReportObjects[objectName] as T;
            }
            catch
            {
                return null;
            }
        }

        public FlowLayoutPanel Panel_SBRefNumber_CR()
        {
            int cardInnerWidth = sideBarWidth - 44;
            panel_RefNumberCrystalReport = CreateCardPanel("🔍  REFERENCE NUMBER", sideBarWidth - 24);
            panel_RefNumberCrystalReport.Visible = false;

            Label label_RefNumberText = new Label
            {
                Text = "Enter Reference Number:",
                Font = FontInputLabel,
                ForeColor = ColorTextMuted,
                Width = cardInnerWidth,
                Height = 18,
                TextAlign = ContentAlignment.MiddleLeft,
                Margin = new Padding(0, 0, 0, 3)
            };
            panel_RefNumberCrystalReport.Controls.Add(label_RefNumberText);

            TextBox textBox_ReferenceNumber_CR = CreateModernTextBox(cardInnerWidth);
            panel_RefNumberCrystalReport.Controls.Add(textBox_ReferenceNumber_CR);

            Button button_SearchRefNum_CR = CreateModernButton("🔍  SEARCH & LOAD", ColorPrimaryBtn, ColorPrimaryHover, Color.White, cardInnerWidth, 32);
            button_SearchRefNum_CR.Margin = new Padding(0, 2, 0, 4);

            Action performSearch = () =>
            {
                if (comboBox_Forms.SelectedIndex == 0)
                {
                    MessageBox.Show("Please select a form type first.", "Notice", MessageBoxButtons.OK, MessageBoxIcon.Information);
                    return;
                }

                string refNumberCR = textBox_ReferenceNumber_CR.Text.Trim();
                if (string.IsNullOrEmpty(refNumberCR))
                {
                    MessageBox.Show("Please enter a reference number.", "Notice", MessageBoxButtons.OK, MessageBoxIcon.Warning);
                    return;
                }

                Cursor.Current = Cursors.WaitCursor;
                button_SearchRefNum_CR.Enabled = false;
                button_SearchRefNum_CR.Text = "⏳ SEARCHING...";

                try
                {
                    if (GlobalVariables.client == "INT")
                    {
                        if (comboBox_Forms.SelectedIndex == 1) // Check Voucher
                        {
                            bool cvDataExists = false;
                            try
                            {
                                CRCV_INT cRCV_INT = new CRCV_INT();
                                string databasePath = Path.Combine(Application.StartupPath, "CheckDatabase.accdb");
                                SetDatabaseLocation(cRCV_INT, databasePath);

                                AccessQueries_INT accessQueries = new AccessQueries_INT();
                                cvData = accessQueries.GetCheckExpensesAndItemsData_INT(refNumberCR);

                                if (cvData != null && cvData.Count > 0)
                                {
                                    cvDataExists = true;

                                    if (GetReportObjectSafe<TextObject>(cRCV_INT, "TextCVPayee") is TextObject textObject_CVPayee)
                                        textObject_CVPayee.Text = cvData[0].PayeeFullName;
                                    if (GetReportObjectSafe<TextObject>(cRCV_INT, "TextCVDatenow") is TextObject textObject_CVDatenow)
                                        textObject_CVDatenow.Text = DateTime.Now.ToString("MMMM dd, yyyy");

                                    var b = cvData[0];
                                    string streetLine = string.Join(", ", new[] { b.AddressBlockAddr1, b.AddressBlockAddr2, b.AddressBlockAddr3, b.AddressBlockAddr4 }.Where(s => !string.IsNullOrWhiteSpace(s)));
                                    string cityLine = string.Join(" ", new[] { b.AddressCity }.Where(s => !string.IsNullOrWhiteSpace(s)));
                                    string fullAddress = string.Join(Environment.NewLine, new[] { streetLine, cityLine }.Where(s => !string.IsNullOrWhiteSpace(s)));

                                    if (GetReportObjectSafe<TextObject>(cRCV_INT, "TextCVAddress") is TextObject textObject_CVAddress)
                                        textObject_CVAddress.Text = fullAddress;

                                    AccessToDatabase_INT accessToDatabase = new AccessToDatabase_INT();
                                    var signatories = accessToDatabase.RetrieveAllSignatoryData();

                                    if (GetReportObjectSafe<TextObject>(cRCV_INT, "TextPreparedBy") is TextObject textObject_PreparedBy)
                                        textObject_PreparedBy.Text = signatories.PreparedByName;
                                    if (GetReportObjectSafe<TextObject>(cRCV_INT, "TextPreparedByPosition") is TextObject textObject_PreparedByPos)
                                        textObject_PreparedByPos.Text = signatories.PreparedByPosition;
                                    if (GetReportObjectSafe<TextObject>(cRCV_INT, "TextCheckedBy") is TextObject textObject_CheckedBy)
                                        textObject_CheckedBy.Text = signatories.ReviewedByName;
                                    if (GetReportObjectSafe<TextObject>(cRCV_INT, "TextCheckedByPosition") is TextObject textObject_CheckedByPos)
                                        textObject_CheckedByPos.Text = signatories.ReviewedByPosition;
                                    if (GetReportObjectSafe<TextObject>(cRCV_INT, "TextApprovedBy") is TextObject textObject_ApprovedBy)
                                        textObject_ApprovedBy.Text = signatories.ApprovedByName;
                                    if (GetReportObjectSafe<TextObject>(cRCV_INT, "TextApprovedByPosition") is TextObject textObject_ApprovedByPos)
                                        textObject_ApprovedByPos.Text = signatories.ApprovedByPosition;

                                    double amount = cvData[0].TotalAmount;
                                    string amountInWords = AccessToDatabase_INT.AmountToWordsConverter.Convert(amount);
                                    string rawBank = cvData[0].BankAccount ?? "";
                                    string bank = rawBank.Contains(":") ? rawBank.Split(':').Last().Trim() : rawBank;

                                    if (GetReportObjectSafe<TextObject>(cRCV_INT, "TextCVRefNumber") is TextObject textObject_CVRefNumber)
                                        textObject_CVRefNumber.Text = refNumberCR;
                                    if (GetReportObjectSafe<TextObject>(cRCV_INT, "TextCVCheckDate") is TextObject textObject_CVCheckDate)
                                        textObject_CVCheckDate.Text = cvData[0].DueDate.ToString("MMMM dd, yyyy");
                                    if ((GetReportObjectSafe<TextObject>(cRCV_INT, "TextCVAmount") ?? GetReportObjectSafe<TextObject>(cRCV_INT, "TextCVTotal")) is TextObject textObject_CVTotal)
                                        textObject_CVTotal.Text = cvData[0].TotalAmount.ToString("N2");
                                    if (GetReportObjectSafe<TextObject>(cRCV_INT, "TextCVAmountInWords") is TextObject textObject_CVAmountinWords)
                                        textObject_CVAmountinWords.Text = "          " + amountInWords;
                                    if (GetReportObjectSafe<TextObject>(cRCV_INT, "TextCVBankAccount") is TextObject textObject_CVBankAccount)
                                        textObject_CVBankAccount.Text = bank;

                                    if (GetReportObjectSafe<SubreportObject>(cRCV_INT, "SubreportCVDetailsIVP") is SubreportObject subreportObject)
                                    {
                                        ReportDocument subReportDocument = cRCV_INT.OpenSubreport(subreportObject.SubreportName);
                                        InsertDataToCheckVoucherCompiledINT(refNumberCR, cvData);
                                    }
                                    if (GetReportObjectSafe<SubreportObject>(cRCV_INT, "SubreportCVDetailsINTCredit") is SubreportObject subreportObjectcredit)
                                    {
                                        ReportDocument subReportDocumentcredit = cRCV_INT.OpenSubreport(subreportObjectcredit.SubreportName);
                                        InsertDataToCheckVoucherCompiledINT(refNumberCR, cvData);
                                    }
                                    if (GetReportObjectSafe<SubreportObject>(cRCV_INT, "SubreportCVDetailsINT") is SubreportObject subreportObject2)
                                    {
                                        ReportDocument subReportDocument2 = cRCV_INT.OpenSubreport(subreportObject2.SubreportName);
                                        if (GetReportObjectSafe<TextObject>(subReportDocument2, "TextRemarks") is TextObject textObject_Remarks)
                                            textObject_Remarks.Text = SafeTruncate(cvData[0].Memo, 500);
                                        if (GetReportObjectSafe<TextObject>(subReportDocument2, "TextCVSubTotalAmount") is TextObject textObject_CVSubTotal)
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
                                MessageBox.Show($"INT CV ERROR:\n{ex.Message}", "Error", MessageBoxButtons.OK, MessageBoxIcon.Error);
                            }

                            if (!cvDataExists)
                            {
                                GenerateBillPaymentReport_INT(refNumberCR);
                            }
                        }
                        else if (comboBox_Forms.SelectedIndex == 3) // Journal Voucher (JV)
                        {
                            CRJV_INT cRJV_INT = new CRJV_INT();
                            string databasePath = Path.Combine(Application.StartupPath, "CheckDatabase.accdb");
                            SetDatabaseLocation(cRJV_INT, databasePath);

                            AccessQueries_INT accessQueries = new AccessQueries_INT();
                            journal = accessQueries.GetJournalEntryForGrid(refNumberCR);

                            if (journal != null && journal.Count > 0)
                            {
                                string selectedVoucherTitle = comboBox_VoucherType?.SelectedItem?.ToString() ?? "JOURNAL VOUCHER";
                                string voucherNoLabel = (selectedVoucherTitle == "EMPLOYEES SUPPLIES VOUCHER" || selectedVoucherTitle == "EMPLOYEE SUPPLIES VOUCHER") ? "E.S.V. No.:" : "J.V. No.:";

                                if (GetReportObjectSafe<TextObject>(cRJV_INT, "TextReportTitle") is TextObject textObject_ReportTitle)
                                    textObject_ReportTitle.Text = selectedVoucherTitle;
                                if (GetReportObjectSafe<TextObject>(cRJV_INT, "TextVoucherNoLabel") is TextObject textObject_VoucherNoLabel)
                                    textObject_VoucherNoLabel.Text = voucherNoLabel;

                                if (GetReportObjectSafe<TextObject>(cRJV_INT, "TextJVCheckDate") is TextObject textObject_JVCheckDate)
                                    textObject_JVCheckDate.Text = journal[0].Date.ToString("MMMM dd, yyyy");
                                if (GetReportObjectSafe<TextObject>(cRJV_INT, "TextJVRefnumber") is TextObject textObject_JVRefnumber)
                                    textObject_JVRefnumber.Text = refNumberCR;

                                if (GetReportObjectSafe<TextObject>(cRJV_INT, "TextCompanyName") is TextObject textObject_CompanyName && comboBox_Company?.SelectedItem != null)
                                    textObject_CompanyName.Text = comboBox_Company.SelectedItem.ToString();

                                AccessToDatabase_INT accessToDatabase = new AccessToDatabase_INT();
                                var signatories = accessToDatabase.RetrieveAllSignatoryData();

                                if (GetReportObjectSafe<TextObject>(cRJV_INT, "TextPreparedBy") is TextObject textObject_PreparedBy)
                                    textObject_PreparedBy.Text = signatories.PreparedByName;
                                if (GetReportObjectSafe<TextObject>(cRJV_INT, "TextCheckedBy") is TextObject textObject_CheckedBy)
                                    textObject_CheckedBy.Text = signatories.ReviewedByName;
                                if (GetReportObjectSafe<TextObject>(cRJV_INT, "TextApprovedBy") is TextObject textObject_ApprovedBy)
                                    textObject_ApprovedBy.Text = signatories.ApprovedByName;

                                if (GetReportObjectSafe<SubreportObject>(cRJV_INT, "SubreportJVDetailsIVP") is SubreportObject subreportObject)
                                {
                                    ReportDocument subReportDocument = cRJV_INT.OpenSubreport(subreportObject.SubreportName);
                                }

                                InsertDataToJournalCompiled(refNumberCR, journal);

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
                                MessageBox.Show("No Journal Entry found for this Reference Number.", "Not Found", MessageBoxButtons.OK, MessageBoxIcon.Information);
                            }
                        }
                        else if (comboBox_Forms.SelectedIndex == 4) // APV
                        {
                            GenerateAPVReport_INT(refNumberCR);
                        }
                    }
                }
                finally
                {
                    Cursor.Current = Cursors.Default;
                    button_SearchRefNum_CR.Enabled = true;
                    button_SearchRefNum_CR.Text = "🔍  SEARCH & LOAD";
                }
            };

            button_SearchRefNum_CR.Click += (sender, e) => performSearch();
            textBox_ReferenceNumber_CR.KeyDown += (sender, e) =>
            {
                if (e.KeyCode == Keys.Enter)
                {
                    e.SuppressKeyPress = true;
                    performSearch();
                }
            };

            panel_RefNumberCrystalReport.Controls.Add(button_SearchRefNum_CR);
            return panel_RefNumberCrystalReport;
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

                TextObject textObject_CVBILLAddress = GetReportObjectSafe<TextObject>(cRAPV_INTBILL, "TextCVBILLAddress");
                TextObject textObject_CVBILLPayee = GetReportObjectSafe<TextObject>(cRAPV_INTBILL, "TextCVBILLPayee");
                TextObject textObject_CVBILLDate = GetReportObjectSafe<TextObject>(cRAPV_INTBILL, "TextCVBILLDate");
                TextObject textObject_CVBILLAmountInWords = GetReportObjectSafe<TextObject>(cRAPV_INTBILL, "TextCVBILLAmountInWords");
                TextObject textObject_CVBILLCheckDate = GetReportObjectSafe<TextObject>(cRAPV_INTBILL, "TextCVBILLCheckDate");
                TextObject textObject_CVBILLRefNumber = GetReportObjectSafe<TextObject>(cRAPV_INTBILL, "TextCVBILLRefNumber");
                TextObject textObject_CVBILLTotalAmount = GetReportObjectSafe<TextObject>(cRAPV_INTBILL, "TextCVBILLTotalAmount");
                TextObject textObject_CVBILLBankAccount = GetReportObjectSafe<TextObject>(cRAPV_INTBILL, "TextCVBILLBankAccount");

                TextObject textObject_PreparedBy = GetReportObjectSafe<TextObject>(cRAPV_INTBILL, "TextPreparedBy");
                TextObject textObject_PreparedByPos = GetReportObjectSafe<TextObject>(cRAPV_INTBILL, "TextPreparedByPosition");
                TextObject textObject_CheckedBy = GetReportObjectSafe<TextObject>(cRAPV_INTBILL, "TextCheckedBy");
                TextObject textObject_CheckedByPos = GetReportObjectSafe<TextObject>(cRAPV_INTBILL, "TextCheckedByPosition");
                TextObject textObject_ApprovedBy = GetReportObjectSafe<TextObject>(cRAPV_INTBILL, "TextApprovedBy");
                TextObject textObject_ApprovedByPos = GetReportObjectSafe<TextObject>(cRAPV_INTBILL, "TextApprovedByPosition");

                AccessToDatabase_INT accessToDatabase = new AccessToDatabase_INT();
                var signatories = accessToDatabase.RetrieveAllSignatoryData();

                if (textObject_PreparedBy != null) textObject_PreparedBy.Text = signatories.PreparedByName;
                if (textObject_PreparedByPos != null) textObject_PreparedByPos.Text = signatories.PreparedByPosition;
                if (textObject_CheckedBy != null) textObject_CheckedBy.Text = signatories.ReviewedByName;
                if (textObject_CheckedByPos != null) textObject_CheckedByPos.Text = signatories.ReviewedByPosition;
                if (textObject_ApprovedBy != null) textObject_ApprovedBy.Text = signatories.ApprovedByName;
                if (textObject_ApprovedByPos != null) textObject_ApprovedByPos.Text = signatories.ApprovedByPosition;

                // =========================================================================
                // CONSOLIDATED TOTALS ACROSS ALL BILLS
                // =========================================================================
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

                if (textObject_CVBILLRefNumber != null) textObject_CVBILLRefNumber.Text = refNumberCR;
                if (textObject_CVBILLAddress != null) textObject_CVBILLAddress.Text = fullAddress;
                if (textObject_CVBILLDate != null) textObject_CVBILLDate.Text = DateTime.Now.ToString("MMMM dd, yyyy");
                if (textObject_CVBILLAmountInWords != null) textObject_CVBILLAmountInWords.Text = amountInWords;
                if (textObject_CVBILLCheckDate != null) textObject_CVBILLCheckDate.Text = latestDueDate.ToString("MMMM dd, yyyy");
                if (textObject_CVBILLPayee != null) textObject_CVBILLPayee.Text = payeeNames;
                if (textObject_CVBILLTotalAmount != null) textObject_CVBILLTotalAmount.Text = totalVoucherAmount.ToString("N2");
                if (textObject_CVBILLBankAccount != null) textObject_CVBILLBankAccount.Text = bankaccount;

                // =========================================================================
                // SUBREPORT 1: DETAILS / REMARKS & PER-BILL AMOUNTS
                // =========================================================================
                SubreportObject subreportObject = GetReportObjectSafe<SubreportObject>(cRAPV_INTBILL, "SubreportCVDetailsINT");
                if (subreportObject != null)
                {
                    ReportDocument subReportDocument = cRAPV_INTBILL.OpenSubreport(subreportObject.SubreportName);
                    SetDatabaseLocation(subReportDocument, databasePathBILL);

                    TextObject textObject_BILLSubRemarks = GetReportObjectSafe<TextObject>(subReportDocument, "TextBILLRemarks");
                    TextObject textObject_BILLSubAmountPayable = GetReportObjectSafe<TextObject>(subReportDocument, "TextBILLSubAmountPayable");

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
                    {
                        textObject_BILLSubRemarks.Text = string.Join("\r\n", memoLines);
                    }

                    if (textObject_BILLSubAmountPayable != null)
                    {
                        textObject_BILLSubAmountPayable.Text = string.Join("\r\n", amountLines);
                    }
                }

                // Subreport 2: IVP (Debit Details)
                SubreportObject subreportObjectIVP = GetReportObjectSafe<SubreportObject>(cRAPV_INTBILL, "SubreportCVDetailsIVP");
                if (subreportObjectIVP != null)
                {
                    ReportDocument subDocIVP = cRAPV_INTBILL.OpenSubreport(subreportObjectIVP.SubreportName);
                    SetDatabaseLocation(subDocIVP, databasePathBILL);
                }

                // Subreport 3: INTCredit (Credit Details - Notes Payable / Vouchers Payable)
                SubreportObject subreportObjectINTCredit = GetReportObjectSafe<SubreportObject>(cRAPV_INTBILL, "SubreportCVDetailsINTCredit");
                if (subreportObjectINTCredit != null)
                {
                    ReportDocument subDocCredit = cRAPV_INTBILL.OpenSubreport(subreportObjectINTCredit.SubreportName);
                    SetDatabaseLocation(subDocCredit, databasePathBILL);
                }

                // =========================================================================
                // GET SELECTED A/P ACCOUNT FROM DROPDOWN
                // =========================================================================
                string selectedAP = comboBox_APAccount?.SelectedItem?.ToString();
                if (string.IsNullOrWhiteSpace(selectedAP))
                {
                    selectedAP = comboBox_APAccount?.Text;
                }
                if (string.IsNullOrWhiteSpace(selectedAP))
                {
                    selectedAP = "Vouchers Payable";
                }

                // Populate MS Access Staging Database with selected AP Account
                InsertDataToAPVBillCompiled(refNumberCR, bills, selectedAP);

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

                TextObject textObject_CVBILLPayee = GetReportObjectSafe<TextObject>(cRCV_INTBILL, "TextCVBILLPayee");
                TextObject textObject_CVBILLAddress = GetReportObjectSafe<TextObject>(cRCV_INTBILL, "TextCVBILLAddress");
                TextObject textObject_CVBILLDate = GetReportObjectSafe<TextObject>(cRCV_INTBILL, "TextCVBILLDate");

                TextObject textObject_CVBILLAmountInWords = GetReportObjectSafe<TextObject>(cRCV_INTBILL, "TextCVBILLAmountinWords");
                TextObject textObject_CVBILLCheckDate = GetReportObjectSafe<TextObject>(cRCV_INTBILL, "TextCVBILLCheckDate");
                TextObject textObject_CVBILLRefnumber = GetReportObjectSafe<TextObject>(cRCV_INTBILL, "TextCVRefNumber");
                TextObject textObject_CVBILLBankAccount = GetReportObjectSafe<TextObject>(cRCV_INTBILL, "TextCVBILLBankAccount");
                TextObject textObject_CVBILLAmount = GetReportObjectSafe<TextObject>(cRCV_INTBILL, "TextCVBILLLAmount");

                TextObject textObject_PreparedBy = GetReportObjectSafe<TextObject>(cRCV_INTBILL, "TextPreparedBy");
                TextObject textObject_PreparedByPos = GetReportObjectSafe<TextObject>(cRCV_INTBILL, "TextPreparedByPosition");
                TextObject textObject_CheckedBy = GetReportObjectSafe<TextObject>(cRCV_INTBILL, "TextCheckedBy");
                TextObject textObject_CheckedByPos = GetReportObjectSafe<TextObject>(cRCV_INTBILL, "TextCheckedByPosition");
                TextObject textObject_ApprovedBy = GetReportObjectSafe<TextObject>(cRCV_INTBILL, "TextApprovedBy");
                TextObject textObject_ApprovedByPos = GetReportObjectSafe<TextObject>(cRCV_INTBILL, "TextApprovedByPosition");

                AccessToDatabase_INT accessToDatabase = new AccessToDatabase_INT();
                var signatories = accessToDatabase.RetrieveAllSignatoryData();

                if (textObject_PreparedBy != null) textObject_PreparedBy.Text = signatories.PreparedByName;
                if (textObject_PreparedByPos != null) textObject_PreparedByPos.Text = signatories.PreparedByPosition;
                if (textObject_CheckedBy != null) textObject_CheckedBy.Text = signatories.ReviewedByName;
                if (textObject_CheckedByPos != null) textObject_CheckedByPos.Text = signatories.ReviewedByPosition;
                if (textObject_ApprovedBy != null) textObject_ApprovedBy.Text = signatories.ApprovedByName;
                if (textObject_ApprovedByPos != null) textObject_ApprovedByPos.Text = signatories.ApprovedByPosition;

                // =========================================================================
                // 1. CALCULATE INDIVIDUAL BILL AMOUNTS AND SUMMARY REMARKS
                // =========================================================================
                var billSummaryList = bills
                    .Where(x => !string.IsNullOrWhiteSpace(x.RefNumber) || !string.IsNullOrWhiteSpace(x.AppliedRefNumber))
                    .GroupBy(x => !string.IsNullOrWhiteSpace(x.AppliedRefNumber) ? x.AppliedRefNumber.Trim() : x.RefNumber.Trim())
                    .Select(g =>
                    {
                        var firstBill = g.First();
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

                // Build list of lines with header if there are multiple bills
                var remarksLines = new List<string>();

                if (billSummaryList.Count > 1)
                {
                    remarksLines.Add("Payment for the following:");
                }

                remarksLines.AddRange(billSummaryList.Select(b => $"SI#{b.RefNumber,-20}{b.Amount,45:N2}"));

                string billRemarksText = string.Join("\r\n", remarksLines);

                // Real total payout
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

                // =========================================================================
                // GET SELECTED A/P ACCOUNT FROM DROPDOWN
                // =========================================================================
                string selectedAP = comboBox_APAccount?.SelectedItem?.ToString();
                if (string.IsNullOrWhiteSpace(selectedAP))
                {
                    selectedAP = comboBox_APAccount?.Text;
                }
                if (string.IsNullOrWhiteSpace(selectedAP))
                {
                    selectedAP = "Vouchers Payable";
                }

                // Populate staging database table with the chosen AP account
                InsertDataToBillCompiled(refNumberCR, bills, selectedAP);

                // Subreport 1: Debit Details
                SubreportObject subreportObjectIVP = GetReportObjectSafe<SubreportObject>(cRCV_INTBILL, "SubreportCVDetailsIVP");
                if (subreportObjectIVP != null)
                {
                    ReportDocument subDocIVP = cRCV_INTBILL.OpenSubreport(subreportObjectIVP.SubreportName);
                    SetDatabaseLocation(subDocIVP, databasePathBILL);
                }

                // Subreport 2: Credit Details (Displays the Vouchers Payable / Notes Payable line)
                SubreportObject subreportObjectINTCredit = GetReportObjectSafe<SubreportObject>(cRCV_INTBILL, "SubreportCVDetailsINTCredit");
                if (subreportObjectINTCredit != null)
                {
                    ReportDocument subDocCredit = cRCV_INTBILL.OpenSubreport(subreportObjectINTCredit.SubreportName);
                    SetDatabaseLocation(subDocCredit, databasePathBILL);
                }

                // Subreport 3: Remarks and Subtotal
                SubreportObject subreportObjectINT = GetReportObjectSafe<SubreportObject>(cRCV_INTBILL, "SubreportCVDetailsINT");
                if (subreportObjectINT != null)
                {
                    ReportDocument subReportDocumentINT = cRCV_INTBILL.OpenSubreport(subreportObjectINT.SubreportName);
                    SetDatabaseLocation(subReportDocumentINT, databasePathBILL);

                    if (GetReportObjectSafe<TextObject>(subReportDocumentINT, "TextBILLRemarks") is TextObject textObject_BILLSubRemarks)
                        textObject_BILLSubRemarks.Text = billRemarksText;
                }

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
                MessageBox.Show($"Bill Payment Report Error:\n{ex.Message}", "Error", MessageBoxButtons.OK, MessageBoxIcon.Error);
                return false;
            }
        }

        public static void InsertDataToCheckVoucherCompiledINT(string refNumber, List<CheckTableExpensesAndItems> checkData)
        {
            string connectionString = AccessToDatabase_INT.GetAccessConnectionString();
            double debitTotalAmount = 0;
            double creditTotalAmount = 0;

            string SafeTrunc(string value, int maxLength)
            {
                if (string.IsNullOrEmpty(value)) return "";
                return value.Length <= maxLength ? value : value.Substring(0, maxLength);
            }

            using (OleDbConnection connection = new OleDbConnection(connectionString))
            {
                connection.Open();
                using (OleDbTransaction transaction = connection.BeginTransaction())
                {
                    try
                    {
                        // 1. Clear old staging data
                        using (OleDbCommand deleteCommand = new OleDbCommand("DELETE FROM CheckVoucherCompiled", connection, transaction))
                        {
                            deleteCommand.ExecuteNonQuery();
                        }

                        // 2. Prepare Insert Query
                        string insertQuery = @"
                        INSERT INTO CheckVoucherCompiled 
                        (RefNumber, [Particulars], [Class], [Debit], [Credit], [Memo], [CustomerJob]) 
                        VALUES 
                        (@RefNumber, @Particulars, @Class, @Debit, @Credit, @Memo, @CustomerJob)";

                        // PASS 1: Group Positive Item Debits
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

                        foreach (var entry in groupedItemDebits.Concat(groupedExpenseDebits))
                        {
                            debitTotalAmount += entry.TotalAmount;
                            using (OleDbCommand command = new OleDbCommand(insertQuery, connection, transaction))
                            {
                                command.Parameters.AddWithValue("@RefNumber", refNumber);
                                command.Parameters.AddWithValue("@Particulars", SafeTrunc(entry.Particulars, 255));
                                command.Parameters.AddWithValue("@Class", string.IsNullOrEmpty(entry.Class) ? (object)DBNull.Value : entry.Class);
                                command.Parameters.AddWithValue("@Debit", entry.TotalAmount.ToString("N2"));
                                command.Parameters.AddWithValue("@Credit", "");
                                command.Parameters.AddWithValue("@Memo", SafeTrunc(entry.Memo, 255));
                                command.Parameters.AddWithValue("@CustomerJob", SafeTrunc(entry.CustomerJob, 255));
                                command.ExecuteNonQuery();
                            }
                        }

                        // PASS 2: Group Negative Credits
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

                        foreach (var entry in groupedItemCredits.Concat(groupedExpenseCredits))
                        {
                            creditTotalAmount += entry.TotalAmount;
                            using (OleDbCommand command = new OleDbCommand(insertQuery, connection, transaction))
                            {
                                command.Parameters.AddWithValue("@RefNumber", refNumber);
                                command.Parameters.AddWithValue("@Particulars", SafeTrunc(entry.Particulars, 255));
                                command.Parameters.AddWithValue("@Class", string.IsNullOrEmpty(entry.Class) ? (object)DBNull.Value : entry.Class);
                                command.Parameters.AddWithValue("@Debit", "");
                                command.Parameters.AddWithValue("@Credit", entry.TotalAmount.ToString("N2"));
                                command.Parameters.AddWithValue("@Memo", SafeTrunc(entry.Memo, 255));
                                command.Parameters.AddWithValue("@CustomerJob", SafeTrunc(entry.CustomerJob, 255));
                                command.ExecuteNonQuery();
                            }
                        }

                        // PASS 3: Main Bank Account Balancing Credit
                        if (checkData != null && checkData.Count > 0)
                        {
                            var mainCheck = checkData[0];
                            double finalCheckCredit = mainCheck.TotalAmount;
                            creditTotalAmount += finalCheckCredit;

                            using (OleDbCommand command = new OleDbCommand(insertQuery, connection, transaction))
                            {
                                command.Parameters.AddWithValue("@RefNumber", refNumber);

                                string bankName = mainCheck.BankAccount.Contains(":")
                                    ? mainCheck.BankAccount.Split(':').Last().Trim()
                                    : mainCheck.BankAccount;

                                command.Parameters.AddWithValue("@Particulars", SafeTrunc(bankName, 255));
                                command.Parameters.AddWithValue("@Class", (object)DBNull.Value);
                                command.Parameters.AddWithValue("@Debit", "");
                                command.Parameters.AddWithValue("@Credit", finalCheckCredit.ToString("N2"));
                                command.Parameters.AddWithValue("@Memo", SafeTrunc(mainCheck.Memo, 255));
                                command.Parameters.AddWithValue("@CustomerJob", (object)DBNull.Value);

                                command.ExecuteNonQuery();
                            }
                        }

                        // PASS 4: Description-only fallbacks
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
                            using (OleDbCommand command = new OleDbCommand(insertQuery, connection, transaction))
                            {
                                command.Parameters.AddWithValue("@RefNumber", refNumber);
                                command.Parameters.AddWithValue("@Particulars", SafeTrunc(desc.Description, 255));
                                command.Parameters.AddWithValue("@Class", (object)DBNull.Value);
                                command.Parameters.AddWithValue("@Debit", "");
                                command.Parameters.AddWithValue("@Credit", "");
                                command.Parameters.AddWithValue("@Memo", SafeTrunc(desc.Memo, 255));
                                command.Parameters.AddWithValue("@CustomerJob", SafeTrunc(desc.CustomerJob, 255));
                                command.ExecuteNonQuery();
                            }
                        }

                        transaction.Commit();
                    }
                    catch (Exception ex)
                    {
                        transaction.Rollback();
                        MessageBox.Show($"Error compiling Check Voucher data: {ex.Message}", "Database Error", MessageBoxButtons.OK, MessageBoxIcon.Error);
                    }
                }
                connection.Close();
            }
        }

        public static void InsertDataToJournalCompiled(string refNumber, List<JournalGridItem> journalData)
        {
            string connectionString = AccessToDatabase_INT.GetAccessConnectionString();
            double debitTotalAmount = 0;
            double creditTotalAmount = 0;

            string SafeTrunc(string value, int maxLength)
            {
                if (string.IsNullOrEmpty(value)) return "";
                return value.Length <= maxLength ? value : value.Substring(0, maxLength);
            }

            using (OleDbConnection connection = new OleDbConnection(connectionString))
            {
                connection.Open();
                using (OleDbTransaction transaction = connection.BeginTransaction())
                {
                    try
                    {
                        using (OleDbCommand deleteCommand = new OleDbCommand("DELETE FROM JV_Compiled", connection, transaction))
                        {
                            deleteCommand.ExecuteNonQuery();
                        }

                        string insertQuery = @"
                                INSERT INTO JV_Compiled 
                                (RefNumber, [AccountNumber], [Particulars], [Class], [Name], [Debit], [Credit], [Memo]) 
                                VALUES 
                                (@RefNumber, @AccountNumber, @Particulars, @Class, @Name, @Debit, @Credit, @Memo)";

                        foreach (var line in journalData)
                        {
                            string accountNumber = SafeTrunc(line.AccountNumber, 50);
                            string particulars = SafeTrunc(line.AccountName, 500);
                            string className = line.Class;
                            string nameValue = SafeTrunc(line.Name, 255);
                            string memoValue = SafeTrunc(line.Memo, 1000);

                            string debitStr = "";
                            string creditStr = "";

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

                            using (OleDbCommand command = new OleDbCommand(insertQuery, connection, transaction))
                            {
                                command.Parameters.AddWithValue("@RefNumber", refNumber);
                                command.Parameters.AddWithValue("@AccountNumber", string.IsNullOrEmpty(accountNumber) ? (object)DBNull.Value : accountNumber);
                                command.Parameters.AddWithValue("@Particulars", particulars);
                                command.Parameters.AddWithValue("@Class", string.IsNullOrEmpty(className) ? (object)DBNull.Value : className);
                                command.Parameters.AddWithValue("@Name", string.IsNullOrEmpty(nameValue) ? (object)DBNull.Value : nameValue);
                                command.Parameters.AddWithValue("@Debit", debitStr);
                                command.Parameters.AddWithValue("@Credit", creditStr);
                                command.Parameters.AddWithValue("@Memo", memoValue);
                                command.ExecuteNonQuery();
                            }
                        }

                        transaction.Commit();
                    }
                    catch (Exception ex)
                    {
                        transaction.Rollback();
                        MessageBox.Show($"Error compiling Journal data: {ex.Message}", "Database Error", MessageBoxButtons.OK, MessageBoxIcon.Error);
                    }
                }
                connection.Close();
            }
        }

        public static void InsertDataToBillCompiled(string refNumber, List<BillTable> bills, string selectedAPAccount = "Vouchers Payable")
        {
            string connectionString = AccessToDatabase_INT.GetAccessConnectionString();
            double debitTotalAmount = 0;
            double creditTotalAmount = 0;

            string SafeTrunc(string value, int maxLength)
            {
                if (string.IsNullOrEmpty(value)) return "";
                return value.Length <= maxLength ? value : value.Substring(0, maxLength);
            }

            using (OleDbConnection connection = new OleDbConnection(connectionString))
            {
                connection.Open();
                using (OleDbTransaction transaction = connection.BeginTransaction())
                {
                    try
                    {
                        using (OleDbCommand deleteCommand = new OleDbCommand("DELETE FROM Bill_Compiled", connection, transaction))
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

                        // PASS 1: POSITIVE DEBIT EXPENSE / ITEM LINES
                        var groupedItemDebits = allDetails
                            .Where(x => !string.IsNullOrEmpty(x.Detail.ItemLineItemRefFullName) && x.Detail.ItemLineAmount > 0)
                            .GroupBy(x => x.Detail.ItemLineItemRefFullName.Trim())
                            .Select(g => new {
                                Particulars = g.Key,
                                Memo = string.Join("; ", g.Select(x => x.Detail.ItemLineMemo).Where(m => !string.IsNullOrWhiteSpace(m)).Distinct()),
                                TotalAmount = g.Sum(x => x.Detail.ItemLineAmount)
                            });

                        var groupedExpenseDebits = allDetails
                            .Where(x => !string.IsNullOrEmpty(x.Detail.ExpenseLineItemRefFullName) && x.Detail.ExpenseLineAmount > 0)
                            .GroupBy(x => x.Detail.ExpenseLineItemRefFullName.Trim())
                            .Select(g => new {
                                Particulars = g.Key,
                                Memo = string.Join("; ", g.Select(x => x.Detail.ExpenseLineMemo).Where(m => !string.IsNullOrWhiteSpace(m)).Distinct()),
                                TotalAmount = g.Sum(x => x.Detail.ExpenseLineAmount)
                            });

                        foreach (var entry in groupedItemDebits.Concat(groupedExpenseDebits))
                        {
                            debitTotalAmount += entry.TotalAmount;
                            using (OleDbCommand command = new OleDbCommand(insertQuery, connection, transaction))
                            {
                                command.Parameters.AddWithValue("@RefNumber", refNumber ?? (object)DBNull.Value);
                                command.Parameters.AddWithValue("@Particulars", SafeTrunc(entry.Particulars, 255));
                                command.Parameters.AddWithValue("@Class", (object)DBNull.Value);
                                command.Parameters.AddWithValue("@Debit", entry.TotalAmount.ToString("N2"));
                                command.Parameters.AddWithValue("@Credit", "");
                                command.Parameters.AddWithValue("@Memo", SafeTrunc(entry.Memo, 255));
                                command.Parameters.AddWithValue("@CustomerJob", (object)DBNull.Value);
                                command.ExecuteNonQuery();
                            }
                        }

                        // PASS 2: NEGATIVE EXPENSE/ITEM LINES
                        var groupedExpenseCredits = allDetails
                            .Where(x => !string.IsNullOrEmpty(x.Detail.ExpenseLineItemRefFullName) && x.Detail.ExpenseLineAmount < 0)
                            .GroupBy(x => {
                                string rawAcc = x.Detail.ExpenseLineItemRefFullName.Trim();
                                return rawAcc.Contains(":") ? rawAcc.Split(':').Last().Trim() : rawAcc;
                            })
                            .Select(g => new {
                                Particulars = g.Key,
                                Memo = string.Join("; ", g.Select(x => x.Detail.ExpenseLineMemo).Where(m => !string.IsNullOrWhiteSpace(m)).Distinct()),
                                TotalCreditAmount = Math.Abs(g.Sum(x => x.Detail.ExpenseLineAmount))
                            });

                        double totalNegativeCredits = 0;

                        foreach (var entry in groupedExpenseCredits)
                        {
                            totalNegativeCredits += entry.TotalCreditAmount;
                            creditTotalAmount += entry.TotalCreditAmount;

                            using (OleDbCommand command = new OleDbCommand(insertQuery, connection, transaction))
                            {
                                command.Parameters.AddWithValue("@RefNumber", refNumber ?? (object)DBNull.Value);
                                command.Parameters.AddWithValue("@Particulars", SafeTrunc(entry.Particulars, 255));
                                command.Parameters.AddWithValue("@Class", (object)DBNull.Value);
                                command.Parameters.AddWithValue("@Debit", "");
                                command.Parameters.AddWithValue("@Credit", entry.TotalCreditAmount.ToString("N2"));
                                command.Parameters.AddWithValue("@Memo", SafeTrunc(entry.Memo, 255));
                                command.Parameters.AddWithValue("@CustomerJob", (object)DBNull.Value);
                                command.ExecuteNonQuery();
                            }
                        }

                        // Discounts / Withholding
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

                            using (OleDbCommand command = new OleDbCommand(insertQuery, connection, transaction))
                            {
                                command.Parameters.AddWithValue("@RefNumber", refNumber ?? (object)DBNull.Value);
                                command.Parameters.AddWithValue("@Particulars", SafeTrunc(disc.AccountName, 255));
                                command.Parameters.AddWithValue("@Class", (object)DBNull.Value);
                                command.Parameters.AddWithValue("@Debit", "");
                                command.Parameters.AddWithValue("@Credit", disc.TotalDiscount.ToString("N2"));
                                command.Parameters.AddWithValue("@Memo", SafeTrunc("Withholding Tax Applied", 255));
                                command.Parameters.AddWithValue("@CustomerJob", (object)DBNull.Value);
                                command.ExecuteNonQuery();
                            }
                        }

                        // PASS 3: NET ACCOUNTS PAYABLE / NOTES PAYABLE CREDIT
                        if (bills != null && bills.Count > 0)
                        {
                            double netPaymentCredit = debitTotalAmount - totalNegativeCredits;
                            creditTotalAmount += netPaymentCredit;

                            string consolidatedMemo = string.Join(" | ", bills
                                .Select(b => !string.IsNullOrWhiteSpace(b.BillMemo) ? b.BillMemo.Trim() : b.Memo?.Trim())
                                .Where(m => !string.IsNullOrWhiteSpace(m))
                                .Distinct());

                            string apAccount = string.IsNullOrWhiteSpace(selectedAPAccount) ? "Vouchers Payable" : selectedAPAccount;

                            using (OleDbCommand command = new OleDbCommand(insertQuery, connection, transaction))
                            {
                                command.Parameters.AddWithValue("@RefNumber", refNumber ?? (object)DBNull.Value);
                                command.Parameters.AddWithValue("@Particulars", SafeTrunc(apAccount, 255));
                                command.Parameters.AddWithValue("@Class", (object)DBNull.Value);
                                command.Parameters.AddWithValue("@Debit", "");
                                command.Parameters.AddWithValue("@Credit", netPaymentCredit.ToString("N2"));
                                command.Parameters.AddWithValue("@Memo", SafeTrunc(consolidatedMemo, 255));
                                command.Parameters.AddWithValue("@CustomerJob", (object)DBNull.Value);
                                command.ExecuteNonQuery();
                            }
                        }

                        transaction.Commit();
                    }
                    catch (Exception ex)
                    {
                        transaction.Rollback();
                        MessageBox.Show($"Error compiling Bill Payment data: {ex.Message}", "Database Error", MessageBoxButtons.OK, MessageBoxIcon.Error);
                    }
                }
                connection.Close();
            }
        }

        public static void InsertDataToAPVBillCompiled(string refNumber, List<BillTable> bills, string selectedAPAccount = "Vouchers Payable")
        {
            string connectionString = AccessToDatabase_INT.GetAccessConnectionString();
            double debitTotalAmount = 0;
            double creditTotalAmount = 0;

            string SafeTrunc(string value, int maxLength)
            {
                if (string.IsNullOrEmpty(value)) return "";
                return value.Length <= maxLength ? value : value.Substring(0, maxLength);
            }

            using (OleDbConnection connection = new OleDbConnection(connectionString))
            {
                connection.Open();
                using (OleDbTransaction transaction = connection.BeginTransaction())
                {
                    try
                    {
                        using (OleDbCommand deleteCommand = new OleDbCommand("DELETE FROM Bill_Compiled", connection, transaction))
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

                        // PASS 1: POSITIVE DEBIT EXPENSE / ITEM LINES
                        var groupedItemDebits = allDetails
                            .Where(x => !string.IsNullOrEmpty(x.Detail.ItemLineItemRefFullName) && x.Detail.ItemLineAmount > 0)
                            .GroupBy(x => x.Detail.ItemLineItemRefFullName.Trim())
                            .Select(g => new {
                                Particulars = g.Key,
                                Memo = string.Join("; ", g.Select(x => x.Detail.ItemLineMemo).Where(m => !string.IsNullOrWhiteSpace(m)).Distinct()),
                                TotalAmount = g.Sum(x => x.Detail.ItemLineAmount)
                            });

                        var groupedExpenseDebits = allDetails
                            .Where(x => !string.IsNullOrEmpty(x.Detail.ExpenseLineItemRefFullName) && x.Detail.ExpenseLineAmount > 0)
                            .GroupBy(x => x.Detail.ExpenseLineItemRefFullName.Trim())
                            .Select(g => new {
                                Particulars = g.Key,
                                Memo = string.Join("; ", g.Select(x => x.Detail.ExpenseLineMemo).Where(m => !string.IsNullOrWhiteSpace(m)).Distinct()),
                                TotalAmount = g.Sum(x => x.Detail.ExpenseLineAmount)
                            });

                        foreach (var entry in groupedItemDebits.Concat(groupedExpenseDebits))
                        {
                            debitTotalAmount += entry.TotalAmount;
                            using (OleDbCommand command = new OleDbCommand(insertQuery, connection, transaction))
                            {
                                command.Parameters.AddWithValue("@RefNumber", refNumber ?? (object)DBNull.Value);
                                command.Parameters.AddWithValue("@Particulars", SafeTrunc(entry.Particulars, 255));
                                command.Parameters.AddWithValue("@Class", (object)DBNull.Value);
                                command.Parameters.AddWithValue("@Debit", entry.TotalAmount.ToString("N2"));
                                command.Parameters.AddWithValue("@Credit", "");
                                command.Parameters.AddWithValue("@Memo", SafeTrunc(entry.Memo, 255));
                                command.Parameters.AddWithValue("@CustomerJob", (object)DBNull.Value);
                                command.ExecuteNonQuery();
                            }
                        }

                        // PASS 2: NEGATIVE EXPENSE/ITEM LINES
                        var groupedExpenseCredits = allDetails
                            .Where(x => !string.IsNullOrEmpty(x.Detail.ExpenseLineItemRefFullName) && x.Detail.ExpenseLineAmount < 0)
                            .GroupBy(x => {
                                string rawAcc = x.Detail.ExpenseLineItemRefFullName.Trim();
                                return rawAcc.Contains(":") ? rawAcc.Split(':').Last().Trim() : rawAcc;
                            })
                            .Select(g => new {
                                Particulars = g.Key,
                                Memo = string.Join("; ", g.Select(x => x.Detail.ExpenseLineMemo).Where(m => !string.IsNullOrWhiteSpace(m)).Distinct()),
                                TotalCreditAmount = Math.Abs(g.Sum(x => x.Detail.ExpenseLineAmount))
                            });

                        double totalNegativeCredits = 0;

                        foreach (var entry in groupedExpenseCredits)
                        {
                            totalNegativeCredits += entry.TotalCreditAmount;
                            creditTotalAmount += entry.TotalCreditAmount;

                            using (OleDbCommand command = new OleDbCommand(insertQuery, connection, transaction))
                            {
                                command.Parameters.AddWithValue("@RefNumber", refNumber ?? (object)DBNull.Value);
                                command.Parameters.AddWithValue("@Particulars", SafeTrunc(entry.Particulars, 255));
                                command.Parameters.AddWithValue("@Class", (object)DBNull.Value);
                                command.Parameters.AddWithValue("@Debit", "");
                                command.Parameters.AddWithValue("@Credit", entry.TotalCreditAmount.ToString("N2"));
                                command.Parameters.AddWithValue("@Memo", SafeTrunc(entry.Memo, 255));
                                command.Parameters.AddWithValue("@CustomerJob", (object)DBNull.Value);
                                command.ExecuteNonQuery();
                            }
                        }

                        // Discounts / Withholding
                        var groupedDiscounts = bills
                            .Where(b => b.AppliedToTxnDiscountAmount > 0 && !string.IsNullOrEmpty(b.AppliedToTxnDiscountAccountRefFullName))
                            .GroupBy(b => {
                                string rawAcc = b.AppliedToTxnDiscountAccountRefFullName;
                                return rawAcc.Contains(":") ? rawAcc.Split(':').Last().Trim() : rawAcc.Trim();
                            })
                            .Select(g => new {
                                AccountName = g.Key,
                                Memos = string.Join("; ", bills.Select(x => x.Memo ?? x.BillMemo).Where(m => !string.IsNullOrWhiteSpace(m)).Distinct()),
                                TotalDiscount = g.Sum(b => Math.Abs(b.AppliedToTxnDiscountAmount))
                            });

                        foreach (var disc in groupedDiscounts)
                        {
                            totalNegativeCredits += disc.TotalDiscount;
                            creditTotalAmount += disc.TotalDiscount;

                            string discMemo = !string.IsNullOrWhiteSpace(disc.Memos) ? disc.Memos : "Withholding Tax Applied";

                            using (OleDbCommand command = new OleDbCommand(insertQuery, connection, transaction))
                            {
                                command.Parameters.AddWithValue("@RefNumber", refNumber ?? (object)DBNull.Value);
                                command.Parameters.AddWithValue("@Particulars", SafeTrunc(disc.AccountName, 255));
                                command.Parameters.AddWithValue("@Class", (object)DBNull.Value);
                                command.Parameters.AddWithValue("@Debit", "");
                                command.Parameters.AddWithValue("@Credit", disc.TotalDiscount.ToString("N2"));
                                command.Parameters.AddWithValue("@Memo", SafeTrunc(discMemo, 255));
                                command.Parameters.AddWithValue("@CustomerJob", (object)DBNull.Value);
                                command.ExecuteNonQuery();
                            }
                        }

                        // PASS 3: NET ACCOUNTS PAYABLE / NOTES PAYABLE CREDIT
                        if (bills != null && bills.Count > 0)
                        {
                            double netPaymentCredit = debitTotalAmount - totalNegativeCredits;
                            creditTotalAmount += netPaymentCredit;

                            string allHeaderMemos = string.Join(" | ", bills
                                .Select(b => !string.IsNullOrWhiteSpace(b.BillMemo) ? b.BillMemo.Trim() : b.Memo?.Trim())
                                .Where(m => !string.IsNullOrWhiteSpace(m))
                                .Distinct());

                            string apAccount = string.IsNullOrWhiteSpace(selectedAPAccount) ? "Vouchers Payable" : selectedAPAccount;

                            using (OleDbCommand command = new OleDbCommand(insertQuery, connection, transaction))
                            {
                                command.Parameters.AddWithValue("@RefNumber", refNumber ?? (object)DBNull.Value);
                                command.Parameters.AddWithValue("@Particulars", SafeTrunc(apAccount, 255));
                                command.Parameters.AddWithValue("@Class", (object)DBNull.Value);
                                command.Parameters.AddWithValue("@Debit", "");
                                command.Parameters.AddWithValue("@Credit", netPaymentCredit.ToString("N2"));
                                command.Parameters.AddWithValue("@Memo", SafeTrunc(allHeaderMemos, 255));
                                command.Parameters.AddWithValue("@CustomerJob", (object)DBNull.Value);
                                command.ExecuteNonQuery();
                            }
                        }

                        transaction.Commit();
                    }
                    catch (Exception ex)
                    {
                        transaction.Rollback();
                        MessageBox.Show($"Error compiling APV data: {ex.Message}", "Database Error", MessageBoxButtons.OK, MessageBoxIcon.Error);
                    }
                }
                connection.Close();
            }
        }

        // =========================================================================
        // SIGNATORY & PRINTING CONTROLS
        // =========================================================================
        private FlowLayoutPanel Panel_SBRefNumber()
        {
            int cardInnerWidth = sideBarWidth - 44;
            panel_RefNumber = CreateCardPanel("🔍  CHECK REFERENCE", sideBarWidth - 24);
            panel_RefNumber.Visible = false;

            Label label_RefNumberText = new Label
            {
                Text = "Enter Check Ref Number:",
                Font = FontInputLabel,
                ForeColor = ColorTextMuted,
                Width = cardInnerWidth,
                Height = 18,
                TextAlign = ContentAlignment.MiddleLeft,
                Margin = new Padding(0, 0, 0, 3)
            };
            panel_RefNumber.Controls.Add(label_RefNumberText);

            TextBox textBox_ReferenceNumber = CreateModernTextBox(cardInnerWidth);
            panel_RefNumber.Controls.Add(textBox_ReferenceNumber);

            Button button_SearchRefNum = CreateModernButton("🔍  SEARCH & PREVIEW", ColorPrimaryBtn, ColorPrimaryHover, Color.White, cardInnerWidth, 32);
            button_SearchRefNum.Margin = new Padding(0, 2, 0, 4);

            Action performCheckSearch = () =>
            {
                if (comboBox_Forms.SelectedIndex == 0)
                {
                    MessageBox.Show("Please select a form type first.", "Notice", MessageBoxButtons.OK, MessageBoxIcon.Information);
                    return;
                }

                string refNumber = textBox_ReferenceNumber.Text.Trim();
                if (string.IsNullOrEmpty(refNumber))
                {
                    MessageBox.Show("Please enter a reference number.", "Notice", MessageBoxButtons.OK, MessageBoxIcon.Warning);
                    return;
                }

                Cursor.Current = Cursors.WaitCursor;
                button_SearchRefNum.Enabled = false;
                button_SearchRefNum.Text = "⏳ SEARCHING...";

                try
                {
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

                    if (data is System.Collections.ICollection collection && collection.Count > 0)
                    {
                        if (GlobalVariables.client == "INT")
                        {
                            Layouts_INT layouts_INT = new Layouts_INT();
                            System.Drawing.Printing.PaperSize paperSize = new System.Drawing.Printing.PaperSize("Custom", 850, 1100);
                            printDocument = new PrintDocument();
                            printDocument.DefaultPageSettings.PaperSize = paperSize;
                            printDocument.PrinterSettings.DefaultPageSettings.PaperSize = paperSize;

                            int selectedIndex = comboBox_Forms.SelectedIndex;
                            string seriesNumberStr = textBox_SeriesNumber.Text;
                            string payeeOverride = textBox_PayeeOverride.Text;

                            itemCounter = 0;
                            pageCounter = 1;
                            printPreviewControl.StartPage = 0;

                            printDocument.PrintPage += (s, ev) =>
                            {
                                layouts_INT.PrintPage_INT(s, ev, selectedIndex, seriesNumberStr, data, payeeOverride);
                            };
                        }

                        printPreviewControl.Document = printDocument;
                        printPreviewControl.Visible = true;
                        panel_Printing.Visible = true;
                    }
                    else
                    {
                        MessageBox.Show("No check data found for reference number: " + refNumber, "Not Found", MessageBoxButtons.OK, MessageBoxIcon.Information);
                    }
                }
                finally
                {
                    Cursor.Current = Cursors.Default;
                    button_SearchRefNum.Enabled = true;
                    button_SearchRefNum.Text = "🔍  SEARCH & PREVIEW";
                }
            };

            button_SearchRefNum.Click += (sender, e) => performCheckSearch();
            textBox_ReferenceNumber.KeyDown += (sender, e) =>
            {
                if (e.KeyCode == Keys.Enter)
                {
                    e.SuppressKeyPress = true;
                    performCheckSearch();
                }
            };

            panel_RefNumber.Controls.Add(button_SearchRefNum);
            return panel_RefNumber;
        }

        public FlowLayoutPanel Panel_SBSignatory()
        {
            int cardInnerWidth = sideBarWidth - 44;
            panel_Signatory = CreateCardPanel("✍️  SIGNATORIES", sideBarWidth - 24);
            panel_Signatory.Visible = false;

            Label label_SignatoryRole = new Label
            {
                Text = "Signatory Role:",
                Font = FontInputLabel,
                ForeColor = ColorTextMuted,
                Width = cardInnerWidth,
                Height = 18,
                TextAlign = ContentAlignment.MiddleLeft,
                Margin = new Padding(0, 0, 0, 3)
            };
            panel_Signatory.Controls.Add(label_SignatoryRole);

            ComboBox comboBox_Signatory = CreateModernComboBox(cardInnerWidth);
            if (GlobalVariables.client == "INT")
            {
                comboBox_Signatory.Items.AddRange(new string[]
                {
                    "-- Select Signatory Role --",
                    "Prepared By:",
                    "Checked By:",
                    "A/P:",
                });
            }
            else
            {
                comboBox_Signatory.Items.AddRange(new string[]
                {
                    "-- Select Signatory Role --",
                    "Prepared By:",
                    "Checked By:",
                    "Approved By:",
                    "Noted By:",
                });
            }
            comboBox_Signatory.SelectedIndex = 0;
            panel_Signatory.Controls.Add(comboBox_Signatory);

            Label label_SignatoryName = new Label
            {
                Text = "Signatory Name:",
                Font = FontInputLabel,
                ForeColor = ColorTextMuted,
                Width = cardInnerWidth,
                Height = 18,
                TextAlign = ContentAlignment.MiddleLeft,
                Margin = new Padding(0, 4, 0, 2)
            };
            panel_Signatory.Controls.Add(label_SignatoryName);

            TextBox textBox_SignatoryName = CreateModernTextBox(cardInnerWidth);
            panel_Signatory.Controls.Add(textBox_SignatoryName);

            Label label_SignatoryPosition = new Label
            {
                Text = "Position / Designation:",
                Font = FontInputLabel,
                ForeColor = ColorTextMuted,
                Width = cardInnerWidth,
                Height = 18,
                TextAlign = ContentAlignment.MiddleLeft,
                Margin = new Padding(0, 4, 0, 2)
            };
            panel_Signatory.Controls.Add(label_SignatoryPosition);

            TextBox textBox_SignatoryPosition = CreateModernTextBox(cardInnerWidth);
            panel_Signatory.Controls.Add(textBox_SignatoryPosition);

            FlowLayoutPanel saveRow = new FlowLayoutPanel
            {
                Width = cardInnerWidth,
                Height = 34,
                FlowDirection = FlowDirection.LeftToRight,
                WrapContents = false,
                Margin = new Padding(0, 4, 0, 0),
                Padding = new Padding(0),
                BackColor = Color.Transparent
            };

            Button button_SaveSignatory = CreateModernButton("💾  SAVE", ColorSuccessBtn, ColorSuccessHover, Color.White, 90, 30);
            button_SaveSignatory.Margin = new Padding(0, 0, 6, 0);
            saveRow.Controls.Add(button_SaveSignatory);

            Label label_SignatoryStatus = new Label
            {
                Height = 30,
                Width = cardInnerWidth - 96,
                TextAlign = ContentAlignment.MiddleLeft,
                Font = new Font("Segoe UI", 8.5f, FontStyle.Bold),
                ForeColor = Color.FromArgb(16, 185, 129),
                Margin = new Padding(0)
            };
            saveRow.Controls.Add(label_SignatoryStatus);

            button_SaveSignatory.Click += (sender, e) =>
            {
                if (comboBox_Signatory.SelectedIndex == 0)
                {
                    MessageBox.Show("Please select a signatory role.", "Notice", MessageBoxButtons.OK, MessageBoxIcon.Information);
                    return;
                }

                string signatoryName = textBox_SignatoryName.Text.Trim();
                string signatoryPosition = textBox_SignatoryPosition.Text.Trim();
                int choice = comboBox_Signatory.SelectedIndex;

                accessToDatabase.SaveSignatoryData(choice, signatoryName, signatoryPosition);
                label_SignatoryStatus.Text = "✓ Saved!";

                Timer t = new Timer { Interval = 3000 };
                t.Tick += (ts, te) =>
                {
                    label_SignatoryStatus.Text = "";
                    t.Stop();
                    t.Dispose();
                };
                t.Start();
            };

            comboBox_Signatory.SelectedIndexChanged += (sender, e) =>
            {
                if (comboBox_Signatory.SelectedIndex == 0)
                {
                    textBox_SignatoryName.Text = "";
                    textBox_SignatoryPosition.Text = "";
                    label_SignatoryStatus.Text = "";
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

            panel_Signatory.Controls.Add(saveRow);
            return panel_Signatory;
        }

        private FlowLayoutPanel Panel_SBRRSignatory()
        {
            int cardInnerWidth = sideBarWidth - 44;
            panel_RRSignatory = CreateCardPanel("✍️  SIGNATORY (RR)", sideBarWidth - 24);

            Label label_ReceivedBy = new Label
            {
                Text = "Received By:",
                Font = FontInputLabel,
                ForeColor = ColorTextMuted,
                Width = cardInnerWidth,
                Height = 18,
                TextAlign = ContentAlignment.MiddleLeft,
                Margin = new Padding(0, 0, 0, 2)
            };
            panel_RRSignatory.Controls.Add(label_ReceivedBy);

            textBox_ReceivedByRR = CreateModernTextBox(cardInnerWidth);
            panel_RRSignatory.Controls.Add(textBox_ReceivedByRR);

            Label label_CheckedBy = new Label
            {
                Text = "Checked By:",
                Font = FontInputLabel,
                ForeColor = ColorTextMuted,
                Width = cardInnerWidth,
                Height = 18,
                TextAlign = ContentAlignment.MiddleLeft,
                Margin = new Padding(0, 4, 0, 2)
            };
            panel_RRSignatory.Controls.Add(label_CheckedBy);

            textBox_CheckedByRR = CreateModernTextBox(cardInnerWidth);
            panel_RRSignatory.Controls.Add(textBox_CheckedByRR);

            FlowLayoutPanel saveRow = new FlowLayoutPanel
            {
                Width = cardInnerWidth,
                Height = 34,
                FlowDirection = FlowDirection.LeftToRight,
                WrapContents = false,
                Margin = new Padding(0, 4, 0, 0),
                Padding = new Padding(0),
                BackColor = Color.Transparent
            };

            Button button_SaveRRSignatory = CreateModernButton("💾  SAVE", ColorSuccessBtn, ColorSuccessHover, Color.White, 90, 30);
            button_SaveRRSignatory.Margin = new Padding(0, 0, 6, 0);
            saveRow.Controls.Add(button_SaveRRSignatory);

            label_SignatoryRRStatus = new Label
            {
                Height = 30,
                Width = cardInnerWidth - 96,
                TextAlign = ContentAlignment.MiddleLeft,
                Font = new Font("Segoe UI", 8.5f, FontStyle.Bold),
                ForeColor = Color.FromArgb(16, 185, 129),
                Margin = new Padding(0)
            };
            saveRow.Controls.Add(label_SignatoryRRStatus);

            button_SaveRRSignatory.Click += (sender, e) =>
            {
                string signatoryName = textBox_ReceivedByRR.Text.Trim();
                string signatoryPosition = textBox_CheckedByRR.Text.Trim();

                accessToDatabase.SaveSignatoryRRData(signatoryName, signatoryPosition);
                label_SignatoryRRStatus.Text = "✓ Saved!";

                Timer t = new Timer { Interval = 3000 };
                t.Tick += (ts, te) =>
                {
                    label_SignatoryRRStatus.Text = "";
                    t.Stop();
                    t.Dispose();
                };
                t.Start();
            };

            panel_RRSignatory.Controls.Add(saveRow);
            return panel_RRSignatory;
        }

        private FlowLayoutPanel Panel_SBPrinting()
        {
            int cardInnerWidth = sideBarWidth - 44;
            panel_Printing = CreateCardPanel("🖨️  PRINT & PREVIEW", sideBarWidth - 24);
            panel_Printing.Visible = false;

            int halfBtnWidth = (cardInnerWidth - 6) / 2;

            // Zoom row
            FlowLayoutPanel zoomRow = new FlowLayoutPanel
            {
                Width = cardInnerWidth,
                Height = 32,
                FlowDirection = FlowDirection.LeftToRight,
                WrapContents = false,
                Margin = new Padding(0, 0, 0, 4),
                Padding = new Padding(0),
                BackColor = Color.Transparent
            };

            Button button_ZoomOut = CreateModernButton("🔍- Zoom Out", CurrentSecondaryBtn, CurrentSecondaryHover, CurrentSecondaryText, halfBtnWidth, 28, null, "Secondary");
            button_ZoomOut.Margin = new Padding(0, 0, 6, 0);
            button_ZoomOut.Click += (sender, e) =>
            {
                if (printPreviewControl.Zoom >= 0.2)
                {
                    printPreviewControl.Zoom -= 0.1;
                }
            };
            zoomRow.Controls.Add(button_ZoomOut);

            Button button_ZoomIn = CreateModernButton("🔍+ Zoom In", CurrentSecondaryBtn, CurrentSecondaryHover, CurrentSecondaryText, halfBtnWidth, 28, null, "Secondary");
            button_ZoomIn.Margin = new Padding(0);
            button_ZoomIn.Click += (sender, e) =>
            {
                if (printPreviewControl.Zoom <= 3.0)
                {
                    printPreviewControl.Zoom += 0.1;
                }
            };
            zoomRow.Controls.Add(button_ZoomIn);

            panel_Printing.Controls.Add(zoomRow);

            // Page row
            FlowLayoutPanel pageRow = new FlowLayoutPanel
            {
                Width = cardInnerWidth,
                Height = 32,
                FlowDirection = FlowDirection.LeftToRight,
                WrapContents = false,
                Margin = new Padding(0, 0, 0, 6),
                Padding = new Padding(0),
                BackColor = Color.Transparent
            };

            Button button_PreviousPage = CreateModernButton("◀ Prev Page", CurrentSecondaryBtn, CurrentSecondaryHover, CurrentSecondaryText, halfBtnWidth, 28, null, "Secondary");
            button_PreviousPage.Margin = new Padding(0, 0, 6, 0);
            button_PreviousPage.Click += (sender, e) =>
            {
                if (printPreviewControl.StartPage > 0)
                {
                    printPreviewControl.StartPage--;
                }
            };
            pageRow.Controls.Add(button_PreviousPage);

            Button button_NextPage = CreateModernButton("Next Page ▶", CurrentSecondaryBtn, CurrentSecondaryHover, CurrentSecondaryText, halfBtnWidth, 28, null, "Secondary");
            button_NextPage.Margin = new Padding(0);
            button_NextPage.Click += (sender, e) =>
            {
                if (printPreviewControl.StartPage < pageCounter - 1)
                {
                    printPreviewControl.StartPage++;
                }
            };
            pageRow.Controls.Add(button_NextPage);

            panel_Printing.Controls.Add(pageRow);

            // Full-width Print button
            Button button_Print = CreateModernButton("🖨️  PRINT DOCUMENT", ColorSuccessBtn, ColorSuccessHover, Color.White, cardInnerWidth, 34, new Font("Segoe UI", 9.5f, FontStyle.Bold));
            button_Print.Margin = new Padding(0, 2, 0, 2);
            button_Print.Click += (sender, e) =>
            {
                try
                {
                    itemCounter = 0;
                    pageCounter = 1;

                    if (comboBox_Forms.SelectedIndex == 4) // APV
                    {
                        int totalItemDetails = apvData.Sum(apv => apv.ItemDetails.Count);
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

                        if (GlobalVariables.client == "INT")
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
                finally
                {
                    GlobalVariables.includeImage = true;
                }
            };
            panel_Printing.Controls.Add(button_Print);

            return panel_Printing;
        }

        // =========================================================================
        // FORM SWITCHING & VISIBILITY LOGIC
        // =========================================================================
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

                        SetVoucherTypeVisibility(true);
                        break;

                    case 4: // Accounts Payable Voucher (APV)
                        prefix = "APV";
                        panel_SeriesNumber.Visible = false;
                        panel_RefNumber.Visible = false;
                        panel_RefNumberCrystalReport.Visible = true;
                        panel_Signatory.Visible = true;

                        label_SeriesNumberText.Text = "Current Series Number: APV";

                        SetAPAccountVisibility(true);

                        panel_Main.Visible = false;
                        panel_Main_CR.Visible = true;
                        break;

                    case 5: // Online Voucher
                        panel_SeriesNumber.Visible = false;
                        panel_RefNumber.Visible = false;
                        panel_RefNumberCrystalReport.Visible = true;
                        panel_Signatory.Visible = true;

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

        private static string SafeTruncate(string value, int maxLength)
        {
            if (string.IsNullOrEmpty(value)) return "";
            return value.Length > maxLength ? value.Substring(0, maxLength) : value;
        }
    }
}
