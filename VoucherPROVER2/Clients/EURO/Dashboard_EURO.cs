using System;
using System.Collections.Generic;
using System.Data;
using System.Data.OleDb;
using System.Drawing;
using System.Drawing.Printing;
using System.IO;
using System.Linq;
using System.Windows.Forms;
using CrystalDecisions.CrystalReports.Engine;
using CrystalDecisions.Shared;
using CrystalDecisions.Windows.Forms;
using static VoucherPROVER2.Clients.EURO.Dataclass_EURO;

namespace VoucherPROVER2.Clients.EURO
{
    public partial class Dashboard_EURO : Form
    {
        public Dashboard_EURO()
        {
            InitializeComponent();

            accessToDatabase = new AccessToDatabase_EURO();
            this.CreateHandle();
        }

        private PrintDocument printDocument;
        private PrintPreviewControl printPreviewControl;
        private CrystalReportViewer reportViewer;
        private AccessToDatabase_EURO accessToDatabase;

        // UI Controls
        private FlowLayoutPanel panel_Company;
        private ComboBox comboBox_Forms;
        private ComboBox comboBox_Company;
        private Label label_CompanyText;
        private Label label_CurrencyText;
        private ComboBox comboBox_Currency;

        private Label label_SeriesNumberText;
        private TextBox textBox_SeriesNumber;

        private FlowLayoutPanel panel_PayeeOverride;
        private TextBox textBox_PayeeOverride;

        private Panel panel_Main;
        private Panel panel_Main_CR;

        private FlowLayoutPanel panel_Printing;
        private FlowLayoutPanel panel_SeriesNumber;
        private FlowLayoutPanel panel_Signatory;
        private FlowLayoutPanel panel_RefNumber;
        private FlowLayoutPanel panel_RefNumberCrystalReport;

        // Data caches
        private List<CheckTableGrid> checkivp = new List<CheckTableGrid>();
        private List<CheckTableExpensesAndItems> cvData = new List<CheckTableExpensesAndItems>();
        private List<JournalGridItem> journal = new List<JournalGridItem>();

        private const int sideBarWidth = 270;
        private int seriesNumber = 1;
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

        // Accent Colors
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

        // Aliases
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
            panel_Company = CreateCardPanel("🏢  COMPANY & CURRENCY", sideBarWidth - 24);
            panel_Company.Visible = true;

            label_CompanyText = new Label
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
                "EURO-PACIFIC HEALTH CARE DISTRIBUTOR INC."
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
                    string selectedCompany = comboBox_Company.SelectedItem?.ToString() ?? "EURO-PACIFIC HEALTH CARE DISTRIBUTOR INC.";
                    seriesNumber = accessToDatabase.GetSeriesNumberFromDatabase(formType, selectedCompany);
                    UpdateSeriesNumberEURO(formType);
                }
            };
            panel_Company.Controls.Add(comboBox_Company);

            label_CurrencyText = new Label
            {
                Text = "💱  Select Currency:",
                Font = FontInputLabel,
                ForeColor = ColorTextMuted,
                Width = cardInnerWidth,
                Height = 18,
                TextAlign = ContentAlignment.MiddleLeft,
                Margin = new Padding(0, 6, 0, 3)
            };
            panel_Company.Controls.Add(label_CurrencyText);

            comboBox_Currency = CreateModernComboBox(cardInnerWidth);
            comboBox_Currency.Items.AddRange(new string[] { "Peso (₱)", "Dollar ($)" });
            comboBox_Currency.SelectedIndex = 0;
            panel_Company.Controls.Add(comboBox_Currency);

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
            Panel panel_SideBar_Local = SideBarPanel();

            panel_Container.Controls.Add(panel_Main);
            panel_Container.Controls.Add(panel_Main_CR);
            panel_Container.Controls.Add(panel_SideBar_Local);
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
                Text = "EURO",
                Font = new Font("Segoe UI", 8f, FontStyle.Bold),
                ForeColor = Color.FromArgb(59, 130, 246),
                BackColor = Color.FromArgb(30, 41, 59),
                Padding = new Padding(6, 2, 6, 2),
                AutoSize = true,
                Location = new Point(150, 16)
            };

            Label labelCompanyHeader = new Label
            {
                Parent = panel_Title,
                Text = "EURO-PACIFIC HEALTH CARE DISTRIBUTOR INC.",
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
                            string formType = "";
                            if (comboBox_Forms.SelectedIndex == 1) formType = "CV";
                            else if (comboBox_Forms.SelectedIndex == 3) formType = "JV";

                            string selectedCompany = comboBox_Company.SelectedItem?.ToString() ?? "EURO-PACIFIC HEALTH CARE DISTRIBUTOR INC.";

                            if (!string.IsNullOrEmpty(formType) && !string.IsNullOrEmpty(selectedCompany))
                            {
                                seriesNumber++;
                                accessToDatabase.UpdateManualSeriesNumber(formType, seriesNumber, selectedCompany);
                                this.BeginInvoke((MethodInvoker)delegate
                                {
                                    UpdateSeriesNumberEURO(formType);
                                });
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

            // 2. COMPANY & CURRENCY CARD
            panel_Company = Panel_SBCompany();
            panel_SideBar.Controls.Add(panel_Company);

            panel_PayeeOverride = Panel_SBPayeeOverride();
            panel_SideBar.Controls.Add(panel_PayeeOverride);

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

            // 6. PRINTING CONTROLS CARD
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
            comboBox_Forms.Items.AddRange(new string[]
            {
                "-- Select Form --",
                "Check Voucher",
                "Check",
                "Journal Voucher"
            });
            comboBox_Forms.SelectedIndex = 0;
            comboBox_Forms.SelectedIndexChanged += ComboBox_Forms_SelectedIndexChanged;
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
            textBox_SeriesNumber.TextChanged += TextBox_SeriesNumber_TextChanged;
            textBox_SeriesNumber.Leave += TextBox_SeriesNumber_Leave;
            rowPanel.Controls.Add(textBox_SeriesNumber);

            Button button_Decrement = CreateModernButton("-", CurrentSecondaryBtn, CurrentSecondaryHover, CurrentSecondaryText, 32, 28, new Font("Segoe UI", 11f, FontStyle.Bold), "Secondary");
            button_Decrement.Margin = new Padding(0, 0, 4, 0);
            button_Decrement.Click += (sender, e) =>
            {
                seriesNumber--;
                string prefix = "";
                if (comboBox_Forms.SelectedIndex == 1) prefix = "CV";
                else if (comboBox_Forms.SelectedIndex == 3) prefix = "JV";

                UpdateSeriesNumberEURO(prefix);
            };
            rowPanel.Controls.Add(button_Decrement);

            Button button_Increment = CreateModernButton("+", CurrentSecondaryBtn, CurrentSecondaryHover, CurrentSecondaryText, 32, 28, new Font("Segoe UI", 11f, FontStyle.Bold), "Secondary");
            button_Increment.Margin = new Padding(0);
            button_Increment.Click += (sender, e) =>
            {
                seriesNumber++;
                string prefix = "";
                if (comboBox_Forms.SelectedIndex == 1) prefix = "CV";
                else if (comboBox_Forms.SelectedIndex == 3) prefix = "JV";

                UpdateSeriesNumberEURO(prefix);
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
                    // -------------------------------------------------------------
                    // OPTION 1: CHECK VOUCHER
                    // -------------------------------------------------------------
                    if (comboBox_Forms.SelectedIndex == 1)
                    {
                        bool cvDataExists = false;
                        try
                        {
                            CRCV_EURO cRCV_EURO = new CRCV_EURO();
                            string databasePath = Path.Combine(Application.StartupPath, "CheckDatabase.accdb");
                            SetDatabaseLocation(cRCV_EURO, databasePath);

                            AccessQueries_EURO accessQueries = new AccessQueries_EURO();
                            cvData = accessQueries.GetCheckExpensesAndItemsData_EURO(refNumberCR);

                            if (cvData != null && cvData.Count > 0)
                            {
                                cvDataExists = true;

                                if (GetReportObjectSafe<TextObject>(cRCV_EURO, "TextCVRefNumber") is TextObject textObject_CVRefNumber)
                                    textObject_CVRefNumber.Text = textBox_SeriesNumber.Text;
                                if (GetReportObjectSafe<TextObject>(cRCV_EURO, "TextCVDateTime") is TextObject textObject_CVDateTime)
                                    textObject_CVDateTime.Text = DateTime.Now.ToString("MMMM dd, yyyy");
                                if (GetReportObjectSafe<TextObject>(cRCV_EURO, "TextCVPayee") is TextObject textObject_CVPayee)
                                    textObject_CVPayee.Text = cvData[0].PayeeFullName;

                                double totalCheckAmount = cvData[0].TotalAmount;
                                string amountInWords = AccessToDatabase_EURO.AmountToWordsConverter.Convert(totalCheckAmount);

                                if (GetReportObjectSafe<TextObject>(cRCV_EURO, "TextCVAddress") is TextObject textObject_CVAddress_AmountInWords)
                                    textObject_CVAddress_AmountInWords.Text = amountInWords;
                                if (GetReportObjectSafe<TextObject>(cRCV_EURO, "TextCVAmountInWords") is TextObject textObject_CVAmountInWords)
                                    textObject_CVAmountInWords.Text = amountInWords;
                                if (GetReportObjectSafe<TextObject>(cRCV_EURO, "TextCVAmount") is TextObject textObject_CVAmount)
                                    textObject_CVAmount.Text = totalCheckAmount.ToString("N2");

                                if (GetReportObjectSafe<TextObject>(cRCV_EURO, "TextCompanyName") is TextObject textObject_CompanyName && comboBox_Company?.SelectedItem != null)
                                    textObject_CompanyName.Text = comboBox_Company.SelectedItem.ToString();

                                AccessToDatabase_EURO accessToDb = new AccessToDatabase_EURO();
                                var signatories = accessToDb.RetrieveAllSignatoryData();

                                if (GetReportObjectSafe<TextObject>(cRCV_EURO, "TextPreparedBy") is TextObject textObject_PreparedBy)
                                    textObject_PreparedBy.Text = signatories.PreparedByName;
                                if (GetReportObjectSafe<TextObject>(cRCV_EURO, "TextPreparedByPosition") is TextObject textObject_PreparedByPos)
                                    textObject_PreparedByPos.Text = signatories.PreparedByPosition;
                                if (GetReportObjectSafe<TextObject>(cRCV_EURO, "TextCheckedBy") is TextObject textObject_CheckedBy)
                                    textObject_CheckedBy.Text = signatories.ReviewedByName;
                                if (GetReportObjectSafe<TextObject>(cRCV_EURO, "TextCheckedByPosition") is TextObject textObject_CheckedByPos)
                                    textObject_CheckedByPos.Text = signatories.ReviewedByPosition;
                                if (GetReportObjectSafe<TextObject>(cRCV_EURO, "TextApprovedBy") is TextObject textObject_ApprovedBy)
                                    textObject_ApprovedBy.Text = signatories.ApprovedByName;
                                if (GetReportObjectSafe<TextObject>(cRCV_EURO, "TextApprovedByPosition") is TextObject textObject_ApprovedByPos)
                                    textObject_ApprovedByPos.Text = signatories.ApprovedByPosition;
                                if (GetReportObjectSafe<TextObject>(cRCV_EURO, "TextReceivedBy") is TextObject textObject_ReceivedBy)
                                    textObject_ReceivedBy.Text = signatories.ReceivedByName;
                                if (GetReportObjectSafe<TextObject>(cRCV_EURO, "TextReceivedByPosition") is TextObject textObject_ReceivedByPos)
                                    textObject_ReceivedByPos.Text = signatories.ReceivedByPosition;

                                string rawBank = cvData[0].BankAccount ?? "";
                                string bank = rawBank.Contains(":") ? rawBank.Split(':').Last().Trim() : rawBank;

                                if (GetReportObjectSafe<TextObject>(cRCV_EURO, "TextCVCheckNum") is TextObject textObject_CVCheckNumber)
                                    textObject_CVCheckNumber.Text = cvData[0].RefNumber;
                                if (GetReportObjectSafe<TextObject>(cRCV_EURO, "TextCVCheckBank") is TextObject textObject_CVCheckBank)
                                    textObject_CVCheckBank.Text = bank;
                                if (GetReportObjectSafe<TextObject>(cRCV_EURO, "TextCVCheckDate") is TextObject textObject_CVCheckDate)
                                    textObject_CVCheckDate.Text = cvData[0].DueDate.ToString("MMMM dd, yyyy");
                                if (GetReportObjectSafe<TextObject>(cRCV_EURO, "TextCVDuePayment") is TextObject textObject_CVDuePayment)
                                    textObject_CVDuePayment.Text = totalCheckAmount.ToString("N2");

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

                                string currencyPrefix = (comboBox_Currency?.SelectedIndex == 1) ? "USD " : "PHP ";
                                if (GetReportObjectSafe<TextObject>(cRCV_EURO, "TextCVTotalDebitAmount") is TextObject textObject_CVTotalDebitAmount)
                                    textObject_CVTotalDebitAmount.Text = $"{currencyPrefix}{debitTotalAmount:N2}";
                                if (GetReportObjectSafe<TextObject>(cRCV_EURO, "TextCVTotalCreditAmount") is TextObject textObject_CVTotalCreditAmount)
                                    textObject_CVTotalCreditAmount.Text = $"{currencyPrefix}{debitTotalAmount:N2}";

                                if (GetReportObjectSafe<SubreportObject>(cRCV_EURO, "SubreportCVDetailsIVP") is SubreportObject subreportObject)
                                {
                                    ReportDocument subReportDocument = cRCV_EURO.OpenSubreport(subreportObject.SubreportName);
                                    if (GetReportObjectSafe<TextObject>(subReportDocument, "TextRemarks") is TextObject textObject_Remarks)
                                        textObject_Remarks.Text = cvData[0].Memo;
                                    if (GetReportObjectSafe<TextObject>(subReportDocument, "TextSubAccountPayable") is TextObject textObject_SubAccountPayable)
                                        textObject_SubAccountPayable.Text = "Cash in Bank - " + bank;
                                    if (GetReportObjectSafe<TextObject>(subReportDocument, "TextSubAmountPayable") is TextObject textObject_SubAmountPayable)
                                        textObject_SubAmountPayable.Text = debitTotalAmount.ToString("N2");
                                    if (GetReportObjectSafe<TextObject>(subReportDocument, "TextSubAccountCode") is TextObject textObject_SubAccountCode)
                                        textObject_SubAccountCode.Text = "";

                                    InsertDataToCheckVoucherCompiledEURO(refNumberCR, cvData);
                                }

                                cRCV_EURO.SetParameterValue("ReferenceNumber", refNumberCR);

                                panel_Printing.Visible = false;
                                panel_Signatory.Visible = true;
                                panel_Main.Visible = false;
                                panel_Main_CR.Visible = true;

                                reportViewer.ReportSource = cRCV_EURO;
                                reportViewer.RefreshReport();
                            }
                        }
                        catch (Exception ex)
                        {
                            MessageBox.Show($"EURO CV ERROR:\n{ex.Message}", "Error", MessageBoxButtons.OK, MessageBoxIcon.Error);
                        }

                        if (!cvDataExists)
                        {
                            GenerateBillPaymentReport_EURO(refNumberCR);
                        }
                    }
                    // -------------------------------------------------------------
                    // OPTION 3: JOURNAL VOUCHER
                    // -------------------------------------------------------------
                    else if (comboBox_Forms.SelectedIndex == 3)
                    {
                        CRJV_EURO cRJV_EURO = new CRJV_EURO();
                        string databasePath = Path.Combine(Application.StartupPath, "CheckDatabase.accdb");
                        SetDatabaseLocation(cRJV_EURO, databasePath);

                        AccessQueries_EURO accessQueries = new AccessQueries_EURO();
                        journal = accessQueries.GetJournalEntryForGrid(refNumberCR);

                        if (journal != null && journal.Count > 0)
                        {
                            if (GetReportObjectSafe<TextObject>(cRJV_EURO, "TextJVRefNumber") is TextObject textObject_JVRefNumber)
                                textObject_JVRefNumber.Text = textBox_SeriesNumber.Text;
                            if (GetReportObjectSafe<TextObject>(cRJV_EURO, "TextRefnumber") is TextObject textObject_JVRefNumber2)
                                textObject_JVRefNumber2.Text = $"NO: {refNumberCR}";
                            if (GetReportObjectSafe<TextObject>(cRJV_EURO, "TextJVCheckDate") is TextObject textObject_JVCheckDate)
                                textObject_JVCheckDate.Text = DateTime.Now.ToString("MMMM dd, yyyy");
                            if (GetReportObjectSafe<TextObject>(cRJV_EURO, "TextJVTransactDate") is TextObject textObject_JVTransactDate)
                                textObject_JVTransactDate.Text = journal[0].Date.ToString("MMMM dd, yyyy");

                            if (GetReportObjectSafe<TextObject>(cRJV_EURO, "TextCompanyName") is TextObject textObject_CompanyName && comboBox_Company?.SelectedItem != null)
                                textObject_CompanyName.Text = comboBox_Company.SelectedItem.ToString();

                            AccessToDatabase_EURO accessToDb = new AccessToDatabase_EURO();
                            var signatories = accessToDb.RetrieveAllSignatoryData();

                            if (GetReportObjectSafe<TextObject>(cRJV_EURO, "TextPreparedBy") is TextObject textObject_PreparedBy)
                                textObject_PreparedBy.Text = signatories.PreparedByName;
                            if (GetReportObjectSafe<TextObject>(cRJV_EURO, "TextPreparedByPosition") is TextObject textObject_PreparedByPos)
                                textObject_PreparedByPos.Text = signatories.PreparedByPosition;
                            if (GetReportObjectSafe<TextObject>(cRJV_EURO, "TextCheckedBy") is TextObject textObject_CheckedBy)
                                textObject_CheckedBy.Text = signatories.ReviewedByName;
                            if (GetReportObjectSafe<TextObject>(cRJV_EURO, "TextCheckedByPosition") is TextObject textObject_CheckedByPos)
                                textObject_CheckedByPos.Text = signatories.ReviewedByPosition;
                            if (GetReportObjectSafe<TextObject>(cRJV_EURO, "TextApprovedBy") is TextObject textObject_ApprovedBy)
                                textObject_ApprovedBy.Text = signatories.ApprovedByName;
                            if (GetReportObjectSafe<TextObject>(cRJV_EURO, "TextApprovedByPosition") is TextObject textObject_ApprovedByPos)
                                textObject_ApprovedByPos.Text = signatories.ApprovedByPosition;

                            double debitTotalAmount = 0;
                            double creditTotalAmount = 0;

                            foreach (var line in journal)
                            {
                                debitTotalAmount += line.Debit;
                                creditTotalAmount += line.Credit;
                            }

                            if (GetReportObjectSafe<TextObject>(cRJV_EURO, "TextJVTotalDebitAmount") is TextObject textObject_JVTotalDebitAmount)
                                textObject_JVTotalDebitAmount.Text = $"{debitTotalAmount:N2}";
                            if (GetReportObjectSafe<TextObject>(cRJV_EURO, "TextJVTotalCreditAmount") is TextObject textObject_JVTotalCreditAmount)
                                textObject_JVTotalCreditAmount.Text = $"{creditTotalAmount:N2}";

                            if (GetReportObjectSafe<SubreportObject>(cRJV_EURO, "SubreportJVDetailsIVP") is SubreportObject subreportObject)
                            {
                                ReportDocument subReportDocument = cRJV_EURO.OpenSubreport(subreportObject.SubreportName);
                                if (GetReportObjectSafe<TextObject>(subReportDocument, "TextJVSUBAccountsPayable") is TextObject textObject_SubAccountPayable)
                                    textObject_SubAccountPayable.Text = journal[0].AccountName;
                                if (GetReportObjectSafe<TextObject>(subReportDocument, "TextJVSUBAmountPayable") is TextObject textObject_SubAmountPayable)
                                    textObject_SubAmountPayable.Text = debitTotalAmount.ToString("N2");
                            }

                            InsertDataToJournalCompiled(refNumberCR, journal);

                            cRJV_EURO.SetParameterValue("ReferenceNumber", refNumberCR);

                            panel_Printing.Visible = false;
                            panel_Signatory.Visible = true;
                            panel_Main.Visible = false;
                            panel_Main_CR.Visible = true;

                            reportViewer.ReportSource = cRJV_EURO;
                            reportViewer.RefreshReport();
                        }
                        else
                        {
                            MessageBox.Show("No Journal Entry found for this Reference Number.", "Not Found", MessageBoxButtons.OK, MessageBoxIcon.Information);
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

        private bool GenerateBillPaymentReport_EURO(string refNumberCR)
        {
            try
            {
                CRCV_EUROBILL cRCV_EUROBILL = new CRCV_EUROBILL();
                string databasePath = Path.Combine(Application.StartupPath, "CheckDatabase.accdb");
                SetDatabaseLocation(cRCV_EUROBILL, databasePath);

                AccessQueries_EURO queries = new AccessQueries_EURO();
                List<BillTable> bills = queries.GetBillData_EURO(refNumberCR);

                if (bills == null || bills.Count == 0)
                {
                    MessageBox.Show("No Check or Bill Payment found for this Reference Number.", "Not Found", MessageBoxButtons.OK, MessageBoxIcon.Information);
                    return false;
                }

                if (GetReportObjectSafe<TextObject>(cRCV_EUROBILL, "TextCVBILLRefNumber") is TextObject textObject_CVBILLRefNumber)
                    textObject_CVBILLRefNumber.Text = textBox_SeriesNumber.Text;
                if (GetReportObjectSafe<TextObject>(cRCV_EUROBILL, "TextCVRefNumber") is TextObject textObject_CVRefNumber)
                    textObject_CVRefNumber.Text = textBox_SeriesNumber.Text;

                if (GetReportObjectSafe<TextObject>(cRCV_EUROBILL, "TextCVBILLDateTime") is TextObject textObject_CVBILLDateTime)
                    textObject_CVBILLDateTime.Text = DateTime.Now.ToString("MMMM dd, yyyy");
                if (GetReportObjectSafe<TextObject>(cRCV_EUROBILL, "TextCVDateTime") is TextObject textObject_CVDateTime)
                    textObject_CVDateTime.Text = DateTime.Now.ToString("MMMM dd, yyyy");

                if (GetReportObjectSafe<TextObject>(cRCV_EUROBILL, "TextCVBILLPayee") is TextObject textObject_CVBILLPayee)
                    textObject_CVBILLPayee.Text = bills[0].PayeeFullName;
                if (GetReportObjectSafe<TextObject>(cRCV_EUROBILL, "TextCVPayee") is TextObject textObject_CVPayee)
                    textObject_CVPayee.Text = bills[0].PayeeFullName;

                double totalBillAmount = bills[0].Amount;
                string amountInWords = AccessToDatabase_EURO.AmountToWordsConverter.Convert(totalBillAmount);

                if (GetReportObjectSafe<TextObject>(cRCV_EUROBILL, "TextCVBILLAddress") is TextObject textObject_CVBILLAddress)
                    textObject_CVBILLAddress.Text = amountInWords;
                if (GetReportObjectSafe<TextObject>(cRCV_EUROBILL, "TextCVAddress") is TextObject textObject_CVAddress)
                    textObject_CVAddress.Text = amountInWords;
                if (GetReportObjectSafe<TextObject>(cRCV_EUROBILL, "TextCVAmountInWords") is TextObject textObject_CVAmountInWords)
                    textObject_CVAmountInWords.Text = amountInWords;

                if (GetReportObjectSafe<TextObject>(cRCV_EUROBILL, "TextCVBILLDue") is TextObject textObject_CVBILLDue)
                    textObject_CVBILLDue.Text = totalBillAmount.ToString("N2");
                if (GetReportObjectSafe<TextObject>(cRCV_EUROBILL, "TextCVAmount") is TextObject textObject_CVAmount)
                    textObject_CVAmount.Text = totalBillAmount.ToString("N2");

                if (GetReportObjectSafe<TextObject>(cRCV_EUROBILL, "TextCompanyName") is TextObject textObject_CompanyName && comboBox_Company?.SelectedItem != null)
                    textObject_CompanyName.Text = comboBox_Company.SelectedItem.ToString();

                AccessToDatabase_EURO accessToDb = new AccessToDatabase_EURO();
                var signatories = accessToDb.RetrieveAllSignatoryData();

                if (GetReportObjectSafe<TextObject>(cRCV_EUROBILL, "TextPreparedBy") is TextObject textObject_PreparedBy)
                    textObject_PreparedBy.Text = signatories.PreparedByName;
                if (GetReportObjectSafe<TextObject>(cRCV_EUROBILL, "TextPreparedByPosition") is TextObject textObject_PreparedByPos)
                    textObject_PreparedByPos.Text = signatories.PreparedByPosition;
                if (GetReportObjectSafe<TextObject>(cRCV_EUROBILL, "TextCheckedBy") is TextObject textObject_CheckedBy)
                    textObject_CheckedBy.Text = signatories.ReviewedByName;
                if (GetReportObjectSafe<TextObject>(cRCV_EUROBILL, "TextCheckedByPosition") is TextObject textObject_CheckedByPos)
                    textObject_CheckedByPos.Text = signatories.ReviewedByPosition;
                if (GetReportObjectSafe<TextObject>(cRCV_EUROBILL, "TextApprovedBy") is TextObject textObject_ApprovedBy)
                    textObject_ApprovedBy.Text = signatories.ApprovedByName;
                if (GetReportObjectSafe<TextObject>(cRCV_EUROBILL, "TextApprovedByPosition") is TextObject textObject_ApprovedByPos)
                    textObject_ApprovedByPos.Text = signatories.ApprovedByPosition;
                if (GetReportObjectSafe<TextObject>(cRCV_EUROBILL, "TextReceivedBy") is TextObject textObject_ReceivedBy)
                    textObject_ReceivedBy.Text = signatories.ReceivedByName;
                if (GetReportObjectSafe<TextObject>(cRCV_EUROBILL, "TextReceivedByPosition") is TextObject textObject_ReceivedByPos)
                    textObject_ReceivedByPos.Text = signatories.ReceivedByPosition;

                string rawBank = bills[0].BankAccount ?? "";
                string bank = rawBank.Contains(":") ? rawBank.Split(':').Last().Trim() : rawBank;

                if (GetReportObjectSafe<TextObject>(cRCV_EUROBILL, "TextCVBILLNumber") is TextObject textObject_CVBILLNumber)
                    textObject_CVBILLNumber.Text = bills[0].RefNumber;
                if (GetReportObjectSafe<TextObject>(cRCV_EUROBILL, "TextCVCheckNum") is TextObject textObject_CVCheckNum)
                    textObject_CVCheckNum.Text = bills[0].RefNumber;

                if (GetReportObjectSafe<TextObject>(cRCV_EUROBILL, "TextCVBILLBank") is TextObject textObject_CVBILLBank)
                    textObject_CVBILLBank.Text = bank;
                if (GetReportObjectSafe<TextObject>(cRCV_EUROBILL, "TextCVCheckBank") is TextObject textObject_CVCheckBank)
                    textObject_CVCheckBank.Text = bank;

                if (GetReportObjectSafe<TextObject>(cRCV_EUROBILL, "TextCVBILLCheckDate") is TextObject textObject_CVBILLCheckDate)
                    textObject_CVBILLCheckDate.Text = bills[0].DateCreated.ToString("MMMM dd, yyyy");
                if (GetReportObjectSafe<TextObject>(cRCV_EUROBILL, "TextCVCheckDate") is TextObject textObject_CVCheckDate)
                    textObject_CVCheckDate.Text = bills[0].DateCreated.ToString("MMMM dd, yyyy");

                if (GetReportObjectSafe<TextObject>(cRCV_EUROBILL, "TextCVBILLDuePayment") is TextObject textObject_CVBILLDuePayment)
                    textObject_CVBILLDuePayment.Text = totalBillAmount.ToString("N2");
                if (GetReportObjectSafe<TextObject>(cRCV_EUROBILL, "TextCVDuePayment") is TextObject textObject_CVDuePayment)
                    textObject_CVDuePayment.Text = totalBillAmount.ToString("N2");

                double debitTotalAmount = 0;
                double creditTotalAmount = 0;

                foreach (var bill in bills)
                {
                    foreach (var item in bill.ItemDetails)
                    {
                        if (item.ExpenseLineAmount > 0) debitTotalAmount += item.ExpenseLineAmount;
                        else if (item.ExpenseLineAmount < 0) creditTotalAmount += Math.Abs(item.ExpenseLineAmount);

                        if (item.ItemLineAmount > 0) debitTotalAmount += item.ItemLineAmount;
                        else if (item.ItemLineAmount < 0) creditTotalAmount += Math.Abs(item.ItemLineAmount);
                    }
                }

                if (GetReportObjectSafe<TextObject>(cRCV_EUROBILL, "TextCVBILLTotalDebitAmount") is TextObject textObject_CVBILLTotalDebitAmount)
                    textObject_CVBILLTotalDebitAmount.Text = $"PHP {debitTotalAmount:N2}";
                if (GetReportObjectSafe<TextObject>(cRCV_EUROBILL, "TextCVTotalDebitAmount") is TextObject textObject_CVTotalDebitAmount)
                    textObject_CVTotalDebitAmount.Text = $"PHP {debitTotalAmount:N2}";

                if (GetReportObjectSafe<TextObject>(cRCV_EUROBILL, "TextCVBILLTotalCreditAmount") is TextObject textObject_CVBILLTotalCreditAmount)
                    textObject_CVBILLTotalCreditAmount.Text = $"PHP {debitTotalAmount:N2}";
                if (GetReportObjectSafe<TextObject>(cRCV_EUROBILL, "TextCVTotalCreditAmount") is TextObject textObject_CVTotalCreditAmount)
                    textObject_CVTotalCreditAmount.Text = $"PHP {debitTotalAmount:N2}";

                // Subreport population
                SubreportObject subreportObject = GetReportObjectSafe<SubreportObject>(cRCV_EUROBILL, "SubreportCVBILLDetailsIVP") ??
                                                  GetReportObjectSafe<SubreportObject>(cRCV_EUROBILL, "SubreportCVDetailsIVP");
                if (subreportObject != null)
                {
                    ReportDocument subReportDocument = cRCV_EUROBILL.OpenSubreport(subreportObject.SubreportName);
                    if (GetReportObjectSafe<TextObject>(subReportDocument, "TextRemarks") is TextObject textObject_Remarks)
                        textObject_Remarks.Text = bills[0].Memo;
                    if (GetReportObjectSafe<TextObject>(subReportDocument, "TextBILLRemarks") is TextObject textObject_BILLRemarks)
                        textObject_BILLRemarks.Text = bills[0].Memo;

                    if (GetReportObjectSafe<TextObject>(subReportDocument, "TextSubAccountPayable") is TextObject textObject_SubAccountPayable)
                        textObject_SubAccountPayable.Text = "Cash in Bank - " + bank;
                    if (GetReportObjectSafe<TextObject>(subReportDocument, "TextBILLSubAccountPayable") is TextObject textObject_BILLSubAccountPayable)
                        textObject_BILLSubAccountPayable.Text = "Cash in Bank - " + bank;

                    if (GetReportObjectSafe<TextObject>(subReportDocument, "TextSubAmountPayable") is TextObject textObject_SubAmountPayable)
                        textObject_SubAmountPayable.Text = debitTotalAmount.ToString("N2");
                    if (GetReportObjectSafe<TextObject>(subReportDocument, "TextBILLSubAmountPayable") is TextObject textObject_BILLSubAmountPayable)
                        textObject_BILLSubAmountPayable.Text = debitTotalAmount.ToString("N2");

                    if (GetReportObjectSafe<TextObject>(subReportDocument, "TextSubAccountCode") is TextObject textObject_SubAccountCode)
                        textObject_SubAccountCode.Text = "";
                    if (GetReportObjectSafe<TextObject>(subReportDocument, "TextBILLSubAccountCode") is TextObject textObject_BILLSubAccountCode)
                        textObject_BILLSubAccountCode.Text = "";
                }

                InsertDataToCheckVoucherDirectBillEURO(refNumberCR, bills);

                cRCV_EUROBILL.SetParameterValue("ReferenceNumber", refNumberCR);

                panel_Printing.Visible = false;
                panel_Signatory.Visible = true;
                panel_Main.Visible = false;
                panel_Main_CR.Visible = true;

                reportViewer.ReportSource = cRCV_EUROBILL;
                reportViewer.RefreshReport();

                return true;
            }
            catch (Exception ex)
            {
                MessageBox.Show($"Error generating Bill Payment Report: {ex.Message}", "Error", MessageBoxButtons.OK, MessageBoxIcon.Error);
                return false;
            }
        }

        private void InsertDataToCheckVoucherCompiledEURO(string refNumber, List<CheckTableExpensesAndItems> checkData)
        {
            string accessConnectionString = AccessToDatabase_EURO.GetAccessConnectionString();

            using (OleDbConnection connection = new OleDbConnection(accessConnectionString))
            {
                try
                {
                    connection.Open();

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

                    string insertQuery = @"
                        INSERT INTO CheckVoucherCompiled 
                        (RefNumber, [AccountNumber], [Particulars], [Class], [Debit], [Credit], [Memo], [CustomerJob]) 
                        VALUES 
                        (@RefNumber, @AccountNumber, @Particulars, @Class, @Debit, @Credit, @Memo, @CustomerJob)";

                    foreach (var check in checkData)
                    {
                        try
                        {
                            string memoValue = string.IsNullOrEmpty(check.ExpensesMemo) ? "" : check.ExpensesMemo;
                            string customerJob = string.IsNullOrEmpty(check.ExpensesCustomerJob) ? "" : check.ExpensesCustomerJob;

                            if (!string.IsNullOrEmpty(check.Item))
                            {
                                string itemName = check.Item;
                                string itemClass = check.ItemClass;
                                double itemAmount = check.ItemAmount;

                                string debit = itemAmount > 0 ? itemAmount.ToString("N2") : "";
                                string credit = itemAmount < 0 ? Math.Abs(itemAmount).ToString("N2") : "";

                                using (OleDbCommand command = new OleDbCommand(insertQuery, connection))
                                {
                                    command.Parameters.AddWithValue("@RefNumber", refNumber ?? (object)DBNull.Value);
                                    command.Parameters.AddWithValue("@AccountNumber", DBNull.Value);
                                    command.Parameters.AddWithValue("@Particulars", itemName);
                                    command.Parameters.AddWithValue("@Class", string.IsNullOrEmpty(itemClass) ? (object)DBNull.Value : itemClass);
                                    command.Parameters.AddWithValue("@Debit", debit);
                                    command.Parameters.AddWithValue("@Credit", credit);
                                    command.Parameters.AddWithValue("@Memo", memoValue);
                                    command.Parameters.AddWithValue("@CustomerJob", customerJob);

                                    command.ExecuteNonQuery();
                                }
                            }

                            if (!string.IsNullOrEmpty(check.Account))
                            {
                                string accountNumber = check.AccountNumber;
                                string expenseName = check.Account;
                                string expenseClass = check.ExpenseClass;
                                double expenseAmount = check.ExpensesAmount;

                                string debit = expenseAmount > 0 ? expenseAmount.ToString("N2") : "";
                                string credit = expenseAmount < 0 ? Math.Abs(expenseAmount).ToString("N2") : "";

                                using (OleDbCommand command = new OleDbCommand(insertQuery, connection))
                                {
                                    command.Parameters.AddWithValue("@RefNumber", refNumber ?? (object)DBNull.Value);
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

                            if (string.IsNullOrEmpty(check.Item) && string.IsNullOrEmpty(check.Account))
                            {
                                using (OleDbCommand command = new OleDbCommand(insertQuery, connection))
                                {
                                    command.Parameters.AddWithValue("@RefNumber", refNumber ?? (object)DBNull.Value);
                                    command.Parameters.AddWithValue("@AccountNumber", DBNull.Value);
                                    command.Parameters.AddWithValue("@Particulars", DBNull.Value);
                                    command.Parameters.AddWithValue("@Class", DBNull.Value);
                                    command.Parameters.AddWithValue("@Debit", DBNull.Value);
                                    command.Parameters.AddWithValue("@Credit", DBNull.Value);
                                    command.Parameters.AddWithValue("@Memo", memoValue);
                                    command.Parameters.AddWithValue("@CustomerJob", customerJob);

                                    command.ExecuteNonQuery();
                                }
                            }
                        }
                        catch (Exception ex)
                        {
                            MessageBox.Show($"Error inserting row: {ex.Message}", "Error", MessageBoxButtons.OK, MessageBoxIcon.Error);
                        }
                    }

                    connection.Close();
                }
                catch (Exception ex)
                {
                    MessageBox.Show($"Database connection error: {ex.Message}", "Error", MessageBoxButtons.OK, MessageBoxIcon.Error);
                }
            }
        }

        private static string SafeTrunc(string value, int maxLength)
        {
            if (string.IsNullOrEmpty(value)) return "";
            return value.Length <= maxLength ? value : value.Substring(0, maxLength);
        }

        private void InsertDataToJournalCompiled(string refNumber, List<JournalGridItem> journalItems)
        {
            string accessConnectionString = AccessToDatabase_EURO.GetAccessConnectionString();

            using (OleDbConnection connection = new OleDbConnection(accessConnectionString))
            {
                try
                {
                    connection.Open();

                    string deleteQuery = "DELETE FROM JV_Compiled";
                    using (OleDbCommand deleteCommand = new OleDbCommand(deleteQuery, connection))
                    {
                        try
                        {
                            deleteCommand.ExecuteNonQuery();
                        }
                        catch (Exception ex)
                        {
                            MessageBox.Show($"Error deleting data from JV_Compiled: {ex.Message}", "Error", MessageBoxButtons.OK, MessageBoxIcon.Error);
                            return;
                        }
                    }

                    string insertQuery = @"
                    INSERT INTO JV_Compiled 
                    (RefNumber, [Particulars], [Class], [Name], [Debit], [Credit], [Memo]) 
                    VALUES 
                    (?, ?, ?, ?, ?, ?, ?)";

                    foreach (var line in journalItems)
                    {
                        try
                        {
                            string particulars = SafeTrunc(line.AccountName, 255);
                            string className = line.Class;
                            string nameValue = SafeTrunc(line.Name, 255);
                            string memoValue = SafeTrunc(line.Memo, 255);

                            string debitStr = line.Debit != 0 ? line.Debit.ToString("N2") : "";
                            string creditStr = line.Credit != 0 ? line.Credit.ToString("N2") : "";

                            using (OleDbCommand command = new OleDbCommand(insertQuery, connection))
                            {
                                command.Parameters.Add("?", OleDbType.VarWChar).Value = refNumber ?? "";
                                command.Parameters.Add("?", OleDbType.VarWChar).Value = particulars;
                                command.Parameters.Add("?", string.IsNullOrEmpty(className) ? (object)DBNull.Value : className);
                                command.Parameters.Add("?", string.IsNullOrEmpty(nameValue) ? (object)DBNull.Value : nameValue);
                                command.Parameters.Add("?", OleDbType.VarWChar).Value = debitStr;
                                command.Parameters.Add("?", OleDbType.VarWChar).Value = creditStr;
                                command.Parameters.Add("?", OleDbType.VarWChar).Value = memoValue;

                                command.ExecuteNonQuery();
                            }
                        }
                        catch (Exception ex)
                        {
                            MessageBox.Show($"Error inserting journal line: {ex.Message}", "Error", MessageBoxButtons.OK, MessageBoxIcon.Error);
                        }
                    }

                    connection.Close();
                }
                catch (Exception ex)
                {
                    MessageBox.Show($"Error inserting journal compiled data: {ex.Message}", "Database Error", MessageBoxButtons.OK, MessageBoxIcon.Error);
                }
            }
        }

        private void InsertDataToCheckVoucherDirectBillEURO(string refNumber, List<BillTable> bills)
        {
            string accessConnectionString = AccessToDatabase_EURO.GetAccessConnectionString();

            using (OleDbConnection connection = new OleDbConnection(accessConnectionString))
            {
                try
                {
                    connection.Open();

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

                    string insertQuery = @"
                    INSERT INTO CheckVoucherCompiled 
                    (RefNumber, [AccountNumber], [Particulars], [Class], [Debit], [Credit], [Memo], [CustomerJob], [BillAppliedRefNumber], [BillDate], [BillAmountDue]) 
                    VALUES 
                    (?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?)";

                    foreach (var bill in bills)
                    {
                        foreach (var item in bill.ItemDetails)
                        {
                            using (OleDbCommand command = new OleDbCommand(insertQuery, connection))
                            {
                                string debit = "";
                                string credit = "";
                                string particular = "";
                                string itemClass = "";
                                string customerJob = "";
                                string memo = "";
                                string accountNum = "";

                                if (item.ExpenseLineAmount != 0 || !string.IsNullOrEmpty(item.ExpenseLineItemRefFullName))
                                {
                                    particular = item.ExpenseLineItemRefFullName ?? "";
                                    itemClass = item.ExpenseLineClassRefFullName ?? "";
                                    customerJob = item.ExpenseLineCustomerJob ?? "";
                                    memo = item.ExpenseLineMemo ?? "";
                                    accountNum = item.ExpenseLineAccountNumber ?? "";

                                    debit = item.ExpenseLineAmount > 0 ? item.ExpenseLineAmount.ToString("N2") : "";
                                    credit = item.ExpenseLineAmount < 0 ? Math.Abs(item.ExpenseLineAmount).ToString("N2") : "";
                                }
                                else if (item.ItemLineAmount != 0 || !string.IsNullOrEmpty(item.ItemLineItemRefFullName))
                                {
                                    particular = item.ItemLineItemRefFullName ?? "";
                                    itemClass = item.ItemLineClassRefFullName ?? "";
                                    customerJob = item.ItemLineCustomerJob ?? "";
                                    memo = item.ItemLineMemo ?? "";

                                    debit = item.ItemLineAmount > 0 ? item.ItemLineAmount.ToString("N2") : "";
                                    credit = item.ItemLineAmount < 0 ? Math.Abs(item.ItemLineAmount).ToString("N2") : "";
                                }

                                command.Parameters.Add("?", OleDbType.VarWChar).Value = refNumber;
                                command.Parameters.Add("?", string.IsNullOrWhiteSpace(accountNum) ? (object)DBNull.Value : accountNum);
                                command.Parameters.Add("?", OleDbType.VarWChar).Value = particular;
                                command.Parameters.Add("?", string.IsNullOrWhiteSpace(itemClass) ? (object)DBNull.Value : itemClass);
                                command.Parameters.Add("?", debit);
                                command.Parameters.Add("?", credit);
                                command.Parameters.Add("?", memo);
                                command.Parameters.Add("?", customerJob);
                                command.Parameters.Add("?", bill.AppliedRefNumber ?? (object)DBNull.Value);
                                command.Parameters.Add("?", bill.DateCreated.ToString("MM/dd/yyyy"));
                                command.Parameters.Add("?", bill.AmountDue.ToString("N2"));

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
                    AccessQueries_EURO queries = new AccessQueries_EURO();
                    checkivp = new List<CheckTableGrid>();
                    object data = null;

                    if (comboBox_Forms.SelectedIndex == 2) // Check
                    {
                        checkivp = queries.GetCheckDataEURO(refNumber);
                        data = checkivp;
                    }

                    if (data is System.Collections.ICollection collection && collection.Count > 0)
                    {
                        Layouts_EURO layouts_EURO = new Layouts_EURO();
                        System.Drawing.Printing.PaperSize paperSize = new System.Drawing.Printing.PaperSize("Custom", 850, 1100);
                        printDocument = new PrintDocument();
                        printDocument.DefaultPageSettings.PaperSize = paperSize;
                        printDocument.PrinterSettings.DefaultPageSettings.PaperSize = paperSize;

                        int selectedIndex = comboBox_Forms.SelectedIndex;
                        string seriesNumberStr = textBox_SeriesNumber.Text;
                        string payeeOverride = textBox_PayeeOverride.Text;

                        pageCounter = 1;
                        printPreviewControl.StartPage = 0;

                        printDocument.PrintPage += (s, ev) =>
                        {
                            layouts_EURO.PrintPage_EURO(s, ev, selectedIndex, seriesNumberStr, data, payeeOverride);
                        };

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
            comboBox_Signatory.Items.AddRange(new string[]
            {
                "-- Select Signatory Role --",
                "Prepared By:",
                "Checked By:",
                "Approved By:",
                "Received By:"
            });
            comboBox_Signatory.SelectedIndex = 0;
            panel_Signatory.Controls.Add(comboBox_Signatory);

            Label label_SignatoryName = new Label
            {
                Text = "Full Name:",
                Font = FontInputLabel,
                ForeColor = ColorTextMuted,
                Width = cardInnerWidth,
                Height = 18,
                TextAlign = ContentAlignment.MiddleLeft,
                Margin = new Padding(0, 2, 0, 2)
            };
            panel_Signatory.Controls.Add(label_SignatoryName);

            TextBox textBox_SignatoryName = CreateModernTextBox(cardInnerWidth);
            panel_Signatory.Controls.Add(textBox_SignatoryName);

            Label label_SignatoryPosition = new Label
            {
                Text = "Position / Title:",
                Font = FontInputLabel,
                ForeColor = ColorTextMuted,
                Width = cardInnerWidth,
                Height = 18,
                TextAlign = ContentAlignment.MiddleLeft,
                Margin = new Padding(0, 2, 0, 2)
            };
            panel_Signatory.Controls.Add(label_SignatoryPosition);

            TextBox textBox_SignatoryPosition = CreateModernTextBox(cardInnerWidth);
            panel_Signatory.Controls.Add(textBox_SignatoryPosition);

            FlowLayoutPanel saveRow = new FlowLayoutPanel
            {
                Width = cardInnerWidth,
                Height = 32,
                FlowDirection = FlowDirection.LeftToRight,
                WrapContents = false,
                Margin = new Padding(0, 2, 0, 0)
            };

            Button button_SaveSignatory = CreateModernButton("💾  SAVE", ColorSuccessBtn, ColorSuccessHover, Color.White, 90, 28);
            saveRow.Controls.Add(button_SaveSignatory);

            Label label_SignatoryStatus = new Label
            {
                Width = cardInnerWidth - 96,
                Height = 28,
                TextAlign = ContentAlignment.MiddleLeft,
                Font = new Font("Segoe UI", 8.5f, FontStyle.Bold),
                ForeColor = ColorSuccessBtn,
                Margin = new Padding(6, 0, 0, 0),
                Tag = "StatusSuccess"
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

                Timer clearTimer = new Timer { Interval = 2500 };
                clearTimer.Tick += (ts, te) =>
                {
                    label_SignatoryStatus.Text = "";
                    clearTimer.Stop();
                    clearTimer.Dispose();
                };
                clearTimer.Start();
            };

            comboBox_Signatory.SelectedIndexChanged += (sender, e) =>
            {
                label_SignatoryStatus.Text = "";
                if (comboBox_Signatory.SelectedIndex == 0)
                {
                    textBox_SignatoryName.Text = "";
                    textBox_SignatoryPosition.Text = "";
                }
                else
                {
                    int choice = comboBox_Signatory.SelectedIndex;
                    var signatoryData = accessToDatabase.GetSignatoryData(choice);
                    textBox_SignatoryName.Text = signatoryData.Name;
                    textBox_SignatoryPosition.Text = signatoryData.Position;
                }
            };

            panel_Signatory.Controls.Add(saveRow);
            return panel_Signatory;
        }

        private FlowLayoutPanel Panel_SBPrinting()
        {
            int cardInnerWidth = sideBarWidth - 44;
            panel_Printing = CreateCardPanel("🖨️  PRINTING CONTROLS", sideBarWidth - 24);
            panel_Printing.Visible = false;

            int halfBtnWidth = (cardInnerWidth - 6) / 2;

            FlowLayoutPanel zoomRow = new FlowLayoutPanel
            {
                Width = cardInnerWidth,
                Height = 32,
                FlowDirection = FlowDirection.LeftToRight,
                WrapContents = false,
                Margin = new Padding(0, 0, 0, 4)
            };

            Button button_ZoomOut = CreateModernButton("🔍- Zoom Out", CurrentSecondaryBtn, CurrentSecondaryHover, CurrentSecondaryText, halfBtnWidth, 28, null, "Secondary");
            button_ZoomOut.Margin = new Padding(0, 0, 6, 0);
            button_ZoomOut.Click += (sender, e) =>
            {
                if (printPreviewControl.Zoom >= 0.2) printPreviewControl.Zoom -= 0.1;
            };
            zoomRow.Controls.Add(button_ZoomOut);

            Button button_ZoomIn = CreateModernButton("🔍+ Zoom In", CurrentSecondaryBtn, CurrentSecondaryHover, CurrentSecondaryText, halfBtnWidth, 28, null, "Secondary");
            button_ZoomIn.Margin = new Padding(0);
            button_ZoomIn.Click += (sender, e) => { printPreviewControl.Zoom += 0.1; };
            zoomRow.Controls.Add(button_ZoomIn);

            panel_Printing.Controls.Add(zoomRow);

            FlowLayoutPanel pageRow = new FlowLayoutPanel
            {
                Width = cardInnerWidth,
                Height = 32,
                FlowDirection = FlowDirection.LeftToRight,
                WrapContents = false,
                Margin = new Padding(0, 0, 0, 6)
            };

            Button button_PreviousPage = CreateModernButton("◀ Prev Page", CurrentSecondaryBtn, CurrentSecondaryHover, CurrentSecondaryText, halfBtnWidth, 28, null, "Secondary");
            button_PreviousPage.Margin = new Padding(0, 0, 6, 0);
            button_PreviousPage.Click += (sender, e) =>
            {
                if (printPreviewControl.StartPage > 0) printPreviewControl.StartPage--;
            };
            pageRow.Controls.Add(button_PreviousPage);

            Button button_NextPage = CreateModernButton("Next Page ▶", CurrentSecondaryBtn, CurrentSecondaryHover, CurrentSecondaryText, halfBtnWidth, 28, null, "Secondary");
            button_NextPage.Margin = new Padding(0);
            button_NextPage.Click += (sender, e) =>
            {
                if (printPreviewControl.StartPage < pageCounter - 1) printPreviewControl.StartPage++;
            };
            pageRow.Controls.Add(button_NextPage);

            panel_Printing.Controls.Add(pageRow);

            Button button_Print = CreateModernButton("🖨️  PRINT DOCUMENT", ColorSuccessBtn, ColorSuccessHover, Color.White, cardInnerWidth, 34, new Font("Segoe UI", 9.5f, FontStyle.Bold));
            button_Print.Margin = new Padding(0, 2, 0, 2);
            button_Print.Click += (sender, e) =>
            {
                try
                {
                    pageCounter = 1;
                    printPreviewControl.StartPage = 0;

                    PrintDialog printDialog = new PrintDialog
                    {
                        Document = printDocument
                    };

                    if (printDialog.ShowDialog() == DialogResult.OK)
                    {
                        GlobalVariables.includeImage = false;
                        printDialog.Document.Print();

                        printPreviewControl.Visible = false;
                        printPreviewControl.Zoom = 1;
                        panel_Printing.Visible = false;

                        string formType = "";
                        if (comboBox_Forms.SelectedIndex == 1) formType = "CV";
                        else if (comboBox_Forms.SelectedIndex == 3) formType = "JV";

                        if (formType != "")
                        {
                            string selectedCompany = comboBox_Company.SelectedItem?.ToString() ?? "EURO-PACIFIC HEALTH CARE DISTRIBUTOR INC.";
                            if (!string.IsNullOrEmpty(selectedCompany))
                            {
                                seriesNumber++;
                                accessToDatabase.UpdateManualSeriesNumber(formType, seriesNumber, selectedCompany);
                                UpdateSeriesNumberEURO(formType);
                            }
                        }
                    }
                }
                catch (Exception ex)
                {
                    MessageBox.Show($"An error occurred while printing: {ex.Message}", "Print Error", MessageBoxButtons.OK, MessageBoxIcon.Error);
                }
                GlobalVariables.includeImage = true;
            };
            panel_Printing.Controls.Add(button_Print);

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

            if (comboBox_Forms.SelectedIndex == 1 || comboBox_Forms.SelectedIndex == 3)
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

            if (prefix != "")
            {
                string selectedCompany = comboBox_Company.SelectedItem?.ToString() ?? "EURO-PACIFIC HEALTH CARE DISTRIBUTOR INC.";
                if (!string.IsNullOrEmpty(selectedCompany))
                {
                    seriesNumber = accessToDatabase.GetSeriesNumberFromDatabase(prefix, selectedCompany);
                    UpdateSeriesNumberEURO(prefix);
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

        private void TextBox_SeriesNumber_TextChanged(object sender, EventArgs e)
        {
            if (!string.IsNullOrEmpty(textBox_SeriesNumber.Text))
            {
                string formPrefix = "";
                if (comboBox_Forms.SelectedIndex == 1) formPrefix = "CV";
                else if (comboBox_Forms.SelectedIndex == 3) formPrefix = "JV";

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

        private void TextBox_SeriesNumber_Leave(object sender, EventArgs e)
        {
            string formType = "";
            if (comboBox_Forms.SelectedIndex == 1) formType = "CV";
            else if (comboBox_Forms.SelectedIndex == 3) formType = "JV";

            if (!string.IsNullOrEmpty(formType) && comboBox_Company.SelectedItem != null)
            {
                accessToDatabase.UpdateManualSeriesNumber(formType, seriesNumber, comboBox_Company.SelectedItem.ToString());
            }
        }

        private void UpdateSeriesNumberEURO(string formPrefix)
        {
            if (accessToDatabase == null) accessToDatabase = new AccessToDatabase_EURO();
            textBox_SeriesNumber.Text = $"{formPrefix}-{seriesNumber:00000}";
        }
    }
}

