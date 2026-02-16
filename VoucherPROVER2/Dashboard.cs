using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;
using VoucherPROVER2.Clients.IVP;

namespace VoucherPROVER2
{
    public class GlobalVariables
    {
        public static string client = "IVP";
        public static bool includeImage = true;
        public static bool includeItemReceipt = true;
        public static bool testWithoutData = true;
        public static bool isPrinting = false;
        public static bool useCrystalReports_LEADS = true;
        public static int itemsPerPageAPV = 10;
    }
    public partial class Dashboard : Form
    {
        public Dashboard()
        {
            InitializeComponent();

            this.WindowState = FormWindowState.Maximized;
            this.Text = "VoucherPro";

            Panel panel = ContainerPanel();
            this.Controls.Add(panel);
        }

        private Panel ContainerPanel()
        {
            Panel panel = new Panel
            {
                Dock = DockStyle.Fill,
            };

            if (GlobalVariables.client == "IVP")
            {
                // 1. Instantiate the specific dashboard class
                Dashboard_IVP dashboard_IVP = new Dashboard_IVP();

                // 2. Call the method that returns the panel
                Panel ivpContent = dashboard_IVP.ContainerPanel();

                // 3. Add that panel into the current panel's controls
                panel.Controls.Add(ivpContent);
            }
            else
            {
                throw new NotImplementedException("Client not implemented: " + GlobalVariables.client);
            }

            return panel;
        }
    }
}
