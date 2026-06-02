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
using VoucherPROVER2.Clients.MPM;
using VoucherPROVER2.Clients.OWI;

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

                return panel;
            }
            if (GlobalVariables.client == "MPM")
            {
                // 1. Instantiate the specific dashboard class
                Dashboard_MPM dashboard_MPM = new Dashboard_MPM();

                // 2. Call the method that returns the panel
                Panel mpmContent = dashboard_MPM.ContainerPanel();

                // 3. Add that panel into the current panel's controls
                panel.Controls.Add(mpmContent);

                return panel;
            }
            if (GlobalVariables.client == "OWI")
            {
                // 1. Instantiate the specific dashboard class
                Dashboard_OWI dashboard_OWI = new Dashboard_OWI();

                // 2. Call the method that returns the panel
                Panel mpmContent = dashboard_OWI.ContainerPanel();

                // 3. Add that panel into the current panel's controls
                panel.Controls.Add(mpmContent);

                return panel;
            }
            else
            {
                throw new NotImplementedException("Client not implemented: " + GlobalVariables.client);
            }
        }
    }
}
