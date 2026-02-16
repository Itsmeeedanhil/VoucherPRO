using System;
using System.Collections.Generic;
using System.Drawing.Printing;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using static VoucherPROVER2.Clients.IVP.AccessToDatabase_IVP;
using static VoucherPROVER2.Clients.IVP.Dataclass_IVP;

namespace VoucherPROVER2.Clients.IVP
{
    public class Layouts_IVP
    {
        Font font_Ten = new Font("Microsoft Sans Serif", 10, FontStyle.Regular);
        Font font_Nine = new Font("Microsoft Sans Serif", 9, FontStyle.Regular);
        Font font_Eleven = new Font("Microsoft Sans Serif", 11, FontStyle.Regular);
        Font font_Eight = new Font("Microsoft Sans Serif", 8, FontStyle.Regular);

        public void PrintPage_IVP(object sender, PrintPageEventArgs e, int layoutIndex, string seriesNumber, object data, string payeeOverride = "")
        {
            StringFormat sfAlignRight = new StringFormat { Alignment = StringAlignment.Far, LineAlignment = StringAlignment.Far };
            StringFormat sfAlignCenterRight = new StringFormat { Alignment = StringAlignment.Far, LineAlignment = StringAlignment.Center };
            StringFormat sfAlignCenter = new StringFormat { Alignment = StringAlignment.Center, LineAlignment = StringAlignment.Center };
            StringFormat sfAlignLeftCenter = new StringFormat { Alignment = StringAlignment.Near, LineAlignment = StringAlignment.Center };

            switch (layoutIndex)
            {
                case 2: // CHECK (Index 2 for IVP)
                    // Use CheckTableGrid here because that is what GetCheckDataIVP returns
                    Layout_Check(e, data as List<CheckTableGrid>, sfAlignCenterRight, sfAlignCenter, sfAlignLeftCenter, payeeOverride);
                    break;

                default:
                    // If you add CV (Index 1) later, add case 1 here.
                    break;
            }
        }

        private void Layout_Check(PrintPageEventArgs e, List<CheckTableGrid> checkTableData, StringFormat sfAlignCenterRight, StringFormat sfAlignCenter, StringFormat sfAlignLeftCenter, string payeeOverride)
        {
            // Safety Check
            if (checkTableData == null || checkTableData.Count == 0) return;

            DateTime dateCreated = Convert.ToDateTime(checkTableData[0].DateCreated).Date;

            string month = dateCreated.ToString("MM");
            string day = dateCreated.ToString("dd");
            string year = dateCreated.ToString("yyyy");

            // Format Date with spaces
            string formattedMonth = string.Join("   ", month.ToCharArray());
            string formattedDay = string.Join("   ", day.ToCharArray());
            string formattedYear = string.Join("   ", year.ToCharArray());
            string formattedDate = $"{formattedMonth}     {formattedDay}     {formattedYear}";

            string payee = checkTableData[0].PayeeFullName.ToString();

            // If the textbox is not empty, use that instead
            if (!string.IsNullOrEmpty(payeeOverride))
            {
                payee = payeeOverride;
            }

            double amount = checkTableData[0].Amount;
            string amountInWords = AmountToWordsConverter.Convert(amount);

            Font amountinWordsFont = font_Eight;
            Font dateFont = font_Nine;
            Font payeeFont = font_Eleven;
            Font payeeFont2 = font_Ten;

            // ----------------------------------------------------------------------
            // PRINTING LOGIC (Coordinates)
            // ----------------------------------------------------------------------
            if (GlobalVariables.isPrinting)
            {
                // ROTATED PRINTING (For specific printers like EPSON LX-310 depending on paper feed)
                e.Graphics.RotateTransform(-90);
                e.Graphics.TranslateTransform(-e.MarginBounds.Height + 180, 0 - 70);

                e.Graphics.DrawString(payee, payeeFont, Brushes.Black, new PointF(60, 410));
                e.Graphics.DrawString(formattedDate, dateFont, Brushes.Black, new PointF(530, 380));
                e.Graphics.DrawString(amount.ToString("N2"), dateFont, Brushes.Black, new PointF(550, 38 + 345 + 30));
                e.Graphics.DrawString(amountInWords, amountinWordsFont, Brushes.Black, new PointF(25, 430 + 15));
            }
            else
            {
                // PREVIEW MODE / NORMAL PRINTING
                // Adjust "minusX" and "minusY" to shift the whole block left/up
                int minusX = 30;
                int minusY = 50;

                // Payee Name
                e.Graphics.DrawString(payee, payeeFont2, Brushes.Black, new PointF(135 - minusX, 110 - minusY));

                // Date
                e.Graphics.DrawString(formattedDate, payeeFont, Brushes.Black, new PointF(605 - minusX, 79 - minusY));

                // Amount (Number)
                e.Graphics.DrawString(amount.ToString("N2"), payeeFont, Brushes.Black, new PointF(635 - minusX, 114 - minusY));

                // Amount (Words)
                e.Graphics.DrawString(amountInWords, payeeFont2, Brushes.Black, new PointF(95 - minusX, 145 - minusY));
            }
        }
    }
}
