using System;
using System.Collections.Generic;
using System.Drawing.Printing;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using static VoucherPROVER2.Clients.INT.AccessToDatabase_INT;
using static VoucherPROVER2.Clients.INT.Dataclass_INT;

namespace VoucherPROVER2.Clients.INT
{
    public class Layouts_INT
    {
        Font font_Ten = new Font("Microsoft Sans Serif", 10, FontStyle.Regular);
        Font font_Nine = new Font("Microsoft Sans Serif", 9, FontStyle.Regular);
        Font font_Eleven = new Font("Microsoft Sans Serif", 11, FontStyle.Regular);
        Font font_Eight = new Font("Microsoft Sans Serif", 8, FontStyle.Regular);

        public void PrintPage_INT(object sender, PrintPageEventArgs e, int layoutIndex, string seriesNumber, object data, string payeeOverride = "")
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

            // Individual digits for Month (MM), Day (DD), and Year (YYYY)
            string M1 = month[0].ToString();
            string M2 = month[1].ToString();

            string D1 = day[0].ToString();
            string D2 = day[1].ToString();

            string Y1 = year[0].ToString();
            string Y2 = year[1].ToString();
            string Y3 = year[2].ToString();
            string Y4 = year[3].ToString();

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

                // Payee Name
                e.Graphics.DrawString(payee, payeeFont, Brushes.Black, new PointF(60, 410));

                // --- DATE SECTION ---
                // 1. Month (MM)
                e.Graphics.DrawString(M1, dateFont, Brushes.Black, new PointF(530, 380));
                e.Graphics.DrawString(M2, dateFont, Brushes.Black, new PointF(548, 380));

                // 2. Day (DD)
                e.Graphics.DrawString(D1, dateFont, Brushes.Black, new PointF(575, 380));
                e.Graphics.DrawString(D2, dateFont, Brushes.Black, new PointF(593, 380));

                // 3. Year (YYYY)
                e.Graphics.DrawString(Y1, dateFont, Brushes.Black, new PointF(635, 380));
                e.Graphics.DrawString(Y2, dateFont, Brushes.Black, new PointF(653, 380));
                e.Graphics.DrawString(Y3, dateFont, Brushes.Black, new PointF(671, 380));
                e.Graphics.DrawString(Y4, dateFont, Brushes.Black, new PointF(689, 380));

                // Amount (Number)
                e.Graphics.DrawString(amount.ToString("N2"), dateFont, Brushes.Black, new PointF(550, 38 + 345 + 30));

                // Amount (Words)
                e.Graphics.DrawString(amountInWords, amountinWordsFont, Brushes.Black, new PointF(25, 430 + 15));
            }
            else
            {
                // PREVIEW MODE / NORMAL PRINTING
                int minusX = 30;
                int minusY = 50;

                // Payee Name
                e.Graphics.DrawString(payee, payeeFont2, Brushes.Black, new PointF(125 - minusX, 117 - minusY));

                // --- DATE SECTION ---
                // 1. Month (MM)
                e.Graphics.DrawString(M1, payeeFont, Brushes.Black, new PointF(603 - minusX, 79 - minusY));
                e.Graphics.DrawString(M2, payeeFont, Brushes.Black, new PointF(621 - minusX, 79 - minusY));

                // 2. Day (DD)
                e.Graphics.DrawString(D1, payeeFont, Brushes.Black, new PointF(648 - minusX, 79 - minusY));
                e.Graphics.DrawString(D2, payeeFont, Brushes.Black, new PointF(666 - minusX, 79 - minusY));

                // 3. Year (YYYY)
                e.Graphics.DrawString(Y1, payeeFont, Brushes.Black, new PointF(693 - minusX, 79 - minusY));
                e.Graphics.DrawString(Y2, payeeFont, Brushes.Black, new PointF(711 - minusX, 79 - minusY));
                e.Graphics.DrawString(Y3, payeeFont, Brushes.Black, new PointF(729 - minusX, 79 - minusY));
                e.Graphics.DrawString(Y4, payeeFont, Brushes.Black, new PointF(747 - minusX, 79 - minusY));

                // Amount (Number)
                e.Graphics.DrawString(amount.ToString("N2"), payeeFont, Brushes.Black, new PointF(635 - minusX, 117 - minusY));

                // Amount (Words)
                e.Graphics.DrawString(amountInWords, payeeFont2, Brushes.Black, new PointF(95 - minusX, 145 - minusY));
            }
        }
    }
}
