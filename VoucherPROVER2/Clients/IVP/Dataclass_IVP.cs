using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace VoucherPROVER2.Clients.IVP
{
    public class Dataclass_IVP
    {
        public class BillTable // For APV / Bill
        {
            // BillPaymentCheckLine table
            public DateTime DateCreated { get; set; } //TimeCreated
            public string PayeeFullName { get; set; } //PayeeEntityRefFullName
            public string TermsRefFullName { get; set; } //TermsRefFullName
            public string BankAccount { get; set; } //BankAccountRefFullName
            public string APAccountRefFullName { get; set; } //APAccountRefFullName
            public double Amount { get; set; }
            public string RefNumber { get; set; }
            public string Address { get; set; } //AddressAddr
            public string Address2 { get; set; } //AddressAddr2

            public string AppliedRefNumber { get; set; } //AppliedToTxnRefNumber
            public string AppliedToTxnTxnID { get; set; } //AppliedToTxnTxnID

            // Bill table
            public string Vendor { get; set; } //VendorRefFullName
            public string VendorAddressAddr1 { get; set; } //VendorRefFullName
            public string VendorAddressAddr2 { get; set; } //VendorRefFullName
            public string VendorAddressAddr3 { get; set; } //VendorRefFullName
            public string VendorAddressAddr4 { get; set; } //VendorRefFullName
            public string VendorAddressCity { get; set; } //VendorRefFullName
            public string Currency { get; set; } //VendorRefFullName
            public double Exchangerate { get; set; } //VendorRefFullName

            public DateTime DueDate { get; set; }
            public double AmountDue { get; set; }
            public string Memo { get; set; }
            public string BillMemo { get; set; }
            public string Tin { get; set; }
            public string POnumber { get; set; }
            public DateTime DateCreatedHistory { get; set; }
            public string MemoHistory { get; set; }
            public double AmountHistory { get; set; }
            public string RefNumberHistory { get; set; }
            public string HistoryCVNumber { get; set; }
            public string HistoryAPVNumber { get; set; }
            public string Remarks { get; set; }
            public string AccountName { get; set; }
            public string AccountNumber { get; set; }
            public bool IsPaid { get; set; }
            public List<string> AccountNameParticularsList { get; set; }
            public List<string> AccountNumberParticularsList { get; set; }

            public List<ItemDetail> ItemDetails { get; set; }

            public BillTable()
            {
                ItemDetails = new List<ItemDetail>();
                AccountNameParticularsList = new List<string>();
                AccountNumberParticularsList = new List<string>();
            }

            //Increment
            public string IncrementalID { get; set; }
        }

        public class ItemDetail
        {
            public string ItemLineItemRefFullName { get; set; }
            public double ItemLineAmount { get; set; }
            public string ItemLineClassRefFullName { get; set; }
            public string ItemLineMemo { get; set; }
            public string ItemLineCustomerJob { get; set; }


            public string ExpenseLineItemRefFullName { get; set; }
            public double ExpenseLineAmount { get; set; }
            public string ExpenseLineClassRefFullName { get; set; }
            public string ExpenseLineCustomerJob { get; set; }
            public string ExpenseLineMemo { get; set; }
        }

        public class APVData // For Grouping Data in APV LEADS
        {
            public double Amount { get; set; }
            public string Class { get; set; }
        }

        public class CheckTable // For Print Check
        {
            public DateTime DateCreated { get; set; } //TimeCreated
            public string RefNumber { get; set; }
            public double Amount { get; set; }
            public string PayeeFullName { get; set; }
        }

        public class CheckTableExpensesAndItems // For Print Check Voucher
        {
            //Check table
            public DateTime DateCreated { get; set; }
            public string BankAccount { get; set; }
            public string PayeeFullName { get; set; }
            public string RefNumber { get; set; }
            public double TotalAmount { get; set; }
            public string Address { get; set; }
            public string Address2 { get; set; }
            public string Memo { get; set; }
            public string IncrementalID { get; set; }

            public DateTime DueDate { get; set; }

            // Properties specific to items
            public string Item { get; set; }
            public string ItemName { get; set; }
            public string ItemDescription { get; set; }
            public string ItemClass { get; set; }
            public double ItemAmount { get; set; }

            // Properties specific to expenses
            public string Account { get; set; }
            public string AccountName { get; set; }
            public string AccountNameCheck { get; set; }
            public string AccountNumber { get; set; }
            public string BankAccountNumber { get; set; }
            public double ExpensesAmount { get; set; }
            public string ExpensesMemo { get; set; }
            public string ExpensesCustomerJob { get; set; }
            public string ExpenseClass { get; set; }
            public string TinID { get; set; }
            public string POnumber { get; set; }
            public string TxnID { get; set; }
            public DateTime DateCreatedHistory { get; set; }
            public string AssetAccountNumber { get; set; }
            public string AssetAccountName { get; set; }
            public string MemoHistory { get; set; }
            public double AmountHistory { get; set; }
            public string RefNumberHistory { get; set; }
            public string HistoryCVNumber { get; set; }
            public string HistoryAPVNumber { get; set; }
            public string Remarks { get; set; }

            public string AccountClassCheck { get; set; }

            // Indicates whether it's an item or an expense
            public ItemType ItemType { get; set; }
        }

        public enum ItemType
        {
            Item,
            Expense,
            Transaction
        }

        public class ItemReciept
        {
            public string PayeeFullName { get; set; }
            public string Addr1 { get; set; }
            public string Addr2 { get; set; }
            public string Addr3 { get; set; }
            public string Addr4 { get; set; }
            public string AddrCity { get; set; }
            public string RefNumber { get; set; }
            public DateTime DateCreated { get; set; }
            public string Memo { get; set; }
            public double TotalAmount { get; set; }
            public string BankAccount { get; set; }
            public string Account { get; set; }
            public string ExpensesMemo { get; set; }
            public string ItemDescription { get; set; }
            public string ItemClass { get; set; }
            public string ItemCustomerJob { get; set; }
            public string Item { get; set; }
            public string ItemUM { get; set; }
            public double ItemQuantity { get; set; }
            public double ItemCost { get; set; }
            public double ExpensesAmount { get; set; }
            public double ItemAmount { get; set; }
            public ReceiptItemType ReceiptItemType { get; set; }
        }
        public enum ReceiptItemType
        {
            ReceiptItem,
            RecieptExpense
        }

        public class CR_APV_LEADS
        {
            public string Particular { get; set; }
            public string Class { get; set; }
            public string Debit { get; set; }
            public string Credit { get; set; }
        }


        public class JournalGridItem
        {
            // The "Row Header" in your screenshot (e.g., Accounts Receivable, Sales, etc.)
            public string AccountName { get; set; }

            // The Columns
            public string Type { get; set; } = "General Journal";
            public DateTime Date { get; set; }
            public string Num { get; set; }
            public string Name { get; set; }      // EntityRef (The Customer/Vendor)
            public string Memo { get; set; }      // Line Memo
            public string Class { get; set; }     // ClassRef

            // Amounts
            public double Debit { get; set; }
            public double Credit { get; set; }
        }

        public class CheckTableGrid
        {
            public DateTime DateCreated { get; set; }
            public string RefNumber { get; set; }
            public double Amount { get; set; }
            public string PayeeFullName { get; set; }
        }
    }
}
