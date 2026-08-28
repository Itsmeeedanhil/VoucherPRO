using System;
using System.Collections.Generic;

namespace VoucherPROVER2.Clients.EURO
{
    public class Dataclass_EURO
    {
        public class BillTable // For Bill Payment Check
        {
            public DateTime DateCreated { get; set; }
            public string PayeeFullName { get; set; }
            public string TermsRefFullName { get; set; }
            public string BankAccount { get; set; }
            public string APAccountRefFullName { get; set; }
            public double Amount { get; set; }
            public string RefNumber { get; set; }
            public string Address { get; set; }
            public string Address2 { get; set; }

            public string AppliedRefNumber { get; set; }
            public string AppliedToTxnTxnID { get; set; }

            public DateTime DueDate { get; set; }
            public double AmountDue { get; set; }
            public string Memo { get; set; }
            public string BillMemo { get; set; }
            public string AccountNumber { get; set; }

            public List<ItemDetail> ItemDetails { get; set; }

            public BillTable()
            {
                ItemDetails = new List<ItemDetail>();
            }
        }

        public class ItemDetail
        {
            public string ItemLineItemRefFullName { get; set; }
            public double ItemLineAmount { get; set; }
            public string ItemLineClassRefFullName { get; set; }
            public string ItemLineMemo { get; set; }
            public string ItemLineCustomerJob { get; set; }

            public string ExpenseLineItemRefFullName { get; set; }
            public string ExpenseLineAccountNumber { get; set; }
            public double ExpenseLineAmount { get; set; }
            public string ExpenseLineClassRefFullName { get; set; }
            public string ExpenseLineCustomerJob { get; set; }
            public string ExpenseLineMemo { get; set; }
        }

        public class CheckTableExpensesAndItems // For Print Check Voucher
        {
            public DateTime DateCreated { get; set; }
            public string BankAccount { get; set; }
            public string PayeeFullName { get; set; }
            public string RefNumber { get; set; }
            public double TotalAmount { get; set; }
            public string Address { get; set; }
            public string Address2 { get; set; }
            public string Memo { get; set; }

            public string AddressBlockAddr1 { get; set; }
            public string AddressBlockAddr2 { get; set; }
            public string AddressBlockAddr3 { get; set; }
            public string AddressBlockAddr4 { get; set; }
            public string AddressCity { get; set; }
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
            public string AccountNumber { get; set; }
            public double ExpensesAmount { get; set; }
            public string ExpensesMemo { get; set; }
            public string ExpensesCustomerJob { get; set; }
            public string ExpenseClass { get; set; }

            public ItemType ItemType { get; set; }
        }

        public enum ItemType
        {
            Item,
            Expense,
            Transaction
        }

        public class JournalGridItem
        {
            public string AccountName { get; set; }
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

