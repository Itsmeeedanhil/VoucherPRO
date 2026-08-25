using QBFC16Lib;
using System;
using System.Collections.Generic;
using System.Windows.Forms;
using static VoucherPROVER2.Clients.EURO.Dataclass_EURO;

namespace VoucherPROVER2.Clients.EURO
{
    public class AccessQueries_EURO
    {
        public List<CheckTableGrid> GetCheckDataEURO(string refNumber)
        {
            List<CheckTableGrid> checkList = new List<CheckTableGrid>();
            QBSessionManager sessionManager = new QBSessionManager();

            try
            {
                sessionManager.OpenConnection2("", "VoucherPro Check Data", ENConnectionType.ctLocalQBD);
                sessionManager.BeginSession("", ENOpenMode.omDontCare);

                IMsgSetRequest request = sessionManager.CreateMsgSetRequest("US", 13, 0);
                request.Attributes.OnError = ENRqOnError.roeContinue;

                // 1. QUERY FOR REGULAR CHECKS
                ICheckQuery checkQuery = request.AppendCheckQueryRq();
                checkQuery.ORTxnQuery.TxnFilter.ORRefNumberFilter.RefNumberFilter.MatchCriterion.SetValue(ENMatchCriterion.mcStartsWith);
                checkQuery.ORTxnQuery.TxnFilter.ORRefNumberFilter.RefNumberFilter.RefNumber.SetValue(refNumber);

                // 2. QUERY FOR BILL PAYMENT CHECKS
                IBillPaymentCheckQuery billPayQuery = request.AppendBillPaymentCheckQueryRq();
                billPayQuery.ORTxnQuery.TxnFilter.ORRefNumberFilter.RefNumberFilter.MatchCriterion.SetValue(ENMatchCriterion.mcStartsWith);
                billPayQuery.ORTxnQuery.TxnFilter.ORRefNumberFilter.RefNumberFilter.RefNumber.SetValue(refNumber);

                // Execute Requests
                IMsgSetResponse response = sessionManager.DoRequests(request);

                // PROCESS RESPONSE 1: REGULAR CHECKS
                IResponse qbResponseCheck = response.ResponseList.GetAt(0);
                ICheckRetList checkRetList = qbResponseCheck.Detail as ICheckRetList;

                if (checkRetList != null)
                {
                    for (int i = 0; i < checkRetList.Count; i++)
                    {
                        ICheckRet checkRet = checkRetList.GetAt(i);
                        string docNum = checkRet.RefNumber.GetValue();

                        if (docNum != refNumber) continue;

                        CheckTableGrid newCheck = new CheckTableGrid
                        {
                            DateCreated = checkRet.TxnDate.GetValue().Date,
                            RefNumber = docNum,
                            Amount = checkRet.Amount.GetValue(),
                            PayeeFullName = checkRet.PayeeEntityRef != null ? checkRet.PayeeEntityRef.FullName.GetValue() : "No Payee"
                        };
                        checkList.Add(newCheck);
                    }
                }

                // PROCESS RESPONSE 2: BILL PAYMENT CHECKS
                IResponse qbResponseBillPay = response.ResponseList.GetAt(1);
                IBillPaymentCheckRetList billPayRetList = qbResponseBillPay.Detail as IBillPaymentCheckRetList;

                if (billPayRetList != null)
                {
                    for (int i = 0; i < billPayRetList.Count; i++)
                    {
                        IBillPaymentCheckRet billPayRet = billPayRetList.GetAt(i);
                        string docNum = billPayRet.RefNumber.GetValue();

                        if (docNum != refNumber) continue;

                        CheckTableGrid newCheck = new CheckTableGrid
                        {
                            DateCreated = billPayRet.TxnDate.GetValue().Date,
                            RefNumber = docNum,
                            Amount = billPayRet.Amount.GetValue(),
                            PayeeFullName = billPayRet.PayeeEntityRef != null ? billPayRet.PayeeEntityRef.FullName.GetValue() : "No Payee"
                        };
                        checkList.Add(newCheck);
                    }
                }
            }
            catch (Exception ex)
            {
                MessageBox.Show($"Error retrieving data from QuickBooks: {ex.Message}", "Error", MessageBoxButtons.OK, MessageBoxIcon.Error);
            }
            finally
            {
                if (sessionManager != null)
                {
                    try
                    {
                        sessionManager.EndSession();
                        sessionManager.CloseConnection();
                    }
                    catch { }
                }
            }

            return checkList;
        }

        public List<CheckTableExpensesAndItems> GetCheckExpensesAndItemsData_EURO(string refNumber)
        {
            QBSessionManager sessionManager = new QBSessionManager();
            List<CheckTableExpensesAndItems> checks = new List<CheckTableExpensesAndItems>();

            try
            {
                string AppName = "QuickBooks Check Retrieval";
                sessionManager.OpenConnection2("", AppName, ENConnectionType.ctLocalQBD);
                sessionManager.BeginSession("", ENOpenMode.omDontCare);

                // Build request
                IMsgSetRequest request = sessionManager.CreateMsgSetRequest("US", 13, 0);
                request.Attributes.OnError = ENRqOnError.roeContinue;

                ICheckQuery checkQuery = request.AppendCheckQueryRq();

                // Filter by RefNumber
                checkQuery.ORTxnQuery.TxnFilter.ORRefNumberFilter.RefNumberFilter.MatchCriterion
                    .SetValue(ENMatchCriterion.mcStartsWith);

                checkQuery.ORTxnQuery.TxnFilter.ORRefNumberFilter.RefNumberFilter.RefNumber
                    .SetValue(refNumber);

                // Include line items
                checkQuery.IncludeLineItems.SetValue(true);

                IMsgSetResponse response = sessionManager.DoRequests(request);
                IResponse qbResponse = response.ResponseList.GetAt(0);

                ICheckRetList list = qbResponse.Detail as ICheckRetList;

                if (list == null || list.Count == 0)
                {
                    return checks;
                }

                // FETCH ALL ACCOUNT NUMBERS FROM QUICKBOOKS
                Dictionary<string, string> accountNumbersDict = GetAccountNumbersFromQB(sessionManager);

                for (int i = 0; i < list.Count; i++)
                {
                    ICheckRet check = list.GetAt(i);

                    // HEADER DATA
                    DateTime txnDate = check.TxnDate?.GetValue() ?? DateTime.MinValue;
                    string bankAccount = check.AccountRef?.FullName?.GetValue() ?? "";
                    string payee = check.PayeeEntityRef?.FullName?.GetValue() ?? "";
                    string memo = check.Memo?.GetValue() ?? "";
                    string address1 = check.Address?.Addr1?.GetValue() ?? "";
                    string address2 = check.Address?.Addr2?.GetValue() ?? "";
                    string address3 = check.Address?.Addr3?.GetValue() ?? "";
                    string address4 = check.Address?.Addr4?.GetValue() ?? "";
                    string addressCity = check.Address?.City?.GetValue() ?? "";
                    double totalAmount = check.Amount?.GetValue() ?? 0;

                    // EXPENSE LINES
                    if (check.ExpenseLineRetList != null)
                    {
                        for (int e = 0; e < check.ExpenseLineRetList.Count; e++)
                        {
                            IExpenseLineRet exp = check.ExpenseLineRetList.GetAt(e);

                            string expAccount = exp.AccountRef?.FullName?.GetValue() ?? "";
                            string expListID = exp.AccountRef?.ListID?.GetValue() ?? "";
                            double expAmount = exp.Amount?.GetValue() ?? 0;

                            string accNumber = "";
                            if (!string.IsNullOrEmpty(expListID) && accountNumbersDict.ContainsKey(expListID))
                            {
                                accNumber = accountNumbersDict[expListID];
                            }
                            else if (!string.IsNullOrEmpty(expAccount) && accountNumbersDict.ContainsKey(expAccount))
                            {
                                accNumber = accountNumbersDict[expAccount];
                            }

                            checks.Add(new CheckTableExpensesAndItems
                            {
                                DateCreated = txnDate,
                                BankAccount = bankAccount,
                                PayeeFullName = payee,
                                RefNumber = refNumber,
                                TotalAmount = totalAmount,
                                DueDate = txnDate,
                                Memo = memo,
                                AddressBlockAddr1 = address1,
                                AddressBlockAddr2 = address2,
                                AddressBlockAddr3 = address3,
                                AddressBlockAddr4 = address4,
                                AddressCity = addressCity,

                                Account = expAccount,
                                AccountNumber = accNumber,
                                ExpenseClass = exp.ClassRef?.FullName?.GetValue() ?? "",
                                ExpensesAmount = expAmount,
                                ExpensesMemo = exp.Memo?.GetValue() ?? "",
                                ExpensesCustomerJob = exp.CustomerRef?.FullName?.GetValue() ?? "",

                                ItemType = ItemType.Expense
                            });
                        }
                    }

                    // ITEM LINES
                    if (check.ORItemLineRetList != null)
                    {
                        for (int iLine = 0; iLine < check.ORItemLineRetList.Count; iLine++)
                        {
                            IORItemLineRet orItemLine = (IORItemLineRet)check.ORItemLineRetList.GetAt(iLine);

                            if (orItemLine.ItemLineRet != null)
                            {
                                IItemLineRet item = orItemLine.ItemLineRet;

                                string itemName = item.ItemRef?.FullName?.GetValue() ?? "";
                                double itemAmount = item.Amount?.GetValue() ?? 0;

                                checks.Add(new CheckTableExpensesAndItems
                                {
                                    DateCreated = txnDate,
                                    BankAccount = bankAccount,
                                    PayeeFullName = payee,
                                    RefNumber = refNumber,
                                    TotalAmount = totalAmount,
                                    DueDate = txnDate,
                                    Memo = memo,
                                    AddressBlockAddr1 = address1,
                                    AddressBlockAddr2 = address2,
                                    AddressBlockAddr3 = address3,
                                    AddressBlockAddr4 = address4,
                                    AddressCity = addressCity,

                                    Item = itemName,
                                    ItemDescription = item.Desc?.GetValue() ?? "",
                                    ItemClass = item.ClassRef?.FullName?.GetValue() ?? "",
                                    ItemAmount = itemAmount,

                                    ItemType = ItemType.Item
                                });
                            }
                        }
                    }
                }
            }
            catch (Exception ex)
            {
                MessageBox.Show($"Error: {ex.Message}");
            }
            finally
            {
                try { sessionManager.EndSession(); sessionManager.CloseConnection(); }
                catch { }
            }

            return checks;
        }

        public List<BillTable> GetBillData_EURO(string refNumber)
        {
            QBSessionManager sessionManager = new QBSessionManager();
            List<BillTable> bills = new List<BillTable>();

            try
            {
                sessionManager.OpenConnection2("", "Bill Retrieval", ENConnectionType.ctLocalQBD);
                sessionManager.BeginSession("", ENOpenMode.omDontCare);

                Dictionary<string, string> accountNumbersDict = GetAccountNumbersFromQBBILL(sessionManager);

                // 1. QUERY BILL PAYMENT CHECK USING RefNumber
                IMsgSetRequest req1 = sessionManager.CreateMsgSetRequest("US", 13, 0);
                req1.Attributes.OnError = ENRqOnError.roeContinue;

                IBillPaymentCheckQuery bpcQuery = req1.AppendBillPaymentCheckQueryRq();
                bpcQuery.IncludeLineItems.SetValue(true);

                bpcQuery.ORTxnQuery.TxnFilter.ORRefNumberFilter.RefNumberFilter.MatchCriterion.SetValue(ENMatchCriterion.mcStartsWith);
                bpcQuery.ORTxnQuery.TxnFilter.ORRefNumberFilter.RefNumberFilter.RefNumber.SetValue(refNumber);

                IMsgSetResponse resp1 = sessionManager.DoRequests(req1);
                IResponse r1 = resp1.ResponseList.GetAt(0);

                IBillPaymentCheckRetList bpList = r1.Detail as IBillPaymentCheckRetList;

                if (bpList == null || bpList.Count == 0)
                {
                    MessageBox.Show("Bill Payment Check not found: " + refNumber);
                    return bills;
                }

                IBillPaymentCheckRet bp = bpList.GetAt(0);

                DateTime payDate = bp.TxnDate?.GetValue() ?? DateTime.MinValue;
                string payee = bp.PayeeEntityRef?.FullName?.GetValue() ?? "";
                string address1 = bp.Address?.Addr1?.GetValue() ?? "";
                string address2 = bp.Address?.Addr2?.GetValue() ?? "";
                string bankAccount = bp.BankAccountRef?.FullName?.GetValue() ?? "";
                string memo = bp.Memo?.GetValue() ?? "";
                double amountPaid = bp.Amount?.GetValue() ?? 0;

                // GET ALL APPLIED BILL TxnIDs
                List<string> appliedTxnIDs = new List<string>();

                if (bp.AppliedToTxnRetList != null && bp.AppliedToTxnRetList.Count > 0)
                {
                    for (int k = 0; k < bp.AppliedToTxnRetList.Count; k++)
                    {
                        var applied = bp.AppliedToTxnRetList.GetAt(k);
                        string tId = applied.TxnID?.GetValue();
                        if (!string.IsNullOrEmpty(tId))
                        {
                            appliedTxnIDs.Add(tId);
                        }
                    }
                }
                else
                {
                    MessageBox.Show("No Applied Bill found from Bill Payment Check.");
                    return bills;
                }

                // 2. QUERY BILL(S) USING THE COLLECTED TxnIDs
                IMsgSetRequest req2 = sessionManager.CreateMsgSetRequest("US", 13, 0);
                req2.Attributes.OnError = ENRqOnError.roeContinue;

                IBillQuery billQuery = req2.AppendBillQueryRq();
                billQuery.IncludeLineItems.SetValue(true);

                foreach (string id in appliedTxnIDs)
                {
                    billQuery.ORBillQuery.TxnIDList.Add(id);
                }

                IMsgSetResponse resp2 = sessionManager.DoRequests(req2);
                IResponse r2 = resp2.ResponseList.GetAt(0);

                IBillRetList billList = r2.Detail as IBillRetList;

                if (billList == null || billList.Count == 0)
                {
                    MessageBox.Show("Bills not found for the provided TxnIDs.");
                    return bills;
                }

                for (int bIndex = 0; bIndex < billList.Count; bIndex++)
                {
                    IBillRet bill = billList.GetAt(bIndex);

                    DateTime dueDate = bill.DueDate?.GetValue() ?? DateTime.MinValue;
                    double amountDue = bill.AmountDue?.GetValue() ?? 0;
                    string billMemo = bill.Memo?.GetValue() ?? "";
                    string billAPAccount = bill.APAccountRef?.FullName?.GetValue() ?? "";
                    string billRefNumber = bill.RefNumber?.GetValue() ?? "";
                    string specificTxnID = bill.TxnID?.GetValue() ?? "";

                    string apListID = bill.APAccountRef?.ListID?.GetValue() ?? "";
                    string billAccNum = "";
                    if (!string.IsNullOrEmpty(apListID) && accountNumbersDict.ContainsKey(apListID))
                    {
                        billAccNum = accountNumbersDict[apListID];
                    }
                    else if (!string.IsNullOrEmpty(billAPAccount) && accountNumbersDict.ContainsKey(billAPAccount))
                    {
                        billAccNum = accountNumbersDict[billAPAccount];
                    }

                    BillTable bt = new BillTable
                    {
                        DateCreated = payDate,
                        DueDate = payDate,
                        PayeeFullName = payee,
                        Address = address1,
                        Address2 = address2,
                        BankAccount = bankAccount,
                        APAccountRefFullName = billAPAccount,
                        AccountNumber = billAccNum,
                        Amount = amountPaid,
                        RefNumber = refNumber,
                        AppliedRefNumber = billRefNumber,
                        AppliedToTxnTxnID = specificTxnID,
                        Memo = memo,
                        BillMemo = billMemo,
                        AmountDue = amountDue,
                    };

                    if (bill.ExpenseLineRetList != null)
                    {
                        for (int i = 0; i < bill.ExpenseLineRetList.Count; i++)
                        {
                            var exp = bill.ExpenseLineRetList.GetAt(i);
                            string expAccountName = exp.AccountRef?.FullName?.GetValue() ?? "";
                            string expListID = exp.AccountRef?.ListID?.GetValue() ?? "";

                            string expAccNumber = "";
                            if (!string.IsNullOrEmpty(expListID) && accountNumbersDict.ContainsKey(expListID))
                            {
                                expAccNumber = accountNumbersDict[expListID];
                            }
                            else if (!string.IsNullOrEmpty(expAccountName) && accountNumbersDict.ContainsKey(expAccountName))
                            {
                                expAccNumber = accountNumbersDict[expAccountName];
                            }

                            bt.ItemDetails.Add(new ItemDetail
                            {
                                ExpenseLineItemRefFullName = expAccountName,
                                ExpenseLineAccountNumber = expAccNumber,
                                ExpenseLineAmount = exp.Amount?.GetValue() ?? 0,
                                ExpenseLineClassRefFullName = exp.ClassRef?.FullName?.GetValue() ?? "",
                                ExpenseLineCustomerJob = exp.CustomerRef?.FullName?.GetValue() ?? "",
                                ExpenseLineMemo = exp.Memo?.GetValue() ?? "",
                            });
                        }
                    }

                    if (bill.ORItemLineRetList != null)
                    {
                        for (int i = 0; i < bill.ORItemLineRetList.Count; i++)
                        {
                            var orItem = bill.ORItemLineRetList.GetAt(i);
                            if (orItem.ItemLineRet != null)
                            {
                                var item = orItem.ItemLineRet;
                                bt.ItemDetails.Add(new ItemDetail
                                {
                                    ItemLineItemRefFullName = item.ItemRef?.FullName?.GetValue() ?? "",
                                    ItemLineAmount = item.Amount?.GetValue() ?? 0,
                                    ItemLineClassRefFullName = item.ClassRef?.FullName?.GetValue() ?? "",
                                    ItemLineCustomerJob = item.CustomerRef?.FullName?.GetValue() ?? "",
                                    ItemLineMemo = item.Desc?.GetValue() ?? "",
                                });
                            }
                        }
                    }

                    bills.Add(bt);
                }
            }
            catch (Exception ex)
            {
                MessageBox.Show("Error retrieving Bill data: " + ex.Message);
            }
            finally
            {
                try
                {
                    sessionManager.EndSession();
                    sessionManager.CloseConnection();
                }
                catch { }
            }

            return bills;
        }

        public List<JournalGridItem> GetJournalEntryForGrid(string refNumber)
        {
            QBSessionManager sessionManager = new QBSessionManager();
            List<JournalGridItem> gridItems = new List<JournalGridItem>();

            try
            {
                sessionManager.OpenConnection2("", "QB Journal Grid", ENConnectionType.ctLocalQBD);
                sessionManager.BeginSession("", ENOpenMode.omDontCare);

                IMsgSetRequest request = sessionManager.CreateMsgSetRequest("US", 13, 0);
                request.Attributes.OnError = ENRqOnError.roeContinue;

                IJournalEntryQuery jeQuery = request.AppendJournalEntryQueryRq();

                jeQuery.ORTxnQuery.TxnFilter.ORRefNumberFilter.RefNumberFilter.MatchCriterion.SetValue(ENMatchCriterion.mcStartsWith);
                jeQuery.ORTxnQuery.TxnFilter.ORRefNumberFilter.RefNumberFilter.RefNumber.SetValue(refNumber);
                jeQuery.IncludeLineItems.SetValue(true);

                IMsgSetResponse response = sessionManager.DoRequests(request);
                IResponse qbResponse = response.ResponseList.GetAt(0);
                IJournalEntryRetList list = qbResponse.Detail as IJournalEntryRetList;

                if (list != null)
                {
                    for (int i = 0; i < list.Count; i++)
                    {
                        IJournalEntryRet je = list.GetAt(i);
                        string docNum = je.RefNumber.GetValue();

                        if (docNum != refNumber)
                        {
                            continue;
                        }

                        DateTime date = je.TxnDate.GetValue();

                        if (je.ORJournalLineList != null)
                        {
                            for (int j = 0; j < je.ORJournalLineList.Count; j++)
                            {
                                IORJournalLine orLine = je.ORJournalLineList.GetAt(j);
                                JournalGridItem item = new JournalGridItem
                                {
                                    Date = date,
                                    Num = docNum,
                                    Type = "General Journal"
                                };

                                if (orLine.JournalDebitLine != null)
                                {
                                    var line = orLine.JournalDebitLine;
                                    item.AccountName = line.AccountRef?.FullName?.GetValue() ?? "";
                                    item.Name = line.EntityRef?.FullName?.GetValue() ?? "";
                                    item.Memo = Truncate(line.Memo?.GetValue() ?? "", 500);
                                    item.Class = line.ClassRef?.FullName?.GetValue() ?? "";
                                    item.Debit = line.Amount?.GetValue() ?? 0;
                                    item.Credit = 0;
                                }
                                else if (orLine.JournalCreditLine != null)
                                {
                                    var line = orLine.JournalCreditLine;
                                    item.AccountName = line.AccountRef?.FullName?.GetValue() ?? "";
                                    item.Name = line.EntityRef?.FullName?.GetValue() ?? "";
                                    item.Memo = Truncate(line.Memo?.GetValue() ?? "", 500);
                                    item.Class = line.ClassRef?.FullName?.GetValue() ?? "";
                                    item.Debit = 0;
                                    item.Credit = line.Amount?.GetValue() ?? 0;
                                }

                                gridItems.Add(item);
                            }
                        }

                        break;
                    }
                }
            }
            catch (Exception ex)
            {
                MessageBox.Show($"Error retrieving journal entry: {ex.Message}");
            }
            finally
            {
                try { sessionManager.EndSession(); sessionManager.CloseConnection(); } catch { }
            }

            return gridItems;
        }

        private Dictionary<string, string> GetAccountNumbersFromQB(QBSessionManager sessionManager)
        {
            Dictionary<string, string> accountDict = new Dictionary<string, string>();

            try
            {
                IMsgSetRequest req = sessionManager.CreateMsgSetRequest("US", 13, 0);
                req.Attributes.OnError = ENRqOnError.roeContinue;

                IAccountQuery accQuery = req.AppendAccountQueryRq();

                IMsgSetResponse res = sessionManager.DoRequests(req);
                IResponse qbRes = res.ResponseList.GetAt(0);

                IAccountRetList accList = qbRes.Detail as IAccountRetList;

                if (accList != null)
                {
                    for (int i = 0; i < accList.Count; i++)
                    {
                        IAccountRet acc = accList.GetAt(i);
                        string listID = acc.ListID?.GetValue() ?? "";
                        string fullName = acc.FullName?.GetValue() ?? "";
                        string accountNumber = acc.AccountNumber?.GetValue() ?? "";

                        if (!string.IsNullOrEmpty(accountNumber))
                        {
                            if (!string.IsNullOrEmpty(listID) && !accountDict.ContainsKey(listID))
                            {
                                accountDict.Add(listID, accountNumber);
                            }
                            if (!string.IsNullOrEmpty(fullName) && !accountDict.ContainsKey(fullName))
                            {
                                accountDict.Add(fullName, accountNumber);
                            }
                        }
                    }
                }
            }
            catch (Exception ex)
            {
                Console.WriteLine($"Error fetching Account Numbers: {ex.Message}");
            }

            return accountDict;
        }

        private Dictionary<string, string> GetAccountNumbersFromQBBILL(QBSessionManager sessionManager)
        {
            Dictionary<string, string> accountDict = new Dictionary<string, string>();

            try
            {
                IMsgSetRequest req = sessionManager.CreateMsgSetRequest("US", 13, 0);
                req.Attributes.OnError = ENRqOnError.roeContinue;

                IAccountQuery accQuery = req.AppendAccountQueryRq();

                IMsgSetResponse res = sessionManager.DoRequests(req);
                IResponse qbRes = res.ResponseList.GetAt(0);

                IAccountRetList accList = qbRes.Detail as IAccountRetList;

                if (accList != null)
                {
                    for (int i = 0; i < accList.Count; i++)
                    {
                        IAccountRet acc = accList.GetAt(i);
                        string listID = acc.ListID?.GetValue() ?? "";
                        string fullName = acc.FullName?.GetValue() ?? "";
                        string accountNumber = acc.AccountNumber?.GetValue() ?? "";

                        if (!string.IsNullOrEmpty(accountNumber))
                        {
                            if (!string.IsNullOrEmpty(listID) && !accountDict.ContainsKey(listID))
                            {
                                accountDict.Add(listID, accountNumber);
                            }
                            if (!string.IsNullOrEmpty(fullName) && !accountDict.ContainsKey(fullName))
                            {
                                accountDict.Add(fullName, accountNumber);
                            }
                        }
                    }
                }
            }
            catch (Exception ex)
            {
                Console.WriteLine($"Error fetching Account Numbers: {ex.Message}");
            }

            return accountDict;
        }

        private string Truncate(string value, int maxLength)
        {
            if (string.IsNullOrEmpty(value)) return "";
            return value.Length <= maxLength ? value : value.Substring(0, maxLength);
        }
    }
}

