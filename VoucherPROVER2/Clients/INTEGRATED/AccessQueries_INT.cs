using QBFC16Lib;
using System;
using System.Collections.Generic;
using System.Data.OleDb;
using System.Linq;
using System.Text.RegularExpressions;
using System.Windows.Forms;
using static VoucherPROVER2.Clients.INT.Dataclass_INT;
using static VoucherPROVER2.Clients.INT.Dataclass_INT.BillTable;

namespace VoucherPROVER2.Clients.INT
{
    public class AccessQueries_INT : IDisposable
    {
        private QBSessionManager _sessionManager;
        private bool _isConnectionOpen = false;
        private bool _isSessionOpen = false;

        // Cached Chart of Accounts maps (loaded once per form session)
        private Dictionary<string, string> _cachedParentAccountMap;
        private Dictionary<string, string> _cachedAccountNumbersMap;

        private QBSessionManager GetSession()
        {
            if (_sessionManager == null)
            {
                _sessionManager = new QBSessionManager();
            }

            try
            {
                if (!_isConnectionOpen)
                {
                    _sessionManager.OpenConnection2("", "VoucherPro INT", ENConnectionType.ctLocalQBD);
                    _isConnectionOpen = true;
                }

                if (!_isSessionOpen)
                {
                    _sessionManager.BeginSession("", ENOpenMode.omDontCare);
                    _isSessionOpen = true;
                }
            }
            catch
            {
                Dispose();
                throw;
            }

            return _sessionManager;
        }

        public void ClearAccountCache()
        {
            _cachedParentAccountMap = null;
            _cachedAccountNumbersMap = null;
        }

        public void Dispose()
        {
            if (_sessionManager != null)
            {
                try
                {
                    if (_isSessionOpen)
                    {
                        _sessionManager.EndSession();
                    }
                }
                catch { }

                try
                {
                    if (_isConnectionOpen)
                    {
                        _sessionManager.CloseConnection();
                    }
                }
                catch { }

                _isSessionOpen = false;
                _isConnectionOpen = false;
                _sessionManager = null;
            }
        }

        // =========================================================================
        // 1. FAST CHECK & BILL PAYMENT LOOKUP (mcStartsWith + C# EXACT FILTER)
        // =========================================================================
        public List<CheckTableGrid> GetCheckDataINT(string refNumber)
        {
            List<CheckTableGrid> checkList = new List<CheckTableGrid>();
            string targetRef = refNumber.Trim();

            try
            {
                var sessionManager = GetSession();
                IMsgSetRequest request = sessionManager.CreateMsgSetRequest("US", 13, 0);
                request.Attributes.OnError = ENRqOnError.roeContinue;

                ICheckQuery checkQuery = request.AppendCheckQueryRq();
                checkQuery.IncludeLineItems.SetValue(false);
                checkQuery.ORTxnQuery.TxnFilter.ORRefNumberFilter.RefNumberFilter.MatchCriterion.SetValue(ENMatchCriterion.mcStartsWith);
                checkQuery.ORTxnQuery.TxnFilter.ORRefNumberFilter.RefNumberFilter.RefNumber.SetValue(targetRef);
                checkQuery.IncludeRetElementList.Add("TxnDate");
                checkQuery.IncludeRetElementList.Add("RefNumber");
                checkQuery.IncludeRetElementList.Add("Amount");
                checkQuery.IncludeRetElementList.Add("PayeeEntityRef");

                IBillPaymentCheckQuery billPayQuery = request.AppendBillPaymentCheckQueryRq();
                billPayQuery.IncludeLineItems.SetValue(false);
                billPayQuery.ORTxnQuery.TxnFilter.ORRefNumberFilter.RefNumberFilter.MatchCriterion.SetValue(ENMatchCriterion.mcStartsWith);
                billPayQuery.ORTxnQuery.TxnFilter.ORRefNumberFilter.RefNumberFilter.RefNumber.SetValue(targetRef);
                billPayQuery.IncludeRetElementList.Add("TxnDate");
                billPayQuery.IncludeRetElementList.Add("RefNumber");
                billPayQuery.IncludeRetElementList.Add("Amount");
                billPayQuery.IncludeRetElementList.Add("PayeeEntityRef");

                IMsgSetResponse response = sessionManager.DoRequests(request);

                ICheckRetList checkRetList = response.ResponseList.GetAt(0).Detail as ICheckRetList;
                if (checkRetList != null)
                {
                    for (int i = 0; i < checkRetList.Count; i++)
                    {
                        ICheckRet checkRet = checkRetList.GetAt(i);
                        string docNum = checkRet.RefNumber?.GetValue()?.Trim() ?? "";
                        if (!string.Equals(docNum, targetRef, StringComparison.OrdinalIgnoreCase)) continue;

                        checkList.Add(new CheckTableGrid
                        {
                            DateCreated = checkRet.TxnDate.GetValue().Date,
                            RefNumber = docNum,
                            Amount = checkRet.Amount.GetValue(),
                            PayeeFullName = checkRet.PayeeEntityRef != null ? checkRet.PayeeEntityRef.FullName.GetValue() : "No Payee"
                        });
                    }
                }

                IBillPaymentCheckRetList billPayRetList = response.ResponseList.GetAt(1).Detail as IBillPaymentCheckRetList;
                if (billPayRetList != null)
                {
                    for (int i = 0; i < billPayRetList.Count; i++)
                    {
                        IBillPaymentCheckRet billPayRet = billPayRetList.GetAt(i);
                        string docNum = billPayRet.RefNumber?.GetValue()?.Trim() ?? "";
                        if (!string.Equals(docNum, targetRef, StringComparison.OrdinalIgnoreCase)) continue;

                        checkList.Add(new CheckTableGrid
                        {
                            DateCreated = billPayRet.TxnDate.GetValue().Date,
                            RefNumber = docNum,
                            Amount = billPayRet.Amount.GetValue(),
                            PayeeFullName = billPayRet.PayeeEntityRef != null ? billPayRet.PayeeEntityRef.FullName.GetValue() : "No Payee"
                        });
                    }
                }
            }
            catch (Exception ex)
            {
                Dispose();
                MessageBox.Show($"Error retrieving data from QuickBooks: {ex.Message}", "Error", MessageBoxButtons.OK, MessageBoxIcon.Error);
            }

            return checkList;
        }

        // =========================================================================
        // 2. FAST CHECK VOUCHER (REGULAR CHECK) RETRIEVAL
        // =========================================================================
        public List<CheckTableExpensesAndItems> GetCheckExpensesAndItemsData_INT(string refNumber, out bool foundAnyMatchInQB)
        {
            foundAnyMatchInQB = false;
            List<CheckTableExpensesAndItems> checks = new List<CheckTableExpensesAndItems>();
            string targetRef = refNumber.Trim();

            try
            {
                var sessionManager = GetSession();

                // STAGE 1: Lightweight header-only search (NO line items yet!)
                IMsgSetRequest headerReq = sessionManager.CreateMsgSetRequest("US", 13, 0);
                headerReq.Attributes.OnError = ENRqOnError.roeContinue;

                ICheckQuery headerQuery = headerReq.AppendCheckQueryRq();
                headerQuery.IncludeLineItems.SetValue(false);
                headerQuery.ORTxnQuery.TxnFilter.ORRefNumberFilter.RefNumberFilter.MatchCriterion.SetValue(ENMatchCriterion.mcStartsWith);
                headerQuery.ORTxnQuery.TxnFilter.ORRefNumberFilter.RefNumberFilter.RefNumber.SetValue(targetRef);
                headerQuery.IncludeRetElementList.Add("TxnID");
                headerQuery.IncludeRetElementList.Add("RefNumber");

                IMsgSetResponse headerResp = sessionManager.DoRequests(headerReq);
                ICheckRetList list = headerResp.ResponseList.GetAt(0).Detail as ICheckRetList;

                if (list == null || list.Count == 0)
                {
                    return checks;
                }

                foundAnyMatchInQB = true;
                List<string> foundRefs = new List<string>();
                string exactTxnID = null;

                for (int idx = 0; idx < list.Count; idx++)
                {
                    var chk = list.GetAt(idx);
                    string rNum = chk.RefNumber?.GetValue()?.Trim() ?? "";

                    if (!string.IsNullOrEmpty(rNum) && !foundRefs.Contains(rNum))
                        foundRefs.Add(rNum);

                    if (string.Equals(rNum, targetRef, StringComparison.OrdinalIgnoreCase))
                    {
                        exactTxnID = chk.TxnID?.GetValue();
                    }
                }

                if (string.IsNullOrEmpty(exactTxnID))
                {
                    var sortedRefs = foundRefs
                        .OrderBy(r => long.TryParse(r, out long num) ? num : long.MaxValue)
                        .ThenBy(r => r)
                        .ToList();

                    MessageBox.Show(
                        $"Reference Number mismatch for '{targetRef}'!\n\n" +
                        $"No exact match found. QuickBooks returned the following check(s) in sequence:\n - {string.Join("\n - ", sortedRefs)}\n\n" +
                        "Please enter the exact reference number.",
                        "Reference Number Mismatch",
                        MessageBoxButtons.OK,
                        MessageBoxIcon.Warning
                    );
                    return checks;
                }

                // STAGE 2: Fetch full line items ONLY for the 1 exact TxnID
                IMsgSetRequest detailReq = sessionManager.CreateMsgSetRequest("US", 13, 0);
                detailReq.Attributes.OnError = ENRqOnError.roeContinue;

                ICheckQuery detailQuery = detailReq.AppendCheckQueryRq();
                detailQuery.IncludeLineItems.SetValue(true);
                detailQuery.ORTxnQuery.TxnIDList.Add(exactTxnID);

                IMsgSetResponse detailResp = sessionManager.DoRequests(detailReq);
                ICheckRetList detailList = detailResp.ResponseList.GetAt(0).Detail as ICheckRetList;

                if (detailList == null || detailList.Count == 0) return checks;

                ICheckRet check = detailList.GetAt(0);

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
                string currentRef = check.RefNumber?.GetValue() ?? "";

                if (check.ExpenseLineRetList != null)
                {
                    for (int e = 0; e < check.ExpenseLineRetList.Count; e++)
                    {
                        IExpenseLineRet exp = check.ExpenseLineRetList.GetAt(e);
                        checks.Add(new CheckTableExpensesAndItems
                        {
                            DateCreated = txnDate,
                            BankAccount = bankAccount,
                            PayeeFullName = payee,
                            RefNumber = currentRef,
                            TotalAmount = totalAmount,
                            DueDate = txnDate,
                            Memo = memo,
                            AddressBlockAddr1 = address1,
                            AddressBlockAddr2 = address2,
                            AddressBlockAddr3 = address3,
                            AddressBlockAddr4 = address4,
                            AddressCity = addressCity,
                            Account = exp.AccountRef?.FullName?.GetValue() ?? "",
                            ExpenseClass = exp.ClassRef?.FullName?.GetValue() ?? "",
                            ExpensesAmount = exp.Amount?.GetValue() ?? 0,
                            ExpensesMemo = exp.Memo?.GetValue() ?? "",
                            ExpensesCustomerJob = exp.CustomerRef?.FullName?.GetValue() ?? "",
                            ItemType = ItemType.Expense
                        });
                    }
                }

                if (check.ORItemLineRetList != null)
                {
                    for (int iLine = 0; iLine < check.ORItemLineRetList.Count; iLine++)
                    {
                        IORItemLineRet orItemLine = (IORItemLineRet)check.ORItemLineRetList.GetAt(iLine);
                        if (orItemLine.ItemLineRet != null)
                        {
                            IItemLineRet item = orItemLine.ItemLineRet;
                            checks.Add(new CheckTableExpensesAndItems
                            {
                                DateCreated = txnDate,
                                BankAccount = bankAccount,
                                PayeeFullName = payee,
                                RefNumber = currentRef,
                                TotalAmount = totalAmount,
                                DueDate = txnDate,
                                Memo = memo,
                                AddressBlockAddr1 = address1,
                                AddressBlockAddr2 = address2,
                                AddressBlockAddr3 = address3,
                                AddressBlockAddr4 = address4,
                                AddressCity = addressCity,
                                Item = item.ItemRef?.FullName?.GetValue() ?? "",
                                ItemDescription = item.Desc?.GetValue() ?? "",
                                ItemClass = item.ClassRef?.FullName?.GetValue() ?? "",
                                ItemAmount = item.Amount?.GetValue() ?? 0,
                                ItemType = ItemType.Item
                            });
                        }
                    }
                }
            }
            catch (Exception ex)
            {
                Dispose();
                MessageBox.Show($"Error: {ex.Message}");
            }

            return checks;
        }

        // =========================================================================
        // 3. FAST BILL PAYMENT CHECK RETRIEVAL (CACHED COA + TWO-STAGE)
        // =========================================================================
        public List<BillTable> GetBillData_INT(string refNumber)
        {
            List<BillTable> bills = new List<BillTable>();
            string targetRef = refNumber.Trim();

            try
            {
                var sessionManager = GetSession();

                // 0. Load Chart of Accounts ONCE and cache it
                var accountMap = GetParentAccountConsolidationMap(sessionManager);

                string ResolveToParent(string raw)
                {
                    if (string.IsNullOrWhiteSpace(raw)) return "";
                    string trimmed = raw.Trim();

                    if (accountMap.ContainsKey(trimmed))
                        return accountMap[trimmed];

                    if (trimmed.Contains(":"))
                    {
                        string rootPart = trimmed.Split(':')[0].Trim();
                        if (accountMap.ContainsKey(rootPart))
                            return accountMap[rootPart];

                        string leafPart = trimmed.Split(':').Last().Trim();
                        if (accountMap.ContainsKey(leafPart))
                            return accountMap[leafPart];

                        var numMatch = Regex.Match(rootPart, @"^([\d]+(?:[-\.][\dA-Za-z]+)*)\s*[-·:]*\s*(.*)$");
                        if (numMatch.Success)
                        {
                            string num = numMatch.Groups[1].Value;
                            string name = Regex.Replace(numMatch.Groups[2].Value, @"^[-·:]*\s*", "").Trim();
                            return $"{num} - {name}";
                        }
                        return rootPart;
                    }

                    return trimmed;
                }

                // STAGE 1: Lightweight header lookup without line items
                IMsgSetRequest reqHeader = sessionManager.CreateMsgSetRequest("US", 13, 0);
                reqHeader.Attributes.OnError = ENRqOnError.roeContinue;

                IBillPaymentCheckQuery bpcHeaderQuery = reqHeader.AppendBillPaymentCheckQueryRq();
                bpcHeaderQuery.IncludeLineItems.SetValue(false);
                bpcHeaderQuery.ORTxnQuery.TxnFilter.ORRefNumberFilter.RefNumberFilter.MatchCriterion.SetValue(ENMatchCriterion.mcStartsWith);
                bpcHeaderQuery.ORTxnQuery.TxnFilter.ORRefNumberFilter.RefNumberFilter.RefNumber.SetValue(targetRef);
                bpcHeaderQuery.IncludeRetElementList.Add("TxnID");
                bpcHeaderQuery.IncludeRetElementList.Add("RefNumber");

                IMsgSetResponse respHeader = sessionManager.DoRequests(reqHeader);
                IBillPaymentCheckRetList bpList = respHeader.ResponseList.GetAt(0).Detail as IBillPaymentCheckRetList;

                if (bpList == null || bpList.Count == 0)
                    return bills;

                List<string> foundRefNumbers = new List<string>();
                string exactBpTxnID = null;

                for (int idx = 0; idx < bpList.Count; idx++)
                {
                    var item = bpList.GetAt(idx);
                    string rNum = item.RefNumber?.GetValue()?.Trim() ?? "";

                    if (!string.IsNullOrEmpty(rNum) && !foundRefNumbers.Contains(rNum))
                        foundRefNumbers.Add(rNum);

                    if (string.Equals(rNum, targetRef, StringComparison.OrdinalIgnoreCase))
                        exactBpTxnID = item.TxnID?.GetValue();
                }

                if (string.IsNullOrEmpty(exactBpTxnID))
                {
                    var sortedRefNumbers = foundRefNumbers
                        .OrderBy(r => long.TryParse(r, out long num) ? num : long.MaxValue)
                        .ThenBy(r => r)
                        .ToList();

                    MessageBox.Show(
                        $"Reference Number mismatch for '{targetRef}'!\n\n" +
                        $"No exact match found. QuickBooks returned the following check(s) in sequence:\n - {string.Join("\n - ", sortedRefNumbers)}\n\n" +
                        "Please enter the exact reference number.",
                        "Reference Number Mismatch",
                        MessageBoxButtons.OK,
                        MessageBoxIcon.Warning
                    );
                    return bills;
                }

                // STAGE 2: Fetch full details for ONLY that exact BillPaymentCheck
                IMsgSetRequest req1 = sessionManager.CreateMsgSetRequest("US", 13, 0);
                req1.Attributes.OnError = ENRqOnError.roeContinue;

                IBillPaymentCheckQuery bpcQuery = req1.AppendBillPaymentCheckQueryRq();
                bpcQuery.IncludeLineItems.SetValue(true);
                bpcQuery.ORTxnQuery.TxnIDList.Add(exactBpTxnID);

                IMsgSetResponse resp1 = sessionManager.DoRequests(req1);
                IBillPaymentCheckRetList exactList = resp1.ResponseList.GetAt(0).Detail as IBillPaymentCheckRetList;
                if (exactList == null || exactList.Count == 0) return bills;

                IBillPaymentCheckRet bp = exactList.GetAt(0);
                string actualCheckRefNumber = bp.RefNumber?.GetValue() ?? "";
                DateTime payDate = bp.TxnDate?.GetValue() ?? DateTime.MinValue;
                string payee = bp.PayeeEntityRef?.FullName?.GetValue() ?? "";
                string address1 = bp.Address?.Addr1?.GetValue() ?? "";
                string address2 = bp.Address?.Addr2?.GetValue() ?? "";
                string bankAccount = bp.BankAccountRef?.FullName?.GetValue() ?? "";
                string memo = bp.Memo?.GetValue() ?? "";
                double totalCheckAmountPaid = bp.Amount?.GetValue() ?? 0;

                var appliedTxnDetails = new Dictionary<string, (double AppliedAmount, double DiscountAmount, string DiscountAccount)>();

                if (bp.AppliedToTxnRetList != null && bp.AppliedToTxnRetList.Count > 0)
                {
                    for (int k = 0; k < bp.AppliedToTxnRetList.Count; k++)
                    {
                        var applied = bp.AppliedToTxnRetList.GetAt(k);
                        string tId = applied.TxnID?.GetValue();
                        if (!string.IsNullOrEmpty(tId))
                        {
                            double appliedAmt = applied.Amount?.GetValue() ?? 0;
                            double discAmt = applied.DiscountAmount?.GetValue() ?? 0;
                            string rawDiscAcc = applied.DiscountAccountRef?.FullName?.GetValue() ?? "";
                            appliedTxnDetails[tId] = (appliedAmt, discAmt, ResolveToParent(rawDiscAcc));
                        }
                    }
                }
                else
                {
                    MessageBox.Show("No Applied Bill found from Bill Payment Check.");
                    return bills;
                }

                // 2. Query Applied Bills
                IMsgSetRequest req2 = sessionManager.CreateMsgSetRequest("US", 13, 0);
                req2.Attributes.OnError = ENRqOnError.roeContinue;

                IBillQuery billQuery = req2.AppendBillQueryRq();
                billQuery.IncludeLineItems.SetValue(true);
                billQuery.IncludeLinkedTxns.SetValue(true);

                foreach (string id in appliedTxnDetails.Keys)
                {
                    billQuery.ORBillQuery.TxnIDList.Add(id);
                }

                IMsgSetResponse resp2 = sessionManager.DoRequests(req2);
                IBillRetList billList = resp2.ResponseList.GetAt(0).Detail as IBillRetList;
                if (billList == null || billList.Count == 0)
                {
                    MessageBox.Show("Bills not found for the provided TxnIDs.");
                    return bills;
                }

                var linkedCreditTxns = new List<(string BillTxnID, string CreditTxnID, double Amount)>();

                for (int bIndex = 0; bIndex < billList.Count; bIndex++)
                {
                    IBillRet bill = billList.GetAt(bIndex);

                    DateTime billTxnDate = bill.TxnDate?.GetValue() ?? payDate;
                    DateTime billDueDate = bill.DueDate?.GetValue() ?? billTxnDate;
                    if (billDueDate == DateTime.MinValue) billDueDate = payDate;

                    double amountDue = bill.AmountDue?.GetValue() ?? 0;
                    string billMemo = bill.Memo?.GetValue() ?? "";
                    string billAPAccount = bill.APAccountRef?.FullName?.GetValue() ?? "";
                    string billRefNumber = bill.RefNumber?.GetValue() ?? "";
                    string specificTxnID = bill.TxnID?.GetValue() ?? "";
                    string resolvedAPAccount = ResolveToParent(billAPAccount);

                    double individualBillPaidAmt = 0;
                    double discountAmt = 0;
                    string discountAcc = "";

                    if (appliedTxnDetails.ContainsKey(specificTxnID))
                    {
                        individualBillPaidAmt = appliedTxnDetails[specificTxnID].AppliedAmount;
                        discountAmt = appliedTxnDetails[specificTxnID].DiscountAmount;
                        discountAcc = appliedTxnDetails[specificTxnID].DiscountAccount;
                    }

                    if (bill.LinkedTxnList != null)
                    {
                        for (int l = 0; l < bill.LinkedTxnList.Count; l++)
                        {
                            ILinkedTxn linked = bill.LinkedTxnList.GetAt(l);
                            string txnType = linked.TxnType?.GetValue().ToString() ?? "";
                            if (txnType.IndexOf("Credit", StringComparison.OrdinalIgnoreCase) >= 0)
                            {
                                string credTxnID = linked.TxnID?.GetValue();
                                double credAmount = linked.Amount?.GetValue() ?? 0;
                                if (!string.IsNullOrEmpty(credTxnID) && !linkedCreditTxns.Any(x => x.CreditTxnID == credTxnID))
                                {
                                    linkedCreditTxns.Add((specificTxnID, credTxnID, Math.Abs(credAmount)));
                                }
                            }
                        }
                    }

                    BillTable bt = new BillTable
                    {
                        DateCreated = billTxnDate,
                        BillCheckDate = payDate,
                        DueDate = billDueDate,
                        PayeeFullName = payee,
                        Address = address1,
                        Address2 = address2,
                        BankAccount = bankAccount,
                        APAccountRefFullName = resolvedAPAccount,
                        Amount = individualBillPaidAmt,
                        AppliedAmount = individualBillPaidAmt,
                        TotalCheckAmount = totalCheckAmountPaid,
                        RefNumber = actualCheckRefNumber,
                        AppliedRefNumber = billRefNumber,
                        AppliedToTxnTxnID = specificTxnID,
                        Memo = memo,
                        BillMemo = billMemo,
                        AmountDue = amountDue,
                        AppliedToTxnDiscountAmount = discountAmt,
                        AppliedToTxnDiscountAccountRefFullName = discountAcc
                    };

                    if (bill.ExpenseLineRetList != null)
                    {
                        for (int i = 0; i < bill.ExpenseLineRetList.Count; i++)
                        {
                            var exp = bill.ExpenseLineRetList.GetAt(i);
                            string rawAccountName = exp.AccountRef?.FullName?.GetValue() ?? "";
                            bt.ItemDetails.Add(new ItemDetail
                            {
                                ItemLineItemRefFullName = ResolveToParent(rawAccountName),
                                ItemLineAmount = exp.Amount?.GetValue() ?? 0,
                                ItemLineClassRefFullName = exp.ClassRef?.FullName?.GetValue() ?? "",
                                ItemLineCustomerJob = exp.CustomerRef?.FullName?.GetValue() ?? "",
                                ItemLineMemo = exp.Memo?.GetValue() ?? ""
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
                                string rawAccountName = item.ItemRef?.FullName?.GetValue() ?? "";
                                bt.ItemDetails.Add(new ItemDetail
                                {
                                    ItemLineItemRefFullName = ResolveToParent(rawAccountName),
                                    ItemLineAmount = item.Amount?.GetValue() ?? 0,
                                    ItemLineClassRefFullName = item.ClassRef?.FullName?.GetValue() ?? "",
                                    ItemLineCustomerJob = item.CustomerRef?.FullName?.GetValue() ?? "",
                                    ItemLineMemo = item.Desc?.GetValue() ?? ""
                                });
                            }
                        }
                    }

                    bills.Add(bt);
                }

                // 3. Query Vendor Credits ONLY if needed
                double calculatedTotalDebits = bills.Sum(b => b.ItemDetails.Sum(d => d.ItemLineAmount));
                double totalPaidCheck = bp.Amount?.GetValue() ?? 0;
                double totalDiscounts = bills.Sum(b => b.AppliedToTxnDiscountAmount);
                double expectedCreditAmount = calculatedTotalDebits - totalPaidCheck - totalDiscounts;

                if (linkedCreditTxns.Count > 0 || expectedCreditAmount > 0.01)
                {
                    try
                    {
                        IMsgSetRequest reqCred = sessionManager.CreateMsgSetRequest("US", 13, 0);
                        reqCred.Attributes.OnError = ENRqOnError.roeContinue;

                        IVendorCreditQuery vcQuery = reqCred.AppendVendorCreditQueryRq();
                        vcQuery.IncludeLineItems.SetValue(true);

                        if (linkedCreditTxns.Count > 0)
                        {
                            foreach (var c in linkedCreditTxns.Select(x => x.CreditTxnID).Distinct())
                                vcQuery.ORTxnQuery.TxnIDList.Add(c);
                        }
                        else
                        {
                            vcQuery.ORTxnQuery.TxnFilter.EntityFilter.OREntityFilter.FullNameList.Add(payee);
                        }

                        IMsgSetResponse respCred = sessionManager.DoRequests(reqCred);
                        if (respCred.ResponseList != null && respCred.ResponseList.Count > 0)
                        {
                            IVendorCreditRetList vcList = respCred.ResponseList.GetAt(0).Detail as IVendorCreditRetList;
                            if (vcList != null && vcList.Count > 0)
                            {
                                double remainingCreditToAllocate = expectedCreditAmount > 0.01 ? expectedCreditAmount : linkedCreditTxns.Sum(x => x.Amount);

                                for (int c = 0; c < vcList.Count && remainingCreditToAllocate > 0.001; c++)
                                {
                                    IVendorCreditRet vCredit = vcList.GetAt(c);
                                    string crRefNum = vCredit.RefNumber?.GetValue() ?? "";
                                    string crMemo = vCredit.Memo?.GetValue() ?? "";

                                    if (vCredit.ExpenseLineRetList != null)
                                    {
                                        for (int e = 0; e < vCredit.ExpenseLineRetList.Count && remainingCreditToAllocate > 0.001; e++)
                                        {
                                            var exp = vCredit.ExpenseLineRetList.GetAt(e);
                                            string rawAcc = exp.AccountRef?.FullName?.GetValue() ?? "";
                                            double lineAmt = exp.Amount?.GetValue() ?? 0;

                                            double appliedThisLine = Math.Min(lineAmt, remainingCreditToAllocate);
                                            remainingCreditToAllocate -= appliedThisLine;

                                            bills[0].AppliedBillCredits.Add(new BillCreditDetail
                                            {
                                                CreditTxnID = vCredit.TxnID?.GetValue() ?? "",
                                                CreditRefNumber = crRefNum,
                                                AppliedAmount = appliedThisLine,
                                                AccountRefFullName = ResolveToParent(rawAcc),
                                                ClassRefFullName = exp.ClassRef?.FullName?.GetValue() ?? "",
                                                CustomerJob = exp.CustomerRef?.FullName?.GetValue() ?? "",
                                                Memo = !string.IsNullOrWhiteSpace(exp.Memo?.GetValue()) ? exp.Memo.GetValue() : crMemo
                                            });
                                        }
                                    }
                                }
                            }
                        }
                    }
                    catch (Exception exCred)
                    {
                        Console.WriteLine($"Error retrieving Vendor Credits: {exCred.Message}");
                    }
                }
            }
            catch (Exception ex)
            {
                Dispose();
                MessageBox.Show("Error retrieving Bill data: " + ex.Message);
            }

            return bills;
        }

        // =========================================================================
        // 4. FAST APV DIRECT BILL RETRIEVAL (BATCHED BILL + VENDOR IN 1 ROUND TRIP)
        // =========================================================================
        public List<BillTable> GetBillData_INT_DirectBill(string billRefNumber)
        {
            List<BillTable> bills = new List<BillTable>();

            try
            {
                var sessionManager = GetSession();

                IMsgSetRequest request = sessionManager.CreateMsgSetRequest("US", 13, 0);
                request.Attributes.OnError = ENRqOnError.roeContinue;

                IBillQuery billQuery = request.AppendBillQueryRq();
                billQuery.IncludeLineItems.SetValue(true);
                billQuery.ORBillQuery.BillFilter.ORRefNumberFilter.RefNumberFilter.MatchCriterion.SetValue(ENMatchCriterion.mcStartsWith);
                billQuery.ORBillQuery.BillFilter.ORRefNumberFilter.RefNumberFilter.RefNumber.SetValue(billRefNumber);

                IMsgSetResponse response = sessionManager.DoRequests(request);
                if (response.ResponseList == null || response.ResponseList.Count == 0) return bills;

                IBillRetList billList = response.ResponseList.GetAt(0).Detail as IBillRetList;
                if (billList == null || billList.Count == 0)
                {
                    MessageBox.Show("Bill RefNumber not found: " + billRefNumber);
                    return bills;
                }

                // Batch-fetch all unique Vendor TINs in ONE request instead of N separate requests
                var vendorIds = new HashSet<string>();
                for (int i = 0; i < billList.Count; i++)
                {
                    string vId = billList.GetAt(i).VendorRef?.ListID?.GetValue();
                    if (!string.IsNullOrEmpty(vId)) vendorIds.Add(vId);
                }

                var vendorTinMap = new Dictionary<string, string>();
                if (vendorIds.Count > 0)
                {
                    IMsgSetRequest vendorReq = sessionManager.CreateMsgSetRequest("US", 13, 0);
                    IVendorQuery vq = vendorReq.AppendVendorQueryRq();
                    foreach (var vId in vendorIds) vq.ORVendorListQuery.ListIDList.Add(vId);
                    vq.OwnerIDList.Add("0");
                    vq.IncludeRetElementList.Add("ListID");
                    vq.IncludeRetElementList.Add("VendorTaxIdent");
                    vq.IncludeRetElementList.Add("DataExtRet");

                    IMsgSetResponse vResp = sessionManager.DoRequests(vendorReq);
                    IVendorRetList vList = vResp.ResponseList.GetAt(0).Detail as IVendorRetList;
                    if (vList != null)
                    {
                        for (int v = 0; v < vList.Count; v++)
                        {
                            IVendorRet vendor = vList.GetAt(v);
                            string listId = vendor.ListID?.GetValue() ?? "";
                            string tin = vendor.VendorTaxIdent?.GetValue() ?? "";
                            if (string.IsNullOrEmpty(tin) && vendor.DataExtRetList != null)
                            {
                                for (int k = 0; k < vendor.DataExtRetList.Count; k++)
                                {
                                    var dataExt = vendor.DataExtRetList.GetAt(k);
                                    if (dataExt.DataExtName.GetValue().IndexOf("TIN", StringComparison.OrdinalIgnoreCase) >= 0)
                                    {
                                        tin = dataExt.DataExtValue.GetValue();
                                        break;
                                    }
                                }
                            }
                            if (!string.IsNullOrEmpty(listId)) vendorTinMap[listId] = tin;
                        }
                    }
                }

                for (int i = 0; i < billList.Count; i++)
                {
                    IBillRet bill = billList.GetAt(i);
                    string vId = bill.VendorRef?.ListID?.GetValue() ?? "";
                    string vendorTIN = vendorTinMap.TryGetValue(vId, out string t) ? t : "";
                    double billAmountDue = bill.AmountDue?.GetValue() ?? 0;

                    BillTable bt = new BillTable
                    {
                        DateCreated = bill.TxnDate?.GetValue() ?? DateTime.Now,
                        DueDate = bill.DueDate?.GetValue() ?? DateTime.Now,
                        PayeeFullName = bill.VendorRef?.FullName?.GetValue() ?? "",
                        TermsRefFullName = bill.TermsRef?.FullName?.GetValue() ?? "",
                        APAccountRefFullName = bill.APAccountRef?.FullName?.GetValue() ?? "",
                        RefNumber = bill.RefNumber?.GetValue() ?? "",
                        AppliedRefNumber = bill.RefNumber?.GetValue() ?? "",
                        Memo = Truncate(bill.Memo?.GetValue() ?? "", 500),
                        BillMemo = Truncate(bill.Memo?.GetValue() ?? "", 500),
                        AmountDue = billAmountDue,
                        Amount = billAmountDue,
                        IsPaid = bill.IsPaid?.GetValue() ?? false,
                        VendorAddressAddr1 = bill.VendorAddress?.Addr1?.GetValue() ?? "",
                        VendorAddressAddr2 = bill.VendorAddress?.Addr2?.GetValue() ?? "",
                        VendorAddressAddr3 = bill.VendorAddress?.Addr3?.GetValue() ?? "",
                        VendorAddressAddr4 = bill.VendorAddress?.Addr4?.GetValue() ?? "",
                        VendorAddressCity = bill.VendorAddress?.City?.GetValue() ?? "",
                        Tin = vendorTIN,
                        Currency = bill.CurrencyRef?.FullName?.GetValue() ?? "",
                        Exchangerate = bill.ExchangeRate?.GetValue() ?? 1.0,
                        ItemDetails = new List<ItemDetail>()
                    };

                    if (bill.ExpenseLineRetList != null)
                    {
                        for (int j = 0; j < bill.ExpenseLineRetList.Count; j++)
                        {
                            var exp = bill.ExpenseLineRetList.GetAt(j);
                            bt.ItemDetails.Add(new ItemDetail
                            {
                                ExpenseLineItemRefFullName = exp.AccountRef?.FullName?.GetValue() ?? "",
                                ExpenseLineAmount = exp.Amount?.GetValue() ?? 0,
                                ExpenseLineClassRefFullName = exp.ClassRef?.FullName?.GetValue() ?? "",
                                ExpenseLineCustomerJob = exp.CustomerRef?.FullName?.GetValue() ?? "",
                                ExpenseLineMemo = Truncate(exp.Memo?.GetValue() ?? "", 500)
                            });
                        }
                    }

                    if (bill.ORItemLineRetList != null)
                    {
                        for (int j = 0; j < bill.ORItemLineRetList.Count; j++)
                        {
                            var orItem = bill.ORItemLineRetList.GetAt(j);
                            if (orItem.ItemLineRet != null)
                            {
                                var item = orItem.ItemLineRet;
                                bt.ItemDetails.Add(new ItemDetail
                                {
                                    ItemLineItemRefFullName = item.ItemRef?.FullName?.GetValue() ?? "",
                                    ItemLineAmount = item.Amount?.GetValue() ?? 0,
                                    ItemLineClassRefFullName = item.ClassRef?.FullName?.GetValue() ?? "",
                                    ItemLineCustomerJob = item.CustomerRef?.FullName?.GetValue() ?? "",
                                    ItemLineMemo = Truncate(item.Desc?.GetValue() ?? "", 500)
                                });
                            }
                        }
                    }

                    bills.Add(bt);
                }
            }
            catch (Exception ex)
            {
                Dispose();
                MessageBox.Show("Error: " + ex.Message);
            }

            return bills;
        }

        // =========================================================================
        // 5. FAST JOURNAL ENTRY RETRIEVAL (mcStartsWith + TWO-STAGE + CACHED COA)
        // =========================================================================
        public List<JournalGridItem> GetJournalEntryForGrid(string refNumber)
        {
            List<JournalGridItem> gridItems = new List<JournalGridItem>();
            string targetRef = refNumber.Trim();

            try
            {
                var sessionManager = GetSession();
                Dictionary<string, string> accountMap = GetAccountNumbersMap(sessionManager);

                // STAGE 1: Lightweight header lookup with mcStartsWith (no line items yet)
                IMsgSetRequest headerReq = sessionManager.CreateMsgSetRequest("US", 13, 0);
                headerReq.Attributes.OnError = ENRqOnError.roeContinue;

                IJournalEntryQuery jeHeaderQuery = headerReq.AppendJournalEntryQueryRq();
                jeHeaderQuery.IncludeLineItems.SetValue(false);
                jeHeaderQuery.ORTxnQuery.TxnFilter.ORRefNumberFilter.RefNumberFilter.MatchCriterion.SetValue(ENMatchCriterion.mcStartsWith);
                jeHeaderQuery.ORTxnQuery.TxnFilter.ORRefNumberFilter.RefNumberFilter.RefNumber.SetValue(targetRef);
                jeHeaderQuery.IncludeRetElementList.Add("TxnID");
                jeHeaderQuery.IncludeRetElementList.Add("RefNumber");
                jeHeaderQuery.IncludeRetElementList.Add("TimeModified");
                jeHeaderQuery.IncludeRetElementList.Add("TxnDate");

                IMsgSetResponse headerResp = sessionManager.DoRequests(headerReq);
                IJournalEntryRetList list = headerResp.ResponseList.GetAt(0).Detail as IJournalEntryRetList;

                if (list == null || list.Count == 0) return gridItems;

                string targetTxnID = null;
                DateTime latestTimeModified = DateTime.MinValue;

                for (int i = 0; i < list.Count; i++)
                {
                    IJournalEntryRet je = list.GetAt(i);
                    string docNum = je.RefNumber?.GetValue()?.Trim() ?? "";

                    if (string.Equals(docNum, targetRef, StringComparison.OrdinalIgnoreCase))
                    {
                        DateTime modifiedDate = je.TimeModified != null ? je.TimeModified.GetValue() : je.TxnDate.GetValue();
                        if (modifiedDate >= latestTimeModified)
                        {
                            latestTimeModified = modifiedDate;
                            targetTxnID = je.TxnID?.GetValue();
                        }
                    }
                }

                if (string.IsNullOrEmpty(targetTxnID)) return gridItems;

                // STAGE 2: Fetch full line items ONLY for the 1 exact TxnID
                IMsgSetRequest detailReq = sessionManager.CreateMsgSetRequest("US", 13, 0);
                detailReq.Attributes.OnError = ENRqOnError.roeContinue;

                IJournalEntryQuery jeDetailQuery = detailReq.AppendJournalEntryQueryRq();
                jeDetailQuery.IncludeLineItems.SetValue(true);
                jeDetailQuery.ORTxnQuery.TxnIDList.Add(targetTxnID);

                IMsgSetResponse detailResp = sessionManager.DoRequests(detailReq);
                IJournalEntryRetList detailList = detailResp.ResponseList.GetAt(0).Detail as IJournalEntryRetList;
                IJournalEntryRet targetEntry = (detailList != null && detailList.Count > 0) ? detailList.GetAt(0) : null;

                if (targetEntry != null && targetEntry.ORJournalLineList != null)
                {
                    string docNum = targetEntry.RefNumber?.GetValue() ?? targetRef;
                    DateTime date = targetEntry.TxnDate.GetValue();

                    for (int j = 0; j < targetEntry.ORJournalLineList.Count; j++)
                    {
                        IORJournalLine orLine = targetEntry.ORJournalLineList.GetAt(j);
                        JournalGridItem item = new JournalGridItem
                        {
                            Date = date,
                            Num = docNum,
                            Type = "General Journal"
                        };

                        if (orLine.JournalDebitLine != null)
                        {
                            var line = orLine.JournalDebitLine;
                            string fullAccountName = line.AccountRef?.FullName?.GetValue() ?? "";
                            item.AccountName = fullAccountName;
                            item.AccountNumber = GetAccountNumberFromMap(accountMap, fullAccountName);
                            item.Name = line.EntityRef?.FullName?.GetValue() ?? "";
                            item.Memo = Truncate(line.Memo?.GetValue() ?? "", 500);
                            item.Class = line.ClassRef?.FullName?.GetValue() ?? "";
                            item.Debit = line.Amount?.GetValue() ?? 0;
                            item.Credit = 0;
                        }
                        else if (orLine.JournalCreditLine != null)
                        {
                            var line = orLine.JournalCreditLine;
                            string fullAccountName = line.AccountRef?.FullName?.GetValue() ?? "";
                            item.AccountName = fullAccountName;
                            item.AccountNumber = GetAccountNumberFromMap(accountMap, fullAccountName);
                            item.Name = line.EntityRef?.FullName?.GetValue() ?? "";
                            item.Memo = Truncate(line.Memo?.GetValue() ?? "", 500);
                            item.Class = line.ClassRef?.FullName?.GetValue() ?? "";
                            item.Debit = 0;
                            item.Credit = line.Amount?.GetValue() ?? 0;
                        }

                        gridItems.Add(item);
                    }
                }
            }
            catch (Exception ex)
            {
                Dispose();
                Console.WriteLine($"ERROR: {ex.Message}");
            }

            return gridItems;
        }

        // =========================================================================
        // CACHED CHART OF ACCOUNTS HELPERS
        // =========================================================================
        private Dictionary<string, string> GetParentAccountConsolidationMap(QBSessionManager sessionManager)
        {
            if (_cachedParentAccountMap != null) return _cachedParentAccountMap;

            var accountMap = new Dictionary<string, string>(StringComparer.OrdinalIgnoreCase);

            try
            {
                IMsgSetRequest accReq = sessionManager.CreateMsgSetRequest("US", 13, 0);
                accReq.Attributes.OnError = ENRqOnError.roeContinue;
                IAccountQuery accQuery = accReq.AppendAccountQueryRq();
                accQuery.ORAccountListQuery.AccountListFilter.ActiveStatus.SetValue(ENActiveStatus.asAll);
                accQuery.IncludeRetElementList.Add("FullName");
                accQuery.IncludeRetElementList.Add("Name");
                accQuery.IncludeRetElementList.Add("AccountNumber");
                accQuery.IncludeRetElementList.Add("ParentRef");

                IMsgSetResponse accResp = sessionManager.DoRequests(accReq);
                if (accResp.ResponseList != null && accResp.ResponseList.Count > 0)
                {
                    IAccountRetList accList = accResp.ResponseList.GetAt(0).Detail as IAccountRetList;
                    if (accList != null)
                    {
                        var rawAccounts = new Dictionary<string, (string Name, string FullName, string AccNum, string ParentFullName)>(StringComparer.OrdinalIgnoreCase);

                        for (int i = 0; i < accList.Count; i++)
                        {
                            IAccountRet acc = accList.GetAt(i);
                            string fullName = acc.FullName?.GetValue() ?? "";
                            string name = acc.Name?.GetValue() ?? "";
                            string accNum = acc.AccountNumber?.GetValue() ?? "";
                            string parentFullName = acc.ParentRef?.FullName?.GetValue() ?? "";

                            if (!string.IsNullOrEmpty(fullName))
                                rawAccounts[fullName] = (name, fullName, accNum, parentFullName);
                        }

                        foreach (var kvp in rawAccounts)
                        {
                            var current = kvp.Value;
                            var root = current;

                            while (!string.IsNullOrWhiteSpace(root.ParentFullName) && rawAccounts.ContainsKey(root.ParentFullName))
                                root = rawAccounts[root.ParentFullName];

                            string rootAccNum = root.AccNum;
                            if (string.IsNullOrWhiteSpace(rootAccNum))
                            {
                                var matchRoot = Regex.Match(root.FullName, @"^([\d]+(?:[-\.][\dA-Za-z]+)*)");
                                if (matchRoot.Success) rootAccNum = matchRoot.Groups[1].Value;
                                else
                                {
                                    var matchCurrent = Regex.Match(current.FullName, @"^([\d]+(?:[-\.][\dA-Za-z]+)*)");
                                    if (matchCurrent.Success) rootAccNum = matchCurrent.Groups[1].Value;
                                }
                            }

                            string cleanRootName = root.Name;
                            if (!string.IsNullOrWhiteSpace(rootAccNum) && cleanRootName.StartsWith(rootAccNum, StringComparison.OrdinalIgnoreCase))
                                cleanRootName = cleanRootName.Substring(rootAccNum.Length).Trim();

                            cleanRootName = Regex.Replace(cleanRootName, @"^[-·:]*\s*", "").Trim();

                            string consolidatedParentName = !string.IsNullOrWhiteSpace(rootAccNum)
                                ? $"{rootAccNum} - {cleanRootName}"
                                : cleanRootName;

                            accountMap[kvp.Key] = consolidatedParentName;
                            accountMap[current.FullName] = consolidatedParentName;
                            if (!accountMap.ContainsKey(current.Name))
                                accountMap[current.Name] = consolidatedParentName;
                        }
                    }
                }

                _cachedParentAccountMap = accountMap;
            }
            catch (Exception ex)
            {
                Console.WriteLine($"Error building account map: {ex.Message}");
            }

            return accountMap;
        }

        private Dictionary<string, string> GetAccountNumbersMap(QBSessionManager sessionManager)
        {
            if (_cachedAccountNumbersMap != null) return _cachedAccountNumbersMap;

            var map = new Dictionary<string, string>(StringComparer.OrdinalIgnoreCase);

            try
            {
                IMsgSetRequest request = sessionManager.CreateMsgSetRequest("US", 13, 0);
                request.Attributes.OnError = ENRqOnError.roeContinue;

                IAccountQuery accountQuery = request.AppendAccountQueryRq();
                accountQuery.ORAccountListQuery.AccountListFilter.ActiveStatus.SetValue(ENActiveStatus.asAll);
                accountQuery.IncludeRetElementList.Add("FullName");
                accountQuery.IncludeRetElementList.Add("AccountNumber");

                IMsgSetResponse response = sessionManager.DoRequests(request);
                IAccountRetList accountList = response.ResponseList.GetAt(0).Detail as IAccountRetList;

                if (accountList != null)
                {
                    for (int i = 0; i < accountList.Count; i++)
                    {
                        IAccountRet account = accountList.GetAt(i);
                        string fullName = account.FullName?.GetValue() ?? "";
                        string acctNum = account.AccountNumber?.GetValue()?.Trim() ?? "";

                        if (!string.IsNullOrEmpty(fullName) && !map.ContainsKey(fullName))
                            map.Add(fullName, acctNum);
                    }
                }

                _cachedAccountNumbersMap = map;
            }
            catch (Exception ex)
            {
                Console.WriteLine($"Error querying Chart of Accounts: {ex.Message}");
            }

            return map;
        }

        private string GetAccountNumberFromMap(Dictionary<string, string> map, string fullName)
        {
            if (string.IsNullOrWhiteSpace(fullName)) return "";
            string trimmedName = fullName.Trim();

            if (map.TryGetValue(trimmedName, out string acctNum) && !string.IsNullOrWhiteSpace(acctNum))
                return acctNum;

            string currentPath = trimmedName;
            while (currentPath.Contains(':'))
            {
                int lastColon = currentPath.LastIndexOf(':');
                currentPath = currentPath.Substring(0, lastColon).Trim();
                if (map.TryGetValue(currentPath, out string parentAcctNum) && !string.IsNullOrWhiteSpace(parentAcctNum))
                    return parentAcctNum;
            }

            return ExtractAccountNumber(trimmedName);
        }

        private string ExtractAccountNumber(string fullName)
        {
            if (string.IsNullOrWhiteSpace(fullName)) return "";
            var hierarchy = fullName.Split(new[] { ':' }, StringSplitOptions.RemoveEmptyEntries);
            for (int i = hierarchy.Length - 1; i >= 0; i--)
            {
                string segment = hierarchy[i].Trim();
                var parts = segment.Split(new[] { ' ', '·', '-' }, StringSplitOptions.RemoveEmptyEntries);
                if (parts.Length > 0 && parts[0].All(char.IsDigit))
                    return parts[0];
            }
            return "";
        }

        public int GetNextIncrementalID_CV(string accessConnectionString)
        {
            int incrementalID = 0;
            using (OleDbConnection connection = new OleDbConnection(accessConnectionString))
            {
                string query = "SELECT FormattedID FROM CVIncrement";
                OleDbCommand command = new OleDbCommand(query, connection);
                try
                {
                    connection.Open();
                    object result = command.ExecuteScalar();
                    if (result != null) incrementalID = Convert.ToInt32(result);
                    else
                    {
                        command = new OleDbCommand("INSERT INTO CVIncrement (FormattedID) VALUES (0)", connection);
                        if (command.ExecuteNonQuery() > 0) incrementalID = 0;
                    }
                }
                catch (Exception ex) { Console.WriteLine("Error: " + ex.Message); }
            }
            return incrementalID;
        }

        public int GetNextIncrementalID_APV(string accessConnectionString)
        {
            int incrementalID = 0;
            using (OleDbConnection connection = new OleDbConnection(accessConnectionString))
            {
                string query = "SELECT FormattedID FROM APVIncrement";
                OleDbCommand command = new OleDbCommand(query, connection);
                try
                {
                    connection.Open();
                    object result = command.ExecuteScalar();
                    if (result != null) incrementalID = Convert.ToInt32(result);
                    else
                    {
                        command = new OleDbCommand("INSERT INTO APVIncrement (FormattedID) VALUES (0)", connection);
                        if (command.ExecuteNonQuery() > 0) incrementalID = 0;
                    }
                }
                catch (Exception ex) { Console.WriteLine("Error: " + ex.Message); }
            }
            return incrementalID;
        }

        private string Truncate(string value, int maxLength)
        {
            if (string.IsNullOrEmpty(value)) return "";
            return value.Length <= maxLength ? value : value.Substring(0, maxLength);
        }
    }
}