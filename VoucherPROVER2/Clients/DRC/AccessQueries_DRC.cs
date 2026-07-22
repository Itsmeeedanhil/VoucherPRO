using QBFC16Lib;
using System;
using System.Collections.Generic;
using System.Data.Odbc;
using System.Data.OleDb;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using static VoucherPROVER2.Clients.DRC.Dataclass_DRC;
using static VoucherPROVER2.Clients.DRC.AccessToDatabase_DRC;
using System.Windows.Forms;

namespace VoucherPROVER2.Clients.DRC
{
    public class AccessQueries_DRC
    {

        public List<CheckTableGrid> GetCheckDataDRC(string refNumber)
        {
            List<CheckTableGrid> checkList = new List<CheckTableGrid>();
            QBSessionManager sessionManager = new QBSessionManager();

            try
            {
                sessionManager.OpenConnection2("", "VoucherPro Check Data", ENConnectionType.ctLocalQBD);
                sessionManager.BeginSession("", ENOpenMode.omDontCare);

                IMsgSetRequest request = sessionManager.CreateMsgSetRequest("US", 13, 0);
                request.Attributes.OnError = ENRqOnError.roeContinue;

                // ----------------------------------------------------------------
                // 1. QUERY FOR REGULAR CHECKS
                // ----------------------------------------------------------------
                ICheckQuery checkQuery = request.AppendCheckQueryRq();
                checkQuery.ORTxnQuery.TxnFilter.ORRefNumberFilter.RefNumberFilter.MatchCriterion.SetValue(ENMatchCriterion.mcStartsWith);
                checkQuery.ORTxnQuery.TxnFilter.ORRefNumberFilter.RefNumberFilter.RefNumber.SetValue(refNumber);

                // ----------------------------------------------------------------
                // 2. QUERY FOR BILL PAYMENT CHECKS
                // ----------------------------------------------------------------
                IBillPaymentCheckQuery billPayQuery = request.AppendBillPaymentCheckQueryRq();
                billPayQuery.ORTxnQuery.TxnFilter.ORRefNumberFilter.RefNumberFilter.MatchCriterion.SetValue(ENMatchCriterion.mcStartsWith);
                billPayQuery.ORTxnQuery.TxnFilter.ORRefNumberFilter.RefNumberFilter.RefNumber.SetValue(refNumber);

                // Execute Requests
                IMsgSetResponse response = sessionManager.DoRequests(request);

                // ----------------------------------------------------------------
                // PROCESS RESPONSE 1: REGULAR CHECKS
                // ----------------------------------------------------------------
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

                // ----------------------------------------------------------------
                // PROCESS RESPONSE 2: BILL PAYMENT CHECKS
                // ----------------------------------------------------------------
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

        public List<ItemReciept> GetItemReceiptData_DRC(string refNumber)
        {
            QBSessionManager sessionManager = new QBSessionManager();
            List<ItemReciept> receipts = new List<ItemReciept>();

            Console.WriteLine("--------------------------------------------------");
            Console.WriteLine($"[DEBUG] START: GetItemReceiptData_DRC for RefNumber: {refNumber}");

            try
            {
                sessionManager.OpenConnection2("", "Item Receipt Retrieval", ENConnectionType.ctLocalQBD);
                sessionManager.BeginSession("", ENOpenMode.omDontCare);

                // Build MsgSet Request
                IMsgSetRequest request = sessionManager.CreateMsgSetRequest("US", 13, 0);
                request.Attributes.OnError = ENRqOnError.roeContinue;

                // Query Item Receipt
                IItemReceiptQuery irQuery = request.AppendItemReceiptQueryRq();
                irQuery.IncludeLineItems.SetValue(true);

                irQuery.ORTxnQuery.TxnFilter.ORRefNumberFilter.RefNumberFilter.MatchCriterion.SetValue(ENMatchCriterion.mcStartsWith);
                irQuery.ORTxnQuery.TxnFilter.ORRefNumberFilter.RefNumberFilter.RefNumber.SetValue(refNumber);

                IMsgSetResponse response = sessionManager.DoRequests(request);

                if (response.ResponseList == null || response.ResponseList.Count == 0)
                    return receipts;

                IResponse qbResponse = response.ResponseList.GetAt(0);
                IItemReceiptRetList irList = qbResponse.Detail as IItemReceiptRetList;

                if (irList == null || irList.Count == 0)
                {
                    MessageBox.Show($"Item Receipt with RefNumber '{refNumber}' was not found.", "Information", MessageBoxButtons.OK, MessageBoxIcon.Information);
                    return receipts;
                }

                // -------------------------------------------------------------
                // FETCH ACCOUNT NUMBERS FROM QUICKBOOKS
                // -------------------------------------------------------------
                Dictionary<string, string> accountNumbersDict = GetAccountNumbersFromQBIR(sessionManager);

                for (int i = 0; i < irList.Count; i++)
                {
                    IItemReceiptRet ir = irList.GetAt(i);

                    // HEADER DATA
                    DateTime txnDate = ir.TxnDate?.GetValue() ?? DateTime.Now;
                    string vendorName = ir.VendorRef?.FullName?.GetValue() ?? "";
                    string currentRef = ir.RefNumber?.GetValue() ?? "";
                    string memo = ir.Memo?.GetValue() ?? "";
                    double totalAmount = ir.TotalAmount?.GetValue() ?? 0;
                    string apAccount = ir.ORAccountRef?.FullName?.GetValue() ?? "";

                    // FETCH VENDOR ADDRESS
                    string addr1 = "";
                    string addr2 = "";
                    string addr3 = "";
                    string addr4 = "";
                    string city = "";

                    if (ir.VendorRef != null && !string.IsNullOrEmpty(ir.VendorRef.ListID?.GetValue()))
                    {
                        try
                        {
                            IMsgSetRequest vendorReq = sessionManager.CreateMsgSetRequest("US", 13, 0);
                            vendorReq.Attributes.OnError = ENRqOnError.roeContinue;

                            IVendorQuery vq = vendorReq.AppendVendorQueryRq();
                            vq.ORVendorListQuery.ListIDList.Add(ir.VendorRef.ListID.GetValue());

                            IMsgSetResponse vResp = sessionManager.DoRequests(vendorReq);

                            if (vResp.ResponseList != null && vResp.ResponseList.Count > 0)
                            {
                                IResponse vResponseRoot = vResp.ResponseList.GetAt(0);
                                IVendorRetList vList = vResponseRoot.Detail as IVendorRetList;

                                if (vList != null && vList.Count > 0)
                                {
                                    IVendorRet vendor = vList.GetAt(0);

                                    if (vendor.VendorAddress != null)
                                    {
                                        addr1 = vendor.VendorAddress.Addr1?.GetValue() ?? "";
                                        addr2 = vendor.VendorAddress.Addr2?.GetValue() ?? "";
                                        addr3 = vendor.VendorAddress.Addr3?.GetValue() ?? "";
                                        addr4 = vendor.VendorAddress.Addr4?.GetValue() ?? "";
                                        city = vendor.VendorAddress.City?.GetValue() ?? "";
                                    }
                                }
                            }
                        }
                        catch (Exception vEx)
                        {
                            Console.WriteLine($"[DEBUG] Error retrieving vendor address: {vEx.Message}");
                        }
                    }

                    // A. PROCESS ITEM LINES
                    if (ir.ORItemLineRetList != null)
                    {
                        for (int j = 0; j < ir.ORItemLineRetList.Count; j++)
                        {
                            IORItemLineRet orItemLine = (IORItemLineRet)ir.ORItemLineRetList.GetAt(j);

                            if (orItemLine.ItemLineRet != null)
                            {
                                IItemLineRet item = orItemLine.ItemLineRet;

                                receipts.Add(new ItemReciept
                                {
                                    RefNumber = currentRef,
                                    TxnDate = txnDate,
                                    DateCreated = txnDate,
                                    VendorFullName = vendorName,
                                    PayeeFullName = vendorName,
                                    Memo = memo,
                                    TotalAmount = totalAmount,
                                    BankAccount = apAccount,
                                    Addr1 = addr1,
                                    Addr2 = addr2,
                                    Addr3 = addr3,
                                    Addr4 = addr4,
                                    AddrCity = city,

                                    // Item Fields
                                    Item = item.ItemRef?.FullName?.GetValue() ?? "",
                                    ItemDescription = item.Desc?.GetValue() ?? "",
                                    ItemQuantity = item.Quantity?.GetValue() ?? 0,
                                    ItemCost = item.Cost?.GetValue() ?? 0,
                                    ItemAmount = item.Amount?.GetValue() ?? 0,
                                    ItemClass = item.ClassRef?.FullName?.GetValue() ?? "",
                                    ItemCustomerJob = item.CustomerRef?.FullName?.GetValue() ?? "",

                                    ReceiptItemType = ReceiptItemType.ReceiptItem
                                });
                            }
                        }
                    }

                    // B. PROCESS EXPENSE LINES (WITH ACCOUNT NUMBER LOOKUP)
                    if (ir.ExpenseLineRetList != null)
                    {
                        for (int j = 0; j < ir.ExpenseLineRetList.Count; j++)
                        {
                            IExpenseLineRet exp = ir.ExpenseLineRetList.GetAt(j);

                            string expAccount = exp.AccountRef?.FullName?.GetValue() ?? "";
                            string expListID = exp.AccountRef?.ListID?.GetValue() ?? "";

                            // Match AccountNumber from dictionary lookup using ListID or FullName
                            string accNumber = "";
                            if (!string.IsNullOrEmpty(expListID) && accountNumbersDict.ContainsKey(expListID))
                            {
                                accNumber = accountNumbersDict[expListID];
                            }
                            else if (!string.IsNullOrEmpty(expAccount) && accountNumbersDict.ContainsKey(expAccount))
                            {
                                accNumber = accountNumbersDict[expAccount];
                            }

                            receipts.Add(new ItemReciept
                            {
                                RefNumber = currentRef,
                                TxnDate = txnDate,
                                DateCreated = txnDate,
                                VendorFullName = vendorName,
                                PayeeFullName = vendorName,
                                Memo = memo,
                                TotalAmount = totalAmount,
                                BankAccount = apAccount,
                                Addr1 = addr1,
                                Addr2 = addr2,
                                Addr3 = addr3,
                                Addr4 = addr4,
                                AddrCity = city,

                                // Expense Fields
                                Account = expAccount,
                                AccountNumber = accNumber, // Assigned AccountNumber
                                ExpensesMemo = exp.Memo?.GetValue() ?? "",
                                ExpensesAmount = exp.Amount?.GetValue() ?? 0,
                                ItemClass = exp.ClassRef?.FullName?.GetValue() ?? "",
                                ItemCustomerJob = exp.CustomerRef?.FullName?.GetValue() ?? "",

                                ReceiptItemType = ReceiptItemType.RecieptExpense
                            });
                        }
                    }
                }
            }
            catch (Exception ex)
            {
                Console.WriteLine($"[DEBUG] EXCEPTION: {ex.Message}");
                MessageBox.Show($"Error retrieving Item Receipt data:\n{ex.Message}", "Error", MessageBoxButtons.OK, MessageBoxIcon.Error);
            }
            finally
            {
                try { sessionManager.EndSession(); sessionManager.CloseConnection(); } catch { }
            }

            return receipts;
        }

        // -----------------------------------------------------------------
        // HELPER METHOD TO QUERY QUICKBOOKS FOR ACCOUNT NUMBERS
        // -----------------------------------------------------------------
        private Dictionary<string, string> GetAccountNumbersFromQBIR(QBSessionManager sessionManager)
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


        public List<BillTable> GetBillData_DRC_DirectBill(string billRefNumber)
        {
            QBSessionManager sessionManager = new QBSessionManager();
            List<BillTable> bills = new List<BillTable>();

            try
            {
                sessionManager.OpenConnection2("", "APV Retrieval", ENConnectionType.ctLocalQBD);
                sessionManager.BeginSession("", ENOpenMode.omDontCare);

                // Fetch all account numbers from QuickBooks into a lookup dictionary
                Dictionary<string, string> accountNumbersDict = GetAccountNumbersFromQBAPV(sessionManager);

                IMsgSetRequest request = sessionManager.CreateMsgSetRequest("US", 13, 0);
                request.Attributes.OnError = ENRqOnError.roeContinue;

                // 1. Query the Bill
                IBillQuery billQuery = request.AppendBillQueryRq();
                billQuery.IncludeLineItems.SetValue(true);
                billQuery.ORBillQuery.BillFilter.ORRefNumberFilter.RefNumberFilter.MatchCriterion.SetValue(ENMatchCriterion.mcStartsWith);
                billQuery.ORBillQuery.BillFilter.ORRefNumberFilter.RefNumberFilter.RefNumber.SetValue(billRefNumber);

                IMsgSetResponse response = sessionManager.DoRequests(request);

                if (response.ResponseList == null || response.ResponseList.Count == 0) return bills;

                IResponse resp = response.ResponseList.GetAt(0);
                IBillRetList billList = resp.Detail as IBillRetList;

                if (billList == null || billList.Count == 0)
                {
                    MessageBox.Show("Bill RefNumber not found: " + billRefNumber);
                    return bills;
                }

                // 2. Loop through results
                for (int i = 0; i < billList.Count; i++)
                {
                    IBillRet bill = billList.GetAt(i);

                    // --- ROBUST TIN FETCHING ---
                    string vendorTIN = "";

                    if (bill.VendorRef != null)
                    {
                        try
                        {
                            string vendorListID = bill.VendorRef.ListID?.GetValue();

                            if (!string.IsNullOrEmpty(vendorListID))
                            {
                                IMsgSetRequest vendorReq = sessionManager.CreateMsgSetRequest("US", 13, 0);
                                IVendorQuery vq = vendorReq.AppendVendorQueryRq();
                                vq.ORVendorListQuery.ListIDList.Add(vendorListID);
                                vq.OwnerIDList.Add("0");

                                IMsgSetResponse vResp = sessionManager.DoRequests(vendorReq);
                                IResponse vResponseRoot = vResp.ResponseList.GetAt(0);
                                IVendorRetList vList = vResponseRoot.Detail as IVendorRetList;

                                if (vList != null && vList.Count > 0)
                                {
                                    IVendorRet vendor = vList.GetAt(0);

                                    vendorTIN = vendor.VendorTaxIdent?.GetValue() ?? "";

                                    if (string.IsNullOrEmpty(vendorTIN) && vendor.DataExtRetList != null)
                                    {
                                        for (int k = 0; k < vendor.DataExtRetList.Count; k++)
                                        {
                                            var dataExt = vendor.DataExtRetList.GetAt(k);
                                            if (dataExt.DataExtName.GetValue().IndexOf("TIN", StringComparison.OrdinalIgnoreCase) >= 0)
                                            {
                                                vendorTIN = dataExt.DataExtValue.GetValue();
                                                break;
                                            }
                                        }
                                    }
                                }
                            }
                        }
                        catch (Exception tinEx)
                        {
                            MessageBox.Show("Error fetching TIN: " + tinEx.Message);
                        }
                    }

                    // Lookup Account Number for Header A/P Account Ref
                    string apAccountName = bill.APAccountRef?.FullName?.GetValue() ?? "";
                    string apListID = bill.APAccountRef?.ListID?.GetValue() ?? "";
                    string billAccNum = "";

                    if (!string.IsNullOrEmpty(apListID) && accountNumbersDict.ContainsKey(apListID))
                    {
                        billAccNum = accountNumbersDict[apListID];
                    }
                    else if (!string.IsNullOrEmpty(apAccountName) && accountNumbersDict.ContainsKey(apAccountName))
                    {
                        billAccNum = accountNumbersDict[apAccountName];
                    }

                    BillTable bt = new BillTable
                    {
                        // Core Fields
                        DateCreated = bill.TxnDate?.GetValue() ?? DateTime.Now,
                        DueDate = bill.DueDate?.GetValue() ?? DateTime.Now,
                        PayeeFullName = bill.VendorRef?.FullName?.GetValue() ?? "",
                        TermsRefFullName = bill.TermsRef?.FullName?.GetValue() ?? "",
                        APAccountRefFullName = apAccountName,
                        AccountNumber = billAccNum, // Assigned Header AccountNumber
                        RefNumber = bill.RefNumber?.GetValue() ?? "",
                        Memo = Truncate(bill.Memo?.GetValue() ?? "", 500),
                        AmountDue = bill.AmountDue?.GetValue() ?? 0,
                        IsPaid = bill.IsPaid?.GetValue() ?? false,

                        // Address Fields
                        VendorAddressAddr1 = bill.VendorAddress?.Addr1?.GetValue() ?? "",
                        VendorAddressAddr2 = bill.VendorAddress?.Addr2?.GetValue() ?? "",
                        VendorAddressAddr3 = bill.VendorAddress?.Addr3?.GetValue() ?? "",
                        VendorAddressAddr4 = bill.VendorAddress?.Addr4?.GetValue() ?? "",
                        VendorAddressCity = bill.VendorAddress?.City?.GetValue() ?? "",

                        // Additional Fields
                        Tin = vendorTIN,
                        Currency = bill.CurrencyRef?.FullName?.GetValue() ?? "",
                        Exchangerate = bill.ExchangeRate?.GetValue() ?? 1.0
                    };

                    // 3. Process Expense Lines
                    if (bill.ExpenseLineRetList != null)
                    {
                        for (int j = 0; j < bill.ExpenseLineRetList.Count; j++)
                        {
                            var exp = bill.ExpenseLineRetList.GetAt(j);
                            string expAccountName = exp.AccountRef?.FullName?.GetValue() ?? "";
                            string expListID = exp.AccountRef?.ListID?.GetValue() ?? "";

                            // Lookup Account Number for Expense Line Account
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
                                ExpenseLineAccountNumber = expAccNumber, // Assigned Expense Line AccountNumber
                                ExpenseLineAmount = exp.Amount?.GetValue() ?? 0,
                                ExpenseLineClassRefFullName = exp.ClassRef?.FullName?.GetValue() ?? "",
                                ExpenseLineCustomerJob = exp.CustomerRef?.FullName?.GetValue() ?? "",
                                ExpenseLineMemo = Truncate(exp.Memo?.GetValue() ?? "", 500)
                            });
                        }
                    }

                    // 4. Process Item Lines
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
                                    ItemLineMemo = item.Desc?.GetValue() ?? ""
                                });
                            }
                        }
                    }

                    bills.Add(bt);
                }
            }
            catch (Exception ex)
            {
                MessageBox.Show("Error: " + ex.Message);
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

        // ====================================================
        // HELPER METHOD TO QUERY QUICKBOOKS FOR ACCOUNT NUMBERS
        // ====================================================
        private Dictionary<string, string> GetAccountNumbersFromQBAPV(QBSessionManager sessionManager)
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

        public List<BillTable> GetBillData_DRC(string refNumber)
        {
            QBSessionManager sessionManager = new QBSessionManager();
            List<BillTable> bills = new List<BillTable>();

            Console.WriteLine("--------------------------------------------------");
            Console.WriteLine($"[DEBUG] START: GetBillData_IVP for RefNumber: {refNumber}");

            try
            {
                sessionManager.OpenConnection2("", "Bill Retrieval", ENConnectionType.ctLocalQBD);
                sessionManager.BeginSession("", ENOpenMode.omDontCare);
                Console.WriteLine("[DEBUG] Session Opened Successfully.");

                // ====================================================
                // FETCH ALL ACCOUNT NUMBERS FROM QUICKBOOKS
                // ====================================================
                Dictionary<string, string> accountNumbersDict = GetAccountNumbersFromQBBILL(sessionManager);

                // ====================================================
                // 1. QUERY BILL PAYMENT CHECK USING RefNumber
                // ====================================================
                IMsgSetRequest req1 = sessionManager.CreateMsgSetRequest("US", 13, 0);
                req1.Attributes.OnError = ENRqOnError.roeContinue;

                IBillPaymentCheckQuery bpcQuery = req1.AppendBillPaymentCheckQueryRq();
                bpcQuery.IncludeLineItems.SetValue(true);

                // exact match
                bpcQuery.ORTxnQuery.TxnFilter.ORRefNumberFilter.RefNumberFilter.MatchCriterion.SetValue(ENMatchCriterion.mcStartsWith);
                bpcQuery.ORTxnQuery.TxnFilter.ORRefNumberFilter.RefNumberFilter.RefNumber.SetValue(refNumber);

                Console.WriteLine("[DEBUG] Sending BillPaymentCheck Query...");
                IMsgSetResponse resp1 = sessionManager.DoRequests(req1);
                IResponse r1 = resp1.ResponseList.GetAt(0);

                IBillPaymentCheckRetList bpList = r1.Detail as IBillPaymentCheckRetList;

                if (bpList == null || bpList.Count == 0)
                {
                    MessageBox.Show("Bill Payment Check not found: " + refNumber);
                    return bills;
                }

                IBillPaymentCheckRet bp = bpList.GetAt(0);

                // HEADER FROM BILL PAYMENT CHECK
                DateTime payDate = bp.TxnDate?.GetValue() ?? DateTime.MinValue;
                string payee = bp.PayeeEntityRef?.FullName?.GetValue() ?? "";
                string address1 = bp.Address?.Addr1?.GetValue() ?? "";
                string address2 = bp.Address?.Addr2?.GetValue() ?? "";
                string bankAccount = bp.BankAccountRef?.FullName?.GetValue() ?? "";
                string memo = bp.Memo?.GetValue() ?? "";
                double amountPaid = bp.Amount?.GetValue() ?? 0;

                // ====================================================
                // GET ALL APPLIED BILL TxnIDs
                // ====================================================
                List<string> appliedTxnIDs = new List<string>();

                if (bp.AppliedToTxnRetList != null && bp.AppliedToTxnRetList.Count > 0)
                {
                    Console.WriteLine($"[DEBUG] AppliedToTxn List Count: {bp.AppliedToTxnRetList.Count}");
                    for (int k = 0; k < bp.AppliedToTxnRetList.Count; k++)
                    {
                        var applied = bp.AppliedToTxnRetList.GetAt(k);
                        string tId = applied.TxnID?.GetValue();
                        if (!string.IsNullOrEmpty(tId))
                        {
                            appliedTxnIDs.Add(tId);
                            Console.WriteLine($"[DEBUG] Found Applied Bill TxnID: {tId}");
                        }
                    }
                }
                else
                {
                    MessageBox.Show("No Applied Bill found from Bill Payment Check.");
                    return bills;
                }

                // ====================================================
                // 2. QUERY BILL(S) USING THE COLLECTED TxnIDs
                // ====================================================
                IMsgSetRequest req2 = sessionManager.CreateMsgSetRequest("US", 13, 0);
                req2.Attributes.OnError = ENRqOnError.roeContinue;

                IBillQuery billQuery = req2.AppendBillQueryRq();
                billQuery.IncludeLineItems.SetValue(true);

                foreach (string id in appliedTxnIDs)
                {
                    billQuery.ORBillQuery.TxnIDList.Add(id);
                }

                Console.WriteLine($"[DEBUG] Sending Bill Query for {appliedTxnIDs.Count} bills...");
                IMsgSetResponse resp2 = sessionManager.DoRequests(req2);
                IResponse r2 = resp2.ResponseList.GetAt(0);

                IBillRetList billList = r2.Detail as IBillRetList;

                if (billList == null || billList.Count == 0)
                {
                    MessageBox.Show("Bills not found for the provided TxnIDs.");
                    return bills;
                }

                // ====================================================
                // LOOP THROUGH ALL RETRIEVED BILLS
                // ====================================================
                Console.WriteLine($"[DEBUG] Retrieved {billList.Count} Bill(s). Processing...");

                for (int bIndex = 0; bIndex < billList.Count; bIndex++)
                {
                    IBillRet bill = billList.GetAt(bIndex);

                    // BILL HEADER FIELDS
                    DateTime billDate = bill.TxnDate?.GetValue() ?? DateTime.MinValue;
                    DateTime dueDate = bill.DueDate?.GetValue() ?? DateTime.MinValue;
                    double amountDue = bill.AmountDue?.GetValue() ?? 0;
                    string billMemo = bill.Memo?.GetValue() ?? "";
                    string billAPAccount = bill.APAccountRef?.FullName?.GetValue() ?? "";
                    string billRefNumber = bill.RefNumber?.GetValue() ?? "";
                    string specificTxnID = bill.TxnID?.GetValue() ?? "";

                    // Get AccountNumber for A/P Account Ref if available
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

                    Console.WriteLine($"[DEBUG] Processing Bill #{bIndex + 1}: Ref {billRefNumber}");

                    BillTable bt = new BillTable
                    {
                        DateCreated = payDate,
                        DueDate = payDate,
                        PayeeFullName = payee,
                        Address = address1,
                        Address2 = address2,
                        BankAccount = bankAccount,
                        APAccountRefFullName = billAPAccount,
                        AccountNumber = billAccNum, // Header AccountNumber
                        Amount = amountPaid,
                        RefNumber = refNumber,
                        AppliedRefNumber = billRefNumber,
                        AppliedToTxnTxnID = specificTxnID,
                        Memo = memo,
                        BillMemo = billMemo,
                        AmountDue = amountDue,
                    };

                    // Process Expense Lines for THIS bill
                    if (bill.ExpenseLineRetList != null)
                    {
                        for (int i = 0; i < bill.ExpenseLineRetList.Count; i++)
                        {
                            var exp = bill.ExpenseLineRetList.GetAt(i);
                            string expAccountName = exp.AccountRef?.FullName?.GetValue() ?? "";
                            string expListID = exp.AccountRef?.ListID?.GetValue() ?? "";

                            // Lookup Account Number from dictionary
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
                                ExpenseLineAccountNumber = expAccNumber, // Expense Line AccountNumber
                                ExpenseLineAmount = exp.Amount?.GetValue() ?? 0,
                                ExpenseLineClassRefFullName = exp.ClassRef?.FullName?.GetValue() ?? "",
                                ExpenseLineCustomerJob = exp.CustomerRef?.FullName?.GetValue() ?? "",
                                ExpenseLineMemo = exp.Memo?.GetValue() ?? "",
                            });
                        }
                    }

                    // Process Item Lines for THIS bill
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

                Console.WriteLine($"[DEBUG] Successfully added {bills.Count} bills to the return list.");
            }
            catch (Exception ex)
            {
                Console.WriteLine($"[DEBUG] EXCEPTION: {ex.Message}");
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

        // ====================================================
        // HELPER METHOD TO QUERY QUICKBOOKS FOR ACCOUNT NUMBERS
        // ====================================================
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
                Console.WriteLine($"[DEBUG] Error fetching Account Numbers: {ex.Message}");
            }

            return accountDict;
        }


        public List<CheckTableExpensesAndItems> GetCheckExpensesAndItemsData_DRC(string refNumber)
        {
            QBSessionManager sessionManager = new QBSessionManager();
            List<CheckTableExpensesAndItems> checks = new List<CheckTableExpensesAndItems>();

            try
            {
                Console.WriteLine("--- Starting QuickBooks Session ---");
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

                Console.WriteLine($"Querying for RefNumber starting with: {refNumber}");
                IMsgSetResponse response = sessionManager.DoRequests(request);
                IResponse qbResponse = response.ResponseList.GetAt(0);

                ICheckRetList list = qbResponse.Detail as ICheckRetList;

                if (list == null || list.Count == 0)
                {
                    Console.WriteLine("No checks found.");
                    return checks;
                }

                Console.WriteLine($"Found {list.Count} check(s).");

                // -----------------------------------------------------------------
                // FETCH ALL ACCOUNT NUMBERS FROM QUICKBOOKS
                // -----------------------------------------------------------------
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
                    string currentRef = check.RefNumber?.GetValue() ?? "";
                    string duedate = check.TxnDate?.GetValue().ToString("yyyy-MM-dd") ?? "";

                    Console.WriteLine($"\n[Check #{i + 1}] Ref: {currentRef} | Payee: {payee} | Total: {totalAmount}");

                    // EXPENSE LINES
                    if (check.ExpenseLineRetList != null)
                    {
                        for (int e = 0; e < check.ExpenseLineRetList.Count; e++)
                        {
                            IExpenseLineRet exp = check.ExpenseLineRetList.GetAt(e);

                            string expAccount = exp.AccountRef?.FullName?.GetValue() ?? "";
                            string expListID = exp.AccountRef?.ListID?.GetValue() ?? "";
                            double expAmount = exp.Amount?.GetValue() ?? 0;

                            // Match AccountNumber from dictionary lookup using ListID or FullName
                            string accNumber = "";
                            if (!string.IsNullOrEmpty(expListID) && accountNumbersDict.ContainsKey(expListID))
                            {
                                accNumber = accountNumbersDict[expListID];
                            }
                            else if (!string.IsNullOrEmpty(expAccount) && accountNumbersDict.ContainsKey(expAccount))
                            {
                                accNumber = accountNumbersDict[expAccount];
                            }

                            Console.WriteLine($"   -> [Expense Line] Account: {expAccount} | AccountNum: {accNumber} | Amount: {expAmount}");

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
                                AccountNumber = accNumber, // Assigned AccountNumber
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
                            // 1. Cast to the "OR" wrapper first
                            IORItemLineRet orItemLine = (IORItemLineRet)check.ORItemLineRetList.GetAt(iLine);

                            // 2. Check if the wrapper contains a standard ItemLineRet
                            if (orItemLine.ItemLineRet != null)
                            {
                                IItemLineRet item = orItemLine.ItemLineRet;

                                string itemName = item.ItemRef?.FullName?.GetValue() ?? "";
                                double itemAmount = item.Amount?.GetValue() ?? 0;

                                Console.WriteLine($"   -> [Item Line] Item: {itemName} | Amount: {itemAmount}");

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
                            else if (orItemLine.ItemGroupLineRet != null)
                            {
                                Console.WriteLine("   -> [Item Group] Found a Group/Bundle (Skipping logic not implemented)");
                            }
                        }
                    }
                }
            }
            catch (Exception ex)
            {
                Console.WriteLine($"CRITICAL ERROR: {ex.Message}");
                Console.WriteLine(ex.StackTrace);
                MessageBox.Show($"Error: {ex.Message}");
            }
            finally
            {
                Console.WriteLine("--- Closing Session ---");
                try { sessionManager.EndSession(); sessionManager.CloseConnection(); }
                catch { }
            }

            return checks;
        }

        // -----------------------------------------------------------------
        // HELPER METHOD TO QUERY QUICKBOOKS FOR ACCOUNT NUMBERS
        // -----------------------------------------------------------------
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


        public List<JournalGridItem> GetJournalEntryForGrid(string refNumber)
        {
            QBSessionManager sessionManager = new QBSessionManager();
            List<JournalGridItem> gridItems = new List<JournalGridItem>();

            try
            {
                Console.WriteLine("--- [START] DATA RETRIEVAL ---");

                sessionManager.OpenConnection2("", "QB Journal Grid", ENConnectionType.ctLocalQBD);
                sessionManager.BeginSession("", ENOpenMode.omDontCare);

                IMsgSetRequest request = sessionManager.CreateMsgSetRequest("US", 13, 0);
                request.Attributes.OnError = ENRqOnError.roeContinue;

                IJournalEntryQuery jeQuery = request.AppendJournalEntryQueryRq();

                // 1. QUERY BROADLY
                // We are forced to use mcStartsWith because your SDK lacks mcValues
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

                        // 2. FILTER STRICTLY (Manual Exact Match)
                        // If QuickBooks returns "JV0010" but we wanted "JV001", skip it.
                        if (docNum != refNumber)
                        {
                            continue;
                        }

                        // If we get here, it is the correct RefNumber. Extract lines.
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
                                    item.Memo = Truncate(line.Memo?.GetValue() ?? "", 255);
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

                        // 3. STOP IMMEDIATELY
                        // We found one "JV001". Even if there is a duplicate "JV001" later in the list, 
                        // we ignore it to prevent the Double Table issue.
                        break;
                    }
                }
            }
            catch (Exception ex)
            {
                Console.WriteLine($"ERROR: {ex.Message}");
            }
            finally
            {
                try { sessionManager.EndSession(); sessionManager.CloseConnection(); } catch { }
            }

            return gridItems;
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

                    if (result != null)
                    {
                        int currentID = Convert.ToInt32(result);
                        // Increment the ID
                        //incrementalID = "CV" + currentID.ToString("D6"); // Format to CV000001
                        incrementalID = currentID; // Format to CV000001
                    }
                    else
                    {
                        // If no record exists, create one with FormattedID set to 0
                        query = "INSERT INTO CVIncrement (FormattedID) VALUES (0)";
                        command = new OleDbCommand(query, connection);
                        int rowsAffected = command.ExecuteNonQuery();

                        if (rowsAffected > 0)
                        {
                            incrementalID = 0;
                        }
                        else
                        {
                            Console.WriteLine("Error creating a new record.");
                        }
                    }
                }
                catch (Exception ex)
                {
                    Console.WriteLine("Error: " + ex.Message);
                }
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

                    if (result != null)
                    {
                        int currentID = Convert.ToInt32(result);
                        // Increment the ID
                        //incrementalID = "CV" + currentID.ToString("D6"); // Format to CV000001
                        incrementalID = currentID; // Format to CV000001
                    }
                    else
                    {
                        // If no record exists, create one with FormattedID set to 0
                        query = "INSERT INTO APVIncrement (FormattedID) VALUES (0)";
                        command = new OleDbCommand(query, connection);
                        int rowsAffected = command.ExecuteNonQuery();

                        if (rowsAffected > 0)
                        {
                            incrementalID = 0;
                        }
                        else
                        {
                            Console.WriteLine("Error creating a new record.");
                        }
                    }
                }
                catch (Exception ex)
                {
                    Console.WriteLine("Error: " + ex.Message);
                }
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
