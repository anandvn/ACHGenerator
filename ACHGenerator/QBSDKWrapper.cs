using QBFC15Lib;
using QBSDKWrapper;
using System;
using System.Collections.Generic;
using System.Collections.ObjectModel;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace ACHGenerator
{
    public class QBSDKWrapper : QBSDKWrapperBase
    {
        public QBSDKWrapper() : base(ACHGenerator.Properties.Settings.Default.AppId.ToString(), ACHGenerator.Properties.Settings.Default.AppName)
        {

        }
        public async Task<Status> CreateVendorCustomFields()
        {
            try
            {
                IMsgSetRequest requestSet = sessionMgr.getMsgSetRequest();
                requestSet.Attributes.OnError = ENRqOnError.roeStop;
                foreach (var field in Enum.GetValues(typeof(VendorExtendedFields)).Cast<VendorExtendedFields>())
                {
                    IDataExtDefAdd DataExtDefAddReq = requestSet.AppendDataExtDefAddRq();
                    DataExtDefAddReq.DataExtName.SetValue(field.ToDescription());
                    DataExtDefAddReq.DataExtType.SetValue(ENDataExtType.detSTR255TYPE);
                    DataExtDefAddReq.OwnerID.SetValue("0");
                    DataExtDefAddReq.AssignToObjectList.Add(ENAssignToObject.atoVendor);
                }

                Status retstatus = new Status("Unknown Error Saving to Quickbooks.  Please contact Support.", ErrorCode.SaveToQBError, 0);

                bool result = await Task.Run(() =>
                {
                    bool retval = true;
                    try
                    {
                        retval = sessionMgr.doRequests(ref requestSet);

                    }
                    catch (Exception except)
                    {
                        log.Error("CreateVendorCustomFields:doRequests", except);
                    }
                    return retval;
                });
                if (!result)
                {
                    retstatus = new Status("Custom fields created successfully", ErrorCode.SavetoQBOK, 100);
                }
                else
                {
                    IResponseList responselist = sessionMgr.getResponseList();
                    if (responselist.Count > 0)
                    {
                        IResponse response = responselist.GetAt(0);
                        retstatus = new Status(response.StatusMessage, ErrorCode.SaveToQBError, 0);
                    }
                }
                return retstatus;
            }
            catch (Exception except)
            {
                log.Error("CreateVendorCustomFields", except);
                return new Status(except.Message, ErrorCode.SaveToQBError, 0);
            }
        }

        public async Task<Status<ObservableCollection<BillPayment>>> FetchBillPayments(DateTime start, DateTime end)
        {
            try
            {
                IMsgSetRequest requestSet = sessionMgr.getMsgSetRequest();
                requestSet.Attributes.OnError = ENRqOnError.roeStop;

                Status<ObservableCollection<BillPayment>> retstatus;
                IBillPaymentCheckQuery billPaymentReq = requestSet.AppendBillPaymentCheckQueryRq();
                billPaymentReq.ORTxnQuery.TxnFilter.ORDateRangeFilter.TxnDateRangeFilter.ORTxnDateRangeFilter.TxnDateFilter.FromTxnDate.SetValue(start);
                billPaymentReq.ORTxnQuery.TxnFilter.ORDateRangeFilter.TxnDateRangeFilter.ORTxnDateRangeFilter.TxnDateFilter.FromTxnDate.SetValue(end);
                billPaymentReq.IncludeLineItems.SetValue(true);


                bool result = await Task.Run(() =>
                {
                    bool retval = true;
                    try
                    {
                        retval = sessionMgr.doRequests(ref requestSet);
                    }
                    catch (Exception except)
                    {
                        log.Error("fetchBillPayments:doRequests", except);
                    }
                    return retval;
                });

                if (!result)
                {
                    IResponseList responselist = sessionMgr.getResponseList();
                    ObservableCollection<BillPayment> payments = new ObservableCollection<BillPayment>();
                    string responsemessage = "An unknown error has occurred";
                    if (responselist.Count > 0)
                    {
                        IResponse response = responselist.GetAt(0);
                        if (response.Detail is IBillPaymentCheckRetList paymentList)
                        {
                            for (int i = 0; i < paymentList.Count; i++)
                            {
                                IBillPaymentCheckRet paymentChk = paymentList.GetAt(i);
                                if (paymentChk != null)
                                {
                                    string payeenote = string.Empty;
                                    List<string> txnids = new List<string>();
                                    if (paymentChk.AppliedToTxnRetList != null)
                                    {
                                        for (int j = 0; j < paymentChk.AppliedToTxnRetList.Count; j++)
                                        {
                                            IAppliedToTxnRet appliedRef = paymentChk.AppliedToTxnRetList.GetAt(j);
                                            string pmtref = appliedRef.RefNumber != null ? appliedRef.RefNumber.GetValue() : String.Empty;
                                            if (j == 0)
                                                payeenote = pmtref.StripSpecialChars();
                                            else
                                                payeenote += " " + pmtref.StripSpecialChars();
                                            txnids.Add(appliedRef.TxnID.GetValue());
                                        }
                                    }
                                    payments.Add(new BillPayment()
                                    {
                                        PaymentAmount = (decimal)paymentChk.Amount.GetValue(),
                                        VendorListID = paymentChk.PayeeEntityRef.ListID.GetValue(),
                                        PaymentDate = paymentChk.TxnDate.GetValue(),
                                        PaymentTxnId = paymentChk.TxnID.GetValue(),
                                        PayeeNote = payeenote,
                                        PaymentEditSeq = paymentChk.EditSequence.GetValue(),
                                        AppliedToTxnIds = txnids,
                                    });

                                }
                            }
                        }
                        responsemessage = response.StatusMessage;
                    }
                    retstatus = new Status<ObservableCollection<BillPayment>>(responsemessage, ErrorCode.SavetoQBOK, 0, payments);
                }
                else
                {
                    IResponseList responselist = sessionMgr.getResponseList();
                    if (responselist.Count > 0)
                    {
                        IResponse response = responselist.GetAt(0);
                        retstatus = new Status<ObservableCollection<BillPayment>>(response.StatusMessage, ErrorCode.SaveToQBError, 0, null);
                    }
                    else
                    {
                        retstatus = new Status<ObservableCollection<BillPayment>>("Unknown Error", ErrorCode.SaveToQBError, 0, null);
                    }
                }
                return retstatus;
            }
            catch (Exception except)
            {
                log.Error("fetchBillPayments", except);
                return new Status<ObservableCollection<BillPayment>>(except.Message, ErrorCode.SaveToQBError, 0, null);
            }
        }

        public async Task<Status> FetchPayeeInfo(ObservableCollection<BillPayment> payments)
        {
            try
            {
                IMsgSetRequest requestSet = sessionMgr.getMsgSetRequest();
                requestSet.Attributes.OnError = ENRqOnError.roeStop;
                IVendorQuery vendorQuery = requestSet.AppendVendorQueryRq();
                vendorQuery.OwnerIDList.Add("0");
                foreach (BillPayment payment in payments)
                {
                    vendorQuery.ORVendorListQuery.ListIDList.Add(payment.VendorListID);
                }

                Status retstatus;

                bool result = await Task.Run(() =>
                {
                    bool retval = true;
                    try
                    {
                        retval = sessionMgr.doRequests(ref requestSet);

                    }
                    catch (Exception except)
                    {
                        log.Error("fetchPayeeInfo:doRequests", except);
                    }
                    return retval;
                });

                if (!result)
                {
                    if ((sessionMgr.getResponse(0) is IVendorRetList vendorretlist) && (vendorretlist.Count != 0))
                    {
                        for (int ndx = 0; ndx <= (vendorretlist.Count - 1); ndx++)
                        {
                            IVendorRet vendorRet = vendorretlist.GetAt(ndx);
                            string listid = vendorRet.ListID.GetValue();
                            BillPayment pmt = payments.Where(x => x.VendorListID == listid).FirstOrDefault();
                            if (pmt == null) { continue; }
                            string payeename = string.Empty;
                            payeename = vendorRet.NameOnCheck?.IsEmpty() == false ? vendorRet.NameOnCheck.GetValue() : string.Empty;

                            if (payeename == string.Empty)
                                payeename = vendorRet.CompanyName?.IsEmpty() == false ? vendorRet.CompanyName.GetValue() : string.Empty;

                            if (payeename == string.Empty)
                                payeename = vendorRet.Name.GetValue();

                            pmt.PayeeName = payeename.StripSpecialChars();

                            if (vendorRet.DataExtRetList == null) { continue; }

                            for (int i = 0; i < vendorRet.DataExtRetList.Count; i++)
                            {
                                IDataExtRet dataext = vendorRet.DataExtRetList.GetAt(i);
                                string extname = dataext.DataExtName.GetValue();
                                string extvalue = dataext.DataExtValue.GetValue();

                                if (extname == VendorExtendedFields.PayeeType.ToDescription())
                                {
                                    pmt.PayeeType = extvalue;
                                    continue;
                                }
                                if (extname == VendorExtendedFields.PayeeAcctType.ToDescription())
                                {
                                    pmt.PayeeAccountType = extvalue;
                                    continue;
                                }
                                if (extname == VendorExtendedFields.PayeeACHActive.ToDescription())
                                {
                                    pmt.ACHActive = extvalue == "Y";
                                    continue;
                                }
                                if (extname == VendorExtendedFields.PayeeAccountNum.ToDescription())
                                {
                                    pmt.PayeeAccountNum = extvalue;
                                    continue;
                                }
                                if (extname == VendorExtendedFields.PayeeRoutingNum.ToDescription())
                                {
                                    pmt.PayeeRoutingNum = extvalue;
                                    continue;
                                }
                            }
                        }
                    }
                    retstatus = new Status("Download Payee Data Successful", ErrorCode.SavetoQBOK, 100);
                }
                else
                {
                    IResponseList responselist = sessionMgr.getResponseList();
                    Console.WriteLine($"response count: {responselist.Count}");
                    retstatus = new Status("Response List is Empty", ErrorCode.SaveToQBError, 100);
                    if (responselist.Count > 0)
                    {
                        IResponse response = responselist.GetAt(0);
                        retstatus.Message = response.StatusMessage;
                    }
                }
                return retstatus;
            }
            catch (Exception except)
            {
                log.Error("fetchBillPayments", except);
                Console.WriteLine(except.Message);
                return new Status("Error Downloading Vendor Information", ErrorCode.SaveToQBError, 100);
            }
        }

        public async Task<Status> UpdateBillPayments(ObservableCollection<BillPayment> billPayments)
        {
            try
            {
                IMsgSetRequest requestSet = sessionMgr.getMsgSetRequest();
                requestSet.Attributes.OnError = ENRqOnError.roeStop;

                Status retstatus;
                foreach (var payment in billPayments.Where(x => x.ACHActive))
                {
                    IBillPaymentCheckMod billPaymentReq = requestSet.AppendBillPaymentCheckModRq();
                    billPaymentReq.TxnID.SetValue(payment.PaymentTxnId);
                    //Set field value for EditSequence
                    billPaymentReq.EditSequence.SetValue(payment.PaymentEditSeq);
                    billPaymentReq.ORCheckPrint.RefNumber.SetValue("ACH");
                }


                bool result = await Task.Run(() =>
                {
                    bool retval = true;
                    try
                    {
                        retval = sessionMgr.doRequests(ref requestSet);
                    }
                    catch (Exception except)
                    {
                        log.Error("UpdateBillPayments:doRequests", except);
                        Console.Write(except.Message);
                        if (except.InnerException != null)
                            Console.WriteLine(except.InnerException.ToString());
                    }
                    return retval;
                });

                if (!result)
                {
                    IResponseList responselist = sessionMgr.getResponseList();
                    string responsemessage = "Response List is Empty";
                    if (responselist.Count > 0)
                    {
                        IResponse response = responselist.GetAt(0);
                        responsemessage = response.StatusMessage;
                    }
                    retstatus = new Status(responsemessage, ErrorCode.SavetoQBOK, 0);
                }
                else
                {
                    IResponseList responselist = sessionMgr.getResponseList();
                    if (responselist.Count > 0)
                    {
                        IResponse response = responselist.GetAt(0);
                        retstatus = new Status(response.StatusMessage, ErrorCode.SaveToQBError, 0);
                    }
                    else
                    {
                        retstatus = new Status("Unknown Error", ErrorCode.SaveToQBError, 0);
                    }
                }
                return retstatus;
            }
            catch (Exception except)
            {
                log.Error("fetchBillPayments", except);
                return new Status(except.Message, ErrorCode.SaveToQBError, 0);
            }

        }

    }
}
