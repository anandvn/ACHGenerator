using System;
using System.Collections.Generic;
using System.ComponentModel;

namespace ACHGenerator
{

    public enum VendorExtendedFields
    {
        [Description("Payee Type")]
        PayeeType,
        [Description("Payee Routing")]
        PayeeRoutingNum,
        [Description("Payee Account")]
        PayeeAccountNum,
        [Description("Payee Acct Type")]
        PayeeAcctType,
        [Description("ACH Active")]
        PayeeACHActive,
    }

    public class BillPayment
    {
        public string VendorListID { get; set; }
        public string PaymentTxnId { get; set; }
        public string PaymentEditSeq { get; set; }
        public string PayeeName { get; set; }
        public string PayeeType { get; set; }
        public bool ACHActive { get; set; }
        public const string PayeePrivacy = "N";
        public string PayeeRoutingNum { get; set; }
        public string PayeeAccountNum { get; set; }
        public string PayeeAccountType { get; set; }
        public DateTime PaymentDate { get; set; }
        public decimal PaymentAmount { get; set; }
        public const string CreditDebit = "C";
        public string PayeeNote { get; set; }
        public string OriginRoutingNum { get; set; }
        public string OriginAccountNum { get; set; }
        public List<string> AppliedToTxnIds { get; set; }

        public BillPayment()
        {
            AppliedToTxnIds = new List<string>();
        }
    }
}
