using CommandLine;
using QBSDKWrapper;
using System;
using System.Collections.Generic;
using System.Collections.ObjectModel;
using System.IO;
using System.Threading.Tasks;

namespace ACHGenerator
{
    internal class Options
    {
        public Options() { }
        [Option('c', "companyfile", Required = true, HelpText = "Set Company file to access")]
        public string CompanyFile { get; set; }
    }

    [Verb("authorize", HelpText = "Open Company File to authorize connection")]
    internal class AuthOptions : Options
    {
        public AuthOptions() { }
    }

    [Verb("gen", HelpText = "Generate CSV file")]
    internal class ACHGenOptions : Options
    {
        public ACHGenOptions() { }
        [Option('o', "output", HelpText = "Output CSV File", Required = true)]
        public string Output { get; set; }
        [Option('d', "date", HelpText = "Check Date", Required = true)]
        public DateTime CheckDate { get; set; }
    }

    [Verb("init", HelpText = "Initial QB File with Fields required")]
    internal class CreateFieldsOptions : Options
    {
        public CreateFieldsOptions() { }
    }


    internal class Program
    {        
        static async Task<int> Main(string[] args)
        {
            return await CommandLine.Parser.Default.ParseArguments<AuthOptions, CreateFieldsOptions, ACHGenOptions>(args)
                .MapResult(
                    (AuthOptions opts) => authorizeApp(opts),
                    (ACHGenOptions opts) => generateACH(opts),
                    (CreateFieldsOptions opts) => createCustomFields(opts),
                    errs => Task.FromResult(-1));
        }

        private static async Task<int> authorizeApp(AuthOptions opts)
        {
            if (!System.IO.File.Exists(opts.CompanyFile))
            {
                Console.WriteLine($"Company File does not exist: {opts.CompanyFile}");
                return 1;
            }
            Console.WriteLine("Open Quickbooks Company file as admin in Multi-User Mode and press any key to continue...");
            Console.ReadKey();
            Console.WriteLine($"Connecting to {opts.CompanyFile}");
            Status status;
            using (QBSDKWrapper qbconnector = new QBSDKWrapper())
            {
                status = await qbconnector.ConnectAsync(opts.CompanyFile);
                qbconnector.Disconnect();
            }
            Console.WriteLine(status.GetFormattedMessage());
            return status?.Code == ErrorCode.ConnectQBOK ? 0 : 1;
        }

        private static async Task<int> createCustomFields(CreateFieldsOptions opts)
        {
            if (!System.IO.File.Exists(opts.CompanyFile))
            {
                Console.WriteLine($"Company File does not exist: {opts.CompanyFile}");
                return 1;
            }
            Console.WriteLine("Open Quickbooks Company file as admin in Single User Mode and press any key to continue...");
            Console.ReadKey();
            Console.WriteLine($"Connecting to {opts.CompanyFile}");
            Status status;
            using (QBSDKWrapper qbconnector = new QBSDKWrapper())
            {
                status = await qbconnector.ConnectAsync(opts.CompanyFile, true);
                Console.WriteLine(status.GetFormattedMessage());
                if (status.Code == ErrorCode.ConnectQBOK)
                {
                    status = await qbconnector.CreateVendorCustomFields();
                    Console.WriteLine(status.GetFormattedMessage());
                }
                qbconnector.Disconnect();
            }
            Console.WriteLine(status.GetFormattedMessage());
            return status?.Code == ErrorCode.ConnectQBOK ? 0 : 1;
        }

        private static async Task<int> generateACH(ACHGenOptions opts)
        {
            int retstatus = -1;
     
            using (QBSDKWrapper qbconnector = new QBSDKWrapper())
            {
                Status status = await qbconnector.ConnectAsync(opts.CompanyFile);
                Console.WriteLine(status.GetFormattedMessage());
                if (status.Code == ErrorCode.ConnectQBOK)
                {
                    Console.WriteLine($"Fetching Bill Payments for {opts.CheckDate:d}..."); 
                    Status<ObservableCollection<BillPayment>> fetchstatus = await qbconnector.FetchBillPayments(opts.CheckDate, opts.CheckDate);
                    Console.WriteLine($"Result: {fetchstatus.GetFormattedMessage()}");
                    if ((fetchstatus.Code == ErrorCode.SavetoQBOK) && (fetchstatus.ReturnObject?.Count > 0))
                    {
                        Console.WriteLine($"Updating Vendor information for {fetchstatus.ReturnObject.Count} bills...");
                        Status updatestatus = await qbconnector.FetchPayeeInfo(fetchstatus.ReturnObject);
                        Console.WriteLine($"Result: {updatestatus.GetFormattedMessage()}");
                        using (StreamWriter sw = new StreamWriter(opts.Output))
                        {
                            Console.Write("Writing CSV file");
                            foreach (BillPayment payment in fetchstatus.ReturnObject)
                            {

                                if (payment.ACHActive == true)
                                {
                                    //Console.WriteLine($"{payment.PayeeType},N,{payment.PayeeName},{payment.PayeeRoutingNum:9},{payment.PayeeAccountNum:34},{payment.PayeeAccountType:1},{payment.PaymentDate:d},{payment.PaymentAmount:F2},C,{payment.PayeeNote:80}");
                                    sw.WriteLine($"{payment.PayeeType},N,{payment.PayeeName},{payment.PayeeRoutingNum:9},{payment.PayeeAccountNum:34},{payment.PayeeAccountType:1},{payment.PaymentDate:d},{payment.PaymentAmount:F2},C,{payment.PayeeNote:80}");
                                    Console.Write(".");
                                }
                            }
                        }
                        Console.WriteLine();
                        Console.WriteLine("Updating bill payment reference for ACH payments...");
                        updatestatus = await qbconnector.UpdateBillPayments(fetchstatus.ReturnObject);
                        Console.WriteLine($"Result: {updatestatus.GetFormattedMessage()}");
                        retstatus = 0;
                    }
                    else
                    {
                        Console.WriteLine("No Bill Payments Found.");
                    }
                }
                qbconnector.Disconnect();
            }
            return retstatus;
        }
    }
}
