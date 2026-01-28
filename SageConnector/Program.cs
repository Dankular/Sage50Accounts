using System.Runtime.InteropServices;
using System.Reflection;
using System.Diagnostics;
using System.IO.Compression;

namespace SageConnector;

/// <summary>
/// Sage 50 SDO (Sage Data Objects) Connector
/// Uses late-binding COM interop to connect to Sage 50 Accounts
///
/// Prerequisites:
/// - Sage 50 Accounts must be installed on this machine
/// - The SDO Engine must be registered (SDOEngine.32)
/// - You need a valid Sage 50 company data path
/// </summary>
class Program
{
    static void Main(string[] args)
    {
        Console.WriteLine("Sage 50 SDO Connector");
        Console.WriteLine(new string('=', 50));

        // If "explore" is passed, explore the interop assembly
        if (args.Length > 0 && args[0].Equals("explore", StringComparison.OrdinalIgnoreCase))
        {
            TypeExplorer.ExploreInteropAssembly();
            return;
        }

        // SDK management commands (don't require connection)
        if (args.Length > 0 && args[0].Equals("sdk", StringComparison.OrdinalIgnoreCase))
        {
            var subCmd = args.Length > 1 ? args[1].ToLowerInvariant() : "status";

            switch (subCmd)
            {
                case "list":
                    SdkManager.ListAvailableVersions();
                    break;

                case "detect":
                    var detectPath = args.Length > 2 ? args[2] : null;
                    if (string.IsNullOrEmpty(detectPath))
                    {
                        Console.WriteLine("Usage: sdk detect <accdata_path>");
                        return;
                    }
                    var version = SdkManager.DetectVersion(detectPath);
                    if (version != null)
                    {
                        var url = SdkManager.GetDownloadUrl(version);
                        Console.WriteLine($"\nDownload URL: {url}");
                    }
                    break;

                case "download":
                    var dlVersion = args.Length > 2 ? args[2] : null;
                    if (string.IsNullOrEmpty(dlVersion))
                    {
                        Console.WriteLine("Usage: sdk download <version>");
                        Console.WriteLine("Example: sdk download 32.0");
                        return;
                    }
                    SdkManager.DownloadSdkAsync(dlVersion).GetAwaiter().GetResult();
                    break;

                case "install":
                    var instPath = args.Length > 2 ? args[2] : null;
                    var forceInstall = args.Any(a => a.Equals("-force", StringComparison.OrdinalIgnoreCase));
                    if (string.IsNullOrEmpty(instPath))
                    {
                        Console.WriteLine("Usage: sdk install <accdata_path> [-force]");
                        Console.WriteLine("Detects version and installs appropriate SDK");
                        Console.WriteLine("  -force: Re-download even if already extracted");
                        return;
                    }
                    SdkManager.EnsureSdkAsync(instPath, forceInstall).GetAwaiter().GetResult();
                    break;

                case "test":
                    // Test registration-free COM by forcing it even if registered COM exists
                    var testPath = args.Length > 2 ? args[2] : null;
                    if (string.IsNullOrEmpty(testPath))
                    {
                        Console.WriteLine("Usage: sdk test <accdata_path>");
                        Console.WriteLine("Tests registration-free COM loading (bypasses registered SDK)");
                        return;
                    }

                    Console.WriteLine("\n[Testing Registration-Free COM]");

                    // Force extraction even if registered SDK exists
                    if (!SdkManager.EnsureSdkAsync(testPath, force: true).GetAwaiter().GetResult())
                    {
                        Console.WriteLine("Failed to prepare SDK");
                        return;
                    }

                    Console.WriteLine($"\n  ExtractedSdkPath: {SdkManager.ExtractedSdkPath}");

                    // Now try to create connection forcing reg-free COM
                    try
                    {
                        Console.WriteLine("\n  Creating SageConnection with forceRegFreeCom=true...");
                        using var testSage = new SageConnection(forceRegFreeCom: true);
                        Console.WriteLine("\n  SUCCESS: Registration-free COM is working!");
                    }
                    catch (Exception ex)
                    {
                        Console.WriteLine($"\n  FAILED: {ex.Message}");
                    }
                    break;

                case "status":
                default:
                    Console.WriteLine("\n[SDK Status]");
                    Console.WriteLine($"  SDOEngine.32 registered: {SdkManager.IsSdkAvailable()}");
                    Console.WriteLine($"  Extracted SDK path: {SdkManager.ExtractedSdkPath ?? "(none)"}");
                    Console.WriteLine("\nSDK Commands:");
                    Console.WriteLine("  sdk status              - Check if SDK is available");
                    Console.WriteLine("  sdk list                - List downloadable SDK versions");
                    Console.WriteLine("  sdk detect <path>       - Detect version from ACCDATA");
                    Console.WriteLine("  sdk download <version>  - Download SDK installer");
                    Console.WriteLine("  sdk install <path>      - Auto-detect and extract SDK");
                    Console.WriteLine("  sdk test <path>         - Test registration-free COM");
                    break;
            }
            return;
        }

        // Check command line args for company path
        string? companyPath = args.Length > 0 ? args[0] : null;

        try
        {
            using var sage = new SageConnection();

            // List available companies if no path specified
            if (string.IsNullOrEmpty(companyPath))
            {
                Console.WriteLine("\nSearching for Sage companies...");
                sage.ListCompanies();

                Console.WriteLine("\nUsage: SageConnector.exe <company_data_path> [username] [password]");
                Console.WriteLine("Example: SageConnector.exe \"\\\\server\\sage\\accdata\" manager \"\"");
                return;
            }

            // Get optional username/password
            string username = args.Length > 1 ? args[1] : "manager";
            string password = args.Length > 2 ? args[2] : "";

            // Connect to the specified company
            Console.WriteLine($"\nConnecting to: {companyPath}");
            sage.Connect(companyPath, username, password);

            // Show company info
            sage.ShowCompanyInfo();

            // Demonstrate reading customers
            sage.ListCustomers(10);

            // Demonstrate reading invoices
            sage.ListInvoices(10);

            // Show how to create an invoice
            sage.ShowInvoiceCreationExample();

            // If "post" argument is provided, create a test invoice
            if (args.Any(a => a.Equals("post", StringComparison.OrdinalIgnoreCase)))
            {
                Console.WriteLine("\n*** POSTING TEST INVOICE ***");

                // Create test line items
                var testItems = new List<InvoiceLineItem>
                {
                    new InvoiceLineItem
                    {
                        Description = "Test Service from SageConnector",
                        Quantity = 1,
                        UnitPrice = 25.00m,
                        NominalCode = "4000",
                        TaxCode = "T1"
                    }
                };

                // Test with existing customer FALCONLO
                var invoiceNum = sage.CreateSalesInvoice(
                    customerAccount: "FALCONLO",
                    orderNumber: $"SC-{DateTime.Now:yyyyMMddHHmmss}",
                    items: testItems,
                    createCustomerIfMissing: false);

                if (!string.IsNullOrEmpty(invoiceNum))
                {
                    Console.WriteLine($"\nInvoice {invoiceNum} created and posted to ledger!");
                }
            }
            // If "newcust" argument is provided, test customer creation
            else if (args.Any(a => a.Equals("newcust", StringComparison.OrdinalIgnoreCase)))
            {
                Console.WriteLine("\n*** TESTING CUSTOMER CREATION ***");

                var testAcct = $"CU{DateTime.Now:ddHHmm}"; // Max 8 chars
                var created = sage.CreateCustomer(testAcct, "Test Customer", "123 Test Street", "Test City", "TE5 T12");

                if (created)
                {
                    Console.WriteLine($"Customer {testAcct} created. Now posting invoice...");

                    var testItems = new List<InvoiceLineItem>
                    {
                        new InvoiceLineItem
                        {
                            Description = "Test service for new customer",
                            Quantity = 1,
                            UnitPrice = 10.00m,
                            NominalCode = "4000",
                            TaxCode = "T1"
                        }
                    };

                    sage.CreateSalesInvoice(testAcct, "NEW-CUST-001", testItems);
                }
            }
            // Test customer lookup
            else if (args.Any(a => a.Equals("lookup", StringComparison.OrdinalIgnoreCase)))
            {
                Console.WriteLine("\n*** TESTING ACCOUNT LOOKUP ***");

                // Search for customers
                var customers = sage.FindCustomers("", 10); // Empty = list all
                Console.WriteLine($"\nFound {customers.Count} customers:");
                foreach (var c in customers)
                {
                    Console.WriteLine($"  {c.AccountRef,-12} {c.Name,-30} Tel: {c.Telephone}");
                    Console.WriteLine($"    {c.Address1}, {c.Postcode}");
                    Console.WriteLine($"    Balance: {c.Balance:C}, Credit Limit: {c.CreditLimit:C}");
                }

                // Search for suppliers
                var suppliers = sage.FindSuppliers("", 10); // Empty = list all
                Console.WriteLine($"\nFound {suppliers.Count} suppliers:");
                foreach (var s in suppliers)
                {
                    Console.WriteLine($"  {s.AccountRef,-12} {s.Name,-30} Tel: {s.Telephone}");
                    Console.WriteLine($"    Balance: {s.Balance:C}");
                }
            }
            // Discover available SDK posting objects
            else if (args.Any(a => a.Equals("discover", StringComparison.OrdinalIgnoreCase)))
            {
                Console.WriteLine("\n*** DISCOVERING SDK POSTING OBJECTS ***");
                sage.DiscoverPostingObjects();
            }
            // Read audit records to understand transaction structure
            else if (args.Any(a => a.Equals("audit", StringComparison.OrdinalIgnoreCase)))
            {
                Console.WriteLine("\n*** READING AUDIT RECORDS ***");
                sage.ReadAuditRecords(20);
            }
            // Test sales invoice posting (direct to ledger via TransactionPost)
            // Usage: sinv [account] [-auto]
            else if (args.Any(a => a.Equals("sinv", StringComparison.OrdinalIgnoreCase)))
            {
                Console.WriteLine("\n*** POSTING SALES INVOICE (via TransactionPost) ***");

                // Parse optional account and -auto flag
                var sinvIndex = Array.FindIndex(args, a => a.Equals("sinv", StringComparison.OrdinalIgnoreCase));
                var autoCreate = args.Any(a => a.Equals("-auto", StringComparison.OrdinalIgnoreCase));
                var customerRef = "FALCONLO"; // default

                // Check if next arg after sinv is an account ref (not a flag)
                if (sinvIndex + 1 < args.Length && !args[sinvIndex + 1].StartsWith("-"))
                {
                    customerRef = args[sinvIndex + 1].ToUpperInvariant();
                    if (customerRef.Length > 8) customerRef = customerRef[..8];
                }

                Console.WriteLine($"  Account: {customerRef}, Auto-create: {autoCreate}");

                // Check if customer exists, create if -auto flag set
                if (!sage.CustomerExists(customerRef))
                {
                    if (autoCreate)
                    {
                        Console.WriteLine($"  Customer '{customerRef}' not found - creating...");
                        var cust = new CustomerAccount
                        {
                            AccountRef = customerRef,
                            Name = $"Auto-created {customerRef}",
                            Address1 = "Auto-created account",
                            Postcode = "AA1 1AA"
                        };
                        if (!sage.CreateCustomerEx(cust))
                        {
                            Console.WriteLine("  Failed to create customer account.");
                            return;
                        }
                        Console.WriteLine($"  Customer '{customerRef}' created.");
                    }
                    else
                    {
                        Console.WriteLine($"  Customer '{customerRef}' not found. Use -auto to create.");
                        return;
                    }
                }

                // Use TransactionPost which updates customer balance (not InvoicePost which just creates documents)
                var success = sage.PostTransactionSI(
                    customerAccount: customerRef,
                    invoiceRef: $"SI-{DateTime.Now:yyyyMMddHHmmss}",
                    netAmount: 100.00m,
                    taxAmount: 20.00m,
                    nominalCode: "4000",
                    details: "Test sales invoice from SageConnector",
                    taxCode: "T1",
                    postToLedger: true);

                if (success)
                    Console.WriteLine("\nSales Invoice posted successfully!");
                else
                    Console.WriteLine("\nSales Invoice posting failed.");
            }
            // Test InvoicePost (creates invoice documents, does not update balance)
            else if (args.Any(a => a.Equals("invdoc", StringComparison.OrdinalIgnoreCase)))
            {
                Console.WriteLine("\n*** CREATING INVOICE DOCUMENT (via InvoicePost) ***");

                var success = sage.PostSalesInvoice(
                    customerAccount: "FALCONLO",
                    invoiceRef: $"DOC-{DateTime.Now:yyyyMMddHHmmss}",
                    netAmount: 100.00m,
                    taxAmount: 20.00m,
                    nominalCode: "4000",
                    details: "Invoice document from SageConnector",
                    taxCode: "T1",
                    postToLedger: false);

                if (success)
                    Console.WriteLine("\nInvoice document created!");
                else
                    Console.WriteLine("\nInvoice document creation failed.");
            }
            // Test purchase invoice posting (direct to ledger)
            // Usage: pinv [account] [-auto]
            else if (args.Any(a => a.Equals("pinv", StringComparison.OrdinalIgnoreCase)))
            {
                Console.WriteLine("\n*** POSTING PURCHASE INVOICE ***");

                // Parse optional account and -auto flag
                var pinvIndex = Array.FindIndex(args, a => a.Equals("pinv", StringComparison.OrdinalIgnoreCase));
                var autoCreate = args.Any(a => a.Equals("-auto", StringComparison.OrdinalIgnoreCase));
                var supplierRef = "FALCONLO"; // default

                // Check if next arg after pinv is an account ref (not a flag)
                if (pinvIndex + 1 < args.Length && !args[pinvIndex + 1].StartsWith("-"))
                {
                    supplierRef = args[pinvIndex + 1].ToUpperInvariant();
                    if (supplierRef.Length > 8) supplierRef = supplierRef[..8];
                }

                Console.WriteLine($"  Account: {supplierRef}, Auto-create: {autoCreate}");

                // Check if supplier exists, create if -auto flag set
                if (!sage.SupplierExists(supplierRef))
                {
                    if (autoCreate)
                    {
                        Console.WriteLine($"  Supplier '{supplierRef}' not found - creating...");
                        var sup = new SupplierAccount
                        {
                            AccountRef = supplierRef,
                            Name = $"Auto-created {supplierRef}",
                            Address1 = "Auto-created account",
                            Postcode = "AA1 1AA"
                        };
                        if (!sage.CreateSupplierEx(sup))
                        {
                            Console.WriteLine("  Failed to create supplier account.");
                            return;
                        }
                        Console.WriteLine($"  Supplier '{supplierRef}' created.");
                    }
                    else
                    {
                        Console.WriteLine($"  Supplier '{supplierRef}' not found. Use -auto to create.");
                        return;
                    }
                }

                // Use TransactionPost TYPE=6 (PI) - same approach as sales invoice TYPE=1 (SI)
                var success = sage.PostTransactionPI(
                    supplierAccount: supplierRef,
                    invoiceRef: $"PI-{DateTime.Now:yyyyMMddHHmmss}",
                    netAmount: 50.00m,
                    taxAmount: 10.00m,
                    nominalCode: "5000",  // Cost of sales - purchase nominals start at 5xxx
                    details: "Test purchase invoice",
                    taxCode: "T1");

                if (success)
                    Console.WriteLine("\nPurchase Invoice posted successfully!");
                else
                    Console.WriteLine("\nPurchase Invoice posting failed.");
            }
            // Test journal posting
            else if (args.Any(a => a.Equals("journal", StringComparison.OrdinalIgnoreCase)))
            {
                Console.WriteLine("\n*** TESTING JOURNAL POSTING ***");

                // Discover JournalPost fields first
                sage.DiscoverJournalPostFields();

                // Post a simple journal: Debit 7500 (Postage), Credit 1200 (Bank)
                var success = sage.PostSimpleJournal(
                    debitNominal: "7500",
                    creditNominal: "1200",
                    amount: 10.00m,
                    reference: $"JNL-{DateTime.Now:yyyyMMddHHmmss}",
                    details: "Test journal from SageConnector",
                    date: DateTime.Today);

                if (success)
                    Console.WriteLine("\nJournal posted successfully!");
                else
                    Console.WriteLine("\nJournal posting failed.");
            }
            // Test bank payment
            else if (args.Any(a => a.Equals("bankpay", StringComparison.OrdinalIgnoreCase)))
            {
                Console.WriteLine("\n*** TESTING BANK PAYMENT ***");

                var success = sage.PostBankPayment(
                    bankNominal: "1200",      // Bank Current Account
                    expenseNominal: "7500",   // Postage
                    netAmount: 15.00m,
                    reference: $"BP-{DateTime.Now:yyyyMMddHHmmss}",
                    details: "Test bank payment from SageConnector",
                    taxCode: "T0");           // No VAT on postage

                if (success)
                    Console.WriteLine("\nBank payment posted successfully!");
                else
                    Console.WriteLine("\nBank payment posting failed.");
            }
            // Test bank receipt
            else if (args.Any(a => a.Equals("bankrec", StringComparison.OrdinalIgnoreCase)))
            {
                Console.WriteLine("\n*** TESTING BANK RECEIPT ***");

                var success = sage.PostBankReceipt(
                    bankNominal: "1200",      // Bank Current Account
                    incomeNominal: "4000",    // Sales
                    netAmount: 50.00m,
                    reference: $"BR-{DateTime.Now:yyyyMMddHHmmss}",
                    details: "Test bank receipt from SageConnector",
                    taxCode: "T1");           // Standard VAT

                if (success)
                    Console.WriteLine("\nBank receipt posted successfully!");
                else
                    Console.WriteLine("\nBank receipt posting failed.");
            }
            // Test nominal codes listing
            else if (args.Any(a => a.Equals("nominals", StringComparison.OrdinalIgnoreCase)))
            {
                sage.ListNominalCodes(100);
            }
            // Create supplier
            else if (args.Any(a => a.Equals("newsup", StringComparison.OrdinalIgnoreCase)))
            {
                Console.WriteLine("\n*** TESTING SUPPLIER CREATION ***");

                // Account ref must be max 8 chars, no spaces
                var supplier = new SupplierAccount
                {
                    AccountRef = $"SP{DateTime.Now:ddHHmm}",
                    Name = "Test Supplier Ltd",
                    Address1 = "456 Supplier Road",
                    Address2 = "Industrial Estate",
                    Address3 = "Test City",
                    Postcode = "TE5 T99",
                    Telephone = "01234 567890",
                    Email = "accounts@testsupplier.com",
                    ContactName = "Jane Smith",
                    CreditLimit = 5000m
                };

                var created = sage.CreateSupplierEx(supplier);

                if (created)
                {
                    Console.WriteLine($"\nSupplier {supplier.AccountRef} created successfully!");

                    // Verify by lookup
                    var found = sage.GetSupplier(supplier.AccountRef);
                    if (found != null)
                        Console.WriteLine($"Verified: {found.Name} at {found.Address1}");
                }
            }
            // Test purchase invoice
            else if (args.Any(a => a.Equals("purchase", StringComparison.OrdinalIgnoreCase)))
            {
                Console.WriteLine("\n*** POSTING PURCHASE INVOICE ***");

                var testItems = new List<InvoiceLineItem>
                {
                    new InvoiceLineItem
                    {
                        Description = "Office Supplies",
                        Quantity = 1,
                        UnitPrice = 50.00m,
                        NominalCode = "5000",  // Purchases nominal
                        TaxCode = "T1"
                    }
                };

                // Try to use existing supplier or create new one
                var supplierRef = "TESTSUP01";
                var invoiceNum = sage.CreatePurchaseInvoice(
                    supplierAccount: supplierRef,
                    invoiceRef: $"PINV-{DateTime.Now:yyyyMMddHHmmss}",
                    items: testItems,
                    createSupplierIfMissing: true,
                    supplierName: "Test Supplier Ltd");

                if (!string.IsNullOrEmpty(invoiceNum))
                {
                    Console.WriteLine($"\nPurchase invoice {invoiceNum} posted!");
                }
            }
            // Test TransactionPost TYPE codes
            else if (args.Any(a => a.Equals("testtypes", StringComparison.OrdinalIgnoreCase)))
            {
                Console.WriteLine("\n*** TESTING TRANSACTIONPOST TYPE VALUES ***");
                Console.WriteLine("Avoiding TYPE 3 (SA - Sales Receipt) and TYPE 6 (PP - Purchase Payment)\n");

                // Ensure FALCONLO exists as supplier
                var testAccount = "FALCONLO";
                if (!sage.SupplierExists(testAccount))
                {
                    Console.WriteLine($"Creating {testAccount} as supplier...");
                    sage.CreateSupplierEx(new SupplierAccount
                    {
                        AccountRef = testAccount,
                        Name = "Falcon Logistics",
                        Address1 = "Test Address",
                        Postcode = "TE5 T12"
                    });
                }

                // Get supplier balance before tests
                var supplierBefore = sage.GetSupplier(testAccount);
                var customerBefore = sage.GetCustomer(testAccount);
                Console.WriteLine($"Before: Customer Balance={customerBefore?.Balance:C}, Supplier Balance={supplierBefore?.Balance:C}");

                // TransType enum: 1=SI, 2=SC, 3=SR, 4=SA, 5=SD, 6=PI, 7=PC, 8=PP, 9=PA, 10=PD
                // Sales types (1-5), Purchase types (6-10)

                Console.WriteLine("\n=== Testing SALES types (1,2) with customer account ===");
                var salesTypes = new[] { (1, "SI", "Sales Invoice"), (2, "SC", "Sales Credit") };
                foreach (var (typeCode, typeAbbr, typeName) in salesTypes)
                {
                    Console.WriteLine($"\n--- Testing TYPE={typeCode} ({typeAbbr} - {typeName}) ---");
                    var success = sage.TestTransactionType(testAccount, typeCode, "4000", 1.00m, 0.20m);
                    Console.WriteLine($"Result: {(success ? "SUCCESS" : "FAILED")}");
                }

                // CORRECTED TYPE VALUES FROM TransType ENUM:
                // TYPE=6 = PI (Purchase Invoice)
                // TYPE=7 = PC (Purchase Credit)
                // (TYPE=4 was SA - Sales Adjustment, which is why it failed!)

                Console.WriteLine("\n=== Testing CORRECT TYPE=6 (PI - Purchase Invoice) ===");

                Console.WriteLine("\n--- Testing TYPE=6 (PI) with 4000 ---");
                var success6 = sage.TestTransactionType(testAccount, 6, "4000", 1.00m, 0.00m);
                Console.WriteLine($"Result: {(success6 ? "SUCCESS" : "FAILED")}");

                // Check supplier balance after TYPE=6
                var supplierMid = sage.GetSupplier(testAccount);
                Console.WriteLine($"Supplier balance after TYPE=6 (PI): {supplierMid?.Balance:C}");

                // Test TYPE=7 (Purchase Credit) too
                Console.WriteLine("\n--- Testing TYPE=7 (PC - Purchase Credit) with 4000 ---");
                var success7 = sage.TestTransactionType(testAccount, 7, "4000", 1.00m, 0.00m);
                Console.WriteLine($"Result: {(success7 ? "SUCCESS" : "FAILED")}");

                // Check supplier balance after TYPE=7
                var supplierMid2 = sage.GetSupplier(testAccount);
                Console.WriteLine($"Supplier balance after TYPE=7 (PC): {supplierMid2?.Balance:C}");

                // Get balances after tests
                var supplierAfter = sage.GetSupplier(testAccount);
                var customerAfter = sage.GetCustomer(testAccount);
                Console.WriteLine($"\nAfter: Customer Balance={customerAfter?.Balance:C}, Supplier Balance={supplierAfter?.Balance:C}");
                Console.WriteLine($"Customer Change: {(customerAfter?.Balance ?? 0) - (customerBefore?.Balance ?? 0):C}");
                Console.WriteLine($"Supplier Change: {(supplierAfter?.Balance ?? 0) - (supplierBefore?.Balance ?? 0):C}");
            }
            else
            {
                Console.WriteLine("\nCommands:");
                Console.WriteLine("  discover - Discover available SDK posting objects");
                Console.WriteLine("");
                Console.WriteLine("  Invoice Posting (updates ledger balances):");
                Console.WriteLine("  sinv [acct] [-auto] - Post sales invoice (creates customer if -auto)");
                Console.WriteLine("  pinv [acct] [-auto] - Post purchase invoice (creates supplier if -auto)");
                Console.WriteLine("");
                Console.WriteLine("  Order Processing (creates documents only, no ledger update):");
                Console.WriteLine("  invdoc   - Create invoice document via InvoicePost");
                Console.WriteLine("  post     - Post via SopPost (Sales Order Processing)");
                Console.WriteLine("  purchase - Post via PopPost (Purchase Order Processing)");
                Console.WriteLine("");
                Console.WriteLine("  Account Management:");
                Console.WriteLine("  newcust  - Create new test customer");
                Console.WriteLine("  newsup   - Create new test supplier");
                Console.WriteLine("  lookup   - Search customers/suppliers");
                Console.WriteLine("");
                Console.WriteLine("  Nominal Ledger:");
                Console.WriteLine("  nominals - List nominal codes");
                Console.WriteLine("  journal  - Post journal entry");
                Console.WriteLine("  bankpay  - Post bank payment");
                Console.WriteLine("  bankrec  - Post bank receipt");
                Console.WriteLine("");
                Console.WriteLine("  Testing:");
                Console.WriteLine("  testtypes - Test TransactionPost TYPE values");
            }
        }
        catch (COMException ex)
        {
            Console.WriteLine($"\nCOM Error: {ex.Message}");
            Console.WriteLine($"Error Code: 0x{ex.ErrorCode:X8}");

            if (ex.ErrorCode == unchecked((int)0x80040154))
            {
                Console.WriteLine("\nThe Sage SDO Engine is not registered.");
                Console.WriteLine("Make sure Sage 50 is installed on this machine.");
            }
        }
        catch (Exception ex)
        {
            Console.WriteLine($"\nError: {ex.Message}");
            Console.WriteLine(ex.StackTrace);
        }
    }
}

/// <summary>
/// Manages Sage 50 SDK detection, download, and loading
/// Detects version from ACCDATA and downloads appropriate SDK from Sage KB
/// </summary>
/// <summary>
/// Registration-free COM support using Windows Activation Contexts
/// Allows loading COM DLLs without registering them in the registry
/// </summary>
static class RegFreeCom
{
    [DllImport("kernel32.dll", SetLastError = true, CharSet = CharSet.Unicode)]
    private static extern IntPtr CreateActCtxW(ref ACTCTX actctx);

    [DllImport("kernel32.dll", SetLastError = true)]
    private static extern bool ActivateActCtx(IntPtr hActCtx, out IntPtr lpCookie);

    [DllImport("kernel32.dll", SetLastError = true)]
    private static extern bool DeactivateActCtx(int dwFlags, IntPtr lpCookie);

    [DllImport("kernel32.dll", SetLastError = true)]
    private static extern void ReleaseActCtx(IntPtr hActCtx);

    [DllImport("kernel32.dll", SetLastError = true, CharSet = CharSet.Unicode)]
    private static extern bool SetDllDirectoryW(string lpPathName);

    [DllImport("kernel32.dll", SetLastError = true, CharSet = CharSet.Unicode)]
    private static extern int AddDllDirectory(string NewDirectory);

    [DllImport("kernel32.dll", SetLastError = true, CharSet = CharSet.Unicode)]
    private static extern IntPtr LoadLibraryW(string lpLibFileName);

    [DllImport("kernel32.dll", SetLastError = true)]
    private static extern bool FreeLibrary(IntPtr hModule);

    [DllImport("kernel32.dll", SetLastError = true, CharSet = CharSet.Ansi)]
    private static extern IntPtr GetProcAddress(IntPtr hModule, string procName);

    // Delegate for DllGetClassObject
    [UnmanagedFunctionPointer(CallingConvention.StdCall)]
    private delegate int DllGetClassObjectDelegate(
        [MarshalAs(UnmanagedType.LPStruct)] Guid rclsid,
        [MarshalAs(UnmanagedType.LPStruct)] Guid riid,
        out IntPtr ppv);

    [DllImport("ole32.dll")]
    private static extern int CoCreateInstance(
        [MarshalAs(UnmanagedType.LPStruct)] Guid rclsid,
        IntPtr pUnkOuter,
        uint dwClsContext,
        [MarshalAs(UnmanagedType.LPStruct)] Guid riid,
        out IntPtr ppv);

    private const uint ACTCTX_FLAG_ASSEMBLY_DIRECTORY_VALID = 0x004;
    private const uint CLSCTX_INPROC_SERVER = 1;
    private const IntPtr INVALID_HANDLE_VALUE = -1;

    private static string? _sdkDirectory;

    [StructLayout(LayoutKind.Sequential, CharSet = CharSet.Unicode)]
    private struct ACTCTX
    {
        public int cbSize;
        public uint dwFlags;
        [MarshalAs(UnmanagedType.LPWStr)]
        public string lpSource;
        public ushort wProcessorArchitecture;
        public ushort wLangId;
        [MarshalAs(UnmanagedType.LPWStr)]
        public string? lpAssemblyDirectory;
        [MarshalAs(UnmanagedType.LPWStr)]
        public string? lpResourceName;
        [MarshalAs(UnmanagedType.LPWStr)]
        public string? lpApplicationName;
        public IntPtr hModule;
    }

    private static IntPtr _actCtx = IntPtr.Zero;
    private static IntPtr _cookie = IntPtr.Zero;
    private static string? _manifestPath;

    // Known CLSIDs for Sage SDO Engine
    private static readonly Guid CLSID_SDOEngine = new("2CCBA24E-16DD-4F45-9DA9-C32AD42DF091");
    private static readonly Guid IID_IDispatch = new("00020400-0000-0000-C000-000000000046");

    /// <summary>
    /// Create a manifest file for the Sage SDK DLL
    /// </summary>
    public static string CreateManifest(string sdkDllPath)
    {
        var dllDir = Path.GetDirectoryName(sdkDllPath)!;
        var manifestPath = Path.Combine(dllDir, "SageSDO.manifest");

        var manifest = $@"<?xml version=""1.0"" encoding=""UTF-8"" standalone=""yes""?>
<assembly xmlns=""urn:schemas-microsoft-com:asm.v1"" manifestVersion=""1.0"">
  <assemblyIdentity type=""win32"" name=""SageSDO"" version=""1.0.0.0"" />
  <file name=""sg50SdoEngine.dll"">
    <comClass clsid=""{{2CCBA24E-16DD-4F45-9DA9-C32AD42DF091}}""
              threadingModel=""Apartment""
              progid=""SDOEngine.32"" />
    <typelib tlbid=""{{8FA9DD05-DBA0-11D1-A59B-00A0C9B18E7F}}""
             version=""1.0""
             helpdir="""" />
  </file>
</assembly>";

        File.WriteAllText(manifestPath, manifest);
        Console.WriteLine($"  Created manifest: {manifestPath}");
        _manifestPath = manifestPath;
        return manifestPath;
    }

    /// <summary>
    /// Activate the COM context for the SDK DLL
    /// </summary>
    public static bool Activate(string manifestPath)
    {
        if (_actCtx != IntPtr.Zero)
        {
            Console.WriteLine("  Activation context already active");
            return true;
        }

        var assemblyDir = Path.GetDirectoryName(manifestPath);
        Console.WriteLine($"  Assembly directory: {assemblyDir}");

        // Add SDK directory to DLL search path so dependencies can be found
        if (!string.IsNullOrEmpty(assemblyDir))
        {
            _sdkDirectory = assemblyDir;
            if (SetDllDirectoryW(assemblyDir))
            {
                Console.WriteLine($"  Added to DLL search path: {assemblyDir}");
            }
            else
            {
                Console.WriteLine($"  SetDllDirectory failed: {Marshal.GetLastWin32Error()}");
            }
        }

        var actctx = new ACTCTX
        {
            cbSize = Marshal.SizeOf<ACTCTX>(),
            dwFlags = ACTCTX_FLAG_ASSEMBLY_DIRECTORY_VALID,
            lpSource = manifestPath,
            lpAssemblyDirectory = assemblyDir,
            lpResourceName = null,
            lpApplicationName = null,
            hModule = IntPtr.Zero,
            wProcessorArchitecture = 0,
            wLangId = 0
        };

        Console.WriteLine($"  ACTCTX.cbSize: {actctx.cbSize}");
        Console.WriteLine($"  ACTCTX.lpSource: {actctx.lpSource}");

        _actCtx = CreateActCtxW(ref actctx);
        if (_actCtx == IntPtr.Zero || _actCtx == new IntPtr(-1))
        {
            var error = Marshal.GetLastWin32Error();
            Console.WriteLine($"  CreateActCtxW failed with error: {error}");
            Console.WriteLine($"  Error description: {new System.ComponentModel.Win32Exception(error).Message}");
            _actCtx = IntPtr.Zero;
            return false;
        }

        Console.WriteLine($"  CreateActCtxW succeeded, handle: {_actCtx}");

        if (!ActivateActCtx(_actCtx, out _cookie))
        {
            var error = Marshal.GetLastWin32Error();
            Console.WriteLine($"  ActivateActCtx failed with error: {error}");
            ReleaseActCtx(_actCtx);
            _actCtx = IntPtr.Zero;
            return false;
        }

        Console.WriteLine("  Activation context activated successfully");
        return true;
    }

    /// <summary>
    /// Create SDOEngine instance using the activation context
    /// </summary>
    public static object? CreateSDOEngine()
    {
        // First try to load the DLL directly to diagnose dependency issues
        if (!string.IsNullOrEmpty(_sdkDirectory))
        {
            var dllPath = Path.Combine(_sdkDirectory, "sg50SdoEngine.dll");
            Console.WriteLine($"  Attempting to load DLL: {dllPath}");

            var hModule = LoadLibraryW(dllPath);
            if (hModule == IntPtr.Zero)
            {
                var error = Marshal.GetLastWin32Error();
                Console.WriteLine($"  LoadLibrary failed with error: {error}");
                Console.WriteLine($"  Error description: {new System.ComponentModel.Win32Exception(error).Message}");
            }
            else
            {
                Console.WriteLine($"  LoadLibrary succeeded, handle: {hModule}");

                // Try calling DllGetClassObject directly
                var procAddr = GetProcAddress(hModule, "DllGetClassObject");
                if (procAddr != IntPtr.Zero)
                {
                    Console.WriteLine("  Found DllGetClassObject export");

                    var dllGetClassObject = Marshal.GetDelegateForFunctionPointer<DllGetClassObjectDelegate>(procAddr);

                    // IClassFactory IID
                    var IID_IClassFactory = new Guid("00000001-0000-0000-C000-000000000046");

                    var hr = dllGetClassObject(CLSID_SDOEngine, IID_IClassFactory, out var pFactory);
                    Console.WriteLine($"  DllGetClassObject returned: 0x{hr:X8}");

                    if (hr == 0 && pFactory != IntPtr.Zero)
                    {
                        Console.WriteLine("  Got class factory, creating instance...");

                        // Get IClassFactory interface and call CreateInstance
                        var factory = (IClassFactory)Marshal.GetObjectForIUnknown(pFactory);
                        var iidDispatch = IID_IDispatch; // Copy to local for ref parameter
                        var instance = factory.CreateInstance(null, ref iidDispatch);
                        Marshal.Release(pFactory);

                        if (instance != null)
                        {
                            Console.WriteLine("  Created SDOEngine via DllGetClassObject");
                            return instance;
                        }
                    }
                }
                else
                {
                    Console.WriteLine($"  DllGetClassObject not found, error: {Marshal.GetLastWin32Error()}");
                }
            }
        }

        // Try with activation context and CoCreateInstance
        if (_actCtx != IntPtr.Zero)
        {
            try
            {
                var hr = CoCreateInstance(CLSID_SDOEngine, IntPtr.Zero, CLSCTX_INPROC_SERVER,
                    IID_IDispatch, out var pUnk);

                if (hr == 0 && pUnk != IntPtr.Zero)
                {
                    var obj = Marshal.GetObjectForIUnknown(pUnk);
                    Marshal.Release(pUnk);
                    Console.WriteLine("  Created SDOEngine via activation context");
                    return obj;
                }
                Console.WriteLine($"  CoCreateInstance returned: 0x{hr:X8}");
            }
            catch (Exception ex)
            {
                Console.WriteLine($"  CoCreateInstance failed: {ex.Message}");
            }
        }

        // Fallback to Type.GetTypeFromCLSID
        try
        {
            var type = Type.GetTypeFromCLSID(CLSID_SDOEngine);
            if (type != null)
            {
                var obj = Activator.CreateInstance(type);
                if (obj != null)
                {
                    Console.WriteLine("  Created SDOEngine via CLSID");
                    return obj;
                }
            }
        }
        catch (Exception ex)
        {
            Console.WriteLine($"  CLSID activation failed: {ex.Message}");
        }

        return null;
    }

    // COM IClassFactory interface
    [ComImport]
    [Guid("00000001-0000-0000-C000-000000000046")]
    [InterfaceType(ComInterfaceType.InterfaceIsIUnknown)]
    private interface IClassFactory
    {
        [return: MarshalAs(UnmanagedType.Interface)]
        object CreateInstance([MarshalAs(UnmanagedType.IUnknown)] object? pUnkOuter,
            ref Guid riid);
        void LockServer(bool fLock);
    }

    /// <summary>
    /// Deactivate and release the context
    /// </summary>
    public static void Deactivate()
    {
        if (_cookie != IntPtr.Zero)
        {
            DeactivateActCtx(0, _cookie);
            _cookie = IntPtr.Zero;
        }

        if (_actCtx != IntPtr.Zero)
        {
            ReleaseActCtx(_actCtx);
            _actCtx = IntPtr.Zero;
        }

        // Reset DLL search path
        if (_sdkDirectory != null)
        {
            SetDllDirectoryW("");
            _sdkDirectory = null;
        }
    }

    /// <summary>
    /// Check if activation context is active
    /// </summary>
    public static bool IsActive => _actCtx != IntPtr.Zero;

    /// <summary>
    /// Get the manifest path
    /// </summary>
    public static string? ManifestPath => _manifestPath;
}

static class SdkManager
{
    private static readonly string SdkCacheDir = Path.Combine(
        Environment.GetFolderPath(Environment.SpecialFolder.LocalApplicationData),
        "SageConnector", "SDK");

    /// <summary>
    /// Path to extracted SDK DLL (set after successful extraction)
    /// </summary>
    public static string? ExtractedSdkPath { get; private set; }

    /// <summary>
    /// SDK download URLs from Sage KB article 201224120012523
    /// </summary>
    private static readonly Dictionary<string, (string Url64, string Url32)> SdkDownloads = new()
    {
        ["33.0"] = ("https://downloads.sage.co.uk/download/?did=3e47ac24-0cff-402d-8232-7a3296a59a01",
                    "https://downloads.sage.co.uk/download/?did=986d817d-494e-4e7b-9d74-9950f2b576e7"),
        ["32.1"] = ("https://downloads.sage.co.uk/download/?did=c298b7b7-c5fa-479c-b90d-9c6f61b77a97",
                    "https://downloads.sage.co.uk/download/?did=e9f67243-1ec0-4dc1-9cb7-ec2fcae20f06"),
        ["32.0"] = ("https://downloads.sage.co.uk/download/?did=c038d725-22a5-411b-8247-5ee6e1156dba",
                    "https://downloads.sage.co.uk/download/?did=86a3b851-5eca-4844-8d7f-551a2b19844c"),
        ["31.1"] = ("https://downloads.sage.co.uk/download/?did=5b98d38c-1497-4aef-baf0-69bf44e3daf6",
                    "https://downloads.sage.co.uk/download/?did=85c44ae1-8912-448e-b703-583130e0b352"),
        ["30.1"] = ("https://downloads.sage.co.uk/download/?did=2025c04b-f6a2-4b3b-9a0f-1db020062f90",
                    "https://downloads.sage.co.uk/download/?did=49d800c3-ff2f-4b63-8ac4-fe354a80b2eb"),
        ["29.0"] = ("https://downloads.sage.co.uk/download/?did=4ad4deae-b847-458d-9081-e53fde039d3d", ""),
        ["28.0"] = ("https://downloads.sage.co.uk/download/?did=7e19e237-5bd3-481f-9e2a-f73b2c7e6a2c", ""),
        ["27.0"] = ("https://downloads.sage.co.uk/download/?did=7a55278e-2294-4349-834f-9141f68686a1", ""),
        ["26.0"] = ("https://downloads.sage.co.uk/download/?did=f71b5929-1a8f-4c65-82a6-b15a7b01cb7a", ""),
    };

    // Known GUID fingerprints from ACCSTAT.DTA mapped to Sage versions
    // These GUIDs appear to be schema identifiers that change per major version
    private static readonly Dictionary<string, string> KnownSchemaGuids = new()
    {
        // Add known GUIDs here as they're discovered
        // Format: GUID -> Version
        ["5D3EB135-3317-413B-99DE-47C6B044134D"] = "32.0", // Observed in v32 data
    };

    // TABLEMETADATA.DTA header format version mapped to Sage versions
    // First 4 bytes contain a format version number
    private static readonly Dictionary<uint, string> KnownTableMetaVersions = new()
    {
        [0x15] = "32.0", // Format version 21 observed in v32 data
        // Add more mappings as discovered from different Sage versions
    };

    /// <summary>
    /// Detect Sage 50 version from ACCDATA folder by analyzing DTA file structure
    /// </summary>
    public static string? DetectVersion(string? accDataPath = null)
    {
        if (string.IsNullOrEmpty(accDataPath) || !Directory.Exists(accDataPath))
        {
            if (!string.IsNullOrEmpty(accDataPath))
                Console.WriteLine($"  ACCDATA path not found: {accDataPath}");
            return null;
        }

        // Method 1: Extract GUID from ACCSTAT.DTA and lookup known versions
        var accstatPath = Path.Combine(accDataPath, "ACCSTAT.DTA");
        if (File.Exists(accstatPath))
        {
            try
            {
                using var fs = new FileStream(accstatPath, FileMode.Open, FileAccess.Read, FileShare.ReadWrite);
                var buffer = new byte[256];
                fs.Read(buffer, 0, buffer.Length);
                var content = System.Text.Encoding.ASCII.GetString(buffer);

                // Extract GUID pattern
                var guidMatch = System.Text.RegularExpressions.Regex.Match(content,
                    @"([0-9A-Fa-f]{8}-[0-9A-Fa-f]{4}-[0-9A-Fa-f]{4}-[0-9A-Fa-f]{4}-[0-9A-Fa-f]{12})");

                if (guidMatch.Success)
                {
                    var guid = guidMatch.Groups[1].Value.ToUpperInvariant();
                    Console.WriteLine($"  ACCSTAT.DTA schema GUID: {guid}");

                    if (KnownSchemaGuids.TryGetValue(guid, out var version))
                    {
                        Console.WriteLine($"  Detected version from GUID fingerprint: {version}");
                        return version;
                    }
                    else
                    {
                        Console.WriteLine($"  Unknown GUID - add to KnownSchemaGuids mapping");
                    }
                }
            }
            catch (Exception ex)
            {
                Console.WriteLine($"  Could not read ACCSTAT.DTA: {ex.Message}");
            }
        }

        // Method 2: Check TABLEMETADATA.DTA header format version
        var tableMetaPath = Path.Combine(accDataPath, "TABLEMETADATA.DTA");
        if (File.Exists(tableMetaPath))
        {
            try
            {
                using var fs = new FileStream(tableMetaPath, FileMode.Open, FileAccess.Read, FileShare.ReadWrite);
                var header = new byte[4];
                fs.Read(header, 0, 4);
                var formatVersion = BitConverter.ToUInt32(header, 0);
                Console.WriteLine($"  TABLEMETADATA.DTA format version: {formatVersion} (0x{formatVersion:X2})");

                if (KnownTableMetaVersions.TryGetValue(formatVersion, out var version))
                {
                    Console.WriteLine($"  Detected version from TABLEMETADATA format: {version}");
                    return version;
                }
            }
            catch (Exception ex)
            {
                Console.WriteLine($"  Could not read TABLEMETADATA.DTA: {ex.Message}");
            }
        }

        // Method 3: Check ACCDATA.INI for Sage program path
        var accdataIni = Path.Combine(accDataPath, "ACCDATA.INI");
        if (File.Exists(accdataIni))
        {
            try
            {
                var iniContent = File.ReadAllText(accdataIni);
                // Look for version in program path like "SAGE\ACCOUNTS\2024" or similar
                var pathMatch = System.Text.RegularExpressions.Regex.Match(iniContent, @"SAGE\\ACCOUNTS\\(\d{4})",
                    System.Text.RegularExpressions.RegexOptions.IgnoreCase);
                if (pathMatch.Success && int.TryParse(pathMatch.Groups[1].Value, out int year))
                {
                    int ver = year - 1992; // 2024 -> 32, 2025 -> 33
                    if (ver >= 26 && ver <= 35)
                    {
                        Console.WriteLine($"  Detected version from ACCDATA.INI path (year {year}): {ver}.0");
                        return $"{ver}.0";
                    }
                }
            }
            catch { }
        }

        // Method 3: Check HEADER.DTA file signature (less common)
        var headerDta = Path.Combine(accDataPath, "HEADER.DTA");
        if (File.Exists(headerDta))
        {
            try
            {
                using var fs = new FileStream(headerDta, FileMode.Open, FileAccess.Read, FileShare.ReadWrite);
                var header = new byte[256];
                fs.Read(header, 0, Math.Min(256, (int)fs.Length));

                // Look for version patterns in header
                // Sage stores version info typically in first 100 bytes
                var headerStr = System.Text.Encoding.ASCII.GetString(header);

                // Check for version markers like "V32", "V33", etc.
                for (int v = 33; v >= 26; v--)
                {
                    if (headerStr.Contains($"V{v}") || headerStr.Contains($"v{v}"))
                    {
                        Console.WriteLine($"  Detected version from HEADER.DTA: {v}.0");
                        return $"{v}.0";
                    }
                }
            }
            catch (Exception ex)
            {
                Console.WriteLine($"  Could not read HEADER.DTA: {ex.Message}");
            }
        }

        // Method 2: Check SETTINGS.DTA for version info
        var settingsDta = Path.Combine(accDataPath, "SETTINGS.DTA");
        if (File.Exists(settingsDta))
        {
            try
            {
                using var fs = new FileStream(settingsDta, FileMode.Open, FileAccess.Read, FileShare.ReadWrite);
                var buffer = new byte[512];
                fs.Read(buffer, 0, Math.Min(512, (int)fs.Length));
                var content = System.Text.Encoding.ASCII.GetString(buffer);

                for (int v = 33; v >= 26; v--)
                {
                    if (content.Contains($"{v}.") || content.Contains($"V{v}"))
                    {
                        Console.WriteLine($"  Detected version from SETTINGS.DTA: {v}.0");
                        return $"{v}.0";
                    }
                }
            }
            catch { }
        }

        // Method 3: Check for ACCDATA version file
        var versionFile = Path.Combine(accDataPath, "VERSION");
        if (File.Exists(versionFile))
        {
            try
            {
                var versionText = File.ReadAllText(versionFile).Trim();
                Console.WriteLine($"  Detected version from VERSION file: {versionText}");
                return versionText;
            }
            catch { }
        }

        // Method 4: Infer from parent folder structure (e.g., Accounts/2024)
        var parentPath = Path.GetDirectoryName(accDataPath);
        if (parentPath != null)
        {
            var yearMatch = System.Text.RegularExpressions.Regex.Match(parentPath, @"(\d{4})");
            if (yearMatch.Success && int.TryParse(yearMatch.Groups[1].Value, out int year))
            {
                // Map year to version: 2024=v32, 2025=v33, etc.
                int version = year - 1992; // Rough mapping
                if (version >= 26 && version <= 33)
                {
                    Console.WriteLine($"  Inferred version from path year {year}: {version}.0");
                    return $"{version}.0";
                }
            }
        }

        Console.WriteLine("  Could not detect Sage version from ACCDATA");
        return null;
    }

    /// <summary>
    /// Check if SDK is already available (either registered or cached)
    /// </summary>
    public static bool IsSdkAvailable()
    {
        var sdoType = Type.GetTypeFromProgID("SDOEngine.32");
        return sdoType != null;
    }

    /// <summary>
    /// Get the download URL for a specific version
    /// Falls back to closest available version for the same major version
    /// </summary>
    public static string? GetDownloadUrl(string version, bool prefer64Bit = true)
    {
        // Normalize version (e.g., "32" -> "32.0")
        if (!version.Contains('.'))
            version = $"{version}.0";

        // Try exact match first
        if (SdkDownloads.TryGetValue(version, out var urls))
        {
            var url = prefer64Bit && !string.IsNullOrEmpty(urls.Url64) ? urls.Url64 : urls.Url32;
            if (!string.IsNullOrEmpty(url))
                return url;
        }

        // Try to find any version with same major number (e.g., 31.0 -> 31.1)
        var majorVersion = version.Split('.')[0];
        var matchingVersion = SdkDownloads.Keys
            .Where(k => k.StartsWith(majorVersion + "."))
            .OrderByDescending(k => k) // Get highest minor version
            .FirstOrDefault();

        if (matchingVersion != null && SdkDownloads.TryGetValue(matchingVersion, out urls))
        {
            Console.WriteLine($"  Using SDK v{matchingVersion} (closest to v{version})");
            return prefer64Bit && !string.IsNullOrEmpty(urls.Url64) ? urls.Url64 : urls.Url32;
        }

        return null;
    }

    /// <summary>
    /// Download SDK zip to cache directory
    /// </summary>
    public static async Task<string?> DownloadSdkAsync(string version, bool prefer64Bit = true)
    {
        var url = GetDownloadUrl(version, prefer64Bit);
        if (url == null)
        {
            Console.WriteLine($"  No SDK download available for version {version}");
            return null;
        }

        Directory.CreateDirectory(SdkCacheDir);
        var bitness = prefer64Bit ? "x64" : "x86";
        var fileName = $"SageSDO_v{version}_{bitness}.zip";
        var filePath = Path.Combine(SdkCacheDir, fileName);

        if (File.Exists(filePath))
        {
            Console.WriteLine($"  SDK already downloaded: {filePath}");
            return filePath;
        }

        Console.WriteLine($"  Downloading SDK v{version} ({bitness})...");
        Console.WriteLine($"  URL: {url}");

        try
        {
            using var http = new HttpClient();
            http.Timeout = TimeSpan.FromMinutes(10);

            var response = await http.GetAsync(url, HttpCompletionOption.ResponseHeadersRead);
            response.EnsureSuccessStatusCode();

            var totalBytes = response.Content.Headers.ContentLength ?? -1;
            using var contentStream = await response.Content.ReadAsStreamAsync();
            using var fileStream = new FileStream(filePath, FileMode.Create, FileAccess.Write, FileShare.None);

            var buffer = new byte[81920];
            long totalRead = 0;
            int bytesRead;

            while ((bytesRead = await contentStream.ReadAsync(buffer, 0, buffer.Length)) > 0)
            {
                await fileStream.WriteAsync(buffer, 0, bytesRead);
                totalRead += bytesRead;

                if (totalBytes > 0)
                {
                    var progress = (int)((totalRead * 100) / totalBytes);
                    Console.Write($"\r  Downloaded: {totalRead / 1024 / 1024}MB / {totalBytes / 1024 / 1024}MB ({progress}%)");
                }
            }
            Console.WriteLine();

            Console.WriteLine($"  SDK downloaded: {filePath}");
            return filePath;
        }
        catch (Exception ex)
        {
            Console.WriteLine($"  Download failed: {ex.Message}");
            if (File.Exists(filePath))
                File.Delete(filePath);
            return null;
        }
    }

    /// <summary>
    /// Extract SDK DLLs from zip file and MSI
    /// </summary>
    public static bool ExtractSdk(string zipPath, out string? sdkPath)
    {
        sdkPath = null;
        if (!File.Exists(zipPath))
        {
            Console.WriteLine($"  ZIP not found: {zipPath}");
            return false;
        }

        var extractDir = Path.Combine(Path.GetDirectoryName(zipPath)!,
            Path.GetFileNameWithoutExtension(zipPath));

        // Clean existing extraction
        if (Directory.Exists(extractDir))
        {
            Console.WriteLine($"  Cleaning existing extraction: {extractDir}");
            Directory.Delete(extractDir, true);
        }

        Console.WriteLine($"  Extracting ZIP to: {extractDir}");

        try
        {
            ZipFile.ExtractToDirectory(zipPath, extractDir);
            Console.WriteLine("  ZIP extraction successful");

            // Find the MSI inside the ZIP
            var msiFile = Directory.GetFiles(extractDir, "*.msi", SearchOption.AllDirectories).FirstOrDefault();
            if (msiFile != null)
            {
                Console.WriteLine($"  Found MSI: {Path.GetFileName(msiFile)}");

                // Extract the MSI to get all dependencies
                var msiExtractDir = Path.Combine(extractDir, "msi_contents");
                Console.WriteLine($"  Extracting MSI to: {msiExtractDir}");

                var psi = new ProcessStartInfo
                {
                    FileName = "msiexec",
                    Arguments = $"/a \"{msiFile}\" /qn TARGETDIR=\"{msiExtractDir}\"",
                    UseShellExecute = false,
                    CreateNoWindow = true
                };

                using var process = Process.Start(psi);
                process?.WaitForExit(TimeSpan.FromMinutes(2));

                if (process?.ExitCode == 0)
                {
                    Console.WriteLine("  MSI extraction successful");

                    // Find the Sage SDK directory (contains all DLLs)
                    var sageDlls = Directory.GetFiles(msiExtractDir, "sg50SdoEngine.dll", SearchOption.AllDirectories);
                    if (sageDlls.Length > 0)
                    {
                        sdkPath = Path.GetDirectoryName(sageDlls[0]);
                        ExtractedSdkPath = sdkPath;

                        // Count DLLs for verification
                        var dllCount = Directory.GetFiles(sdkPath!, "*.dll").Length;
                        Console.WriteLine($"  SDK DLLs found at: {sdkPath} ({dllCount} DLLs)");

                        // List key DLLs
                        var keyDlls = new[] { "sg50SdoEngine.dll", "sg50BusinessObjectsV32.dll", "sg50DataObjectsV32.dll", "Sage.UtilsLibV32.dll" };
                        foreach (var dll in keyDlls)
                        {
                            var exists = File.Exists(Path.Combine(sdkPath!, dll));
                            Console.WriteLine($"    {dll}: {(exists ? "OK" : "MISSING")}");
                        }

                        // Create manifest for registration-free COM
                        RegFreeCom.CreateManifest(sageDlls[0]);
                        Console.WriteLine($"  SDK ready for registration-free COM");
                        return true;
                    }
                }
                else
                {
                    Console.WriteLine($"  MSI extraction failed with exit code: {process?.ExitCode}");
                }
            }

            // Fallback: Look directly in ZIP extraction
            var files = Directory.GetFiles(extractDir, "*", SearchOption.AllDirectories);
            Console.WriteLine($"  Extracted {files.Length} files from ZIP");

            var sdoDll = files.FirstOrDefault(f => Path.GetFileName(f).Equals("sg50SdoEngine.dll", StringComparison.OrdinalIgnoreCase));
            if (sdoDll != null)
            {
                sdkPath = Path.GetDirectoryName(sdoDll);
                ExtractedSdkPath = sdkPath;
                Console.WriteLine($"  SDK DLLs found at: {sdkPath}");
                RegFreeCom.CreateManifest(sdoDll);
                return true;
            }

            // Try to find type library
            var tlbFile = files.FirstOrDefault(f => f.EndsWith(".tlb", StringComparison.OrdinalIgnoreCase));
            if (tlbFile != null)
            {
                sdkPath = Path.GetDirectoryName(tlbFile);
                Console.WriteLine($"  SDK type library found at: {sdkPath}");
                return true;
            }

            // Just return the extract dir if we got files
            if (files.Length > 0)
            {
                sdkPath = extractDir;
                Console.WriteLine("  SDK DLLs not found but extraction succeeded");
                return true;
            }

            Console.WriteLine("  Extraction succeeded but no files found");
            return false;
        }
        catch (Exception ex)
        {
            Console.WriteLine($"  ZIP extraction failed: {ex.Message}");
            return false;
        }
    }

    /// <summary>
    /// Install/register SDK - extracts zip and registers DLLs
    /// </summary>
    public static bool InstallSdk(string zipPath)
    {
        if (!ExtractSdk(zipPath, out var sdkPath) || sdkPath == null)
        {
            Console.WriteLine("  SDK extraction failed");
            return false;
        }

        // Try to register the COM DLLs
        return RegisterSdkDlls(sdkPath);
    }

    private static bool RegisterSdkDlls(string sdkPath)
    {
        Console.WriteLine($"  Looking for SDK DLLs in: {sdkPath}");

        // Find sg50SdoEngine.dll recursively
        var sdoDll = Directory.GetFiles(sdkPath, "sg50SdoEngine.dll", SearchOption.AllDirectories)
            .FirstOrDefault();

        if (sdoDll == null)
        {
            Console.WriteLine("  sg50SdoEngine.dll not found");
            Console.WriteLine("  Available DLLs:");
            foreach (var dll in Directory.GetFiles(sdkPath, "*.dll", SearchOption.AllDirectories).Take(10))
            {
                Console.WriteLine($"    {Path.GetFileName(dll)}");
            }
            return false;
        }

        Console.WriteLine($"  Found: {sdoDll}");
        Console.WriteLine($"  Registering COM DLL (requires admin)...");

        try
        {
            var psi = new ProcessStartInfo
            {
                FileName = "regsvr32",
                Arguments = $"/s \"{sdoDll}\"",
                UseShellExecute = true,
                Verb = "runas"
            };

            using var process = Process.Start(psi);
            process?.WaitForExit(TimeSpan.FromSeconds(30));

            if (process?.ExitCode == 0)
            {
                Console.WriteLine("  COM DLL registered successfully");
                return true;
            }
            else
            {
                Console.WriteLine($"  regsvr32 exited with code: {process?.ExitCode}");
            }
        }
        catch (Exception ex)
        {
            Console.WriteLine($"  Registration failed: {ex.Message}");
        }

        return false;
    }

    /// <summary>
    /// Ensure SDK is available for the detected version
    /// Downloads and extracts if necessary (no registration required)
    /// </summary>
    public static async Task<bool> EnsureSdkAsync(string accDataPath, bool force = false)
    {
        Console.WriteLine("\n[SDK Manager]");

        // Check if SDK is already registered (can skip download)
        if (!force && IsSdkAvailable())
        {
            Console.WriteLine("  SDK already available (SDOEngine.32 registered)");
            return true;
        }

        // Check if we already have extracted SDK
        if (!force && !string.IsNullOrEmpty(ExtractedSdkPath) && Directory.Exists(ExtractedSdkPath))
        {
            var dllPath = Path.Combine(ExtractedSdkPath, "sg50SdoEngine.dll");
            if (File.Exists(dllPath))
            {
                Console.WriteLine($"  SDK already extracted: {ExtractedSdkPath}");
                return true;
            }
        }

        if (force)
        {
            Console.WriteLine("  Force mode: re-downloading SDK");
        }

        // Detect version from ACCDATA
        var version = DetectVersion(accDataPath);
        if (version == null)
        {
            Console.WriteLine("  Cannot determine required SDK version");
            Console.WriteLine("  Please install Sage 50 SDO manually from:");
            Console.WriteLine("  https://gb-kb.sage.com/portal/app/portlets/results/viewsolution.jsp?solutionid=201224120012523");
            return false;
        }

        Console.WriteLine($"  Required SDK version: {version}");

        // Download SDK
        var is64Bit = Environment.Is64BitOperatingSystem;
        var zipPath = await DownloadSdkAsync(version, is64Bit);
        if (zipPath == null)
            return false;

        // Extract SDK (no registration needed - we use reg-free COM)
        if (!ExtractSdk(zipPath, out var sdkPath) || sdkPath == null)
        {
            Console.WriteLine("  SDK extraction failed");
            return false;
        }

        // Verify extraction
        var sdoDllPath = Path.Combine(sdkPath, "sg50SdoEngine.dll");
        if (!File.Exists(sdoDllPath))
        {
            Console.WriteLine($"  SDK DLL not found: {sdoDllPath}");
            return false;
        }

        Console.WriteLine("  SDK ready for registration-free COM");
        return true;
    }

    /// <summary>
    /// List available SDK versions
    /// </summary>
    public static void ListAvailableVersions()
    {
        Console.WriteLine("\nAvailable Sage 50 SDK versions:");
        foreach (var kvp in SdkDownloads.OrderByDescending(k => k.Key))
        {
            var has64 = !string.IsNullOrEmpty(kvp.Value.Url64) ? "64-bit" : "";
            var has32 = !string.IsNullOrEmpty(kvp.Value.Url32) ? "32-bit" : "";
            var bits = string.Join(", ", new[] { has64, has32 }.Where(s => !string.IsNullOrEmpty(s)));
            Console.WriteLine($"  v{kvp.Key} ({bits})");
        }
    }
}

/// <summary>
/// Wrapper class for Sage 50 SDO COM connection
/// </summary>
class SageConnection : IDisposable
{
    private dynamic? _sdoEngine;
    private dynamic? _workspaces;
    private dynamic? _workspace;
    private bool _disposed;
    private bool _usingRegFreeCom;

    /// <summary>
    /// Create SageConnection
    /// </summary>
    /// <param name="forceRegFreeCom">Force using registration-free COM even if registered COM exists</param>
    public SageConnection(bool forceRegFreeCom = false)
    {
        Type? sdoEngineType = null;
        string? usedProgId = null;

        // Try registered COM first (unless forced to use reg-free)
        if (!forceRegFreeCom)
        {
            string[] progIds = new[]
            {
                "SDOEngine.32",
                "SDOEngine",
                "SageDataObject50.SDOEngine",
                "Sage.Accounts.SDOEngine"
            };

            foreach (var progId in progIds)
            {
                sdoEngineType = Type.GetTypeFromProgID(progId);
                if (sdoEngineType != null)
                {
                    usedProgId = progId;
                    Console.WriteLine($"Found registered COM object: {progId}");
                    break;
                }
            }
        }
        else
        {
            Console.WriteLine("Forcing registration-free COM mode...");
        }

        // If no registered COM (or forced), try registration-free COM with extracted SDK
        if (sdoEngineType == null)
        {
            if (!forceRegFreeCom)
                Console.WriteLine("No registered SDK found, trying registration-free COM...");

            if (!string.IsNullOrEmpty(SdkManager.ExtractedSdkPath))
            {
                var manifestPath = RegFreeCom.ManifestPath
                    ?? Path.Combine(SdkManager.ExtractedSdkPath, "SageSDO.manifest");

                Console.WriteLine($"  Manifest path: {manifestPath}");
                Console.WriteLine($"  Manifest exists: {File.Exists(manifestPath)}");

                if (File.Exists(manifestPath))
                {
                    if (RegFreeCom.Activate(manifestPath))
                    {
                        _sdoEngine = RegFreeCom.CreateSDOEngine();
                        if (_sdoEngine != null)
                        {
                            _usingRegFreeCom = true;
                            Console.WriteLine("SDO Engine created via registration-free COM.");
                        }
                    }
                }
                else
                {
                    Console.WriteLine("  Manifest not found - run 'sdk install' first");
                }
            }
            else
            {
                Console.WriteLine("  No extracted SDK path set - run 'sdk install' first");
            }

            if (_sdoEngine == null)
            {
                throw new InvalidOperationException(
                    "No Sage SDO Engine found. Run 'sdk install <accdata_path>' first to download and extract the SDK.");
            }
        }
        else
        {
            _sdoEngine = Activator.CreateInstance(sdoEngineType);
            if (_sdoEngine == null)
            {
                throw new InvalidOperationException(
                    "Failed to create SDOEngine instance.");
            }
            Console.WriteLine($"SDO Engine created successfully (using {usedProgId}).");
        }

        // List available properties/methods
        try
        {
            Console.WriteLine("\nExploring SDO Engine members...");
            var type = _sdoEngine!.GetType();

            // Try to get type info via IDispatch
            foreach (var member in type.GetMembers())
            {
                if (member.Name.StartsWith("get_") || member.Name.StartsWith("set_"))
                    continue;
                if (member.MemberType == System.Reflection.MemberTypes.Method ||
                    member.MemberType == System.Reflection.MemberTypes.Property)
                {
                    Console.WriteLine($"  {member.MemberType}: {member.Name}");
                }
            }
        }
        catch { }
    }

    /// <summary>
    /// List companies found in common Sage data locations
    /// </summary>
    public void ListCompanies()
    {
        var commonPaths = new[]
        {
            @"C:\ProgramData\Sage\Accounts\2025",
            @"C:\ProgramData\Sage\Accounts\2024",
            @"C:\ProgramData\Sage\Accounts\2023",
            @"C:\SAGELINE50\Company.000\ACCDATA",
            @"C:\Documents and Settings\All Users\Application Data\Sage\Accounts"
        };

        Console.WriteLine("\nSearching common Sage data locations:");

        foreach (var basePath in commonPaths)
        {
            if (Directory.Exists(basePath))
            {
                Console.WriteLine($"\n  Found: {basePath}");

                // Look for Company.XXX folders
                try
                {
                    var companyFolders = Directory.GetDirectories(basePath, "Company.*");
                    foreach (var folder in companyFolders)
                    {
                        var accDataPath = Path.Combine(folder, "ACCDATA");
                        if (Directory.Exists(accDataPath))
                        {
                            Console.WriteLine($"    -> {accDataPath}");
                        }
                    }

                    // Also check for ACCDATA directly
                    var directAccData = Path.Combine(basePath, "ACCDATA");
                    if (Directory.Exists(directAccData))
                    {
                        Console.WriteLine($"    -> {directAccData}");
                    }
                }
                catch (Exception ex)
                {
                    Console.WriteLine($"    Error scanning: {ex.Message}");
                }
            }
        }
    }

    /// <summary>
    /// Connect to a Sage 50 company
    /// </summary>
    public void Connect(string dataPath, string username = "manager", string password = "")
    {
        if (_sdoEngine == null)
            throw new InvalidOperationException("SDO Engine not initialized");

        Console.WriteLine($"  Data path: {dataPath}");
        Console.WriteLine($"  Username: {username}");
        Console.WriteLine($"  Password: {(string.IsNullOrEmpty(password) ? "(empty)" : "****")}");

        // Get workspaces collection using InvokeMember for better COM interop
        var engineType = ((object)_sdoEngine).GetType();

        Console.WriteLine("\n  Getting Workspaces via InvokeMember...");
        _workspaces = engineType.InvokeMember(
            "Workspaces",
            BindingFlags.GetProperty,
            null,
            _sdoEngine,
            null);
        Console.WriteLine("  Got Workspaces collection");

        // Add a new workspace
        var wsType = ((object)_workspaces).GetType();
        Console.WriteLine("\n  Adding workspace...");
        _workspace = wsType.InvokeMember(
            "Add",
            BindingFlags.InvokeMethod,
            null,
            _workspaces,
            new object[] { "SageConnector" });
        Console.WriteLine("  Created workspace");

        // List workspace methods using IDispatch type info
        ListComMethods("Workspace", _workspace);

        // Try multiple connection approaches using InvokeMember for better COM interop
        var workspaceType = ((object)_workspace).GetType();

        // Connection attempts - the 4-param version is correct for Sage 50 v32
        // Parameters: DataPath, UserName, Password, UniqueID
        // Use timestamp+random to ensure unique ID each time
        var uniqueId = $"SageConn_{DateTime.Now.Ticks}_{new Random().Next(1000, 9999)}";

        var connectionAttempts = new (string Name, object[] Args)[]
        {
            ("Standard 4 params", new object[] { dataPath, username, password, uniqueId }),
        };

        bool connected = false;
        foreach (var (name, args) in connectionAttempts)
        {
            try
            {
                Console.WriteLine($"  Trying: {name}...");
                var result = workspaceType.InvokeMember(
                    "Connect",
                    BindingFlags.InvokeMethod,
                    null,
                    _workspace,
                    args);
                Console.WriteLine($"    Result: {result}");
                if (result != null && !result.Equals(false) && !result.Equals(0))
                {
                    connected = true;
                    Console.WriteLine("Connected to Sage 50 successfully!");
                    break;
                }
            }
            catch (TargetInvocationException tie)
            {
                var msg = tie.InnerException?.Message ?? tie.Message;
                Console.WriteLine($"    Failed: {msg}");

                // Check for busy error - connection params are correct but Sage is in use
                if (msg.Contains("busy", StringComparison.OrdinalIgnoreCase))
                {
                    Console.WriteLine("\n*** Sage 50 is currently in use! ***");
                    Console.WriteLine("    Please close Sage 50 on all machines and try again.");
                    Console.WriteLine("    Or wait for other users to finish their session.");

                    // Check for local Sage processes
                    CheckSageProcesses();
                    throw new InvalidOperationException("Sage is busy - close the application and retry");
                }
            }
            catch (Exception ex)
            {
                Console.WriteLine($"    Failed: {ex.Message}");
            }
        }

        if (!connected)
        {
            // Try direct COM approach using SDOApp objects
            Console.WriteLine("\n  Trying SDOApp direct approach...");
            TryDirectSdoAppConnection(dataPath, username, password);
        }
    }

    /// <summary>
    /// Check if Sage processes are running locally
    /// </summary>
    private void CheckSageProcesses()
    {
        try
        {
            var sageProcesses = System.Diagnostics.Process.GetProcesses()
                .Where(p => p.ProcessName.Contains("Sage", StringComparison.OrdinalIgnoreCase) ||
                           p.ProcessName.Contains("sg50", StringComparison.OrdinalIgnoreCase))
                .ToList();

            if (sageProcesses.Any())
            {
                Console.WriteLine("\n    Local Sage processes found:");
                foreach (var proc in sageProcesses)
                {
                    Console.WriteLine($"      - {proc.ProcessName} (PID: {proc.Id})");
                }
            }
            else
            {
                Console.WriteLine("\n    No local Sage processes found.");
                Console.WriteLine("    The lock may be on the server or from another workstation.");
            }
        }
        catch { }
    }

    /// <summary>
    /// List COM methods using ITypeInfo
    /// </summary>
    private void ListComMethods(string name, object comObject)
    {
        Console.WriteLine($"\n  {name} methods (via TypeInfo):");
        try
        {
            var type = comObject.GetType();

            // Try to get type info from COM object
            if (comObject is IDispatch dispatch)
            {
                Console.WriteLine("    (IDispatch available - would need TypeLib)");
            }

            // List what we can see through reflection
            var members = type.GetMembers(BindingFlags.Public | BindingFlags.Instance);
            foreach (var m in members.Where(m =>
                m.MemberType == MemberTypes.Method ||
                m.MemberType == MemberTypes.Property))
            {
                if (!m.Name.StartsWith("get_") && !m.Name.StartsWith("set_") &&
                    m.DeclaringType != typeof(object) &&
                    m.DeclaringType != typeof(MarshalByRefObject))
                {
                    Console.WriteLine($"    {m.MemberType}: {m.Name}");
                }
            }
        }
        catch (Exception ex)
        {
            Console.WriteLine($"    Error: {ex.Message}");
        }
    }

    // IDispatch interface for COM type info
    [ComImport]
    [Guid("00020400-0000-0000-C000-000000000046")]
    [InterfaceType(ComInterfaceType.InterfaceIsIUnknown)]
    interface IDispatch
    {
        int GetTypeInfoCount(out uint pctinfo);
        int GetTypeInfo(uint iTInfo, uint lcid, out IntPtr ppTInfo);
    }

    /// <summary>
    /// Try direct connection using SDOApp COM objects
    /// </summary>
    private void TryDirectSdoAppConnection(string dataPath, string username, string password)
    {
        // Try SDOApp.Invoicing - it might handle its own connection
        string[] sdoAppProgIds = new[]
        {
            "SDOApp.Invoicing",
            "SDOApp.Invoicing10",
            "SDOApp.Customer",
            "SDOApp.Customer10",
            "SDOApp.Control",
            "SDOApp.Controls"
        };

        foreach (var progId in sdoAppProgIds)
        {
            try
            {
                Console.WriteLine($"    Trying {progId}...");
                var type = Type.GetTypeFromProgID(progId);
                if (type == null)
                {
                    Console.WriteLine($"      Not registered");
                    continue;
                }

                dynamic obj = Activator.CreateInstance(type)!;
                Console.WriteLine($"      Created instance");

                // Try to explore methods
                try
                {
                    // Some SDOApp objects have a Connect or SetDataPath method
                    obj.SetDataPath(dataPath);
                    Console.WriteLine($"      SetDataPath succeeded!");
                }
                catch (Exception ex)
                {
                    Console.WriteLine($"      SetDataPath: {ex.Message}");
                }

                try
                {
                    obj.Connect(dataPath, username, password);
                    Console.WriteLine($"      Connect succeeded!");
                    _workspace = obj;
                    return;
                }
                catch (Exception ex)
                {
                    Console.WriteLine($"      Connect: {ex.Message}");
                }
            }
            catch (Exception ex)
            {
                Console.WriteLine($"      Error: {ex.Message}");
            }
        }

        throw new InvalidOperationException("Could not connect using any method");
    }

    /// <summary>
    /// Helper to invoke COM property getter
    /// </summary>
    private object? GetComProperty(object obj, string propertyName)
    {
        return obj.GetType().InvokeMember(
            propertyName,
            BindingFlags.GetProperty,
            null,
            obj,
            null);
    }

    /// <summary>
    /// Helper to invoke COM method
    /// </summary>
    private object? InvokeComMethod(object obj, string methodName, params object[] args)
    {
        return obj.GetType().InvokeMember(
            methodName,
            BindingFlags.InvokeMethod,
            null,
            obj,
            args);
    }

    /// <summary>
    /// Display company information
    /// </summary>
    public void ShowCompanyInfo()
    {
        if (_workspace == null) return;

        Console.WriteLine("\n[COMPANY INFORMATION]");

        try
        {
            // Try to get company data from SetupData or CompanyData record
            var setupData = InvokeComMethod(_workspace, "CreateObject", "SetupData");
            if (setupData != null)
            {
                // Open and read company setup
                InvokeComMethod(setupData, "Open", 0); // 0 = read mode

                var fields = GetComProperty(setupData, "Fields");
                if (fields != null)
                {
                    // SetupData fields typically include company name, address, etc.
                    DiscoverFields(setupData, "Company Setup Data", 15);
                }

                InvokeComMethod(setupData, "Close");
            }
        }
        catch (Exception ex)
        {
            Console.WriteLine($"  Error reading company info: {ex.Message}");

            // Fallback - just show the data path
            try
            {
                Console.WriteLine($"  Data Path: X:\\ACCDATA");
            }
            catch { }
        }
    }

    /// <summary>
    /// Helper to get a field value from a record's Fields collection
    /// </summary>
    private object? GetFieldValue(object record, string fieldName)
    {
        var fields = GetComProperty(record, "Fields");
        if (fields == null) return null;

        // Access field by name using Item property
        var field = fields.GetType().InvokeMember(
            "Item",
            BindingFlags.GetProperty,
            null,
            fields,
            new object[] { fieldName });

        if (field == null) return null;

        return GetComProperty(field, "Value");
    }

    /// <summary>
    /// List customers from Sales Ledger
    /// </summary>
    public void ListCustomers(int maxCount = 10)
    {
        if (_workspace == null) return;

        Console.WriteLine($"\n[CUSTOMERS (first {maxCount})]");

        try
        {
            // Access Sales Ledger using CreateObject
            var salesRecords = InvokeComMethod(_workspace, "CreateObject", "SalesRecord");
            if (salesRecords == null)
            {
                Console.WriteLine("  Could not create SalesRecord object");
                return;
            }

            InvokeComMethod(salesRecords, "MoveFirst");
            int count = 0;

            while (count < maxCount)
            {
                // Use IsEOF() method, not EOF property
                var eof = InvokeComMethod(salesRecords, "IsEOF");
                if (eof != null && (bool)eof) break;

                var acctRef = GetFieldValue(salesRecords, "ACCOUNT_REF");
                var name = GetFieldValue(salesRecords, "NAME");
                var balance = GetFieldValue(salesRecords, "BALANCE");

                Console.WriteLine($"  {acctRef,-12} {name,-30} Balance: {balance:C}");

                InvokeComMethod(salesRecords, "MoveNext");
                count++;
            }

            Console.WriteLine($"  (Showing {count} records)");
        }
        catch (Exception ex)
        {
            Console.WriteLine($"  Error listing customers: {ex.Message}");
            if (ex.InnerException != null)
                Console.WriteLine($"    Inner: {ex.InnerException.Message}");
        }
    }

    /// <summary>
    /// List recent invoices
    /// </summary>
    public void ListInvoices(int maxCount = 10)
    {
        if (_workspace == null) return;

        Console.WriteLine($"\n[RECENT INVOICES (first {maxCount})]");

        try
        {
            // Access Invoice records using CreateObject
            var invoiceRecords = InvokeComMethod(_workspace, "CreateObject", "InvoiceRecord");
            if (invoiceRecords == null)
            {
                Console.WriteLine("  Could not create InvoiceRecord object");
                return;
            }

            // Get record count
            try
            {
                var recCount = GetComProperty(invoiceRecords, "Count");
                Console.WriteLine($"  Total invoice records: {recCount}");
            }
            catch { }

            // Field discovery (commented out - use only when debugging)
            // DiscoverFields(invoiceRecords, "InvoiceRecord");

            InvokeComMethod(invoiceRecords, "MoveFirst");
            int count = 0;

            while (count < maxCount)
            {
                var eof = InvokeComMethod(invoiceRecords, "IsEOF");
                if (eof != null && (bool)eof) break;

                try
                {
                    // Known field indices (1-based):
                    // [1] ACCOUNT_REF, [60] INVOICE_DATE, [62] INVOICE_NUMBER
                    // [65] ITEMS_NET, [64] INVOICE_TYPE_CODE
                    var fields = GetComProperty(invoiceRecords, "Fields");
                    if (fields != null)
                    {
                        var acctRef = GetFieldByIndex(fields, 1);   // ACCOUNT_REF
                        var invDate = GetFieldByIndex(fields, 60);  // INVOICE_DATE
                        var invNum = GetFieldByIndex(fields, 62);   // INVOICE_NUMBER
                        var invType = GetFieldByIndex(fields, 64);  // INVOICE_TYPE_CODE
                        var netAmt = GetFieldByIndex(fields, 65);   // ITEMS_NET

                        var dateStr = invDate is DateTime dt ? dt.ToString("dd/MM/yyyy") : invDate?.ToString() ?? "";
                        Console.WriteLine($"  Inv# {invNum,-10} Account: {acctRef,-15} Date: {dateStr,-12} Type: {invType} Net: {netAmt:C}");
                    }
                }
                catch (Exception fieldEx)
                {
                    if (count == 0)
                    {
                        Console.WriteLine($"  Field error: {fieldEx.Message}");
                        if (fieldEx.InnerException != null)
                            Console.WriteLine($"    Inner: {fieldEx.InnerException.Message}");
                    }
                }

                InvokeComMethod(invoiceRecords, "MoveNext");
                count++;
            }

            Console.WriteLine($"  (Showing {count} records)");
        }
        catch (Exception ex)
        {
            Console.WriteLine($"  Error listing invoices: {ex.Message}");
        }
    }

    /// <summary>
    /// Read audit (transaction) records to understand field structure
    /// </summary>
    public void ReadAuditRecords(int maxCount = 20)
    {
        if (_workspace == null)
            throw new InvalidOperationException("Not connected to Sage");

        Console.WriteLine("\n[AUDIT RECORDS (Transactions)]");

        try
        {
            // Try different transaction record objects
            object? auditRecords = null;
            string? objectUsed = null;

            foreach (var objName in new[] { "AuditTrail", "TransactionHistory", "NominalTranRecord", "BankTranRecord", "SalesDebtorsTran", "PurchaseCreditorsTran" })
            {
                try
                {
                    auditRecords = InvokeComMethod(_workspace, "CreateObject", objName);
                    if (auditRecords != null)
                    {
                        objectUsed = objName;
                        Console.WriteLine($"  Found: {objName}");
                        break;
                    }
                }
                catch { }
            }

            if (auditRecords == null)
            {
                Console.WriteLine("  Could not find transaction records object");
                Console.WriteLine("  Let's check PurchaseRecord to see supplier transactions...");

                // Use PurchaseRecord to read supplier transactions
                auditRecords = InvokeComMethod(_workspace, "CreateObject", "PurchaseRecord");
                objectUsed = "PurchaseRecord";
            }

            Console.WriteLine($"  Using: {objectUsed}");

            var fields = GetComProperty(auditRecords, "Fields");
            if (fields != null)
            {
                // Discover all field names
                var count = (int)GetComProperty(fields, "Count")!;
                Console.WriteLine($"  Total fields: {count}");

                for (int i = 1; i <= Math.Min(count, 30); i++)
                {
                    try
                    {
                        var field = fields.GetType().InvokeMember("Item", BindingFlags.GetProperty, null, fields, new object[] { i });
                        var name = GetComProperty(field, "Name");
                        var value = GetComProperty(field, "Value");
                        Console.WriteLine($"    [{i}] {name} = {value}");
                    }
                    catch { }
                }
                if (count > 30)
                    Console.WriteLine($"    ... and {count - 30} more fields");
            }

            // Move to first record
            InvokeComMethod(auditRecords, "MoveFirst");

            Console.WriteLine("\n  Recent transactions:");
            int recordCount = 0;

            while (recordCount < maxCount)
            {
                var eof = GetComProperty(auditRecords, "Eof");
                if (eof != null && (bool)eof)
                    break;

                try
                {
                    // Key fields to identify transaction type
                    var typeVal = GetFieldByName(auditRecords, "TYPE");
                    var acctRef = GetFieldByName(auditRecords, "ACCOUNT_REF");
                    var date = GetFieldByName(auditRecords, "DATE");
                    var details = GetFieldByName(auditRecords, "DETAILS");
                    var net = GetFieldByName(auditRecords, "NET_AMOUNT");
                    var nominal = GetFieldByName(auditRecords, "NOMINAL_CODE");

                    var dateStr = date is DateTime dt ? dt.ToString("dd/MM/yyyy") : date?.ToString() ?? "";
                    Console.WriteLine($"    TYPE={typeVal} ACCT={acctRef} DATE={dateStr} NOM={nominal} NET={net:C} DETAILS={details}");
                }
                catch (Exception ex)
                {
                    Console.WriteLine($"    Error reading record: {ex.Message}");
                }

                InvokeComMethod(auditRecords, "MoveNext");
                recordCount++;
            }

            Console.WriteLine($"  (Showed {recordCount} records)");
        }
        catch (Exception ex)
        {
            Console.WriteLine($"  Error reading audit records: {ex.Message}");
            if (ex.InnerException != null)
                Console.WriteLine($"    Inner: {ex.InnerException.Message}");
        }
    }

    /// <summary>
    /// Get field value by numeric index (1-based)
    /// </summary>
    private object? GetFieldByIndex(object fields, int index)
    {
        var field = fields.GetType().InvokeMember(
            "Item",
            BindingFlags.GetProperty,
            null,
            fields,
            new object[] { index });

        if (field == null) return null;
        return GetComProperty(field, "Value");
    }

    /// <summary>
    /// Get field value by field name
    /// </summary>
    private object? GetFieldByName(object fieldsOrContainer, string fieldName)
    {
        object? field = null;

        try
        {
            // Try direct Item access (if it's a Fields collection)
            field = fieldsOrContainer.GetType().InvokeMember(
                "Item",
                BindingFlags.GetProperty,
                null,
                fieldsOrContainer,
                new object[] { fieldName });
        }
        catch
        {
            // Try getting Fields property first (if it's a header/item object)
            var fields = GetComProperty(fieldsOrContainer, "Fields");
            if (fields != null)
            {
                field = fields.GetType().InvokeMember(
                    "Item",
                    BindingFlags.GetProperty,
                    null,
                    fields,
                    new object[] { fieldName });
            }
        }

        if (field == null) return null;
        return GetComProperty(field, "Value");
    }

    /// <summary>
    /// Discover available field names on a record object
    /// </summary>
    private void DiscoverFields(object record, string recordName, int maxFields = 0)
    {
        Console.WriteLine($"\n  [{recordName}]");
        try
        {
            var fields = GetComProperty(record, "Fields");
            if (fields == null)
            {
                Console.WriteLine("    Could not get Fields collection");
                return;
            }

            // Get count of fields
            var fieldCount = fields.GetType().InvokeMember(
                "Count",
                BindingFlags.GetProperty,
                null,
                fields,
                null);

            int count = fieldCount is int fc ? fc : 0;
            int limit = maxFields > 0 ? Math.Min(count, maxFields) : count;

            Console.WriteLine($"    Total fields: {count}");

            // List field names (1-based indexing)
            for (int i = 1; i <= limit; i++)
            {
                try
                {
                    var field = fields.GetType().InvokeMember(
                        "Item",
                        BindingFlags.GetProperty,
                        null,
                        fields,
                        new object[] { i });

                    if (field != null)
                    {
                        var name = GetComProperty(field, "Name");
                        var value = GetComProperty(field, "Value");
                        Console.WriteLine($"    [{i}] {name} = {value}");
                    }
                }
                catch { }
            }

            if (maxFields > 0 && count > maxFields)
                Console.WriteLine($"    ... and {count - maxFields} more fields");
        }
        catch (Exception ex)
        {
            Console.WriteLine($"    Error discovering fields: {ex.Message}");
        }
    }

    /// <summary>
    /// Show example code for creating an invoice (does not execute)
    /// </summary>
    public void ShowInvoiceCreationExample()
    {
        Console.WriteLine("\n[INVOICE CREATION EXAMPLE]");
        Console.WriteLine("  The following code demonstrates how to create a Sales Invoice:");
        Console.WriteLine();
        Console.WriteLine(@"  // Create invoice object
  dynamic invoice = _workspace.CreateObject(""SopInvoice"");

  // Set header fields
  invoice.Fields[""ACCOUNT_REF""].Value = ""CUST001"";
  invoice.Fields[""INVOICE_DATE""].Value = DateTime.Now;
  invoice.Fields[""ORDER_NUMBER""].Value = ""PO12345"";

  // Add invoice items
  dynamic items = invoice.Items;
  dynamic item = items.Add();

  item.Fields[""STOCK_CODE""].Value = ""PROD001"";
  item.Fields[""DESCRIPTION""].Value = ""Product Description"";
  item.Fields[""QTY_ORDER""].Value = 1;
  item.Fields[""UNIT_PRICE""].Value = 100.00m;
  item.Fields[""NOMINAL_CODE""].Value = ""4000"";
  item.Fields[""TAX_CODE""].Value = ""T1"";

  // Post the invoice
  bool success = invoice.Post();

  if (success)
  {
      Console.WriteLine($""Invoice created: {invoice.Fields[""INVOICE_NUMBER""].Value}"");
  }
  else
  {
      Console.WriteLine(""Failed to post invoice"");
  }");

        Console.WriteLine("\n  NOTE: Actual posting commented out for safety.");
        Console.WriteLine("  Uncomment and modify the code above to post real invoices.");
    }

    /// <summary>
    /// Check if a customer account exists by iterating through records
    /// </summary>
    public bool CustomerExists(string accountRef)
    {
        if (_workspace == null) return false;

        try
        {
            var salesRecord = InvokeComMethod(_workspace, "CreateObject", "SalesRecord");
            if (salesRecord == null) return false;

            InvokeComMethod(salesRecord, "MoveFirst");

            // Iterate through all customers to find match
            while (true)
            {
                var eof = InvokeComMethod(salesRecord, "IsEOF");
                if (eof != null && (bool)eof) break;

                var acctRef = GetFieldValue(salesRecord, "ACCOUNT_REF");
                if (acctRef?.ToString()?.Equals(accountRef, StringComparison.OrdinalIgnoreCase) ?? false)
                {
                    return true;
                }

                InvokeComMethod(salesRecord, "MoveNext");
            }

            return false;
        }
        catch (Exception ex)
        {
            Console.WriteLine($"  CustomerExists error: {ex.Message}");
            return false;
        }
    }

    /// <summary>
    /// Check if a supplier account exists by iterating through records
    /// </summary>
    public bool SupplierExists(string accountRef)
    {
        if (_workspace == null) return false;

        try
        {
            var purchaseRecord = InvokeComMethod(_workspace, "CreateObject", "PurchaseRecord");
            if (purchaseRecord == null) return false;

            InvokeComMethod(purchaseRecord, "MoveFirst");

            while (true)
            {
                var eof = InvokeComMethod(purchaseRecord, "IsEOF");
                if (eof != null && (bool)eof) break;

                var acctRef = GetFieldValue(purchaseRecord, "ACCOUNT_REF");
                if (acctRef?.ToString()?.Equals(accountRef, StringComparison.OrdinalIgnoreCase) ?? false)
                {
                    return true;
                }

                InvokeComMethod(purchaseRecord, "MoveNext");
            }

            return false;
        }
        catch
        {
            return false;
        }
    }

    /// <summary>
    /// Create a new customer account
    /// </summary>
    public bool CreateCustomer(string accountRef, string name, string address1 = "", string address2 = "", string postcode = "")
    {
        if (_workspace == null) return false;

        Console.WriteLine($"\n[CREATING CUSTOMER: {accountRef}]");

        try
        {
            var salesRecord = InvokeComMethod(_workspace, "CreateObject", "SalesRecord");
            if (salesRecord == null)
            {
                Console.WriteLine("  Could not create SalesRecord");
                return false;
            }

            // Must call AddNew() to prepare for new record
            try { InvokeComMethod(salesRecord, "AddNew"); } catch { }

            // Set customer fields (ACCOUNT_REF must be max 8 chars, no spaces)
            SetFieldByName(salesRecord, "ACCOUNT_REF", accountRef);
            SetFieldByName(salesRecord, "NAME", name);
            if (!string.IsNullOrEmpty(address1))
                SetFieldByName(salesRecord, "ADDRESS_1", address1);
            if (!string.IsNullOrEmpty(address2))
                SetFieldByName(salesRecord, "ADDRESS_2", address2);
            if (!string.IsNullOrEmpty(postcode))
                SetFieldByName(salesRecord, "ADDRESS_5", postcode);

            Console.WriteLine($"  Set fields: {accountRef}, {name}");

            // Save the record
            var updateResult = InvokeComMethod(salesRecord, "Update");
            if (updateResult != null && (bool)updateResult)
            {
                Console.WriteLine("  Customer created successfully");
                return true;
            }
            else
            {
                Console.WriteLine("  Update failed");
                return false;
            }
        }
        catch (Exception ex)
        {
            Console.WriteLine($"  Error: {ex.Message}");
            return false;
        }
    }

    /// <summary>
    /// Create a new supplier account
    /// </summary>
    public bool CreateSupplier(string accountRef, string name, string address1 = "", string address2 = "", string postcode = "")
    {
        if (_workspace == null) return false;

        Console.WriteLine($"\n[CREATING SUPPLIER: {accountRef}]");

        try
        {
            var purchaseRecord = InvokeComMethod(_workspace, "CreateObject", "PurchaseRecord");
            if (purchaseRecord == null)
            {
                Console.WriteLine("  Could not create PurchaseRecord");
                return false;
            }

            // Must call AddNew() to prepare for new record
            try { InvokeComMethod(purchaseRecord, "AddNew"); } catch { }

            // Set supplier fields (ACCOUNT_REF must be max 8 chars, no spaces)
            SetFieldByName(purchaseRecord, "ACCOUNT_REF", accountRef);
            SetFieldByName(purchaseRecord, "NAME", name);
            if (!string.IsNullOrEmpty(address1))
                SetFieldByName(purchaseRecord, "ADDRESS_1", address1);
            if (!string.IsNullOrEmpty(address2))
                SetFieldByName(purchaseRecord, "ADDRESS_2", address2);
            if (!string.IsNullOrEmpty(postcode))
                SetFieldByName(purchaseRecord, "ADDRESS_5", postcode);

            Console.WriteLine($"  Set fields: {accountRef}, {name}");

            var updateResult = InvokeComMethod(purchaseRecord, "Update");
            if (updateResult != null && (bool)updateResult)
            {
                Console.WriteLine("  Supplier created successfully");
                return true;
            }
            else
            {
                Console.WriteLine("  Update failed");
                return false;
            }
        }
        catch (Exception ex)
        {
            Console.WriteLine($"  Error: {ex.Message}");
            return false;
        }
    }

    /// <summary>
    /// Get full customer account details by account reference
    /// </summary>
    public CustomerAccount? GetCustomer(string accountRef)
    {
        if (_workspace == null) return null;

        try
        {
            var salesRecord = InvokeComMethod(_workspace, "CreateObject", "SalesRecord");
            if (salesRecord == null) return null;

            InvokeComMethod(salesRecord, "MoveFirst");

            while (true)
            {
                var eof = InvokeComMethod(salesRecord, "IsEOF");
                if (eof != null && (bool)eof) break;

                var acctRef = GetFieldValue(salesRecord, "ACCOUNT_REF");
                if (acctRef?.ToString()?.Equals(accountRef, StringComparison.OrdinalIgnoreCase) ?? false)
                {
                    return ExtractCustomerFromRecord(salesRecord);
                }

                InvokeComMethod(salesRecord, "MoveNext");
            }

            return null;
        }
        catch (Exception ex)
        {
            Console.WriteLine($"  GetCustomer error: {ex.Message}");
            return null;
        }
    }

    /// <summary>
    /// Extract customer data from a SalesRecord
    /// </summary>
    private CustomerAccount ExtractCustomerFromRecord(object salesRecord)
    {
        var customer = new CustomerAccount();

        // Helper to safely get field value
        string SafeGetString(string fieldName)
        {
            try { return GetFieldValue(salesRecord, fieldName)?.ToString() ?? ""; }
            catch { return ""; }
        }

        customer.AccountRef = SafeGetString("ACCOUNT_REF");
        customer.Name = SafeGetString("NAME");
        customer.Address1 = SafeGetString("ADDRESS_1");
        customer.Address2 = SafeGetString("ADDRESS_2");
        customer.Address3 = SafeGetString("ADDRESS_3");
        customer.Address4 = SafeGetString("ADDRESS_4");
        customer.Postcode = SafeGetString("ADDRESS_5");
        customer.Telephone = SafeGetString("TELEPHONE");
        customer.Fax = SafeGetString("FAX");
        customer.Email = SafeGetString("E_MAIL");
        customer.ContactName = SafeGetString("CONTACT_NAME");
        customer.VatNumber = SafeGetString("VAT_REG_NUMBER");
        customer.NominalCode = SafeGetString("DEF_NOM_CODE");
        customer.Currency = SafeGetString("CURRENCY");

        try
        {
            var balance = GetFieldValue(salesRecord, "BALANCE");
            if (balance != null && decimal.TryParse(balance.ToString(), out var bal))
                customer.Balance = bal;
        }
        catch { }

        try
        {
            var creditLimit = GetFieldValue(salesRecord, "CREDIT_LIMIT");
            if (creditLimit != null && decimal.TryParse(creditLimit.ToString(), out var limit))
                customer.CreditLimit = limit;
        }
        catch { }

        try
        {
            var dateOpened = GetFieldValue(salesRecord, "DATE_OPENED");
            if (dateOpened is DateTime dt)
                customer.DateOpened = dt;
        }
        catch { }

        return customer;
    }

    /// <summary>
    /// Search customers by name or account reference (partial match)
    /// </summary>
    public List<CustomerAccount> FindCustomers(string searchTerm, int maxResults = 50)
    {
        var results = new List<CustomerAccount>();
        if (_workspace == null) return results;

        var search = searchTerm.ToUpperInvariant();
        Console.WriteLine($"  Searching for: '{search}'");

        try
        {
            var salesRecord = InvokeComMethod(_workspace, "CreateObject", "SalesRecord");
            if (salesRecord == null)
            {
                Console.WriteLine("  Could not create SalesRecord");
                return results;
            }

            InvokeComMethod(salesRecord, "MoveFirst");
            int scanned = 0;

            while (results.Count < maxResults)
            {
                var eof = InvokeComMethod(salesRecord, "IsEOF");
                if (eof != null && (bool)eof) break;

                try
                {
                    var acctRef = GetFieldValue(salesRecord, "ACCOUNT_REF")?.ToString() ?? "";
                    var name = GetFieldValue(salesRecord, "NAME")?.ToString() ?? "";
                    scanned++;

                    var acctUpper = acctRef.ToUpperInvariant();
                    var nameUpper = name.ToUpperInvariant();

                    if (acctUpper.Contains(search) || nameUpper.Contains(search))
                    {
                        var customer = ExtractCustomerFromRecord(salesRecord);
                        results.Add(customer);
                    }
                }
                catch (Exception fieldEx)
                {
                    Console.WriteLine($"  Field read error: {fieldEx.Message}");
                    scanned++;
                }

                InvokeComMethod(salesRecord, "MoveNext");
            }

            Console.WriteLine($"  Scanned {scanned} records");
        }
        catch (Exception ex)
        {
            Console.WriteLine($"  FindCustomers error: {ex.Message}");
            if (ex.InnerException != null)
                Console.WriteLine($"    Inner: {ex.InnerException.Message}");
        }

        return results;
    }

    /// <summary>
    /// Get full supplier account details by account reference
    /// </summary>
    public SupplierAccount? GetSupplier(string accountRef)
    {
        if (_workspace == null) return null;

        try
        {
            var purchaseRecord = InvokeComMethod(_workspace, "CreateObject", "PurchaseRecord");
            if (purchaseRecord == null) return null;

            InvokeComMethod(purchaseRecord, "MoveFirst");

            while (true)
            {
                var eof = InvokeComMethod(purchaseRecord, "IsEOF");
                if (eof != null && (bool)eof) break;

                var acctRef = GetFieldValue(purchaseRecord, "ACCOUNT_REF");
                if (acctRef?.ToString()?.Equals(accountRef, StringComparison.OrdinalIgnoreCase) ?? false)
                {
                    return ExtractSupplierFromRecord(purchaseRecord);
                }

                InvokeComMethod(purchaseRecord, "MoveNext");
            }

            return null;
        }
        catch (Exception ex)
        {
            Console.WriteLine($"  GetSupplier error: {ex.Message}");
            return null;
        }
    }

    /// <summary>
    /// Extract supplier data from a PurchaseRecord
    /// </summary>
    private SupplierAccount ExtractSupplierFromRecord(object purchaseRecord)
    {
        var supplier = new SupplierAccount();

        // Helper to safely get field value
        string SafeGetString(string fieldName)
        {
            try { return GetFieldValue(purchaseRecord, fieldName)?.ToString() ?? ""; }
            catch { return ""; }
        }

        supplier.AccountRef = SafeGetString("ACCOUNT_REF");
        supplier.Name = SafeGetString("NAME");
        supplier.Address1 = SafeGetString("ADDRESS_1");
        supplier.Address2 = SafeGetString("ADDRESS_2");
        supplier.Address3 = SafeGetString("ADDRESS_3");
        supplier.Address4 = SafeGetString("ADDRESS_4");
        supplier.Postcode = SafeGetString("ADDRESS_5");
        supplier.Telephone = SafeGetString("TELEPHONE");
        supplier.Fax = SafeGetString("FAX");
        supplier.Email = SafeGetString("E_MAIL");
        supplier.ContactName = SafeGetString("CONTACT_NAME");
        supplier.VatNumber = SafeGetString("VAT_REG_NUMBER");
        supplier.NominalCode = SafeGetString("DEF_NOM_CODE");
        supplier.Currency = SafeGetString("CURRENCY");

        try
        {
            var balance = GetFieldValue(purchaseRecord, "BALANCE");
            if (balance != null && decimal.TryParse(balance.ToString(), out var bal))
                supplier.Balance = bal;
        }
        catch { }

        try
        {
            var creditLimit = GetFieldValue(purchaseRecord, "CREDIT_LIMIT");
            if (creditLimit != null && decimal.TryParse(creditLimit.ToString(), out var limit))
                supplier.CreditLimit = limit;
        }
        catch { }

        try
        {
            var dateOpened = GetFieldValue(purchaseRecord, "DATE_OPENED");
            if (dateOpened is DateTime dt)
                supplier.DateOpened = dt;
        }
        catch { }

        return supplier;
    }

    /// <summary>
    /// Search suppliers by name or account reference (partial match)
    /// </summary>
    public List<SupplierAccount> FindSuppliers(string searchTerm, int maxResults = 50)
    {
        var results = new List<SupplierAccount>();
        if (_workspace == null) return results;

        var search = searchTerm.ToUpperInvariant();
        int scanned = 0;

        try
        {
            var purchaseRecord = InvokeComMethod(_workspace, "CreateObject", "PurchaseRecord");
            if (purchaseRecord == null) return results;

            InvokeComMethod(purchaseRecord, "MoveFirst");

            while (results.Count < maxResults)
            {
                var eof = InvokeComMethod(purchaseRecord, "IsEOF");
                if (eof != null && (bool)eof) break;

                try
                {
                    var acctRef = GetFieldValue(purchaseRecord, "ACCOUNT_REF")?.ToString() ?? "";
                    var name = GetFieldValue(purchaseRecord, "NAME")?.ToString() ?? "";
                    scanned++;

                    var acctUpper = acctRef.ToUpperInvariant();
                    var nameUpper = name.ToUpperInvariant();

                    if (acctUpper.Contains(search) || nameUpper.Contains(search))
                    {
                        results.Add(ExtractSupplierFromRecord(purchaseRecord));
                    }
                }
                catch { scanned++; }

                InvokeComMethod(purchaseRecord, "MoveNext");
            }
        }
        catch (Exception ex)
        {
            Console.WriteLine($"  FindSuppliers error: {ex.Message}");
        }

        return results;
    }

    /// <summary>
    /// Create a customer with extended fields
    /// </summary>
    public bool CreateCustomerEx(CustomerAccount customer)
    {
        if (_workspace == null) return false;
        if (string.IsNullOrWhiteSpace(customer.AccountRef))
            throw new ArgumentException("AccountRef is required");

        Console.WriteLine($"\n[CREATING CUSTOMER: {customer.AccountRef}]");

        try
        {
            var salesRecord = InvokeComMethod(_workspace, "CreateObject", "SalesRecord");
            if (salesRecord == null)
            {
                Console.WriteLine("  Could not create SalesRecord");
                return false;
            }

            // Must call AddNew() to prepare for new record
            try { InvokeComMethod(salesRecord, "AddNew"); } catch { }

            // Set all fields (ACCOUNT_REF must be max 8 chars, no spaces)
            SetFieldByName(salesRecord, "ACCOUNT_REF", customer.AccountRef);
            SetFieldByName(salesRecord, "NAME", customer.Name);

            if (!string.IsNullOrEmpty(customer.Address1))
                SetFieldByName(salesRecord, "ADDRESS_1", customer.Address1);
            if (!string.IsNullOrEmpty(customer.Address2))
                SetFieldByName(salesRecord, "ADDRESS_2", customer.Address2);
            if (!string.IsNullOrEmpty(customer.Address3))
                SetFieldByName(salesRecord, "ADDRESS_3", customer.Address3);
            if (!string.IsNullOrEmpty(customer.Address4))
                SetFieldByName(salesRecord, "ADDRESS_4", customer.Address4);
            if (!string.IsNullOrEmpty(customer.Postcode))
                SetFieldByName(salesRecord, "ADDRESS_5", customer.Postcode);
            if (!string.IsNullOrEmpty(customer.Country))
                SetFieldByName(salesRecord, "COUNTRY", customer.Country);
            if (!string.IsNullOrEmpty(customer.Telephone))
                SetFieldByName(salesRecord, "TELEPHONE", customer.Telephone);
            if (!string.IsNullOrEmpty(customer.Fax))
                SetFieldByName(salesRecord, "FAX", customer.Fax);
            if (!string.IsNullOrEmpty(customer.Email))
                SetFieldByName(salesRecord, "E_MAIL", customer.Email);
            if (!string.IsNullOrEmpty(customer.ContactName))
                SetFieldByName(salesRecord, "CONTACT_NAME", customer.ContactName);
            if (!string.IsNullOrEmpty(customer.VatNumber))
                SetFieldByName(salesRecord, "VAT_REG_NUMBER", customer.VatNumber);
            if (!string.IsNullOrEmpty(customer.NominalCode))
                SetFieldByName(salesRecord, "DEF_NOM_CODE", customer.NominalCode);
            if (!string.IsNullOrEmpty(customer.Currency))
                SetFieldByName(salesRecord, "CURRENCY", customer.Currency);
            if (customer.CreditLimit > 0)
                SetFieldByName(salesRecord, "CREDIT_LIMIT", customer.CreditLimit);

            Console.WriteLine($"  Set fields for: {customer.AccountRef} - {customer.Name}");

            var updateResult = InvokeComMethod(salesRecord, "Update");
            if (updateResult != null && (bool)updateResult)
            {
                Console.WriteLine("  Customer created successfully");
                return true;
            }
            else
            {
                Console.WriteLine("  Update failed");
                TryShowLastError();
                return false;
            }
        }
        catch (Exception ex)
        {
            Console.WriteLine($"  Error: {ex.Message}");
            return false;
        }
    }

    /// <summary>
    /// Create a supplier with extended fields
    /// </summary>
    public bool CreateSupplierEx(SupplierAccount supplier)
    {
        if (_workspace == null) return false;
        if (string.IsNullOrWhiteSpace(supplier.AccountRef))
            throw new ArgumentException("AccountRef is required");

        Console.WriteLine($"\n[CREATING SUPPLIER: {supplier.AccountRef}]");

        try
        {
            var purchaseRecord = InvokeComMethod(_workspace, "CreateObject", "PurchaseRecord");
            if (purchaseRecord == null)
            {
                Console.WriteLine("  Could not create PurchaseRecord");
                return false;
            }

            // Must call AddNew() to prepare for new record
            try
            {
                InvokeComMethod(purchaseRecord, "AddNew");
            }
            catch { }

            // Set all fields (ACCOUNT_REF must be max 8 chars, no spaces)
            SetFieldByName(purchaseRecord, "ACCOUNT_REF", supplier.AccountRef);
            SetFieldByName(purchaseRecord, "NAME", supplier.Name);

            if (!string.IsNullOrEmpty(supplier.Address1))
                SetFieldByName(purchaseRecord, "ADDRESS_1", supplier.Address1);
            if (!string.IsNullOrEmpty(supplier.Address2))
                SetFieldByName(purchaseRecord, "ADDRESS_2", supplier.Address2);
            if (!string.IsNullOrEmpty(supplier.Address3))
                SetFieldByName(purchaseRecord, "ADDRESS_3", supplier.Address3);
            if (!string.IsNullOrEmpty(supplier.Address4))
                SetFieldByName(purchaseRecord, "ADDRESS_4", supplier.Address4);
            if (!string.IsNullOrEmpty(supplier.Postcode))
                SetFieldByName(purchaseRecord, "ADDRESS_5", supplier.Postcode);
            if (!string.IsNullOrEmpty(supplier.Country))
                SetFieldByName(purchaseRecord, "COUNTRY", supplier.Country);
            if (!string.IsNullOrEmpty(supplier.Telephone))
                SetFieldByName(purchaseRecord, "TELEPHONE", supplier.Telephone);
            if (!string.IsNullOrEmpty(supplier.Fax))
                SetFieldByName(purchaseRecord, "FAX", supplier.Fax);
            if (!string.IsNullOrEmpty(supplier.Email))
                SetFieldByName(purchaseRecord, "E_MAIL", supplier.Email);
            if (!string.IsNullOrEmpty(supplier.ContactName))
                SetFieldByName(purchaseRecord, "CONTACT_NAME", supplier.ContactName);
            if (!string.IsNullOrEmpty(supplier.VatNumber))
                SetFieldByName(purchaseRecord, "VAT_REG_NUMBER", supplier.VatNumber);
            if (!string.IsNullOrEmpty(supplier.NominalCode))
                SetFieldByName(purchaseRecord, "DEF_NOM_CODE", supplier.NominalCode);
            if (!string.IsNullOrEmpty(supplier.Currency))
                SetFieldByName(purchaseRecord, "CURRENCY", supplier.Currency);
            if (supplier.CreditLimit > 0)
                SetFieldByName(purchaseRecord, "CREDIT_LIMIT", supplier.CreditLimit);

            Console.WriteLine($"  Set fields for: {supplier.AccountRef} - {supplier.Name}");

            Console.WriteLine("  Calling Update()...");
            var updateResult = InvokeComMethod(purchaseRecord, "Update");
            Console.WriteLine($"  Update result: {updateResult}");

            if (updateResult != null && Convert.ToBoolean(updateResult))
            {
                Console.WriteLine("  Supplier created successfully");
                return true;
            }
            else
            {
                Console.WriteLine("  Update failed");
                TryShowLastError();
                return false;
            }
        }
        catch (Exception ex)
        {
            Console.WriteLine($"  Error: {ex.Message}");
            if (ex.InnerException != null)
                Console.WriteLine($"    Inner: {ex.InnerException.Message}");
            return false;
        }
    }

    /// <summary>
    /// Post journal entries to the Nominal Ledger
    /// Journal entries must balance (total debits = total credits)
    /// </summary>
    /// <param name="journalLines">List of journal lines - must balance</param>
    /// <param name="reference">Transaction reference</param>
    /// <param name="date">Journal date (defaults to today)</param>
    /// <returns>True if posted successfully</returns>
    public bool PostJournal(List<JournalLine> journalLines, string reference, DateTime? date = null)
    {
        if (_workspace == null)
            throw new InvalidOperationException("Not connected to Sage");

        if (journalLines == null || journalLines.Count == 0)
            throw new ArgumentException("At least one journal line is required");

        // Validate journal balances
        var totalDebit = journalLines.Sum(j => j.Debit);
        var totalCredit = journalLines.Sum(j => j.Credit);

        if (Math.Abs(totalDebit - totalCredit) > 0.01m)
            throw new ArgumentException($"Journal does not balance. Debits: {totalDebit:C}, Credits: {totalCredit:C}");

        Console.WriteLine("\n[POSTING JOURNAL]");
        Console.WriteLine($"  Reference: {reference}");
        Console.WriteLine($"  Date: {(date ?? DateTime.Today):d}");
        Console.WriteLine($"  Lines: {journalLines.Count}");
        Console.WriteLine($"  Total: {totalDebit:C}");

        try
        {
            // Create JournalPost object
            var journalPost = InvokeComMethod(_workspace, "CreateObject", "JournalPost");
            if (journalPost == null)
            {
                Console.WriteLine("  ERROR: Could not create JournalPost object");
                return false;
            }

            Console.WriteLine("  Created JournalPost object");

            // Get header and set properties
            var header = GetComProperty(journalPost, "Header");
            if (header != null)
            {
                SetFieldByName(header, "DATE", date ?? DateTime.Today);
                SetFieldByName(header, "DETAILS", $"Journal: {reference}");
            }

            // Get items collection
            var items = GetComProperty(journalPost, "Items");
            if (items == null)
            {
                Console.WriteLine("  ERROR: Could not get Items collection");
                return false;
            }

            // Add each journal line
            // In Sage SDK, journal debits use positive NET_AMOUNT, credits use negative
            // Or use TYPE field: typically 0=debit, 1=credit for journals
            foreach (var line in journalLines)
            {
                var item = InvokeComMethod(items, "Add");
                if (item != null)
                {
                    SetFieldByName(item, "NOMINAL_CODE", line.NominalCode);
                    SetFieldByName(item, "DETAILS", line.Details);
                    SetFieldByName(item, "DATE", line.Date ?? date ?? DateTime.Today);

                    // Convert tax code
                    int taxCode = line.TaxCode switch
                    {
                        "T0" => 0, "T1" => 1, "T2" => 2, "T5" => 5, "T9" => 9, _ => 9
                    };
                    SetFieldByName(item, "TAX_CODE", taxCode);

                    // For journals: Debit = positive NET_AMOUNT with TYPE=0
                    //               Credit = positive NET_AMOUNT with TYPE=1
                    if (line.Debit > 0)
                    {
                        SetFieldByName(item, "NET_AMOUNT", line.Debit);
                        SetFieldByName(item, "TYPE", 0);  // Debit
                        Console.WriteLine($"    DR {line.NominalCode}: {line.Debit:C} - {line.Details}");
                    }
                    else if (line.Credit > 0)
                    {
                        SetFieldByName(item, "NET_AMOUNT", line.Credit);
                        SetFieldByName(item, "TYPE", 1);  // Credit
                        Console.WriteLine($"    CR {line.NominalCode}: {line.Credit:C} - {line.Details}");
                    }
                }
            }

            // Post the journal
            Console.WriteLine("  Calling Update()...");
            var result = InvokeComMethod(journalPost, "Update");
            Console.WriteLine($"  Update result: {result}");

            if (result != null && (bool)result)
            {
                Console.WriteLine("  SUCCESS! Journal posted to Nominal Ledger");
                return true;
            }
            else
            {
                Console.WriteLine("  Update() returned false");
                TryShowLastError();
                return false;
            }
        }
        catch (Exception ex)
        {
            Console.WriteLine($"  Journal posting error: {ex.Message}");
            if (ex.InnerException != null)
                Console.WriteLine($"    Inner: {ex.InnerException.Message}");
            return false;
        }
    }

    /// <summary>
    /// Post a simple debit/credit journal entry to two nominal codes
    /// </summary>
    public bool PostSimpleJournal(
        string debitNominal,
        string creditNominal,
        decimal amount,
        string reference,
        string details,
        DateTime? date = null)
    {
        var lines = new List<JournalLine>
        {
            new JournalLine
            {
                NominalCode = debitNominal,
                Debit = amount,
                Details = details,
                Reference = reference,
                Date = date
            },
            new JournalLine
            {
                NominalCode = creditNominal,
                Credit = amount,
                Details = details,
                Reference = reference,
                Date = date
            }
        };

        return PostJournal(lines, reference, date);
    }

    /// <summary>
    /// Post a bank payment transaction
    /// </summary>
    public bool PostBankPayment(
        string bankNominal,
        string expenseNominal,
        decimal netAmount,
        string reference,
        string details,
        string taxCode = "T1",
        DateTime? date = null)
    {
        if (_workspace == null)
            throw new InvalidOperationException("Not connected to Sage");

        Console.WriteLine("\n[POSTING BANK PAYMENT]");
        Console.WriteLine($"  Bank: {bankNominal}");
        Console.WriteLine($"  Expense: {expenseNominal}");
        Console.WriteLine($"  Amount: {netAmount:C}");
        Console.WriteLine($"  Reference: {reference}");

        try
        {
            // Create BankPost object for bank transactions
            var bankPost = InvokeComMethod(_workspace, "CreateObject", "BankPost");
            if (bankPost == null)
            {
                // Fallback to JournalPost
                Console.WriteLine("  BankPost not available, using JournalPost");
                return PostSimpleJournal(expenseNominal, bankNominal, netAmount, reference, details, date);
            }

            var header = GetComProperty(bankPost, "Header");
            if (header != null)
            {
                SetFieldByName(header, "BANK_CODE", bankNominal);
                SetFieldByName(header, "DATE", date ?? DateTime.Today);
                SetFieldByName(header, "REFERENCE", reference);
                SetFieldByName(header, "PAYMENT_TYPE", 1); // 1 = Payment, 2 = Receipt
            }

            var items = GetComProperty(bankPost, "Items");
            if (items != null)
            {
                var item = InvokeComMethod(items, "Add");
                if (item != null)
                {
                    SetFieldByName(item, "NOMINAL_CODE", expenseNominal);
                    SetFieldByName(item, "DETAILS", details);
                    SetFieldByName(item, "NET_AMOUNT", netAmount);

                    int tc = taxCode switch
                    {
                        "T0" => 0, "T1" => 1, "T2" => 2, "T5" => 5, "T9" => 9, _ => 1
                    };
                    SetFieldByName(item, "TAX_CODE", tc);
                }
            }

            Console.WriteLine("  Calling Update()...");
            var result = InvokeComMethod(bankPost, "Update");
            Console.WriteLine($"  Update result: {result}");

            if (result != null && (bool)result)
            {
                Console.WriteLine("  SUCCESS! Bank payment posted");
                return true;
            }
            else
            {
                Console.WriteLine("  Update() returned false");
                TryShowLastError();
                return false;
            }
        }
        catch (Exception ex)
        {
            Console.WriteLine($"  Bank payment error: {ex.Message}");
            return false;
        }
    }

    /// <summary>
    /// Post a bank receipt transaction
    /// </summary>
    public bool PostBankReceipt(
        string bankNominal,
        string incomeNominal,
        decimal netAmount,
        string reference,
        string details,
        string taxCode = "T1",
        DateTime? date = null)
    {
        if (_workspace == null)
            throw new InvalidOperationException("Not connected to Sage");

        Console.WriteLine("\n[POSTING BANK RECEIPT]");
        Console.WriteLine($"  Bank: {bankNominal}");
        Console.WriteLine($"  Income: {incomeNominal}");
        Console.WriteLine($"  Amount: {netAmount:C}");
        Console.WriteLine($"  Reference: {reference}");

        try
        {
            var bankPost = InvokeComMethod(_workspace, "CreateObject", "BankPost");
            if (bankPost == null)
            {
                Console.WriteLine("  BankPost not available, using JournalPost");
                return PostSimpleJournal(bankNominal, incomeNominal, netAmount, reference, details, date);
            }

            var header = GetComProperty(bankPost, "Header");
            if (header != null)
            {
                SetFieldByName(header, "BANK_CODE", bankNominal);
                SetFieldByName(header, "DATE", date ?? DateTime.Today);
                SetFieldByName(header, "REFERENCE", reference);
                SetFieldByName(header, "PAYMENT_TYPE", 2); // 2 = Receipt
            }

            var items = GetComProperty(bankPost, "Items");
            if (items != null)
            {
                var item = InvokeComMethod(items, "Add");
                if (item != null)
                {
                    SetFieldByName(item, "NOMINAL_CODE", incomeNominal);
                    SetFieldByName(item, "DETAILS", details);
                    SetFieldByName(item, "NET_AMOUNT", netAmount);

                    int tc = taxCode switch
                    {
                        "T0" => 0, "T1" => 1, "T2" => 2, "T5" => 5, "T9" => 9, _ => 1
                    };
                    SetFieldByName(item, "TAX_CODE", tc);
                }
            }

            Console.WriteLine("  Calling Update()...");
            var result = InvokeComMethod(bankPost, "Update");
            Console.WriteLine($"  Update result: {result}");

            if (result != null && (bool)result)
            {
                Console.WriteLine("  SUCCESS! Bank receipt posted");
                return true;
            }
            else
            {
                Console.WriteLine("  Update() returned false");
                TryShowLastError();
                return false;
            }
        }
        catch (Exception ex)
        {
            Console.WriteLine($"  Bank receipt error: {ex.Message}");
            return false;
        }
    }

    /// <summary>
    /// Discover the field structure of JournalPost object
    /// </summary>
    public void DiscoverJournalPostFields()
    {
        if (_workspace == null) return;

        Console.WriteLine("\n[DISCOVERING JOURNALPOST FIELDS]");

        try
        {
            var journalPost = InvokeComMethod(_workspace, "CreateObject", "JournalPost");
            if (journalPost == null)
            {
                Console.WriteLine("  Could not create JournalPost");
                return;
            }

            // Try to get Header
            try
            {
                var header = GetComProperty(journalPost, "Header");
                if (header != null)
                {
                    DiscoverFields(header, "JournalPost Header", 50);
                }
            }
            catch (Exception ex)
            {
                Console.WriteLine($"  Header error: {ex.Message}");
            }

            // Try to get Items and add one to see item fields
            try
            {
                var items = GetComProperty(journalPost, "Items");
                if (items != null)
                {
                    Console.WriteLine("\n  Items collection found");
                    var item = InvokeComMethod(items, "Add");
                    if (item != null)
                    {
                        DiscoverFields(item, "JournalPost Item", 50);
                    }
                }
            }
            catch (Exception ex)
            {
                Console.WriteLine($"  Items error: {ex.Message}");
            }

            // Also try JournalDebit and JournalCredit objects
            try
            {
                var jrnlDebit = InvokeComMethod(_workspace, "CreateObject", "JournalDebit");
                if (jrnlDebit != null)
                {
                    Console.WriteLine("\n  JournalDebit object available");
                    DiscoverFields(jrnlDebit, "JournalDebit", 30);
                }
            }
            catch { Console.WriteLine("  JournalDebit not available"); }

            try
            {
                var jrnlCredit = InvokeComMethod(_workspace, "CreateObject", "JournalCredit");
                if (jrnlCredit != null)
                {
                    Console.WriteLine("\n  JournalCredit object available");
                    DiscoverFields(jrnlCredit, "JournalCredit", 30);
                }
            }
            catch { Console.WriteLine("  JournalCredit not available"); }
        }
        catch (Exception ex)
        {
            Console.WriteLine($"  Error: {ex.Message}");
        }
    }

    /// <summary>
    /// Discover all available posting objects in Sage SDK
    /// </summary>
    public void DiscoverPostingObjects()
    {
        if (_workspace == null) return;

        Console.WriteLine("\n[DISCOVERING AVAILABLE POSTING OBJECTS]");

        // SDK posting object names (from PeNet analysis of SageSDO.Interop.dll)
        var postingObjects = new[]
        {
            // Direct posting objects (from SDK)
            "InvoicePost",        // IInvoicePost - direct invoice posting
            "TransactionPost",    // ITransactionPost - general transaction posting
            "JournalPost",        // IJournalPost - journal entries
            "StockPost",          // IStockPost - stock transactions
            "SDOIspPost",         // ISDOIspPost - unknown

            // Order processing
            "SopPost",            // ISopPost - Sales Order Processing
            "PopPost",            // IPopPost - Purchase Order Processing

            // Try other potential names
            "SalesInvoicePost", "SInvoicePost", "SIPost",
            "SalesDebitPost", "SalesDebit", "SDebitPost",
            "SalesCreditPost", "SalesCredit", "SCreditPost",
            "SalesReceiptPost", "SalesReceipt", "SRPost",
            "PurchaseInvoicePost", "PInvoicePost", "PIPost",
            "PurchaseDebitPost", "PurchaseDebit", "PDebitPost",
            "PurchaseCreditPost", "PurchaseCredit", "PCreditPost",
            "PurchasePaymentPost", "PurchasePayment", "PPPost",
            "BankPost", "BankPaymentPost", "BankReceiptPost",
            "NominalPost",

            // Order Processing (for reference)
            "SopPost", "SopItem", "SopInvoice",
            "PopPost", "PopItem", "PopInvoice",

            // Other transaction types
            "AuditPost", "StockPost", "ProjectPost"
        };

        var available = new List<string>();

        foreach (var objName in postingObjects)
        {
            try
            {
                var obj = InvokeComMethod(_workspace, "CreateObject", objName);
                if (obj != null)
                {
                    available.Add(objName);
                    Console.WriteLine($"  ✓ {objName}");
                }
            }
            catch
            {
                // Not available
            }
        }

        Console.WriteLine($"\nFound {available.Count} posting objects");

        // Show fields for key posting objects (the ones that actually exist)
        foreach (var objName in new[] { "InvoicePost", "TransactionPost", "JournalPost" })
        {
            if (available.Contains(objName))
            {
                try
                {
                    var obj = InvokeComMethod(_workspace, "CreateObject", objName);
                    if (obj != null)
                    {
                        Console.WriteLine($"\n--- {objName} Fields ---");

                        // Try Header
                        try
                        {
                            var header = GetComProperty(obj, "Header");
                            if (header != null)
                                DiscoverFields(header, $"{objName} Header", 70);
                        }
                        catch { }

                        // Try Items
                        try
                        {
                            var items = GetComProperty(obj, "Items");
                            if (items != null)
                            {
                                var item = InvokeComMethod(items, "Add");
                                if (item != null)
                                    DiscoverFields(item, $"{objName} Item", 50);
                            }
                        }
                        catch { }

                        // Try direct Fields (for single-record objects)
                        try
                        {
                            DiscoverFields(obj, $"{objName} (direct)", 30);
                        }
                        catch { }
                    }
                }
                catch { }
            }
        }
    }

    /// <summary>
    /// Post a Sales Invoice directly to the Sales Ledger using TransactionPost
    /// This creates an actual financial transaction that updates the customer balance
    /// </summary>
    /// <param name="customerAccount">Customer account reference (max 8 chars)</param>
    /// <param name="invoiceRef">Invoice reference number</param>
    /// <param name="netAmount">Net amount (ex VAT)</param>
    /// <param name="taxAmount">VAT amount</param>
    /// <param name="nominalCode">Nominal code for sales (default 4000)</param>
    /// <param name="details">Transaction details/description</param>
    /// <param name="taxCode">Tax code (T0, T1, T2, etc.)</param>
    /// <param name="date">Invoice date</param>
    /// <param name="postToLedger">If true, posts to nominal ledger (always true for TransactionPost)</param>
    public bool PostSalesInvoice(
        string customerAccount,
        string invoiceRef,
        decimal netAmount,
        decimal taxAmount,
        string nominalCode = "4000",
        string details = "",
        string taxCode = "T1",
        DateTime? date = null,
        bool postToLedger = true)
    {
        // TransactionPost is the correct SDK object for posting to the ledger
        // InvoicePost creates order documents but doesn't update customer balances
        return PostTransactionSI(customerAccount, invoiceRef, netAmount, taxAmount,
                                nominalCode, details, taxCode, date, postToLedger);
    }

    /// <summary>
    /// Post a Sales Invoice using TransactionPost
    /// TransactionPost is a generic transaction posting object in Sage 50
    /// </summary>
    public bool PostTransactionSI(
        string customerAccount,
        string invoiceRef,
        decimal netAmount,
        decimal taxAmount,
        string nominalCode = "4000",
        string details = "",
        string taxCode = "T1",
        DateTime? date = null,
        bool postToLedger = true)
    {
        if (_workspace == null)
            throw new InvalidOperationException("Not connected to Sage");

        Console.WriteLine("\n[POSTING VIA TRANSACTIONPOST]");
        Console.WriteLine($"  Customer: {customerAccount}");
        Console.WriteLine($"  Invoice Ref: {invoiceRef}");
        Console.WriteLine($"  Net: {netAmount:C}, VAT: {taxAmount:C}");

        try
        {
            var txPost = InvokeComMethod(_workspace, "CreateObject", "TransactionPost");
            if (txPost == null)
            {
                Console.WriteLine("  TransactionPost not available");
                return false;
            }

            Console.WriteLine("  Using: TransactionPost");

            // Set header fields
            var header = GetComProperty(txPost, "Header");
            if (header == null)
            {
                Console.WriteLine("  ERROR: Could not get Header");
                return false;
            }

            // Set transaction header - TYPE=1 is Sales Invoice
            SetFieldByName(header, "ACCOUNT_REF", customerAccount);
            SetFieldByName(header, "DATE", date ?? DateTime.Today);
            SetFieldByName(header, "INV_REF", invoiceRef);  // Use INV_REF for invoice reference
            SetFieldByName(header, "DETAILS", string.IsNullOrEmpty(details) ? "Sales Invoice" : details);
            SetFieldByName(header, "TYPE", 1); // 1 = Sales Invoice (SI)

            int tc = taxCode switch { "T0" => 0, "T1" => 1, "T2" => 2, "T5" => 5, "T9" => 9, _ => 1 };
            decimal grossAmount = netAmount + taxAmount;
            SetFieldByName(header, "NET_AMOUNT", netAmount);
            SetFieldByName(header, "TAX_AMOUNT", taxAmount);

            // Set item (split) for nominal posting
            var items = GetComProperty(txPost, "Items");
            if (items != null)
            {
                var item = InvokeComMethod(items, "Add");
                if (item != null)
                {
                    SetFieldByName(item, "NOMINAL_CODE", nominalCode);
                    SetFieldByName(item, "DETAILS", string.IsNullOrEmpty(details) ? "Sales Invoice" : details);
                    SetFieldByName(item, "NET_AMOUNT", netAmount);
                    SetFieldByName(item, "TAX_AMOUNT", taxAmount);
                    SetFieldByName(item, "TAX_CODE", tc);
                }
            }

            Console.WriteLine($"  Header: Account={customerAccount}, Type=1 (SI), Gross={grossAmount:C}");

            Console.WriteLine("  Calling Update()...");
            var result = InvokeComMethod(txPost, "Update");
            Console.WriteLine($"  Update result: {result}");

            if (result != null && Convert.ToBoolean(result))
            {
                Console.WriteLine("  SUCCESS! Transaction posted");
                return true;
            }
            else
            {
                TryShowLastError();
                return false;
            }
        }
        catch (Exception ex)
        {
            Console.WriteLine($"  Error: {ex.Message}");
            return false;
        }
    }

    /// <summary>
    /// Post a Sales Debit (alternative method for sales invoices)
    /// </summary>
    public bool PostSalesDebit(
        string customerAccount,
        string reference,
        decimal netAmount,
        decimal taxAmount,
        string nominalCode = "4000",
        string details = "",
        string taxCode = "T1",
        DateTime? date = null,
        bool postToLedger = true)
    {
        // Delegate to PostTransactionSI
        return PostTransactionSI(customerAccount, reference, netAmount, taxAmount,
                                nominalCode, details, taxCode, date, postToLedger);
    }

    /// <summary>
    /// Test TransactionPost with a specific TYPE value
    /// TransType enum: 1=SI, 2=SC, 3=SR, 4=SA, 5=SD, 6=PI, 7=PC, 8=PP, 9=PA, 10=PD
    /// </summary>
    public bool TestTransactionType(string accountRef, int transType, string nominalCode, decimal netAmount, decimal taxAmount)
    {
        if (_workspace == null)
            throw new InvalidOperationException("Not connected to Sage");

        var invoiceRef = $"TEST-{transType}-{DateTime.Now:HHmmss}";
        Console.WriteLine($"  Account: {accountRef}, Nominal: {nominalCode}, Net: {netAmount:C}");

        try
        {
            var txPost = InvokeComMethod(_workspace, "CreateObject", "TransactionPost");
            if (txPost == null)
            {
                Console.WriteLine("  TransactionPost not available");
                return false;
            }

            var header = GetComProperty(txPost, "Header");
            if (header == null)
            {
                Console.WriteLine("  ERROR: Could not get Header");
                return false;
            }

            // Set header - same pattern as working sales invoice
            SetFieldByName(header, "ACCOUNT_REF", accountRef);
            SetFieldByName(header, "DATE", DateTime.Today);
            SetFieldByName(header, "INV_REF", invoiceRef);
            SetFieldByName(header, "DETAILS", $"Test TYPE={transType}");
            SetFieldByName(header, "TYPE", transType);
            SetFieldByName(header, "NET_AMOUNT", netAmount);
            SetFieldByName(header, "TAX_AMOUNT", taxAmount);

            // Set item/split
            var items = GetComProperty(txPost, "Items");
            if (items != null)
            {
                var item = InvokeComMethod(items, "Add");
                if (item != null)
                {
                    SetFieldByName(item, "NOMINAL_CODE", nominalCode);
                    SetFieldByName(item, "DETAILS", $"Test TYPE={transType}");
                    SetFieldByName(item, "NET_AMOUNT", netAmount);
                    SetFieldByName(item, "TAX_AMOUNT", taxAmount);
                    // CORRECTED: Purchase types are 6-10, not 4-5
                    // TransType enum: 6=PI, 7=PC, 8=PP, 9=PA, 10=PD
                    int taxCode = (transType >= 6 && transType <= 10) ? 9 : 1;
                    SetFieldByName(item, "TAX_CODE", taxCode);
                    Console.WriteLine($"  Using TAX_CODE: T{taxCode}");
                }
            }

            var result = InvokeComMethod(txPost, "Update");
            if (result != null && Convert.ToBoolean(result))
            {
                Console.WriteLine($"  Posted: {invoiceRef}");
                return true;
            }
            else
            {
                TryShowLastError();
                return false;
            }
        }
        catch (Exception ex)
        {
            Console.WriteLine($"  Error: {ex.Message}");
            if (ex.InnerException != null)
                Console.WriteLine($"  Inner: {ex.InnerException.Message}");
            TryShowLastError();
            return false;
        }
    }

    /// <summary>
    /// Test TransactionPost without setting TYPE on item
    /// </summary>
    public bool TestTransactionTypeNoItemType(string accountRef, int transType, string nominalCode, decimal netAmount, decimal taxAmount)
    {
        if (_workspace == null) return false;

        var invoiceRef = $"MIN-{transType}-{DateTime.Now:HHmmss}";
        Console.WriteLine($"  Account: {accountRef}, Nominal: {nominalCode}, Net: {netAmount:C}");

        try
        {
            var txPost = InvokeComMethod(_workspace, "CreateObject", "TransactionPost");
            if (txPost == null) return false;

            var header = GetComProperty(txPost, "Header");
            if (header == null) return false;

            // Set TYPE first
            SetFieldByName(header, "TYPE", transType);
            SetFieldByName(header, "ACCOUNT_REF", accountRef);
            SetFieldByName(header, "DATE", DateTime.Today);
            SetFieldByName(header, "INV_REF", invoiceRef);
            SetFieldByName(header, "DETAILS", $"Minimal TYPE={transType}");
            SetFieldByName(header, "NET_AMOUNT", netAmount);
            SetFieldByName(header, "TAX_AMOUNT", taxAmount);

            // Add item but DON'T set TYPE on item
            var items = GetComProperty(txPost, "Items");
            if (items != null)
            {
                var item = InvokeComMethod(items, "Add");
                if (item != null)
                {
                    SetFieldByName(item, "NOMINAL_CODE", nominalCode);
                    SetFieldByName(item, "NET_AMOUNT", netAmount);
                    SetFieldByName(item, "TAX_AMOUNT", taxAmount);
                    int taxCode = (transType == 4 || transType == 5) ? 9 : 1;
                    SetFieldByName(item, "TAX_CODE", taxCode);
                    // NOT setting TYPE on item
                }
            }

            var result = InvokeComMethod(txPost, "Update");
            if (result != null && Convert.ToBoolean(result))
            {
                Console.WriteLine($"  Posted: {invoiceRef}");
                return true;
            }
            else
            {
                TryShowLastError();
                return false;
            }
        }
        catch (Exception ex)
        {
            Console.WriteLine($"  Error: {ex.Message}");
            if (ex.InnerException != null)
                Console.WriteLine($"  Inner: {ex.InnerException.Message}");
            return false;
        }
    }

    /// <summary>
    /// Post a Purchase Invoice using JournalPost with supplier linking
    /// Alternative method using journal entries (TransactionPost TYPE=6 is preferred)
    /// </summary>
    public bool PostPurchaseInvoiceViaJournal(
        string supplierAccount,
        string invoiceRef,
        decimal netAmount,
        decimal taxAmount,
        string nominalCode = "4000",
        string details = "",
        DateTime? date = null)
    {
        if (_workspace == null)
            throw new InvalidOperationException("Not connected to Sage");

        Console.WriteLine("\n[POSTING PURCHASE INVOICE VIA JOURNALPOST]");
        Console.WriteLine($"  Supplier: {supplierAccount}");
        Console.WriteLine($"  Invoice Ref: {invoiceRef}");
        Console.WriteLine($"  Net: {netAmount:C}, VAT: {taxAmount:C}");
        Console.WriteLine($"  Nominal: {nominalCode}");

        try
        {
            var journalPost = InvokeComMethod(_workspace, "CreateObject", "JournalPost");
            if (journalPost == null)
            {
                Console.WriteLine("  ERROR: Could not create JournalPost");
                return false;
            }

            var header = GetComProperty(journalPost, "Header");
            if (header != null)
            {
                SetFieldByName(header, "DATE", date ?? DateTime.Today);
                SetFieldByName(header, "REFERENCE", invoiceRef);
                SetFieldByName(header, "DETAILS", details);
                // Try setting ACCOUNT_REF on header for supplier link
                SetFieldByName(header, "ACCOUNT_REF", supplierAccount);
                // Set TYPE=4 for Purchase Invoice
                SetFieldByName(header, "TYPE", 4);
            }

            var items = GetComProperty(journalPost, "Items");
            if (items != null)
            {
                decimal grossAmount = netAmount + taxAmount;

                // Debit the expense/cost account
                var debitItem = InvokeComMethod(items, "Add");
                if (debitItem != null)
                {
                    SetFieldByName(debitItem, "NOMINAL_CODE", nominalCode);
                    SetFieldByName(debitItem, "DETAILS", details);
                    SetFieldByName(debitItem, "NET_AMOUNT", netAmount);
                    SetFieldByName(debitItem, "TAX_AMOUNT", taxAmount);
                    SetFieldByName(debitItem, "TAX_CODE", 1); // T1 standard VAT
                    SetFieldByName(debitItem, "TYPE", 0); // Debit
                    Console.WriteLine($"  DR {nominalCode}: {grossAmount:C}");
                }

                // Credit the creditors control account
                var creditItem = InvokeComMethod(items, "Add");
                if (creditItem != null)
                {
                    SetFieldByName(creditItem, "NOMINAL_CODE", "2100"); // Creditors Control
                    SetFieldByName(creditItem, "DETAILS", details);
                    SetFieldByName(creditItem, "NET_AMOUNT", grossAmount);
                    SetFieldByName(creditItem, "TAX_CODE", 9); // T9 no VAT on control account
                    SetFieldByName(creditItem, "TYPE", 1); // Credit
                    // Try to link to supplier
                    SetFieldByName(creditItem, "ACCOUNT_REF", supplierAccount);
                    Console.WriteLine($"  CR 2100: {grossAmount:C} (supplier: {supplierAccount})");
                }
            }

            Console.WriteLine("  Calling Update()...");
            var result = InvokeComMethod(journalPost, "Update");
            Console.WriteLine($"  Update result: {result}");

            if (result != null && Convert.ToBoolean(result))
            {
                Console.WriteLine("  SUCCESS! Purchase Invoice posted via Journal");
                return true;
            }
            else
            {
                TryShowLastError();
                return false;
            }
        }
        catch (Exception ex)
        {
            Console.WriteLine($"  Error: {ex.Message}");
            if (ex.InnerException != null)
                Console.WriteLine($"    Inner: {ex.InnerException.Message}");
            return false;
        }
    }

    /// <summary>
    /// Post a Purchase Invoice using TransactionPost
    /// TYPE=6 is Purchase Invoice (PI) in Sage TransType enum
    /// Tax code T9 is required for purchase transaction types (6-10)
    /// </summary>
    public bool PostTransactionPI(
        string supplierAccount,
        string invoiceRef,
        decimal netAmount,
        decimal taxAmount,
        string nominalCode = "5000",
        string details = "",
        string taxCode = "T1",
        DateTime? date = null,
        bool postToLedger = true)
    {
        if (_workspace == null)
            throw new InvalidOperationException("Not connected to Sage");

        Console.WriteLine("\n[POSTING PURCHASE INVOICE VIA TRANSACTIONPOST]");
        Console.WriteLine($"  Supplier: {supplierAccount}");
        Console.WriteLine($"  Invoice Ref: {invoiceRef}");
        Console.WriteLine($"  Net: {netAmount:C}, VAT: {taxAmount:C}");

        try
        {
            var txPost = InvokeComMethod(_workspace, "CreateObject", "TransactionPost");
            if (txPost == null)
            {
                Console.WriteLine("  TransactionPost not available");
                return false;
            }

            Console.WriteLine("  Using: TransactionPost");

            // Set header fields
            var header = GetComProperty(txPost, "Header");
            if (header == null)
            {
                Console.WriteLine("  ERROR: Could not get Header");
                return false;
            }

            // Match EXACTLY the sales invoice structure that works
            // Set ACCOUNT_REF first, then TYPE (like sales invoice)
            SetFieldByName(header, "ACCOUNT_REF", supplierAccount);
            SetFieldByName(header, "DATE", date ?? DateTime.Today);
            SetFieldByName(header, "INV_REF", invoiceRef);
            SetFieldByName(header, "DETAILS", string.IsNullOrEmpty(details) ? "Purchase Invoice" : details);
            SetFieldByName(header, "TYPE", 6); // 6 = Purchase Invoice (PI) per TransType enum

            // Purchase types (6-10) require T9 tax code
            int tc = 9; // T9 required for purchase transactions
            decimal grossAmount = netAmount + taxAmount;
            SetFieldByName(header, "NET_AMOUNT", netAmount);
            SetFieldByName(header, "TAX_AMOUNT", taxAmount);

            // Set item (split) for nominal posting - exactly like sales invoice
            var items = GetComProperty(txPost, "Items");
            if (items != null)
            {
                var item = InvokeComMethod(items, "Add");
                if (item != null)
                {
                    SetFieldByName(item, "NOMINAL_CODE", nominalCode);
                    SetFieldByName(item, "DETAILS", string.IsNullOrEmpty(details) ? "Purchase Invoice" : details);
                    SetFieldByName(item, "NET_AMOUNT", netAmount);
                    SetFieldByName(item, "TAX_AMOUNT", taxAmount);
                    SetFieldByName(item, "TAX_CODE", tc);
                }
            }

            Console.WriteLine($"  Header: Account={supplierAccount}, Type=6 (PI), Nominal={nominalCode}, Tax=T9, Gross={grossAmount:C}");

            Console.WriteLine("  Calling Update()...");
            try
            {
                var result = InvokeComMethod(txPost, "Update");
                Console.WriteLine($"  Update result: {result}");

                if (result != null && Convert.ToBoolean(result))
                {
                    Console.WriteLine("  SUCCESS! Purchase Invoice posted");
                    return true;
                }
                else
                {
                    Console.WriteLine("  Update returned false");
                    TryShowLastError();
                    return false;
                }
            }
            catch (Exception updateEx)
            {
                Console.WriteLine($"  Update error: {updateEx.Message}");
                if (updateEx.InnerException != null)
                    Console.WriteLine($"    Inner: {updateEx.InnerException.Message}");
                TryShowLastError();
                return false;
            }
        }
        catch (Exception ex)
        {
            Console.WriteLine($"  Error: {ex.Message}");
            if (ex.InnerException != null)
                Console.WriteLine($"    Inner: {ex.InnerException.Message}");
            return false;
        }
    }

    /// <summary>
    /// Post a Purchase Invoice directly to the Purchase Ledger using TransactionPost
    /// This creates an actual financial transaction that updates the supplier balance
    /// </summary>
    /// <param name="supplierAccount">Supplier account reference (max 8 chars)</param>
    /// <param name="invoiceRef">Invoice reference number</param>
    /// <param name="netAmount">Net amount (ex VAT)</param>
    /// <param name="taxAmount">VAT amount</param>
    /// <param name="nominalCode">Nominal code for purchases (default 5000)</param>
    /// <param name="details">Transaction details/description</param>
    /// <param name="taxCode">Tax code (T0, T1, T2, etc.)</param>
    /// <param name="date">Invoice date</param>
    /// <param name="postToLedger">If true, posts to nominal ledger (always true for TransactionPost)</param>
    public bool PostPurchaseInvoice(
        string supplierAccount,
        string invoiceRef,
        decimal netAmount,
        decimal taxAmount,
        string nominalCode = "5000",
        string details = "",
        string taxCode = "T1",
        DateTime? date = null,
        bool postToLedger = true)
    {
        // TransactionPost is the correct SDK object for posting to the purchase ledger
        return PostTransactionPI(supplierAccount, invoiceRef, netAmount, taxAmount,
                                nominalCode, details, taxCode, date, postToLedger);
    }

    /// <summary>
    /// Post a Purchase Credit (alternative method for purchase invoices)
    /// </summary>
    public bool PostPurchaseCredit(
        string supplierAccount,
        string reference,
        decimal netAmount,
        decimal taxAmount,
        string nominalCode = "5000",
        string details = "",
        string taxCode = "T1",
        DateTime? date = null,
        bool postToLedger = true)
    {
        if (_workspace == null)
            throw new InvalidOperationException("Not connected to Sage");

        Console.WriteLine("\n[POSTING PURCHASE CREDIT]");
        Console.WriteLine($"  Supplier: {supplierAccount}");
        Console.WriteLine($"  Reference: {reference}");
        Console.WriteLine($"  Net: {netAmount:C}, VAT: {taxAmount:C}");

        try
        {
            var pcPost = InvokeComMethod(_workspace, "CreateObject", "PurchaseCreditPost");
            if (pcPost == null)
            {
                Console.WriteLine("  PurchaseCreditPost not available");
                return false;
            }

            var header = GetComProperty(pcPost, "Header");
            if (header != null)
            {
                SetFieldByName(header, "ACCOUNT_REF", supplierAccount);
                SetFieldByName(header, "DATE", date ?? DateTime.Today);
                SetFieldByName(header, "REFERENCE", reference);
                SetFieldByName(header, "DETAILS", details);
                SetFieldByName(header, "NET_AMOUNT", netAmount);
                SetFieldByName(header, "TAX_AMOUNT", taxAmount);

                int tc = taxCode switch { "T0" => 0, "T1" => 1, "T2" => 2, "T5" => 5, "T9" => 9, _ => 1 };
                SetFieldByName(header, "TAX_CODE", tc);
            }

            try
            {
                var items = GetComProperty(pcPost, "Items");
                if (items != null)
                {
                    var item = InvokeComMethod(items, "Add");
                    if (item != null)
                    {
                        SetFieldByName(item, "NOMINAL_CODE", nominalCode);
                        SetFieldByName(item, "DETAILS", details);
                        SetFieldByName(item, "NET_AMOUNT", netAmount);
                        SetFieldByName(item, "TAX_AMOUNT", taxAmount);
                    }
                }
            }
            catch { }

            Console.WriteLine("  Calling Update()...");
            var result = InvokeComMethod(pcPost, "Update");
            Console.WriteLine($"  Update result: {result}");

            if (result != null && Convert.ToBoolean(result))
            {
                Console.WriteLine("  SUCCESS! Purchase Credit posted");
                return true;
            }
            else
            {
                TryShowLastError();
                return false;
            }
        }
        catch (Exception ex)
        {
            Console.WriteLine($"  Error: {ex.Message}");
            return false;
        }
    }

    /// <summary>
    /// List nominal codes/accounts
    /// </summary>
    public void ListNominalCodes(int maxCount = 50)
    {
        if (_workspace == null) return;

        Console.WriteLine($"\n[NOMINAL CODES (first {maxCount})]");

        try
        {
            var nominalRecord = InvokeComMethod(_workspace, "CreateObject", "NominalRecord");
            if (nominalRecord == null)
            {
                Console.WriteLine("  Could not create NominalRecord object");
                return;
            }

            InvokeComMethod(nominalRecord, "MoveFirst");
            int count = 0;

            while (count < maxCount)
            {
                var eof = InvokeComMethod(nominalRecord, "IsEOF");
                if (eof != null && (bool)eof) break;

                var nomCode = GetFieldValue(nominalRecord, "ACCOUNT_REF");
                var name = GetFieldValue(nominalRecord, "NAME");
                var balance = GetFieldValue(nominalRecord, "BALANCE");

                Console.WriteLine($"  {nomCode,-8} {name,-40} Balance: {balance:C}");

                InvokeComMethod(nominalRecord, "MoveNext");
                count++;
            }

            Console.WriteLine($"  (Showing {count} records)");
        }
        catch (Exception ex)
        {
            Console.WriteLine($"  Error listing nominal codes: {ex.Message}");
        }
    }

    /// <summary>
    /// Create a new nominal code
    /// </summary>
    public bool CreateNominal(string nominalCode, string name)
    {
        if (_workspace == null) return false;

        Console.WriteLine($"\n[CREATING NOMINAL CODE: {nominalCode}]");

        try
        {
            var nominalRecord = InvokeComMethod(_workspace, "CreateObject", "NominalRecord");
            if (nominalRecord == null)
            {
                Console.WriteLine("  Could not create NominalRecord object");
                return false;
            }

            // Add new record
            InvokeComMethod(nominalRecord, "Add");

            SetFieldByName(nominalRecord, "ACCOUNT_REF", nominalCode);
            SetFieldByName(nominalRecord, "NAME", name);

            var result = InvokeComMethod(nominalRecord, "Update");
            if (result != null && Convert.ToBoolean(result))
            {
                Console.WriteLine($"  Created: {nominalCode} - {name}");
                return true;
            }
            else
            {
                TryShowLastError();
                return false;
            }
        }
        catch (Exception ex)
        {
            Console.WriteLine($"  Error creating nominal: {ex.Message}");
            return false;
        }
    }

    /// <summary>
    /// Check if a nominal code exists
    /// </summary>
    public bool NominalExists(string nominalCode)
    {
        if (_workspace == null) return false;

        try
        {
            var nominalRecord = InvokeComMethod(_workspace, "CreateObject", "NominalRecord");
            if (nominalRecord == null) return false;

            InvokeComMethod(nominalRecord, "MoveFirst");
            while (true)
            {
                var eof = InvokeComMethod(nominalRecord, "IsEOF");
                if (eof != null && (bool)eof) break;

                var code = GetFieldValue(nominalRecord, "ACCOUNT_REF")?.ToString();
                if (code == nominalCode) return true;

                InvokeComMethod(nominalRecord, "MoveNext");
            }
            return false;
        }
        catch
        {
            return false;
        }
    }

    /// <summary>
    /// Create and post a Purchase Invoice
    /// </summary>
    /// <param name="supplierAccount">Supplier account reference</param>
    /// <param name="invoiceRef">Supplier's invoice reference</param>
    /// <param name="items">Line items</param>
    /// <param name="createSupplierIfMissing">If true, creates supplier if not found</param>
    /// <param name="supplierName">Name for new supplier</param>
    public string? CreatePurchaseInvoice(
        string supplierAccount,
        string invoiceRef,
        List<InvoiceLineItem> items,
        bool createSupplierIfMissing = false,
        string? supplierName = null)
    {
        if (_workspace == null)
            throw new InvalidOperationException("Not connected to Sage");

        Console.WriteLine("\n[POSTING PURCHASE INVOICE]");
        Console.WriteLine($"  Supplier: {supplierAccount}");
        Console.WriteLine($"  Invoice Ref: {invoiceRef}");
        Console.WriteLine($"  Items: {items.Count}");

        // Check if supplier exists
        if (!SupplierExists(supplierAccount))
        {
            Console.WriteLine($"  Supplier '{supplierAccount}' not found");

            if (createSupplierIfMissing)
            {
                var name = supplierName ?? supplierAccount;
                if (!CreateSupplier(supplierAccount, name))
                {
                    Console.WriteLine("  Failed to create supplier - aborting invoice");
                    return null;
                }
            }
            else
            {
                Console.WriteLine("  Set createSupplierIfMissing=true to auto-create");
                return null;
            }
        }
        else
        {
            Console.WriteLine($"  Supplier verified");
        }

        try
        {
            // Create PopPost object for Purchase Order Processing
            var popPost = InvokeComMethod(_workspace, "CreateObject", "PopPost");
            if (popPost == null)
            {
                Console.WriteLine("  ERROR: Could not create PopPost object");
                return null;
            }

            Console.WriteLine("  Created PopPost object");
            return PostWithPopPost(popPost, supplierAccount, invoiceRef, items);
        }
        catch (Exception ex)
        {
            Console.WriteLine($"  Exception: {ex.Message}");
            if (ex.InnerException != null)
                Console.WriteLine($"    Inner: {ex.InnerException.Message}");
        }

        return null;
    }

    /// <summary>
    /// Post purchase invoice using PopPost object
    /// </summary>
    private string? PostWithPopPost(object popPost, string supplierAccount, string invoiceRef, List<InvoiceLineItem> items)
    {
        try
        {
            var header = GetComProperty(popPost, "Header");
            if (header == null)
            {
                Console.WriteLine("  ERROR: Could not get Header");
                return null;
            }

            // Set header fields - ORDER_TYPE 2 = Purchase Invoice
            SetFieldByName(header, "ACCOUNT_REF", supplierAccount);
            SetFieldByName(header, "ORDER_DATE", DateTime.Now);
            SetFieldByName(header, "SUPP_ORDER_NUMBER", invoiceRef);
            SetFieldByName(header, "ORDER_TYPE", 2);

            Console.WriteLine($"  Header: Account={supplierAccount}, Date={DateTime.Now:d}, OrderType=2 (Invoice)");

            // Add line items
            var popItems = GetComProperty(popPost, "Items");
            if (popItems != null)
            {
                foreach (var lineItem in items)
                {
                    var item = InvokeComMethod(popItems, "Add");
                    if (item != null)
                    {
                        var netAmount = lineItem.UnitPrice * lineItem.Quantity;

                        SetFieldByName(item, "DESCRIPTION", lineItem.Description);
                        SetFieldByName(item, "NOMINAL_CODE", lineItem.NominalCode);

                        int taxCode = lineItem.TaxCode switch
                        {
                            "T0" => 0, "T1" => 1, "T2" => 2, "T5" => 5, "T9" => 9, _ => 1
                        };
                        SetFieldByName(item, "TAX_CODE", taxCode);
                        SetFieldByName(item, "NET_AMOUNT", netAmount);
                        SetFieldByName(item, "FULL_NET_AMOUNT", netAmount);
                        SetFieldByName(item, "UNIT_PRICE", lineItem.UnitPrice);
                        SetFieldByName(item, "QTY_ORDER", lineItem.Quantity);
                        SetFieldByName(item, "SERVICE_FLAG", 1); // Service item

                        Console.WriteLine($"  Item: {lineItem.Description} x{lineItem.Quantity} @ {lineItem.UnitPrice:C} = {netAmount:C}");
                    }
                }
            }

            // Call Update to commit
            Console.WriteLine("  Calling Update()...");
            var result = InvokeComMethod(popPost, "Update");
            Console.WriteLine($"  Update result: {result}");

            if (result != null && (bool)result)
            {
                var invoiceNum = GetFieldByName(header, "INVOICE_NUMBER");
                if (invoiceNum != null && !string.IsNullOrEmpty(invoiceNum.ToString()) && invoiceNum.ToString() != "0")
                {
                    Console.WriteLine($"  SUCCESS! Invoice Number: {invoiceNum}");
                    return invoiceNum.ToString();
                }
                else
                {
                    Console.WriteLine("  NOTE: Update() returned true - purchase invoice posted.");
                    return "POSTED";
                }
            }
            else
            {
                Console.WriteLine("  Update() returned false");
                TryShowLastError();
            }
        }
        catch (Exception ex)
        {
            Console.WriteLine($"  PopPost error: {ex.Message}");
        }

        return null;
    }

    /// <summary>
    /// Create and post a Sales Invoice
    /// </summary>
    /// <param name="customerAccount">Customer account reference</param>
    /// <param name="orderNumber">Order/reference number</param>
    /// <param name="items">Line items</param>
    /// <param name="createCustomerIfMissing">If true, creates customer if not found</param>
    /// <param name="customerName">Name for new customer (required if createCustomerIfMissing is true)</param>
    public string? CreateSalesInvoice(
        string customerAccount,
        string orderNumber,
        List<InvoiceLineItem> items,
        bool createCustomerIfMissing = false,
        string? customerName = null)
    {
        if (_workspace == null)
            throw new InvalidOperationException("Not connected to Sage");

        Console.WriteLine("\n[POSTING SALES INVOICE]");
        Console.WriteLine($"  Customer: {customerAccount}");
        Console.WriteLine($"  Order: {orderNumber}");
        Console.WriteLine($"  Items: {items.Count}");

        // Check if customer exists
        if (!CustomerExists(customerAccount))
        {
            Console.WriteLine($"  Customer '{customerAccount}' not found");

            if (createCustomerIfMissing)
            {
                var name = customerName ?? customerAccount;
                if (!CreateCustomer(customerAccount, name))
                {
                    Console.WriteLine("  Failed to create customer - aborting invoice");
                    return null;
                }
            }
            else
            {
                Console.WriteLine("  Set createCustomerIfMissing=true to auto-create");
                return null;
            }
        }
        else
        {
            Console.WriteLine($"  Customer verified");
        }

        try
        {
            // Create SopPost object for Sales Order Processing
            var sopPost = InvokeComMethod(_workspace, "CreateObject", "SopPost");
            if (sopPost == null)
            {
                Console.WriteLine("  ERROR: Could not create SopPost object");
                return null;
            }

            Console.WriteLine("  Created SopPost object");
            return PostWithSopPost(sopPost, customerAccount, orderNumber, items);
        }
        catch (Exception ex)
        {
            Console.WriteLine($"  Exception: {ex.Message}");
            if (ex.InnerException != null)
                Console.WriteLine($"    Inner: {ex.InnerException.Message}");
        }

        return null;
    }

    /// <summary>
    /// Post invoice using SopPost object
    /// Note: The Sage 50 SDK's SopPost.Update() returns true but may not create actual invoices.
    /// This appears to be a limitation requiring stock records or additional SDK configuration.
    /// </summary>
    private string? PostWithSopPost(object sopPost, string customerAccount, string orderNumber, List<InvoiceLineItem> items)
    {
        try
        {
            var header = GetComProperty(sopPost, "Header");
            if (header == null)
            {
                Console.WriteLine("  ERROR: Could not get Header");
                return null;
            }

            // Set header fields - ORDER_TYPE 2 = Sales Invoice
            SetFieldByName(header, "ACCOUNT_REF", customerAccount);
            SetFieldByName(header, "ORDER_DATE", DateTime.Now);
            SetFieldByName(header, "CUST_ORDER_NUMBER", orderNumber);
            SetFieldByName(header, "ORDER_TYPE", 2);

            Console.WriteLine($"  Header: Account={customerAccount}, Date={DateTime.Now:d}, OrderType=2");

            // Add line items
            var sopItems = GetComProperty(sopPost, "Items");
            if (sopItems != null)
            {
                foreach (var lineItem in items)
                {
                    var item = InvokeComMethod(sopItems, "Add");
                    if (item != null)
                    {
                        var netAmount = lineItem.UnitPrice * lineItem.Quantity;

                        SetFieldByName(item, "DESCRIPTION", lineItem.Description);
                        SetFieldByName(item, "NOMINAL_CODE", lineItem.NominalCode);

                        // TAX_CODE uses integer (T1=1, T0=0, etc.)
                        int taxCode = lineItem.TaxCode switch
                        {
                            "T0" => 0, "T1" => 1, "T2" => 2, "T5" => 5, "T9" => 9, _ => 1
                        };
                        SetFieldByName(item, "TAX_CODE", taxCode);
                        SetFieldByName(item, "NET_AMOUNT", netAmount);
                        SetFieldByName(item, "FULL_NET_AMOUNT", netAmount);
                        SetFieldByName(item, "UNIT_PRICE", lineItem.UnitPrice);
                        SetFieldByName(item, "QTY_ORDER", lineItem.Quantity);
                        SetFieldByName(item, "SERVICE_FLAG", 1); // Service item

                        Console.WriteLine($"  Item: {lineItem.Description} x{lineItem.Quantity} @ {lineItem.UnitPrice:C} = {netAmount:C}");
                    }
                }
            }

            // Call Update to commit
            Console.WriteLine("  Calling Update()...");
            var result = InvokeComMethod(sopPost, "Update");
            Console.WriteLine($"  Update result: {result}");

            if (result != null && (bool)result)
            {
                var invoiceNum = GetFieldByName(header, "INVOICE_NUMBER");
                if (invoiceNum != null && !string.IsNullOrEmpty(invoiceNum.ToString()) && invoiceNum.ToString() != "0")
                {
                    Console.WriteLine($"  SUCCESS! Invoice Number: {invoiceNum}");
                    return invoiceNum.ToString();
                }
                else
                {
                    Console.WriteLine("  NOTE: Update() returned true but no invoice number assigned.");
                    Console.WriteLine("        This may require stock records or additional SDK configuration.");
                    return "PENDING";
                }
            }
            else
            {
                Console.WriteLine("  Update() returned false");
                TryShowLastError();
            }
        }
        catch (Exception ex)
        {
            Console.WriteLine($"  SopPost error: {ex.Message}");
        }

        return null;
    }

    /// <summary>
    /// Try to show the last error from the workspace
    /// </summary>
    private void TryShowLastError()
    {
        try
        {
            if (_workspace == null) return;
            var lastError = GetComProperty(_workspace, "LastError");
            if (lastError != null)
            {
                var errNum = GetComProperty(lastError, "ErrorNumber");
                var errDesc = GetComProperty(lastError, "ErrorDescription");
                Console.WriteLine($"  Error {errNum}: {errDesc}");
            }
        }
        catch (Exception errEx)
        {
            Console.WriteLine($"  Could not get error info: {errEx.Message}");
        }
    }

    /// <summary>
    /// Set a COM property value
    /// </summary>
    private void SetComProperty(object obj, string propertyName, object value)
    {
        obj.GetType().InvokeMember(
            propertyName,
            BindingFlags.SetProperty,
            null,
            obj,
            new object[] { value });
    }

    /// <summary>
    /// Set a field value by name
    /// </summary>
    private void SetFieldByName(object fieldsOrItem, string fieldName, object value)
    {
        // Get the Fields collection
        object? fields;
        try
        {
            fields = GetComProperty(fieldsOrItem, "Fields");
            if (fields == null) fields = fieldsOrItem;
        }
        catch
        {
            fields = fieldsOrItem;
        }

        // Try direct access by field name first
        try
        {
            var field = fields!.GetType().InvokeMember(
                "Item",
                BindingFlags.GetProperty,
                null,
                fields,
                new object[] { fieldName });

            if (field != null)
            {
                SetComProperty(field, "Value", value);
                return;
            }
        }
        catch { }

        // Fallback: search by iterating through fields (1-based index)
        try
        {
            var fc = fields!.GetType().InvokeMember("Count", BindingFlags.GetProperty, null, fields, null);
            int fieldCount = fc is int i ? i : 0;

            for (int idx = 1; idx <= fieldCount; idx++)
            {
                try
                {
                    var field = fields.GetType().InvokeMember(
                        "Item",
                        BindingFlags.GetProperty,
                        null,
                        fields,
                        new object[] { idx });

                    if (field != null)
                    {
                        var name = GetComProperty(field, "Name")?.ToString();
                        if (name == fieldName)
                        {
                            SetComProperty(field, "Value", value);
                            return;
                        }
                    }
                }
                catch { }
            }
        }
        catch { }

        // Field not found - silent for non-critical fields
    }

    public void Dispose()
    {
        if (_disposed) return;

        try
        {
            if (_workspace != null)
            {
                _workspace.Disconnect();
                Console.WriteLine("\nDisconnected from Sage.");
            }
        }
        catch { }

        if (_sdoEngine != null)
        {
            Marshal.ReleaseComObject(_sdoEngine);
        }

        // Deactivate registration-free COM context if used
        if (_usingRegFreeCom)
        {
            RegFreeCom.Deactivate();
        }

        _disposed = true;
    }
}

/// <summary>
/// Represents a line item for an invoice
/// </summary>
public class InvoiceLineItem
{
    public string StockCode { get; set; } = "";
    public string Description { get; set; } = "";
    public decimal Quantity { get; set; }
    public decimal UnitPrice { get; set; }
    public string NominalCode { get; set; } = "4000";
    public string TaxCode { get; set; } = "T1";
}

/// <summary>
/// Represents a customer account from Sales Ledger
/// </summary>
public class CustomerAccount
{
    public string AccountRef { get; set; } = "";
    public string Name { get; set; } = "";
    public string Address1 { get; set; } = "";
    public string Address2 { get; set; } = "";
    public string Address3 { get; set; } = "";
    public string Address4 { get; set; } = "";
    public string Postcode { get; set; } = "";
    public string Country { get; set; } = "";
    public string Telephone { get; set; } = "";
    public string Fax { get; set; } = "";
    public string Email { get; set; } = "";
    public string ContactName { get; set; } = "";
    public string VatNumber { get; set; } = "";
    public string NominalCode { get; set; } = "";
    public decimal Balance { get; set; }
    public decimal CreditLimit { get; set; }
    public string Currency { get; set; } = "";
    public DateTime? DateOpened { get; set; }
}

/// <summary>
/// Represents a supplier account from Purchase Ledger
/// </summary>
public class SupplierAccount
{
    public string AccountRef { get; set; } = "";
    public string Name { get; set; } = "";
    public string Address1 { get; set; } = "";
    public string Address2 { get; set; } = "";
    public string Address3 { get; set; } = "";
    public string Address4 { get; set; } = "";
    public string Postcode { get; set; } = "";
    public string Country { get; set; } = "";
    public string Telephone { get; set; } = "";
    public string Fax { get; set; } = "";
    public string Email { get; set; } = "";
    public string ContactName { get; set; } = "";
    public string VatNumber { get; set; } = "";
    public string NominalCode { get; set; } = "";
    public decimal Balance { get; set; }
    public decimal CreditLimit { get; set; }
    public string Currency { get; set; } = "";
    public DateTime? DateOpened { get; set; }
}

/// <summary>
/// Represents a journal line for nominal ledger posting
/// </summary>
public class JournalLine
{
    public string NominalCode { get; set; } = "";
    public string Reference { get; set; } = "";
    public string Details { get; set; } = "";
    public decimal Debit { get; set; }
    public decimal Credit { get; set; }
    public string TaxCode { get; set; } = "T9"; // T9 = no VAT by default for journals
    public DateTime? Date { get; set; }
}
