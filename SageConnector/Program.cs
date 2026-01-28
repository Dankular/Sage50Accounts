using System.Runtime.InteropServices;
using System.Reflection;
using System.Diagnostics;

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
            // Test sales invoice posting (direct to ledger via TransactionPost)
            else if (args.Any(a => a.Equals("sinv", StringComparison.OrdinalIgnoreCase)))
            {
                Console.WriteLine("\n*** POSTING SALES INVOICE (via TransactionPost) ***");

                // Use TransactionPost which updates customer balance (not InvoicePost which just creates documents)
                var success = sage.PostTransactionSI(
                    customerAccount: "FALCONLO",
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
            else if (args.Any(a => a.Equals("pinv", StringComparison.OrdinalIgnoreCase)))
            {
                Console.WriteLine("\n*** POSTING PURCHASE INVOICE ***");

                // Use account that exists as both customer and supplier
                var supplierRef = "IMPHAM"; // Import Helpers gmbh - in both customer and supplier lists

                var success = sage.PostPurchaseInvoice(
                    supplierAccount: supplierRef,
                    invoiceRef: $"PI-{DateTime.Now:yyyyMMddHHmmss}",
                    netAmount: 50.00m,
                    taxAmount: 10.00m,
                    nominalCode: "5000",
                    details: "Test purchase invoice from SageConnector",
                    taxCode: "T1",
                    postToLedger: true);

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
            else
            {
                Console.WriteLine("\nCommands:");
                Console.WriteLine("  discover - Discover available SDK posting objects");
                Console.WriteLine("");
                Console.WriteLine("  Invoice Posting (updates ledger balances):");
                Console.WriteLine("  sinv     - Post sales invoice (updates customer balance)");
                Console.WriteLine("  pinv     - Post purchase invoice (updates supplier balance)");
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
/// Wrapper class for Sage 50 SDO COM connection
/// </summary>
class SageConnection : IDisposable
{
    private dynamic? _sdoEngine;
    private dynamic? _workspaces;
    private dynamic? _workspace;
    private bool _disposed;

    public SageConnection()
    {
        // Try different Sage COM objects
        string[] progIds = new[]
        {
            "SDOEngine.32",
            "SDOEngine",
            "SageDataObject50.SDOEngine",
            "Sage.Accounts.SDOEngine"
        };

        Type? sdoEngineType = null;
        string? usedProgId = null;

        foreach (var progId in progIds)
        {
            sdoEngineType = Type.GetTypeFromProgID(progId);
            if (sdoEngineType != null)
            {
                usedProgId = progId;
                Console.WriteLine($"Found COM object: {progId}");
                break;
            }
        }

        if (sdoEngineType == null)
        {
            throw new InvalidOperationException(
                "No Sage SDO Engine found. Is Sage 50 installed?");
        }

        _sdoEngine = Activator.CreateInstance(sdoEngineType);
        if (_sdoEngine == null)
        {
            throw new InvalidOperationException(
                "Failed to create SDOEngine instance.");
        }

        Console.WriteLine($"SDO Engine created successfully (using {usedProgId}).");

        // List available properties/methods
        try
        {
            Console.WriteLine("\nExploring SDO Engine members...");
            var type = _sdoEngine.GetType();

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
                                DiscoverFields(header, $"{objName} Header", 30);
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
                                    DiscoverFields(item, $"{objName} Item", 30);
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
    /// Post a Purchase Invoice using TransactionPost
    /// TYPE=4 is Purchase Invoice (PI) in Sage transaction types
    ///
    /// NOTE: TransactionPost has limitations with purchase invoices:
    /// - Accounts must exist in BOTH customer and supplier tables
    /// - Multi-currency accounts may cause "currency mismatch" errors
    /// - For reliable purchase posting, consider using JournalPost instead
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

            // Set TYPE first - this determines which account table to use
            // TYPE=4 is Purchase Invoice (PI) - uses Supplier accounts
            SetFieldByName(header, "TYPE", 4);

            // Now set the account reference (Sage knows to look in Suppliers table)
            SetFieldByName(header, "ACCOUNT_REF", supplierAccount);
            SetFieldByName(header, "DATE", date ?? DateTime.Today);
            SetFieldByName(header, "INV_REF", invoiceRef);
            SetFieldByName(header, "DETAILS", string.IsNullOrEmpty(details) ? "Purchase Invoice" : details);

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
                    SetFieldByName(item, "DETAILS", string.IsNullOrEmpty(details) ? "Purchase Invoice" : details);
                    SetFieldByName(item, "NET_AMOUNT", netAmount);
                    SetFieldByName(item, "TAX_AMOUNT", taxAmount);
                    SetFieldByName(item, "TAX_CODE", tc);
                }
            }

            Console.WriteLine($"  Header: Account={supplierAccount}, Type=4 (PI), Gross={grossAmount:C}");

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
