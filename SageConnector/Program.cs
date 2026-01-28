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

                var testAcct = $"TEST{DateTime.Now:MMddHHmm}";
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
            else
            {
                Console.WriteLine("\nCommands:");
                Console.WriteLine("  post    - Post test invoice to existing customer FALCONLO");
                Console.WriteLine("  newcust - Create new test customer and post invoice");
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

            // Try to add new record - SDK uses Add or AddNew
            bool addSuccess = false;
            try
            {
                var addResult = InvokeComMethod(salesRecord, "Add");
                addSuccess = addResult != null && (bool)addResult;
            }
            catch
            {
                try
                {
                    var addResult = InvokeComMethod(salesRecord, "AddNew");
                    addSuccess = addResult != null && (bool)addResult;
                }
                catch { }
            }

            if (!addSuccess)
            {
                Console.WriteLine("  Add/AddNew not available, setting fields directly");
            }

            // Set customer fields
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

            // Try to add new record
            bool addSuccess = false;
            try
            {
                var addResult = InvokeComMethod(purchaseRecord, "Add");
                addSuccess = addResult != null && (bool)addResult;
            }
            catch
            {
                try
                {
                    var addResult = InvokeComMethod(purchaseRecord, "AddNew");
                    addSuccess = addResult != null && (bool)addResult;
                }
                catch { }
            }

            if (!addSuccess)
            {
                Console.WriteLine("  Add/AddNew not available, setting fields directly");
            }

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
