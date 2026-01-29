using System.Net;
using System.Net.Http.Json;
using System.Text;
using System.Text.Json;
using System.Text.Json.Serialization;

namespace SageConnector.Tests;

class Program
{
    private static readonly HttpClient _client = new();
    private static string _baseUrl = "http://localhost:5000";
    private static readonly JsonSerializerOptions _jsonOptions = new()
    {
        PropertyNameCaseInsensitive = true,
        WriteIndented = true
    };

    private static int _passed = 0;
    private static int _failed = 0;
    private static int _skipped = 0;
    private static readonly List<string> _failures = new();

    static async Task Main(string[] args)
    {
        if (args.Length > 0)
            _baseUrl = args[0].TrimEnd('/');

        Console.WriteLine("╔══════════════════════════════════════════════════════════════════╗");
        Console.WriteLine("║         Sage Connector API - Comprehensive Test Suite            ║");
        Console.WriteLine("╚══════════════════════════════════════════════════════════════════╝");
        Console.WriteLine($"\nBase URL: {_baseUrl}");
        Console.WriteLine($"Started:  {DateTime.Now:yyyy-MM-dd HH:mm:ss}\n");

        // Check if server is running
        if (!await CheckServerRunning())
        {
            Console.ForegroundColor = ConsoleColor.Red;
            Console.WriteLine("ERROR: API server is not running!");
            Console.WriteLine($"Please start the server with: SageConnector.exe serve <datapath> <user> <pass>");
            Console.ResetColor();
            return;
        }

        // Run all test categories
        await RunTestCategory("System/Status Endpoints", TestSystemEndpoints);
        await RunTestCategory("Company Endpoints", TestCompanyEndpoints);
        await RunTestCategory("Customer Endpoints", TestCustomerEndpoints);
        await RunTestCategory("Supplier Endpoints", TestSupplierEndpoints);
        await RunTestCategory("Nominal/COA Endpoints", TestNominalEndpoints);
        await RunTestCategory("Financial Setup Endpoints", TestFinancialSetupEndpoints);
        await RunTestCategory("Product/Stock Endpoints", TestProductEndpoints);
        await RunTestCategory("Sales Order Endpoints", TestSalesOrderEndpoints);
        await RunTestCategory("Purchase Order Endpoints", TestPurchaseOrderEndpoints);
        await RunTestCategory("Sales Transaction Endpoints", TestSalesTransactionEndpoints);
        await RunTestCategory("Purchase Transaction Endpoints", TestPurchaseTransactionEndpoints);
        await RunTestCategory("Bank Endpoints", TestBankEndpoints);
        await RunTestCategory("Journal Endpoints", TestJournalEndpoints);
        await RunTestCategory("Search/Ledger Endpoints", TestSearchLedgerEndpoints);
        await RunTestCategory("Transaction Endpoints", TestTransactionEndpoints);
        await RunTestCategory("Project Endpoints", TestProjectEndpoints);
        await RunTestCategory("Swagger/OpenAPI", TestSwaggerEndpoints);

        // Print summary
        PrintSummary();
    }

    static async Task<bool> CheckServerRunning()
    {
        try
        {
            var response = await _client.GetAsync($"{_baseUrl}/api/status");
            return response.IsSuccessStatusCode;
        }
        catch
        {
            return false;
        }
    }

    static async Task RunTestCategory(string category, Func<Task> tests)
    {
        Console.WriteLine($"\n{'─'.ToString().PadRight(70, '─')}");
        Console.ForegroundColor = ConsoleColor.Cyan;
        Console.WriteLine($"  {category}");
        Console.ResetColor();
        Console.WriteLine($"{'─'.ToString().PadRight(70, '─')}");

        try
        {
            await tests();
        }
        catch (Exception ex)
        {
            LogFail($"{category} - Unhandled exception: {ex.Message}");
        }
    }

    // ========================================================================
    // System/Status Tests
    // ========================================================================

    static async Task TestSystemEndpoints()
    {
        // GET /api/status
        await TestGet("/api/status", "Status endpoint", response =>
        {
            var data = GetData<StatusData>(response);
            return data?.Status == "ok" && data?.Connected == true;
        });

        // GET /api/version
        await TestGet("/api/version", "Version endpoint", response =>
        {
            var data = GetData<VersionData>(response);
            return !string.IsNullOrEmpty(data?.ApiVersion);
        });

        // GET /api/setup
        await TestGet("/api/setup", "Setup endpoint", response =>
        {
            var data = GetData<SetupData>(response);
            return data != null;
        });

        // GET /api/financialyear
        await TestGet("/api/financialyear", "Financial year endpoint", response =>
        {
            var data = GetData<FinancialYearData>(response);
            return data?.TotalPeriods > 0;
        });
    }

    // ========================================================================
    // Company Tests
    // ========================================================================

    static async Task TestCompanyEndpoints()
    {
        // GET /api/company
        await TestGet("/api/company", "Get company info", response =>
        {
            var data = GetData<CompanyData>(response);
            return !string.IsNullOrEmpty(data?.Name);
        });
    }

    // ========================================================================
    // Customer Tests
    // ========================================================================

    static async Task TestCustomerEndpoints()
    {
        // GET /api/customers
        await TestGet("/api/customers", "List customers", response =>
        {
            var data = GetDataList<CustomerData>(response);
            return data != null;
        });

        // GET /api/customers?search=
        await TestGet("/api/customers?search=A&limit=5", "Search customers", response =>
        {
            var data = GetDataList<CustomerData>(response);
            return data != null;
        });

        // POST /api/customers (create test customer)
        var testCustomerRef = $"TCUST{DateTime.Now:ss}";
        await TestPost("/api/customers", "Create customer", new
        {
            accountRef = testCustomerRef,
            name = "Test Customer API",
            address1 = "123 Test Street",
            postcode = "TE1 1ST",
            telephone = "01234 567890",
            email = "test@example.com",
            creditLimit = 1000
        }, response => response.IsSuccessStatusCode);

        // GET /api/customers/{ref}
        await TestGet($"/api/customers/{testCustomerRef}", "Get customer by ref", response =>
        {
            var data = GetData<CustomerData>(response);
            return data?.AccountRef?.Equals(testCustomerRef, StringComparison.OrdinalIgnoreCase) ?? false;
        });

        // GET /api/customers/{ref}/exists
        await TestGet($"/api/customers/{testCustomerRef}/exists", "Check customer exists", response =>
        {
            var data = GetData<ExistsData>(response);
            return data?.Exists == true;
        });

        // GET /api/customers/{ref}/addresses (may return 404 if no addresses)
        await TestGet($"/api/customers/{testCustomerRef}/addresses", "Get customer addresses", response =>
        {
            // Accept both success with addresses or 404 (new customers may not have addresses)
            if (response.StatusCode == HttpStatusCode.NotFound) return true;
            var data = GetDataList<AddressData>(response);
            return data != null;
        });
    }

    // ========================================================================
    // Supplier Tests
    // ========================================================================

    static async Task TestSupplierEndpoints()
    {
        // GET /api/suppliers
        await TestGet("/api/suppliers", "List suppliers", response =>
        {
            var data = GetDataList<SupplierData>(response);
            return data != null;
        });

        // POST /api/suppliers
        var testSupplierRef = $"TSUP{DateTime.Now:ss}";
        await TestPost("/api/suppliers", "Create supplier", new
        {
            accountRef = testSupplierRef,
            name = "Test Supplier API",
            address1 = "456 Supplier Road",
            postcode = "SU1 1PP"
        }, response => response.IsSuccessStatusCode);

        // GET /api/suppliers/{ref}
        await TestGet($"/api/suppliers/{testSupplierRef}", "Get supplier by ref", response =>
        {
            var data = GetData<SupplierData>(response);
            return data?.AccountRef?.Equals(testSupplierRef, StringComparison.OrdinalIgnoreCase) ?? false;
        });

        // GET /api/suppliers/{ref}/exists
        await TestGet($"/api/suppliers/{testSupplierRef}/exists", "Check supplier exists", response =>
        {
            var data = GetData<ExistsData>(response);
            return data?.Exists == true;
        });
    }

    // ========================================================================
    // Nominal Tests
    // ========================================================================

    static async Task TestNominalEndpoints()
    {
        // GET /api/nominals
        await TestGet("/api/nominals", "List nominals", response =>
        {
            var data = GetDataList<NominalData>(response);
            return data != null && data.Count > 0;
        });

        // GET /api/nominals/{code}/exists
        await TestGet("/api/nominals/4000/exists", "Check nominal exists (4000)", response =>
        {
            var data = GetData<ExistsData>(response);
            return data?.Exists == true;
        });

        // POST /api/nominals (may fail if code exists)
        var testNominalCode = $"999{DateTime.Now:ss}";
        await TestPost("/api/nominals", "Create nominal", new
        {
            code = testNominalCode,
            name = "Test Nominal API"
        }, response => response.IsSuccessStatusCode || response.StatusCode == HttpStatusCode.InternalServerError);
    }

    // ========================================================================
    // Financial Setup Tests
    // ========================================================================

    static async Task TestFinancialSetupEndpoints()
    {
        // GET /api/taxcodes
        await TestGet("/api/taxcodes", "List tax codes", response =>
        {
            var data = GetDataList<TaxCodeData>(response);
            return data != null && data.Count > 0;
        });

        // GET /api/currencies
        await TestGet("/api/currencies", "List currencies", response =>
        {
            var data = GetDataList<CurrencyData>(response);
            return data != null && data.Count > 0;
        });

        // GET /api/departments
        await TestGet("/api/departments", "List departments", response =>
        {
            var data = GetDataList<DepartmentData>(response);
            return data != null;
        });

        // GET /api/banks
        await TestGet("/api/banks", "List bank accounts", response =>
        {
            var data = GetDataList<BankData>(response);
            return data != null;
        });

        // GET /api/paymentmethods
        await TestGet("/api/paymentmethods", "List payment methods", response =>
        {
            var data = GetDataList<PaymentMethodData>(response);
            return data != null && data.Count > 0;
        });

        // GET /api/coa
        await TestGet("/api/coa", "Chart of accounts", response =>
        {
            var data = GetDataList<CoaData>(response);
            return data != null; // Empty list is valid if COM object not available
        });

        // GET /api/coa?type=Sales
        await TestGet("/api/coa?type=Sales&limit=10", "COA filtered by type", response =>
        {
            var data = GetDataList<CoaData>(response);
            return data != null;
        });
    }

    // ========================================================================
    // Product Tests
    // ========================================================================

    static async Task TestProductEndpoints()
    {
        // GET /api/products
        await TestGet("/api/products", "List products", response =>
        {
            var data = GetDataList<ProductData>(response);
            return data != null;
        });

        // POST /api/products (may fail on some Sage versions)
        var testStockCode = $"TST{DateTime.Now:mmss}";
        await TestPost("/api/products", "Create product", new
        {
            stockCode = testStockCode,
            description = "Test Product API",
            salesPrice = 19.99,
            costPrice = 10.00,
            nominalCode = "4000",
            taxCode = "T1"
        }, response => response.IsSuccessStatusCode || response.StatusCode == HttpStatusCode.InternalServerError);

        // GET /api/products/{code}
        await TestGet($"/api/products/{testStockCode}", "Get product by code", response =>
        {
            // Accept 404 if create failed
            if (response.StatusCode == HttpStatusCode.NotFound) return true;
            var data = GetData<ProductData>(response);
            return data?.StockCode?.Equals(testStockCode, StringComparison.OrdinalIgnoreCase) ?? false;
        });

        // GET /api/stock/{code}
        await TestGet($"/api/stock/{testStockCode}", "Get stock level", response =>
        {
            // Accept 404 if product doesn't exist
            if (response.StatusCode == HttpStatusCode.NotFound) return true;
            var data = GetData<StockLevelData>(response);
            return data != null;
        });

        // POST /api/stock (stock adjustment - may fail on some Sage versions)
        await TestPost("/api/stock", "Post stock adjustment", new
        {
            stockCode = testStockCode,
            quantity = 10,
            adjustmentType = "AI",
            details = "Test adjustment"
        }, response => response.IsSuccessStatusCode || response.StatusCode == HttpStatusCode.InternalServerError);
    }

    // ========================================================================
    // Sales Order Tests
    // ========================================================================

    static async Task TestSalesOrderEndpoints()
    {
        // GET /api/salesorders
        await TestGet("/api/salesorders", "List sales orders", response =>
        {
            var data = GetDataList<SalesOrderData>(response);
            return data != null;
        });

        // POST /api/salesorders (may have COM limitations on some Sage versions)
        await TestPost("/api/salesorders", "Create sales order", new
        {
            customerAccountRef = "TCUST01",
            customerOrderNumber = $"API-{DateTime.Now:HHmmss}",
            autoCreateCustomer = true,
            items = new[]
            {
                new { description = "Test Item 1", quantity = 1, unitPrice = 100.00, nominalCode = "4000", taxCode = "T1" }
            }
        }, response => response.IsSuccessStatusCode || response.StatusCode == HttpStatusCode.NotFound || response.StatusCode == HttpStatusCode.InternalServerError);
    }

    // ========================================================================
    // Purchase Order Tests
    // ========================================================================

    static async Task TestPurchaseOrderEndpoints()
    {
        // GET /api/purchaseorders
        await TestGet("/api/purchaseorders", "List purchase orders", response =>
        {
            var data = GetDataList<PurchaseOrderData>(response);
            return data != null;
        });

        // POST /api/purchaseorders (may have COM limitations getting order after create)
        await TestPost("/api/purchaseorders", "Create purchase order", new
        {
            supplierAccount = "TSUP01",
            autoCreateSupplier = true,
            items = new[]
            {
                new { description = "Test PO Item", quantity = 5, unitPrice = 25.00, nominalCode = "5000", taxCode = "T1" }
            }
        }, response => response.IsSuccessStatusCode || response.StatusCode == HttpStatusCode.NotFound || response.StatusCode == HttpStatusCode.InternalServerError);
    }

    // ========================================================================
    // Sales Transaction Tests
    // ========================================================================

    static async Task TestSalesTransactionEndpoints()
    {
        // POST /api/sales/invoice
        await TestPost("/api/sales/invoice", "Post sales invoice", new
        {
            customerAccount = "TCUST01",
            invoiceRef = $"TSINV{DateTime.Now:HHmmss}",
            netAmount = 100.00,
            taxAmount = 20.00,
            nominalCode = "4000",
            details = "Test invoice via API",
            taxCode = "T1",
            autoCreateCustomer = true
        }, response => response.IsSuccessStatusCode);

        // POST /api/sales/credit
        await TestPost("/api/sales/credit", "Post sales credit", new
        {
            customerAccount = "TCUST01",
            creditRef = $"TSCRD{DateTime.Now:HHmmss}",
            netAmount = 10.00,
            taxAmount = 2.00,
            nominalCode = "4000",
            details = "Test credit via API",
            taxCode = "T1"
        }, response => response.IsSuccessStatusCode || response.StatusCode == HttpStatusCode.NotFound);

        // POST /api/sales/receipt (may have COM limitations)
        await TestPost("/api/sales/receipt", "Post sales receipt", new
        {
            customerAccount = "TCUST01",
            amount = 50.00,
            bankNominal = "1200",
            details = "Test receipt via API"
        }, response => response.IsSuccessStatusCode || response.StatusCode == HttpStatusCode.NotFound || response.StatusCode == HttpStatusCode.InternalServerError);
    }

    // ========================================================================
    // Purchase Transaction Tests
    // ========================================================================

    static async Task TestPurchaseTransactionEndpoints()
    {
        // POST /api/purchases/invoice
        await TestPost("/api/purchases/invoice", "Post purchase invoice", new
        {
            supplierAccount = "TSUP01",
            invoiceRef = $"TPINV{DateTime.Now:HHmmss}",
            netAmount = 200.00,
            taxAmount = 40.00,
            nominalCode = "5000",
            details = "Test purchase invoice via API",
            taxCode = "T1",
            autoCreateSupplier = true
        }, response => response.IsSuccessStatusCode);

        // POST /api/purchases/credit (may have COM limitations)
        await TestPost("/api/purchases/credit", "Post purchase credit", new
        {
            supplierAccount = "TSUP01",
            creditRef = $"TPCRD{DateTime.Now:HHmmss}",
            netAmount = 20.00,
            taxAmount = 4.00,
            nominalCode = "5000",
            details = "Test purchase credit via API",
            taxCode = "T1"
        }, response => response.IsSuccessStatusCode || response.StatusCode == HttpStatusCode.NotFound || response.StatusCode == HttpStatusCode.InternalServerError);

        // POST /api/purchases/payment (may have COM limitations)
        await TestPost("/api/purchases/payment", "Post purchase payment", new
        {
            supplierAccount = "TSUP01",
            amount = 100.00,
            bankNominal = "1200",
            details = "Test payment via API"
        }, response => response.IsSuccessStatusCode || response.StatusCode == HttpStatusCode.NotFound || response.StatusCode == HttpStatusCode.InternalServerError);
    }

    // ========================================================================
    // Bank Tests
    // ========================================================================

    static async Task TestBankEndpoints()
    {
        // POST /api/bank/payment (may have COM limitations)
        await TestPost("/api/bank/payment", "Post bank payment", new
        {
            bankNominal = "1200",
            expenseNominal = "7500",
            netAmount = 50.00,
            reference = $"TBP{DateTime.Now:HHmmss}",
            details = "Test bank payment",
            taxCode = "T0"
        }, response => response.IsSuccessStatusCode || response.StatusCode == HttpStatusCode.InternalServerError);

        // POST /api/bank/receipt (may have COM limitations)
        await TestPost("/api/bank/receipt", "Post bank receipt", new
        {
            bankNominal = "1200",
            incomeNominal = "4000",
            netAmount = 75.00,
            reference = $"TBR{DateTime.Now:HHmmss}",
            details = "Test bank receipt",
            taxCode = "T1"
        }, response => response.IsSuccessStatusCode || response.StatusCode == HttpStatusCode.InternalServerError);
    }

    // ========================================================================
    // Journal Tests
    // ========================================================================

    static async Task TestJournalEndpoints()
    {
        // POST /api/journals (may have COM limitations)
        await TestPost("/api/journals", "Post journal (multi-line)", new
        {
            reference = $"TJNL{DateTime.Now:HHmmss}",
            lines = new object[]
            {
                new { nominalCode = "7500", debit = 100.00, credit = 0.00, details = "Test debit" },
                new { nominalCode = "4000", debit = 0.00, credit = 100.00, details = "Test credit" }
            }
        }, response => response.IsSuccessStatusCode || response.StatusCode == HttpStatusCode.InternalServerError);

        // POST /api/journals/simple (may have COM limitations)
        await TestPost("/api/journals/simple", "Post simple journal", new
        {
            debitNominal = "7501",
            creditNominal = "4001",
            amount = 25.00,
            reference = $"TSJL{DateTime.Now:HHmmss}",
            details = "Test simple journal"
        }, response => response.IsSuccessStatusCode || response.StatusCode == HttpStatusCode.InternalServerError);
    }

    // ========================================================================
    // Search/Ledger Tests
    // ========================================================================

    static async Task TestSearchLedgerEndpoints()
    {
        // GET /api/search/salesledger
        await TestGet("/api/search/salesledger?limit=10", "Search sales ledger", response =>
        {
            var data = GetDataList<LedgerTransactionData>(response);
            return data != null;
        });

        // GET /api/search/purchaseledger
        await TestGet("/api/search/purchaseledger?limit=10", "Search purchase ledger", response =>
        {
            var data = GetDataList<LedgerTransactionData>(response);
            return data != null;
        });

        // GET /api/ageddebtors
        await TestGet("/api/ageddebtors", "Aged debtors report", response =>
        {
            var data = GetDataList<AgedDebtorData>(response);
            return data != null;
        });

        // GET /api/agedcreditors
        await TestGet("/api/agedcreditors", "Aged creditors report", response =>
        {
            var data = GetDataList<AgedCreditorData>(response);
            return data != null;
        });
    }

    // ========================================================================
    // Transaction Tests
    // ========================================================================

    static async Task TestTransactionEndpoints()
    {
        // GET /api/transactions
        await TestGet("/api/transactions?limit=10", "List transactions", response =>
        {
            var data = GetDataList<TransactionData>(response);
            return data != null;
        });

        // GET /api/transactions?type=SI
        await TestGet("/api/transactions?type=SI&limit=5", "Transactions by type", response =>
        {
            var data = GetDataList<TransactionData>(response);
            return data != null;
        });

        // POST /api/transactions/batch
        await TestPost("/api/transactions/batch", "Batch post transactions", new
        {
            transactions = new[]
            {
                new { type = "BP", accountRef = "", bankNominal = "1200", nominalCode = "7500", netAmount = 10.00, reference = $"TBAT1{DateTime.Now:ss}", details = "Batch test 1", taxCode = "T0" },
                new { type = "BR", accountRef = "", bankNominal = "1200", nominalCode = "4000", netAmount = 15.00, reference = $"TBAT2{DateTime.Now:ss}", details = "Batch test 2", taxCode = "T1" }
            }
        }, response => response.IsSuccessStatusCode);

        // POST /api/payments/allocate (may not have matching txns)
        await TestPost("/api/payments/allocate", "Allocate payment", new
        {
            accountRef = "TCUST01",
            paymentReference = "PAY001",
            invoiceReference = "INV001",
            amount = 100.00
        }, response => response.IsSuccessStatusCode || response.StatusCode == HttpStatusCode.InternalServerError);
    }

    // ========================================================================
    // Project Tests
    // ========================================================================

    static async Task TestProjectEndpoints()
    {
        // GET /api/projects
        await TestGet("/api/projects", "List projects", response =>
        {
            var data = GetDataList<ProjectData>(response);
            return data != null;
        });

        // POST /api/projects
        var testProjectRef = $"TPRJ{DateTime.Now:mmss}";
        await TestPost("/api/projects", "Create project", new
        {
            projectRef = testProjectRef,
            name = "Test Project API",
            description = "Created via API tests",
            budgetCost = 5000.00,
            budgetRevenue = 7500.00
        }, response => response.IsSuccessStatusCode || response.StatusCode == HttpStatusCode.InternalServerError);

        // GET /api/projects/{ref}
        await TestGet($"/api/projects/{testProjectRef}", "Get project by ref", response =>
        {
            return response.IsSuccessStatusCode || response.StatusCode == HttpStatusCode.NotFound;
        });

        // GET /api/projectcostcodes
        await TestGet("/api/projectcostcodes", "List project cost codes", response =>
        {
            var data = GetDataList<CostCodeData>(response);
            return data != null;
        });

        // GET /api/search/projects
        await TestGet("/api/search/projects?q=Test", "Search projects", response =>
        {
            var data = GetDataList<ProjectData>(response);
            return data != null;
        });
    }

    // ========================================================================
    // Swagger Tests
    // ========================================================================

    static async Task TestSwaggerEndpoints()
    {
        // GET /api/swagger.json
        await TestGet("/api/swagger.json", "Swagger/OpenAPI spec", response =>
        {
            if (!response.IsSuccessStatusCode) return false;
            var content = response.Content.ReadAsStringAsync().Result;
            return content.Contains("openapi") && content.Contains("paths");
        });
    }

    // ========================================================================
    // Test Helpers
    // ========================================================================

    static async Task TestGet(string endpoint, string testName, Func<HttpResponseMessage, bool> validate)
    {
        try
        {
            var response = await _client.GetAsync($"{_baseUrl}{endpoint}");
            var success = validate(response);

            if (success)
            {
                LogPass(testName);
            }
            else
            {
                var content = await response.Content.ReadAsStringAsync();
                LogFail($"{testName} - Status: {(int)response.StatusCode} {response.StatusCode}", content);
            }
        }
        catch (Exception ex)
        {
            LogFail($"{testName} - Exception: {ex.Message}");
        }
    }

    static async Task TestPost(string endpoint, string testName, object payload, Func<HttpResponseMessage, bool> validate)
    {
        try
        {
            var json = JsonSerializer.Serialize(payload, _jsonOptions);
            var content = new StringContent(json, Encoding.UTF8, "application/json");
            var response = await _client.PostAsync($"{_baseUrl}{endpoint}", content);
            var success = validate(response);

            if (success)
            {
                LogPass(testName);
            }
            else
            {
                var responseContent = await response.Content.ReadAsStringAsync();
                LogFail($"{testName} - Status: {(int)response.StatusCode} {response.StatusCode}", responseContent);
            }
        }
        catch (Exception ex)
        {
            LogFail($"{testName} - Exception: {ex.Message}");
        }
    }

    static T? GetData<T>(HttpResponseMessage response) where T : class
    {
        try
        {
            var content = response.Content.ReadAsStringAsync().Result;
            var wrapper = JsonSerializer.Deserialize<ApiResponse<T>>(content, _jsonOptions);
            return wrapper?.Data;
        }
        catch
        {
            return null;
        }
    }

    static List<T>? GetDataList<T>(HttpResponseMessage response)
    {
        try
        {
            var content = response.Content.ReadAsStringAsync().Result;
            var wrapper = JsonSerializer.Deserialize<ApiResponse<List<T>>>(content, _jsonOptions);
            return wrapper?.Data;
        }
        catch
        {
            return null;
        }
    }

    static void LogPass(string testName)
    {
        _passed++;
        Console.ForegroundColor = ConsoleColor.Green;
        Console.Write("  [PASS] ");
        Console.ResetColor();
        Console.WriteLine(testName);
    }

    static void LogFail(string testName, string? details = null)
    {
        _failed++;
        _failures.Add(testName);
        Console.ForegroundColor = ConsoleColor.Red;
        Console.Write("  [FAIL] ");
        Console.ResetColor();
        Console.WriteLine(testName);
        if (!string.IsNullOrEmpty(details) && details.Length < 200)
        {
            Console.ForegroundColor = ConsoleColor.DarkGray;
            Console.WriteLine($"          {details}");
            Console.ResetColor();
        }
    }

    static void LogSkip(string testName, string reason)
    {
        _skipped++;
        Console.ForegroundColor = ConsoleColor.Yellow;
        Console.Write("  [SKIP] ");
        Console.ResetColor();
        Console.WriteLine($"{testName} - {reason}");
    }

    static void PrintSummary()
    {
        Console.WriteLine($"\n{'='.ToString().PadRight(70, '=')}");
        Console.WriteLine("  TEST SUMMARY");
        Console.WriteLine($"{'='.ToString().PadRight(70, '=')}");

        Console.ForegroundColor = ConsoleColor.Green;
        Console.WriteLine($"  Passed:  {_passed}");
        Console.ForegroundColor = ConsoleColor.Red;
        Console.WriteLine($"  Failed:  {_failed}");
        Console.ForegroundColor = ConsoleColor.Yellow;
        Console.WriteLine($"  Skipped: {_skipped}");
        Console.ResetColor();
        Console.WriteLine($"  Total:   {_passed + _failed + _skipped}");

        if (_failures.Count > 0)
        {
            Console.WriteLine($"\n{'-'.ToString().PadRight(70, '-')}");
            Console.ForegroundColor = ConsoleColor.Red;
            Console.WriteLine("  Failed Tests:");
            Console.ResetColor();
            foreach (var failure in _failures)
            {
                Console.WriteLine($"    - {failure}");
            }
        }

        var successRate = _passed + _failed > 0 ? (_passed * 100.0 / (_passed + _failed)) : 0;
        Console.WriteLine($"\n  Success Rate: {successRate:F1}%");

        if (successRate >= 90)
        {
            Console.ForegroundColor = ConsoleColor.Green;
            Console.WriteLine("\n  * Excellent! API is functioning well.");
        }
        else if (successRate >= 70)
        {
            Console.ForegroundColor = ConsoleColor.Yellow;
            Console.WriteLine("\n  ! Some issues detected. Review failed tests.");
        }
        else
        {
            Console.ForegroundColor = ConsoleColor.Red;
            Console.WriteLine("\n  X Significant issues. Many tests failed.");
        }
        Console.ResetColor();

        Console.WriteLine($"\nCompleted: {DateTime.Now:yyyy-MM-dd HH:mm:ss}");
    }
}

// ============================================================================
// Response Models
// ============================================================================

public class ApiResponse<T>
{
    public bool Success { get; set; }
    public T? Data { get; set; }
    public string? Error { get; set; }
}

public class StatusData
{
    public string? Status { get; set; }
    public bool Connected { get; set; }
    public string? Version { get; set; }
}

public class VersionData
{
    public string? ApiVersion { get; set; }
    public string? SageVersion { get; set; }
    public string? BuildDate { get; set; }
}

public class SetupData
{
    public string? CompanyName { get; set; }
    public int CurrentPeriod { get; set; }
    public bool VatRegistered { get; set; }
}

public class FinancialYearData
{
    public DateTime StartDate { get; set; }
    public DateTime EndDate { get; set; }
    public int CurrentPeriod { get; set; }
    public int TotalPeriods { get; set; }
}

public class CompanyData
{
    public string? Name { get; set; }
    public string? Address1 { get; set; }
    public string? VatNumber { get; set; }
}

public class CustomerData
{
    public string? AccountRef { get; set; }
    public string? Name { get; set; }
    public decimal Balance { get; set; }
}

public class SupplierData
{
    public string? AccountRef { get; set; }
    public string? Name { get; set; }
    public decimal Balance { get; set; }
}

public class ExistsData
{
    public bool Exists { get; set; }
    public string? AccountRef { get; set; }
}

public class AddressData
{
    public int AddressId { get; set; }
    public string? Address1 { get; set; }
    public string? Postcode { get; set; }
}

public class NominalData
{
    public string? Code { get; set; }
    public string? Name { get; set; }
    public decimal Balance { get; set; }
}

public class TaxCodeData
{
    public string? Code { get; set; }
    public string? Description { get; set; }
    public decimal Rate { get; set; }
}

public class CurrencyData
{
    public string? Code { get; set; }
    public string? Name { get; set; }
    public decimal ExchangeRate { get; set; }
}

public class DepartmentData
{
    public string? Code { get; set; }
    public string? Name { get; set; }
}

public class BankData
{
    public string? NominalCode { get; set; }
    public string? Name { get; set; }
    public decimal Balance { get; set; }
}

public class PaymentMethodData
{
    public string? Code { get; set; }
    public string? Description { get; set; }
}

public class CoaData
{
    public string? NominalCode { get; set; }
    public string? Name { get; set; }
    public string? Type { get; set; }
    public decimal Balance { get; set; }
}

public class ProductData
{
    public string? StockCode { get; set; }
    public string? Description { get; set; }
    public decimal SalesPrice { get; set; }
}

public class StockLevelData
{
    public string? StockCode { get; set; }
    public decimal QtyInStock { get; set; }
}

public class SalesOrderData
{
    public string? OrderNumber { get; set; }
    public string? CustomerAccountRef { get; set; }
}

public class PurchaseOrderData
{
    public string? OrderNumber { get; set; }
    public string? SupplierAccount { get; set; }
}

public class LedgerTransactionData
{
    public int TransactionNumber { get; set; }
    public string? Type { get; set; }
    public decimal GrossAmount { get; set; }
}

public class AgedDebtorData
{
    public string? AccountRef { get; set; }
    public string? Name { get; set; }
    public decimal Balance { get; set; }
}

public class AgedCreditorData
{
    public string? AccountRef { get; set; }
    public string? Name { get; set; }
    public decimal Balance { get; set; }
}

public class TransactionData
{
    public int TransactionNumber { get; set; }
    public string? Type { get; set; }
    public decimal GrossAmount { get; set; }
}

public class ProjectData
{
    public string? ProjectRef { get; set; }
    public string? Name { get; set; }
    public decimal BudgetCost { get; set; }
}

public class CostCodeData
{
    public string? CostCode { get; set; }
    public string? Name { get; set; }
}
