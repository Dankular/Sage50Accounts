using System.Net;
using System.Text;
using System.Text.Json;
using System.Text.RegularExpressions;

namespace SageConnector;

/// <summary>
/// HTTP API server for Sage 50 Accounts using HttpListener
/// </summary>
public class ApiServer : IDisposable
{
    private readonly HttpListener _listener;
    private readonly SageConnection _sage;
    private readonly string _prefix;
    private readonly JsonSerializerOptions _jsonOptions;
    private bool _disposed;

    public ApiServer(string dataPath, string username, string password, int port = 5000)
    {
        _prefix = $"http://+:{port}/";
        _listener = new HttpListener();
        _listener.Prefixes.Add(_prefix);

        _jsonOptions = new JsonSerializerOptions
        {
            PropertyNamingPolicy = JsonNamingPolicy.CamelCase,
            WriteIndented = true
        };

        // Initialize Sage connection
        _sage = new SageConnection();
        _sage.Connect(dataPath, username, password);

        Console.WriteLine($"Connected to Sage 50 at {dataPath}");
    }

    public async Task StartAsync(CancellationToken ct)
    {
        _listener.Start();
        Console.WriteLine($"API server listening at {_prefix}api/");
        Console.WriteLine("Press Ctrl+C to stop");
        Console.WriteLine();
        Console.WriteLine("Endpoints:");
        Console.WriteLine("  GET  /api/swagger.json     - OpenAPI specification");
        Console.WriteLine("  GET  /api/company          - Company information");
        Console.WriteLine("  GET  /api/customers        - List customers");
        Console.WriteLine("  POST /api/customers        - Create customer");
        Console.WriteLine("  GET  /api/suppliers        - List suppliers");
        Console.WriteLine("  POST /api/sales/invoice    - Post sales invoice");
        Console.WriteLine("  POST /api/purchases/invoice - Post purchase invoice");
        Console.WriteLine("  ... and more (see swagger.json)");
        Console.WriteLine();

        while (!ct.IsCancellationRequested)
        {
            try
            {
                var context = await _listener.GetContextAsync().WaitAsync(ct);
                _ = HandleRequestAsync(context);
            }
            catch (OperationCanceledException)
            {
                break;
            }
            catch (Exception ex)
            {
                Console.WriteLine($"Error accepting request: {ex.Message}");
            }
        }

        _listener.Stop();
        Console.WriteLine("API server stopped");
    }

    private async Task HandleRequestAsync(HttpListenerContext context)
    {
        var request = context.Request;
        var response = context.Response;
        var path = request.Url?.AbsolutePath ?? "";
        var method = request.HttpMethod;

        // Add CORS headers
        response.AddHeader("Access-Control-Allow-Origin", "*");
        response.AddHeader("Access-Control-Allow-Methods", "GET, POST, OPTIONS");
        response.AddHeader("Access-Control-Allow-Headers", "Content-Type");

        // Handle preflight
        if (method == "OPTIONS")
        {
            response.StatusCode = 204;
            response.Close();
            return;
        }

        Console.WriteLine($"{DateTime.Now:HH:mm:ss} {method} {path}");

        try
        {
            var (statusCode, data) = await RouteAsync(method, path, request);

            response.ContentType = "application/json; charset=utf-8";
            response.StatusCode = statusCode;

            var json = JsonSerializer.Serialize(data, _jsonOptions);
            var buffer = Encoding.UTF8.GetBytes(json);
            response.ContentLength64 = buffer.Length;
            await response.OutputStream.WriteAsync(buffer);
        }
        catch (Exception ex)
        {
            Console.WriteLine($"  Error: {ex.Message}");
            response.StatusCode = 500;
            response.ContentType = "application/json; charset=utf-8";
            var error = JsonSerializer.Serialize(ApiResponse.Error(ex.Message), _jsonOptions);
            var buffer = Encoding.UTF8.GetBytes(error);
            response.ContentLength64 = buffer.Length;
            await response.OutputStream.WriteAsync(buffer);
        }
        finally
        {
            response.Close();
        }
    }

    private async Task<(int StatusCode, object Data)> RouteAsync(string method, string path, HttpListenerRequest request)
    {
        // Swagger
        if (path == "/api/swagger.json" && method == "GET")
            return (200, SwaggerGenerator.Generate(_prefix.TrimEnd('/')));

        // Company
        if (path == "/api/company" && method == "GET")
            return GetCompanyInfo();

        // Customers
        if (path == "/api/customers")
        {
            if (method == "GET") return GetCustomers(request);
            if (method == "POST") return await CreateCustomerAsync(request);
        }

        var customerMatch = Regex.Match(path, @"^/api/customers/([^/]+)$");
        if (customerMatch.Success)
        {
            var accountRef = Uri.UnescapeDataString(customerMatch.Groups[1].Value);
            if (method == "GET") return GetCustomer(accountRef);
        }

        var customerExistsMatch = Regex.Match(path, @"^/api/customers/([^/]+)/exists$");
        if (customerExistsMatch.Success && method == "GET")
        {
            var accountRef = Uri.UnescapeDataString(customerExistsMatch.Groups[1].Value);
            return CustomerExists(accountRef);
        }

        // Suppliers
        if (path == "/api/suppliers")
        {
            if (method == "GET") return GetSuppliers(request);
            if (method == "POST") return await CreateSupplierAsync(request);
        }

        var supplierMatch = Regex.Match(path, @"^/api/suppliers/([^/]+)$");
        if (supplierMatch.Success)
        {
            var accountRef = Uri.UnescapeDataString(supplierMatch.Groups[1].Value);
            if (method == "GET") return GetSupplier(accountRef);
        }

        var supplierExistsMatch = Regex.Match(path, @"^/api/suppliers/([^/]+)/exists$");
        if (supplierExistsMatch.Success && method == "GET")
        {
            var accountRef = Uri.UnescapeDataString(supplierExistsMatch.Groups[1].Value);
            return SupplierExists(accountRef);
        }

        // Sales
        if (path == "/api/sales/invoice" && method == "POST")
            return await PostSalesInvoiceAsync(request);

        if (path == "/api/sales/credit" && method == "POST")
            return await PostSalesCreditAsync(request);

        // Purchases
        if (path == "/api/purchases/invoice" && method == "POST")
            return await PostPurchaseInvoiceAsync(request);

        if (path == "/api/purchases/credit" && method == "POST")
            return await PostPurchaseCreditAsync(request);

        // Nominals
        if (path == "/api/nominals")
        {
            if (method == "GET") return GetNominals(request);
            if (method == "POST") return await CreateNominalAsync(request);
        }

        var nominalExistsMatch = Regex.Match(path, @"^/api/nominals/([^/]+)/exists$");
        if (nominalExistsMatch.Success && method == "GET")
        {
            var code = Uri.UnescapeDataString(nominalExistsMatch.Groups[1].Value);
            return NominalExists(code);
        }

        // Bank
        if (path == "/api/bank/payment" && method == "POST")
            return await PostBankPaymentAsync(request);

        if (path == "/api/bank/receipt" && method == "POST")
            return await PostBankReceiptAsync(request);

        // Journals
        if (path == "/api/journals" && method == "POST")
            return await PostJournalAsync(request);

        if (path == "/api/journals/simple" && method == "POST")
            return await PostSimpleJournalAsync(request);

        // Products
        if (path == "/api/products")
        {
            if (method == "GET") return GetProducts(request);
            if (method == "POST") return await CreateProductAsync(request);
        }

        var productMatch = Regex.Match(path, @"^/api/products/([^/]+)$");
        if (productMatch.Success && method == "GET")
        {
            var stockCode = Uri.UnescapeDataString(productMatch.Groups[1].Value);
            return GetProduct(stockCode);
        }

        // Stock
        if (path == "/api/stock" && method == "POST")
            return await PostStockAdjustmentAsync(request);

        var stockMatch = Regex.Match(path, @"^/api/stock/([^/]+)$");
        if (stockMatch.Success && method == "GET")
        {
            var stockCode = Uri.UnescapeDataString(stockMatch.Groups[1].Value);
            return GetStockLevel(stockCode);
        }

        // Sales Orders
        if (path == "/api/salesorders")
        {
            if (method == "GET") return GetSalesOrders(request);
            if (method == "POST") return await CreateSalesOrderAsync(request);
        }

        var salesOrderMatch = Regex.Match(path, @"^/api/salesorders/([^/]+)$");
        if (salesOrderMatch.Success)
        {
            var orderNumber = Uri.UnescapeDataString(salesOrderMatch.Groups[1].Value);
            if (method == "GET") return GetSalesOrder(orderNumber);
            if (method == "PATCH") return await UpdateSalesOrderAsync(orderNumber, request);
            if (method == "DELETE") return DeleteSalesOrder(orderNumber);
        }

        var salesOrderCompleteMatch = Regex.Match(path, @"^/api/salesorders/([^/]+)/complete$");
        if (salesOrderCompleteMatch.Success && method == "POST")
        {
            var orderNumber = Uri.UnescapeDataString(salesOrderCompleteMatch.Groups[1].Value);
            return CompleteSalesOrder(orderNumber);
        }

        // Purchase Orders
        if (path == "/api/purchaseorders")
        {
            if (method == "GET") return GetPurchaseOrders(request);
            if (method == "POST") return await CreatePurchaseOrderAsync(request);
        }

        var purchaseOrderMatch = Regex.Match(path, @"^/api/purchaseorders/([^/]+)$");
        if (purchaseOrderMatch.Success)
        {
            var orderNumber = Uri.UnescapeDataString(purchaseOrderMatch.Groups[1].Value);
            if (method == "GET") return GetPurchaseOrder(orderNumber);
            if (method == "PATCH") return await UpdatePurchaseOrderAsync(orderNumber, request);
            if (method == "DELETE") return DeletePurchaseOrder(orderNumber);
        }

        // System/Status endpoints
        if (path == "/api/status" && method == "GET")
            return GetStatus();

        if (path == "/api/version" && method == "GET")
            return GetVersion();

        if (path == "/api/setup" && method == "GET")
            return GetSetup();

        if (path == "/api/financialyear" && method == "GET")
            return GetFinancialYear();

        // Financial Setup endpoints
        if (path == "/api/taxcodes" && method == "GET")
            return GetTaxCodes();

        if (path == "/api/currencies" && method == "GET")
            return GetCurrencies();

        if (path == "/api/departments" && method == "GET")
            return GetDepartments();

        if (path == "/api/banks" && method == "GET")
            return GetBanks();

        if (path == "/api/paymentmethods" && method == "GET")
            return GetPaymentMethods();

        if (path == "/api/coa" && method == "GET")
            return GetChartOfAccounts(request);

        // Search/Ledger endpoints
        if (path == "/api/search/salesledger" && method == "GET")
            return SearchSalesLedger(request);

        if (path == "/api/search/purchaseledger" && method == "GET")
            return SearchPurchaseLedger(request);

        if (path == "/api/ageddebtors" && method == "GET")
            return GetAgedDebtors(request);

        if (path == "/api/agedcreditors" && method == "GET")
            return GetAgedCreditors(request);

        var customerAddressesMatch = Regex.Match(path, @"^/api/customers/([^/]+)/addresses$");
        if (customerAddressesMatch.Success && method == "GET")
        {
            var accountRef = Uri.UnescapeDataString(customerAddressesMatch.Groups[1].Value);
            return GetCustomerAddresses(accountRef);
        }

        // Transaction/Payment endpoints
        if (path == "/api/transactions" && method == "GET")
            return GetTransactions(request);

        if (path == "/api/transactions/batch" && method == "POST")
            return await PostTransactionBatchAsync(request);

        if (path == "/api/payments/allocate" && method == "POST")
            return await AllocatePaymentAsync(request);

        if (path == "/api/sales/receipt" && method == "POST")
            return await PostSalesReceiptAsync(request);

        if (path == "/api/purchases/payment" && method == "POST")
            return await PostPurchasePaymentAsync(request);

        // Project endpoints
        if (path == "/api/projects")
        {
            if (method == "GET") return GetProjects(request);
            if (method == "POST") return await CreateProjectAsync(request);
        }

        var projectMatch = Regex.Match(path, @"^/api/projects/([^/]+)$");
        if (projectMatch.Success && method == "GET")
        {
            var projectRef = Uri.UnescapeDataString(projectMatch.Groups[1].Value);
            return GetProject(projectRef);
        }

        if (path == "/api/projectcostcodes" && method == "GET")
            return GetProjectCostCodes(request);

        if (path == "/api/search/projects" && method == "GET")
            return SearchProjects(request);

        // Not found
        return (404, ApiResponse.Error($"Endpoint not found: {method} {path}"));
    }

    // =========================================================================
    // Company Handlers
    // =========================================================================

    private (int, object) GetCompanyInfo()
    {
        var info = _sage.GetCompanyInfo();
        if (info == null)
            return (500, ApiResponse.Error("Failed to get company info"));

        return (200, ApiResponse.Ok(info));
    }

    // =========================================================================
    // Customer Handlers
    // =========================================================================

    private (int, object) GetCustomers(HttpListenerRequest request)
    {
        var search = request.QueryString["search"] ?? "";
        var limitStr = request.QueryString["limit"] ?? "50";
        int.TryParse(limitStr, out var limit);
        if (limit <= 0) limit = 50;

        var customers = _sage.FindCustomers(search, limit);
        var response = customers.Select(c => new CustomerResponse(
            c.AccountRef, c.Name, c.Address1, c.Address2, c.Address3,
            c.Postcode, c.Telephone, c.Email, c.ContactName, c.Balance, c.CreditLimit
        )).ToList();

        return (200, ApiResponse.Ok(response));
    }

    private (int, object) GetCustomer(string accountRef)
    {
        var customer = _sage.GetCustomer(accountRef);
        if (customer == null)
            return (404, ApiResponse.Error($"Customer not found: {accountRef}"));

        var response = new CustomerResponse(
            customer.AccountRef, customer.Name, customer.Address1, customer.Address2, customer.Address3,
            customer.Postcode, customer.Telephone, customer.Email, customer.ContactName, customer.Balance, customer.CreditLimit
        );

        return (200, ApiResponse.Ok(response));
    }

    private (int, object) CustomerExists(string accountRef)
    {
        var exists = _sage.CustomerExists(accountRef);
        return (200, ApiResponse.Ok(new ExistsResponse(exists, accountRef)));
    }

    private async Task<(int, object)> CreateCustomerAsync(HttpListenerRequest request)
    {
        var body = await ReadBodyAsync<CreateCustomerRequest>(request);
        if (body == null)
            return (400, ApiResponse.Error("Invalid request body"));

        if (string.IsNullOrWhiteSpace(body.AccountRef))
            return (400, ApiResponse.Error("accountRef is required"));

        if (body.AccountRef.Length > 8)
            return (400, ApiResponse.Error("accountRef must be max 8 characters"));

        var customer = new CustomerAccount
        {
            AccountRef = body.AccountRef.ToUpperInvariant(),
            Name = body.Name,
            Address1 = body.Address1 ?? "",
            Address2 = body.Address2 ?? "",
            Address3 = body.Address3 ?? "",
            Postcode = body.Postcode ?? "",
            Telephone = body.Telephone ?? "",
            Email = body.Email ?? "",
            ContactName = body.ContactName ?? "",
            CreditLimit = body.CreditLimit ?? 0
        };

        var success = _sage.CreateCustomerEx(customer);
        if (!success)
            return (500, ApiResponse.Error("Failed to create customer"));

        // Return the created customer
        var created = _sage.GetCustomer(customer.AccountRef);
        if (created == null)
            return (200, ApiResponse.Ok(new TransactionResponse(true, customer.AccountRef, "Customer created")));

        var response = new CustomerResponse(
            created.AccountRef, created.Name, created.Address1, created.Address2, created.Address3,
            created.Postcode, created.Telephone, created.Email, created.ContactName, created.Balance, created.CreditLimit
        );

        return (201, ApiResponse.Ok(response));
    }

    // =========================================================================
    // Supplier Handlers
    // =========================================================================

    private (int, object) GetSuppliers(HttpListenerRequest request)
    {
        var search = request.QueryString["search"] ?? "";
        var limitStr = request.QueryString["limit"] ?? "50";
        int.TryParse(limitStr, out var limit);
        if (limit <= 0) limit = 50;

        var suppliers = _sage.FindSuppliers(search, limit);
        var response = suppliers.Select(s => new SupplierResponse(
            s.AccountRef, s.Name, s.Address1, s.Address2, s.Address3,
            s.Postcode, s.Telephone, s.Email, s.ContactName, s.Balance, s.CreditLimit
        )).ToList();

        return (200, ApiResponse.Ok(response));
    }

    private (int, object) GetSupplier(string accountRef)
    {
        var supplier = _sage.GetSupplier(accountRef);
        if (supplier == null)
            return (404, ApiResponse.Error($"Supplier not found: {accountRef}"));

        var response = new SupplierResponse(
            supplier.AccountRef, supplier.Name, supplier.Address1, supplier.Address2, supplier.Address3,
            supplier.Postcode, supplier.Telephone, supplier.Email, supplier.ContactName, supplier.Balance, supplier.CreditLimit
        );

        return (200, ApiResponse.Ok(response));
    }

    private (int, object) SupplierExists(string accountRef)
    {
        var exists = _sage.SupplierExists(accountRef);
        return (200, ApiResponse.Ok(new ExistsResponse(exists, accountRef)));
    }

    private async Task<(int, object)> CreateSupplierAsync(HttpListenerRequest request)
    {
        var body = await ReadBodyAsync<CreateSupplierRequest>(request);
        if (body == null)
            return (400, ApiResponse.Error("Invalid request body"));

        if (string.IsNullOrWhiteSpace(body.AccountRef))
            return (400, ApiResponse.Error("accountRef is required"));

        if (body.AccountRef.Length > 8)
            return (400, ApiResponse.Error("accountRef must be max 8 characters"));

        var supplier = new SupplierAccount
        {
            AccountRef = body.AccountRef.ToUpperInvariant(),
            Name = body.Name,
            Address1 = body.Address1 ?? "",
            Address2 = body.Address2 ?? "",
            Address3 = body.Address3 ?? "",
            Postcode = body.Postcode ?? "",
            Telephone = body.Telephone ?? "",
            Email = body.Email ?? "",
            ContactName = body.ContactName ?? "",
            CreditLimit = body.CreditLimit ?? 0
        };

        var success = _sage.CreateSupplierEx(supplier);
        if (!success)
            return (500, ApiResponse.Error("Failed to create supplier"));

        // Return the created supplier
        var created = _sage.GetSupplier(supplier.AccountRef);
        if (created == null)
            return (200, ApiResponse.Ok(new TransactionResponse(true, supplier.AccountRef, "Supplier created")));

        var response = new SupplierResponse(
            created.AccountRef, created.Name, created.Address1, created.Address2, created.Address3,
            created.Postcode, created.Telephone, created.Email, created.ContactName, created.Balance, created.CreditLimit
        );

        return (201, ApiResponse.Ok(response));
    }

    // =========================================================================
    // Sales Handlers
    // =========================================================================

    private async Task<(int, object)> PostSalesInvoiceAsync(HttpListenerRequest request)
    {
        var body = await ReadBodyAsync<PostSalesInvoiceRequest>(request);
        if (body == null)
            return (400, ApiResponse.Error("Invalid request body"));

        if (string.IsNullOrWhiteSpace(body.CustomerAccount))
            return (400, ApiResponse.Error("customerAccount is required"));

        // Check customer exists or auto-create
        if (!_sage.CustomerExists(body.CustomerAccount))
        {
            if (body.AutoCreateCustomer)
            {
                var cust = new CustomerAccount
                {
                    AccountRef = body.CustomerAccount.ToUpperInvariant(),
                    Name = $"Auto-created {body.CustomerAccount}",
                    Address1 = "Auto-created via API"
                };
                if (!_sage.CreateCustomerEx(cust))
                    return (400, ApiResponse.Error($"Failed to auto-create customer: {body.CustomerAccount}"));
            }
            else
            {
                return (404, ApiResponse.Error($"Customer not found: {body.CustomerAccount}. Set autoCreateCustomer=true to create."));
            }
        }

        var invoiceRef = body.InvoiceRef ?? $"SI-{DateTime.Now:yyyyMMddHHmmss}";

        var success = _sage.PostTransactionSI(
            customerAccount: body.CustomerAccount,
            invoiceRef: invoiceRef,
            netAmount: body.NetAmount,
            taxAmount: body.TaxAmount,
            nominalCode: body.NominalCode,
            details: body.Details ?? "Posted via API",
            taxCode: body.TaxCode,
            postToLedger: true
        );

        if (!success)
            return (500, ApiResponse.Error("Failed to post sales invoice"));

        return (200, ApiResponse.Ok(new TransactionResponse(true, invoiceRef, "Sales invoice posted")));
    }

    private async Task<(int, object)> PostSalesCreditAsync(HttpListenerRequest request)
    {
        var body = await ReadBodyAsync<PostSalesCreditRequest>(request);
        if (body == null)
            return (400, ApiResponse.Error("Invalid request body"));

        if (string.IsNullOrWhiteSpace(body.CustomerAccount))
            return (400, ApiResponse.Error("customerAccount is required"));

        if (!_sage.CustomerExists(body.CustomerAccount))
            return (404, ApiResponse.Error($"Customer not found: {body.CustomerAccount}"));

        var creditRef = body.CreditRef ?? $"SC-{DateTime.Now:yyyyMMddHHmmss}";

        var success = _sage.PostSalesDebit(
            customerAccount: body.CustomerAccount,
            reference: creditRef,
            netAmount: body.NetAmount,
            taxAmount: body.TaxAmount,
            nominalCode: body.NominalCode,
            details: body.Details ?? "Credit note via API",
            taxCode: body.TaxCode
        );

        if (!success)
            return (500, ApiResponse.Error("Failed to post sales credit"));

        return (200, ApiResponse.Ok(new TransactionResponse(true, creditRef, "Sales credit posted")));
    }

    // =========================================================================
    // Purchase Handlers
    // =========================================================================

    private async Task<(int, object)> PostPurchaseInvoiceAsync(HttpListenerRequest request)
    {
        var body = await ReadBodyAsync<PostPurchaseInvoiceRequest>(request);
        if (body == null)
            return (400, ApiResponse.Error("Invalid request body"));

        if (string.IsNullOrWhiteSpace(body.SupplierAccount))
            return (400, ApiResponse.Error("supplierAccount is required"));

        // Check supplier exists or auto-create
        if (!_sage.SupplierExists(body.SupplierAccount))
        {
            if (body.AutoCreateSupplier)
            {
                var sup = new SupplierAccount
                {
                    AccountRef = body.SupplierAccount.ToUpperInvariant(),
                    Name = $"Auto-created {body.SupplierAccount}",
                    Address1 = "Auto-created via API"
                };
                if (!_sage.CreateSupplierEx(sup))
                    return (400, ApiResponse.Error($"Failed to auto-create supplier: {body.SupplierAccount}"));
            }
            else
            {
                return (404, ApiResponse.Error($"Supplier not found: {body.SupplierAccount}. Set autoCreateSupplier=true to create."));
            }
        }

        var invoiceRef = body.InvoiceRef ?? $"PI-{DateTime.Now:yyyyMMddHHmmss}";

        var success = _sage.PostTransactionPI(
            supplierAccount: body.SupplierAccount,
            invoiceRef: invoiceRef,
            netAmount: body.NetAmount,
            taxAmount: body.TaxAmount,
            nominalCode: body.NominalCode,
            details: body.Details ?? "Posted via API",
            taxCode: body.TaxCode
        );

        if (!success)
            return (500, ApiResponse.Error("Failed to post purchase invoice"));

        return (200, ApiResponse.Ok(new TransactionResponse(true, invoiceRef, "Purchase invoice posted")));
    }

    private async Task<(int, object)> PostPurchaseCreditAsync(HttpListenerRequest request)
    {
        var body = await ReadBodyAsync<PostPurchaseCreditRequest>(request);
        if (body == null)
            return (400, ApiResponse.Error("Invalid request body"));

        if (string.IsNullOrWhiteSpace(body.SupplierAccount))
            return (400, ApiResponse.Error("supplierAccount is required"));

        if (!_sage.SupplierExists(body.SupplierAccount))
            return (404, ApiResponse.Error($"Supplier not found: {body.SupplierAccount}"));

        var creditRef = body.CreditRef ?? $"PC-{DateTime.Now:yyyyMMddHHmmss}";

        var success = _sage.PostPurchaseCredit(
            supplierAccount: body.SupplierAccount,
            reference: creditRef,
            netAmount: body.NetAmount,
            taxAmount: body.TaxAmount,
            nominalCode: body.NominalCode,
            details: body.Details ?? "Credit note via API",
            taxCode: body.TaxCode
        );

        if (!success)
            return (500, ApiResponse.Error("Failed to post purchase credit"));

        return (200, ApiResponse.Ok(new TransactionResponse(true, creditRef, "Purchase credit posted")));
    }

    // =========================================================================
    // Nominal Handlers
    // =========================================================================

    private (int, object) GetNominals(HttpListenerRequest request)
    {
        var limitStr = request.QueryString["limit"] ?? "100";
        int.TryParse(limitStr, out var limit);
        if (limit <= 0) limit = 100;

        var nominals = _sage.GetNominalCodes(limit);
        var response = nominals.Select(n => new NominalCodeResponse(n.Code, n.Name, n.Balance)).ToList();

        return (200, ApiResponse.Ok(response));
    }

    private (int, object) NominalExists(string code)
    {
        var exists = _sage.NominalExists(code);
        return (200, ApiResponse.Ok(new ExistsResponse(exists, code)));
    }

    private async Task<(int, object)> CreateNominalAsync(HttpListenerRequest request)
    {
        var body = await ReadBodyAsync<CreateNominalRequest>(request);
        if (body == null)
            return (400, ApiResponse.Error("Invalid request body"));

        if (string.IsNullOrWhiteSpace(body.Code) || string.IsNullOrWhiteSpace(body.Name))
            return (400, ApiResponse.Error("code and name are required"));

        var success = _sage.CreateNominal(body.Code, body.Name);
        if (!success)
            return (500, ApiResponse.Error("Failed to create nominal code"));

        return (201, ApiResponse.Ok(new TransactionResponse(true, body.Code, "Nominal code created")));
    }

    // =========================================================================
    // Bank Handlers
    // =========================================================================

    private async Task<(int, object)> PostBankPaymentAsync(HttpListenerRequest request)
    {
        var body = await ReadBodyAsync<PostBankPaymentRequest>(request);
        if (body == null)
            return (400, ApiResponse.Error("Invalid request body"));

        var reference = body.Reference ?? $"BP-{DateTime.Now:yyyyMMddHHmmss}";

        var success = _sage.PostBankPayment(
            bankNominal: body.BankNominal,
            expenseNominal: body.ExpenseNominal,
            netAmount: body.NetAmount,
            reference: reference,
            details: body.Details ?? "Bank payment via API",
            taxCode: body.TaxCode
        );

        if (!success)
            return (500, ApiResponse.Error("Failed to post bank payment"));

        return (200, ApiResponse.Ok(new TransactionResponse(true, reference, "Bank payment posted")));
    }

    private async Task<(int, object)> PostBankReceiptAsync(HttpListenerRequest request)
    {
        var body = await ReadBodyAsync<PostBankReceiptRequest>(request);
        if (body == null)
            return (400, ApiResponse.Error("Invalid request body"));

        var reference = body.Reference ?? $"BR-{DateTime.Now:yyyyMMddHHmmss}";

        var success = _sage.PostBankReceipt(
            bankNominal: body.BankNominal,
            incomeNominal: body.IncomeNominal,
            netAmount: body.NetAmount,
            reference: reference,
            details: body.Details ?? "Bank receipt via API",
            taxCode: body.TaxCode
        );

        if (!success)
            return (500, ApiResponse.Error("Failed to post bank receipt"));

        return (200, ApiResponse.Ok(new TransactionResponse(true, reference, "Bank receipt posted")));
    }

    // =========================================================================
    // Journal Handlers
    // =========================================================================

    private async Task<(int, object)> PostJournalAsync(HttpListenerRequest request)
    {
        var body = await ReadBodyAsync<PostJournalRequest>(request);
        if (body == null)
            return (400, ApiResponse.Error("Invalid request body"));

        if (body.Lines == null || body.Lines.Count < 2)
            return (400, ApiResponse.Error("Journal must have at least 2 lines"));

        var journalLines = body.Lines.Select(l => new JournalLine
        {
            NominalCode = l.NominalCode,
            Debit = l.Debit,
            Credit = l.Credit,
            Details = l.Details ?? ""
        }).ToList();

        var reference = body.Reference ?? $"JNL-{DateTime.Now:yyyyMMddHHmmss}";

        var success = _sage.PostJournal(journalLines, reference, body.Date);
        if (!success)
            return (500, ApiResponse.Error("Failed to post journal"));

        return (200, ApiResponse.Ok(new TransactionResponse(true, reference, "Journal posted")));
    }

    private async Task<(int, object)> PostSimpleJournalAsync(HttpListenerRequest request)
    {
        var body = await ReadBodyAsync<PostSimpleJournalRequest>(request);
        if (body == null)
            return (400, ApiResponse.Error("Invalid request body"));

        if (string.IsNullOrWhiteSpace(body.DebitNominal) || string.IsNullOrWhiteSpace(body.CreditNominal))
            return (400, ApiResponse.Error("debitNominal and creditNominal are required"));

        var reference = body.Reference ?? $"JNL-{DateTime.Now:yyyyMMddHHmmss}";

        var success = _sage.PostSimpleJournal(
            debitNominal: body.DebitNominal,
            creditNominal: body.CreditNominal,
            amount: body.Amount,
            reference: reference,
            details: body.Details ?? "Journal via API",
            date: body.Date ?? DateTime.Today
        );

        if (!success)
            return (500, ApiResponse.Error("Failed to post journal"));

        return (200, ApiResponse.Ok(new TransactionResponse(true, reference, "Journal posted")));
    }

    // =========================================================================
    // Product Handlers
    // =========================================================================

    private (int, object) GetProducts(HttpListenerRequest request)
    {
        var search = request.QueryString["search"] ?? "";
        var limitStr = request.QueryString["limit"] ?? "50";
        int.TryParse(limitStr, out var limit);
        if (limit <= 0) limit = 50;

        var products = _sage.GetProducts(search, limit);
        return (200, ApiResponse.Ok(products));
    }

    private (int, object) GetProduct(string stockCode)
    {
        var product = _sage.GetProduct(stockCode);
        if (product == null)
            return (404, ApiResponse.Error($"Product not found: {stockCode}"));

        return (200, ApiResponse.Ok(product));
    }

    private async Task<(int, object)> CreateProductAsync(HttpListenerRequest request)
    {
        var body = await ReadBodyAsync<CreateProductRequest>(request);
        if (body == null)
            return (400, ApiResponse.Error("Invalid request body"));

        if (string.IsNullOrWhiteSpace(body.StockCode))
            return (400, ApiResponse.Error("stockCode is required"));

        var result = _sage.CreateProduct(body);
        if (result == null)
            return (500, ApiResponse.Error("Failed to create product"));

        return (201, ApiResponse.Ok(result));
    }

    // =========================================================================
    // Stock Handlers
    // =========================================================================

    private (int, object) GetStockLevel(string stockCode)
    {
        var stock = _sage.GetStockLevel(stockCode);
        if (stock == null)
            return (404, ApiResponse.Error($"Stock not found: {stockCode}"));

        return (200, ApiResponse.Ok(stock));
    }

    private async Task<(int, object)> PostStockAdjustmentAsync(HttpListenerRequest request)
    {
        var body = await ReadBodyAsync<StockAdjustmentRequest>(request);
        if (body == null)
            return (400, ApiResponse.Error("Invalid request body"));

        if (string.IsNullOrWhiteSpace(body.StockCode))
            return (400, ApiResponse.Error("stockCode is required"));

        var reference = body.Reference ?? $"ADJ-{DateTime.Now:yyyyMMddHHmmss}";

        var success = _sage.PostStockAdjustment(
            body.StockCode,
            body.Quantity,
            body.AdjustmentType,
            reference,
            body.Details ?? "Stock adjustment via API",
            body.CostPrice
        );

        if (!success)
            return (500, ApiResponse.Error("Failed to post stock adjustment"));

        return (200, ApiResponse.Ok(new TransactionResponse(true, reference, "Stock adjustment posted")));
    }

    // =========================================================================
    // Sales Order Handlers
    // =========================================================================

    private (int, object) GetSalesOrders(HttpListenerRequest request)
    {
        var search = request.QueryString["search"] ?? "";
        var limitStr = request.QueryString["limit"] ?? "50";
        int.TryParse(limitStr, out var limit);
        if (limit <= 0) limit = 50;

        var orders = _sage.GetSalesOrders(search, limit);
        return (200, ApiResponse.Ok(orders));
    }

    private (int, object) GetSalesOrder(string orderNumber)
    {
        var order = _sage.GetSalesOrder(orderNumber);
        if (order == null)
            return (404, ApiResponse.Error($"Sales order not found: {orderNumber}"));

        return (200, ApiResponse.Ok(order));
    }

    private async Task<(int, object)> CreateSalesOrderAsync(HttpListenerRequest request)
    {
        var body = await ReadBodyAsync<CreateSalesOrderRequest>(request);
        if (body == null)
            return (400, ApiResponse.Error("Invalid request body"));

        if (string.IsNullOrWhiteSpace(body.CustomerAccountRef))
            return (400, ApiResponse.Error("customerAccountRef is required"));

        // Check customer exists or auto-create
        if (!_sage.CustomerExists(body.CustomerAccountRef))
        {
            if (body.AutoCreateCustomer)
            {
                var cust = new CustomerAccount
                {
                    AccountRef = body.CustomerAccountRef.ToUpperInvariant(),
                    Name = $"Auto-created {body.CustomerAccountRef}",
                    Address1 = "Auto-created via API"
                };
                if (!_sage.CreateCustomerEx(cust))
                    return (400, ApiResponse.Error($"Failed to auto-create customer: {body.CustomerAccountRef}"));
            }
            else
            {
                return (404, ApiResponse.Error($"Customer not found: {body.CustomerAccountRef}. Set autoCreateCustomer=true to create."));
            }
        }

        var result = _sage.CreateSalesOrder(body);
        if (result == null)
            return (500, ApiResponse.Error("Failed to create sales order"));

        return (201, ApiResponse.Ok(result));
    }

    private async Task<(int, object)> UpdateSalesOrderAsync(string orderNumber, HttpListenerRequest request)
    {
        var body = await ReadBodyAsync<UpdateSalesOrderRequest>(request);
        if (body == null)
            return (400, ApiResponse.Error("Invalid request body"));

        var success = _sage.UpdateSalesOrder(orderNumber, body);
        if (!success)
            return (500, ApiResponse.Error($"Failed to update sales order: {orderNumber}"));

        return (200, ApiResponse.Ok(new TransactionResponse(true, orderNumber, "Sales order updated")));
    }

    private (int, object) DeleteSalesOrder(string orderNumber)
    {
        var success = _sage.DeleteSalesOrder(orderNumber);
        if (!success)
            return (500, ApiResponse.Error($"Failed to delete sales order: {orderNumber}"));

        return (200, ApiResponse.Ok(new TransactionResponse(true, orderNumber, "Sales order deleted")));
    }

    private (int, object) CompleteSalesOrder(string orderNumber)
    {
        var success = _sage.CompleteSalesOrder(orderNumber);
        if (!success)
            return (500, ApiResponse.Error($"Failed to complete sales order: {orderNumber}"));

        return (200, ApiResponse.Ok(new TransactionResponse(true, orderNumber, "Sales order completed")));
    }

    // =========================================================================
    // Purchase Order Handlers
    // =========================================================================

    private (int, object) GetPurchaseOrders(HttpListenerRequest request)
    {
        var search = request.QueryString["search"] ?? "";
        var limitStr = request.QueryString["limit"] ?? "50";
        int.TryParse(limitStr, out var limit);
        if (limit <= 0) limit = 50;

        var orders = _sage.GetPurchaseOrders(search, limit);
        return (200, ApiResponse.Ok(orders));
    }

    private (int, object) GetPurchaseOrder(string orderNumber)
    {
        var order = _sage.GetPurchaseOrder(orderNumber);
        if (order == null)
            return (404, ApiResponse.Error($"Purchase order not found: {orderNumber}"));

        return (200, ApiResponse.Ok(order));
    }

    private async Task<(int, object)> CreatePurchaseOrderAsync(HttpListenerRequest request)
    {
        var body = await ReadBodyAsync<CreatePurchaseOrderRequest>(request);
        if (body == null)
            return (400, ApiResponse.Error("Invalid request body"));

        if (string.IsNullOrWhiteSpace(body.SupplierAccount))
            return (400, ApiResponse.Error("supplierAccount is required"));

        // Check supplier exists or auto-create
        if (!_sage.SupplierExists(body.SupplierAccount))
        {
            if (body.AutoCreateSupplier)
            {
                var sup = new SupplierAccount
                {
                    AccountRef = body.SupplierAccount.ToUpperInvariant(),
                    Name = $"Auto-created {body.SupplierAccount}",
                    Address1 = "Auto-created via API"
                };
                if (!_sage.CreateSupplierEx(sup))
                    return (400, ApiResponse.Error($"Failed to auto-create supplier: {body.SupplierAccount}"));
            }
            else
            {
                return (404, ApiResponse.Error($"Supplier not found: {body.SupplierAccount}. Set autoCreateSupplier=true to create."));
            }
        }

        var result = _sage.CreatePurchaseOrder(body);
        if (result == null)
            return (500, ApiResponse.Error("Failed to create purchase order"));

        return (201, ApiResponse.Ok(result));
    }

    private async Task<(int, object)> UpdatePurchaseOrderAsync(string orderNumber, HttpListenerRequest request)
    {
        var body = await ReadBodyAsync<UpdatePurchaseOrderRequest>(request);
        if (body == null)
            return (400, ApiResponse.Error("Invalid request body"));

        var success = _sage.UpdatePurchaseOrder(orderNumber, body);
        if (!success)
            return (500, ApiResponse.Error($"Failed to update purchase order: {orderNumber}"));

        return (200, ApiResponse.Ok(new TransactionResponse(true, orderNumber, "Purchase order updated")));
    }

    private (int, object) DeletePurchaseOrder(string orderNumber)
    {
        var success = _sage.DeletePurchaseOrder(orderNumber);
        if (!success)
            return (500, ApiResponse.Error($"Failed to delete purchase order: {orderNumber}"));

        return (200, ApiResponse.Ok(new TransactionResponse(true, orderNumber, "Purchase order deleted")));
    }

    // =========================================================================
    // System/Status Handlers
    // =========================================================================

    private (int, object) GetStatus()
    {
        var status = new StatusResponse(
            Status: "ok",
            Connected: _sage.IsConnected,
            Timestamp: DateTime.UtcNow,
            Version: GetType().Assembly.GetName().Version?.ToString() ?? "1.0.0"
        );
        return (200, ApiResponse.Ok(status));
    }

    private (int, object) GetVersion()
    {
        var version = new VersionResponse(
            ApiVersion: GetType().Assembly.GetName().Version?.ToString() ?? "1.0.0",
            SageVersion: _sage.GetSageVersion(),
            BuildDate: File.GetLastWriteTime(GetType().Assembly.Location).ToString("yyyy-MM-dd")
        );
        return (200, ApiResponse.Ok(version));
    }

    private (int, object) GetSetup()
    {
        var setup = _sage.GetSetup();
        if (setup == null)
            return (500, ApiResponse.Error("Failed to get setup information"));

        return (200, ApiResponse.Ok(setup));
    }

    private (int, object) GetFinancialYear()
    {
        var financialYear = _sage.GetFinancialYear();
        if (financialYear == null)
            return (500, ApiResponse.Error("Failed to get financial year information"));

        return (200, ApiResponse.Ok(financialYear));
    }

    // =========================================================================
    // Financial Setup Handlers
    // =========================================================================

    private (int, object) GetTaxCodes()
    {
        var taxCodes = _sage.GetTaxCodes();
        return (200, ApiResponse.Ok(taxCodes));
    }

    private (int, object) GetCurrencies()
    {
        var currencies = _sage.GetCurrencies();
        return (200, ApiResponse.Ok(currencies));
    }

    private (int, object) GetDepartments()
    {
        var departments = _sage.GetDepartments();
        return (200, ApiResponse.Ok(departments));
    }

    private (int, object) GetBanks()
    {
        var banks = _sage.GetBanks();
        return (200, ApiResponse.Ok(banks));
    }

    private (int, object) GetPaymentMethods()
    {
        var methods = _sage.GetPaymentMethods();
        return (200, ApiResponse.Ok(methods));
    }

    private (int, object) GetChartOfAccounts(HttpListenerRequest request)
    {
        var typeFilter = request.QueryString["type"] ?? "";
        var limitStr = request.QueryString["limit"] ?? "500";
        int.TryParse(limitStr, out var limit);
        if (limit <= 0) limit = 500;

        var accounts = _sage.GetChartOfAccounts(typeFilter, limit);
        return (200, ApiResponse.Ok(accounts));
    }

    // =========================================================================
    // Search/Ledger Handlers
    // =========================================================================

    private (int, object) SearchSalesLedger(HttpListenerRequest request)
    {
        var accountRef = request.QueryString["account"] ?? "";
        var fromDate = request.QueryString["from"];
        var toDate = request.QueryString["to"];
        var limitStr = request.QueryString["limit"] ?? "100";
        int.TryParse(limitStr, out var limit);
        if (limit <= 0) limit = 100;

        DateTime? from = null, to = null;
        if (!string.IsNullOrEmpty(fromDate) && DateTime.TryParse(fromDate, out DateTime fromParsed))
            from = fromParsed;
        if (!string.IsNullOrEmpty(toDate) && DateTime.TryParse(toDate, out DateTime toParsed))
            to = toParsed;

        var transactions = _sage.SearchSalesLedger(accountRef, from, to, limit);
        return (200, ApiResponse.Ok(transactions));
    }

    private (int, object) SearchPurchaseLedger(HttpListenerRequest request)
    {
        var accountRef = request.QueryString["account"] ?? "";
        var fromDate = request.QueryString["from"];
        var toDate = request.QueryString["to"];
        var limitStr = request.QueryString["limit"] ?? "100";
        int.TryParse(limitStr, out var limit);
        if (limit <= 0) limit = 100;

        DateTime? from = null, to = null;
        if (!string.IsNullOrEmpty(fromDate) && DateTime.TryParse(fromDate, out DateTime fromParsed))
            from = fromParsed;
        if (!string.IsNullOrEmpty(toDate) && DateTime.TryParse(toDate, out DateTime toParsed))
            to = toParsed;

        var transactions = _sage.SearchPurchaseLedger(accountRef, from, to, limit);
        return (200, ApiResponse.Ok(transactions));
    }

    private (int, object) GetAgedDebtors(HttpListenerRequest request)
    {
        var limitStr = request.QueryString["limit"] ?? "100";
        int.TryParse(limitStr, out var limit);
        if (limit <= 0) limit = 100;

        var debtors = _sage.GetAgedDebtors(limit);
        return (200, ApiResponse.Ok(debtors));
    }

    private (int, object) GetAgedCreditors(HttpListenerRequest request)
    {
        var limitStr = request.QueryString["limit"] ?? "100";
        int.TryParse(limitStr, out var limit);
        if (limit <= 0) limit = 100;

        var creditors = _sage.GetAgedCreditors(limit);
        return (200, ApiResponse.Ok(creditors));
    }

    private (int, object) GetCustomerAddresses(string accountRef)
    {
        var addresses = _sage.GetCustomerAddresses(accountRef);
        if (addresses == null || addresses.Count == 0)
            return (404, ApiResponse.Error($"No addresses found for customer: {accountRef}"));

        return (200, ApiResponse.Ok(addresses));
    }

    // =========================================================================
    // Transaction/Payment Handlers
    // =========================================================================

    private (int, object) GetTransactions(HttpListenerRequest request)
    {
        var type = request.QueryString["type"] ?? "";
        var fromDate = request.QueryString["from"];
        var toDate = request.QueryString["to"];
        var limitStr = request.QueryString["limit"] ?? "100";
        int.TryParse(limitStr, out var limit);
        if (limit <= 0) limit = 100;

        DateTime? from = null, to = null;
        if (!string.IsNullOrEmpty(fromDate) && DateTime.TryParse(fromDate, out DateTime fromParsed))
            from = fromParsed;
        if (!string.IsNullOrEmpty(toDate) && DateTime.TryParse(toDate, out DateTime toParsed))
            to = toParsed;

        var transactions = _sage.GetTransactions(type, from, to, limit);
        return (200, ApiResponse.Ok(transactions));
    }

    private async Task<(int, object)> PostTransactionBatchAsync(HttpListenerRequest request)
    {
        var body = await ReadBodyAsync<PostTransactionBatchRequest>(request);
        if (body == null || body.Transactions == null || body.Transactions.Count == 0)
            return (400, ApiResponse.Error("Invalid request body or empty transactions list"));

        var results = new List<TransactionResponse>();
        int successCount = 0, failCount = 0;

        foreach (var txn in body.Transactions)
        {
            bool success = false;
            string message = "";

            try
            {
                switch (txn.Type?.ToUpperInvariant())
                {
                    case "SI": // Sales Invoice
                        success = _sage.PostTransactionSI(txn.AccountRef, txn.Reference ?? $"SI-{DateTime.Now.Ticks}",
                            txn.NetAmount, txn.TaxAmount, txn.NominalCode ?? "4000", txn.Details ?? "", txn.TaxCode ?? "T1", null, true);
                        message = success ? "Sales invoice posted" : "Failed to post sales invoice";
                        break;

                    case "PI": // Purchase Invoice
                        success = _sage.PostTransactionPI(txn.AccountRef, txn.Reference ?? $"PI-{DateTime.Now.Ticks}",
                            txn.NetAmount, txn.TaxAmount, txn.NominalCode ?? "5000", txn.Details ?? "", txn.TaxCode ?? "T1");
                        message = success ? "Purchase invoice posted" : "Failed to post purchase invoice";
                        break;

                    case "BP": // Bank Payment
                        success = _sage.PostBankPayment(txn.BankNominal ?? "1200", txn.NominalCode ?? "7500",
                            txn.NetAmount, txn.Reference ?? $"BP-{DateTime.Now.Ticks}", txn.Details ?? "", txn.TaxCode ?? "T0");
                        message = success ? "Bank payment posted" : "Failed to post bank payment";
                        break;

                    case "BR": // Bank Receipt
                        success = _sage.PostBankReceipt(txn.BankNominal ?? "1200", txn.NominalCode ?? "4000",
                            txn.NetAmount, txn.Reference ?? $"BR-{DateTime.Now.Ticks}", txn.Details ?? "", txn.TaxCode ?? "T1");
                        message = success ? "Bank receipt posted" : "Failed to post bank receipt";
                        break;

                    default:
                        message = $"Unknown transaction type: {txn.Type}";
                        break;
                }
            }
            catch (Exception ex)
            {
                message = ex.Message;
            }

            if (success) successCount++;
            else failCount++;

            results.Add(new TransactionResponse(success, txn.Reference, message));
        }

        return (200, ApiResponse.Ok(new BatchTransactionResponse(successCount, failCount, results)));
    }

    private async Task<(int, object)> AllocatePaymentAsync(HttpListenerRequest request)
    {
        var body = await ReadBodyAsync<AllocatePaymentRequest>(request);
        if (body == null)
            return (400, ApiResponse.Error("Invalid request body"));

        if (string.IsNullOrWhiteSpace(body.PaymentReference) || string.IsNullOrWhiteSpace(body.InvoiceReference))
            return (400, ApiResponse.Error("paymentReference and invoiceReference are required"));

        var success = _sage.AllocatePayment(body.AccountRef, body.PaymentReference, body.InvoiceReference, body.Amount);
        if (!success)
            return (500, ApiResponse.Error("Failed to allocate payment"));

        return (200, ApiResponse.Ok(new TransactionResponse(true, body.PaymentReference, "Payment allocated")));
    }

    private async Task<(int, object)> PostSalesReceiptAsync(HttpListenerRequest request)
    {
        var body = await ReadBodyAsync<PostSalesReceiptRequest>(request);
        if (body == null)
            return (400, ApiResponse.Error("Invalid request body"));

        if (string.IsNullOrWhiteSpace(body.CustomerAccount))
            return (400, ApiResponse.Error("customerAccount is required"));

        if (!_sage.CustomerExists(body.CustomerAccount))
            return (404, ApiResponse.Error($"Customer not found: {body.CustomerAccount}"));

        var receiptRef = body.ReceiptRef ?? $"SR-{DateTime.Now:yyyyMMddHHmmss}";

        var success = _sage.PostSalesReceipt(
            customerAccount: body.CustomerAccount,
            reference: receiptRef,
            amount: body.Amount,
            bankNominal: body.BankNominal,
            details: body.Details ?? "Receipt via API"
        );

        if (!success)
            return (500, ApiResponse.Error("Failed to post sales receipt"));

        return (200, ApiResponse.Ok(new TransactionResponse(true, receiptRef, "Sales receipt posted")));
    }

    private async Task<(int, object)> PostPurchasePaymentAsync(HttpListenerRequest request)
    {
        var body = await ReadBodyAsync<PostPurchasePaymentRequest>(request);
        if (body == null)
            return (400, ApiResponse.Error("Invalid request body"));

        if (string.IsNullOrWhiteSpace(body.SupplierAccount))
            return (400, ApiResponse.Error("supplierAccount is required"));

        if (!_sage.SupplierExists(body.SupplierAccount))
            return (404, ApiResponse.Error($"Supplier not found: {body.SupplierAccount}"));

        var paymentRef = body.PaymentRef ?? $"PP-{DateTime.Now:yyyyMMddHHmmss}";

        var success = _sage.PostPurchasePayment(
            supplierAccount: body.SupplierAccount,
            reference: paymentRef,
            amount: body.Amount,
            bankNominal: body.BankNominal,
            details: body.Details ?? "Payment via API"
        );

        if (!success)
            return (500, ApiResponse.Error("Failed to post purchase payment"));

        return (200, ApiResponse.Ok(new TransactionResponse(true, paymentRef, "Purchase payment posted")));
    }

    // =========================================================================
    // Project Handlers
    // =========================================================================

    private (int, object) GetProjects(HttpListenerRequest request)
    {
        var search = request.QueryString["search"] ?? "";
        var limitStr = request.QueryString["limit"] ?? "50";
        int.TryParse(limitStr, out var limit);
        if (limit <= 0) limit = 50;

        var projects = _sage.GetProjects(search, limit);
        return (200, ApiResponse.Ok(projects));
    }

    private (int, object) GetProject(string projectRef)
    {
        var project = _sage.GetProject(projectRef);
        if (project == null)
            return (404, ApiResponse.Error($"Project not found: {projectRef}"));

        return (200, ApiResponse.Ok(project));
    }

    private async Task<(int, object)> CreateProjectAsync(HttpListenerRequest request)
    {
        var body = await ReadBodyAsync<CreateProjectRequest>(request);
        if (body == null)
            return (400, ApiResponse.Error("Invalid request body"));

        if (string.IsNullOrWhiteSpace(body.ProjectRef))
            return (400, ApiResponse.Error("projectRef is required"));

        var result = _sage.CreateProject(body);
        if (result == null)
            return (500, ApiResponse.Error("Failed to create project"));

        return (201, ApiResponse.Ok(result));
    }

    private (int, object) GetProjectCostCodes(HttpListenerRequest request)
    {
        var projectRef = request.QueryString["project"] ?? "";
        var limitStr = request.QueryString["limit"] ?? "100";
        int.TryParse(limitStr, out var limit);
        if (limit <= 0) limit = 100;

        var costCodes = _sage.GetProjectCostCodes(projectRef, limit);
        return (200, ApiResponse.Ok(costCodes));
    }

    private (int, object) SearchProjects(HttpListenerRequest request)
    {
        var search = request.QueryString["q"] ?? request.QueryString["search"] ?? "";
        var status = request.QueryString["status"] ?? "";
        var limitStr = request.QueryString["limit"] ?? "50";
        int.TryParse(limitStr, out var limit);
        if (limit <= 0) limit = 50;

        var projects = _sage.SearchProjects(search, status, limit);
        return (200, ApiResponse.Ok(projects));
    }

    // =========================================================================
    // Helpers
    // =========================================================================

    private async Task<T?> ReadBodyAsync<T>(HttpListenerRequest request) where T : class
    {
        if (!request.HasEntityBody)
            return null;

        try
        {
            using var reader = new StreamReader(request.InputStream, request.ContentEncoding);
            var body = await reader.ReadToEndAsync();
            return JsonSerializer.Deserialize<T>(body, _jsonOptions);
        }
        catch
        {
            return null;
        }
    }

    public void Dispose()
    {
        if (_disposed) return;
        _disposed = true;

        _listener.Close();
        _sage.Dispose();
    }
}
