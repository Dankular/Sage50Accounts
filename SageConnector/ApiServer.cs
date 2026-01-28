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
        _prefix = $"http://localhost:{port}/";
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
