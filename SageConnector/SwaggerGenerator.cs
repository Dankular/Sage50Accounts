namespace SageConnector;

/// <summary>
/// Generates OpenAPI 3.0 specification for the Sage 50 API
/// </summary>
public static class SwaggerGenerator
{
    public static object Generate(string serverUrl = "http://localhost:5000")
    {
        return new Dictionary<string, object>
        {
            ["openapi"] = "3.0.0",
            ["info"] = new Dictionary<string, object>
            {
                ["title"] = "Sage 50 Accounts API",
                ["description"] = "REST API for Sage 50 Accounts (UK) via SDO interface",
                ["version"] = "1.0.0",
                ["contact"] = new Dictionary<string, string>
                {
                    ["name"] = "Sage 50 SDK Tools"
                }
            },
            ["servers"] = new[]
            {
                new Dictionary<string, string> { ["url"] = serverUrl }
            },
            ["paths"] = GeneratePaths(),
            ["components"] = new Dictionary<string, object>
            {
                ["schemas"] = GenerateSchemas()
            }
        };
    }

    private static Dictionary<string, object> GeneratePaths()
    {
        return new Dictionary<string, object>
        {
            // Company
            ["/api/company"] = new Dictionary<string, object>
            {
                ["get"] = new Dictionary<string, object>
                {
                    ["tags"] = new[] { "Company" },
                    ["summary"] = "Get company information",
                    ["operationId"] = "getCompanyInfo",
                    ["responses"] = SuccessResponse("CompanyInfoResponse", "Company information retrieved")
                }
            },

            // Customers
            ["/api/customers"] = new Dictionary<string, object>
            {
                ["get"] = new Dictionary<string, object>
                {
                    ["tags"] = new[] { "Customers" },
                    ["summary"] = "List customers",
                    ["operationId"] = "listCustomers",
                    ["parameters"] = new[]
                    {
                        QueryParam("search", "Search term for customer name or reference"),
                        QueryParam("limit", "Maximum number of results", "integer", "50")
                    },
                    ["responses"] = SuccessResponse("CustomerListResponse", "List of customers")
                },
                ["post"] = new Dictionary<string, object>
                {
                    ["tags"] = new[] { "Customers" },
                    ["summary"] = "Create a new customer",
                    ["operationId"] = "createCustomer",
                    ["requestBody"] = RequestBody("CreateCustomerRequest"),
                    ["responses"] = SuccessResponse("CustomerResponse", "Customer created")
                }
            },
            ["/api/customers/{accountRef}"] = new Dictionary<string, object>
            {
                ["get"] = new Dictionary<string, object>
                {
                    ["tags"] = new[] { "Customers" },
                    ["summary"] = "Get customer by account reference",
                    ["operationId"] = "getCustomer",
                    ["parameters"] = new[] { PathParam("accountRef", "Customer account reference (max 8 chars)") },
                    ["responses"] = SuccessResponse("CustomerResponse", "Customer details")
                }
            },
            ["/api/customers/{accountRef}/exists"] = new Dictionary<string, object>
            {
                ["get"] = new Dictionary<string, object>
                {
                    ["tags"] = new[] { "Customers" },
                    ["summary"] = "Check if customer exists",
                    ["operationId"] = "customerExists",
                    ["parameters"] = new[] { PathParam("accountRef", "Customer account reference") },
                    ["responses"] = SuccessResponse("ExistsResponse", "Existence check result")
                }
            },

            // Suppliers
            ["/api/suppliers"] = new Dictionary<string, object>
            {
                ["get"] = new Dictionary<string, object>
                {
                    ["tags"] = new[] { "Suppliers" },
                    ["summary"] = "List suppliers",
                    ["operationId"] = "listSuppliers",
                    ["parameters"] = new[]
                    {
                        QueryParam("search", "Search term for supplier name or reference"),
                        QueryParam("limit", "Maximum number of results", "integer", "50")
                    },
                    ["responses"] = SuccessResponse("SupplierListResponse", "List of suppliers")
                },
                ["post"] = new Dictionary<string, object>
                {
                    ["tags"] = new[] { "Suppliers" },
                    ["summary"] = "Create a new supplier",
                    ["operationId"] = "createSupplier",
                    ["requestBody"] = RequestBody("CreateSupplierRequest"),
                    ["responses"] = SuccessResponse("SupplierResponse", "Supplier created")
                }
            },
            ["/api/suppliers/{accountRef}"] = new Dictionary<string, object>
            {
                ["get"] = new Dictionary<string, object>
                {
                    ["tags"] = new[] { "Suppliers" },
                    ["summary"] = "Get supplier by account reference",
                    ["operationId"] = "getSupplier",
                    ["parameters"] = new[] { PathParam("accountRef", "Supplier account reference (max 8 chars)") },
                    ["responses"] = SuccessResponse("SupplierResponse", "Supplier details")
                }
            },
            ["/api/suppliers/{accountRef}/exists"] = new Dictionary<string, object>
            {
                ["get"] = new Dictionary<string, object>
                {
                    ["tags"] = new[] { "Suppliers" },
                    ["summary"] = "Check if supplier exists",
                    ["operationId"] = "supplierExists",
                    ["parameters"] = new[] { PathParam("accountRef", "Supplier account reference") },
                    ["responses"] = SuccessResponse("ExistsResponse", "Existence check result")
                }
            },

            // Sales Ledger
            ["/api/sales/invoice"] = new Dictionary<string, object>
            {
                ["post"] = new Dictionary<string, object>
                {
                    ["tags"] = new[] { "Sales" },
                    ["summary"] = "Post sales invoice to ledger",
                    ["description"] = "Posts a sales invoice using TransactionPost (updates customer balance)",
                    ["operationId"] = "postSalesInvoice",
                    ["requestBody"] = RequestBody("PostSalesInvoiceRequest"),
                    ["responses"] = SuccessResponse("TransactionResponse", "Invoice posted")
                }
            },
            ["/api/sales/credit"] = new Dictionary<string, object>
            {
                ["post"] = new Dictionary<string, object>
                {
                    ["tags"] = new[] { "Sales" },
                    ["summary"] = "Post sales credit note",
                    ["operationId"] = "postSalesCredit",
                    ["requestBody"] = RequestBody("PostSalesCreditRequest"),
                    ["responses"] = SuccessResponse("TransactionResponse", "Credit note posted")
                }
            },

            // Purchase Ledger
            ["/api/purchases/invoice"] = new Dictionary<string, object>
            {
                ["post"] = new Dictionary<string, object>
                {
                    ["tags"] = new[] { "Purchases" },
                    ["summary"] = "Post purchase invoice to ledger",
                    ["description"] = "Posts a purchase invoice using TransactionPost (updates supplier balance)",
                    ["operationId"] = "postPurchaseInvoice",
                    ["requestBody"] = RequestBody("PostPurchaseInvoiceRequest"),
                    ["responses"] = SuccessResponse("TransactionResponse", "Invoice posted")
                }
            },
            ["/api/purchases/credit"] = new Dictionary<string, object>
            {
                ["post"] = new Dictionary<string, object>
                {
                    ["tags"] = new[] { "Purchases" },
                    ["summary"] = "Post purchase credit note",
                    ["operationId"] = "postPurchaseCredit",
                    ["requestBody"] = RequestBody("PostPurchaseCreditRequest"),
                    ["responses"] = SuccessResponse("TransactionResponse", "Credit note posted")
                }
            },

            // Nominal Ledger
            ["/api/nominals"] = new Dictionary<string, object>
            {
                ["get"] = new Dictionary<string, object>
                {
                    ["tags"] = new[] { "Nominals" },
                    ["summary"] = "List nominal codes",
                    ["operationId"] = "listNominals",
                    ["parameters"] = new[]
                    {
                        QueryParam("limit", "Maximum number of results", "integer", "100")
                    },
                    ["responses"] = SuccessResponse("NominalListResponse", "List of nominal codes")
                },
                ["post"] = new Dictionary<string, object>
                {
                    ["tags"] = new[] { "Nominals" },
                    ["summary"] = "Create a new nominal code",
                    ["operationId"] = "createNominal",
                    ["requestBody"] = RequestBody("CreateNominalRequest"),
                    ["responses"] = SuccessResponse("TransactionResponse", "Nominal code created")
                }
            },
            ["/api/nominals/{code}/exists"] = new Dictionary<string, object>
            {
                ["get"] = new Dictionary<string, object>
                {
                    ["tags"] = new[] { "Nominals" },
                    ["summary"] = "Check if nominal code exists",
                    ["operationId"] = "nominalExists",
                    ["parameters"] = new[] { PathParam("code", "Nominal code") },
                    ["responses"] = SuccessResponse("ExistsResponse", "Existence check result")
                }
            },

            // Bank
            ["/api/bank/payment"] = new Dictionary<string, object>
            {
                ["post"] = new Dictionary<string, object>
                {
                    ["tags"] = new[] { "Bank" },
                    ["summary"] = "Post bank payment",
                    ["operationId"] = "postBankPayment",
                    ["requestBody"] = RequestBody("PostBankPaymentRequest"),
                    ["responses"] = SuccessResponse("TransactionResponse", "Payment posted")
                }
            },
            ["/api/bank/receipt"] = new Dictionary<string, object>
            {
                ["post"] = new Dictionary<string, object>
                {
                    ["tags"] = new[] { "Bank" },
                    ["summary"] = "Post bank receipt",
                    ["operationId"] = "postBankReceipt",
                    ["requestBody"] = RequestBody("PostBankReceiptRequest"),
                    ["responses"] = SuccessResponse("TransactionResponse", "Receipt posted")
                }
            },

            // Journals
            ["/api/journals"] = new Dictionary<string, object>
            {
                ["post"] = new Dictionary<string, object>
                {
                    ["tags"] = new[] { "Journals" },
                    ["summary"] = "Post journal entry",
                    ["operationId"] = "postJournal",
                    ["requestBody"] = RequestBody("PostJournalRequest"),
                    ["responses"] = SuccessResponse("TransactionResponse", "Journal posted")
                }
            },
            ["/api/journals/simple"] = new Dictionary<string, object>
            {
                ["post"] = new Dictionary<string, object>
                {
                    ["tags"] = new[] { "Journals" },
                    ["summary"] = "Post simple two-line journal",
                    ["description"] = "Posts a simple journal with one debit and one credit line",
                    ["operationId"] = "postSimpleJournal",
                    ["requestBody"] = RequestBody("PostSimpleJournalRequest"),
                    ["responses"] = SuccessResponse("TransactionResponse", "Journal posted")
                }
            }
        };
    }

    private static Dictionary<string, object> GenerateSchemas()
    {
        return new Dictionary<string, object>
        {
            // Response wrappers
            ["ApiResponse"] = new Dictionary<string, object>
            {
                ["type"] = "object",
                ["properties"] = new Dictionary<string, object>
                {
                    ["success"] = new Dictionary<string, string> { ["type"] = "boolean" },
                    ["data"] = new Dictionary<string, string> { ["type"] = "object" },
                    ["error"] = new Dictionary<string, string> { ["type"] = "string", ["nullable"] = "true" }
                }
            },

            // Company
            ["CompanyInfoResponse"] = new Dictionary<string, object>
            {
                ["type"] = "object",
                ["properties"] = new Dictionary<string, object>
                {
                    ["name"] = StringProp(),
                    ["address1"] = StringProp(true),
                    ["address2"] = StringProp(true),
                    ["address3"] = StringProp(true),
                    ["postcode"] = StringProp(true),
                    ["telephone"] = StringProp(true),
                    ["vatNumber"] = StringProp(true)
                }
            },

            // Customer
            ["CustomerResponse"] = new Dictionary<string, object>
            {
                ["type"] = "object",
                ["properties"] = new Dictionary<string, object>
                {
                    ["accountRef"] = StringProp(),
                    ["name"] = StringProp(),
                    ["address1"] = StringProp(true),
                    ["address2"] = StringProp(true),
                    ["address3"] = StringProp(true),
                    ["postcode"] = StringProp(true),
                    ["telephone"] = StringProp(true),
                    ["email"] = StringProp(true),
                    ["contactName"] = StringProp(true),
                    ["balance"] = NumberProp(),
                    ["creditLimit"] = NumberProp()
                }
            },
            ["CustomerListResponse"] = ArrayOfRef("CustomerResponse"),
            ["CreateCustomerRequest"] = new Dictionary<string, object>
            {
                ["type"] = "object",
                ["required"] = new[] { "accountRef", "name" },
                ["properties"] = new Dictionary<string, object>
                {
                    ["accountRef"] = StringProp(description: "Max 8 characters, no spaces"),
                    ["name"] = StringProp(),
                    ["address1"] = StringProp(true),
                    ["address2"] = StringProp(true),
                    ["address3"] = StringProp(true),
                    ["postcode"] = StringProp(true),
                    ["telephone"] = StringProp(true),
                    ["email"] = StringProp(true),
                    ["contactName"] = StringProp(true),
                    ["creditLimit"] = NumberProp(true)
                }
            },

            // Supplier
            ["SupplierResponse"] = new Dictionary<string, object>
            {
                ["type"] = "object",
                ["properties"] = new Dictionary<string, object>
                {
                    ["accountRef"] = StringProp(),
                    ["name"] = StringProp(),
                    ["address1"] = StringProp(true),
                    ["address2"] = StringProp(true),
                    ["address3"] = StringProp(true),
                    ["postcode"] = StringProp(true),
                    ["telephone"] = StringProp(true),
                    ["email"] = StringProp(true),
                    ["contactName"] = StringProp(true),
                    ["balance"] = NumberProp(),
                    ["creditLimit"] = NumberProp()
                }
            },
            ["SupplierListResponse"] = ArrayOfRef("SupplierResponse"),
            ["CreateSupplierRequest"] = new Dictionary<string, object>
            {
                ["type"] = "object",
                ["required"] = new[] { "accountRef", "name" },
                ["properties"] = new Dictionary<string, object>
                {
                    ["accountRef"] = StringProp(description: "Max 8 characters, no spaces"),
                    ["name"] = StringProp(),
                    ["address1"] = StringProp(true),
                    ["address2"] = StringProp(true),
                    ["address3"] = StringProp(true),
                    ["postcode"] = StringProp(true),
                    ["telephone"] = StringProp(true),
                    ["email"] = StringProp(true),
                    ["contactName"] = StringProp(true),
                    ["creditLimit"] = NumberProp(true)
                }
            },

            // Sales
            ["PostSalesInvoiceRequest"] = new Dictionary<string, object>
            {
                ["type"] = "object",
                ["required"] = new[] { "customerAccount", "netAmount" },
                ["properties"] = new Dictionary<string, object>
                {
                    ["customerAccount"] = StringProp(description: "Customer account reference"),
                    ["invoiceRef"] = StringProp(true, "Invoice reference (auto-generated if omitted)"),
                    ["netAmount"] = NumberProp(description: "Net amount before VAT"),
                    ["taxAmount"] = NumberProp(description: "VAT amount"),
                    ["nominalCode"] = StringProp(description: "Sales nominal code (default: 4000)"),
                    ["details"] = StringProp(true, "Transaction details"),
                    ["taxCode"] = StringProp(description: "Tax code (default: T1)"),
                    ["autoCreateCustomer"] = BoolProp("Auto-create customer if not found")
                }
            },
            ["PostSalesCreditRequest"] = new Dictionary<string, object>
            {
                ["type"] = "object",
                ["required"] = new[] { "customerAccount", "netAmount" },
                ["properties"] = new Dictionary<string, object>
                {
                    ["customerAccount"] = StringProp(),
                    ["creditRef"] = StringProp(true),
                    ["netAmount"] = NumberProp(),
                    ["taxAmount"] = NumberProp(),
                    ["nominalCode"] = StringProp(),
                    ["details"] = StringProp(true),
                    ["taxCode"] = StringProp()
                }
            },

            // Purchases
            ["PostPurchaseInvoiceRequest"] = new Dictionary<string, object>
            {
                ["type"] = "object",
                ["required"] = new[] { "supplierAccount", "netAmount" },
                ["properties"] = new Dictionary<string, object>
                {
                    ["supplierAccount"] = StringProp(description: "Supplier account reference"),
                    ["invoiceRef"] = StringProp(true, "Invoice reference (auto-generated if omitted)"),
                    ["netAmount"] = NumberProp(description: "Net amount before VAT"),
                    ["taxAmount"] = NumberProp(description: "VAT amount"),
                    ["nominalCode"] = StringProp(description: "Purchase nominal code (default: 5000)"),
                    ["details"] = StringProp(true, "Transaction details"),
                    ["taxCode"] = StringProp(description: "Tax code (default: T1)"),
                    ["autoCreateSupplier"] = BoolProp("Auto-create supplier if not found")
                }
            },
            ["PostPurchaseCreditRequest"] = new Dictionary<string, object>
            {
                ["type"] = "object",
                ["required"] = new[] { "supplierAccount", "netAmount" },
                ["properties"] = new Dictionary<string, object>
                {
                    ["supplierAccount"] = StringProp(),
                    ["creditRef"] = StringProp(true),
                    ["netAmount"] = NumberProp(),
                    ["taxAmount"] = NumberProp(),
                    ["nominalCode"] = StringProp(),
                    ["details"] = StringProp(true),
                    ["taxCode"] = StringProp()
                }
            },

            // Nominals
            ["NominalCodeResponse"] = new Dictionary<string, object>
            {
                ["type"] = "object",
                ["properties"] = new Dictionary<string, object>
                {
                    ["code"] = StringProp(),
                    ["name"] = StringProp(),
                    ["balance"] = NumberProp()
                }
            },
            ["NominalListResponse"] = ArrayOfRef("NominalCodeResponse"),
            ["CreateNominalRequest"] = new Dictionary<string, object>
            {
                ["type"] = "object",
                ["required"] = new[] { "code", "name" },
                ["properties"] = new Dictionary<string, object>
                {
                    ["code"] = StringProp(description: "Nominal code (e.g., 4000)"),
                    ["name"] = StringProp(description: "Nominal code name")
                }
            },

            // Bank
            ["PostBankPaymentRequest"] = new Dictionary<string, object>
            {
                ["type"] = "object",
                ["required"] = new[] { "netAmount" },
                ["properties"] = new Dictionary<string, object>
                {
                    ["bankNominal"] = StringProp(description: "Bank nominal code (default: 1200)"),
                    ["expenseNominal"] = StringProp(description: "Expense nominal code (default: 7500)"),
                    ["netAmount"] = NumberProp(),
                    ["reference"] = StringProp(true),
                    ["details"] = StringProp(true),
                    ["taxCode"] = StringProp(description: "Tax code (default: T0)")
                }
            },
            ["PostBankReceiptRequest"] = new Dictionary<string, object>
            {
                ["type"] = "object",
                ["required"] = new[] { "netAmount" },
                ["properties"] = new Dictionary<string, object>
                {
                    ["bankNominal"] = StringProp(description: "Bank nominal code (default: 1200)"),
                    ["incomeNominal"] = StringProp(description: "Income nominal code (default: 4000)"),
                    ["netAmount"] = NumberProp(),
                    ["reference"] = StringProp(true),
                    ["details"] = StringProp(true),
                    ["taxCode"] = StringProp(description: "Tax code (default: T1)")
                }
            },

            // Journals
            ["JournalLineRequest"] = new Dictionary<string, object>
            {
                ["type"] = "object",
                ["required"] = new[] { "nominalCode" },
                ["properties"] = new Dictionary<string, object>
                {
                    ["nominalCode"] = StringProp(),
                    ["debit"] = NumberProp(description: "Debit amount (0 if credit)"),
                    ["credit"] = NumberProp(description: "Credit amount (0 if debit)"),
                    ["details"] = StringProp(true)
                }
            },
            ["PostJournalRequest"] = new Dictionary<string, object>
            {
                ["type"] = "object",
                ["required"] = new[] { "lines" },
                ["properties"] = new Dictionary<string, object>
                {
                    ["lines"] = new Dictionary<string, object>
                    {
                        ["type"] = "array",
                        ["items"] = new Dictionary<string, string> { ["$ref"] = "#/components/schemas/JournalLineRequest" }
                    },
                    ["reference"] = StringProp(true),
                    ["date"] = new Dictionary<string, object> { ["type"] = "string", ["format"] = "date", ["nullable"] = true }
                }
            },
            ["PostSimpleJournalRequest"] = new Dictionary<string, object>
            {
                ["type"] = "object",
                ["required"] = new[] { "debitNominal", "creditNominal", "amount" },
                ["properties"] = new Dictionary<string, object>
                {
                    ["debitNominal"] = StringProp(description: "Nominal code to debit"),
                    ["creditNominal"] = StringProp(description: "Nominal code to credit"),
                    ["amount"] = NumberProp(description: "Journal amount"),
                    ["reference"] = StringProp(true),
                    ["details"] = StringProp(true),
                    ["date"] = new Dictionary<string, object> { ["type"] = "string", ["format"] = "date", ["nullable"] = true }
                }
            },

            // Common
            ["TransactionResponse"] = new Dictionary<string, object>
            {
                ["type"] = "object",
                ["properties"] = new Dictionary<string, object>
                {
                    ["success"] = BoolProp(),
                    ["reference"] = StringProp(true),
                    ["message"] = StringProp(true)
                }
            },
            ["ExistsResponse"] = new Dictionary<string, object>
            {
                ["type"] = "object",
                ["properties"] = new Dictionary<string, object>
                {
                    ["exists"] = BoolProp(),
                    ["accountRef"] = StringProp(true)
                }
            }
        };
    }

    // Helper methods
    private static Dictionary<string, object> PathParam(string name, string description) => new()
    {
        ["name"] = name,
        ["in"] = "path",
        ["required"] = true,
        ["schema"] = new Dictionary<string, string> { ["type"] = "string" },
        ["description"] = description
    };

    private static Dictionary<string, object> QueryParam(string name, string description, string type = "string", string? defaultValue = null)
    {
        var param = new Dictionary<string, object>
        {
            ["name"] = name,
            ["in"] = "query",
            ["required"] = false,
            ["schema"] = new Dictionary<string, object> { ["type"] = type },
            ["description"] = description
        };
        if (defaultValue != null)
            ((Dictionary<string, object>)param["schema"])["default"] = type == "integer" ? int.Parse(defaultValue) : defaultValue;
        return param;
    }

    private static Dictionary<string, object> RequestBody(string schemaRef) => new()
    {
        ["required"] = true,
        ["content"] = new Dictionary<string, object>
        {
            ["application/json"] = new Dictionary<string, object>
            {
                ["schema"] = new Dictionary<string, string> { ["$ref"] = $"#/components/schemas/{schemaRef}" }
            }
        }
    };

    private static Dictionary<string, object> SuccessResponse(string schemaRef, string description) => new()
    {
        ["200"] = new Dictionary<string, object>
        {
            ["description"] = description,
            ["content"] = new Dictionary<string, object>
            {
                ["application/json"] = new Dictionary<string, object>
                {
                    ["schema"] = new Dictionary<string, object>
                    {
                        ["type"] = "object",
                        ["properties"] = new Dictionary<string, object>
                        {
                            ["success"] = new Dictionary<string, string> { ["type"] = "boolean" },
                            ["data"] = new Dictionary<string, string> { ["$ref"] = $"#/components/schemas/{schemaRef}" },
                            ["error"] = new Dictionary<string, object> { ["type"] = "string", ["nullable"] = true }
                        }
                    }
                }
            }
        },
        ["400"] = ErrorResponse("Bad request"),
        ["404"] = ErrorResponse("Not found"),
        ["500"] = ErrorResponse("Internal server error")
    };

    private static Dictionary<string, object> ErrorResponse(string description) => new()
    {
        ["description"] = description,
        ["content"] = new Dictionary<string, object>
        {
            ["application/json"] = new Dictionary<string, object>
            {
                ["schema"] = new Dictionary<string, object>
                {
                    ["type"] = "object",
                    ["properties"] = new Dictionary<string, object>
                    {
                        ["success"] = new Dictionary<string, object> { ["type"] = "boolean", ["example"] = false },
                        ["error"] = new Dictionary<string, string> { ["type"] = "string" }
                    }
                }
            }
        }
    };

    private static Dictionary<string, object> StringProp(bool nullable = false, string? description = null)
    {
        var prop = new Dictionary<string, object> { ["type"] = "string" };
        if (nullable) prop["nullable"] = true;
        if (description != null) prop["description"] = description;
        return prop;
    }

    private static Dictionary<string, object> NumberProp(bool nullable = false, string? description = null)
    {
        var prop = new Dictionary<string, object> { ["type"] = "number", ["format"] = "decimal" };
        if (nullable) prop["nullable"] = true;
        if (description != null) prop["description"] = description;
        return prop;
    }

    private static Dictionary<string, object> BoolProp(string? description = null)
    {
        var prop = new Dictionary<string, object> { ["type"] = "boolean" };
        if (description != null) prop["description"] = description;
        return prop;
    }

    private static Dictionary<string, object> ArrayOfRef(string schemaRef) => new()
    {
        ["type"] = "array",
        ["items"] = new Dictionary<string, string> { ["$ref"] = $"#/components/schemas/{schemaRef}" }
    };
}
