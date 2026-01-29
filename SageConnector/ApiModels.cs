using System.Text.Json.Serialization;

namespace SageConnector;

// ============================================================================
// API Response Wrapper
// ============================================================================

public record ApiResponse<T>(
    [property: JsonPropertyName("success")] bool Success,
    [property: JsonPropertyName("data")] T? Data,
    [property: JsonPropertyName("error")] string? Error = null
);

public static class ApiResponse
{
    public static ApiResponse<T> Ok<T>(T data) => new(true, data);
    public static ApiResponse<object> Error(string message) => new(false, null, message);
}

// ============================================================================
// Customer DTOs
// ============================================================================

public record CustomerResponse(
    [property: JsonPropertyName("accountRef")] string AccountRef,
    [property: JsonPropertyName("name")] string Name,
    [property: JsonPropertyName("address1")] string? Address1,
    [property: JsonPropertyName("address2")] string? Address2,
    [property: JsonPropertyName("address3")] string? Address3,
    [property: JsonPropertyName("postcode")] string? Postcode,
    [property: JsonPropertyName("telephone")] string? Telephone,
    [property: JsonPropertyName("email")] string? Email,
    [property: JsonPropertyName("contactName")] string? ContactName,
    [property: JsonPropertyName("balance")] decimal Balance,
    [property: JsonPropertyName("creditLimit")] decimal CreditLimit
);

public record CreateCustomerRequest(
    [property: JsonPropertyName("accountRef")] string AccountRef,
    [property: JsonPropertyName("name")] string Name,
    [property: JsonPropertyName("address1")] string? Address1 = null,
    [property: JsonPropertyName("address2")] string? Address2 = null,
    [property: JsonPropertyName("address3")] string? Address3 = null,
    [property: JsonPropertyName("postcode")] string? Postcode = null,
    [property: JsonPropertyName("telephone")] string? Telephone = null,
    [property: JsonPropertyName("email")] string? Email = null,
    [property: JsonPropertyName("contactName")] string? ContactName = null,
    [property: JsonPropertyName("creditLimit")] decimal? CreditLimit = null
);

// ============================================================================
// Supplier DTOs
// ============================================================================

public record SupplierResponse(
    [property: JsonPropertyName("accountRef")] string AccountRef,
    [property: JsonPropertyName("name")] string Name,
    [property: JsonPropertyName("address1")] string? Address1,
    [property: JsonPropertyName("address2")] string? Address2,
    [property: JsonPropertyName("address3")] string? Address3,
    [property: JsonPropertyName("postcode")] string? Postcode,
    [property: JsonPropertyName("telephone")] string? Telephone,
    [property: JsonPropertyName("email")] string? Email,
    [property: JsonPropertyName("contactName")] string? ContactName,
    [property: JsonPropertyName("balance")] decimal Balance,
    [property: JsonPropertyName("creditLimit")] decimal CreditLimit
);

public record CreateSupplierRequest(
    [property: JsonPropertyName("accountRef")] string AccountRef,
    [property: JsonPropertyName("name")] string Name,
    [property: JsonPropertyName("address1")] string? Address1 = null,
    [property: JsonPropertyName("address2")] string? Address2 = null,
    [property: JsonPropertyName("address3")] string? Address3 = null,
    [property: JsonPropertyName("postcode")] string? Postcode = null,
    [property: JsonPropertyName("telephone")] string? Telephone = null,
    [property: JsonPropertyName("email")] string? Email = null,
    [property: JsonPropertyName("contactName")] string? ContactName = null,
    [property: JsonPropertyName("creditLimit")] decimal? CreditLimit = null
);

// ============================================================================
// Sales Invoice DTOs
// ============================================================================

public record PostSalesInvoiceRequest(
    [property: JsonPropertyName("customerAccount")] string CustomerAccount,
    [property: JsonPropertyName("invoiceRef")] string? InvoiceRef = null,
    [property: JsonPropertyName("netAmount")] decimal NetAmount = 0,
    [property: JsonPropertyName("taxAmount")] decimal TaxAmount = 0,
    [property: JsonPropertyName("nominalCode")] string NominalCode = "4000",
    [property: JsonPropertyName("details")] string? Details = null,
    [property: JsonPropertyName("taxCode")] string TaxCode = "T1",
    [property: JsonPropertyName("autoCreateCustomer")] bool AutoCreateCustomer = false
);

public record PostSalesCreditRequest(
    [property: JsonPropertyName("customerAccount")] string CustomerAccount,
    [property: JsonPropertyName("creditRef")] string? CreditRef = null,
    [property: JsonPropertyName("netAmount")] decimal NetAmount = 0,
    [property: JsonPropertyName("taxAmount")] decimal TaxAmount = 0,
    [property: JsonPropertyName("nominalCode")] string NominalCode = "4000",
    [property: JsonPropertyName("details")] string? Details = null,
    [property: JsonPropertyName("taxCode")] string TaxCode = "T1"
);

public record PostSalesReceiptRequest(
    [property: JsonPropertyName("customerAccount")] string CustomerAccount,
    [property: JsonPropertyName("receiptRef")] string? ReceiptRef = null,
    [property: JsonPropertyName("amount")] decimal Amount = 0,
    [property: JsonPropertyName("bankNominal")] string BankNominal = "1200",
    [property: JsonPropertyName("details")] string? Details = null
);

// ============================================================================
// Purchase Invoice DTOs
// ============================================================================

public record PostPurchaseInvoiceRequest(
    [property: JsonPropertyName("supplierAccount")] string SupplierAccount,
    [property: JsonPropertyName("invoiceRef")] string? InvoiceRef = null,
    [property: JsonPropertyName("netAmount")] decimal NetAmount = 0,
    [property: JsonPropertyName("taxAmount")] decimal TaxAmount = 0,
    [property: JsonPropertyName("nominalCode")] string NominalCode = "5000",
    [property: JsonPropertyName("details")] string? Details = null,
    [property: JsonPropertyName("taxCode")] string TaxCode = "T1",
    [property: JsonPropertyName("autoCreateSupplier")] bool AutoCreateSupplier = false
);

public record PostPurchaseCreditRequest(
    [property: JsonPropertyName("supplierAccount")] string SupplierAccount,
    [property: JsonPropertyName("creditRef")] string? CreditRef = null,
    [property: JsonPropertyName("netAmount")] decimal NetAmount = 0,
    [property: JsonPropertyName("taxAmount")] decimal TaxAmount = 0,
    [property: JsonPropertyName("nominalCode")] string NominalCode = "5000",
    [property: JsonPropertyName("details")] string? Details = null,
    [property: JsonPropertyName("taxCode")] string TaxCode = "T1"
);

public record PostPurchasePaymentRequest(
    [property: JsonPropertyName("supplierAccount")] string SupplierAccount,
    [property: JsonPropertyName("paymentRef")] string? PaymentRef = null,
    [property: JsonPropertyName("amount")] decimal Amount = 0,
    [property: JsonPropertyName("bankNominal")] string BankNominal = "1200",
    [property: JsonPropertyName("details")] string? Details = null
);

// ============================================================================
// Bank DTOs
// ============================================================================

public record PostBankPaymentRequest(
    [property: JsonPropertyName("bankNominal")] string BankNominal = "1200",
    [property: JsonPropertyName("expenseNominal")] string ExpenseNominal = "7500",
    [property: JsonPropertyName("netAmount")] decimal NetAmount = 0,
    [property: JsonPropertyName("reference")] string? Reference = null,
    [property: JsonPropertyName("details")] string? Details = null,
    [property: JsonPropertyName("taxCode")] string TaxCode = "T0"
);

public record PostBankReceiptRequest(
    [property: JsonPropertyName("bankNominal")] string BankNominal = "1200",
    [property: JsonPropertyName("incomeNominal")] string IncomeNominal = "4000",
    [property: JsonPropertyName("netAmount")] decimal NetAmount = 0,
    [property: JsonPropertyName("reference")] string? Reference = null,
    [property: JsonPropertyName("details")] string? Details = null,
    [property: JsonPropertyName("taxCode")] string TaxCode = "T1"
);

// ============================================================================
// Journal DTOs
// ============================================================================

public record JournalLineRequest(
    [property: JsonPropertyName("nominalCode")] string NominalCode,
    [property: JsonPropertyName("debit")] decimal Debit = 0,
    [property: JsonPropertyName("credit")] decimal Credit = 0,
    [property: JsonPropertyName("details")] string? Details = null
);

public record PostJournalRequest(
    [property: JsonPropertyName("lines")] List<JournalLineRequest> Lines,
    [property: JsonPropertyName("reference")] string? Reference = null,
    [property: JsonPropertyName("date")] DateTime? Date = null
);

public record PostSimpleJournalRequest(
    [property: JsonPropertyName("debitNominal")] string DebitNominal,
    [property: JsonPropertyName("creditNominal")] string CreditNominal,
    [property: JsonPropertyName("amount")] decimal Amount,
    [property: JsonPropertyName("reference")] string? Reference = null,
    [property: JsonPropertyName("details")] string? Details = null,
    [property: JsonPropertyName("date")] DateTime? Date = null
);

// ============================================================================
// Nominal DTOs
// ============================================================================

public record NominalCodeResponse(
    [property: JsonPropertyName("code")] string Code,
    [property: JsonPropertyName("name")] string Name,
    [property: JsonPropertyName("balance")] decimal Balance
);

public record CreateNominalRequest(
    [property: JsonPropertyName("code")] string Code,
    [property: JsonPropertyName("name")] string Name
);

// ============================================================================
// Company DTOs
// ============================================================================

public record CompanyInfoResponse(
    [property: JsonPropertyName("name")] string Name,
    [property: JsonPropertyName("address1")] string? Address1,
    [property: JsonPropertyName("address2")] string? Address2,
    [property: JsonPropertyName("address3")] string? Address3,
    [property: JsonPropertyName("postcode")] string? Postcode,
    [property: JsonPropertyName("telephone")] string? Telephone,
    [property: JsonPropertyName("vatNumber")] string? VatNumber
);

// ============================================================================
// Transaction Response
// ============================================================================

public record TransactionResponse(
    [property: JsonPropertyName("success")] bool Success,
    [property: JsonPropertyName("reference")] string? Reference,
    [property: JsonPropertyName("message")] string? Message
);

// ============================================================================
// Exists Response
// ============================================================================

public record ExistsResponse(
    [property: JsonPropertyName("exists")] bool Exists,
    [property: JsonPropertyName("accountRef")] string? AccountRef = null
);

// ============================================================================
// Purchase Order DTOs
// ============================================================================

public record PurchaseOrderResponse(
    [property: JsonPropertyName("orderNumber")] string OrderNumber,
    [property: JsonPropertyName("supplierAccount")] string SupplierAccount,
    [property: JsonPropertyName("supplierName")] string? SupplierName,
    [property: JsonPropertyName("orderDate")] DateTime? OrderDate,
    [property: JsonPropertyName("deliveryDate")] DateTime? DeliveryDate,
    [property: JsonPropertyName("reference")] string? Reference,
    [property: JsonPropertyName("notes")] string? Notes,
    [property: JsonPropertyName("netAmount")] decimal NetAmount,
    [property: JsonPropertyName("taxAmount")] decimal TaxAmount,
    [property: JsonPropertyName("grossAmount")] decimal GrossAmount,
    [property: JsonPropertyName("status")] int Status,
    [property: JsonPropertyName("items")] List<PurchaseOrderItemResponse>? Items = null
);

public record PurchaseOrderItemResponse(
    [property: JsonPropertyName("lineNumber")] int LineNumber,
    [property: JsonPropertyName("stockCode")] string? StockCode,
    [property: JsonPropertyName("description")] string? Description,
    [property: JsonPropertyName("quantity")] decimal Quantity,
    [property: JsonPropertyName("unitPrice")] decimal UnitPrice,
    [property: JsonPropertyName("netAmount")] decimal NetAmount,
    [property: JsonPropertyName("taxCode")] string? TaxCode,
    [property: JsonPropertyName("taxAmount")] decimal TaxAmount,
    [property: JsonPropertyName("nominalCode")] string? NominalCode
);

public record PurchaseOrderItemRequest(
    [property: JsonPropertyName("stockCode")] string? StockCode = null,
    [property: JsonPropertyName("description")] string Description = "",
    [property: JsonPropertyName("quantity")] decimal Quantity = 1,
    [property: JsonPropertyName("unitPrice")] decimal UnitPrice = 0,
    [property: JsonPropertyName("taxCode")] string TaxCode = "T1",
    [property: JsonPropertyName("nominalCode")] string NominalCode = "5000"
);

public record CreatePurchaseOrderRequest(
    [property: JsonPropertyName("supplierAccount")] string SupplierAccount,
    [property: JsonPropertyName("orderDate")] DateTime? OrderDate = null,
    [property: JsonPropertyName("deliveryDate")] DateTime? DeliveryDate = null,
    [property: JsonPropertyName("reference")] string? Reference = null,
    [property: JsonPropertyName("notes")] string? Notes = null,
    [property: JsonPropertyName("items")] List<PurchaseOrderItemRequest>? Items = null,
    [property: JsonPropertyName("autoCreateSupplier")] bool AutoCreateSupplier = false
);

public record UpdatePurchaseOrderRequest(
    [property: JsonPropertyName("deliveryDate")] DateTime? DeliveryDate = null,
    [property: JsonPropertyName("reference")] string? Reference = null,
    [property: JsonPropertyName("notes")] string? Notes = null,
    [property: JsonPropertyName("items")] List<PurchaseOrderItemRequest>? Items = null
);

// ============================================================================
// Sales Order DTOs
// ============================================================================

public record SalesOrderItemResponse(
    [property: JsonPropertyName("lineNumber")] int LineNumber,
    [property: JsonPropertyName("stockCode")] string? StockCode,
    [property: JsonPropertyName("description")] string? Description,
    [property: JsonPropertyName("quantity")] decimal Quantity,
    [property: JsonPropertyName("unitPrice")] decimal UnitPrice,
    [property: JsonPropertyName("netAmount")] decimal NetAmount,
    [property: JsonPropertyName("taxCode")] string? TaxCode,
    [property: JsonPropertyName("taxAmount")] decimal TaxAmount,
    [property: JsonPropertyName("nominalCode")] string? NominalCode
);

public record SalesOrderResponse(
    [property: JsonPropertyName("orderNumber")] string OrderNumber,
    [property: JsonPropertyName("customerAccountRef")] string CustomerAccountRef,
    [property: JsonPropertyName("customerName")] string? CustomerName,
    [property: JsonPropertyName("orderDate")] DateTime? OrderDate,
    [property: JsonPropertyName("deliveryDate")] DateTime? DeliveryDate,
    [property: JsonPropertyName("customerOrderNumber")] string? CustomerOrderNumber,
    [property: JsonPropertyName("netAmount")] decimal NetAmount,
    [property: JsonPropertyName("taxAmount")] decimal TaxAmount,
    [property: JsonPropertyName("grossAmount")] decimal GrossAmount,
    [property: JsonPropertyName("status")] string? Status,
    [property: JsonPropertyName("notes")] string? Notes,
    [property: JsonPropertyName("items")] List<SalesOrderItemResponse>? Items = null
);

public record SalesOrderItemRequest(
    [property: JsonPropertyName("stockCode")] string? StockCode = null,
    [property: JsonPropertyName("description")] string? Description = null,
    [property: JsonPropertyName("quantity")] decimal Quantity = 1,
    [property: JsonPropertyName("unitPrice")] decimal UnitPrice = 0,
    [property: JsonPropertyName("nominalCode")] string NominalCode = "4000",
    [property: JsonPropertyName("taxCode")] string TaxCode = "T1"
);

public record CreateSalesOrderRequest(
    [property: JsonPropertyName("customerAccountRef")] string CustomerAccountRef,
    [property: JsonPropertyName("customerOrderNumber")] string? CustomerOrderNumber = null,
    [property: JsonPropertyName("orderDate")] DateTime? OrderDate = null,
    [property: JsonPropertyName("deliveryDate")] DateTime? DeliveryDate = null,
    [property: JsonPropertyName("notes")] string? Notes = null,
    [property: JsonPropertyName("items")] List<SalesOrderItemRequest>? Items = null,
    [property: JsonPropertyName("autoCreateCustomer")] bool AutoCreateCustomer = false
);

public record UpdateSalesOrderRequest(
    [property: JsonPropertyName("customerOrderNumber")] string? CustomerOrderNumber = null,
    [property: JsonPropertyName("deliveryDate")] DateTime? DeliveryDate = null,
    [property: JsonPropertyName("notes")] string? Notes = null
);

// ============================================================================
// Product DTOs
// ============================================================================

public record ProductResponse(
    [property: JsonPropertyName("stockCode")] string StockCode,
    [property: JsonPropertyName("description")] string? Description,
    [property: JsonPropertyName("salesPrice")] decimal SalesPrice,
    [property: JsonPropertyName("costPrice")] decimal CostPrice,
    [property: JsonPropertyName("qtyInStock")] decimal QtyInStock,
    [property: JsonPropertyName("qtyOnOrder")] decimal QtyOnOrder,
    [property: JsonPropertyName("qtyAllocated")] decimal QtyAllocated,
    [property: JsonPropertyName("reorderLevel")] decimal ReorderLevel,
    [property: JsonPropertyName("reorderQty")] decimal ReorderQty,
    [property: JsonPropertyName("nominalCode")] string? NominalCode,
    [property: JsonPropertyName("taxCode")] string? TaxCode,
    [property: JsonPropertyName("category")] string? Category,
    [property: JsonPropertyName("unitOfSale")] string? UnitOfSale
);

public record CreateProductRequest(
    [property: JsonPropertyName("stockCode")] string StockCode,
    [property: JsonPropertyName("description")] string? Description = null,
    [property: JsonPropertyName("salesPrice")] decimal SalesPrice = 0,
    [property: JsonPropertyName("costPrice")] decimal CostPrice = 0,
    [property: JsonPropertyName("reorderLevel")] decimal ReorderLevel = 0,
    [property: JsonPropertyName("reorderQty")] decimal ReorderQty = 0,
    [property: JsonPropertyName("nominalCode")] string NominalCode = "4000",
    [property: JsonPropertyName("taxCode")] string TaxCode = "T1",
    [property: JsonPropertyName("category")] string? Category = null,
    [property: JsonPropertyName("unitOfSale")] string? UnitOfSale = null
);

// ============================================================================
// Stock DTOs
// ============================================================================

public record StockLevelResponse(
    [property: JsonPropertyName("stockCode")] string StockCode,
    [property: JsonPropertyName("description")] string? Description,
    [property: JsonPropertyName("qtyInStock")] decimal QtyInStock,
    [property: JsonPropertyName("qtyOnOrder")] decimal QtyOnOrder,
    [property: JsonPropertyName("qtyAllocated")] decimal QtyAllocated,
    [property: JsonPropertyName("freeStock")] decimal FreeStock,
    [property: JsonPropertyName("costPrice")] decimal CostPrice
);

public record StockAdjustmentRequest(
    [property: JsonPropertyName("stockCode")] string StockCode,
    [property: JsonPropertyName("quantity")] decimal Quantity,
    [property: JsonPropertyName("adjustmentType")] string AdjustmentType = "AI",
    [property: JsonPropertyName("reference")] string? Reference = null,
    [property: JsonPropertyName("details")] string? Details = null,
    [property: JsonPropertyName("costPrice")] decimal? CostPrice = null
);

// ============================================================================
// System/Status DTOs
// ============================================================================

public record StatusResponse(
    [property: JsonPropertyName("status")] string Status,
    [property: JsonPropertyName("connected")] bool Connected,
    [property: JsonPropertyName("timestamp")] DateTime Timestamp,
    [property: JsonPropertyName("version")] string Version
);

public record VersionResponse(
    [property: JsonPropertyName("apiVersion")] string ApiVersion,
    [property: JsonPropertyName("sageVersion")] string? SageVersion,
    [property: JsonPropertyName("buildDate")] string BuildDate
);

public record SetupResponse(
    [property: JsonPropertyName("companyName")] string? CompanyName,
    [property: JsonPropertyName("financialYearStart")] DateTime? FinancialYearStart,
    [property: JsonPropertyName("financialYearEnd")] DateTime? FinancialYearEnd,
    [property: JsonPropertyName("currentPeriod")] int CurrentPeriod,
    [property: JsonPropertyName("vatRegistered")] bool VatRegistered,
    [property: JsonPropertyName("vatNumber")] string? VatNumber,
    [property: JsonPropertyName("vatScheme")] string? VatScheme,
    [property: JsonPropertyName("baseCurrency")] string? BaseCurrency,
    [property: JsonPropertyName("defaultTaxCode")] string? DefaultTaxCode
);

public record FinancialYearResponse(
    [property: JsonPropertyName("startDate")] DateTime StartDate,
    [property: JsonPropertyName("endDate")] DateTime EndDate,
    [property: JsonPropertyName("currentPeriod")] int CurrentPeriod,
    [property: JsonPropertyName("totalPeriods")] int TotalPeriods,
    [property: JsonPropertyName("periodStart")] DateTime? PeriodStart,
    [property: JsonPropertyName("periodEnd")] DateTime? PeriodEnd
);

// ============================================================================
// Financial Setup DTOs
// ============================================================================

public record TaxCodeResponse(
    [property: JsonPropertyName("code")] string Code,
    [property: JsonPropertyName("description")] string? Description,
    [property: JsonPropertyName("rate")] decimal Rate,
    [property: JsonPropertyName("isEcCode")] bool IsEcCode,
    [property: JsonPropertyName("inputNominal")] string? InputNominal,
    [property: JsonPropertyName("outputNominal")] string? OutputNominal
);

public record CurrencyResponse(
    [property: JsonPropertyName("code")] string Code,
    [property: JsonPropertyName("name")] string? Name,
    [property: JsonPropertyName("symbol")] string? Symbol,
    [property: JsonPropertyName("exchangeRate")] decimal ExchangeRate,
    [property: JsonPropertyName("isBase")] bool IsBase
);

public record DepartmentResponse(
    [property: JsonPropertyName("code")] string Code,
    [property: JsonPropertyName("name")] string? Name,
    [property: JsonPropertyName("number")] int Number
);

public record BankAccountResponse(
    [property: JsonPropertyName("nominalCode")] string NominalCode,
    [property: JsonPropertyName("name")] string? Name,
    [property: JsonPropertyName("accountNumber")] string? AccountNumber,
    [property: JsonPropertyName("sortCode")] string? SortCode,
    [property: JsonPropertyName("balance")] decimal Balance,
    [property: JsonPropertyName("bankName")] string? BankName
);

public record PaymentMethodResponse(
    [property: JsonPropertyName("code")] string Code,
    [property: JsonPropertyName("description")] string? Description,
    [property: JsonPropertyName("type")] string? Type
);

public record ChartOfAccountsResponse(
    [property: JsonPropertyName("nominalCode")] string NominalCode,
    [property: JsonPropertyName("name")] string? Name,
    [property: JsonPropertyName("type")] string? Type,
    [property: JsonPropertyName("category")] string? Category,
    [property: JsonPropertyName("balance")] decimal Balance,
    [property: JsonPropertyName("budgetBalance")] decimal BudgetBalance,
    [property: JsonPropertyName("priorYearBalance")] decimal PriorYearBalance
);

// ============================================================================
// Search/Ledger DTOs
// ============================================================================

public record LedgerTransactionResponse(
    [property: JsonPropertyName("transactionNumber")] int TransactionNumber,
    [property: JsonPropertyName("accountRef")] string AccountRef,
    [property: JsonPropertyName("type")] string? Type,
    [property: JsonPropertyName("date")] DateTime? Date,
    [property: JsonPropertyName("reference")] string? Reference,
    [property: JsonPropertyName("details")] string? Details,
    [property: JsonPropertyName("netAmount")] decimal NetAmount,
    [property: JsonPropertyName("taxAmount")] decimal TaxAmount,
    [property: JsonPropertyName("grossAmount")] decimal GrossAmount,
    [property: JsonPropertyName("outstanding")] decimal Outstanding,
    [property: JsonPropertyName("paid")] bool Paid
);

public record AgedDebtorResponse(
    [property: JsonPropertyName("accountRef")] string AccountRef,
    [property: JsonPropertyName("name")] string? Name,
    [property: JsonPropertyName("balance")] decimal Balance,
    [property: JsonPropertyName("current")] decimal Current,
    [property: JsonPropertyName("period1")] decimal Period1,
    [property: JsonPropertyName("period2")] decimal Period2,
    [property: JsonPropertyName("period3")] decimal Period3,
    [property: JsonPropertyName("older")] decimal Older,
    [property: JsonPropertyName("creditLimit")] decimal CreditLimit
);

public record AgedCreditorResponse(
    [property: JsonPropertyName("accountRef")] string AccountRef,
    [property: JsonPropertyName("name")] string? Name,
    [property: JsonPropertyName("balance")] decimal Balance,
    [property: JsonPropertyName("current")] decimal Current,
    [property: JsonPropertyName("period1")] decimal Period1,
    [property: JsonPropertyName("period2")] decimal Period2,
    [property: JsonPropertyName("period3")] decimal Period3,
    [property: JsonPropertyName("older")] decimal Older,
    [property: JsonPropertyName("creditLimit")] decimal CreditLimit
);

public record CustomerAddressResponse(
    [property: JsonPropertyName("addressId")] int AddressId,
    [property: JsonPropertyName("description")] string? Description,
    [property: JsonPropertyName("address1")] string? Address1,
    [property: JsonPropertyName("address2")] string? Address2,
    [property: JsonPropertyName("address3")] string? Address3,
    [property: JsonPropertyName("address4")] string? Address4,
    [property: JsonPropertyName("postcode")] string? Postcode,
    [property: JsonPropertyName("country")] string? Country,
    [property: JsonPropertyName("contactName")] string? ContactName,
    [property: JsonPropertyName("telephone")] string? Telephone,
    [property: JsonPropertyName("email")] string? Email,
    [property: JsonPropertyName("isDefault")] bool IsDefault
);

// ============================================================================
// Transaction/Payment DTOs
// ============================================================================

public record TransactionRecordResponse(
    [property: JsonPropertyName("transactionNumber")] int TransactionNumber,
    [property: JsonPropertyName("type")] string? Type,
    [property: JsonPropertyName("accountRef")] string? AccountRef,
    [property: JsonPropertyName("nominalCode")] string? NominalCode,
    [property: JsonPropertyName("date")] DateTime? Date,
    [property: JsonPropertyName("reference")] string? Reference,
    [property: JsonPropertyName("details")] string? Details,
    [property: JsonPropertyName("netAmount")] decimal NetAmount,
    [property: JsonPropertyName("taxAmount")] decimal TaxAmount,
    [property: JsonPropertyName("grossAmount")] decimal GrossAmount,
    [property: JsonPropertyName("taxCode")] string? TaxCode
);

public record BatchTransactionRequest(
    [property: JsonPropertyName("type")] string Type,
    [property: JsonPropertyName("accountRef")] string AccountRef,
    [property: JsonPropertyName("reference")] string? Reference = null,
    [property: JsonPropertyName("netAmount")] decimal NetAmount = 0,
    [property: JsonPropertyName("taxAmount")] decimal TaxAmount = 0,
    [property: JsonPropertyName("nominalCode")] string? NominalCode = null,
    [property: JsonPropertyName("bankNominal")] string? BankNominal = null,
    [property: JsonPropertyName("details")] string? Details = null,
    [property: JsonPropertyName("taxCode")] string? TaxCode = null
);

public record PostTransactionBatchRequest(
    [property: JsonPropertyName("transactions")] List<BatchTransactionRequest> Transactions
);

public record BatchTransactionResponse(
    [property: JsonPropertyName("successCount")] int SuccessCount,
    [property: JsonPropertyName("failCount")] int FailCount,
    [property: JsonPropertyName("results")] List<TransactionResponse> Results
);

public record AllocatePaymentRequest(
    [property: JsonPropertyName("accountRef")] string AccountRef,
    [property: JsonPropertyName("paymentReference")] string PaymentReference,
    [property: JsonPropertyName("invoiceReference")] string InvoiceReference,
    [property: JsonPropertyName("amount")] decimal Amount = 0
);

// ============================================================================
// Project DTOs
// ============================================================================

public record ProjectResponse(
    [property: JsonPropertyName("projectRef")] string ProjectRef,
    [property: JsonPropertyName("name")] string? Name,
    [property: JsonPropertyName("description")] string? Description,
    [property: JsonPropertyName("status")] string? Status,
    [property: JsonPropertyName("startDate")] DateTime? StartDate,
    [property: JsonPropertyName("endDate")] DateTime? EndDate,
    [property: JsonPropertyName("customerAccountRef")] string? CustomerAccountRef,
    [property: JsonPropertyName("customerName")] string? CustomerName,
    [property: JsonPropertyName("budgetCost")] decimal BudgetCost,
    [property: JsonPropertyName("actualCost")] decimal ActualCost,
    [property: JsonPropertyName("budgetRevenue")] decimal BudgetRevenue,
    [property: JsonPropertyName("actualRevenue")] decimal ActualRevenue,
    [property: JsonPropertyName("percentComplete")] decimal PercentComplete
);

public record CreateProjectRequest(
    [property: JsonPropertyName("projectRef")] string ProjectRef,
    [property: JsonPropertyName("name")] string? Name = null,
    [property: JsonPropertyName("description")] string? Description = null,
    [property: JsonPropertyName("startDate")] DateTime? StartDate = null,
    [property: JsonPropertyName("endDate")] DateTime? EndDate = null,
    [property: JsonPropertyName("customerAccountRef")] string? CustomerAccountRef = null,
    [property: JsonPropertyName("budgetCost")] decimal BudgetCost = 0,
    [property: JsonPropertyName("budgetRevenue")] decimal BudgetRevenue = 0
);

public record ProjectCostCodeResponse(
    [property: JsonPropertyName("costCode")] string CostCode,
    [property: JsonPropertyName("name")] string? Name,
    [property: JsonPropertyName("projectRef")] string? ProjectRef,
    [property: JsonPropertyName("budgetAmount")] decimal BudgetAmount,
    [property: JsonPropertyName("actualAmount")] decimal ActualAmount,
    [property: JsonPropertyName("variance")] decimal Variance
);
