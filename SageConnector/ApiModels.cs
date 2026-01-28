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
