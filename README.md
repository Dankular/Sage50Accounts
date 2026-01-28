# Sage 50 Accounts SDK Tools

C# console application for interacting with Sage 50 Accounts (UK) via the SDO (Sage Data Objects) COM interface.

## SageConnector

A console application that connects to Sage 50 Accounts and provides:

- **Post sales invoices to ledger** - Direct posting via TransactionPost (updates customer balances)
- **Post purchase invoices** - Via TransactionPost (with limitations, see below)
- **Create customers/suppliers** - Add accounts to Sales/Purchase ledgers
- **Search accounts** - Find customers/suppliers by name or account reference
- **Read company information** - View company setup data (294 fields)
- **List customers/suppliers** - Browse ledger accounts with balances
- **List invoices** - View recent invoice records with details
- **Create invoice documents** - Via InvoicePost/SopPost (order processing system)

## Requirements

- **Sage 50 Accounts** installed with SDO Engine registered (SDOEngine.32)
- **.NET 9.0** or later (Windows x64)
- Access to a Sage 50 company data folder (e.g., `X:\ACCDATA`)

## Usage

### Basic Connection

```bash
# Connect and view company data
SageConnector.exe "X:\ACCDATA" "manager" "password"
```

### Available Commands

```bash
# Show all commands
SageConnector.exe "X:\ACCDATA" "manager" "password"

# Invoice Posting (updates ledger balances)
SageConnector.exe ... sinv      # Post sales invoice to FALCONLO
SageConnector.exe ... pinv      # Post purchase invoice

# Order Processing (creates documents only)
SageConnector.exe ... invdoc    # Create invoice document via InvoicePost
SageConnector.exe ... post      # Post via SopPost (Sales Order Processing)
SageConnector.exe ... purchase  # Post via PopPost (Purchase Order Processing)

# Account Management
SageConnector.exe ... newcust   # Create new test customer
SageConnector.exe ... newsup    # Create new test supplier
SageConnector.exe ... lookup    # Search customers/suppliers

# Nominal Ledger
SageConnector.exe ... nominals  # List nominal codes
SageConnector.exe ... journal   # Post journal entry
SageConnector.exe ... bankpay   # Post bank payment
SageConnector.exe ... bankrec   # Post bank receipt

# SDK Discovery
SageConnector.exe ... discover  # Discover available SDK posting objects
```

## Code Examples

### Connect to Sage 50

```csharp
using var sage = new SageConnection();
sage.Connect(@"X:\ACCDATA", "manager", "password");
sage.ShowCompanyInfo();
sage.ListCustomers(10);
sage.ListInvoices(10);
```

### Post a Sales Invoice (Updates Customer Balance)

```csharp
// This uses TransactionPost which actually updates the ledger
var success = sage.PostSalesInvoice(
    customerAccount: "CUST001",
    invoiceRef: "SI-001",
    netAmount: 100.00m,
    taxAmount: 20.00m,      // VAT at 20%
    nominalCode: "4000",    // Sales nominal
    details: "Professional Services",
    taxCode: "T1",          // Standard VAT
    postToLedger: true
);
// Customer balance increases by £120.00 (gross)
```

### Create a Customer

```csharp
var success = sage.CreateCustomerEx(
    accountRef: "NEWCUST1",   // Max 8 characters
    name: "New Customer Ltd",
    address1: "123 Business Street",
    address2: "Business Park",
    postcode: "AB1 2CD",
    telephone: "01onal234 567890",
    email: "info@customer.com"
);
```

### Search for Customers

```csharp
// Search by name
var customers = sage.FindCustomers("Falcon", 10);
foreach (var c in customers)
{
    Console.WriteLine($"{c.AccountRef}: {c.Name} - Balance: {c.Balance:C}");
}

// Get specific customer
var customer = sage.GetCustomer("FALCONLO");
if (customer != null)
{
    Console.WriteLine($"Balance: {customer.Balance:C}");
}
```

## SDK Posting Objects

### TransactionPost (Recommended for Ledger Posting)

The `TransactionPost` object is the correct way to post transactions that update ledger balances:

| TYPE | Code | Description |
|------|------|-------------|
| 1 | SI | Sales Invoice - Updates customer balance |
| 2 | SC | Sales Credit |
| 3 | SR | Sales Receipt |
| 4 | SA | Sales Adjustment |
| 5 | SD | Sales Discount |
| 6 | PI | Purchase Invoice - Updates supplier balance |
| 7 | PC | Purchase Credit |
| 8 | PP | Purchase Payment |
| 9 | PA | Purchase Adjustment |
| 10 | PD | Purchase Discount |

**Note:** Sales types (1-5) use T1 tax code, Purchase types (6-10) require T9 tax code.

**TransactionPost Header Fields** (66 total):
- `TYPE` - Transaction type (1-10, see TransType enum above)
- `ACCOUNT_REF` - Customer/Supplier account reference
- `DATE` - Transaction date
- `INV_REF` - Invoice reference
- `DETAILS` - Transaction description
- `NET_AMOUNT` - Net amount
- `TAX_AMOUNT` - VAT amount

**TransactionPost Item Fields** (47 total):
- `NOMINAL_CODE` - Nominal account code
- `NET_AMOUNT` - Line net amount
- `TAX_AMOUNT` - Line VAT amount
- `TAX_CODE` - Tax code (0=T0, 1=T1, etc.)
- `DETAILS` - Line description

### InvoicePost / SopPost (Order Processing)

These objects create invoice/order documents but do NOT update ledger balances:

- `InvoicePost` - Creates invoice documents
- `SopPost` - Sales Order Processing
- `PopPost` - Purchase Order Processing

Use these when you need to create order documents that will be posted later through Sage 50.

## Known Limitations

### Account References
- Maximum 8 characters
- No spaces allowed
- Must be unique within the ledger

### Purchase Invoice Posting
The `TransactionPost` with TYPE=6 (Purchase Invoice) works reliably:
- Accounts must exist as suppliers in the Purchase Ledger
- Use T9 tax code for all purchase transaction types (6-10)
- Multi-currency supplier accounts may cause "Foreign Currency mismatch" errors

### What Works Reliably

**Read Operations:**
- Connect to Sage 50
- Read company information (294 fields)
- List customers/suppliers with balances
- List invoices with details
- Search for accounts by name/reference

**Write Operations:**
- Create customer accounts
- Create supplier accounts
- Post sales invoices (via TransactionPost TYPE=1)
- Post purchase invoices (via TransactionPost TYPE=6)
- Create invoice documents (via InvoicePost/SopPost)

## Troubleshooting

### "Username is already in use"
The SDK only allows one connection per username. Wait a few seconds for the previous session to timeout, or ensure proper disconnection.

### "Accounts application is busy"
Sage 50 is open on another machine. Close Sage or wait for other users to finish.

### "SDOEngine.32 not found"
Sage 50 is not installed or the COM components are not registered.

### "Empty key specified" when creating accounts
Account reference exceeds 8 characters or contains invalid characters.

### "Foreign Currency mismatch"
The account is set up with a different currency. Ensure the transaction currency matches the account currency.

## Type Library

The `SdoEng.tlb` file can be used to generate a typed interop assembly:

```bash
tlbimp SdoEng.tlb /out:SageSDO.Interop.dll /namespace:SageSDO
```

## License

MIT License - See LICENSE file for details.

## Contributing

Pull requests welcome. For major changes, please open an issue first.
