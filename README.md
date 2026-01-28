# Sage 50 Accounts SDK Tools

C# console applications for interacting with Sage 50 Accounts (UK) via the SDO (Sage Data Objects) COM interface.

## Projects

### SageConnector

A console application that connects to Sage 50 Accounts and provides:

- **Read company information** - View company setup data (294 fields)
- **List customers** - Browse sales ledger accounts with balances
- **List invoices** - View recent invoice records with details
- **Check customer exists** - Lookup customer accounts by reference
- **Check supplier exists** - Lookup supplier accounts by reference
- **Create customers/suppliers** - Add accounts to ledgers (see Known Limitations)
- **Post sales invoices** - Create sales invoices via SopPost (see Known Limitations)

### SagePostViewer

A PE file analysis tool using PeNet to examine the Sage 50 SDK DLLs and discover available COM interfaces.

## Requirements

- **Sage 50 Accounts** installed with SDO Engine registered (SDOEngine.32)
- **.NET 9.0** or later (Windows)
- Access to a Sage 50 company data folder (e.g., `X:\ACCDATA`)

## Usage

### Basic Connection

```bash
# Connect and view data
SageConnector.exe "X:\ACCDATA" "manager" "password"
```

### Post a Test Invoice

```bash
# Post invoice to existing customer
SageConnector.exe "X:\ACCDATA" "manager" "password" post
```

### Create Customer and Post Invoice

```bash
# Create new customer and post invoice
SageConnector.exe "X:\ACCDATA" "manager" "password" newcust
```

### Explore SDK Types

```bash
# Explore the interop assembly types
SageConnector.exe explore
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

### Create a Customer

```csharp
sage.CreateCustomer(
    accountRef: "NEWCUST01",
    name: "New Customer Ltd",
    address1: "123 Business Street",
    address2: "Business Park",
    postcode: "AB1 2CD"
);
```

### Post a Sales Invoice

```csharp
var items = new List<InvoiceLineItem>
{
    new InvoiceLineItem
    {
        Description = "Professional Services",
        Quantity = 1,
        UnitPrice = 100.00m,
        NominalCode = "4000",  // Sales nominal
        TaxCode = "T1"         // Standard VAT
    }
};

var invoiceNum = sage.CreateSalesInvoice(
    customerAccount: "NEWCUST01",
    orderNumber: "PO-12345",
    items: items,
    createCustomerIfMissing: true,
    customerName: "New Customer Ltd"
);
```

## SDK Field Reference

### Common SalesRecord Fields
- `ACCOUNT_REF` - Customer account code
- `NAME` - Customer name
- `ADDRESS_1` to `ADDRESS_5` - Address lines (5 is typically postcode)
- `BALANCE` - Current balance

### Common InvoiceRecord Fields
- `INVOICE_NUMBER` - Invoice number
- `ACCOUNT_REF` - Customer account
- `INVOICE_DATE` - Invoice date
- `ITEMS_NET` - Net amount
- `INVOICE_TYPE_CODE` - Type (2 = Sales Invoice)

### SopPost Header Fields (101 total)
- `ACCOUNT_REF` - Customer account
- `ORDER_DATE` - Transaction date
- `CUST_ORDER_NUMBER` - Customer reference
- `ORDER_TYPE` - 0=Quote, 1=Proforma, 2=Sales Order/Invoice
- `INVOICE_NUMBER` - Generated invoice number (read-only after Update)

### SopPost Item Fields (47 total)
- `DESCRIPTION` - Line description
- `NET_AMOUNT` - Net value
- `FULL_NET_AMOUNT` - Full net amount
- `NOMINAL_CODE` - Nominal account code
- `TAX_CODE` - Tax code as integer (0=T0, 1=T1, 2=T2, etc.)
- `UNIT_PRICE` - Unit price
- `QTY_ORDER` - Quantity ordered
- `SERVICE_FLAG` - 1 for service items (no stock)

## Type Library

The `SdoEng.tlb` file can be used to generate a typed interop assembly:

```bash
tlbimp SdoEng.tlb /out:SageSDO.Interop.dll /namespace:SageSDO
```

## Known Limitations

### Write Operations
The SDK's write operations (creating customers, posting invoices) have limitations:

**Customer/Supplier Creation**: The `SalesRecord.Update()` method throws an exception. This may require:
- Opening records in a specific edit mode
- Additional SDK configuration
- Running as a specific user or with elevated permissions

**Invoice Posting**: The `SopPost.Update()` method returns `true` but may not create actual invoices. This appears to require:
- Stock records to be set up in Sage
- Additional SDK licensing or configuration
- A different posting object or mechanism

### What Works Reliably
All **read operations** work correctly:
- Connect to Sage 50
- Read company information (294 fields)
- List customers with balances
- List invoices with details
- Check if customer/supplier accounts exist

## Troubleshooting

### "Username is already in use"
The SDK only allows one connection per username. Wait a few seconds for the previous session to timeout, or ensure proper disconnection. The connector uses a unique timestamp-based ID to avoid conflicts.

### "Accounts application is busy"
Sage 50 is open on another machine. Close Sage or wait for other users to finish.

### "SDOEngine.32 not found"
Sage 50 is not installed or the COM components are not registered.

## License

MIT License - See LICENSE file for details.

## Contributing

Pull requests welcome. For major changes, please open an issue first.
