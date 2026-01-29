# Sage 50 Accounts SDK Tools

C# console application and REST API for interacting with Sage 50 Accounts (UK) via the SDO (Sage Data Objects) COM interface.

**No Sage 50 Accounts SDO installation required** - The SDK is automatically downloaded and loaded without COM registration.

## Features

- **REST API Server** - HTTP API with Swagger/OpenAPI documentation
- **Post sales invoices to ledger** - Direct posting via TransactionPost (updates customer balances)
- **Post purchase invoices** - Via TransactionPost (with limitations, see below)
- **Create customers/suppliers** - Add accounts to Sales/Purchase ledgers
- **Search accounts** - Find customers/suppliers by name or account reference
- **Read company information** - View company setup data
- **List customers/suppliers** - Browse ledger accounts with balances
- **Post journals and bank transactions** - Direct nominal ledger posting

## Requirements

- **.NET 9.0** or later (Windows x64)
- Access to a Sage 50 company data folder (e.g., `X:\ACCDATA`)

**No Sage 50 Accounts SDO installation needed** - The SDK Manager automatically:
1. Detects the Sage version from ACCDATA
2. Downloads the matching SDK from Sage KB
3. Extracts all dependencies from the MSI
4. Loads the SDK using registration-free COM

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
SageConnector.exe ... sinv              # Post sales invoice
SageConnector.exe ... sinv CUST01       # Post to specific customer
SageConnector.exe ... sinv CUST01 -auto # Auto-create customer if missing
SageConnector.exe ... pinv              # Post purchase invoice
SageConnector.exe ... pinv SUP01 -auto  # Auto-create supplier if missing

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

### SDK Management

The SDK Manager automatically detects the Sage 50 version from ACCDATA and downloads the appropriate SDK. No COM registration or admin rights required.

```bash
# Check SDK status
SageConnector.exe sdk status

# List available SDK versions (v26-v33)
SageConnector.exe sdk list

# Detect version from ACCDATA
SageConnector.exe sdk detect "X:\ACCDATA"

# Download specific SDK version
SageConnector.exe sdk download 32.0

# Auto-detect, download and extract SDK
SageConnector.exe sdk install "X:\ACCDATA"

# Test registration-free COM loading
SageConnector.exe sdk test "X:\ACCDATA"
```

**How it works:**
1. Detects Sage version from ACCSTAT.DTA schema GUID
2. Downloads SDK ZIP from official Sage KB article
3. Extracts MSI from ZIP to get all 31 dependency DLLs
4. Loads `sg50SdoEngine.dll` via `LoadLibrary`
5. Creates COM objects via `DllGetClassObject` (no registration)

SDK files are cached in `%LOCALAPPDATA%\SageConnector\SDK\`.

## HTTP API Server

Start the REST API server to access Sage 50 via HTTP:

```bash
# Start API server on default port 5000
SageConnector.exe serve "X:\ACCDATA" manager password

# Start on custom port
SageConnector.exe serve "X:\ACCDATA" manager password --port=8080
```

### API Endpoints

#### Documentation
| Method | Endpoint | Description |
|--------|----------|-------------|
| GET | `/api/swagger.json` | OpenAPI 3.0 specification |

#### System & Status
| Method | Endpoint | Description |
|--------|----------|-------------|
| GET | `/api/status` | Health check and connection status |
| GET | `/api/version` | Sage version information |
| GET | `/api/company` | Get company information |

#### Customers
| Method | Endpoint | Description |
|--------|----------|-------------|
| GET | `/api/customers` | List customers (?search=&limit=) |
| GET | `/api/customers/{ref}` | Get customer by account ref |
| POST | `/api/customers` | Create new customer |
| GET | `/api/customers/{ref}/exists` | Check if customer exists |
| GET | `/api/customers/{ref}/addresses` | Get customer delivery addresses |

#### Suppliers
| Method | Endpoint | Description |
|--------|----------|-------------|
| GET | `/api/suppliers` | List suppliers (?search=&limit=) |
| GET | `/api/suppliers/{ref}` | Get supplier by account ref |
| POST | `/api/suppliers` | Create new supplier |
| GET | `/api/suppliers/{ref}/exists` | Check if supplier exists |

#### Products & Stock
| Method | Endpoint | Description |
|--------|----------|-------------|
| GET | `/api/products` | List products (?search=&limit=) |
| GET | `/api/products/{code}` | Get product by code |
| POST | `/api/products` | Create new product |
| GET | `/api/products/{code}/exists` | Check if product exists |
| GET | `/api/products/{code}/stock` | Get stock levels |
| PUT | `/api/products/{code}/stock` | Update stock quantity |

#### Financial Setup
| Method | Endpoint | Description |
|--------|----------|-------------|
| GET | `/api/setup` | Get company setup/preferences |
| GET | `/api/setup/financialyear` | Get financial year dates |
| GET | `/api/taxcodes` | List tax codes |
| GET | `/api/currencies` | List currencies |
| GET | `/api/departments` | List departments |
| GET | `/api/banks` | List bank accounts |
| GET | `/api/paymentmethods` | List payment methods |
| GET | `/api/chartofaccounts` | Full chart of accounts |

#### Nominal Ledger
| Method | Endpoint | Description |
|--------|----------|-------------|
| GET | `/api/nominals` | List nominal codes |
| POST | `/api/nominals` | Create nominal code |
| GET | `/api/nominals/{code}/exists` | Check if nominal exists |

#### Sales Transactions
| Method | Endpoint | Description |
|--------|----------|-------------|
| POST | `/api/sales/invoice` | Post sales invoice to ledger |
| POST | `/api/sales/credit` | Post sales credit note |
| POST | `/api/sales/receipt` | Post sales receipt/payment |
| POST | `/api/sales/order` | Create sales order |
| GET | `/api/sales/orders` | List sales orders |
| GET | `/api/sales/ledger` | Search sales ledger (?account=&from=&to=) |

#### Purchase Transactions
| Method | Endpoint | Description |
|--------|----------|-------------|
| POST | `/api/purchases/invoice` | Post purchase invoice |
| POST | `/api/purchases/credit` | Post purchase credit note |
| POST | `/api/purchases/payment` | Post purchase payment |
| POST | `/api/purchases/order` | Create purchase order |
| GET | `/api/purchases/orders` | List purchase orders |
| GET | `/api/purchases/ledger` | Search purchase ledger |

#### Bank & Journals
| Method | Endpoint | Description |
|--------|----------|-------------|
| POST | `/api/bank/payment` | Post bank payment |
| POST | `/api/bank/receipt` | Post bank receipt |
| POST | `/api/journals` | Post journal entry |
| POST | `/api/journals/simple` | Post simple two-line journal |

#### Aged Analysis
| Method | Endpoint | Description |
|--------|----------|-------------|
| GET | `/api/debtors` | Aged debtors report |
| GET | `/api/creditors` | Aged creditors report |

#### Transactions
| Method | Endpoint | Description |
|--------|----------|-------------|
| GET | `/api/transactions` | Query transactions (?type=&account=&from=&to=) |
| POST | `/api/allocate` | Allocate payment to invoice |

#### Projects
| Method | Endpoint | Description |
|--------|----------|-------------|
| GET | `/api/projects` | List projects (?search=&status=) |
| GET | `/api/projects/{ref}` | Get project by reference |
| POST | `/api/projects` | Create new project |
| GET | `/api/projects/{ref}/costcodes` | Get project cost codes |

### Example API Calls

```bash
# Get company info
curl http://localhost:5000/api/company

# List customers
curl http://localhost:5000/api/customers?search=falcon&limit=10

# Create a customer
curl -X POST http://localhost:5000/api/customers \
  -H "Content-Type: application/json" \
  -d '{"accountRef":"TEST01","name":"Test Customer Ltd","postcode":"AB1 2CD"}'

# Post a sales invoice
curl -X POST http://localhost:5000/api/sales/invoice \
  -H "Content-Type: application/json" \
  -d '{"customerAccount":"TEST01","netAmount":100,"taxAmount":20,"nominalCode":"4000"}'

# Get OpenAPI spec
curl http://localhost:5000/api/swagger.json
```

### Response Format

All responses use a consistent JSON format:

```json
{
  "success": true,
  "data": { ... },
  "error": null
}
```

On error:
```json
{
  "success": false,
  "data": null,
  "error": "Error message here"
}
```

## Test Suite

A comprehensive test suite is included to validate all API endpoints:

```bash
# Build and run tests (requires API server running)
dotnet build SageConnector.Tests
dotnet run --project SageConnector.Tests
```

The test suite includes 58 tests covering:
- System/Status endpoints
- Company information
- Customer CRUD operations
- Supplier CRUD operations
- Nominal codes
- Financial setup (tax codes, currencies, departments)
- Products and stock management
- Sales and purchase orders
- Sales and purchase transactions
- Bank payments and receipts
- Journal entries
- Ledger search and aged analysis
- Transaction queries
- Projects

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
// Customer balance increases by 120.00 (gross)
```

### Create a Customer

```csharp
var success = sage.CreateCustomerEx(
    accountRef: "NEWCUST1",   // Max 8 characters
    name: "New Customer Ltd",
    address1: "123 Business Street",
    address2: "Business Park",
    postcode: "AB1 2CD",
    telephone: "01234 567890",
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

## Troubleshooting

### "Username is already in use"
The SDK only allows one connection per username. Wait a few seconds for the previous session to timeout, or ensure proper disconnection.

### "Accounts application is busy"
Sage 50 is open on another machine. Close Sage or wait for other users to finish.

### "No Sage SDO Engine found"
Run `sdk install "X:\ACCDATA"` to download and extract the SDK automatically.

### "Empty key specified" when creating accounts
Account reference exceeds 8 characters or contains invalid characters.

### "Foreign Currency mismatch"
The account is set up with a different currency. Ensure the transaction currency matches the account currency.

## Version Detection

The SDK Manager detects Sage version using multiple methods:

1. **ACCSTAT.DTA GUID** - Schema identifier unique to each major version
2. **TABLEMETADATA.DTA** - Format version in file header
3. **Path inference** - Year from folder structure (e.g., `Accounts/2024` = v32)

Known schema GUIDs:
- `5D3EB135-3317-413B-99DE-47C6B044134D` = v32.0

## License

MIT License - See LICENSE file for details.

## Contributing

Pull requests welcome. For major changes, please open an issue first.
