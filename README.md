# Mini-Grocery Store POS & Daily Inventory System

A complete, standalone desktop Point-of-Sale (POS) and Daily Inventory Management System built in **Visual Basic .NET (VB.NET)** for Windows Forms, backed by a local **SQLite** database.

---

## 🏛️ Academic Information

- **Institution:** University of Caloocan City – Computer Studies Department
- **Course Code & Title:** ITP 102 – Object-Oriented Programming
- **Instructor:** Prof. Joble, Jayson P.
- **Proponents / Student Developers:**
  - Alon
  - Caboboy
  - Olpenda
  - Paguia

---

## 📋 Table of Contents

1. [Features by Role](#-features-by-role)
2. [Object-Oriented Programming (OOP) Implementation](#-object-oriented-programming-oop-implementation)
3. [Technology Stack & Prerequisites](#-technology-stack--prerequisites)
4. [Solution Architecture & Folder Structure](#-solution-architecture--folder-structure)
5. [Database Schema & Seed Data](#-database-schema--seed-data)
6. [Default Login Accounts](#-default-login-accounts)
7. [How to Open, Build, and Run](#-how-to-open-build-and-run)
8. [How to Edit Forms in Visual Studio Designer (Section 5A)](#-how-to-edit-forms-in-visual-studio-designer-section-5a)
9. [Database Location & Reset Instructions](#-database-location--reset-instructions)

---

## 🛒 Features by Role

### 1. Store Owner (Full System Control)
- **Summary Cards:** Real-time metrics on Today's Sales, Transactions Count, Low-Stock items count, and Expiring-Soon items count.
- **Staff Management:** Create, edit, deactivate, and reset passwords for Cashier and Inventory Clerk accounts.
- **Product Management (CRUD):** Add, edit, adjust prices/costs, set reorder levels, set expiry dates, and deactivate products.
- **Category & Supplier Management (CRUD):** Full maintenance of grocery categories and vendor/supplier directory.
- **Transaction History:** Search sales by receipt number, filter by date ranges, and inspect past itemized sale receipts.
- **Comprehensive Reporting:**
  - *Daily Sales Report:* Filter by date, totals, discounts, top-selling products, and cashier breakdown.
  - *Inventory Valuation Report:* Current stock counts, cost values, and retail values.
  - *Low Stock & Expiry Alerts:* Immediate visibility of urgent items.
  - *Export & Print:* Export report grids to CSV files and print formatted reports.

### 2. Cashier (POS & Checkout)
- **Product Lookup:** Search by item name or exact product code (e.g., `PRD001`).
- **Cart Management:** Add products, update quantities, delete line items, clear cart, with automated stock validation.
- **Polymorphic Discounts:**
  - *No Discount (0%)*
  - *Senior / PWD (20% statutory discount)*
  - *Regular Promo (5% configurable discount)*
- **Customer Linking:** Optionally attach a customer username so purchases record to their order history.
- **Transaction-Safe Checkout:** Atomic SQLite database transaction updates stock, logs stock movement records, and creates the sale header and line items.
- **Receipt Printing & Preview:** Built-in `PrintPreviewDialog` and `PrintDocument` showing store header, UCC department branding, line items, cash received, and change.
- **Shift Activity:** View today's personal transactions.

### 3. Inventory Clerk (Stock Operations & Monitoring)
- **Stock-In:** Record deliveries from suppliers, adjust unit quantities, and record batch expiration dates.
- **Stock Adjustment:** Correct discrepancies with mandatory audit reasons (*Damaged*, *Expired*, *Correction*).
- **Monitoring Dashboards:**
  - ⚠️ *Low Stock:* Items at or below reorder threshold (highlighted in Orange/Yellow).
  - ⏳ *Expiring Soon:* Items expiring within 30 days (highlighted in Orange).
  - ❌ *Expired:* Expired products flagged for disposal (highlighted in Red).
- **Stock Movement Ledger:** Complete date-filterable audit trail of stock ins, sales deductions, and adjustments.
- **Read-Only Catalog:** Browse product specifications without permission to alter retail prices or delete catalog items.

### 4. Customer Portal
- **Self-Registration:** Register an account directly from the login screen.
- **Live Inventory Catalog:** Browse available products with stock indicators (*In Stock*, *Low Stock*, *Out of Stock*).
- **Purchase History:** Review past transactions and double-click to view itemized receipts.
- **Profile Management:** Update personal name and change account password.

---

## 🧬 Object-Oriented Programming (OOP) Implementation

| OOP Concept | Application in Code | Location |
|---|---|---|
| **Abstraction** | `User` is an abstract base class (`MustInherit`) defining `GetRoleName()`, `GetDashboardForm()`, and `GetPermissions()`. Callers interact strictly through base abstractions without needing to know concrete types. `Discount` is an abstract strategy class with `MustOverride Function Apply()`. | `Models/User.vb`<br>`Models/Discount.vb` |
| **Inheritance** | `Customer`, `Cashier`, `InventoryClerk`, and `StoreOwner` inherit from `User`. `NoDiscount`, `RegularDiscount`, and `SeniorPwdDiscount` inherit from `Discount`. | `Models/Customer.vb`<br>`Models/Cashier.vb`<br>`Models/InventoryClerk.vb`<br>`Models/StoreOwner.vb`<br>`Models/Discount.vb` |
| **Polymorphism** | 1. **Dynamic Dashboard Routing:** After login, the application calls `authenticatedUser.GetDashboardForm()` to launch the correct UI without conditional `If/Else` or `Select Case` role branches.<br>2. **Polymorphic Pricing:** `CartService` calculates net totals via `CurrentDiscount.Apply(subtotal)`, dynamically swapping discount behaviors at runtime. | `Forms/LoginForm.vb`<br>`Services/CartService.vb` |
| **Encapsulation** | Private fields with validated public properties (rejecting negative prices, negative stock, blank names). Passwords are never stored or exposed in plain text; authenticated via PBKDF2 (10,000 iterations + 16-byte cryptographic salt). | All `Models/` classes<br>`Services/AuthService.vb` |
| **Interfaces** | `IReportable` contract defines `GetReportHeaders()` and `ToReportRow()` implemented by reporting models for uniform CSV exports. | `Models/IReportable.vb`<br>`Models/Sale.vb` |

---

## 🛠️ Technology Stack & Prerequisites

- **Language:** Visual Basic .NET (VB.NET)
- **Target Framework:** .NET Framework 4.8
- **IDE:** Microsoft Visual Studio 2022 (Community, Professional, or Enterprise)
- **Database Engine:** SQLite 3 via NuGet package `System.Data.SQLite.Core` (v1.0.119.0)
- **UI Framework:** Windows Forms (WinForms) with native Visual Studio 2022 Designer support

---

## 📁 Solution Architecture & Folder Structure

```
oopfinals/
├── oopfinals.sln                       # Solution file for Visual Studio 2022
├── nuget.exe                           # NuGet command-line executable
├── packages/                           # NuGet package dependencies
└── oopfinals/
    ├── oopfinals.vbproj                # Project configuration file
    ├── packages.config                 # SQLite package references
    ├── Models/                         # Domain Entities & OOP Hierarchies
    │   ├── IReportable.vb
    │   ├── Discount.vb
    │   ├── User.vb
    │   ├── Customer.vb
    │   ├── Cashier.vb
    │   ├── InventoryClerk.vb
    │   ├── StoreOwner.vb
    │   ├── Product.vb
    │   ├── Category.vb
    │   ├── Supplier.vb
    │   ├── Sale.vb
    │   ├── SaleItem.vb
    │   └── StockMovement.vb
    ├── Data/                           # Database Access Layer & Repositories
    │   ├── DatabaseHelper.vb           # Connection management, DDL schema & auto-seed
    │   ├── UserRepository.vb           # User persistence & credential queries
    │   ├── ProductRepository.vb        # Product queries, stock & expiry filters
    │   ├── CategoryRepository.vb       # Category management
    │   ├── SupplierRepository.vb       # Supplier management
    │   ├── SaleRepository.vb           # Transaction-wrapped sale recording
    │   └── StockRepository.vb          # Stock movement tracking
    ├── Services/                       # Business Logic Layer
    │   ├── AuthService.vb              # PBKDF2 hashing, salting, authentication
    │   ├── CartService.vb              # POS shopping cart calculation engine
    │   ├── ReportService.vb            # CSV export and data formatting
    │   └── ReceiptPrinter.vb           # PrintDocument receipt rendering engine
    └── Forms/                          # 3-File Windows Forms Partial Classes
        ├── LoginForm[.vb, .Designer.vb, .resx]
        ├── RegisterForm[.vb, .Designer.vb, .resx]
        ├── CustomerDashboardForm[.vb, .Designer.vb, .resx]
        ├── CashierDashboardForm[.vb, .Designer.vb, .resx]
        ├── ClerkDashboardForm[.vb, .Designer.vb, .resx]
        ├── OwnerDashboardForm[.vb, .Designer.vb, .resx]
        ├── ReceiptPreviewForm[.vb, .Designer.vb, .resx]
        └── PastReceiptForm[.vb, .Designer.vb, .resx]
```

---

## 🗄️ Database Schema & Seed Data

On the first application run, the SQLite database `grocery.db` is automatically created and populated inside a single database transaction.

### Tables
1. **Users:** `UserId`, `Username`, `PasswordHash`, `Salt`, `FullName`, `Role`, `IsActive`, `CreatedAt`
2. **Categories:** `CategoryId`, `Name` (6 default grocery categories)
3. **Suppliers:** `SupplierId`, `Name`, `ContactNumber`, `Address` (3 default suppliers)
4. **Products:** `ProductId`, `ProductCode`, `Name`, `CategoryId`, `SupplierId`, `Price`, `CostPrice`, `StockQty`, `ReorderLevel`, `ExpiryDate`, `IsActive` (30 realistic Philippine grocery products)
5. **Sales:** `SaleId`, `ReceiptNo`, `SaleDate`, `CashierId`, `CustomerId`, `Subtotal`, `DiscountType`, `DiscountAmount`, `Total`, `CashReceived`, `ChangeAmount`
6. **SaleItems:** `SaleItemId`, `SaleId`, `ProductId`, `Quantity`, `UnitPrice`, `LineTotal`
7. **StockMovements:** `MovementId`, `ProductId`, `MovementType`, `QuantityChange`, `Reason`, `SupplierId`, `UserId`, `MovementDate`

---

## 🔑 Default Login Accounts

| Role | Username | Password | Full Name | Access Level |
|---|---|---|---|---|
| **Store Owner** | `owner` | `owner123` | Maria Santos (Owner) | Full access to all modules, staff, settings, inventory, POS, and financial reports |
| **Cashier** | `cashier` | `cashier123` | Juan Dela Cruz (Cashier) | POS Screen, cart management, checkout, discount application, receipts |
| **Inventory Clerk** | `clerk` | `clerk123` | Pedro Penduko (Clerk) | Stock-in, adjustments, inventory monitoring, movement history, read-only catalog |
| **Customer** | `customer` | `customer123` | Ana Batumbakal (Customer) | Product catalog availability, order history, receipt viewing, profile update |

*Note: New customers can register at any time via the "Register as Customer" link on the login window.*

---

## 🚀 How to Open, Build, and Run

### Running in Visual Studio 2022
1. Open Visual Studio 2022.
2. Click **Open a project or solution**.
3. Select `oopfinals.sln`.
4. Visual Studio will automatically detect the packages in `packages.config`. If prompted, restore NuGet packages (or right-click the solution > **Restore NuGet Packages**).
5. Ensure the configuration is set to **Debug** or **Release** and platform to **Any CPU**.
6. Press <kbd>F5</kbd> (or <kbd>Ctrl</kbd> + <kbd>F5</kbd>) to start the application.

### Running via Command Line
If you prefer running via PowerShell:
```powershell
# Restore NuGet dependencies
.\nuget.exe restore oopfinals.sln

# Build the project
& "D:\Microsoft Visual Studio 18 Enterprise\MSBuild\Current\Bin\MSBuild.exe" oopfinals.sln /p:Configuration=Debug /v:m

# Launch the executable
.\oopfinals\bin\Debug\oopfinals.exe
```

---

## 🎨 How to Edit Forms in Visual Studio Designer (Section 5A)

Every form has been constructed to be **100% editable** in Visual Studio 2022's visual designer without any designer errors:

1. **Viewing Forms in Solution Explorer:**
   - In Visual Studio's **Solution Explorer**, expand the `Forms` folder.
   - Click the arrow next to any form (e.g., `OwnerDashboardForm.vb`) to see its grouped child files:
     - `OwnerDashboardForm.Designer.vb` (auto-generated control layout)
     - `OwnerDashboardForm.resx` (form resources)
2. **Visual Designer Mode:**
   - Right-click `FormName.vb` > choose **View Designer** (or press <kbd>Shift</kbd> + <kbd>F7</kbd>).
   - You can drag, resize, restyle, and inspect properties using the standard Visual Studio Toolbox and Properties panel.
3. **Code Mode:**
   - Right-click `FormName.vb` > choose **View Code** (or press <kbd>F7</kbd>).
4. **Important Designer Rules:**
   - Controls are declared as `Friend WithEvents` in `FormName.Designer.vb`.
   - Event handlers in `FormName.vb` use the standard `Handles ControlName.Event` syntax.
   - Do not rename controls inside the `.Designer.vb` file without also updating the corresponding `Handles` clause in `FormName.vb`.

---

## 💾 Database Location & Reset Instructions

- **Database File:** `grocery.db`
- **Location:**  
  `oopfinals/bin/Debug/grocery.db` (when running under Debug)  
  `oopfinals/bin/Release/grocery.db` (when running under Release)

### Resetting to Factory Seed Data
If you test sales, stock adjustments, or staff modifications and wish to return to the clean starting state:
1. Close the application.
2. Delete the `grocery.db` file in `oopfinals/bin/Debug/`.
3. Start the application again.  
The application will detect that the file is missing and automatically recreate all tables and re-seed the default 4 accounts, 6 categories, 3 suppliers, and 30 grocery products.
