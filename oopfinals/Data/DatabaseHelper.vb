' ==============================================================================
' Project: Mini-Grocery Store POS & Daily Inventory System
' Course: Object-Oriented Programming (ITP 102) - University of Caloocan City
' Class: DatabaseHelper
' OOP Concept: DATA ACCESS ABSTRACTION & ENCAPSULATION
' Manages the SQLite database connection, table schema creation, parameterized
' query execution, and automated seeding on first run.
' ==============================================================================

Option Strict On
Option Explicit On

Imports System
Imports System.Data
Imports System.Data.SQLite
Imports System.IO
Imports System.Windows.Forms
Imports oopfinals.Services

Namespace Data
    Public Class DatabaseHelper
        Public Shared ReadOnly DbFileName As String = "grocery.db"
        Public Shared Property DbFilePath As String = Path.Combine(AppDomain.CurrentDomain.BaseDirectory, DbFileName)

        Public Shared ReadOnly Property ConnectionString As String
            Get
                Return $"Data Source={DbFilePath};Version=3;Foreign Keys=True;"
            End Get
        End Property

        ''' <summary>
        ''' Gets an active, open SQLite database connection.
        ''' </summary>
        Public Shared Function GetConnection() As SQLiteConnection
            Dim conn As New SQLiteConnection(ConnectionString)
            conn.Open()
            Return conn
        End Function

        ''' <summary>
        ''' Initializes database tables and seed data if not yet created.
        ''' Automatically called during application startup.
        ''' </summary>
        Public Shared Sub InitializeDatabase()
            Try
                Dim isNewDb As Boolean = Not File.Exists(DbFilePath)

                Using conn As SQLiteConnection = GetConnection()
                    Using cmd As SQLiteCommand = conn.CreateCommand()
                        ' 1. Users table
                        cmd.CommandText = "
                            CREATE TABLE IF NOT EXISTS Users (
                                UserId INTEGER PRIMARY KEY AUTOINCREMENT,
                                Username TEXT UNIQUE NOT NULL,
                                PasswordHash TEXT NOT NULL,
                                Salt TEXT NOT NULL,
                                FullName TEXT NOT NULL,
                                Role TEXT NOT NULL,
                                IsActive INTEGER NOT NULL DEFAULT 1,
                                CreatedAt TEXT NOT NULL
                            );"
                        cmd.ExecuteNonQuery()

                        ' 2. Categories table
                        cmd.CommandText = "
                            CREATE TABLE IF NOT EXISTS Categories (
                                CategoryId INTEGER PRIMARY KEY AUTOINCREMENT,
                                Name TEXT UNIQUE NOT NULL
                            );"
                        cmd.ExecuteNonQuery()

                        ' 3. Suppliers table
                        cmd.CommandText = "
                            CREATE TABLE IF NOT EXISTS Suppliers (
                                SupplierId INTEGER PRIMARY KEY AUTOINCREMENT,
                                Name TEXT NOT NULL,
                                ContactNumber TEXT,
                                Address TEXT
                            );"
                        cmd.ExecuteNonQuery()

                        ' 4. Products table
                        cmd.CommandText = "
                            CREATE TABLE IF NOT EXISTS Products (
                                ProductId INTEGER PRIMARY KEY AUTOINCREMENT,
                                ProductCode TEXT UNIQUE NOT NULL,
                                Name TEXT NOT NULL,
                                CategoryId INTEGER NOT NULL,
                                SupplierId INTEGER NOT NULL,
                                Price REAL NOT NULL,
                                CostPrice REAL NOT NULL,
                                StockQty INTEGER NOT NULL,
                                ReorderLevel INTEGER NOT NULL,
                                ExpiryDate TEXT,
                                IsActive INTEGER NOT NULL DEFAULT 1,
                                FOREIGN KEY (CategoryId) REFERENCES Categories(CategoryId),
                                FOREIGN KEY (SupplierId) REFERENCES Suppliers(SupplierId)
                            );"
                        cmd.ExecuteNonQuery()

                        ' 5. Sales table
                        cmd.CommandText = "
                            CREATE TABLE IF NOT EXISTS Sales (
                                SaleId INTEGER PRIMARY KEY AUTOINCREMENT,
                                ReceiptNo TEXT UNIQUE NOT NULL,
                                SaleDate TEXT NOT NULL,
                                CashierId INTEGER NOT NULL,
                                CustomerId INTEGER,
                                Subtotal REAL NOT NULL,
                                DiscountType TEXT NOT NULL,
                                DiscountAmount REAL NOT NULL,
                                Total REAL NOT NULL,
                                CashReceived REAL NOT NULL,
                                ChangeAmount REAL NOT NULL,
                                FOREIGN KEY (CashierId) REFERENCES Users(UserId),
                                FOREIGN KEY (CustomerId) REFERENCES Users(UserId)
                            );"
                        cmd.ExecuteNonQuery()

                        ' 6. SaleItems table
                        cmd.CommandText = "
                            CREATE TABLE IF NOT EXISTS SaleItems (
                                SaleItemId INTEGER PRIMARY KEY AUTOINCREMENT,
                                SaleId INTEGER NOT NULL,
                                ProductId INTEGER NOT NULL,
                                Quantity INTEGER NOT NULL,
                                UnitPrice REAL NOT NULL,
                                LineTotal REAL NOT NULL,
                                FOREIGN KEY (SaleId) REFERENCES Sales(SaleId),
                                FOREIGN KEY (ProductId) REFERENCES Products(ProductId)
                            );"
                        cmd.ExecuteNonQuery()

                        ' 7. StockMovements table
                        cmd.CommandText = "
                            CREATE TABLE IF NOT EXISTS StockMovements (
                                MovementId INTEGER PRIMARY KEY AUTOINCREMENT,
                                ProductId INTEGER NOT NULL,
                                MovementType TEXT NOT NULL,
                                QuantityChange INTEGER NOT NULL,
                                Reason TEXT,
                                SupplierId INTEGER,
                                UserId INTEGER NOT NULL,
                                MovementDate TEXT NOT NULL,
                                FOREIGN KEY (ProductId) REFERENCES Products(ProductId),
                                FOREIGN KEY (SupplierId) REFERENCES Suppliers(SupplierId),
                                FOREIGN KEY (UserId) REFERENCES Users(UserId)
                            );"
                        cmd.ExecuteNonQuery()
                    End Using

                    ' Check if seeding is required
                    SeedInitialDataIfEmpty(conn)
                End Using
            Catch ex As Exception
                If Application.MessageLoop Then
                    MessageBox.Show($"Failed to initialize database: {ex.Message}", "Database Error", MessageBoxButtons.OK, MessageBoxIcon.Error)
                End If
                Throw
            End Try
        End Sub

        ''' <summary>
        ''' Populates initial seed data (Users, Categories, Suppliers, 30 Products) inside a transaction.
        ''' </summary>
        Private Shared Sub SeedInitialDataIfEmpty(conn As SQLiteConnection)
            Using checkCmd As New SQLiteCommand("SELECT COUNT(*) FROM Users;", conn)
                Dim userCount As Long = Convert.ToInt64(checkCmd.ExecuteScalar())
                If userCount > 0 Then
                    Return ' Database already populated
                End If
            End Using

            Using trans As SQLiteTransaction = conn.BeginTransaction()
                Try
                    ' 1. Seed Users (1 Owner, 1 Cashier, 1 Clerk, 1 Customer)
                    Dim usersToSeed As New List(Of SeedUser) From {
                        New SeedUser("owner", "owner123", "Maria Santos (Owner)", "StoreOwner"),
                        New SeedUser("cashier", "cashier123", "Juan Dela Cruz (Cashier)", "Cashier"),
                        New SeedUser("clerk", "clerk123", "Pedro Penduko (Clerk)", "InventoryClerk"),
                        New SeedUser("customer", "customer123", "Ana Batumbakal (Customer)", "Customer")
                    }

                    For Each u In usersToSeed
                        Dim salt As String = AuthService.GenerateSalt()
                        Dim hash As String = AuthService.HashPassword(u.P, salt)

                        Using cmd As New SQLiteCommand("
                            INSERT INTO Users (Username, PasswordHash, Salt, FullName, Role, IsActive, CreatedAt)
                            VALUES (@user, @hash, @salt, @name, @role, 1, @created);", conn, trans)
                            cmd.Parameters.AddWithValue("@user", u.U)
                            cmd.Parameters.AddWithValue("@hash", hash)
                            cmd.Parameters.AddWithValue("@salt", salt)
                            cmd.Parameters.AddWithValue("@name", u.N)
                            cmd.Parameters.AddWithValue("@role", u.R)
                            cmd.Parameters.AddWithValue("@created", DateTime.Now.ToString("yyyy-MM-dd HH:mm:ss"))
                            cmd.ExecuteNonQuery()
                        End Using
                    Next

                    ' 2. Seed 6 Categories
                    Dim categories As String() = {"Rice & Grains", "Canned Goods", "Noodles & Pasta", "Beverages & Drinks", "Snacks & Biscuits", "Personal Care & Toiletries"}
                    For Each cat In categories
                        Using cmd As New SQLiteCommand("INSERT INTO Categories (Name) VALUES (@name);", conn, trans)
                            cmd.Parameters.AddWithValue("@name", cat)
                            cmd.ExecuteNonQuery()
                        End Using
                    Next

                    ' 3. Seed 3 Suppliers
                    Dim suppliers As New List(Of SeedSupplier) From {
                        New SeedSupplier("San Miguel Distribution", "0917-123-4567", "EDSA, Mandaluyong City"),
                        New SeedSupplier("Universal Robina Logistics", "0918-234-5678", "Bagong Ilog, Pasig City"),
                        New SeedSupplier("Mega Global Supply Co.", "0922-345-6789", "Valenzuela City, Metro Manila")
                    }
                    For Each sup In suppliers
                        Using cmd As New SQLiteCommand("INSERT INTO Suppliers (Name, ContactNumber, Address) VALUES (@n, @c, @a);", conn, trans)
                            cmd.Parameters.AddWithValue("@n", sup.Name)
                            cmd.Parameters.AddWithValue("@c", sup.Contact)
                            cmd.Parameters.AddWithValue("@a", sup.Address)
                            cmd.ExecuteNonQuery()
                        End Using
                    Next

                    ' 4. Seed 30 Realistic Grocery Products with Philippine Peso prices
                    ' Category IDs: 1=Rice, 2=Canned, 3=Noodles, 4=Beverages, 5=Snacks, 6=Personal Care
                    ' Supplier IDs: 1=San Miguel, 2=Universal Robina, 3=Mega Global
                    Dim products As List(Of SeedProduct) = GetSeedProducts()

                    Dim prodIdCounter As Integer = 1
                    For Each p In products
                        Using cmd As New SQLiteCommand("
                            INSERT INTO Products (ProductCode, Name, CategoryId, SupplierId, Price, CostPrice, StockQty, ReorderLevel, ExpiryDate, IsActive)
                            VALUES (@code, @name, @cat, @sup, @price, @cost, @qty, @reorder, @exp, 1);", conn, trans)
                            cmd.Parameters.AddWithValue("@code", p.Code)
                            cmd.Parameters.AddWithValue("@name", p.Name)
                            cmd.Parameters.AddWithValue("@cat", p.Cat)
                            cmd.Parameters.AddWithValue("@sup", p.Sup)
                            cmd.Parameters.AddWithValue("@price", p.Price)
                            cmd.Parameters.AddWithValue("@cost", p.Cost)
                            cmd.Parameters.AddWithValue("@qty", p.Qty)
                            cmd.Parameters.AddWithValue("@reorder", p.Reorder)
                            cmd.Parameters.AddWithValue("@exp", p.Exp.ToString("yyyy-MM-dd"))
                            cmd.ExecuteNonQuery()
                        End Using

                        ' Record initial stock movement audit entry
                        Using moveCmd As New SQLiteCommand("
                            INSERT INTO StockMovements (ProductId, MovementType, QuantityChange, Reason, SupplierId, UserId, MovementDate)
                            VALUES (@pid, 'StockIn', @qty, 'Initial Store Inventory Stock-In', @sup, 1, @dt);", conn, trans)
                            moveCmd.Parameters.AddWithValue("@pid", prodIdCounter)
                            moveCmd.Parameters.AddWithValue("@qty", p.Qty)
                            moveCmd.Parameters.AddWithValue("@sup", p.Sup)
                            moveCmd.Parameters.AddWithValue("@dt", DateTime.Now.ToString("yyyy-MM-dd HH:mm:ss"))
                            moveCmd.ExecuteNonQuery()
                        End Using

                        prodIdCounter += 1
                    Next

                    trans.Commit()
                Catch
                    trans.Rollback()
                    Throw
                End Try
            End Using
        End Sub

        Private Shared Function GetSeedProducts() As List(Of SeedProduct)
            Dim list As New List(Of SeedProduct)()
            list.Add(New SeedProduct("PRD001", "Dinorado Special Rice 5kg", 1, 1, 280.0D, 240.0D, 45, 10, DateTime.Today.AddMonths(12)))
            list.Add(New SeedProduct("PRD002", "Sinandomeng Rice 5kg", 1, 1, 250.0D, 215.0D, 8, 10, DateTime.Today.AddMonths(10)))
            list.Add(New SeedProduct("PRD003", "Jasmine Fragrant Rice 1kg", 1, 1, 62.0D, 52.0D, 30, 15, DateTime.Today.AddMonths(8)))
            list.Add(New SeedProduct("PRD004", "Brown Rice Healthy 1kg", 1, 1, 75.0D, 60.0D, 5, 10, DateTime.Today.AddDays(15)))
            list.Add(New SeedProduct("PRD005", "Mega Sardines in Tomato Sauce 155g", 2, 3, 24.5D, 20.0D, 120, 25, DateTime.Today.AddYears(2)))
            list.Add(New SeedProduct("PRD006", "Mega Sardines with Chili 155g", 2, 3, 25.0D, 20.5D, 90, 20, DateTime.Today.AddYears(2)))
            list.Add(New SeedProduct("PRD007", "Purefoods Corned Beef 150g", 2, 1, 68.0D, 58.0D, 50, 15, DateTime.Today.AddMonths(18)))
            list.Add(New SeedProduct("PRD008", "Argentina Meat Loaf 150g", 2, 1, 28.0D, 23.0D, 4, 15, DateTime.Today.AddMonths(6)))
            list.Add(New SeedProduct("PRD009", "San Marino Corned Tuna 180g", 2, 3, 42.0D, 35.0D, 65, 20, DateTime.Today.AddMonths(14)))
            list.Add(New SeedProduct("PRD010", "Century Tuna Flakes in Oil 155g", 2, 3, 38.5D, 32.0D, 12, 20, DateTime.Today.AddDays(20)))
            list.Add(New SeedProduct("PRD011", "Lucky Me! Instant Pancit Canton Kalamansi 80g", 3, 2, 16.0D, 13.0D, 150, 30, DateTime.Today.AddMonths(8)))
            list.Add(New SeedProduct("PRD012", "Lucky Me! Pancit Canton Chilimansi 80g", 3, 2, 16.0D, 13.0D, 140, 30, DateTime.Today.AddMonths(8)))
            list.Add(New SeedProduct("PRD013", "Lucky Me! Beef Mami Noodle Soup 55g", 3, 2, 14.5D, 11.5D, 80, 25, DateTime.Today.AddMonths(7)))
            list.Add(New SeedProduct("PRD014", "Payless Extra Big Pancit Canton 130g", 3, 2, 21.0D, 17.0D, 6, 15, DateTime.Today.AddMonths(5)))
            list.Add(New SeedProduct("PRD015", "Royal Spaghetti Pasta 900g", 3, 2, 95.0D, 78.0D, 25, 10, DateTime.Today.AddMonths(12)))
            list.Add(New SeedProduct("PRD016", "Coca-Cola Original Taste 1.5L", 4, 1, 75.0D, 62.0D, 40, 15, DateTime.Today.AddMonths(6)))
            list.Add(New SeedProduct("PRD017", "Royal Tru-Orange 1.5L", 4, 1, 72.0D, 60.0D, 35, 15, DateTime.Today.AddMonths(6)))
            list.Add(New SeedProduct("PRD018", "Sprite Lemon-Lime 1.5L", 4, 1, 72.0D, 60.0D, 3, 10, DateTime.Today.AddMonths(5)))
            list.Add(New SeedProduct("PRD019", "C2 Green Tea Apple 500ml", 4, 2, 26.0D, 20.0D, 55, 20, DateTime.Today.AddDays(10)))
            list.Add(New SeedProduct("PRD020", "San Miguel Pale Pilsen Can 330ml", 4, 1, 55.0D, 45.0D, 48, 12, DateTime.Today.AddMonths(9)))
            list.Add(New SeedProduct("PRD021", "Piattos Cheese Potato Crisps 85g", 5, 2, 38.0D, 31.0D, 60, 20, DateTime.Today.AddMonths(6)))
            list.Add(New SeedProduct("PRD022", "Nova Country Cheddar Multigrain 78g", 5, 2, 37.0D, 30.0D, 50, 15, DateTime.Today.AddMonths(6)))
            list.Add(New SeedProduct("PRD023", "Chippy Barbecue Corn Chips 110g", 5, 2, 27.0D, 22.0D, 7, 20, DateTime.Today.AddMonths(4)))
            list.Add(New SeedProduct("PRD024", "SkyFlakes Crackers Box 24 Packs", 5, 2, 72.0D, 58.0D, 30, 10, DateTime.Today.AddMonths(10)))
            list.Add(New SeedProduct("PRD025", "Rebisco Sandwich Choco Crackers 10 Packs", 5, 2, 65.0D, 52.0D, 28, 10, DateTime.Today.AddDays(-5)))
            list.Add(New SeedProduct("PRD026", "Safeguard Pure White Bar Soap 130g", 6, 1, 48.0D, 38.0D, 70, 20, DateTime.Today.AddYears(3)))
            list.Add(New SeedProduct("PRD027", "Palmolive Naturals Shampoo 180ml", 6, 1, 115.0D, 95.0D, 25, 10, DateTime.Today.AddYears(2)))
            list.Add(New SeedProduct("PRD028", "Colgate Maximum Cavity Toothpaste 145g", 6, 1, 98.0D, 80.0D, 35, 15, DateTime.Today.AddYears(2)))
            list.Add(New SeedProduct("PRD029", "Sunsilk Smooth & Manageable Sachet 12s", 6, 1, 72.0D, 60.0D, 5, 15, DateTime.Today.AddMonths(12)))
            list.Add(New SeedProduct("PRD030", "Head & Shoulders Cool Menthol Sachet 12s", 6, 1, 84.0D, 70.0D, 2, 15, DateTime.Today.AddDays(25)))
            Return list
        End Function

        Private Structure SeedUser
            Public U As String
            Public P As String
            Public N As String
            Public R As String
            Public Sub New(username As String, password As String, fullName As String, role As String)
                U = username : P = password : N = fullName : R = role
            End Sub
        End Structure

        Private Structure SeedSupplier
            Public Name As String
            Public Contact As String
            Public Address As String
            Public Sub New(n As String, c As String, a As String)
                Name = n : Contact = c : Address = a
            End Sub
        End Structure

        Private Structure SeedProduct
            Public Code As String
            Public Name As String
            Public Cat As Integer
            Public Sup As Integer
            Public Price As Decimal
            Public Cost As Decimal
            Public Qty As Integer
            Public Reorder As Integer
            Public Exp As DateTime
            Public Sub New(c As String, n As String, catId As Integer, supId As Integer, pr As Decimal, co As Decimal, q As Integer, re As Integer, ex As DateTime)
                Code = c : Name = n : Cat = catId : Sup = supId : Price = pr : Cost = co : Qty = q : Reorder = re : Exp = ex
            End Sub
        End Structure
    End Class
End Namespace
