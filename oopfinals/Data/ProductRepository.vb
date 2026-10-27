' ==============================================================================
' Project: Mini-Grocery Store POS & Daily Inventory System
' Course: Object-Oriented Programming (ITP 102) - University of Caloocan City
' Class: ProductRepository
' OOP Concept: DATA ACCESS ENCAPSULATION
' Manages Product persistence, search, stock thresholds, and expiry queries
' using parameterized SQLite commands.
' ==============================================================================

Option Strict On
Option Explicit On

Imports System
Imports System.Collections.Generic
Imports System.Data.SQLite
Imports oopfinals.Models

Namespace Data
    Public Class ProductRepository
        Private Function MapProduct(reader As SQLiteDataReader) As Product
            Dim p As New Product()
            p.ProductId = Convert.ToInt32(reader("ProductId"))
            p.ProductCode = Convert.ToString(reader("ProductCode"))
            p.Name = Convert.ToString(reader("Name"))
            p.CategoryId = Convert.ToInt32(reader("CategoryId"))
            p.CategoryName = If(reader("CategoryName") IsNot DBNull.Value, Convert.ToString(reader("CategoryName")), "Uncategorized")
            p.SupplierId = Convert.ToInt32(reader("SupplierId"))
            p.SupplierName = If(reader("SupplierName") IsNot DBNull.Value, Convert.ToString(reader("SupplierName")), "Unknown Supplier")
            p.Price = Convert.ToDecimal(reader("Price"))
            p.CostPrice = Convert.ToDecimal(reader("CostPrice"))
            p.StockQty = Convert.ToInt32(reader("StockQty"))
            p.ReorderLevel = Convert.ToInt32(reader("ReorderLevel"))

            If reader("ExpiryDate") IsNot DBNull.Value AndAlso Not String.IsNullOrWhiteSpace(Convert.ToString(reader("ExpiryDate"))) Then
                Dim exp As DateTime
                If DateTime.TryParse(Convert.ToString(reader("ExpiryDate")), exp) Then
                    p.ExpiryDate = exp
                End If
            End If

            p.IsActive = Convert.ToInt32(reader("IsActive")) = 1
            Return p
        End Function

        Public Function GetAll(Optional activeOnly As Boolean = True) As List(Of Product)
            Dim list As New List(Of Product)()
            Using conn As SQLiteConnection = DatabaseHelper.GetConnection()
                Dim sql As String = "
                    SELECT p.ProductId, p.ProductCode, p.Name, p.CategoryId, c.Name AS CategoryName,
                           p.SupplierId, s.Name AS SupplierName, p.Price, p.CostPrice, p.StockQty,
                           p.ReorderLevel, p.ExpiryDate, p.IsActive
                    FROM Products p
                    LEFT JOIN Categories c ON p.CategoryId = c.CategoryId
                    LEFT JOIN Suppliers s ON p.SupplierId = s.SupplierId "

                If activeOnly Then
                    sql &= " WHERE p.IsActive = 1 "
                End If
                sql &= " ORDER BY p.Name;"

                Using cmd As New SQLiteCommand(sql, conn)
                    Using reader As SQLiteDataReader = cmd.ExecuteReader()
                        While reader.Read()
                            list.Add(MapProduct(reader))
                        End While
                    End Using
                End Using
            End Using
            Return list
        End Function

        Public Function Search(keyword As String, Optional categoryId As Nullable(Of Integer) = Nothing, Optional activeOnly As Boolean = True) As List(Of Product)
            Dim list As New List(Of Product)()
            Using conn As SQLiteConnection = DatabaseHelper.GetConnection()
                Dim sql As String = "
                    SELECT p.ProductId, p.ProductCode, p.Name, p.CategoryId, c.Name AS CategoryName,
                           p.SupplierId, s.Name AS SupplierName, p.Price, p.CostPrice, p.StockQty,
                           p.ReorderLevel, p.ExpiryDate, p.IsActive
                    FROM Products p
                    LEFT JOIN Categories c ON p.CategoryId = c.CategoryId
                    LEFT JOIN Suppliers s ON p.SupplierId = s.SupplierId
                    WHERE 1=1 "

                If activeOnly Then
                    sql &= " AND p.IsActive = 1 "
                End If

                If categoryId.HasValue AndAlso categoryId.Value > 0 Then
                    sql &= " AND p.CategoryId = @catId "
                End If

                If Not String.IsNullOrWhiteSpace(keyword) Then
                    sql &= " AND (LOWER(p.Name) LIKE @kw OR LOWER(p.ProductCode) LIKE @kw) "
                End If

                sql &= " ORDER BY p.Name;"

                Using cmd As New SQLiteCommand(sql, conn)
                    If categoryId.HasValue AndAlso categoryId.Value > 0 Then
                        cmd.Parameters.AddWithValue("@catId", categoryId.Value)
                    End If
                    If Not String.IsNullOrWhiteSpace(keyword) Then
                        cmd.Parameters.AddWithValue("@kw", "%" & keyword.Trim().ToLowerInvariant() & "%")
                    End If

                    Using reader As SQLiteDataReader = cmd.ExecuteReader()
                        While reader.Read()
                            list.Add(MapProduct(reader))
                        End While
                    End Using
                End Using
            End Using
            Return list
        End Function

        Public Function GetById(productId As Integer) As Product
            Using conn As SQLiteConnection = DatabaseHelper.GetConnection()
                Using cmd As New SQLiteCommand("
                    SELECT p.ProductId, p.ProductCode, p.Name, p.CategoryId, c.Name AS CategoryName,
                           p.SupplierId, s.Name AS SupplierName, p.Price, p.CostPrice, p.StockQty,
                           p.ReorderLevel, p.ExpiryDate, p.IsActive
                    FROM Products p
                    LEFT JOIN Categories c ON p.CategoryId = c.CategoryId
                    LEFT JOIN Suppliers s ON p.SupplierId = s.SupplierId
                    WHERE p.ProductId = @id
                    LIMIT 1;", conn)
                    cmd.Parameters.AddWithValue("@id", productId)
                    Using reader As SQLiteDataReader = cmd.ExecuteReader()
                        If reader.Read() Then
                            Return MapProduct(reader)
                        End If
                    End Using
                End Using
            End Using
            Return Nothing
        End Function

        Public Function GetByCode(code As String) As Product
            Using conn As SQLiteConnection = DatabaseHelper.GetConnection()
                Using cmd As New SQLiteCommand("
                    SELECT p.ProductId, p.ProductCode, p.Name, p.CategoryId, c.Name AS CategoryName,
                           p.SupplierId, s.Name AS SupplierName, p.Price, p.CostPrice, p.StockQty,
                           p.ReorderLevel, p.ExpiryDate, p.IsActive
                    FROM Products p
                    LEFT JOIN Categories c ON p.CategoryId = c.CategoryId
                    LEFT JOIN Suppliers s ON p.SupplierId = s.SupplierId
                    WHERE UPPER(p.ProductCode) = UPPER(@code)
                    LIMIT 1;", conn)
                    cmd.Parameters.AddWithValue("@code", code.Trim())
                    Using reader As SQLiteDataReader = cmd.ExecuteReader()
                        If reader.Read() Then
                            Return MapProduct(reader)
                        End If
                    End Using
                End Using
            End Using
            Return Nothing
        End Function

        Public Function Add(p As Product) As Integer
            Using conn As SQLiteConnection = DatabaseHelper.GetConnection()
                Using cmd As New SQLiteCommand("
                    INSERT INTO Products (ProductCode, Name, CategoryId, SupplierId, Price, CostPrice, StockQty, ReorderLevel, ExpiryDate, IsActive)
                    VALUES (@code, @name, @cat, @sup, @price, @cost, @qty, @reorder, @exp, @active);
                    SELECT last_insert_rowid();", conn)
                    cmd.Parameters.AddWithValue("@code", p.ProductCode)
                    cmd.Parameters.AddWithValue("@name", p.Name)
                    cmd.Parameters.AddWithValue("@cat", p.CategoryId)
                    cmd.Parameters.AddWithValue("@sup", p.SupplierId)
                    cmd.Parameters.AddWithValue("@price", p.Price)
                    cmd.Parameters.AddWithValue("@cost", p.CostPrice)
                    cmd.Parameters.AddWithValue("@qty", p.StockQty)
                    cmd.Parameters.AddWithValue("@reorder", p.ReorderLevel)
                    cmd.Parameters.AddWithValue("@exp", If(p.ExpiryDate.HasValue, CObj(p.ExpiryDate.Value.ToString("yyyy-MM-dd")), DBNull.Value))
                    cmd.Parameters.AddWithValue("@active", If(p.IsActive, 1, 0))

                    Dim newId As Integer = Convert.ToInt32(cmd.ExecuteScalar())
                    p.ProductId = newId
                    Return newId
                End Using
            End Using
        End Function

        Public Function Update(p As Product) As Boolean
            Using conn As SQLiteConnection = DatabaseHelper.GetConnection()
                Using cmd As New SQLiteCommand("
                    UPDATE Products
                    SET ProductCode = @code, Name = @name, CategoryId = @cat, SupplierId = @sup,
                        Price = @price, CostPrice = @cost, StockQty = @qty, ReorderLevel = @reorder,
                        ExpiryDate = @exp, IsActive = @active
                    WHERE ProductId = @id;", conn)
                    cmd.Parameters.AddWithValue("@code", p.ProductCode)
                    cmd.Parameters.AddWithValue("@name", p.Name)
                    cmd.Parameters.AddWithValue("@cat", p.CategoryId)
                    cmd.Parameters.AddWithValue("@sup", p.SupplierId)
                    cmd.Parameters.AddWithValue("@price", p.Price)
                    cmd.Parameters.AddWithValue("@cost", p.CostPrice)
                    cmd.Parameters.AddWithValue("@qty", p.StockQty)
                    cmd.Parameters.AddWithValue("@reorder", p.ReorderLevel)
                    cmd.Parameters.AddWithValue("@exp", If(p.ExpiryDate.HasValue, CObj(p.ExpiryDate.Value.ToString("yyyy-MM-dd")), DBNull.Value))
                    cmd.Parameters.AddWithValue("@active", If(p.IsActive, 1, 0))
                    cmd.Parameters.AddWithValue("@id", p.ProductId)

                    Return cmd.ExecuteNonQuery() > 0
                End Using
            End Using
        End Function

        Public Function UpdatePrice(productId As Integer, newPrice As Decimal, newCostPrice As Decimal) As Boolean
            Using conn As SQLiteConnection = DatabaseHelper.GetConnection()
                Using cmd As New SQLiteCommand("
                    UPDATE Products
                    SET Price = @price, CostPrice = @cost
                    WHERE ProductId = @id;", conn)
                    cmd.Parameters.AddWithValue("@price", newPrice)
                    cmd.Parameters.AddWithValue("@cost", newCostPrice)
                    cmd.Parameters.AddWithValue("@id", productId)
                    Return cmd.ExecuteNonQuery() > 0
                End Using
            End Using
        End Function

        Public Function DeleteOrDeactivate(productId As Integer) As Boolean
            Using conn As SQLiteConnection = DatabaseHelper.GetConnection()
                Using cmd As New SQLiteCommand("UPDATE Products SET IsActive = 0 WHERE ProductId = @id;", conn)
                    cmd.Parameters.AddWithValue("@id", productId)
                    Return cmd.ExecuteNonQuery() > 0
                End Using
            End Using
        End Function

        Public Function GetLowStockProducts() As List(Of Product)
            Dim list As New List(Of Product)()
            For Each p In GetAll(activeOnly:=True)
                If p.IsLowStock() Then
                    list.Add(p)
                End If
            Next
            Return list
        End Function

        Public Function GetExpiringProducts(Optional daysThreshold As Integer = 30) As List(Of Product)
            Dim list As New List(Of Product)()
            For Each p In GetAll(activeOnly:=True)
                If p.IsExpiringSoon(daysThreshold) Then
                    list.Add(p)
                End If
            Next
            Return list
        End Function

        Public Function GetExpiredProducts() As List(Of Product)
            Dim list As New List(Of Product)()
            For Each p In GetAll(activeOnly:=True)
                If p.IsExpired() Then
                    list.Add(p)
                End If
            Next
            Return list
        End Function
    End Class
End Namespace
