' ==============================================================================
' Project: Mini-Grocery Store POS & Daily Inventory System
' Course: Object-Oriented Programming (ITP 102) - University of Caloocan City
' Class: SaleRepository
' OOP Concept: TRANSACTION MANAGEMENT & DATA ACCESS
' Manages POS checkout recording inside an atomic SQLite transaction,
' ensuring Sales, SaleItems, Stock updates, and StockMovements stay consistent.
' ==============================================================================

Option Strict On
Option Explicit On

Imports System
Imports System.Collections.Generic
Imports System.Data.SQLite
Imports oopfinals.Models

Namespace Data
    Public Class SaleRepository
        ''' <summary>
        ''' Generates a unique sequential receipt number: RCP-yyyyMMdd-XXXX
        ''' </summary>
        Public Function GenerateReceiptNumber() As String
            Dim todayPrefix As String = "RCP-" & DateTime.Today.ToString("yyyyMMdd") & "-"
            Using conn As SQLiteConnection = DatabaseHelper.GetConnection()
                Using cmd As New SQLiteCommand("
                    SELECT COUNT(*)
                    FROM Sales
                    WHERE ReceiptNo LIKE @pfx;", conn)
                    cmd.Parameters.AddWithValue("@pfx", todayPrefix & "%")
                    Dim countToday As Long = Convert.ToInt64(cmd.ExecuteScalar())
                    Return todayPrefix & (countToday + 1).ToString("D4")
                End Using
            End Using
        End Function

        ''' <summary>
        ''' Records a complete POS sale inside a database TRANSACTION.
        ''' Inserts Sale, inserts SaleItems, decrements product inventory, and records StockMovements.
        ''' </summary>
        Public Function RecordSale(sale As Sale, items As List(Of SaleItem)) As Integer
            If sale Is Nothing Then Throw New ArgumentNullException(NameOf(sale))
            If items Is Nothing OrElse items.Count = 0 Then Throw New ArgumentException("Sale must have at least one line item.")

            Using conn As SQLiteConnection = DatabaseHelper.GetConnection()
                Using trans As SQLiteTransaction = conn.BeginTransaction()
                    Try
                        ' 1. Insert into Sales
                        Dim saleId As Integer
                        Using cmdSale As New SQLiteCommand("
                            INSERT INTO Sales (ReceiptNo, SaleDate, CashierId, CustomerId, Subtotal, DiscountType, DiscountAmount, Total, CashReceived, ChangeAmount)
                            VALUES (@receipt, @date, @cashier, @customer, @subtotal, @discountType, @discountAmount, @total, @cash, @change);
                            SELECT last_insert_rowid();", conn, trans)
                            cmdSale.Parameters.AddWithValue("@receipt", sale.ReceiptNo)
                            cmdSale.Parameters.AddWithValue("@date", sale.SaleDate.ToString("yyyy-MM-dd HH:mm:ss"))
                            cmdSale.Parameters.AddWithValue("@cashier", sale.CashierId)
                            cmdSale.Parameters.AddWithValue("@customer", If(sale.CustomerId.HasValue, CObj(sale.CustomerId.Value), DBNull.Value))
                            cmdSale.Parameters.AddWithValue("@subtotal", sale.Subtotal)
                            cmdSale.Parameters.AddWithValue("@discountType", sale.DiscountType)
                            cmdSale.Parameters.AddWithValue("@discountAmount", sale.DiscountAmount)
                            cmdSale.Parameters.AddWithValue("@total", sale.Total)
                            cmdSale.Parameters.AddWithValue("@cash", sale.CashReceived)
                            cmdSale.Parameters.AddWithValue("@change", sale.ChangeAmount)

                            saleId = Convert.ToInt32(cmdSale.ExecuteScalar())
                            sale.SaleId = saleId
                        End Using

                        ' 2. Insert line items, deduct inventory, and record stock movements
                        For Each item In items
                            item.SaleId = saleId

                            ' Insert SaleItem
                            Using cmdItem As New SQLiteCommand("
                                INSERT INTO SaleItems (SaleId, ProductId, Quantity, UnitPrice, LineTotal)
                                VALUES (@saleId, @productId, @qty, @price, @total);
                                SELECT last_insert_rowid();", conn, trans)
                                cmdItem.Parameters.AddWithValue("@saleId", saleId)
                                cmdItem.Parameters.AddWithValue("@productId", item.ProductId)
                                cmdItem.Parameters.AddWithValue("@qty", item.Quantity)
                                cmdItem.Parameters.AddWithValue("@price", item.UnitPrice)
                                cmdItem.Parameters.AddWithValue("@total", item.LineTotal)
                                item.SaleItemId = Convert.ToInt32(cmdItem.ExecuteScalar())
                            End Using

                            ' Deduct inventory from Products (guard against insufficient stock)
                            Using cmdStock As New SQLiteCommand("
                                UPDATE Products
                                SET StockQty = StockQty - @qty
                                WHERE ProductId = @productId AND StockQty >= @qty;", conn, trans)
                                cmdStock.Parameters.AddWithValue("@qty", item.Quantity)
                                cmdStock.Parameters.AddWithValue("@productId", item.ProductId)
                                Dim affected As Integer = cmdStock.ExecuteNonQuery()
                                If affected = 0 Then
                                    Throw New InvalidOperationException($"Insufficient stock for a cart item during checkout. Transaction cancelled.")
                                End If
                            End Using

                            ' Log audit trail in StockMovements
                            Using cmdMove As New SQLiteCommand("
                                INSERT INTO StockMovements (ProductId, MovementType, QuantityChange, Reason, SupplierId, UserId, MovementDate)
                                VALUES (@productId, 'Sale', @qtyChange, @reason, NULL, @userId, @moveDate);", conn, trans)
                                cmdMove.Parameters.AddWithValue("@productId", item.ProductId)
                                cmdMove.Parameters.AddWithValue("@qtyChange", -item.Quantity)
                                cmdMove.Parameters.AddWithValue("@reason", "POS Sale: " & sale.ReceiptNo)
                                cmdMove.Parameters.AddWithValue("@userId", sale.CashierId)
                                cmdMove.Parameters.AddWithValue("@moveDate", sale.SaleDate.ToString("yyyy-MM-dd HH:mm:ss"))
                                cmdMove.ExecuteNonQuery()
                            End Using
                        Next

                        trans.Commit()
                        Return saleId
                    Catch
                        trans.Rollback()
                        Throw
                    End Try
                End Using
            End Using
        End Function

        Private Function MapSale(reader As SQLiteDataReader) As Sale
            Dim s As New Sale()
            s.SaleId = Convert.ToInt32(reader("SaleId"))
            s.ReceiptNo = Convert.ToString(reader("ReceiptNo"))
            s.SaleDate = DateTime.Parse(Convert.ToString(reader("SaleDate")))
            s.CashierId = Convert.ToInt32(reader("CashierId"))
            s.CashierName = If(reader("CashierName") IsNot DBNull.Value, Convert.ToString(reader("CashierName")), "Cashier")

            If reader("CustomerId") IsNot DBNull.Value Then
                s.CustomerId = Convert.ToInt32(reader("CustomerId"))
            End If
            s.CustomerName = If(reader("CustomerName") IsNot DBNull.Value, Convert.ToString(reader("CustomerName")), "Walk-In")

            s.Subtotal = Convert.ToDecimal(reader("Subtotal"))
            s.DiscountType = Convert.ToString(reader("DiscountType"))
            s.DiscountAmount = Convert.ToDecimal(reader("DiscountAmount"))
            s.Total = Convert.ToDecimal(reader("Total"))
            s.CashReceived = Convert.ToDecimal(reader("CashReceived"))
            s.ChangeAmount = Convert.ToDecimal(reader("ChangeAmount"))
            Return s
        End Function

        Public Function GetById(saleId As Integer) As Sale
            Using conn As SQLiteConnection = DatabaseHelper.GetConnection()
                Using cmd As New SQLiteCommand("
                    SELECT s.SaleId, s.ReceiptNo, s.SaleDate, s.CashierId, uCash.FullName AS CashierName,
                           s.CustomerId, uCust.FullName AS CustomerName, s.Subtotal, s.DiscountType,
                           s.DiscountAmount, s.Total, s.CashReceived, s.ChangeAmount
                    FROM Sales s
                    LEFT JOIN Users uCash ON s.CashierId = uCash.UserId
                    LEFT JOIN Users uCust ON s.CustomerId = uCust.UserId
                    WHERE s.SaleId = @id;", conn)
                    cmd.Parameters.AddWithValue("@id", saleId)
                    Using reader As SQLiteDataReader = cmd.ExecuteReader()
                        If reader.Read() Then
                            Dim sale As Sale = MapSale(reader)
                            sale.Items = GetSaleItems(sale.SaleId)
                            Return sale
                        End If
                    End Using
                End Using
            End Using
            Return Nothing
        End Function

        Public Function GetByReceiptNo(receiptNo As String) As Sale
            Using conn As SQLiteConnection = DatabaseHelper.GetConnection()
                Using cmd As New SQLiteCommand("
                    SELECT s.SaleId, s.ReceiptNo, s.SaleDate, s.CashierId, uCash.FullName AS CashierName,
                           s.CustomerId, uCust.FullName AS CustomerName, s.Subtotal, s.DiscountType,
                           s.DiscountAmount, s.Total, s.CashReceived, s.ChangeAmount
                    FROM Sales s
                    LEFT JOIN Users uCash ON s.CashierId = uCash.UserId
                    LEFT JOIN Users uCust ON s.CustomerId = uCust.UserId
                    WHERE s.ReceiptNo = @receipt;", conn)
                    cmd.Parameters.AddWithValue("@receipt", receiptNo.Trim())
                    Using reader As SQLiteDataReader = cmd.ExecuteReader()
                        If reader.Read() Then
                            Dim sale As Sale = MapSale(reader)
                            sale.Items = GetSaleItems(sale.SaleId)
                            Return sale
                        End If
                    End Using
                End Using
            End Using
            Return Nothing
        End Function

        Public Function GetSaleItems(saleId As Integer) As List(Of SaleItem)
            Dim list As New List(Of SaleItem)()
            Using conn As SQLiteConnection = DatabaseHelper.GetConnection()
                Using cmd As New SQLiteCommand("
                    SELECT si.SaleItemId, si.SaleId, si.ProductId, p.ProductCode, p.Name AS ProductName,
                           si.Quantity, si.UnitPrice, si.LineTotal
                    FROM SaleItems si
                    LEFT JOIN Products p ON si.ProductId = p.ProductId
                    WHERE si.SaleId = @saleId
                    ORDER BY si.SaleItemId;", conn)
                    cmd.Parameters.AddWithValue("@saleId", saleId)
                    Using reader As SQLiteDataReader = cmd.ExecuteReader()
                        While reader.Read()
                            Dim item As New SaleItem()
                            item.SaleItemId = Convert.ToInt32(reader("SaleItemId"))
                            item.SaleId = Convert.ToInt32(reader("SaleId"))
                            item.ProductId = Convert.ToInt32(reader("ProductId"))
                            item.ProductCode = If(reader("ProductCode") IsNot DBNull.Value, Convert.ToString(reader("ProductCode")), "N/A")
                            item.ProductName = If(reader("ProductName") IsNot DBNull.Value, Convert.ToString(reader("ProductName")), "Unknown Item")
                            item.Quantity = Convert.ToInt32(reader("Quantity"))
                            item.UnitPrice = Convert.ToDecimal(reader("UnitPrice"))
                            list.Add(item)
                        End While
                    End Using
                End Using
            End Using
            Return list
        End Function

        Public Function GetSalesByDate(targetDate As DateTime) As List(Of Sale)
            Dim startStr As String = targetDate.ToString("yyyy-MM-dd") & " 00:00:00"
            Dim endStr As String = targetDate.ToString("yyyy-MM-dd") & " 23:59:59"
            Return GetSalesBetween(startStr, endStr)
        End Function

        Public Function GetSalesByDateRange(startDate As DateTime, endDate As DateTime) As List(Of Sale)
            Dim startStr As String = startDate.ToString("yyyy-MM-dd") & " 00:00:00"
            Dim endStr As String = endDate.ToString("yyyy-MM-dd") & " 23:59:59"
            Return GetSalesBetween(startStr, endStr)
        End Function

        Private Function GetSalesBetween(startStr As String, endStr As String) As List(Of Sale)
            Dim list As New List(Of Sale)()
            Using conn As SQLiteConnection = DatabaseHelper.GetConnection()
                Using cmd As New SQLiteCommand("
                    SELECT s.SaleId, s.ReceiptNo, s.SaleDate, s.CashierId, uCash.FullName AS CashierName,
                           s.CustomerId, uCust.FullName AS CustomerName, s.Subtotal, s.DiscountType,
                           s.DiscountAmount, s.Total, s.CashReceived, s.ChangeAmount
                    FROM Sales s
                    LEFT JOIN Users uCash ON s.CashierId = uCash.UserId
                    LEFT JOIN Users uCust ON s.CustomerId = uCust.UserId
                    WHERE s.SaleDate BETWEEN @start AND @end
                    ORDER BY s.SaleDate DESC;", conn)
                    cmd.Parameters.AddWithValue("@start", startStr)
                    cmd.Parameters.AddWithValue("@end", endStr)
                    Using reader As SQLiteDataReader = cmd.ExecuteReader()
                        While reader.Read()
                            list.Add(MapSale(reader))
                        End While
                    End Using
                End Using
            End Using
            Return list
        End Function

        Public Function GetSalesByCustomer(customerId As Integer) As List(Of Sale)
            Dim list As New List(Of Sale)()
            Using conn As SQLiteConnection = DatabaseHelper.GetConnection()
                Using cmd As New SQLiteCommand("
                    SELECT s.SaleId, s.ReceiptNo, s.SaleDate, s.CashierId, uCash.FullName AS CashierName,
                           s.CustomerId, uCust.FullName AS CustomerName, s.Subtotal, s.DiscountType,
                           s.DiscountAmount, s.Total, s.CashReceived, s.ChangeAmount
                    FROM Sales s
                    LEFT JOIN Users uCash ON s.CashierId = uCash.UserId
                    LEFT JOIN Users uCust ON s.CustomerId = uCust.UserId
                    WHERE s.CustomerId = @custId
                    ORDER BY s.SaleDate DESC;", conn)
                    cmd.Parameters.AddWithValue("@custId", customerId)
                    Using reader As SQLiteDataReader = cmd.ExecuteReader()
                        While reader.Read()
                            list.Add(MapSale(reader))
                        End While
                    End Using
                End Using
            End Using
            Return list
        End Function

        Public Function GetSalesByCashierAndDate(cashierId As Integer, targetDate As DateTime) As List(Of Sale)
            Dim startStr As String = targetDate.ToString("yyyy-MM-dd") & " 00:00:00"
            Dim endStr As String = targetDate.ToString("yyyy-MM-dd") & " 23:59:59"
            Dim list As New List(Of Sale)()
            Using conn As SQLiteConnection = DatabaseHelper.GetConnection()
                Using cmd As New SQLiteCommand("
                    SELECT s.SaleId, s.ReceiptNo, s.SaleDate, s.CashierId, uCash.FullName AS CashierName,
                           s.CustomerId, uCust.FullName AS CustomerName, s.Subtotal, s.DiscountType,
                           s.DiscountAmount, s.Total, s.CashReceived, s.ChangeAmount
                    FROM Sales s
                    LEFT JOIN Users uCash ON s.CashierId = uCash.UserId
                    LEFT JOIN Users uCust ON s.CustomerId = uCust.UserId
                    WHERE s.CashierId = @cashierId AND s.SaleDate BETWEEN @start AND @end
                    ORDER BY s.SaleDate DESC;", conn)
                    cmd.Parameters.AddWithValue("@cashierId", cashierId)
                    cmd.Parameters.AddWithValue("@start", startStr)
                    cmd.Parameters.AddWithValue("@end", endStr)
                    Using reader As SQLiteDataReader = cmd.ExecuteReader()
                        While reader.Read()
                            list.Add(MapSale(reader))
                        End While
                    End Using
                End Using
            End Using
            Return list
        End Function
    End Class
End Namespace
