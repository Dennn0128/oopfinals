' ==============================================================================
' Project: Mini-Grocery Store POS & Daily Inventory System
' Course: Object-Oriented Programming (ITP 102) - University of Caloocan City
' Class: StockRepository
' OOP Concept: TRANSACTION MANAGEMENT & DATA ACCESS
' Manages inventory deliveries (Stock-In), adjustments, and audit trail movements
' inside atomic SQLite transactions.
' ==============================================================================

Option Strict On
Option Explicit On

Imports System
Imports System.Collections.Generic
Imports System.Data.SQLite
Imports oopfinals.Models

Namespace Data
    Public Class StockRepository
        ''' <summary>
        ''' Records a stock delivery from a supplier inside an atomic database TRANSACTION.
        ''' </summary>
        Public Function RecordStockIn(productId As Integer, quantity As Integer, supplierId As Nullable(Of Integer),
                                      userId As Integer, newExpiryDate As Nullable(Of DateTime), reason As String) As Boolean
            If quantity <= 0 Then Throw New ArgumentException("Stock-in quantity must be greater than zero.")
            If quantity > 10000 Then Throw New ArgumentException("Stock-in quantity must not exceed 10,000 units per delivery.")
            If newExpiryDate.HasValue AndAlso newExpiryDate.Value.Date <= DateTime.Today Then
                Throw New InvalidOperationException("Expired stock must be rejected, not received. Record a disposal adjustment instead.")
            End If

            Using conn As SQLiteConnection = DatabaseHelper.GetConnection()
                Using trans As SQLiteTransaction = conn.BeginTransaction()
                    Try
                        ' 1. Update product quantity and optionally expiry date
                        Dim updateSql As String = "UPDATE Products SET StockQty = StockQty + @qty "
                        If newExpiryDate.HasValue Then
                            updateSql &= ", ExpiryDate = @exp "
                        End If
                        updateSql &= " WHERE ProductId = @pid;"

                        Using cmdProd As New SQLiteCommand(updateSql, conn, trans)
                            cmdProd.Parameters.AddWithValue("@qty", quantity)
                            cmdProd.Parameters.AddWithValue("@pid", productId)
                            If newExpiryDate.HasValue Then
                                cmdProd.Parameters.AddWithValue("@exp", newExpiryDate.Value.ToString("yyyy-MM-dd"))
                            End If
                            cmdProd.ExecuteNonQuery()
                        End Using

                        ' 2. Log StockMovement
                        Using cmdMove As New SQLiteCommand("
                            INSERT INTO StockMovements (ProductId, MovementType, QuantityChange, Reason, SupplierId, UserId, MovementDate)
                            VALUES (@pid, 'StockIn', @qty, @reason, @sup, @uid, @dt);", conn, trans)
                            cmdMove.Parameters.AddWithValue("@pid", productId)
                            cmdMove.Parameters.AddWithValue("@qty", quantity)
                            cmdMove.Parameters.AddWithValue("@reason", If(String.IsNullOrWhiteSpace(reason), "Supplier Delivery", reason.Trim()))
                            cmdMove.Parameters.AddWithValue("@sup", If(supplierId.HasValue, CObj(supplierId.Value), DBNull.Value))
                            cmdMove.Parameters.AddWithValue("@uid", userId)
                            cmdMove.Parameters.AddWithValue("@dt", DateTime.Now.ToString("yyyy-MM-dd HH:mm:ss"))
                            cmdMove.ExecuteNonQuery()
                        End Using

                        trans.Commit()
                        Return True
                    Catch
                        trans.Rollback()
                        Throw
                    End Try
                End Using
            End Using
        End Function

        ''' <summary>
        ''' Records a manual stock correction/adjustment inside an atomic database TRANSACTION.
        ''' Prevents negative final stock.
        ''' </summary>
        Public Function RecordAdjustment(productId As Integer, quantityChange As Integer, reason As String, userId As Integer) As Boolean
            If quantityChange = 0 Then Throw New ArgumentException("Quantity change cannot be zero.")
            If String.IsNullOrWhiteSpace(reason) Then Throw New ArgumentException("A specific adjustment reason is required.")

            Using conn As SQLiteConnection = DatabaseHelper.GetConnection()
                Using trans As SQLiteTransaction = conn.BeginTransaction()
                    Try
                        ' Check current stock
                        Using checkCmd As New SQLiteCommand("SELECT StockQty FROM Products WHERE ProductId = @pid;", conn, trans)
                            checkCmd.Parameters.AddWithValue("@pid", productId)
                            Dim currentQty As Integer = Convert.ToInt32(checkCmd.ExecuteScalar())
                            If currentQty + quantityChange < 0 Then
                                Throw New InvalidOperationException($"Adjustment would result in negative stock. Current: {currentQty}, Adjustment: {quantityChange}.")
                            End If
                        End Using

                        ' Update product stock
                        Using cmdProd As New SQLiteCommand("
                            UPDATE Products
                            SET StockQty = StockQty + @change
                            WHERE ProductId = @pid;", conn, trans)
                            cmdProd.Parameters.AddWithValue("@change", quantityChange)
                            cmdProd.Parameters.AddWithValue("@pid", productId)
                            cmdProd.ExecuteNonQuery()
                        End Using

                        ' Log Movement
                        Using cmdMove As New SQLiteCommand("
                            INSERT INTO StockMovements (ProductId, MovementType, QuantityChange, Reason, SupplierId, UserId, MovementDate)
                            VALUES (@pid, 'Adjustment', @change, @reason, NULL, @uid, @dt);", conn, trans)
                            cmdMove.Parameters.AddWithValue("@pid", productId)
                            cmdMove.Parameters.AddWithValue("@change", quantityChange)
                            cmdMove.Parameters.AddWithValue("@reason", reason.Trim())
                            cmdMove.Parameters.AddWithValue("@uid", userId)
                            cmdMove.Parameters.AddWithValue("@dt", DateTime.Now.ToString("yyyy-MM-dd HH:mm:ss"))
                            cmdMove.ExecuteNonQuery()
                        End Using

                        trans.Commit()
                        Return True
                    Catch
                        trans.Rollback()
                        Throw
                    End Try
                End Using
            End Using
        End Function

        ''' <summary>
        ''' Retrieves audit movements with optional filters.
        ''' </summary>
        Public Function GetMovements(Optional productId As Nullable(Of Integer) = Nothing,
                                     Optional movementType As String = Nothing,
                                     Optional startDate As Nullable(Of DateTime) = Nothing,
                                     Optional endDate As Nullable(Of DateTime) = Nothing) As List(Of StockMovement)
            Dim list As New List(Of StockMovement)()
            Using conn As SQLiteConnection = DatabaseHelper.GetConnection()
                Dim sql As String = "
                    SELECT sm.MovementId, sm.ProductId, p.ProductCode, p.Name AS ProductName,
                           sm.MovementType, sm.QuantityChange, sm.Reason, sm.SupplierId,
                           s.Name AS SupplierName, sm.UserId, u.FullName AS UserName, sm.MovementDate
                    FROM StockMovements sm
                    LEFT JOIN Products p ON sm.ProductId = p.ProductId
                    LEFT JOIN Suppliers s ON sm.SupplierId = s.SupplierId
                    LEFT JOIN Users u ON sm.UserId = u.UserId
                    WHERE 1=1 "

                If productId.HasValue AndAlso productId.Value > 0 Then
                    sql &= " AND sm.ProductId = @pid "
                End If

                If Not String.IsNullOrWhiteSpace(movementType) AndAlso movementType <> "All" Then
                    sql &= " AND sm.MovementType = @type "
                End If

                If startDate.HasValue Then
                    sql &= " AND sm.MovementDate >= @start "
                End If

                If endDate.HasValue Then
                    sql &= " AND sm.MovementDate <= @end "
                End If

                sql &= " ORDER BY sm.MovementDate DESC;"

                Using cmd As New SQLiteCommand(sql, conn)
                    If productId.HasValue AndAlso productId.Value > 0 Then
                        cmd.Parameters.AddWithValue("@pid", productId.Value)
                    End If
                    If Not String.IsNullOrWhiteSpace(movementType) AndAlso movementType <> "All" Then
                        cmd.Parameters.AddWithValue("@type", movementType)
                    End If
                    If startDate.HasValue Then
                        cmd.Parameters.AddWithValue("@start", startDate.Value.ToString("yyyy-MM-dd 00:00:00"))
                    End If
                    If endDate.HasValue Then
                        cmd.Parameters.AddWithValue("@end", endDate.Value.ToString("yyyy-MM-dd 23:59:59"))
                    End If

                    Using reader As SQLiteDataReader = cmd.ExecuteReader()
                        While reader.Read()
                            Dim m As New StockMovement()
                            m.MovementId = Convert.ToInt32(reader("MovementId"))
                            m.ProductId = Convert.ToInt32(reader("ProductId"))
                            m.ProductCode = If(reader("ProductCode") IsNot DBNull.Value, Convert.ToString(reader("ProductCode")), "N/A")
                            m.ProductName = If(reader("ProductName") IsNot DBNull.Value, Convert.ToString(reader("ProductName")), "N/A")
                            m.MovementType = Convert.ToString(reader("MovementType"))
                            m.QuantityChange = Convert.ToInt32(reader("QuantityChange"))
                            m.Reason = If(reader("Reason") IsNot DBNull.Value, Convert.ToString(reader("Reason")), "")
                            If reader("SupplierId") IsNot DBNull.Value Then
                                m.SupplierId = Convert.ToInt32(reader("SupplierId"))
                            End If
                            m.SupplierName = If(reader("SupplierName") IsNot DBNull.Value, Convert.ToString(reader("SupplierName")), "N/A")
                            m.UserId = Convert.ToInt32(reader("UserId"))
                            m.UserName = If(reader("UserName") IsNot DBNull.Value, Convert.ToString(reader("UserName")), "System")
                            m.MovementDate = DateTime.Parse(Convert.ToString(reader("MovementDate")))
                            list.Add(m)
                        End While
                    End Using
                End Using
            End Using
            Return list
        End Function
    End Class
End Namespace
