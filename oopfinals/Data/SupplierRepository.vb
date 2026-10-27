' ==============================================================================
' Project: Mini-Grocery Store POS & Daily Inventory System
' Course: Object-Oriented Programming (ITP 102) - University of Caloocan City
' Class: SupplierRepository
' OOP Concept: DATA ACCESS ENCAPSULATION
' Manages Supplier CRUD operations with parameterized SQLite queries.
' ==============================================================================

Option Strict On
Option Explicit On

Imports System
Imports System.Collections.Generic
Imports System.Data.SQLite
Imports oopfinals.Models

Namespace Data
    Public Class SupplierRepository
        Public Function GetAll() As List(Of Supplier)
            Dim list As New List(Of Supplier)()
            Using conn As SQLiteConnection = DatabaseHelper.GetConnection()
                Using cmd As New SQLiteCommand("SELECT SupplierId, Name, ContactNumber, Address FROM Suppliers ORDER BY Name;", conn)
                    Using reader As SQLiteDataReader = cmd.ExecuteReader()
                        While reader.Read()
                            list.Add(New Supplier(Convert.ToInt32(reader("SupplierId")),
                                                  Convert.ToString(reader("Name")),
                                                  Convert.ToString(reader("ContactNumber")),
                                                  Convert.ToString(reader("Address"))))
                        End While
                    End Using
                End Using
            End Using
            Return list
        End Function

        Public Function GetById(id As Integer) As Supplier
            Using conn As SQLiteConnection = DatabaseHelper.GetConnection()
                Using cmd As New SQLiteCommand("SELECT SupplierId, Name, ContactNumber, Address FROM Suppliers WHERE SupplierId = @id;", conn)
                    cmd.Parameters.AddWithValue("@id", id)
                    Using reader As SQLiteDataReader = cmd.ExecuteReader()
                        If reader.Read() Then
                            Return New Supplier(Convert.ToInt32(reader("SupplierId")),
                                                Convert.ToString(reader("Name")),
                                                Convert.ToString(reader("ContactNumber")),
                                                Convert.ToString(reader("Address")))
                        End If
                    End Using
                End Using
            End Using
            Return Nothing
        End Function

        Public Function Add(supplier As Supplier) As Integer
            Using conn As SQLiteConnection = DatabaseHelper.GetConnection()
                Using cmd As New SQLiteCommand("
                    INSERT INTO Suppliers (Name, ContactNumber, Address)
                    VALUES (@name, @contact, @addr);
                    SELECT last_insert_rowid();", conn)
                    cmd.Parameters.AddWithValue("@name", supplier.Name)
                    cmd.Parameters.AddWithValue("@contact", supplier.ContactNumber)
                    cmd.Parameters.AddWithValue("@addr", supplier.Address)
                    Return Convert.ToInt32(cmd.ExecuteScalar())
                End Using
            End Using
        End Function

        Public Function Update(supplier As Supplier) As Boolean
            Using conn As SQLiteConnection = DatabaseHelper.GetConnection()
                Using cmd As New SQLiteCommand("
                    UPDATE Suppliers
                    SET Name = @name, ContactNumber = @contact, Address = @addr
                    WHERE SupplierId = @id;", conn)
                    cmd.Parameters.AddWithValue("@name", supplier.Name)
                    cmd.Parameters.AddWithValue("@contact", supplier.ContactNumber)
                    cmd.Parameters.AddWithValue("@addr", supplier.Address)
                    cmd.Parameters.AddWithValue("@id", supplier.SupplierId)
                    Return cmd.ExecuteNonQuery() > 0
                End Using
            End Using
        End Function

        Public Function Delete(id As Integer) As Boolean
            Using conn As SQLiteConnection = DatabaseHelper.GetConnection()
                Using checkCmd As New SQLiteCommand("SELECT COUNT(*) FROM Products WHERE SupplierId = @id AND IsActive = 1;", conn)
                    checkCmd.Parameters.AddWithValue("@id", id)
                    If Convert.ToInt32(checkCmd.ExecuteScalar()) > 0 Then
                        Throw New InvalidOperationException("Cannot delete supplier because active products are assigned to this supplier.")
                    End If
                End Using

                Using cmd As New SQLiteCommand("DELETE FROM Suppliers WHERE SupplierId = @id;", conn)
                    cmd.Parameters.AddWithValue("@id", id)
                    Return cmd.ExecuteNonQuery() > 0
                End Using
            End Using
        End Function
    End Class
End Namespace
