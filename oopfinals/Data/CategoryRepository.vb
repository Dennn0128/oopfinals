' ==============================================================================
' Project: Mini-Grocery Store POS & Daily Inventory System
' Course: Object-Oriented Programming (ITP 102) - University of Caloocan City
' Class: CategoryRepository
' OOP Concept: DATA ACCESS ENCAPSULATION
' Manages Category CRUD operations with parameterized SQLite queries.
' ==============================================================================

Option Strict On
Option Explicit On

Imports System
Imports System.Collections.Generic
Imports System.Data.SQLite
Imports oopfinals.Models

Namespace Data
    Public Class CategoryRepository
        Public Function GetAll() As List(Of Category)
            Dim list As New List(Of Category)()
            Using conn As SQLiteConnection = DatabaseHelper.GetConnection()
                Using cmd As New SQLiteCommand("SELECT CategoryId, Name FROM Categories ORDER BY Name;", conn)
                    Using reader As SQLiteDataReader = cmd.ExecuteReader()
                        While reader.Read()
                            list.Add(New Category(Convert.ToInt32(reader("CategoryId")), Convert.ToString(reader("Name"))))
                        End While
                    End Using
                End Using
            End Using
            Return list
        End Function

        Public Function GetById(id As Integer) As Category
            Using conn As SQLiteConnection = DatabaseHelper.GetConnection()
                Using cmd As New SQLiteCommand("SELECT CategoryId, Name FROM Categories WHERE CategoryId = @id;", conn)
                    cmd.Parameters.AddWithValue("@id", id)
                    Using reader As SQLiteDataReader = cmd.ExecuteReader()
                        If reader.Read() Then
                            Return New Category(Convert.ToInt32(reader("CategoryId")), Convert.ToString(reader("Name")))
                        End If
                    End Using
                End Using
            End Using
            Return Nothing
        End Function

        Public Function Add(name As String) As Integer
            Using conn As SQLiteConnection = DatabaseHelper.GetConnection()
                Using cmd As New SQLiteCommand("INSERT INTO Categories (Name) VALUES (@name); SELECT last_insert_rowid();", conn)
                    cmd.Parameters.AddWithValue("@name", name.Trim())
                    Return Convert.ToInt32(cmd.ExecuteScalar())
                End Using
            End Using
        End Function

        Public Function Update(id As Integer, name As String) As Boolean
            Using conn As SQLiteConnection = DatabaseHelper.GetConnection()
                Using cmd As New SQLiteCommand("UPDATE Categories SET Name = @name WHERE CategoryId = @id;", conn)
                    cmd.Parameters.AddWithValue("@name", name.Trim())
                    cmd.Parameters.AddWithValue("@id", id)
                    Return cmd.ExecuteNonQuery() > 0
                End Using
            End Using
        End Function

        Public Function Delete(id As Integer) As Boolean
            ' Verify no products attached
            Using conn As SQLiteConnection = DatabaseHelper.GetConnection()
                Using checkCmd As New SQLiteCommand("SELECT COUNT(*) FROM Products WHERE CategoryId = @id AND IsActive = 1;", conn)
                    checkCmd.Parameters.AddWithValue("@id", id)
                    If Convert.ToInt32(checkCmd.ExecuteScalar()) > 0 Then
                        Throw New InvalidOperationException("Cannot delete category because it contains active products.")
                    End If
                End Using

                Using cmd As New SQLiteCommand("DELETE FROM Categories WHERE CategoryId = @id;", conn)
                    cmd.Parameters.AddWithValue("@id", id)
                    Return cmd.ExecuteNonQuery() > 0
                End Using
            End Using
        End Function
    End Class
End Namespace
