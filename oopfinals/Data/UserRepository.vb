' ==============================================================================
' Project: Mini-Grocery Store POS & Daily Inventory System
' Course: Object-Oriented Programming (ITP 102) - University of Caloocan City
' Class: UserRepository
' OOP Concept: DATA ACCESS & FACTORY POLYMORPHISM
' Handles database operations for User entities using parameterized SQLite queries.
' Dynamically instantiates the concrete polymorphic subclass (StoreOwner,
' Cashier, InventoryClerk, or Customer) based on the database role.
' ==============================================================================

Option Strict On
Option Explicit On

Imports System
Imports System.Collections.Generic
Imports System.Data.SQLite
Imports oopfinals.Models

Namespace Data
    Public Class UserRepository
        ''' <summary>
        ''' Maps a SQLiteDataReader record into the appropriate polymorphic User subclass.
        ''' Demonstrates FACTORY PATTERN and POLYMORPHISM.
        ''' </summary>
        Private Function MapUser(reader As SQLiteDataReader) As User
            Dim role As String = Convert.ToString(reader("Role"))
            Dim user As User

            Select Case role.ToLowerInvariant()
                Case "storeowner", "owner"
                    user = New StoreOwner()
                Case "cashier"
                    user = New Cashier()
                Case "inventoryclerk", "clerk"
                    user = New InventoryClerk()
                Case "customer"
                    user = New Customer()
                Case Else
                    user = New Customer()
            End Select

            user.UserId = Convert.ToInt32(reader("UserId"))
            user.Username = Convert.ToString(reader("Username"))
            user.PasswordHash = Convert.ToString(reader("PasswordHash"))
            user.Salt = Convert.ToString(reader("Salt"))
            user.FullName = Convert.ToString(reader("FullName"))
            user.Role = role
            user.IsActive = Convert.ToInt32(reader("IsActive")) = 1
            user.CreatedAt = DateTime.Parse(Convert.ToString(reader("CreatedAt")))

            Return user
        End Function

        ''' <summary>
        ''' Retrieves a user by their username (case-insensitive).
        ''' </summary>
        Public Function GetByUsername(username As String) As User
            Using conn As SQLiteConnection = DatabaseHelper.GetConnection()
                Using cmd As New SQLiteCommand("
                    SELECT UserId, Username, PasswordHash, Salt, FullName, Role, IsActive, CreatedAt
                    FROM Users
                    WHERE LOWER(Username) = LOWER(@username)
                    LIMIT 1;", conn)
                    cmd.Parameters.AddWithValue("@username", username.Trim())

                    Using reader As SQLiteDataReader = cmd.ExecuteReader()
                        If reader.Read() Then
                            Return MapUser(reader)
                        End If
                    End Using
                End Using
            End Using
            Return Nothing
        End Function

        ''' <summary>
        ''' Retrieves a user by their unique primary key ID.
        ''' </summary>
        Public Function GetById(userId As Integer) As User
            Using conn As SQLiteConnection = DatabaseHelper.GetConnection()
                Using cmd As New SQLiteCommand("
                    SELECT UserId, Username, PasswordHash, Salt, FullName, Role, IsActive, CreatedAt
                    FROM Users
                    WHERE UserId = @userId
                    LIMIT 1;", conn)
                    cmd.Parameters.AddWithValue("@userId", userId)

                    Using reader As SQLiteDataReader = cmd.ExecuteReader()
                        If reader.Read() Then
                            Return MapUser(reader)
                        End If
                    End Using
                End Using
            End Using
            Return Nothing
        End Function

        ''' <summary>
        ''' Retrieves all staff members (Store Owner, Cashier, Inventory Clerk) for administrative management.
        ''' </summary>
        Public Function GetAllStaff() As List(Of User)
            Dim list As New List(Of User)()
            Using conn As SQLiteConnection = DatabaseHelper.GetConnection()
                Using cmd As New SQLiteCommand("
                    SELECT UserId, Username, PasswordHash, Salt, FullName, Role, IsActive, CreatedAt
                    FROM Users
                    WHERE Role IN ('Cashier', 'InventoryClerk', 'StoreOwner')
                    ORDER BY Role, FullName;", conn)

                    Using reader As SQLiteDataReader = cmd.ExecuteReader()
                        While reader.Read()
                            list.Add(MapUser(reader))
                        End While
                    End Using
                End Using
            End Using
            Return list
        End Function

        ''' <summary>
        ''' Retrieves all registered customers.
        ''' </summary>
        Public Function GetAllCustomers() As List(Of Customer)
            Dim list As New List(Of Customer)()
            Using conn As SQLiteConnection = DatabaseHelper.GetConnection()
                Using cmd As New SQLiteCommand("
                    SELECT UserId, Username, PasswordHash, Salt, FullName, Role, IsActive, CreatedAt
                    FROM Users
                    WHERE Role = 'Customer'
                    ORDER BY FullName;", conn)

                    Using reader As SQLiteDataReader = cmd.ExecuteReader()
                        While reader.Read()
                            list.Add(DirectCast(MapUser(reader), Customer))
                        End While
                    End Using
                End Using
            End Using
            Return list
        End Function

        ''' <summary>
        ''' Checks if a username is already taken.
        ''' </summary>
        Public Function UsernameExists(username As String) As Boolean
            Using conn As SQLiteConnection = DatabaseHelper.GetConnection()
                Using cmd As New SQLiteCommand("SELECT COUNT(*) FROM Users WHERE LOWER(Username) = LOWER(@user);", conn)
                    cmd.Parameters.AddWithValue("@user", username.Trim())
                    Return Convert.ToInt32(cmd.ExecuteScalar()) > 0
                End Using
            End Using
        End Function

        ''' <summary>
        ''' Inserts a new user record.
        ''' </summary>
        Public Function CreateUser(user As User) As Integer
            Using conn As SQLiteConnection = DatabaseHelper.GetConnection()
                Using cmd As New SQLiteCommand("
                    INSERT INTO Users (Username, PasswordHash, Salt, FullName, Role, IsActive, CreatedAt)
                    VALUES (@user, @hash, @salt, @name, @role, @active, @created);
                    SELECT last_insert_rowid();", conn)
                    cmd.Parameters.AddWithValue("@user", user.Username)
                    cmd.Parameters.AddWithValue("@hash", user.PasswordHash)
                    cmd.Parameters.AddWithValue("@salt", user.Salt)
                    cmd.Parameters.AddWithValue("@name", user.FullName)
                    cmd.Parameters.AddWithValue("@role", user.Role)
                    cmd.Parameters.AddWithValue("@active", If(user.IsActive, 1, 0))
                    cmd.Parameters.AddWithValue("@created", DateTime.Now.ToString("yyyy-MM-dd HH:mm:ss"))

                    Dim newId As Integer = Convert.ToInt32(cmd.ExecuteScalar())
                    user.UserId = newId
                    Return newId
                End Using
            End Using
        End Function

        ''' <summary>
        ''' Updates an existing user's profile and active status.
        ''' </summary>
        Public Function UpdateUser(user As User) As Boolean
            Using conn As SQLiteConnection = DatabaseHelper.GetConnection()
                Using cmd As New SQLiteCommand("
                    UPDATE Users
                    SET FullName = @name, Role = @role, IsActive = @active
                    WHERE UserId = @userId;", conn)
                    cmd.Parameters.AddWithValue("@name", user.FullName)
                    cmd.Parameters.AddWithValue("@role", user.Role)
                    cmd.Parameters.AddWithValue("@active", If(user.IsActive, 1, 0))
                    cmd.Parameters.AddWithValue("@userId", user.UserId)

                    Return cmd.ExecuteNonQuery() > 0
                End Using
            End Using
        End Function

        ''' <summary>
        ''' Updates the user's password hash and salt.
        ''' </summary>
        Public Function UpdatePassword(userId As Integer, newHash As String, newSalt As String) As Boolean
            Using conn As SQLiteConnection = DatabaseHelper.GetConnection()
                Using cmd As New SQLiteCommand("
                    UPDATE Users
                    SET PasswordHash = @hash, Salt = @salt
                    WHERE UserId = @userId;", conn)
                    cmd.Parameters.AddWithValue("@hash", newHash)
                    cmd.Parameters.AddWithValue("@salt", newSalt)
                    cmd.Parameters.AddWithValue("@userId", userId)

                    Return cmd.ExecuteNonQuery() > 0
                End Using
            End Using
        End Function

        ''' <summary>
        ''' Toggles the active status of a user (activate or deactivate).
        ''' </summary>
        Public Function SetActiveStatus(userId As Integer, isActive As Boolean) As Boolean
            Using conn As SQLiteConnection = DatabaseHelper.GetConnection()
                Using cmd As New SQLiteCommand("UPDATE Users SET IsActive = @active WHERE UserId = @userId;", conn)
                    cmd.Parameters.AddWithValue("@active", If(isActive, 1, 0))
                    cmd.Parameters.AddWithValue("@userId", userId)
                    Return cmd.ExecuteNonQuery() > 0
                End Using
            End Using
        End Function
    End Class
End Namespace
