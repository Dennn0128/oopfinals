' ==============================================================================
' Project: Mini-Grocery Store POS & Daily Inventory System
' Course: Object-Oriented Programming (ITP 102) - University of Caloocan City
' Class: User (Abstract Base Class)
' OOP Concept: ABSTRACTION & ENCAPSULATION
' Defines common state and behavior for all user accounts in the grocery system.
' Abstract methods enforce polymorphic role behavior across derived classes.
' ==============================================================================

Option Strict On
Option Explicit On

Imports System
Imports System.Collections.Generic
Imports System.Windows.Forms

Namespace Models
    ''' <summary>
    ''' Abstract base class representing a generic user account in the system.
    ''' Demonstrates ABSTRACTION by defining MustInherit contract and MustOverride methods.
    ''' Demonstrates ENCAPSULATION through private fields, validation properties, and protected state.
    ''' </summary>
    Public MustInherit Class User
        ' --- ENCAPSULATION: Private fields protecting internal state ---
        Private _userId As Integer
        Private _username As String
        Private _passwordHash As String
        Private _salt As String
        Private _fullName As String
        Private _role As String
        Private _isActive As Boolean
        Private _createdAt As DateTime

        ' --- Public Properties with validation ---
        Public Property UserId As Integer
            Get
                Return _userId
            End Get
            Set(value As Integer)
                _userId = value
            End Set
        End Property

        Public Property Username As String
            Get
                Return _username
            End Get
            Set(value As String)
                If String.IsNullOrWhiteSpace(value) Then
                    Throw New ArgumentException("Username cannot be empty or whitespace.")
                End If
                _username = value.Trim().ToLowerInvariant()
            End Set
        End Property

        ''' <summary>
        ''' The cryptographic password hash. Encapsulated to prevent plain-text exposure.
        ''' </summary>
        Public Property PasswordHash As String
            Get
                Return _passwordHash
            End Get
            Set(value As String)
                If String.IsNullOrWhiteSpace(value) Then
                    Throw New ArgumentException("Password hash cannot be empty.")
                End If
                _passwordHash = value
            End Set
        End Property

        ''' <summary>
        ''' Cryptographic salt used during hashing.
        ''' </summary>
        Public Property Salt As String
            Get
                Return _salt
            End Get
            Set(value As String)
                _salt = value
            End Set
        End Property

        Public Property FullName As String
            Get
                Return _fullName
            End Get
            Set(value As String)
                If String.IsNullOrWhiteSpace(value) Then
                    Throw New ArgumentException("Full name cannot be empty.")
                End If
                _fullName = value.Trim()
            End Set
        End Property

        Public Property Role As String
            Get
                Return _role
            End Get
            Set(value As String)
                _role = value
            End Set
        End Property

        Public Property IsActive As Boolean
            Get
                Return _isActive
            End Get
            Set(value As Boolean)
                _isActive = value
            End Set
        End Property

        Public Property CreatedAt As DateTime
            Get
                Return _createdAt
            End Get
            Set(value As DateTime)
                _createdAt = value
            End Set
        End Property

        ''' <summary>
        ''' Default constructor
        ''' </summary>
        Protected Sub New()
            _isActive = True
            _createdAt = DateTime.Now
        End Sub

        ''' <summary>
        ''' Parameterized constructor enforcing valid initial data.
        ''' </summary>
        Protected Sub New(userId As Integer, username As String, fullName As String, role As String, isActive As Boolean)
            Me.New()
            Me.UserId = userId
            Me.Username = username
            Me.FullName = fullName
            Me.Role = role
            Me.IsActive = isActive
        End Sub

        ' --- ABSTRACTION: MustOverride members to be realized by subclasses ---

        ''' <summary>
        ''' Gets the formal human-readable display role name.
        ''' </summary>
        Public MustOverride Function GetRoleName() As String

        ''' <summary>
        ''' Factory method returning the tailored dashboard WinForm for this user type.
        ''' Demonstrates POLYMORPHISM: UI routing is determined dynamically by runtime type.
        ''' </summary>
        Public MustOverride Function GetDashboardForm() As Form

        ''' <summary>
        ''' Returns the set of permissions authorized for this role.
        ''' </summary>
        Public MustOverride Function GetPermissions() As List(Of String)

        ''' <summary>
        ''' Checks if the user is authorized for a specific system action.
        ''' </summary>
        Public Overridable Function HasPermission(permissionKey As String) As Boolean
            If String.IsNullOrWhiteSpace(permissionKey) Then Return False
            Dim perms As List(Of String) = GetPermissions()
            Return perms IsNot Nothing AndAlso perms.Contains(permissionKey)
        End Function

        Public Overrides Function ToString() As String
            Return $"{FullName} ({GetRoleName()}) - @{Username}"
        End Function
    End Class
End Namespace
