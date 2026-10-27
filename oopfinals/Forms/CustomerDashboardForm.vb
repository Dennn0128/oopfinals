' ==============================================================================
' Project: Mini-Grocery Store POS & Daily Inventory System
' Course: Object-Oriented Programming (ITP 102) - University of Caloocan City
' Form: CustomerDashboardForm (Code-Behind)
' OOP Concept: POLYMORPHISM & ENCAPSULATION
' Customer interface allowing product catalog browsing, real-time availability
' status viewing, personal purchase history auditing, and profile updates.
' ==============================================================================

Option Strict On
Option Explicit On

Imports System
Imports System.Collections.Generic
Imports System.Drawing
Imports System.Windows.Forms
Imports oopfinals.Data
Imports oopfinals.Models
Imports oopfinals.Services

Namespace Forms
    Public Class CustomerDashboardForm
        Private ReadOnly _customer As Customer
        Private ReadOnly _productRepo As New ProductRepository()
        Private ReadOnly _categoryRepo As New CategoryRepository()
        Private ReadOnly _saleRepo As New SaleRepository()
        Private ReadOnly _userRepo As New UserRepository()

        Public Sub New(customer As Customer)
            InitializeComponent()
            _customer = customer
            Me.DoubleBuffered = True
        End Sub

        Private Sub CustomerDashboardForm_Resize(sender As Object, e As EventArgs) Handles MyBase.Resize
            Me.Refresh()
        End Sub

        Private Sub CustomerDashboardForm_Load(sender As Object, e As EventArgs) Handles MyBase.Load
            ' SECURITY: verify role on form load (not just navigation/button-hiding)
            If _customer.Role <> "Customer" Then
                MessageBox.Show("Access denied: this dashboard is restricted to Customer accounts.", "Access Denied", MessageBoxButtons.OK, MessageBoxIcon.Warning)
                TryCast(Me, Form).BeginInvoke(New MethodInvoker(Sub() Me.Close()))
                Return
            End If

            lblUserInfo.Text = $"Welcome, {_customer.FullName} ({_customer.GetRoleName()})"
            lblProfUsernameVal.Text = _customer.Username
            txtProfFullName.Text = _customer.FullName

            LoadCategories()
            LoadProducts()
            LoadHistory()
        End Sub

        Private Sub LoadCategories()
            cboCategory.Items.Clear()
            cboCategory.Items.Add("All Categories")
            Dim cats As List(Of Category) = _categoryRepo.GetAll()
            For Each c In cats
                cboCategory.Items.Add(c)
            Next
            cboCategory.SelectedIndex = 0
        End Sub

        Private Sub LoadProducts()
            dgvProducts.Rows.Clear()
            Dim kw As String = txtSearch.Text.Trim()
            Dim catId As Nullable(Of Integer) = Nothing

            If cboCategory.SelectedIndex > 0 Then
                Dim selectedCat As Category = TryCast(cboCategory.SelectedItem, Category)
                If selectedCat IsNot Nothing Then
                    catId = selectedCat.CategoryId
                End If
            End If

            Dim list As List(Of Product) = _productRepo.Search(kw, catId, activeOnly:=True)
            For Each p In list
                Dim rowIdx As Integer = dgvProducts.Rows.Add(p.ProductCode, p.Name, p.CategoryName, p.Price.ToString("N2"), p.GetStockStatus())
                Dim status As String = p.GetStockStatus()
                If status = "Out of Stock" Then
                    dgvProducts.Rows(rowIdx).DefaultCellStyle.ForeColor = Color.Red
                ElseIf status = "Low Stock" Then
                    dgvProducts.Rows(rowIdx).DefaultCellStyle.ForeColor = Color.DarkOrange
                Else
                    dgvProducts.Rows(rowIdx).DefaultCellStyle.ForeColor = Color.DarkGreen
                End If
            Next
        End Sub

        Private Sub LoadHistory()
            dgvHistory.Rows.Clear()
            Dim sales As List(Of Sale) = _saleRepo.GetSalesByCustomer(_customer.UserId)
            For Each s In sales
                dgvHistory.Rows.Add(s.ReceiptNo, s.SaleDate.ToString("yyyy-MM-dd hh:mm tt"), s.CashierName, s.Subtotal.ToString("N2"), s.DiscountAmount.ToString("N2"), s.Total.ToString("N2"))
            Next
        End Sub

        Private Sub txtSearch_TextChanged(sender As Object, e As EventArgs) Handles txtSearch.TextChanged
            LoadProducts()
        End Sub

        Private Sub cboCategory_SelectedIndexChanged(sender As Object, e As EventArgs) Handles cboCategory.SelectedIndexChanged
            LoadProducts()
        End Sub

        Private Sub btnRefreshHistory_Click(sender As Object, e As EventArgs) Handles btnRefreshHistory.Click
            LoadHistory()
        End Sub

        Private Sub btnViewReceiptDetails_Click(sender As Object, e As EventArgs) Handles btnViewReceiptDetails.Click
            If dgvHistory.SelectedRows.Count = 0 Then
                MessageBox.Show("Please select a transaction from the history table to view details.", "Select Receipt", MessageBoxButtons.OK, MessageBoxIcon.Information)
                Return
            End If

            Dim receiptNo As String = Convert.ToString(dgvHistory.SelectedRows(0).Cells(0).Value)
            Dim sale As Sale = _saleRepo.GetByReceiptNo(receiptNo)
            If sale IsNot Nothing Then
                Using dlg As New PastReceiptForm(sale)
                    dlg.ShowDialog(Me)
                End Using
            End If
        End Sub

        Private Sub btnSaveProfile_Click(sender As Object, e As EventArgs) Handles btnSaveProfile.Click
            lblProfileMsg.Text = String.Empty
            Dim newName As String = txtProfFullName.Text.Trim()
            If String.IsNullOrWhiteSpace(newName) Then
                lblProfileMsg.Text = "Full name cannot be empty."
                lblProfileMsg.ForeColor = Color.Red
                Return
            End If

            If Not ValidationRules.IsValidFullName(newName) Then
                lblProfileMsg.Text = "Full name may only contain letters, spaces, hyphens, apostrophes, dots and commas."
                lblProfileMsg.ForeColor = Color.Red
                Return
            End If

            Try
                _customer.FullName = newName
                If _userRepo.UpdateUser(_customer) Then
                    lblUserInfo.Text = $"Welcome, {_customer.FullName} ({_customer.GetRoleName()})"
                    lblProfileMsg.Text = "Profile updated successfully!"
                    lblProfileMsg.ForeColor = Color.DarkGreen
                Else
                    lblProfileMsg.Text = "Failed to update profile."
                    lblProfileMsg.ForeColor = Color.Red
                End If
            Catch ex As Exception
                lblProfileMsg.Text = ex.Message
                lblProfileMsg.ForeColor = Color.Red
            End Try
        End Sub

        Private Sub btnChangePassword_Click(sender As Object, e As EventArgs) Handles btnChangePassword.Click
            lblPassMsg.Text = String.Empty
            Dim curPass As String = txtCurPass.Text
            Dim newPass As String = txtNewPass.Text
            Dim confPass As String = txtConfPass.Text

            If String.IsNullOrEmpty(curPass) OrElse String.IsNullOrEmpty(newPass) Then
                lblPassMsg.Text = "Please fill in all password fields."
                lblPassMsg.ForeColor = Color.Red
                Return
            End If

            If Not AuthService.VerifyPassword(curPass, _customer.Salt, _customer.PasswordHash) Then
                lblPassMsg.Text = "Current password is incorrect."
                lblPassMsg.ForeColor = Color.Red
                Return
            End If

            If newPass.Length < 6 Then
                lblPassMsg.Text = "New password must be at least 6 characters long."
                lblPassMsg.ForeColor = Color.Red
                Return
            End If

            If Not ValidationRules.IsValidPassword(newPass) Then
                lblPassMsg.Text = "New password cannot contain spaces."
                lblPassMsg.ForeColor = Color.Red
                Return
            End If

            If newPass <> confPass Then
                lblPassMsg.Text = "New passwords do not match."
                lblPassMsg.ForeColor = Color.Red
                Return
            End If

            Try
                Dim newSalt As String = AuthService.GenerateSalt()
                Dim newHash As String = AuthService.HashPassword(newPass, newSalt)
                If _userRepo.UpdatePassword(_customer.UserId, newHash, newSalt) Then
                    _customer.PasswordHash = newHash
                    _customer.Salt = newSalt
                    lblPassMsg.Text = "Password changed successfully!"
                    lblPassMsg.ForeColor = Color.DarkGreen
                    txtCurPass.Clear()
                    txtNewPass.Clear()
                    txtConfPass.Clear()
                Else
                    lblPassMsg.Text = "Failed to update password."
                    lblPassMsg.ForeColor = Color.Red
                End If
            Catch ex As Exception
                lblPassMsg.Text = ex.Message
                lblPassMsg.ForeColor = Color.Red
            End Try
        End Sub

        ' Typing-level validation: profile name accepts letters only, password fields block spaces
        Private Sub BlockDigitsInProfileName(sender As Object, e As KeyPressEventArgs) Handles txtProfFullName.KeyPress
            Dim allowedPunct As String = " '-.,"
            If Not (Char.IsControl(e.KeyChar) OrElse Char.IsLetter(e.KeyChar) OrElse
                    Char.IsWhiteSpace(e.KeyChar) OrElse allowedPunct.Contains(e.KeyChar.ToString())) Then
                e.Handled = True
            End If
        End Sub

        Private Sub BlockSpaceTyping(sender As Object, e As KeyPressEventArgs) _
            Handles txtCurPass.KeyPress, txtNewPass.KeyPress, txtConfPass.KeyPress
            If Char.IsWhiteSpace(e.KeyChar) Then
                e.Handled = True
            End If
        End Sub

        Private Sub btnLogout_Click(sender As Object, e As EventArgs) Handles btnLogout.Click
            If MessageBox.Show("Are you sure you want to log out?", "Confirm Logout", MessageBoxButtons.YesNo, MessageBoxIcon.Question) = DialogResult.Yes Then
                Me.Close()
            End If
        End Sub
    End Class
End Namespace
