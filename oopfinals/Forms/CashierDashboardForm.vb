' ==============================================================================
' Project: Mini-Grocery Store POS & Daily Inventory System
' Course: Object-Oriented Programming (ITP 102) - University of Caloocan City
' Form: CashierDashboardForm (Code-Behind)
' OOP Concept: POLYMORPHISM (DISCOUNTS), TRANSACTION INTEGRITY, ENCAPSULATION
' Provides the full cashier checkout terminal: item catalog lookup, live cart
' management, polymorphic discount calculation, transaction checkout, receipt
' printing, and daily cashier sales auditing.
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
    Public Class CashierDashboardForm
        Private ReadOnly _cashier As Cashier
        Private ReadOnly _cartService As New CartService()
        Private ReadOnly _productRepo As New ProductRepository()
        Private ReadOnly _categoryRepo As New CategoryRepository()
        Private ReadOnly _saleRepo As New SaleRepository()
        Private ReadOnly _userRepo As New UserRepository()

        Private _lastCompletedSale As Sale = Nothing
        Private _isCheckingOut As Boolean = False

        Public Sub New(cashier As Cashier)
            InitializeComponent()
            _cashier = cashier
            Me.DoubleBuffered = True
        End Sub

        Private Sub CashierDashboardForm_Resize(sender As Object, e As EventArgs) Handles MyBase.Resize
            Me.Refresh()
        End Sub

        Private Sub CashierDashboardForm_Load(sender As Object, e As EventArgs) Handles MyBase.Load
            ' SECURITY: verify role on form load
            If _cashier.Role <> "Cashier" Then
                MessageBox.Show("Access denied: this dashboard is restricted to Cashier accounts.", "Access Denied", MessageBoxButtons.OK, MessageBoxIcon.Warning)
                Me.BeginInvoke(New MethodInvoker(Sub() Me.Close()))
                Return
            End If

            lblCashierInfo.Text = $"Cashier: {_cashier.FullName} (@{_cashier.Username})"

            PopulateCategories()
            PopulateDiscounts()
            PopulateCustomers()
            LoadPosProducts()
            RefreshCartGrid()
            LoadTodayTransactions()
        End Sub

        Private Sub PopulateCategories()
            cboPosCategory.Items.Clear()
            cboPosCategory.Items.Add("All Categories")
            Dim cats As List(Of Category) = _categoryRepo.GetAll()
            For Each c In cats
                cboPosCategory.Items.Add(c)
            Next
            cboPosCategory.SelectedIndex = 0
        End Sub

        ''' <summary>
        ''' Demonstrates POLYMORPHISM: Populate discount strategies inheriting from base Discount.
        ''' </summary>
        Private Sub PopulateDiscounts()
            cboDiscount.Items.Clear()
            cboDiscount.Items.Add(New NoDiscount())
            cboDiscount.Items.Add(New SeniorPwdDiscount())
            cboDiscount.Items.Add(New RegularDiscount(5D))
            cboDiscount.Items.Add(New RegularDiscount(10D))
            cboDiscount.SelectedIndex = 0
        End Sub

        Private Sub PopulateCustomers()
            cboCustomer.Items.Clear()
            cboCustomer.Items.Add("Walk-In Customer")
            Dim customers As List(Of Customer) = _userRepo.GetAllCustomers()
            For Each cust In customers
                cboCustomer.Items.Add(cust)
            Next
            cboCustomer.SelectedIndex = 0
        End Sub

        Private Sub LoadPosProducts()
            dgvPosProducts.Rows.Clear()
            Dim kw As String = txtPosSearch.Text.Trim()
            Dim catId As Nullable(Of Integer) = Nothing

            If cboPosCategory.SelectedIndex > 0 Then
                Dim selectedCat As Category = TryCast(cboPosCategory.SelectedItem, Category)
                If selectedCat IsNot Nothing Then
                    catId = selectedCat.CategoryId
                End If
            End If

            Dim list As List(Of Product) = _productRepo.Search(kw, catId, activeOnly:=True)
            For Each p In list
                Dim rowIdx As Integer = dgvPosProducts.Rows.Add(p.ProductCode, p.Name, p.Price.ToString("N2"), p.StockQty)
                ' Visual highlight for zero or low stock
                If p.StockQty <= 0 Then
                    dgvPosProducts.Rows(rowIdx).DefaultCellStyle.ForeColor = Color.Red
                ElseIf p.StockQty <= p.ReorderLevel Then
                    dgvPosProducts.Rows(rowIdx).DefaultCellStyle.ForeColor = Color.DarkOrange
                End If
            Next
        End Sub

        Private Sub txtPosSearch_TextChanged(sender As Object, e As EventArgs) Handles txtPosSearch.TextChanged
            LoadPosProducts()
        End Sub

        Private Sub cboPosCategory_SelectedIndexChanged(sender As Object, e As EventArgs) Handles cboPosCategory.SelectedIndexChanged
            LoadPosProducts()
        End Sub

        Private Sub btnAddToCart_Click(sender As Object, e As EventArgs) Handles btnAddToCart.Click
            If dgvPosProducts.SelectedRows.Count = 0 Then
                MessageBox.Show("Please select a product from the list.", "Notice", MessageBoxButtons.OK, MessageBoxIcon.Information)
                Return
            End If

            Dim code As String = Convert.ToString(dgvPosProducts.SelectedRows(0).Cells(0).Value)
            Dim prod As Product = _productRepo.GetByCode(code)
            If prod Is Nothing Then Return

            ' BUSINESS RULE: expired products cannot be sold
            If prod.IsExpired() Then
                MessageBox.Show($"'{prod.Name}' expired on {prod.ExpiryDate.Value:yyyy-MM-dd} and CANNOT be sold." & vbCrLf &
                                "Please remove it from the sellable list (notify the Inventory Clerk).",
                                "Expired Product", MessageBoxButtons.OK, MessageBoxIcon.Warning)
                Return
            End If

            ' BUSINESS RULE: warn for near-expiry products (within 30 days)
            If prod.IsExpiringSoon(30) Then
                If MessageBox.Show($"'{prod.Name}' is expiring soon ({prod.ExpiryDate.Value:yyyy-MM-dd})." & vbCrLf &
                                   "Add it to the cart anyway?", "Near-Expiry Warning",
                                   MessageBoxButtons.YesNo, MessageBoxIcon.Question) <> DialogResult.Yes Then
                    Return
                End If
            End If

            ' BUSINESS RULE: out-of-stock items blocked up-front
            If prod.StockQty <= 0 Then
                MessageBox.Show($"'{prod.Name}' is out of stock.", "Out of Stock", MessageBoxButtons.OK, MessageBoxIcon.Warning)
                Return
            End If

            Dim qty As Integer = Convert.ToInt32(numQuantity.Value)

            ' BUSINESS RULE: quantity must be a positive whole number and within stock
            If qty <= 0 Then
                MessageBox.Show("Quantity must be a positive whole number.", "Validation", MessageBoxButtons.OK, MessageBoxIcon.Warning)
                Return
            End If

            Try
                _cartService.AddItem(prod, qty)
                RefreshCartGrid()
                numQuantity.Value = 1
            Catch ex As Exception
                MessageBox.Show(ex.Message, "Stock Warning", MessageBoxButtons.OK, MessageBoxIcon.Warning)
            End Try
        End Sub

        Private Sub RefreshCartGrid()
            dgvCart.Rows.Clear()
            For Each itm In _cartService.Items
                dgvCart.Rows.Add(itm.ProductCode, itm.ProductName, itm.UnitPrice.ToString("N2"), itm.Quantity, itm.LineTotal.ToString("N2"))
            Next

            UpdateCheckoutTotals()
        End Sub

        ''' <summary>
        ''' Updates subtotal, polymorphic discount deduction, net total, and change.
        ''' </summary>
        Private Sub UpdateCheckoutTotals()
            Dim subtotal As Decimal = _cartService.GetSubtotal()
            Dim discountAmt As Decimal = _cartService.GetDiscountAmount()
            Dim netTotal As Decimal = _cartService.GetNetTotal()

            lblSubtotal.Text = $"PHP {subtotal:N2}"
            lblDiscountAmount.Text = $"-PHP {discountAmt:N2}"
            lblNetTotal.Text = $"PHP {netTotal:N2}"

            Dim cash As Decimal = 0D
            If Decimal.TryParse(txtCashReceived.Text.Trim(), cash) Then
                If cash >= netTotal Then
                    Dim change As Decimal = cash - netTotal
                    lblChange.Text = $"PHP {change:N2}"
                    lblChange.ForeColor = Color.FromArgb(30, 94, 58)
                Else
                    lblChange.Text = "Insufficient Cash"
                    lblChange.ForeColor = Color.Red
                End If
            Else
                lblChange.Text = "PHP 0.00"
                lblChange.ForeColor = Color.Black
            End If
        End Sub

        ''' <summary>
        ''' Demonstrates POLYMORPHISM: Dynamically updates the discount calculation strategy.
        ''' </summary>
        Private Sub cboDiscount_SelectedIndexChanged(sender As Object, e As EventArgs) Handles cboDiscount.SelectedIndexChanged
            Dim selectedDiscount As Discount = TryCast(cboDiscount.SelectedItem, Discount)
            If selectedDiscount IsNot Nothing Then
                _cartService.CurrentDiscount = selectedDiscount
            Else
                _cartService.CurrentDiscount = New NoDiscount()
            End If
            UpdateCheckoutTotals()
        End Sub

        Private Sub txtCashReceived_TextChanged(sender As Object, e As EventArgs) Handles txtCashReceived.TextChanged
            UpdateCheckoutTotals()
        End Sub

        Private Sub btnIncQty_Click(sender As Object, e As EventArgs) Handles btnIncQty.Click
            If dgvCart.SelectedRows.Count = 0 Then Return
            Dim code As String = Convert.ToString(dgvCart.SelectedRows(0).Cells(0).Value)
            Dim prod As Product = _productRepo.GetByCode(code)
            If prod Is Nothing Then Return

            Dim currentItem = _cartService.Items.FirstOrDefault(Function(i) i.ProductId = prod.ProductId)
            If currentItem IsNot Nothing Then
                Try
                    _cartService.UpdateItemQuantity(prod.ProductId, currentItem.Quantity + 1, prod.StockQty)
                    RefreshCartGrid()
                Catch ex As Exception
                    MessageBox.Show(ex.Message, "Stock Limit", MessageBoxButtons.OK, MessageBoxIcon.Warning)
                End Try
            End If
        End Sub

        Private Sub btnDecQty_Click(sender As Object, e As EventArgs) Handles btnDecQty.Click
            If dgvCart.SelectedRows.Count = 0 Then Return
            Dim code As String = Convert.ToString(dgvCart.SelectedRows(0).Cells(0).Value)
            Dim prod As Product = _productRepo.GetByCode(code)
            If prod Is Nothing Then Return

            Dim currentItem = _cartService.Items.FirstOrDefault(Function(i) i.ProductId = prod.ProductId)
            If currentItem IsNot Nothing Then
                _cartService.UpdateItemQuantity(prod.ProductId, currentItem.Quantity - 1, prod.StockQty)
                RefreshCartGrid()
            End If
        End Sub

        Private Sub btnRemoveItem_Click(sender As Object, e As EventArgs) Handles btnRemoveItem.Click
            If dgvCart.SelectedRows.Count = 0 Then Return
            Dim code As String = Convert.ToString(dgvCart.SelectedRows(0).Cells(0).Value)
            Dim prod As Product = _productRepo.GetByCode(code)
            If prod IsNot Nothing Then
                _cartService.RemoveItem(prod.ProductId)
                RefreshCartGrid()
            End If
        End Sub

        Private Sub btnClearCart_Click(sender As Object, e As EventArgs) Handles btnClearCart.Click
            If _cartService.Items.Count > 0 Then
                If MessageBox.Show("Are you sure you want to clear all items in the cart?", "Clear Cart", MessageBoxButtons.YesNo, MessageBoxIcon.Question) = DialogResult.Yes Then
                    _cartService.ClearCart()
                    cboDiscount.SelectedIndex = 0
                    txtCashReceived.Clear()
                    RefreshCartGrid()
                End If
            End If
        End Sub

        ''' <summary>
        ''' Executes POS checkout inside an atomic database TRANSACTION.
        ''' </summary>
        Private Sub btnCheckout_Click(sender As Object, e As EventArgs) Handles btnCheckout.Click
            ' Double-click / re-entry guard: prevent duplicate transactions
            If _isCheckingOut Then Return

            If _cartService.Items.Count = 0 Then
                MessageBox.Show("Cart is empty. Please add products before checking out.", "Empty Cart", MessageBoxButtons.OK, MessageBoxIcon.Warning)
                Return
            End If

            Dim netTotal As Decimal = _cartService.GetNetTotal()
            Dim cashReceived As Decimal = 0D

            If Not Decimal.TryParse(txtCashReceived.Text.Trim(), cashReceived) Then
                MessageBox.Show("Cash tendered must be a valid number.", "Payment Validation", MessageBoxButtons.OK, MessageBoxIcon.Warning)
                txtCashReceived.Focus()
                Return
            End If

            If cashReceived < 0D Then
                MessageBox.Show("Cash tendered cannot be negative.", "Payment Validation", MessageBoxButtons.OK, MessageBoxIcon.Warning)
                txtCashReceived.Focus()
                Return
            End If

            If cashReceived > ValidationRules.MaxMoney Then
                MessageBox.Show($"Cash tendered is unreasonably large (max PHP {ValidationRules.MaxMoney:N2}).", "Payment Validation", MessageBoxButtons.OK, MessageBoxIcon.Warning)
                txtCashReceived.Focus()
                Return
            End If

            If cashReceived < netTotal Then
                MessageBox.Show($"Cash tendered is insufficient. Total due is PHP {netTotal:N2}.", "Payment Validation", MessageBoxButtons.OK, MessageBoxIcon.Warning)
                txtCashReceived.Focus()
                Return
            End If

            Dim changeDue As Decimal = cashReceived - netTotal

            ' Confirm with Cashier
            Dim promptMsg As String = $"Confirm Checkout:" & vbCrLf &
                                      $"Total Due : PHP {netTotal:N2}" & vbCrLf &
                                      $"Cash Rec. : PHP {cashReceived:N2}" & vbCrLf &
                                      $"Change Due: PHP {changeDue:N2}"

            If MessageBox.Show(promptMsg, "Confirm Sale Checkout", MessageBoxButtons.YesNo, MessageBoxIcon.Question) <> DialogResult.Yes Then
                Return
            End If

            Try
                _isCheckingOut = True
                btnCheckout.Enabled = False

                Try
                    ' Prepare Sale Model
                    Dim newSale As New Sale() With {
                        .ReceiptNo = _saleRepo.GenerateReceiptNumber(),
                        .SaleDate = DateTime.Now,
                        .CashierId = _cashier.UserId,
                        .CashierName = _cashier.FullName,
                        .Subtotal = _cartService.GetSubtotal(),
                        .DiscountType = _cartService.CurrentDiscount.DiscountName,
                        .DiscountAmount = _cartService.GetDiscountAmount(),
                        .Total = netTotal,
                        .CashReceived = cashReceived,
                        .ChangeAmount = changeDue,
                        .Items = _cartService.Items.ToList()
                    }

                    ' Optional attached customer
                    If cboCustomer.SelectedIndex > 0 Then
                        Dim attachedCust As Customer = TryCast(cboCustomer.SelectedItem, Customer)
                        If attachedCust IsNot Nothing Then
                            newSale.CustomerId = attachedCust.UserId
                            newSale.CustomerName = attachedCust.FullName
                        End If
                    End If

                    ' Save inside atomic database transaction
                    Dim newSaleId As Integer = _saleRepo.RecordSale(newSale, newSale.Items)
                    newSale.SaleId = newSaleId
                    _lastCompletedSale = newSale

                    MessageBox.Show($"Checkout successful! Receipt #{newSale.ReceiptNo} generated." & vbCrLf & $"Change: PHP {changeDue:N2}", "Transaction Complete", MessageBoxButtons.OK, MessageBoxIcon.Information)

                    ' Prompt receipt preview / printing
                    Using receiptForm As New ReceiptPreviewForm(newSale)
                        receiptForm.ShowDialog(Me)
                    End Using

                    ' Reset POS state for next customer
                    _cartService.ClearCart()
                    cboDiscount.SelectedIndex = 0
                    txtCashReceived.Clear()
                    cboCustomer.SelectedIndex = 0
                    RefreshCartGrid()
                    LoadPosProducts()
                    LoadTodayTransactions()

                Catch ex As Exception
                    MessageBox.Show($"Failed to process checkout: {ex.Message}", "Transaction Error", MessageBoxButtons.OK, MessageBoxIcon.Error)
                End Try
            Finally
                _isCheckingOut = False
                btnCheckout.Enabled = True
            End Try
        End Sub

        Private Sub btnPrintLast_Click(sender As Object, e As EventArgs) Handles btnPrintLast.Click
            If _lastCompletedSale IsNot Nothing Then
                Using receiptForm As New ReceiptPreviewForm(_lastCompletedSale)
                    receiptForm.ShowDialog(Me)
                End Using
            Else
                MessageBox.Show("No transaction has been completed in this session yet.", "No Receipt", MessageBoxButtons.OK, MessageBoxIcon.Information)
            End If
        End Sub

        Private Sub LoadTodayTransactions()
            dgvTodaySales.Rows.Clear()
            Dim sales As List(Of Sale) = _saleRepo.GetSalesByCashierAndDate(_cashier.UserId, DateTime.Today)
            Dim totalRevenue As Decimal = 0D

            For Each s In sales
                dgvTodaySales.Rows.Add(s.ReceiptNo, s.SaleDate.ToString("hh:mm tt"), If(String.IsNullOrEmpty(s.CustomerName), "Walk-In", s.CustomerName), s.Subtotal.ToString("N2"), s.DiscountAmount.ToString("N2"), s.Total.ToString("N2"))
                totalRevenue += s.Total
            Next

            lblTodaySummary.Text = $"Today: {sales.Count} transaction(s) | Total Sales: PHP {totalRevenue:N2}"
        End Sub

        Private Sub btnRefreshTodaySales_Click(sender As Object, e As EventArgs) Handles btnRefreshTodaySales.Click
            LoadTodayTransactions()
        End Sub

        Private Sub btnViewSaleDetails_Click(sender As Object, e As EventArgs) Handles btnViewSaleDetails.Click
            If dgvTodaySales.SelectedRows.Count = 0 Then
                MessageBox.Show("Please select a transaction from the list.", "Notice", MessageBoxButtons.OK, MessageBoxIcon.Information)
                Return
            End If

            Dim receiptNo As String = Convert.ToString(dgvTodaySales.SelectedRows(0).Cells(0).Value)
            Dim sale As Sale = _saleRepo.GetByReceiptNo(receiptNo)
            If sale IsNot Nothing Then
                Using dlg As New PastReceiptForm(sale)
                    dlg.ShowDialog(Me)
                End Using
            End If
        End Sub

        Private Sub btnLogout_Click(sender As Object, e As EventArgs) Handles btnLogout.Click
            If _cartService.Items.Count > 0 Then
                If MessageBox.Show("There are items currently in the active cart. Are you sure you want to log out and discard the cart?", "Confirm Logout", MessageBoxButtons.YesNo, MessageBoxIcon.Warning) <> DialogResult.Yes Then
                    Return
                End If
            Else
                If MessageBox.Show("Are you sure you want to log out?", "Confirm Logout", MessageBoxButtons.YesNo, MessageBoxIcon.Question) <> DialogResult.Yes Then
                    Return
                End If
            End If
            Me.Close()
        End Sub
    End Class
End Namespace
