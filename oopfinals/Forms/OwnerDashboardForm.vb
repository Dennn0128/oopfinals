' ==============================================================================
' Project: Mini-Grocery Store POS & Daily Inventory System
' Course: Object-Oriented Programming (ITP 102) - University of Caloocan City
' Form: OwnerDashboardForm (Code-Behind)
' OOP Concept: FULL SYSTEM CONTROL, ABSTRACTION, REPORTING & SECURITY
' Provides total administrative access: KPI dashboard cards, staff management,
' complete product/pricing/supplier CRUD, historical sales auditing, and
' business reporting with CSV export and printing.
' ==============================================================================

Option Strict On
Option Explicit On

Imports System
Imports System.Collections.Generic
Imports System.Data
Imports System.Drawing
Imports System.IO
Imports System.Windows.Forms
Imports oopfinals.Data
Imports oopfinals.Models
Imports oopfinals.Services

Namespace Forms
    Public Class OwnerDashboardForm
        Private ReadOnly _owner As StoreOwner
        Private ReadOnly _userRepo As New UserRepository()
        Private ReadOnly _productRepo As New ProductRepository()
        Private ReadOnly _categoryRepo As New CategoryRepository()
        Private ReadOnly _supplierRepo As New SupplierRepository()
        Private ReadOnly _saleRepo As New SaleRepository()

        Private _selectedStaffId As Integer = 0
        Private _selectedProductId As Integer = 0
        Private _selectedCategoryId As Integer = 0
        Private _selectedSupplierId As Integer = 0
        Private _currentReportData As DataTable = Nothing

        Public Sub New(owner As StoreOwner)
            InitializeComponent()
            _owner = owner
            Me.DoubleBuffered = True
            LayoutKpiCards()
        End Sub

        Private Sub OwnerDashboardForm_Resize(sender As Object, e As EventArgs) Handles MyBase.Resize
            Me.Refresh()
        End Sub

        Private Sub pnlCards_Resize(sender As Object, e As EventArgs) Handles pnlCards.Resize
            LayoutKpiCards()
        End Sub

        Private Sub LayoutKpiCards()
            Dim cards() As Panel = {pnlCard1, pnlCard2, pnlCard3, pnlCard4}
            Dim lblTitles() As Label = {lblCard1Title, lblCard2Title, lblCard3Title, lblCard4Title}
            Dim lblVals() As Label = {lblCard1Val, lblCard2Val, lblCard3Val, lblCard4Val}

            Dim panelWidth As Integer = pnlCards.ClientSize.Width
            Const margin As Integer = 15
            Const cardGap As Integer = 15
            Dim n As Integer = cards.Length
            Dim avail As Integer = panelWidth - (margin * 2)

            Dim cardW As Integer = (avail - (cardGap * (n - 1))) \ n
            If cardW < 200 Then
                cardW = 200
            End If

            Dim gap As Integer = (avail - (cardW * n)) \ (n - 1)
            If gap < 6 Then gap = 6
            Dim spacing As Integer = cardW + gap

            For i As Integer = 0 To n - 1
                cards(i).Size = New Size(cardW, cards(i).Height)
                cards(i).Left = margin + (spacing * i)
                lblTitles(i).Size = New Size(cardW - 20, lblTitles(i).Height)
                lblVals(i).Size = New Size(cardW - 20, lblVals(i).Height)
            Next
        End Sub

        Private Sub OwnerDashboardForm_Load(sender As Object, e As EventArgs) Handles MyBase.Load
            ' SECURITY: verify role on form load
            If _owner.Role <> "StoreOwner" Then
                MessageBox.Show("Access denied: this dashboard is restricted to Store Owner accounts.", "Access Denied", MessageBoxButtons.OK, MessageBoxIcon.Warning)
                Me.BeginInvoke(New MethodInvoker(Sub() Me.Close()))
                Return
            End If

            lblOwnerInfo.Text = $"Store Owner: {_owner.FullName} (@{_owner.Username})"

            ' Set initial date ranges
            dtpTxStart.Value = DateTime.Today.AddDays(-7)
            dtpTxEnd.Value = DateTime.Today
            dtpReportDate.Value = DateTime.Today

            ' Populate drop-down lists
            PopulateStaffRoles()
            PopulateCategoriesAndSuppliers()
            PopulateReportTypes()

            ' Load initial data views
            LoadOverview()
            LoadStaffList()
            LoadProductsList()
            LoadCategoriesList()
            LoadSuppliersList()
            LoadTransactions()
        End Sub

        ' =====================================================================
        ' TAB 1: OVERVIEW & KPI CARDS
        ' =====================================================================
        Private Sub LoadOverview()
            Dim todaySales As List(Of Sale) = _saleRepo.GetSalesByDate(DateTime.Today)
            Dim totalRev As Decimal = 0D
            For Each s In todaySales
                totalRev += s.Total
            Next

            Dim lowStockCount As Integer = _productRepo.GetLowStockProducts().Count
            Dim expSoonCount As Integer = _productRepo.GetExpiringProducts(30).Count

            lblCard1Val.Text = $"PHP {totalRev:N2}"
            lblCard2Val.Text = todaySales.Count.ToString()
            lblCard3Val.Text = $"{lowStockCount} items"
            lblCard4Val.Text = $"{expSoonCount} items"

            dgvOverviewRecent.Rows.Clear()
            For Each s In todaySales
                dgvOverviewRecent.Rows.Add(s.ReceiptNo, s.SaleDate.ToString("hh:mm tt"), s.CashierName, s.Total.ToString("N2"))
            Next
        End Sub

        ' =====================================================================
        ' TAB 2: STAFF MANAGEMENT
        ' =====================================================================
        Private Sub PopulateStaffRoles()
            cboStaffRole.Items.Clear()
            cboStaffRole.Items.Add("Cashier")
            cboStaffRole.Items.Add("InventoryClerk")
            cboStaffRole.SelectedIndex = 0
        End Sub

        Private Sub LoadStaffList()
            dgvStaff.Rows.Clear()
            Dim staffList As List(Of User) = _userRepo.GetAllStaff()
            For Each u In staffList
                Dim statusStr As String = If(u.IsActive, "Active", "Inactive")
                Dim rowIdx As Integer = dgvStaff.Rows.Add(u.UserId, u.Username, u.FullName, u.Role, statusStr, u.CreatedAt.ToString("yyyy-MM-dd"))
                If Not u.IsActive Then
                    dgvStaff.Rows(rowIdx).DefaultCellStyle.ForeColor = Color.Red
                End If
            Next
        End Sub

        Private Sub dgvStaff_SelectionChanged(sender As Object, e As EventArgs) Handles dgvStaff.SelectionChanged
            If dgvStaff.SelectedRows.Count = 0 Then Return
            Dim row As DataGridViewRow = dgvStaff.SelectedRows(0)
            _selectedStaffId = Convert.ToInt32(row.Cells(0).Value)
            txtStaffUsername.Text = Convert.ToString(row.Cells(1).Value)
            txtStaffName.Text = Convert.ToString(row.Cells(2).Value)

            Dim role As String = Convert.ToString(row.Cells(3).Value)
            If cboStaffRole.Items.Contains(role) Then
                cboStaffRole.SelectedItem = role
            End If

            txtStaffUsername.ReadOnly = True ' Do not change existing username
            txtStaffPassword.Clear()
            lblStaffMsg.Text = $"Editing staff ID #{_selectedStaffId}"
            lblStaffMsg.ForeColor = Color.Black
        End Sub

        Private Sub btnStaffNew_Click(sender As Object, e As EventArgs) Handles btnStaffNew.Click
            _selectedStaffId = 0
            txtStaffUsername.Clear()
            txtStaffUsername.ReadOnly = False
            txtStaffName.Clear()
            txtStaffPassword.Clear()
            cboStaffRole.SelectedIndex = 0
            lblStaffMsg.Text = "Creating new staff account."
            lblStaffMsg.ForeColor = Color.Black
            txtStaffUsername.Focus()
        End Sub

        Private Sub btnStaffSave_Click(sender As Object, e As EventArgs) Handles btnStaffSave.Click
            lblStaffMsg.Text = String.Empty
            Dim username As String = txtStaffUsername.Text.Trim().ToLowerInvariant()
            Dim fullName As String = txtStaffName.Text.Trim()
            Dim role As String = Convert.ToString(cboStaffRole.SelectedItem)
            Dim pass As String = txtStaffPassword.Text

            If String.IsNullOrWhiteSpace(username) OrElse String.IsNullOrWhiteSpace(fullName) Then
                lblStaffMsg.Text = "Username and Full Name are required."
                lblStaffMsg.ForeColor = Color.Red
                Return
            End If

            If Not ValidationRules.IsValidUsername(username) Then
                lblStaffMsg.Text = "Username must be 3-20 characters; letters, numbers, dots and underscores only (no spaces)."
                lblStaffMsg.ForeColor = Color.Red
                Return
            End If

            If Not ValidationRules.IsValidFullName(fullName) Then
                lblStaffMsg.Text = "Full name must contain only letters, spaces, hyphens, apostrophes, dots and commas."
                lblStaffMsg.ForeColor = Color.Red
                Return
            End If

            ' Fixed role list validation: only Cashier or InventoryClerk may be assigned
            If role <> "Cashier" AndAlso role <> "InventoryClerk" Then
                lblStaffMsg.Text = "Role must be either 'Cashier' or 'InventoryClerk'."
                lblStaffMsg.ForeColor = Color.Red
                Return
            End If

            Try
                If _selectedStaffId = 0 Then
                    ' Creating new staff
                    If Not ValidationRules.IsValidPassword(pass) Then
                        lblStaffMsg.Text = "Initial password must be at least 6 characters and contain no spaces."
                        lblStaffMsg.ForeColor = Color.Red
                        Return
                    End If

                    If _userRepo.UsernameExists(username) Then
                        lblStaffMsg.Text = "Username already exists."
                        lblStaffMsg.ForeColor = Color.Red
                        Return
                    End If

                    Dim salt As String = AuthService.GenerateSalt()
                    Dim hash As String = AuthService.HashPassword(pass, salt)

                    Dim newStaff As User
                    If role = "InventoryClerk" Then
                        newStaff = New InventoryClerk()
                    Else
                        newStaff = New Cashier()
                    End If

                    newStaff.Username = username
                    newStaff.FullName = fullName
                    newStaff.Role = role
                    newStaff.PasswordHash = hash
                    newStaff.Salt = salt
                    newStaff.IsActive = True
                    newStaff.CreatedAt = DateTime.Now

                    _userRepo.CreateUser(newStaff)
                    lblStaffMsg.Text = "Staff account created successfully!"
                    lblStaffMsg.ForeColor = Color.DarkGreen
                Else
                    ' Updating existing staff
                    Dim staff As User = _userRepo.GetById(_selectedStaffId)
                    If staff IsNot Nothing Then
                        staff.FullName = fullName
                        staff.Role = role
                        _userRepo.UpdateUser(staff)
                        lblStaffMsg.Text = "Staff profile updated successfully!"
                        lblStaffMsg.ForeColor = Color.DarkGreen
                    End If
                End If

                LoadStaffList()
            Catch ex As Exception
                lblStaffMsg.Text = ex.Message
                lblStaffMsg.ForeColor = Color.Red
            End Try
        End Sub

        Private Sub btnStaffResetPass_Click(sender As Object, e As EventArgs) Handles btnStaffResetPass.Click
            If _selectedStaffId <= 0 Then
                MessageBox.Show("Please select a staff member to reset password.", "Select Staff", MessageBoxButtons.OK, MessageBoxIcon.Information)
                Return
            End If

            Dim newPass As String = txtStaffPassword.Text
            If Not ValidationRules.IsValidPassword(newPass) Then
                MessageBox.Show("New password must be at least 6 characters and contain no spaces.", "Validation", MessageBoxButtons.OK, MessageBoxIcon.Warning)
                txtStaffPassword.Focus()
                Return
            End If

            Try
                Dim newSalt As String = AuthService.GenerateSalt()
                Dim newHash As String = AuthService.HashPassword(newPass, newSalt)
                If _userRepo.UpdatePassword(_selectedStaffId, newHash, newSalt) Then
                    lblStaffMsg.Text = "Password reset successfully!"
                    lblStaffMsg.ForeColor = Color.DarkGreen
                    txtStaffPassword.Clear()
                End If
            Catch ex As Exception
                lblStaffMsg.Text = ex.Message
                lblStaffMsg.ForeColor = Color.Red
            End Try
        End Sub

        Private Sub btnStaffToggleActive_Click(sender As Object, e As EventArgs) Handles btnStaffToggleActive.Click
            If _selectedStaffId <= 0 Then Return
            Dim staff As User = _userRepo.GetById(_selectedStaffId)
            If staff Is Nothing Then Return

            If staff.UserId = _owner.UserId Then
                MessageBox.Show("Cannot deactivate the primary Store Owner account.", "Restriction", MessageBoxButtons.OK, MessageBoxIcon.Warning)
                Return
            End If

            Dim newStatus As Boolean = Not staff.IsActive
            Dim actionName As String = If(newStatus, "activate", "deactivate")

            If MessageBox.Show($"Are you sure you want to {actionName} account '{staff.Username}'?", "Confirm Status Change", MessageBoxButtons.YesNo, MessageBoxIcon.Question) = DialogResult.Yes Then
                _userRepo.SetActiveStatus(staff.UserId, newStatus)
                LoadStaffList()
                lblStaffMsg.Text = $"Account '{staff.Username}' has been {actionName}d."
                lblStaffMsg.ForeColor = Color.DarkGreen
            End If
        End Sub

        ' =====================================================================
        ' TAB 3: PRODUCT MANAGEMENT (CRUD)
        ' =====================================================================
        Private Sub PopulateCategoriesAndSuppliers()
            cboProdCategory.Items.Clear()
            cboProdCatFilter.Items.Clear()
            cboProdCatFilter.Items.Add("All Categories")

            Dim cats As List(Of Category) = _categoryRepo.GetAll()
            For Each c In cats
                cboProdCategory.Items.Add(c)
                cboProdCatFilter.Items.Add(c)
            Next
            If cboProdCategory.Items.Count > 0 Then cboProdCategory.SelectedIndex = 0
            cboProdCatFilter.SelectedIndex = 0

            cboProdSupplier.Items.Clear()
            Dim sups As List(Of Supplier) = _supplierRepo.GetAll()
            For Each s In sups
                cboProdSupplier.Items.Add(s)
            Next
            If cboProdSupplier.Items.Count > 0 Then cboProdSupplier.SelectedIndex = 0
        End Sub

        Private Sub LoadProductsList()
            dgvProducts.Rows.Clear()
            Dim kw As String = txtProdSearch.Text.Trim()
            Dim catId As Nullable(Of Integer) = Nothing

            If cboProdCatFilter.SelectedIndex > 0 Then
                Dim c = TryCast(cboProdCatFilter.SelectedItem, Category)
                If c IsNot Nothing Then catId = c.CategoryId
            End If

            Dim prods As List(Of Product) = _productRepo.Search(kw, catId, activeOnly:=False)
            For Each p In prods
                Dim statusStr As String = If(p.IsActive, "Yes", "Inactive")
                Dim rowIdx As Integer = dgvProducts.Rows.Add(p.ProductCode, p.Name, p.CategoryName, p.Price.ToString("N2"), p.CostPrice.ToString("N2"), p.StockQty, statusStr)
                If Not p.IsActive Then
                    dgvProducts.Rows(rowIdx).DefaultCellStyle.ForeColor = Color.Gray
                ElseIf p.IsLowStock() Then
                    dgvProducts.Rows(rowIdx).DefaultCellStyle.ForeColor = Color.DarkOrange
                End If
            Next
        End Sub

        Private Sub txtProdSearch_TextChanged(sender As Object, e As EventArgs) Handles txtProdSearch.TextChanged
            LoadProductsList()
        End Sub

        Private Sub cboProdCatFilter_SelectedIndexChanged(sender As Object, e As EventArgs) Handles cboProdCatFilter.SelectedIndexChanged
            LoadProductsList()
        End Sub

        Private Sub dgvProducts_SelectionChanged(sender As Object, e As EventArgs) Handles dgvProducts.SelectionChanged
            If dgvProducts.SelectedRows.Count = 0 Then Return
            Dim code As String = Convert.ToString(dgvProducts.SelectedRows(0).Cells(0).Value)
            Dim prod As Product = _productRepo.GetByCode(code)
            If prod Is Nothing Then Return

            _selectedProductId = prod.ProductId
            txtProdCode.Text = prod.ProductCode
            txtProdName.Text = prod.Name
            txtProdPrice.Text = prod.Price.ToString("F2")
            txtProdCost.Text = prod.CostPrice.ToString("F2")
            numProdStock.Value = Math.Max(0, prod.StockQty)
            numProdReorder.Value = Math.Max(0, prod.ReorderLevel)
            chkProdActive.Checked = prod.IsActive

            If prod.ExpiryDate.HasValue Then
                dtpProdExpiry.Checked = True
                dtpProdExpiry.Value = prod.ExpiryDate.Value
            Else
                dtpProdExpiry.Checked = False
            End If

            ' Match Category
            For i As Integer = 0 To cboProdCategory.Items.Count - 1
                Dim c = TryCast(cboProdCategory.Items(i), Category)
                If c IsNot Nothing AndAlso c.CategoryId = prod.CategoryId Then
                    cboProdCategory.SelectedIndex = i
                    Exit For
                End If
            Next

            ' Match Supplier
            For i As Integer = 0 To cboProdSupplier.Items.Count - 1
                Dim s = TryCast(cboProdSupplier.Items(i), Supplier)
                If s IsNot Nothing AndAlso s.SupplierId = prod.SupplierId Then
                    cboProdSupplier.SelectedIndex = i
                    Exit For
                End If
            Next

            lblProdMsg.Text = $"Editing product #{prod.ProductCode}"
            lblProdMsg.ForeColor = Color.Black
        End Sub

        Private Sub btnProdNew_Click(sender As Object, e As EventArgs) Handles btnProdNew.Click
            _selectedProductId = 0
            txtProdCode.Clear()
            txtProdName.Clear()
            txtProdPrice.Clear()
            txtProdCost.Clear()
            numProdStock.Value = 10
            numProdReorder.Value = 5
            dtpProdExpiry.Checked = False
            chkProdActive.Checked = True
            lblProdMsg.Text = "Ready to add new product."
            lblProdMsg.ForeColor = Color.Black
            txtProdCode.Focus()
        End Sub

        Private Sub btnProdSave_Click(sender As Object, e As EventArgs) Handles btnProdSave.Click
            lblProdMsg.Text = String.Empty

            Dim code As String = txtProdCode.Text.Trim().ToUpperInvariant()
            Dim name As String = txtProdName.Text.Trim()
            Dim price As Decimal = 0D
            Dim cost As Decimal = 0D

            If String.IsNullOrWhiteSpace(code) OrElse String.IsNullOrWhiteSpace(name) Then
                lblProdMsg.Text = "Product code and name are required."
                lblProdMsg.ForeColor = Color.Red
                Return
            End If

            ' BUSINESS RULE: product code charset limited to letters, digits and dash
            If Not System.Text.RegularExpressions.Regex.IsMatch(code, "^[A-Z0-9-]+$") Then
                lblProdMsg.Text = "Product code may only contain letters, numbers and dashes."
                lblProdMsg.ForeColor = Color.Red
                Return
            End If

            If Not Decimal.TryParse(txtProdPrice.Text.Trim(), price) OrElse Not ValidationRules.IsValidMoney(price) OrElse price <= 0D Then
                lblProdMsg.Text = $"Retail price must be a positive amount (max PHP {ValidationRules.MaxMoney:N2}, up to 2 decimals)."
                lblProdMsg.ForeColor = Color.Red
                Return
            End If

            If Not Decimal.TryParse(txtProdCost.Text.Trim(), cost) OrElse Not ValidationRules.IsValidMoney(cost) OrElse cost <= 0D Then
                lblProdMsg.Text = $"Cost price must be a positive amount (max PHP {ValidationRules.MaxMoney:N2}, up to 2 decimals)."
                lblProdMsg.ForeColor = Color.Red
                Return
            End If

            Dim expDate As Nullable(Of DateTime) = If(dtpProdExpiry.Checked, dtpProdExpiry.Value.Date, CType(Nothing, Nullable(Of DateTime)))

            ' BUSINESS RULE: expiry date cannot be in the past
            If expDate.HasValue AndAlso expDate.Value.Date <= DateTime.Today Then
                lblProdMsg.Text = "Expiry date must be in the future."
                lblProdMsg.ForeColor = Color.Red
                Return
            End If

            Dim selectedCat = TryCast(cboProdCategory.SelectedItem, Category)
            Dim selectedSup = TryCast(cboProdSupplier.SelectedItem, Supplier)

            If selectedCat Is Nothing OrElse selectedSup Is Nothing Then
                lblProdMsg.Text = "Please select valid Category and Supplier."
                lblProdMsg.ForeColor = Color.Red
                Return
            End If

            ' BUSINESS RULE: warn if selling price is below the cost price
            If price < cost Then
                Dim warnRes As DialogResult = MessageBox.Show(
                    $"WARNING: Retail price (PHP {price:N2}) is BELOW cost price (PHP {cost:N2})." & vbCrLf &
                    "Saving this product will cause a loss on every sale. Save anyway?",
                    "Price Below Cost", MessageBoxButtons.YesNo, MessageBoxIcon.Warning)
                If warnRes <> DialogResult.Yes Then
                    Return
                End If
            End If

            Try
                Dim p As New Product() With {
                    .ProductId = _selectedProductId,
                    .ProductCode = code,
                    .Name = name,
                    .CategoryId = selectedCat.CategoryId,
                    .SupplierId = selectedSup.SupplierId,
                    .Price = price,
                    .CostPrice = cost,
                    .StockQty = Convert.ToInt32(numProdStock.Value),
                    .ReorderLevel = Convert.ToInt32(numProdReorder.Value),
                    .ExpiryDate = expDate,
                    .IsActive = chkProdActive.Checked
                }

                If _selectedProductId = 0 Then
                    Dim existing = _productRepo.GetByCode(code)
                    If existing IsNot Nothing Then
                        lblProdMsg.Text = "Product code already exists."
                        lblProdMsg.ForeColor = Color.Red
                        Return
                    End If

                    ' BUSINESS RULE: duplicate product name (case-insensitive) rejected
                    Dim dupeName = _productRepo.GetAll(activeOnly:=False).FirstOrDefault(
                        Function(x) x.Name.Trim().Equals(name.Trim(), StringComparison.OrdinalIgnoreCase))
                    If dupeName IsNot Nothing Then
                        lblProdMsg.Text = $"Product name already exists (code {dupeName.ProductCode})."
                        lblProdMsg.ForeColor = Color.Red
                        Return
                    End If

                    _productRepo.Add(p)
                    lblProdMsg.Text = "Product added successfully!"
                    lblProdMsg.ForeColor = Color.DarkGreen
                Else
                    ' Duplicate guard on update (exclude the product being edited)
                    Dim dupeNameUpd = _productRepo.GetAll(activeOnly:=False).FirstOrDefault(
                        Function(x) x.ProductId <> _selectedProductId AndAlso
                                    x.Name.Trim().Equals(name.Trim(), StringComparison.OrdinalIgnoreCase))
                    If dupeNameUpd IsNot Nothing Then
                        lblProdMsg.Text = $"Product name already exists (code {dupeNameUpd.ProductCode})."
                        lblProdMsg.ForeColor = Color.Red
                        Return
                    End If

                    _productRepo.Update(p)
                    lblProdMsg.Text = "Product updated successfully!"
                    lblProdMsg.ForeColor = Color.DarkGreen
                End If

                LoadProductsList()
                LoadOverview()
            Catch ex As Exception
                lblProdMsg.Text = ex.Message
                lblProdMsg.ForeColor = Color.Red
            End Try
        End Sub

        Private Sub btnProdDelete_Click(sender As Object, e As EventArgs) Handles btnProdDelete.Click
            If _selectedProductId <= 0 Then Return

            If MessageBox.Show("Are you sure you want to deactivate this product?", "Confirm Deactivation", MessageBoxButtons.YesNo, MessageBoxIcon.Question) = DialogResult.Yes Then
                _productRepo.DeleteOrDeactivate(_selectedProductId)
                lblProdMsg.Text = "Product has been deactivated."
                lblProdMsg.ForeColor = Color.DarkGreen
                LoadProductsList()
                LoadOverview()
            End If
        End Sub

        ' =====================================================================
        ' TAB 4: CATEGORIES & SUPPLIERS CRUD
        ' =====================================================================
        Private Sub LoadCategoriesList()
            dgvCategories.Rows.Clear()
            Dim cats As List(Of Category) = _categoryRepo.GetAll()
            For Each c In cats
                dgvCategories.Rows.Add(c.CategoryId, c.Name)
            Next
        End Sub

        Private Sub dgvCategories_SelectionChanged(sender As Object, e As EventArgs) Handles dgvCategories.SelectionChanged
            If dgvCategories.SelectedRows.Count = 0 Then Return
            _selectedCategoryId = Convert.ToInt32(dgvCategories.SelectedRows(0).Cells(0).Value)
            txtCatName.Text = Convert.ToString(dgvCategories.SelectedRows(0).Cells(1).Value)
        End Sub

        Private Sub btnCatSave_Click(sender As Object, e As EventArgs) Handles btnCatSave.Click
            Dim name As String = txtCatName.Text.Trim()
            If String.IsNullOrWhiteSpace(name) Then
                MessageBox.Show("Category name cannot be blank.", "Validation", MessageBoxButtons.OK, MessageBoxIcon.Warning)
                Return
            End If

            ' BUSINESS RULE: duplicate category name (case-insensitive) rejected
            Dim dupe As Category = _categoryRepo.GetAll().FirstOrDefault(
                Function(c) c.CategoryId <> _selectedCategoryId AndAlso
                            c.Name.Trim().Equals(name, StringComparison.OrdinalIgnoreCase))
            If dupe IsNot Nothing Then
                MessageBox.Show($"Category '{name}' already exists.", "Validation", MessageBoxButtons.OK, MessageBoxIcon.Warning)
                Return
            End If

            Try
                If _selectedCategoryId = 0 Then
                    _categoryRepo.Add(name)
                Else
                    _categoryRepo.Update(_selectedCategoryId, name)
                End If
                txtCatName.Clear()
                _selectedCategoryId = 0
                LoadCategoriesList()
                PopulateCategoriesAndSuppliers()
            Catch ex As Exception
                MessageBox.Show(ex.Message, "Error", MessageBoxButtons.OK, MessageBoxIcon.Error)
            End Try
        End Sub

        Private Sub btnCatDelete_Click(sender As Object, e As EventArgs) Handles btnCatDelete.Click
            If _selectedCategoryId <= 0 Then Return
            If MessageBox.Show("Are you sure you want to delete this category?", "Confirm Delete", MessageBoxButtons.YesNo, MessageBoxIcon.Question) = DialogResult.Yes Then
                Try
                    _categoryRepo.Delete(_selectedCategoryId)
                    txtCatName.Clear()
                    _selectedCategoryId = 0
                    LoadCategoriesList()
                    PopulateCategoriesAndSuppliers()
                Catch ex As Exception
                    MessageBox.Show(ex.Message, "Cannot Delete", MessageBoxButtons.OK, MessageBoxIcon.Warning)
                End Try
            End If
        End Sub

        Private Sub LoadSuppliersList()
            dgvSuppliers.Rows.Clear()
            Dim sups As List(Of Supplier) = _supplierRepo.GetAll()
            For Each s In sups
                dgvSuppliers.Rows.Add(s.SupplierId, s.Name, s.ContactNumber, s.Address)
            Next
        End Sub

        Private Sub dgvSuppliers_SelectionChanged(sender As Object, e As EventArgs) Handles dgvSuppliers.SelectionChanged
            If dgvSuppliers.SelectedRows.Count = 0 Then Return
            _selectedSupplierId = Convert.ToInt32(dgvSuppliers.SelectedRows(0).Cells(0).Value)
            txtSupName.Text = Convert.ToString(dgvSuppliers.SelectedRows(0).Cells(1).Value)
            txtSupContact.Text = Convert.ToString(dgvSuppliers.SelectedRows(0).Cells(2).Value)
            txtSupAddr.Text = Convert.ToString(dgvSuppliers.SelectedRows(0).Cells(3).Value)
        End Sub

        Private Sub btnSupNew_Click(sender As Object, e As EventArgs) Handles btnSupNew.Click
            _selectedSupplierId = 0
            txtSupName.Clear()
            txtSupContact.Clear()
            txtSupAddr.Clear()
            txtSupName.Focus()
        End Sub

        Private Sub btnSupSave_Click(sender As Object, e As EventArgs) Handles btnSupSave.Click
            Dim name As String = txtSupName.Text.Trim()
            If String.IsNullOrWhiteSpace(name) Then
                MessageBox.Show("Supplier name cannot be blank.", "Validation", MessageBoxButtons.OK, MessageBoxIcon.Warning)
                Return
            End If

            If Not ValidationRules.IsValidContactNumber(txtSupContact.Text.Trim()) Then
                MessageBox.Show("Contact number must be 7-15 digits (optionally starting with +).", "Validation", MessageBoxButtons.OK, MessageBoxIcon.Warning)
                txtSupContact.Focus()
                Return
            End If

            ' BUSINESS RULE: duplicate supplier name (case-insensitive) rejected
            Dim dupe As Supplier = _supplierRepo.GetAll().FirstOrDefault(
                Function(s) s.SupplierId <> _selectedSupplierId AndAlso
                            s.Name.Trim().Equals(name, StringComparison.OrdinalIgnoreCase))
            If dupe IsNot Nothing Then
                MessageBox.Show($"Supplier '{name}' already exists.", "Validation", MessageBoxButtons.OK, MessageBoxIcon.Warning)
                Return
            End If

            Dim sup As New Supplier(_selectedSupplierId, name, txtSupContact.Text.Trim(), txtSupAddr.Text.Trim())

            Try
                If _selectedSupplierId = 0 Then
                    _supplierRepo.Add(sup)
                Else
                    _supplierRepo.Update(sup)
                End If
                btnSupNew.PerformClick()
                LoadSuppliersList()
                PopulateCategoriesAndSuppliers()
            Catch ex As Exception
                MessageBox.Show(ex.Message, "Error", MessageBoxButtons.OK, MessageBoxIcon.Error)
            End Try
        End Sub

        Private Sub btnSupDelete_Click(sender As Object, e As EventArgs) Handles btnSupDelete.Click
            If _selectedSupplierId <= 0 Then Return
            If MessageBox.Show("Are you sure you want to delete this supplier?", "Confirm Delete", MessageBoxButtons.YesNo, MessageBoxIcon.Question) = DialogResult.Yes Then
                Try
                    _supplierRepo.Delete(_selectedSupplierId)
                    btnSupNew.PerformClick()
                    LoadSuppliersList()
                    PopulateCategoriesAndSuppliers()
                Catch ex As Exception
                    MessageBox.Show(ex.Message, "Cannot Delete", MessageBoxButtons.OK, MessageBoxIcon.Warning)
                End Try
            End If
        End Sub

        ' =====================================================================
        ' TAB 5: TRANSACTION HISTORY
        ' =====================================================================
        Private Sub LoadTransactions()
            dgvTransactions.Rows.Clear()
            Dim sDate As DateTime = dtpTxStart.Value.Date
            Dim eDate As DateTime = dtpTxEnd.Value.Date.AddDays(1).AddSeconds(-1)

            Dim sales As List(Of Sale) = _saleRepo.GetSalesByDateRange(sDate, eDate)
            For Each s In sales
                dgvTransactions.Rows.Add(s.ReceiptNo, s.SaleDate.ToString("yyyy-MM-dd hh:mm tt"), s.CashierName, s.CustomerName, s.Subtotal.ToString("N2"), s.DiscountAmount.ToString("N2"), s.Total.ToString("N2"))
            Next
        End Sub

        Private Sub TxDateFilter_Changed(sender As Object, e As EventArgs) Handles dtpTxStart.ValueChanged, dtpTxEnd.ValueChanged
            LoadTransactions()
        End Sub

        Private Sub btnTxViewDetails_Click(sender As Object, e As EventArgs) Handles btnTxViewDetails.Click
            If dgvTransactions.SelectedRows.Count = 0 Then
                MessageBox.Show("Please select a transaction from the list.", "Notice", MessageBoxButtons.OK, MessageBoxIcon.Information)
                Return
            End If

            Dim receiptNo As String = Convert.ToString(dgvTransactions.SelectedRows(0).Cells(0).Value)
            Dim sale As Sale = _saleRepo.GetByReceiptNo(receiptNo)
            If sale IsNot Nothing Then
                Using dlg As New PastReceiptForm(sale)
                    dlg.ShowDialog(Me)
                End Using
            End If
        End Sub

        ' =====================================================================
        ' TAB 6: REPORTS & ANALYTICS
        ' =====================================================================
        Private Sub PopulateReportTypes()
            cboReportType.Items.Clear()
            cboReportType.Items.Add("Daily Sales Report")
            cboReportType.Items.Add("Inventory Valuation Report")
            cboReportType.Items.Add("Low-Stock & Expiry Alert Report")
            cboReportType.SelectedIndex = 0
        End Sub

        Private Sub cboReportType_SelectedIndexChanged(sender As Object, e As EventArgs) Handles cboReportType.SelectedIndexChanged
            dtpReportDate.Enabled = (cboReportType.SelectedIndex = 0)
        End Sub

        Private Sub btnGenerateReport_Click(sender As Object, e As EventArgs) Handles btnGenerateReport.Click
            Dim stats As Dictionary(Of String, Object) = Nothing
            Dim selectedIdx As Integer = cboReportType.SelectedIndex

            Select Case selectedIdx
                Case 0
                    _currentReportData = ReportService.GenerateDailySalesReport(dtpReportDate.Value.Date, stats)
                    lblReportSummary.Text = $"Daily Sales for {dtpReportDate.Value:yyyy-MM-dd}: {stats("TransactionCount")} Transactions | Gross: PHP {CDec(stats("TotalGross")):N2} | Discounts: PHP {CDec(stats("TotalDiscount")):N2} | Net Total: PHP {CDec(stats("TotalNet")):N2}"
                Case 1
                    _currentReportData = ReportService.GenerateInventoryReport(stats)
                    lblReportSummary.Text = $"Inventory Valuation: {stats("TotalProducts")} products ({stats("TotalUnits")} units) | Cost Valuation: PHP {CDec(stats("TotalCostValue")):N2} | Retail Valuation: PHP {CDec(stats("TotalRetailValue")):N2} | Potential Gross Margin: PHP {CDec(stats("PotentialProfit")):N2}"
                Case 2
                    _currentReportData = ReportService.GenerateLowStockAndExpiryReport()
                    lblReportSummary.Text = $"Alert Summary: {_currentReportData.Rows.Count} critical item(s) flagged for low-stock or expiration."
            End Select

            dgvReport.DataSource = _currentReportData
        End Sub

        Private Sub btnExportCsv_Click(sender As Object, e As EventArgs) Handles btnExportCsv.Click
            If _currentReportData Is Nothing OrElse _currentReportData.Rows.Count = 0 Then
                MessageBox.Show("Please generate a report with data before exporting to CSV.", "No Data", MessageBoxButtons.OK, MessageBoxIcon.Information)
                Return
            End If

            Using sfd As New SaveFileDialog()
                sfd.Filter = "CSV Files (*.csv)|*.csv"
                sfd.FileName = $"{cboReportType.SelectedItem.ToString().Replace(" ", "_")}_{DateTime.Now:yyyyMMdd_HHmmss}.csv"
                If sfd.ShowDialog(Me) = DialogResult.OK Then
                    Try
                        ReportService.ExportDataTableToCsv(_currentReportData, sfd.FileName)
                        MessageBox.Show($"Report successfully exported to:{vbCrLf}{sfd.FileName}", "Export Complete", MessageBoxButtons.OK, MessageBoxIcon.Information)
                    Catch ex As Exception
                        MessageBox.Show($"Export error: {ex.Message}", "Error", MessageBoxButtons.OK, MessageBoxIcon.Error)
                    End Try
                End If
            End Using
        End Sub

        Private Sub btnPrintReport_Click(sender As Object, e As EventArgs) Handles btnPrintReport.Click
            If _currentReportData Is Nothing OrElse _currentReportData.Rows.Count = 0 Then
                MessageBox.Show("Please generate a report before printing.", "No Data", MessageBoxButtons.OK, MessageBoxIcon.Information)
                Return
            End If

            Try
                ReportService.PrintReport(cboReportType.SelectedItem.ToString(), _currentReportData, Me)
            Catch ex As Exception
                MessageBox.Show($"Print preview error: {ex.Message}", "Print Error", MessageBoxButtons.OK, MessageBoxIcon.Error)
            End Try
        End Sub

        Private Sub btnLogout_Click(sender As Object, e As EventArgs) Handles btnLogout.Click
            If MessageBox.Show("Are you sure you want to log out of the owner administration?", "Confirm Logout", MessageBoxButtons.YesNo, MessageBoxIcon.Question) = DialogResult.Yes Then
                Me.Close()
            End If
        End Sub

        ' =====================================================================
        ' TYPING-LEVEL INPUT VALIDATION
        ' =====================================================================
        Private Sub BlockSpaceTyping(sender As Object, e As KeyPressEventArgs) _
            Handles txtStaffUsername.KeyPress, txtStaffPassword.KeyPress
            If Char.IsWhiteSpace(e.KeyChar) Then
                e.Handled = True
                lblStaffMsg.Text = "Spaces are not allowed in this field."
                lblStaffMsg.ForeColor = Color.Red
            End If
        End Sub

        Private Sub BlockDigitsInNames(sender As Object, e As KeyPressEventArgs) _
            Handles txtStaffName.KeyPress
            Dim allowedPunct As String = " '-.,"
            If Not (Char.IsControl(e.KeyChar) OrElse Char.IsLetter(e.KeyChar) OrElse
                    Char.IsWhiteSpace(e.KeyChar) OrElse allowedPunct.Contains(e.KeyChar.ToString())) Then
                e.Handled = True
                lblStaffMsg.Text = "Numbers and symbols are not allowed in this field."
                lblStaffMsg.ForeColor = Color.Red
            End If
        End Sub
    End Class
End Namespace
