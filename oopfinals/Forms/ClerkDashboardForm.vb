' ==============================================================================
' Project: Mini-Grocery Store POS & Daily Inventory System
' Course: Object-Oriented Programming (ITP 102) - University of Caloocan City
' Form: ClerkDashboardForm (Code-Behind)
' OOP Concept: TRANSACTION INTEGRITY, AUDIT TRAIL, ROLE RESTRICTION
' Provides inventory operations for the Clerk: stock alerts with color coding,
' stock-in inbound deliveries, manual adjustments, audit logs, and read-only catalog.
' ==============================================================================

Option Strict On
Option Explicit On

Imports System
Imports System.Collections.Generic
Imports System.Drawing
Imports System.Windows.Forms
Imports oopfinals.Data
Imports oopfinals.Models

Namespace Forms
    Public Class ClerkDashboardForm
        Private ReadOnly _clerk As InventoryClerk
        Private ReadOnly _productRepo As New ProductRepository()
        Private ReadOnly _supplierRepo As New SupplierRepository()
        Private ReadOnly _stockRepo As New StockRepository()

        Public Sub New(clerk As InventoryClerk)
            InitializeComponent()
            _clerk = clerk
            Me.DoubleBuffered = True
        End Sub

        Private Sub ClerkDashboardForm_Resize(sender As Object, e As EventArgs) Handles MyBase.Resize
            Me.Refresh()
        End Sub

        Private Sub ClerkDashboardForm_Load(sender As Object, e As EventArgs) Handles MyBase.Load
            ' SECURITY: verify role on form load
            If _clerk.Role <> "InventoryClerk" Then
                MessageBox.Show("Access denied: this dashboard is restricted to Inventory Clerk accounts.", "Access Denied", MessageBoxButtons.OK, MessageBoxIcon.Warning)
                Me.BeginInvoke(New MethodInvoker(Sub() Me.Close()))
                Return
            End If

            lblClerkInfo.Text = $"Inventory Clerk: {_clerk.FullName} (@{_clerk.Username})"

            PopulateSuppliers()
            PopulateAdjustmentReasons()
            PopulateMovementTypes()
            PopulateProductSelectors()

            dtpMovStart.Value = DateTime.Today.AddDays(-7)
            dtpMovEnd.Value = DateTime.Today

            LoadAlerts()
            LoadMovements()
            LoadCatalog()
        End Sub

        Private Sub PopulateSuppliers()
            cboStockInSupplier.Items.Clear()
            Dim suppliers As List(Of Supplier) = _supplierRepo.GetAll()
            For Each s In suppliers
                cboStockInSupplier.Items.Add(s)
            Next
            If cboStockInSupplier.Items.Count > 0 Then cboStockInSupplier.SelectedIndex = 0
        End Sub

        Private Sub PopulateProductSelectors()
            cboStockInProduct.Items.Clear()
            cboAdjProduct.Items.Clear()

            Dim prods As List(Of Product) = _productRepo.GetAll(activeOnly:=True)
            For Each p In prods
                cboStockInProduct.Items.Add(p)
                cboAdjProduct.Items.Add(p)
            Next

            If cboStockInProduct.Items.Count > 0 Then cboStockInProduct.SelectedIndex = 0
            If cboAdjProduct.Items.Count > 0 Then cboAdjProduct.SelectedIndex = 0
        End Sub

        Private Sub PopulateAdjustmentReasons()
            cboAdjType.Items.Clear()
            cboAdjType.Items.Add("Deduct / Decrease (-)")
            cboAdjType.Items.Add("Add / Increase (+)")
            cboAdjType.SelectedIndex = 0

            cboAdjReason.Items.Clear()
            cboAdjReason.Items.Add("Damaged Goods / Spillage")
            cboAdjReason.Items.Add("Expired Stock Disposal")
            cboAdjReason.Items.Add("Physical Inventory Count Correction")
            cboAdjReason.Items.Add("Customer Return / Restocked")
            cboAdjReason.Items.Add("Defective / Factory Recall")
            cboAdjReason.SelectedIndex = 0
        End Sub

        Private Sub PopulateMovementTypes()
            cboMovType.Items.Clear()
            cboMovType.Items.Add("All")
            cboMovType.Items.Add("StockIn")
            cboMovType.Items.Add("Adjustment")
            cboMovType.Items.Add("Sale")
            cboMovType.SelectedIndex = 0
        End Sub

        ''' <summary>
        ''' Loads inventory alerts and highlights items with color coding:
        ''' Red = Expired, DarkOrange = Low Stock, Gold = Expiring Soon.
        ''' </summary>
        Private Sub LoadAlerts()
            dgvAlerts.Rows.Clear()
            Dim products As List(Of Product) = _productRepo.GetAll(activeOnly:=True)

            Dim lowCount As Integer = 0
            Dim expSoonCount As Integer = 0
            Dim expiredCount As Integer = 0

            For Each p In products
                Dim isLow As Boolean = p.IsLowStock()
                Dim isExp As Boolean = p.IsExpired()
                Dim isExpSoon As Boolean = p.IsExpiringSoon(30)

                If isLow Then lowCount += 1
                If isExpSoon Then expSoonCount += 1
                If isExp Then expiredCount += 1

                ' Filter selection
                Dim showRow As Boolean = False
                If rbAllAlerts.Checked AndAlso (isLow OrElse isExp OrElse isExpSoon) Then
                    showRow = True
                ElseIf rbLowStock.Checked AndAlso isLow Then
                    showRow = True
                ElseIf rbExpiring.Checked AndAlso isExpSoon Then
                    showRow = True
                ElseIf rbExpired.Checked AndAlso isExp Then
                    showRow = True
                End If

                If showRow Then
                    Dim statusText As String = ""
                    Dim statusColor As Color = Color.Black

                    If isExp Then
                        statusText = "⛔ EXPIRED"
                        statusColor = Color.Red
                    ElseIf isLow AndAlso isExpSoon Then
                        statusText = "⚠️ LOW & EXPIRING"
                        statusColor = Color.DarkRed
                    ElseIf isLow Then
                        statusText = "⚠️ LOW STOCK"
                        statusColor = Color.DarkOrange
                    ElseIf isExpSoon Then
                        statusText = "⏳ EXPIRING SOON"
                        statusColor = Color.DarkGoldenrod
                    End If

                    Dim expDateStr As String = If(p.ExpiryDate.HasValue, p.ExpiryDate.Value.ToString("yyyy-MM-dd"), "N/A")
                    Dim rowIdx As Integer = dgvAlerts.Rows.Add(p.ProductCode, p.Name, p.CategoryName, p.StockQty, p.ReorderLevel, expDateStr, statusText)
                    dgvAlerts.Rows(rowIdx).DefaultCellStyle.ForeColor = statusColor
                    dgvAlerts.Rows(rowIdx).DefaultCellStyle.Font = New Font(dgvAlerts.Font, FontStyle.Bold)
                End If
            Next

            lblKpiSummary.Text = $"Alerts: {lowCount} Low Stock | {expSoonCount} Expiring Soon | {expiredCount} Expired"
        End Sub

        Private Sub rbAlertFilter_CheckedChanged(sender As Object, e As EventArgs) Handles rbAllAlerts.CheckedChanged, rbLowStock.CheckedChanged, rbExpiring.CheckedChanged, rbExpired.CheckedChanged
            Dim rb = TryCast(sender, RadioButton)
            If rb IsNot Nothing AndAlso rb.Checked Then
                LoadAlerts()
            End If
        End Sub

        Private Sub btnRefreshAlerts_Click(sender As Object, e As EventArgs) Handles btnRefreshAlerts.Click
            LoadAlerts()
        End Sub

        Private Sub chkUpdateExpiry_CheckedChanged(sender As Object, e As EventArgs) Handles chkUpdateExpiry.CheckedChanged
            dtpStockInExpiry.Enabled = chkUpdateExpiry.Checked
        End Sub

        Private Sub btnSubmitStockIn_Click(sender As Object, e As EventArgs) Handles btnSubmitStockIn.Click
            lblStockInMsg.Text = String.Empty

            Dim selectedProd As Product = TryCast(cboStockInProduct.SelectedItem, Product)
            If selectedProd Is Nothing Then
                lblStockInMsg.Text = "Please select a product."
                lblStockInMsg.ForeColor = Color.Red
                Return
            End If

            Dim selectedSup As Supplier = TryCast(cboStockInSupplier.SelectedItem, Supplier)
            Dim supId As Nullable(Of Integer) = If(selectedSup IsNot Nothing, selectedSup.SupplierId, CType(Nothing, Nullable(Of Integer)))

            Dim qty As Integer = Convert.ToInt32(numStockInQty.Value)
            Dim newExpiry As Nullable(Of DateTime) = Nothing
            If chkUpdateExpiry.Checked Then
                newExpiry = dtpStockInExpiry.Value.Date

                ' BUSINESS RULE: expired stock received must be rejected (record disposal instead)
                If newExpiry.Value.Date <= DateTime.Today Then
                    lblStockInMsg.Text = "Stock with an expired or today-expiring date must be rejected, not received." & vbCrLf &
                                         "Use the Disposal adjustment for expired stock."
                    lblStockInMsg.ForeColor = Color.Red
                    Return
                End If
            End If

            Dim notes As String = txtStockInNotes.Text.Trim()
            If String.IsNullOrWhiteSpace(notes) Then notes = "Supplier Delivery"

            Try
                If _stockRepo.RecordStockIn(selectedProd.ProductId, qty, supId, _clerk.UserId, newExpiry, notes) Then
                    lblStockInMsg.Text = $"Successfully stocked in {qty} units of '{selectedProd.Name}'!"
                    lblStockInMsg.ForeColor = Color.DarkGreen
                    txtStockInNotes.Clear()
                    numStockInQty.Value = 10
                    chkUpdateExpiry.Checked = False

                    ' Reload dependent views
                    LoadAlerts()
                    LoadMovements()
                    LoadCatalog()
                    PopulateProductSelectors()
                Else
                    lblStockInMsg.Text = "Failed to record stock-in."
                    lblStockInMsg.ForeColor = Color.Red
                End If
            Catch ex As Exception
                lblStockInMsg.Text = ex.Message
                lblStockInMsg.ForeColor = Color.Red
            End Try
        End Sub

        Private Sub cboAdjProduct_SelectedIndexChanged(sender As Object, e As EventArgs) Handles cboAdjProduct.SelectedIndexChanged
            Dim p As Product = TryCast(cboAdjProduct.SelectedItem, Product)
            If p IsNot Nothing Then
                lblAdjCurrentStock.Text = $"Current Stock: {p.StockQty} unit(s)"
            End If
        End Sub

        Private Sub btnSubmitAdjustment_Click(sender As Object, e As EventArgs) Handles btnSubmitAdjustment.Click
            lblAdjMsg.Text = String.Empty

            Dim selectedProd As Product = TryCast(cboAdjProduct.SelectedItem, Product)
            If selectedProd Is Nothing Then
                lblAdjMsg.Text = "Please select a product."
                lblAdjMsg.ForeColor = Color.Red
                Return
            End If

            Dim rawQty As Integer = Convert.ToInt32(numAdjQty.Value)
            Dim isDeduct As Boolean = cboAdjType.SelectedIndex = 0
            Dim qtyChange As Integer = If(isDeduct, -rawQty, rawQty)

            Dim reasonCategory As String = Convert.ToString(cboAdjReason.SelectedItem)
            Dim userDetails As String = txtAdjNotes.Text.Trim()

            ' BUSINESS RULE: spoilage/damage/expiry deductions require a specific written reason
            Dim isSpoilage As Boolean = isDeduct AndAlso
                (reasonCategory = "Damaged Goods / Spillage" OrElse reasonCategory = "Expired Stock Disposal")
            If isSpoilage AndAlso String.IsNullOrWhiteSpace(userDetails) Then
                lblAdjMsg.Text = $"A reason detail in Notes is required for '{reasonCategory}' deductions."
                lblAdjMsg.ForeColor = Color.Red
                txtAdjNotes.Focus()
                Return
            End If

            Dim fullReason As String = If(String.IsNullOrWhiteSpace(userDetails), reasonCategory, $"{reasonCategory} - {userDetails}")

            Try
                If _stockRepo.RecordAdjustment(selectedProd.ProductId, qtyChange, fullReason, _clerk.UserId) Then
                    lblAdjMsg.Text = $"Adjustment applied successfully: {qtyChange:+0;-0} unit(s) for '{selectedProd.Name}'!"
                    lblAdjMsg.ForeColor = Color.DarkGreen
                    txtAdjNotes.Clear()
                    numAdjQty.Value = 1

                    ' Reload dependent views
                    LoadAlerts()
                    LoadMovements()
                    LoadCatalog()
                    PopulateProductSelectors()
                Else
                    lblAdjMsg.Text = "Failed to apply adjustment."
                    lblAdjMsg.ForeColor = Color.Red
                End If
            Catch ex As Exception
                lblAdjMsg.Text = ex.Message
                lblAdjMsg.ForeColor = Color.Red
            End Try
        End Sub

        Private Sub LoadMovements()
            dgvMovements.Rows.Clear()
            Dim mType As String = Convert.ToString(cboMovType.SelectedItem)
            Dim sDate As DateTime = dtpMovStart.Value.Date
            Dim eDate As DateTime = dtpMovEnd.Value.Date.AddDays(1).AddSeconds(-1)

            Dim movements As List(Of StockMovement) = _stockRepo.GetMovements(Nothing, mType, sDate, eDate)
            For Each m In movements
                Dim rowIdx As Integer = dgvMovements.Rows.Add(m.MovementId, m.MovementDate.ToString("yyyy-MM-dd hh:mm tt"), m.ProductCode, m.ProductName, m.MovementType, m.QuantityChange.ToString("+0;-0;0"), m.Reason, m.SupplierName, m.UserName)

                If m.QuantityChange > 0 Then
                    dgvMovements.Rows(rowIdx).Cells(5).Style.ForeColor = Color.DarkGreen
                Else
                    dgvMovements.Rows(rowIdx).Cells(5).Style.ForeColor = Color.Red
                End If
            Next
        End Sub

        Private Sub MovementFilter_Changed(sender As Object, e As EventArgs) Handles cboMovType.SelectedIndexChanged, dtpMovStart.ValueChanged, dtpMovEnd.ValueChanged
            LoadMovements()
        End Sub

        Private Sub LoadCatalog()
            dgvCatalog.Rows.Clear()
            Dim kw As String = txtCatalogSearch.Text.Trim()
            Dim prods As List(Of Product) = _productRepo.Search(kw, Nothing, activeOnly:=True)

            For Each p In prods
                Dim expStr As String = If(p.ExpiryDate.HasValue, p.ExpiryDate.Value.ToString("yyyy-MM-dd"), "N/A")
                dgvCatalog.Rows.Add(p.ProductCode, p.Name, p.CategoryName, p.Price.ToString("N2"), p.CostPrice.ToString("N2"), p.StockQty, p.ReorderLevel, expStr)
            Next
        End Sub

        Private Sub txtCatalogSearch_TextChanged(sender As Object, e As EventArgs) Handles txtCatalogSearch.TextChanged
            LoadCatalog()
        End Sub

        Private Sub btnLogout_Click(sender As Object, e As EventArgs) Handles btnLogout.Click
            If MessageBox.Show("Are you sure you want to log out?", "Confirm Logout", MessageBoxButtons.YesNo, MessageBoxIcon.Question) = DialogResult.Yes Then
                Me.Close()
            End If
        End Sub
    End Class
End Namespace
