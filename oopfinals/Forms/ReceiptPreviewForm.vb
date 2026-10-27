' ==============================================================================
' Project: Mini-Grocery Store POS & Daily Inventory System
' Course: Object-Oriented Programming (ITP 102) - University of Caloocan City
' Form: ReceiptPreviewForm (Code-Behind)
' Displays the formatted official receipt and facilitates printing.
' ==============================================================================

Option Strict On
Option Explicit On

Imports System
Imports System.Windows.Forms
Imports oopfinals.Models
Imports oopfinals.Services

Namespace Forms
    Public Class ReceiptPreviewForm
        Private ReadOnly _sale As Sale
        Private ReadOnly _printer As ReceiptPrinter

        Public Sub New(sale As Sale)
            InitializeComponent()
            _sale = sale
            _printer = New ReceiptPrinter(_sale)
        End Sub

        Private Sub ReceiptPreviewForm_Load(sender As Object, e As EventArgs) Handles MyBase.Load
            If _sale IsNot Nothing Then
                txtReceipt.Text = _printer.GenerateReceiptText()
                txtReceipt.SelectionStart = 0
                txtReceipt.SelectionLength = 0
            End If
        End Sub

        Private Sub btnPrint_Click(sender As Object, e As EventArgs) Handles btnPrint.Click
            Try
                _printer.ShowPrintPreview(Me)
            Catch ex As Exception
                MessageBox.Show($"Print preview error: {ex.Message}", "Print Error", MessageBoxButtons.OK, MessageBoxIcon.Error)
            End Try
        End Sub

        Private Sub btnClose_Click(sender As Object, e As EventArgs) Handles btnClose.Click
            Me.Close()
        End Sub
    End Class
End Namespace
