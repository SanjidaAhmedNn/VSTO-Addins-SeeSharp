Imports Excel = Microsoft.Office.Interop.Excel
Imports System.Runtime.InteropServices
Imports System.Drawing
Imports Microsoft.Office.Interop.Excel
Imports System.Net.Mime.MediaTypeNames
Imports System.Reflection
Imports System.Windows.Forms
Imports System.Windows.Forms.VisualStyles.VisualStyleElement
Imports System.Diagnostics
Imports System.Text.RegularExpressions

Public Class Form8
    Dim WithEvents excelApp As Excel.Application

    Dim workBook As Excel.Workbook
    Dim workbook2 As Excel.Workbook

    Dim workSheet As Excel.Worksheet
    Dim workSheet2 As Excel.Worksheet

    Dim rng As Excel.Range
    Dim rng2 As Excel.Range

    Dim opened As Integer
    Dim FocusedTextBox As Integer

    Private Sub Form8_Load(sender As Object, e As EventArgs) Handles Me.Load

        Try
            excelApp = Globals.ThisAddIn.Application
            workBook = excelApp.ActiveWorkbook
            workbook2 = excelApp.ActiveWorkbook
            workSheet = workBook.ActiveSheet
            workSheet2 = workbook2.ActiveSheet

            AddHandler excelApp.SheetSelectionChange, AddressOf excelApp_SheetSelectionChange

            opened = opened + 1

            Me.Label2.Enabled = False
            Me.TextBox3.Enabled = False
            Me.PictureBox6.Enabled = False

        Catch ex As Exception

        End Try

    End Sub

    Private Sub excelApp_SheetSelectionChange(ByVal Sh As Object, ByVal Target As Excel.Range)

        Try

            Dim selectedRange As Excel.Range
            selectedRange = excelApp.Selection

            If FocusedTextBox = 1 Then
                TextBox1.Text = selectedRange.Address
                workSheet = workBook.ActiveSheet
                rng = selectedRange
                TextBox1.Focus()

            ElseIf FocusedTextBox = 3 Then
                TextBox3.Text = selectedRange.Address
                workSheet2 = workbook2.ActiveSheet
                rng2 = selectedRange
                TextBox3.Focus()
            End If

        Catch ex As Exception

        End Try

    End Sub

End Class