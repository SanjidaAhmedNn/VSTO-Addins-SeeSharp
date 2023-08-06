Imports System.Drawing

Public Class Form22_Merge_Duplicate_Rows

    Public WithEvents excelApp As Excel.Application
    Public workBook As Excel.Workbook
    Public workSheet As Excel.Worksheet
    Public workSheet2 As Excel.Worksheet
    Public rng As Excel.Range
    Public rng2 As Excel.Range
    Public selectedRange As Excel.Range

    Public opened As Integer
    Public FocusedTextBox As Integer

    Private Sub Label3_Click(sender As Object, e As EventArgs) Handles Label3.Click
        Label3.BackColor = SystemColors.Control
    End Sub

End Class