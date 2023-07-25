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
Public Class Form10

    Dim WithEvents excelApp As Excel.Application
    Dim workBook As Excel.Workbook
    Dim workSheet As Excel.Worksheet
    Dim workSheet2 As Excel.Worksheet
    Dim rng As Excel.Range
    Dim rng2 As Excel.Range
    Dim selectedRange As Excel.Range

    Dim opened As Integer
    Dim FocusedTextBox As Integer
    Private Function SearchInArray(i, j, Arr)

        Dim Result As Integer = 0

        For k = LBound(Arr, 1) To UBound(Arr, 1)
            If Arr(k, 0) = i And Arr(k, 1) = j Then
                Result = Arr(k, 2)
                Exit For
            End If
        Next k

        SearchInArray = Result

    End Function
    Private Function Available(i, j, Arr)

        Dim Result As Boolean = False

        For k = LBound(Arr, 1) To UBound(Arr, 1)
            If Arr(k, 0) = i And Arr(k, 1) = j Then
                Result = True
                Exit For
            End If
        Next k

        Available = Result

    End Function
    Private Sub Display()

        CustomPanel1.Controls.Clear()
        CustomPanel2.Controls.Clear()

        Dim Rng As Excel.Range
        Rng = workSheet.Range(TextBox1.Text)
        Rng.Select()

        Dim displayRng As Excel.Range

        If Rng.Rows.Count > 50 Then
            displayRng = workSheet.Range(Rng.Cells(1, 1), Rng.Cells(50, Rng.Columns.Count))
        Else
            displayRng = workSheet.Range(Rng.Cells(1, 1), Rng.Cells(Rng.Rows.Count, Rng.Columns.Count))
        End If

        Dim r As Integer
        Dim c As Integer

        r = displayRng.Rows.Count
        c = displayRng.Columns.Count

        Dim height As Integer
        Dim width As Integer

        If r <= 6 Then
            height = CustomPanel1.Height / r
        Else
            height = CustomPanel1.Height / 6
        End If

        If c <= 4 Then
            width = CustomPanel1.Width / c
        Else
            width = CustomPanel1.Width / 4
        End If

        Dim Arr((r * c) - 1, 1) As Object

        Dim Count As Integer = 0

        For i = 1 To r
            For j = 1 To C
                If Available(i, j, Arr) = False Then
                    If rng.Cells(i, j).MergeCells = True Then
                        For k = 2 To rng.Cells(i, j).MergeArea.Columns.Count
                            If Available(i, j + k - 1, Arr) = False Then
                                Arr(Count, 0) = i
                                Arr(Count, 1) = j + k - 1
                                Count = Count + 1
                            End If
                        Next k
                        For m = 2 To rng.Cells(i, j).MergeArea.Rows.Count
                            For n = 1 To rng.Cells(i, j).MergeArea.Columns.Count
                                If Available(i + m - 1, j + n - 1, Arr) = False Then
                                    Arr(Count, 0) = i + m - 1
                                    Arr(Count, 1) = j + n - 1
                                    Count = Count + 1
                                End If
                            Next n
                        Next m
                    End If
                End If
            Next j
        Next i

        For i = 1 To r
            For j = 1 To C
                If Available(i, j, Arr) = False Then
                    l2 = l * rng.Cells(i, j).MergeArea.Rows.Count
                    w2 = w * rng.Cells(i, j).MergeArea.Columns.Count
            Set cmdLots = UserForm12.Frame5.Controls.Add("Forms.Label.1", "lbl1")
            With cmdLots
                        .Top = (i - 1) * l
                        .Left = ((j - 1) * w)
                        .BackColor = TColor
                        .Caption = rng.Cells(i, j)
                        .Width = w2
                        .Height = l2
                        .BorderStyle = fmBorderStyleSingle
                        .TextAlign = fmTextAlignCenter
                        .Font.Name = "Times New Roman"
                        .Font.Bold = True
                        .ForeColor = FColor
                    End With
                End If
            Next j
        Next i


        With UserForm12.Frame5
            .ScrollBars = fmScrollBarsBoth
            .ScrollHeight = l * r
            .ScrollWidth = w * C
        End With

        Dim Arr2() As Variant
        ReDim Arr2((r * C) - 1, 2)

        Count = 0

        For i = 1 To r
            For j = 1 To C
                If rng.Cells(i, j).MergeCells = True And Available(i, j, Arr2) = False Then
                    For m = 0 To rng.Cells(i, j).MergeArea.Rows.Count - 1
                        For n = 0 To rng.Cells(i, j).MergeArea.Columns.Count - 1
                            Arr2(Count, 0) = i + m
                            Arr2(Count, 1) = j + n
                            Arr2(Count, 2) = rng.Cells(i, j)
                            Count = Count + 1
                        Next n
                    Next m
                End If
            Next j
        Next i

        For i = 1 To r
            For j = 1 To C
        Set cmdLots = UserForm12.Frame4.Controls.Add("Forms.Label.1", "lbl1")
        With cmdLots
                    .Top = (i - 1) * l
                    .Left = ((j - 1) * w)
                    .BackColor = TColor
                    If rng.Cells(i, j).MergeCells = True Then
                        .Caption = SearchInArray(i, j, Arr2)
                    Else
                        .Caption = rng.Cells(i, j)
                    End If
                    .Width = w
                    .Height = l
                    .BorderStyle = fmBorderStyleSingle
                    .TextAlign = fmTextAlignCenter
                    .Font.Name = "Times New Roman"
                    .Font.Bold = True
                    .ForeColor = FColor
                End With
            Next j
        Next i

        With UserForm12.Frame4
            .ScrollBars = fmScrollBarsBoth
            .ScrollHeight = l * r
            .ScrollWidth = w * C
        End With
    End Sub

    Private Sub Button2_Click(sender As Object, e As EventArgs) Handles Button2.Click

    End Sub
End Class