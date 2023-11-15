Imports System.Data
Imports System.Runtime.InteropServices
Imports System.Windows.Forms
Imports Microsoft.Office
Imports System.Drawing
Imports Microsoft.Office.Interop.Excel
Imports DataTable = System.Data.DataTable
Imports Point = System.Drawing.Point
Imports System.Windows.Forms.VisualStyles.VisualStyleElement

Public Class Form40

    Dim WithEvents excelApp As Excel.Application
    Dim workBook As Excel.Workbook
    Public Shared workSheet As Excel.Worksheet

    Dim src_rng As Excel.Range
    Public des_rng As Excel.Range
    Dim selectedRange As Excel.Range

    Public validationRange As Excel.Range
    Private allItems As New List(Of String)()

    Private processingEvent As Boolean = False
    'Public focuschange As Boolean

    Private Declare Function SetWindowPos Lib "user32" (ByVal hWnd As IntPtr, ByVal hWndInsertAfter As IntPtr, ByVal X As Integer, ByVal Y As Integer, ByVal cx As Integer, ByVal cy As Integer, ByVal uFlags As UInteger) As Boolean
    Private Const SWP_NOMOVE As UInteger = &H2
    Private Const SWP_NOSIZE As UInteger = &H1
    Private Const SWP_NOACTIVATE As UInteger = &H10
    Private Const HWND_TOPMOST As Integer = -1

    'Public Target As Excel.Range


    Private Structure RECT
        Public Left As Integer
        Public Top As Integer
        Public Right As Integer
        Public Bottom As Integer
    End Structure

    <DllImport("user32.dll", SetLastError:=True)>
    Private Shared Function GetWindowRect(ByVal hWnd As IntPtr, ByRef lpRect As RECT) As Boolean
    End Function



    Private Sub Form40_Load(sender As Object, e As EventArgs) Handles Me.Load
        excelApp = Globals.ThisAddIn.Application
        Dim workbook As Excel.Workbook = excelApp.ActiveWorkbook
        Dim worksheet As Excel.Worksheet = workbook.ActiveSheet

        Dim cell As Excel.Range = worksheet.Range(TargetVar3) ' In TargetVar, there is address about Target cell
        Dim validationFormula As String = cell.Validation.Formula1
        Dim items As New List(Of String)()
        'MsgBox(validationFormula)
        ' Dim items As New List(Of String)()

        If Not validationFormula.Contains(",") AndAlso Not validationFormula.Contains("!") Then
            ' It's a range on the same sheet
            Dim range As Excel.Range = worksheet.Range(validationFormula)

            For Each cellInRange As Excel.Range In range.Cells
                If Not String.IsNullOrEmpty(cellInRange.Value?.ToString()) Then
                    items.Add(cellInRange.Value.ToString())
                    allItems.Add(cellInRange.Value.ToString()) ' Add to the master list as well
                End If
            Next
        ElseIf validationFormula.Contains(",") Then
            ' Direct values separated by commas
            items.AddRange(validationFormula.Split(New Char() {","c}))
            allItems.AddRange(validationFormula.Split(New Char() {","c}))
        End If

        ListBox1.Items.Clear()
        ListBox1.Items.AddRange(items.ToArray())

        Me.BringToFront()


    End Sub

    Private Sub ListBox1_SelectedIndexChanged(sender As Object, e As EventArgs) Handles ListBox1.SelectedIndexChanged
        excelApp = Globals.ThisAddIn.Application
        Dim workbook As Excel.Workbook = excelApp.ActiveWorkbook
        Dim worksheet As Excel.Worksheet = workbook.ActiveSheet

        If ListBox1.SelectedItem IsNot Nothing Then
            ' Set the value in B1 cell to the selected item
            Dim selectedItem As String = ListBox1.SelectedItem.ToString()
            worksheet.Range(TargetVar3).Value = selectedItem
        End If
    End Sub

    Private Sub PictureBox3_Click(sender As Object, e As EventArgs) Handles PictureBox3.Click
        Me.Close()
    End Sub

    Private Sub txtSearch_TextChanged(sender As Object, e As EventArgs) Handles txtSearch.TextChanged
        Dim searchTerm As String = txtSearch.Text.ToLower()

        ' Filter items based on the search term
        Dim filteredItems = allItems.Where(Function(item) item.ToLower().Contains(searchTerm)).ToList()

        ' Update the ListBox
        ListBox1.Items.Clear()
        ListBox1.Items.AddRange(filteredItems.ToArray())


    End Sub

    'For position of the form
    Private Sub Form40_Activated(sender As Object, e As EventArgs) Handles Me.Activated
        excelApp = Globals.ThisAddIn.Application
        Dim workbook As Excel.Workbook = excelApp.ActiveWorkbook
        Dim worksheet As Excel.Worksheet = workbook.ActiveSheet


        Dim excelWindow = excelApp.ActiveWindow

        Dim cell As Range = worksheet.Range(TargetVar3).Offset(1, 1)
        Dim zoomFactor = excelWindow.Zoom / 100
        'var ws = cell.Worksheet;

        Dim ap = excelWindow.ActivePane ' might be split panes
        Dim origScrollCol = ap.ScrollColumn
        Dim origScrollRow = ap.ScrollRow
        excelApp.ScreenUpdating = False
        ' when FreezePanes == true, ap.ScrollColumn/Row will only reset
        ' as much as the location of the frozen splitter
        ap.ScrollColumn = 1
        ap.ScrollRow = 1

        ' PointsToScreenPixels returns different values if the scroll Is Not currently 1
        ' Temporarily set the scroll back to 1 so that PointsToScreenPixels returns a
        ' value we know how to handle.
        ' (x,y) are screen coordinates for the top left corner of the top left cell
        Dim x As Integer = ap.PointsToScreenPixelsX(0) ' e.g. window.x + row header width
        Dim y As Integer = ap.PointsToScreenPixelsY(0) ' e.g. window.y + ribbon height + column headers height

        Dim dpiX As Single = 0
        Dim dpiY As Single = 0

        Using g As Graphics = Graphics.FromHwnd(IntPtr.Zero)
            dpiX = g.DpiX
            dpiY = g.DpiY
        End Using

        Dim deltaRow As Integer = 0
        Dim deltaCol As Integer = 0
        Dim fromCol As Integer = origScrollCol
        Dim fromRow As Integer = origScrollRow
        If (excelWindow.FreezePanes) Then
            fromCol = 1
            fromRow = 1
            deltaCol = origScrollCol - ap.ScrollColumn '// Note: ap.ScrollColumn/ Row <> 1
            deltaRow = origScrollRow - ap.ScrollRow  '// see comment: when FreezePanes == true ...
        End If

        '// Note Each column width / row height has to be calculated individually.
        '// Before, tried to use this approach:
        '// var r2 = (Microsoft.Office.Interop.Excel.Range) cell.Worksheet.Cells[origScrollRow, origScrollCol];
        '// double dw = cell.Left - r2.Left;
        '// double dh = cell.Top - r2.Top;
        '// However, that only works when the zoom factor Is a whole number.
        '// A fractional zoom (e.g. 1.27) causes each individual row Or column to round to the closest whole number,
        '// which means having to loop through.

        Dim col As Excel.Range
        Dim ww As Double
        Dim newW As Double
        Dim i As Integer
        For i = fromCol To cell.Column - 1
            ' skip the columns between the frozen split and the first visible column
            If i >= ap.ScrollColumn AndAlso i < ap.ScrollColumn + deltaCol Then
                Continue For
            End If

            col = CType(worksheet.Cells(cell.Row, i), Microsoft.Office.Interop.Excel.Range)
            ww = col.Width * dpiX / 72
            newW = zoomFactor * ww
            x += CInt(Math.Round(newW))
        Next


        Dim row As Range
        Dim hh As Double
        Dim newH As Double

        For i = fromRow To cell.Row - 1
            ' skip the rows between the frozen split and the first visible row
            If i >= ap.ScrollRow AndAlso i < ap.ScrollRow + deltaRow Then
                Continue For
            End If

            row = CType(worksheet.Cells(i, cell.Column), Excel.Range)
            hh = row.Height * dpiY / 72
            newH = zoomFactor * hh
            y += CInt(Math.Round(newH))
        Next

        ap.ScrollColumn = origScrollCol
        ap.ScrollRow = origScrollRow
        excelApp.ScreenUpdating = True

        'myFormInstance = New Form36()
        'yourFormInstance.Show()

        'Form f = New Form();
        'Me.Show()
        'Me.StartPosition = FormStartPosition.Manual
        Me.Location = New Point(x, y) + New Point(2, 2)
        ' MsgBox(Me.Location.ToString)

    End Sub


    Private Sub form_enter(sender As Object, e As KeyEventArgs) Handles Me.KeyDown

        Try
            If e.KeyCode = Keys.Enter Then

                Me.Close()

            End If

        Catch ex As Exception

        End Try

    End Sub


    Private Sub listbox_enter(sender As Object, e As KeyEventArgs) Handles ListBox1.KeyDown

        Try
            If e.KeyCode = Keys.Enter Then

                Me.Close()

            End If

        Catch ex As Exception

        End Try

    End Sub

    Private Sub Form40_Shown(sender As Object, e As EventArgs) Handles Me.Shown
        'Me.BringToFront()
        Me.Focus()
    End Sub

    Private Sub Panel1_Paint(sender As Object, e As PaintEventArgs) Handles Panel1.Paint

    End Sub

    Private Sub PictureBox4_Click(sender As Object, e As EventArgs)

    End Sub
End Class