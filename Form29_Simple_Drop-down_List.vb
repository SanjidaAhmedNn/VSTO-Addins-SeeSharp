Imports System.Threading
Imports System.Windows.Forms
Imports System.Windows.Forms.VisualStyles.VisualStyleElement
Imports Microsoft.Office.Interop
Imports System.Runtime.InteropServices


Public Class Form29_Simple_Drop_down_List

    Dim WithEvents excelApp As Excel.Application
    Dim workBook As Excel.Workbook
    Dim workSheet As Excel.Worksheet
    Dim workSheet2 As Excel.Worksheet
    Dim src_rng As Excel.Range
    Public des_rng As Excel.Range
    Dim selectedRange As Excel.Range


    Private Declare Function SetWindowPos Lib "user32" (ByVal hWnd As IntPtr, ByVal hWndInsertAfter As IntPtr, ByVal X As Integer, ByVal Y As Integer, ByVal cx As Integer, ByVal cy As Integer, ByVal uFlags As UInteger) As Boolean
    Private Const SWP_NOMOVE As UInteger = &H2
    Private Const SWP_NOSIZE As UInteger = &H1
    Private Const SWP_NOACTIVATE As UInteger = &H10
    Private Const HWND_TOPMOST As Integer = -1


    Dim opened As Integer
    Private Sub Info_Click(sender As Object, e As EventArgs) Handles Info.Click

    End Sub

    Private Sub ListBox1_SelectedIndexChanged(sender As Object, e As EventArgs) Handles ListBox1.SelectedIndexChanged
        ' Clear the list box
        ListBox2.Items.Clear()
        Dim selectedItem As String = ListBox1.SelectedItem.ToString()
        ' Split the string into an array of strings
        Dim items As String() = selectedItem.Split(","c)

        ListBox2.Items.AddRange(items)
        Label7.Visible = True
        Label7.Text = items.Count

    End Sub

    Private Sub ComboBox1_SelectedIndexChanged(sender As Object, e As EventArgs) Handles ComboBox1.SelectedIndexChanged

    End Sub

    Private Sub ComboBox1_TextUpdate(sender As Object, e As EventArgs) Handles ComboBox1.TextUpdate

    End Sub

    Private Sub ComboBox1_MouseClick(sender As Object, e As MouseEventArgs) Handles ComboBox1.MouseClick
        If ComboBox1.Text = "" Then
            'Do nothing
        Else
            ' Clear the list box
            ListBox2.Items.Clear()
            Dim selectedItem As String = ComboBox1.Text
            ' Split the string into an array of strings
            Dim items As String() = selectedItem.Split(","c)

            ListBox2.Items.AddRange(items)
            Label7.Visible = True
            Label7.Text = items.Count
        End If
    End Sub

    Private Sub ComboBox1_Enter(sender As Object, e As EventArgs) Handles ComboBox1.Enter
        If ComboBox1.Text = "" Then
            'Do nothing
        Else
            ' Clear the list box
            ListBox2.Items.Clear()
            Dim selectedItem As String = ComboBox1.Text
            ' Split the string into an array of strings
            Dim items As String() = selectedItem.Split(","c)

            ListBox2.Items.AddRange(items)
            Label7.Visible = True
            Label7.Text = items.Count
        End If
    End Sub

    Private Sub ComboBox1_DragLeave(sender As Object, e As EventArgs) Handles ComboBox1.DragLeave

    End Sub

    Private Sub ComboBox1_KeyPress(sender As Object, e As KeyPressEventArgs) Handles ComboBox1.KeyPress
        If ComboBox1.Text = "" Then
            'Do nothing
        Else
            ' Clear the list box
            ListBox2.Items.Clear()
            Dim selectedItem As String = ComboBox1.Text
            ' Split the string into an array of strings
            Dim items As String() = selectedItem.Split(","c)

            ListBox2.Items.AddRange(items)
            Label7.Visible = True
            Label7.Text = items.Count
        End If
    End Sub

    Private Sub Btn_OK_Click(sender As Object, e As EventArgs) Handles Btn_OK.Click
        excelApp = Globals.ThisAddIn.Application
        workBook = excelApp.ActiveWorkbook
        workSheet = workBook.ActiveSheet

        'Dim items As String() = {"Item 1", "Item 2", "Item 3"}
        Dim stringItems As New List(Of String)()

        ' Join the items into a comma-separated string
        'Dim formula As String = String.Join(",", ListBox2.Items)
        For Each item As Object In ListBox2.Items
            stringItems.Add(item.ToString())
        Next

        ' Join the string representations into a single string
        Dim items As String = String.Join(", ", stringItems)

        'Dim items As String = String.Join(",", ListBox2.Items.Cast(Of String)().ToArray())


        ' Define the cell that will contain the drop-down list (for example, cell A1)
        'Dim range As Excel.Range = des_rng
        ' Delete existing validation rules
        'MsgBox(des_rng.Address)
        ' des_rng.Address= TB_dest_range.text
        des_rng.Validation.Delete()

        ' Create a new validation rule
        Dim validation As Excel.Validation = des_rng.Validation

        ' Add a drop-down list validation rule
        validation.Delete()
        validation.Add(Excel.XlDVType.xlValidateList, Excel.XlDVAlertStyle.xlValidAlertStop, Excel.XlFormatConditionOperator.xlBetween, items, Type.Missing)
        validation.IgnoreBlank = True
        validation.InCellDropdown = True

        Me.Close()

    End Sub

    Private Sub Btn_Cancel_Click(sender As Object, e As EventArgs) Handles Btn_Cancel.Click
        Me.Close()
    End Sub

    Private Sub Selection_Source_Click(sender As Object, e As EventArgs) Handles Selection_Source.Click
        If selectedRange Is Nothing Then
        Else
            ' TB_src_range.Text = selectedRange.Address


            Me.Hide()

            excelApp = Globals.ThisAddIn.Application
            workBook = excelApp.ActiveWorkbook

            'Dim userInput As String = excelApp.InputBox("Select a range", "Select range", "=$A$1")


            Dim userInput As Excel.Range = excelApp.InputBox("Select a range", "Select a range", "=$A$1", Type:=8)
            src_rng = userInput

            Dim sheetName As String
            sheetName = Split(src_rng.Address(True, True, Excel.XlReferenceStyle.xlA1, True), "]")(1)
            sheetName = Split(sheetName, "!")(0)

            If Mid(sheetName, Len(sheetName), 1) = "'" Then
                sheetName = Mid(sheetName, 1, Len(sheetName) - 1)
            End If

            workSheet = workBook.Worksheets(sheetName)
            workSheet.Activate()

            src_rng.Select()
            'MsgBox(src_rng.Address)

            TB_src_range.Text = src_rng.Address

            Me.Show()
            TB_src_range.Focus()

            ' Define the range of cells to read (for example, cells A1 to A10)
            Dim range As Excel.Range = src_rng

            ' Clear the ListBox
            ListBox2.Items.Clear()

            ' Iterate over each cell in the range
            For Each cell As Excel.Range In range
                ' Add the cell's value to the ListBox
                ListBox2.Items.Add(cell.Value)
            Next

            Label7.Visible = True
            Label7.Text = ListBox2.Items.Count
            TB_src_range.Focus()
            Me.Activate()

        End If

    End Sub


    Private Sub Selection_Click(sender As Object, e As EventArgs) Handles Selection_destination.Click
        Try
            If selectedRange Is Nothing Then
            Else

                TB_dest_range.Text = selectedRange.Address


                'FocusedTextBox = 1
                Me.Hide()

                excelApp = Globals.ThisAddIn.Application
                workBook = excelApp.ActiveWorkbook

                'Dim userInput As String = excelApp.InputBox("Select a range", "Select range", "=$A$1")


                Dim userInput As Excel.Range = excelApp.InputBox("Select a range", "Select a range", "=$A$1", Type:=8)
                des_rng = userInput

                Dim sheetName As String
                sheetName = Split(des_rng.Address(True, True, Excel.XlReferenceStyle.xlA1, True), "]")(1)
                sheetName = Split(sheetName, "!")(0)

                If Mid(sheetName, Len(sheetName), 1) = "'" Then
                    sheetName = Mid(sheetName, 1, Len(sheetName) - 1)
                End If

                workSheet = workBook.Worksheets(sheetName)
                workSheet.Activate()

                des_rng.Select()

                TB_dest_range.Text = des_rng.Address

                Me.Show()
                TB_dest_range.Focus()
            End If

        Catch ex As Exception

            Me.Show()
            TB_dest_range.Focus()

        End Try
    End Sub



    Private Sub Form1_Load(sender As Object, e As EventArgs) Handles MyBase.Load

        Try

            excelApp = Globals.ThisAddIn.Application

            AddHandler excelApp.SheetSelectionChange, AddressOf excelApp_SheetSelectionChange

            opened = opened + 1

            If excelApp.Selection IsNot Nothing Then
                selectedRange = excelApp.Selection
                des_rng = selectedRange
                TB_dest_range.Text = selectedRange.Address
            End If

        Catch ex As Exception

        End Try

    End Sub


    Private Sub excelApp_SheetSelectionChange(ByVal Sh As Object, ByVal selectionRange1 As Excel.Range) Handles excelApp.SheetSelectionChange
        Try

            excelApp = Globals.ThisAddIn.Application

            If Me.ActiveControl Is TB_dest_range Then
                des_rng = selectionRange1
                ' This will run on the Excel thread, so you need to use Invoke to update the UI
                'Me.BeginInvoke(New System.Action(Sub() TB_dest_range.Text = selectionRange1.Address))
                Me.Activate()
                Me.BeginInvoke(New System.Action(Sub()
                                                     TB_dest_range.Text = des_rng.Address
                                                     SetWindowPos(Me.Handle, New IntPtr(HWND_TOPMOST), 0, 0, 0, 0, SWP_NOACTIVATE Or SWP_NOMOVE Or SWP_NOSIZE)
                                                 End Sub))

            ElseIf Me.ActiveControl Is TB_src_range Then
                src_rng = selectionRange1
                'workSheet = workBook.ActiveSheet
                'TB_src_range.Text = src_rng.Address
                'TB_src_range.Focus()
                'Me.Activate()
                'ActiveForm.Select()


                Me.BeginInvoke(New System.Action(Sub()
                                                    TB_src_range.Text = src_rng.Address
                                                     SetWindowPos(Me.Handle, New IntPtr(HWND_TOPMOST), 0, 0, 0, 0, SWP_NOACTIVATE Or SWP_NOMOVE Or SWP_NOSIZE)
                                                 End Sub))
            End If



        Catch ex As Exception

        End Try

    End Sub

    Private Sub TB_src_range_TextChanged(sender As Object, e As EventArgs) Handles TB_src_range.TextChanged
        If TB_src_range.Text IsNot Nothing Then

            ' Define the range of cells to read (for example, cells A1 to A10)
            Dim range As Excel.Range = src_rng

            ' Clear the ListBox
            ListBox2.Items.Clear()

            ' Iterate over each cell in the range
            For Each cell As Excel.Range In range
                ' Add the cell's value to the ListBox
                ListBox2.Items.Add(cell.Value)
            Next

            Label7.Visible = True
            Label7.Text = ListBox2.Items.Count
            TB_src_range.Focus()
            Me.Activate()
        End If
    End Sub

    Private Sub TB_dest_range_TextChanged(sender As Object, e As EventArgs) Handles TB_dest_range.TextChanged

    End Sub
End Class

