Imports VSTO_Addins
Imports Excel = Microsoft.Office.Interop.Excel
Imports Xunit

Namespace flipTest
    Public Class integrationTest

        Public Function CompareRange(range1 As Excel.Range, range2 As Excel.Range) As Boolean

            If range1.GetLength(0) <> range2.GetLength(0) OrElse range1.GetLength(1) <> range2.GetLength(1) Then
                Return False
            End If

            For i As Integer = 0 To range1.GetLength(0) - 1
                For j As Integer = 0 To range1.GetLength(1) - 1
                    If Not Object.Equals(range1(i, j), range2(i, j)) Then
                        Return False
                    End If
                Next
            Next

            Return True
        End Function
        <Fact>
        Public Sub TestSub01(sender As Object, e As EventArgs)

            Dim excelApp As New Excel.Application
            Dim workbook1 As Excel.Workbook = excelApp.Workbooks("C:\Users\USER\Desktop\sqa\Flip_Ranges.xlsm")
            Dim worksheet1 As Excel.Worksheet = workbook1.Worksheets("Sheet1")


            Dim myform As New Form1()
            Dim sourceRange As Excel.Range = worksheet1.Range("A1:F14")
            Dim destinationRange As Excel.Range = worksheet1.Range("H18")
            Dim expectedd As Excel.Range = worksheet1.Range("H1:M14")
            Dim rng As Excel.Range
            Dim rng2 As Excel.Range

            ' Act
            myform.TextBox1.Text = sourceRange.Address
            myform.RadioButton3.Checked = True
            myform.RadioButton1.Checked = True
            myform.TextBox2.Text = destinationRange.Address
            rng = sourceRange
            rng2 = destinationRange
            Call myform.btn_OK_Click(sender, e)

            ' Assert
            Assert.True(CompareRange(expectedd, destinationRange))

        End Sub
    End Class
End Namespace


