Imports Microsoft.Office.Tools.Ribbon

Public Class Ribbon1

    Private Sub Ribbon1_Load(ByVal sender As System.Object, ByVal e As RibbonUIEventArgs) Handles MyBase.Load

    End Sub

    Private Sub Button1_Click(sender As Object, e As RibbonControlEventArgs) Handles Button1.Click
        Dim form As New Form1
        form.Show()
    End Sub

    Private Sub Button2_Click(sender As Object, e As RibbonControlEventArgs) Handles Button2.Click
        Dim form As New Form2
        form.Show()
    End Sub

    Private Sub Button3_Click(sender As Object, e As RibbonControlEventArgs) Handles Button3.Click
        Dim form As New Form3
        form.Show()
    End Sub

    Private Sub Button5_Click(sender As Object, e As RibbonControlEventArgs) Handles Button5.Click
        Dim form As New Form8
        form.Show()
    End Sub

    Private Sub Button6_Click(sender As Object, e As RibbonControlEventArgs) Handles Button6.Click
        Dim form As New Form10

        form.Show()
    End Sub

    Private Sub Button7_Click(sender As Object, e As RibbonControlEventArgs) Handles Button7.Click
        Dim form As New Form7

        form.Show()
    End Sub

    Private Sub Button8_Click(sender As Object, e As RibbonControlEventArgs) Handles Button8.Click
        Dim form As New Form11SwapRanges

        form.Show()
    End Sub

    Private Sub Button4_Click(sender As Object, e As RibbonControlEventArgs) Handles Button4.Click
        Dim form As New Form12HideRanges

        form.Show()
    End Sub

    Private Sub Button11_Click(sender As Object, e As RibbonControlEventArgs) Handles Button11.Click
        Dim form As New Form13HideAllExceptSelectedRange

        form.Show()
    End Sub

    Private Sub Button12_Click(sender As Object, e As RibbonControlEventArgs) Handles Button12.Click
        Dim form As New Form14SpecifyScrollArea
        form.Show()
    End Sub

    Private Sub Button13_Click(sender As Object, e As RibbonControlEventArgs) Handles Button13.Click

        Dim form As New Form15CompareCells
        form.Show()

    End Sub

    Private Sub Button14_Click(sender As Object, e As RibbonControlEventArgs) Handles Button14.Click

        Dim form As New Form16PasteintoVisibleRange
        form.Show()

    End Sub

    Private Sub Button15_Click(sender As Object, e As RibbonControlEventArgs) Handles Button15.Click

        Dim form As New Form17DivideNames
        form.Show()

    End Sub

    Private Sub Button16_Click_1(sender As Object, e As RibbonControlEventArgs) Handles Button16.Click
        Dim form As New Form18_CombineRanges
        form.Show()
    End Sub
End Class
