Imports System.Drawing
Imports System.Drawing.Drawing2D
Imports System.Windows.Forms

Public Class CustomBorderButton
    Inherits Button

    Public Property BorderColor As Color = Color.Red
    Public Property BorderThickness As Integer = 1
    Public Property HoverColor As Color = Color.LightBlue
    Public Property PressedColor As Color = Color.DarkBlue
    Public Property HoverTextColor As Color = Color.Black
    Public Property PressedTextColor As Color = Color.White

    Private IsMouseOver As Boolean = False
    Private IsMouseDown As Boolean = False

    Protected Overrides Sub OnMouseEnter(e As EventArgs)
        IsMouseOver = True
        Invalidate()  ' Redraw the button
    End Sub

    Protected Overrides Sub OnMouseLeave(e As EventArgs)
        IsMouseOver = False
        IsMouseDown = False
        Invalidate()  ' Redraw the button
    End Sub

    Protected Overrides Sub OnMouseDown(mevent As MouseEventArgs)
        IsMouseDown = True
        Invalidate()  ' Redraw the button
    End Sub

    Protected Overrides Sub OnMouseUp(mevent As MouseEventArgs)
        IsMouseDown = False
        Invalidate()  ' Redraw the button
    End Sub

    Protected Overrides Sub OnPaint(ByVal e As PaintEventArgs)
        ' Default color
        Dim currentColor As Color = Me.BackColor
        Dim currentTextColor As Color = Me.ForeColor

        If IsMouseOver Then
            ' Change to hover color if mouse is over the button
            currentColor = HoverColor
            currentTextColor = HoverTextColor
        End If

        If IsMouseDown Then
            ' Change to pressed color if mouse is down
            currentColor = PressedColor
            currentTextColor = PressedTextColor
        End If

        ' Draw the button with the current color
        e.Graphics.Clear(currentColor)

        Dim buttonPath As GraphicsPath = New GraphicsPath()
        Dim myRectangle As Rectangle = ClientRectangle
        myRectangle.Inflate(0, -1)
        buttonPath.AddRectangle(myRectangle)
        Region = New Region(buttonPath)

        Dim borderPen As Pen = New Pen(BorderColor, BorderThickness)
        Dim borderRectangle As Rectangle = DisplayRectangle
        borderRectangle.Inflate(-1 * BorderThickness, -1 * BorderThickness)

        e.Graphics.DrawRectangle(borderPen, borderRectangle)

        ' Draw the text with the current text color
        TextRenderer.DrawText(e.Graphics, Me.Text, Me.Font, borderRectangle, currentTextColor)

        ' MyBase.OnPaint(e)  ' Comment this line out
    End Sub
End Class
