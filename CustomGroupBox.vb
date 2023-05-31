
Imports System.Drawing
Imports System.Windows.Forms

Public Class CustomGroupBox
    Inherits GroupBox

    Private ReadOnly flags As TextFormatFlags =
        TextFormatFlags.Top Or
        TextFormatFlags.Left Or
        TextFormatFlags.LeftAndRightPadding Or
        TextFormatFlags.EndEllipsis
    Private _Bordercolor As Color = SystemColors.Window

    Public Property BorderColor As Color
        Get
            Return _Bordercolor

        End Get
        Set(value As Color)
            _Bordercolor = value
            Me.Invalidate()
        End Set
    End Property

    Protected Overrides Sub Onpaint(e As PaintEventArgs)
        Dim mTxt = TextRenderer.MeasureText(e.Graphics, Text, Font, ClientSize).Height \ 2 + 2
        Dim r As New Rectangle(0, mTxt, ClientSize.Width, ClientSize.Height - mTxt)
        ControlPaint.DrawBorder(e.Graphics, r, BorderColor, ButtonBorderStyle.Solid)

        Dim textrect = Rectangle.Inflate(ClientRectangle, -4, 0)
        TextRenderer.DrawText(e.Graphics, Me.Text, Font, textrect, ForeColor, BackColor, flags)

    End Sub


End Class
