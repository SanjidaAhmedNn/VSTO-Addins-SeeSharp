Imports System.Drawing
Imports System.Windows.Forms

Public Class CustomLabel
    Inherits Label

    Private _borderColor As Color = Color.Black
    Private _borderWidth As Double = 0.1

    Public Property BorderColor As Color
        Get
            Return _borderColor
        End Get
        Set(value As Color)
            _borderColor = value
            Invalidate() ' Redraw the control
        End Set
    End Property

    Public Property BorderWidth As Integer
        Get
            Return _borderWidth
        End Get
        Set(value As Integer)
            _borderWidth = value
            Invalidate() ' Redraw the control
        End Set
    End Property

    Protected Overrides Sub OnPaint(e As PaintEventArgs)
        MyBase.OnPaint(e)

        Dim borderRect As Rectangle = New Rectangle(
            New Point(_borderWidth, _borderWidth),
            New Size(ClientSize.Width - (2 * _borderWidth), ClientSize.Height - (2 * _borderWidth))
        )

        ' Draw border
        Using borderPen As New Pen(_borderColor, _borderWidth)
            e.Graphics.DrawRectangle(borderPen, borderRect)
        End Using
    End Sub
End Class

