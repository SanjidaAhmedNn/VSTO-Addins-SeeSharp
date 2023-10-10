Imports System.Windows.Forms
Imports System.Drawing

Public Class CustomButton
    Inherits Button

    Private _borderColor As Color = Color.Black
    Private _borderWidth As Integer = 1

    Public Property BorderColor As Color
        Get
            Return _borderColor
        End Get
        Set(ByVal value As Color)
            _borderColor = value
            Me.Invalidate() ' Forces control to be redrawn
        End Set
    End Property

    Public Property BorderWidth As Integer
        Get
            Return _borderWidth
        End Get
        Set(ByVal value As Integer)
            _borderWidth = value
            Me.Invalidate() ' Forces control to be redrawn
        End Set
    End Property

    Protected Overrides Sub OnPaint(ByVal e As PaintEventArgs)
        MyBase.OnPaint(e)

        ' Create border using BorderColor and BorderWidth properties
        Dim borderPen As New Pen(_borderColor, _borderWidth)
        Dim borderRectangle As New Rectangle(0, 0, Me.ClientSize.Width - 1, Me.ClientSize.Height - 1)

        ' Draw border
        e.Graphics.DrawRectangle(borderPen, borderRectangle)
    End Sub

    Private Sub InitializeComponent()
        Me.SuspendLayout()
        Me.ResumeLayout(False)

    End Sub
End Class
