Imports System
Imports System.Security.Cryptography.X509Certificates
Imports VSTO_Addins
Imports Xunit

Namespace flipTest
    Public Class integrationTest
        <Fact>
        Public Sub TestSub01()

            Dim form1 As New TestForm1()
            Dim sourceRange As String = "A1:B2"
            Dim destinationRange As String = "C1:D2"

            ' Act
            form1.TextBox1_TextChanged(Nothing, Nothing) ' Simulate TextBox1's text changing.
            form1.btn_OK_Click(RadioButton3.checked)
            form1.TextBox2_TextChanged(Nothing, Nothing) ' Simulate TextBox2's text changing.
            Dim result As Boolean = form1.ClickOKButton()

            ' Assert
            Assert.True(result)
        End Sub
    End Class

    Public Class TestForm1
        Inherits Form1

        Public Overrides Sub TextBox1_TextChanged(sender As Object, e As EventArgs)
            TextBox1.Text = "A1:B2" ' Simulate the user typing "A1:B2" into TextBox1.
            MyBase.TextBox1_TextChanged(sender, e)
        End Sub
        End Sub
    End Class
End Namespace

