Imports System.DirectoryServices.AccountManagement
Public Class TrialGlobalLogin
    Private Sub Button1_Click(sender As Object, e As EventArgs) Handles Button1.Click
        Dim validCreds As Boolean = False
        Using context As New PrincipalContext(ContextType.Domain)
            validCreds = context.ValidateCredentials(TextBox1.Text, TextBox2.Text)
        End Using
        Dim val As String = validCreds.ToString()
        MsgBox(val)
    End Sub

    Private Sub TextBox2_TextChanged(sender As Object, e As EventArgs) Handles TextBox2.TextChanged, TextBox4.TextChanged
        TextBox2.PasswordChar = "*"
    End Sub

    Private Sub Button2_Click(sender As Object, e As EventArgs) Handles Button2.Click
        Dim trimmedName As String = TextBox3.Text.TrimEnd()
        If Not String.IsNullOrWhiteSpace(trimmedName) Then
            trimmedName &= ""
            TextBox4.Text = String.Join("", trimmedName.Split(" "c).Where(Function(word) Not String.IsNullOrWhiteSpace(word)).Select(Function(word) word.Substring(0, 1))).ToUpper()
        Else
            MsgBox("invalid")
        End If
    End Sub
End Class