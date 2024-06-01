Public Class PASSWORD_RST
    Dim productversion = Application.ProductVersion.ToString
    Dim file1 As System.IO.StreamWriter
    Dim file2 As System.IO.StreamWriter
    Private Sub PASSWORD_RST_Load(sender As Object, e As EventArgs) Handles MyBase.Load
        TextBox1.Enabled = False
        TextBox2.Clear()
        TextBox3.Clear()
        Button1.Enabled = False
        Label3.Text = productversion
    End Sub

    Private Sub TextBox3_TextChanged(sender As Object, e As EventArgs) Handles TextBox3.TextChanged
        If TextBox3.Text = "" Then
            Button1.Enabled = False
        Else
            Button1.Enabled = True
        End If
    End Sub

    Private Sub TextBox2_TextChanged(sender As Object, e As EventArgs) Handles TextBox2.TextChanged
        If TextBox2.Text = "" Then
            TextBox3.Enabled = False
        Else
            TextBox3.Enabled = True
        End If
    End Sub

    Private Sub Button2_Click(sender As Object, e As EventArgs) Handles Button2.Click
        Me.Close()
    End Sub

    Private Sub Button1_Click(sender As Object, e As EventArgs) Handles Button1.Click
        Dim line1 As Label
        MAIN.ListBox1.Items.Add("PASSWORD RESET REQUESTED FOR : " & TextBox1.Text)
        MAIN.ListBox1.SelectedIndex = MAIN.ListBox1.Items.Count - 1
line1:
        If FUNNCTIONLIB.SecqurityReset(TextBox1.Text, TextBox2.Text, TextBox3.Text, CONFIG.TextBox1.Text, CONFIG.TextBox2.Text, CONFIG.TextBox3.Text) = "INVALID USER" Then
            Dim iret As Object = MsgBox("INVALID CREDITIONALS! PASSWORD RECOVERY FAILED", vbRetryCancel + vbInformation, "SECQURITY RESET FAIL")
            If iret = vbRetry Then
                GoTo line1
            ElseIf iret = vbCancel Then
                TextBox2.Clear()
                TextBox3.Clear()
            End If
        Else
            Dim Password As String = FUNNCTIONLIB.SecqurityReset(TextBox1.Text, TextBox2.Text, TextBox3.Text, CONFIG.TextBox1.Text, CONFIG.TextBox2.Text, CONFIG.TextBox3.Text)
            Dim iret1 As Object = MsgBox("PASSWORD RECOVERSUCCESS! YOUR PASSWORD IS : " & Password, vbOKOnly + vbInformation, "SECQURITY RESETED")
            If iret1 = vbOK Then
                TextBox2.Clear()
                TextBox3.Clear()
                Me.Close()
            End If
        End If
    End Sub
End Class