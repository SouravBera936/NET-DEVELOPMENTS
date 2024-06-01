Public Class MAIN
    Private Sub MAIN_Load(sender As Object, e As EventArgs) Handles MyBase.Load
        TextBox2.Enabled = False
        TextBox1.Enabled = False
    End Sub

    Private Sub Button1_Click(sender As Object, e As EventArgs) Handles Button1.Click
        TESTFORM.TextBox3.Text = Me.TextBox2.Text
        Me.Hide()
        Call TESTFORM.Show()
    End Sub

    Private Sub Button2_Click(sender As Object, e As EventArgs) Handles Button2.Click
        Me.Hide()
        Call SETTING.Show()
        SETTING.TextBox1.Text = Me.TextBox2.Text
    End Sub

    Private Sub Label3_Click(sender As Object, e As EventArgs) Handles Label3.Click

    End Sub

    Private Sub TextBox2_TextChanged(sender As Object, e As EventArgs) Handles TextBox2.TextChanged

    End Sub

    Private Sub Button3_Click(sender As Object, e As EventArgs) Handles Button3.Click
        Dim iret As Object = MsgBox("DO YOU WANT TO EXIT?", vbQuestion + vbYesNo, "EXIT CONFIRMATION")
        If iret = vbYes Then
            Application.Exit()
        Else
            If iret = vbNo Then
                'donothing
            End If
        End If
    End Sub
End Class