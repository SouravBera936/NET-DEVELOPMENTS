Public Class SETTING
    Private Sub SETTING_Load(sender As Object, e As EventArgs) Handles MyBase.Load
        TextBox1.Enabled = False
        Button5.Enabled = False
        Button6.Enabled = False
        If MAIN.TextBox2.Text = "ADMIN" Then
            Button1.Enabled = True
            Button2.Enabled = True
            Button3.Enabled = False
            Button4.Enabled = True
        Else
            Button1.Enabled = False
            Button2.Enabled = False
            Button3.Enabled = True
            Button4.Enabled = False
        End If

    End Sub

    Private Sub Button1_Click(sender As Object, e As EventArgs) Handles Button1.Click
        Me.Hide()
        Call CUSER.Show()
    End Sub

    Private Sub Button7_Click(sender As Object, e As EventArgs) Handles Button7.Click
        Me.Close()
        Call MAIN.Show()
    End Sub

    Private Sub Button4_Click(sender As Object, e As EventArgs) Handles Button4.Click
        Me.Hide()
        Call CONFIG.Show()
    End Sub

    Private Sub Button2_Click(sender As Object, e As EventArgs) Handles Button2.Click
        Me.Hide()
        Call REMUSER.Show()
    End Sub

    Private Sub Button3_Click(sender As Object, e As EventArgs) Handles Button3.Click
        Me.Hide()
        Call PSCH.Show()
    End Sub
End Class