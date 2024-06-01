Public Class SETTING
    Private Sub SETTING_Load(sender As Object, e As EventArgs) Handles MyBase.Load
        Me.TextBox1.Enabled = False
        If TextBox1.Text = "ADMINISTRATOR" Then
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

    Private Sub Button3_Click(sender As Object, e As EventArgs) Handles Button3.Click
        PSCH.TextBox1.Text = Me.TextBox1.Text
        Me.Hide()
        PSCH.ShowDialog()
    End Sub

    Private Sub Button1_Click(sender As Object, e As EventArgs) Handles Button1.Click
        Me.Hide()
        CUSER.ShowDialog()
    End Sub

    Private Sub Button4_Click(sender As Object, e As EventArgs) Handles Button4.Click
        Me.Hide()
        CONFIG.ShowDialog()
    End Sub

    Private Sub Button5_Click(sender As Object, e As EventArgs) Handles Button5.Click
        Me.Close()
        Call MAIN.Show()
    End Sub

    Private Sub Button2_Click(sender As Object, e As EventArgs) Handles Button2.Click
        Me.Hide()
        REMUSER.ShowDialog()
    End Sub
End Class