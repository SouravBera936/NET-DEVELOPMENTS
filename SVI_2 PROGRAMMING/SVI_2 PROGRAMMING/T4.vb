Public Class T4
    Private Sub textBox2_TextChanged(sender As Object, e As EventArgs) Handles textBox2.TextChanged

    End Sub

    Private Sub T4_Load(sender As Object, e As EventArgs) Handles MyBase.Load
        textBox2.Enabled = False
    End Sub

    Private Sub Button1_Click(sender As Object, e As EventArgs) Handles Button1.Click
        Me.Close()
        Call T3.Show()
    End Sub
End Class