Public Class RebbotInitiaizer
    Private countdown As Integer
    Private timer As New Timer()
    Private Sub RebbotInitiaizer_Load(sender As Object, e As EventArgs) Handles MyBase.Load
        PictureBox1.Visible = False
        PictureBox2.Visible = False
        If Me.BackColor = Color.LimeGreen Then
            BunifuLabel1.Text = "Success"
            BunifuLabel2.Text = "Rebooting In"
            PictureBox2.Visible = False
            PictureBox1.Visible = True
        ElseIf Me.BackColor = Color.Crimson Then
            BunifuLabel1.Text = "Failed"
            BunifuLabel2.Text = "Rebooting In"
            PictureBox1.Visible = False
            PictureBox2.Visible = True
        End If
        timer.Interval = 1000
        TextBox1.Visible = False
        AddHandler timer.Tick, AddressOf Timer_Tick
        countdown = TextBox1.Text
        timer.Start()
    End Sub
    Private Sub Timer_Tick(sender As Object, e As EventArgs)
        countdown -= 1
        If countdown >= 0 Then
            BunifuLabel3.Text = countdown.ToString()
        Else
            timer.Stop()
            Application.Restart()
        End If
    End Sub

    Private Sub BunifuButton1_Click(sender As Object, e As EventArgs) Handles BunifuButton1.Click
        timer.Stop()
        Me.Close()
    End Sub
End Class