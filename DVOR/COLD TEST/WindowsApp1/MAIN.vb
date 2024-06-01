Public Class MAIN
    Private Sub MAIN_Load(sender As Object, e As EventArgs) Handles MyBase.Load
        TextBox1.Enabled = False
        TextBox2.Enabled = False
        Dim myProcesses() As Process
        Dim myProcess As Process
        myProcesses = Process.GetProcessesByName("EXCEL")
        If myProcesses.Length > 0 Then
            For Each myProcess In myProcesses
                If myProcess IsNot Nothing Then
                    myProcess.Kill()
                End If
            Next
        End If
    End Sub

    Private Sub Button1_Click(sender As Object, e As EventArgs) Handles Button1.Click
        TESTFORM.TextBox3.Text = Me.TextBox2.Text
        TESTFORM.TextBox4.Text = Me.TextBox1.Text
        Me.Hide()
        TESTFORM.ShowDialog()
    End Sub

    Private Sub Button3_Click(sender As Object, e As EventArgs) Handles Button3.Click
        Dim iret As Object = MsgBox("DO YOU WANT TO EXIT?", vbQuestion + vbYesNo, "EXIT CONFIRMATION")
        If iret = vbYes Then
            Application.Exit()
        Else
            If iret = vbNo Then
                'donoting
            End If
        End If
    End Sub

    Private Sub Button2_Click(sender As Object, e As EventArgs) Handles Button2.Click
        Me.Hide()
        SETTING.TextBox1.Text = Me.TextBox1.Text
        SETTING.ShowDialog()
    End Sub
End Class