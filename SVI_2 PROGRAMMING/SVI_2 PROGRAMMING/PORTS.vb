Imports Ivi.Visa
Public Class PORTS
    Dim result As String
    Dim tcpip As Ivi.Visa.Interop.ITcpipInstr
    Dim visa_string As String
    Private Sub PORTS_Load(sender As Object, e As EventArgs) Handles MyBase.Load
        ListBox1.Items.Clear()
        For Each sp As String In My.Computer.Ports.SerialPortNames
            ListBox1.Items.Add(sp)
        Next
    End Sub

    Private Sub Button1_Click(sender As Object, e As EventArgs) Handles Button1.Click
        If TextBox1.Text = "" Then
            Dim iret As Object = MsgBox("INSTR ADDRESS NOT FOUND")
        Else
            CONFIG.TextBox4.Text = Me.TextBox1.Text
            Me.Close()
        End If
    End Sub
End Class