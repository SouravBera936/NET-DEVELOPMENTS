Imports Ivi.Visa.Interop
Imports System.Diagnostics
Imports System.IO
Public Class T2
    Private Sub Button1_Click(sender As Object, e As EventArgs) Handles Button1.Click
        Dim port As String = CONFIG.TextBox4.Text.ToString
        Dim iomgr As New Ivi.Visa.Interop.ResourceManager
        Dim instrany As New Ivi.Visa.Interop.FormattedIO488
        instrany.IO = iomgr.Open(port)
        instrany.WriteString("*RST")
        instrany.WriteString("*CLS")
        instrany.WriteString("SYST:RWL")
        instrany.WriteString("VOLT 24.00")
        instrany.WriteString("CURR 0.500")
        instrany.WriteString("OUTP ON")
        TESTFORM.TextBox15.Text = "RUNNING"
        Dim MnMMATest As Process() = Process.GetProcessesByName("MnMMATest")
        If MnMMATest.Length = 0 Then
            Dim proc As New System.Diagnostics.Process()
            proc = Process.Start(CONFIG.TextBox5.Text.ToString)
            Me.Close()
            Call T4.Show()
        Else
            Me.Close()
            Call T3.Show()
        End If
    End Sub

    Private Sub T2_Load(sender As Object, e As EventArgs) Handles MyBase.Load
        TESTFORM.TextBox15.Text = "RUNNING"
    End Sub
End Class