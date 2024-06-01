Public Class T3
    Private Sub Button1_Click(sender As Object, e As EventArgs) Handles Button1.Click
        Dim name As String = TESTFORM.TextBox4.Text
        Dim file As System.IO.StreamWriter
        Dim path As String = (CONFIG.TextBox6.Text.ToString) & name & "_" & "PASS" + ".txt"
        If System.IO.File.Exists(path) Then
            file = My.Computer.FileSystem.OpenTextFileWriter(path, True)
            file.WriteLine("")
            file.WriteLine("1 -------USERNAME -------" & TESTFORM.TextBox3.Text & " -------")
            file.WriteLine("2 -------SERIAL NUMBER -------" & TESTFORM.TextBox4.Text & " -------")
            file.WriteLine("3 -------RESULT -------PASS -------")
            file.WriteLine("*****END OF DATA****" & DateAndTime.Now.ToString)
            file.Close()
        Else
            System.IO.File.CreateText(path).Dispose()
            file = My.Computer.FileSystem.OpenTextFileWriter(path, True)
            file.WriteLine("")
            file.WriteLine("1 -------USERNAME -------" & TESTFORM.TextBox3.Text & " -------")
            file.WriteLine("2 -------SERIAL NUMBER -------" & TESTFORM.TextBox4.Text & " -------")
            file.WriteLine("3 -------RESULT -------PASS -------")
            file.WriteLine("*****END OF DATA****" & DateAndTime.Now.ToString)
            file.Close()
        End If
        Dim port As String = CONFIG.TextBox4.Text.ToString
        Dim iomgr As New Ivi.Visa.Interop.ResourceManager
        Dim instrany As New Ivi.Visa.Interop.FormattedIO488
        instrany.IO = iomgr.Open(port)
        instrany.WriteString("*RST")
        instrany.WriteString("*CLS")
        instrany.WriteString("OUTP OFF")
        TESTFORM.TextBox17.Text = "TRUE"
        TESTFORM.TextBox17.Enabled = False
        TESTFORM.TextBox15.Text = "PASS"
        TESTFORM.Button1.Text = "PASS"
        TESTFORM.TextBox4.Enabled = True
        TESTFORM.TextBox4.Text = ""
        TESTFORM.TextBox12.Enabled = True
        TESTFORM.TextBox17.Enabled = True
        Me.Close()
    End Sub

    Private Sub Button2_Click(sender As Object, e As EventArgs) Handles Button2.Click
        Dim name As String = TESTFORM.TextBox4.Text
        Dim file As System.IO.StreamWriter
        Dim path As String = (CONFIG.TextBox6.Text.ToString) & name & "_" & "FAIL" + ".txt"
        If System.IO.File.Exists(path) Then
            file = My.Computer.FileSystem.OpenTextFileWriter(path, True)
            file.WriteLine("")
            file.WriteLine("1 -------USERNAME -------" & TESTFORM.TextBox3.Text & " -------")
            file.WriteLine("2 -------SERIAL NUMBER -------" & TESTFORM.TextBox4.Text & " -------")
            file.WriteLine("3 -------RESULT -------FAIL -------")
            file.WriteLine("*****END OF DATA****" & DateAndTime.Now.ToString)
            file.Close()
        Else
            System.IO.File.CreateText(path).Dispose()
            file = My.Computer.FileSystem.OpenTextFileWriter(path, True)
            file.WriteLine("")
            file.WriteLine("1 -------USERNAME -------" & TESTFORM.TextBox3.Text & " -------")
            file.WriteLine("2 -------SERIAL NUMBER -------" & TESTFORM.TextBox4.Text & " -------")
            file.WriteLine("3 -------RESULT -------FAIL -------")
            file.WriteLine("*****END OF DATA****" & DateAndTime.Now.ToString)
            file.Close()
        End If
        Dim port As String = CONFIG.TextBox4.Text.ToString
        Dim iomgr As New Ivi.Visa.Interop.ResourceManager
        Dim instrany As New Ivi.Visa.Interop.FormattedIO488
        instrany.IO = iomgr.Open(port)
        instrany.WriteString("*RST")
        instrany.WriteString("*CLS")
        instrany.WriteString("OUTP OFF")
        TESTFORM.TextBox17.Text = "FALSE"
        TESTFORM.TextBox15.Text = "FAIL"
        TESTFORM.Button1.Text = "FAIL"
        TESTFORM.TextBox4.Enabled = True
        TESTFORM.TextBox4.Text = ""
        TESTFORM.TextBox12.Enabled = True
        TESTFORM.TextBox17.Enabled = True
        Me.Close()
    End Sub
End Class