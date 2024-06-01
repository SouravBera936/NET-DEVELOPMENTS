Imports System.Data.OleDb
Public Class SYSREG
    Dim pcname = System.Net.Dns.GetHostName
    Dim cn As OleDb.OleDbConnection
    Dim oleCommand As OleDbCommand
    Dim da As OleDbDataAdapter
    Dim productversion = Application.ProductVersion.ToString
    Dim file1 As System.IO.StreamWriter
    Dim file2 As System.IO.StreamWriter
    Private Sub SYSREG_Load(sender As Object, e As EventArgs) Handles MyBase.Load
        MAIN.ListBox1.Items.Add("Loading System Variables")
        Dim line1 As Label
line1:
        Dim path1 As String = CONFIG.TextBox4.Text
        file1 = My.Computer.FileSystem.OpenTextFileWriter(path1, True)
        Try
            cn = New OleDbConnection
            cn.ConnectionString = "Provider=" & CONFIG.TextBox1.Text & ";Data Source=" & CONFIG.TextBox2.Text & "; Persist Security Info=" & CONFIG.TextBox3.Text & ";"
            cn.Open()
            Dim SQL As String = "SELECT * FROM SYSTEM_1 WHERE (HOST_NAME = '" & pcname & "')"
            oleCommand = New OleDbCommand(SQL, cn)
            Dim reader As OleDbDataReader = oleCommand.ExecuteReader
            If reader.HasRows Then
                reader.Read()
                MAIN.ListBox1.Items.Add("HostName : " & pcname)
                MAIN.ListBox1.Items.Add("Previous Used Serial Number : " & reader("SERIAL_NUM"))
                TextBox2.Text = reader("SERIAL_NUM")
                MAIN.ListBox1.Items.Add("Previous Insalled Version : " & reader("SW_VERSION"))
                TextBox3.Text = reader("SW_VERSION")
                reader.Close()
                TextBox1.Text = pcname
                TextBox1.Enabled = False
                TextBox3.Enabled = False
                TextBox4.Clear()
                TextBox4.Text = DateTime.Now
                TextBox4.Enabled = False
                TextBox5.Clear()
                TextBox5.Enabled = False
                Button1.Enabled = False
                MAIN.ListBox1.Items.Add("Variables Loaded Succesfully")
                file1.Close()
            Else
                MAIN.ListBox1.Items.Add("Uploading PcName : " & pcname)
                TextBox1.Text = pcname
                TextBox1.Enabled = False
                TextBox2.Clear()
                TextBox3.Clear()
                TextBox3.Enabled = False
                TextBox4.Clear()
                TextBox4.Text = DateTime.Now
                TextBox4.Enabled = False
                TextBox5.Clear()
                TextBox5.Enabled = False
                Button1.Enabled = False
                MAIN.ListBox1.Items.Add("Variables Loaded For New System Registration")
                file1.Close()
            End If
        Catch ex As Exception
            Dim iret As Object = MsgBox("UNABLE TO CHECK SYSTEM LISCENCING : " & ex.Message, vbCritical + vbRetryCancel, "FUNNCTIONLIB.CheckLiscence error")
            If iret = vbRetry Then
                MAIN.ListBox1.Items.Add("CheckLiscence.Status : False" & "[" & ex.Message & "]")
                file1.Close()
                GoTo line1
            ElseIf iret = vbCancel Then
                file1.WriteLine("UNABLE TO CHECK SYSTEM LISCENCING : " & ex.Message & " " & "@" & DateTime.Now.ToString("yyyy/MM/dd/HH:mm:ss"))
                file1.Close()
                MAIN.ListBox1.Items.Add("CheckLiscence.Status : False" & "[" & ex.Message & "]")
            End If
        End Try

    End Sub

    Private Sub TextBox2_TextChanged(sender As Object, e As EventArgs) Handles TextBox2.TextChanged
        If TextBox2.Text = "SOURAV" Then
            TextBox3.Text = "TRIAL"
        Else
            If TextBox2.Text = "6RH4V-HNTWC-JQKG8-RFR3R-36498" Then
                TextBox3.Text = "PROFFESIONAL"
            Else
                If TextBox2.Text = "3FCND-JTWFM-24VQ8-QXTMB-TXT67" Then
                    TextBox3.Text = "PROFFESIONAL"
                Else
                    If TextBox2.Text = "MMRNH-BMB4F-87JR9-D72RY-MY2KV" Then
                        TextBox3.Text = "PROFFESIONAL"
                    Else
                        If TextBox2.Text = "T3NJK-3P683-2T7BJ-2X27F-8B2KV" Then
                            TextBox3.Text = "PROFFESIONAL"
                        Else
                            If TextBox2.Text = "MBFBV-W3DP2-2MVKN-PJCQD-KKTF7" Then
                                TextBox3.Text = "PROFFESIONAL"
                            Else
                                If TextBox2.Text = "GX9N8-4H2FH-D987T-BQ9GK-XKT67" Then
                                    TextBox3.Text = "PROFFESIONAL"
                                Else
                                    If TextBox2.Text = "KQWNF-XPMXP-HDK3M-GBV69-Y7RDH" Then
                                        TextBox3.Text = "PROFFESIONAL"
                                    Else
                                        If TextBox2.Text = "T3NJK-3P683-2T7BJ-2X27F-8B2KV" Then
                                            TextBox3.Text = "PROFFESIONAL"
                                        Else
                                            If TextBox2.Text = "334NH-RXG76-64THK-C7CKG-D3VPT" Then
                                                TextBox3.Text = "PROFFESIONAL"
                                            Else
                                                If TextBox2.Text = "KKPMN-469HY-H6V43-T8VX2-8W8XV" Then
                                                    TextBox3.Text = "PROFFESIONAL"
                                                Else
                                                    If TextBox2.Text = "ND8P2-BD2PB-DD8HM-2926R-CRYQH" Then
                                                        TextBox3.Text = "PROFFESIONAL"
                                                    Else
                                                        If TextBox2.Text = "RFQ3N-4Y4XR-JY9PV-883PR-BY2KV" Then
                                                            TextBox3.Text = "PROFFESIONAL"
                                                        Else
                                                            TextBox3.Text = "INVALID KEY"
                                                        End If
                                                    End If
                                                End If
                                            End If
                                        End If
                                    End If
                                End If
                            End If
                        End If
                    End If
                End If
            End If
        End If
    End Sub

    Private Sub TextBox3_TextChanged(sender As Object, e As EventArgs) Handles TextBox3.TextChanged
        If TextBox3.Text = "TRIAL" Then
            TextBox5.Text = DateTime.Now.AddDays(30)
            Button1.Enabled = True
        Else
            If TextBox3.Text = "PROFFESIONAL" Then
                TextBox5.Text = DateTime.Now.AddDays(2555)
                Button1.Enabled = True
            Else
                If TextBox3.Text = "INVALID KEY" Then
                    TextBox5.Clear()
                    Button1.Enabled = False
                End If
            End If
        End If
    End Sub

    Private Sub Button1_Click(sender As Object, e As EventArgs) Handles Button1.Click
        Dim line1 As Label
line1:
        Dim path1 As String = CONFIG.TextBox4.Text
        file1 = My.Computer.FileSystem.OpenTextFileWriter(path1, True)
        Dim path2 As String = CONFIG.TextBox36.Text
        file2 = My.Computer.FileSystem.OpenTextFileWriter(path2, True)
        Try
            cn = New OleDbConnection
            cn.ConnectionString = "Provider=" & CONFIG.TextBox1.Text & ";Data Source=" & CONFIG.TextBox2.Text & "; Persist Security Info=" & CONFIG.TextBox3.Text & ";"
            cn.Open()
            Dim SQL As String = "SELECT * FROM SYSTEM_1 WHERE (HOST_NAME = '" & pcname & "')"
            Dim SQL1 As String = "DELETE * FROM SYSTEM_1 WHERE (HOST_NAME = '" & pcname & "')"
            Dim SQL2 As String = "INSERT INTO SYSTEM_1 (HOST_NAME, SERIAL_NUM, INS_DATE, EXP_DATE, SW_VERSION) VALUES ('" & TextBox1.Text & "','" & TextBox2.Text & "','" & TextBox4.Text & "','" & TextBox5.Text & "','" & TextBox3.Text & "')"
            oleCommand = New OleDbCommand(SQL, cn)
            Dim reader As OleDbDataReader = oleCommand.ExecuteReader
            If reader.HasRows Then
                reader.Read()
                If reader("SW_VERSION") = "TRIAL" And TextBox3.Text = "TRIAL" Then
                    reader.Close()
                    Dim iret As Object = MsgBox("TRIAL PERIODE EXPIRE. PLEASE UPGRADE TO FULL VERSION", vbInformation + vbOKOnly, "LISCENCE MANAGER")
                    If iret = vbOK Then
                        MAIN.ListBox1.Items.Add("FUNNCLIB.InstallLiscence returns : False [Previous Insalled Version : TRIAL / Installing Version : TRIAL]")
                        file2.WriteLine("ATTEMPTED FOR LISCENCE INSTALATTION : False [Previous Installed Version : TRIAL / Installing Version : TRIAL]")
                        file2.WriteLine("PCNAME : " & pcname)
                        file2.WriteLine("DATETIME : " & DateTime.Now)
                        file2.Dispose()
                        file2.Close()
                        file1.Close()
                        TextBox2.Clear()
                        cn.Close()
                    End If
                Else
                    If reader("SW_VERSION") = "TRIAL" And TextBox3.Text = "PROFFESIONAL" Then
                        reader.Close()
                        Dim command As New OleDbCommand(SQL1, cn)
                        command.ExecuteNonQuery()
                        Dim command1 As New OleDbCommand(SQL2, cn)
                        command1.ExecuteNonQuery()
                        Dim iret1 As Object = MsgBox("SOFTWARE ACTIVATION SUCCESSFULL. PLEASE RESART THE SOFTWARE", vbInformation + vbOKOnly, "LISCENCE MANAGER")
                        If iret1 = vbOK Then
                            MAIN.ListBox1.Items.Add("FUNNCLIB.InstallLiscence returns : True [Previous Insalled Version : TRIAL / Installing Version : PROFFESIONAL]")
                            file2.WriteLine("ATTEMPTED FOR LISCENCE INSTALATTION : True [Previous Installed Version : TRIAL / Installing Version : PROFFESIONAL]")
                            file2.WriteLine("PCNAME : " & pcname)
                            file2.WriteLine("DATETIME : " & DateTime.Now)
                            file2.WriteLine("EXPIRY DATE : " & TextBox5.Text)
                            file2.WriteLine("USED SERIAL NUMBER : " & TextBox2.Text)
                            file2.Dispose()
                            file2.Close()
                            file1.Close()
                            Application.Restart()
                        End If
                    Else
                        If reader("SW_VERSION") = "PROFFESIONAL" And TextBox3.Text = "PROFFESIONAL" Then
                            reader.Close()
                            Dim command As New OleDbCommand(SQL1, cn)
                            command.ExecuteNonQuery()
                            Dim command1 As New OleDbCommand(SQL2, cn)
                            command1.ExecuteNonQuery()
                            Dim iret2 As Object = MsgBox("SOFTWARE ACTIVATION SUCCESSFULL. PLEASE RESART THE SOFTWARE", vbInformation + vbOKOnly, "LISCENCE MANAGER")
                            If iret2 = vbOK Then
                                MAIN.ListBox1.Items.Add("FUNNCLIB.InstallLiscence returns : True [Previous Insalled Version : PROFFESIONAL / Installing Version : PROFFESIONAL]")
                                file2.WriteLine("ATTEMPTED FOR LISCENCE UPGRADE : True [Previous Installed Version : PROFFESIONAL / Installing Version : PROFFESIONAL]")
                                file2.WriteLine("PCNAME : " & pcname)
                                file2.WriteLine("DATETIME : " & DateTime.Now)
                                file2.WriteLine("EXPIRY DATE : " & TextBox5.Text)
                                file2.WriteLine("USED SERIAL NUMBER : " & TextBox2.Text)
                                file2.Dispose()
                                file2.Close()
                                file1.Close()
                                Application.Restart()
                            End If
                        End If
                    End If
                End If
            Else
                reader.Close()
                Dim command1 As New OleDbCommand(SQL2, cn)
                command1.ExecuteNonQuery()
                Dim iret3 As Object = MsgBox("SOFTWARE ACTIVATION SUCCESSFULL. PLEASE RESART THE SOFTWARE", vbInformation + vbOKOnly, "LISCENCE MANAGER")
                If iret3 = vbOK Then
                    MAIN.ListBox1.Items.Add("FUNNCLIB.InstallLiscence returns : True [Previous Insalled Version : NA / Installing Version : " & TextBox3.Text & "]")
                    file2.WriteLine("ATTEMPTED FOR LISCENCE INSTALATTION : True [Previous Installed Version : TRIAL / Installing Version :" & TextBox3.Text & "]")
                    file2.WriteLine("PCNAME : " & pcname)
                    file2.WriteLine("DATETIME : " & DateTime.Now)
                    file2.WriteLine("EXPIRY DATE : " & TextBox5.Text)
                    file2.WriteLine("USED SERIAL NUMBER : " & TextBox2.Text)
                    file2.Dispose()
                    file2.Close()
                    file1.Close()
                    Application.Restart()
                End If
            End If
        Catch ex As Exception
            Dim iret As Object = MsgBox("Error Installing Liscence. Contact to Administrator...", vbRetryCancel + vbCritical, "Error Installing Liscence")
            If iret = vbRetry Then
                MAIN.ListBox1.Items.Add("InstallLisce.Status : False" & "[" & ex.Message & "]")
                file1.Close()
                file2.Close()
                GoTo line1
            Else
                file1.WriteLine("Error Installing Liscence : " & ex.Message & " " & "@" & DateTime.Now.ToString("yyyy/MM/dd/HH:mm:ss"))
                file1.Close()
                file2.Close()
                Application.Exit()
            End If
        End Try
    End Sub

    Private Sub Button2_Click(sender As Object, e As EventArgs) Handles Button2.Click
        Me.Close()
    End Sub
End Class