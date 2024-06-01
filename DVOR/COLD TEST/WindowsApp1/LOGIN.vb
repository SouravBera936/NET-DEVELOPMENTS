Imports MadMilkman.Ini
Imports System.Diagnostics
Public Class LOGIN
    Dim line1 As Label
    Dim wb As Object
    Dim xlapp As Object
    Dim st As Object
    Private Sub LOGIN_Load(sender As Object, e As EventArgs) Handles MyBase.Load
        TextBox1.Clear()
        TextBox2.Clear()
        Try
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
            Dim ini As New IniFile()
            ini.Load("C:\DVOR\COLD TEST\WindowsApp1\obj\CONFRIGURATIONS.ini")
            CONFIG.TextBox1.Text = ini.Sections("USER_DATA").Keys("Count").Value
            CONFIG.TextBox2.Text = ini.Sections("TRACKER").Keys("Count").Value
            CONFIG.TextBox3.Text = ini.Sections("COUPLER_TEMPLATE").Keys("Count").Value
            CONFIG.TextBox4.Text = ini.Sections("COUPLER_LOG").Keys("Count").Value
            CONFIG.TextBox5.Text = ini.Sections("COUPLER_1").Keys("Count").Value
            CONFIG.TextBox6.Text = ini.Sections("COUPLER_2").Keys("Count").Value
            CONFIG.TextBox7.Text = ini.Sections("COUPLER_3").Keys("Count").Value
            CONFIG.TextBox8.Text = ini.Sections("COUPLER_4").Keys("Count").Value
            CONFIG.TextBox9.Text = ini.Sections("COUPLER_5").Keys("Count").Value
            CONFIG.TextBox10.Text = ini.Sections("COUPLER_6").Keys("Count").Value
            CONFIG.TextBox13.Text = ini.Sections("STATUS_TEMPLATE").Keys("Count").Value
            CONFIG.TextBox12.Text = ini.Sections("STATUS_LOG").Keys("Count").Value
            CONFIG.TextBox11.Text = ini.Sections("STATUS_1").Keys("Count").Value
        Catch ex As Exception
            Dim iret As Object = MsgBox("ERROR LOADING CONFRIGURATIONS.ini FILE." & "-" & ex.Message, vbCritical + vbOKOnly, "INITIALIZATION ERROR")
            If iret = vbOK Then
                Dim file As System.IO.StreamWriter
                Dim path As String = "C:\DVOR\COLD TEST\WindowsApp1\obj\TRACKER.txt"
                file = My.Computer.FileSystem.OpenTextFileWriter(path, True)
                file.WriteLine("INITIALIZATION ERROR :" & ex.Message & "_" & DateTime.Now.ToString("yyyy/MM/dd/HH:mm:ss"))
                file.Close()
                Application.Exit()
            End If
        End Try

    End Sub

    Private Sub TextBox2_TextChanged(sender As Object, e As EventArgs) Handles TextBox2.TextChanged
        TextBox2.PasswordChar = "*"
    End Sub

    Private Sub Button2_Click(sender As Object, e As EventArgs) Handles Button2.Click
        Dim iret As Object = MsgBox("DO YOU WANT TO EXIT SOFTWARE?", vbQuestion + vbYesNo, "EXIT LOGIC")
        If iret = vbYes Then
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
            Application.Exit()
        Else
            If iret = vbNo Then
                'donothing
            End If
        End If
    End Sub

    Private Sub Button1_Click(sender As Object, e As EventArgs) Handles Button1.Click
        Try
            Dim file1 As System.IO.StreamWriter
            Dim path1 As String = CONFIG.TextBox2.Text.ToString()
            xlapp = CreateObject("excel.application")
            wb = xlapp.workbooks.open(CONFIG.TextBox1.Text.ToString())
            st = wb.worksheets("USER DATA")
            file1 = My.Computer.FileSystem.OpenTextFileWriter(path1, True)
            If TextBox1.Text = "" And TextBox2.Text = "" Then
                Dim iret As Object = MsgBox("INVALID USER OR PASSWORD", vbCritical + vbOKOnly, "LOGIN FAIL")
                If iret = vbOK Then
                    file1.WriteLine("LOGIN FAILED :" + "=" + TextBox1.Text + "_" & DateTime.Now.ToString("yyyy/MM/dd/HH:mm:ss"))
                    TextBox1.Clear()
                    TextBox2.Clear()
                    file1.Close()
                    GoTo line1
                End If
            Else
                If TextBox1.Text = "USER 1" And TextBox2.Text = "" Then
                    Dim iret1 As Object = MsgBox("INVALID USER OR PASSWORD", vbCritical + vbOKOnly, "LOGIN FAIL")
                    If iret1 = vbOK Then
                        file1.WriteLine("LOGIN FAILED :" + "=" + TextBox1.Text + "_" & DateTime.Now.ToString("yyyy/MM/dd/HH:mm:ss"))
                        file1.Close()
                        TextBox1.Clear()
                        TextBox2.Clear()
                        GoTo line1
                    End If
                Else
                    If TextBox1.Text = "USER 2" And TextBox2.Text = "" Then
                        Dim iret2 As Object = MsgBox("INVALID USER OR PASSWORD", vbCritical + vbOKOnly, "LOGIN FAIL")
                        If iret2 = vbOK Then
                            file1.WriteLine("LOGIN FAILED :" + "=" + TextBox1.Text + "_" & DateTime.Now.ToString("yyyy/MM/dd/HH:mm:ss"))
                            file1.Close()
                            TextBox1.Clear()
                            TextBox2.Clear()
                            GoTo line1
                        End If
                    Else
                        If TextBox1.Text = "USER 3" And TextBox2.Text = "" Then
                            Dim iret3 As Object = MsgBox("INVALID USER OR PASSWORD", vbCritical + vbOKOnly, "LOGIN FAIL")
                            If iret3 = vbOK Then
                                file1.WriteLine("LOGIN FAILED :" + "=" + TextBox1.Text + "_" & DateTime.Now.ToString("yyyy/MM/dd/HH:mm:ss"))
                                file1.Close()
                                TextBox1.Clear()
                                TextBox2.Clear()
                                GoTo line1
                            End If
                        Else
                            If TextBox1.Text = "USER 4" And TextBox2.Text = "" Then
                                Dim iret4 As Object = MsgBox("INVALID USER OR PASSWORD", vbCritical + vbOKOnly, "LOGIN FAIL")
                                If iret4 = vbOK Then
                                    file1.WriteLine("LOGIN FAILED :" + "=" + TextBox1.Text + "_" & DateTime.Now.ToString("yyyy/MM/dd/HH:mm:ss"))
                                    file1.Close()
                                    TextBox1.Clear()
                                    TextBox2.Clear()
                                    GoTo line1
                                End If
                            Else
                                If TextBox1.Text = "USER 5" And TextBox2.Text = "" Then
                                    Dim iret5 As Object = MsgBox("INVALID USER OR PASSWORD", vbCritical + vbOKOnly, "LOGIN FAIL")
                                    If iret5 = vbOK Then
                                        file1.WriteLine("LOGIN FAILED :" + "=" + TextBox1.Text + "_" & DateTime.Now.ToString("yyyy/MM/dd/HH:mm:ss"))
                                        file1.Close()
                                        TextBox1.Clear()
                                        TextBox2.Clear()
                                        GoTo line1
                                    End If
                                Else
                                    If TextBox1.Text = "USER 6" And TextBox2.Text = "" Then
                                        Dim iret6 As Object = MsgBox("INVALID USER OR PASSWORD", vbCritical + vbOKOnly, "LOGIN FAIL")
                                        If iret6 = vbOK Then
                                            file1.WriteLine("LOGIN FAILED :" + "=" + TextBox1.Text + "_" & DateTime.Now.ToString("yyyy/MM/dd/HH:mm:ss"))
                                            file1.Close()
                                            TextBox1.Clear()
                                            TextBox2.Clear()
                                            GoTo line1
                                        End If
                                    Else
                                        If TextBox1.Text = "USER 7" And TextBox2.Text = "" Then
                                            Dim iret7 As Object = MsgBox("INVALID USER OR PASSWORD", vbCritical + vbOKOnly, "LOGIN FAIL")
                                            If iret7 = vbOK Then
                                                file1.WriteLine("LOGIN FAILED :" + "=" + TextBox1.Text + "_" & DateTime.Now.ToString("yyyy/MM/dd/HH:mm:ss"))
                                                file1.Close()
                                                TextBox1.Clear()
                                                TextBox2.Clear()
                                                GoTo line1
                                            End If
                                        Else
                                            If TextBox1.Text = "USER 8" And TextBox2.Text = "" Then
                                                Dim iret8 As Object = MsgBox("INVALID USER OR PASSWORD", vbCritical + vbOKOnly, "LOGIN FAIL")
                                                If iret8 = vbOK Then
                                                    file1.WriteLine("LOGIN FAILED :" + "=" + TextBox1.Text + "_" & DateTime.Now.ToString("yyyy/MM/dd/HH:mm:ss"))
                                                    file1.Close()
                                                    TextBox1.Clear()
                                                    TextBox2.Clear()
                                                    GoTo line1
                                                End If
                                            Else
                                                If TextBox1.Text = "USER 9" And TextBox2.Text = "" Then
                                                    Dim iret9 As Object = MsgBox("INVALID USER OR PASSWORD", vbCritical + vbOKOnly, "LOGIN FAIL")
                                                    If iret9 = vbOK Then
                                                        file1.WriteLine("LOGIN FAILED :" + "=" + TextBox1.Text + "_" & DateTime.Now.ToString("yyyy/MM/dd/HH:mm:ss"))
                                                        file1.Close()
                                                        TextBox1.Clear()
                                                        TextBox2.Clear()
                                                        GoTo line1
                                                    End If
                                                Else
                                                    If TextBox1.Text = "USER 10" And TextBox2.Text = "" Then
                                                        Dim iret10 As Object = MsgBox("INVALID USER OR PASSWORD", vbCritical + vbOKOnly, "LOGIN FAIL")
                                                        If iret10 = vbOK Then
                                                            file1.WriteLine("LOGIN FAILED :" + "=" + TextBox1.Text + "_" & DateTime.Now.ToString("yyyy/MM/dd/HH:mm:ss"))
                                                            file1.Close()
                                                            TextBox1.Clear()
                                                            TextBox2.Clear()
                                                            GoTo line1
                                                        End If
                                                    Else
                                                        If TextBox1.Text = "ADMINISTRATOR" And TextBox2.Text = "Test@1234" Then
                                                            Dim iret11 As Object = MsgBox("LOGIN SUCCESS AS ADMINISTRATOR", vbInformation + vbOKOnly, "LOGIN SUCCESS")
                                                            If iret11 = vbOK Then
                                                                file1.WriteLine("LOGIN SUCCESS :" & "=" & TextBox1.Text & "_" & DateTime.Now.ToString("yyyy/MM/dd/HH:mm:ss"))
                                                                file1.Close()
                                                                MAIN.TextBox1.Text = TextBox1.Text
                                                                MAIN.TextBox2.Text = "DEVELOPER"
                                                                Me.Hide()
                                                                MAIN.ShowDialog()
                                                                GoTo line1
                                                            End If
                                                        Else
                                                            If TextBox1.Text = st.range("B2").value And TextBox2.Text = st.range("C2").value Then
                                                                Dim iret12 As Object = MsgBox("LOGIN SUCCESS AS" & " " & st.range("A2").value, vbInformation + vbOKOnly, "LOGIN LOGIC")
                                                                If iret12 = vbOK Then
                                                                    file1.WriteLine("LOGIN SUCCESS :" & "=" & TextBox1.Text & "_" & DateTime.Now.ToString("yyyy/MM/dd/HH:mm:ss"))
                                                                    file1.Close()
                                                                    MAIN.TextBox1.Text = TextBox1.Text
                                                                    MAIN.TextBox2.Text = st.range("A2").value
                                                                    Me.Hide()
                                                                    MAIN.ShowDialog()
                                                                    GoTo line1
                                                                End If
                                                            Else
                                                                If TextBox1.Text = st.range("B3").value And TextBox2.Text = st.range("C3").value Then
                                                                    Dim iret13 As Object = MsgBox("LOGIN SUCCESS AS" & " " & st.range("A3").value, vbInformation + vbOKOnly, "LOGIN LOGIC")
                                                                    If iret13 = vbOK Then
                                                                        file1.WriteLine("LOGIN SUCCESS :" & "=" & TextBox1.Text & "_" & DateTime.Now.ToString("yyyy/MM/dd/HH:mm:ss"))
                                                                        file1.Close()
                                                                        MAIN.TextBox1.Text = TextBox1.Text
                                                                        MAIN.TextBox2.Text = st.range("A3").value
                                                                        Me.Hide()
                                                                        MAIN.ShowDialog()
                                                                        GoTo line1
                                                                    End If
                                                                Else
                                                                    If TextBox1.Text = st.range("B4").value And TextBox2.Text = st.range("C4").value Then
                                                                        Dim iret14 As Object = MsgBox("LOGIN SUCCESS AS" & " " & st.range("A4").value, vbInformation + vbOKOnly, "LOGIN LOGIC")
                                                                        If iret14 = vbOK Then
                                                                            file1.WriteLine("LOGIN SUCCESS :" & "=" & TextBox1.Text & "_" & DateTime.Now.ToString("yyyy/MM/dd/HH:mm:ss"))
                                                                            file1.Close()
                                                                            MAIN.TextBox1.Text = TextBox1.Text
                                                                            MAIN.TextBox2.Text = st.range("A4").value
                                                                            Me.Hide()
                                                                            MAIN.ShowDialog()
                                                                            GoTo line1
                                                                        End If
                                                                    Else
                                                                        If TextBox1.Text = st.range("B5").value And TextBox2.Text = st.range("C5").value Then
                                                                            Dim iret15 As Object = MsgBox("LOGIN SUCCESS AS" & " " & st.range("A5").value, vbInformation + vbOKOnly, "LOGIN LOGIC")
                                                                            If iret15 = vbOK Then
                                                                                file1.WriteLine("LOGIN SUCCESS :" & "=" & TextBox1.Text & "_" & DateTime.Now.ToString("yyyy/MM/dd/HH:mm:ss"))
                                                                                file1.Close()
                                                                                MAIN.TextBox1.Text = TextBox1.Text
                                                                                MAIN.TextBox2.Text = st.range("A5").value
                                                                                Me.Hide()
                                                                                MAIN.ShowDialog()
                                                                                GoTo line1
                                                                            End If
                                                                        Else
                                                                            If TextBox1.Text = st.range("B6").value And TextBox2.Text = st.range("C6").value Then
                                                                                Dim iret16 As Object = MsgBox("LOGIN SUCCESS AS" & " " & st.range("A6").value, vbInformation + vbOKOnly, "LOGIN LOGIC")
                                                                                If iret16 = vbOK Then
                                                                                    file1.WriteLine("LOGIN SUCCESS :" & "=" & TextBox1.Text & "_" & DateTime.Now.ToString("yyyy/MM/dd/HH:mm:ss"))
                                                                                    file1.Close()
                                                                                    MAIN.TextBox1.Text = TextBox1.Text
                                                                                    MAIN.TextBox2.Text = st.range("A6").value
                                                                                    Me.Hide()
                                                                                    MAIN.ShowDialog()
                                                                                    GoTo line1
                                                                                End If
                                                                            Else
                                                                                If TextBox1.Text = st.range("B7").value And TextBox2.Text = st.range("C7").value Then
                                                                                    Dim iret17 As Object = MsgBox("LOGIN SUCCESS AS" & " " & st.range("A7").value, vbInformation + vbOKOnly, "LOGIN LOGIC")
                                                                                    If iret17 = vbOK Then
                                                                                        file1.WriteLine("LOGIN SUCCESS :" & "=" & TextBox1.Text & "_" & DateTime.Now.ToString("yyyy/MM/dd/HH:mm:ss"))
                                                                                        file1.Close()
                                                                                        MAIN.TextBox1.Text = TextBox1.Text
                                                                                        MAIN.TextBox2.Text = st.range("A7").value
                                                                                        Me.Hide()
                                                                                        MAIN.ShowDialog()
                                                                                        GoTo line1
                                                                                    End If
                                                                                Else
                                                                                    If TextBox1.Text = st.range("B8").value And TextBox2.Text = st.range("C8").value Then
                                                                                        Dim iret18 As Object = MsgBox("LOGIN SUCCESS AS" & " " & st.range("A8").value, vbInformation + vbOKOnly, "LOGIN LOGIC")
                                                                                        If iret18 = vbOK Then
                                                                                            file1.WriteLine("LOGIN SUCCESS :" & "=" & TextBox1.Text & "_" & DateTime.Now.ToString("yyyy/MM/dd/HH:mm:ss"))
                                                                                            file1.Close()
                                                                                            MAIN.TextBox1.Text = TextBox1.Text
                                                                                            MAIN.TextBox2.Text = st.range("A8").value
                                                                                            Me.Hide()
                                                                                            MAIN.ShowDialog()
                                                                                            GoTo line1
                                                                                        End If
                                                                                    Else
                                                                                        If TextBox1.Text = st.range("B9").value And TextBox2.Text = st.range("C9").value Then
                                                                                            Dim iret19 As Object = MsgBox("LOGIN SUCCESS AS" & " " & st.range("A9").value, vbInformation + vbOKOnly, "LOGIN LOGIC")
                                                                                            If iret19 = vbOK Then
                                                                                                file1.WriteLine("LOGIN SUCCESS :" & "=" & TextBox1.Text & "_" & DateTime.Now.ToString("yyyy/MM/dd/HH:mm:ss"))
                                                                                                file1.Close()
                                                                                                MAIN.TextBox1.Text = TextBox1.Text
                                                                                                MAIN.TextBox2.Text = st.range("A9").value
                                                                                                Me.Hide()
                                                                                                MAIN.ShowDialog()
                                                                                                GoTo line1
                                                                                            End If
                                                                                        Else
                                                                                            If TextBox1.Text = st.range("B10").value And TextBox2.Text = st.range("C10").value Then
                                                                                                Dim iret20 As Object = MsgBox("LOGIN SUCCESS AS" & " " & st.range("A10").value, vbInformation + vbOKOnly, "LOGIN LOGIC")
                                                                                                If iret20 = vbOK Then
                                                                                                    file1.WriteLine("LOGIN SUCCESS :" & "=" & TextBox1.Text & "_" & DateTime.Now.ToString("yyyy/MM/dd/HH:mm:ss"))
                                                                                                    file1.Close()
                                                                                                    MAIN.TextBox1.Text = TextBox1.Text
                                                                                                    MAIN.TextBox2.Text = st.range("A10").value
                                                                                                    Me.Hide()
                                                                                                    MAIN.ShowDialog()
                                                                                                    GoTo line1
                                                                                                End If
                                                                                            Else
                                                                                                If TextBox1.Text = st.range("B11").value And TextBox2.Text = st.range("C11").value Then
                                                                                                    Dim iret21 As Object = MsgBox("LOGIN SUCCESS AS" & " " & st.range("A11").value, vbInformation + vbOKOnly, "LOGIN LOGIC")
                                                                                                    If iret21 = vbOK Then
                                                                                                        file1.WriteLine("LOGIN SUCCESS :" & "=" & TextBox1.Text & "_" & DateTime.Now.ToString("yyyy/MM/dd/HH:mm:ss"))
                                                                                                        file1.Close()
                                                                                                        MAIN.TextBox1.Text = TextBox1.Text
                                                                                                        MAIN.TextBox2.Text = st.range("A11").value
                                                                                                        Me.Hide()
                                                                                                        MAIN.ShowDialog()
                                                                                                        GoTo line1
                                                                                                    End If
                                                                                                Else
                                                                                                    Dim iret23 As Object = MsgBox("INVALID USER ID OR PASSWORD", vbExclamation + vbOKOnly, "LOGIN FAIL")
                                                                                                    If iret23 = vbOK Then
                                                                                                        file1.WriteLine("LOGIN FAILED :" & "=" & TextBox1.Text & "_" & DateTime.Now.ToString("yyyy/MM/dd/HH:mm:ss"))
                                                                                                        file1.Close()
                                                                                                        TextBox1.Clear()
                                                                                                        TextBox2.Clear()
                                                                                                        GoTo line1
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


line1:
            wb.close(True)
            System.Runtime.InteropServices.Marshal.ReleaseComObject(wb)
            wb = Nothing
            xlapp.quit()
            System.Runtime.InteropServices.Marshal.ReleaseComObject(xlapp)
            xlapp = Nothing
            Process.GetProcessesByName("EXCEL")(0).Kill()
        Catch ex As Exception
            Dim iret22 As Object = MsgBox("ERROR LOGING INTO THE SOFTWARE :" & "-" & ex.Message, vbCritical + vbOKOnly, "LOGIN ERROR")
            If iret22 = vbOK Then
                Dim file2 As System.IO.StreamWriter
                Dim path2 As String = CONFIG.TextBox2.Text.ToString()
                file2 = My.Computer.FileSystem.OpenTextFileWriter(path2, True)
                file2.WriteLine("LOGIN ERROR :" & ex.Message & "_" & DateTime.Now.ToString("yyyy/MM/dd/HH:mm:ss"))
                file2.Close()
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
                    Application.Exit()
            End If
        End Try

    End Sub


End Class
