Imports MadMilkman.Ini
Imports System.Diagnostics
Public Class LOGIN
    Private Sub TextBox2_TextChanged(sender As Object, e As EventArgs) Handles TextBox2.TextChanged
        TextBox2.PasswordChar = "*"
    End Sub

    Private Sub LOGIN_Load(sender As Object, e As EventArgs) Handles MyBase.Load
        TextBox1.Text = ""
        TextBox2.Text = ""
        Dim ini As New IniFile()
        ini.Load("C:\SVI 2 PROGRAMMING SOFTWARE\SVI_2 PROGRAMMING\SVI_2 PROGRAMMING\obj\CONFRIGURATIONS\CONFIG.ini")
        CONFIG.TextBox1.Text = ini.Sections("USER_TEMPLATE").Keys("Count").Value
        CONFIG.TextBox2.Text = ini.Sections("ACTIVITY_LOG").Keys("Count").Value
        CONFIG.TextBox7.Text = ini.Sections("DB_PATH").Keys("Count").Value
        CONFIG.TextBox4.Text = ini.Sections("POWER_SUPPLY").Keys("Count").Value
        CONFIG.TextBox5.Text = ini.Sections("MNMMA_PATH").Keys("Count").Value
        CONFIG.TextBox6.Text = ini.Sections("LOG_PATH").Keys("Count").Value
    End Sub

    Private Sub Button1_Click(sender As Object, e As EventArgs) Handles Button1.Click
        Dim line1 As Label
        Dim iret As Object
        Dim iret1 As Object
        Dim iret2 As Object
        Dim iret3 As Object
        Dim iret4 As Object
        Dim iret5 As Object
        Dim iret6 As Object
        Dim iret7 As Object
        Dim iret8 As Object
        Dim iret9 As Object
        Dim iret10 As Object
        Dim iret11 As Object
        Dim iret12 As Object
        Dim iret13 As Object
        Dim iret14 As Object
        Dim iret15 As Object
        Dim iret16 As Object
        Dim iret17 As Object
        Dim iret18 As Object
        Dim iret19 As Object
        Dim iret20 As Object
        Dim iret21 As Object
        Dim iret22 As Object
        Dim iret23 As Object
        Dim iret24 As Object
        Dim iret25 As Object
        Dim iret26 As Object
        Dim iret27 As Object
        Dim iret28 As Object
        Dim iret29 As Object
        Dim iret30 As Object
        Dim iret31 As Object
        Dim iret32 As Object
        Dim iret33 As Object
        Dim iret34 As Object
        Dim iret35 As Object
        Dim iret36 As Object
        Dim iret37 As Object
        Dim iret38 As Object
        Dim iret39 As Object
        Dim iret40 As Object
        Dim iret41 As Object
        Dim iret42 As Object
        Dim iret43 As Object
        Dim iret44 As Object
        Dim iret45 As Object
        Dim iret46 As Object
        Dim iret47 As Object
        Dim iret48 As Object
        Dim iret49 As Object
        Dim iret50 As Object
        Dim iret51 As Object
        Dim iret52 As Object
        Dim iret53 As Object
        Dim iret54 As Object
        Dim iret55 As Object
        Dim iret56 As Object
        Dim iret57 As Object
        Dim iret58 As Object
        Dim iret59 As Object
        Dim iret60 As Object
        Dim iret61 As Object
        Dim iret62 As Object
        Dim wb As Object
        Dim xlapp As Object
        Dim st As Object
        xlapp = CreateObject("excel.application")
        wb = xlapp.workbooks.open(CONFIG.TextBox1.Text.ToString())
        st = wb.worksheets("Sheet1")
        Dim file As System.IO.StreamWriter
        Dim path As String = CONFIG.TextBox2.Text.ToString()
        file = My.Computer.FileSystem.OpenTextFileWriter(path, True)
        If TextBox1.Text = "" And TextBox2.Text = "" Then
            iret = MsgBox("INVALID USER OR PASSWORD", vbCritical + vbOKOnly, "LOGIN FAILED")
            If iret = vbOK Then
                file.WriteLine("LOGIN FAILED :" + "=" + TextBox1.Text + "_" & DateTime.Now.ToString("yyyy/MM/dd/HH:mm:ss"))
                file.Close()
                TextBox1.Text = ""
                TextBox2.Text = ""
                GoTo line1
            End If
        Else
            If TextBox1.Text = "USER 1" Then
                iret1 = MsgBox("INVALID USER OR PASSWORD", vbCritical + vbOKOnly, "LOGIN FAILED")
                If iret1 = vbOK Then
                    file.WriteLine("LOGIN FAILED :" + "=" + TextBox1.Text + "_" & DateTime.Now.ToString("yyyy/MM/dd/HH:mm:ss"))
                    file.Close()
                    TextBox1.Text = ""
                    TextBox2.Text = ""
                    GoTo line1
                End If
            Else
                If TextBox1.Text = "USER 2" Then
                    iret2 = MsgBox("INVALID USER OR PASSWORD", vbCritical + vbOKOnly, "LOGIN FAILED")
                    If iret2 = vbOK Then
                        file.WriteLine("LOGIN FAILED :" & "=" & TextBox1.Text & "_" & DateTime.Now.ToString("yyyy/MM/dd/HH:mm:ss"))
                        file.Close()
                        TextBox1.Text = ""
                        TextBox2.Text = ""
                        GoTo line1
                    End If
                Else
                    If TextBox1.Text = "USER 3" Then
                        iret3 = MsgBox("INVALID USER OR PASSWORD", vbCritical + vbOKOnly, "LOGIN FAILED")
                        If iret3 = vbOK Then
                            file.WriteLine("LOGIN FAILED :" & "=" & TextBox1.Text & "_" & DateTime.Now.ToString("yyyy/MM/dd/HH:mm:ss"))
                            file.Close()
                            TextBox1.Text = ""
                            TextBox2.Text = ""
                            GoTo line1
                        End If
                    Else
                        If TextBox1.Text = "USER 4" Then
                            iret4 = MsgBox("INVALID USER OR PASSWORD", vbCritical + vbOKOnly, "LOGIN FAILED")
                            If iret4 = vbOK Then
                                file.WriteLine("LOGIN FAILED :" & "=" & TextBox1.Text & "_" & DateTime.Now.ToString("yyyy/MM/dd/HH:mm:ss"))
                                file.Close()
                                TextBox1.Text = ""
                                TextBox2.Text = ""
                                GoTo line1
                            End If
                        Else
                            If TextBox1.Text = "USER 5" Then
                                iret5 = MsgBox("INVALID USER OR PASSWORD", vbCritical + vbOKOnly, "LOGIN FAILED")
                                If iret5 = vbOK Then
                                    file.WriteLine("LOGIN FAILED :" & "=" & TextBox1.Text & "_" & DateTime.Now.ToString("yyyy/MM/dd/HH:mm:ss"))
                                    file.Close()
                                    TextBox1.Text = ""
                                    TextBox2.Text = ""
                                    GoTo line1
                                End If
                            Else
                                If TextBox1.Text = "USER 6" Then
                                    iret6 = MsgBox("INVALID USER OR PASSWORD", vbCritical + vbOKOnly, "LOGIN FAILED")
                                    If iret6 = vbOK Then
                                        file.WriteLine("LOGIN FAILED :" & "=" & TextBox1.Text & "_" & DateTime.Now.ToString("yyyy/MM/dd/HH:mm:ss"))
                                        file.Close()
                                        TextBox1.Text = ""
                                        TextBox2.Text = ""
                                        GoTo line1
                                    End If
                                Else
                                    If TextBox1.Text = "USER 7" Then
                                        iret7 = MsgBox("INVALID USER OR PASSWORD", vbCritical + vbOKOnly, "LOGIN FAILED")
                                        If iret7 = vbOK Then
                                            file.WriteLine("LOGIN FAILED :" & "=" & TextBox1.Text & "_" & DateTime.Now.ToString("yyyy/MM/dd/HH:mm:ss"))
                                            file.Close()
                                            TextBox1.Text = ""
                                            TextBox2.Text = ""
                                            GoTo line1
                                        End If
                                    Else
                                        If TextBox1.Text = "USER 8" Then
                                            iret8 = MsgBox("INVALID USER OR PASSWORD", vbCritical + vbOKOnly, "LOGIN FAILED")
                                            If iret8 = vbOK Then
                                                file.WriteLine("LOGIN FAILED :" & "=" & TextBox1.Text & "_" & DateTime.Now.ToString("yyyy/MM/dd/HH:mm:ss"))
                                                file.Close()
                                                TextBox1.Text = ""
                                                TextBox2.Text = ""
                                                GoTo line1
                                            End If
                                        Else
                                            If TextBox1.Text = "USER 9" Then
                                                iret9 = MsgBox("INVALID USER OR PASSWORD", vbCritical + vbOKOnly, "LOGIN FAILED")
                                                If iret9 = vbOK Then
                                                    file.WriteLine("LOGIN FAILED :" & "=" & TextBox1.Text & "_" & DateTime.Now.ToString("yyyy/MM/dd/HH:mm:ss"))
                                                    file.Close()
                                                    TextBox1.Text = ""
                                                    TextBox2.Text = ""
                                                    GoTo line1
                                                End If
                                            Else
                                                If TextBox1.Text = "USER 10" Then
                                                    iret10 = MsgBox("INVALID USER OR PASSWORD", vbCritical + vbOKOnly, "LOGIN FAILED")
                                                    If iret10 = vbOK Then
                                                        file.WriteLine("LOGIN FAILED :" & "=" & TextBox1.Text & "_" & DateTime.Now.ToString("yyyy/MM/dd/HH:mm:ss"))
                                                        file.Close()
                                                        TextBox1.Text = ""
                                                        TextBox2.Text = ""
                                                        GoTo line1
                                                    End If
                                                Else
                                                    If TextBox1.Text = "USER 11" Then
                                                        iret11 = MsgBox("INVALID USER OR PASSWORD", vbCritical + vbOKOnly, "LOGIN FAILED")
                                                        If iret11 = vbOK Then
                                                            file.WriteLine("LOGIN FAILED :" & "=" & TextBox1.Text & "_" & DateTime.Now.ToString("yyyy/MM/dd/HH:mm:ss"))
                                                            file.Close()
                                                            TextBox1.Text = ""
                                                            TextBox2.Text = ""
                                                            GoTo line1
                                                        End If
                                                    Else
                                                        If TextBox1.Text = "USER 12" Then
                                                            iret12 = MsgBox("INVALID USER OR PASSWORD", vbCritical + vbOKOnly, "LOGIN FAILED")
                                                            If iret12 = vbOK Then
                                                                file.WriteLine("LOGIN FAILED :" & "=" & TextBox1.Text & "_" & DateTime.Now.ToString("yyyy/MM/dd/HH:mm:ss"))
                                                                file.Close()
                                                                TextBox1.Text = ""
                                                                TextBox2.Text = ""
                                                                GoTo line1
                                                            End If
                                                        Else
                                                            If TextBox1.Text = "USER 13" Then
                                                                iret13 = MsgBox("INVALID USER OR PASSWORD", vbCritical + vbOKOnly, "LOGIN FAILED")
                                                                If iret13 = vbOK Then
                                                                    file.WriteLine("LOGIN FAILED :" & "=" & TextBox1.Text & "_" & DateTime.Now.ToString("yyyy/MM/dd/HH:mm:ss"))
                                                                    file.Close()
                                                                    TextBox1.Text = ""
                                                                    TextBox2.Text = ""
                                                                    GoTo line1
                                                                End If
                                                            Else
                                                                If TextBox1.Text = "USER 14" Then
                                                                    iret14 = MsgBox("INVALID USER OR PASSWORD", vbCritical + vbOKOnly, "LOGIN FAILED")
                                                                    If iret14 = vbOK Then
                                                                        file.WriteLine("LOGIN FAILED :" & "=" & TextBox1.Text & "_" & DateTime.Now.ToString("yyyy/MM/dd/HH:mm:ss"))
                                                                        file.Close()
                                                                        TextBox1.Text = ""
                                                                        TextBox2.Text = ""
                                                                        GoTo line1
                                                                    End If
                                                                Else
                                                                    If TextBox1.Text = "USER 15" Then
                                                                        iret15 = MsgBox("INVALID USER OR PASSWORD", vbCritical + vbOKOnly, "LOGIN FAILED")
                                                                        If iret15 = vbOK Then
                                                                            file.WriteLine("LOGIN FAILED :" & "=" & TextBox1.Text & "_" & DateTime.Now.ToString("yyyy/MM/dd/HH:mm:ss"))
                                                                            file.Close()
                                                                            TextBox1.Text = ""
                                                                            TextBox2.Text = ""
                                                                            GoTo line1
                                                                        End If
                                                                    Else
                                                                        If TextBox1.Text = "USER 16" Then
                                                                            iret16 = MsgBox("INVALID USER OR PASSWORD", vbCritical + vbOKOnly, "LOGIN FAILED")
                                                                            If iret16 = vbOK Then
                                                                                file.WriteLine("LOGIN FAILED :" & "=" & TextBox1.Text & "_" & DateTime.Now.ToString("yyyy/MM/dd/HH:mm:ss"))
                                                                                file.Close()
                                                                                TextBox1.Text = ""
                                                                                TextBox2.Text = ""
                                                                                GoTo line1
                                                                            End If
                                                                        Else
                                                                            If TextBox1.Text = "USER 17" Then
                                                                                iret17 = MsgBox("INVALID USER OR PASSWORD", vbCritical + vbOKOnly, "LOGIN FAILED")
                                                                                If iret17 = vbOK Then
                                                                                    file.WriteLine("LOGIN FAILED :" & "=" & TextBox1.Text & "_" & DateTime.Now.ToString("yyyy/MM/dd/HH:mm:ss"))
                                                                                    file.Close()
                                                                                    TextBox1.Text = ""
                                                                                    TextBox2.Text = ""
                                                                                    GoTo line1
                                                                                End If
                                                                            Else
                                                                                If TextBox1.Text = "USER 18" Then
                                                                                    iret18 = MsgBox("INVALID USER OR PASSWORD", vbCritical + vbOKOnly, "LOGIN FAILED")
                                                                                    If iret18 = vbOK Then
                                                                                        file.WriteLine("LOGIN FAILED :" & "=" & TextBox1.Text & "_" & DateTime.Now.ToString("yyyy/MM/dd/HH:mm:ss"))
                                                                                        file.Close()
                                                                                        TextBox1.Text = ""
                                                                                        TextBox2.Text = ""
                                                                                        GoTo line1
                                                                                    End If
                                                                                Else
                                                                                    If TextBox1.Text = "USER 19" Then
                                                                                        iret19 = MsgBox("INVALID USER OR PASSWORD", vbCritical + vbOKOnly, "LOGIN FAILED")
                                                                                        If iret19 = vbOK Then
                                                                                            file.WriteLine("LOGIN FAILED :" & "=" & TextBox1.Text & "_" & DateTime.Now.ToString("yyyy/MM/dd/HH:mm:ss"))
                                                                                            file.Close()
                                                                                            TextBox1.Text = ""
                                                                                            TextBox2.Text = ""
                                                                                            GoTo line1
                                                                                        End If
                                                                                    Else
                                                                                        If TextBox1.Text = "USER 20" Then
                                                                                            iret20 = MsgBox("INVALID USER OR PASSWORD", vbCritical + vbOKOnly, "LOGIN FAILED")
                                                                                            If iret20 = vbOK Then
                                                                                                file.WriteLine("LOGIN FAILED :" & "=" & TextBox1.Text & "_" & DateTime.Now.ToString("yyyy/MM/dd/HH:mm:ss"))
                                                                                                file.Close()
                                                                                                TextBox1.Text = ""
                                                                                                TextBox2.Text = ""
                                                                                                GoTo line1
                                                                                            End If
                                                                                        Else
                                                                                            If TextBox1.Text = "USER 21" Then
                                                                                                iret21 = MsgBox("INVALID USER OR PASSWORD", vbCritical + vbOKOnly, "LOGIN FAILED")
                                                                                                If iret21 = vbOK Then
                                                                                                    file.WriteLine("LOGIN FAILED :" & "=" & TextBox1.Text & "_" & DateTime.Now.ToString("yyyy/MM/dd/HH:mm:ss"))
                                                                                                    file.Close()
                                                                                                    TextBox1.Text = ""
                                                                                                    TextBox2.Text = ""
                                                                                                    GoTo line1
                                                                                                End If
                                                                                            Else
                                                                                                If TextBox1.Text = "USER 22" Then
                                                                                                    iret22 = MsgBox("INVALID USER OR PASSWORD", vbCritical + vbOKOnly, "LOGIN FAILED")
                                                                                                    If iret22 = vbOK Then
                                                                                                        file.WriteLine("LOGIN FAILED :" & "=" & TextBox1.Text & "_" & DateTime.Now.ToString("yyyy/MM/dd/HH:mm:ss"))
                                                                                                        file.Close()
                                                                                                        TextBox1.Text = ""
                                                                                                        TextBox2.Text = ""
                                                                                                        GoTo line1
                                                                                                    End If
                                                                                                Else
                                                                                                    If TextBox1.Text = "USER 23" Then
                                                                                                        iret23 = MsgBox("INVALID USER OR PASSWORD", vbCritical + vbOKOnly, "LOGIN FAILED")
                                                                                                        If iret23 = vbOK Then
                                                                                                            file.WriteLine("LOGIN FAILED :" & "=" & TextBox1.Text & "_" & DateTime.Now.ToString("yyyy/MM/dd/HH:mm:ss"))
                                                                                                            file.Close()
                                                                                                            TextBox1.Text = ""
                                                                                                            TextBox2.Text = ""
                                                                                                            GoTo line1
                                                                                                        End If
                                                                                                    Else
                                                                                                        If TextBox1.Text = "USER 24" Then
                                                                                                            iret24 = MsgBox("INVALID USER OR PASSWORD", vbCritical + vbOKOnly, "LOGIN FAILED")
                                                                                                            If iret24 = vbOK Then
                                                                                                                file.WriteLine("LOGIN FAILED :" & "=" & TextBox1.Text & "_" & DateTime.Now.ToString("yyyy/MM/dd/HH:mm:ss"))
                                                                                                                file.Close()
                                                                                                                TextBox1.Text = ""
                                                                                                                TextBox2.Text = ""
                                                                                                                GoTo line1
                                                                                                            End If
                                                                                                        Else
                                                                                                            If TextBox1.Text = "USER 25" Then
                                                                                                                iret25 = MsgBox("INVALID USER OR PASSWORD", vbCritical + vbOKOnly, "LOGIN FAILED")
                                                                                                                If iret25 = vbOK Then
                                                                                                                    file.WriteLine("LOGIN FAILED :" & "=" & TextBox1.Text & "_" & DateTime.Now.ToString("yyyy/MM/dd/HH:mm:ss"))
                                                                                                                    file.Close()
                                                                                                                    TextBox1.Text = ""
                                                                                                                    TextBox2.Text = ""
                                                                                                                    GoTo line1
                                                                                                                End If
                                                                                                            Else
                                                                                                                If TextBox1.Text = "USER 26" Then
                                                                                                                    iret26 = MsgBox("INVALID USER OR PASSWORD", vbCritical + vbOKOnly, "LOGIN FAILED")
                                                                                                                    If iret26 = vbOK Then
                                                                                                                        file.WriteLine("LOGIN FAILED :" & "=" & TextBox1.Text & "_" & DateTime.Now.ToString("yyyy/MM/dd/HH:mm:ss"))
                                                                                                                        file.Close()
                                                                                                                        TextBox1.Text = ""
                                                                                                                        TextBox2.Text = ""
                                                                                                                        GoTo line1
                                                                                                                    End If
                                                                                                                Else
                                                                                                                    If TextBox1.Text = "USER 27" Then
                                                                                                                        iret27 = MsgBox("INVALID USER OR PASSWORD", vbCritical + vbOKOnly, "LOGIN FAILED")
                                                                                                                        If iret27 = vbOK Then
                                                                                                                            file.WriteLine("LOGIN FAILED :" & "=" & TextBox1.Text & "_" & DateTime.Now.ToString("yyyy/MM/dd/HH:mm:ss"))
                                                                                                                            file.Close()
                                                                                                                            TextBox1.Text = ""
                                                                                                                            TextBox2.Text = ""
                                                                                                                            GoTo line1
                                                                                                                        End If
                                                                                                                    Else
                                                                                                                        If TextBox1.Text = "USER 28" Then
                                                                                                                            iret28 = MsgBox("INVALID USER OR PASSWORD", vbCritical + vbOKOnly, "LOGIN FAILED")
                                                                                                                            If iret28 = vbOK Then
                                                                                                                                file.WriteLine("LOGIN FAILED :" & "=" & TextBox1.Text & "_" & DateTime.Now.ToString("yyyy/MM/dd/HH:mm:ss"))
                                                                                                                                file.Close()
                                                                                                                                TextBox1.Text = ""
                                                                                                                                TextBox2.Text = ""
                                                                                                                                GoTo line1
                                                                                                                            End If
                                                                                                                        Else
                                                                                                                            If TextBox1.Text = "USER 29" Then
                                                                                                                                iret29 = MsgBox("INVALID USER OR PASSWORD", vbCritical + vbOKOnly, "LOGIN FAILED")
                                                                                                                                If iret29 = vbOK Then
                                                                                                                                    file.WriteLine("LOGIN FAILED :" & "=" & TextBox1.Text & "_" & DateTime.Now.ToString("yyyy/MM/dd/HH:mm:ss"))
                                                                                                                                    file.Close()
                                                                                                                                    TextBox1.Text = ""
                                                                                                                                    TextBox2.Text = ""
                                                                                                                                    GoTo line1
                                                                                                                                End If
                                                                                                                            Else
                                                                                                                                If TextBox1.Text = "USER 30" Then
                                                                                                                                    iret30 = MsgBox("INVALID USER OR PASSWORD", vbCritical + vbOKOnly, "LOGIN FAILED")
                                                                                                                                    If iret30 = vbOK Then
                                                                                                                                        file.WriteLine("LOGIN FAILED :" & "=" & TextBox1.Text & "_" & DateTime.Now.ToString("yyyy/MM/dd/HH:mm:ss"))
                                                                                                                                        file.Close()
                                                                                                                                        TextBox1.Text = ""
                                                                                                                                        TextBox2.Text = ""
                                                                                                                                        GoTo line1
                                                                                                                                    End If
                                                                                                                                Else
                                                                                                                                    If TextBox1.Text = "ADMINISTRATOR" And TextBox2.Text = "BabaMaaDidi@1097" Then
                                                                                                                                        iret31 = MsgBox("LOGIN SUCCESS AS ADMIN", vbInformation + vbOKOnly, "LOGIN_LOGIC")
                                                                                                                                        If iret31 = vbOK Then
                                                                                                                                            file.WriteLine("LOGIN SUCCESS :" & "=" & TextBox1.Text & "_" & DateTime.Now.ToString("yyyy/MM/dd/HH:mm:ss"))
                                                                                                                                            file.Close()
                                                                                                                                            MAIN.TextBox2.Text = "ADMIN"
                                                                                                                                            Me.Hide()
                                                                                                                                            Call MAIN.Show()
                                                                                                                                            GoTo line1
                                                                                                                                        End If
                                                                                                                                    Else
                                                                                                                                        If TextBox1.Text = st.range("A2").value And TextBox2.Text = st.range("C2").value Then
                                                                                                                                            iret32 = MsgBox("LOGIN SUCCESS AS" & st.range("A2").value, vbInformation + vbOKOnly, "LOGIN LOGIC")
                                                                                                                                            If iret32 = vbOK Then
                                                                                                                                                file.WriteLine("LOGIN SUCCESS :" & "=" & TextBox1.Text & "_" & DateTime.Now.ToString("yyyy/MM/dd/HH:mm:ss"))
                                                                                                                                                file.Close()
                                                                                                                                                MAIN.TextBox2.Text = TextBox1.Text
                                                                                                                                                MAIN.TextBox1.Text = st.range("D2").value
                                                                                                                                                Me.Hide()
                                                                                                                                                Call MAIN.Show()
                                                                                                                                                GoTo line1
                                                                                                                                            End If
                                                                                                                                        Else
                                                                                                                                            If TextBox1.Text = st.range("A3").value And TextBox2.Text = st.range("C3").value Then
                                                                                                                                                iret33 = MsgBox("LOGIN SUCCESS AS" & st.range("A3").value, vbInformation + vbOKOnly, "LOGIN LOGIC")
                                                                                                                                                If iret33 = vbOK Then
                                                                                                                                                    file.WriteLine("LOGIN SUCCESS :" & "=" & TextBox1.Text & "_" & DateTime.Now.ToString("yyyy/MM/dd/HH:mm:ss"))
                                                                                                                                                    file.Close()
                                                                                                                                                    MAIN.TextBox2.Text = TextBox1.Text
                                                                                                                                                    MAIN.TextBox1.Text = st.range("D3").value
                                                                                                                                                    Me.Hide()
                                                                                                                                                    Call MAIN.Show()
                                                                                                                                                    GoTo line1
                                                                                                                                                End If
                                                                                                                                            Else
                                                                                                                                                If TextBox1.Text = st.range("A4").value And TextBox2.Text = st.range("C4").value Then
                                                                                                                                                    iret34 = MsgBox("LOGIN SUCCESS AS" & st.range("A4").value, vbInformation + vbOKOnly, "LOGIN LOGIC")
                                                                                                                                                    If iret34 = vbOK Then
                                                                                                                                                        file.WriteLine("LOGIN SUCCESS :" & "=" & TextBox1.Text & "_" & DateTime.Now.ToString("yyyy/MM/dd/HH:mm:ss"))
                                                                                                                                                        file.Close()
                                                                                                                                                        MAIN.TextBox2.Text = TextBox1.Text
                                                                                                                                                        MAIN.TextBox1.Text = st.range("D4").value
                                                                                                                                                        Me.Hide()
                                                                                                                                                        Call MAIN.Show()
                                                                                                                                                        GoTo line1
                                                                                                                                                    End If
                                                                                                                                                Else
                                                                                                                                                    If TextBox1.Text = st.range("A5").value And TextBox2.Text = st.range("C5").value Then
                                                                                                                                                        iret35 = MsgBox("LOGIN SUCCESS AS" & st.range("A5").value, vbInformation + vbOKOnly, "LOGIN LOGIC")
                                                                                                                                                        If iret35 = vbOK Then
                                                                                                                                                            file.WriteLine("LOGIN SUCCESS :" & "=" & TextBox1.Text & "_" & DateTime.Now.ToString("yyyy/MM/dd/HH:mm:ss"))
                                                                                                                                                            file.Close()
                                                                                                                                                            MAIN.TextBox2.Text = TextBox1.Text
                                                                                                                                                            MAIN.TextBox1.Text = st.range("D5").value
                                                                                                                                                            Me.Hide()
                                                                                                                                                            Call MAIN.Show()
                                                                                                                                                            GoTo line1
                                                                                                                                                        End If
                                                                                                                                                    Else
                                                                                                                                                        If TextBox1.Text = st.range("A6").value And TextBox2.Text = st.range("C6").value Then
                                                                                                                                                            iret36 = MsgBox("LOGIN SUCCESS AS" & st.range("A6").value, vbInformation + vbOKOnly, "LOGIN LOGIC")
                                                                                                                                                            If iret36 = vbOK Then
                                                                                                                                                                file.WriteLine("LOGIN SUCCESS :" & "=" & TextBox1.Text & "_" & DateTime.Now.ToString("yyyy/MM/dd/HH:mm:ss"))
                                                                                                                                                                file.Close()
                                                                                                                                                                MAIN.TextBox2.Text = TextBox1.Text
                                                                                                                                                                MAIN.TextBox1.Text = st.range("D6").value
                                                                                                                                                                Me.Hide()
                                                                                                                                                                Call MAIN.Show()
                                                                                                                                                                GoTo line1
                                                                                                                                                            End If
                                                                                                                                                        Else
                                                                                                                                                            If TextBox1.Text = st.range("A7").value And TextBox2.Text = st.range("C7").value Then
                                                                                                                                                                iret37 = MsgBox("LOGIN SUCCESS AS" & st.range("A7").value, vbInformation + vbOKOnly, "LOGIN LOGIC")
                                                                                                                                                                If iret37 = vbOK Then
                                                                                                                                                                    file.WriteLine("LOGIN SUCCESS :" & "=" & TextBox1.Text & "_" & DateTime.Now.ToString("yyyy/MM/dd/HH:mm:ss"))
                                                                                                                                                                    file.Close()
                                                                                                                                                                    MAIN.TextBox2.Text = TextBox1.Text
                                                                                                                                                                    MAIN.TextBox1.Text = st.range("D7").value
                                                                                                                                                                    Me.Hide()
                                                                                                                                                                    Call MAIN.Show()
                                                                                                                                                                    GoTo line1
                                                                                                                                                                End If
                                                                                                                                                            Else
                                                                                                                                                                If TextBox1.Text = st.range("A8").value And TextBox2.Text = st.range("C8").value Then
                                                                                                                                                                    iret38 = MsgBox("LOGIN SUCCESS AS" & st.range("A8").value, vbInformation + vbOKOnly, "LOGIN LOGIC")
                                                                                                                                                                    If iret38 = vbOK Then
                                                                                                                                                                        file.WriteLine("LOGIN SUCCESS :" & "=" & TextBox1.Text & "_" & DateTime.Now.ToString("yyyy/MM/dd/HH:mm:ss"))
                                                                                                                                                                        file.Close()
                                                                                                                                                                        MAIN.TextBox2.Text = TextBox1.Text
                                                                                                                                                                        MAIN.TextBox1.Text = st.range("D8").value
                                                                                                                                                                        Me.Hide()
                                                                                                                                                                        Call MAIN.Show()
                                                                                                                                                                        GoTo line1
                                                                                                                                                                    End If
                                                                                                                                                                Else
                                                                                                                                                                    If TextBox1.Text = st.range("A9").value And TextBox2.Text = st.range("C9").value Then
                                                                                                                                                                        iret39 = MsgBox("LOGIN SUCCESS AS" & st.range("A9").value, vbInformation + vbOKOnly, "LOGIN LOGIC")
                                                                                                                                                                        If iret39 = vbOK Then
                                                                                                                                                                            file.WriteLine("LOGIN SUCCESS :" & "=" & TextBox1.Text & "_" & DateTime.Now.ToString("yyyy/MM/dd/HH:mm:ss"))
                                                                                                                                                                            file.Close()
                                                                                                                                                                            MAIN.TextBox2.Text = TextBox1.Text
                                                                                                                                                                            MAIN.TextBox1.Text = st.range("D9").value
                                                                                                                                                                            Me.Hide()
                                                                                                                                                                            Call MAIN.Show()
                                                                                                                                                                            GoTo line1
                                                                                                                                                                        End If
                                                                                                                                                                    Else
                                                                                                                                                                        If TextBox1.Text = st.range("A10").value And TextBox2.Text = st.range("C10").value Then
                                                                                                                                                                            iret40 = MsgBox("LOGIN SUCCESS AS" & st.range("A10").value, vbInformation + vbOKOnly, "LOGIN LOGIC")
                                                                                                                                                                            If iret40 = vbOK Then
                                                                                                                                                                                file.WriteLine("LOGIN SUCCESS :" & "=" & TextBox1.Text & "_" & DateTime.Now.ToString("yyyy/MM/dd/HH:mm:ss"))
                                                                                                                                                                                file.Close()
                                                                                                                                                                                MAIN.TextBox2.Text = TextBox1.Text
                                                                                                                                                                                MAIN.TextBox1.Text = st.range("D10").value
                                                                                                                                                                                Me.Hide()
                                                                                                                                                                                Call MAIN.Show()
                                                                                                                                                                                GoTo line1
                                                                                                                                                                            End If
                                                                                                                                                                        Else
                                                                                                                                                                            If TextBox1.Text = st.range("A11").value And TextBox2.Text = st.range("C11").value Then
                                                                                                                                                                                iret41 = MsgBox("LOGIN SUCCESS AS" & st.range("A11").value, vbInformation + vbOKOnly, "LOGIN LOGIC")
                                                                                                                                                                                If iret41 = vbOK Then
                                                                                                                                                                                    file.WriteLine("LOGIN SUCCESS :" & "=" & TextBox1.Text & "_" & DateTime.Now.ToString("yyyy/MM/dd/HH:mm:ss"))
                                                                                                                                                                                    file.Close()
                                                                                                                                                                                    MAIN.TextBox2.Text = TextBox1.Text
                                                                                                                                                                                    MAIN.TextBox1.Text = st.range("D11").value
                                                                                                                                                                                    Me.Hide()
                                                                                                                                                                                    Call MAIN.Show()
                                                                                                                                                                                    GoTo line1
                                                                                                                                                                                End If
                                                                                                                                                                            Else
                                                                                                                                                                                If TextBox1.Text = st.range("A12").value And TextBox2.Text = st.range("C12").value Then
                                                                                                                                                                                    iret42 = MsgBox("LOGIN SUCCESS AS" & st.range("A12").value, vbInformation + vbOKOnly, "LOGIN LOGIC")
                                                                                                                                                                                    If iret42 = vbOK Then
                                                                                                                                                                                        file.WriteLine("LOGIN SUCCESS :" & "=" & TextBox1.Text & "_" & DateTime.Now.ToString("yyyy/MM/dd/HH:mm:ss"))
                                                                                                                                                                                        file.Close()
                                                                                                                                                                                        MAIN.TextBox2.Text = TextBox1.Text
                                                                                                                                                                                        MAIN.TextBox1.Text = st.range("D12").value
                                                                                                                                                                                        Me.Hide()
                                                                                                                                                                                        Call MAIN.Show()
                                                                                                                                                                                        GoTo line1
                                                                                                                                                                                    End If
                                                                                                                                                                                Else
                                                                                                                                                                                    If TextBox1.Text = st.range("A13").value And TextBox2.Text = st.range("C13").value Then
                                                                                                                                                                                        iret43 = MsgBox("LOGIN SUCCESS AS" & st.range("A2").value, vbInformation + vbOKOnly, "LOGIN LOGIC")
                                                                                                                                                                                        If iret43 = vbOK Then
                                                                                                                                                                                            file.WriteLine("LOGIN SUCCESS :" & "=" & TextBox1.Text & "_" & DateTime.Now.ToString("yyyy/MM/dd/HH:mm:ss"))
                                                                                                                                                                                            file.Close()
                                                                                                                                                                                            MAIN.TextBox2.Text = TextBox1.Text
                                                                                                                                                                                            MAIN.TextBox1.Text = st.range("D13").value
                                                                                                                                                                                            Me.Hide()
                                                                                                                                                                                            Call MAIN.Show()
                                                                                                                                                                                            GoTo line1
                                                                                                                                                                                        End If
                                                                                                                                                                                    Else
                                                                                                                                                                                        If TextBox1.Text = st.range("A14").value And TextBox2.Text = st.range("C14").value Then
                                                                                                                                                                                            iret44 = MsgBox("LOGIN SUCCESS AS" & st.range("A14").value, vbInformation + vbOKOnly, "LOGIN LOGIC")
                                                                                                                                                                                            If iret44 = vbOK Then
                                                                                                                                                                                                file.WriteLine("LOGIN SUCCESS :" & "=" & TextBox1.Text & "_" & DateTime.Now.ToString("yyyy/MM/dd/HH:mm:ss"))
                                                                                                                                                                                                file.Close()
                                                                                                                                                                                                MAIN.TextBox2.Text = TextBox1.Text
                                                                                                                                                                                                MAIN.TextBox1.Text = st.range("D14").value
                                                                                                                                                                                                Me.Hide()
                                                                                                                                                                                                Call MAIN.Show()
                                                                                                                                                                                                GoTo line1
                                                                                                                                                                                            End If
                                                                                                                                                                                        Else
                                                                                                                                                                                            If TextBox1.Text = st.range("A15").value And TextBox2.Text = st.range("C15").value Then
                                                                                                                                                                                                iret45 = MsgBox("LOGIN SUCCESS AS" & st.range("A15").value, vbInformation + vbOKOnly, "LOGIN LOGIC")
                                                                                                                                                                                                If iret45 = vbOK Then
                                                                                                                                                                                                    file.WriteLine("LOGIN SUCCESS :" & "=" & TextBox1.Text & "_" & DateTime.Now.ToString("yyyy/MM/dd/HH:mm:ss"))
                                                                                                                                                                                                    file.Close()
                                                                                                                                                                                                    MAIN.TextBox2.Text = TextBox1.Text
                                                                                                                                                                                                    MAIN.TextBox1.Text = st.range("D15").value
                                                                                                                                                                                                    Me.Hide()
                                                                                                                                                                                                    Call MAIN.Show()
                                                                                                                                                                                                    GoTo line1
                                                                                                                                                                                                End If
                                                                                                                                                                                            Else
                                                                                                                                                                                                If TextBox1.Text = st.range("A16").value And TextBox2.Text = st.range("C16").value Then
                                                                                                                                                                                                    iret46 = MsgBox("LOGIN SUCCESS AS" & st.range("A16").value, vbInformation + vbOKOnly, "LOGIN LOGIC")
                                                                                                                                                                                                    If iret46 = vbOK Then
                                                                                                                                                                                                        file.WriteLine("LOGIN SUCCESS :" & "=" & TextBox1.Text & "_" & DateTime.Now.ToString("yyyy/MM/dd/HH:mm:ss"))
                                                                                                                                                                                                        file.Close()
                                                                                                                                                                                                        MAIN.TextBox2.Text = TextBox1.Text
                                                                                                                                                                                                        MAIN.TextBox1.Text = st.range("D16").value
                                                                                                                                                                                                        Me.Hide()
                                                                                                                                                                                                        Call MAIN.Show()
                                                                                                                                                                                                        GoTo line1
                                                                                                                                                                                                    End If
                                                                                                                                                                                                Else
                                                                                                                                                                                                    If TextBox1.Text = st.range("A17").value And TextBox2.Text = st.range("C17").value Then
                                                                                                                                                                                                        iret47 = MsgBox("LOGIN SUCCESS AS" & st.range("A17").value, vbInformation + vbOKOnly, "LOGIN LOGIC")
                                                                                                                                                                                                        If iret47 = vbOK Then
                                                                                                                                                                                                            file.WriteLine("LOGIN SUCCESS :" & "=" & TextBox1.Text & "_" & DateTime.Now.ToString("yyyy/MM/dd/HH:mm:ss"))
                                                                                                                                                                                                            file.Close()
                                                                                                                                                                                                            MAIN.TextBox2.Text = TextBox1.Text
                                                                                                                                                                                                            MAIN.TextBox1.Text = st.range("D17").value
                                                                                                                                                                                                            Me.Hide()
                                                                                                                                                                                                            Call MAIN.Show()
                                                                                                                                                                                                            GoTo line1
                                                                                                                                                                                                        End If
                                                                                                                                                                                                    Else
                                                                                                                                                                                                        If TextBox1.Text = st.range("A18").value And TextBox2.Text = st.range("C18").value Then
                                                                                                                                                                                                            iret48 = MsgBox("LOGIN SUCCESS AS" & st.range("A18").value, vbInformation + vbOKOnly, "LOGIN LOGIC")
                                                                                                                                                                                                            If iret48 = vbOK Then
                                                                                                                                                                                                                file.WriteLine("LOGIN SUCCESS :" & "=" & TextBox1.Text & "_" & DateTime.Now.ToString("yyyy/MM/dd/HH:mm:ss"))
                                                                                                                                                                                                                file.Close()
                                                                                                                                                                                                                MAIN.TextBox2.Text = TextBox1.Text
                                                                                                                                                                                                                MAIN.TextBox1.Text = st.range("D18").value
                                                                                                                                                                                                                Me.Hide()
                                                                                                                                                                                                                Call MAIN.Show()
                                                                                                                                                                                                                GoTo line1
                                                                                                                                                                                                            End If
                                                                                                                                                                                                        Else
                                                                                                                                                                                                            If TextBox1.Text = st.range("A19").value And TextBox2.Text = st.range("C19").value Then
                                                                                                                                                                                                                iret49 = MsgBox("LOGIN SUCCESS AS" & st.range("A19").value, vbInformation + vbOKOnly, "LOGIN LOGIC")
                                                                                                                                                                                                                If iret49 = vbOK Then
                                                                                                                                                                                                                    file.WriteLine("LOGIN SUCCESS :" & "=" & TextBox1.Text & "_" & DateTime.Now.ToString("yyyy/MM/dd/HH:mm:ss"))
                                                                                                                                                                                                                    file.Close()
                                                                                                                                                                                                                    MAIN.TextBox2.Text = TextBox1.Text
                                                                                                                                                                                                                    MAIN.TextBox1.Text = st.range("D19").value
                                                                                                                                                                                                                    Me.Hide()
                                                                                                                                                                                                                    Call MAIN.Show()
                                                                                                                                                                                                                    GoTo line1
                                                                                                                                                                                                                End If
                                                                                                                                                                                                            Else
                                                                                                                                                                                                                If TextBox1.Text = st.range("A20").value And TextBox2.Text = st.range("C20").value Then
                                                                                                                                                                                                                    iret50 = MsgBox("LOGIN SUCCESS AS" & st.range("A20").value, vbInformation + vbOKOnly, "LOGIN LOGIC")
                                                                                                                                                                                                                    If iret50 = vbOK Then
                                                                                                                                                                                                                        file.WriteLine("LOGIN SUCCESS :" & "=" & TextBox1.Text & "_" & DateTime.Now.ToString("yyyy/MM/dd/HH:mm:ss"))
                                                                                                                                                                                                                        file.Close()
                                                                                                                                                                                                                        MAIN.TextBox2.Text = TextBox1.Text
                                                                                                                                                                                                                        MAIN.TextBox1.Text = st.range("D20").value
                                                                                                                                                                                                                        Me.Hide()
                                                                                                                                                                                                                        Call MAIN.Show()
                                                                                                                                                                                                                        GoTo line1
                                                                                                                                                                                                                    End If
                                                                                                                                                                                                                Else
                                                                                                                                                                                                                    If TextBox1.Text = st.range("A21").value And TextBox2.Text = st.range("C21").value Then
                                                                                                                                                                                                                        iret51 = MsgBox("LOGIN SUCCESS AS" & st.range("A21").value, vbInformation + vbOKOnly, "LOGIN LOGIC")
                                                                                                                                                                                                                        If iret51 = vbOK Then
                                                                                                                                                                                                                            file.WriteLine("LOGIN SUCCESS :" & "=" & TextBox1.Text & "_" & DateTime.Now.ToString("yyyy/MM/dd/HH:mm:ss"))
                                                                                                                                                                                                                            file.Close()
                                                                                                                                                                                                                            MAIN.TextBox2.Text = TextBox1.Text
                                                                                                                                                                                                                            MAIN.TextBox1.Text = st.range("D21").value
                                                                                                                                                                                                                            Me.Hide()
                                                                                                                                                                                                                            Call MAIN.Show()
                                                                                                                                                                                                                            GoTo line1
                                                                                                                                                                                                                        End If
                                                                                                                                                                                                                    Else
                                                                                                                                                                                                                        If TextBox1.Text = st.range("A22").value And TextBox2.Text = st.range("C22").value Then
                                                                                                                                                                                                                            iret52 = MsgBox("LOGIN SUCCESS AS" & st.range("A22").value, vbInformation + vbOKOnly, "LOGIN LOGIC")
                                                                                                                                                                                                                            If iret52 = vbOK Then
                                                                                                                                                                                                                                file.WriteLine("LOGIN SUCCESS :" & "=" & TextBox1.Text & "_" & DateTime.Now.ToString("yyyy/MM/dd/HH:mm:ss"))
                                                                                                                                                                                                                                file.Close()
                                                                                                                                                                                                                                MAIN.TextBox2.Text = TextBox1.Text
                                                                                                                                                                                                                                MAIN.TextBox1.Text = st.range("D22").value
                                                                                                                                                                                                                                Me.Hide()
                                                                                                                                                                                                                                Call MAIN.Show()
                                                                                                                                                                                                                                GoTo line1
                                                                                                                                                                                                                            End If
                                                                                                                                                                                                                        Else
                                                                                                                                                                                                                            If TextBox1.Text = st.range("A23").value And TextBox2.Text = st.range("C23").value Then
                                                                                                                                                                                                                                iret53 = MsgBox("LOGIN SUCCESS AS" & st.range("A2").value, vbInformation + vbOKOnly, "LOGIN LOGIC")
                                                                                                                                                                                                                                If iret53 = vbOK Then
                                                                                                                                                                                                                                    file.WriteLine("LOGIN SUCCESS :" & "=" & TextBox1.Text & "_" & DateTime.Now.ToString("yyyy/MM/dd/HH:mm:ss"))
                                                                                                                                                                                                                                    file.Close()
                                                                                                                                                                                                                                    MAIN.TextBox2.Text = TextBox1.Text
                                                                                                                                                                                                                                    MAIN.TextBox1.Text = st.range("D23").value
                                                                                                                                                                                                                                    Me.Hide()
                                                                                                                                                                                                                                    Call MAIN.Show()
                                                                                                                                                                                                                                    GoTo line1
                                                                                                                                                                                                                                End If
                                                                                                                                                                                                                            Else
                                                                                                                                                                                                                                If TextBox1.Text = st.range("A24").value And TextBox2.Text = st.range("C24").value Then
                                                                                                                                                                                                                                    iret54 = MsgBox("LOGIN SUCCESS AS" & st.range("A24").value, vbInformation + vbOKOnly, "LOGIN LOGIC")
                                                                                                                                                                                                                                    If iret54 = vbOK Then
                                                                                                                                                                                                                                        file.WriteLine("LOGIN SUCCESS :" & "=" & TextBox1.Text & "_" & DateTime.Now.ToString("yyyy/MM/dd/HH:mm:ss"))
                                                                                                                                                                                                                                        file.Close()
                                                                                                                                                                                                                                        MAIN.TextBox2.Text = TextBox1.Text
                                                                                                                                                                                                                                        MAIN.TextBox1.Text = st.range("D24").value
                                                                                                                                                                                                                                        Me.Hide()
                                                                                                                                                                                                                                        Call MAIN.Show()
                                                                                                                                                                                                                                        GoTo line1
                                                                                                                                                                                                                                    End If
                                                                                                                                                                                                                                Else
                                                                                                                                                                                                                                    If TextBox1.Text = st.range("A25").value And TextBox2.Text = st.range("C25").value Then
                                                                                                                                                                                                                                        iret55 = MsgBox("LOGIN SUCCESS AS" & st.range("A25").value, vbInformation + vbOKOnly, "LOGIN LOGIC")
                                                                                                                                                                                                                                        If iret55 = vbOK Then
                                                                                                                                                                                                                                            file.WriteLine("LOGIN SUCCESS :" & "=" & TextBox1.Text & "_" & DateTime.Now.ToString("yyyy/MM/dd/HH:mm:ss"))
                                                                                                                                                                                                                                            file.Close()
                                                                                                                                                                                                                                            MAIN.TextBox2.Text = TextBox1.Text
                                                                                                                                                                                                                                            MAIN.TextBox1.Text = st.range("D25").value
                                                                                                                                                                                                                                            Me.Hide()
                                                                                                                                                                                                                                            Call MAIN.Show()
                                                                                                                                                                                                                                            GoTo line1
                                                                                                                                                                                                                                        End If
                                                                                                                                                                                                                                    Else
                                                                                                                                                                                                                                        If TextBox1.Text = st.range("A26").value And TextBox2.Text = st.range("C26").value Then
                                                                                                                                                                                                                                            iret56 = MsgBox("LOGIN SUCCESS AS" & st.range("A26").value, vbInformation + vbOKOnly, "LOGIN LOGIC")
                                                                                                                                                                                                                                            If iret56 = vbOK Then
                                                                                                                                                                                                                                                file.WriteLine("LOGIN SUCCESS :" & "=" & TextBox1.Text & "_" & DateTime.Now.ToString("yyyy/MM/dd/HH:mm:ss"))
                                                                                                                                                                                                                                                file.Close()
                                                                                                                                                                                                                                                MAIN.TextBox2.Text = TextBox1.Text
                                                                                                                                                                                                                                                MAIN.TextBox1.Text = st.range("D26").value
                                                                                                                                                                                                                                                Me.Hide()
                                                                                                                                                                                                                                                Call MAIN.Show()
                                                                                                                                                                                                                                                GoTo line1
                                                                                                                                                                                                                                            End If
                                                                                                                                                                                                                                        Else
                                                                                                                                                                                                                                            If TextBox1.Text = st.range("A27").value And TextBox2.Text = st.range("C27").value Then
                                                                                                                                                                                                                                                iret57 = MsgBox("LOGIN SUCCESS AS" & st.range("A27").value, vbInformation + vbOKOnly, "LOGIN LOGIC")
                                                                                                                                                                                                                                                If iret57 = vbOK Then
                                                                                                                                                                                                                                                    file.WriteLine("LOGIN SUCCESS :" & "=" & TextBox1.Text & "_" & DateTime.Now.ToString("yyyy/MM/dd/HH:mm:ss"))
                                                                                                                                                                                                                                                    file.Close()
                                                                                                                                                                                                                                                    MAIN.TextBox2.Text = TextBox1.Text
                                                                                                                                                                                                                                                    MAIN.TextBox1.Text = st.range("D27").value
                                                                                                                                                                                                                                                    Me.Hide()
                                                                                                                                                                                                                                                    Call MAIN.Show()
                                                                                                                                                                                                                                                    GoTo line1
                                                                                                                                                                                                                                                End If
                                                                                                                                                                                                                                            Else
                                                                                                                                                                                                                                                If TextBox1.Text = st.range("A28").value And TextBox2.Text = st.range("C28").value Then
                                                                                                                                                                                                                                                    iret58 = MsgBox("LOGIN SUCCESS AS" & st.range("A28").value, vbInformation + vbOKOnly, "LOGIN LOGIC")
                                                                                                                                                                                                                                                    If iret58 = vbOK Then
                                                                                                                                                                                                                                                        file.WriteLine("LOGIN SUCCESS :" & "=" & TextBox1.Text & "_" & DateTime.Now.ToString("yyyy/MM/dd/HH:mm:ss"))
                                                                                                                                                                                                                                                        file.Close()
                                                                                                                                                                                                                                                        MAIN.TextBox2.Text = TextBox1.Text
                                                                                                                                                                                                                                                        MAIN.TextBox1.Text = st.range("D28").value
                                                                                                                                                                                                                                                        Me.Hide()
                                                                                                                                                                                                                                                        Call MAIN.Show()
                                                                                                                                                                                                                                                        GoTo line1
                                                                                                                                                                                                                                                    End If
                                                                                                                                                                                                                                                Else
                                                                                                                                                                                                                                                    If TextBox1.Text = st.range("A29").value And TextBox2.Text = st.range("C29").value Then
                                                                                                                                                                                                                                                        iret59 = MsgBox("LOGIN SUCCESS AS" & st.range("A29").value, vbInformation + vbOKOnly, "LOGIN LOGIC")
                                                                                                                                                                                                                                                        If iret59 = vbOK Then
                                                                                                                                                                                                                                                            file.WriteLine("LOGIN SUCCESS :" & "=" & TextBox1.Text & "_" & DateTime.Now.ToString("yyyy/MM/dd/HH:mm:ss"))
                                                                                                                                                                                                                                                            file.Close()
                                                                                                                                                                                                                                                            MAIN.TextBox2.Text = TextBox1.Text
                                                                                                                                                                                                                                                            MAIN.TextBox1.Text = st.range("D29").value
                                                                                                                                                                                                                                                            Me.Hide()
                                                                                                                                                                                                                                                            Call MAIN.Show()
                                                                                                                                                                                                                                                            GoTo line1
                                                                                                                                                                                                                                                        End If
                                                                                                                                                                                                                                                    Else
                                                                                                                                                                                                                                                        If TextBox1.Text = st.range("A30").value And TextBox2.Text = st.range("C30").value Then
                                                                                                                                                                                                                                                            iret60 = MsgBox("LOGIN SUCCESS AS" & st.range("A30").value, vbInformation + vbOKOnly, "LOGIN LOGIC")
                                                                                                                                                                                                                                                            If iret60 = vbOK Then
                                                                                                                                                                                                                                                                file.WriteLine("LOGIN SUCCESS :" & "=" & TextBox1.Text & "_" & DateTime.Now.ToString("yyyy/MM/dd/HH:mm:ss"))
                                                                                                                                                                                                                                                                file.Close()
                                                                                                                                                                                                                                                                MAIN.TextBox2.Text = TextBox1.Text
                                                                                                                                                                                                                                                                MAIN.TextBox1.Text = st.range("D30").value
                                                                                                                                                                                                                                                                Me.Hide()
                                                                                                                                                                                                                                                                Call MAIN.Show()
                                                                                                                                                                                                                                                                GoTo line1
                                                                                                                                                                                                                                                            End If
                                                                                                                                                                                                                                                        Else
                                                                                                                                                                                                                                                            If TextBox1.Text = st.range("A31").value And TextBox2.Text = st.range("C31").value Then
                                                                                                                                                                                                                                                                iret61 = MsgBox("LOGIN SUCCESS AS" & st.range("A31").value, vbInformation + vbOKOnly, "LOGIN LOGIC")
                                                                                                                                                                                                                                                                If iret61 = vbOK Then
                                                                                                                                                                                                                                                                    file.WriteLine("LOGIN SUCCESS :" & "=" & TextBox1.Text & "_" & DateTime.Now.ToString("yyyy/MM/dd/HH:mm:ss"))
                                                                                                                                                                                                                                                                    file.Close()
                                                                                                                                                                                                                                                                    MAIN.TextBox2.Text = TextBox1.Text
                                                                                                                                                                                                                                                                    MAIN.TextBox1.Text = st.range("D31").value
                                                                                                                                                                                                                                                                    Me.Hide()
                                                                                                                                                                                                                                                                    Call MAIN.Show()
                                                                                                                                                                                                                                                                    GoTo line1
                                                                                                                                                                                                                                                                End If
                                                                                                                                                                                                                                                            Else
                                                                                                                                                                                                                                                                iret62 = MsgBox("INVALID USER ID OR PASSWORD", vbCritical + vbOKOnly, "LOGIN FAILED")
                                                                                                                                                                                                                                                                If iret62 = vbOK Then
                                                                                                                                                                                                                                                                    file.WriteLine("LOGIN FAILED :" & "=" & TextBox1.Text & "_" & DateTime.Now.ToString("yyyy/MM/dd/HH:mm:ss"))
                                                                                                                                                                                                                                                                    file.Close()
                                                                                                                                                                                                                                                                    TextBox1.Text = ""
                                                                                                                                                                                                                                                                    TextBox2.Text = ""
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
    End Sub

    Private Sub Button3_Click(sender As Object, e As EventArgs) Handles Button3.Click
        Me.Hide()
        Call FP.Show()
    End Sub

    Private Sub Button2_Click(sender As Object, e As EventArgs) Handles Button2.Click
        Dim iret As Object = MsgBox("DO YOU WANT TO EXIT THE SOFTWARE", vbQuestion + vbYesNo, "EXIT_LOGIC")
        If iret = vbYes Then
            Application.Exit()
        Else
            If iret = vbNo Then
                'do nothing
            End If
        End If
    End Sub
End Class
