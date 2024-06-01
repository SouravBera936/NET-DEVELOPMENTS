Imports System.Diagnostics
Public Class FP
    Private Sub Button1_Click(sender As Object, e As EventArgs) Handles Button1.Click
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
        Dim xlapp As Object
        Dim wb As Object
        Dim st As Object
        Dim line1 As Label
        xlapp = CreateObject("excel.application")
        wb = xlapp.workbooks.open(CONFIG.TextBox1.Text.ToString())
        st = wb.worksheets("Sheet1")
        Dim file As System.IO.StreamWriter
        Dim path As String = CONFIG.TextBox2.Text.ToString()
        file = My.Computer.FileSystem.OpenTextFileWriter(path, True)
        If TextBox1.Text = "USER 1" Then
            iret = MsgBox("INVALID USER OR PASSWORD", vbCritical + vbOKOnly, "LOGIN_LOGIC ERROR")
            If iret = vbOK Then
                file.WriteLine("PASSWORD RECOVER FAILED :" + "=" + TextBox1.Text + "_" & DateTime.Now.ToString("yyyy/MM/dd/HH:mm:ss"))
                file.Close()
                GoTo line1
                TextBox1.Text = ""
                TextBox2.Text = ""
                Me.Close()
                Call LOGIN.Show()
            End If
        Else
            If TextBox1.Text = "USER 2" Then
                iret1 = MsgBox("INVALID USER OR PASSWORD", vbCritical + vbOKOnly, "LOGIN_LOGIC ERROR")
                If iret1 = vbOK Then
                    file.WriteLine("PASSWORD RECOVER FAILED :" + "=" + TextBox1.Text + "_" & DateTime.Now.ToString("yyyy/MM/dd/HH:mm:ss"))
                    file.Close()
                    TextBox1.Text = ""
                    TextBox2.Text = ""
                    GoTo line1
                    Me.Close()
                    Call LOGIN.Show()
                End If
            Else
                If TextBox1.Text = "USER 3" Then
                    iret2 = MsgBox("INVALID USER OR PASSWORD", vbCritical + vbOKOnly, "LOGIN_LOGIC ERROR")
                    If iret2 = vbOK Then
                        file.WriteLine("PASSWORD RECOVER FAILED :" + "=" + TextBox1.Text + "_" & DateTime.Now.ToString("yyyy/MM/dd/HH:mm:ss"))
                        file.Close()
                        TextBox1.Text = ""
                        TextBox2.Text = ""
                        GoTo line1
                        Me.Close()
                        Call LOGIN.Show()
                    End If
                Else
                    If TextBox1.Text = "USER 4" Then
                        iret3 = MsgBox("INVALID USER OR PASSWORD", vbCritical + vbOKOnly, "LOGIN_LOGIC ERROR")
                        If iret3 = vbOK Then
                            file.WriteLine("PASSWORD RECOVER FAILED :" + "=" + TextBox1.Text + "_" & DateTime.Now.ToString("yyyy/MM/dd/HH:mm:ss"))
                            file.Close()
                            TextBox1.Text = ""
                            TextBox2.Text = ""
                            GoTo line1
                            Me.Close()
                            Call LOGIN.Show()
                        End If
                    Else
                        If TextBox1.Text = "USER 5" Then
                            iret4 = MsgBox("INVALID USER OR PASSWORD", vbCritical + vbOKOnly, "LOGIN_LOGIC ERROR")
                            If iret4 = vbOK Then
                                file.WriteLine("PASSWORD RECOVER FAILED :" + "=" + TextBox1.Text + "_" & DateTime.Now.ToString("yyyy/MM/dd/HH:mm:ss"))
                                file.Close()
                                TextBox1.Text = ""
                                TextBox2.Text = ""
                                GoTo line1
                                Me.Close()
                                Call LOGIN.Show()
                            End If
                        Else
                            If TextBox1.Text = "USER 6" Then
                                iret5 = MsgBox("INVALID USER OR PASSWORD", vbCritical + vbOKOnly, "LOGIN_LOGIC ERROR")
                                If iret5 = vbOK Then
                                    file.WriteLine("PASSWORD RECOVER FAILED :" + "=" + TextBox1.Text + "_" & DateTime.Now.ToString("yyyy/MM/dd/HH:mm:ss"))
                                    file.Close()
                                    TextBox1.Text = ""
                                    TextBox2.Text = ""
                                    GoTo line1
                                    Me.Close()
                                    Call LOGIN.Show()
                                End If
                            Else
                                If TextBox1.Text = "USER 7" Then
                                    iret6 = MsgBox("INVALID USER OR PASSWORD", vbCritical + vbOKOnly, "LOGIN_LOGIC ERROR")
                                    If iret6 = vbOK Then
                                        file.WriteLine("PASSWORD RECOVER FAILED :" + "=" + TextBox1.Text + "_" & DateTime.Now.ToString("yyyy/MM/dd/HH:mm:ss"))
                                        file.Close()
                                        TextBox1.Text = ""
                                        TextBox2.Text = ""
                                        GoTo line1
                                        Me.Close()
                                        Call LOGIN.Show()
                                    End If
                                Else
                                    If TextBox1.Text = "USER 8" Then
                                        iret7 = MsgBox("INVALID USER OR PASSWORD", vbCritical + vbOKOnly, "LOGIN_LOGIC ERROR")
                                        If iret7 = vbOK Then
                                            file.WriteLine("PASSWORD RECOVER FAILED :" + "=" + TextBox1.Text + "_" & DateTime.Now.ToString("yyyy/MM/dd/HH:mm:ss"))
                                            file.Close()
                                            TextBox1.Text = ""
                                            TextBox2.Text = ""
                                            GoTo line1
                                            Me.Close()
                                            Call LOGIN.Show()
                                        End If
                                    Else
                                        If TextBox1.Text = "USER 9" Then
                                            iret8 = MsgBox("INVALID USER OR PASSWORD", vbCritical + vbOKOnly, "LOGIN_LOGIC ERROR")
                                            If iret8 = vbOK Then
                                                file.WriteLine("PASSWORD RECOVER FAILED :" + "=" + TextBox1.Text + "_" & DateTime.Now.ToString("yyyy/MM/dd/HH:mm:ss"))
                                                file.Close()
                                                TextBox1.Text = ""
                                                TextBox2.Text = ""
                                                GoTo line1
                                                Me.Close()
                                                Call LOGIN.Show()
                                            End If
                                        Else
                                            If TextBox1.Text = "USER 10" Then
                                                iret9 = MsgBox("INVALID USER OR PASSWORD", vbCritical + vbOKOnly, "LOGIN_LOGIC ERROR")
                                                If iret9 = vbOK Then
                                                    file.WriteLine("PASSWORD RECOVER FAILED :" + "=" + TextBox1.Text + "_" & DateTime.Now.ToString("yyyy/MM/dd/HH:mm:ss"))
                                                    file.Close()
                                                    TextBox1.Text = ""
                                                    TextBox2.Text = ""
                                                    GoTo line1
                                                    Me.Close()
                                                    Call LOGIN.Show()
                                                End If
                                            Else
                                                If TextBox1.Text = "USER 11" Then
                                                    iret10 = MsgBox("INVALID USER OR PASSWORD", vbCritical + vbOKOnly, "LOGIN_LOGIC ERROR")
                                                    If iret10 = vbOK Then
                                                        file.WriteLine("PASSWORD RECOVER FAILED :" + "=" + TextBox1.Text + "_" & DateTime.Now.ToString("yyyy/MM/dd/HH:mm:ss"))
                                                        file.Close()
                                                        TextBox1.Text = ""
                                                        TextBox2.Text = ""
                                                        GoTo line1
                                                        Me.Close()
                                                        Call LOGIN.Show()
                                                    End If
                                                Else
                                                    If TextBox1.Text = "USER 12" Then
                                                        iret11 = MsgBox("INVALID USER OR PASSWORD", vbCritical + vbOKOnly, "LOGIN_LOGIC ERROR")
                                                        If iret11 = vbOK Then
                                                            file.WriteLine("PASSWORD RECOVER FAILED :" + "=" + TextBox1.Text + "_" & DateTime.Now.ToString("yyyy/MM/dd/HH:mm:ss"))
                                                            file.Close()
                                                            TextBox1.Text = ""
                                                            TextBox2.Text = ""
                                                            GoTo line1
                                                            Me.Close()
                                                            Call LOGIN.Show()
                                                        End If
                                                    Else
                                                        If TextBox1.Text = "USER 13" Then
                                                            iret12 = MsgBox("INVALID USER OR PASSWORD", vbCritical + vbOKOnly, "LOGIN_LOGIC ERROR")
                                                            If iret12 = vbOK Then
                                                                file.WriteLine("PASSWORD RECOVER FAILED :" + "=" + TextBox1.Text + "_" & DateTime.Now.ToString("yyyy/MM/dd/HH:mm:ss"))
                                                                file.Close()
                                                                TextBox1.Text = ""
                                                                TextBox2.Text = ""
                                                                GoTo line1
                                                                Me.Close()
                                                                Call LOGIN.Show()
                                                            End If
                                                        Else
                                                            If TextBox1.Text = "USER 14" Then
                                                                iret13 = MsgBox("INVALID USER OR PASSWORD", vbCritical + vbOKOnly, "LOGIN_LOGIC ERROR")
                                                                If iret13 = vbOK Then
                                                                    file.WriteLine("PASSWORD RECOVER FAILED :" + "=" + TextBox1.Text + "_" & DateTime.Now.ToString("yyyy/MM/dd/HH:mm:ss"))
                                                                    file.Close()
                                                                    TextBox1.Text = ""
                                                                    TextBox2.Text = ""
                                                                    GoTo line1
                                                                    Me.Close()
                                                                    Call LOGIN.Show()
                                                                End If
                                                            Else
                                                                If TextBox1.Text = "USER 15" Then
                                                                    iret14 = MsgBox("INVALID USER OR PASSWORD", vbCritical + vbOKOnly, "LOGIN_LOGIC ERROR")
                                                                    If iret14 = vbOK Then
                                                                        file.WriteLine("PASSWORD RECOVER FAILED :" + "=" + TextBox1.Text + "_" & DateTime.Now.ToString("yyyy/MM/dd/HH:mm:ss"))
                                                                        file.Close()
                                                                        TextBox1.Text = ""
                                                                        TextBox2.Text = ""
                                                                        GoTo line1
                                                                        Me.Close()
                                                                        Call LOGIN.Show()
                                                                    End If
                                                                Else
                                                                    If TextBox1.Text = "USER 16" Then
                                                                        iret15 = MsgBox("INVALID USER OR PASSWORD", vbCritical + vbOKOnly, "LOGIN_LOGIC ERROR")
                                                                        If iret15 = vbOK Then
                                                                            file.WriteLine("PASSWORD RECOVER FAILED :" + "=" + TextBox1.Text + "_" & DateTime.Now.ToString("yyyy/MM/dd/HH:mm:ss"))
                                                                            file.Close()
                                                                            TextBox1.Text = ""
                                                                            TextBox2.Text = ""
                                                                            GoTo line1
                                                                            Me.Close()
                                                                            Call LOGIN.Show()
                                                                        End If
                                                                    Else
                                                                        If TextBox1.Text = "USER 17" Then
                                                                            iret16 = MsgBox("INVALID USER OR PASSWORD", vbCritical + vbOKOnly, "LOGIN_LOGIC ERROR")
                                                                            If iret16 = vbOK Then
                                                                                file.WriteLine("PASSWORD RECOVER FAILED :" + "=" + TextBox1.Text + "_" & DateTime.Now.ToString("yyyy/MM/dd/HH:mm:ss"))
                                                                                file.Close()
                                                                                TextBox1.Text = ""
                                                                                TextBox2.Text = ""
                                                                                GoTo line1
                                                                                Me.Close()
                                                                                Call LOGIN.Show()
                                                                            End If
                                                                        Else
                                                                            If TextBox1.Text = "USER 18" Then
                                                                                iret17 = MsgBox("INVALID USER OR PASSWORD", vbCritical + vbOKOnly, "LOGIN_LOGIC ERROR")
                                                                                If iret17 = vbOK Then
                                                                                    file.WriteLine("PASSWORD RECOVER FAILED :" + "=" + TextBox1.Text + "_" & DateTime.Now.ToString("yyyy/MM/dd/HH:mm:ss"))
                                                                                    file.Close()
                                                                                    TextBox1.Text = ""
                                                                                    TextBox2.Text = ""
                                                                                    GoTo line1
                                                                                    Me.Close()
                                                                                    Call LOGIN.Show()
                                                                                End If
                                                                            Else
                                                                                If TextBox1.Text = "USER 19" Then
                                                                                    iret18 = MsgBox("INVALID USER OR PASSWORD", vbCritical + vbOKOnly, "LOGIN_LOGIC ERROR")
                                                                                    If iret18 = vbOK Then
                                                                                        file.WriteLine("PASSWORD RECOVER FAILED :" + "=" + TextBox1.Text + "_" & DateTime.Now.ToString("yyyy/MM/dd/HH:mm:ss"))
                                                                                        file.Close()
                                                                                        TextBox1.Text = ""
                                                                                        TextBox2.Text = ""
                                                                                        GoTo line1
                                                                                        Me.Close()
                                                                                        Call LOGIN.Show()
                                                                                    End If
                                                                                Else
                                                                                    If TextBox1.Text = "USER 20" Then
                                                                                        iret19 = MsgBox("INVALID USER OR PASSWORD", vbCritical + vbOKOnly, "LOGIN_LOGIC ERROR")
                                                                                        If iret19 = vbOK Then
                                                                                            file.WriteLine("PASSWORD RECOVER FAILED :" + "=" + TextBox1.Text + "_" & DateTime.Now.ToString("yyyy/MM/dd/HH:mm:ss"))
                                                                                            file.Close()
                                                                                            TextBox1.Text = ""
                                                                                            TextBox2.Text = ""
                                                                                            GoTo line1
                                                                                            Me.Close()
                                                                                            Call LOGIN.Show()
                                                                                        End If
                                                                                    Else
                                                                                        If TextBox1.Text = "USER 21" Then
                                                                                            iret20 = MsgBox("INVALID USER OR PASSWORD", vbCritical + vbOKOnly, "LOGIN_LOGIC ERROR")
                                                                                            If iret20 = vbOK Then
                                                                                                file.WriteLine("PASSWORD RECOVER FAILED :" + "=" + TextBox1.Text + "_" & DateTime.Now.ToString("yyyy/MM/dd/HH:mm:ss"))
                                                                                                file.Close()
                                                                                                TextBox1.Text = ""
                                                                                                TextBox2.Text = ""
                                                                                                GoTo line1
                                                                                                Me.Close()
                                                                                                Call LOGIN.Show()
                                                                                            End If
                                                                                        Else
                                                                                            If TextBox1.Text = "USER 22" Then
                                                                                                iret21 = MsgBox("INVALID USER OR PASSWORD", vbCritical + vbOKOnly, "LOGIN_LOGIC ERROR")
                                                                                                If iret21 = vbOK Then
                                                                                                    file.WriteLine("PASSWORD RECOVER FAILED :" + "=" + TextBox1.Text + "_" & DateTime.Now.ToString("yyyy/MM/dd/HH:mm:ss"))
                                                                                                    file.Close()
                                                                                                    TextBox1.Text = ""
                                                                                                    TextBox2.Text = ""
                                                                                                    GoTo line1
                                                                                                    Me.Close()
                                                                                                    Call LOGIN.Show()
                                                                                                End If
                                                                                            Else
                                                                                                If TextBox1.Text = "USER 23" Then
                                                                                                    iret22 = MsgBox("INVALID USER OR PASSWORD", vbCritical + vbOKOnly, "LOGIN_LOGIC ERROR")
                                                                                                    If iret22 = vbOK Then
                                                                                                        file.WriteLine("PASSWORD RECOVER FAILED :" + "=" + TextBox1.Text + "_" & DateTime.Now.ToString("yyyy/MM/dd/HH:mm:ss"))
                                                                                                        file.Close()
                                                                                                        TextBox1.Text = ""
                                                                                                        TextBox2.Text = ""
                                                                                                        GoTo line1
                                                                                                        Me.Close()
                                                                                                        Call LOGIN.Show()
                                                                                                    End If
                                                                                                Else
                                                                                                    If TextBox1.Text = "USER 24" Then
                                                                                                        iret23 = MsgBox("INVALID USER OR PASSWORD", vbCritical + vbOKOnly, "LOGIN_LOGIC ERROR")
                                                                                                        If iret23 = vbOK Then
                                                                                                            file.WriteLine("PASSWORD RECOVER FAILED :" + "=" + TextBox1.Text + "_" & DateTime.Now.ToString("yyyy/MM/dd/HH:mm:ss"))
                                                                                                            file.Close()
                                                                                                            TextBox1.Text = ""
                                                                                                            TextBox2.Text = ""
                                                                                                            GoTo line1
                                                                                                            Me.Close()
                                                                                                            Call LOGIN.Show()
                                                                                                        End If
                                                                                                    Else
                                                                                                        If TextBox1.Text = "USER 25" Then
                                                                                                            iret24 = MsgBox("INVALID USER OR PASSWORD", vbCritical + vbOKOnly, "LOGIN_LOGIC ERROR")
                                                                                                            If iret24 = vbOK Then
                                                                                                                file.WriteLine("PASSWORD RECOVER FAILED :" + "=" + TextBox1.Text + "_" & DateTime.Now.ToString("yyyy/MM/dd/HH:mm:ss"))
                                                                                                                file.Close()
                                                                                                                TextBox1.Text = ""
                                                                                                                TextBox2.Text = ""
                                                                                                                GoTo line1
                                                                                                                Me.Close()
                                                                                                                Call LOGIN.Show()
                                                                                                            End If
                                                                                                        Else
                                                                                                            If TextBox1.Text = "USER 26" Then
                                                                                                                iret25 = MsgBox("INVALID USER OR PASSWORD", vbCritical + vbOKOnly, "LOGIN_LOGIC ERROR")
                                                                                                                If iret25 = vbOK Then
                                                                                                                    file.WriteLine("PASSWORD RECOVER FAILED :" + "=" + TextBox1.Text + "_" & DateTime.Now.ToString("yyyy/MM/dd/HH:mm:ss"))
                                                                                                                    file.Close()
                                                                                                                    TextBox1.Text = ""
                                                                                                                    TextBox2.Text = ""
                                                                                                                    GoTo line1
                                                                                                                    Me.Close()
                                                                                                                    Call LOGIN.Show()
                                                                                                                End If
                                                                                                            Else
                                                                                                                If TextBox1.Text = "USER 27" Then
                                                                                                                    iret26 = MsgBox("INVALID USER OR PASSWORD", vbCritical + vbOKOnly, "LOGIN_LOGIC ERROR")
                                                                                                                    If iret26 = vbOK Then
                                                                                                                        file.WriteLine("PASSWORD RECOVER FAILED :" + "=" + TextBox1.Text + "_" & DateTime.Now.ToString("yyyy/MM/dd/HH:mm:ss"))
                                                                                                                        file.Close()
                                                                                                                        TextBox1.Text = ""
                                                                                                                        TextBox2.Text = ""
                                                                                                                        GoTo line1
                                                                                                                        Me.Close()
                                                                                                                        Call LOGIN.Show()
                                                                                                                    End If
                                                                                                                Else
                                                                                                                    If TextBox1.Text = "USER 28" Then
                                                                                                                        iret27 = MsgBox("INVALID USER OR PASSWORD", vbCritical + vbOKOnly, "LOGIN_LOGIC ERROR")
                                                                                                                        If iret27 = vbOK Then
                                                                                                                            file.WriteLine("PASSWORD RECOVER FAILED :" + "=" + TextBox1.Text + "_" & DateTime.Now.ToString("yyyy/MM/dd/HH:mm:ss"))
                                                                                                                            file.Close()
                                                                                                                            TextBox1.Text = ""
                                                                                                                            TextBox2.Text = ""
                                                                                                                            GoTo line1
                                                                                                                            Me.Close()
                                                                                                                            Call LOGIN.Show()
                                                                                                                        End If
                                                                                                                    Else
                                                                                                                        If TextBox1.Text = "USER 29" Then
                                                                                                                            iret28 = MsgBox("INVALID USER OR PASSWORD", vbCritical + vbOKOnly, "LOGIN_LOGIC ERROR")
                                                                                                                            If iret28 = vbOK Then
                                                                                                                                file.WriteLine("PASSWORD RECOVER FAILED :" + "=" + TextBox1.Text + "_" & DateTime.Now.ToString("yyyy/MM/dd/HH:mm:ss"))
                                                                                                                                file.Close()
                                                                                                                                TextBox1.Text = ""
                                                                                                                                TextBox2.Text = ""
                                                                                                                                GoTo line1
                                                                                                                                Me.Close()
                                                                                                                                Call LOGIN.Show()
                                                                                                                            End If
                                                                                                                        Else
                                                                                                                            If TextBox1.Text = "USER 30" Then
                                                                                                                                iret29 = MsgBox("INVALID USER OR PASSWORD", vbCritical + vbOKOnly, "LOGIN_LOGIC ERROR")
                                                                                                                                If iret29 = vbOK Then
                                                                                                                                    file.WriteLine("PASSWORD RECOVER FAILED :" + "=" + TextBox1.Text + "_" & DateTime.Now.ToString("yyyy/MM/dd/HH:mm:ss"))
                                                                                                                                    file.Close()
                                                                                                                                    TextBox1.Text = ""
                                                                                                                                    TextBox2.Text = ""
                                                                                                                                    GoTo line1
                                                                                                                                    Me.Close()
                                                                                                                                    Call LOGIN.Show()
                                                                                                                                End If
                                                                                                                            Else
                                                                                                                                If TextBox1.Text = st.range("A2").value And TextBox2.Text = st.range("B2").value Then
                                                                                                                                    iret30 = MsgBox("THE PASSWORD IS" & st.range("C2").value, vbInformation + vbOKOnly, "LOGIN_LOGIC")
                                                                                                                                    If iret30 = vbOK Then
                                                                                                                                        file.WriteLine("PASSWORD RECOVER SUCCESSFULLY :" + "=" + TextBox1.Text + "_" & DateTime.Now.ToString("yyyy/MM/dd/HH:mm:ss"))
                                                                                                                                        file.Close()
                                                                                                                                        TextBox1.Text = ""
                                                                                                                                        TextBox2.Text = ""
                                                                                                                                        GoTo line1
                                                                                                                                        Me.Close()
                                                                                                                                        Call LOGIN.Show()
                                                                                                                                    End If
                                                                                                                                Else
                                                                                                                                    If TextBox1.Text = st.range("A3").value And TextBox2.Text = st.range("B3").value Then
                                                                                                                                        iret31 = MsgBox("THE PASSWORD IS" & st.range("C3").value, vbInformation + vbOKOnly, "LOGIN_LOGIC")
                                                                                                                                        If iret31 = vbOK Then
                                                                                                                                            file.WriteLine("PASSWORD RECOVER SUCCESSFULLY :" + "=" + TextBox1.Text + "_" & DateTime.Now.ToString("yyyy/MM/dd/HH:mm:ss"))
                                                                                                                                            file.Close()
                                                                                                                                            TextBox1.Text = ""
                                                                                                                                            TextBox2.Text = ""
                                                                                                                                            GoTo line1
                                                                                                                                            Me.Close()
                                                                                                                                            Call LOGIN.Show()
                                                                                                                                        End If
                                                                                                                                    Else
                                                                                                                                        If TextBox1.Text = st.range("A4").value And TextBox2.Text = st.range("B4").value Then
                                                                                                                                            iret32 = MsgBox("THE PASSWORD IS" & st.range("C4").value, vbInformation + vbOKOnly, "LOGIN_LOGIC")
                                                                                                                                            If iret32 = vbOK Then
                                                                                                                                                file.WriteLine("PASSWORD RECOVER SUCCESSFULLY :" + "=" + TextBox1.Text + "_" & DateTime.Now.ToString("yyyy/MM/dd/HH:mm:ss"))
                                                                                                                                                file.Close()
                                                                                                                                                TextBox1.Text = ""
                                                                                                                                                TextBox2.Text = ""
                                                                                                                                                GoTo line1
                                                                                                                                                Me.Close()
                                                                                                                                                Call LOGIN.Show()
                                                                                                                                            End If
                                                                                                                                        Else
                                                                                                                                            If TextBox1.Text = st.range("A5").value And TextBox2.Text = st.range("B5").value Then
                                                                                                                                                iret33 = MsgBox("THE PASSWORD IS" & st.range("C5").value, vbInformation + vbOKOnly, "LOGIN_LOGIC")
                                                                                                                                                If iret33 = vbOK Then
                                                                                                                                                    file.WriteLine("PASSWORD RECOVER SUCCESSFULLY :" + "=" + TextBox1.Text + "_" & DateTime.Now.ToString("yyyy/MM/dd/HH:mm:ss"))
                                                                                                                                                    file.Close()
                                                                                                                                                    TextBox1.Text = ""
                                                                                                                                                    TextBox2.Text = ""
                                                                                                                                                    GoTo line1
                                                                                                                                                    Me.Close()
                                                                                                                                                    Call LOGIN.Show()
                                                                                                                                                End If
                                                                                                                                            Else
                                                                                                                                                If TextBox1.Text = st.range("A6").value And TextBox2.Text = st.range("B6").value Then
                                                                                                                                                    iret34 = MsgBox("THE PASSWORD IS" & st.range("C6").value, vbInformation + vbOKOnly, "LOGIN_LOGIC")
                                                                                                                                                    If iret34 = vbOK Then
                                                                                                                                                        file.WriteLine("PASSWORD RECOVER SUCCESSFULLY :" + "=" + TextBox1.Text + "_" & DateTime.Now.ToString("yyyy/MM/dd/HH:mm:ss"))
                                                                                                                                                        file.Close()
                                                                                                                                                        TextBox1.Text = ""
                                                                                                                                                        TextBox2.Text = ""
                                                                                                                                                        GoTo line1
                                                                                                                                                        Me.Close()
                                                                                                                                                        Call LOGIN.Show()
                                                                                                                                                    End If
                                                                                                                                                Else
                                                                                                                                                    If TextBox1.Text = st.range("A7").value And TextBox2.Text = st.range("B7").value Then
                                                                                                                                                        iret35 = MsgBox("THE PASSWORD IS" & st.range("C7").value, vbInformation + vbOKOnly, "LOGIN_LOGIC")
                                                                                                                                                        If iret35 = vbOK Then
                                                                                                                                                            file.WriteLine("PASSWORD RECOVER SUCCESSFULLY :" + "=" + TextBox1.Text + "_" & DateTime.Now.ToString("yyyy/MM/dd/HH:mm:ss"))
                                                                                                                                                            file.Close()
                                                                                                                                                            TextBox1.Text = ""
                                                                                                                                                            TextBox2.Text = ""
                                                                                                                                                            GoTo line1
                                                                                                                                                            Me.Close()
                                                                                                                                                            Call LOGIN.Show()
                                                                                                                                                        End If
                                                                                                                                                    Else
                                                                                                                                                        If TextBox1.Text = st.range("A8").value And TextBox2.Text = st.range("B8").value Then
                                                                                                                                                            iret36 = MsgBox("THE PASSWORD IS" & st.range("C8").value, vbInformation + vbOKOnly, "LOGIN_LOGIC")
                                                                                                                                                            If iret36 = vbOK Then
                                                                                                                                                                file.WriteLine("PASSWORD RECOVER SUCCESSFULLY :" + "=" + TextBox1.Text + "_" & DateTime.Now.ToString("yyyy/MM/dd/HH:mm:ss"))
                                                                                                                                                                file.Close()
                                                                                                                                                                TextBox1.Text = ""
                                                                                                                                                                TextBox2.Text = ""
                                                                                                                                                                GoTo line1
                                                                                                                                                                Me.Close()
                                                                                                                                                                Call LOGIN.Show()
                                                                                                                                                            End If
                                                                                                                                                        Else
                                                                                                                                                            If TextBox1.Text = st.range("A9").value And TextBox2.Text = st.range("B9").value Then
                                                                                                                                                                iret37 = MsgBox("THE PASSWORD IS" & st.range("C9").value, vbInformation + vbOKOnly, "LOGIN_LOGIC")
                                                                                                                                                                If iret37 = vbOK Then
                                                                                                                                                                    file.WriteLine("PASSWORD RECOVER SUCCESSFULLY :" + "=" + TextBox1.Text + "_" & DateTime.Now.ToString("yyyy/MM/dd/HH:mm:ss"))
                                                                                                                                                                    file.Close()
                                                                                                                                                                    TextBox1.Text = ""
                                                                                                                                                                    TextBox2.Text = ""
                                                                                                                                                                    GoTo line1
                                                                                                                                                                    Me.Close()
                                                                                                                                                                    Call LOGIN.Show()
                                                                                                                                                                End If
                                                                                                                                                            Else
                                                                                                                                                                If TextBox1.Text = st.range("A10").value And TextBox2.Text = st.range("B10").value Then
                                                                                                                                                                    iret38 = MsgBox("THE PASSWORD IS" & st.range("C10").value, vbInformation + vbOKOnly, "LOGIN_LOGIC")
                                                                                                                                                                    If iret38 = vbOK Then
                                                                                                                                                                        file.WriteLine("PASSWORD RECOVER SUCCESSFULLY :" + "=" + TextBox1.Text + "_" & DateTime.Now.ToString("yyyy/MM/dd/HH:mm:ss"))
                                                                                                                                                                        file.Close()
                                                                                                                                                                        TextBox1.Text = ""
                                                                                                                                                                        TextBox2.Text = ""
                                                                                                                                                                        GoTo line1
                                                                                                                                                                        Me.Close()
                                                                                                                                                                        Call LOGIN.Show()
                                                                                                                                                                    End If
                                                                                                                                                                Else
                                                                                                                                                                    If TextBox1.Text = st.range("A11").value And TextBox2.Text = st.range("B11").value Then
                                                                                                                                                                        iret39 = MsgBox("THE PASSWORD IS" & st.range("C11").value, vbInformation + vbOKOnly, "LOGIN_LOGIC")
                                                                                                                                                                        If iret39 = vbOK Then
                                                                                                                                                                            file.WriteLine("PASSWORD RECOVER SUCCESSFULLY :" + "=" + TextBox1.Text + "_" & DateTime.Now.ToString("yyyy/MM/dd/HH:mm:ss"))
                                                                                                                                                                            file.Close()
                                                                                                                                                                            TextBox1.Text = ""
                                                                                                                                                                            TextBox2.Text = ""
                                                                                                                                                                            GoTo line1
                                                                                                                                                                            Me.Close()
                                                                                                                                                                            Call LOGIN.Show()
                                                                                                                                                                        End If
                                                                                                                                                                    Else
                                                                                                                                                                        If TextBox1.Text = st.range("A12").value And TextBox2.Text = st.range("B12").value Then
                                                                                                                                                                            iret40 = MsgBox("THE PASSWORD IS" & st.range("C12").value, vbInformation + vbOKOnly, "LOGIN_LOGIC")
                                                                                                                                                                            If iret40 = vbOK Then
                                                                                                                                                                                file.WriteLine("PASSWORD RECOVER SUCCESSFULLY :" + "=" + TextBox1.Text + "_" & DateTime.Now.ToString("yyyy/MM/dd/HH:mm:ss"))
                                                                                                                                                                                file.Close()
                                                                                                                                                                                TextBox1.Text = ""
                                                                                                                                                                                TextBox2.Text = ""
                                                                                                                                                                                GoTo line1
                                                                                                                                                                                Me.Close()
                                                                                                                                                                                Call LOGIN.Show()
                                                                                                                                                                            End If
                                                                                                                                                                        Else
                                                                                                                                                                            If TextBox1.Text = st.range("A13").value And TextBox2.Text = st.range("B13").value Then
                                                                                                                                                                                iret41 = MsgBox("THE PASSWORD IS" & st.range("C13").value, vbInformation + vbOKOnly, "LOGIN_LOGIC")
                                                                                                                                                                                If iret41 = vbOK Then
                                                                                                                                                                                    file.WriteLine("PASSWORD RECOVER SUCCESSFULLY :" + "=" + TextBox1.Text + "_" & DateTime.Now.ToString("yyyy/MM/dd/HH:mm:ss"))
                                                                                                                                                                                    file.Close()
                                                                                                                                                                                    TextBox1.Text = ""
                                                                                                                                                                                    TextBox2.Text = ""
                                                                                                                                                                                    GoTo line1
                                                                                                                                                                                    Me.Close()
                                                                                                                                                                                    Call LOGIN.Show()
                                                                                                                                                                                End If
                                                                                                                                                                            Else
                                                                                                                                                                                If TextBox1.Text = st.range("A14").value And TextBox2.Text = st.range("B14").value Then
                                                                                                                                                                                    iret42 = MsgBox("THE PASSWORD IS" & st.range("C14").value, vbInformation + vbOKOnly, "LOGIN_LOGIC")
                                                                                                                                                                                    If iret42 = vbOK Then
                                                                                                                                                                                        file.WriteLine("PASSWORD RECOVER SUCCESSFULLY :" + "=" + TextBox1.Text + "_" & DateTime.Now.ToString("yyyy/MM/dd/HH:mm:ss"))
                                                                                                                                                                                        file.Close()
                                                                                                                                                                                        TextBox1.Text = ""
                                                                                                                                                                                        TextBox2.Text = ""
                                                                                                                                                                                        GoTo line1
                                                                                                                                                                                        Me.Close()
                                                                                                                                                                                        Call LOGIN.Show()
                                                                                                                                                                                    End If
                                                                                                                                                                                Else
                                                                                                                                                                                    If TextBox1.Text = st.range("A15").value And TextBox2.Text = st.range("B15").value Then
                                                                                                                                                                                        iret43 = MsgBox("THE PASSWORD IS" & st.range("C15").value, vbInformation + vbOKOnly, "LOGIN_LOGIC")
                                                                                                                                                                                        If iret43 = vbOK Then
                                                                                                                                                                                            file.WriteLine("PASSWORD RECOVER SUCCESSFULLY :" + "=" + TextBox1.Text + "_" & DateTime.Now.ToString("yyyy/MM/dd/HH:mm:ss"))
                                                                                                                                                                                            file.Close()
                                                                                                                                                                                            TextBox1.Text = ""
                                                                                                                                                                                            TextBox2.Text = ""
                                                                                                                                                                                            GoTo line1
                                                                                                                                                                                            Me.Close()
                                                                                                                                                                                            Call LOGIN.Show()
                                                                                                                                                                                        End If
                                                                                                                                                                                    Else
                                                                                                                                                                                        If TextBox1.Text = st.range("A16").value And TextBox2.Text = st.range("B16").value Then
                                                                                                                                                                                            iret44 = MsgBox("THE PASSWORD IS" & st.range("C16").value, vbInformation + vbOKOnly, "LOGIN_LOGIC")
                                                                                                                                                                                            If iret44 = vbOK Then
                                                                                                                                                                                                file.WriteLine("PASSWORD RECOVER SUCCESSFULLY :" + "=" + TextBox1.Text + "_" & DateTime.Now.ToString("yyyy/MM/dd/HH:mm:ss"))
                                                                                                                                                                                                file.Close()
                                                                                                                                                                                                TextBox1.Text = ""
                                                                                                                                                                                                TextBox2.Text = ""
                                                                                                                                                                                                GoTo line1
                                                                                                                                                                                                Me.Close()
                                                                                                                                                                                                Call LOGIN.Show()
                                                                                                                                                                                            End If
                                                                                                                                                                                        Else
                                                                                                                                                                                            If TextBox1.Text = st.range("A17").value And TextBox2.Text = st.range("B17").value Then
                                                                                                                                                                                                iret45 = MsgBox("THE PASSWORD IS" & st.range("C17").value, vbInformation + vbOKOnly, "LOGIN_LOGIC")
                                                                                                                                                                                                If iret45 = vbOK Then
                                                                                                                                                                                                    file.WriteLine("PASSWORD RECOVER SUCCESSFULLY :" + "=" + TextBox1.Text + "_" & DateTime.Now.ToString("yyyy/MM/dd/HH:mm:ss"))
                                                                                                                                                                                                    file.Close()
                                                                                                                                                                                                    TextBox1.Text = ""
                                                                                                                                                                                                    TextBox2.Text = ""
                                                                                                                                                                                                    GoTo line1
                                                                                                                                                                                                    Me.Close()
                                                                                                                                                                                                    Call LOGIN.Show()
                                                                                                                                                                                                End If
                                                                                                                                                                                            Else
                                                                                                                                                                                                If TextBox1.Text = st.range("A18").value And TextBox2.Text = st.range("B18").value Then
                                                                                                                                                                                                    iret46 = MsgBox("THE PASSWORD IS" & st.range("C18").value, vbInformation + vbOKOnly, "LOGIN_LOGIC")
                                                                                                                                                                                                    If iret46 = vbOK Then
                                                                                                                                                                                                        file.WriteLine("PASSWORD RECOVER SUCCESSFULLY :" + "=" + TextBox1.Text + "_" & DateTime.Now.ToString("yyyy/MM/dd/HH:mm:ss"))
                                                                                                                                                                                                        file.Close()
                                                                                                                                                                                                        TextBox1.Text = ""
                                                                                                                                                                                                        TextBox2.Text = ""
                                                                                                                                                                                                        GoTo line1
                                                                                                                                                                                                        Me.Close()
                                                                                                                                                                                                        Call LOGIN.Show()
                                                                                                                                                                                                    End If
                                                                                                                                                                                                Else
                                                                                                                                                                                                    If TextBox1.Text = st.range("A19").value And TextBox2.Text = st.range("B19").value Then
                                                                                                                                                                                                        iret47 = MsgBox("THE PASSWORD IS" & st.range("C19").value, vbInformation + vbOKOnly, "LOGIN_LOGIC")
                                                                                                                                                                                                        If iret47 = vbOK Then
                                                                                                                                                                                                            file.WriteLine("PASSWORD RECOVER SUCCESSFULLY :" + "=" + TextBox1.Text + "_" & DateTime.Now.ToString("yyyy/MM/dd/HH:mm:ss"))
                                                                                                                                                                                                            file.Close()
                                                                                                                                                                                                            TextBox1.Text = ""
                                                                                                                                                                                                            TextBox2.Text = ""
                                                                                                                                                                                                            GoTo line1
                                                                                                                                                                                                            Me.Close()
                                                                                                                                                                                                            Call LOGIN.Show()
                                                                                                                                                                                                        End If
                                                                                                                                                                                                    Else
                                                                                                                                                                                                        If TextBox1.Text = st.range("A20").value And TextBox2.Text = st.range("B20").value Then
                                                                                                                                                                                                            iret48 = MsgBox("THE PASSWORD IS" & st.range("C20").value, vbInformation + vbOKOnly, "LOGIN_LOGIC")
                                                                                                                                                                                                            If iret48 = vbOK Then
                                                                                                                                                                                                                file.WriteLine("PASSWORD RECOVER SUCCESSFULLY :" + "=" + TextBox1.Text + "_" & DateTime.Now.ToString("yyyy/MM/dd/HH:mm:ss"))
                                                                                                                                                                                                                file.Close()
                                                                                                                                                                                                                TextBox1.Text = ""
                                                                                                                                                                                                                TextBox2.Text = ""
                                                                                                                                                                                                                GoTo line1
                                                                                                                                                                                                                Me.Close()
                                                                                                                                                                                                                Call LOGIN.Show()
                                                                                                                                                                                                            End If
                                                                                                                                                                                                        Else
                                                                                                                                                                                                            If TextBox1.Text = st.range("A21").value And TextBox2.Text = st.range("B21").value Then
                                                                                                                                                                                                                iret49 = MsgBox("THE PASSWORD IS" & st.range("C21").value, vbInformation + vbOKOnly, "LOGIN_LOGIC")
                                                                                                                                                                                                                If iret49 = vbOK Then
                                                                                                                                                                                                                    file.WriteLine("PASSWORD RECOVER SUCCESSFULLY :" + "=" + TextBox1.Text + "_" & DateTime.Now.ToString("yyyy/MM/dd/HH:mm:ss"))
                                                                                                                                                                                                                    file.Close()
                                                                                                                                                                                                                    TextBox1.Text = ""
                                                                                                                                                                                                                    TextBox2.Text = ""
                                                                                                                                                                                                                    GoTo line1
                                                                                                                                                                                                                    Me.Close()
                                                                                                                                                                                                                    Call LOGIN.Show()
                                                                                                                                                                                                                End If
                                                                                                                                                                                                            Else
                                                                                                                                                                                                                If TextBox1.Text = st.range("A22").value And TextBox2.Text = st.range("B22").value Then
                                                                                                                                                                                                                    iret50 = MsgBox("THE PASSWORD IS" & st.range("C22").value, vbInformation + vbOKOnly, "LOGIN_LOGIC")
                                                                                                                                                                                                                    If iret50 = vbOK Then
                                                                                                                                                                                                                        file.WriteLine("PASSWORD RECOVER SUCCESSFULLY :" + "=" + TextBox1.Text + "_" & DateTime.Now.ToString("yyyy/MM/dd/HH:mm:ss"))
                                                                                                                                                                                                                        file.Close()
                                                                                                                                                                                                                        TextBox1.Text = ""
                                                                                                                                                                                                                        TextBox2.Text = ""
                                                                                                                                                                                                                        GoTo line1
                                                                                                                                                                                                                        Me.Close()
                                                                                                                                                                                                                        Call LOGIN.Show()
                                                                                                                                                                                                                    End If
                                                                                                                                                                                                                Else
                                                                                                                                                                                                                    If TextBox1.Text = st.range("A23").value And TextBox2.Text = st.range("B23").value Then
                                                                                                                                                                                                                        iret51 = MsgBox("THE PASSWORD IS" & st.range("C23").value, vbInformation + vbOKOnly, "LOGIN_LOGIC")
                                                                                                                                                                                                                        If iret51 = vbOK Then
                                                                                                                                                                                                                            file.WriteLine("PASSWORD RECOVER SUCCESSFULLY :" + "=" + TextBox1.Text + "_" & DateTime.Now.ToString("yyyy/MM/dd/HH:mm:ss"))
                                                                                                                                                                                                                            file.Close()
                                                                                                                                                                                                                            TextBox1.Text = ""
                                                                                                                                                                                                                            TextBox2.Text = ""
                                                                                                                                                                                                                            GoTo line1
                                                                                                                                                                                                                            Me.Close()
                                                                                                                                                                                                                            Call LOGIN.Show()
                                                                                                                                                                                                                        End If
                                                                                                                                                                                                                    Else
                                                                                                                                                                                                                        If TextBox1.Text = st.range("A24").value And TextBox2.Text = st.range("B24").value Then
                                                                                                                                                                                                                            iret52 = MsgBox("THE PASSWORD IS" & st.range("C24").value, vbInformation + vbOKOnly, "LOGIN_LOGIC")
                                                                                                                                                                                                                            If iret52 = vbOK Then
                                                                                                                                                                                                                                file.WriteLine("PASSWORD RECOVER SUCCESSFULLY :" + "=" + TextBox1.Text + "_" & DateTime.Now.ToString("yyyy/MM/dd/HH:mm:ss"))
                                                                                                                                                                                                                                file.Close()
                                                                                                                                                                                                                                TextBox1.Text = ""
                                                                                                                                                                                                                                TextBox2.Text = ""
                                                                                                                                                                                                                                GoTo line1
                                                                                                                                                                                                                                Me.Close()
                                                                                                                                                                                                                                Call LOGIN.Show()
                                                                                                                                                                                                                            End If
                                                                                                                                                                                                                        Else
                                                                                                                                                                                                                            If TextBox1.Text = st.range("A25").value And TextBox2.Text = st.range("B25").value Then
                                                                                                                                                                                                                                iret53 = MsgBox("THE PASSWORD IS" & st.range("C25").value, vbInformation + vbOKOnly, "LOGIN_LOGIC")
                                                                                                                                                                                                                                If iret53 = vbOK Then
                                                                                                                                                                                                                                    file.WriteLine("PASSWORD RECOVER SUCCESSFULLY :" + "=" + TextBox1.Text + "_" & DateTime.Now.ToString("yyyy/MM/dd/HH:mm:ss"))
                                                                                                                                                                                                                                    file.Close()
                                                                                                                                                                                                                                    TextBox1.Text = ""
                                                                                                                                                                                                                                    TextBox2.Text = ""
                                                                                                                                                                                                                                    GoTo line1
                                                                                                                                                                                                                                    Me.Close()
                                                                                                                                                                                                                                    Call LOGIN.Show()
                                                                                                                                                                                                                                End If
                                                                                                                                                                                                                            Else
                                                                                                                                                                                                                                If TextBox1.Text = st.range("A26").value And TextBox2.Text = st.range("B26").value Then
                                                                                                                                                                                                                                    iret54 = MsgBox("THE PASSWORD IS" & st.range("C26").value, vbInformation + vbOKOnly, "LOGIN_LOGIC")
                                                                                                                                                                                                                                    If iret54 = vbOK Then
                                                                                                                                                                                                                                        file.WriteLine("PASSWORD RECOVER SUCCESSFULLY :" + "=" + TextBox1.Text + "_" & DateTime.Now.ToString("yyyy/MM/dd/HH:mm:ss"))
                                                                                                                                                                                                                                        file.Close()
                                                                                                                                                                                                                                        TextBox1.Text = ""
                                                                                                                                                                                                                                        TextBox2.Text = ""
                                                                                                                                                                                                                                        GoTo line1
                                                                                                                                                                                                                                        Me.Close()
                                                                                                                                                                                                                                        Call LOGIN.Show()
                                                                                                                                                                                                                                    End If
                                                                                                                                                                                                                                Else
                                                                                                                                                                                                                                    If TextBox1.Text = st.range("A27").value And TextBox2.Text = st.range("B27").value Then
                                                                                                                                                                                                                                        iret55 = MsgBox("THE PASSWORD IS" & st.range("C27").value, vbInformation + vbOKOnly, "LOGIN_LOGIC")
                                                                                                                                                                                                                                        If iret55 = vbOK Then
                                                                                                                                                                                                                                            file.WriteLine("PASSWORD RECOVER SUCCESSFULLY :" + "=" + TextBox1.Text + "_" & DateTime.Now.ToString("yyyy/MM/dd/HH:mm:ss"))
                                                                                                                                                                                                                                            file.Close()
                                                                                                                                                                                                                                            TextBox1.Text = ""
                                                                                                                                                                                                                                            TextBox2.Text = ""
                                                                                                                                                                                                                                            GoTo line1
                                                                                                                                                                                                                                            Me.Close()
                                                                                                                                                                                                                                            Call LOGIN.Show()
                                                                                                                                                                                                                                        End If
                                                                                                                                                                                                                                    Else
                                                                                                                                                                                                                                        If TextBox1.Text = st.range("A28").value And TextBox2.Text = st.range("B28").value Then
                                                                                                                                                                                                                                            iret56 = MsgBox("THE PASSWORD IS" & st.range("C28").value, vbInformation + vbOKOnly, "LOGIN_LOGIC")
                                                                                                                                                                                                                                            If iret56 = vbOK Then
                                                                                                                                                                                                                                                file.WriteLine("PASSWORD RECOVER SUCCESSFULLY :" + "=" + TextBox1.Text + "_" & DateTime.Now.ToString("yyyy/MM/dd/HH:mm:ss"))
                                                                                                                                                                                                                                                file.Close()
                                                                                                                                                                                                                                                TextBox1.Text = ""
                                                                                                                                                                                                                                                TextBox2.Text = ""
                                                                                                                                                                                                                                                GoTo line1
                                                                                                                                                                                                                                                Me.Close()
                                                                                                                                                                                                                                                Call LOGIN.Show()
                                                                                                                                                                                                                                            End If
                                                                                                                                                                                                                                        Else
                                                                                                                                                                                                                                            If TextBox1.Text = st.range("A29").value And TextBox2.Text = st.range("B29").value Then
                                                                                                                                                                                                                                                iret57 = MsgBox("THE PASSWORD IS" & st.range("C29").value, vbInformation + vbOKOnly, "LOGIN_LOGIC")
                                                                                                                                                                                                                                                If iret57 = vbOK Then
                                                                                                                                                                                                                                                    file.WriteLine("PASSWORD RECOVER SUCCESSFULLY :" + "=" + TextBox1.Text + "_" & DateTime.Now.ToString("yyyy/MM/dd/HH:mm:ss"))
                                                                                                                                                                                                                                                    file.Close()
                                                                                                                                                                                                                                                    TextBox1.Text = ""
                                                                                                                                                                                                                                                    TextBox2.Text = ""
                                                                                                                                                                                                                                                    GoTo line1
                                                                                                                                                                                                                                                    Me.Close()
                                                                                                                                                                                                                                                    Call LOGIN.Show()
                                                                                                                                                                                                                                                End If
                                                                                                                                                                                                                                            Else
                                                                                                                                                                                                                                                If TextBox1.Text = st.range("A30").value And TextBox2.Text = st.range("B30").value Then
                                                                                                                                                                                                                                                    iret58 = MsgBox("THE PASSWORD IS" & st.range("C30").value, vbInformation + vbOKOnly, "LOGIN_LOGIC")
                                                                                                                                                                                                                                                    If iret58 = vbOK Then
                                                                                                                                                                                                                                                        file.WriteLine("PASSWORD RECOVER SUCCESSFULLY :" + "=" + TextBox1.Text + "_" & DateTime.Now.ToString("yyyy/MM/dd/HH:mm:ss"))
                                                                                                                                                                                                                                                        file.Close()
                                                                                                                                                                                                                                                        TextBox1.Text = ""
                                                                                                                                                                                                                                                        TextBox2.Text = ""
                                                                                                                                                                                                                                                        GoTo line1
                                                                                                                                                                                                                                                        Me.Close()
                                                                                                                                                                                                                                                        Call LOGIN.Show()
                                                                                                                                                                                                                                                    End If
                                                                                                                                                                                                                                                Else
                                                                                                                                                                                                                                                    If TextBox1.Text = st.range("A31").value And TextBox2.Text = st.range("B31").value Then
                                                                                                                                                                                                                                                        iret59 = MsgBox("THE PASSWORD IS" & st.range("C31").value, vbInformation + vbOKOnly, "LOGIN_LOGIC")
                                                                                                                                                                                                                                                        If iret59 = vbOK Then
                                                                                                                                                                                                                                                            file.WriteLine("PASSWORD RECOVER SUCCESSFULLY :" + "=" + TextBox1.Text + "_" & DateTime.Now.ToString("yyyy/MM/dd/HH:mm:ss"))
                                                                                                                                                                                                                                                            file.Close()
                                                                                                                                                                                                                                                            TextBox1.Text = ""
                                                                                                                                                                                                                                                            TextBox2.Text = ""
                                                                                                                                                                                                                                                            GoTo line1
                                                                                                                                                                                                                                                            Me.Close()
                                                                                                                                                                                                                                                            Call LOGIN.Show()
                                                                                                                                                                                                                                                        End If
                                                                                                                                                                                                                                                    Else
                                                                                                                                                                                                                                                        iret60 = MsgBox("USER NAME OR USER ID NOT MATCHED", vbCritical + vbOKOnly, "LOGIN_LOGIC ERROR")
                                                                                                                                                                                                                                                        If iret60 = vbOK Then
                                                                                                                                                                                                                                                            file.WriteLine("PASSWORD RECOVER FAILED :" + "=" + TextBox1.Text + "_" & DateTime.Now.ToString("yyyy/MM/dd/HH:mm:ss"))
                                                                                                                                                                                                                                                            file.Close()
                                                                                                                                                                                                                                                            TextBox1.Text = ""
                                                                                                                                                                                                                                                            TextBox2.Text = ""
                                                                                                                                                                                                                                                            GoTo line1
                                                                                                                                                                                                                                                            Me.Close()
                                                                                                                                                                                                                                                            Call LOGIN.Show()
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

    Private Sub Button2_Click(sender As Object, e As EventArgs) Handles Button2.Click
        Me.Close()
        Call LOGIN.Show()
    End Sub
End Class