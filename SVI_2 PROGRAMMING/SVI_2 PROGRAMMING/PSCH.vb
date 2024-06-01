Public Class PSCH
    Private Sub TextBox3_TextChanged(sender As Object, e As EventArgs) Handles TextBox3.TextChanged
        TextBox3.PasswordChar = "*"
    End Sub

    Private Sub TextBox4_TextChanged(sender As Object, e As EventArgs) Handles TextBox4.TextChanged
        TextBox4.PasswordChar = "*"
        If TextBox3.Text = TextBox4.Text Then
            Button4.BackColor = Color.Green
        Else
            Button4.BackColor = Color.Red
        End If
    End Sub

    Private Sub TextBox2_TextChanged(sender As Object, e As EventArgs) Handles TextBox2.TextChanged
        TextBox2.PasswordChar = "*"
        TextBox2.BackColor = Color.White
    End Sub

    Private Sub PSCH_Load(sender As Object, e As EventArgs) Handles MyBase.Load
        Me.TextBox1.Text = SETTING.TextBox1.Text
        TextBox1.Enabled = False
        TextBox2.Enabled = True
        TextBox3.Enabled = False
        TextBox4.Enabled = False
        Button3.Enabled = True
        Button4.Enabled = False
        Button1.Enabled = False
    End Sub

    Private Sub Button3_Click(sender As Object, e As EventArgs) Handles Button3.Click
        Dim line1 As Label
        Dim xlapp As Object
        Dim wb As Object
        Dim st As Object
        xlapp = CreateObject("excel.application")
        wb = xlapp.workbooks.open(CONFIG.TextBox1.Text.ToString())
        st = wb.worksheets("Sheet1")
        If TextBox2.Text = "" Then
            Dim iret As Object = MsgBox("INVALID USER PASSWORD", vbCritical + vbOKOnly, "CHANGE PASSWORD ERROR")
        Else
            If TextBox1.Text = st.range("A2").value And TextBox2.Text = st.range("C2").value Then
                TextBox2.BackColor = Color.Green
                TextBox3.Enabled = True
                TextBox4.Enabled = True
                TextBox2.Enabled = False
                Button3.Enabled = False
                GoTo line1
            Else
                If TextBox1.Text = st.range("A3").value And TextBox2.Text = st.range("C3").value Then
                    TextBox2.BackColor = Color.Green
                    TextBox3.Enabled = True
                    TextBox4.Enabled = True
                    TextBox2.Enabled = False
                    Button3.Enabled = False
                    GoTo line1
                Else
                    If TextBox1.Text = st.range("A4").value And TextBox2.Text = st.range("C4").value Then
                        TextBox2.BackColor = Color.Green
                        TextBox3.Enabled = True
                        TextBox4.Enabled = True
                        TextBox2.Enabled = False
                        Button3.Enabled = False
                        GoTo line1
                    Else
                        If TextBox1.Text = st.range("A5").value And TextBox2.Text = st.range("C5").value Then
                            TextBox2.BackColor = Color.Green
                            TextBox3.Enabled = True
                            TextBox4.Enabled = True
                            TextBox2.Enabled = False
                            Button3.Enabled = False
                            GoTo line1
                        Else
                            If TextBox1.Text = st.range("A6").value And TextBox2.Text = st.range("C6").value Then
                                TextBox2.BackColor = Color.Green
                                TextBox3.Enabled = True
                                TextBox4.Enabled = True
                                TextBox2.Enabled = False
                                Button3.Enabled = False
                                GoTo line1
                            Else
                                If TextBox1.Text = st.range("A7").value And TextBox2.Text = st.range("C7").value Then
                                    TextBox2.BackColor = Color.Green
                                    TextBox3.Enabled = True
                                    TextBox4.Enabled = True
                                    TextBox2.Enabled = False
                                    Button3.Enabled = False
                                    GoTo line1
                                Else
                                    If TextBox1.Text = st.range("A8").value And TextBox2.Text = st.range("C8").value Then
                                        TextBox2.BackColor = Color.Green
                                        TextBox3.Enabled = True
                                        TextBox4.Enabled = True
                                        TextBox2.Enabled = False
                                        Button3.Enabled = False
                                        GoTo line1
                                    Else
                                        If TextBox1.Text = st.range("A9").value And TextBox2.Text = st.range("C9").value Then
                                            TextBox2.BackColor = Color.Green
                                            TextBox3.Enabled = True
                                            TextBox4.Enabled = True
                                            TextBox2.Enabled = False
                                            Button3.Enabled = False
                                            GoTo line1
                                        Else
                                            If TextBox1.Text = st.range("A10").value And TextBox2.Text = st.range("C10").value Then
                                                TextBox2.BackColor = Color.Green
                                                TextBox3.Enabled = True
                                                TextBox4.Enabled = True
                                                TextBox2.Enabled = False
                                                Button3.Enabled = False
                                                GoTo line1
                                            Else
                                                If TextBox1.Text = st.range("A11").value And TextBox2.Text = st.range("C11").value Then
                                                    TextBox2.BackColor = Color.Green
                                                    TextBox3.Enabled = True
                                                    TextBox4.Enabled = True
                                                    TextBox2.Enabled = False
                                                    Button3.Enabled = False
                                                    GoTo line1
                                                Else
                                                    If TextBox1.Text = st.range("A12").value And TextBox2.Text = st.range("C12").value Then
                                                        TextBox2.BackColor = Color.Green
                                                        TextBox3.Enabled = True
                                                        TextBox4.Enabled = True
                                                        TextBox2.Enabled = False
                                                        Button3.Enabled = False
                                                        GoTo line1
                                                    Else
                                                        If TextBox1.Text = st.range("A13").value And TextBox2.Text = st.range("C13").value Then
                                                            TextBox2.BackColor = Color.Green
                                                            TextBox3.Enabled = True
                                                            TextBox4.Enabled = True
                                                            TextBox2.Enabled = False
                                                            Button3.Enabled = False
                                                            GoTo line1
                                                        Else
                                                            If TextBox1.Text = st.range("A14").value And TextBox2.Text = st.range("C14").value Then
                                                                TextBox2.BackColor = Color.Green
                                                                TextBox3.Enabled = True
                                                                TextBox4.Enabled = True
                                                                TextBox2.Enabled = False
                                                                Button3.Enabled = False
                                                                GoTo line1
                                                            Else
                                                                If TextBox1.Text = st.range("A15").value And TextBox2.Text = st.range("C15").value Then
                                                                    TextBox2.BackColor = Color.Green
                                                                    TextBox3.Enabled = True
                                                                    TextBox4.Enabled = True
                                                                    TextBox2.Enabled = False
                                                                    Button3.Enabled = False
                                                                    GoTo line1
                                                                Else
                                                                    If TextBox1.Text = st.range("A16").value And TextBox2.Text = st.range("C16").value Then
                                                                        TextBox2.BackColor = Color.Green
                                                                        TextBox3.Enabled = True
                                                                        TextBox4.Enabled = True
                                                                        TextBox2.Enabled = False
                                                                        Button3.Enabled = False
                                                                        GoTo line1
                                                                    Else
                                                                        If TextBox1.Text = st.range("A17").value And TextBox2.Text = st.range("C17").value Then
                                                                            TextBox2.BackColor = Color.Green
                                                                            TextBox3.Enabled = True
                                                                            TextBox4.Enabled = True
                                                                            TextBox2.Enabled = False
                                                                            Button3.Enabled = False
                                                                            GoTo line1
                                                                        Else
                                                                            If TextBox1.Text = st.range("A18").value And TextBox2.Text = st.range("C18").value Then
                                                                                TextBox2.BackColor = Color.Green
                                                                                TextBox3.Enabled = True
                                                                                TextBox4.Enabled = True
                                                                                TextBox2.Enabled = False
                                                                                Button3.Enabled = False
                                                                                GoTo line1
                                                                            Else
                                                                                If TextBox1.Text = st.range("A19").value And TextBox2.Text = st.range("C19").value Then
                                                                                    TextBox2.BackColor = Color.Green
                                                                                    TextBox3.Enabled = True
                                                                                    TextBox4.Enabled = True
                                                                                    TextBox2.Enabled = False
                                                                                    Button3.Enabled = False
                                                                                    GoTo line1
                                                                                Else
                                                                                    If TextBox1.Text = st.range("A20").value And TextBox2.Text = st.range("C20").value Then
                                                                                        TextBox2.BackColor = Color.Green
                                                                                        TextBox3.Enabled = True
                                                                                        TextBox4.Enabled = True
                                                                                        TextBox2.Enabled = False
                                                                                        Button3.Enabled = False
                                                                                        GoTo line1
                                                                                    Else
                                                                                        If TextBox1.Text = st.range("A21").value And TextBox2.Text = st.range("C21").value Then
                                                                                            TextBox2.BackColor = Color.Green
                                                                                            TextBox3.Enabled = True
                                                                                            TextBox4.Enabled = True
                                                                                            TextBox2.Enabled = False
                                                                                            Button3.Enabled = False
                                                                                            GoTo line1
                                                                                        Else
                                                                                            If TextBox1.Text = st.range("A22").value And TextBox2.Text = st.range("C22").value Then
                                                                                                TextBox2.BackColor = Color.Green
                                                                                                TextBox3.Enabled = True
                                                                                                TextBox4.Enabled = True
                                                                                                TextBox2.Enabled = False
                                                                                                Button3.Enabled = False
                                                                                                GoTo line1
                                                                                            Else
                                                                                                If TextBox1.Text = st.range("A23").value And TextBox2.Text = st.range("C23").value Then
                                                                                                    TextBox2.BackColor = Color.Green
                                                                                                    TextBox3.Enabled = True
                                                                                                    TextBox4.Enabled = True
                                                                                                    TextBox2.Enabled = False
                                                                                                    Button3.Enabled = False
                                                                                                    GoTo line1
                                                                                                Else
                                                                                                    If TextBox1.Text = st.range("A24").value And TextBox2.Text = st.range("C24").value Then
                                                                                                        TextBox2.BackColor = Color.Green
                                                                                                        TextBox3.Enabled = True
                                                                                                        TextBox4.Enabled = True
                                                                                                        TextBox2.Enabled = False
                                                                                                        Button3.Enabled = False
                                                                                                        GoTo line1
                                                                                                    Else
                                                                                                        If TextBox1.Text = st.range("A25").value And TextBox2.Text = st.range("C25").value Then
                                                                                                            TextBox2.BackColor = Color.Green
                                                                                                            TextBox3.Enabled = True
                                                                                                            TextBox4.Enabled = True
                                                                                                            TextBox2.Enabled = False
                                                                                                            Button3.Enabled = False
                                                                                                            GoTo line1
                                                                                                        Else
                                                                                                            If TextBox1.Text = st.range("A26").value And TextBox2.Text = st.range("C26").value Then
                                                                                                                TextBox2.BackColor = Color.Green
                                                                                                                TextBox3.Enabled = True
                                                                                                                TextBox4.Enabled = True
                                                                                                                TextBox2.Enabled = False
                                                                                                                Button3.Enabled = False
                                                                                                                GoTo line1
                                                                                                            Else
                                                                                                                If TextBox1.Text = st.range("A27").value And TextBox2.Text = st.range("C27").value Then
                                                                                                                    TextBox2.BackColor = Color.Green
                                                                                                                    TextBox3.Enabled = True
                                                                                                                    TextBox4.Enabled = True
                                                                                                                    TextBox2.Enabled = False
                                                                                                                    Button3.Enabled = False
                                                                                                                    GoTo line1
                                                                                                                Else
                                                                                                                    If TextBox1.Text = st.range("A28").value And TextBox2.Text = st.range("C28").value Then
                                                                                                                        TextBox2.BackColor = Color.Green
                                                                                                                        TextBox3.Enabled = True
                                                                                                                        TextBox4.Enabled = True
                                                                                                                        TextBox2.Enabled = False
                                                                                                                        Button3.Enabled = False
                                                                                                                        GoTo line1
                                                                                                                    Else
                                                                                                                        If TextBox1.Text = st.range("A29").value And TextBox2.Text = st.range("C29").value Then
                                                                                                                            TextBox2.BackColor = Color.Green
                                                                                                                            TextBox3.Enabled = True
                                                                                                                            TextBox4.Enabled = True
                                                                                                                            TextBox2.Enabled = False
                                                                                                                            Button3.Enabled = False
                                                                                                                            GoTo line1
                                                                                                                        Else
                                                                                                                            If TextBox1.Text = st.range("A30").value And TextBox2.Text = st.range("C30").value Then
                                                                                                                                TextBox2.BackColor = Color.Green
                                                                                                                                TextBox3.Enabled = True
                                                                                                                                TextBox4.Enabled = True
                                                                                                                                TextBox2.Enabled = False
                                                                                                                                Button3.Enabled = False
                                                                                                                                GoTo line1
                                                                                                                            Else
                                                                                                                                If TextBox1.Text = st.range("A31").value And TextBox2.Text = st.range("C31").value Then
                                                                                                                                    TextBox2.BackColor = Color.Green
                                                                                                                                    TextBox3.Enabled = True
                                                                                                                                    TextBox4.Enabled = True
                                                                                                                                    TextBox2.Enabled = False
                                                                                                                                    Button3.Enabled = False
                                                                                                                                    GoTo line1
                                                                                                                                Else
                                                                                                                                    TextBox2.BackColor = Color.Red
                                                                                                                                    TextBox3.Enabled = False
                                                                                                                                    TextBox4.Enabled = False
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

    Private Sub Button1_Click(sender As Object, e As EventArgs) Handles Button1.Click
        Dim line1 As Label
        Dim xlapp As Object
        Dim wb As Object
        Dim st As Object
        xlapp = CreateObject("excel.application")
        wb = xlapp.workbooks.open(CONFIG.TextBox1.Text.ToString())
        st = wb.worksheets("Sheet1")
        st.unprotect("Test@123")
        Dim file As System.IO.StreamWriter
        Dim path As String = CONFIG.TextBox2.Text.ToString()
        file = My.Computer.FileSystem.OpenTextFileWriter(path, True)
        If TextBox1.Text = st.range("A2").value And TextBox2.BackColor = Color.Green Then
            st.range("C2").value = Me.TextBox4.Text
            Dim iret As Object = MsgBox("PASSWORD SUCCESSSFULLY CHANGED FOR :" & " " & TextBox1.Text, vbInformation + vbOKOnly, "PASSWORD CHANGE")
            If iret = vbOK Then
                file.WriteLine("PASSWORD CHANGED FOR :" & " " & TextBox1.Text & " " & "AT" & " " & DateTime.Now.ToString("yyyy/MM/dd/HH:mm:ss"))
                file.Close()
                TextBox2.Text = ""
                TextBox3.Text = ""
                TextBox4.Text = ""
                TextBox2.Enabled = True
                Button3.Enabled = True
                TextBox3.Enabled = False
                TextBox4.Enabled = False
                Button1.Enabled = False
                GoTo line1
            End If
        Else
            If TextBox1.Text = st.range("A3").value And TextBox2.BackColor = Color.Green Then
                st.range("C3").value = Me.TextBox4.Text
                Dim iret1 As Object = MsgBox("PASSWORD SUCCESSSFULLY CHANGED FOR :" & " " & TextBox1.Text, vbInformation + vbOKOnly, "PASSWORD CHANGE")
                If iret1 = vbOK Then
                    file.WriteLine("PASSWORD CHANGED FOR :" & " " & TextBox1.Text & " " & "AT" & " " & DateTime.Now.ToString("yyyy/MM/dd/HH:mm:ss"))
                    file.Close()
                    TextBox2.Text = ""
                    TextBox3.Text = ""
                    TextBox4.Text = ""
                    TextBox2.Enabled = True
                    Button3.Enabled = True
                    TextBox3.Enabled = False
                    TextBox4.Enabled = False
                    Button1.Enabled = False
                    GoTo line1
                End If
            Else
                If TextBox1.Text = st.range("A4").value And TextBox2.BackColor = Color.Green Then
                    st.range("C4").value = Me.TextBox4.Text
                    Dim iret2 As Object = MsgBox("PASSWORD SUCCESSSFULLY CHANGED FOR :" & " " & TextBox1.Text, vbInformation + vbOKOnly, "PASSWORD CHANGE")
                    If iret2 = vbOK Then
                        file.WriteLine("PASSWORD CHANGED FOR :" & " " & TextBox1.Text & " " & "AT" & " " & DateTime.Now.ToString("yyyy/MM/dd/HH:mm:ss"))
                        file.Close()
                        TextBox2.Text = ""
                        TextBox3.Text = ""
                        TextBox4.Text = ""
                        TextBox2.Enabled = True
                        Button3.Enabled = True
                        TextBox3.Enabled = False
                        TextBox4.Enabled = False
                        Button1.Enabled = False
                        GoTo line1
                    End If
                Else
                    If TextBox1.Text = st.range("A5").value And TextBox2.BackColor = Color.Green Then
                        st.range("C5").value = Me.TextBox4.Text
                        Dim iret3 As Object = MsgBox("PASSWORD SUCCESSSFULLY CHANGED FOR :" & " " & TextBox1.Text, vbInformation + vbOKOnly, "PASSWORD CHANGE")
                        If iret3 = vbOK Then
                            file.WriteLine("PASSWORD CHANGED FOR :" & " " & TextBox1.Text & " " & "AT" & " " & DateTime.Now.ToString("yyyy/MM/dd/HH:mm:ss"))
                            file.Close()
                            TextBox2.Text = ""
                            TextBox3.Text = ""
                            TextBox4.Text = ""
                            TextBox2.Enabled = True
                            Button3.Enabled = True
                            TextBox3.Enabled = False
                            TextBox4.Enabled = False
                            Button1.Enabled = False
                            GoTo line1
                        End If
                    Else
                        If TextBox1.Text = st.range("A6").value And TextBox2.BackColor = Color.Green Then
                            st.range("C6").value = Me.TextBox4.Text
                            Dim iret4 As Object = MsgBox("PASSWORD SUCCESSSFULLY CHANGED FOR :" & " " & TextBox1.Text, vbInformation + vbOKOnly, "PASSWORD CHANGE")
                            If iret4 = vbOK Then
                                file.WriteLine("PASSWORD CHANGED FOR :" & " " & TextBox1.Text & " " & "AT" & " " & DateTime.Now.ToString("yyyy/MM/dd/HH:mm:ss"))
                                file.Close()
                                TextBox2.Text = ""
                                TextBox3.Text = ""
                                TextBox4.Text = ""
                                TextBox2.Enabled = True
                                Button3.Enabled = True
                                TextBox3.Enabled = False
                                TextBox4.Enabled = False
                                Button1.Enabled = False
                                GoTo line1
                            End If
                        Else
                            If TextBox1.Text = st.range("A7").value And TextBox2.BackColor = Color.Green Then
                                st.range("C7").value = Me.TextBox4.Text
                                Dim iret5 As Object = MsgBox("PASSWORD SUCCESSSFULLY CHANGED FOR :" & " " & TextBox1.Text, vbInformation + vbOKOnly, "PASSWORD CHANGE")
                                If iret5 = vbOK Then
                                    file.WriteLine("PASSWORD CHANGED FOR :" & " " & TextBox1.Text & " " & "AT" & " " & DateTime.Now.ToString("yyyy/MM/dd/HH:mm:ss"))
                                    file.Close()
                                    TextBox2.Text = ""
                                    TextBox3.Text = ""
                                    TextBox4.Text = ""
                                    TextBox2.Enabled = True
                                    Button3.Enabled = True
                                    TextBox3.Enabled = False
                                    TextBox4.Enabled = False
                                    Button1.Enabled = False
                                    GoTo line1
                                End If
                            Else
                                If TextBox1.Text = st.range("A8").value And TextBox2.BackColor = Color.Green Then
                                    st.range("C8").value = Me.TextBox4.Text
                                    Dim iret6 As Object = MsgBox("PASSWORD SUCCESSSFULLY CHANGED FOR :" & " " & TextBox1.Text, vbInformation + vbOKOnly, "PASSWORD CHANGE")
                                    If iret6 = vbOK Then
                                        file.WriteLine("PASSWORD CHANGED FOR :" & " " & TextBox1.Text & " " & "AT" & " " & DateTime.Now.ToString("yyyy/MM/dd/HH:mm:ss"))
                                        file.Close()
                                        TextBox2.Text = ""
                                        TextBox3.Text = ""
                                        TextBox4.Text = ""
                                        TextBox2.Enabled = True
                                        Button3.Enabled = True
                                        TextBox3.Enabled = False
                                        TextBox4.Enabled = False
                                        Button1.Enabled = False
                                        GoTo line1
                                    End If
                                Else
                                    If TextBox1.Text = st.range("A9").value And TextBox2.BackColor = Color.Green Then
                                        st.range("C9").value = Me.TextBox4.Text
                                        Dim iret7 As Object = MsgBox("PASSWORD SUCCESSSFULLY CHANGED FOR :" & " " & TextBox1.Text, vbInformation + vbOKOnly, "PASSWORD CHANGE")
                                        If iret7 = vbOK Then
                                            file.WriteLine("PASSWORD CHANGED FOR :" & " " & TextBox1.Text & " " & "AT" & " " & DateTime.Now.ToString("yyyy/MM/dd/HH:mm:ss"))
                                            file.Close()
                                            TextBox2.Text = ""
                                            TextBox3.Text = ""
                                            TextBox4.Text = ""
                                            TextBox2.Enabled = True
                                            Button3.Enabled = True
                                            TextBox3.Enabled = False
                                            TextBox4.Enabled = False
                                            Button1.Enabled = False
                                            GoTo line1
                                        End If
                                    Else
                                        If TextBox1.Text = st.range("A10").value And TextBox2.BackColor = Color.Green Then
                                            st.range("C10").value = Me.TextBox4.Text
                                            Dim iret8 As Object = MsgBox("PASSWORD SUCCESSSFULLY CHANGED FOR :" & " " & TextBox1.Text, vbInformation + vbOKOnly, "PASSWORD CHANGE")
                                            If iret8 = vbOK Then
                                                file.WriteLine("PASSWORD CHANGED FOR :" & " " & TextBox1.Text & " " & "AT" & " " & DateTime.Now.ToString("yyyy/MM/dd/HH:mm:ss"))
                                                file.Close()
                                                TextBox2.Text = ""
                                                TextBox3.Text = ""
                                                TextBox4.Text = ""
                                                TextBox2.Enabled = True
                                                Button3.Enabled = True
                                                TextBox3.Enabled = False
                                                TextBox4.Enabled = False
                                                Button1.Enabled = False
                                                GoTo line1
                                            End If
                                        Else
                                            If TextBox1.Text = st.range("A11").value And TextBox2.BackColor = Color.Green Then
                                                st.range("C11").value = Me.TextBox4.Text
                                                Dim iret9 As Object = MsgBox("PASSWORD SUCCESSSFULLY CHANGED FOR :" & " " & TextBox1.Text, vbInformation + vbOKOnly, "PASSWORD CHANGE")
                                                If iret9 = vbOK Then
                                                    file.WriteLine("PASSWORD CHANGED FOR :" & " " & TextBox1.Text & " " & "AT" & " " & DateTime.Now.ToString("yyyy/MM/dd/HH:mm:ss"))
                                                    file.Close()
                                                    TextBox2.Text = ""
                                                    TextBox3.Text = ""
                                                    TextBox4.Text = ""
                                                    TextBox2.Enabled = True
                                                    Button3.Enabled = True
                                                    TextBox3.Enabled = False
                                                    TextBox4.Enabled = False
                                                    Button1.Enabled = False
                                                    GoTo line1
                                                End If
                                            Else
                                                If TextBox1.Text = st.range("A12").value And TextBox2.BackColor = Color.Green Then
                                                    st.range("C12").value = Me.TextBox4.Text
                                                    Dim iret10 As Object = MsgBox("PASSWORD SUCCESSSFULLY CHANGED FOR :" & " " & TextBox1.Text, vbInformation + vbOKOnly, "PASSWORD CHANGE")
                                                    If iret10 = vbOK Then
                                                        file.WriteLine("PASSWORD CHANGED FOR :" & " " & TextBox1.Text & " " & "AT" & " " & DateTime.Now.ToString("yyyy/MM/dd/HH:mm:ss"))
                                                        file.Close()
                                                        TextBox2.Text = ""
                                                        TextBox3.Text = ""
                                                        TextBox4.Text = ""
                                                        TextBox2.Enabled = True
                                                        Button3.Enabled = True
                                                        TextBox3.Enabled = False
                                                        TextBox4.Enabled = False
                                                        Button1.Enabled = False
                                                        GoTo line1
                                                    End If
                                                Else
                                                    If TextBox1.Text = st.range("A13").value And TextBox2.BackColor = Color.Green Then
                                                        st.range("C13").value = Me.TextBox4.Text
                                                        Dim iret11 As Object = MsgBox("PASSWORD SUCCESSSFULLY CHANGED FOR :" & " " & TextBox1.Text, vbInformation + vbOKOnly, "PASSWORD CHANGE")
                                                        If iret11 = vbOK Then
                                                            file.WriteLine("PASSWORD CHANGED FOR :" & " " & TextBox1.Text & " " & "AT" & " " & DateTime.Now.ToString("yyyy/MM/dd/HH:mm:ss"))
                                                            file.Close()
                                                            TextBox2.Text = ""
                                                            TextBox3.Text = ""
                                                            TextBox4.Text = ""
                                                            TextBox2.Enabled = True
                                                            Button3.Enabled = True
                                                            TextBox3.Enabled = False
                                                            TextBox4.Enabled = False
                                                            Button1.Enabled = False
                                                            GoTo line1
                                                        End If
                                                    Else
                                                        If TextBox1.Text = st.range("A14").value And TextBox2.BackColor = Color.Green Then
                                                            st.range("C14").value = Me.TextBox4.Text
                                                            Dim iret12 As Object = MsgBox("PASSWORD SUCCESSSFULLY CHANGED FOR :" & " " & TextBox1.Text, vbInformation + vbOKOnly, "PASSWORD CHANGE")
                                                            If iret12 = vbOK Then
                                                                file.WriteLine("PASSWORD CHANGED FOR :" & " " & TextBox1.Text & " " & "AT" & " " & DateTime.Now.ToString("yyyy/MM/dd/HH:mm:ss"))
                                                                file.Close()
                                                                TextBox2.Text = ""
                                                                TextBox3.Text = ""
                                                                TextBox4.Text = ""
                                                                TextBox2.Enabled = True
                                                                Button3.Enabled = True
                                                                TextBox3.Enabled = False
                                                                TextBox4.Enabled = False
                                                                Button1.Enabled = False
                                                                GoTo line1
                                                            End If
                                                        Else
                                                            If TextBox1.Text = st.range("A15").value And TextBox2.BackColor = Color.Green Then
                                                                st.range("C15").value = Me.TextBox4.Text
                                                                Dim iret13 As Object = MsgBox("PASSWORD SUCCESSSFULLY CHANGED FOR :" & " " & TextBox1.Text, vbInformation + vbOKOnly, "PASSWORD CHANGE")
                                                                If iret13 = vbOK Then
                                                                    file.WriteLine("PASSWORD CHANGED FOR :" & " " & TextBox1.Text & " " & "AT" & " " & DateTime.Now.ToString("yyyy/MM/dd/HH:mm:ss"))
                                                                    file.Close()
                                                                    TextBox2.Text = ""
                                                                    TextBox3.Text = ""
                                                                    TextBox4.Text = ""
                                                                    TextBox2.Enabled = True
                                                                    Button3.Enabled = True
                                                                    TextBox3.Enabled = False
                                                                    TextBox4.Enabled = False
                                                                    Button1.Enabled = False
                                                                    GoTo line1
                                                                End If
                                                            Else
                                                                If TextBox1.Text = st.range("A16").value And TextBox2.BackColor = Color.Green Then
                                                                    st.range("C16").value = Me.TextBox4.Text
                                                                    Dim iret14 As Object = MsgBox("PASSWORD SUCCESSSFULLY CHANGED FOR :" & " " & TextBox1.Text, vbInformation + vbOKOnly, "PASSWORD CHANGE")
                                                                    If iret14 = vbOK Then
                                                                        file.WriteLine("PASSWORD CHANGED FOR :" & " " & TextBox1.Text & " " & "AT" & " " & DateTime.Now.ToString("yyyy/MM/dd/HH:mm:ss"))
                                                                        file.Close()
                                                                        TextBox2.Text = ""
                                                                        TextBox3.Text = ""
                                                                        TextBox4.Text = ""
                                                                        TextBox2.Enabled = True
                                                                        Button3.Enabled = True
                                                                        TextBox3.Enabled = False
                                                                        TextBox4.Enabled = False
                                                                        Button1.Enabled = False
                                                                        GoTo line1
                                                                    End If
                                                                Else
                                                                    If TextBox1.Text = st.range("A17").value And TextBox2.BackColor = Color.Green Then
                                                                        st.range("C17").value = Me.TextBox4.Text
                                                                        Dim iret15 As Object = MsgBox("PASSWORD SUCCESSSFULLY CHANGED FOR :" & " " & TextBox1.Text, vbInformation + vbOKOnly, "PASSWORD CHANGE")
                                                                        If iret15 = vbOK Then
                                                                            file.WriteLine("PASSWORD CHANGED FOR :" & " " & TextBox1.Text & " " & "AT" & " " & DateTime.Now.ToString("yyyy/MM/dd/HH:mm:ss"))
                                                                            file.Close()
                                                                            TextBox2.Text = ""
                                                                            TextBox3.Text = ""
                                                                            TextBox4.Text = ""
                                                                            TextBox2.Enabled = True
                                                                            Button3.Enabled = True
                                                                            TextBox3.Enabled = False
                                                                            TextBox4.Enabled = False
                                                                            Button1.Enabled = False
                                                                            GoTo line1
                                                                        End If
                                                                    Else
                                                                        If TextBox1.Text = st.range("A18").value And TextBox2.BackColor = Color.Green Then
                                                                            st.range("C18").value = Me.TextBox4.Text
                                                                            Dim iret16 As Object = MsgBox("PASSWORD SUCCESSSFULLY CHANGED FOR :" & " " & TextBox1.Text, vbInformation + vbOKOnly, "PASSWORD CHANGE")
                                                                            If iret16 = vbOK Then
                                                                                file.WriteLine("PASSWORD CHANGED FOR :" & " " & TextBox1.Text & " " & "AT" & " " & DateTime.Now.ToString("yyyy/MM/dd/HH:mm:ss"))
                                                                                file.Close()
                                                                                TextBox2.Text = ""
                                                                                TextBox3.Text = ""
                                                                                TextBox4.Text = ""
                                                                                TextBox2.Enabled = True
                                                                                Button3.Enabled = True
                                                                                TextBox3.Enabled = False
                                                                                TextBox4.Enabled = False
                                                                                Button1.Enabled = False
                                                                                GoTo line1
                                                                            End If
                                                                        Else
                                                                            If TextBox1.Text = st.range("A19").value And TextBox2.BackColor = Color.Green Then
                                                                                st.range("C19").value = Me.TextBox4.Text
                                                                                Dim iret17 As Object = MsgBox("PASSWORD SUCCESSSFULLY CHANGED FOR :" & " " & TextBox1.Text, vbInformation + vbOKOnly, "PASSWORD CHANGE")
                                                                                If iret17 = vbOK Then
                                                                                    file.WriteLine("PASSWORD CHANGED FOR :" & " " & TextBox1.Text & " " & "AT" & " " & DateTime.Now.ToString("yyyy/MM/dd/HH:mm:ss"))
                                                                                    file.Close()
                                                                                    TextBox2.Text = ""
                                                                                    TextBox3.Text = ""
                                                                                    TextBox4.Text = ""
                                                                                    TextBox2.Enabled = True
                                                                                    Button3.Enabled = True
                                                                                    TextBox3.Enabled = False
                                                                                    TextBox4.Enabled = False
                                                                                    Button1.Enabled = False
                                                                                    GoTo line1
                                                                                End If
                                                                            Else
                                                                                If TextBox1.Text = st.range("A20").value And TextBox2.BackColor = Color.Green Then
                                                                                    st.range("C20").value = Me.TextBox4.Text
                                                                                    Dim iret18 As Object = MsgBox("PASSWORD SUCCESSSFULLY CHANGED FOR :" & " " & TextBox1.Text, vbInformation + vbOKOnly, "PASSWORD CHANGE")
                                                                                    If iret18 = vbOK Then
                                                                                        file.WriteLine("PASSWORD CHANGED FOR :" & " " & TextBox1.Text & " " & "AT" & " " & DateTime.Now.ToString("yyyy/MM/dd/HH:mm:ss"))
                                                                                        file.Close()
                                                                                        TextBox2.Text = ""
                                                                                        TextBox3.Text = ""
                                                                                        TextBox4.Text = ""
                                                                                        TextBox2.Enabled = True
                                                                                        Button3.Enabled = True
                                                                                        TextBox3.Enabled = False
                                                                                        TextBox4.Enabled = False
                                                                                        Button1.Enabled = False
                                                                                        GoTo line1
                                                                                    End If
                                                                                Else
                                                                                    If TextBox1.Text = st.range("A21").value And TextBox2.BackColor = Color.Green Then
                                                                                        st.range("C21").value = Me.TextBox4.Text
                                                                                        Dim iret19 As Object = MsgBox("PASSWORD SUCCESSSFULLY CHANGED FOR :" & " " & TextBox1.Text, vbInformation + vbOKOnly, "PASSWORD CHANGE")
                                                                                        If iret19 = vbOK Then
                                                                                            file.WriteLine("PASSWORD CHANGED FOR :" & " " & TextBox1.Text & " " & "AT" & " " & DateTime.Now.ToString("yyyy/MM/dd/HH:mm:ss"))
                                                                                            file.Close()
                                                                                            TextBox2.Text = ""
                                                                                            TextBox3.Text = ""
                                                                                            TextBox4.Text = ""
                                                                                            TextBox2.Enabled = True
                                                                                            Button3.Enabled = True
                                                                                            TextBox3.Enabled = False
                                                                                            TextBox4.Enabled = False
                                                                                            Button1.Enabled = False
                                                                                            GoTo line1
                                                                                        End If
                                                                                    Else
                                                                                        If TextBox1.Text = st.range("A22").value And TextBox2.BackColor = Color.Green Then
                                                                                            st.range("C22").value = Me.TextBox4.Text
                                                                                            Dim iret20 As Object = MsgBox("PASSWORD SUCCESSSFULLY CHANGED FOR :" & " " & TextBox1.Text, vbInformation + vbOKOnly, "PASSWORD CHANGE")
                                                                                            If iret20 = vbOK Then
                                                                                                file.WriteLine("PASSWORD CHANGED FOR :" & " " & TextBox1.Text & " " & "AT" & " " & DateTime.Now.ToString("yyyy/MM/dd/HH:mm:ss"))
                                                                                                file.Close()
                                                                                                TextBox2.Text = ""
                                                                                                TextBox3.Text = ""
                                                                                                TextBox4.Text = ""
                                                                                                TextBox2.Enabled = True
                                                                                                Button3.Enabled = True
                                                                                                TextBox3.Enabled = False
                                                                                                TextBox4.Enabled = False
                                                                                                Button1.Enabled = False
                                                                                                GoTo line1
                                                                                            End If
                                                                                        Else
                                                                                            If TextBox1.Text = st.range("A23").value And TextBox2.BackColor = Color.Green Then
                                                                                                st.range("C23").value = Me.TextBox4.Text
                                                                                                Dim iret21 As Object = MsgBox("PASSWORD SUCCESSSFULLY CHANGED FOR :" & " " & TextBox1.Text, vbInformation + vbOKOnly, "PASSWORD CHANGE")
                                                                                                If iret21 = vbOK Then
                                                                                                    file.WriteLine("PASSWORD CHANGED FOR :" & " " & TextBox1.Text & " " & "AT" & " " & DateTime.Now.ToString("yyyy/MM/dd/HH:mm:ss"))
                                                                                                    file.Close()
                                                                                                    TextBox2.Text = ""
                                                                                                    TextBox3.Text = ""
                                                                                                    TextBox4.Text = ""
                                                                                                    TextBox2.Enabled = True
                                                                                                    Button3.Enabled = True
                                                                                                    TextBox3.Enabled = False
                                                                                                    TextBox4.Enabled = False
                                                                                                    Button1.Enabled = False
                                                                                                    GoTo line1
                                                                                                End If
                                                                                            Else
                                                                                                If TextBox1.Text = st.range("A24").value And TextBox2.BackColor = Color.Green Then
                                                                                                    st.range("C24").value = Me.TextBox4.Text
                                                                                                    Dim iret22 As Object = MsgBox("PASSWORD SUCCESSSFULLY CHANGED FOR :" & " " & TextBox1.Text, vbInformation + vbOKOnly, "PASSWORD CHANGE")
                                                                                                    If iret22 = vbOK Then
                                                                                                        file.WriteLine("PASSWORD CHANGED FOR :" & " " & TextBox1.Text & " " & "AT" & " " & DateTime.Now.ToString("yyyy/MM/dd/HH:mm:ss"))
                                                                                                        file.Close()
                                                                                                        TextBox2.Text = ""
                                                                                                        TextBox3.Text = ""
                                                                                                        TextBox4.Text = ""
                                                                                                        TextBox2.Enabled = True
                                                                                                        Button3.Enabled = True
                                                                                                        TextBox3.Enabled = False
                                                                                                        TextBox4.Enabled = False
                                                                                                        Button1.Enabled = False
                                                                                                        GoTo line1
                                                                                                    End If
                                                                                                Else
                                                                                                    If TextBox1.Text = st.range("A25").value And TextBox2.BackColor = Color.Green Then
                                                                                                        st.range("C25").value = Me.TextBox4.Text
                                                                                                        Dim iret23 As Object = MsgBox("PASSWORD SUCCESSSFULLY CHANGED FOR :" & " " & TextBox1.Text, vbInformation + vbOKOnly, "PASSWORD CHANGE")
                                                                                                        If iret23 = vbOK Then
                                                                                                            file.WriteLine("PASSWORD CHANGED FOR :" & " " & TextBox1.Text & " " & "AT" & " " & DateTime.Now.ToString("yyyy/MM/dd/HH:mm:ss"))
                                                                                                            file.Close()
                                                                                                            TextBox2.Text = ""
                                                                                                            TextBox3.Text = ""
                                                                                                            TextBox4.Text = ""
                                                                                                            TextBox2.Enabled = True
                                                                                                            Button3.Enabled = True
                                                                                                            TextBox3.Enabled = False
                                                                                                            TextBox4.Enabled = False
                                                                                                            Button1.Enabled = False
                                                                                                            GoTo line1
                                                                                                        End If
                                                                                                    Else
                                                                                                        If TextBox1.Text = st.range("A26").value And TextBox2.BackColor = Color.Green Then
                                                                                                            st.range("C26").value = Me.TextBox4.Text
                                                                                                            Dim iret24 As Object = MsgBox("PASSWORD SUCCESSSFULLY CHANGED FOR :" & " " & TextBox1.Text, vbInformation + vbOKOnly, "PASSWORD CHANGE")
                                                                                                            If iret24 = vbOK Then
                                                                                                                file.WriteLine("PASSWORD CHANGED FOR :" & " " & TextBox1.Text & " " & "AT" & " " & DateTime.Now.ToString("yyyy/MM/dd/HH:mm:ss"))
                                                                                                                file.Close()
                                                                                                                TextBox2.Text = ""
                                                                                                                TextBox3.Text = ""
                                                                                                                TextBox4.Text = ""
                                                                                                                TextBox2.Enabled = True
                                                                                                                Button3.Enabled = True
                                                                                                                TextBox3.Enabled = False
                                                                                                                TextBox4.Enabled = False
                                                                                                                Button1.Enabled = False
                                                                                                                GoTo line1
                                                                                                            End If
                                                                                                        Else
                                                                                                            If TextBox1.Text = st.range("A27").value And TextBox2.BackColor = Color.Green Then
                                                                                                                st.range("C27").value = Me.TextBox4.Text
                                                                                                                Dim iret25 As Object = MsgBox("PASSWORD SUCCESSSFULLY CHANGED FOR :" & " " & TextBox1.Text, vbInformation + vbOKOnly, "PASSWORD CHANGE")
                                                                                                                If iret25 = vbOK Then
                                                                                                                    file.WriteLine("PASSWORD CHANGED FOR :" & " " & TextBox1.Text & " " & "AT" & " " & DateTime.Now.ToString("yyyy/MM/dd/HH:mm:ss"))
                                                                                                                    file.Close()
                                                                                                                    TextBox2.Text = ""
                                                                                                                    TextBox3.Text = ""
                                                                                                                    TextBox4.Text = ""
                                                                                                                    TextBox2.Enabled = True
                                                                                                                    Button3.Enabled = True
                                                                                                                    TextBox3.Enabled = False
                                                                                                                    TextBox4.Enabled = False
                                                                                                                    Button1.Enabled = False
                                                                                                                    GoTo line1
                                                                                                                End If
                                                                                                            Else
                                                                                                                If TextBox1.Text = st.range("A28").value And TextBox2.BackColor = Color.Green Then
                                                                                                                    st.range("C28").value = Me.TextBox4.Text
                                                                                                                    Dim iret26 As Object = MsgBox("PASSWORD SUCCESSSFULLY CHANGED FOR :" & " " & TextBox1.Text, vbInformation + vbOKOnly, "PASSWORD CHANGE")
                                                                                                                    If iret26 = vbOK Then
                                                                                                                        file.WriteLine("PASSWORD CHANGED FOR :" & " " & TextBox1.Text & " " & "AT" & " " & DateTime.Now.ToString("yyyy/MM/dd/HH:mm:ss"))
                                                                                                                        file.Close()
                                                                                                                        TextBox2.Text = ""
                                                                                                                        TextBox3.Text = ""
                                                                                                                        TextBox4.Text = ""
                                                                                                                        TextBox2.Enabled = True
                                                                                                                        Button3.Enabled = True
                                                                                                                        TextBox3.Enabled = False
                                                                                                                        TextBox4.Enabled = False
                                                                                                                        Button1.Enabled = False
                                                                                                                        GoTo line1
                                                                                                                    End If
                                                                                                                Else
                                                                                                                    If TextBox1.Text = st.range("A29").value And TextBox2.BackColor = Color.Green Then
                                                                                                                        st.range("C29").value = Me.TextBox4.Text
                                                                                                                        Dim iret27 As Object = MsgBox("PASSWORD SUCCESSSFULLY CHANGED FOR :" & " " & TextBox1.Text, vbInformation + vbOKOnly, "PASSWORD CHANGE")
                                                                                                                        If iret27 = vbOK Then
                                                                                                                            file.WriteLine("PASSWORD CHANGED FOR :" & " " & TextBox1.Text & " " & "AT" & " " & DateTime.Now.ToString("yyyy/MM/dd/HH:mm:ss"))
                                                                                                                            file.Close()
                                                                                                                            TextBox2.Text = ""
                                                                                                                            TextBox3.Text = ""
                                                                                                                            TextBox4.Text = ""
                                                                                                                            TextBox2.Enabled = True
                                                                                                                            Button3.Enabled = True
                                                                                                                            TextBox3.Enabled = False
                                                                                                                            TextBox4.Enabled = False
                                                                                                                            Button1.Enabled = False
                                                                                                                            GoTo line1
                                                                                                                        End If
                                                                                                                    Else
                                                                                                                        If TextBox1.Text = st.range("A30").value And TextBox2.BackColor = Color.Green Then
                                                                                                                            st.range("C30").value = Me.TextBox4.Text
                                                                                                                            Dim iret28 As Object = MsgBox("PASSWORD SUCCESSSFULLY CHANGED FOR :" & " " & TextBox1.Text, vbInformation + vbOKOnly, "PASSWORD CHANGE")
                                                                                                                            If iret28 = vbOK Then
                                                                                                                                file.WriteLine("PASSWORD CHANGED FOR :" & " " & TextBox1.Text & " " & "AT" & " " & DateTime.Now.ToString("yyyy/MM/dd/HH:mm:ss"))
                                                                                                                                file.Close()
                                                                                                                                TextBox2.Text = ""
                                                                                                                                TextBox3.Text = ""
                                                                                                                                TextBox4.Text = ""
                                                                                                                                TextBox2.Enabled = True
                                                                                                                                Button3.Enabled = True
                                                                                                                                TextBox3.Enabled = False
                                                                                                                                TextBox4.Enabled = False
                                                                                                                                Button1.Enabled = False
                                                                                                                                GoTo line1
                                                                                                                            End If
                                                                                                                        Else
                                                                                                                            If TextBox1.Text = st.range("A31").value And TextBox2.BackColor = Color.Green Then
                                                                                                                                st.range("C31").value = Me.TextBox4.Text
                                                                                                                                Dim iret29 As Object = MsgBox("PASSWORD SUCCESSSFULLY CHANGED FOR :" & " " & TextBox1.Text, vbInformation + vbOKOnly, "PASSWORD CHANGE")
                                                                                                                                If iret29 = vbOK Then
                                                                                                                                    file.WriteLine("PASSWORD CHANGED FOR :" & " " & TextBox1.Text & " " & "AT" & " " & DateTime.Now.ToString("yyyy/MM/dd/HH:mm:ss"))
                                                                                                                                    file.Close()
                                                                                                                                    TextBox2.Text = ""
                                                                                                                                    TextBox3.Text = ""
                                                                                                                                    TextBox4.Text = ""
                                                                                                                                    TextBox2.Enabled = True
                                                                                                                                    Button3.Enabled = True
                                                                                                                                    TextBox3.Enabled = False
                                                                                                                                    TextBox4.Enabled = False
                                                                                                                                    Button1.Enabled = False
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

line1:
        st.protect("Test@123")
        wb.close(True)
        System.Runtime.InteropServices.Marshal.ReleaseComObject(wb)
        wb = Nothing
        xlapp.quit()
        System.Runtime.InteropServices.Marshal.ReleaseComObject(xlapp)
        xlapp = Nothing
        Process.GetProcessesByName("EXCEL")(0).Kill()
    End Sub

    Private Sub Button4_Click(sender As Object, e As EventArgs) Handles Button4.Click

    End Sub

    Private Sub Button4_BackColorChanged(sender As Object, e As EventArgs) Handles Button4.BackColorChanged
        If Button4.BackColor = Color.Green Then
            Button1.Enabled = True
        Else
            Button1.Enabled = False
        End If
    End Sub

    Private Sub Button2_Click(sender As Object, e As EventArgs) Handles Button2.Click
        Me.Close()
        Call SETTING.Show()
    End Sub
End Class