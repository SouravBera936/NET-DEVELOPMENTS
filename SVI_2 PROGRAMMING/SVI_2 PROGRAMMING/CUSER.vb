Public Class CUSER
    Private Sub CUSER_Load(sender As Object, e As EventArgs) Handles MyBase.Load
        ComboBox1.Enabled = True
        TextBox1.Enabled = False
        TextBox2.Enabled = False
        TextBox3.Enabled = False
        TextBox4.Enabled = False
        RadioButton1.Enabled = False
        RadioButton2.Enabled = False
        RadioButton3.Enabled = False
        Button1.Enabled = False
    End Sub

    Private Sub Button2_Click(sender As Object, e As EventArgs) Handles Button2.Click
        Me.Close()
        Call SETTING.Show()
    End Sub

    Private Sub ComboBox1_SelectedIndexChanged(sender As Object, e As EventArgs) Handles ComboBox1.SelectedIndexChanged
        Me.TextBox1.Enabled = True
    End Sub

    Private Sub ComboBox1_Click(sender As Object, e As EventArgs) Handles ComboBox1.Click
        Dim line1 As Label
        Dim wb As Object
        Dim xlapp As Object
        Dim st As Object
        xlapp = CreateObject("excel.application")
        wb = xlapp.workbooks.open(CONFIG.TextBox1.Text.ToString())
        st = wb.worksheets("Sheet1")
        If st.range("C2").value = "" Then
            ComboBox1.Items.Add(st.range("A2").value)
            GoTo line1
        Else
            If st.range("C3").value = "" Then
                ComboBox1.Items.Add(st.range("A3").value)
                GoTo line1
            Else
                If st.range("C4").value = "" Then
                    ComboBox1.Items.Add(st.range("A4").value)
                    GoTo line1
                Else
                    If st.range("C5").value = "" Then
                        ComboBox1.Items.Add(st.range("A5").value)
                        GoTo line1
                    Else
                        If st.range("C6").value = "" Then
                            ComboBox1.Items.Add(st.range("A6").value)
                            GoTo line1
                        Else
                            If st.range("C7").value = "" Then
                                ComboBox1.Items.Add(st.range("A7").value)
                                GoTo line1
                            Else
                                If st.range("C8").value = "" Then
                                    ComboBox1.Items.Add(st.range("A8").value)
                                    GoTo line1
                                Else
                                    If st.range("C9").value = "" Then
                                        ComboBox1.Items.Add(st.range("A9").value)
                                        GoTo line1
                                    Else
                                        If st.range("C10").value = "" Then
                                            ComboBox1.Items.Add(st.range("A10").value)
                                            GoTo line1
                                        Else
                                            If st.range("C11").value = "" Then
                                                ComboBox1.Items.Add(st.range("A11").value)
                                                GoTo line1
                                            Else
                                                If st.range("C12").value = "" Then
                                                    ComboBox1.Items.Add(st.range("A12").value)
                                                    GoTo line1
                                                Else
                                                    If st.range("C13").value = "" Then
                                                        ComboBox1.Items.Add(st.range("A13").value)
                                                        GoTo line1
                                                    Else
                                                        If st.range("C14").value = "" Then
                                                            ComboBox1.Items.Add(st.range("A14").value)
                                                            GoTo line1
                                                        Else
                                                            If st.range("C15").value = "" Then
                                                                ComboBox1.Items.Add(st.range("A15").value)
                                                                GoTo line1
                                                            Else
                                                                If st.range("C16").value = "" Then
                                                                    ComboBox1.Items.Add(st.range("A16").value)
                                                                    GoTo line1
                                                                Else
                                                                    If st.range("C17").value = "" Then
                                                                        ComboBox1.Items.Add(st.range("A17").value)
                                                                        GoTo line1
                                                                    Else
                                                                        If st.range("C18").value = "" Then
                                                                            ComboBox1.Items.Add(st.range("A18").value)
                                                                            GoTo line1
                                                                        Else
                                                                            If st.range("C19").value = "" Then
                                                                                ComboBox1.Items.Add(st.range("A19").value)
                                                                                GoTo line1
                                                                            Else
                                                                                If st.range("C20").value = "" Then
                                                                                    ComboBox1.Items.Add(st.range("A20").value)
                                                                                    GoTo line1
                                                                                Else
                                                                                    If st.range("C21").value = "" Then
                                                                                        ComboBox1.Items.Add(st.range("A21").value)
                                                                                        GoTo line1
                                                                                    Else
                                                                                        If st.range("C22").value = "" Then
                                                                                            ComboBox1.Items.Add(st.range("A22").value)
                                                                                            GoTo line1
                                                                                        Else
                                                                                            If st.range("C23").value = "" Then
                                                                                                ComboBox1.Items.Add(st.range("A23").value)
                                                                                                GoTo line1
                                                                                            Else
                                                                                                If st.range("C24").value = "" Then
                                                                                                    ComboBox1.Items.Add(st.range("A24").value)
                                                                                                    GoTo line1
                                                                                                Else
                                                                                                    If st.range("C25").value = "" Then
                                                                                                        ComboBox1.Items.Add(st.range("A25").value)
                                                                                                        GoTo line1
                                                                                                    Else
                                                                                                        If st.range("C26").value = "" Then
                                                                                                            ComboBox1.Items.Add(st.range("A26").value)
                                                                                                            GoTo line1
                                                                                                        Else
                                                                                                            If st.range("C27").value = "" Then
                                                                                                                ComboBox1.Items.Add(st.range("A27").value)
                                                                                                                GoTo line1
                                                                                                            Else
                                                                                                                If st.range("C28").value = "" Then
                                                                                                                    ComboBox1.Items.Add(st.range("A28").value)
                                                                                                                    GoTo line1
                                                                                                                Else
                                                                                                                    If st.range("C29").value = "" Then
                                                                                                                        ComboBox1.Items.Add(st.range("A29").value)
                                                                                                                        GoTo line1
                                                                                                                    Else
                                                                                                                        If st.range("C30").value = "" Then
                                                                                                                            ComboBox1.Items.Add(st.range("A30").value)
                                                                                                                            GoTo line1
                                                                                                                        Else
                                                                                                                            If st.range("C31").value = "" Then
                                                                                                                                ComboBox1.Items.Add(st.range("A31").value)
                                                                                                                                GoTo line1
                                                                                                                                Dim iret As Object = MsgBox("USER LIST IS OUT OF RANGE", vbCritical + vbOKOnly, "CUSER ERROR")
                                                                                                                                If iret = vbOK Then
                                                                                                                                    GoTo line1
                                                                                                                                    Me.Close()
                                                                                                                                    Call SETTING.Show()
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

    Private Sub TextBox1_TextChanged(sender As Object, e As EventArgs) Handles TextBox1.TextChanged
        TextBox2.Enabled = True
        If TextBox1.Text = "" Then
            TextBox2.Enabled = False
            Button1.Enabled = False
        End If
    End Sub

    Private Sub TextBox2_TextChanged(sender As Object, e As EventArgs) Handles TextBox2.TextChanged
        RadioButton1.Enabled = True
        RadioButton2.Enabled = True
        RadioButton3.Enabled = True
        If TextBox2.Text = "" Then
            RadioButton1.Enabled = False
            RadioButton2.Enabled = False
            RadioButton3.Enabled = False
            Button1.Enabled = False
        End If
    End Sub

    Private Sub RadioButton1_CheckedChanged(sender As Object, e As EventArgs) Handles RadioButton1.CheckedChanged
        If RadioButton1.Checked = True Then
            TextBox3.Enabled = True
        Else
            TextBox3.Enabled = False
            Button1.Enabled = False
        End If
    End Sub

    Private Sub RadioButton2_CheckedChanged(sender As Object, e As EventArgs) Handles RadioButton2.CheckedChanged
        If RadioButton2.Checked = True Then
            TextBox3.Enabled = True
        Else
            TextBox3.Enabled = False
            Button1.Enabled = False
        End If
    End Sub

    Private Sub RadioButton3_CheckedChanged(sender As Object, e As EventArgs) Handles RadioButton3.CheckedChanged
        If RadioButton3.Checked = True Then
            TextBox3.Enabled = True
        Else
            TextBox3.Enabled = False
            Button1.Enabled = False
        End If
    End Sub

    Private Sub TextBox3_TextChanged(sender As Object, e As EventArgs) Handles TextBox3.TextChanged
        TextBox3.PasswordChar = "X"
        TextBox4.Enabled = True
        If TextBox3.Text = "" Then
            TextBox4.Enabled = False
            Button1.Enabled = False
        End If
    End Sub

    Private Sub TextBox4_TextChanged(sender As Object, e As EventArgs) Handles TextBox4.TextChanged
        TextBox4.PasswordChar = "X"
        If TextBox3.Text = TextBox4.Text Then
            TextBox4.BackColor = Color.Green
            Button1.Enabled = True
        Else
            TextBox4.BackColor = Color.Red
            Button1.Enabled = False
        End If
    End Sub

    Private Sub Button1_Click(sender As Object, e As EventArgs) Handles Button1.Click
        Dim line1 As Label
        Dim wb As Object
        Dim xlapp As Object
        Dim st As Object
        xlapp = CreateObject("excel.application")
        wb = xlapp.workbooks.open(CONFIG.TextBox1.Text.ToString())
        st = wb.worksheets("Sheet1")
        Dim file As System.IO.StreamWriter
        Dim path As String = CONFIG.TextBox2.Text.ToString()
        file = My.Computer.FileSystem.OpenTextFileWriter(path, True)
        If ComboBox1.Text = st.range("A2").value Then
            st.unprotect("Test@123")
            st.range("A2").value = TextBox1.Text
            st.range("B2").value = TextBox2.Text
            st.range("C2").value = TextBox3.Text
            If RadioButton1.Checked = True Then
                st.range("D2").value = RadioButton1.Text
                Dim iret As Object = MsgBox("ACCESS GRANTED FOR" & " " & TextBox1.Text, vbInformation + vbOKOnly, "USER CREATED")
                file.WriteLine("USER CREATED :" & "=" & TextBox1.Text & "_" & DateTime.Now.ToString("yyyy/MM/dd/HH:mm:ss"))
                file.Close()
                If iret = vbOK Then
                    GoTo line1
                End If
            Else
                If RadioButton2.Checked = True Then
                    st.range("D2").value = RadioButton2.Text
                    Dim iret1 As Object = MsgBox("ACCESS GRANTED FOR" & " " & TextBox1.Text, vbInformation + vbOKOnly, "USER CREATED")
                    file.WriteLine("USER CREATED :" & "=" & TextBox1.Text & "_" & DateTime.Now.ToString("yyyy/MM/dd/HH:mm:ss"))
                    file.Close()
                    If iret1 = vbOK Then
                        GoTo line1
                    End If
                Else
                    If RadioButton3.Checked = True Then
                        st.range("D2").value = RadioButton3.Text
                        Dim iret2 As Object = MsgBox("ACCESS GRANTED FOR" & " " & TextBox1.Text, vbInformation + vbOKOnly, "USER CREATED")
                        file.WriteLine("USER CREATED :" & "=" & TextBox1.Text & "_" & DateTime.Now.ToString("yyyy/MM/dd/HH:mm:ss"))
                        file.Close()
                        If iret2 = vbOK Then
                            GoTo line1
                        End If
                    End If
                End If
            End If
        Else
            If ComboBox1.Text = st.range("A3").value Then
                st.unprotect("Test@123")
                st.range("A3").value = TextBox1.Text
                st.range("B3").value = TextBox2.Text
                st.range("C3").value = TextBox3.Text
                If RadioButton1.Checked = True Then
                    st.range("D3").value = RadioButton1.Text
                    Dim iret As Object = MsgBox("ACCESS GRANTED FOR" & " " & TextBox1.Text, vbInformation + vbOKOnly, "USER CREATED")
                    file.WriteLine("USER CREATED :" & "=" & TextBox1.Text & "_" & DateTime.Now.ToString("yyyy/MM/dd/HH:mm:ss"))
                    file.Close()
                    If iret = vbOK Then
                        GoTo line1
                    End If
                Else
                    If RadioButton2.Checked = True Then
                        st.range("D3").value = RadioButton2.Text
                        Dim iret1 As Object = MsgBox("ACCESS GRANTED FOR" & " " & TextBox1.Text, vbInformation + vbOKOnly, "USER CREATED")
                        file.WriteLine("USER CREATED :" & "=" & TextBox1.Text & "_" & DateTime.Now.ToString("yyyy/MM/dd/HH:mm:ss"))
                        file.Close()
                        If iret1 = vbOK Then
                            GoTo line1
                        End If
                    Else
                        If RadioButton3.Checked = True Then
                            st.range("D3").value = RadioButton3.Text
                            Dim iret2 As Object = MsgBox("ACCESS GRANTED FOR" & " " & TextBox1.Text, vbInformation + vbOKOnly, "USER CREATED")
                            file.WriteLine("USER CREATED :" & "=" & TextBox1.Text & "_" & DateTime.Now.ToString("yyyy/MM/dd/HH:mm:ss"))
                            file.Close()
                            If iret2 = vbOK Then
                                GoTo line1
                            End If
                        End If
                    End If
                End If
            Else
                If ComboBox1.Text = st.range("A4").value Then
                    st.unprotect("Test@123")
                    st.range("A4").value = TextBox1.Text
                    st.range("B4").value = TextBox2.Text
                    st.range("C4").value = TextBox3.Text
                    If RadioButton1.Checked = True Then
                        st.range("D4").value = RadioButton1.Text
                        Dim iret As Object = MsgBox("ACCESS GRANTED FOR" & " " & TextBox1.Text, vbInformation + vbOKOnly, "USER CREATED")
                        file.WriteLine("USER CREATED :" & "=" & TextBox1.Text & "_" & DateTime.Now.ToString("yyyy/MM/dd/HH:mm:ss"))
                        file.Close()
                        If iret = vbOK Then
                            GoTo line1
                        End If
                    Else
                        If RadioButton2.Checked = True Then
                            st.range("D4").value = RadioButton2.Text
                            Dim iret1 As Object = MsgBox("ACCESS GRANTED FOR" & " " & TextBox1.Text, vbInformation + vbOKOnly, "USER CREATED")
                            file.WriteLine("USER CREATED :" & "=" & TextBox1.Text & "_" & DateTime.Now.ToString("yyyy/MM/dd/HH:mm:ss"))
                            file.Close()
                            If iret1 = vbOK Then
                                GoTo line1
                            End If
                        Else
                            If RadioButton3.Checked = True Then
                                st.range("D4").value = RadioButton3.Text
                                Dim iret2 As Object = MsgBox("ACCESS GRANTED FOR" & " " & TextBox1.Text, vbInformation + vbOKOnly, "USER CREATED")
                                file.WriteLine("USER CREATED :" & "=" & TextBox1.Text & "_" & DateTime.Now.ToString("yyyy/MM/dd/HH:mm:ss"))
                                file.Close()
                                If iret2 = vbOK Then
                                    GoTo line1
                                End If
                            End If
                        End If
                    End If
                Else
                    If ComboBox1.Text = st.range("A5").value Then
                        st.unprotect("Test@123")
                        st.range("A5").value = TextBox1.Text
                        st.range("B5").value = TextBox2.Text
                        st.range("C5").value = TextBox3.Text
                        If RadioButton1.Checked = True Then
                            st.range("D5").value = RadioButton1.Text
                            Dim iret As Object = MsgBox("ACCESS GRANTED FOR" & " " & TextBox1.Text, vbInformation + vbOKOnly, "USER CREATED")
                            file.WriteLine("USER CREATED :" & "=" & TextBox1.Text & "_" & DateTime.Now.ToString("yyyy/MM/dd/HH:mm:ss"))
                            file.Close()
                            If iret = vbOK Then
                                GoTo line1
                            End If
                        Else
                            If RadioButton2.Checked = True Then
                                st.range("D5").value = RadioButton2.Text
                                Dim iret1 As Object = MsgBox("ACCESS GRANTED FOR" & " " & TextBox1.Text, vbInformation + vbOKOnly, "USER CREATED")
                                file.WriteLine("USER CREATED :" & "=" & TextBox1.Text & "_" & DateTime.Now.ToString("yyyy/MM/dd/HH:mm:ss"))
                                file.Close()
                                If iret1 = vbOK Then
                                    GoTo line1
                                End If
                            Else
                                If RadioButton3.Checked = True Then
                                    st.range("D5").value = RadioButton3.Text
                                    Dim iret2 As Object = MsgBox("ACCESS GRANTED FOR" & " " & TextBox1.Text, vbInformation + vbOKOnly, "USER CREATED")
                                    file.WriteLine("USER CREATED :" & "=" & TextBox1.Text & "_" & DateTime.Now.ToString("yyyy/MM/dd/HH:mm:ss"))
                                    file.Close()
                                    If iret2 = vbOK Then
                                        GoTo line1
                                    End If
                                End If
                            End If
                        End If
                    Else
                        If ComboBox1.Text = st.range("A6").value Then
                            st.unprotect("Test@123")
                            st.range("A6").value = TextBox1.Text
                            st.range("B6").value = TextBox2.Text
                            st.range("C6").value = TextBox3.Text
                            If RadioButton1.Checked = True Then
                                st.range("D6").value = RadioButton1.Text
                                Dim iret As Object = MsgBox("ACCESS GRANTED FOR" & " " & TextBox1.Text, vbInformation + vbOKOnly, "USER CREATED")
                                file.WriteLine("USER CREATED :" & "=" & TextBox1.Text & "_" & DateTime.Now.ToString("yyyy/MM/dd/HH:mm:ss"))
                                file.Close()
                                If iret = vbOK Then
                                    GoTo line1
                                End If
                            Else
                                If RadioButton2.Checked = True Then
                                    st.range("D6").value = RadioButton2.Text
                                    Dim iret1 As Object = MsgBox("ACCESS GRANTED FOR" & " " & TextBox1.Text, vbInformation + vbOKOnly, "USER CREATED")
                                    file.WriteLine("USER CREATED :" & "=" & TextBox1.Text & "_" & DateTime.Now.ToString("yyyy/MM/dd/HH:mm:ss"))
                                    file.Close()
                                    If iret1 = vbOK Then
                                        GoTo line1
                                    End If
                                Else
                                    If RadioButton3.Checked = True Then
                                        st.range("D6").value = RadioButton3.Text
                                        Dim iret2 As Object = MsgBox("ACCESS GRANTED FOR" & " " & TextBox1.Text, vbInformation + vbOKOnly, "USER CREATED")
                                        file.WriteLine("USER CREATED :" & "=" & TextBox1.Text & "_" & DateTime.Now.ToString("yyyy/MM/dd/HH:mm:ss"))
                                        file.Close()
                                        If iret2 = vbOK Then
                                            GoTo line1
                                        End If
                                    End If
                                End If
                            End If
                        Else
                            If ComboBox1.Text = st.range("A7").value Then
                                st.unprotect("Test@123")
                                st.range("A7").value = TextBox1.Text
                                st.range("B7").value = TextBox2.Text
                                st.range("C7").value = TextBox3.Text
                                If RadioButton1.Checked = True Then
                                    st.range("D7").value = RadioButton1.Text
                                    Dim iret As Object = MsgBox("ACCESS GRANTED FOR" & " " & TextBox1.Text, vbInformation + vbOKOnly, "USER CREATED")
                                    file.WriteLine("USER CREATED :" & "=" & TextBox1.Text & "_" & DateTime.Now.ToString("yyyy/MM/dd/HH:mm:ss"))
                                    file.Close()
                                    If iret = vbOK Then
                                        GoTo line1
                                    End If
                                Else
                                    If RadioButton2.Checked = True Then
                                        st.range("D7").value = RadioButton2.Text
                                        Dim iret1 As Object = MsgBox("ACCESS GRANTED FOR" & " " & TextBox1.Text, vbInformation + vbOKOnly, "USER CREATED")
                                        file.WriteLine("USER CREATED :" & "=" & TextBox1.Text & "_" & DateTime.Now.ToString("yyyy/MM/dd/HH:mm:ss"))
                                        file.Close()
                                        If iret1 = vbOK Then
                                            GoTo line1
                                        End If
                                    Else
                                        If RadioButton3.Checked = True Then
                                            st.range("D7").value = RadioButton3.Text
                                            Dim iret2 As Object = MsgBox("ACCESS GRANTED FOR" & " " & TextBox1.Text, vbInformation + vbOKOnly, "USER CREATED")
                                            file.WriteLine("USER CREATED :" & "=" & TextBox1.Text & "_" & DateTime.Now.ToString("yyyy/MM/dd/HH:mm:ss"))
                                            file.Close()
                                            If iret2 = vbOK Then
                                                GoTo line1
                                            End If
                                        End If
                                    End If
                                End If
                            Else
                                If ComboBox1.Text = st.range("A8").value Then
                                    st.unprotect("Test@123")
                                    st.range("A8").value = TextBox1.Text
                                    st.range("B8").value = TextBox2.Text
                                    st.range("C8").value = TextBox3.Text
                                    If RadioButton1.Checked = True Then
                                        st.range("D8").value = RadioButton1.Text
                                        Dim iret As Object = MsgBox("ACCESS GRANTED FOR" & " " & TextBox1.Text, vbInformation + vbOKOnly, "USER CREATED")
                                        file.WriteLine("USER CREATED :" & "=" & TextBox1.Text & "_" & DateTime.Now.ToString("yyyy/MM/dd/HH:mm:ss"))
                                        file.Close()
                                        If iret = vbOK Then
                                            GoTo line1
                                        End If
                                    Else
                                        If RadioButton2.Checked = True Then
                                            st.range("D8").value = RadioButton2.Text
                                            Dim iret1 As Object = MsgBox("ACCESS GRANTED FOR" & " " & TextBox1.Text, vbInformation + vbOKOnly, "USER CREATED")
                                            file.WriteLine("USER CREATED :" & "=" & TextBox1.Text & "_" & DateTime.Now.ToString("yyyy/MM/dd/HH:mm:ss"))
                                            file.Close()
                                            If iret1 = vbOK Then
                                                GoTo line1
                                            End If
                                        Else
                                            If RadioButton3.Checked = True Then
                                                st.range("D8").value = RadioButton3.Text
                                                Dim iret2 As Object = MsgBox("ACCESS GRANTED FOR" & " " & TextBox1.Text, vbInformation + vbOKOnly, "USER CREATED")
                                                file.WriteLine("USER CREATED :" & "=" & TextBox1.Text & "_" & DateTime.Now.ToString("yyyy/MM/dd/HH:mm:ss"))
                                                file.Close()
                                                If iret2 = vbOK Then
                                                    GoTo line1
                                                End If
                                            End If
                                        End If
                                    End If
                                Else
                                    If ComboBox1.Text = st.range("A9").value Then
                                        st.unprotect("Test@123")
                                        st.range("A9").value = TextBox1.Text
                                        st.range("B9").value = TextBox2.Text
                                        st.range("C9").value = TextBox3.Text
                                        If RadioButton1.Checked = True Then
                                            st.range("D9").value = RadioButton1.Text
                                            Dim iret As Object = MsgBox("ACCESS GRANTED FOR" & " " & TextBox1.Text, vbInformation + vbOKOnly, "USER CREATED")
                                            file.WriteLine("USER CREATED :" & "=" & TextBox1.Text & "_" & DateTime.Now.ToString("yyyy/MM/dd/HH:mm:ss"))
                                            file.Close()
                                            If iret = vbOK Then
                                                GoTo line1
                                            End If
                                        Else
                                            If RadioButton2.Checked = True Then
                                                st.range("D9").value = RadioButton2.Text
                                                Dim iret1 As Object = MsgBox("ACCESS GRANTED FOR" & " " & TextBox1.Text, vbInformation + vbOKOnly, "USER CREATED")
                                                file.WriteLine("USER CREATED :" & "=" & TextBox1.Text & "_" & DateTime.Now.ToString("yyyy/MM/dd/HH:mm:ss"))
                                                file.Close()
                                                If iret1 = vbOK Then
                                                    GoTo line1
                                                End If
                                            Else
                                                If RadioButton3.Checked = True Then
                                                    st.range("D9").value = RadioButton3.Text
                                                    Dim iret2 As Object = MsgBox("ACCESS GRANTED FOR" & " " & TextBox1.Text, vbInformation + vbOKOnly, "USER CREATED")
                                                    file.WriteLine("USER CREATED :" & "=" & TextBox1.Text & "_" & DateTime.Now.ToString("yyyy/MM/dd/HH:mm:ss"))
                                                    file.Close()
                                                    If iret2 = vbOK Then
                                                        GoTo line1
                                                    End If
                                                End If
                                            End If
                                        End If
                                    Else
                                        If ComboBox1.Text = st.range("A10").value Then
                                            st.unprotect("Test@123")
                                            st.range("A10").value = TextBox1.Text
                                            st.range("B10").value = TextBox2.Text
                                            st.range("C10").value = TextBox3.Text
                                            If RadioButton1.Checked = True Then
                                                st.range("D10").value = RadioButton1.Text
                                                Dim iret As Object = MsgBox("ACCESS GRANTED FOR" & " " & TextBox1.Text, vbInformation + vbOKOnly, "USER CREATED")
                                                file.WriteLine("USER CREATED :" & "=" & TextBox1.Text & "_" & DateTime.Now.ToString("yyyy/MM/dd/HH:mm:ss"))
                                                file.Close()
                                                If iret = vbOK Then
                                                    GoTo line1
                                                End If
                                            Else
                                                If RadioButton2.Checked = True Then
                                                    st.range("D10").value = RadioButton2.Text
                                                    Dim iret1 As Object = MsgBox("ACCESS GRANTED FOR" & " " & TextBox1.Text, vbInformation + vbOKOnly, "USER CREATED")
                                                    file.WriteLine("USER CREATED :" & "=" & TextBox1.Text & "_" & DateTime.Now.ToString("yyyy/MM/dd/HH:mm:ss"))
                                                    file.Close()
                                                    If iret1 = vbOK Then
                                                        GoTo line1
                                                    End If
                                                Else
                                                    If RadioButton3.Checked = True Then
                                                        st.range("D10").value = RadioButton3.Text
                                                        Dim iret2 As Object = MsgBox("ACCESS GRANTED FOR" & " " & TextBox1.Text, vbInformation + vbOKOnly, "USER CREATED")
                                                        file.WriteLine("USER CREATED :" & "=" & TextBox1.Text & "_" & DateTime.Now.ToString("yyyy/MM/dd/HH:mm:ss"))
                                                        file.Close()
                                                        If iret2 = vbOK Then
                                                            GoTo line1
                                                        End If
                                                    End If
                                                End If
                                            End If
                                        Else
                                            If ComboBox1.Text = st.range("A11").value Then
                                                st.unprotect("Test@123")
                                                st.range("A11").value = TextBox1.Text
                                                st.range("B11").value = TextBox2.Text
                                                st.range("C11").value = TextBox3.Text
                                                If RadioButton1.Checked = True Then
                                                    st.range("D11").value = RadioButton1.Text
                                                    Dim iret As Object = MsgBox("ACCESS GRANTED FOR" & " " & TextBox1.Text, vbInformation + vbOKOnly, "USER CREATED")
                                                    file.WriteLine("USER CREATED :" & "=" & TextBox1.Text & "_" & DateTime.Now.ToString("yyyy/MM/dd/HH:mm:ss"))
                                                    file.Close()
                                                    If iret = vbOK Then
                                                        GoTo line1
                                                    End If
                                                Else
                                                    If RadioButton2.Checked = True Then
                                                        st.range("D11").value = RadioButton2.Text
                                                        Dim iret1 As Object = MsgBox("ACCESS GRANTED FOR" & " " & TextBox1.Text, vbInformation + vbOKOnly, "USER CREATED")
                                                        file.WriteLine("USER CREATED :" & "=" & TextBox1.Text & "_" & DateTime.Now.ToString("yyyy/MM/dd/HH:mm:ss"))
                                                        file.Close()
                                                        If iret1 = vbOK Then
                                                            GoTo line1
                                                        End If
                                                    Else
                                                        If RadioButton3.Checked = True Then
                                                            st.range("D11").value = RadioButton3.Text
                                                            Dim iret2 As Object = MsgBox("ACCESS GRANTED FOR" & " " & TextBox1.Text, vbInformation + vbOKOnly, "USER CREATED")
                                                            file.WriteLine("USER CREATED :" & "=" & TextBox1.Text & "_" & DateTime.Now.ToString("yyyy/MM/dd/HH:mm:ss"))
                                                            file.Close()
                                                            If iret2 = vbOK Then
                                                                GoTo line1
                                                            End If
                                                        End If
                                                    End If
                                                End If
                                            Else
                                                If ComboBox1.Text = st.range("A12").value Then
                                                    st.unprotect("Test@123")
                                                    st.range("A12").value = TextBox1.Text
                                                    st.range("B12").value = TextBox2.Text
                                                    st.range("C12").value = TextBox3.Text
                                                    If RadioButton1.Checked = True Then
                                                        st.range("D12").value = RadioButton1.Text
                                                        Dim iret As Object = MsgBox("ACCESS GRANTED FOR" & " " & TextBox1.Text, vbInformation + vbOKOnly, "USER CREATED")
                                                        file.WriteLine("USER CREATED :" & "=" & TextBox1.Text & "_" & DateTime.Now.ToString("yyyy/MM/dd/HH:mm:ss"))
                                                        file.Close()
                                                        If iret = vbOK Then
                                                            GoTo line1
                                                        End If
                                                    Else
                                                        If RadioButton2.Checked = True Then
                                                            st.range("D12").value = RadioButton2.Text
                                                            Dim iret1 As Object = MsgBox("ACCESS GRANTED FOR" & " " & TextBox1.Text, vbInformation + vbOKOnly, "USER CREATED")
                                                            file.WriteLine("USER CREATED :" & "=" & TextBox1.Text & "_" & DateTime.Now.ToString("yyyy/MM/dd/HH:mm:ss"))
                                                            file.Close()
                                                            If iret1 = vbOK Then
                                                                GoTo line1
                                                            End If
                                                        Else
                                                            If RadioButton3.Checked = True Then
                                                                st.range("D12").value = RadioButton3.Text
                                                                Dim iret2 As Object = MsgBox("ACCESS GRANTED FOR" & " " & TextBox1.Text, vbInformation + vbOKOnly, "USER CREATED")
                                                                file.WriteLine("USER CREATED :" & "=" & TextBox1.Text & "_" & DateTime.Now.ToString("yyyy/MM/dd/HH:mm:ss"))
                                                                file.Close()
                                                                If iret2 = vbOK Then
                                                                    GoTo line1
                                                                End If
                                                            End If
                                                        End If
                                                    End If
                                                Else
                                                    If ComboBox1.Text = st.range("A13").value Then
                                                        st.unprotect("Test@123")
                                                        st.range("A13").value = TextBox1.Text
                                                        st.range("B13").value = TextBox2.Text
                                                        st.range("C13").value = TextBox3.Text
                                                        If RadioButton1.Checked = True Then
                                                            st.range("D13").value = RadioButton1.Text
                                                            Dim iret As Object = MsgBox("ACCESS GRANTED FOR" & " " & TextBox1.Text, vbInformation + vbOKOnly, "USER CREATED")
                                                            file.WriteLine("USER CREATED :" & "=" & TextBox1.Text & "_" & DateTime.Now.ToString("yyyy/MM/dd/HH:mm:ss"))
                                                            file.Close()
                                                            If iret = vbOK Then
                                                                GoTo line1
                                                            End If
                                                        Else
                                                            If RadioButton2.Checked = True Then
                                                                st.range("D13").value = RadioButton2.Text
                                                                Dim iret1 As Object = MsgBox("ACCESS GRANTED FOR" & " " & TextBox1.Text, vbInformation + vbOKOnly, "USER CREATED")
                                                                file.WriteLine("USER CREATED :" & "=" & TextBox1.Text & "_" & DateTime.Now.ToString("yyyy/MM/dd/HH:mm:ss"))
                                                                file.Close()
                                                                If iret1 = vbOK Then
                                                                    GoTo line1
                                                                End If
                                                            Else
                                                                If RadioButton3.Checked = True Then
                                                                    st.range("D13").value = RadioButton3.Text
                                                                    Dim iret2 As Object = MsgBox("ACCESS GRANTED FOR" & " " & TextBox1.Text, vbInformation + vbOKOnly, "USER CREATED")
                                                                    file.WriteLine("USER CREATED :" & "=" & TextBox1.Text & "_" & DateTime.Now.ToString("yyyy/MM/dd/HH:mm:ss"))
                                                                    file.Close()
                                                                    If iret2 = vbOK Then
                                                                        GoTo line1
                                                                    End If
                                                                End If
                                                            End If
                                                        End If
                                                    Else
                                                        If ComboBox1.Text = st.range("A14").value Then
                                                            st.unprotect("Test@123")
                                                            st.range("A14").value = TextBox1.Text
                                                            st.range("B14").value = TextBox2.Text
                                                            st.range("C14").value = TextBox3.Text
                                                            If RadioButton1.Checked = True Then
                                                                st.range("D14").value = RadioButton1.Text
                                                                Dim iret As Object = MsgBox("ACCESS GRANTED FOR" & " " & TextBox1.Text, vbInformation + vbOKOnly, "USER CREATED")
                                                                file.WriteLine("USER CREATED :" & "=" & TextBox1.Text & "_" & DateTime.Now.ToString("yyyy/MM/dd/HH:mm:ss"))
                                                                file.Close()
                                                                If iret = vbOK Then
                                                                    GoTo line1
                                                                End If
                                                            Else
                                                                If RadioButton2.Checked = True Then
                                                                    st.range("D14").value = RadioButton2.Text
                                                                    Dim iret1 As Object = MsgBox("ACCESS GRANTED FOR" & " " & TextBox1.Text, vbInformation + vbOKOnly, "USER CREATED")
                                                                    file.WriteLine("USER CREATED :" & "=" & TextBox1.Text & "_" & DateTime.Now.ToString("yyyy/MM/dd/HH:mm:ss"))
                                                                    file.Close()
                                                                    If iret1 = vbOK Then
                                                                        GoTo line1
                                                                    End If
                                                                Else
                                                                    If RadioButton3.Checked = True Then
                                                                        st.range("D14").value = RadioButton3.Text
                                                                        Dim iret2 As Object = MsgBox("ACCESS GRANTED FOR" & " " & TextBox1.Text, vbInformation + vbOKOnly, "USER CREATED")
                                                                        file.WriteLine("USER CREATED :" & "=" & TextBox1.Text & "_" & DateTime.Now.ToString("yyyy/MM/dd/HH:mm:ss"))
                                                                        file.Close()
                                                                        If iret2 = vbOK Then
                                                                            GoTo line1
                                                                        End If
                                                                    End If
                                                                End If
                                                            End If
                                                        Else
                                                            If ComboBox1.Text = st.range("A15").value Then
                                                                st.range("A15").value = TextBox1.Text
                                                                st.range("B15").value = TextBox2.Text
                                                                st.range("C15").value = TextBox3.Text
                                                                If RadioButton1.Checked = True Then
                                                                    st.range("D15").value = RadioButton1.Text
                                                                    Dim iret As Object = MsgBox("ACCESS GRANTED FOR" & " " & TextBox1.Text, vbInformation + vbOKOnly, "USER CREATED")
                                                                    file.WriteLine("USER CREATED :" & "=" & TextBox1.Text & "_" & DateTime.Now.ToString("yyyy/MM/dd/HH:mm:ss"))
                                                                    file.Close()
                                                                    If iret = vbOK Then
                                                                        GoTo line1
                                                                    End If
                                                                Else
                                                                    If RadioButton2.Checked = True Then
                                                                        st.range("D15").value = RadioButton2.Text
                                                                        Dim iret1 As Object = MsgBox("ACCESS GRANTED FOR" & " " & TextBox1.Text, vbInformation + vbOKOnly, "USER CREATED")
                                                                        file.WriteLine("USER CREATED :" & "=" & TextBox1.Text & "_" & DateTime.Now.ToString("yyyy/MM/dd/HH:mm:ss"))
                                                                        file.Close()
                                                                        If iret1 = vbOK Then
                                                                            GoTo line1
                                                                        End If
                                                                    Else
                                                                        If RadioButton3.Checked = True Then
                                                                            st.range("D15").value = RadioButton3.Text
                                                                            Dim iret2 As Object = MsgBox("ACCESS GRANTED FOR" & " " & TextBox1.Text, vbInformation + vbOKOnly, "USER CREATED")
                                                                            file.WriteLine("USER CREATED :" & "=" & TextBox1.Text & "_" & DateTime.Now.ToString("yyyy/MM/dd/HH:mm:ss"))
                                                                            file.Close()
                                                                            If iret2 = vbOK Then
                                                                                GoTo line1
                                                                            End If
                                                                        End If
                                                                    End If
                                                                End If
                                                            Else
                                                                If ComboBox1.Text = st.range("A16").value Then
                                                                    st.unprotect("Test@123")
                                                                    st.range("A16").value = TextBox1.Text
                                                                    st.range("B16").value = TextBox2.Text
                                                                    st.range("C16").value = TextBox3.Text
                                                                    If RadioButton1.Checked = True Then
                                                                        st.range("D16").value = RadioButton1.Text
                                                                        Dim iret As Object = MsgBox("ACCESS GRANTED FOR" & " " & TextBox1.Text, vbInformation + vbOKOnly, "USER CREATED")
                                                                        file.WriteLine("USER CREATED :" & "=" & TextBox1.Text & "_" & DateTime.Now.ToString("yyyy/MM/dd/HH:mm:ss"))
                                                                        file.Close()
                                                                        If iret = vbOK Then
                                                                            GoTo line1
                                                                        End If
                                                                    Else
                                                                        If RadioButton2.Checked = True Then
                                                                            st.range("D16").value = RadioButton2.Text
                                                                            Dim iret1 As Object = MsgBox("ACCESS GRANTED FOR" & " " & TextBox1.Text, vbInformation + vbOKOnly, "USER CREATED")
                                                                            file.WriteLine("USER CREATED :" & "=" & TextBox1.Text & "_" & DateTime.Now.ToString("yyyy/MM/dd/HH:mm:ss"))
                                                                            file.Close()
                                                                            If iret1 = vbOK Then
                                                                                GoTo line1
                                                                            End If
                                                                        Else
                                                                            If RadioButton3.Checked = True Then
                                                                                st.range("D16").value = RadioButton3.Text
                                                                                Dim iret2 As Object = MsgBox("ACCESS GRANTED FOR" & " " & TextBox1.Text, vbInformation + vbOKOnly, "USER CREATED")
                                                                                file.WriteLine("USER CREATED :" & "=" & TextBox1.Text & "_" & DateTime.Now.ToString("yyyy/MM/dd/HH:mm:ss"))
                                                                                file.Close()
                                                                                If iret2 = vbOK Then
                                                                                    GoTo line1
                                                                                End If
                                                                            End If
                                                                        End If
                                                                    End If
                                                                Else
                                                                    If ComboBox1.Text = st.range("A17").value Then
                                                                        st.unprotect("Test@123")
                                                                        st.range("A17").value = TextBox1.Text
                                                                        st.range("B17").value = TextBox2.Text
                                                                        st.range("C17").value = TextBox3.Text
                                                                        If RadioButton1.Checked = True Then
                                                                            st.range("D17").value = RadioButton1.Text
                                                                            Dim iret As Object = MsgBox("ACCESS GRANTED FOR" & " " & TextBox1.Text, vbInformation + vbOKOnly, "USER CREATED")
                                                                            file.WriteLine("USER CREATED :" & "=" & TextBox1.Text & "_" & DateTime.Now.ToString("yyyy/MM/dd/HH:mm:ss"))
                                                                            file.Close()
                                                                            If iret = vbOK Then
                                                                                GoTo line1
                                                                            End If
                                                                        Else
                                                                            If RadioButton2.Checked = True Then
                                                                                st.range("D17").value = RadioButton2.Text
                                                                                Dim iret1 As Object = MsgBox("ACCESS GRANTED FOR" & " " & TextBox1.Text, vbInformation + vbOKOnly, "USER CREATED")
                                                                                file.WriteLine("USER CREATED :" & "=" & TextBox1.Text & "_" & DateTime.Now.ToString("yyyy/MM/dd/HH:mm:ss"))
                                                                                file.Close()
                                                                                If iret1 = vbOK Then
                                                                                    GoTo line1
                                                                                End If
                                                                            Else
                                                                                If RadioButton3.Checked = True Then
                                                                                    st.range("D17").value = RadioButton3.Text
                                                                                    Dim iret2 As Object = MsgBox("ACCESS GRANTED FOR" & " " & TextBox1.Text, vbInformation + vbOKOnly, "USER CREATED")
                                                                                    file.WriteLine("USER CREATED :" & "=" & TextBox1.Text & "_" & DateTime.Now.ToString("yyyy/MM/dd/HH:mm:ss"))
                                                                                    file.Close()
                                                                                    If iret2 = vbOK Then
                                                                                        GoTo line1
                                                                                    End If
                                                                                End If
                                                                            End If
                                                                        End If
                                                                    Else
                                                                        If ComboBox1.Text = st.range("A18").value Then
                                                                            st.unprotect("Test@123")
                                                                            st.range("A18").value = TextBox1.Text
                                                                            st.range("B18").value = TextBox2.Text
                                                                            st.range("C18").value = TextBox3.Text
                                                                            If RadioButton1.Checked = True Then
                                                                                st.range("D18").value = RadioButton1.Text
                                                                                Dim iret As Object = MsgBox("ACCESS GRANTED FOR" & " " & TextBox1.Text, vbInformation + vbOKOnly, "USER CREATED")
                                                                                file.WriteLine("USER CREATED :" & "=" & TextBox1.Text & "_" & DateTime.Now.ToString("yyyy/MM/dd/HH:mm:ss"))
                                                                                file.Close()
                                                                                If iret = vbOK Then
                                                                                    GoTo line1
                                                                                End If
                                                                            Else
                                                                                If RadioButton2.Checked = True Then
                                                                                    st.range("D18").value = RadioButton2.Text
                                                                                    Dim iret1 As Object = MsgBox("ACCESS GRANTED FOR" & " " & TextBox1.Text, vbInformation + vbOKOnly, "USER CREATED")
                                                                                    file.WriteLine("USER CREATED :" & "=" & TextBox1.Text & "_" & DateTime.Now.ToString("yyyy/MM/dd/HH:mm:ss"))
                                                                                    file.Close()
                                                                                    If iret1 = vbOK Then
                                                                                        GoTo line1
                                                                                    End If
                                                                                Else
                                                                                    If RadioButton3.Checked = True Then
                                                                                        st.range("D18").value = RadioButton3.Text
                                                                                        Dim iret2 As Object = MsgBox("ACCESS GRANTED FOR" & " " & TextBox1.Text, vbInformation + vbOKOnly, "USER CREATED")
                                                                                        file.WriteLine("USER CREATED :" & "=" & TextBox1.Text & "_" & DateTime.Now.ToString("yyyy/MM/dd/HH:mm:ss"))
                                                                                        file.Close()
                                                                                        If iret2 = vbOK Then
                                                                                            GoTo line1
                                                                                        End If
                                                                                    End If
                                                                                End If
                                                                            End If
                                                                        Else
                                                                            If ComboBox1.Text = st.range("A19").value Then
                                                                                st.unprotect("Test@123")
                                                                                st.range("A19").value = TextBox1.Text
                                                                                st.range("B19").value = TextBox2.Text
                                                                                st.range("C19").value = TextBox3.Text
                                                                                If RadioButton1.Checked = True Then
                                                                                    st.range("D19").value = RadioButton1.Text
                                                                                    Dim iret As Object = MsgBox("ACCESS GRANTED FOR" & " " & TextBox1.Text, vbInformation + vbOKOnly, "USER CREATED")
                                                                                    file.WriteLine("USER CREATED :" & "=" & TextBox1.Text & "_" & DateTime.Now.ToString("yyyy/MM/dd/HH:mm:ss"))
                                                                                    file.Close()
                                                                                    If iret = vbOK Then
                                                                                        GoTo line1
                                                                                    End If
                                                                                Else
                                                                                    If RadioButton2.Checked = True Then
                                                                                        st.range("D19").value = RadioButton2.Text
                                                                                        Dim iret1 As Object = MsgBox("ACCESS GRANTED FOR" & " " & TextBox1.Text, vbInformation + vbOKOnly, "USER CREATED")
                                                                                        file.WriteLine("USER CREATED :" & "=" & TextBox1.Text & "_" & DateTime.Now.ToString("yyyy/MM/dd/HH:mm:ss"))
                                                                                        file.Close()
                                                                                        If iret1 = vbOK Then
                                                                                            GoTo line1
                                                                                        End If
                                                                                    Else
                                                                                        If RadioButton3.Checked = True Then
                                                                                            st.range("D19").value = RadioButton3.Text
                                                                                            Dim iret2 As Object = MsgBox("ACCESS GRANTED FOR" & " " & TextBox1.Text, vbInformation + vbOKOnly, "USER CREATED")
                                                                                            file.WriteLine("USER CREATED :" & "=" & TextBox1.Text & "_" & DateTime.Now.ToString("yyyy/MM/dd/HH:mm:ss"))
                                                                                            file.Close()
                                                                                            If iret2 = vbOK Then
                                                                                                GoTo line1
                                                                                            End If
                                                                                        End If
                                                                                    End If
                                                                                End If
                                                                            Else
                                                                                If ComboBox1.Text = st.range("A20").value Then
                                                                                    st.unprotect("Test@123")
                                                                                    st.range("A20").value = TextBox1.Text
                                                                                    st.range("B20").value = TextBox2.Text
                                                                                    st.range("C20").value = TextBox3.Text
                                                                                    If RadioButton1.Checked = True Then
                                                                                        st.range("D20").value = RadioButton1.Text
                                                                                        Dim iret As Object = MsgBox("ACCESS GRANTED FOR" & " " & TextBox1.Text, vbInformation + vbOKOnly, "USER CREATED")
                                                                                        file.WriteLine("USER CREATED :" & "=" & TextBox1.Text & "_" & DateTime.Now.ToString("yyyy/MM/dd/HH:mm:ss"))
                                                                                        file.Close()
                                                                                        If iret = vbOK Then
                                                                                            GoTo line1
                                                                                        End If
                                                                                    Else
                                                                                        If RadioButton2.Checked = True Then
                                                                                            st.range("D20").value = RadioButton2.Text
                                                                                            Dim iret1 As Object = MsgBox("ACCESS GRANTED FOR" & " " & TextBox1.Text, vbInformation + vbOKOnly, "USER CREATED")
                                                                                            file.WriteLine("USER CREATED :" & "=" & TextBox1.Text & "_" & DateTime.Now.ToString("yyyy/MM/dd/HH:mm:ss"))
                                                                                            file.Close()
                                                                                            If iret1 = vbOK Then
                                                                                                GoTo line1
                                                                                            End If
                                                                                        Else
                                                                                            If RadioButton3.Checked = True Then
                                                                                                st.range("D20").value = RadioButton3.Text
                                                                                                Dim iret2 As Object = MsgBox("ACCESS GRANTED FOR" & " " & TextBox1.Text, vbInformation + vbOKOnly, "USER CREATED")
                                                                                                file.WriteLine("USER CREATED :" & "=" & TextBox1.Text & "_" & DateTime.Now.ToString("yyyy/MM/dd/HH:mm:ss"))
                                                                                                file.Close()
                                                                                                If iret2 = vbOK Then
                                                                                                    GoTo line1
                                                                                                End If
                                                                                            End If
                                                                                        End If
                                                                                    End If
                                                                                Else
                                                                                    If ComboBox1.Text = st.range("A21").value Then
                                                                                        st.unprotect("Test@123")
                                                                                        st.range("A21").value = TextBox1.Text
                                                                                        st.range("B21").value = TextBox2.Text
                                                                                        st.range("C21").value = TextBox3.Text
                                                                                        If RadioButton1.Checked = True Then
                                                                                            st.range("D21").value = RadioButton1.Text
                                                                                            Dim iret As Object = MsgBox("ACCESS GRANTED FOR" & " " & TextBox1.Text, vbInformation + vbOKOnly, "USER CREATED")
                                                                                            file.WriteLine("USER CREATED :" & "=" & TextBox1.Text & "_" & DateTime.Now.ToString("yyyy/MM/dd/HH:mm:ss"))
                                                                                            file.Close()
                                                                                            If iret = vbOK Then
                                                                                                GoTo line1
                                                                                            End If
                                                                                        Else
                                                                                            If RadioButton2.Checked = True Then
                                                                                                st.range("D21").value = RadioButton2.Text
                                                                                                Dim iret1 As Object = MsgBox("ACCESS GRANTED FOR" & " " & TextBox1.Text, vbInformation + vbOKOnly, "USER CREATED")
                                                                                                file.WriteLine("USER CREATED :" & "=" & TextBox1.Text & "_" & DateTime.Now.ToString("yyyy/MM/dd/HH:mm:ss"))
                                                                                                file.Close()
                                                                                                If iret1 = vbOK Then
                                                                                                    GoTo line1
                                                                                                End If
                                                                                            Else
                                                                                                If RadioButton3.Checked = True Then
                                                                                                    st.range("D21").value = RadioButton3.Text
                                                                                                    Dim iret2 As Object = MsgBox("ACCESS GRANTED FOR" & " " & TextBox1.Text, vbInformation + vbOKOnly, "USER CREATED")
                                                                                                    file.WriteLine("USER CREATED :" & "=" & TextBox1.Text & "_" & DateTime.Now.ToString("yyyy/MM/dd/HH:mm:ss"))
                                                                                                    file.Close()
                                                                                                    If iret2 = vbOK Then
                                                                                                        GoTo line1
                                                                                                    End If
                                                                                                End If
                                                                                            End If
                                                                                        End If
                                                                                    Else
                                                                                        If ComboBox1.Text = st.range("A22").value Then
                                                                                            st.unprotect("Test@123")
                                                                                            st.range("A22").value = TextBox1.Text
                                                                                            st.range("B22").value = TextBox2.Text
                                                                                            st.range("C22").value = TextBox3.Text
                                                                                            If RadioButton1.Checked = True Then
                                                                                                st.range("D22").value = RadioButton1.Text
                                                                                                Dim iret As Object = MsgBox("ACCESS GRANTED FOR" & " " & TextBox1.Text, vbInformation + vbOKOnly, "USER CREATED")
                                                                                                file.WriteLine("USER CREATED :" & "=" & TextBox1.Text & "_" & DateTime.Now.ToString("yyyy/MM/dd/HH:mm:ss"))
                                                                                                file.Close()
                                                                                                If iret = vbOK Then
                                                                                                    GoTo line1
                                                                                                End If
                                                                                            Else
                                                                                                If RadioButton2.Checked = True Then
                                                                                                    st.range("D22").value = RadioButton2.Text
                                                                                                    Dim iret1 As Object = MsgBox("ACCESS GRANTED FOR" & " " & TextBox1.Text, vbInformation + vbOKOnly, "USER CREATED")
                                                                                                    file.WriteLine("USER CREATED :" & "=" & TextBox1.Text & "_" & DateTime.Now.ToString("yyyy/MM/dd/HH:mm:ss"))
                                                                                                    file.Close()
                                                                                                    If iret1 = vbOK Then
                                                                                                        GoTo line1
                                                                                                    End If
                                                                                                Else
                                                                                                    If RadioButton3.Checked = True Then
                                                                                                        st.range("D22").value = RadioButton3.Text
                                                                                                        Dim iret2 As Object = MsgBox("ACCESS GRANTED FOR" & " " & TextBox1.Text, vbInformation + vbOKOnly, "USER CREATED")
                                                                                                        file.WriteLine("USER CREATED :" & "=" & TextBox1.Text & "_" & DateTime.Now.ToString("yyyy/MM/dd/HH:mm:ss"))
                                                                                                        file.Close()
                                                                                                        If iret2 = vbOK Then
                                                                                                            GoTo line1
                                                                                                        End If
                                                                                                    End If
                                                                                                End If
                                                                                            End If
                                                                                        Else
                                                                                            If ComboBox1.Text = st.range("A23").value Then
                                                                                                st.unprotect("Test@123")
                                                                                                st.range("A23").value = TextBox1.Text
                                                                                                st.range("B23").value = TextBox2.Text
                                                                                                st.range("C23").value = TextBox3.Text
                                                                                                If RadioButton1.Checked = True Then
                                                                                                    st.range("D23").value = RadioButton1.Text
                                                                                                    Dim iret As Object = MsgBox("ACCESS GRANTED FOR" & " " & TextBox1.Text, vbInformation + vbOKOnly, "USER CREATED")
                                                                                                    file.WriteLine("USER CREATED :" & "=" & TextBox1.Text & "_" & DateTime.Now.ToString("yyyy/MM/dd/HH:mm:ss"))
                                                                                                    file.Close()
                                                                                                    If iret = vbOK Then
                                                                                                        GoTo line1
                                                                                                    End If
                                                                                                Else
                                                                                                    If RadioButton2.Checked = True Then
                                                                                                        st.range("D23").value = RadioButton2.Text
                                                                                                        Dim iret1 As Object = MsgBox("ACCESS GRANTED FOR" & " " & TextBox1.Text, vbInformation + vbOKOnly, "USER CREATED")
                                                                                                        file.WriteLine("USER CREATED :" & "=" & TextBox1.Text & "_" & DateTime.Now.ToString("yyyy/MM/dd/HH:mm:ss"))
                                                                                                        file.Close()
                                                                                                        If iret1 = vbOK Then
                                                                                                            GoTo line1
                                                                                                        End If
                                                                                                    Else
                                                                                                        If RadioButton3.Checked = True Then
                                                                                                            st.range("D23").value = RadioButton3.Text
                                                                                                            Dim iret2 As Object = MsgBox("ACCESS GRANTED FOR" & " " & TextBox1.Text, vbInformation + vbOKOnly, "USER CREATED")
                                                                                                            file.WriteLine("USER CREATED :" & "=" & TextBox1.Text & "_" & DateTime.Now.ToString("yyyy/MM/dd/HH:mm:ss"))
                                                                                                            file.Close()
                                                                                                            If iret2 = vbOK Then
                                                                                                                GoTo line1
                                                                                                            End If
                                                                                                        End If
                                                                                                    End If
                                                                                                End If
                                                                                            Else
                                                                                                If ComboBox1.Text = st.range("A24").value Then
                                                                                                    st.unprotect("Test@123")
                                                                                                    st.range("A24").value = TextBox1.Text
                                                                                                    st.range("B24").value = TextBox2.Text
                                                                                                    st.range("C24").value = TextBox3.Text
                                                                                                    If RadioButton1.Checked = True Then
                                                                                                        st.range("D24").value = RadioButton1.Text
                                                                                                        Dim iret As Object = MsgBox("ACCESS GRANTED FOR" & " " & TextBox1.Text, vbInformation + vbOKOnly, "USER CREATED")
                                                                                                        file.WriteLine("USER CREATED :" & "=" & TextBox1.Text & "_" & DateTime.Now.ToString("yyyy/MM/dd/HH:mm:ss"))
                                                                                                        file.Close()
                                                                                                        If iret = vbOK Then
                                                                                                            GoTo line1
                                                                                                        End If
                                                                                                    Else
                                                                                                        If RadioButton2.Checked = True Then
                                                                                                            st.range("D24").value = RadioButton2.Text
                                                                                                            Dim iret1 As Object = MsgBox("ACCESS GRANTED FOR" & " " & TextBox1.Text, vbInformation + vbOKOnly, "USER CREATED")
                                                                                                            file.WriteLine("USER CREATED :" & "=" & TextBox1.Text & "_" & DateTime.Now.ToString("yyyy/MM/dd/HH:mm:ss"))
                                                                                                            file.Close()
                                                                                                            If iret1 = vbOK Then
                                                                                                                GoTo line1
                                                                                                            End If
                                                                                                        Else
                                                                                                            If RadioButton3.Checked = True Then
                                                                                                                st.range("D24").value = RadioButton3.Text
                                                                                                                Dim iret2 As Object = MsgBox("ACCESS GRANTED FOR" & " " & TextBox1.Text, vbInformation + vbOKOnly, "USER CREATED")
                                                                                                                file.WriteLine("USER CREATED :" & "=" & TextBox1.Text & "_" & DateTime.Now.ToString("yyyy/MM/dd/HH:mm:ss"))
                                                                                                                file.Close()
                                                                                                                If iret2 = vbOK Then
                                                                                                                    GoTo line1
                                                                                                                End If
                                                                                                            End If
                                                                                                        End If
                                                                                                    End If
                                                                                                Else
                                                                                                    If ComboBox1.Text = st.range("A25").value Then
                                                                                                        st.unprotect("Test@123")
                                                                                                        st.range("A25").value = TextBox1.Text
                                                                                                        st.range("B25").value = TextBox2.Text
                                                                                                        st.range("C25").value = TextBox3.Text
                                                                                                        If RadioButton1.Checked = True Then
                                                                                                            st.range("D25").value = RadioButton1.Text
                                                                                                            Dim iret As Object = MsgBox("ACCESS GRANTED FOR" & " " & TextBox1.Text, vbInformation + vbOKOnly, "USER CREATED")
                                                                                                            file.WriteLine("USER CREATED :" & "=" & TextBox1.Text & "_" & DateTime.Now.ToString("yyyy/MM/dd/HH:mm:ss"))
                                                                                                            file.Close()
                                                                                                            If iret = vbOK Then
                                                                                                                GoTo line1
                                                                                                            End If
                                                                                                        Else
                                                                                                            If RadioButton2.Checked = True Then
                                                                                                                st.range("D25").value = RadioButton2.Text
                                                                                                                Dim iret1 As Object = MsgBox("ACCESS GRANTED FOR" & " " & TextBox1.Text, vbInformation + vbOKOnly, "USER CREATED")
                                                                                                                file.WriteLine("USER CREATED :" & "=" & TextBox1.Text & "_" & DateTime.Now.ToString("yyyy/MM/dd/HH:mm:ss"))
                                                                                                                file.Close()
                                                                                                                If iret1 = vbOK Then
                                                                                                                    GoTo line1
                                                                                                                End If
                                                                                                            Else
                                                                                                                If RadioButton3.Checked = True Then
                                                                                                                    st.range("D25").value = RadioButton3.Text
                                                                                                                    Dim iret2 As Object = MsgBox("ACCESS GRANTED FOR" & " " & TextBox1.Text, vbInformation + vbOKOnly, "USER CREATED")
                                                                                                                    file.WriteLine("USER CREATED :" & "=" & TextBox1.Text & "_" & DateTime.Now.ToString("yyyy/MM/dd/HH:mm:ss"))
                                                                                                                    file.Close()
                                                                                                                    If iret2 = vbOK Then
                                                                                                                        GoTo line1
                                                                                                                    End If
                                                                                                                End If
                                                                                                            End If
                                                                                                        End If
                                                                                                    Else
                                                                                                        If ComboBox1.Text = st.range("A26").value Then
                                                                                                            st.unprotect("Test@123")
                                                                                                            st.range("A26").value = TextBox1.Text
                                                                                                            st.range("B26").value = TextBox2.Text
                                                                                                            st.range("C26").value = TextBox3.Text
                                                                                                            If RadioButton1.Checked = True Then
                                                                                                                st.range("D26").value = RadioButton1.Text
                                                                                                                Dim iret As Object = MsgBox("ACCESS GRANTED FOR" & " " & TextBox1.Text, vbInformation + vbOKOnly, "USER CREATED")
                                                                                                                file.WriteLine("USER CREATED :" & "=" & TextBox1.Text & "_" & DateTime.Now.ToString("yyyy/MM/dd/HH:mm:ss"))
                                                                                                                file.Close()
                                                                                                                If iret = vbOK Then
                                                                                                                    GoTo line1
                                                                                                                End If
                                                                                                            Else
                                                                                                                If RadioButton2.Checked = True Then
                                                                                                                    st.range("D26").value = RadioButton2.Text
                                                                                                                    Dim iret1 As Object = MsgBox("ACCESS GRANTED FOR" & " " & TextBox1.Text, vbInformation + vbOKOnly, "USER CREATED")
                                                                                                                    file.WriteLine("USER CREATED :" & "=" & TextBox1.Text & "_" & DateTime.Now.ToString("yyyy/MM/dd/HH:mm:ss"))
                                                                                                                    file.Close()
                                                                                                                    If iret1 = vbOK Then
                                                                                                                        GoTo line1
                                                                                                                    End If
                                                                                                                Else
                                                                                                                    If RadioButton3.Checked = True Then
                                                                                                                        st.range("D26").value = RadioButton3.Text
                                                                                                                        Dim iret2 As Object = MsgBox("ACCESS GRANTED FOR" & " " & TextBox1.Text, vbInformation + vbOKOnly, "USER CREATED")
                                                                                                                        file.WriteLine("USER CREATED :" & "=" & TextBox1.Text & "_" & DateTime.Now.ToString("yyyy/MM/dd/HH:mm:ss"))
                                                                                                                        file.Close()
                                                                                                                        If iret2 = vbOK Then
                                                                                                                            GoTo line1
                                                                                                                        End If
                                                                                                                    End If
                                                                                                                End If
                                                                                                            End If
                                                                                                        Else
                                                                                                            If ComboBox1.Text = st.range("A27").value Then
                                                                                                                st.unprotect("Test@123")
                                                                                                                st.range("A27").value = TextBox1.Text
                                                                                                                st.range("B27").value = TextBox2.Text
                                                                                                                st.range("C27").value = TextBox3.Text
                                                                                                                If RadioButton1.Checked = True Then
                                                                                                                    st.range("D27").value = RadioButton1.Text
                                                                                                                    Dim iret As Object = MsgBox("ACCESS GRANTED FOR" & " " & TextBox1.Text, vbInformation + vbOKOnly, "USER CREATED")
                                                                                                                    file.WriteLine("USER CREATED :" & "=" & TextBox1.Text & "_" & DateTime.Now.ToString("yyyy/MM/dd/HH:mm:ss"))
                                                                                                                    file.Close()
                                                                                                                    If iret = vbOK Then
                                                                                                                        GoTo line1
                                                                                                                    End If
                                                                                                                Else
                                                                                                                    If RadioButton2.Checked = True Then
                                                                                                                        st.range("D27").value = RadioButton2.Text
                                                                                                                        Dim iret1 As Object = MsgBox("ACCESS GRANTED FOR" & " " & TextBox1.Text, vbInformation + vbOKOnly, "USER CREATED")
                                                                                                                        file.WriteLine("USER CREATED :" & "=" & TextBox1.Text & "_" & DateTime.Now.ToString("yyyy/MM/dd/HH:mm:ss"))
                                                                                                                        file.Close()
                                                                                                                        If iret1 = vbOK Then
                                                                                                                            GoTo line1
                                                                                                                        End If
                                                                                                                    Else
                                                                                                                        If RadioButton3.Checked = True Then
                                                                                                                            st.range("D27").value = RadioButton3.Text
                                                                                                                            Dim iret2 As Object = MsgBox("ACCESS GRANTED FOR" & " " & TextBox1.Text, vbInformation + vbOKOnly, "USER CREATED")
                                                                                                                            file.WriteLine("USER CREATED :" & "=" & TextBox1.Text & "_" & DateTime.Now.ToString("yyyy/MM/dd/HH:mm:ss"))
                                                                                                                            file.Close()
                                                                                                                            If iret2 = vbOK Then
                                                                                                                                GoTo line1
                                                                                                                            End If
                                                                                                                        End If
                                                                                                                    End If
                                                                                                                End If
                                                                                                            Else
                                                                                                                If ComboBox1.Text = st.range("A28").value Then
                                                                                                                    st.unprotect("Test@123")
                                                                                                                    st.range("A28").value = TextBox1.Text
                                                                                                                    st.range("B28").value = TextBox2.Text
                                                                                                                    st.range("C28").value = TextBox3.Text
                                                                                                                    If RadioButton1.Checked = True Then
                                                                                                                        st.range("D28").value = RadioButton1.Text
                                                                                                                        Dim iret As Object = MsgBox("ACCESS GRANTED FOR" & " " & TextBox1.Text, vbInformation + vbOKOnly, "USER CREATED")
                                                                                                                        file.WriteLine("USER CREATED :" & "=" & TextBox1.Text & "_" & DateTime.Now.ToString("yyyy/MM/dd/HH:mm:ss"))
                                                                                                                        file.Close()
                                                                                                                        If iret = vbOK Then
                                                                                                                            GoTo line1
                                                                                                                        End If
                                                                                                                    Else
                                                                                                                        If RadioButton2.Checked = True Then
                                                                                                                            st.range("D28").value = RadioButton2.Text
                                                                                                                            Dim iret1 As Object = MsgBox("ACCESS GRANTED FOR" & " " & TextBox1.Text, vbInformation + vbOKOnly, "USER CREATED")
                                                                                                                            file.WriteLine("USER CREATED :" & "=" & TextBox1.Text & "_" & DateTime.Now.ToString("yyyy/MM/dd/HH:mm:ss"))
                                                                                                                            file.Close()
                                                                                                                            If iret1 = vbOK Then
                                                                                                                                GoTo line1
                                                                                                                            End If
                                                                                                                        Else
                                                                                                                            If RadioButton3.Checked = True Then
                                                                                                                                st.range("D28").value = RadioButton3.Text
                                                                                                                                Dim iret2 As Object = MsgBox("ACCESS GRANTED FOR" & " " & TextBox1.Text, vbInformation + vbOKOnly, "USER CREATED")
                                                                                                                                file.WriteLine("USER CREATED :" & "=" & TextBox1.Text & "_" & DateTime.Now.ToString("yyyy/MM/dd/HH:mm:ss"))
                                                                                                                                file.Close()
                                                                                                                                If iret2 = vbOK Then
                                                                                                                                    GoTo line1
                                                                                                                                End If
                                                                                                                            End If
                                                                                                                        End If
                                                                                                                    End If
                                                                                                                Else
                                                                                                                    If ComboBox1.Text = st.range("A29").value Then
                                                                                                                        st.unprotect("Test@123")
                                                                                                                        st.range("A29").value = TextBox1.Text
                                                                                                                        st.range("B29").value = TextBox2.Text
                                                                                                                        st.range("C29").value = TextBox3.Text
                                                                                                                        If RadioButton1.Checked = True Then
                                                                                                                            st.range("D29").value = RadioButton1.Text
                                                                                                                            Dim iret As Object = MsgBox("ACCESS GRANTED FOR" & " " & TextBox1.Text, vbInformation + vbOKOnly, "USER CREATED")
                                                                                                                            file.WriteLine("USER CREATED :" & "=" & TextBox1.Text & "_" & DateTime.Now.ToString("yyyy/MM/dd/HH:mm:ss"))
                                                                                                                            file.Close()
                                                                                                                            If iret = vbOK Then
                                                                                                                                GoTo line1
                                                                                                                            End If
                                                                                                                        Else
                                                                                                                            If RadioButton2.Checked = True Then
                                                                                                                                st.range("D29").value = RadioButton2.Text
                                                                                                                                Dim iret1 As Object = MsgBox("ACCESS GRANTED FOR" & " " & TextBox1.Text, vbInformation + vbOKOnly, "USER CREATED")
                                                                                                                                file.WriteLine("USER CREATED :" & "=" & TextBox1.Text & "_" & DateTime.Now.ToString("yyyy/MM/dd/HH:mm:ss"))
                                                                                                                                file.Close()
                                                                                                                                If iret1 = vbOK Then
                                                                                                                                    GoTo line1
                                                                                                                                End If
                                                                                                                            Else
                                                                                                                                If RadioButton3.Checked = True Then
                                                                                                                                    st.range("D29").value = RadioButton3.Text
                                                                                                                                    Dim iret2 As Object = MsgBox("ACCESS GRANTED FOR" & " " & TextBox1.Text, vbInformation + vbOKOnly, "USER CREATED")
                                                                                                                                    file.WriteLine("USER CREATED :" & "=" & TextBox1.Text & "_" & DateTime.Now.ToString("yyyy/MM/dd/HH:mm:ss"))
                                                                                                                                    file.Close()
                                                                                                                                    If iret2 = vbOK Then
                                                                                                                                        GoTo line1
                                                                                                                                    End If
                                                                                                                                End If
                                                                                                                            End If
                                                                                                                        End If
                                                                                                                    Else
                                                                                                                        If ComboBox1.Text = st.range("A30").value Then
                                                                                                                            st.unprotect("Test@123")
                                                                                                                            st.range("A30").value = TextBox1.Text
                                                                                                                            st.range("B30").value = TextBox2.Text
                                                                                                                            st.range("C30").value = TextBox3.Text
                                                                                                                            If RadioButton1.Checked = True Then
                                                                                                                                st.range("D30").value = RadioButton1.Text
                                                                                                                                Dim iret As Object = MsgBox("ACCESS GRANTED FOR" & " " & TextBox1.Text, vbInformation + vbOKOnly, "USER CREATED")
                                                                                                                                file.WriteLine("USER CREATED :" & "=" & TextBox1.Text & "_" & DateTime.Now.ToString("yyyy/MM/dd/HH:mm:ss"))
                                                                                                                                file.Close()
                                                                                                                                If iret = vbOK Then
                                                                                                                                    GoTo line1
                                                                                                                                End If
                                                                                                                            Else
                                                                                                                                If RadioButton2.Checked = True Then
                                                                                                                                    st.range("D30").value = RadioButton2.Text
                                                                                                                                    Dim iret1 As Object = MsgBox("ACCESS GRANTED FOR" & " " & TextBox1.Text, vbInformation + vbOKOnly, "USER CREATED")
                                                                                                                                    file.WriteLine("USER CREATED :" & "=" & TextBox1.Text & "_" & DateTime.Now.ToString("yyyy/MM/dd/HH:mm:ss"))
                                                                                                                                    file.Close()
                                                                                                                                    If iret1 = vbOK Then
                                                                                                                                        GoTo line1
                                                                                                                                    End If
                                                                                                                                Else
                                                                                                                                    If RadioButton3.Checked = True Then
                                                                                                                                        st.range("D30").value = RadioButton3.Text
                                                                                                                                        Dim iret2 As Object = MsgBox("ACCESS GRANTED FOR" & " " & TextBox1.Text, vbInformation + vbOKOnly, "USER CREATED")
                                                                                                                                        file.WriteLine("USER CREATED :" & "=" & TextBox1.Text & "_" & DateTime.Now.ToString("yyyy/MM/dd/HH:mm:ss"))
                                                                                                                                        file.Close()
                                                                                                                                        If iret2 = vbOK Then
                                                                                                                                            GoTo line1
                                                                                                                                        End If
                                                                                                                                    End If
                                                                                                                                End If
                                                                                                                            End If
                                                                                                                        Else
                                                                                                                            If ComboBox1.Text = st.range("A31").value Then
                                                                                                                                st.unprotect("Test@123")
                                                                                                                                st.range("A31").value = TextBox1.Text
                                                                                                                                st.range("B31").value = TextBox2.Text
                                                                                                                                st.range("C31").value = TextBox3.Text
                                                                                                                                If RadioButton1.Checked = True Then
                                                                                                                                    st.range("D31").value = RadioButton1.Text
                                                                                                                                    Dim iret As Object = MsgBox("ACCESS GRANTED FOR" & " " & TextBox1.Text, vbInformation + vbOKOnly, "USER CREATED")
                                                                                                                                    file.WriteLine("USER CREATED :" & "=" & TextBox1.Text & "_" & DateTime.Now.ToString("yyyy/MM/dd/HH:mm:ss"))
                                                                                                                                    file.Close()
                                                                                                                                    If iret = vbOK Then
                                                                                                                                        GoTo line1
                                                                                                                                    End If
                                                                                                                                Else
                                                                                                                                    If RadioButton2.Checked = True Then
                                                                                                                                        st.range("D31").value = RadioButton2.Text
                                                                                                                                        Dim iret1 As Object = MsgBox("ACCESS GRANTED FOR" & " " & TextBox1.Text, vbInformation + vbOKOnly, "USER CREATED")
                                                                                                                                        file.WriteLine("USER CREATED :" & "=" & TextBox1.Text & "_" & DateTime.Now.ToString("yyyy/MM/dd/HH:mm:ss"))
                                                                                                                                        file.Close()
                                                                                                                                        If iret1 = vbOK Then
                                                                                                                                            GoTo line1
                                                                                                                                        End If
                                                                                                                                    Else
                                                                                                                                        If RadioButton3.Checked = True Then
                                                                                                                                            st.range("D31").value = RadioButton3.Text
                                                                                                                                            Dim iret2 As Object = MsgBox("ACCESS GRANTED FOR" & " " & TextBox1.Text, vbInformation + vbOKOnly, "USER CREATED")
                                                                                                                                            file.WriteLine("USER CREATED :" & "=" & TextBox1.Text & "_" & DateTime.Now.ToString("yyyy/MM/dd/HH:mm:ss"))
                                                                                                                                            file.Close()
                                                                                                                                            If iret2 = vbOK Then
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

line1:
        st.protect("Test@123")
        wb.close(True)
        System.Runtime.InteropServices.Marshal.ReleaseComObject(wb)
        wb = Nothing
        xlapp.quit()
        System.Runtime.InteropServices.Marshal.ReleaseComObject(xlapp)
        xlapp = Nothing
        Process.GetProcessesByName("EXCEL")(0).Kill()
        Me.Close()
        Call SETTING.Show()
    End Sub
End Class