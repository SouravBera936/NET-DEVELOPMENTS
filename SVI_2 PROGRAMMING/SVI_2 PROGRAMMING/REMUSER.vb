Public Class REMUSER
    Dim xlapp As Object
    Dim wb As Object
    Dim st As Object
    Private Sub REMUSER_Load(sender As Object, e As EventArgs) Handles MyBase.Load
        Dim line1 As Label
        ListBox1.Items.Clear()
        TextBox1.Text = ""
        xlapp = CreateObject("excel.application")
        wb = xlapp.workbooks.open(CONFIG.TextBox1.Text.ToString())
        st = wb.worksheets("Sheet1")
        If st.range("A2").value = "USER 1" Then
            'donothing
        Else
            ListBox1.Items.Add(st.range("A2").value)
        End If
        If st.range("A3").value = "USER 2" Then
            'donothing
        Else
            ListBox1.Items.Add(st.range("A3").value)
        End If
        If st.range("A4").value = "USER 3" Then
            'donothing
        Else
            ListBox1.Items.Add(st.range("A4").value)
        End If
        If st.range("A5").value = "USER 4" Then
            'donothing
        Else
            ListBox1.Items.Add(st.range("A5").value)
        End If
        If st.range("A6").value = "USER 5" Then
            'donothing
        Else
            ListBox1.Items.Add(st.range("A6").value)
        End If
        If st.range("A7").value = "USER 6" Then
            'donothing
        Else
            ListBox1.Items.Add(st.range("A7").value)
        End If
        If st.range("A8").value = "USER 7" Then
            'donothing
        Else
            ListBox1.Items.Add(st.range("A8").value)
        End If
        If st.range("A9").value = "USER 8" Then
            'donothing
        Else
            ListBox1.Items.Add(st.range("A9").value)
        End If
        If st.range("A10").value = "USER 9" Then
            'donothing
        Else
            ListBox1.Items.Add(st.range("A10").value)
        End If
        If st.range("A11").value = "USER 10" Then
            'donothing
        Else
            ListBox1.Items.Add(st.range("A11").value)
        End If
        If st.range("A12").value = "USER 11" Then
            'donothing
        Else
            ListBox1.Items.Add(st.range("A12").value)
        End If
        If st.range("A13").value = "USER 12" Then
            'donothing
        Else
            ListBox1.Items.Add(st.range("A13").value)
        End If
        If st.range("A14").value = "USER 13" Then
            'donothing
        Else
            ListBox1.Items.Add(st.range("A14").value)
        End If
        If st.range("A15").value = "USER 14" Then
            'donothing
        Else
            ListBox1.Items.Add(st.range("A15").value)
        End If
        If st.range("A16").value = "USER 15" Then
            'donothing
        Else
            ListBox1.Items.Add(st.range("A16").value)
        End If
        If st.range("A17").value = "USER 16" Then
            'donothing
        Else
            ListBox1.Items.Add(st.range("A17").value)
        End If
        If st.range("A18").value = "USER 17" Then
            'donothing
        Else
            ListBox1.Items.Add(st.range("A18").value)
        End If
        If st.range("A19").value = "USER 18" Then
            'donothing
        Else
            ListBox1.Items.Add(st.range("A19").value)
        End If
        If st.range("A20").value = "USER 19" Then
            'donothing
        Else
            ListBox1.Items.Add(st.range("A20").value)
        End If
        If st.range("A21").value = "USER 20" Then
            'donothing
        Else
            ListBox1.Items.Add(st.range("A21").value)
        End If
        If st.range("A22").value = "USER 21" Then
            'donothing
        Else
            ListBox1.Items.Add(st.range("A22").value)
        End If
        If st.range("A23").value = "USER 22" Then
            'donothing
        Else
            ListBox1.Items.Add(st.range("A23").value)
        End If
        If st.range("A24").value = "USER 23" Then
            'donothing
        Else
            ListBox1.Items.Add(st.range("A24").value)
        End If
        If st.range("A25").value = "USER 24" Then
            'donothing
        Else
            ListBox1.Items.Add(st.range("A25").value)
        End If
        If st.range("A26").value = "USER 25" Then
            'donothing
        Else
            ListBox1.Items.Add(st.range("A26").value)
        End If
        If st.range("A27").value = "USER 26" Then
            'donothing
        Else
            ListBox1.Items.Add(st.range("A27").value)
        End If
        If st.range("A28").value = "USER 27" Then
            'donothing
        Else
            ListBox1.Items.Add(st.range("A28").value)
        End If
        If st.range("A29").value = "USER 28" Then
            'donothing
        Else
            ListBox1.Items.Add(st.range("A29").value)
        End If
        If st.range("A30").value = "USER 29" Then
            'donothing
        Else
            ListBox1.Items.Add(st.range("A30").value)
        End If
        If st.range("A31").value = "USER 30" Then
            'donothing
        Else
            ListBox1.Items.Add(st.range("A31").value)
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
        xlapp = CreateObject("excel.application")
        wb = xlapp.workbooks.open(CONFIG.TextBox1.Text.ToString())
        st = wb.worksheets("Sheet1")
        st.unprotect("Test@123")
        Dim file As System.IO.StreamWriter
        Dim path As String = CONFIG.TextBox2.Text.ToString()
        file = My.Computer.FileSystem.OpenTextFileWriter(path, True)
        If ListBox1.SelectedItem = Nothing Then
            Dim iret As Object = MsgBox("INVALID USER SELECTION", vbCritical + vbOKOnly, "USER REMOVAL ERROR")
            file.WriteLine("USER REMOVAL FAILED :" & DateTime.Now.ToString("yyyy/MM/dd/HH:mm:ss"))
            file.Close()
            GoTo line1
        Else
            If TextBox1.Text = "" Then
                Dim iret1 As Object = MsgBox("INVALID REASON", vbCritical + vbOKOnly, "USER REMOVAL ERROR")
                file.WriteLine("USER REMOVAL FAILED :" & DateTime.Now.ToString("yyyy/MM/dd/HH:mm:ss"))
                file.Close()
                GoTo line1
            Else
                If ListBox1.SelectedItem = st.range("A2").value Then
                    Dim iret2 As Object = MsgBox("DO YOU WANT TO REMOVE USER ACCESS FROM :" & ListBox1.SelectedItem, vbQuestion + vbYesNo, "USER REMOVAL")
                    If iret2 = vbYes Then
                        st.range("B2").value = ""
                        st.range("C2").value = ""
                        st.range("D2").value = ""
                        st.range("A2").value = "USER 1"
                        file.WriteLine("USER REMOVED :" & " " & ListBox1.SelectedItem.ToString() & " " & "FOR" & " " & TextBox1.Text & "AT" & " " & DateTime.Now.ToString("yyyy/MM/dd/HH:mm:ss"))
                        file.Close()
                        MsgBox("USER REMOVED SUCCESSFULLY")
                        GoTo line1
                        Me.Close()
                        Call SETTING.Show()
                    ElseIf iret2 = vbNo Then
                        'donothing
                    End If
                Else
                    If ListBox1.SelectedItem = st.range("A3").value Then
                        Dim iret3 As Object = MsgBox("DO YOU WANT TO REMOVE USER ACCESS FROM :" & ListBox1.SelectedItem, vbQuestion + vbYesNo, "USER REMOVAL")
                        If iret3 = vbYes Then
                            st.range("B3").value = ""
                            st.range("C3").value = ""
                            st.range("D3").value = ""
                            st.range("A3").value = "USER 2"
                            file.WriteLine("USER REMOVED :" & " " & ListBox1.SelectedItem.ToString() & " " & "FOR" & " " & TextBox1.Text & "AT" & " " & DateTime.Now.ToString("yyyy/MM/dd/HH:mm:ss"))
                            file.Close()
                            MsgBox("USER REMOVED SUCCESSFULLY")
                            GoTo line1
                            Me.Close()
                            Call SETTING.Show()
                        ElseIf iret3 = vbNo Then
                            'donothing
                        End If
                    Else
                        If ListBox1.SelectedItem = st.range("A4").value Then
                            Dim iret4 As Object = MsgBox("DO YOU WANT TO REMOVE USER ACCESS FROM :" & ListBox1.SelectedItem, vbQuestion + vbYesNo, "USER REMOVAL")
                            If iret4 = vbYes Then
                                st.range("B4").value = ""
                                st.range("C4").value = ""
                                st.range("D4").value = ""
                                st.range("A4").value = "USER 3"
                                file.WriteLine("USER REMOVED :" & " " & ListBox1.SelectedItem.ToString() & " " & "FOR" & " " & TextBox1.Text & "AT" & " " & DateTime.Now.ToString("yyyy/MM/dd/HH:mm:ss"))
                                file.Close()
                                MsgBox("USER REMOVED SUCCESSFULLY")
                                GoTo line1
                                Me.Close()
                                Call SETTING.Show()
                            ElseIf iret4 = vbNo Then
                                'donothing
                            End If
                        Else
                            If ListBox1.SelectedItem = st.range("A5").value Then
                                Dim iret5 As Object = MsgBox("DO YOU WANT TO REMOVE USER ACCESS FROM :" & ListBox1.SelectedItem, vbQuestion + vbYesNo, "USER REMOVAL")
                                If iret5 = vbYes Then
                                    st.range("B5").value = ""
                                    st.range("C5").value = ""
                                    st.range("D5").value = ""
                                    st.range("A5").value = "USER 4"
                                    file.WriteLine("USER REMOVED :" & " " & ListBox1.SelectedItem.ToString() & " " & "FOR" & " " & TextBox1.Text & "AT" & " " & DateTime.Now.ToString("yyyy/MM/dd/HH:mm:ss"))
                                    file.Close()
                                    MsgBox("USER REMOVED SUCCESSFULLY")
                                    GoTo line1
                                    Me.Close()
                                    Call SETTING.Show()
                                ElseIf iret5 = vbNo Then
                                    'donothing
                                End If
                            Else
                                If ListBox1.SelectedItem = st.range("A6").value Then
                                    Dim iret6 As Object = MsgBox("DO YOU WANT TO REMOVE USER ACCESS FROM :" & ListBox1.SelectedItem, vbQuestion + vbYesNo, "USER REMOVAL")
                                    If iret6 = vbYes Then
                                        st.range("B6").value = ""
                                        st.range("C6").value = ""
                                        st.range("D6").value = ""
                                        st.range("A6").value = "USER 5"
                                        file.WriteLine("USER REMOVED :" & " " & ListBox1.SelectedItem.ToString() & " " & "FOR" & " " & TextBox1.Text & "AT" & " " & DateTime.Now.ToString("yyyy/MM/dd/HH:mm:ss"))
                                        file.Close()
                                        MsgBox("USER REMOVED SUCCESSFULLY")
                                        GoTo line1
                                        Me.Close()
                                        Call SETTING.Show()
                                    ElseIf iret6 = vbNo Then
                                        'donothing
                                    End If
                                Else
                                    If ListBox1.SelectedItem = st.range("A7").value Then
                                        Dim iret7 As Object = MsgBox("DO YOU WANT TO REMOVE USER ACCESS FROM :" & ListBox1.SelectedItem, vbQuestion + vbYesNo, "USER REMOVAL")
                                        If iret7 = vbYes Then
                                            st.range("B7").value = ""
                                            st.range("C7").value = ""
                                            st.range("D7").value = ""
                                            st.range("A7").value = "USER 6"
                                            file.WriteLine("USER REMOVED :" & " " & ListBox1.SelectedItem.ToString() & " " & "FOR" & " " & TextBox1.Text & "AT" & " " & DateTime.Now.ToString("yyyy/MM/dd/HH:mm:ss"))
                                            file.Close()
                                            MsgBox("USER REMOVED SUCCESSFULLY")
                                            GoTo line1
                                            Me.Close()
                                            Call SETTING.Show()
                                        ElseIf iret7 = vbNo Then
                                            'donothing
                                        End If
                                    Else
                                        If ListBox1.SelectedItem = st.range("A8").value Then
                                            Dim iret8 As Object = MsgBox("DO YOU WANT TO REMOVE USER ACCESS FROM :" & ListBox1.SelectedItem, vbQuestion + vbYesNo, "USER REMOVAL")
                                            If iret8 = vbYes Then
                                                st.range("B8").value = ""
                                                st.range("C8").value = ""
                                                st.range("D8").value = ""
                                                st.range("A8").value = "USER 7"
                                                file.WriteLine("USER REMOVED :" & " " & ListBox1.SelectedItem.ToString() & " " & "FOR" & " " & TextBox1.Text & "AT" & " " & DateTime.Now.ToString("yyyy/MM/dd/HH:mm:ss"))
                                                file.Close()
                                                MsgBox("USER REMOVED SUCCESSFULLY")
                                                GoTo line1
                                                Me.Close()
                                                Call SETTING.Show()
                                            ElseIf iret8 = vbNo Then
                                                'donothing
                                            End If
                                        Else
                                            If ListBox1.SelectedItem = st.range("A9").value Then
                                                Dim iret9 As Object = MsgBox("DO YOU WANT TO REMOVE USER ACCESS FROM :" & ListBox1.SelectedItem, vbQuestion + vbYesNo, "USER REMOVAL")
                                                If iret9 = vbYes Then
                                                    st.range("B9").value = ""
                                                    st.range("C9").value = ""
                                                    st.range("D9").value = ""
                                                    st.range("A9").value = "USER 8"
                                                    file.WriteLine("USER REMOVED :" & " " & ListBox1.SelectedItem.ToString() & " " & "FOR" & " " & TextBox1.Text & "AT" & " " & DateTime.Now.ToString("yyyy/MM/dd/HH:mm:ss"))
                                                    file.Close()
                                                    MsgBox("USER REMOVED SUCCESSFULLY")
                                                    GoTo line1
                                                    Me.Close()
                                                    Call SETTING.Show()
                                                ElseIf iret9 = vbNo Then
                                                    'donothing
                                                End If
                                            Else
                                                If ListBox1.SelectedItem = st.range("A10").value Then
                                                    Dim iret10 As Object = MsgBox("DO YOU WANT TO REMOVE USER ACCESS FROM :" & ListBox1.SelectedItem, vbQuestion + vbYesNo, "USER REMOVAL")
                                                    If iret10 = vbYes Then
                                                        st.range("B10").value = ""
                                                        st.range("C10").value = ""
                                                        st.range("D10").value = ""
                                                        st.range("A10").value = "USER 9"
                                                        file.WriteLine("USER REMOVED :" & " " & ListBox1.SelectedItem.ToString() & " " & "FOR" & " " & TextBox1.Text & "AT" & " " & DateTime.Now.ToString("yyyy/MM/dd/HH:mm:ss"))
                                                        file.Close()
                                                        MsgBox("USER REMOVED SUCCESSFULLY")
                                                        GoTo line1
                                                        Me.Close()
                                                        Call SETTING.Show()
                                                    ElseIf iret10 = vbNo Then
                                                        'donothing
                                                    End If
                                                Else
                                                    If ListBox1.SelectedItem = st.range("A11").value Then
                                                        Dim iret11 As Object = MsgBox("DO YOU WANT TO REMOVE USER ACCESS FROM :" & ListBox1.SelectedItem, vbQuestion + vbYesNo, "USER REMOVAL")
                                                        If iret11 = vbYes Then
                                                            st.range("B11").value = ""
                                                            st.range("C11").value = ""
                                                            st.range("D11").value = ""
                                                            st.range("A11").value = "USER 10"
                                                            file.WriteLine("USER REMOVED :" & " " & ListBox1.SelectedItem.ToString() & " " & "FOR" & " " & TextBox1.Text & "AT" & " " & DateTime.Now.ToString("yyyy/MM/dd/HH:mm:ss"))
                                                            file.Close()
                                                            MsgBox("USER REMOVED SUCCESSFULLY")
                                                            GoTo line1
                                                            Me.Close()
                                                            Call SETTING.Show()
                                                        ElseIf iret11 = vbNo Then
                                                            'donothing
                                                        End If
                                                    Else
                                                        If ListBox1.SelectedItem = st.range("A12").value Then
                                                            Dim iret12 As Object = MsgBox("DO YOU WANT TO REMOVE USER ACCESS FROM :" & ListBox1.SelectedItem, vbQuestion + vbYesNo, "USER REMOVAL")
                                                            If iret12 = vbYes Then
                                                                st.range("B12").value = ""
                                                                st.range("C12").value = ""
                                                                st.range("D12").value = ""
                                                                st.range("A12").value = "USER 11"
                                                                file.WriteLine("USER REMOVED :" & " " & ListBox1.SelectedItem.ToString() & " " & "FOR" & " " & TextBox1.Text & "AT" & " " & DateTime.Now.ToString("yyyy/MM/dd/HH:mm:ss"))
                                                                file.Close()
                                                                MsgBox("USER REMOVED SUCCESSFULLY")
                                                                GoTo line1
                                                                Me.Close()
                                                                Call SETTING.Show()
                                                            ElseIf iret12 = vbNo Then
                                                                'donothing
                                                            End If
                                                        Else
                                                            If ListBox1.SelectedItem = st.range("A13").value Then
                                                                Dim iret13 As Object = MsgBox("DO YOU WANT TO REMOVE USER ACCESS FROM :" & ListBox1.SelectedItem, vbQuestion + vbYesNo, "USER REMOVAL")
                                                                If iret13 = vbYes Then
                                                                    st.range("B13").value = ""
                                                                    st.range("C13").value = ""
                                                                    st.range("D13").value = ""
                                                                    st.range("A13").value = "USER 12"
                                                                    file.WriteLine("USER REMOVED :" & " " & ListBox1.SelectedItem.ToString() & " " & "FOR" & " " & TextBox1.Text & "AT" & " " & DateTime.Now.ToString("yyyy/MM/dd/HH:mm:ss"))
                                                                    file.Close()
                                                                    MsgBox("USER REMOVED SUCCESSFULLY")
                                                                    GoTo line1
                                                                    Me.Close()
                                                                    Call SETTING.Show()
                                                                ElseIf iret13 = vbNo Then
                                                                    'donothing
                                                                End If
                                                            Else
                                                                If ListBox1.SelectedItem = st.range("A14").value Then
                                                                    Dim iret14 As Object = MsgBox("DO YOU WANT TO REMOVE USER ACCESS FROM :" & ListBox1.SelectedItem, vbQuestion + vbYesNo, "USER REMOVAL")
                                                                    If iret14 = vbYes Then
                                                                        st.range("B14").value = ""
                                                                        st.range("C14").value = ""
                                                                        st.range("D14").value = ""
                                                                        st.range("A14").value = "USER 13"
                                                                        file.WriteLine("USER REMOVED :" & " " & ListBox1.SelectedItem.ToString() & " " & "FOR" & " " & TextBox1.Text & "AT" & " " & DateTime.Now.ToString("yyyy/MM/dd/HH:mm:ss"))
                                                                        file.Close()
                                                                        MsgBox("USER REMOVED SUCCESSFULLY")
                                                                        GoTo line1
                                                                        Me.Close()
                                                                        Call SETTING.Show()
                                                                    ElseIf iret14 = vbNo Then
                                                                        'donothing
                                                                    End If
                                                                Else
                                                                    If ListBox1.SelectedItem = st.range("A15").value Then
                                                                        Dim iret15 As Object = MsgBox("DO YOU WANT TO REMOVE USER ACCESS FROM :" & ListBox1.SelectedItem, vbQuestion + vbYesNo, "USER REMOVAL")
                                                                        If iret15 = vbYes Then
                                                                            st.range("B15").value = ""
                                                                            st.range("C15").value = ""
                                                                            st.range("D15").value = ""
                                                                            st.range("A15").value = "USER 14"
                                                                            file.WriteLine("USER REMOVED :" & " " & ListBox1.SelectedItem.ToString() & " " & "FOR" & " " & TextBox1.Text & "AT" & " " & DateTime.Now.ToString("yyyy/MM/dd/HH:mm:ss"))
                                                                            file.Close()
                                                                            MsgBox("USER REMOVED SUCCESSFULLY")
                                                                            GoTo line1
                                                                            Me.Close()
                                                                            Call SETTING.Show()
                                                                        ElseIf iret15 = vbNo Then
                                                                            'donothing
                                                                        End If
                                                                    Else
                                                                        If ListBox1.SelectedItem = st.range("A16").value Then
                                                                            Dim iret16 As Object = MsgBox("DO YOU WANT TO REMOVE USER ACCESS FROM :" & ListBox1.SelectedItem, vbQuestion + vbYesNo, "USER REMOVAL")
                                                                            If iret16 = vbYes Then
                                                                                st.range("B16").value = ""
                                                                                st.range("C16").value = ""
                                                                                st.range("D16").value = ""
                                                                                st.range("A16").value = "USER 15"
                                                                                file.WriteLine("USER REMOVED :" & " " & ListBox1.SelectedItem.ToString() & " " & "FOR" & " " & TextBox1.Text & "AT" & " " & DateTime.Now.ToString("yyyy/MM/dd/HH:mm:ss"))
                                                                                file.Close()
                                                                                MsgBox("USER REMOVED SUCCESSFULLY")
                                                                                GoTo line1
                                                                                Me.Close()
                                                                                Call SETTING.Show()
                                                                            ElseIf iret16 = vbNo Then
                                                                                'donothing
                                                                            End If
                                                                        Else
                                                                            If ListBox1.SelectedItem = st.range("A17").value Then
                                                                                Dim iret17 As Object = MsgBox("DO YOU WANT TO REMOVE USER ACCESS FROM :" & ListBox1.SelectedItem, vbQuestion + vbYesNo, "USER REMOVAL")
                                                                                If iret17 = vbYes Then
                                                                                    st.range("B17").value = ""
                                                                                    st.range("C17").value = ""
                                                                                    st.range("D17").value = ""
                                                                                    st.range("A17").value = "USER 16"
                                                                                    file.WriteLine("USER REMOVED :" & " " & ListBox1.SelectedItem.ToString() & " " & "FOR" & " " & TextBox1.Text & "AT" & " " & DateTime.Now.ToString("yyyy/MM/dd/HH:mm:ss"))
                                                                                    file.Close()
                                                                                    MsgBox("USER REMOVED SUCCESSFULLY")
                                                                                    GoTo line1
                                                                                    Me.Close()
                                                                                    Call SETTING.Show()
                                                                                ElseIf iret17 = vbNo Then
                                                                                    'donothing
                                                                                End If
                                                                            Else
                                                                                If ListBox1.SelectedItem = st.range("A18").value Then
                                                                                    Dim iret18 As Object = MsgBox("DO YOU WANT TO REMOVE USER ACCESS FROM :" & ListBox1.SelectedItem, vbQuestion + vbYesNo, "USER REMOVAL")
                                                                                    If iret18 = vbYes Then
                                                                                        st.range("B18").value = ""
                                                                                        st.range("C18").value = ""
                                                                                        st.range("D18").value = ""
                                                                                        st.range("A18").value = "USER 17"
                                                                                        file.WriteLine("USER REMOVED :" & " " & ListBox1.SelectedItem.ToString() & " " & "FOR" & " " & TextBox1.Text & "AT" & " " & DateTime.Now.ToString("yyyy/MM/dd/HH:mm:ss"))
                                                                                        file.Close()
                                                                                        MsgBox("USER REMOVED SUCCESSFULLY")
                                                                                        GoTo line1
                                                                                        Me.Close()
                                                                                        Call SETTING.Show()
                                                                                    ElseIf iret18 = vbNo Then
                                                                                        'donothing
                                                                                    End If
                                                                                Else
                                                                                    If ListBox1.SelectedItem = st.range("A19").value Then
                                                                                        Dim iret19 As Object = MsgBox("DO YOU WANT TO REMOVE USER ACCESS FROM :" & ListBox1.SelectedItem, vbQuestion + vbYesNo, "USER REMOVAL")
                                                                                        If iret19 = vbYes Then
                                                                                            st.range("B19").value = ""
                                                                                            st.range("C19").value = ""
                                                                                            st.range("D19").value = ""
                                                                                            st.range("A19").value = "USER 18"
                                                                                            file.WriteLine("USER REMOVED :" & " " & ListBox1.SelectedItem.ToString() & " " & "FOR" & " " & TextBox1.Text & "AT" & " " & DateTime.Now.ToString("yyyy/MM/dd/HH:mm:ss"))
                                                                                            file.Close()
                                                                                            MsgBox("USER REMOVED SUCCESSFULLY")
                                                                                            GoTo line1
                                                                                            Me.Close()
                                                                                            Call SETTING.Show()
                                                                                        ElseIf iret19 = vbNo Then
                                                                                            'donothing
                                                                                        End If
                                                                                    Else
                                                                                        If ListBox1.SelectedItem = st.range("A20").value Then
                                                                                            Dim iret20 As Object = MsgBox("DO YOU WANT TO REMOVE USER ACCESS FROM :" & ListBox1.SelectedItem, vbQuestion + vbYesNo, "USER REMOVAL")
                                                                                            If iret20 = vbYes Then
                                                                                                st.range("B20").value = ""
                                                                                                st.range("C20").value = ""
                                                                                                st.range("D20").value = ""
                                                                                                st.range("A20").value = "USER 19"
                                                                                                file.WriteLine("USER REMOVED :" & " " & ListBox1.SelectedItem.ToString() & " " & "FOR" & " " & TextBox1.Text & "AT" & " " & DateTime.Now.ToString("yyyy/MM/dd/HH:mm:ss"))
                                                                                                file.Close()
                                                                                                MsgBox("USER REMOVED SUCCESSFULLY")
                                                                                                GoTo line1
                                                                                                Me.Close()
                                                                                                Call SETTING.Show()
                                                                                            ElseIf iret20 = vbNo Then
                                                                                                'donothing
                                                                                            End If
                                                                                        Else
                                                                                            If ListBox1.SelectedItem = st.range("A21").value Then
                                                                                                Dim iret21 As Object = MsgBox("DO YOU WANT TO REMOVE USER ACCESS FROM :" & ListBox1.SelectedItem, vbQuestion + vbYesNo, "USER REMOVAL")
                                                                                                If iret21 = vbYes Then
                                                                                                    st.range("B21").value = ""
                                                                                                    st.range("C21").value = ""
                                                                                                    st.range("D21").value = ""
                                                                                                    st.range("A21").value = "USER 20"
                                                                                                    file.WriteLine("USER REMOVED :" & " " & ListBox1.SelectedItem.ToString() & " " & "FOR" & " " & TextBox1.Text & "AT" & " " & DateTime.Now.ToString("yyyy/MM/dd/HH:mm:ss"))
                                                                                                    file.Close()
                                                                                                    MsgBox("USER REMOVED SUCCESSFULLY")
                                                                                                    GoTo line1
                                                                                                    Me.Close()
                                                                                                    Call SETTING.Show()
                                                                                                ElseIf iret21 = vbNo Then
                                                                                                    'donothing
                                                                                                End If
                                                                                            Else
                                                                                                If ListBox1.SelectedItem = st.range("A22").value Then
                                                                                                    Dim iret22 As Object = MsgBox("DO YOU WANT TO REMOVE USER ACCESS FROM :" & ListBox1.SelectedItem, vbQuestion + vbYesNo, "USER REMOVAL")
                                                                                                    If iret22 = vbYes Then
                                                                                                        st.range("B22").value = ""
                                                                                                        st.range("C22").value = ""
                                                                                                        st.range("D22").value = ""
                                                                                                        st.range("A22").value = "USER 21"
                                                                                                        file.WriteLine("USER REMOVED :" & " " & ListBox1.SelectedItem.ToString() & " " & "FOR" & " " & TextBox1.Text & "AT" & " " & DateTime.Now.ToString("yyyy/MM/dd/HH:mm:ss"))
                                                                                                        file.Close()
                                                                                                        MsgBox("USER REMOVED SUCCESSFULLY")
                                                                                                        GoTo line1
                                                                                                        Me.Close()
                                                                                                        Call SETTING.Show()
                                                                                                    ElseIf iret22 = vbNo Then
                                                                                                        'donothing
                                                                                                    End If
                                                                                                Else
                                                                                                    If ListBox1.SelectedItem = st.range("A23").value Then
                                                                                                        Dim iret23 As Object = MsgBox("DO YOU WANT TO REMOVE USER ACCESS FROM :" & ListBox1.SelectedItem, vbQuestion + vbYesNo, "USER REMOVAL")
                                                                                                        If iret23 = vbYes Then
                                                                                                            st.range("B23").value = ""
                                                                                                            st.range("C23").value = ""
                                                                                                            st.range("D23").value = ""
                                                                                                            st.range("A23").value = "USER 22"
                                                                                                            file.WriteLine("USER REMOVED :" & " " & ListBox1.SelectedItem.ToString() & " " & "FOR" & " " & TextBox1.Text & "AT" & " " & DateTime.Now.ToString("yyyy/MM/dd/HH:mm:ss"))
                                                                                                            file.Close()
                                                                                                            MsgBox("USER REMOVED SUCCESSFULLY")
                                                                                                            GoTo line1
                                                                                                            Me.Close()
                                                                                                            Call SETTING.Show()
                                                                                                        ElseIf iret23 = vbNo Then
                                                                                                            'donothing
                                                                                                        End If
                                                                                                    Else
                                                                                                        If ListBox1.SelectedItem = st.range("A24").value Then
                                                                                                            Dim iret24 As Object = MsgBox("DO YOU WANT TO REMOVE USER ACCESS FROM :" & ListBox1.SelectedItem, vbQuestion + vbYesNo, "USER REMOVAL")
                                                                                                            If iret24 = vbYes Then
                                                                                                                st.range("B24").value = ""
                                                                                                                st.range("C24").value = ""
                                                                                                                st.range("D24").value = ""
                                                                                                                st.range("A24").value = "USER 23"
                                                                                                                file.WriteLine("USER REMOVED :" & " " & ListBox1.SelectedItem.ToString() & " " & "FOR" & " " & TextBox1.Text & "AT" & " " & DateTime.Now.ToString("yyyy/MM/dd/HH:mm:ss"))
                                                                                                                file.Close()
                                                                                                                MsgBox("USER REMOVED SUCCESSFULLY")
                                                                                                                GoTo line1
                                                                                                                Me.Close()
                                                                                                                Call SETTING.Show()
                                                                                                            ElseIf iret24 = vbNo Then
                                                                                                                'donothing
                                                                                                            End If
                                                                                                        Else
                                                                                                            If ListBox1.SelectedItem = st.range("A25").value Then
                                                                                                                Dim iret25 As Object = MsgBox("DO YOU WANT TO REMOVE USER ACCESS FROM :" & ListBox1.SelectedItem, vbQuestion + vbYesNo, "USER REMOVAL")
                                                                                                                If iret25 = vbYes Then
                                                                                                                    st.range("B25").value = ""
                                                                                                                    st.range("C25").value = ""
                                                                                                                    st.range("D25").value = ""
                                                                                                                    st.range("A25").value = "USER 24"
                                                                                                                    file.WriteLine("USER REMOVED :" & " " & ListBox1.SelectedItem.ToString() & " " & "FOR" & " " & TextBox1.Text & "AT" & " " & DateTime.Now.ToString("yyyy/MM/dd/HH:mm:ss"))
                                                                                                                    file.Close()
                                                                                                                    MsgBox("USER REMOVED SUCCESSFULLY")
                                                                                                                    GoTo line1
                                                                                                                    Me.Close()
                                                                                                                    Call SETTING.Show()
                                                                                                                ElseIf iret25 = vbNo Then
                                                                                                                    'donothing
                                                                                                                End If
                                                                                                            Else
                                                                                                                If ListBox1.SelectedItem = st.range("A26").value Then
                                                                                                                    Dim iret26 As Object = MsgBox("DO YOU WANT TO REMOVE USER ACCESS FROM :" & ListBox1.SelectedItem, vbQuestion + vbYesNo, "USER REMOVAL")
                                                                                                                    If iret26 = vbYes Then
                                                                                                                        st.range("B26").value = ""
                                                                                                                        st.range("C26").value = ""
                                                                                                                        st.range("D26").value = ""
                                                                                                                        st.range("A26").value = "USER 25"
                                                                                                                        file.WriteLine("USER REMOVED :" & " " & ListBox1.SelectedItem.ToString() & " " & "FOR" & " " & TextBox1.Text & "AT" & " " & DateTime.Now.ToString("yyyy/MM/dd/HH:mm:ss"))
                                                                                                                        file.Close()
                                                                                                                        MsgBox("USER REMOVED SUCCESSFULLY")
                                                                                                                        GoTo line1
                                                                                                                        Me.Close()
                                                                                                                        Call SETTING.Show()
                                                                                                                    ElseIf iret26 = vbNo Then
                                                                                                                        'donothing
                                                                                                                    End If
                                                                                                                Else
                                                                                                                    If ListBox1.SelectedItem = st.range("A27").value Then
                                                                                                                        Dim iret27 As Object = MsgBox("DO YOU WANT TO REMOVE USER ACCESS FROM :" & ListBox1.SelectedItem, vbQuestion + vbYesNo, "USER REMOVAL")
                                                                                                                        If iret27 = vbYes Then
                                                                                                                            st.range("B27").value = ""
                                                                                                                            st.range("C27").value = ""
                                                                                                                            st.range("D27").value = ""
                                                                                                                            st.range("A27").value = "USER 26"
                                                                                                                            file.WriteLine("USER REMOVED :" & " " & ListBox1.SelectedItem.ToString() & " " & "FOR" & " " & TextBox1.Text & "AT" & " " & DateTime.Now.ToString("yyyy/MM/dd/HH:mm:ss"))
                                                                                                                            file.Close()
                                                                                                                            MsgBox("USER REMOVED SUCCESSFULLY")
                                                                                                                            GoTo line1
                                                                                                                            Me.Close()
                                                                                                                            Call SETTING.Show()
                                                                                                                        ElseIf iret27 = vbNo Then
                                                                                                                            'donothing
                                                                                                                        End If
                                                                                                                    Else
                                                                                                                        If ListBox1.SelectedItem = st.range("A28").value Then
                                                                                                                            Dim iret28 As Object = MsgBox("DO YOU WANT TO REMOVE USER ACCESS FROM :" & ListBox1.SelectedItem, vbQuestion + vbYesNo, "USER REMOVAL")
                                                                                                                            If iret28 = vbYes Then
                                                                                                                                st.range("B28").value = ""
                                                                                                                                st.range("C28").value = ""
                                                                                                                                st.range("D28").value = ""
                                                                                                                                st.range("A28").value = "USER 27"
                                                                                                                                file.WriteLine("USER REMOVED :" & " " & ListBox1.SelectedItem.ToString() & " " & "FOR" & " " & TextBox1.Text & "AT" & " " & DateTime.Now.ToString("yyyy/MM/dd/HH:mm:ss"))
                                                                                                                                file.Close()
                                                                                                                                MsgBox("USER REMOVED SUCCESSFULLY")
                                                                                                                                GoTo line1
                                                                                                                                Me.Close()
                                                                                                                                Call SETTING.Show()
                                                                                                                            ElseIf iret28 = vbNo Then
                                                                                                                                'donothing
                                                                                                                            End If
                                                                                                                        Else
                                                                                                                            If ListBox1.SelectedItem = st.range("A29").value Then
                                                                                                                                Dim iret29 As Object = MsgBox("DO YOU WANT TO REMOVE USER ACCESS FROM :" & ListBox1.SelectedItem, vbQuestion + vbYesNo, "USER REMOVAL")
                                                                                                                                If iret29 = vbYes Then
                                                                                                                                    st.range("B29").value = ""
                                                                                                                                    st.range("C29").value = ""
                                                                                                                                    st.range("D29").value = ""
                                                                                                                                    st.range("A29").value = "USER 28"
                                                                                                                                    file.WriteLine("USER REMOVED :" & " " & ListBox1.SelectedItem.ToString() & " " & "FOR" & " " & TextBox1.Text & "AT" & " " & DateTime.Now.ToString("yyyy/MM/dd/HH:mm:ss"))
                                                                                                                                    file.Close()
                                                                                                                                    MsgBox("USER REMOVED SUCCESSFULLY")
                                                                                                                                    GoTo line1
                                                                                                                                    Me.Close()
                                                                                                                                    Call SETTING.Show()
                                                                                                                                ElseIf iret29 = vbNo Then
                                                                                                                                    'donothing
                                                                                                                                End If
                                                                                                                            Else
                                                                                                                                If ListBox1.SelectedItem = st.range("A30").value Then
                                                                                                                                    Dim iret30 As Object = MsgBox("DO YOU WANT TO REMOVE USER ACCESS FROM :" & ListBox1.SelectedItem, vbQuestion + vbYesNo, "USER REMOVAL")
                                                                                                                                    If iret30 = vbYes Then
                                                                                                                                        st.range("B30").value = ""
                                                                                                                                        st.range("C30").value = ""
                                                                                                                                        st.range("D30").value = ""
                                                                                                                                        st.range("A30").value = "USER 29"
                                                                                                                                        file.WriteLine("USER REMOVED :" & " " & ListBox1.SelectedItem.ToString() & " " & "FOR" & " " & TextBox1.Text & "AT" & " " & DateTime.Now.ToString("yyyy/MM/dd/HH:mm:ss"))
                                                                                                                                        file.Close()
                                                                                                                                        MsgBox("USER REMOVED SUCCESSFULLY")
                                                                                                                                        GoTo line1
                                                                                                                                        Me.Close()
                                                                                                                                        Call SETTING.Show()
                                                                                                                                    ElseIf iret30 = vbNo Then
                                                                                                                                        'donothing
                                                                                                                                    End If
                                                                                                                                Else
                                                                                                                                    If ListBox1.SelectedItem = st.range("A31").value Then
                                                                                                                                        Dim iret31 As Object = MsgBox("DO YOU WANT TO REMOVE USER ACCESS FROM :" & ListBox1.SelectedItem, vbQuestion + vbYesNo, "USER REMOVAL")
                                                                                                                                        If iret31 = vbYes Then
                                                                                                                                            st.range("B31").value = ""
                                                                                                                                            st.range("C31").value = ""
                                                                                                                                            st.range("D31").value = ""
                                                                                                                                            st.range("A31").value = "USER 30"
                                                                                                                                            file.WriteLine("USER REMOVED :" & " " & ListBox1.SelectedItem.ToString() & " " & "FOR" & " " & TextBox1.Text & "AT" & " " & DateTime.Now.ToString("yyyy/MM/dd/HH:mm:ss"))
                                                                                                                                            file.Close()
                                                                                                                                            MsgBox("USER REMOVED SUCCESSFULLY")
                                                                                                                                            GoTo line1
                                                                                                                                            Me.Close()
                                                                                                                                            Call SETTING.Show()
                                                                                                                                        ElseIf iret31 = vbNo Then
                                                                                                                                            'donothing
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
    End Sub

    Private Sub Button2_Click(sender As Object, e As EventArgs) Handles Button2.Click
        Me.Close()
        Call SETTING.Show()
    End Sub
End Class