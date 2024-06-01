Imports Microsoft.Office.Interop.Excel
Public Class PSCH
    Private Sub PSCH_Load(sender As Object, e As EventArgs) Handles MyBase.Load
        TextBox1.Enabled = False
        TextBox3.Enabled = False
        TextBox4.Enabled = False
        Button1.Enabled = False
        Button3.Enabled = False
        TextBox2.BackColor = Color.White
        TextBox2.Clear()
        TextBox3.Clear()
        TextBox4.Clear()
    End Sub

    Private Sub TextBox2_TextChanged(sender As Object, e As EventArgs) Handles TextBox2.TextChanged
        TextBox2.PasswordChar = "*"
        If TextBox2.Text = "" Then
            Button3.Enabled = False
        Else
            Button3.Enabled = True
        End If
    End Sub

    Private Sub TextBox3_TextChanged(sender As Object, e As EventArgs) Handles TextBox3.TextChanged
        TextBox3.PasswordChar = "*"
    End Sub

    Private Sub TextBox4_TextChanged(sender As Object, e As EventArgs) Handles TextBox4.TextChanged
        TextBox4.PasswordChar = "*"
        If TextBox3.Text = TextBox4.Text Then
            TextBox4.BackColor = Color.Green
            Button1.Enabled = True
        Else
            TextBox4.BackColor = Color.Red
            Button1.Enabled = False
        End If
    End Sub

    Private Sub Button3_Click(sender As Object, e As EventArgs) Handles Button3.Click
        Dim line1 As Label
        Dim xlapp = New Microsoft.Office.Interop.Excel.Application
        Dim workbook = xlapp.Workbooks.Open(CONFIG.TextBox1.Text.ToString)
        Dim ws As Worksheet
        ws = workbook.Sheets("USER DATA")
        If TextBox1.Text = ws.Range("B2").Value And TextBox2.Text = ws.Range("C2").Value Then
            TextBox2.BackColor = Color.Green
            TextBox2.Enabled = False
            TextBox3.Enabled = True
            TextBox4.Enabled = True
            GoTo line1
        Else
            If TextBox1.Text = ws.Range("B3").Value And TextBox2.Text = ws.Range("C3").Value Then
                TextBox2.BackColor = Color.Green
                TextBox2.Enabled = False
                TextBox3.Enabled = True
                TextBox4.Enabled = True
                GoTo line1
            Else
                If TextBox1.Text = ws.Range("B4").Value And TextBox2.Text = ws.Range("C4").Value Then
                    TextBox2.BackColor = Color.Green
                    TextBox2.Enabled = False
                    TextBox3.Enabled = True
                    TextBox4.Enabled = True
                    GoTo line1
                Else
                    If TextBox1.Text = ws.Range("B5").Value And TextBox2.Text = ws.Range("C5").Value Then
                        TextBox2.BackColor = Color.Green
                        TextBox2.Enabled = False
                        TextBox3.Enabled = True
                        TextBox4.Enabled = True
                        GoTo line1
                    Else
                        If TextBox1.Text = ws.Range("B6").Value And TextBox2.Text = ws.Range("C6").Value Then
                            TextBox2.BackColor = Color.Green
                            TextBox2.Enabled = False
                            TextBox3.Enabled = True
                            TextBox4.Enabled = True
                            GoTo line1
                        Else
                            If TextBox1.Text = ws.Range("B7").Value And TextBox2.Text = ws.Range("C7").Value Then
                                TextBox2.BackColor = Color.Green
                                TextBox2.Enabled = False
                                TextBox3.Enabled = True
                                TextBox4.Enabled = True
                                GoTo line1
                            Else
                                If TextBox1.Text = ws.Range("B8").Value And TextBox2.Text = ws.Range("C8").Value Then
                                    TextBox2.BackColor = Color.Green
                                    TextBox2.Enabled = False
                                    TextBox3.Enabled = True
                                    TextBox4.Enabled = True
                                    GoTo line1
                                Else
                                    If TextBox1.Text = ws.Range("B9").Value And TextBox2.Text = ws.Range("C9").Value Then
                                        TextBox2.BackColor = Color.Green
                                        TextBox2.Enabled = False
                                        TextBox3.Enabled = True
                                        TextBox4.Enabled = True
                                        GoTo line1
                                    Else
                                        If TextBox1.Text = ws.Range("B10").Value And TextBox2.Text = ws.Range("C10").Value Then
                                            TextBox2.BackColor = Color.Green
                                            TextBox2.Enabled = False
                                            TextBox3.Enabled = True
                                            TextBox4.Enabled = True
                                            GoTo line1
                                        Else
                                            If TextBox1.Text = ws.Range("B11").Value And TextBox2.Text = ws.Range("C11").Value Then
                                                TextBox2.BackColor = Color.Green
                                                TextBox2.Enabled = False
                                                TextBox3.Enabled = True
                                                TextBox4.Enabled = True
                                                GoTo line1
                                            Else
                                                TextBox2.BackColor = Color.Red
                                                TextBox2.Enabled = True
                                                TextBox3.Enabled = False
                                                TextBox4.Enabled = False
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
line1:
        workbook.Application.DisplayAlerts = False
        workbook.Close(True)
        System.Runtime.InteropServices.Marshal.ReleaseComObject(workbook)
        workbook = Nothing
        xlapp.Quit()
        System.Runtime.InteropServices.Marshal.ReleaseComObject(xlapp)
        xlapp = Nothing
        Process.GetProcessesByName("EXCEL")(0).Kill()
    End Sub

    Private Sub Button1_Click(sender As Object, e As EventArgs) Handles Button1.Click
        Dim line1 As Label
        Dim xlapp = New Microsoft.Office.Interop.Excel.Application
        Dim workbook = xlapp.Workbooks.Open(CONFIG.TextBox1.Text.ToString)
        Dim ws As Worksheet
        ws = workbook.Sheets("USER DATA")
        If TextBox1.Text = ws.Range("B2").Value Then
            ws.Range("C2").Value = TextBox4.Text
            Dim iret As Object = MsgBox("PASSWORD UPDAED SUCCESSFULLY. LOGGING OUT NOW TO UPDATE THE PASSWORD.", vbInformation + vbOKOnly, "PASSWORD UPDATE")
            If iret = vbOK Then
                GoTo line1
            End If
        Else
            If TextBox1.Text = ws.Range("B3").Value Then
                ws.Range("C3").Value = TextBox4.Text
                Dim iret1 As Object = MsgBox("PASSWORD UPDAED SUCCESSFULLY. LOGGING OUT NOW TO UPDATE THE PASSWORD.", vbInformation + vbOKOnly, "PASSWORD UPDATE")
                If iret1 = vbOK Then
                    GoTo line1
                End If
            Else
                If TextBox1.Text = ws.Range("B4").Value Then
                    ws.Range("C4").Value = TextBox4.Text
                    Dim iret2 As Object = MsgBox("PASSWORD UPDAED SUCCESSFULLY. LOGGING OUT NOW TO UPDATE THE PASSWORD.", vbInformation + vbOKOnly, "PASSWORD UPDATE")
                    If iret2 = vbOK Then
                        GoTo line1
                    End If
                Else
                    If TextBox1.Text = ws.Range("B5").Value Then
                        ws.Range("C5").Value = TextBox4.Text
                        Dim iret3 As Object = MsgBox("PASSWORD UPDAED SUCCESSFULLY. LOGGING OUT NOW TO UPDATE THE PASSWORD.", vbInformation + vbOKOnly, "PASSWORD UPDATE")
                        If iret3 = vbOK Then
                            GoTo line1
                        End If
                    Else
                        If TextBox1.Text = ws.Range("B6").Value Then
                            ws.Range("C6").Value = TextBox4.Text
                            Dim iret4 As Object = MsgBox("PASSWORD UPDAED SUCCESSFULLY. LOGGING OUT NOW TO UPDATE THE PASSWORD.", vbInformation + vbOKOnly, "PASSWORD UPDATE")
                            If iret4 = vbOK Then
                                GoTo line1
                            End If
                        Else
                            If TextBox1.Text = ws.Range("B7").Value Then
                                ws.Range("C7").Value = TextBox4.Text
                                Dim iret5 As Object = MsgBox("PASSWORD UPDAED SUCCESSFULLY. LOGGING OUT NOW TO UPDATE THE PASSWORD.", vbInformation + vbOKOnly, "PASSWORD UPDATE")
                                If iret5 = vbOK Then
                                    GoTo line1
                                End If
                            Else
                                If TextBox1.Text = ws.Range("B8").Value Then
                                    ws.Range("C8").Value = TextBox4.Text
                                    Dim iret6 As Object = MsgBox("PASSWORD UPDAED SUCCESSFULLY. LOGGING OUT NOW TO UPDATE THE PASSWORD.", vbInformation + vbOKOnly, "PASSWORD UPDATE")
                                    If iret6 = vbOK Then
                                        GoTo line1
                                    End If
                                Else
                                    If TextBox1.Text = ws.Range("B9").Value Then
                                        ws.Range("C9").Value = TextBox4.Text
                                        Dim iret7 As Object = MsgBox("PASSWORD UPDAED SUCCESSFULLY. LOGGING OUT NOW TO UPDATE THE PASSWORD.", vbInformation + vbOKOnly, "PASSWORD UPDATE")
                                        If iret7 = vbOK Then
                                            GoTo line1
                                        End If
                                    Else
                                        If TextBox1.Text = ws.Range("B10").Value Then
                                            ws.Range("C10").Value = TextBox4.Text
                                            Dim iret8 As Object = MsgBox("PASSWORD UPDAED SUCCESSFULLY. LOGGING OUT NOW TO UPDATE THE PASSWORD.", vbInformation + vbOKOnly, "PASSWORD UPDATE")
                                            If iret8 = vbOK Then
                                                GoTo line1
                                            End If
                                        Else
                                            If TextBox1.Text = ws.Range("B11").Value Then
                                                ws.Range("C11").Value = TextBox4.Text
                                                Dim iret9 As Object = MsgBox("PASSWORD UPDAED SUCCESSFULLY. LOGGING OUT NOW TO UPDATE THE PASSWORD.", vbInformation + vbOKOnly, "PASSWORD UPDATE")
                                                If iret9 = vbOK Then
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
line1:
        workbook.Application.DisplayAlerts = False
        workbook.Close(True)
        System.Runtime.InteropServices.Marshal.ReleaseComObject(workbook)
        workbook = Nothing
        xlapp.Quit()
        System.Runtime.InteropServices.Marshal.ReleaseComObject(xlapp)
        xlapp = Nothing
        Process.GetProcessesByName("EXCEL")(0).Kill()
        Me.Close()
        Call SETTING.Show()
    End Sub

    Private Sub Button2_Click(sender As Object, e As EventArgs) Handles Button2.Click
        Me.Close()
        Call SETTING.Show()
    End Sub
End Class