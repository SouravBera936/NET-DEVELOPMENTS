Imports Microsoft.Office.Interop.Excel
Public Class CUSER
    Private Sub CUSER_Load(sender As Object, e As EventArgs) Handles MyBase.Load
        TextBox1.Enabled = False
        TextBox2.Enabled = False
        TextBox3.Enabled = False
        TextBox4.Enabled = False
        Button1.Enabled = False
        Dim xlapp = New Microsoft.Office.Interop.Excel.Application
        Dim workbook = xlapp.Workbooks.Open(CONFIG.TextBox1.Text.ToString)
        Dim ws As Worksheet
        ws = workbook.Sheets("USER DATA")
        If ws.Range("C2").Value = "" Then
            ComboBox1.Items.Add(ws.Range("A2").Value)
        End If
        If ws.Range("C3").Value = "" Then
            ComboBox1.Items.Add(ws.Range("A3").Value)
        End If
        If ws.Range("C4").Value = "" Then
            ComboBox1.Items.Add(ws.Range("A4").Value)
        End If
        If ws.Range("C5").Value = "" Then
            ComboBox1.Items.Add(ws.Range("A5").Value)
        End If
        If ws.Range("C6").Value = "" Then
            ComboBox1.Items.Add(ws.Range("A6").Value)
        End If
        If ws.Range("C7").Value = "" Then
            ComboBox1.Items.Add(ws.Range("A7").Value)
        End If
        If ws.Range("C8").Value = "" Then
            ComboBox1.Items.Add(ws.Range("A8").Value)
        End If
        If ws.Range("C9").Value = "" Then
            ComboBox1.Items.Add(ws.Range("A9").Value)
        End If
        If ws.Range("C10").Value = "" Then
            ComboBox1.Items.Add(ws.Range("A10").Value)
        End If
        If ws.Range("C11").Value = "" Then
            ComboBox1.Items.Add(ws.Range("A11").Value)
        End If
        workbook.Close(True)
        System.Runtime.InteropServices.Marshal.ReleaseComObject(workbook)
        workbook = Nothing
        xlapp.Quit()
        System.Runtime.InteropServices.Marshal.ReleaseComObject(xlapp)
        xlapp = Nothing
        Process.GetProcessesByName("EXCEL")(0).Kill()
    End Sub

    Private Sub TextBox3_TextChanged(sender As Object, e As EventArgs) Handles TextBox3.TextChanged
        TextBox3.PasswordChar = "*"
        If TextBox3.Text = "" Then
            TextBox4.Enabled = False
        Else
            TextBox4.Enabled = True
        End If
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

    Private Sub ComboBox1_SelectedIndexChanged(sender As Object, e As EventArgs) Handles ComboBox1.SelectedIndexChanged
        TextBox1.Clear()
        TextBox2.Clear()
        TextBox3.Clear()
        TextBox4.Clear()
        Button1.Enabled = False
        TextBox4.BackColor = Color.White
        If ComboBox1.Text = "" Then
            TextBox1.Enabled = False
        Else
            TextBox1.Enabled = True
        End If
    End Sub

    Private Sub TextBox1_TextChanged(sender As Object, e As EventArgs) Handles TextBox1.TextChanged
        If TextBox1.Text = "" Then
            TextBox2.Enabled = False
        Else
            TextBox2.Enabled = True
        End If
    End Sub

    Private Sub TextBox2_TextChanged(sender As Object, e As EventArgs) Handles TextBox2.TextChanged
        If TextBox2.Text = "" Then
            TextBox3.Enabled = False
        Else
            TextBox3.Enabled = True
        End If
    End Sub

    Private Sub Button1_Click(sender As Object, e As EventArgs) Handles Button1.Click
        Dim line1 As Label
        Dim xlapp = New Microsoft.Office.Interop.Excel.Application
        Dim workbook = xlapp.Workbooks.Open(CONFIG.TextBox1.Text.ToString)
        Dim ws As Worksheet
        ws = workbook.Sheets("USER DATA")
        If ComboBox1.SelectedItem = ws.Range("A2").Value Then
            ws.Range("A2").Value = TextBox1.Text
            ws.Range("B2").Value = TextBox2.Text
            ws.Range("C2").Value = TextBox3.Text
        Else
            If ComboBox1.SelectedItem = ws.Range("A3").Value Then
                ws.Range("A3").Value = TextBox1.Text
                ws.Range("B3").Value = TextBox2.Text
                ws.Range("C3").Value = TextBox3.Text
            Else
                If ComboBox1.SelectedItem = ws.Range("A4").Value Then
                    ws.Range("A4").Value = TextBox1.Text
                    ws.Range("B4").Value = TextBox2.Text
                    ws.Range("C4").Value = TextBox3.Text
                Else
                    If ComboBox1.SelectedItem = ws.Range("A5").Value Then
                        ws.Range("A5").Value = TextBox1.Text
                        ws.Range("B5").Value = TextBox2.Text
                        ws.Range("C5").Value = TextBox3.Text
                    Else
                        If ComboBox1.SelectedItem = ws.Range("A6").Value Then
                            ws.Range("A6").Value = TextBox1.Text
                            ws.Range("B6").Value = TextBox2.Text
                            ws.Range("C6").Value = TextBox3.Text
                        Else
                            If ComboBox1.SelectedItem = ws.Range("A7").Value Then
                                ws.Range("A7").Value = TextBox1.Text
                                ws.Range("B7").Value = TextBox2.Text
                                ws.Range("C7").Value = TextBox3.Text
                            Else
                                If ComboBox1.SelectedItem = ws.Range("A8").Value Then
                                    ws.Range("A8").Value = TextBox1.Text
                                    ws.Range("B8").Value = TextBox2.Text
                                    ws.Range("C8").Value = TextBox3.Text
                                Else
                                    If ComboBox1.SelectedItem = ws.Range("A9").Value Then
                                        ws.Range("A9").Value = TextBox1.Text
                                        ws.Range("B9").Value = TextBox2.Text
                                        ws.Range("C9").Value = TextBox3.Text
                                    Else
                                        If ComboBox1.SelectedItem = ws.Range("A10").Value Then
                                            ws.Range("A10").Value = TextBox1.Text
                                            ws.Range("B10").Value = TextBox2.Text
                                            ws.Range("C10").Value = TextBox3.Text
                                        Else
                                            If ComboBox1.SelectedItem = ws.Range("A11").Value Then
                                                ws.Range("A11").Value = TextBox1.Text
                                                ws.Range("B11").Value = TextBox2.Text
                                                ws.Range("C11").Value = TextBox3.Text
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
        Dim iret As Object = MsgBox("USER CREATEC SUCCESSFULLY", vbInformation + vbOKOnly, "USER CREATION")
        If iret = vbOK Then
            TextBox1.Clear()
            TextBox2.Clear()
            TextBox3.Clear()
            TextBox4.Clear()
            ComboBox1.Text = ""
            ComboBox1.Items.Clear()
        End If
line1:
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
        TextBox1.Clear()
        TextBox2.Clear()
        TextBox3.Clear()
        TextBox4.Clear()
        ComboBox1.Items.Clear()
        Me.Close()
        Call SETTING.Show()
    End Sub
End Class