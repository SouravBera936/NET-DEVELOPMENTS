Imports Microsoft.Office.Interop.Excel
Public Class REMUSER
    Private Sub REMUSER_Load(sender As Object, e As EventArgs) Handles MyBase.Load
        ListBox1.Items.Clear()
        Dim line1 As Label
        Dim xlapp = New Microsoft.Office.Interop.Excel.Application
        Dim workbook = xlapp.Workbooks.Open(CONFIG.TextBox1.Text.ToString)
        Dim ws As Worksheet
        ws = workbook.Sheets("USER DATA")
        If ws.Range("B2").Value = Nothing Then
            'donothing
        Else
            ListBox1.Items.Add(ws.Range("A2").Value)
        End If
        If ws.Range("B3").Value = "" Then
            'donothing
            GoTo line1
        Else
            ListBox1.Items.Add(ws.Range("A3").Value)
            GoTo line1
        End If
        If ws.Range("B4").Value = "" Then
            'donothing
            GoTo line1
        Else
            ListBox1.Items.Add(ws.Range("A4").Value)
            GoTo line1
        End If
        If ws.Range("B5").Value = "" Then
            'donothing
            GoTo line1
        Else
            ListBox1.Items.Add(ws.Range("A5").Value)
            GoTo line1
        End If
        If ws.Range("B6").Value = "" Then
            'donothing
            GoTo line1
        Else
            ListBox1.Items.Add(ws.Range("A6").Value)
            GoTo line1
        End If
        If ws.Range("B7").Value = "" Then
            'donothing
            GoTo line1
        Else
            ListBox1.Items.Add(ws.Range("A7").Value)
            GoTo line1
        End If
        If ws.Range("B8").Value = "" Then
            'donothing
            GoTo line1
        Else
            ListBox1.Items.Add(ws.Range("A8").Value)
            GoTo line1
        End If
        If ws.Range("B9").Value = "" Then
            'donothing
            GoTo line1
        Else
            ListBox1.Items.Add(ws.Range("A9").Value)
            GoTo line1
        End If
        If ws.Range("B10").Value = "" Then
            'donothing
            GoTo line1
        Else
            ListBox1.Items.Add(ws.Range("A10").Value)
            GoTo line1
        End If
        If ws.Range("B11").Value = "" Then
            'donothing
            GoTo line1
        Else
            ListBox1.Items.Add(ws.Range("A11").Value)
            GoTo line1
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

    Private Sub Button2_Click(sender As Object, e As EventArgs) Handles Button2.Click
        Me.Close()
        Call SETTING.Show()
    End Sub

    Private Sub Button1_Click(sender As Object, e As EventArgs) Handles Button1.Click
        Dim xlapp = New Microsoft.Office.Interop.Excel.Application
        Dim workbook = xlapp.Workbooks.Open(CONFIG.TextBox1.Text.ToString)
        Dim ws As Worksheet
        ws = workbook.Sheets("USER DATA")
        If ListBox1.SelectedItem = ws.Range("A2").Value Then
            ws.Range("B2", "C2").ClearContents()
            ws.Range("A2").Value = "USER 1"
            GoTo line1
        Else
            If ListBox1.SelectedItem = ws.Range("A3").Value Then
                ws.Range("B3", "C3").ClearContents()
                ws.Range("A3").Value = "USER 2"
                GoTo line1
            Else
                If ListBox1.SelectedItem = ws.Range("A4").Value Then
                    ws.Range("B4", "C4").ClearContents()
                    ws.Range("A4").Value = "USER 3"
                    GoTo line1
                Else
                    If ListBox1.SelectedItem = ws.Range("A5").Value Then
                        ws.Range("B5", "C5").ClearContents()
                        ws.Range("A5").Value = "USER 4"
                        GoTo line1
                    Else
                        If ListBox1.SelectedItem = ws.Range("A6").Value Then
                            ws.Range("B6", "C6").ClearContents()
                            ws.Range("A6").Value = "USER 5"
                            GoTo line1
                        Else
                            If ListBox1.SelectedItem = ws.Range("A7").Value Then
                                ws.Range("B7", "C7").ClearContents()
                                ws.Range("A7").Value = "USER 6"
                                GoTo line1
                            Else
                                If ListBox1.SelectedItem = ws.Range("A8").Value Then
                                    ws.Range("B8", "C8").ClearContents()
                                    ws.Range("A8").Value = "USER 7"
                                    GoTo line1
                                Else
                                    If ListBox1.SelectedItem = ws.Range("A9").Value Then
                                        ws.Range("B9", "C9").ClearContents()
                                        ws.Range("A9").Value = "USER 8"
                                        GoTo line1
                                    Else
                                        If ListBox1.SelectedItem = ws.Range("A10").Value Then
                                            ws.Range("B10", "C10").ClearContents()
                                            ws.Range("A10").Value = "USER 9"
                                            GoTo line1
                                        Else
                                            If ListBox1.SelectedItem = ws.Range("A11").Value Then
                                                ws.Range("B11", "C11").ClearContents()
                                                ws.Range("A11").Value = "USER 10"
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
        Dim iret As Object = MsgBox("USER REMOVED SUCCESS FULLY", vbOKOnly + vbInformation, "USER REMOVAL")
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
End Class