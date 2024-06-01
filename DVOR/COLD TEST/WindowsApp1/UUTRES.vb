Imports Microsoft.Office.Interop.Excel
Public Class UUTRES
    Private Sub UUTRES_Load(sender As Object, e As EventArgs) Handles MyBase.Load
        If TESTFORM.Button1.Text = "PASS" Then
            Me.BackColor = Color.Green
            Me.Label1.Text = "TEST SEQUENCE PASSED"
        Else
            If TESTFORM.Button1.Text = "FAIL" Then
                Me.BackColor = Color.Red
                Me.Label1.Text = "TEST SEQUENCE FAILED"
            End If
        End If
    End Sub

    Private Sub Button1_Click(sender As Object, e As EventArgs) Handles Button1.Click
        Dim line1 As Label
        Try
line1:
            If TESTFORM.ComboBox1.SelectedItem = "T11120920-0001-RF-COUPLER" Then
                Dim xlapp = New Microsoft.Office.Interop.Excel.Application
                Dim workbook = xlapp.Workbooks.Open(CONFIG.TextBox3.Text.ToString)
                Dim ws As Worksheet
                ws = workbook.Sheets("COUPLER")
                ws.Range("D6").Value = TESTFORM.TextBox3.Text & "_" & TESTFORM.TextBox4.Text
                ws.Range("D7").Value = "CONTINUITY TEST"
                ws.Range("D8").Value = TESTFORM.TextBox7.Text
                ws.Range("D9").Value = DateTime.Now.ToString
                ws.Range("D10").Value = TESTFORM.Button1.Text
                ws.Range("G13").Value = TESTFORM.TextBox19.Text
                ws.Range("J13").Value = TESTFORM.Button10.Text
                ws.Range("G14").Value = TESTFORM.TextBox24.Text
                ws.Range("J14").Value = TESTFORM.Button11.Text
                ws.Range("G15").Value = TESTFORM.TextBox37.Text
                ws.Range("J15").Value = TESTFORM.Button12.Text
                ws.Range("G16").Value = TESTFORM.TextBox50.Text
                ws.Range("J16").Value = TESTFORM.Button13.Text
                ws.Range("G17").Value = TESTFORM.TextBox63.Text
                ws.Range("J17").Value = TESTFORM.Button14.Text
                ws.Range("G18").Value = TESTFORM.TextBox76.Text
                ws.Range("J18").Value = TESTFORM.Button15.Text
                ws.Range("F6").Value = TESTFORM.TextBox5.Text
                ws.Range("F7").Value = TESTFORM.TextBox6.Text
                With ws.PageSetup
                    .Orientation = XlPageOrientation.xlLandscape
                    .Zoom = False
                    .FitToPagesTall = 1
                    .FitToPagesWide = 1
                End With
                workbook.ExportAsFixedFormat(XlFixedFormatType.xlTypePDF, CONFIG.TextBox4.Text.ToString & TESTFORM.TextBox7.Text.ToString & "_" & TESTFORM.Button1.Text.ToString & Format(Now(), "hhmmssmmddyy") & ".pdf", vbNormal)
                workbook.Application.DisplayAlerts = False
                workbook.Close(True)
                System.Runtime.InteropServices.Marshal.ReleaseComObject(workbook)
                workbook = Nothing
                xlapp.Quit()
                System.Runtime.InteropServices.Marshal.ReleaseComObject(xlapp)
                xlapp = Nothing
                Process.GetProcessesByName("EXCEL")(0).Kill()
                TESTFORM.ComboBox1.Enabled = True
                TESTFORM.Button2.Enabled = True
                TESTFORM.TextBox7.Clear()
                TESTFORM.TextBox19.Clear()
                TESTFORM.Button10.Text = "IDLE"
                TESTFORM.TextBox24.Clear()
                TESTFORM.Button11.Text = "IDLE"
                TESTFORM.TextBox37.Clear()
                TESTFORM.Button12.Text = "IDLE"
                TESTFORM.TextBox50.Clear()
                TESTFORM.Button13.Text = "IDLE"
                TESTFORM.TextBox63.Clear()
                TESTFORM.Button14.Text = "IDLE"
                TESTFORM.TextBox76.Clear()
                TESTFORM.Button15.Text = "IDLE"
                TESTFORM.Button4.BackColor = Color.White
                TESTFORM.Button5.BackColor = Color.White
                TESTFORM.Button6.BackColor = Color.White
                TESTFORM.Button7.BackColor = Color.White
                TESTFORM.Button8.BackColor = Color.White
                TESTFORM.Button9.BackColor = Color.White
                TESTFORM.TabControl2.SelectedTab = TESTFORM.TabPage3
                TESTFORM.Button1.Text = "IDLE"
                TESTFORM.TextBox6.Clear()
                TESTFORM.TextBox7.Clear()
                Me.Close()
            Else
                If TESTFORM.ComboBox1.SelectedItem = "120888-0001 STATUS PANEL" Then
                    Dim xlapp1 = New Microsoft.Office.Interop.Excel.Application
                    Dim workbook1 = xlapp1.Workbooks.Open(CONFIG.TextBox13.Text.ToString)
                    Dim ws1 As Worksheet
                    ws1 = workbook1.Sheets("STATUS")
                    ws1.Range("D6").Value = TESTFORM.TextBox3.Text & "_" & TESTFORM.TextBox4.Text
                    ws1.Range("D7").Value = "COLD TEST"
                    ws1.Range("D8").Value = TESTFORM.TextBox7.Text
                    ws1.Range("D9").Value = DateTime.Now.ToString
                    ws1.Range("D10").Value = TESTFORM.Button1.Text
                    ws1.Range("G13").Value = TESTFORM.TextBox154.Text
                    ws1.Range("J13").Value = TESTFORM.Button39.Text
                    ws1.Range("F6").Value = TESTFORM.TextBox5.Text
                    ws1.Range("F7").Value = TESTFORM.TextBox6.Text
                    With ws1.PageSetup
                        .Orientation = XlPageOrientation.xlLandscape
                        .Zoom = False
                        .FitToPagesTall = 1
                        .FitToPagesWide = 1
                    End With
                    workbook1.ExportAsFixedFormat(XlFixedFormatType.xlTypePDF, CONFIG.TextBox12.Text.ToString & TESTFORM.TextBox7.Text.ToString & "_" & TESTFORM.Button1.Text.ToString & Format(Now(), "hhmmssmmddyy") & ".pdf", vbNormal)
                    workbook1.Application.DisplayAlerts = False
                    workbook1.Close(True)
                    System.Runtime.InteropServices.Marshal.ReleaseComObject(workbook1)
                    workbook1 = Nothing
                    xlapp1.Quit()
                    System.Runtime.InteropServices.Marshal.ReleaseComObject(xlapp1)
                    xlapp1 = Nothing
                    Process.GetProcessesByName("EXCEL")(0).Kill()
                    TESTFORM.ComboBox1.Enabled = True
                    TESTFORM.Button2.Enabled = True
                    TESTFORM.TextBox7.Clear()
                    TESTFORM.TextBox154.Clear()
                    TESTFORM.Button39.Text = "IDLE"
                    TESTFORM.Button38.BackColor = Color.White
                    TESTFORM.TabControl4.SelectedTab = TESTFORM.TabPage15
                    TESTFORM.Button1.Text = "IDLE"
                    TESTFORM.TextBox6.Clear()
                    TESTFORM.TextBox7.Clear()
                    Me.Close()
                End If
            End If
        Catch ex As Exception
            Dim iret As Object = MsgBox(ex.Message, vbRetryCancel + vbCritical, "ERROR GENERATING REPORT")
            If iret = vbRetry Then
                GoTo line1
            Else
                If iret = vbCancel Then
                    Exit Sub
                End If
            End If
        End Try


    End Sub
End Class