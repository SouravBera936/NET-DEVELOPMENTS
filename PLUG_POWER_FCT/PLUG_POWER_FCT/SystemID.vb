Imports Microsoft.Office.Interop
Imports System.Data.OleDb
Module SystemID
    Dim myProcesses() As Process
    Dim myProcess As Process
    Dim cn As OleDb.OleDbConnection
    Dim oleCommand As OleDbCommand
    Dim CapturedImg As Bitmap
    Dim Reference As Bitmap
    Dim file1 As System.IO.StreamWriter
    Dim file2 As System.IO.StreamWriter
    Dim SystIDBox = {MAIN.TextBox7, MAIN.TextBox8, MAIN.TextBox9, MAIN.TextBox10, MAIN.TextBox11, MAIN.TextBox12, MAIN.TextBox13, MAIN.TextBox14, MAIN.TextBox15, MAIN.Button8, MAIN.TextBox159}
    Function TestInitialize()
        MAIN.Button3.Enabled = False
        MAIN.Button1.Enabled = True
        FUNNCTIONLIB.CalculateProgress(5, 1)
        MAIN.Button100.Visible = False
        MAIN.Button101.Visible = False
        Dim line1 As Label
        MAIN.ListBox1.Items.Add($"{DateTime.Now.ToString} Software Initialization Started")
        MAIN.ListBox1.SelectedIndex = MAIN.ListBox1.Items.Count - 1
        IMAGE_STR.PictureBox22.Image = Nothing
line1:
        Try
            If MAIN.RadioButton1.Checked = True Then
                MAIN.TabControl3.SelectedTab = MAIN.TabPage9
                MAIN.GroupBox4.Visible = False
                MAIN.GroupBox3.Visible = True
                MAIN.Button1.Enabled = True
                MAIN.Button43.Text = "RUNNING"
                MAIN.TimerFunction("START")
                MAIN.Button3.Enabled = False
                MAIN.GroupBox2.Visible = True
                For Each tb In SystIDBox
                    tb.enabled = False
                Next
                For Each tn In MAIN.TreeView1.Nodes
                    If tn.Text = MAIN.TextBox1.Text Then
                        MAIN.TreeView1.SelectedNode = tn
                        MAIN.TreeView1.SelectedNode.Nodes(0).ExpandAll()
                        Exit For
                    End If
                    FindNode(tn)
                Next
                MAIN.TabControl1.TabPages.Remove(MAIN.TabPage1)
                MAIN.TabControl1.TabPages.Remove(MAIN.TabPage2)
                MAIN.TabControl1.TabPages.Remove(MAIN.TabPage3)
                MAIN.TabControl1.TabPages.Remove(MAIN.TabPage4)
                MAIN.TabControl1.TabPages.Remove(MAIN.TabPage5)
                MAIN.TabControl1.TabPages.Add(MAIN.TabPage1)
                MAIN.TextBox13.Clear()
                MAIN.Button8.Text = "RUNNING"
                MAIN.Button8.Enabled = False
                MAIN.TextBox159.Clear()
                MAIN.TextBox159.Visible = False
                MAIN.Button9.Visible = False
                SystemIDNode0()
            ElseIf MAIN.RadioButton2.Checked = True Then
                MAIN.TreeView1.Enabled = True
                MAIN.GroupBox4.Visible = True
                MAIN.GroupBox3.Visible = True
                MAIN.Button1.Enabled = True
                MAIN.Button43.Text = "RUNNING"
                FUNNCTIONLIB.ProgressReset(True)
                MAIN.TimerFunction("START")
                MAIN.Button3.Enabled = False
                MAIN.GroupBox2.Visible = True
                For Each tb In SystIDBox
                    tb.enabled = False
                Next
                For Each tn In MAIN.TreeView1.Nodes
                    If tn.Text = MAIN.TextBox1.Text Then
                        MAIN.TreeView1.SelectedNode = tn
                        MAIN.TreeView1.SelectedNode.Nodes(0).ExpandAll()
                        Exit For
                    End If
                    FindNode(tn)
                Next
                MAIN.TabControl1.TabPages.Remove(MAIN.TabPage1)
                MAIN.TabControl1.TabPages.Remove(MAIN.TabPage2)
                MAIN.TabControl1.TabPages.Remove(MAIN.TabPage3)
                MAIN.TabControl1.TabPages.Remove(MAIN.TabPage4)
                MAIN.TabControl1.TabPages.Remove(MAIN.TabPage5)
                MAIN.TabControl1.TabPages.Add(MAIN.TabPage1)
                MAIN.TextBox13.Clear()
                MAIN.Button8.Text = "RUNNING"
                MAIN.Button8.Enabled = False
                MAIN.TextBox159.Clear()
                MAIN.TextBox159.Visible = False
                MAIN.Button9.Visible = False
            End If
        Catch ex As Exception
            Dim iret As Object = MsgBox("Error Executing SystemID.TestInitialize : " & ex.Message, vbCritical + vbRetryCancel, "Error Executing SystemID.TestInitialize")
            If iret = vbRetry Then
                GoTo line1
            ElseIf iret = vbCancel Then
                MAIN.Button43.Text = "ERROR"
                MAIN.Button8.Text = "ERROR"
                MAIN.TextBox13.Text = "ERROR"
                MAIN.ListBox1.Items.Add($"{DateTime.Now.ToString} Error Executing SystemID.TestInitialize :" & ex.Message)
                MAIN.ListBox1.SelectedIndex = MAIN.ListBox1.Items.Count - 1
                UUT_RESULT.RadioButton3.Checked = True
                UUT_RESULT.TextBox2.Text = "SYSTEMID_INIT"
                MAIN.TreeView1.SelectedNode.Nodes(0).ExpandAll()
                MAIN.TreeView1.SelectedNode.Nodes(0).Nodes(0).BackColor = Color.Red
                UUT_RESULT.ShowDialog()
            End If
        End Try
    End Function
    Function SystemIDNode0()
        FUNNCTIONLIB.CalculateProgress(5, 2)
        If MAIN.RadioButton1.Checked = True Then
            MAIN.TreeView1.SelectedNode.Nodes(0).Nodes(0).BackColor = Color.Yellow
            MAIN.PictureBox1.Visible = True
            MAIN.PictureBox1.Image = Image.FromFile(CONFIG.TextBox15.Text.ToString)
            MAIN.TextBox159.Visible = True
            MAIN.TextBox159.Text = "Connect The DUT as Shown in Below Figure"
            MAIN.Button9.Visible = True
            MAIN.TreeView1.SelectedNode.Nodes(0).Nodes(0).BackColor = Color.Green
        Else
            MAIN.PictureBox1.Image = Image.FromFile(CONFIG.TextBox15.Text.ToString)
            MAIN.PictureBox1.Visible = True
            MAIN.TextBox159.Visible = True
            MAIN.TextBox159.Text = "Connect The DUT as Shown in Below Figure"
            MAIN.Button9.Visible = True
        End If
    End Function
    Function SystemIDNode1()
        FUNNCTIONLIB.CalculateProgress(5, 3)
        Dim line1 As Label
line1:
        Try
            If MAIN.RadioButton1.Checked = True Then
                MAIN.TreeView1.SelectedNode.Nodes(0).Nodes(1).BackColor = Color.Yellow
                myProcesses = Process.GetProcessesByName(CONFIG.TextBox22.Text.ToString)
                If myProcesses.Length > 0 Then
                    MAIN.ListBox1.Items.Add($"{DateTime.Now.ToString} SystemID Programmer Running Status : {myProcesses.Length}")
                    MAIN.ListBox1.SelectedIndex = MAIN.ListBox1.Items.Count - 1
                    For Each myProcess In myProcesses
                        If myProcess IsNot Nothing Then
                            myProcess.Kill()
                        End If
                    Next
                End If
                System.Threading.Thread.Sleep(2000)
                Dim p As Process = Process.Start(CONFIG.TextBox16.Text.ToString)
                myProcesses = Process.GetProcessesByName(CONFIG.TextBox22.Text.ToString)
                System.Threading.Thread.Sleep(5000)
                If myProcesses.Length > 0 Then
                    MAIN.ListBox1.Items.Add($"{DateTime.Now.ToString} PROGRAMME EXECUTED SUCCESSFULLY")
                    MAIN.ListBox1.SelectedIndex = MAIN.ListBox1.Items.Count - 1
                    System.Threading.Thread.Sleep(3000)
                    SendKeys.Send("^(/)")
                    MAIN.ListBox1.Items.Add($"{DateTime.Now.ToString} SEND KEY COMMAND : ^(/)")
                    MAIN.ListBox1.SelectedIndex = MAIN.ListBox1.Items.Count - 1
                    MAIN.TreeView1.SelectedNode.Nodes(0).Nodes(1).BackColor = Color.Green
                    SystemIDNode2()
                Else
                    Dim iret1 = MsgBox("Error Executing SystemID Programmer", vbRetryCancel + vbCritical, "Error Launching SystemID Programmer")
                    If iret1 = vbRetry Then
                        MAIN.ListBox1.Items.Add($"{DateTime.Now.ToString} Error Launching SystemID Programmer")
                        MAIN.ListBox1.SelectedIndex = MAIN.ListBox1.Items.Count - 1
                        GoTo line1
                    ElseIf iret1 = vbCancel Then
                        MAIN.TextBox13.Text = "FALSE"
                        MAIN.Button8.Text = "FAIL"
                        MAIN.Button43.Text = "FAIL"
                        MAIN.ListBox1.Items.Add($"{DateTime.Now.ToString} Error Launching SystemID Programmer")
                        MAIN.ListBox1.SelectedIndex = MAIN.ListBox1.Items.Count - 1
                        UUT_RESULT.RadioButton2.Checked = True
                        UUT_RESULT.TextBox2.Text = "SYSTEMID_NODE1"
                        MAIN.TreeView1.SelectedNode.Nodes(0).ExpandAll()
                        MAIN.TreeView1.SelectedNode.Nodes(0).Nodes(1).BackColor = Color.Red
                        UUT_RESULT.ShowDialog()
                    End If
                End If
            Else
                myProcesses = Process.GetProcessesByName(CONFIG.TextBox22.Text.ToString)
                If myProcesses.Length > 0 Then
                    MAIN.ListBox1.Items.Add($"{DateTime.Now.ToString} SystemID Programmer Running Status : {myProcesses.Length}")
                    MAIN.ListBox1.SelectedIndex = MAIN.ListBox1.Items.Count - 1
                    For Each myProcess In myProcesses
                        If myProcess IsNot Nothing Then
                            myProcess.Kill()
                        End If
                    Next
                End If
                System.Threading.Thread.Sleep(2000)
                Dim p As Process = Process.Start(CONFIG.TextBox16.Text.ToString)
                myProcesses = Process.GetProcessesByName(CONFIG.TextBox22.Text.ToString)
                If myProcesses.Length > 0 Then
                    MAIN.ListBox1.Items.Add($"{DateTime.Now.ToString} PROGRAMME EXECUTED SUCCESSFULLY")
                    MAIN.ListBox1.SelectedIndex = MAIN.ListBox1.Items.Count - 1
                    System.Threading.Thread.Sleep(7000)
                    SendKeys.Send("^(/)")
                    MAIN.ListBox1.Items.Add($"{DateTime.Now.ToString} SEND KEY COMMAND : ^(/)")
                    MAIN.ListBox1.SelectedIndex = MAIN.ListBox1.Items.Count - 1
                Else
                    Dim iret1 = MsgBox("Error Executing SystemID Programmer", vbRetryCancel + vbCritical, "Error Launching SystemID Programmer")
                    If iret1 = vbRetry Then
                        MAIN.ListBox1.Items.Add($"{DateTime.Now.ToString} Error Launching SystemID Programmer :")
                        MAIN.ListBox1.SelectedIndex = MAIN.ListBox1.Items.Count - 1
                        GoTo line1
                    ElseIf iret1 = vbCancel Then
                        MAIN.TextBox13.Text = "FALSE"
                        MAIN.Button8.Text = "FAIL"
                        MAIN.Button43.Text = "FAIL"
                        MAIN.ListBox1.Items.Add($"{DateTime.Now.ToString} Error Launching SystemID Programmer :")
                        MAIN.ListBox1.SelectedIndex = MAIN.ListBox1.Items.Count - 1
                        Exit Function
                    End If
                End If
            End If
        Catch ex As Exception
            Dim iret As Object = MsgBox("Error Executing SystemID.SystemIDNode1 : " & ex.Message, vbCritical + vbRetryCancel, "Error Executing SystemID.SystemIDNode1")
            If iret = vbRetry Then
                MAIN.ListBox1.Items.Add($"{DateTime.Now.ToString} Error Executing SystemID.SystemIDNode1 " & ex.Message)
                MAIN.ListBox1.SelectedIndex = MAIN.ListBox1.Items.Count - 1
                GoTo line1
            ElseIf iret = vbCancel Then
                MAIN.TextBox13.Text = "ERROR"
                MAIN.Button8.Text = "ERROR"
                MAIN.Button43.Text = "ERROR"
                MAIN.ListBox1.Items.Add($"{DateTime.Now.ToString} Error Executing SystemID.SystemIDNode1 " & ex.Message)
                MAIN.ListBox1.SelectedIndex = MAIN.ListBox1.Items.Count - 1
                UUT_RESULT.RadioButton3.Checked = True
                UUT_RESULT.TextBox2.Text = "SYSTEMID_NODE1"
                MAIN.TreeView1.SelectedNode.Nodes(0).ExpandAll()
                MAIN.TreeView1.SelectedNode.Nodes(0).Nodes(1).BackColor = Color.FromArgb(255, 127, 127)
                UUT_RESULT.ShowDialog()
            End If
        End Try
    End Function
    Function SystemIDNode2()
        FUNNCTIONLIB.CalculateProgress(5, 4)
        Dim line1 As Label
line1:
        Try
            If MAIN.RadioButton1.Checked = True Then
                MAIN.TreeView1.SelectedNode.Nodes(0).Nodes(2).BackColor = Color.Yellow
                If SETUP.CheckBox6.Checked = True Then
                    MAIN.ListBox1.Items.Add($"{DateTime.Now.ToString} AUTO COMPARISON FUNCTION CHECK : ENABLED")
                    MAIN.ListBox1.SelectedIndex = MAIN.ListBox1.Items.Count - 1
                    FUNNCTIONLIB.RestoreApp(CONFIG.TextBox90.Text)
                    System.Threading.Thread.Sleep(2000)
                    CapturedImg = FUNNCTIONLIB.TakeSnap(True)
                    System.Threading.Thread.Sleep(2000)
                    IMAGE_STR.PictureBox22.Image = CapturedImg
                    MAIN.ListBox1.Items.Add($"{DateTime.Now.ToString} FUNNCTIONLIB.TakeSnapofActiveApp Executed SuccessFully")
                    MAIN.ListBox1.SelectedIndex = MAIN.ListBox1.Items.Count - 1
                    myProcesses = Process.GetProcessesByName(CONFIG.TextBox22.Text.ToString)
                    System.Threading.Thread.Sleep(2000)
                    If myProcesses.Length > 0 Then
                        For Each myProcess In myProcesses
                            If myProcess IsNot Nothing Then
                                myProcess.Kill()
                            End If
                        Next
                    End If
                    MAIN.TreeView1.SelectedNode.Nodes(0).Nodes(2).BackColor = Color.Green
                    SystemIDNode3()
                ElseIf SETUP.CheckBox6.Checked = False Then
                    MAIN.ListBox1.Items.Add($"{DateTime.Now.ToString} AUTO COMPARISON FUNCTION CHECK : DISABLED")
                    MAIN.ListBox1.SelectedIndex = MAIN.ListBox1.Items.Count - 1
                    MAIN.Button100.Visible = True
                    MAIN.Button101.Visible = True
                    MAIN.PictureBox1.Image = Image.FromFile(CONFIG.TextBox19.Text.ToString)
                    MAIN.PictureBox1.Visible = True
                    MAIN.TextBox159.Text = "Compare the Snap And If Found Same Click On CONFIRM Otherwise Click On NONCONFIRM"
                    MAIN.TextBox159.Visible = True
                End If
            End If
        Catch ex As Exception
            Dim iret As Object = MsgBox("Error Executing SystemID.SystemIDNode2 : " & ex.Message, vbCritical + vbRetryCancel, "Error Executing SystemID.SystemIDNode2")
            If iret = vbRetry Then
                MAIN.ListBox1.Items.Add($"{DateTime.Now.ToString} Error Executing SystemID.SystemIDNode2 " & ex.Message)
                MAIN.ListBox1.SelectedIndex = MAIN.ListBox1.Items.Count - 1
                GoTo line1
            ElseIf iret = vbCancel Then
                MAIN.TextBox13.Text = "ERROR"
                MAIN.Button8.Text = "ERROR"
                MAIN.Button43.Text = "ERROR"
                MAIN.ListBox1.Items.Add($"{DateTime.Now.ToString} Error Executing SystemID.SystemIDNode2 " & ex.Message)
                MAIN.ListBox1.SelectedIndex = MAIN.ListBox1.Items.Count - 1
                UUT_RESULT.RadioButton3.Checked = True
                UUT_RESULT.TextBox2.Text = "SYSTEMID_NODE2"
                MAIN.TreeView1.SelectedNode.Nodes(0).ExpandAll()
                MAIN.TreeView1.SelectedNode.Nodes(0).Nodes(2).BackColor = Color.FromArgb(255, 127, 127)
                UUT_RESULT.ShowDialog()
            End If
        End Try
    End Function
    Function SystemIDNode3()
        FUNNCTIONLIB.CalculateProgress(5, 5)
        Reference = Image.FromFile(CONFIG.TextBox19.Text)
        Dim line1 As Label
line1:
        Try
            If MAIN.RadioButton1.Checked = True Then
                MAIN.TreeView1.SelectedNode.Nodes(0).Nodes(3).BackColor = Color.Yellow
                If FUNNCTIONLIB.CompareImages(CapturedImg, Reference) = True Then
                    MAIN.ListBox1.Items.Add($"{DateTime.Now.ToString} AUTO COMPARISON FUNCTION RETURNED : True")
                    MAIN.ListBox1.SelectedIndex = MAIN.ListBox1.Items.Count - 1
                    MAIN.TextBox13.Text = "TRUE"
                    MAIN.Button8.Text = "PASS"
                    MAIN.Button43.Text = "PASS"
                    MAIN.TreeView1.SelectedNode.Nodes(0).Nodes(3).BackColor = Color.Green
                    UUT_RESULT.RadioButton1.Checked = True
                    UUT_RESULT.TextBox2.Text = "SYSTEMID_NODE3"
                    UUT_RESULT.ShowDialog()
                Else
                    MAIN.ListBox1.Items.Add($"{DateTime.Now.ToString} AUTO COMPARISON FUNCTION RETURNED : False")
                    MAIN.ListBox1.SelectedIndex = MAIN.ListBox1.Items.Count - 1
                    MAIN.TextBox13.Text = "FALSE"
                    MAIN.Button8.Text = "FAIL"
                    MAIN.Button43.Text = "FAIL"
                    MAIN.TreeView1.SelectedNode.Nodes(0).Nodes(3).BackColor = Color.Red
                    UUT_RESULT.RadioButton2.Checked = True
                    UUT_RESULT.TextBox2.Text = "SYSTEMID_NODE3"
                    UUT_RESULT.ShowDialog()
                End If
            End If
        Catch ex As Exception
            Dim iret As Object = MsgBox("Error Executing SystemID.SystemIDNode3 : " & ex.Message, vbCritical + vbRetryCancel, "Error Executing SystemID.SystemIDNode3")
            If iret = vbRetry Then
                MAIN.ListBox1.Items.Add($"{DateTime.Now.ToString} Error Executing SystemID.SystemIDNode3 " & ex.Message)
                MAIN.ListBox1.SelectedIndex = MAIN.ListBox1.Items.Count - 1
                GoTo line1
            ElseIf iret = vbCancel Then
                MAIN.TextBox13.Text = "ERROR"
                MAIN.Button8.Text = "ERROR"
                MAIN.Button43.Text = "ERROR"
                MAIN.ListBox1.Items.Add($"{DateTime.Now.ToString} Error Executing SystemID.SystemIDNode3 " & ex.Message)
                MAIN.ListBox1.SelectedIndex = MAIN.ListBox1.Items.Count - 1
                UUT_RESULT.RadioButton3.Checked = True
                UUT_RESULT.TextBox2.Text = "SYSTEMID_NODE3"
                MAIN.TreeView1.SelectedNode.Nodes(0).ExpandAll()
                MAIN.TreeView1.SelectedNode.Nodes(0).Nodes(3).BackColor = Color.FromArgb(255, 127, 127)
                UUT_RESULT.ShowDialog()
            End If
        End Try
    End Function
    Private Sub FindNode(ByVal tNode As TreeNode)
        Dim tn As TreeNode
        For Each tn In tNode.Nodes
            If tn.Text = MAIN.TextBox1.Text Then
                MAIN.TreeView1.SelectedNode = tn
                MAIN.TreeView1.SelectedNode.ExpandAll()
                Exit For
            End If
            FindNode(tn)
        Next
    End Sub
    Function GenerateLogReport(Path As String, SerialNumber As String)
        Dim SW As IO.StreamWriter = IO.File.CreateText(Path & SerialNumber & "_" & Format(Now(), "hhmmssmmddyy"))
        For Each S As String In MAIN.ListBox1.Items
            SW.WriteLine(S)
        Next
        SW.Close()
    End Function
    Function CreateRecord(Provider As String, DataSource As String, Secqurity As String, SerialNumber As String, TestMethode As String, User As String, DutID As String, Mark As String, ProgrammeTest As String, Result As String, CycleTime As String, TesterID As String)
        Dim line1 As Label
line1:
        Dim path1 As String = CONFIG.TextBox4.Text
        file1 = My.Computer.FileSystem.OpenTextFileWriter(path1, True)
        Try
            cn = New OleDbConnection
            cn.ConnectionString = "Provider=" & Provider & ";Data Source=" & DataSource & "; Persist Security Info=" & Secqurity & ";"
            cn.Open()
            Dim SQL1 As String = "INSERT INTO SYSTEM_ID (SERIAL_NUMBER, TEST_STEP, DONE_BY, DUT_ID, MARK, PROGRAMMING_TEST, TEST_RESULT, DATE_TIME, CYCLE_TIME, TESTER_ID ) VALUES ('" & SerialNumber & "','" & TestMethode & "','" & User & "','" & DutID & "','" & Mark & "','" & ProgrammeTest & "','" & Result & "','" & DateTime.Now & "','" & CycleTime & "','" & TesterID & "')"
            Dim command As New OleDbCommand(SQL1, cn)
            command.ExecuteNonQuery()
            MAIN.ListBox1.Items.Add($"{DateTime.Now.ToString} Test Record Generated Success : " & SerialNumber & TestMethode & User & DutID & Mark & ProgrammeTest & Result & DateTime.Now & CycleTime & TesterID)
            MAIN.ListBox1.SelectedIndex = MAIN.ListBox1.Items.Count - 1
            file1.Close()
        Catch ex As Exception
            Dim iret1 As Object = MsgBox("ERROR WHILE CREATING TEST RECORD: " & ex.Message, vbRetryCancel, "SystemID.CreateRecord error")
            MAIN.ListBox1.Items.Add("Error Creating Test Report : " & ex.Message)
            MAIN.ListBox1.SelectedIndex = MAIN.ListBox1.Items.Count - 1
            If iret1 = vbRetry Then
                file1.Close()
                GoTo line1
            ElseIf iret1 = vbCancel Then
                MAIN.ListBox1.Items.Add("Error Creating Test Report : " & ex.Message)
                MAIN.ListBox1.SelectedIndex = MAIN.ListBox1.Items.Count - 1
                file1.WriteLine("SystemID.CreateRecord FAILED DUE TO : " & ex.Message & " " & "@" & DateTime.Now.ToString("yyyy/MM/dd/HH:mm:ss"))
                file1.Close()
                Exit Function
            End If
        End Try
    End Function
    Function GenerateReport(User As String, TestMethode As String, SerialNumber As String, Status As String, CycleTime As String, TesterID As String, TestStatus As String, Errormsg As String)
        Dim line1 As Label
line1:
        Dim path1 As String = CONFIG.TextBox4.Text
        file1 = My.Computer.FileSystem.OpenTextFileWriter(path1, True)
        Try
            Dim xlapp = New Excel.Application
            xlapp.DisplayAlerts = False
            Dim workbook = xlapp.Workbooks.Open(CONFIG.TextBox20.Text.ToString)
            Dim worksheet As Excel.Worksheet
            Dim worksheet1 As Excel.Worksheet
            worksheet = workbook.Sheets("SYSTEM-ID")
            worksheet1 = workbook.Sheets.Add(After:=worksheet)
            worksheet.Range("D6").Value = User
            worksheet.Range("D7").Value = TestMethode
            worksheet.Range("D8").Value = SerialNumber
            worksheet.Range("D9").Value = DateTime.Now.ToString
            worksheet.Range("D10").Value = TestStatus
            worksheet.Range("F6").Value = CycleTime
            worksheet.Range("F7").Value = TesterID
            worksheet.Range("G10").Value = Errormsg
            worksheet.Range("G13").Value = Status
            Clipboard.Clear()
            Clipboard.SetImage(IMAGE_STR.PictureBox22.Image)
            worksheet1.Paste()
            With worksheet1.PageSetup
                .Orientation = Excel.XlPageOrientation.xlLandscape
                .Zoom = False
                .FitToPagesTall = 1
                .FitToPagesWide = 1
            End With
            workbook.ExportAsFixedFormat(Excel.XlFixedFormatType.xlTypePDF, CONFIG.TextBox21.Text.ToString & SerialNumber & "_" & TestMethode & "_" & TestStatus & Format(Now(), "hhmmssmmddyy") & ".pdf", vbNormal)
            worksheet1.Delete()
            Clipboard.Clear()
            workbook.Close(True)
            System.Runtime.InteropServices.Marshal.ReleaseComObject(workbook)
            workbook = Nothing
            xlapp.Quit()
            System.Runtime.InteropServices.Marshal.ReleaseComObject(xlapp)
            xlapp = Nothing
            Process.GetProcessesByName("EXCEL")(0).Kill()
            MAIN.ListBox1.Items.Add("Test Report Generated success : " & CONFIG.TextBox21.Text.ToString & SerialNumber & "_" & TestMethode & "_" & Status & Format(Now(), "hhmmssmmddyy"))
            file1.Close()
        Catch ex As Exception
            Dim iret1 As Object = MsgBox("ERROR WHILE GENERATING TEST REPORT: " & ex.Message, vbRetryCancel, "SystemID.GenerateReport error")
            MAIN.ListBox1.Items.Add("Error Generating Test Report : " & ex.Message)
            If iret1 = vbRetry Then
                file1.Close()
                GoTo line1
            ElseIf iret1 = vbCancel Then
                file1.WriteLine("SystemID.GenerateReport FAILED DUE TO : " & ex.Message & " " & "@" & DateTime.Now.ToString("yyyy/MM/dd/HH:mm:ss"))
                file1.Close()
                Exit Function
            End If
        End Try

    End Function
    Function GenerateReportOnErrorOfInit(User As String, TestMethode As String, SerialNumber As String, Status As String, CycleTime As String, TesterID As String, TestStatus As String, Errormsg As String)

        Dim line1 As Label
line1:
        Dim path1 As String = CONFIG.TextBox4.Text
        file1 = My.Computer.FileSystem.OpenTextFileWriter(path1, True)
        Try
            Dim xlapp = New Excel.Application
            xlapp.DisplayAlerts = False
            Dim workbook = xlapp.Workbooks.Open(CONFIG.TextBox20.Text.ToString)
            Dim worksheet As Excel.Worksheet
            worksheet = workbook.Sheets("SYSTEM-ID")
            worksheet.Range("D6").Value = User
            worksheet.Range("D7").Value = TestMethode
            worksheet.Range("D8").Value = SerialNumber
            worksheet.Range("D9").Value = DateTime.Now.ToString
            worksheet.Range("D10").Value = TestStatus
            worksheet.Range("F6").Value = CycleTime
            worksheet.Range("F7").Value = TesterID
            worksheet.Range("G10").Value = Errormsg
            worksheet.Range("G13").Value = Status
            workbook.ExportAsFixedFormat(Excel.XlFixedFormatType.xlTypePDF, CONFIG.TextBox21.Text.ToString & SerialNumber & "_" & TestMethode & "_" & TestStatus & Format(Now(), "hhmmssmmddyy") & ".pdf", vbNormal)
            workbook.Close(True)
            System.Runtime.InteropServices.Marshal.ReleaseComObject(workbook)
            workbook = Nothing
            xlapp.Quit()
            System.Runtime.InteropServices.Marshal.ReleaseComObject(xlapp)
            xlapp = Nothing
            Process.GetProcessesByName("EXCEL")(0).Kill()
            MAIN.ListBox1.Items.Add("Test Report Generated success : " & CONFIG.TextBox21.Text.ToString & SerialNumber & "_" & TestMethode & "_" & Status & Format(Now(), "hhmmssmmddyy"))
            file1.Close()
        Catch ex As Exception
            Dim iret1 As Object = MsgBox("ERROR WHILE GENERATING TEST REPORT: " & ex.Message, vbRetryCancel, "SystemID.GenerateReport error")
            MAIN.ListBox1.Items.Add("Error Generating Test Report : " & ex.Message)
            If iret1 = vbRetry Then
                file1.Close()
                GoTo line1
            ElseIf iret1 = vbCancel Then
                file1.WriteLine("SystemID.GenerateReport FAILED DUE TO : " & ex.Message & " " & "@" & DateTime.Now.ToString("yyyy/MM/dd/HH:mm:ss"))
                file1.Close()
                Exit Function
            End If
        End Try
    End Function
    Function ClearReportBeforeLogGeberate()
        Dim line1 As Label
line1:
        Dim path1 As String = CONFIG.TextBox4.Text
        file1 = My.Computer.FileSystem.OpenTextFileWriter(path1, True)
        Try
            Dim xlapp = New Excel.Application
            xlapp.DisplayAlerts = False
            Dim workbook = xlapp.Workbooks.Open(CONFIG.TextBox20.Text.ToString)
            Dim worksheet As Excel.Worksheet
            worksheet = workbook.Sheets("SYSTEM-ID")
            worksheet.Range("D6", "D10").ClearContents()
            worksheet.Range("F6", "F7").ClearContents()
            worksheet.Range("G7").Value = ""
            worksheet.Range("G13").Value = ""
            workbook.Close(True)
            System.Runtime.InteropServices.Marshal.ReleaseComObject(workbook)
            workbook = Nothing
            xlapp.Quit()
            System.Runtime.InteropServices.Marshal.ReleaseComObject(xlapp)
            xlapp = Nothing
            Process.GetProcessesByName("EXCEL")(0).Kill()
            file1.Close()
        Catch ex As Exception
            Dim iret1 As Object = MsgBox("Error Executing BatterID.ClearReportBeforeLogGeberate: " & ex.Message, vbRetryCancel, "Error Executing BatterID.ClearReportBeforeLogGeberate")
            MAIN.ListBox1.Items.Add($"{DateTime.Now.ToString} Error Generating Test Report : " & ex.Message)
            MAIN.ListBox1.SelectedIndex = MAIN.ListBox1.Items.Count - 1
            If iret1 = vbRetry Then
                file1.Close()
                GoTo line1
            ElseIf iret1 = vbCancel Then
                MAIN.ListBox1.Items.Add($"{DateTime.Now.ToString} Error Generating Test Report : " & ex.Message)
                MAIN.ListBox1.SelectedIndex = MAIN.ListBox1.Items.Count - 1
                file1.WriteLine("Error Executing BatterID.ClearReportBeforeLogGeberate : " & ex.Message & " " & "@" & DateTime.Now.ToString("yyyy/MM/dd/HH:mm:ss"))
                file1.Close()
                Exit Function
            End If
        End Try
    End Function
    Function ClearSequence()
        MAIN.Button3.Enabled = True
        MAIN.Button1.Enabled = False
        MAIN.Button9.Visible = False
        MAIN.Button100.Visible = False
        MAIN.Button101.Visible = False
        If MAIN.RadioButton1.Checked = True Then
            MAIN.TextBox4.Clear()
            MAIN.TimerFunction("RESET")
            MAIN.RadioButton1.Checked = False
            MAIN.RadioButton2.Checked = False
            MAIN.Button43.Text = "IDLE"
            MAIN.TextBox13.Clear()
            MAIN.TextBox6.Clear()
            MAIN.TextBox153.Clear()
            MAIN.Button8.Text = "IDLE"
            MAIN.PictureBox1.Visible = False
            MAIN.TextBox159.Visible = False
            MAIN.TreeView1.SelectedNode.Nodes(0).Nodes(0).BackColor = Color.WhiteSmoke
            MAIN.TreeView1.SelectedNode.Nodes(0).Nodes(1).BackColor = Color.WhiteSmoke
            MAIN.TreeView1.SelectedNode.Nodes(0).Nodes(2).BackColor = Color.WhiteSmoke
            MAIN.TreeView1.SelectedNode.Nodes(0).Nodes(3).BackColor = Color.WhiteSmoke
            MAIN.TreeView1.SelectedNode.Nodes(0).Nodes(4).BackColor = Color.WhiteSmoke
            MAIN.Button1.Enabled = False
            MAIN.Button3.Enabled = True
            FUNNCTIONLIB.ProgressReset(True)
            myProcesses = Process.GetProcessesByName(CONFIG.TextBox22.Text.ToString)
            If myProcesses.Length > 0 Then
                For Each myProcess In myProcesses
                    If myProcess IsNot Nothing Then
                        myProcess.Kill()
                    End If
                Next
            End If
        Else
            MAIN.TextBox4.Clear()
            MAIN.TimerFunction("RESET")
            MAIN.RadioButton1.Checked = False
            MAIN.RadioButton2.Checked = False
            MAIN.Button43.Text = "IDLE"
            MAIN.TextBox13.Clear()
            MAIN.TextBox6.Clear()
            MAIN.TextBox153.Clear()
            MAIN.Button8.Text = "IDLE"
            MAIN.PictureBox1.Visible = False
            MAIN.TextBox159.Visible = False
            MAIN.Button1.Enabled = False
            MAIN.Button3.Enabled = True
            FUNNCTIONLIB.ProgressReset(True)
            myProcesses = Process.GetProcessesByName(CONFIG.TextBox22.Text.ToString)
            If myProcesses.Length > 0 Then
                For Each myProcess In myProcesses
                    If myProcess IsNot Nothing Then
                        myProcess.Kill()
                    End If
                Next
            End If
        End If

    End Function
    Function CalculateCount(TestResult As String)
        Dim TotalCount As Integer = MAIN.TextBox174.Text
        Dim PassCount As Integer = MAIN.TextBox175.Text
        Dim FailCount As Integer = MAIN.TextBox176.Text
        If TestResult = "PASS" Then
            TotalCount = TotalCount + 1
            PassCount = PassCount + 1
            FailCount = FailCount
        ElseIf TestResult = "FAIL" Then
            TotalCount = TotalCount + 1
            PassCount = PassCount
            FailCount = FailCount + 1
        End If
        MAIN.TextBox174.Text = TotalCount
        MAIN.TextBox175.Text = PassCount
        MAIN.TextBox176.Text = FailCount
        Dim Yield As Double = (TotalCount - FailCount) / TotalCount
        Dim formattedValue As String = (Yield * 100).ToString("0.00") & "%"
        MAIN.TextBox177.Text = formattedValue
    End Function
End Module
