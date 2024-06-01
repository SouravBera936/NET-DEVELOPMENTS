Imports Microsoft.Office.Interop
Imports System.Data.OleDb
Imports System.Runtime.InteropServices
Imports MadMilkman.Ini

Module SMC
    Dim currentDateTime As DateTime = DateTime.Now
    Dim ini As New IniFile
    Dim batchFilePath As String = CONFIG.TextBox34.Text
    Dim CapturedImg As Bitmap
    Dim Reference As Bitmap
    Dim cn As OleDb.OleDbConnection
    Dim oleCommand As OleDbCommand
    Dim file1 As System.IO.StreamWriter
    Dim filepath1 As String = "C:\PLUG_POWER_FCT\PLUG_POWER_FCT\Project Files\Temp\SNAP\1.jpeg"
    Dim SmcBox = {MAIN.Button5, MAIN.TextBox162, MAIN.TextBox167, MAIN.Button6, MAIN.TextBox172, MAIN.TextBox171, MAIN.TextBox170, MAIN.TextBox169, MAIN.Button5, MAIN.TextBox168, MAIN.TextBox166,
        MAIN.TextBox165, MAIN.Button4, MAIN.TextBox161, MAIN.TextBox164, MAIN.TextBox163, MAIN.TextBox30, MAIN.TextBox31, MAIN.TextBox32, MAIN.TextBox33, MAIN.TextBox34,
        MAIN.TextBox26, MAIN.TextBox27, MAIN.TextBox28, MAIN.TextBox29,
       MAIN.TextBox35, MAIN.TextBox36, MAIN.TextBox37, MAIN.TextBox38, MAIN.Button16, MAIN.Button17, MAIN.TextBox152, MAIN.TextBox47, MAIN.TextBox48, MAIN.TextBox49,
       MAIN.TextBox50, MAIN.TextBox51, MAIN.TextBox41, MAIN.TextBox43, MAIN.TextBox45, MAIN.TextBox46, MAIN.TextBox39, MAIN.TextBox40, MAIN.TextBox42, MAIN.TextBox44,
       MAIN.TextBox52, MAIN.TextBox53, MAIN.TextBox54, MAIN.TextBox55, MAIN.TextBox56, MAIN.TextBox57, MAIN.TextBox58, MAIN.TextBox59, MAIN.TextBox60, MAIN.TextBox61,
       MAIN.TextBox62, MAIN.TextBox63, MAIN.TextBox64, MAIN.TextBox65, MAIN.TextBox66, MAIN.TextBox67, MAIN.TextBox71, MAIN.TextBox70, MAIN.TextBox69, MAIN.TextBox68,
       MAIN.TextBox72, MAIN.TextBox73, MAIN.TextBox74, MAIN.TextBox75, MAIN.Button19, MAIN.Button18, MAIN.Button20, MAIN.Button21, MAIN.Button22, MAIN.Button23, MAIN.Button24,
       MAIN.Button25, MAIN.TextBox76, MAIN.TextBox77, MAIN.TextBox78, MAIN.TextBox79, MAIN.TextBox80, MAIN.TextBox81, MAIN.TextBox82, MAIN.TextBox83, MAIN.TextBox84,
       MAIN.TextBox85, MAIN.TextBox86, MAIN.TextBox87, MAIN.TextBox88, MAIN.TextBox89, MAIN.TextBox90, MAIN.TextBox91, MAIN.TextBox92, MAIN.TextBox93, MAIN.TextBox94, MAIN.TextBox95,
       MAIN.TextBox96, MAIN.TextBox97, MAIN.TextBox98, MAIN.TextBox99, MAIN.TextBox100, MAIN.TextBox101, MAIN.TextBox102, MAIN.TextBox103, MAIN.TextBox104, MAIN.TextBox105,
       MAIN.TextBox106, MAIN.TextBox107, MAIN.TextBox108, MAIN.TextBox109, MAIN.TextBox110, MAIN.TextBox111, MAIN.TextBox112, MAIN.TextBox113, MAIN.TextBox114, MAIN.TextBox115,
       MAIN.TextBox116, MAIN.TextBox117, MAIN.TextBox118, MAIN.TextBox119, MAIN.TextBox120, MAIN.TextBox121, MAIN.TextBox122, MAIN.TextBox123, MAIN.TextBox124, MAIN.Button26,
       MAIN.Button27, MAIN.Button28, MAIN.Button29, MAIN.Button30, MAIN.Button31, MAIN.Button32, MAIN.Button33, MAIN.Button34, MAIN.Button35, MAIN.Button36, MAIN.TextBox129,
       MAIN.TextBox130, MAIN.TextBox131, MAIN.TextBox132, MAIN.TextBox133, MAIN.TextBox125, MAIN.TextBox126, MAIN.TextBox127, MAIN.TextBox128, MAIN.Button37, MAIN.TextBox138, MAIN.TextBox139,
       MAIN.TextBox140, MAIN.TextBox141, MAIN.TextBox142, MAIN.TextBox134, MAIN.TextBox135, MAIN.TextBox136, MAIN.TextBox137, MAIN.Button38, MAIN.TextBox147, MAIN.TextBox148, MAIN.TextBox149,
       MAIN.TextBox150, MAIN.TextBox151, MAIN.TextBox143, MAIN.TextBox144, MAIN.TextBox145, MAIN.TextBox146, MAIN.Button39, MAIN.Button26}
    Dim SmcTextBox = {MAIN.TextBox167, MAIN.TextBox27, MAIN.TextBox37, MAIN.TextBox43, MAIN.TextBox40, MAIN.TextBox54, MAIN.TextBox58, MAIN.TextBox62, MAIN.TextBox66, MAIN.TextBox69, MAIN.TextBox74, MAIN.TextBox86,
        MAIN.TextBox82, MAIN.TextBox90, MAIN.TextBox95, MAIN.TextBox99, MAIN.TextBox103, MAIN.TextBox107, MAIN.TextBox111, MAIN.TextBox115, MAIN.TextBox119, MAIN.TextBox123, MAIN.TextBox126,
        MAIN.TextBox135, MAIN.TextBox144, MAIN.TextBox162, MAIN.TextBox171}
    Dim SmcButton = {MAIN.Button6, MAIN.Button5, MAIN.Button4, MAIN.Button16, MAIN.Button17, MAIN.Button19, MAIN.Button18, MAIN.Button20, MAIN.Button21, MAIN.Button22, MAIN.Button23, MAIN.Button24,
        MAIN.Button25, MAIN.Button26, MAIN.Button28, MAIN.Button29, MAIN.Button30, MAIN.Button31, MAIN.Button32, MAIN.Button33, MAIN.Button34,
        MAIN.Button35, MAIN.Button36, MAIN.Button39, MAIN.Button37, MAIN.Button27, MAIN.Button38}
    Dim SmcFUnctionButton = {MAIN.Button40, MAIN.Button7, MAIN.Button13, MAIN.Button14, MAIN.Button44, MAIN.Button45, MAIN.Button46, MAIN.Button47, MAIN.Button48,
        MAIN.Button49, MAIN.Button50, MAIN.Button51, MAIN.Button52, MAIN.Button53, MAIN.Button54, MAIN.Button55, MAIN.Button56, MAIN.Button57, MAIN.Button58,
        MAIN.Button59, MAIN.Button60, MAIN.Button61, MAIN.Button62, MAIN.Button63, MAIN.Button41, MAIN.Button42, MAIN.Button10, MAIN.Button11,
        MAIN.Button64, MAIN.Button65, MAIN.Button66, MAIN.Button67, MAIN.Button68, MAIN.Button69, MAIN.Button70, MAIN.Button71, MAIN.Button72, MAIN.Button73,
        MAIN.Button74, MAIN.Button75, MAIN.Button76, MAIN.Button77, MAIN.Button78, MAIN.Button79, MAIN.Button80, MAIN.Button81, MAIN.Button82, MAIN.Button83,
        MAIN.Button84, MAIN.Button85, MAIN.Button86, MAIN.Button87, MAIN.Button88, MAIN.Button89, MAIN.Button90, MAIN.Button91, MAIN.Button92, MAIN.Button93,
        MAIN.Button94, MAIN.Button95, MAIN.Button96, MAIN.Button97}
    Function TestInitialize() 'INITIALIZE THE TEST ENVIROMENT
        MAIN.Button3.Enabled = False
        MAIN.Button1.Enabled = True
        FUNNCTIONLIB.CalculateProgress(34, 1)
        MAIN.ListBox1.Items.Add($"{DateTime.Now.ToString} Software Initialization Started")
        MAIN.ListBox1.SelectedIndex = MAIN.ListBox1.Items.Count - 1
        Dim line1 As Label
line1:
        Try
            IMAGE_STR.PictureBox1.Image = Nothing
            IMAGE_STR.PictureBox2.Image = Nothing
            IMAGE_STR.PictureBox3.Image = Nothing
            IMAGE_STR.PictureBox4.Image = Nothing
            IMAGE_STR.PictureBox5.Image = Nothing
            IMAGE_STR.PictureBox6.Image = Nothing
            IMAGE_STR.PictureBox7.Image = Nothing
            IMAGE_STR.PictureBox8.Image = Nothing
            IMAGE_STR.PictureBox9.Image = Nothing
            IMAGE_STR.PictureBox10.Image = Nothing
            IMAGE_STR.PictureBox11.Image = Nothing
            IMAGE_STR.PictureBox12.Image = Nothing
            IMAGE_STR.PictureBox13.Image = Nothing
            IMAGE_STR.PictureBox14.Image = Nothing
            IMAGE_STR.PictureBox15.Image = Nothing
            IMAGE_STR.PictureBox16.Image = Nothing
            IMAGE_STR.PictureBox17.Image = Nothing
            IMAGE_STR.PictureBox18.Image = Nothing
            MAIN.ListBox2.Items.Clear()
            For Each tb3 In SmcFUnctionButton
                tb3.visible = False
            Next
            For Each tb In SmcBox
                tb.enabled = False
            Next
            For Each tb1 In SmcTextBox
                tb1.enabled = True
                tb1.text = "IDLE"
            Next
            For Each tb2 In SmcButton
                tb2.text = "IDLE"
            Next
            MAIN.TabControl1.TabPages.Remove(MAIN.TabPage1)
            MAIN.TabControl1.TabPages.Remove(MAIN.TabPage2)
            MAIN.TabControl1.TabPages.Remove(MAIN.TabPage3)
            MAIN.TabControl1.TabPages.Remove(MAIN.TabPage4)
            MAIN.TabControl1.TabPages.Remove(MAIN.TabPage5)
            MAIN.TabControl1.TabPages.Add(MAIN.TabPage3)
            MAIN.TabControl2.SelectedTab = MAIN.TabPage6
            MAIN.TabControl3.SelectedTab = MAIN.TabPage14
            MAIN.GroupBox2.Visible = True
            MAIN.GroupBox3.Visible = True
            For Each tn In MAIN.TreeView1.Nodes
                If tn.Text = STARTUP.ComboBox1.SelectedItem Then
                    MAIN.TreeView1.SelectedNode = tn
                    MAIN.TreeView1.SelectedNode.Nodes(0).Expand()
                    Exit For
                End If
                FindNode(tn)
            Next
            MAIN.Button43.Text = "RUNNING"
            MAIN.TextBox173.Clear()
            MAIN.TextBox173.Visible = False
            MAIN.Button40.Visible = False
            MAIN.PictureBox3.Visible = False
            If POWER_SUPPLY.Initialize() = True Then
                If DMM.Initialize() = True Then
                    MAIN.ListBox1.Items.Add($"{DateTime.Now.ToString} Software Initialization Success")
                    MAIN.ListBox1.SelectedIndex = MAIN.ListBox1.Items.Count - 1
                    MAIN.TreeView1.SelectedNode.Nodes(0).ExpandAll()
                    MAIN.TreeView1.SelectedNode.Nodes(0).Nodes(0).BackColor = Color.Yellow
                    SMC_DUT_INFIXTURE.PictureBox1.Image = Image.FromFile(CONFIG.TextBox37.Text)
                    SMC_DUT_INFIXTURE.ShowDialog()
                Else
                    MAIN.ListBox1.Items.Add($"{DateTime.Now.ToString} Fail to Initialize DMM")
                    MAIN.ListBox1.SelectedIndex = MAIN.ListBox1.Items.Count - 1
                    MAIN.Button43.Text = "FAIL"
                    UUT_RESULT.RadioButton2.Checked = True
                    UUT_RESULT.TextBox2.Text = "SMC_INIT"
                    FUNNCTIONLIB.ProgressReset(True)
                    MAIN.TreeView1.SelectedNode.Nodes(0).ExpandAll()
                    MAIN.TreeView1.SelectedNode.Nodes(0).Nodes(0).BackColor = Color.Red
                    UUT_RESULT.ShowDialog()
                End If
            Else
                MAIN.ListBox1.Items.Add($"{DateTime.Now.ToString} Fail to Initialize POWER SUPPLY")
                MAIN.ListBox1.SelectedIndex = MAIN.ListBox1.Items.Count - 1
                MAIN.Button43.Text = "FAIL"
                UUT_RESULT.RadioButton2.Checked = True
                UUT_RESULT.TextBox2.Text = "SMC_INIT"
                FUNNCTIONLIB.ProgressReset(True)
                MAIN.TreeView1.SelectedNode.Nodes(0).ExpandAll()
                MAIN.TreeView1.SelectedNode.Nodes(0).Nodes(0).BackColor = Color.Red
                UUT_RESULT.ShowDialog()
            End If
        Catch ex As Exception
            Dim iret As Object = MsgBox("Error Executing SMC.TestInitialize : " & ex.Message, vbCritical + vbRetryCancel, "Error executing SMC.TestInitialize")
            If iret = vbRetry Then
                GoTo line1
            ElseIf iret = vbCancel Then
                MAIN.TimerFunction("STOP")
                MAIN.Button43.Text = "ERROR"
                MAIN.ListBox1.Items.Add($"{DateTime.Now.ToString} Error Executing SMC.TestInitialize :" & ex.Message)
                MAIN.ListBox1.SelectedIndex = MAIN.ListBox1.Items.Count - 1
                UUT_RESULT.RadioButton3.Checked = True
                UUT_RESULT.TextBox2.Text = "SMC_INIT"
                MAIN.TreeView1.SelectedNode.Nodes(0).ExpandAll()
                MAIN.TreeView1.SelectedNode.Nodes(0).Nodes(0).BackColor = Color.Red
                UUT_RESULT.ShowDialog()
            End If
        End Try
    End Function
    Function SMCProgrammingNode0() 'DUT CONNECT
        FUNNCTIONLIB.CalculateProgress(34, 2)
        MAIN.ListBox1.Items.Add($"{DateTime.Now.ToString} Waiting For Response From User To Connect The Programmer Cable ")
        MAIN.ListBox1.SelectedIndex = MAIN.ListBox1.Items.Count - 1
        If MAIN.RadioButton1.Checked = True Then
            MAIN.PictureBox3.Image = Image.FromFile(CONFIG.TextBox30.Text.ToString)
            MAIN.PictureBox3.Visible = True
            MAIN.Button40.Visible = True
            MAIN.TextBox173.Text = "CONNECT THE PROGRAMMER CABLE TO THE BOARD AND CLICK ON CONTINUE"
            MAIN.TextBox173.Enabled = False
            MAIN.TextBox173.Visible = True
            MAIN.TreeView1.SelectedNode.Nodes(0).Nodes(0).BackColor = Color.Green
            MAIN.TreeView1.SelectedNode.Nodes(0).Nodes(1).BackColor = Color.Yellow
        Else

        End If

    End Function
    Function SMCProgrammingNode1() 'SET POWER
        FUNNCTIONLIB.CalculateProgress(34, 3)
        MAIN.ListBox1.Items.Add($"{DateTime.Now.ToString} Started Setting Limit To The Power Supply ")
        MAIN.ListBox1.SelectedIndex = MAIN.ListBox1.Items.Count - 1
        Dim line1 As Label
line1:
        Try
            If MAIN.RadioButton1.Checked = True Then
                MAIN.PictureBox3.Visible = False
                MAIN.Button40.Visible = False
                MAIN.TextBox173.Visible = False
                If POWER_SUPPLY.SetLimit() = True Then
                    MAIN.ListBox1.Items.Add($"{DateTime.Now.ToString} POWER_SUPPLY.SetLimit Returns : true ")
                    MAIN.ListBox1.SelectedIndex = MAIN.ListBox1.Items.Count - 1
                    MAIN.TreeView1.SelectedNode.Nodes(0).Nodes(1).BackColor = Color.Green
                    SMCProgrammingNode2()
                    MAIN.TreeView1.SelectedNode.Nodes(0).Nodes(2).BackColor = Color.Yellow
                Else
                    MAIN.ListBox1.Items.Add($"{DateTime.Now.ToString} POWER_SUPPLY.SetLimit Returns : false ")
                    MAIN.ListBox1.SelectedIndex = MAIN.ListBox1.Items.Count - 1
                    MAIN.TreeView1.SelectedNode.Nodes(0).Nodes(1).BackColor = Color.Red
                    MAIN.TimerFunction("STOP")
                    MAIN.Button43.Text = "FAIL"
                    UUT_RESULT.RadioButton2.Checked = True
                    UUT_RESULT.TextBox2.Text = "SET_POWER"
                    UUT_RESULT.ShowDialog()
                End If
            Else

            End If
        Catch ex As Exception
            Dim iret As Object = MsgBox("Error Executing SMC.SMCProgrammingNode1 : " & ex.Message, vbCritical + vbRetryCancel, "Error executing SMC.SMCProgrammingNode1")
            If iret = vbRetry Then
                GoTo line1
            ElseIf iret = vbCancel Then
                MAIN.TimerFunction("STOP")
                MAIN.Button43.Text = "ERROR"
                MAIN.ListBox1.Items.Add($"{DateTime.Now.ToString} Error Executing SMC.SMCProgrammingNode1 :" & ex.Message)
                MAIN.ListBox1.SelectedIndex = MAIN.ListBox1.Items.Count - 1
                UUT_RESULT.RadioButton3.Checked = True
                UUT_RESULT.TextBox2.Text = "SET_POWER"
                MAIN.TreeView1.SelectedNode.Nodes(0).Nodes(1).BackColor = Color.Red
                UUT_RESULT.ShowDialog()
            End If
        End Try
    End Function
    Function SMCProgrammingNode2() 'POWER ON
        FUNNCTIONLIB.CalculateProgress(34, 4)
        MAIN.ListBox1.Items.Add($"{DateTime.Now.ToString} Startrd Power Supply ON stage")
        MAIN.ListBox1.SelectedIndex = MAIN.ListBox1.Items.Count - 1
        Dim line1 As Label
line1:
        Try
            If MAIN.RadioButton1.Checked = True Then
                MAIN.ListBox1.Items.Add($"{DateTime.Now.ToString} POWER_SUPPLY.Ps_ON Returns : true")
                MAIN.ListBox1.SelectedIndex = MAIN.ListBox1.Items.Count - 1
                If POWER_SUPPLY.PS_ON() = True Then
                    MAIN.TreeView1.SelectedNode.Nodes(0).Nodes(2).BackColor = Color.Green
                    MAIN.Button41.Visible = True
                    MAIN.PictureBox3.Image = Image.FromFile(CONFIG.TextBox84.Text.ToString)
                    MAIN.PictureBox3.Visible = True
                    MAIN.TextBox173.Text = "CONNECT THE USB CABLE TO THE PROGRAMMER AND CLICK ON CONTINUE"
                    MAIN.TextBox173.Enabled = False
                    MAIN.TextBox173.Visible = True
                    MAIN.TreeView1.SelectedNode.Nodes(0).Nodes(3).BackColor = Color.Yellow
                Else
                    MAIN.ListBox1.Items.Add($"{DateTime.Now.ToString} POWER_SUPPLY.Ps_ON Returns : false")
                    MAIN.ListBox1.SelectedIndex = MAIN.ListBox1.Items.Count - 1
                    MAIN.TreeView1.SelectedNode.Nodes(0).Nodes(2).BackColor = Color.Red
                    MAIN.TimerFunction("STOP")
                    MAIN.Button43.Text = "FAIL"
                    UUT_RESULT.RadioButton2.Checked = True
                    UUT_RESULT.TextBox2.Text = "POWER_ON"
                    UUT_RESULT.ShowDialog()
                End If
            Else
                POWER_SUPPLY.PS_ON()
            End If
        Catch ex As Exception
            Dim iret As Object = MsgBox("Error Executing SMC.SMCProgrammingNode2 : " & ex.Message, vbCritical + vbRetryCancel, "Error executing SMC.SMCProgrammingNode2")
            If iret = vbRetry Then
                GoTo line1
            ElseIf iret = vbCancel Then
                MAIN.TimerFunction("STOP")
                MAIN.Button43.Text = "ERROR"
                MAIN.ListBox1.Items.Add($"{DateTime.Now.ToString} Error Executing SMC.SMCProgrammingNode1 :" & ex.Message)
                MAIN.ListBox1.SelectedIndex = MAIN.ListBox1.Items.Count - 1
                UUT_RESULT.RadioButton3.Checked = True
                UUT_RESULT.TextBox2.Text = "POWER_ON"
                MAIN.TreeView1.SelectedNode.Nodes(0).Nodes(2).BackColor = Color.Red
                UUT_RESULT.ShowDialog()
            End If
        End Try
    End Function
    Function SMCProgrammimgNode3() 'UNIFLASH PROGRAMMING
        FUNNCTIONLIB.CalculateProgress(34, 5)
        MAIN.ListBox1.Items.Add($"{DateTime.Now.ToString} : Started Uniflash Programming Step")
        MAIN.ListBox1.SelectedIndex = MAIN.ListBox1.Items.Count - 1
        Dim line1 As Label
line1:
        Try
            MAIN.Button16.Text = "RUNNING"
            If MAIN.RadioButton1.Checked = True Then
                If FUNNCTIONLIB.UnifalshProgramme(batchFilePath) = True Then
                    MAIN.TreeView1.SelectedNode.Nodes(0).Nodes(3).BackColor = Color.Green
                    MAIN.Button16.Text = "PASS"
                    MAIN.TextBox27.Text = "TRUE"
                    MAIN.ListBox1.Items.Add($"{DateTime.Now.ToString} FUNNCLIB.UnifalProgramme Returns : True")
                    MAIN.ListBox1.SelectedIndex = MAIN.ListBox1.Items.Count - 1
                    SMCProgrammingNode4()
                Else
                    MAIN.ListBox1.Items.Add($"{DateTime.Now.ToString} FUNNCLIB.UnifalProgramme Returns : false")
                    MAIN.ListBox1.SelectedIndex = MAIN.ListBox1.Items.Count - 1
                    MAIN.TreeView1.SelectedNode.Nodes(0).Nodes(3).BackColor = Color.Red
                    MAIN.TimerFunction("STOP")
                    MAIN.Button43.Text = "FAIL"
                    MAIN.Button16.Text = "FAIL"
                    MAIN.TextBox27.Text = "FALSE"
                    UUT_RESULT.RadioButton2.Checked = True
                    UUT_RESULT.TextBox2.Text = "UNIFLASH_PROG"
                    UUT_RESULT.ShowDialog()
                End If
            End If
        Catch ex As Exception
            Dim iret As Object = MsgBox("Error Executing SMC.SMCProgrammimgNode3 : " & ex.Message, vbCritical + vbRetryCancel, "Error executing SMC.SMCProgrammimgNode3")
            If iret = vbRetry Then
                GoTo line1
            ElseIf iret = vbCancel Then
                MAIN.TimerFunction("STOP")
                MAIN.Button16.Text = "ERROR"
                MAIN.TextBox27.Text = "ERROR"
                MAIN.Button43.Text = "ERROR"
                MAIN.ListBox1.Items.Add($"{DateTime.Now.ToString} Error Executing SMC.SMCProgrammimgNode3 :" & ex.Message)
                MAIN.ListBox1.SelectedIndex = MAIN.ListBox1.Items.Count - 1
                UUT_RESULT.RadioButton3.Checked = True
                UUT_RESULT.TextBox2.Text = "UNIFLASH_PROG"
                MAIN.TreeView1.SelectedNode.Nodes(0).Nodes(3).BackColor = Color.Red
                UUT_RESULT.ShowDialog()
            End If
        End Try

    End Function
    Function SMCProgrammingNode4() 'POWER RECYCLE
        FUNNCTIONLIB.CalculateProgress(34, 6)
        MAIN.ListBox1.Items.Add($"{DateTime.Now.ToString} : Started Power Recycling Step")
        MAIN.ListBox1.SelectedIndex = MAIN.ListBox1.Items.Count - 1
        Dim line1 As Label
        MAIN.Button17.Text = "RUNNING"
line1:
        Try
            If MAIN.RadioButton1.Checked = True Then
                MAIN.TreeView1.SelectedNode.Nodes(0).Nodes(4).BackColor = Color.Yellow
                If POWER_SUPPLY.PS_RECYCLE(5000) = True Then
                    MAIN.ListBox1.Items.Add($"{DateTime.Now.ToString} : POWER_SUPPLY.PS_RECYCLE Returns : true")
                    MAIN.ListBox1.SelectedIndex = MAIN.ListBox1.Items.Count - 1
                    MAIN.TreeView1.SelectedNode.Nodes(0).Nodes(4).BackColor = Color.Green
                    MAIN.Button42.Visible = True
                    MAIN.PictureBox3.Image = Image.FromFile(CONFIG.TextBox85.Text.ToString)
                    MAIN.PictureBox3.Visible = True
                    MAIN.TextBox173.Text = "DISCONNECT THE UNIFLASH PROGRAMMER AND CONNECT THE USB TO THE BOARD AND CLICK ON CONTINUE"
                    MAIN.TextBox173.Enabled = False
                    MAIN.TextBox173.Visible = True
                    MAIN.TreeView1.SelectedNode.Nodes(0).Nodes(5).BackColor = Color.Yellow
                Else
                    MAIN.ListBox1.Items.Add($"{DateTime.Now.ToString} : POWER_SUPPLY.PS_RECYCLE Returns : false")
                    MAIN.ListBox1.SelectedIndex = MAIN.ListBox1.Items.Count - 1
                    MAIN.TreeView1.SelectedNode.Nodes(0).Nodes(4).BackColor = Color.Red
                    MAIN.TimerFunction("STOP")
                    MAIN.TextBox37.Text = "FALSE"
                    MAIN.Button17.Text = "FAIL"
                    MAIN.Button43.Text = "FAIL"
                    UUT_RESULT.RadioButton2.Checked = True
                    UUT_RESULT.TextBox2.Text = "POWER_REC"
                    UUT_RESULT.ShowDialog()
                End If
            Else

            End If
        Catch ex As Exception
            Dim iret As Object = MsgBox("Error Executing SMC.SMCProgrammingNode4 : " & ex.Message, vbCritical + vbRetryCancel, "Error executing SMC.SMCProgrammingNode4")
            If iret = vbRetry Then
                GoTo line1
            ElseIf iret = vbCancel Then
                MAIN.TimerFunction("STOP")
                MAIN.Button43.Text = "ERROR"
                MAIN.Button17.Text = "ERROR"
                MAIN.Button43.Text = "ERROR"
                MAIN.ListBox1.Items.Add($"{DateTime.Now.ToString} Error Executing SMC.SMCProgrammingNode4 :" & ex.Message)
                MAIN.ListBox1.SelectedIndex = MAIN.ListBox1.Items.Count - 1
                UUT_RESULT.RadioButton3.Checked = True
                UUT_RESULT.TextBox2.Text = "POWER_REC"
                MAIN.TreeView1.SelectedNode.Nodes(0).Nodes(4).BackColor = Color.Red
                UUT_RESULT.ShowDialog()
            End If
        End Try
    End Function
    Function SMCProgrammingNode5() 'FTDI PROGRAMMING
        FUNNCTIONLIB.CalculateProgress(34, 7)
        MAIN.ListBox1.Items.Add($"{DateTime.Now.ToString} : Started FTDI PROGRAMMING Step")
        MAIN.ListBox1.SelectedIndex = MAIN.ListBox1.Items.Count - 1
        Dim line1 As Label
line1:
        Try
            If MAIN.RadioButton1.Checked = True Then
                If FUNNCTIONLIB.FTDIDetectDevices() = True Then
                    MAIN.ListBox1.Items.Add($"{DateTime.Now.ToString} : FUNNCLIB.FTDIDetectDevices Returns : true")
                    MAIN.ListBox1.SelectedIndex = MAIN.ListBox1.Items.Count - 1
                    MAIN.ListBox1.Items.Add($"{DateTime.Now.ToString} : Waiting For User To Select The Proper FTDI Device To Programme")
                    MAIN.ListBox1.SelectedIndex = MAIN.ListBox1.Items.Count - 1
                    Call FTDI_SELECTOR.Show()
                Else
                    MAIN.ListBox1.Items.Add($"{DateTime.Now.ToString} : FUNNCLIB.FTDIDetectDevices Returns : false")
                    MAIN.ListBox1.SelectedIndex = MAIN.ListBox1.Items.Count - 1
                    MAIN.TreeView1.SelectedNode.Nodes(0).Nodes(5).BackColor = Color.Red
                    MAIN.TextBox37.Text = "FALSE"
                    MAIN.Button17.Text = "FAIL"
                    MAIN.Button43.Text = "FAIL"
                    MAIN.TimerFunction("STOP")
                    UUT_RESULT.RadioButton2.Checked = True
                    UUT_RESULT.TextBox2.Text = "FTDI_PROG"
                    UUT_RESULT.ShowDialog()
                End If
            End If
        Catch ex As Exception
            Dim iret As Object = MsgBox("Error Executing SMC.SMCProgrammingNode5 : " & ex.Message, vbCritical + vbRetryCancel, "Error executing SMC.SMCProgrammingNode5")
            If iret = vbRetry Then
                GoTo line1
            ElseIf iret = vbCancel Then
                MAIN.TimerFunction("STOP")
                MAIN.Button43.Text = "ERROR"
                MAIN.Button17.Text = "ERROR"
                MAIN.Button43.Text = "ERROR"
                MAIN.ListBox1.Items.Add($"{DateTime.Now.ToString} Error Executing SMC.SMCProgrammingNode5 :" & ex.Message)
                MAIN.ListBox1.SelectedIndex = MAIN.ListBox1.Items.Count - 1
                UUT_RESULT.RadioButton3.Checked = True
                UUT_RESULT.TextBox2.Text = "FTDI_PROG"
                MAIN.TreeView1.SelectedNode.Nodes(0).Nodes(5).BackColor = Color.Red
                UUT_RESULT.ShowDialog()
            End If
        End Try
    End Function
    Function SMCProgrammingNode6() 'POWER RECYCLE
        FUNNCTIONLIB.CalculateProgress(34, 8)
        MAIN.ListBox1.Items.Add($"{DateTime.Now.ToString} : Started Power Recycling Step")
        MAIN.ListBox1.SelectedIndex = MAIN.ListBox1.Items.Count - 1
        Dim line1 As Label
line1:
        If MAIN.RadioButton1.Checked = True Then
            MAIN.TreeView1.SelectedNode.Nodes(0).Nodes(6).BackColor = Color.Yellow
            If POWER_SUPPLY.PS_RECYCLE(10000) = True Then
                MAIN.TreeView1.SelectedNode.Nodes(0).Nodes(6).BackColor = Color.Green
                MAIN.ListBox1.Items.Add($"{DateTime.Now.ToString} : Power Supply Reclying Success")
                MAIN.ListBox1.SelectedIndex = MAIN.ListBox1.Items.Count - 1
                SMCPowerTestNode0()
            Else
                MAIN.ListBox1.Items.Add($"{DateTime.Now.ToString} : Power Supply Reclying Failed")
                MAIN.ListBox1.SelectedIndex = MAIN.ListBox1.Items.Count - 1
                MAIN.TreeView1.SelectedNode.Nodes(0).Nodes(6).BackColor = Color.Red
                Exit Function
            End If
        Else
            POWER_SUPPLY.PS_RECYCLE(5000)
        End If
    End Function
    Function SMCPowerTestNode0() 'INITIAL POWER
        FUNNCTIONLIB.CalculateProgress(34, 9)
        Dim line1 As Label
        MAIN.ListBox1.Items.Add($"{DateTime.Now.ToString} Started Running INITIAL POWER Step")
        MAIN.ListBox1.SelectedIndex = MAIN.ListBox1.Items.Count - 1
        If MAIN.RadioButton1.Checked = True Then
            MAIN.TreeView1.SelectedNode.Nodes(0).Collapse()
            MAIN.TreeView1.SelectedNode.Nodes(1).ExpandAll()
            MAIN.TreeView1.SelectedNode.Nodes(1).Nodes(0).BackColor = Color.Yellow
            MAIN.TabControl2.SelectedTab = MAIN.TabPage7
            MAIN.Button19.Text = "RUNNING"
line1:
            MAIN.ListBox2.Items.Clear()
            Try
                Dim myProcesses() As Process
                Dim myProcess As Process
                myProcesses = Process.GetProcessesByName(CONFIG.TextBox48.Text.ToString)
                If myProcesses.Length > 0 Then
                    For Each myProcess In myProcesses
                        If myProcess IsNot Nothing Then
                            myProcess.Kill()
                        End If
                    Next
                End If
                Dim p As Process = Process.Start(CONFIG.TextBox38.Text.ToString)
                System.Threading.Thread.Sleep(7000)
                If myProcesses.Length > 0 Then
                    MAIN.ListBox1.Items.Add($"{DateTime.Now.ToString} SMC Application {CONFIG.TextBox38.Text} has been satrted after {MAIN.TextBox5.Text}")
                    MAIN.ListBox1.SelectedIndex = MAIN.ListBox1.Items.Count - 1
                Else
                    Dim iret As Object = MsgBox("Error Starting SMC Application", vbCritical + vbRetryCancel, "Error Executing SMC.SMCPowerTestNode0")
                    If iret = vbRetry Then
                        GoTo line1
                    ElseIf iret = vbCancel Then
                        MAIN.ListBox1.Items.Add($"{DateTime.Now.ToString} Error Occured while starting SMC Application")
                        MAIN.ListBox1.SelectedIndex = MAIN.ListBox1.Items.Count - 1
                        MAIN.Button19.Text = "FAIL"
                        MAIN.TreeView1.SelectedNode.Nodes(1).Nodes(0).BackColor = Color.Red
                        MAIN.Button43.Text = "FAIL"
                        MAIN.TextBox43.Text = "FALSE"
                        MAIN.TimerFunction("STOP")
                        UUT_RESULT.RadioButton2.Checked = True
                        UUT_RESULT.TextBox2.Text = "INIT_POWER"
                        UUT_RESULT.ShowDialog()
                    End If
                End If
                If SETUP.CheckBox6.Checked = True Then
                    MAIN.ListBox1.Items.Add($"{DateTime.Now.ToString} AutoImageComparison Enabled")
                    MAIN.ListBox1.SelectedIndex = MAIN.ListBox1.Items.Count - 1
                    MAIN.ListBox1.Items.Add($"{DateTime.Now.ToString} Loading Coordinates from Confriguration File")
                    MAIN.ListBox1.SelectedIndex = MAIN.ListBox1.Items.Count - 1
                    ini.Load("C:\PLUG_POWER_FCT\PLUG_POWER_FCT\Project Files\Confrigurations.ini")
                    MAIN.ListBox2.Items.Add(ini.Sections("SMC_INITIALPOWER_Coordinates").Keys("Coordinate1").Value)
                    MAIN.ListBox1.Items.Add($"{DateTime.Now.ToString} {ini.Sections("SMC_INITIALPOWER_Coordinates").Keys("Coordinate1").Value}")
                    MAIN.ListBox1.SelectedIndex = MAIN.ListBox1.Items.Count - 1
                    MAIN.ListBox2.Items.Add(ini.Sections("SMC_INITIALPOWER_Coordinates").Keys("Coordinate2").Value)
                    MAIN.ListBox1.Items.Add($"{DateTime.Now.ToString} {ini.Sections("SMC_INITIALPOWER_Coordinates").Keys("Coordinate2").Value}")
                    MAIN.ListBox1.SelectedIndex = MAIN.ListBox1.Items.Count - 1
                    MAIN.ListBox2.Items.Add(ini.Sections("SMC_INITIALPOWER_Coordinates").Keys("Coordinate3").Value)
                    MAIN.ListBox1.Items.Add($"{DateTime.Now.ToString} {ini.Sections("SMC_INITIALPOWER_Coordinates").Keys("Coordinate3").Value}")
                    MAIN.ListBox1.SelectedIndex = MAIN.ListBox1.Items.Count - 1
                    MAIN.ListBox2.Items.Add(ini.Sections("SMC_INITIALPOWER_Coordinates").Keys("Coordinate4").Value)
                    MAIN.ListBox1.Items.Add($"{DateTime.Now.ToString} {ini.Sections("SMC_INITIALPOWER_Coordinates").Keys("Coordinate4").Value}")
                    MAIN.ListBox1.SelectedIndex = MAIN.ListBox1.Items.Count - 1
                    MAIN.ListBox2.Items.Add(ini.Sections("SMC_INITIALPOWER_Coordinates").Keys("Coordinate5").Value)
                    MAIN.ListBox1.Items.Add($"{DateTime.Now.ToString} {ini.Sections("SMC_INITIALPOWER_Coordinates").Keys("Coordinate5").Value}")
                    MAIN.ListBox1.SelectedIndex = MAIN.ListBox1.Items.Count - 1
                    MAIN.ListBox2.Items.Add(ini.Sections("SMC_INITIALPOWER_Coordinates").Keys("Coordinate6").Value)
                    MAIN.ListBox1.Items.Add($"{DateTime.Now.ToString} {ini.Sections("SMC_INITIALPOWER_Coordinates").Keys("Coordinate6").Value}")
                    MAIN.ListBox1.SelectedIndex = MAIN.ListBox1.Items.Count - 1
                    MAIN.ListBox2.Items.Add(ini.Sections("SMC_INITIALPOWER_Coordinates").Keys("Coordinate7").Value)
                    MAIN.ListBox1.Items.Add($"{DateTime.Now.ToString} {ini.Sections("SMC_INITIALPOWER_Coordinates").Keys("Coordinate7").Value}")
                    MAIN.ListBox1.SelectedIndex = MAIN.ListBox1.Items.Count - 1
                    MAIN.ListBox1.Items.Add($"{DateTime.Now.ToString} Confriguration Loaded successfully")
                    MAIN.ListBox1.SelectedIndex = MAIN.ListBox1.Items.Count - 1
                    CapturedImg = FUNNCTIONLIB.TakeSnapofActiveApp()
                    Reference = Image.FromFile(CONFIG.TextBox40.Text)
                    MAIN.ListBox1.Items.Add($"{DateTime.Now.ToString} Starting Comparison Function of Coordintaes")
                    MAIN.ListBox1.SelectedIndex = MAIN.ListBox1.Items.Count - 1
                    SMC_IMAGECOMP.PictureBox1.Image = CapturedImg
                    SMC_IMAGECOMP.PictureBox2.Image = Reference
                    SMC_IMAGECOMP.TextBox3.Text = "SMC_INITIAL POWER"
                    SMC_IMAGECOMP.ShowDialog()
                    If MAIN.Button19.Text = "PASS" Then
                        IMAGE_STR.PictureBox1.Image = CapturedImg
                        SMCPowerTestNode1()
                    Else
                        MAIN.TimerFunction("STOP")
                        IMAGE_STR.PictureBox1.Image = CapturedImg
                        UUT_RESULT.RadioButton2.Checked = True
                        UUT_RESULT.TextBox2.Text = "INIT_POWER"
                        UUT_RESULT.ShowDialog()
                    End If
                ElseIf SETUP.CheckBox6.Checked = False Then
                    MAIN.ListBox1.Items.Add($"{DateTime.Now.ToString} AutoImageComparison Disabled")
                    MAIN.ListBox1.SelectedIndex = MAIN.ListBox1.Items.Count - 1
                    MAIN.PictureBox3.Image = Image.FromFile(CONFIG.TextBox39.Text.ToString)
                    MAIN.TextBox173.Text = "CONFIRM THE TEST SCREEN AS SHOWN IN THE IMAGE IF FOUND SAME CLICK ON CONFIRM OTHERWISE CLICK ON NON-CONFIRM"
                    MAIN.Button10.Visible = True
                    MAIN.Button11.Visible = True
                    MAIN.PictureBox3.Visible = True
                    MAIN.TextBox173.Visible = True
                    MAIN.TextBox173.Enabled = False
                End If
            Catch ex As Exception
                Dim iret As Object = MsgBox("Error Executing SMC.SMCPowerTestNode0 : " & ex.Message, vbCritical + vbRetryCancel, "Error executing SMC.SMCPowerTestNode0")
                If iret = vbRetry Then
                    GoTo line1
                ElseIf iret = vbCancel Then
                    MAIN.TimerFunction("STOP")
                    MAIN.Button19.Text = "ERROR"
                    MAIN.TreeView1.SelectedNode.Nodes(1).Nodes(0).BackColor = Color.Red
                    MAIN.Button43.Text = "ERROR"
                    MAIN.TextBox43.Text = "ERROR"
                    MAIN.ListBox1.Items.Add($"{DateTime.Now.ToString} Error Executing SMC.SMCPowerTestNode0 :" & ex.Message)
                    MAIN.ListBox1.SelectedIndex = MAIN.ListBox1.Items.Count - 1
                    IMAGE_STR.PictureBox1.Image = CapturedImg
                    UUT_RESULT.RadioButton3.Checked = True
                    UUT_RESULT.TextBox2.Text = "INIT_POWER"
                    UUT_RESULT.ShowDialog()
                End If
            End Try
        End If
    End Function
    Function SMCPowerTestNode1() 'CAN670_5V
        FUNNCTIONLIB.CalculateProgress(34, 10)
        If MAIN.RadioButton1.Checked = True Then
            MAIN.TreeView1.SelectedNode.Nodes(1).Nodes(1).BackColor = Color.Yellow
            MAIN.Button18.Text = "RUNNING"
            MAIN.PictureBox3.Image = Image.FromFile(CONFIG.TextBox50.Text)
            MAIN.PictureBox3.Visible = True
            MAIN.Button7.Visible = True
            MAIN.TextBox173.Text = "ROTATE THE KNOB TO CAN 670_5V AND CLICK ON CONTINUE"
            MAIN.TextBox173.Enabled = False
            MAIN.TextBox173.Visible = True
        Else
            MAIN.Button18.Text = "RUNNING"
            MAIN.PictureBox3.Image = Image.FromFile(CONFIG.TextBox50.Text)
            MAIN.PictureBox3.Visible = True
            MAIN.Button7.Visible = True
        End If
    End Function
    Function SMCPowerTestNode2() 'ISOLATED_V+12
        FUNNCTIONLIB.CalculateProgress(34, 11)
        If MAIN.RadioButton1.Checked = True Then
            MAIN.TreeView1.SelectedNode.Nodes(1).Nodes(2).BackColor = Color.Yellow
            MAIN.Button20.Text = "RUNNING"
            MAIN.PictureBox3.Image = Image.FromFile(CONFIG.TextBox51.Text)
            MAIN.PictureBox3.Visible = True
            MAIN.Button13.Visible = True
            MAIN.TextBox173.Text = "ROTATE THE KNOB TO ISOLATED V+12 AND CLICK ON CONTINUE"
            MAIN.TextBox173.Enabled = False
            MAIN.TextBox173.Visible = True
        Else
            MAIN.Button20.Text = "RUNNING"
            MAIN.PictureBox3.Image = Image.FromFile(CONFIG.TextBox51.Text)
            MAIN.PictureBox3.Visible = True
            MAIN.Button13.Visible = True
        End If
    End Function
    Function SMCPowerTestNode3() 'SENSOR_V+12
        FUNNCTIONLIB.CalculateProgress(34, 12)
        If MAIN.RadioButton1.Checked = True Then
            MAIN.TreeView1.SelectedNode.Nodes(1).Nodes(3).BackColor = Color.Yellow
            MAIN.Button21.Text = "RUNNING"
            MAIN.PictureBox3.Image = Image.FromFile(CONFIG.TextBox52.Text)
            MAIN.PictureBox3.Visible = True
            MAIN.Button14.Visible = True
            MAIN.TextBox173.Text = "ROTATE THE KNOB TO SENSOR V+12 AND CLICK ON CONTINUE"
            MAIN.TextBox173.Enabled = False
            MAIN.TextBox173.Visible = True
        Else
            MAIN.Button21.Text = "RUNNING"
            MAIN.PictureBox3.Image = Image.FromFile(CONFIG.TextBox52.Text)
            MAIN.PictureBox3.Visible = True
            MAIN.Button14.Visible = True
        End If

    End Function
    Function SMCPowerTestNode4() 'TT620_5V
        FUNNCTIONLIB.CalculateProgress(34, 13)
        If MAIN.RadioButton1.Checked = True Then
            MAIN.TreeView1.SelectedNode.Nodes(1).Nodes(4).BackColor = Color.Yellow
            MAIN.Button22.Text = "RUNNING"
            MAIN.PictureBox3.Image = Image.FromFile(CONFIG.TextBox53.Text)
            MAIN.PictureBox3.Visible = True
            MAIN.Button44.Visible = True
            MAIN.TextBox173.Text = "ROTATE THE KNOB TO TT620 5V AND CLICK ON CONTINUE"
            MAIN.TextBox173.Enabled = False
            MAIN.TextBox173.Visible = True
        Else
            MAIN.Button22.Text = "RUNNING"
            MAIN.PictureBox3.Image = Image.FromFile(CONFIG.TextBox53.Text)
            MAIN.PictureBox3.Visible = True
            MAIN.Button44.Visible = True
        End If
    End Function
    Function SMCPowerTestNode5() 'IT661_5V
        FUNNCTIONLIB.CalculateProgress(34, 15)
        If MAIN.RadioButton1.Checked = True Then
            MAIN.TreeView1.SelectedNode.Nodes(1).Nodes(5).BackColor = Color.Yellow
            MAIN.Button23.Text = "RUNNING"
            MAIN.PictureBox3.Image = Image.FromFile(CONFIG.TextBox54.Text)
            MAIN.PictureBox3.Visible = True
            MAIN.Button45.Visible = True
            MAIN.TextBox173.Text = "ROTATE THE KNOB TO IT661_5V AND CLICK ON CONTINUE"
            MAIN.TextBox173.Enabled = False
            MAIN.TextBox173.Visible = True
        Else
            MAIN.Button23.Text = "RUNNING"
            MAIN.PictureBox3.Image = Image.FromFile(CONFIG.TextBox54.Text)
            MAIN.PictureBox3.Visible = True
            MAIN.Button45.Visible = True
        End If
    End Function
    Function SMCPowerTestNode6() 'A0691
        FUNNCTIONLIB.CalculateProgress(34, 16)
        If MAIN.RadioButton1.Checked = True Then
            MAIN.TreeView1.SelectedNode.Nodes(1).Nodes(6).BackColor = Color.Yellow
            MAIN.Button24.Text = "RUNNING"
            MAIN.PictureBox3.Image = Image.FromFile(CONFIG.TextBox55.Text)
            MAIN.PictureBox3.Visible = True
            MAIN.Button46.Visible = True
            MAIN.TextBox173.Text = "ROTATE THE KNOB TO A0691 AND CLICK ON CONTINUE"
            MAIN.TextBox173.Visible = True
            MAIN.TextBox173.Enabled = False
        Else
            MAIN.Button24.Text = "RUNNING"
            MAIN.PictureBox3.Image = Image.FromFile(CONFIG.TextBox55.Text)
            MAIN.PictureBox3.Visible = True
            MAIN.Button46.Visible = True
        End If
    End Function
    Function SMCPowerTestNode7() 'D0691
        FUNNCTIONLIB.CalculateProgress(34, 17)
        If MAIN.RadioButton1.Checked = True Then
            MAIN.TreeView1.SelectedNode.Nodes(1).Nodes(7).BackColor = Color.Yellow
            MAIN.Button25.Text = "RUNNING"
            MAIN.PictureBox3.Image = Image.FromFile(CONFIG.TextBox56.Text)
            MAIN.TextBox173.Text = "CLICK ON THE D0691 BUTTON ON THE SOFTWARE AND CLICK ON CONTINUE"
            MAIN.PictureBox3.Visible = True
            MAIN.TextBox173.Visible = True
            MAIN.TextBox173.Enabled = False
            FUNNCTIONLIB.RestoreApp(CONFIG.TextBox57.Text)
            MAIN.Button47.Visible = True
        Else
            MAIN.Button25.Text = "RUNNING"
            MAIN.PictureBox3.Image = Image.FromFile(CONFIG.TextBox56.Text)
            MAIN.TextBox173.Text = "Click on the D0691 button"
            MAIN.PictureBox3.Visible = True
            MAIN.TextBox173.Visible = True
            MAIN.TextBox173.Enabled = False
            FUNNCTIONLIB.RestoreApp(CONFIG.TextBox57.Text)
            MAIN.Button47.Visible = True
        End If
    End Function
    Function SMCSwitchTestNode0() 'DO691 ON
        FUNNCTIONLIB.CalculateProgress(34, 18)
        MAIN.ListBox1.Items.Add($"{DateTime.Now.ToString} Started Running  D0691_CMD_ON Step")
        MAIN.ListBox1.SelectedIndex = MAIN.ListBox1.Items.Count - 1
        If MAIN.RadioButton1.Checked = True Then
            MAIN.TreeView1.SelectedNode.Nodes(1).Collapse()
            MAIN.TreeView1.SelectedNode.Nodes(2).ExpandAll()
            MAIN.TreeView1.SelectedNode.Nodes(2).Nodes(0).BackColor = Color.Yellow
            MAIN.TabControl2.SelectedTab = MAIN.TabPage8
            MAIN.Button4.Text = "RUNNING"
            Dim line1 As Label
line1:
            MAIN.ListBox2.Items.Clear()
            Try
                FUNNCTIONLIB.RestoreApp(CONFIG.TextBox57.Text)
                System.Threading.Thread.Sleep(3000)
                If SETUP.CheckBox6.Checked = True Then
                    MAIN.ListBox1.Items.Add($"{DateTime.Now.ToString} AutoImageComparison Enabled")
                    MAIN.ListBox1.SelectedIndex = MAIN.ListBox1.Items.Count - 1
                    CapturedImg = FUNNCTIONLIB.TakeSnapofActiveApp()
                    Reference = Image.FromFile(CONFIG.TextBox58.Text)
                    System.Threading.Thread.Sleep(3000)
                    MAIN.ListBox1.Items.Add($"{DateTime.Now.ToString} Loading Coordinates from Confriguration File")
                    MAIN.ListBox1.SelectedIndex = MAIN.ListBox1.Items.Count - 1
                    ini.Load("C:\PLUG_POWER_FCT\PLUG_POWER_FCT\Project Files\Confrigurations.ini")
                    MAIN.ListBox2.Items.Add(ini.Sections("SMC_D0691_CMD").Keys("Coordinate1").Value)
                    MAIN.ListBox1.Items.Add($"{DateTime.Now.ToString} {ini.Sections("SMC_D0691_CMD").Keys("Coordinate1").Value}")
                    MAIN.ListBox1.SelectedIndex = MAIN.ListBox1.Items.Count - 1
                    MAIN.ListBox2.Items.Add(ini.Sections("SMC_D0691_CMD").Keys("Coordinate2").Value)
                    MAIN.ListBox1.Items.Add($"{DateTime.Now.ToString} {ini.Sections("SMC_D0691_CMD").Keys("Coordinate2").Value}")
                    MAIN.ListBox1.SelectedIndex = MAIN.ListBox1.Items.Count - 1
                    SMC_IMAGECOMP.PictureBox1.Image = CapturedImg
                    SMC_IMAGECOMP.PictureBox2.Image = Reference
                    SMC_IMAGECOMP.TextBox3.Text = "SMC_D0691_CMD"
                    MAIN.ListBox1.Items.Add($"{DateTime.Now.ToString} Confriguration Loaded successfully")
                    MAIN.ListBox1.SelectedIndex = MAIN.ListBox1.Items.Count - 1
                    MAIN.ListBox1.Items.Add($"{DateTime.Now.ToString} Starting Comparison Function of Coordintaes")
                    MAIN.ListBox1.SelectedIndex = MAIN.ListBox1.Items.Count - 1
                    SMC_IMAGECOMP.ShowDialog()
                    If MAIN.Button4.Text = "PASS" Then
                        IMAGE_STR.PictureBox2.Image = CapturedImg
                        MAIN.PictureBox3.Image = Image.FromFile(CONFIG.TextBox41.Text)
                        MAIN.TextBox173.Text = "CLICK ON D0661 CMD SWITCH & CLICK ON CONTINUE"
                        MAIN.PictureBox3.Visible = True
                        MAIN.TextBox173.Visible = True
                        MAIN.Button48.Visible = True
                    Else
                        IMAGE_STR.PictureBox2.Image = CapturedImg
                        MAIN.TimerFunction("STOP")
                        UUT_RESULT.RadioButton2.Checked = True
                        UUT_RESULT.TextBox2.Text = "D0691_ON"
                        UUT_RESULT.ShowDialog()
                    End If
                ElseIf SETUP.CheckBox6.Checked = False Then
                    MAIN.ListBox1.Items.Add($"{DateTime.Now.ToString} AutoImageComparison Disabled")
                    MAIN.ListBox1.SelectedIndex = MAIN.ListBox1.Items.Count - 1
                    MAIN.PictureBox3.Image = Image.FromFile(CONFIG.TextBox56.Text)
                    MAIN.TextBox173.Text = "CONFIRM THE SCREEN AS SHOWN AND CLICK ON CONFIRM OTHERWISE CLICK ON NON-CONFIRM"
                    MAIN.PictureBox3.Visible = True
                    MAIN.TextBox173.Visible = True
                    MAIN.TextBox173.Enabled = False
                    MAIN.Button64.Visible = True
                    MAIN.Button65.Visible = True
                End If
            Catch ex As Exception
                Dim iret As Object = MsgBox("Error Executing SMC.SMCSwitchTestNode0 : " & ex.Message, vbCritical + vbRetryCancel, "Error executing SMC.SMCSwitchTestNode0")
                If iret = vbRetry Then
                    GoTo line1
                ElseIf iret = vbCancel Then
                    MAIN.TimerFunction("STOP")
                    MAIN.Button4.Text = "ERROR"
                    MAIN.TreeView1.SelectedNode.Nodes(2).Nodes(0).BackColor = Color.Red
                    MAIN.Button43.Text = "ERROR"
                    MAIN.TextBox162.Text = "ERROR"
                    MAIN.ListBox1.Items.Add($"{DateTime.Now.ToString} Error Executing SMC.SMCSwitchTestNode0 :" & ex.Message)
                    MAIN.ListBox1.SelectedIndex = MAIN.ListBox1.Items.Count - 1
                    IMAGE_STR.PictureBox2.Image = CapturedImg
                    UUT_RESULT.RadioButton3.Checked = True
                    UUT_RESULT.TextBox2.Text = "D0691_ON"
                    UUT_RESULT.ShowDialog()
                End If
            End Try
        Else
            MAIN.TabControl2.SelectedTab = MAIN.TabPage8
            MAIN.Button4.Text = "RUNNING"
        End If
    End Function
    Function SMCSwitchTestNode1() 'D0661 ON
        FUNNCTIONLIB.CalculateProgress(34, 19)
        MAIN.ListBox1.Items.Add($"{DateTime.Now.ToString} Started Running  D0661_CMD_ON Step")
        MAIN.ListBox1.SelectedIndex = MAIN.ListBox1.Items.Count - 1
        Dim line1 As Label
        If MAIN.RadioButton1.Checked = True Then
            MAIN.TreeView1.SelectedNode.Nodes(2).Nodes(1).BackColor = Color.Yellow
            MAIN.Button26.Text = "RUNNING"
line1:
            Try
                MAIN.ListBox2.Items.Clear()
                FUNNCTIONLIB.RestoreApp(CONFIG.TextBox57.Text)
                System.Threading.Thread.Sleep(3000)
                If SETUP.CheckBox6.Checked = True Then
                    MAIN.ListBox1.Items.Add($"{DateTime.Now.ToString} AutoImageComparison Enabled")
                    MAIN.ListBox1.SelectedIndex = MAIN.ListBox1.Items.Count - 1
                    CapturedImg = FUNNCTIONLIB.TakeSnapofActiveApp()
                    Reference = Image.FromFile(CONFIG.TextBox42.Text)
                    System.Threading.Thread.Sleep(3000)
                    MAIN.ListBox1.Items.Add($"{DateTime.Now.ToString} Loading Coordinates from Confriguration File")
                    MAIN.ListBox1.SelectedIndex = MAIN.ListBox1.Items.Count - 1
                    ini.Load("C:\PLUG_POWER_FCT\PLUG_POWER_FCT\Project Files\Confrigurations.ini")
                    MAIN.ListBox2.Items.Add(ini.Sections("SMC_D0661_CMDON").Keys("Coordinate1").Value)
                    MAIN.ListBox1.Items.Add($"{DateTime.Now.ToString} {ini.Sections("SMC_D0661_CMDON").Keys("Coordinate1").Value}")
                    MAIN.ListBox1.SelectedIndex = MAIN.ListBox1.Items.Count - 1
                    MAIN.ListBox2.Items.Add(ini.Sections("SMC_D0661_CMDON").Keys("Coordinate2").Value)
                    MAIN.ListBox1.Items.Add($"{DateTime.Now.ToString} {ini.Sections("SMC_D0661_CMDON").Keys("Coordinate2").Value}")
                    MAIN.ListBox1.SelectedIndex = MAIN.ListBox1.Items.Count - 1
                    SMC_IMAGECOMP.PictureBox1.Image = CapturedImg
                    SMC_IMAGECOMP.PictureBox2.Image = Reference
                    SMC_IMAGECOMP.TextBox3.Text = "SMC_D0661_CMDON"
                    MAIN.ListBox1.Items.Add($"{DateTime.Now.ToString} Confriguration Loaded successfully")
                    MAIN.ListBox1.SelectedIndex = MAIN.ListBox1.Items.Count - 1
                    MAIN.ListBox1.Items.Add($"{DateTime.Now.ToString} Starting Comparison Function of Coordintaes")
                    MAIN.ListBox1.SelectedIndex = MAIN.ListBox1.Items.Count - 1
                    SMC_IMAGECOMP.ShowDialog()
                    If MAIN.Button26.Text = "PASS" Then
                        IMAGE_STR.PictureBox3.Image = CapturedImg
                        MAIN.PictureBox3.Image = Image.FromFile(CONFIG.TextBox41.Text)
                        MAIN.TextBox173.Text = "CLICK ON D0661 CMD AGAIN AS SHOWN AND CLICK ON CONTINUE"
                        MAIN.PictureBox3.Visible = True
                        MAIN.TextBox173.Visible = True
                        MAIN.Button49.Visible = True
                    Else
                        IMAGE_STR.PictureBox3.Image = CapturedImg
                        MAIN.TimerFunction("STOP")
                        UUT_RESULT.RadioButton2.Checked = True
                        UUT_RESULT.TextBox2.Text = "D0661_ON"
                        UUT_RESULT.ShowDialog()
                    End If
                ElseIf SETUP.CheckBox6.Checked = False Then
                    MAIN.ListBox1.Items.Add($"{DateTime.Now.ToString} AutoImageComparison Disabled")
                    MAIN.ListBox1.SelectedIndex = MAIN.ListBox1.Items.Count - 1
                    MAIN.Button66.Visible = True
                    MAIN.Button67.Visible = True
                    MAIN.PictureBox3.Image = Image.FromFile(CONFIG.TextBox41.Text)
                    MAIN.TextBox173.Text = "CONFIRM THE SCREEN AS SHOWN AND CLICK ON CONFIRM OTHERWISE CLICK ON NON-CONFIRM"
                    MAIN.TextBox173.Enabled = False
                    MAIN.PictureBox3.Visible = True
                    MAIN.TextBox173.Visible = True
                End If
            Catch ex As Exception
                Dim iret As Object = MsgBox("Error Executing SMC.SMCSwitchTestNode1 : " & ex.Message, vbCritical + vbRetryCancel, "Error executing SMC.SMCSwitchTestNode1")
                If iret = vbRetry Then
                    GoTo line1
                ElseIf iret = vbCancel Then
                    MAIN.TimerFunction("STOP")
                    MAIN.Button26.Text = "ERROR"
                    MAIN.TreeView1.SelectedNode.Nodes(2).Nodes(0).BackColor = Color.Red
                    MAIN.Button43.Text = "ERROR"
                    MAIN.TextBox82.Text = "ERROR"
                    MAIN.ListBox1.Items.Add($"{DateTime.Now.ToString} Error Executing SMC.SMCSwitchTestNode1 :" & ex.Message)
                    MAIN.ListBox1.SelectedIndex = MAIN.ListBox1.Items.Count - 1
                    IMAGE_STR.PictureBox3.Image = CapturedImg
                    UUT_RESULT.RadioButton3.Checked = True
                    UUT_RESULT.TextBox2.Text = "D0661_ON"
                    UUT_RESULT.ShowDialog()
                End If
            End Try
        Else
            MAIN.TabControl2.SelectedTab = MAIN.TabPage8
            MAIN.Button4.Text = "RUNNING"
        End If
    End Function
    Function SMCSwitchTestNode2() 'DO661 OFF
        FUNNCTIONLIB.CalculateProgress(34, 20)
        MAIN.ListBox1.Items.Add($"{DateTime.Now.ToString} Started Running  D0661_CMD_OFF Step")
        MAIN.ListBox1.SelectedIndex = MAIN.ListBox1.Items.Count - 1
        Dim line1 As Label
        If MAIN.RadioButton1.Checked = True Then
            MAIN.TreeView1.SelectedNode.Nodes(2).Nodes(2).BackColor = Color.Yellow
            MAIN.Button27.Text = "RUNNING"
line1:
            Try
                MAIN.ListBox2.Items.Clear()
                FUNNCTIONLIB.RestoreApp(CONFIG.TextBox57.Text)
                System.Threading.Thread.Sleep(3000)
                If SETUP.CheckBox6.Checked = True Then
                    MAIN.ListBox1.Items.Add($"{DateTime.Now.ToString} AutoImageComparison Enabled")
                    MAIN.ListBox1.SelectedIndex = MAIN.ListBox1.Items.Count - 1
                    CapturedImg = FUNNCTIONLIB.TakeSnapofActiveApp()
                    Reference = Image.FromFile(CONFIG.TextBox43.Text)
                    System.Threading.Thread.Sleep(3000)
                    MAIN.ListBox1.Items.Add($"{DateTime.Now.ToString} Loading Coordinates from Confriguration File")
                    MAIN.ListBox1.SelectedIndex = MAIN.ListBox1.Items.Count - 1
                    ini.Load("C:\PLUG_POWER_FCT\PLUG_POWER_FCT\Project Files\Confrigurations.ini")
                    MAIN.ListBox2.Items.Add(ini.Sections("SMC_D0661_CMDOFF").Keys("Coordinate1").Value)
                    MAIN.ListBox1.Items.Add($"{DateTime.Now.ToString} {ini.Sections("SMC_D0661_CMDOFF").Keys("Coordinate1").Value}")
                    MAIN.ListBox1.SelectedIndex = MAIN.ListBox1.Items.Count - 1
                    MAIN.ListBox2.Items.Add(ini.Sections("SMC_D0661_CMDOFF").Keys("Coordinate2").Value)
                    MAIN.ListBox1.Items.Add($"{DateTime.Now.ToString} {ini.Sections("SMC_D0661_CMDOFF").Keys("Coordinate2").Value}")
                    MAIN.ListBox1.SelectedIndex = MAIN.ListBox1.Items.Count - 1
                    SMC_IMAGECOMP.PictureBox1.Image = CapturedImg
                    SMC_IMAGECOMP.PictureBox2.Image = Reference
                    SMC_IMAGECOMP.TextBox3.Text = "SMC_D0661_CMDOFF"
                    MAIN.ListBox1.Items.Add($"{DateTime.Now.ToString} Confriguration Loaded successfully")
                    MAIN.ListBox1.SelectedIndex = MAIN.ListBox1.Items.Count - 1
                    MAIN.ListBox1.Items.Add($"{DateTime.Now.ToString} Starting Comparison Function of Coordintaes")
                    MAIN.ListBox1.SelectedIndex = MAIN.ListBox1.Items.Count - 1
                    SMC_IMAGECOMP.ShowDialog()
                    If MAIN.Button27.Text = "PASS" Then
                        IMAGE_STR.PictureBox4.Image = CapturedImg
                        MAIN.PictureBox3.Image = Image.FromFile(CONFIG.TextBox45.Text)
                        MAIN.TextBox173.Text = "CLICK ON AUX_EN_CMD AS SHOWN THEN CLICK ON CONTINUE"
                        MAIN.PictureBox3.Visible = True
                        MAIN.TextBox173.Visible = True
                        MAIN.Button50.Visible = True
                    Else
                        MAIN.TimerFunction("STOP")
                        IMAGE_STR.PictureBox4.Image = CapturedImg
                        UUT_RESULT.RadioButton2.Checked = True
                        UUT_RESULT.TextBox2.Text = "D0661_OFF"
                        UUT_RESULT.ShowDialog()
                    End If
                ElseIf SETUP.CheckBox6.Checked = False Then
                    MAIN.Button68.Visible = True
                    MAIN.Button69.Visible = True
                    MAIN.TextBox173.Text = "CONFIRM THE SCREEN AS SHOWN AND CLICK ON CONFIRM OTHERWISE CLICK ON NON-CONFIRM"
                    MAIN.TextBox173.Enabled = False
                    MAIN.TextBox173.Visible = True
                    MAIN.PictureBox3.Image = Image.FromFile(CONFIG.TextBox44.Text)
                    MAIN.PictureBox3.Visible = True
                End If
            Catch ex As Exception
                Dim iret As Object = MsgBox("Error Executing SMC.SMCSwitchTestNode2 : " & ex.Message, vbCritical + vbRetryCancel, "Error executing SMC.SMCSwitchTestNode2")
                If iret = vbRetry Then
                    GoTo line1
                ElseIf iret = vbCancel Then
                    MAIN.TimerFunction("STOP")
                    MAIN.Button27.Text = "ERROR"
                    MAIN.TreeView1.SelectedNode.Nodes(2).Nodes(2).BackColor = Color.Red
                    MAIN.Button43.Text = "ERROR"
                    MAIN.TextBox86.Text = "ERROR"
                    MAIN.ListBox1.Items.Add($"{DateTime.Now.ToString} Error Executing SMC.SMCSwitchTestNode2 :" & ex.Message)
                    MAIN.ListBox1.SelectedIndex = MAIN.ListBox1.Items.Count - 1
                    IMAGE_STR.PictureBox4.Image = CapturedImg
                    UUT_RESULT.RadioButton3.Checked = True
                    UUT_RESULT.TextBox2.Text = "D0661_OFF"
                    UUT_RESULT.ShowDialog()
                End If
            End Try
        End If
    End Function
    Function SMCSwitchTestNode3() 'AUXEN CMD ON
        FUNNCTIONLIB.CalculateProgress(34, 21)
        MAIN.ListBox1.Items.Add($"{DateTime.Now.ToString} Started Running  AUXEN_CMD_ON Step")
        MAIN.ListBox1.SelectedIndex = MAIN.ListBox1.Items.Count - 1
        Dim line1 As Label
        If MAIN.RadioButton1.Checked = True Then
            MAIN.TreeView1.SelectedNode.Nodes(2).Nodes(3).BackColor = Color.Yellow
            MAIN.Button28.Text = "RUNNING"
line1:
            MAIN.ListBox2.Items.Clear()
            Try
                If SETUP.CheckBox6.Checked = True Then
                    MAIN.ListBox1.Items.Add($"{DateTime.Now.ToString} AutoImageComparison Enabled")
                    MAIN.ListBox1.SelectedIndex = MAIN.ListBox1.Items.Count - 1
                    FUNNCTIONLIB.RestoreApp(CONFIG.TextBox57.Text)
                    System.Threading.Thread.Sleep(1000)
                    CapturedImg = FUNNCTIONLIB.TakeSnapofActiveApp()
                    Reference = Image.FromFile(CONFIG.TextBox46.Text)
                    System.Threading.Thread.Sleep(3000)
                    MAIN.ListBox1.Items.Add($"{DateTime.Now.ToString} Loading Coordinates from Confriguration File")
                    MAIN.ListBox1.SelectedIndex = MAIN.ListBox1.Items.Count - 1
                    ini.Load("C:\PLUG_POWER_FCT\PLUG_POWER_FCT\Project Files\Confrigurations.ini")
                    MAIN.ListBox2.Items.Add(ini.Sections("SMC_AUXENCMD_ON").Keys("Coordinate1").Value)
                    MAIN.ListBox1.Items.Add($"{DateTime.Now.ToString} {ini.Sections("SMC_AUXENCMD_ON").Keys("Coordinate1").Value}")
                    MAIN.ListBox1.SelectedIndex = MAIN.ListBox1.Items.Count - 1
                    SMC_IMAGECOMP.PictureBox1.Image = CapturedImg
                    SMC_IMAGECOMP.PictureBox2.Image = Reference
                    SMC_IMAGECOMP.TextBox3.Text = "AUXENCMD_ON"
                    MAIN.ListBox1.Items.Add($"{DateTime.Now.ToString} Confriguration Loaded successfully")
                    MAIN.ListBox1.SelectedIndex = MAIN.ListBox1.Items.Count - 1
                    MAIN.ListBox1.Items.Add($"{DateTime.Now.ToString} Starting Comparison Function of Coordintaes")
                    MAIN.ListBox1.SelectedIndex = MAIN.ListBox1.Items.Count - 1
                    SMC_IMAGECOMP.ShowDialog()
                    If MAIN.Button28.Text = "PASS" Then
                        IMAGE_STR.PictureBox5.Image = CapturedImg
                        MAIN.PictureBox3.Image = Image.FromFile(CONFIG.TextBox47.Text)
                        MAIN.TextBox173.Text = "CLICK ON D0100_CMD THEN CLICK ON CONTINUE"
                        MAIN.PictureBox3.Visible = True
                        MAIN.TextBox173.Visible = True
                        MAIN.Button51.Visible = True
                    Else
                        MAIN.TimerFunction("STOP")
                        IMAGE_STR.PictureBox5.Image = CapturedImg
                        UUT_RESULT.RadioButton2.Checked = True
                        UUT_RESULT.TextBox2.Text = "AUXENCMD_ON"
                        UUT_RESULT.ShowDialog()
                    End If
                ElseIf SETUP.CheckBox6.Checked = False Then
                    MAIN.Button70.Visible = True
                    MAIN.Button71.Visible = True
                    MAIN.TextBox173.Text = "CONFIRM THE SCREEN AS SHOWN AND CLICK ON CONFIRM OTHERWISE CLICK ON NON-CONFIRM"
                    MAIN.TextBox173.Enabled = False
                    MAIN.TextBox173.Visible = True
                    MAIN.PictureBox3.Image = Image.FromFile(CONFIG.TextBox45.Text)
                    MAIN.PictureBox3.Visible = True
                End If
            Catch ex As Exception
                Dim iret As Object = MsgBox("Error Executing SMC.SMCSwitchTestNode3 : " & ex.Message, vbCritical + vbRetryCancel, "Error executing SMC.SMCSwitchTestNode3")
                If iret = vbRetry Then
                    GoTo line1
                ElseIf iret = vbCancel Then
                    MAIN.TimerFunction("STOP")
                    MAIN.Button28.Text = "ERROR"
                    MAIN.TreeView1.SelectedNode.Nodes(2).Nodes(3).BackColor = Color.Red
                    MAIN.Button43.Text = "ERROR"
                    MAIN.TextBox90.Text = "ERROR"
                    MAIN.ListBox1.Items.Add($"{DateTime.Now.ToString} Error Executing SMC.SMCSwitchTestNode3 :" & ex.Message)
                    MAIN.ListBox1.SelectedIndex = MAIN.ListBox1.Items.Count - 1
                    IMAGE_STR.PictureBox5.Image = CapturedImg
                    UUT_RESULT.RadioButton3.Checked = True
                    UUT_RESULT.TextBox2.Text = "AUXENCMD_ON"
                    UUT_RESULT.ShowDialog()
                End If
            End Try
        End If
    End Function
    Function SMCSwitchTestNode4() 'DO100 ON
        FUNNCTIONLIB.CalculateProgress(34, 22)
        MAIN.ListBox1.Items.Add($"{DateTime.Now.ToString} Started Running  D0100_CMD_ON Step")
        MAIN.ListBox1.SelectedIndex = MAIN.ListBox1.Items.Count - 1
        Dim line1 As Label
        If MAIN.RadioButton1.Checked = True Then
            MAIN.TreeView1.SelectedNode.Nodes(2).Nodes(4).BackColor = Color.Yellow
            MAIN.Button29.Text = "RUNNING"
line1:
            MAIN.ListBox2.Items.Clear()
            Try
                If SETUP.CheckBox6.Checked = True Then
                    MAIN.ListBox1.Items.Add($"{DateTime.Now.ToString} AutoImageComparison Enabled")
                    MAIN.ListBox1.SelectedIndex = MAIN.ListBox1.Items.Count - 1
                    FUNNCTIONLIB.RestoreApp(CONFIG.TextBox57.Text)
                    System.Threading.Thread.Sleep(1000)
                    CapturedImg = FUNNCTIONLIB.TakeSnapofActiveApp()
                    Reference = Image.FromFile(CONFIG.TextBox59.Text)
                    System.Threading.Thread.Sleep(3000)
                    MAIN.ListBox1.Items.Add($"{DateTime.Now.ToString} Loading Coordinates from Confriguration File")
                    MAIN.ListBox1.SelectedIndex = MAIN.ListBox1.Items.Count - 1
                    ini.Load("C:\PLUG_POWER_FCT\PLUG_POWER_FCT\Project Files\Confrigurations.ini")
                    MAIN.ListBox2.Items.Add(ini.Sections("SMC_D0100_ON").Keys("Coordinate1").Value)
                    MAIN.ListBox1.Items.Add($"{DateTime.Now.ToString} {ini.Sections("SMC_D0100_ON").Keys("Coordinate1").Value}")
                    MAIN.ListBox1.SelectedIndex = MAIN.ListBox1.Items.Count - 1
                    SMC_IMAGECOMP.PictureBox1.Image = CapturedImg
                    SMC_IMAGECOMP.PictureBox2.Image = Reference
                    SMC_IMAGECOMP.TextBox3.Text = "D0100_ON"
                    MAIN.ListBox1.Items.Add($"{DateTime.Now.ToString} Confriguration Loaded successfully")
                    MAIN.ListBox1.SelectedIndex = MAIN.ListBox1.Items.Count - 1
                    MAIN.ListBox1.Items.Add($"{DateTime.Now.ToString} Starting Comparison Function of Coordintaes")
                    MAIN.ListBox1.SelectedIndex = MAIN.ListBox1.Items.Count - 1
                    SMC_IMAGECOMP.ShowDialog()
                    If MAIN.Button29.Text = "PASS" Then
                        IMAGE_STR.PictureBox6.Image = CapturedImg
                        MAIN.PictureBox3.Image = Image.FromFile(CONFIG.TextBox60.Text)
                        MAIN.TextBox173.Text = "CLICK ON D0100_CMD THEN CLICK ON CONTINUE"
                        MAIN.PictureBox3.Visible = True
                        MAIN.TextBox173.Visible = True
                        MAIN.Button52.Visible = True
                    Else
                        IMAGE_STR.PictureBox6.Image = CapturedImg
                        MAIN.TimerFunction("STOP")
                        UUT_RESULT.RadioButton2.Checked = True
                        UUT_RESULT.TextBox2.Text = "D0100_ON"
                        UUT_RESULT.ShowDialog()
                    End If
                ElseIf SETUP.CheckBox6.Checked = False Then
                    MAIN.Button72.Visible = True
                    MAIN.Button73.Visible = True
                    MAIN.TextBox173.Text = "CONFIRM THE SCREEN AS SHOWN AND CLICK ON CONFIRM OTHERWISE CLICK ON NON-CONFIRM"
                    MAIN.TextBox173.Enabled = False
                    MAIN.TextBox173.Visible = True
                    MAIN.PictureBox3.Image = Image.FromFile(CONFIG.TextBox47.Text)
                    MAIN.PictureBox3.Visible = True
                End If
            Catch ex As Exception
                Dim iret As Object = MsgBox("Error Executing SMC.SMCSwitchTestNode4 : " & ex.Message, vbCritical + vbRetryCancel, "Error executing SMC.SMCSwitchTestNode4")
                If iret = vbRetry Then
                    GoTo line1
                ElseIf iret = vbCancel Then
                    MAIN.TimerFunction("STOP")
                    MAIN.Button29.Text = "ERROR"
                    MAIN.TreeView1.SelectedNode.Nodes(2).Nodes(4).BackColor = Color.Red
                    MAIN.Button43.Text = "ERROR"
                    MAIN.TextBox95.Text = "ERROR"
                    MAIN.ListBox1.Items.Add($"{DateTime.Now.ToString} Error Executing SMC.SMCSwitchTestNode4 :" & ex.Message)
                    MAIN.ListBox1.SelectedIndex = MAIN.ListBox1.Items.Count - 1
                    IMAGE_STR.PictureBox6.Image = CapturedImg
                    UUT_RESULT.RadioButton3.Checked = True
                    UUT_RESULT.TextBox2.Text = "D0100_ON"
                    UUT_RESULT.ShowDialog()
                End If
            End Try
        End If
    End Function
    Function SMCSwitchTestNode5() 'DO100 OFF
        FUNNCTIONLIB.CalculateProgress(34, 23)
        MAIN.ListBox1.Items.Add($"{DateTime.Now.ToString} Started Running  D0100_CMD_OFF Step")
        MAIN.ListBox1.SelectedIndex = MAIN.ListBox1.Items.Count - 1
        Dim line1 As Label
        If MAIN.RadioButton1.Checked = True Then
            MAIN.TreeView1.SelectedNode.Nodes(2).Nodes(5).BackColor = Color.Yellow
            MAIN.Button30.Text = "RUNNING"
line1:
            MAIN.ListBox2.Items.Clear()
            Try
                If SETUP.CheckBox6.Checked = True Then
                    MAIN.ListBox1.Items.Add($"{DateTime.Now.ToString} AutoImageComparison Enabled")
                    MAIN.ListBox1.SelectedIndex = MAIN.ListBox1.Items.Count - 1
                    FUNNCTIONLIB.RestoreApp(CONFIG.TextBox57.Text)
                    System.Threading.Thread.Sleep(3000)
                    CapturedImg = FUNNCTIONLIB.TakeSnapofActiveApp()
                    Reference = Image.FromFile(CONFIG.TextBox61.Text)
                    System.Threading.Thread.Sleep(3000)
                    MAIN.ListBox1.Items.Add($"{DateTime.Now.ToString} Loading Coordinates from Confriguration File")
                    MAIN.ListBox1.SelectedIndex = MAIN.ListBox1.Items.Count - 1
                    ini.Load("C:\PLUG_POWER_FCT\PLUG_POWER_FCT\Project Files\Confrigurations.ini")
                    MAIN.ListBox2.Items.Add(ini.Sections("SMC_D0100_OFF").Keys("Coordinate1").Value)
                    MAIN.ListBox1.Items.Add($"{DateTime.Now.ToString} {ini.Sections("SMC_D0100_OFF").Keys("Coordinate1").Value}")
                    MAIN.ListBox1.SelectedIndex = MAIN.ListBox1.Items.Count - 1
                    SMC_IMAGECOMP.PictureBox1.Image = CapturedImg
                    SMC_IMAGECOMP.PictureBox2.Image = Reference
                    SMC_IMAGECOMP.TextBox3.Text = "D0100_OFF"
                    MAIN.ListBox1.Items.Add($"{DateTime.Now.ToString} Confriguration Loaded successfully")
                    MAIN.ListBox1.SelectedIndex = MAIN.ListBox1.Items.Count - 1
                    MAIN.ListBox1.Items.Add($"{DateTime.Now.ToString} Starting Comparison Function of Coordintaes")
                    MAIN.ListBox1.SelectedIndex = MAIN.ListBox1.Items.Count - 1
                    SMC_IMAGECOMP.ShowDialog()
                    If MAIN.Button30.Text = "PASS" Then
                        IMAGE_STR.PictureBox7.Image = CapturedImg
                        MAIN.PictureBox3.Image = Image.FromFile(CONFIG.TextBox62.Text)
                        MAIN.TextBox173.Text = "CLICK ON D0101_CMD THEN CLICK ON CONTINUE"
                        MAIN.PictureBox3.Visible = True
                        MAIN.TextBox173.Visible = True
                        MAIN.Button53.Visible = True
                    Else
                        IMAGE_STR.PictureBox7.Image = CapturedImg
                        MAIN.TimerFunction("STOP")
                        UUT_RESULT.RadioButton2.Checked = True
                        UUT_RESULT.TextBox2.Text = "D0100_OFF"
                        UUT_RESULT.ShowDialog()
                    End If
                ElseIf SETUP.CheckBox6.Checked = False Then
                    MAIN.Button74.Visible = True
                    MAIN.Button75.Visible = True
                    MAIN.TextBox173.Text = "CONFIRM THE SCREEN AS SHOWN AND CLICK ON CONFIRM OTHERWISE CLICK ON NON-CONFIRM"
                    MAIN.TextBox173.Enabled = False
                    MAIN.TextBox173.Visible = True
                    MAIN.PictureBox3.Image = Image.FromFile(CONFIG.TextBox60.Text)
                    MAIN.PictureBox3.Visible = True
                End If
            Catch ex As Exception
                Dim iret As Object = MsgBox("Error Executing SMC.SMCSwitchTestNode5 : " & ex.Message, vbCritical + vbRetryCancel, "Error executing SMC.SMCSwitchTestNode5")
                If iret = vbRetry Then
                    GoTo line1
                ElseIf iret = vbCancel Then
                    MAIN.TimerFunction("STOP")
                    MAIN.Button30.Text = "ERROR"
                    MAIN.TreeView1.SelectedNode.Nodes(2).Nodes(5).BackColor = Color.Red
                    MAIN.Button43.Text = "ERROR"
                    MAIN.TextBox99.Text = "ERROR"
                    MAIN.ListBox1.Items.Add($"{DateTime.Now.ToString} Error Executing SMC.SMCSwitchTestNode5 :" & ex.Message)
                    MAIN.ListBox1.SelectedIndex = MAIN.ListBox1.Items.Count - 1
                    IMAGE_STR.PictureBox7.Image = CapturedImg
                    UUT_RESULT.RadioButton3.Checked = True
                    UUT_RESULT.TextBox2.Text = "D0100_OFF"
                    UUT_RESULT.ShowDialog()
                End If
            End Try
        End If
    End Function
    Function SMCSwitchTestNode6() 'DO101 ON
        FUNNCTIONLIB.CalculateProgress(34, 24)
        MAIN.ListBox1.Items.Add($"{DateTime.Now.ToString} Started Running  D0101_CMD_ON Step")
        MAIN.ListBox1.SelectedIndex = MAIN.ListBox1.Items.Count - 1
        Dim line1 As Label
        If MAIN.RadioButton1.Checked = True Then
            MAIN.TreeView1.SelectedNode.Nodes(2).Nodes(6).BackColor = Color.Yellow
            MAIN.Button31.Text = "RUNNING"
line1:
            Try
                MAIN.ListBox2.Items.Clear()
                If SETUP.CheckBox6.Checked = True Then
                    MAIN.ListBox1.Items.Add($"{DateTime.Now.ToString} AutoImageComparison Enabled")
                    MAIN.ListBox1.SelectedIndex = MAIN.ListBox1.Items.Count - 1
                    FUNNCTIONLIB.RestoreApp(CONFIG.TextBox57.Text)
                    System.Threading.Thread.Sleep(3000)
                    CapturedImg = FUNNCTIONLIB.TakeSnapofActiveApp()
                    Reference = Image.FromFile(CONFIG.TextBox63.Text)
                    System.Threading.Thread.Sleep(3000)
                    MAIN.ListBox1.Items.Add($"{DateTime.Now.ToString} Loading Coordinates from Confriguration File")
                    MAIN.ListBox1.SelectedIndex = MAIN.ListBox1.Items.Count - 1
                    ini.Load("C:\PLUG_POWER_FCT\PLUG_POWER_FCT\Project Files\Confrigurations.ini")
                    MAIN.ListBox2.Items.Add(ini.Sections("SMC_D0101_ON").Keys("Coordinate1").Value)
                    MAIN.ListBox1.Items.Add($"{DateTime.Now.ToString} {ini.Sections("SMC_D0101_ON").Keys("Coordinate1").Value}")
                    MAIN.ListBox1.SelectedIndex = MAIN.ListBox1.Items.Count - 1
                    SMC_IMAGECOMP.PictureBox1.Image = CapturedImg
                    SMC_IMAGECOMP.PictureBox2.Image = Reference
                    SMC_IMAGECOMP.TextBox3.Text = "D0101 ON"
                    MAIN.ListBox1.Items.Add($"{DateTime.Now.ToString} Confriguration Loaded successfully")
                    MAIN.ListBox1.SelectedIndex = MAIN.ListBox1.Items.Count - 1
                    MAIN.ListBox1.Items.Add($"{DateTime.Now.ToString} Starting Comparison Function of Coordintaes")
                    MAIN.ListBox1.SelectedIndex = MAIN.ListBox1.Items.Count - 1
                    SMC_IMAGECOMP.ShowDialog()
                    If MAIN.Button31.Text = "PASS" Then
                        IMAGE_STR.PictureBox8.Image = CapturedImg
                        MAIN.PictureBox3.Image = Image.FromFile(CONFIG.TextBox64.Text)
                        MAIN.TextBox173.Text = "CLICK ON D0101_CMD THEN CLICK ON CONTINUE"
                        MAIN.PictureBox3.Visible = True
                        MAIN.TextBox173.Visible = True
                        MAIN.Button54.Visible = True
                    Else
                        IMAGE_STR.PictureBox8.Image = CapturedImg
                        MAIN.TimerFunction("STOP")
                        UUT_RESULT.RadioButton2.Checked = True
                        UUT_RESULT.TextBox2.Text = "D0101_ON"
                        UUT_RESULT.ShowDialog()
                    End If
                ElseIf SETUP.CheckBox6.Checked = False Then
                    MAIN.Button76.Visible = True
                    MAIN.Button77.Visible = True
                    MAIN.TextBox173.Text = "CONFIRM THE SCREEN AS SHOWN AND CLICK ON CONFIRM OTHERWISE CLICK ON NON-CONFIRM"
                    MAIN.TextBox173.Enabled = False
                    MAIN.TextBox173.Visible = True
                    MAIN.PictureBox3.Image = Image.FromFile(CONFIG.TextBox62.Text)
                    MAIN.PictureBox3.Visible = True
                End If
            Catch ex As Exception
                Dim iret As Object = MsgBox("Error Executing SMC.SMCSwitchTestNode6 : " & ex.Message, vbCritical + vbRetryCancel, "Error executing SMC.SMCSwitchTestNode6")
                If iret = vbRetry Then
                    GoTo line1
                ElseIf iret = vbCancel Then
                    MAIN.TimerFunction("STOP")
                    MAIN.Button31.Text = "ERROR"
                    MAIN.TreeView1.SelectedNode.Nodes(2).Nodes(6).BackColor = Color.Red
                    MAIN.Button43.Text = "ERROR"
                    MAIN.TextBox103.Text = "ERROR"
                    MAIN.ListBox1.Items.Add($"{DateTime.Now.ToString} Error Executing SMC.SMCSwitchTestNode6 :" & ex.Message)
                    MAIN.ListBox1.SelectedIndex = MAIN.ListBox1.Items.Count - 1
                    IMAGE_STR.PictureBox8.Image = CapturedImg
                    UUT_RESULT.RadioButton3.Checked = True
                    UUT_RESULT.TextBox2.Text = "D0101_ON"
                    UUT_RESULT.ShowDialog()
                End If
            End Try

        End If
    End Function
    Function SMCSwitchTestNode7() 'D0101 OFF
        FUNNCTIONLIB.CalculateProgress(34, 25)
        MAIN.ListBox1.Items.Add($"{DateTime.Now.ToString} Started Running  D0101_CMD_OFF Step")
        MAIN.ListBox1.SelectedIndex = MAIN.ListBox1.Items.Count - 1
        Dim line1 As Label
        If MAIN.RadioButton1.Checked = True Then
            MAIN.TreeView1.SelectedNode.Nodes(2).Nodes(7).BackColor = Color.Yellow
            MAIN.Button32.Text = "RUNNING"
line1:
            Try
                MAIN.ListBox2.Items.Clear()
                If SETUP.CheckBox6.Checked = True Then
                    MAIN.ListBox1.Items.Add($"{DateTime.Now.ToString} AutoImageComparison Enabled")
                    MAIN.ListBox1.SelectedIndex = MAIN.ListBox1.Items.Count - 1
                    FUNNCTIONLIB.RestoreApp(CONFIG.TextBox57.Text)
                    System.Threading.Thread.Sleep(3000)
                    CapturedImg = FUNNCTIONLIB.TakeSnapofActiveApp()
                    Reference = Image.FromFile(CONFIG.TextBox65.Text)
                    System.Threading.Thread.Sleep(3000)
                    MAIN.ListBox1.Items.Add($"{DateTime.Now.ToString} Loading Coordinates from Confriguration File")
                    MAIN.ListBox1.SelectedIndex = MAIN.ListBox1.Items.Count - 1
                    ini.Load("C:\PLUG_POWER_FCT\PLUG_POWER_FCT\Project Files\Confrigurations.ini")
                    MAIN.ListBox2.Items.Add(ini.Sections("SMC_D0101_OFF").Keys("Coordinate1").Value)
                    MAIN.ListBox1.Items.Add($"{DateTime.Now.ToString} {ini.Sections("SMC_D0101_OFF").Keys("Coordinate1").Value}")
                    MAIN.ListBox1.SelectedIndex = MAIN.ListBox1.Items.Count - 1
                    SMC_IMAGECOMP.PictureBox1.Image = CapturedImg
                    SMC_IMAGECOMP.PictureBox2.Image = Reference
                    SMC_IMAGECOMP.TextBox3.Text = "D0101 OFF"
                    MAIN.ListBox1.Items.Add($"{DateTime.Now.ToString} Confriguration Loaded successfully")
                    MAIN.ListBox1.SelectedIndex = MAIN.ListBox1.Items.Count - 1
                    MAIN.ListBox1.Items.Add($"{DateTime.Now.ToString} Starting Comparison Function of Coordintaes")
                    MAIN.ListBox1.SelectedIndex = MAIN.ListBox1.Items.Count - 1
                    SMC_IMAGECOMP.ShowDialog()
                    If MAIN.Button32.Text = "PASS" Then
                        IMAGE_STR.PictureBox9.Image = CapturedImg
                        MAIN.PictureBox3.Image = Image.FromFile(CONFIG.TextBox66.Text)
                        MAIN.TextBox173.Text = "TURN OFF ALL THE CMD SWITCHES"
                        MAIN.PictureBox3.Visible = True
                        MAIN.TextBox173.Visible = True
                        MAIN.TextBox173.Enabled = False
                        MAIN.Button55.Visible = True
                    Else
                        IMAGE_STR.PictureBox9.Image = CapturedImg
                        MAIN.TimerFunction("STOP")
                        UUT_RESULT.RadioButton2.Checked = True
                        UUT_RESULT.TextBox2.Text = "D0101_OFF"
                        UUT_RESULT.ShowDialog()
                    End If
                ElseIf SETUP.CheckBox6.Checked = False Then
                    MAIN.Button78.Visible = True
                    MAIN.Button79.Visible = True
                    MAIN.TextBox173.Text = "CONFIRM THE SCREEN AS SHOWN AND CLICK ON CONFIRM OTHERWISE CLICK ON NON-CONFIRM"
                    MAIN.TextBox173.Enabled = False
                    MAIN.TextBox173.Visible = True
                    MAIN.PictureBox3.Image = Image.FromFile(CONFIG.TextBox64.Text)
                    MAIN.PictureBox3.Visible = True
                End If
            Catch ex As Exception
                Dim iret As Object = MsgBox("Error Executing SMC.SMCSwitchTestNode7 : " & ex.Message, vbCritical + vbRetryCancel, "Error executing SMC.SMCSwitchTestNode7")
                If iret = vbRetry Then
                    GoTo line1
                ElseIf iret = vbCancel Then
                    MAIN.TimerFunction("STOP")
                    MAIN.Button32.Text = "ERROR"
                    MAIN.TreeView1.SelectedNode.Nodes(2).Nodes(7).BackColor = Color.Red
                    MAIN.Button43.Text = "ERROR"
                    MAIN.TextBox107.Text = "ERROR"
                    MAIN.ListBox1.Items.Add($"{DateTime.Now.ToString} Error Executing SMC.SMCSwitchTestNode7 :" & ex.Message)
                    MAIN.ListBox1.SelectedIndex = MAIN.ListBox1.Items.Count - 1
                    IMAGE_STR.PictureBox9.Image = CapturedImg
                    UUT_RESULT.RadioButton3.Checked = True
                    UUT_RESULT.TextBox2.Text = "D0101_OFF"
                    UUT_RESULT.ShowDialog()
                End If
            End Try
        End If
    End Function
    Function SMCSwitchTestNode8() 'CMD OFF
        FUNNCTIONLIB.CalculateProgress(34, 26)
        MAIN.ListBox1.Items.Add($"{DateTime.Now.ToString} Started Running  CMD_SWITCH_OFF Step")
        MAIN.ListBox1.SelectedIndex = MAIN.ListBox1.Items.Count - 1
        Dim line1 As Label
        If MAIN.RadioButton1.Checked = True Then
            MAIN.TreeView1.SelectedNode.Nodes(2).Nodes(8).BackColor = Color.Yellow
            MAIN.Button33.Text = "RUNNING"
line1:
            Try
                MAIN.ListBox2.Items.Clear()
                If SETUP.CheckBox6.Checked = True Then
                    MAIN.ListBox1.Items.Add($"{DateTime.Now.ToString} AutoImageComparison Enabled")
                    MAIN.ListBox1.SelectedIndex = MAIN.ListBox1.Items.Count - 1
                    FUNNCTIONLIB.RestoreApp(CONFIG.TextBox57.Text)
                    System.Threading.Thread.Sleep(3000)
                    CapturedImg = FUNNCTIONLIB.TakeSnapofActiveApp()
                    Reference = Image.FromFile(CONFIG.TextBox67.Text)
                    System.Threading.Thread.Sleep(3000)
                    MAIN.ListBox1.Items.Add($"{DateTime.Now.ToString} Loading Coordinates from Confriguration File")
                    MAIN.ListBox1.SelectedIndex = MAIN.ListBox1.Items.Count - 1
                    ini.Load("C:\PLUG_POWER_FCT\PLUG_POWER_FCT\Project Files\Confrigurations.ini")
                    MAIN.ListBox2.Items.Add(ini.Sections("SMC_CMD_OFF").Keys("Coordinate1").Value)
                    MAIN.ListBox1.Items.Add($"{DateTime.Now.ToString} {ini.Sections("SMC_CMD_OFF").Keys("Coordinate1").Value}")
                    MAIN.ListBox1.SelectedIndex = MAIN.ListBox1.Items.Count - 1
                    SMC_IMAGECOMP.PictureBox1.Image = CapturedImg
                    SMC_IMAGECOMP.PictureBox2.Image = Reference
                    SMC_IMAGECOMP.TextBox3.Text = "CMD OFF"
                    MAIN.ListBox1.Items.Add($"{DateTime.Now.ToString} Confriguration Loaded successfully")
                    MAIN.ListBox1.SelectedIndex = MAIN.ListBox1.Items.Count - 1
                    MAIN.ListBox1.Items.Add($"{DateTime.Now.ToString} Starting Comparison Function of Coordintaes")
                    MAIN.ListBox1.SelectedIndex = MAIN.ListBox1.Items.Count - 1
                    SMC_IMAGECOMP.ShowDialog()
                    If MAIN.Button33.Text = "PASS" Then
                        IMAGE_STR.PictureBox10.Image = CapturedImg
                        MAIN.PictureBox3.Image = Image.FromFile(CONFIG.TextBox68.Text)
                        MAIN.TextBox173.Text = "TURN ON THE SW2 SWITCH AND CLICK ON CONTINUE"
                        MAIN.PictureBox3.Visible = True
                        MAIN.TextBox173.Visible = True
                        MAIN.TextBox173.Enabled = False
                        MAIN.Button56.Visible = True
                    Else
                        IMAGE_STR.PictureBox10.Image = CapturedImg
                        MAIN.TimerFunction("STOP")
                        UUT_RESULT.RadioButton2.Checked = True
                        UUT_RESULT.TextBox2.Text = "CMD_OFF"
                        UUT_RESULT.ShowDialog()
                    End If
                ElseIf SETUP.CheckBox6.Checked = False Then
                    MAIN.Button80.Visible = True
                    MAIN.Button81.Visible = True
                    MAIN.TextBox173.Text = "CONFIRM THE SCREEN AS SHOWN AND CLICK ON CONFIRM OTHERWISE CLICK ON NON-CONFIRM"
                    MAIN.TextBox173.Enabled = False
                    MAIN.TextBox173.Visible = True
                    MAIN.PictureBox3.Image = Image.FromFile(CONFIG.TextBox66.Text)
                    MAIN.PictureBox3.Visible = True
                End If
            Catch ex As Exception
                Dim iret As Object = MsgBox("Error Executing SMC.SMCSwitchTestNode8 : " & ex.Message, vbCritical + vbRetryCancel, "Error executing SMC.SMCSwitchTestNode8")
                If iret = vbRetry Then
                    GoTo line1
                ElseIf iret = vbCancel Then
                    MAIN.TimerFunction("STOP")
                    MAIN.Button33.Text = "ERROR"
                    MAIN.TreeView1.SelectedNode.Nodes(2).Nodes(8).BackColor = Color.Red
                    MAIN.Button43.Text = "ERROR"
                    MAIN.TextBox111.Text = "ERROR"
                    MAIN.ListBox1.Items.Add($"{DateTime.Now.ToString} Error Executing SMC.SMCSwitchTestNode8 :" & ex.Message)
                    MAIN.ListBox1.SelectedIndex = MAIN.ListBox1.Items.Count - 1
                    IMAGE_STR.PictureBox10.Image = CapturedImg
                    UUT_RESULT.RadioButton3.Checked = True
                    UUT_RESULT.TextBox2.Text = "CMD_OFF"
                    UUT_RESULT.ShowDialog()
                End If
            End Try
        End If
    End Function
    Function SMCSwitchTestNode9() 'SW2 ON
        FUNNCTIONLIB.CalculateProgress(34, 27)
        MAIN.ListBox1.Items.Add($"{DateTime.Now.ToString} Started Running  SW2_ON Step")
        MAIN.ListBox1.SelectedIndex = MAIN.ListBox1.Items.Count - 1
        Dim line1 As Label
        If MAIN.RadioButton1.Checked = True Then
            MAIN.TreeView1.SelectedNode.Nodes(2).Nodes(9).BackColor = Color.Yellow
            MAIN.Button5.Text = "RUNNING"
line1:
            Try
                MAIN.ListBox2.Items.Clear()
                If SETUP.CheckBox6.Checked = True Then
                    MAIN.ListBox1.Items.Add($"{DateTime.Now.ToString} AutoImageComparison Enabled")
                    MAIN.ListBox1.SelectedIndex = MAIN.ListBox1.Items.Count - 1
                    FUNNCTIONLIB.RestoreApp(CONFIG.TextBox57.Text)
                    System.Threading.Thread.Sleep(3000)
                    CapturedImg = FUNNCTIONLIB.TakeSnapofActiveApp()
                    Reference = Image.FromFile(CONFIG.TextBox69.Text)
                    System.Threading.Thread.Sleep(3000)
                    MAIN.ListBox1.Items.Add($"{DateTime.Now.ToString} Loading Coordinates from Confriguration File")
                    MAIN.ListBox1.SelectedIndex = MAIN.ListBox1.Items.Count - 1
                    ini.Load("C:\PLUG_POWER_FCT\PLUG_POWER_FCT\Project Files\Confrigurations.ini")
                    MAIN.ListBox2.Items.Add(ini.Sections("SMC_SW2_ON").Keys("Coordinate1").Value)
                    MAIN.ListBox1.Items.Add($"{DateTime.Now.ToString} {ini.Sections("SMC_SW2_ON").Keys("Coordinate1").Value}")
                    MAIN.ListBox1.SelectedIndex = MAIN.ListBox1.Items.Count - 1
                    SMC_IMAGECOMP.PictureBox1.Image = CapturedImg
                    SMC_IMAGECOMP.PictureBox2.Image = Reference
                    SMC_IMAGECOMP.TextBox3.Text = "SW2 ON"
                    MAIN.ListBox1.Items.Add($"{DateTime.Now.ToString} Confriguration Loaded successfully")
                    MAIN.ListBox1.SelectedIndex = MAIN.ListBox1.Items.Count - 1
                    MAIN.ListBox1.Items.Add($"{DateTime.Now.ToString} Starting Comparison Function of Coordintaes")
                    MAIN.ListBox1.SelectedIndex = MAIN.ListBox1.Items.Count - 1
                    SMC_IMAGECOMP.ShowDialog()
                    If MAIN.Button5.Text = "PASS" Then
                        IMAGE_STR.PictureBox11.Image = CapturedImg
                        MAIN.PictureBox3.Image = Image.FromFile(CONFIG.TextBox70.Text)
                        MAIN.TextBox173.Text = "TURN ON THE SW3 SWITCH AND CLICK ON CONTINUE"
                        MAIN.PictureBox3.Visible = True
                        MAIN.TextBox173.Visible = True
                        MAIN.TextBox173.Enabled = False
                        MAIN.Button57.Visible = True
                    Else
                        IMAGE_STR.PictureBox11.Image = CapturedImg
                        MAIN.TimerFunction("STOP")
                        UUT_RESULT.RadioButton2.Checked = True
                        UUT_RESULT.TextBox2.Text = "SW2_ON"
                        UUT_RESULT.ShowDialog()
                    End If
                ElseIf SETUP.CheckBox6.Checked = False Then
                    MAIN.Button82.Visible = True
                    MAIN.Button83.Visible = True
                    MAIN.TextBox173.Text = "CONFIRM THE SCREEN AS SHOWN AND CLICK ON CONFIRM OTHERWISE CLICK ON NON-CONFIRM"
                    MAIN.TextBox173.Enabled = False
                    MAIN.TextBox173.Visible = True
                    MAIN.PictureBox3.Image = Image.FromFile(CONFIG.TextBox68.Text)
                    MAIN.PictureBox3.Visible = True
                End If
            Catch ex As Exception
                Dim iret As Object = MsgBox("Error Executing SMC.SMCSwitchTestNode9 : " & ex.Message, vbCritical + vbRetryCancel, "Error executing SMC.SMCSwitchTestNode9")
                If iret = vbRetry Then
                    GoTo line1
                ElseIf iret = vbCancel Then
                    MAIN.TimerFunction("STOP")
                    MAIN.Button5.Text = "ERROR"
                    MAIN.TreeView1.SelectedNode.Nodes(2).Nodes(9).BackColor = Color.Red
                    MAIN.Button43.Text = "ERROR"
                    MAIN.TextBox167.Text = "ERROR"
                    MAIN.ListBox1.Items.Add($"{DateTime.Now.ToString} Error Executing SMC.SMCSwitchTestNode9 :" & ex.Message)
                    MAIN.ListBox1.SelectedIndex = MAIN.ListBox1.Items.Count - 1
                    IMAGE_STR.PictureBox11.Image = CapturedImg
                    MAIN.TimerFunction("STOP")
                    UUT_RESULT.RadioButton3.Checked = True
                    UUT_RESULT.TextBox2.Text = "SW2_ON"
                    UUT_RESULT.ShowDialog()
                End If
            End Try
        End If
    End Function
    Function SMCSwitchTestNode10() 'SW3 ON
        FUNNCTIONLIB.CalculateProgress(34, 28)
        MAIN.ListBox1.Items.Add($"{DateTime.Now.ToString} Started Running  SW3_ON Step")
        MAIN.ListBox1.SelectedIndex = MAIN.ListBox1.Items.Count - 1
        Dim line1 As Label
        If MAIN.RadioButton1.Checked = True Then
            MAIN.TreeView1.SelectedNode.Nodes(2).Nodes(10).BackColor = Color.Yellow
            MAIN.Button6.Text = "RUNNING"
line1:
            Try
                MAIN.ListBox2.Items.Clear()
                If SETUP.CheckBox6.Checked = True Then
                    MAIN.ListBox1.Items.Add($"{DateTime.Now.ToString} AutoImageComparison Enabled")
                    MAIN.ListBox1.SelectedIndex = MAIN.ListBox1.Items.Count - 1
                    FUNNCTIONLIB.RestoreApp(CONFIG.TextBox57.Text)
                    System.Threading.Thread.Sleep(3000)
                    CapturedImg = FUNNCTIONLIB.TakeSnapofActiveApp()
                    Reference = Image.FromFile(CONFIG.TextBox71.Text.ToString)
                    System.Threading.Thread.Sleep(3000)
                    MAIN.ListBox1.Items.Add($"{DateTime.Now.ToString} Loading Coordinates from Confriguration File")
                    MAIN.ListBox1.SelectedIndex = MAIN.ListBox1.Items.Count - 1
                    ini.Load("C:\PLUG_POWER_FCT\PLUG_POWER_FCT\Project Files\Confrigurations.ini")
                    MAIN.ListBox2.Items.Add(ini.Sections("SMC_SW3_ON").Keys("Coordinate1").Value)
                    MAIN.ListBox1.Items.Add($"{DateTime.Now.ToString} {ini.Sections("SMC_SW3_ON").Keys("Coordinate1").Value}")
                    MAIN.ListBox1.SelectedIndex = MAIN.ListBox1.Items.Count - 1
                    SMC_IMAGECOMP.PictureBox1.Image = CapturedImg
                    SMC_IMAGECOMP.PictureBox2.Image = Reference
                    SMC_IMAGECOMP.TextBox3.Text = "SW3 ON"
                    MAIN.ListBox1.Items.Add($"{DateTime.Now.ToString} Confriguration Loaded successfully")
                    MAIN.ListBox1.SelectedIndex = MAIN.ListBox1.Items.Count - 1
                    MAIN.ListBox1.Items.Add($"{DateTime.Now.ToString} Starting Comparison Function of Coordintaes")
                    MAIN.ListBox1.SelectedIndex = MAIN.ListBox1.Items.Count - 1
                    SMC_IMAGECOMP.ShowDialog()
                    If MAIN.Button6.Text = "PASS" Then
                        IMAGE_STR.PictureBox12.Image = CapturedImg
                        MAIN.PictureBox3.Image = Image.FromFile(CONFIG.TextBox72.Text)
                        MAIN.TextBox173.Text = "TURN OFF ALL THE CMD SWITCHES AND CLICK ON CONTINUE"
                        MAIN.PictureBox3.Visible = True
                        MAIN.TextBox173.Visible = True
                        MAIN.TextBox173.Enabled = False
                        MAIN.Button58.Visible = True
                    Else
                        IMAGE_STR.PictureBox12.Image = CapturedImg
                        MAIN.TimerFunction("STOP")
                        UUT_RESULT.RadioButton2.Checked = True
                        UUT_RESULT.TextBox2.Text = "SW3_ON"
                        UUT_RESULT.ShowDialog()
                    End If
                ElseIf SETUP.CheckBox6.Checked = False Then
                    MAIN.Button84.Visible = True
                    MAIN.Button85.Visible = True
                    MAIN.TextBox173.Text = "CONFIRM THE SCREEN AS SHOWN AND CLICK ON CONFIRM OTHERWISE CLICK ON NON-CONFIRM"
                    MAIN.TextBox173.Enabled = False
                    MAIN.TextBox173.Visible = True
                    MAIN.PictureBox3.Image = Image.FromFile(CONFIG.TextBox70.Text)
                    MAIN.PictureBox3.Visible = True
                End If
            Catch ex As Exception
                Dim iret As Object = MsgBox("Error Executing SMC.SMCSwitchTestNode10 : " & ex.Message, vbCritical + vbRetryCancel, "Error executing SMC.SMCSwitchTestNode10")
                If iret = vbRetry Then
                    GoTo line1
                ElseIf iret = vbCancel Then
                    MAIN.TimerFunction("STOP")
                    MAIN.Button6.Text = "ERROR"
                    MAIN.TreeView1.SelectedNode.Nodes(2).Nodes(10).BackColor = Color.Red
                    MAIN.Button43.Text = "ERROR"
                    MAIN.TextBox171.Text = "ERROR"
                    MAIN.ListBox1.Items.Add($"{DateTime.Now.ToString} Error Executing SMC.SMCSwitchTestNode10 :" & ex.Message)
                    MAIN.ListBox1.SelectedIndex = MAIN.ListBox1.Items.Count - 1
                    IMAGE_STR.PictureBox12.Image = CapturedImg
                    UUT_RESULT.RadioButton3.Checked = True
                    UUT_RESULT.TextBox2.Text = "SW3_ON"
                    UUT_RESULT.ShowDialog()
                End If
            End Try
        End If
    End Function
    Function SMCSwitchTestNode11() 'TOOGLE CMD ON
        FUNNCTIONLIB.CalculateProgress(34, 29)
        MAIN.ListBox1.Items.Add($"{DateTime.Now.ToString} Started Running  TOOGLE_CMD_ON Step")
        MAIN.ListBox1.SelectedIndex = MAIN.ListBox1.Items.Count - 1
        Dim line1 As Label
        If MAIN.RadioButton1.Checked = True Then
            MAIN.TreeView1.SelectedNode.Nodes(2).Nodes(11).BackColor = Color.Yellow
            MAIN.Button34.Text = "RUNNING"
line1:
            Try
                MAIN.ListBox2.Items.Clear()
                If SETUP.CheckBox6.Checked = True Then
                    MAIN.ListBox1.Items.Add($"{DateTime.Now.ToString} AutoImageComparison Enabled")
                    MAIN.ListBox1.SelectedIndex = MAIN.ListBox1.Items.Count - 1
                    FUNNCTIONLIB.RestoreApp(CONFIG.TextBox57.Text)
                    System.Threading.Thread.Sleep(3000)
                    CapturedImg = FUNNCTIONLIB.TakeSnapofActiveApp()
                    Reference = Image.FromFile(CONFIG.TextBox73.Text.ToString)
                    System.Threading.Thread.Sleep(3000)
                    MAIN.ListBox1.Items.Add($"{DateTime.Now.ToString} Loading Coordinates from Confriguration File")
                    MAIN.ListBox1.SelectedIndex = MAIN.ListBox1.Items.Count - 1
                    ini.Load("C:\PLUG_POWER_FCT\PLUG_POWER_FCT\Project Files\Confrigurations.ini")
                    MAIN.ListBox2.Items.Add(ini.Sections("SMC_TOOGLECMD_ON").Keys("Coordinate1").Value)
                    MAIN.ListBox1.Items.Add($"{DateTime.Now.ToString} {ini.Sections("SMC_TOOGLECMD_ON").Keys("Coordinate1").Value}")
                    MAIN.ListBox1.SelectedIndex = MAIN.ListBox1.Items.Count - 1
                    SMC_IMAGECOMP.PictureBox1.Image = CapturedImg
                    SMC_IMAGECOMP.PictureBox1.Image = CapturedImg
                    SMC_IMAGECOMP.PictureBox2.Image = Reference
                    SMC_IMAGECOMP.TextBox3.Text = "TOOGLECMD ON"
                    MAIN.ListBox1.Items.Add($"{DateTime.Now.ToString} Confriguration Loaded successfully")
                    MAIN.ListBox1.SelectedIndex = MAIN.ListBox1.Items.Count - 1
                    MAIN.ListBox1.Items.Add($"{DateTime.Now.ToString} Starting Comparison Function of Coordintaes")
                    MAIN.ListBox1.SelectedIndex = MAIN.ListBox1.Items.Count - 1
                    SMC_IMAGECOMP.ShowDialog()
                    If MAIN.Button34.Text = "PASS" Then
                        IMAGE_STR.PictureBox13.Image = CapturedImg
                        MAIN.PictureBox3.Image = Image.FromFile(CONFIG.TextBox74.Text)
                        MAIN.TextBox173.Text = "TURN ON CAN1_CMD SILENCE THEN CLICK ON CONTINUE"
                        MAIN.PictureBox3.Visible = True
                        MAIN.TextBox173.Visible = True
                        MAIN.TextBox173.Enabled = False
                        MAIN.Button59.Visible = True
                    Else
                        IMAGE_STR.PictureBox13.Image = CapturedImg
                        MAIN.TimerFunction("STOP")
                        UUT_RESULT.RadioButton2.Checked = True
                        UUT_RESULT.TextBox2.Text = "TOOGLECMD_ON"
                        UUT_RESULT.ShowDialog()
                    End If
                ElseIf SETUP.CheckBox6.Checked = False Then
                    MAIN.Button86.Visible = True
                    MAIN.Button87.Visible = True
                    MAIN.TextBox173.Text = "CONFIRM THE SCREEN AS SHOWN AND CLICK ON CONFIRM OTHERWISE CLICK ON NON-CONFIRM"
                    MAIN.TextBox173.Enabled = False
                    MAIN.TextBox173.Visible = True
                    MAIN.PictureBox3.Image = Image.FromFile(CONFIG.TextBox72.Text)
                    MAIN.PictureBox3.Visible = True
                End If
            Catch ex As Exception
                Dim iret As Object = MsgBox("Error Executing SMC.SMCSwitchTestNode11 : " & ex.Message, vbCritical + vbRetryCancel, "Error executing SMC.SMCSwitchTestNode11")
                If iret = vbRetry Then
                    GoTo line1
                ElseIf iret = vbCancel Then
                    MAIN.TimerFunction("STOP")
                    MAIN.Button34.Text = "ERROR"
                    MAIN.TreeView1.SelectedNode.Nodes(2).Nodes(11).BackColor = Color.Red
                    MAIN.Button43.Text = "ERROR"
                    MAIN.TextBox115.Text = "ERROR"
                    MAIN.ListBox1.Items.Add($"{DateTime.Now.ToString} Error Executing SMC.SMCSwitchTestNode11 :" & ex.Message)
                    MAIN.ListBox1.SelectedIndex = MAIN.ListBox1.Items.Count - 1
                    IMAGE_STR.PictureBox13.Image = CapturedImg
                    UUT_RESULT.RadioButton3.Checked = True
                    UUT_RESULT.TextBox2.Text = "TOOGLECMD_ON"
                    UUT_RESULT.ShowDialog()
                End If
            End Try
        End If
    End Function
    Function SMCSwitchTestNode12() 'CAN1 SILENCE ON
        FUNNCTIONLIB.CalculateProgress(34, 30)
        MAIN.ListBox1.Items.Add($"{DateTime.Now.ToString} Started Running  CAN1_SILENCE_CMD_ON Step")
        MAIN.ListBox1.SelectedIndex = MAIN.ListBox1.Items.Count - 1
        Dim line1 As Label
        If MAIN.RadioButton1.Checked = True Then
            MAIN.TreeView1.SelectedNode.Nodes(2).Nodes(12).BackColor = Color.Yellow
            MAIN.Button35.Text = "RUNNING"
line1:
            Try
                MAIN.ListBox2.Items.Clear()
                If SETUP.CheckBox6.Checked = True Then
                    MAIN.ListBox1.Items.Add($"{DateTime.Now.ToString} AutoImageComparison Enabled")
                    MAIN.ListBox1.SelectedIndex = MAIN.ListBox1.Items.Count - 1
                    FUNNCTIONLIB.RestoreApp(CONFIG.TextBox57.Text)
                    System.Threading.Thread.Sleep(3000)
                    CapturedImg = FUNNCTIONLIB.TakeSnapofActiveApp()
                    Reference = Image.FromFile(CONFIG.TextBox75.Text.ToString)
                    System.Threading.Thread.Sleep(3000)
                    MAIN.ListBox1.Items.Add($"{DateTime.Now.ToString} Loading Coordinates from Confriguration File")
                    MAIN.ListBox1.SelectedIndex = MAIN.ListBox1.Items.Count - 1
                    ini.Load("C:\PLUG_POWER_FCT\PLUG_POWER_FCT\Project Files\Confrigurations.ini")
                    MAIN.ListBox2.Items.Add(ini.Sections("SMC_CAN1SMC_ON").Keys("Coordinate1").Value)
                    MAIN.ListBox1.Items.Add($"{DateTime.Now.ToString} {ini.Sections("SMC_CAN1SMC_ON").Keys("Coordinate1").Value}")
                    MAIN.ListBox1.SelectedIndex = MAIN.ListBox1.Items.Count - 1
                    SMC_IMAGECOMP.PictureBox1.Image = CapturedImg
                    SMC_IMAGECOMP.PictureBox2.Image = Reference
                    SMC_IMAGECOMP.TextBox3.Text = "CAN1SILENCE ON"
                    MAIN.ListBox1.Items.Add($"{DateTime.Now.ToString} Confriguration Loaded successfully")
                    MAIN.ListBox1.SelectedIndex = MAIN.ListBox1.Items.Count - 1
                    MAIN.ListBox1.Items.Add($"{DateTime.Now.ToString} Starting Comparison Function of Coordintaes")
                    MAIN.ListBox1.SelectedIndex = MAIN.ListBox1.Items.Count - 1
                    SMC_IMAGECOMP.ShowDialog()
                    If MAIN.Button35.Text = "PASS" Then
                        IMAGE_STR.PictureBox14.Image = CapturedImg
                        MAIN.PictureBox3.Image = Image.FromFile(CONFIG.TextBox76.Text)
                        MAIN.TextBox173.Text = "TURN ON CAN2_CMD SILENCE THEN CLICK ON CONTINUE"
                        MAIN.PictureBox3.Visible = True
                        MAIN.TextBox173.Visible = True
                        MAIN.TextBox173.Enabled = False
                        MAIN.Button60.Visible = True
                    Else
                        IMAGE_STR.PictureBox14.Image = CapturedImg
                        MAIN.TimerFunction("STOP")
                        UUT_RESULT.RadioButton2.Checked = True
                        UUT_RESULT.TextBox2.Text = "CAN1SILENCE_ON"
                        UUT_RESULT.ShowDialog()
                    End If
                ElseIf SETUP.CheckBox6.Checked = False Then
                    MAIN.Button88.Visible = True
                    MAIN.Button89.Visible = True
                    MAIN.TextBox173.Text = "CONFIRM THE SCREEN AS SHOWN AND CLICK ON CONFIRM OTHERWISE CLICK ON NON-CONFIRM"
                    MAIN.TextBox173.Enabled = False
                    MAIN.TextBox173.Visible = True
                    MAIN.PictureBox3.Image = Image.FromFile(CONFIG.TextBox74.Text)
                    MAIN.PictureBox3.Visible = True
                End If
            Catch ex As Exception
                Dim iret As Object = MsgBox("Error Executing SMC.SMCSwitchTestNode12 : " & ex.Message, vbCritical + vbRetryCancel, "Error executing SMC.SMCSwitchTestNode12")
                If iret = vbRetry Then
                    GoTo line1
                ElseIf iret = vbCancel Then
                    MAIN.TimerFunction("STOP")
                    MAIN.Button35.Text = "ERROR"
                    MAIN.TreeView1.SelectedNode.Nodes(2).Nodes(11).BackColor = Color.Red
                    MAIN.Button43.Text = "ERROR"
                    MAIN.TextBox119.Text = "ERROR"
                    MAIN.ListBox1.Items.Add($"{DateTime.Now.ToString} Error Executing SMC.SMCSwitchTestNode12 :" & ex.Message)
                    MAIN.ListBox1.SelectedIndex = MAIN.ListBox1.Items.Count - 1
                    IMAGE_STR.PictureBox14.Image = CapturedImg
                    UUT_RESULT.RadioButton3.Checked = True
                    UUT_RESULT.TextBox2.Text = "CAN1SILENCE_ON"
                    UUT_RESULT.ShowDialog()
                End If
            End Try
        End If
    End Function
    Function SMCSwitchTestNode13() 'CAN2 SILENCE ON
        FUNNCTIONLIB.CalculateProgress(34, 31)
        MAIN.ListBox1.Items.Add($"{DateTime.Now.ToString} Started Running  CAN2_SILENCE_CMD_ON Step")
        MAIN.ListBox1.SelectedIndex = MAIN.ListBox1.Items.Count - 1
        Dim line1 As Label
        If MAIN.RadioButton1.Checked = True Then
            MAIN.TreeView1.SelectedNode.Nodes(2).Nodes(13).BackColor = Color.Yellow
            MAIN.Button36.Text = "RUNNING"
line1:
            Try
                MAIN.ListBox2.Items.Clear()
                If SETUP.CheckBox6.Checked = True Then
                    MAIN.ListBox1.Items.Add($"{DateTime.Now.ToString} AutoImageComparison Enabled")
                    MAIN.ListBox1.SelectedIndex = MAIN.ListBox1.Items.Count - 1
                    FUNNCTIONLIB.RestoreApp(CONFIG.TextBox57.Text)
                    System.Threading.Thread.Sleep(3000)
                    CapturedImg = FUNNCTIONLIB.TakeSnapofActiveApp()
                    Reference = Image.FromFile(CONFIG.TextBox76.Text.ToString)
                    System.Threading.Thread.Sleep(3000)
                    MAIN.ListBox1.Items.Add($"{DateTime.Now.ToString} Loading Coordinates from Confriguration File")
                    MAIN.ListBox1.SelectedIndex = MAIN.ListBox1.Items.Count - 1
                    ini.Load("C:\PLUG_POWER_FCT\PLUG_POWER_FCT\Project Files\Confrigurations.ini")
                    MAIN.ListBox2.Items.Add(ini.Sections("SMC_CAN2SMC_ON").Keys("Coordinate1").Value)
                    MAIN.ListBox1.Items.Add($"{DateTime.Now.ToString} {ini.Sections("SMC_CAN2SMC_ON").Keys("Coordinate1").Value}")
                    MAIN.ListBox1.SelectedIndex = MAIN.ListBox1.Items.Count - 1
                    SMC_IMAGECOMP.PictureBox1.Image = CapturedImg
                    SMC_IMAGECOMP.PictureBox2.Image = Reference
                    SMC_IMAGECOMP.TextBox3.Text = "CAN2SILENCE ON"
                    MAIN.ListBox1.Items.Add($"{DateTime.Now.ToString} Confriguration Loaded successfully")
                    MAIN.ListBox1.SelectedIndex = MAIN.ListBox1.Items.Count - 1
                    MAIN.ListBox1.Items.Add($"{DateTime.Now.ToString} Starting Comparison Function of Coordintaes")
                    MAIN.ListBox1.SelectedIndex = MAIN.ListBox1.Items.Count - 1
                    SMC_IMAGECOMP.ShowDialog()
                    If MAIN.Button36.Text = "PASS" Then
                        IMAGE_STR.PictureBox15.Image = CapturedImg
                        MAIN.PictureBox3.Image = Image.FromFile(CONFIG.TextBox78.Text)
                        MAIN.TextBox173.Text = "CLCIK ON THE SET TIME AS SHOWN"
                        MAIN.PictureBox3.Visible = True
                        MAIN.TextBox173.Visible = True
                        MAIN.TextBox173.Enabled = False
                        MAIN.Button61.Visible = True
                    Else
                        IMAGE_STR.PictureBox15.Image = CapturedImg
                        MAIN.TimerFunction("STOP")
                        UUT_RESULT.RadioButton2.Checked = True
                        UUT_RESULT.TextBox2.Text = "CAN2SILENCE_ON"
                        UUT_RESULT.ShowDialog()
                    End If
                ElseIf SETUP.CheckBox6.Checked = False Then
                    MAIN.Button90.Visible = True
                    MAIN.Button91.Visible = True
                    MAIN.TextBox173.Text = "CONFIRM THE SCREEN AS SHOWN AND CLICK ON CONFIRM OTHERWISE CLICK ON NON-CONFIRM"
                    MAIN.TextBox173.Enabled = False
                    MAIN.TextBox173.Visible = True
                    MAIN.PictureBox3.Image = Image.FromFile(CONFIG.TextBox76.Text)
                    MAIN.PictureBox3.Visible = True
                End If
            Catch ex As Exception
                Dim iret As Object = MsgBox("Error Executing SMC.SMCSwitchTestNode13 : " & ex.Message, vbCritical + vbRetryCancel, "Error executing SMC.SMCSwitchTestNode13")
                If iret = vbRetry Then
                    GoTo line1
                ElseIf iret = vbCancel Then
                    MAIN.TimerFunction("STOP")
                    MAIN.Button36.Text = "ERROR"
                    MAIN.TreeView1.SelectedNode.Nodes(2).Nodes(11).BackColor = Color.Red
                    MAIN.Button43.Text = "ERROR"
                    MAIN.TextBox123.Text = "ERROR"
                    MAIN.ListBox1.Items.Add($"{DateTime.Now.ToString} Error Executing SMC.SMCSwitchTestNode13 :" & ex.Message)
                    MAIN.ListBox1.SelectedIndex = MAIN.ListBox1.Items.Count - 1
                    IMAGE_STR.PictureBox15.Image = CapturedImg
                    UUT_RESULT.RadioButton3.Checked = True
                    UUT_RESULT.TextBox2.Text = "CAN2SILENCE_ON"
                    UUT_RESULT.ShowDialog()
                End If
            End Try
        End If
    End Function
    Function SMCTimeNode0() 'TIME TEST
        FUNNCTIONLIB.CalculateProgress(34, 32)
        MAIN.ListBox1.Items.Add($"{DateTime.Now.ToString} Started Running  SET_TIME_TEST Step")
        MAIN.ListBox1.SelectedIndex = MAIN.ListBox1.Items.Count - 1
        Dim line1 As Label
        If MAIN.RadioButton1.Checked = True Then
            MAIN.TreeView1.SelectedNode.Nodes(2).Collapse()
            MAIN.TreeView1.SelectedNode.Nodes(3).BackColor = Color.Yellow
            MAIN.TabControl2.SelectedTab = MAIN.TabPage10
            MAIN.Button37.Text = "RUNNING"
line1:
            Try
                MAIN.ListBox2.Items.Clear()
                If SETUP.CheckBox6.Checked = True Then
                    MAIN.ListBox1.Items.Add($"{DateTime.Now.ToString} AutoImageComparison Enabled")
                    MAIN.ListBox1.SelectedIndex = MAIN.ListBox1.Items.Count - 1
                    FUNNCTIONLIB.RestoreApp(CONFIG.TextBox57.Text)
                    System.Threading.Thread.Sleep(3000)
                    CapturedImg = FUNNCTIONLIB.TakeSnapofActiveApp()
                    System.Threading.Thread.Sleep(3000)
                    SMC_TIMETEST.PictureBox1.Image = CapturedImg
                    MAIN.ListBox1.Items.Add($"{DateTime.Now.ToString} Loading Coordinates from Confriguration File")
                    MAIN.ListBox1.SelectedIndex = MAIN.ListBox1.Items.Count - 1
                    ini.Load("C:\PLUG_POWER_FCT\PLUG_POWER_FCT\Project Files\Confrigurations.ini")
                    MAIN.ListBox2.Items.Add(ini.Sections("SMC_TIMETEST").Keys("Coordinate1").Value)
                    MAIN.ListBox1.Items.Add($"{DateTime.Now.ToString} {ini.Sections("SMC_TIMETEST").Keys("Coordinate1").Value}")
                    MAIN.ListBox1.SelectedIndex = MAIN.ListBox1.Items.Count - 1
                    MAIN.ListBox1.Items.Add($"{DateTime.Now.ToString} Confriguration Loaded successfully")
                    MAIN.ListBox1.SelectedIndex = MAIN.ListBox1.Items.Count - 1
                    MAIN.ListBox1.Items.Add($"{DateTime.Now.ToString} Starting Comparison Function of Coordintaes")
                    MAIN.ListBox1.SelectedIndex = MAIN.ListBox1.Items.Count - 1
                    SMC_TIMETEST.ShowDialog()
                    If MAIN.Button37.Text = "PASS" Then
                        IMAGE_STR.PictureBox16.Image = CapturedImg
                        MAIN.TreeView1.SelectedNode.Nodes(3).BackColor = Color.Green
                        MAIN.PictureBox3.Image = Image.FromFile(CONFIG.TextBox79.Text)
                        MAIN.TextBox173.Text = "CLCIK ON THE STEPPER ON AS SHOWN"
                        MAIN.PictureBox3.Visible = True
                        MAIN.TextBox173.Visible = True
                        MAIN.TextBox173.Enabled = False
                        MAIN.Button62.Visible = True
                    Else
                        IMAGE_STR.PictureBox16.Image = CapturedImg
                        MAIN.TreeView1.SelectedNode.Nodes(3).BackColor = Color.Red
                        MAIN.Button43.Text = "FAIL"
                        MAIN.TimerFunction("STOP")
                        UUT_RESULT.RadioButton2.Checked = True
                        UUT_RESULT.TextBox2.Text = "TIME_TEST"
                        UUT_RESULT.ShowDialog()
                    End If
                Else
                    MAIN.Button92.Visible = True
                    MAIN.Button93.Visible = True
                    MAIN.TextBox173.Text = "CONFIRM THE SCREEN AS SHOWN AND CLICK ON CONFIRM OTHERWISE CLICK ON NON-CONFIRM"
                    MAIN.TextBox173.Enabled = False
                    MAIN.TextBox173.Visible = True
                    MAIN.PictureBox3.Image = Image.FromFile(CONFIG.TextBox78.Text)
                    MAIN.PictureBox3.Visible = True
                End If
            Catch ex As Exception
                Dim iret As Object = MsgBox("Error Executing SMC.SMCTimeNode0 : " & ex.Message, vbCritical + vbRetryCancel, "Error executing SMC.SMCTimeNode0")
                If iret = vbRetry Then
                    GoTo line1
                ElseIf iret = vbCancel Then
                    MAIN.TimerFunction("STOP")
                    MAIN.Button37.Text = "ERROR"
                    MAIN.TreeView1.SelectedNode.Nodes(2).Nodes(11).BackColor = Color.Red
                    MAIN.Button43.Text = "ERROR"
                    MAIN.TextBox126.Text = "ERROR"
                    MAIN.ListBox1.Items.Add($"{DateTime.Now.ToString} Error Executing SMC.SMCTimeNode0 :" & ex.Message)
                    MAIN.ListBox1.SelectedIndex = MAIN.ListBox1.Items.Count - 1
                    IMAGE_STR.PictureBox16.Image = CapturedImg
                    UUT_RESULT.RadioButton3.Checked = True
                    UUT_RESULT.TextBox2.Text = "TIME_TEST"
                    UUT_RESULT.ShowDialog()
                End If
            End Try
        End If
    End Function
    Function SMCStepperTestNode0()
        FUNNCTIONLIB.CalculateProgress(34, 33)
        MAIN.ListBox1.Items.Add($"{DateTime.Now.ToString} Started Running  STEPPER_TEST Step")
        MAIN.ListBox1.SelectedIndex = MAIN.ListBox1.Items.Count - 1
        Dim line1 As Label
        If MAIN.RadioButton1.Checked = True Then
            MAIN.TreeView1.SelectedNode.Nodes(4).BackColor = Color.Yellow
            MAIN.TabControl2.SelectedTab = MAIN.TabPage11
            MAIN.Button38.Text = "RUNNING"
line1:
            Try
                MAIN.ListBox2.Items.Clear()
                If SETUP.CheckBox6.Checked = True Then
                    MAIN.ListBox1.Items.Add($"{DateTime.Now.ToString} AutoImageComparison Enabled")
                    MAIN.ListBox1.SelectedIndex = MAIN.ListBox1.Items.Count - 1
                    FUNNCTIONLIB.RestoreApp(CONFIG.TextBox57.Text)
                    System.Threading.Thread.Sleep(3000)
                    CapturedImg = FUNNCTIONLIB.TakeSnapofActiveApp()
                    Reference = Image.FromFile(CONFIG.TextBox80.Text.ToString)
                    System.Threading.Thread.Sleep(3000)
                    MAIN.ListBox1.Items.Add($"{DateTime.Now.ToString} Loading Coordinates from Confriguration File")
                    MAIN.ListBox1.SelectedIndex = MAIN.ListBox1.Items.Count - 1
                    ini.Load("C:\PLUG_POWER_FCT\PLUG_POWER_FCT\Project Files\Confrigurations.ini")
                    MAIN.ListBox2.Items.Add(ini.Sections("SMC_STEPPER_TEST").Keys("Coordinate1").Value)
                    MAIN.ListBox1.Items.Add($"{DateTime.Now.ToString} {ini.Sections("SMC_STEPPER_TEST").Keys("Coordinate1").Value}")
                    MAIN.ListBox1.SelectedIndex = MAIN.ListBox1.Items.Count - 1
                    SMC_IMAGECOMP.PictureBox1.Image = CapturedImg
                    SMC_IMAGECOMP.PictureBox2.Image = Reference
                    SMC_IMAGECOMP.TextBox3.Text = "STEPPER TEST"
                    MAIN.ListBox1.Items.Add($"{DateTime.Now.ToString} Confriguration Loaded successfully")
                    MAIN.ListBox1.SelectedIndex = MAIN.ListBox1.Items.Count - 1
                    MAIN.ListBox1.Items.Add($"{DateTime.Now.ToString} Starting Extrction Of DateTime From Captured Image")
                    MAIN.ListBox1.SelectedIndex = MAIN.ListBox1.Items.Count - 1
                    SMC_IMAGECOMP.ShowDialog()
                    If MAIN.Button38.Text = "PASS" Then
                        IMAGE_STR.PictureBox17.Image = CapturedImg
                        MAIN.PictureBox3.Image = Image.FromFile(CONFIG.TextBox87.Text)
                        MAIN.TextBox173.Text = "CLCIK ON THE SHUT DOWN AS SHOWN"
                        MAIN.PictureBox3.Visible = True
                        MAIN.TextBox173.Visible = True
                        MAIN.TextBox173.Enabled = False
                        MAIN.Button63.Visible = True
                    Else
                        IMAGE_STR.PictureBox17.Image = CapturedImg
                        MAIN.TimerFunction("STOP")
                        UUT_RESULT.RadioButton2.Checked = True
                        UUT_RESULT.TextBox2.Text = "STEPPER_TEST"
                        UUT_RESULT.ShowDialog()
                    End If
                ElseIf SETUP.CheckBox6.Checked = False Then
                    MAIN.Button94.Visible = True
                    MAIN.Button95.Visible = True
                    MAIN.TextBox173.Text = "CONFIRM THE SCREEN AS SHOWN AND CLICK ON CONFIRM OTHERWISE CLICK ON NON-CONFIRM"
                    MAIN.TextBox173.Enabled = False
                    MAIN.TextBox173.Visible = True
                    MAIN.PictureBox3.Image = Image.FromFile(CONFIG.TextBox79.Text)
                    MAIN.PictureBox3.Visible = True
                End If
            Catch ex As Exception
                Dim iret As Object = MsgBox("Error Executing SMC.SMCStepperTestNode0 : " & ex.Message, vbCritical + vbRetryCancel, "Error executing SMC.SMCStepperTestNode0")
                If iret = vbRetry Then
                    GoTo line1
                ElseIf iret = vbCancel Then
                    MAIN.TimerFunction("STOP")
                    MAIN.Button38.Text = "ERROR"
                    MAIN.TreeView1.SelectedNode.Nodes(2).Nodes(11).BackColor = Color.Red
                    MAIN.Button43.Text = "ERROR"
                    MAIN.TextBox135.Text = "ERROR"
                    MAIN.ListBox1.Items.Add($"{DateTime.Now.ToString} Error Executing SMC.SMCStepperTestNode0 :" & ex.Message)
                    MAIN.ListBox1.SelectedIndex = MAIN.ListBox1.Items.Count - 1
                    IMAGE_STR.PictureBox17.Image = CapturedImg
                    UUT_RESULT.RadioButton3.Checked = True
                    UUT_RESULT.TextBox2.Text = "STEPPER_TEST"
                    UUT_RESULT.ShowDialog()
                End If
            End Try
        End If
    End Function
    Function SMCShutDownNod0()
        FUNNCTIONLIB.CalculateProgress(34, 34)
        MAIN.ListBox1.Items.Add($"{DateTime.Now.ToString} Started Running  SHUTDOWN_TEST Step")
        MAIN.ListBox1.SelectedIndex = MAIN.ListBox1.Items.Count - 1
        Dim line1 As Label
        MAIN.TabControl2.SelectedTab = MAIN.TabPage12
        If MAIN.RadioButton1.Checked = True Then
            MAIN.TreeView1.SelectedNode.Nodes(5).BackColor = Color.Yellow
line1:
            Try
                MAIN.ListBox2.Items.Clear()
                MAIN.Button39.Text = "RUNNING"
                If SETUP.CheckBox6.Checked = True Then
                    MAIN.ListBox1.Items.Add($"{DateTime.Now.ToString} AutoImageComparison Enabled")
                    MAIN.ListBox1.SelectedIndex = MAIN.ListBox1.Items.Count - 1
                    System.Threading.Thread.Sleep(3000)
                    FUNNCTIONLIB.RestoreApp(CONFIG.TextBox57.Text)
                    System.Threading.Thread.Sleep(3000)
                    CapturedImg = FUNNCTIONLIB.TakeSnapofActiveApp()
                    Reference = Image.FromFile(CONFIG.TextBox86.Text)
                    If FUNNCTIONLIB.CompareImages(CapturedImg, Reference) = True Then
                        IMAGE_STR.PictureBox18.Image = CapturedImg
                        MAIN.TextBox144.Text = "TRUE"
                        MAIN.Button39.Text = "PASS"
                        MAIN.Button43.Text = "PASS"
                        MAIN.TreeView1.SelectedNode.Nodes(5).BackColor = Color.Green
                        MAIN.TimerFunction("STOP")
                        UUT_RESULT.RadioButton1.Checked = True
                        UUT_RESULT.TextBox2.Text = "SHUTDOWN_TEST"
                        UUT_RESULT.ShowDialog()
                    Else
                        IMAGE_STR.PictureBox18.Image = CapturedImg
                        MAIN.TextBox144.Text = "FALSE"
                        MAIN.Button39.Text = "FAIL"
                        MAIN.Button43.Text = "FAIL"
                        MAIN.TreeView1.SelectedNode.Nodes(5).BackColor = Color.Red
                        MAIN.TimerFunction("STOP")
                        UUT_RESULT.RadioButton2.Checked = True
                        UUT_RESULT.TextBox2.Text = "SHUTDOWN_TEST"
                        UUT_RESULT.ShowDialog()
                    End If
                ElseIf SETUP.CheckBox6.Checked = False Then
                    MAIN.Button96.Visible = True
                    MAIN.Button97.Visible = True
                    MAIN.TextBox173.Text = "CONFIRM THE SCREEN AS SHOWN AND CLICK ON CONFIRM OTHERWISE CLICK ON NON-CONFIRM"
                    MAIN.TextBox173.Enabled = False
                    MAIN.TextBox173.Visible = True
                    MAIN.PictureBox3.Image = Image.FromFile(CONFIG.TextBox87.Text)
                    MAIN.PictureBox3.Visible = True
                End If
            Catch ex As Exception
                Dim iret As Object = MsgBox("Error Executing SMC.SMCShutDownNod0 : " & ex.Message, vbCritical + vbRetryCancel, "Error executing SMC.SMCShutDownNod0")
                If iret = vbRetry Then
                    GoTo line1
                ElseIf iret = vbCancel Then
                    MAIN.TimerFunction("STOP")
                    MAIN.Button39.Text = "ERROR"
                    MAIN.TreeView1.SelectedNode.Nodes(5).BackColor = Color.Red
                    MAIN.Button43.Text = "ERROR"
                    MAIN.TextBox144.Text = "ERROR"
                    MAIN.ListBox1.Items.Add($"{DateTime.Now.ToString} Error Executing SMC.SMCShutDownNod0 :" & ex.Message)
                    MAIN.ListBox1.SelectedIndex = MAIN.ListBox1.Items.Count - 1
                    IMAGE_STR.PictureBox18.Image = CapturedImg
                    UUT_RESULT.RadioButton3.Checked = True
                    UUT_RESULT.TextBox2.Text = "SHUTDOWN_TEST"
                    UUT_RESULT.ShowDialog()
                End If
            End Try
        End If
    End Function
    Function GenerateReportOnerrorInit(User As String, TestMode As String, SerialNumber As String, Status As String, CycleTime As String, TesterID As String, Mark As String, ErrorMsg As String)
        Dim line1 As Label
line1:
        Dim path1 As String = CONFIG.TextBox4.Text
        file1 = My.Computer.FileSystem.OpenTextFileWriter(path1, True)
        Try
            Dim xlapp = New Microsoft.Office.Interop.Excel.Application
            xlapp.DisplayAlerts = False
            Dim workbook = xlapp.Workbooks.Open(CONFIG.TextBox82.Text.ToString)
            Dim worksheet As Excel.Worksheet
            worksheet = workbook.Sheets("SMC")
            worksheet.Range("D6").Value = User
            worksheet.Range("D7").Value = TestMode
            worksheet.Range("D8").Value = SerialNumber
            worksheet.Range("D9").Value = DateTime.Now.ToString
            worksheet.Range("D10").Value = Status
            worksheet.Range("F6").Value = CycleTime
            worksheet.Range("F7").Value = TesterID
            worksheet.Range("F8").Value = Mark
            worksheet.Range("G7").Value = ErrorMsg
            workbook.ExportAsFixedFormat(Excel.XlFixedFormatType.xlTypePDF, CONFIG.TextBox83.Text.ToString & SerialNumber & "_" & TestMode & "_" & Status & Format(Now(), "hhmmssmmddyy") & ".pdf", vbNormal)
            workbook.Close(True)
            System.Runtime.InteropServices.Marshal.ReleaseComObject(workbook)
            workbook = Nothing
            xlapp.Quit()
            System.Runtime.InteropServices.Marshal.ReleaseComObject(xlapp)
            xlapp = Nothing
            Process.GetProcessesByName("EXCEL")(0).Kill()
            file1.Close()
            MAIN.ListBox1.Items.Add($"{DateTime.Now.ToString} Test Report Generated success : " & CONFIG.TextBox83.Text.ToString & SerialNumber & "_" & TestMode & "_" & Status & Format(Now(), "hhmmssmmddyy"))
            MAIN.ListBox1.SelectedIndex = MAIN.ListBox1.Items.Count - 1
            Dim SW As IO.StreamWriter = IO.File.CreateText(CONFIG.TextBox83.Text.ToString & SerialNumber)
            For Each S As String In MAIN.ListBox1.Items
                SW.WriteLine(S, True)
            Next
            SW.Close()
            file1.Close()

        Catch ex As Exception
            Dim iret As Object = MsgBox("Error generating Report : " & ex.Message, vbCritical + vbRetryCancel, "Error executing SMC.GenerateReportOnerrorInit")
            If iret = vbRetry Then
                file1.Close()
                GoTo line1
            ElseIf iret = vbCancel Then
                file1.WriteLine("Error generating Report :  " & ex.Message & " " & "@" & DateTime.Now.ToString("yyyy/MM/dd/HH:mm:ss"))
                file1.Close()
                MAIN.ListBox1.Items.Add($"{DateTime.Now.ToString} Error generating Report :" & ex.Message)
                MAIN.ListBox1.SelectedIndex = MAIN.ListBox1.Items.Count - 1
                Exit Function
            End If
        End Try
    End Function
    Function GenerateReportOnerrorSMCProgrammingNode1(User As String, TestMode As String, SerialNumber As String, Status As String, CycleTime As String, TesterID As String, Mark As String, ErrorMsg As String, DutConnect As String, SetPower As String)
        Dim line1 As Label
line1:
        Dim path1 As String = CONFIG.TextBox4.Text
        file1 = My.Computer.FileSystem.OpenTextFileWriter(path1, True)
        Try
            Dim xlapp = New Microsoft.Office.Interop.Excel.Application
            xlapp.DisplayAlerts = False
            Dim workbook = xlapp.Workbooks.Open(CONFIG.TextBox82.Text.ToString)
            Dim worksheet As Excel.Worksheet
            worksheet = workbook.Sheets("SMC")
            worksheet.Range("D6").Value = User
            worksheet.Range("D7").Value = TestMode
            worksheet.Range("D8").Value = SerialNumber
            worksheet.Range("D9").Value = DateTime.Now.ToString
            worksheet.Range("D10").Value = Status
            worksheet.Range("F6").Value = CycleTime
            worksheet.Range("F7").Value = TesterID
            worksheet.Range("F8").Value = Mark
            worksheet.Range("G7").Value = ErrorMsg
            worksheet.Range("G13").Value = DutConnect
            worksheet.Range("G14").Value = SetPower
            workbook.ExportAsFixedFormat(Excel.XlFixedFormatType.xlTypePDF, CONFIG.TextBox83.Text.ToString & SerialNumber & "_" & TestMode & "_" & Status & Format(Now(), "hhmmssmmddyy") & ".pdf", vbNormal)
            workbook.Close(True)
            System.Runtime.InteropServices.Marshal.ReleaseComObject(workbook)
            workbook = Nothing
            xlapp.Quit()
            System.Runtime.InteropServices.Marshal.ReleaseComObject(xlapp)
            xlapp = Nothing
            Process.GetProcessesByName("EXCEL")(0).Kill()
            file1.Close()
            MAIN.ListBox1.Items.Add($"{DateTime.Now.ToString} Test Report Generated success : " & CONFIG.TextBox83.Text.ToString & SerialNumber & "_" & TestMode & "_" & Status & Format(Now(), "hhmmssmmddyy"))
            MAIN.ListBox1.SelectedIndex = MAIN.ListBox1.Items.Count - 1
            Dim SW As IO.StreamWriter = IO.File.CreateText(CONFIG.TextBox83.Text.ToString & SerialNumber)
            For Each S As String In MAIN.ListBox1.Items
                SW.WriteLine(S, True)
            Next
            SW.Close()
            file1.Close()
        Catch ex As Exception
            Dim iret As Object = MsgBox("Error generating Report : " & ex.Message, vbCritical + vbRetryCancel, "Error executing SMC.GenerateReportOnerrorInit")
            If iret = vbRetry Then
                file1.Close()
                GoTo line1
            ElseIf iret = vbCancel Then
                file1.WriteLine("Error generating Report :  " & ex.Message & " " & "@" & DateTime.Now.ToString("yyyy/MM/dd/HH:mm:ss"))
                file1.Close()
                MAIN.ListBox1.Items.Add($"{DateTime.Now.ToString} Error generating Report :" & ex.Message)
                MAIN.ListBox1.SelectedIndex = MAIN.ListBox1.Items.Count - 1
                Exit Function
            End If
        End Try
    End Function
    Function GenerateReportOnErrorUniflashProg(User As String, TestMode As String, SerialNumber As String, Status As String, CycleTime As String, TesterID As String, Mark As String, ErrorMsg As String, DutConnect As String, SetPower As String, UniflashProgramme As String)
        Dim line1 As Label
line1:
        Dim path1 As String = CONFIG.TextBox4.Text
        file1 = My.Computer.FileSystem.OpenTextFileWriter(path1, True)
        Try
            Dim xlapp = New Microsoft.Office.Interop.Excel.Application
            xlapp.DisplayAlerts = False
            Dim workbook = xlapp.Workbooks.Open(CONFIG.TextBox82.Text.ToString)
            Dim worksheet As Excel.Worksheet
            worksheet = workbook.Sheets("SMC")
            worksheet.Range("D6").Value = User
            worksheet.Range("D7").Value = TestMode
            worksheet.Range("D8").Value = SerialNumber
            worksheet.Range("D9").Value = DateTime.Now.ToString
            worksheet.Range("D10").Value = Status
            worksheet.Range("F6").Value = CycleTime
            worksheet.Range("F7").Value = TesterID
            worksheet.Range("F8").Value = Mark
            worksheet.Range("G7").Value = ErrorMsg
            worksheet.Range("G13").Value = DutConnect
            worksheet.Range("G14").Value = SetPower
            worksheet.Range("G15").Value = UniflashProgramme
            workbook.ExportAsFixedFormat(Excel.XlFixedFormatType.xlTypePDF, CONFIG.TextBox83.Text.ToString & SerialNumber & "_" & TestMode & "_" & Status & Format(Now(), "hhmmssmmddyy") & ".pdf", vbNormal)
            workbook.Close(True)
            System.Runtime.InteropServices.Marshal.ReleaseComObject(workbook)
            workbook = Nothing
            xlapp.Quit()
            System.Runtime.InteropServices.Marshal.ReleaseComObject(xlapp)
            xlapp = Nothing
            Process.GetProcessesByName("EXCEL")(0).Kill()
            file1.Close()
            MAIN.ListBox1.Items.Add($"{DateTime.Now.ToString} Test Report Generated success : " & CONFIG.TextBox83.Text.ToString & SerialNumber & "_" & TestMode & "_" & Status & Format(Now(), "hhmmssmmddyy"))
            MAIN.ListBox1.SelectedIndex = MAIN.ListBox1.Items.Count - 1
            Dim SW As IO.StreamWriter = IO.File.CreateText(CONFIG.TextBox83.Text.ToString & SerialNumber)
            For Each S As String In MAIN.ListBox1.Items
                SW.WriteLine(S, True)
            Next
            SW.Close()
            file1.Close()
        Catch ex As Exception
            Dim iret As Object = MsgBox("Error generating Report : " & ex.Message, vbCritical + vbRetryCancel, "Error executing SMC.GenerateReportOnerrorInit")
            If iret = vbRetry Then
                file1.Close()
                GoTo line1
            ElseIf iret = vbCancel Then
                file1.WriteLine("Error generating Report :  " & ex.Message & " " & "@" & DateTime.Now.ToString("yyyy/MM/dd/HH:mm:ss"))
                file1.Close()
                MAIN.ListBox1.Items.Add($"{DateTime.Now.ToString} Error generating Report :" & ex.Message)
                MAIN.ListBox1.SelectedIndex = MAIN.ListBox1.Items.Count - 1
                Exit Function
            End If
        End Try
    End Function
    Function GenerateReportOnErrorOfFTDIProg(User As String, TestMode As String, SerialNumber As String, Status As String, CycleTime As String, TesterID As String, Mark As String, ErrorMsg As String, DutConnect As String, SetPower As String, UniflashProgramme As String, FTDIProg As String)
        Dim line1 As Label
line1:
        Dim path1 As String = CONFIG.TextBox4.Text
        file1 = My.Computer.FileSystem.OpenTextFileWriter(path1, True)
        Try
            Dim xlapp = New Microsoft.Office.Interop.Excel.Application
            xlapp.DisplayAlerts = False
            Dim workbook = xlapp.Workbooks.Open(CONFIG.TextBox82.Text.ToString)
            Dim worksheet As Excel.Worksheet
            worksheet = workbook.Sheets("SMC")
            worksheet.Range("D6").Value = User
            worksheet.Range("D7").Value = TestMode
            worksheet.Range("D8").Value = SerialNumber
            worksheet.Range("D9").Value = DateTime.Now.ToString
            worksheet.Range("D10").Value = Status
            worksheet.Range("F6").Value = CycleTime
            worksheet.Range("F7").Value = TesterID
            worksheet.Range("F8").Value = Mark
            worksheet.Range("G7").Value = ErrorMsg
            worksheet.Range("G13").Value = DutConnect
            worksheet.Range("G14").Value = SetPower
            worksheet.Range("G15").Value = UniflashProgramme
            worksheet.Range("G16").Value = FTDIProg
            workbook.ExportAsFixedFormat(Excel.XlFixedFormatType.xlTypePDF, CONFIG.TextBox83.Text.ToString & SerialNumber & "_" & TestMode & "_" & Status & Format(Now(), "hhmmssmmddyy") & ".pdf", vbNormal)
            workbook.Close(True)
            System.Runtime.InteropServices.Marshal.ReleaseComObject(workbook)
            workbook = Nothing
            xlapp.Quit()
            System.Runtime.InteropServices.Marshal.ReleaseComObject(xlapp)
            xlapp = Nothing
            Process.GetProcessesByName("EXCEL")(0).Kill()
            file1.Close()
            MAIN.ListBox1.Items.Add($"{DateTime.Now.ToString} Test Report Generated success : " & CONFIG.TextBox83.Text.ToString & SerialNumber & "_" & TestMode & "_" & Status & Format(Now(), "hhmmssmmddyy"))
            MAIN.ListBox1.SelectedIndex = MAIN.ListBox1.Items.Count - 1
            Dim SW As IO.StreamWriter = IO.File.CreateText(CONFIG.TextBox83.Text.ToString & SerialNumber)
            For Each S As String In MAIN.ListBox1.Items
                SW.WriteLine(S, True)
            Next
            SW.Close()
            file1.Close()
        Catch ex As Exception
            Dim iret As Object = MsgBox("Error generating Report : " & ex.Message, vbCritical + vbRetryCancel, "Error executing SMC.GenerateReportOnerrorInit")
            If iret = vbRetry Then
                file1.Close()
                GoTo line1
            ElseIf iret = vbCancel Then
                file1.WriteLine("Error generating Report :  " & ex.Message & " " & "@" & DateTime.Now.ToString("yyyy/MM/dd/HH:mm:ss"))
                file1.Close()
                MAIN.ListBox1.Items.Add($"{DateTime.Now.ToString} Error generating Report :" & ex.Message)
                MAIN.ListBox1.SelectedIndex = MAIN.ListBox1.Items.Count - 1
                Exit Function
            End If
        End Try
    End Function
    Function GenerateReportonErrorSMCPowerTestNode0(User As String, TestMode As String, SerialNumber As String, Status As String, CycleTime As String, TesterID As String, Mark As String, ErrorMsg As String, DutConnect As String, SetPower As String, UniflashProgramme As String, FTDIProg As String, InitialPower As String)
        Dim line1 As Label
line1:
        Dim path1 As String = CONFIG.TextBox4.Text
        file1 = My.Computer.FileSystem.OpenTextFileWriter(path1, True)
        Try
            Dim xlapp = New Microsoft.Office.Interop.Excel.Application
            xlapp.DisplayAlerts = False
            Dim workbook = xlapp.Workbooks.Open(CONFIG.TextBox82.Text.ToString)
            Dim worksheet As Excel.Worksheet
            worksheet = workbook.Sheets("SMC")
            worksheet.Range("D6").Value = User
            worksheet.Range("D7").Value = TestMode
            worksheet.Range("D8").Value = SerialNumber
            worksheet.Range("D9").Value = DateTime.Now.ToString
            worksheet.Range("D10").Value = Status
            worksheet.Range("F6").Value = CycleTime
            worksheet.Range("F7").Value = TesterID
            worksheet.Range("F8").Value = Mark
            worksheet.Range("G7").Value = ErrorMsg
            worksheet.Range("G13").Value = DutConnect
            worksheet.Range("G14").Value = SetPower
            worksheet.Range("G15").Value = UniflashProgramme
            worksheet.Range("G16").Value = FTDIProg
            worksheet.Range("G18").Value = InitialPower
            Dim worksheet1 As Excel.Worksheet = workbook.Sheets.Add(After:=worksheet)
            worksheet1.Name = "1"
            Clipboard.Clear()
            Clipboard.SetImage(IMAGE_STR.PictureBox1.Image)
            worksheet1.Paste()
            With worksheet1.PageSetup
                .Orientation = Excel.XlPageOrientation.xlLandscape
                .Zoom = False
                .FitToPagesTall = 1
                .FitToPagesWide = 1
            End With
            workbook.ExportAsFixedFormat(Excel.XlFixedFormatType.xlTypePDF, CONFIG.TextBox83.Text.ToString & SerialNumber & "_" & TestMode & "_" & Status & Format(Now(), "hhmmssmmddyy") & ".pdf", vbNormal)
            workbook.Worksheets("1").Delete()
            workbook.Save()
            workbook.Close(True)
            System.Runtime.InteropServices.Marshal.ReleaseComObject(workbook)
            workbook = Nothing
            xlapp.Quit()
            System.Runtime.InteropServices.Marshal.ReleaseComObject(xlapp)
            xlapp = Nothing
            Process.GetProcessesByName("EXCEL")(0).Kill()
            file1.Close()
            MAIN.ListBox1.Items.Add($"{DateTime.Now.ToString} Test Report Generated success : " & CONFIG.TextBox83.Text.ToString & SerialNumber & "_" & TestMode & "_" & Status & Format(Now(), "hhmmssmmddyy"))
            MAIN.ListBox1.SelectedIndex = MAIN.ListBox1.Items.Count - 1
            Dim SW As IO.StreamWriter = IO.File.CreateText(CONFIG.TextBox83.Text.ToString & SerialNumber)
            For Each S As String In MAIN.ListBox1.Items
                SW.WriteLine(S, True)
            Next
            SW.Close()

        Catch ex As Exception
            Dim iret As Object = MsgBox("Error generating Report : " & ex.Message, vbCritical + vbRetryCancel, "Error executing SMC.GenerateReportOnerrorInit")
            If iret = vbRetry Then
                file1.Close()
                GoTo line1
            ElseIf iret = vbCancel Then
                file1.WriteLine("Error generating Report :  " & ex.Message & " " & "@" & DateTime.Now.ToString("yyyy/MM/dd/HH:mm:ss"))
                file1.Close()
                MAIN.ListBox1.Items.Add($"{DateTime.Now.ToString} Error generating Report :" & ex.Message)
                MAIN.ListBox1.SelectedIndex = MAIN.ListBox1.Items.Count - 1
                Exit Function
            End If
        End Try
    End Function
    Function GenerateReportOnErrorSMCPowerTestNode1(User As String, TestMode As String, SerialNumber As String, Status As String, CycleTime As String, TesterID As String, Mark As String, ErrorMsg As String, DutConnect As String, SetPower As String, UniflashProgramme As String, FTDIProg As String, InitialPower As String, Can6705v As String)
        Dim line1 As Label
line1:
        Dim path1 As String = CONFIG.TextBox4.Text
        file1 = My.Computer.FileSystem.OpenTextFileWriter(path1, True)
        Try
            Dim xlapp = New Microsoft.Office.Interop.Excel.Application
            xlapp.DisplayAlerts = False
            Dim workbook = xlapp.Workbooks.Open(CONFIG.TextBox82.Text.ToString)
            Dim worksheet As Excel.Worksheet
            worksheet = workbook.Sheets("SMC")
            worksheet.Range("D6").Value = User
            worksheet.Range("D7").Value = TestMode
            worksheet.Range("D8").Value = SerialNumber
            worksheet.Range("D9").Value = DateTime.Now.ToString
            worksheet.Range("D10").Value = Status
            worksheet.Range("F6").Value = CycleTime
            worksheet.Range("F7").Value = TesterID
            worksheet.Range("F8").Value = Mark
            worksheet.Range("G7").Value = ErrorMsg
            worksheet.Range("G13").Value = DutConnect
            worksheet.Range("G14").Value = SetPower
            worksheet.Range("G15").Value = UniflashProgramme
            worksheet.Range("G16").Value = FTDIProg
            worksheet.Range("G18").Value = InitialPower
            worksheet.Range("G19").Value = Can6705v
            Dim worksheet1 As Excel.Worksheet = workbook.Sheets.Add(After:=worksheet)
            worksheet1.Name = "1"
            Clipboard.Clear()
            Clipboard.SetImage(IMAGE_STR.PictureBox1.Image)
            worksheet1.Paste()
            With worksheet1.PageSetup
                .Orientation = Excel.XlPageOrientation.xlLandscape
                .Zoom = False
                .FitToPagesTall = 1
                .FitToPagesWide = 1
            End With
            workbook.ExportAsFixedFormat(Excel.XlFixedFormatType.xlTypePDF, CONFIG.TextBox83.Text.ToString & SerialNumber & "_" & TestMode & "_" & Status & Format(Now(), "hhmmssmmddyy") & ".pdf", vbNormal)
            workbook.Worksheets("1").Delete()
            workbook.Close(True)
            System.Runtime.InteropServices.Marshal.ReleaseComObject(workbook)
            workbook = Nothing
            xlapp.Quit()
            System.Runtime.InteropServices.Marshal.ReleaseComObject(xlapp)
            xlapp = Nothing
            Process.GetProcessesByName("EXCEL")(0).Kill()
            file1.Close()
            MAIN.ListBox1.Items.Add($"{DateTime.Now.ToString} Test Report Generated success : " & CONFIG.TextBox83.Text.ToString & SerialNumber & "_" & TestMode & "_" & Status & Format(Now(), "hhmmssmmddyy"))
            MAIN.ListBox1.SelectedIndex = MAIN.ListBox1.Items.Count - 1
            Dim SW As IO.StreamWriter = IO.File.CreateText(CONFIG.TextBox83.Text.ToString & SerialNumber)
            For Each S As String In MAIN.ListBox1.Items
                SW.WriteLine(S, True)
            Next
            SW.Close()
            file1.Close()
        Catch ex As Exception
            Dim iret As Object = MsgBox("Error generating Report : " & ex.Message, vbCritical + vbRetryCancel, "Error executing SMC.GenerateReportOnerrorInit")
            If iret = vbRetry Then
                file1.Close()
                GoTo line1
            ElseIf iret = vbCancel Then
                file1.WriteLine("Error generating Report :  " & ex.Message & " " & "@" & DateTime.Now.ToString("yyyy/MM/dd/HH:mm:ss"))
                file1.Close()
                MAIN.ListBox1.Items.Add($"{DateTime.Now.ToString} Error generating Report :" & ex.Message)
                MAIN.ListBox1.SelectedIndex = MAIN.ListBox1.Items.Count - 1
                Exit Function
            End If
        End Try
    End Function
    Function GenerateReportOnErrorSMCPowerTestNode2(User As String, TestMode As String, SerialNumber As String, Status As String, CycleTime As String, TesterID As String, Mark As String, ErrorMsg As String, DutConnect As String, SetPower As String, UniflashProgramme As String, FTDIProg As String, InitialPower As String, Can6705v As String, Isolatedv12 As String)
        Dim line1 As Label
line1:
        Dim path1 As String = CONFIG.TextBox4.Text
        file1 = My.Computer.FileSystem.OpenTextFileWriter(path1, True)
        Try
            Dim xlapp = New Microsoft.Office.Interop.Excel.Application
            xlapp.DisplayAlerts = False
            Dim workbook = xlapp.Workbooks.Open(CONFIG.TextBox82.Text.ToString)
            Dim worksheet As Excel.Worksheet
            worksheet = workbook.Sheets("SMC")
            worksheet.Range("D6").Value = User
            worksheet.Range("D7").Value = TestMode
            worksheet.Range("D8").Value = SerialNumber
            worksheet.Range("D9").Value = DateTime.Now.ToString
            worksheet.Range("D10").Value = Status
            worksheet.Range("F6").Value = CycleTime
            worksheet.Range("F7").Value = TesterID
            worksheet.Range("F8").Value = Mark
            worksheet.Range("G7").Value = ErrorMsg
            worksheet.Range("G13").Value = DutConnect
            worksheet.Range("G14").Value = SetPower
            worksheet.Range("G15").Value = UniflashProgramme
            worksheet.Range("G16").Value = FTDIProg
            worksheet.Range("G18").Value = InitialPower
            worksheet.Range("G19").Value = Can6705v
            worksheet.Range("G20").Value = Isolatedv12
            Dim worksheet1 As Excel.Worksheet = workbook.Sheets.Add(After:=worksheet)
            worksheet1.Name = "1"
            Clipboard.Clear()
            Clipboard.SetImage(IMAGE_STR.PictureBox1.Image)
            worksheet1.Paste()
            With worksheet1.PageSetup
                .Orientation = Excel.XlPageOrientation.xlLandscape
                .Zoom = False
                .FitToPagesTall = 1
                .FitToPagesWide = 1
            End With
            workbook.ExportAsFixedFormat(Excel.XlFixedFormatType.xlTypePDF, CONFIG.TextBox83.Text.ToString & SerialNumber & "_" & TestMode & "_" & Status & Format(Now(), "hhmmssmmddyy") & ".pdf", vbNormal)
            workbook.Worksheets("1").Delete()
            workbook.Close(True)
            System.Runtime.InteropServices.Marshal.ReleaseComObject(workbook)
            workbook = Nothing
            xlapp.Quit()
            System.Runtime.InteropServices.Marshal.ReleaseComObject(xlapp)
            xlapp = Nothing
            Process.GetProcessesByName("EXCEL")(0).Kill()
            file1.Close()
            MAIN.ListBox1.Items.Add($"{DateTime.Now.ToString} Test Report Generated success : " & CONFIG.TextBox83.Text.ToString & SerialNumber & "_" & TestMode & "_" & Status & Format(Now(), "hhmmssmmddyy"))
            MAIN.ListBox1.SelectedIndex = MAIN.ListBox1.Items.Count - 1
            Dim SW As IO.StreamWriter = IO.File.CreateText(CONFIG.TextBox83.Text.ToString & SerialNumber)
            For Each S As String In MAIN.ListBox1.Items
                SW.WriteLine(S, True)
            Next
            SW.Close()
            file1.Close()
        Catch ex As Exception
            Dim iret As Object = MsgBox("Error generating Report : " & ex.Message, vbCritical + vbRetryCancel, "Error executing SMC.GenerateReportOnerrorInit")
            If iret = vbRetry Then
                file1.Close()
                GoTo line1
            ElseIf iret = vbCancel Then
                file1.WriteLine("Error generating Report :  " & ex.Message & " " & "@" & DateTime.Now.ToString("yyyy/MM/dd/HH:mm:ss"))
                file1.Close()
                MAIN.ListBox1.Items.Add($"{DateTime.Now.ToString} Error generating Report :" & ex.Message)
                MAIN.ListBox1.SelectedIndex = MAIN.ListBox1.Items.Count - 1
                Exit Function
            End If
        End Try
    End Function
    Function GenerateReportOnErrorSMCPowerTestNode3(User As String, TestMode As String, SerialNumber As String, Status As String, CycleTime As String, TesterID As String, Mark As String, ErrorMsg As String, DutConnect As String, SetPower As String, UniflashProgramme As String, FTDIProg As String, InitialPower As String, Can6705v As String, Isolatedv12 As String, Sensorv12 As String)
        Dim line1 As Label
line1:
        Dim path1 As String = CONFIG.TextBox4.Text
        file1 = My.Computer.FileSystem.OpenTextFileWriter(path1, True)
        Try
            Dim xlapp = New Microsoft.Office.Interop.Excel.Application
            xlapp.DisplayAlerts = False
            Dim workbook = xlapp.Workbooks.Open(CONFIG.TextBox82.Text.ToString)
            Dim worksheet As Excel.Worksheet
            worksheet = workbook.Sheets("SMC")
            worksheet.Range("D6").Value = User
            worksheet.Range("D7").Value = TestMode
            worksheet.Range("D8").Value = SerialNumber
            worksheet.Range("D9").Value = DateTime.Now.ToString
            worksheet.Range("D10").Value = Status
            worksheet.Range("F6").Value = CycleTime
            worksheet.Range("F7").Value = TesterID
            worksheet.Range("F8").Value = Mark
            worksheet.Range("G7").Value = ErrorMsg
            worksheet.Range("G13").Value = DutConnect
            worksheet.Range("G14").Value = SetPower
            worksheet.Range("G15").Value = UniflashProgramme
            worksheet.Range("G16").Value = FTDIProg
            worksheet.Range("G18").Value = InitialPower
            worksheet.Range("G19").Value = Can6705v
            worksheet.Range("G20").Value = Isolatedv12
            worksheet.Range("G21").Value = Sensorv12
            Dim worksheet1 As Excel.Worksheet = workbook.Sheets.Add(After:=worksheet)
            worksheet1.Name = "1"
            Clipboard.Clear()
            Clipboard.SetImage(IMAGE_STR.PictureBox1.Image)
            worksheet1.Paste()
            With worksheet1.PageSetup
                .Orientation = Excel.XlPageOrientation.xlLandscape
                .Zoom = False
                .FitToPagesTall = 1
                .FitToPagesWide = 1
            End With
            workbook.ExportAsFixedFormat(Excel.XlFixedFormatType.xlTypePDF, CONFIG.TextBox83.Text.ToString & SerialNumber & "_" & TestMode & "_" & Status & Format(Now(), "hhmmssmmddyy") & ".pdf", vbNormal)
            workbook.Worksheets("1").Delete()
            workbook.Close(True)
            System.Runtime.InteropServices.Marshal.ReleaseComObject(workbook)
            workbook = Nothing
            xlapp.Quit()
            System.Runtime.InteropServices.Marshal.ReleaseComObject(xlapp)
            xlapp = Nothing
            Process.GetProcessesByName("EXCEL")(0).Kill()
            file1.Close()
            MAIN.ListBox1.Items.Add($"{DateTime.Now.ToString} Test Report Generated success : " & CONFIG.TextBox83.Text.ToString & SerialNumber & "_" & TestMode & "_" & Status & Format(Now(), "hhmmssmmddyy"))
            MAIN.ListBox1.SelectedIndex = MAIN.ListBox1.Items.Count - 1
            Dim SW As IO.StreamWriter = IO.File.CreateText(CONFIG.TextBox83.Text.ToString & SerialNumber)
            For Each S As String In MAIN.ListBox1.Items
                SW.WriteLine(S, True)
            Next
            SW.Close()
            file1.Close()
        Catch ex As Exception
            Dim iret As Object = MsgBox("Error generating Report : " & ex.Message, vbCritical + vbRetryCancel, "Error executing SMC.GenerateReportOnerrorInit")
            If iret = vbRetry Then
                file1.Close()
                GoTo line1
            ElseIf iret = vbCancel Then
                file1.WriteLine("Error generating Report :  " & ex.Message & " " & "@" & DateTime.Now.ToString("yyyy/MM/dd/HH:mm:ss"))
                file1.Close()
                MAIN.ListBox1.Items.Add($"{DateTime.Now.ToString} Error generating Report :" & ex.Message)
                MAIN.ListBox1.SelectedIndex = MAIN.ListBox1.Items.Count - 1
                Exit Function
            End If
        End Try
    End Function
    Function GenerateReportOnErrorSMCPowerTestNode4(User As String, TestMode As String, SerialNumber As String, Status As String, CycleTime As String, TesterID As String, Mark As String, ErrorMsg As String, DutConnect As String, SetPower As String, UniflashProgramme As String, FTDIProg As String, InitialPower As String, Can6705v As String, Isolatedv12 As String, Sensorv12 As String, TT6205v As String)
        Dim line1 As Label
line1:
        Dim path1 As String = CONFIG.TextBox4.Text
        file1 = My.Computer.FileSystem.OpenTextFileWriter(path1, True)
        Try
            Dim xlapp = New Microsoft.Office.Interop.Excel.Application
            xlapp.DisplayAlerts = False
            Dim workbook = xlapp.Workbooks.Open(CONFIG.TextBox82.Text.ToString)
            Dim worksheet As Excel.Worksheet
            worksheet = workbook.Sheets("SMC")
            worksheet.Range("D6").Value = User
            worksheet.Range("D7").Value = TestMode
            worksheet.Range("D8").Value = SerialNumber
            worksheet.Range("D9").Value = DateTime.Now.ToString
            worksheet.Range("D10").Value = Status
            worksheet.Range("F6").Value = CycleTime
            worksheet.Range("F7").Value = TesterID
            worksheet.Range("F8").Value = Mark
            worksheet.Range("G7").Value = ErrorMsg
            worksheet.Range("G13").Value = DutConnect
            worksheet.Range("G14").Value = SetPower
            worksheet.Range("G15").Value = UniflashProgramme
            worksheet.Range("G16").Value = FTDIProg
            worksheet.Range("G18").Value = InitialPower
            worksheet.Range("G19").Value = Can6705v
            worksheet.Range("G20").Value = Isolatedv12
            worksheet.Range("G21").Value = Sensorv12
            worksheet.Range("G22").Value = TT6205v
            Dim worksheet1 As Excel.Worksheet = workbook.Sheets.Add(After:=worksheet)
            worksheet1.Name = "1"
            Clipboard.Clear()
            Clipboard.SetImage(IMAGE_STR.PictureBox1.Image)
            worksheet1.Paste()
            With worksheet1.PageSetup
                .Orientation = Excel.XlPageOrientation.xlLandscape
                .Zoom = False
                .FitToPagesTall = 1
                .FitToPagesWide = 1
            End With
            workbook.ExportAsFixedFormat(Excel.XlFixedFormatType.xlTypePDF, CONFIG.TextBox83.Text.ToString & SerialNumber & "_" & TestMode & "_" & Status & Format(Now(), "hhmmssmmddyy") & ".pdf", vbNormal)
            workbook.Worksheets("1").Delete()
            workbook.Close(True)
            System.Runtime.InteropServices.Marshal.ReleaseComObject(workbook)
            workbook = Nothing
            xlapp.Quit()
            System.Runtime.InteropServices.Marshal.ReleaseComObject(xlapp)
            xlapp = Nothing
            Process.GetProcessesByName("EXCEL")(0).Kill()
            file1.Close()
            MAIN.ListBox1.Items.Add($"{DateTime.Now.ToString} Test Report Generated success : " & CONFIG.TextBox83.Text.ToString & SerialNumber & "_" & TestMode & "_" & Status & Format(Now(), "hhmmssmmddyy"))
            MAIN.ListBox1.SelectedIndex = MAIN.ListBox1.Items.Count - 1
            Dim SW As IO.StreamWriter = IO.File.CreateText(CONFIG.TextBox83.Text.ToString & SerialNumber)
            For Each S As String In MAIN.ListBox1.Items
                SW.WriteLine(S, True)
            Next
            SW.Close()
            file1.Close()
        Catch ex As Exception
            Dim iret As Object = MsgBox("Error generating Report : " & ex.Message, vbCritical + vbRetryCancel, "Error executing SMC.GenerateReportOnerrorInit")
            If iret = vbRetry Then
                file1.Close()
                GoTo line1
            ElseIf iret = vbCancel Then
                file1.WriteLine("Error generating Report :  " & ex.Message & " " & "@" & DateTime.Now.ToString("yyyy/MM/dd/HH:mm:ss"))
                file1.Close()
                MAIN.ListBox1.Items.Add($"{DateTime.Now.ToString} Error generating Report :" & ex.Message)
                MAIN.ListBox1.SelectedIndex = MAIN.ListBox1.Items.Count - 1
                Exit Function
            End If
        End Try
    End Function
    Function GenerateReportOnErrorSMCPowerTestNode5(User As String, TestMode As String, SerialNumber As String, Status As String, CycleTime As String, TesterID As String, Mark As String, ErrorMsg As String, DutConnect As String, SetPower As String, UniflashProgramme As String, FTDIProg As String, InitialPower As String, Can6705v As String, Isolatedv12 As String, Sensorv12 As String, TT6205v As String, IT6615v As String)
        Dim line1 As Label
line1:
        Dim path1 As String = CONFIG.TextBox4.Text
        file1 = My.Computer.FileSystem.OpenTextFileWriter(path1, True)
        Try
            Dim xlapp = New Microsoft.Office.Interop.Excel.Application
            xlapp.DisplayAlerts = False
            Dim workbook = xlapp.Workbooks.Open(CONFIG.TextBox82.Text.ToString)
            Dim worksheet As Excel.Worksheet
            worksheet = workbook.Sheets("SMC")
            worksheet.Range("D6").Value = User
            worksheet.Range("D7").Value = TestMode
            worksheet.Range("D8").Value = SerialNumber
            worksheet.Range("D9").Value = DateTime.Now.ToString
            worksheet.Range("D10").Value = Status
            worksheet.Range("F6").Value = CycleTime
            worksheet.Range("F7").Value = TesterID
            worksheet.Range("F8").Value = Mark
            worksheet.Range("G7").Value = ErrorMsg
            worksheet.Range("G13").Value = DutConnect
            worksheet.Range("G14").Value = SetPower
            worksheet.Range("G15").Value = UniflashProgramme
            worksheet.Range("G16").Value = FTDIProg
            worksheet.Range("G18").Value = InitialPower
            worksheet.Range("G19").Value = Can6705v
            worksheet.Range("G20").Value = Isolatedv12
            worksheet.Range("G21").Value = Sensorv12
            worksheet.Range("G22").Value = TT6205v
            worksheet.Range("G23").Value = IT6615v
            Dim worksheet1 As Excel.Worksheet = workbook.Sheets.Add(After:=worksheet)
            worksheet1.Name = "1"
            Clipboard.Clear()
            Clipboard.SetImage(IMAGE_STR.PictureBox1.Image)
            worksheet1.Paste()
            With worksheet1.PageSetup
                .Orientation = Excel.XlPageOrientation.xlLandscape
                .Zoom = False
                .FitToPagesTall = 1
                .FitToPagesWide = 1
            End With
            workbook.ExportAsFixedFormat(Excel.XlFixedFormatType.xlTypePDF, CONFIG.TextBox83.Text.ToString & SerialNumber & "_" & TestMode & "_" & Status & Format(Now(), "hhmmssmmddyy") & ".pdf", vbNormal)
            workbook.Worksheets("1").Delete()
            workbook.Close(True)
            System.Runtime.InteropServices.Marshal.ReleaseComObject(workbook)
            workbook = Nothing
            xlapp.Quit()
            System.Runtime.InteropServices.Marshal.ReleaseComObject(xlapp)
            xlapp = Nothing
            Process.GetProcessesByName("EXCEL")(0).Kill()
            file1.Close()
            MAIN.ListBox1.Items.Add($"{DateTime.Now.ToString} Test Report Generated success : " & CONFIG.TextBox83.Text.ToString & SerialNumber & "_" & TestMode & "_" & Status & Format(Now(), "hhmmssmmddyy"))
            MAIN.ListBox1.SelectedIndex = MAIN.ListBox1.Items.Count - 1
            Dim SW As IO.StreamWriter = IO.File.CreateText(CONFIG.TextBox83.Text.ToString & SerialNumber)
            For Each S As String In MAIN.ListBox1.Items
                SW.WriteLine(S, True)
            Next
            SW.Close()
            file1.Close()
        Catch ex As Exception
            Dim iret As Object = MsgBox("Error generating Report : " & ex.Message, vbCritical + vbRetryCancel, "Error executing SMC.GenerateReportOnerrorInit")
            If iret = vbRetry Then
                file1.Close()
                GoTo line1
            ElseIf iret = vbCancel Then
                file1.WriteLine("Error generating Report :  " & ex.Message & " " & "@" & DateTime.Now.ToString("yyyy/MM/dd/HH:mm:ss"))
                file1.Close()
                MAIN.ListBox1.Items.Add($"{DateTime.Now.ToString} Error generating Report :" & ex.Message)
                MAIN.ListBox1.SelectedIndex = MAIN.ListBox1.Items.Count - 1
                Exit Function
            End If
        End Try
    End Function
    Function GenerateReportOnErrorSMCPowerTestNode6(User As String, TestMode As String, SerialNumber As String, Status As String, CycleTime As String, TesterID As String, Mark As String, ErrorMsg As String, DutConnect As String, SetPower As String, UniflashProgramme As String, FTDIProg As String, InitialPower As String, Can6705v As String, Isolatedv12 As String, Sensorv12 As String, TT6205v As String, IT6615v As String, A0691 As String)
        Dim line1 As Label
line1:
        Dim path1 As String = CONFIG.TextBox4.Text
        file1 = My.Computer.FileSystem.OpenTextFileWriter(path1, True)
        Try
            Dim xlapp = New Microsoft.Office.Interop.Excel.Application
            xlapp.DisplayAlerts = False
            Dim workbook = xlapp.Workbooks.Open(CONFIG.TextBox82.Text.ToString)
            Dim worksheet As Excel.Worksheet
            worksheet = workbook.Sheets("SMC")
            worksheet.Range("D6").Value = User
            worksheet.Range("D7").Value = TestMode
            worksheet.Range("D8").Value = SerialNumber
            worksheet.Range("D9").Value = DateTime.Now.ToString
            worksheet.Range("D10").Value = Status
            worksheet.Range("F6").Value = CycleTime
            worksheet.Range("F7").Value = TesterID
            worksheet.Range("F8").Value = Mark
            worksheet.Range("G7").Value = ErrorMsg
            worksheet.Range("G13").Value = DutConnect
            worksheet.Range("G14").Value = SetPower
            worksheet.Range("G15").Value = UniflashProgramme
            worksheet.Range("G16").Value = FTDIProg
            worksheet.Range("G18").Value = InitialPower
            worksheet.Range("G19").Value = Can6705v
            worksheet.Range("G20").Value = Isolatedv12
            worksheet.Range("G21").Value = Sensorv12
            worksheet.Range("G22").Value = TT6205v
            worksheet.Range("G23").Value = IT6615v
            worksheet.Range("G24").Value = A0691
            Dim worksheet1 As Excel.Worksheet = workbook.Sheets.Add(After:=worksheet)
            worksheet1.Name = "1"
            Clipboard.Clear()
            Clipboard.SetImage(IMAGE_STR.PictureBox1.Image)
            worksheet1.Paste()
            With worksheet1.PageSetup
                .Orientation = Excel.XlPageOrientation.xlLandscape
                .Zoom = False
                .FitToPagesTall = 1
                .FitToPagesWide = 1
            End With
            workbook.ExportAsFixedFormat(Excel.XlFixedFormatType.xlTypePDF, CONFIG.TextBox83.Text.ToString & SerialNumber & "_" & TestMode & "_" & Status & Format(Now(), "hhmmssmmddyy") & ".pdf", vbNormal)
            workbook.Worksheets("1").Delete()
            workbook.Close(True)
            System.Runtime.InteropServices.Marshal.ReleaseComObject(workbook)
            workbook = Nothing
            xlapp.Quit()
            System.Runtime.InteropServices.Marshal.ReleaseComObject(xlapp)
            xlapp = Nothing
            Process.GetProcessesByName("EXCEL")(0).Kill()
            file1.Close()
            MAIN.ListBox1.Items.Add($"{DateTime.Now.ToString} Test Report Generated success : " & CONFIG.TextBox83.Text.ToString & SerialNumber & "_" & TestMode & "_" & Status & Format(Now(), "hhmmssmmddyy"))
            MAIN.ListBox1.SelectedIndex = MAIN.ListBox1.Items.Count - 1
            Dim SW As IO.StreamWriter = IO.File.CreateText(CONFIG.TextBox83.Text.ToString & SerialNumber)
            For Each S As String In MAIN.ListBox1.Items
                SW.WriteLine(S, True)
            Next
            SW.Close()
            file1.Close()
        Catch ex As Exception
            Dim iret As Object = MsgBox("Error generating Report : " & ex.Message, vbCritical + vbRetryCancel, "Error executing SMC.GenerateReportOnerrorInit")
            If iret = vbRetry Then
                file1.Close()
                GoTo line1
            ElseIf iret = vbCancel Then
                file1.WriteLine("Error generating Report :  " & ex.Message & " " & "@" & DateTime.Now.ToString("yyyy/MM/dd/HH:mm:ss"))
                file1.Close()
                MAIN.ListBox1.Items.Add($"{DateTime.Now.ToString} Error generating Report :" & ex.Message)
                MAIN.ListBox1.SelectedIndex = MAIN.ListBox1.Items.Count - 1
                Exit Function
            End If
        End Try
    End Function
    Function GenerateReportOnErrorSMCPowerTestNode7(User As String, TestMode As String, SerialNumber As String, Status As String, CycleTime As String, TesterID As String, Mark As String, ErrorMsg As String, DutConnect As String, SetPower As String, UniflashProgramme As String, FTDIProg As String, InitialPower As String, Can6705v As String, Isolatedv12 As String, Sensorv12 As String, TT6205v As String, IT6615v As String, A0691 As String, D0691 As String)
        Dim line1 As Label
line1:
        Dim path1 As String = CONFIG.TextBox4.Text
        file1 = My.Computer.FileSystem.OpenTextFileWriter(path1, True)
        Try
            Dim xlapp = New Microsoft.Office.Interop.Excel.Application
            xlapp.DisplayAlerts = False
            Dim workbook = xlapp.Workbooks.Open(CONFIG.TextBox82.Text.ToString)
            Dim worksheet As Excel.Worksheet
            worksheet = workbook.Sheets("SMC")
            worksheet.Range("D6").Value = User
            worksheet.Range("D7").Value = TestMode
            worksheet.Range("D8").Value = SerialNumber
            worksheet.Range("D9").Value = DateTime.Now.ToString
            worksheet.Range("D10").Value = Status
            worksheet.Range("F6").Value = CycleTime
            worksheet.Range("F7").Value = TesterID
            worksheet.Range("F8").Value = Mark
            worksheet.Range("G7").Value = ErrorMsg
            worksheet.Range("G13").Value = DutConnect
            worksheet.Range("G14").Value = SetPower
            worksheet.Range("G15").Value = UniflashProgramme
            worksheet.Range("G16").Value = FTDIProg
            worksheet.Range("G18").Value = InitialPower
            worksheet.Range("G19").Value = Can6705v
            worksheet.Range("G20").Value = Isolatedv12
            worksheet.Range("G21").Value = Sensorv12
            worksheet.Range("G22").Value = TT6205v
            worksheet.Range("G23").Value = IT6615v
            worksheet.Range("G24").Value = A0691
            worksheet.Range("G25").Value = D0691
            Dim worksheet1 As Excel.Worksheet = workbook.Sheets.Add(After:=worksheet)
            worksheet1.Name = "1"
            Clipboard.Clear()
            Clipboard.SetImage(IMAGE_STR.PictureBox1.Image)
            worksheet1.Paste()
            With worksheet1.PageSetup
                .Orientation = Excel.XlPageOrientation.xlLandscape
                .Zoom = False
                .FitToPagesTall = 1
                .FitToPagesWide = 1
            End With
            workbook.ExportAsFixedFormat(Excel.XlFixedFormatType.xlTypePDF, CONFIG.TextBox83.Text.ToString & SerialNumber & "_" & TestMode & "_" & Status & Format(Now(), "hhmmssmmddyy") & ".pdf", vbNormal)
            workbook.Worksheets("1").Delete()
            workbook.Close(True)
            System.Runtime.InteropServices.Marshal.ReleaseComObject(workbook)
            workbook = Nothing
            xlapp.Quit()
            System.Runtime.InteropServices.Marshal.ReleaseComObject(xlapp)
            xlapp = Nothing
            Process.GetProcessesByName("EXCEL")(0).Kill()
            file1.Close()
            MAIN.ListBox1.Items.Add($"{DateTime.Now.ToString} Test Report Generated success : " & CONFIG.TextBox83.Text.ToString & SerialNumber & "_" & TestMode & "_" & Status & Format(Now(), "hhmmssmmddyy"))
            MAIN.ListBox1.SelectedIndex = MAIN.ListBox1.Items.Count - 1
            Dim SW As IO.StreamWriter = IO.File.CreateText(CONFIG.TextBox83.Text.ToString & SerialNumber)
            For Each S As String In MAIN.ListBox1.Items
                SW.WriteLine(S, True)
            Next
            SW.Close()
            file1.Close()
        Catch ex As Exception
            Dim iret As Object = MsgBox("Error generating Report : " & ex.Message, vbCritical + vbRetryCancel, "Error executing SMC.GenerateReportOnerrorInit")
            If iret = vbRetry Then
                file1.Close()
                GoTo line1
            ElseIf iret = vbCancel Then
                file1.WriteLine("Error generating Report :  " & ex.Message & " " & "@" & DateTime.Now.ToString("yyyy/MM/dd/HH:mm:ss"))
                file1.Close()
                MAIN.ListBox1.Items.Add($"{DateTime.Now.ToString} Error generating Report :" & ex.Message)
                MAIN.ListBox1.SelectedIndex = MAIN.ListBox1.Items.Count - 1
                Exit Function
            End If
        End Try
    End Function
    Function GenerateReportOnErrorSMCSwitchTestNode0(User As String, TestMode As String, SerialNumber As String, Status As String, CycleTime As String, TesterID As String, Mark As String, ErrorMsg As String, DutConnect As String, SetPower As String, UniflashProgramme As String, FTDIProg As String, InitialPower As String, Can6705v As String, Isolatedv12 As String, Sensorv12 As String, TT6205v As String, IT6615v As String, A0691 As String, D0691 As String, D0691ON As String)
        Dim line1 As Label
line1:
        Dim path1 As String = CONFIG.TextBox4.Text
        file1 = My.Computer.FileSystem.OpenTextFileWriter(path1, True)
        Try
            Dim xlapp = New Microsoft.Office.Interop.Excel.Application
            xlapp.DisplayAlerts = False
            Dim workbook = xlapp.Workbooks.Open(CONFIG.TextBox82.Text.ToString)
            Dim worksheet As Excel.Worksheet
            worksheet = workbook.Sheets("SMC")
            worksheet.Range("D6").Value = User
            worksheet.Range("D7").Value = TestMode
            worksheet.Range("D8").Value = SerialNumber
            worksheet.Range("D9").Value = DateTime.Now.ToString
            worksheet.Range("D10").Value = Status
            worksheet.Range("F6").Value = CycleTime
            worksheet.Range("F7").Value = TesterID
            worksheet.Range("F8").Value = Mark
            worksheet.Range("G7").Value = ErrorMsg
            worksheet.Range("G13").Value = DutConnect
            worksheet.Range("G14").Value = SetPower
            worksheet.Range("G15").Value = UniflashProgramme
            worksheet.Range("G16").Value = FTDIProg
            worksheet.Range("G18").Value = InitialPower
            worksheet.Range("G19").Value = Can6705v
            worksheet.Range("G20").Value = Isolatedv12
            worksheet.Range("G21").Value = Sensorv12
            worksheet.Range("G22").Value = TT6205v
            worksheet.Range("G23").Value = IT6615v
            worksheet.Range("G24").Value = A0691
            worksheet.Range("G25").Value = D0691
            worksheet.Range("G27").Value = D0691ON
            Dim worksheet1 As Excel.Worksheet = workbook.Sheets.Add(After:=worksheet)
            worksheet1.Name = "1"
            Dim worksheet2 As Excel.Worksheet = workbook.Sheets.Add(After:=worksheet1)
            worksheet2.Name = "2"
            Clipboard.Clear()
            Clipboard.SetImage(IMAGE_STR.PictureBox1.Image)
            worksheet1.Paste()
            Clipboard.Clear()
            Clipboard.SetImage(IMAGE_STR.PictureBox2.Image)
            worksheet2.Paste()
            Clipboard.Clear()
            With worksheet1.PageSetup
                .Orientation = Excel.XlPageOrientation.xlLandscape
                .Zoom = False
                .FitToPagesTall = 1
                .FitToPagesWide = 1
            End With
            With worksheet2.PageSetup
                .Orientation = Excel.XlPageOrientation.xlLandscape
                .Zoom = False
                .FitToPagesTall = 1
                .FitToPagesWide = 1
            End With
            workbook.ExportAsFixedFormat(Excel.XlFixedFormatType.xlTypePDF, CONFIG.TextBox83.Text.ToString & SerialNumber & "_" & TestMode & "_" & Status & Format(Now(), "hhmmssmmddyy") & ".pdf", vbNormal)
            workbook.Worksheets("1").Delete()
            workbook.Worksheets("2").Delete()
            workbook.Close(True)
            System.Runtime.InteropServices.Marshal.ReleaseComObject(workbook)
            workbook = Nothing
            xlapp.Quit()
            System.Runtime.InteropServices.Marshal.ReleaseComObject(xlapp)
            xlapp = Nothing
            Process.GetProcessesByName("EXCEL")(0).Kill()
            file1.Close()
            MAIN.ListBox1.Items.Add($"{DateTime.Now.ToString} Test Report Generated success : " & CONFIG.TextBox83.Text.ToString & SerialNumber & "_" & TestMode & "_" & Status & Format(Now(), "hhmmssmmddyy"))
            MAIN.ListBox1.SelectedIndex = MAIN.ListBox1.Items.Count - 1
            Dim SW As IO.StreamWriter = IO.File.CreateText(CONFIG.TextBox83.Text.ToString & SerialNumber)
            For Each S As String In MAIN.ListBox1.Items
                SW.WriteLine(S, True)
            Next
            SW.Close()
            file1.Close()
        Catch ex As Exception
            Dim iret As Object = MsgBox("Error generating Report : " & ex.Message, vbCritical + vbRetryCancel, "Error executing SMC.GenerateReportOnerrorInit")
            If iret = vbRetry Then
                file1.Close()
                GoTo line1
            ElseIf iret = vbCancel Then
                file1.WriteLine("Error generating Report :  " & ex.Message & " " & "@" & DateTime.Now.ToString("yyyy/MM/dd/HH:mm:ss"))
                file1.Close()
                MAIN.ListBox1.Items.Add($"{DateTime.Now.ToString} Error generating Report :" & ex.Message)
                MAIN.ListBox1.SelectedIndex = MAIN.ListBox1.Items.Count - 1
                Exit Function
            End If
        End Try
    End Function
    Function GenerateReportOnErrorSMCSwitchTestNode1(User As String, TestMode As String, SerialNumber As String, Status As String, CycleTime As String, TesterID As String, Mark As String, ErrorMsg As String, DutConnect As String, SetPower As String, UniflashProgramme As String, FTDIProg As String, InitialPower As String, Can6705v As String, Isolatedv12 As String, Sensorv12 As String, TT6205v As String, IT6615v As String, A0691 As String, D0691 As String, D0691ON As String, D0661ON As String)
        Dim line1 As Label
line1:
        Dim path1 As String = CONFIG.TextBox4.Text
        file1 = My.Computer.FileSystem.OpenTextFileWriter(path1, True)
        Try
            Dim xlapp = New Microsoft.Office.Interop.Excel.Application
            xlapp.DisplayAlerts = False
            Dim workbook = xlapp.Workbooks.Open(CONFIG.TextBox82.Text.ToString)
            Dim worksheet As Excel.Worksheet
            worksheet = workbook.Sheets("SMC")
            worksheet.Range("D6").Value = User
            worksheet.Range("D7").Value = TestMode
            worksheet.Range("D8").Value = SerialNumber
            worksheet.Range("D9").Value = DateTime.Now.ToString
            worksheet.Range("D10").Value = Status
            worksheet.Range("F6").Value = CycleTime
            worksheet.Range("F7").Value = TesterID
            worksheet.Range("F8").Value = Mark
            worksheet.Range("G7").Value = ErrorMsg
            worksheet.Range("G13").Value = DutConnect
            worksheet.Range("G14").Value = SetPower
            worksheet.Range("G15").Value = UniflashProgramme
            worksheet.Range("G16").Value = FTDIProg
            worksheet.Range("G18").Value = InitialPower
            worksheet.Range("G19").Value = Can6705v
            worksheet.Range("G20").Value = Isolatedv12
            worksheet.Range("G21").Value = Sensorv12
            worksheet.Range("G22").Value = TT6205v
            worksheet.Range("G23").Value = IT6615v
            worksheet.Range("G24").Value = A0691
            worksheet.Range("G25").Value = D0691
            worksheet.Range("G27").Value = D0691ON
            worksheet.Range("G28").Value = D0661ON
            Dim worksheet1 As Excel.Worksheet = workbook.Sheets.Add(After:=worksheet)
            worksheet1.Name = "1"
            Dim worksheet2 As Excel.Worksheet = workbook.Sheets.Add(After:=worksheet1)
            worksheet2.Name = "2"
            Dim worksheet3 As Excel.Worksheet = workbook.Sheets.Add(After:=worksheet2)
            worksheet3.Name = "3"
            Clipboard.Clear()
            Clipboard.SetImage(IMAGE_STR.PictureBox1.Image)
            worksheet1.Paste()
            Clipboard.Clear()
            Clipboard.SetImage(IMAGE_STR.PictureBox2.Image)
            worksheet2.Paste()
            Clipboard.SetImage(IMAGE_STR.PictureBox3.Image)
            worksheet3.Paste()
            With worksheet1.PageSetup
                .Orientation = Excel.XlPageOrientation.xlLandscape
                .Zoom = False
                .FitToPagesTall = 1
                .FitToPagesWide = 1
            End With
            With worksheet2.PageSetup
                .Orientation = Excel.XlPageOrientation.xlLandscape
                .Zoom = False
                .FitToPagesTall = 1
                .FitToPagesWide = 1
            End With
            With worksheet3.PageSetup
                .Orientation = Excel.XlPageOrientation.xlLandscape
                .Zoom = False
                .FitToPagesTall = 1
                .FitToPagesWide = 1
            End With
            workbook.ExportAsFixedFormat(Excel.XlFixedFormatType.xlTypePDF, CONFIG.TextBox83.Text.ToString & SerialNumber & "_" & TestMode & "_" & Status & Format(Now(), "hhmmssmmddyy") & ".pdf", vbNormal)
            workbook.Worksheets("1").Delete()
            workbook.Worksheets("2").Delete()
            workbook.Worksheets("3").Delete()
            workbook.Close(True)
            System.Runtime.InteropServices.Marshal.ReleaseComObject(workbook)
            workbook = Nothing
            xlapp.Quit()
            System.Runtime.InteropServices.Marshal.ReleaseComObject(xlapp)
            xlapp = Nothing
            Process.GetProcessesByName("EXCEL")(0).Kill()
            file1.Close()
            MAIN.ListBox1.Items.Add($"{DateTime.Now.ToString} Test Report Generated success : " & CONFIG.TextBox83.Text.ToString & SerialNumber & "_" & TestMode & "_" & Status & Format(Now(), "hhmmssmmddyy"))
            MAIN.ListBox1.SelectedIndex = MAIN.ListBox1.Items.Count - 1
            Dim SW As IO.StreamWriter = IO.File.CreateText(CONFIG.TextBox83.Text.ToString & SerialNumber)
            For Each S As String In MAIN.ListBox1.Items
                SW.WriteLine(S, True)
            Next
            SW.Close()
            file1.Close()
        Catch ex As Exception
            Dim iret As Object = MsgBox("Error generating Report : " & ex.Message, vbCritical + vbRetryCancel, "Error executing SMC.GenerateReportOnerrorInit")
            If iret = vbRetry Then
                file1.Close()
                GoTo line1
            ElseIf iret = vbCancel Then
                file1.WriteLine("Error generating Report :  " & ex.Message & " " & "@" & DateTime.Now.ToString("yyyy/MM/dd/HH:mm:ss"))
                file1.Close()
                MAIN.ListBox1.Items.Add($"{DateTime.Now.ToString} Error generating Report :" & ex.Message)
                MAIN.ListBox1.SelectedIndex = MAIN.ListBox1.Items.Count - 1
                Exit Function
            End If
        End Try
    End Function
    Function GenerateReportOnErrorSMCSwitchTestNode2(User As String, TestMode As String, SerialNumber As String, Status As String, CycleTime As String, TesterID As String, Mark As String, ErrorMsg As String, DutConnect As String, SetPower As String, UniflashProgramme As String, FTDIProg As String, InitialPower As String, Can6705v As String, Isolatedv12 As String, Sensorv12 As String, TT6205v As String, IT6615v As String, A0691 As String, D0691 As String, D0691ON As String, D0661ON As String, D0661OFF As String)
        Dim line1 As Label
line1:
        Dim path1 As String = CONFIG.TextBox4.Text
        file1 = My.Computer.FileSystem.OpenTextFileWriter(path1, True)
        Try
            Dim xlapp = New Microsoft.Office.Interop.Excel.Application
            xlapp.DisplayAlerts = False
            Dim workbook = xlapp.Workbooks.Open(CONFIG.TextBox82.Text.ToString)
            Dim worksheet As Excel.Worksheet
            worksheet = workbook.Sheets("SMC")
            worksheet.Range("D6").Value = User
            worksheet.Range("D7").Value = TestMode
            worksheet.Range("D8").Value = SerialNumber
            worksheet.Range("D9").Value = DateTime.Now.ToString
            worksheet.Range("D10").Value = Status
            worksheet.Range("F6").Value = CycleTime
            worksheet.Range("F7").Value = TesterID
            worksheet.Range("F8").Value = Mark
            worksheet.Range("G7").Value = ErrorMsg
            worksheet.Range("G13").Value = DutConnect
            worksheet.Range("G14").Value = SetPower
            worksheet.Range("G15").Value = UniflashProgramme
            worksheet.Range("G16").Value = FTDIProg
            worksheet.Range("G18").Value = InitialPower
            worksheet.Range("G19").Value = Can6705v
            worksheet.Range("G20").Value = Isolatedv12
            worksheet.Range("G21").Value = Sensorv12
            worksheet.Range("G22").Value = TT6205v
            worksheet.Range("G23").Value = IT6615v
            worksheet.Range("G24").Value = A0691
            worksheet.Range("G25").Value = D0691
            worksheet.Range("G27").Value = D0691ON
            worksheet.Range("G28").Value = D0661ON
            worksheet.Range("G29").Value = D0661OFF
            Dim worksheet1 As Excel.Worksheet = workbook.Sheets.Add(After:=worksheet)
            worksheet1.Name = "1"
            Dim worksheet2 As Excel.Worksheet = workbook.Sheets.Add(After:=worksheet1)
            worksheet2.Name = "2"
            Dim worksheet3 As Excel.Worksheet = workbook.Sheets.Add(After:=worksheet2)
            worksheet3.Name = "3"
            Dim worksheet4 As Excel.Worksheet = workbook.Sheets.Add(After:=worksheet3)
            worksheet4.Name = "4"
            Clipboard.Clear()
            Clipboard.SetImage(IMAGE_STR.PictureBox1.Image)
            worksheet1.Paste()
            Clipboard.Clear()
            Clipboard.SetImage(IMAGE_STR.PictureBox2.Image)
            worksheet2.Paste()
            Clipboard.Clear()
            Clipboard.SetImage(IMAGE_STR.PictureBox3.Image)
            worksheet3.Paste()
            Clipboard.Clear()
            Clipboard.SetImage(IMAGE_STR.PictureBox4.Image)
            worksheet4.Paste()
            Clipboard.Clear()
            With worksheet1.PageSetup
                .Orientation = Excel.XlPageOrientation.xlLandscape
                .Zoom = False
                .FitToPagesTall = 1
                .FitToPagesWide = 1
            End With
            With worksheet2.PageSetup
                .Orientation = Excel.XlPageOrientation.xlLandscape
                .Zoom = False
                .FitToPagesTall = 1
                .FitToPagesWide = 1
            End With
            With worksheet3.PageSetup
                .Orientation = Excel.XlPageOrientation.xlLandscape
                .Zoom = False
                .FitToPagesTall = 1
                .FitToPagesWide = 1
            End With
            With worksheet4.PageSetup
                .Orientation = Excel.XlPageOrientation.xlLandscape
                .Zoom = False
                .FitToPagesTall = 1
                .FitToPagesWide = 1
            End With
            workbook.ExportAsFixedFormat(Excel.XlFixedFormatType.xlTypePDF, CONFIG.TextBox83.Text.ToString & SerialNumber & "_" & TestMode & "_" & Status & Format(Now(), "hhmmssmmddyy") & ".pdf", vbNormal)
            workbook.Worksheets("1").Delete()
            workbook.Worksheets("2").Delete()
            workbook.Worksheets("3").Delete()
            workbook.Worksheets("4").Delete()
            workbook.Close(True)
            System.Runtime.InteropServices.Marshal.ReleaseComObject(workbook)
            workbook = Nothing
            xlapp.Quit()
            System.Runtime.InteropServices.Marshal.ReleaseComObject(xlapp)
            xlapp = Nothing
            Process.GetProcessesByName("EXCEL")(0).Kill()
            file1.Close()
            MAIN.ListBox1.Items.Add($"{DateTime.Now.ToString} Test Report Generated success : " & CONFIG.TextBox83.Text.ToString & SerialNumber & "_" & TestMode & "_" & Status & Format(Now(), "hhmmssmmddyy"))
            MAIN.ListBox1.SelectedIndex = MAIN.ListBox1.Items.Count - 1
            Dim SW As IO.StreamWriter = IO.File.CreateText(CONFIG.TextBox83.Text.ToString & SerialNumber)
            For Each S As String In MAIN.ListBox1.Items
                SW.WriteLine(S, True)
            Next
            SW.Close()
            file1.Close()
        Catch ex As Exception
            Dim iret As Object = MsgBox("Error generating Report : " & ex.Message, vbCritical + vbRetryCancel, "Error executing SMC.GenerateReportOnerrorInit")
            If iret = vbRetry Then
                file1.Close()
                GoTo line1
            ElseIf iret = vbCancel Then
                file1.WriteLine("Error generating Report :  " & ex.Message & " " & "@" & DateTime.Now.ToString("yyyy/MM/dd/HH:mm:ss"))
                file1.Close()
                MAIN.ListBox1.Items.Add($"{DateTime.Now.ToString} Error generating Report :" & ex.Message)
                MAIN.ListBox1.SelectedIndex = MAIN.ListBox1.Items.Count - 1
                Exit Function
            End If
        End Try
    End Function
    Function GenerateReportOnErrorSMCSwitchTestNode3(User As String, TestMode As String, SerialNumber As String, Status As String, CycleTime As String, TesterID As String, Mark As String, ErrorMsg As String, DutConnect As String, SetPower As String, UniflashProgramme As String, FTDIProg As String, InitialPower As String, Can6705v As String, Isolatedv12 As String, Sensorv12 As String, TT6205v As String, IT6615v As String, A0691 As String, D0691 As String, D0691ON As String, D0661ON As String, D0661OFF As String, AUXENCMDON As String)
        Dim line1 As Label
line1:
        Dim path1 As String = CONFIG.TextBox4.Text
        file1 = My.Computer.FileSystem.OpenTextFileWriter(path1, True)
        Try
            Dim xlapp = New Microsoft.Office.Interop.Excel.Application
            xlapp.DisplayAlerts = False
            Dim workbook = xlapp.Workbooks.Open(CONFIG.TextBox82.Text.ToString)
            Dim worksheet As Excel.Worksheet
            worksheet = workbook.Sheets("SMC")
            worksheet.Range("D6").Value = User
            worksheet.Range("D7").Value = TestMode
            worksheet.Range("D8").Value = SerialNumber
            worksheet.Range("D9").Value = DateTime.Now.ToString
            worksheet.Range("D10").Value = Status
            worksheet.Range("F6").Value = CycleTime
            worksheet.Range("F7").Value = TesterID
            worksheet.Range("F8").Value = Mark
            worksheet.Range("G7").Value = ErrorMsg
            worksheet.Range("G13").Value = DutConnect
            worksheet.Range("G14").Value = SetPower
            worksheet.Range("G15").Value = UniflashProgramme
            worksheet.Range("G16").Value = FTDIProg
            worksheet.Range("G18").Value = InitialPower
            worksheet.Range("G19").Value = Can6705v
            worksheet.Range("G20").Value = Isolatedv12
            worksheet.Range("G21").Value = Sensorv12
            worksheet.Range("G22").Value = TT6205v
            worksheet.Range("G23").Value = IT6615v
            worksheet.Range("G24").Value = A0691
            worksheet.Range("G25").Value = D0691
            worksheet.Range("G27").Value = D0691ON
            worksheet.Range("G28").Value = D0661ON
            worksheet.Range("G29").Value = D0661OFF
            worksheet.Range("G30").Value = AUXENCMDON
            Dim worksheet1 As Excel.Worksheet = workbook.Sheets.Add(After:=worksheet)
            worksheet1.Name = "1"
            Dim worksheet2 As Excel.Worksheet = workbook.Sheets.Add(After:=worksheet1)
            worksheet2.Name = "2"
            Dim worksheet3 As Excel.Worksheet = workbook.Sheets.Add(After:=worksheet2)
            worksheet3.Name = "3"
            Dim worksheet4 As Excel.Worksheet = workbook.Sheets.Add(After:=worksheet3)
            worksheet4.Name = "4"
            Dim worksheet5 As Excel.Worksheet = workbook.Sheets.Add(After:=worksheet4)
            worksheet5.Name = "5"
            Clipboard.Clear()
            Clipboard.SetImage(IMAGE_STR.PictureBox1.Image)
            worksheet1.Paste()
            Clipboard.Clear()
            Clipboard.SetImage(IMAGE_STR.PictureBox2.Image)
            worksheet2.Paste()
            Clipboard.Clear()
            Clipboard.SetImage(IMAGE_STR.PictureBox3.Image)
            worksheet3.Paste()
            Clipboard.Clear()
            Clipboard.SetImage(IMAGE_STR.PictureBox4.Image)
            worksheet4.Paste()
            Clipboard.Clear()
            Clipboard.SetImage(IMAGE_STR.PictureBox5.Image)
            worksheet5.Paste()
            Clipboard.Clear()
            With worksheet1.PageSetup
                .Orientation = Excel.XlPageOrientation.xlLandscape
                .Zoom = False
                .FitToPagesTall = 1
                .FitToPagesWide = 1
            End With
            With worksheet2.PageSetup
                .Orientation = Excel.XlPageOrientation.xlLandscape
                .Zoom = False
                .FitToPagesTall = 1
                .FitToPagesWide = 1
            End With
            With worksheet3.PageSetup
                .Orientation = Excel.XlPageOrientation.xlLandscape
                .Zoom = False
                .FitToPagesTall = 1
                .FitToPagesWide = 1
            End With
            With worksheet4.PageSetup
                .Orientation = Excel.XlPageOrientation.xlLandscape
                .Zoom = False
                .FitToPagesTall = 1
                .FitToPagesWide = 1
            End With
            With worksheet5.PageSetup
                .Orientation = Excel.XlPageOrientation.xlLandscape
                .Zoom = False
                .FitToPagesTall = 1
                .FitToPagesWide = 1
            End With
            workbook.ExportAsFixedFormat(Excel.XlFixedFormatType.xlTypePDF, CONFIG.TextBox83.Text.ToString & SerialNumber & "_" & TestMode & "_" & Status & Format(Now(), "hhmmssmmddyy") & ".pdf", vbNormal)
            workbook.Worksheets("1").Delete()
            workbook.Worksheets("2").Delete()
            workbook.Worksheets("3").Delete()
            workbook.Worksheets("4").Delete()
            workbook.Worksheets("5").Delete()
            workbook.Close(True)
            System.Runtime.InteropServices.Marshal.ReleaseComObject(workbook)
            workbook = Nothing
            xlapp.Quit()
            System.Runtime.InteropServices.Marshal.ReleaseComObject(xlapp)
            xlapp = Nothing
            Process.GetProcessesByName("EXCEL")(0).Kill()
            file1.Close()
            MAIN.ListBox1.Items.Add($"{DateTime.Now.ToString} Test Report Generated success : " & CONFIG.TextBox83.Text.ToString & SerialNumber & "_" & TestMode & "_" & Status & Format(Now(), "hhmmssmmddyy"))
            MAIN.ListBox1.SelectedIndex = MAIN.ListBox1.Items.Count - 1
            Dim SW As IO.StreamWriter = IO.File.CreateText(CONFIG.TextBox83.Text.ToString & SerialNumber)
            For Each S As String In MAIN.ListBox1.Items
                SW.WriteLine(S, True)
            Next
            SW.Close()
            file1.Close()
        Catch ex As Exception
            Dim iret As Object = MsgBox("Error generating Report : " & ex.Message, vbCritical + vbRetryCancel, "Error executing SMC.GenerateReportOnerrorInit")
            If iret = vbRetry Then
                file1.Close()
                GoTo line1
            ElseIf iret = vbCancel Then
                file1.WriteLine("Error generating Report :  " & ex.Message & " " & "@" & DateTime.Now.ToString("yyyy/MM/dd/HH:mm:ss"))
                file1.Close()
                MAIN.ListBox1.Items.Add($"{DateTime.Now.ToString} Error generating Report :" & ex.Message)
                MAIN.ListBox1.SelectedIndex = MAIN.ListBox1.Items.Count - 1
                Exit Function
            End If
        End Try
    End Function
    Function GenerateReportOnErrorSMCSwitchTestNode4(User As String, TestMode As String, SerialNumber As String, Status As String, CycleTime As String, TesterID As String, Mark As String, ErrorMsg As String, DutConnect As String, SetPower As String, UniflashProgramme As String, FTDIProg As String, InitialPower As String, Can6705v As String, Isolatedv12 As String, Sensorv12 As String, TT6205v As String, IT6615v As String, A0691 As String, D0691 As String, D0691ON As String, D0661ON As String, D0661OFF As String, AUXENCMDON As String, D0100ON As String)
        Dim line1 As Label
line1:
        Dim path1 As String = CONFIG.TextBox4.Text
        file1 = My.Computer.FileSystem.OpenTextFileWriter(path1, True)
        Try
            Dim xlapp = New Microsoft.Office.Interop.Excel.Application
            xlapp.DisplayAlerts = False
            Dim workbook = xlapp.Workbooks.Open(CONFIG.TextBox82.Text.ToString)
            Dim worksheet As Excel.Worksheet
            worksheet = workbook.Sheets("SMC")
            worksheet.Range("D6").Value = User
            worksheet.Range("D7").Value = TestMode
            worksheet.Range("D8").Value = SerialNumber
            worksheet.Range("D9").Value = DateTime.Now.ToString
            worksheet.Range("D10").Value = Status
            worksheet.Range("F6").Value = CycleTime
            worksheet.Range("F7").Value = TesterID
            worksheet.Range("F8").Value = Mark
            worksheet.Range("G7").Value = ErrorMsg
            worksheet.Range("G13").Value = DutConnect
            worksheet.Range("G14").Value = SetPower
            worksheet.Range("G15").Value = UniflashProgramme
            worksheet.Range("G16").Value = FTDIProg
            worksheet.Range("G18").Value = InitialPower
            worksheet.Range("G19").Value = Can6705v
            worksheet.Range("G20").Value = Isolatedv12
            worksheet.Range("G21").Value = Sensorv12
            worksheet.Range("G22").Value = TT6205v
            worksheet.Range("G23").Value = IT6615v
            worksheet.Range("G24").Value = A0691
            worksheet.Range("G25").Value = D0691
            worksheet.Range("G27").Value = D0691ON
            worksheet.Range("G28").Value = D0661ON
            worksheet.Range("G29").Value = D0661OFF
            worksheet.Range("G30").Value = AUXENCMDON
            worksheet.Range("G31").Value = D0100ON
            Dim worksheet1 As Excel.Worksheet = workbook.Sheets.Add(After:=worksheet)
            worksheet1.Name = "1"
            Dim worksheet2 As Excel.Worksheet = workbook.Sheets.Add(After:=worksheet1)
            worksheet2.Name = "2"
            Dim worksheet3 As Excel.Worksheet = workbook.Sheets.Add(After:=worksheet2)
            worksheet3.Name = "3"
            Dim worksheet4 As Excel.Worksheet = workbook.Sheets.Add(After:=worksheet3)
            worksheet4.Name = "4"
            Dim worksheet5 As Excel.Worksheet = workbook.Sheets.Add(After:=worksheet4)
            worksheet5.Name = "5"
            Dim worksheet6 As Excel.Worksheet = workbook.Sheets.Add(After:=worksheet5)
            worksheet6.Name = "6"
            Clipboard.Clear()
            Clipboard.SetImage(IMAGE_STR.PictureBox1.Image)
            worksheet1.Paste()
            Clipboard.Clear()
            Clipboard.SetImage(IMAGE_STR.PictureBox2.Image)
            worksheet2.Paste()
            Clipboard.Clear()
            Clipboard.SetImage(IMAGE_STR.PictureBox3.Image)
            worksheet3.Paste()
            Clipboard.Clear()
            Clipboard.SetImage(IMAGE_STR.PictureBox4.Image)
            worksheet4.Paste()
            Clipboard.Clear()
            Clipboard.SetImage(IMAGE_STR.PictureBox5.Image)
            worksheet5.Paste()
            Clipboard.Clear()
            Clipboard.SetImage(IMAGE_STR.PictureBox6.Image)
            worksheet6.Paste()
            Clipboard.Clear()
            With worksheet1.PageSetup
                .Orientation = Excel.XlPageOrientation.xlLandscape
                .Zoom = False
                .FitToPagesTall = 1
                .FitToPagesWide = 1
            End With
            With worksheet2.PageSetup
                .Orientation = Excel.XlPageOrientation.xlLandscape
                .Zoom = False
                .FitToPagesTall = 1
                .FitToPagesWide = 1
            End With
            With worksheet3.PageSetup
                .Orientation = Excel.XlPageOrientation.xlLandscape
                .Zoom = False
                .FitToPagesTall = 1
                .FitToPagesWide = 1
            End With
            With worksheet4.PageSetup
                .Orientation = Excel.XlPageOrientation.xlLandscape
                .Zoom = False
                .FitToPagesTall = 1
                .FitToPagesWide = 1
            End With
            With worksheet5.PageSetup
                .Orientation = Excel.XlPageOrientation.xlLandscape
                .Zoom = False
                .FitToPagesTall = 1
                .FitToPagesWide = 1
            End With
            With worksheet6.PageSetup
                .Orientation = Excel.XlPageOrientation.xlLandscape
                .Zoom = False
                .FitToPagesTall = 1
                .FitToPagesWide = 1
            End With
            workbook.ExportAsFixedFormat(Excel.XlFixedFormatType.xlTypePDF, CONFIG.TextBox83.Text.ToString & SerialNumber & "_" & TestMode & "_" & Status & Format(Now(), "hhmmssmmddyy") & ".pdf", vbNormal)
            workbook.Worksheets("1").Delete()
            workbook.Worksheets("2").Delete()
            workbook.Worksheets("3").Delete()
            workbook.Worksheets("4").Delete()
            workbook.Worksheets("5").Delete()
            workbook.Worksheets("6").Delete()
            workbook.Close(True)
            System.Runtime.InteropServices.Marshal.ReleaseComObject(workbook)
            workbook = Nothing
            xlapp.Quit()
            System.Runtime.InteropServices.Marshal.ReleaseComObject(xlapp)
            xlapp = Nothing
            Process.GetProcessesByName("EXCEL")(0).Kill()
            file1.Close()
            MAIN.ListBox1.Items.Add($"{DateTime.Now.ToString} Test Report Generated success : " & CONFIG.TextBox83.Text.ToString & SerialNumber & "_" & TestMode & "_" & Status & Format(Now(), "hhmmssmmddyy"))
            MAIN.ListBox1.SelectedIndex = MAIN.ListBox1.Items.Count - 1
            Dim SW As IO.StreamWriter = IO.File.CreateText(CONFIG.TextBox83.Text.ToString & SerialNumber)
            For Each S As String In MAIN.ListBox1.Items
                SW.WriteLine(S, True)
            Next
            SW.Close()
            file1.Close()
        Catch ex As Exception
            Dim iret As Object = MsgBox("Error generating Report : " & ex.Message, vbCritical + vbRetryCancel, "Error executing SMC.GenerateReportOnerrorInit")
            If iret = vbRetry Then
                file1.Close()
                GoTo line1
            ElseIf iret = vbCancel Then
                file1.WriteLine("Error generating Report :  " & ex.Message & " " & "@" & DateTime.Now.ToString("yyyy/MM/dd/HH:mm:ss"))
                file1.Close()
                MAIN.ListBox1.Items.Add($"{DateTime.Now.ToString} Error generating Report :" & ex.Message)
                MAIN.ListBox1.SelectedIndex = MAIN.ListBox1.Items.Count - 1
                Exit Function
            End If
        End Try
    End Function
    Function GenerateReportOnErrorSMCSwitchTestNode5(User As String, TestMode As String, SerialNumber As String, Status As String, CycleTime As String, TesterID As String, Mark As String, ErrorMsg As String, DutConnect As String, SetPower As String, UniflashProgramme As String, FTDIProg As String, InitialPower As String, Can6705v As String, Isolatedv12 As String, Sensorv12 As String, TT6205v As String, IT6615v As String, A0691 As String, D0691 As String, D0691ON As String, D0661ON As String, D0661OFF As String, AUXENCMDON As String, D0100ON As String, D0100OFF As String)
        Dim line1 As Label
line1:
        Dim path1 As String = CONFIG.TextBox4.Text
        file1 = My.Computer.FileSystem.OpenTextFileWriter(path1, True)
        Try
            Dim xlapp = New Microsoft.Office.Interop.Excel.Application
            xlapp.DisplayAlerts = False
            Dim workbook = xlapp.Workbooks.Open(CONFIG.TextBox82.Text.ToString)
            Dim worksheet As Excel.Worksheet
            worksheet = workbook.Sheets("SMC")
            worksheet.Range("D6").Value = User
            worksheet.Range("D7").Value = TestMode
            worksheet.Range("D8").Value = SerialNumber
            worksheet.Range("D9").Value = DateTime.Now.ToString
            worksheet.Range("D10").Value = Status
            worksheet.Range("F6").Value = CycleTime
            worksheet.Range("F7").Value = TesterID
            worksheet.Range("F8").Value = Mark
            worksheet.Range("G7").Value = ErrorMsg
            worksheet.Range("G13").Value = DutConnect
            worksheet.Range("G14").Value = SetPower
            worksheet.Range("G15").Value = UniflashProgramme
            worksheet.Range("G16").Value = FTDIProg
            worksheet.Range("G18").Value = InitialPower
            worksheet.Range("G19").Value = Can6705v
            worksheet.Range("G20").Value = Isolatedv12
            worksheet.Range("G21").Value = Sensorv12
            worksheet.Range("G22").Value = TT6205v
            worksheet.Range("G23").Value = IT6615v
            worksheet.Range("G24").Value = A0691
            worksheet.Range("G25").Value = D0691
            worksheet.Range("G27").Value = D0691ON
            worksheet.Range("G28").Value = D0661ON
            worksheet.Range("G29").Value = D0661OFF
            worksheet.Range("G30").Value = AUXENCMDON
            worksheet.Range("G31").Value = D0100ON
            worksheet.Range("G32").Value = D0100OFF
            Dim worksheet1 As Excel.Worksheet = workbook.Sheets.Add(After:=worksheet)
            worksheet1.Name = "1"
            Dim worksheet2 As Excel.Worksheet = workbook.Sheets.Add(After:=worksheet1)
            worksheet2.Name = "2"
            Dim worksheet3 As Excel.Worksheet = workbook.Sheets.Add(After:=worksheet2)
            worksheet3.Name = "3"
            Dim worksheet4 As Excel.Worksheet = workbook.Sheets.Add(After:=worksheet3)
            worksheet4.Name = "4"
            Dim worksheet5 As Excel.Worksheet = workbook.Sheets.Add(After:=worksheet4)
            worksheet5.Name = "5"
            Dim worksheet6 As Excel.Worksheet = workbook.Sheets.Add(After:=worksheet5)
            worksheet6.Name = "6"
            Dim worksheet7 As Excel.Worksheet = workbook.Sheets.Add(After:=worksheet6)
            worksheet7.Name = "7"
            Clipboard.Clear()
            Clipboard.SetImage(IMAGE_STR.PictureBox1.Image)
            worksheet1.Paste()
            Clipboard.Clear()
            Clipboard.SetImage(IMAGE_STR.PictureBox2.Image)
            worksheet2.Paste()
            Clipboard.Clear()
            Clipboard.SetImage(IMAGE_STR.PictureBox3.Image)
            worksheet3.Paste()
            Clipboard.Clear()
            Clipboard.SetImage(IMAGE_STR.PictureBox4.Image)
            worksheet4.Paste()
            Clipboard.Clear()
            Clipboard.SetImage(IMAGE_STR.PictureBox5.Image)
            worksheet5.Paste()
            Clipboard.Clear()
            Clipboard.SetImage(IMAGE_STR.PictureBox6.Image)
            worksheet6.Paste()
            Clipboard.Clear()
            Clipboard.SetImage(IMAGE_STR.PictureBox7.Image)
            worksheet7.Paste()
            Clipboard.Clear()
            With worksheet1.PageSetup
                .Orientation = Excel.XlPageOrientation.xlLandscape
                .Zoom = False
                .FitToPagesTall = 1
                .FitToPagesWide = 1
            End With
            With worksheet2.PageSetup
                .Orientation = Excel.XlPageOrientation.xlLandscape
                .Zoom = False
                .FitToPagesTall = 1
                .FitToPagesWide = 1
            End With
            With worksheet3.PageSetup
                .Orientation = Excel.XlPageOrientation.xlLandscape
                .Zoom = False
                .FitToPagesTall = 1
                .FitToPagesWide = 1
            End With
            With worksheet4.PageSetup
                .Orientation = Excel.XlPageOrientation.xlLandscape
                .Zoom = False
                .FitToPagesTall = 1
                .FitToPagesWide = 1
            End With
            With worksheet5.PageSetup
                .Orientation = Excel.XlPageOrientation.xlLandscape
                .Zoom = False
                .FitToPagesTall = 1
                .FitToPagesWide = 1
            End With
            With worksheet6.PageSetup
                .Orientation = Excel.XlPageOrientation.xlLandscape
                .Zoom = False
                .FitToPagesTall = 1
                .FitToPagesWide = 1
            End With
            With worksheet7.PageSetup
                .Orientation = Excel.XlPageOrientation.xlLandscape
                .Zoom = False
                .FitToPagesTall = 1
                .FitToPagesWide = 1
            End With
            workbook.ExportAsFixedFormat(Excel.XlFixedFormatType.xlTypePDF, CONFIG.TextBox83.Text.ToString & SerialNumber & "_" & TestMode & "_" & Status & Format(Now(), "hhmmssmmddyy") & ".pdf", vbNormal)
            workbook.Worksheets("1").Delete()
            workbook.Worksheets("2").Delete()
            workbook.Worksheets("3").Delete()
            workbook.Worksheets("4").Delete()
            workbook.Worksheets("5").Delete()
            workbook.Worksheets("6").Delete()
            workbook.Worksheets("7").Delete()
            workbook.Close(True)
            System.Runtime.InteropServices.Marshal.ReleaseComObject(workbook)
            workbook = Nothing
            xlapp.Quit()
            System.Runtime.InteropServices.Marshal.ReleaseComObject(xlapp)
            xlapp = Nothing
            Process.GetProcessesByName("EXCEL")(0).Kill()
            file1.Close()
            MAIN.ListBox1.Items.Add($"{DateTime.Now.ToString} Test Report Generated success : " & CONFIG.TextBox83.Text.ToString & SerialNumber & "_" & TestMode & "_" & Status & Format(Now(), "hhmmssmmddyy"))
            MAIN.ListBox1.SelectedIndex = MAIN.ListBox1.Items.Count - 1
            Dim SW As IO.StreamWriter = IO.File.CreateText(CONFIG.TextBox83.Text.ToString & SerialNumber)
            For Each S As String In MAIN.ListBox1.Items
                SW.WriteLine(S, True)
            Next
            SW.Close()
            file1.Close()
        Catch ex As Exception
            Dim iret As Object = MsgBox("Error generating Report : " & ex.Message, vbCritical + vbRetryCancel, "Error executing SMC.GenerateReportOnerrorInit")
            If iret = vbRetry Then
                file1.Close()
                GoTo line1
            ElseIf iret = vbCancel Then
                file1.WriteLine("Error generating Report :  " & ex.Message & " " & "@" & DateTime.Now.ToString("yyyy/MM/dd/HH:mm:ss"))
                file1.Close()
                MAIN.ListBox1.Items.Add($"{DateTime.Now.ToString} Error generating Report :" & ex.Message)
                MAIN.ListBox1.SelectedIndex = MAIN.ListBox1.Items.Count - 1
                Exit Function
            End If
        End Try
    End Function
    Function GenerateReportOnErrorSMCSwitchTestNode6(User As String, TestMode As String, SerialNumber As String, Status As String, CycleTime As String, TesterID As String, Mark As String, ErrorMsg As String, DutConnect As String, SetPower As String, UniflashProgramme As String, FTDIProg As String, InitialPower As String, Can6705v As String, Isolatedv12 As String, Sensorv12 As String, TT6205v As String, IT6615v As String, A0691 As String, D0691 As String, D0691ON As String, D0661ON As String, D0661OFF As String, AUXENCMDON As String, D0100ON As String, D0100OFF As String, D0101ON As String)
        Dim line1 As Label
line1:
        Dim path1 As String = CONFIG.TextBox4.Text
        file1 = My.Computer.FileSystem.OpenTextFileWriter(path1, True)
        Try
            Dim xlapp = New Microsoft.Office.Interop.Excel.Application
            xlapp.DisplayAlerts = False
            Dim workbook = xlapp.Workbooks.Open(CONFIG.TextBox82.Text.ToString)
            Dim worksheet As Excel.Worksheet
            worksheet = workbook.Sheets("SMC")
            worksheet.Range("D6").Value = User
            worksheet.Range("D7").Value = TestMode
            worksheet.Range("D8").Value = SerialNumber
            worksheet.Range("D9").Value = DateTime.Now.ToString
            worksheet.Range("D10").Value = Status
            worksheet.Range("F6").Value = CycleTime
            worksheet.Range("F7").Value = TesterID
            worksheet.Range("F8").Value = Mark
            worksheet.Range("G7").Value = ErrorMsg
            worksheet.Range("G13").Value = DutConnect
            worksheet.Range("G14").Value = SetPower
            worksheet.Range("G15").Value = UniflashProgramme
            worksheet.Range("G16").Value = FTDIProg
            worksheet.Range("G18").Value = InitialPower
            worksheet.Range("G19").Value = Can6705v
            worksheet.Range("G20").Value = Isolatedv12
            worksheet.Range("G21").Value = Sensorv12
            worksheet.Range("G22").Value = TT6205v
            worksheet.Range("G23").Value = IT6615v
            worksheet.Range("G24").Value = A0691
            worksheet.Range("G25").Value = D0691
            worksheet.Range("G27").Value = D0691ON
            worksheet.Range("G28").Value = D0661ON
            worksheet.Range("G29").Value = D0661OFF
            worksheet.Range("G30").Value = AUXENCMDON
            worksheet.Range("G31").Value = D0100ON
            worksheet.Range("G32").Value = D0100OFF
            worksheet.Range("G33").Value = D0101ON
            Dim worksheet1 As Excel.Worksheet = workbook.Sheets.Add(After:=worksheet)
            worksheet1.Name = "1"
            Dim worksheet2 As Excel.Worksheet = workbook.Sheets.Add(After:=worksheet1)
            worksheet2.Name = "2"
            Dim worksheet3 As Excel.Worksheet = workbook.Sheets.Add(After:=worksheet2)
            worksheet3.Name = "3"
            Dim worksheet4 As Excel.Worksheet = workbook.Sheets.Add(After:=worksheet3)
            worksheet4.Name = "4"
            Dim worksheet5 As Excel.Worksheet = workbook.Sheets.Add(After:=worksheet4)
            worksheet5.Name = "5"
            Dim worksheet6 As Excel.Worksheet = workbook.Sheets.Add(After:=worksheet5)
            worksheet6.Name = "6"
            Dim worksheet7 As Excel.Worksheet = workbook.Sheets.Add(After:=worksheet6)
            worksheet7.Name = "7"
            Dim worksheet8 As Excel.Worksheet = workbook.Sheets.Add(After:=worksheet7)
            worksheet8.Name = "8"
            Clipboard.Clear()
            Clipboard.SetImage(IMAGE_STR.PictureBox1.Image)
            worksheet1.Paste()
            Clipboard.Clear()
            Clipboard.SetImage(IMAGE_STR.PictureBox2.Image)
            worksheet2.Paste()
            Clipboard.Clear()
            Clipboard.SetImage(IMAGE_STR.PictureBox3.Image)
            worksheet3.Paste()
            Clipboard.Clear()
            Clipboard.SetImage(IMAGE_STR.PictureBox4.Image)
            worksheet4.Paste()
            Clipboard.Clear()
            Clipboard.SetImage(IMAGE_STR.PictureBox5.Image)
            worksheet5.Paste()
            Clipboard.Clear()
            Clipboard.SetImage(IMAGE_STR.PictureBox6.Image)
            worksheet6.Paste()
            Clipboard.Clear()
            Clipboard.SetImage(IMAGE_STR.PictureBox7.Image)
            worksheet7.Paste()
            Clipboard.Clear()
            Clipboard.SetImage(IMAGE_STR.PictureBox8.Image)
            worksheet8.Paste()
            Clipboard.Clear()
            With worksheet1.PageSetup
                .Orientation = Excel.XlPageOrientation.xlLandscape
                .Zoom = False
                .FitToPagesTall = 1
                .FitToPagesWide = 1
            End With
            With worksheet2.PageSetup
                .Orientation = Excel.XlPageOrientation.xlLandscape
                .Zoom = False
                .FitToPagesTall = 1
                .FitToPagesWide = 1
            End With
            With worksheet3.PageSetup
                .Orientation = Excel.XlPageOrientation.xlLandscape
                .Zoom = False
                .FitToPagesTall = 1
                .FitToPagesWide = 1
            End With
            With worksheet4.PageSetup
                .Orientation = Excel.XlPageOrientation.xlLandscape
                .Zoom = False
                .FitToPagesTall = 1
                .FitToPagesWide = 1
            End With
            With worksheet5.PageSetup
                .Orientation = Excel.XlPageOrientation.xlLandscape
                .Zoom = False
                .FitToPagesTall = 1
                .FitToPagesWide = 1
            End With
            With worksheet6.PageSetup
                .Orientation = Excel.XlPageOrientation.xlLandscape
                .Zoom = False
                .FitToPagesTall = 1
                .FitToPagesWide = 1
            End With
            With worksheet7.PageSetup
                .Orientation = Excel.XlPageOrientation.xlLandscape
                .Zoom = False
                .FitToPagesTall = 1
                .FitToPagesWide = 1
            End With
            With worksheet8.PageSetup
                .Orientation = Excel.XlPageOrientation.xlLandscape
                .Zoom = False
                .FitToPagesTall = 1
                .FitToPagesWide = 1
            End With
            workbook.ExportAsFixedFormat(Excel.XlFixedFormatType.xlTypePDF, CONFIG.TextBox83.Text.ToString & SerialNumber & "_" & TestMode & "_" & Status & Format(Now(), "hhmmssmmddyy") & ".pdf", vbNormal)
            workbook.Worksheets("1").Delete()
            workbook.Worksheets("2").Delete()
            workbook.Worksheets("3").Delete()
            workbook.Worksheets("4").Delete()
            workbook.Worksheets("5").Delete()
            workbook.Worksheets("6").Delete()
            workbook.Worksheets("7").Delete()
            workbook.Worksheets("8").Delete()
            workbook.Close(True)
            System.Runtime.InteropServices.Marshal.ReleaseComObject(workbook)
            workbook = Nothing
            xlapp.Quit()
            System.Runtime.InteropServices.Marshal.ReleaseComObject(xlapp)
            xlapp = Nothing
            Process.GetProcessesByName("EXCEL")(0).Kill()
            file1.Close()
            MAIN.ListBox1.Items.Add($"{DateTime.Now.ToString} Test Report Generated success : " & CONFIG.TextBox83.Text.ToString & SerialNumber & "_" & TestMode & "_" & Status & Format(Now(), "hhmmssmmddyy"))
            MAIN.ListBox1.SelectedIndex = MAIN.ListBox1.Items.Count - 1
            Dim SW As IO.StreamWriter = IO.File.CreateText(CONFIG.TextBox83.Text.ToString & SerialNumber)
            For Each S As String In MAIN.ListBox1.Items
                SW.WriteLine(S, True)
            Next
            SW.Close()
            file1.Close()
        Catch ex As Exception
            Dim iret As Object = MsgBox("Error generating Report : " & ex.Message, vbCritical + vbRetryCancel, "Error executing SMC.GenerateReportOnerrorInit")
            If iret = vbRetry Then
                file1.Close()
                GoTo line1
            ElseIf iret = vbCancel Then
                file1.WriteLine("Error generating Report :  " & ex.Message & " " & "@" & DateTime.Now.ToString("yyyy/MM/dd/HH:mm:ss"))
                file1.Close()
                MAIN.ListBox1.Items.Add($"{DateTime.Now.ToString} Error generating Report :" & ex.Message)
                MAIN.ListBox1.SelectedIndex = MAIN.ListBox1.Items.Count - 1
                Exit Function
            End If
        End Try
    End Function
    Function GenerateReportOnErrorSMCSwitchTestNode7(User As String, TestMode As String, SerialNumber As String, Status As String, CycleTime As String, TesterID As String, Mark As String, ErrorMsg As String, DutConnect As String, SetPower As String, UniflashProgramme As String, FTDIProg As String, InitialPower As String, Can6705v As String, Isolatedv12 As String, Sensorv12 As String, TT6205v As String, IT6615v As String, A0691 As String, D0691 As String, D0691ON As String, D0661ON As String, D0661OFF As String, AUXENCMDON As String, D0100ON As String, D0100OFF As String, D0101ON As String, D0101OFF As String)
        Dim line1 As Label
line1:
        Dim path1 As String = CONFIG.TextBox4.Text
        file1 = My.Computer.FileSystem.OpenTextFileWriter(path1, True)
        Try
            Dim xlapp = New Microsoft.Office.Interop.Excel.Application
            xlapp.DisplayAlerts = False
            Dim workbook = xlapp.Workbooks.Open(CONFIG.TextBox82.Text.ToString)
            Dim worksheet As Excel.Worksheet
            worksheet = workbook.Sheets("SMC")
            worksheet.Range("D6").Value = User
            worksheet.Range("D7").Value = TestMode
            worksheet.Range("D8").Value = SerialNumber
            worksheet.Range("D9").Value = DateTime.Now.ToString
            worksheet.Range("D10").Value = Status
            worksheet.Range("F6").Value = CycleTime
            worksheet.Range("F7").Value = TesterID
            worksheet.Range("F8").Value = Mark
            worksheet.Range("G7").Value = ErrorMsg
            worksheet.Range("G13").Value = DutConnect
            worksheet.Range("G14").Value = SetPower
            worksheet.Range("G15").Value = UniflashProgramme
            worksheet.Range("G16").Value = FTDIProg
            worksheet.Range("G18").Value = InitialPower
            worksheet.Range("G19").Value = Can6705v
            worksheet.Range("G20").Value = Isolatedv12
            worksheet.Range("G21").Value = Sensorv12
            worksheet.Range("G22").Value = TT6205v
            worksheet.Range("G23").Value = IT6615v
            worksheet.Range("G24").Value = A0691
            worksheet.Range("G25").Value = D0691
            worksheet.Range("G27").Value = D0691ON
            worksheet.Range("G28").Value = D0661ON
            worksheet.Range("G29").Value = D0661OFF
            worksheet.Range("G30").Value = AUXENCMDON
            worksheet.Range("G31").Value = D0100ON
            worksheet.Range("G32").Value = D0100OFF
            worksheet.Range("G33").Value = D0101ON
            worksheet.Range("G34").Value = D0101OFF
            Dim worksheet1 As Excel.Worksheet = workbook.Sheets.Add(After:=worksheet)
            worksheet1.Name = "1"
            Dim worksheet2 As Excel.Worksheet = workbook.Sheets.Add(After:=worksheet1)
            worksheet2.Name = "2"
            Dim worksheet3 As Excel.Worksheet = workbook.Sheets.Add(After:=worksheet2)
            worksheet3.Name = "3"
            Dim worksheet4 As Excel.Worksheet = workbook.Sheets.Add(After:=worksheet3)
            worksheet4.Name = "4"
            Dim worksheet5 As Excel.Worksheet = workbook.Sheets.Add(After:=worksheet4)
            worksheet5.Name = "5"
            Dim worksheet6 As Excel.Worksheet = workbook.Sheets.Add(After:=worksheet5)
            worksheet6.Name = "6"
            Dim worksheet7 As Excel.Worksheet = workbook.Sheets.Add(After:=worksheet6)
            worksheet7.Name = "7"
            Dim worksheet8 As Excel.Worksheet = workbook.Sheets.Add(After:=worksheet7)
            worksheet8.Name = "8"
            Dim worksheet9 As Excel.Worksheet = workbook.Sheets.Add(After:=worksheet8)
            worksheet9.Name = "9"
            Clipboard.Clear()
            Clipboard.SetImage(IMAGE_STR.PictureBox1.Image)
            worksheet1.Paste()
            Clipboard.Clear()
            Clipboard.SetImage(IMAGE_STR.PictureBox2.Image)
            worksheet2.Paste()
            Clipboard.Clear()
            Clipboard.SetImage(IMAGE_STR.PictureBox3.Image)
            worksheet3.Paste()
            Clipboard.Clear()
            Clipboard.SetImage(IMAGE_STR.PictureBox4.Image)
            worksheet4.Paste()
            Clipboard.Clear()
            Clipboard.SetImage(IMAGE_STR.PictureBox5.Image)
            worksheet5.Paste()
            Clipboard.Clear()
            Clipboard.SetImage(IMAGE_STR.PictureBox6.Image)
            worksheet6.Paste()
            Clipboard.Clear()
            Clipboard.SetImage(IMAGE_STR.PictureBox7.Image)
            worksheet7.Paste()
            Clipboard.Clear()
            Clipboard.SetImage(IMAGE_STR.PictureBox8.Image)
            worksheet8.Paste()
            Clipboard.Clear()
            Clipboard.SetImage(IMAGE_STR.PictureBox9.Image)
            worksheet9.Paste()
            Clipboard.Clear()
            With worksheet1.PageSetup
                .Orientation = Excel.XlPageOrientation.xlLandscape
                .Zoom = False
                .FitToPagesTall = 1
                .FitToPagesWide = 1
            End With
            With worksheet2.PageSetup
                .Orientation = Excel.XlPageOrientation.xlLandscape
                .Zoom = False
                .FitToPagesTall = 1
                .FitToPagesWide = 1
            End With
            With worksheet3.PageSetup
                .Orientation = Excel.XlPageOrientation.xlLandscape
                .Zoom = False
                .FitToPagesTall = 1
                .FitToPagesWide = 1
            End With
            With worksheet4.PageSetup
                .Orientation = Excel.XlPageOrientation.xlLandscape
                .Zoom = False
                .FitToPagesTall = 1
                .FitToPagesWide = 1
            End With
            With worksheet5.PageSetup
                .Orientation = Excel.XlPageOrientation.xlLandscape
                .Zoom = False
                .FitToPagesTall = 1
                .FitToPagesWide = 1
            End With
            With worksheet6.PageSetup
                .Orientation = Excel.XlPageOrientation.xlLandscape
                .Zoom = False
                .FitToPagesTall = 1
                .FitToPagesWide = 1
            End With
            With worksheet7.PageSetup
                .Orientation = Excel.XlPageOrientation.xlLandscape
                .Zoom = False
                .FitToPagesTall = 1
                .FitToPagesWide = 1
            End With
            With worksheet8.PageSetup
                .Orientation = Excel.XlPageOrientation.xlLandscape
                .Zoom = False
                .FitToPagesTall = 1
                .FitToPagesWide = 1
            End With
            With worksheet9.PageSetup
                .Orientation = Excel.XlPageOrientation.xlLandscape
                .Zoom = False
                .FitToPagesTall = 1
                .FitToPagesWide = 1
            End With
            workbook.ExportAsFixedFormat(Excel.XlFixedFormatType.xlTypePDF, CONFIG.TextBox83.Text.ToString & SerialNumber & "_" & TestMode & "_" & Status & Format(Now(), "hhmmssmmddyy") & ".pdf", vbNormal)
            workbook.Worksheets("1").Delete()
            workbook.Worksheets("2").Delete()
            workbook.Worksheets("3").Delete()
            workbook.Worksheets("4").Delete()
            workbook.Worksheets("5").Delete()
            workbook.Worksheets("6").Delete()
            workbook.Worksheets("7").Delete()
            workbook.Worksheets("8").Delete()
            workbook.Worksheets("9").Delete()
            workbook.Close(True)
            System.Runtime.InteropServices.Marshal.ReleaseComObject(workbook)
            workbook = Nothing
            xlapp.Quit()
            System.Runtime.InteropServices.Marshal.ReleaseComObject(xlapp)
            xlapp = Nothing
            Process.GetProcessesByName("EXCEL")(0).Kill()
            file1.Close()
            MAIN.ListBox1.Items.Add($"{DateTime.Now.ToString} Test Report Generated success : " & CONFIG.TextBox83.Text.ToString & SerialNumber & "_" & TestMode & "_" & Status & Format(Now(), "hhmmssmmddyy"))
            MAIN.ListBox1.SelectedIndex = MAIN.ListBox1.Items.Count - 1
            Dim SW As IO.StreamWriter = IO.File.CreateText(CONFIG.TextBox83.Text.ToString & SerialNumber)
            For Each S As String In MAIN.ListBox1.Items
                SW.WriteLine(S, True)
            Next
            SW.Close()
            file1.Close()
        Catch ex As Exception
            Dim iret As Object = MsgBox("Error generating Report : " & ex.Message, vbCritical + vbRetryCancel, "Error executing SMC.GenerateReportOnerrorInit")
            If iret = vbRetry Then
                file1.Close()
                GoTo line1
            ElseIf iret = vbCancel Then
                file1.WriteLine("Error generating Report :  " & ex.Message & " " & "@" & DateTime.Now.ToString("yyyy/MM/dd/HH:mm:ss"))
                file1.Close()
                MAIN.ListBox1.Items.Add($"{DateTime.Now.ToString} Error generating Report :" & ex.Message)
                MAIN.ListBox1.SelectedIndex = MAIN.ListBox1.Items.Count - 1
                Exit Function
            End If
        End Try
    End Function
    Function GenerateReportOnErrorSMCSwitchTestNode8(User As String, TestMode As String, SerialNumber As String, Status As String, CycleTime As String, TesterID As String, Mark As String, ErrorMsg As String, DutConnect As String, SetPower As String, UniflashProgramme As String, FTDIProg As String, InitialPower As String, Can6705v As String, Isolatedv12 As String, Sensorv12 As String, TT6205v As String, IT6615v As String, A0691 As String, D0691 As String, D0691ON As String, D0661ON As String, D0661OFF As String, AUXENCMDON As String, D0100ON As String, D0100OFF As String, D0101ON As String, D0101OFF As String, CMDOFF As String)
        Dim line1 As Label
line1:
        Dim path1 As String = CONFIG.TextBox4.Text
        file1 = My.Computer.FileSystem.OpenTextFileWriter(path1, True)
        Try
            Dim xlapp = New Microsoft.Office.Interop.Excel.Application
            xlapp.DisplayAlerts = False
            Dim workbook = xlapp.Workbooks.Open(CONFIG.TextBox82.Text.ToString)
            Dim worksheet As Excel.Worksheet
            worksheet = workbook.Sheets("SMC")
            worksheet.Range("D6").Value = User
            worksheet.Range("D7").Value = TestMode
            worksheet.Range("D8").Value = SerialNumber
            worksheet.Range("D9").Value = DateTime.Now.ToString
            worksheet.Range("D10").Value = Status
            worksheet.Range("F6").Value = CycleTime
            worksheet.Range("F7").Value = TesterID
            worksheet.Range("F8").Value = Mark
            worksheet.Range("G7").Value = ErrorMsg
            worksheet.Range("G13").Value = DutConnect
            worksheet.Range("G14").Value = SetPower
            worksheet.Range("G15").Value = UniflashProgramme
            worksheet.Range("G16").Value = FTDIProg
            worksheet.Range("G18").Value = InitialPower
            worksheet.Range("G19").Value = Can6705v
            worksheet.Range("G20").Value = Isolatedv12
            worksheet.Range("G21").Value = Sensorv12
            worksheet.Range("G22").Value = TT6205v
            worksheet.Range("G23").Value = IT6615v
            worksheet.Range("G24").Value = A0691
            worksheet.Range("G25").Value = D0691
            worksheet.Range("G27").Value = D0691ON
            worksheet.Range("G28").Value = D0661ON
            worksheet.Range("G29").Value = D0661OFF
            worksheet.Range("G30").Value = AUXENCMDON
            worksheet.Range("G31").Value = D0100ON
            worksheet.Range("G32").Value = D0100OFF
            worksheet.Range("G33").Value = D0101ON
            worksheet.Range("G34").Value = D0101OFF
            worksheet.Range("G35").Value = CMDOFF
            Dim worksheet1 As Excel.Worksheet = workbook.Sheets.Add(After:=worksheet)
            worksheet1.Name = "1"
            Dim worksheet2 As Excel.Worksheet = workbook.Sheets.Add(After:=worksheet1)
            worksheet2.Name = "2"
            Dim worksheet3 As Excel.Worksheet = workbook.Sheets.Add(After:=worksheet2)
            worksheet3.Name = "3"
            Dim worksheet4 As Excel.Worksheet = workbook.Sheets.Add(After:=worksheet3)
            worksheet4.Name = "4"
            Dim worksheet5 As Excel.Worksheet = workbook.Sheets.Add(After:=worksheet4)
            worksheet5.Name = "5"
            Dim worksheet6 As Excel.Worksheet = workbook.Sheets.Add(After:=worksheet5)
            worksheet6.Name = "6"
            Dim worksheet7 As Excel.Worksheet = workbook.Sheets.Add(After:=worksheet6)
            worksheet7.Name = "7"
            Dim worksheet8 As Excel.Worksheet = workbook.Sheets.Add(After:=worksheet7)
            worksheet8.Name = "8"
            Dim worksheet9 As Excel.Worksheet = workbook.Sheets.Add(After:=worksheet8)
            worksheet9.Name = "9"
            Dim worksheet10 As Excel.Worksheet = workbook.Sheets.Add(After:=worksheet9)
            worksheet10.Name = "10"
            Clipboard.Clear()
            Clipboard.SetImage(IMAGE_STR.PictureBox1.Image)
            worksheet1.Paste()
            Clipboard.Clear()
            Clipboard.SetImage(IMAGE_STR.PictureBox2.Image)
            worksheet2.Paste()
            Clipboard.Clear()
            Clipboard.SetImage(IMAGE_STR.PictureBox3.Image)
            worksheet3.Paste()
            Clipboard.Clear()
            Clipboard.SetImage(IMAGE_STR.PictureBox4.Image)
            worksheet4.Paste()
            Clipboard.Clear()
            Clipboard.SetImage(IMAGE_STR.PictureBox5.Image)
            worksheet5.Paste()
            Clipboard.Clear()
            Clipboard.SetImage(IMAGE_STR.PictureBox6.Image)
            worksheet6.Paste()
            Clipboard.Clear()
            Clipboard.SetImage(IMAGE_STR.PictureBox7.Image)
            worksheet7.Paste()
            Clipboard.Clear()
            Clipboard.SetImage(IMAGE_STR.PictureBox8.Image)
            worksheet8.Paste()
            Clipboard.Clear()
            Clipboard.SetImage(IMAGE_STR.PictureBox9.Image)
            worksheet9.Paste()
            Clipboard.Clear()
            Clipboard.SetImage(IMAGE_STR.PictureBox10.Image)
            worksheet10.Paste()
            Clipboard.Clear()
            With worksheet1.PageSetup
                .Orientation = Excel.XlPageOrientation.xlLandscape
                .Zoom = False
                .FitToPagesTall = 1
                .FitToPagesWide = 1
            End With
            With worksheet2.PageSetup
                .Orientation = Excel.XlPageOrientation.xlLandscape
                .Zoom = False
                .FitToPagesTall = 1
                .FitToPagesWide = 1
            End With
            With worksheet3.PageSetup
                .Orientation = Excel.XlPageOrientation.xlLandscape
                .Zoom = False
                .FitToPagesTall = 1
                .FitToPagesWide = 1
            End With
            With worksheet4.PageSetup
                .Orientation = Excel.XlPageOrientation.xlLandscape
                .Zoom = False
                .FitToPagesTall = 1
                .FitToPagesWide = 1
            End With
            With worksheet5.PageSetup
                .Orientation = Excel.XlPageOrientation.xlLandscape
                .Zoom = False
                .FitToPagesTall = 1
                .FitToPagesWide = 1
            End With
            With worksheet6.PageSetup
                .Orientation = Excel.XlPageOrientation.xlLandscape
                .Zoom = False
                .FitToPagesTall = 1
                .FitToPagesWide = 1
            End With
            With worksheet7.PageSetup
                .Orientation = Excel.XlPageOrientation.xlLandscape
                .Zoom = False
                .FitToPagesTall = 1
                .FitToPagesWide = 1
            End With
            With worksheet8.PageSetup
                .Orientation = Excel.XlPageOrientation.xlLandscape
                .Zoom = False
                .FitToPagesTall = 1
                .FitToPagesWide = 1
            End With
            With worksheet9.PageSetup
                .Orientation = Excel.XlPageOrientation.xlLandscape
                .Zoom = False
                .FitToPagesTall = 1
                .FitToPagesWide = 1
            End With
            With worksheet10.PageSetup
                .Orientation = Excel.XlPageOrientation.xlLandscape
                .Zoom = False
                .FitToPagesTall = 1
                .FitToPagesWide = 1
            End With
            workbook.ExportAsFixedFormat(Excel.XlFixedFormatType.xlTypePDF, CONFIG.TextBox83.Text.ToString & SerialNumber & "_" & TestMode & "_" & Status & Format(Now(), "hhmmssmmddyy") & ".pdf", vbNormal)
            workbook.Worksheets("1").Delete()
            workbook.Worksheets("2").Delete()
            workbook.Worksheets("3").Delete()
            workbook.Worksheets("4").Delete()
            workbook.Worksheets("5").Delete()
            workbook.Worksheets("6").Delete()
            workbook.Worksheets("7").Delete()
            workbook.Worksheets("8").Delete()
            workbook.Worksheets("9").Delete()
            workbook.Worksheets("10").Delete()
            workbook.Close(True)
            System.Runtime.InteropServices.Marshal.ReleaseComObject(workbook)
            workbook = Nothing
            xlapp.Quit()
            System.Runtime.InteropServices.Marshal.ReleaseComObject(xlapp)
            xlapp = Nothing
            Process.GetProcessesByName("EXCEL")(0).Kill()
            file1.Close()
            MAIN.ListBox1.Items.Add($"{DateTime.Now.ToString} Test Report Generated success : " & CONFIG.TextBox83.Text.ToString & SerialNumber & "_" & TestMode & "_" & Status & Format(Now(), "hhmmssmmddyy"))
            MAIN.ListBox1.SelectedIndex = MAIN.ListBox1.Items.Count - 1
            Dim SW As IO.StreamWriter = IO.File.CreateText(CONFIG.TextBox83.Text.ToString & SerialNumber)
            For Each S As String In MAIN.ListBox1.Items
                SW.WriteLine(S, True)
            Next
            SW.Close()
            file1.Close()
        Catch ex As Exception
            Dim iret As Object = MsgBox("Error generating Report : " & ex.Message, vbCritical + vbRetryCancel, "Error executing SMC.GenerateReportOnerrorInit")
            If iret = vbRetry Then
                file1.Close()
                GoTo line1
            ElseIf iret = vbCancel Then
                file1.WriteLine("Error generating Report :  " & ex.Message & " " & "@" & DateTime.Now.ToString("yyyy/MM/dd/HH:mm:ss"))
                file1.Close()
                MAIN.ListBox1.Items.Add($"{DateTime.Now.ToString} Error generating Report :" & ex.Message)
                MAIN.ListBox1.SelectedIndex = MAIN.ListBox1.Items.Count - 1
                Exit Function
            End If
        End Try
    End Function
    Function GenerateReportOnErrorSMCSwitchTestNode9(User As String, TestMode As String, SerialNumber As String, Status As String, CycleTime As String, TesterID As String, Mark As String, ErrorMsg As String, DutConnect As String, SetPower As String, UniflashProgramme As String, FTDIProg As String, InitialPower As String, Can6705v As String, Isolatedv12 As String, Sensorv12 As String, TT6205v As String, IT6615v As String, A0691 As String, D0691 As String, D0691ON As String, D0661ON As String, D0661OFF As String, AUXENCMDON As String, D0100ON As String, D0100OFF As String, D0101ON As String, D0101OFF As String, CMDOFF As String, SW2ON As String)
        Dim line1 As Label
line1:
        Dim path1 As String = CONFIG.TextBox4.Text
        file1 = My.Computer.FileSystem.OpenTextFileWriter(path1, True)
        Try
            Dim xlapp = New Microsoft.Office.Interop.Excel.Application
            xlapp.DisplayAlerts = False
            Dim workbook = xlapp.Workbooks.Open(CONFIG.TextBox82.Text.ToString)
            Dim worksheet As Excel.Worksheet
            worksheet = workbook.Sheets("SMC")
            worksheet.Range("D6").Value = User
            worksheet.Range("D7").Value = TestMode
            worksheet.Range("D8").Value = SerialNumber
            worksheet.Range("D9").Value = DateTime.Now.ToString
            worksheet.Range("D10").Value = Status
            worksheet.Range("F6").Value = CycleTime
            worksheet.Range("F7").Value = TesterID
            worksheet.Range("F8").Value = Mark
            worksheet.Range("G7").Value = ErrorMsg
            worksheet.Range("G13").Value = DutConnect
            worksheet.Range("G14").Value = SetPower
            worksheet.Range("G15").Value = UniflashProgramme
            worksheet.Range("G16").Value = FTDIProg
            worksheet.Range("G18").Value = InitialPower
            worksheet.Range("G19").Value = Can6705v
            worksheet.Range("G20").Value = Isolatedv12
            worksheet.Range("G21").Value = Sensorv12
            worksheet.Range("G22").Value = TT6205v
            worksheet.Range("G23").Value = IT6615v
            worksheet.Range("G24").Value = A0691
            worksheet.Range("G25").Value = D0691
            worksheet.Range("G27").Value = D0691ON
            worksheet.Range("G28").Value = D0661ON
            worksheet.Range("G29").Value = D0661OFF
            worksheet.Range("G30").Value = AUXENCMDON
            worksheet.Range("G31").Value = D0100ON
            worksheet.Range("G32").Value = D0100OFF
            worksheet.Range("G33").Value = D0101ON
            worksheet.Range("G34").Value = D0101OFF
            worksheet.Range("G35").Value = CMDOFF
            worksheet.Range("G36").Value = SW2ON
            Dim worksheet1 As Excel.Worksheet = workbook.Sheets.Add(After:=worksheet)
            worksheet1.Name = "1"
            Dim worksheet2 As Excel.Worksheet = workbook.Sheets.Add(After:=worksheet1)
            worksheet2.Name = "2"
            Dim worksheet3 As Excel.Worksheet = workbook.Sheets.Add(After:=worksheet2)
            worksheet3.Name = "3"
            Dim worksheet4 As Excel.Worksheet = workbook.Sheets.Add(After:=worksheet3)
            worksheet4.Name = "4"
            Dim worksheet5 As Excel.Worksheet = workbook.Sheets.Add(After:=worksheet4)
            worksheet5.Name = "5"
            Dim worksheet6 As Excel.Worksheet = workbook.Sheets.Add(After:=worksheet5)
            worksheet6.Name = "6"
            Dim worksheet7 As Excel.Worksheet = workbook.Sheets.Add(After:=worksheet6)
            worksheet7.Name = "7"
            Dim worksheet8 As Excel.Worksheet = workbook.Sheets.Add(After:=worksheet7)
            worksheet8.Name = "8"
            Dim worksheet9 As Excel.Worksheet = workbook.Sheets.Add(After:=worksheet8)
            worksheet9.Name = "9"
            Dim worksheet10 As Excel.Worksheet = workbook.Sheets.Add(After:=worksheet9)
            worksheet10.Name = "10"
            Dim worksheet11 As Excel.Worksheet = workbook.Sheets.Add(After:=worksheet10)
            worksheet11.Name = "11"
            Clipboard.Clear()
            Clipboard.SetImage(IMAGE_STR.PictureBox1.Image)
            worksheet1.Paste()
            Clipboard.Clear()
            Clipboard.SetImage(IMAGE_STR.PictureBox2.Image)
            worksheet2.Paste()
            Clipboard.Clear()
            Clipboard.SetImage(IMAGE_STR.PictureBox3.Image)
            worksheet3.Paste()
            Clipboard.Clear()
            Clipboard.SetImage(IMAGE_STR.PictureBox4.Image)
            worksheet4.Paste()
            Clipboard.Clear()
            Clipboard.SetImage(IMAGE_STR.PictureBox5.Image)
            worksheet5.Paste()
            Clipboard.Clear()
            Clipboard.SetImage(IMAGE_STR.PictureBox6.Image)
            worksheet6.Paste()
            Clipboard.Clear()
            Clipboard.SetImage(IMAGE_STR.PictureBox7.Image)
            worksheet7.Paste()
            Clipboard.Clear()
            Clipboard.SetImage(IMAGE_STR.PictureBox8.Image)
            worksheet8.Paste()
            Clipboard.Clear()
            Clipboard.SetImage(IMAGE_STR.PictureBox9.Image)
            worksheet9.Paste()
            Clipboard.Clear()
            Clipboard.SetImage(IMAGE_STR.PictureBox10.Image)
            worksheet10.Paste()
            Clipboard.Clear()
            Clipboard.SetImage(IMAGE_STR.PictureBox11.Image)
            worksheet11.Paste()
            Clipboard.Clear()
            With worksheet1.PageSetup
                .Orientation = Excel.XlPageOrientation.xlLandscape
                .Zoom = False
                .FitToPagesTall = 1
                .FitToPagesWide = 1
            End With
            With worksheet2.PageSetup
                .Orientation = Excel.XlPageOrientation.xlLandscape
                .Zoom = False
                .FitToPagesTall = 1
                .FitToPagesWide = 1
            End With
            With worksheet3.PageSetup
                .Orientation = Excel.XlPageOrientation.xlLandscape
                .Zoom = False
                .FitToPagesTall = 1
                .FitToPagesWide = 1
            End With
            With worksheet4.PageSetup
                .Orientation = Excel.XlPageOrientation.xlLandscape
                .Zoom = False
                .FitToPagesTall = 1
                .FitToPagesWide = 1
            End With
            With worksheet5.PageSetup
                .Orientation = Excel.XlPageOrientation.xlLandscape
                .Zoom = False
                .FitToPagesTall = 1
                .FitToPagesWide = 1
            End With
            With worksheet6.PageSetup
                .Orientation = Excel.XlPageOrientation.xlLandscape
                .Zoom = False
                .FitToPagesTall = 1
                .FitToPagesWide = 1
            End With
            With worksheet7.PageSetup
                .Orientation = Excel.XlPageOrientation.xlLandscape
                .Zoom = False
                .FitToPagesTall = 1
                .FitToPagesWide = 1
            End With
            With worksheet8.PageSetup
                .Orientation = Excel.XlPageOrientation.xlLandscape
                .Zoom = False
                .FitToPagesTall = 1
                .FitToPagesWide = 1
            End With
            With worksheet9.PageSetup
                .Orientation = Excel.XlPageOrientation.xlLandscape
                .Zoom = False
                .FitToPagesTall = 1
                .FitToPagesWide = 1
            End With
            With worksheet10.PageSetup
                .Orientation = Excel.XlPageOrientation.xlLandscape
                .Zoom = False
                .FitToPagesTall = 1
                .FitToPagesWide = 1
            End With
            With worksheet11.PageSetup
                .Orientation = Excel.XlPageOrientation.xlLandscape
                .Zoom = False
                .FitToPagesTall = 1
                .FitToPagesWide = 1
            End With
            workbook.ExportAsFixedFormat(Excel.XlFixedFormatType.xlTypePDF, CONFIG.TextBox83.Text.ToString & SerialNumber & "_" & TestMode & "_" & Status & Format(Now(), "hhmmssmmddyy") & ".pdf", vbNormal)
            workbook.Worksheets("1").Delete()
            workbook.Worksheets("2").Delete()
            workbook.Worksheets("3").Delete()
            workbook.Worksheets("4").Delete()
            workbook.Worksheets("5").Delete()
            workbook.Worksheets("6").Delete()
            workbook.Worksheets("7").Delete()
            workbook.Worksheets("8").Delete()
            workbook.Worksheets("9").Delete()
            workbook.Worksheets("10").Delete()
            workbook.Worksheets("11").Delete()
            workbook.Close(True)
            System.Runtime.InteropServices.Marshal.ReleaseComObject(workbook)
            workbook = Nothing
            xlapp.Quit()
            System.Runtime.InteropServices.Marshal.ReleaseComObject(xlapp)
            xlapp = Nothing
            Process.GetProcessesByName("EXCEL")(0).Kill()
            file1.Close()
            MAIN.ListBox1.Items.Add($"{DateTime.Now.ToString} Test Report Generated success : " & CONFIG.TextBox83.Text.ToString & SerialNumber & "_" & TestMode & "_" & Status & Format(Now(), "hhmmssmmddyy"))
            MAIN.ListBox1.SelectedIndex = MAIN.ListBox1.Items.Count - 1
            Dim SW As IO.StreamWriter = IO.File.CreateText(CONFIG.TextBox83.Text.ToString & SerialNumber)
            For Each S As String In MAIN.ListBox1.Items
                SW.WriteLine(S, True)
            Next
            SW.Close()
            file1.Close()
        Catch ex As Exception
            Dim iret As Object = MsgBox("Error generating Report : " & ex.Message, vbCritical + vbRetryCancel, "Error executing SMC.GenerateReportOnerrorInit")
            If iret = vbRetry Then
                file1.Close()
                GoTo line1
            ElseIf iret = vbCancel Then
                file1.WriteLine("Error generating Report :  " & ex.Message & " " & "@" & DateTime.Now.ToString("yyyy/MM/dd/HH:mm:ss"))
                file1.Close()
                MAIN.ListBox1.Items.Add($"{DateTime.Now.ToString} Error generating Report :" & ex.Message)
                MAIN.ListBox1.SelectedIndex = MAIN.ListBox1.Items.Count - 1
                Exit Function
            End If
        End Try
    End Function
    Function GenerateReportOnErrorSMCSwitchTestNode10(User As String, TestMode As String, SerialNumber As String, Status As String, CycleTime As String, TesterID As String, Mark As String, ErrorMsg As String, DutConnect As String, SetPower As String, UniflashProgramme As String, FTDIProg As String, InitialPower As String, Can6705v As String, Isolatedv12 As String, Sensorv12 As String, TT6205v As String, IT6615v As String, A0691 As String, D0691 As String, D0691ON As String, D0661ON As String, D0661OFF As String, AUXENCMDON As String, D0100ON As String, D0100OFF As String, D0101ON As String, D0101OFF As String, CMDOFF As String, SW2ON As String, SW3ON As String)
        Dim line1 As Label
line1:
        Dim path1 As String = CONFIG.TextBox4.Text
        file1 = My.Computer.FileSystem.OpenTextFileWriter(path1, True)
        Try
            Dim xlapp = New Microsoft.Office.Interop.Excel.Application
            xlapp.DisplayAlerts = False
            Dim workbook = xlapp.Workbooks.Open(CONFIG.TextBox82.Text.ToString)
            Dim worksheet As Excel.Worksheet
            worksheet = workbook.Sheets("SMC")
            worksheet.Range("D6").Value = User
            worksheet.Range("D7").Value = TestMode
            worksheet.Range("D8").Value = SerialNumber
            worksheet.Range("D9").Value = DateTime.Now.ToString
            worksheet.Range("D10").Value = Status
            worksheet.Range("F6").Value = CycleTime
            worksheet.Range("F7").Value = TesterID
            worksheet.Range("F8").Value = Mark
            worksheet.Range("G7").Value = ErrorMsg
            worksheet.Range("G13").Value = DutConnect
            worksheet.Range("G14").Value = SetPower
            worksheet.Range("G15").Value = UniflashProgramme
            worksheet.Range("G16").Value = FTDIProg
            worksheet.Range("G18").Value = InitialPower
            worksheet.Range("G19").Value = Can6705v
            worksheet.Range("G20").Value = Isolatedv12
            worksheet.Range("G21").Value = Sensorv12
            worksheet.Range("G22").Value = TT6205v
            worksheet.Range("G23").Value = IT6615v
            worksheet.Range("G24").Value = A0691
            worksheet.Range("G25").Value = D0691
            worksheet.Range("G27").Value = D0691ON
            worksheet.Range("G28").Value = D0661ON
            worksheet.Range("G29").Value = D0661OFF
            worksheet.Range("G30").Value = AUXENCMDON
            worksheet.Range("G31").Value = D0100ON
            worksheet.Range("G32").Value = D0100OFF
            worksheet.Range("G33").Value = D0101ON
            worksheet.Range("G34").Value = D0101OFF
            worksheet.Range("G35").Value = CMDOFF
            worksheet.Range("G36").Value = SW2ON
            worksheet.Range("G37").Value = SW3ON
            Dim worksheet1 As Excel.Worksheet = workbook.Sheets.Add(After:=worksheet)
            worksheet1.Name = "1"
            Dim worksheet2 As Excel.Worksheet = workbook.Sheets.Add(After:=worksheet1)
            worksheet2.Name = "2"
            Dim worksheet3 As Excel.Worksheet = workbook.Sheets.Add(After:=worksheet2)
            worksheet3.Name = "3"
            Dim worksheet4 As Excel.Worksheet = workbook.Sheets.Add(After:=worksheet3)
            worksheet4.Name = "4"
            Dim worksheet5 As Excel.Worksheet = workbook.Sheets.Add(After:=worksheet4)
            worksheet5.Name = "5"
            Dim worksheet6 As Excel.Worksheet = workbook.Sheets.Add(After:=worksheet5)
            worksheet6.Name = "6"
            Dim worksheet7 As Excel.Worksheet = workbook.Sheets.Add(After:=worksheet6)
            worksheet7.Name = "7"
            Dim worksheet8 As Excel.Worksheet = workbook.Sheets.Add(After:=worksheet7)
            worksheet8.Name = "8"
            Dim worksheet9 As Excel.Worksheet = workbook.Sheets.Add(After:=worksheet8)
            worksheet9.Name = "9"
            Dim worksheet10 As Excel.Worksheet = workbook.Sheets.Add(After:=worksheet9)
            worksheet10.Name = "10"
            Dim worksheet11 As Excel.Worksheet = workbook.Sheets.Add(After:=worksheet10)
            worksheet11.Name = "11"
            Dim worksheet12 As Excel.Worksheet = workbook.Sheets.Add(After:=worksheet11)
            worksheet12.Name = "12"
            Clipboard.Clear()
            Clipboard.SetImage(IMAGE_STR.PictureBox1.Image)
            worksheet1.Paste()
            Clipboard.Clear()
            Clipboard.SetImage(IMAGE_STR.PictureBox2.Image)
            worksheet2.Paste()
            Clipboard.Clear()
            Clipboard.SetImage(IMAGE_STR.PictureBox3.Image)
            worksheet3.Paste()
            Clipboard.Clear()
            Clipboard.SetImage(IMAGE_STR.PictureBox4.Image)
            worksheet4.Paste()
            Clipboard.Clear()
            Clipboard.SetImage(IMAGE_STR.PictureBox5.Image)
            worksheet5.Paste()
            Clipboard.Clear()
            Clipboard.SetImage(IMAGE_STR.PictureBox6.Image)
            worksheet6.Paste()
            Clipboard.Clear()
            Clipboard.SetImage(IMAGE_STR.PictureBox7.Image)
            worksheet7.Paste()
            Clipboard.Clear()
            Clipboard.SetImage(IMAGE_STR.PictureBox8.Image)
            worksheet8.Paste()
            Clipboard.Clear()
            Clipboard.SetImage(IMAGE_STR.PictureBox9.Image)
            worksheet9.Paste()
            Clipboard.Clear()
            Clipboard.SetImage(IMAGE_STR.PictureBox10.Image)
            worksheet10.Paste()
            Clipboard.Clear()
            Clipboard.SetImage(IMAGE_STR.PictureBox11.Image)
            worksheet11.Paste()
            Clipboard.Clear()
            Clipboard.SetImage(IMAGE_STR.PictureBox12.Image)
            worksheet12.Paste()
            Clipboard.Clear()
            With worksheet1.PageSetup
                .Orientation = Excel.XlPageOrientation.xlLandscape
                .Zoom = False
                .FitToPagesTall = 1
                .FitToPagesWide = 1
            End With
            With worksheet2.PageSetup
                .Orientation = Excel.XlPageOrientation.xlLandscape
                .Zoom = False
                .FitToPagesTall = 1
                .FitToPagesWide = 1
            End With
            With worksheet3.PageSetup
                .Orientation = Excel.XlPageOrientation.xlLandscape
                .Zoom = False
                .FitToPagesTall = 1
                .FitToPagesWide = 1
            End With
            With worksheet4.PageSetup
                .Orientation = Excel.XlPageOrientation.xlLandscape
                .Zoom = False
                .FitToPagesTall = 1
                .FitToPagesWide = 1
            End With
            With worksheet5.PageSetup
                .Orientation = Excel.XlPageOrientation.xlLandscape
                .Zoom = False
                .FitToPagesTall = 1
                .FitToPagesWide = 1
            End With
            With worksheet6.PageSetup
                .Orientation = Excel.XlPageOrientation.xlLandscape
                .Zoom = False
                .FitToPagesTall = 1
                .FitToPagesWide = 1
            End With
            With worksheet7.PageSetup
                .Orientation = Excel.XlPageOrientation.xlLandscape
                .Zoom = False
                .FitToPagesTall = 1
                .FitToPagesWide = 1
            End With
            With worksheet8.PageSetup
                .Orientation = Excel.XlPageOrientation.xlLandscape
                .Zoom = False
                .FitToPagesTall = 1
                .FitToPagesWide = 1
            End With
            With worksheet9.PageSetup
                .Orientation = Excel.XlPageOrientation.xlLandscape
                .Zoom = False
                .FitToPagesTall = 1
                .FitToPagesWide = 1
            End With
            With worksheet10.PageSetup
                .Orientation = Excel.XlPageOrientation.xlLandscape
                .Zoom = False
                .FitToPagesTall = 1
                .FitToPagesWide = 1
            End With
            With worksheet11.PageSetup
                .Orientation = Excel.XlPageOrientation.xlLandscape
                .Zoom = False
                .FitToPagesTall = 1
                .FitToPagesWide = 1
            End With
            With worksheet12.PageSetup
                .Orientation = Excel.XlPageOrientation.xlLandscape
                .Zoom = False
                .FitToPagesTall = 1
                .FitToPagesWide = 1
            End With
            workbook.ExportAsFixedFormat(Excel.XlFixedFormatType.xlTypePDF, CONFIG.TextBox83.Text.ToString & SerialNumber & "_" & TestMode & "_" & Status & Format(Now(), "hhmmssmmddyy") & ".pdf", vbNormal)
            workbook.Worksheets("1").Delete()
            workbook.Worksheets("2").Delete()
            workbook.Worksheets("3").Delete()
            workbook.Worksheets("4").Delete()
            workbook.Worksheets("5").Delete()
            workbook.Worksheets("6").Delete()
            workbook.Worksheets("7").Delete()
            workbook.Worksheets("8").Delete()
            workbook.Worksheets("9").Delete()
            workbook.Worksheets("10").Delete()
            workbook.Worksheets("11").Delete()
            workbook.Worksheets("12").Delete()
            workbook.Close(True)
            System.Runtime.InteropServices.Marshal.ReleaseComObject(workbook)
            workbook = Nothing
            xlapp.Quit()
            System.Runtime.InteropServices.Marshal.ReleaseComObject(xlapp)
            xlapp = Nothing
            Process.GetProcessesByName("EXCEL")(0).Kill()
            file1.Close()
            MAIN.ListBox1.Items.Add($"{DateTime.Now.ToString} Test Report Generated success : " & CONFIG.TextBox83.Text.ToString & SerialNumber & "_" & TestMode & "_" & Status & Format(Now(), "hhmmssmmddyy"))
            MAIN.ListBox1.SelectedIndex = MAIN.ListBox1.Items.Count - 1
            Dim SW As IO.StreamWriter = IO.File.CreateText(CONFIG.TextBox83.Text.ToString & SerialNumber)
            For Each S As String In MAIN.ListBox1.Items
                SW.WriteLine(S, True)
            Next
            SW.Close()
            file1.Close()
        Catch ex As Exception
            Dim iret As Object = MsgBox("Error generating Report : " & ex.Message, vbCritical + vbRetryCancel, "Error executing SMC.GenerateReportOnerrorInit")
            If iret = vbRetry Then
                file1.Close()
                GoTo line1
            ElseIf iret = vbCancel Then
                file1.WriteLine("Error generating Report :  " & ex.Message & " " & "@" & DateTime.Now.ToString("yyyy/MM/dd/HH:mm:ss"))
                file1.Close()
                MAIN.ListBox1.Items.Add($"{DateTime.Now.ToString} Error generating Report :" & ex.Message)
                MAIN.ListBox1.SelectedIndex = MAIN.ListBox1.Items.Count - 1
                Exit Function
            End If
        End Try
    End Function
    Function GenerateReportOnErrorSMCSwitchTestNode11(User As String, TestMode As String, SerialNumber As String, Status As String, CycleTime As String, TesterID As String, Mark As String, ErrorMsg As String, DutConnect As String, SetPower As String, UniflashProgramme As String, FTDIProg As String, InitialPower As String, Can6705v As String, Isolatedv12 As String, Sensorv12 As String, TT6205v As String, IT6615v As String, A0691 As String, D0691 As String, D0691ON As String, D0661ON As String, D0661OFF As String, AUXENCMDON As String, D0100ON As String, D0100OFF As String, D0101ON As String, D0101OFF As String, CMDOFF As String, SW2ON As String, SW3ON As String, TOOGLECMDON As String)
        Dim line1 As Label
line1:
        Dim path1 As String = CONFIG.TextBox4.Text
        file1 = My.Computer.FileSystem.OpenTextFileWriter(path1, True)
        Try
            Dim xlapp = New Microsoft.Office.Interop.Excel.Application
            xlapp.DisplayAlerts = False
            Dim workbook = xlapp.Workbooks.Open(CONFIG.TextBox82.Text.ToString)
            Dim worksheet As Excel.Worksheet
            worksheet = workbook.Sheets("SMC")
            worksheet.Range("D6").Value = User
            worksheet.Range("D7").Value = TestMode
            worksheet.Range("D8").Value = SerialNumber
            worksheet.Range("D9").Value = DateTime.Now.ToString
            worksheet.Range("D10").Value = Status
            worksheet.Range("F6").Value = CycleTime
            worksheet.Range("F7").Value = TesterID
            worksheet.Range("F8").Value = Mark
            worksheet.Range("G7").Value = ErrorMsg
            worksheet.Range("G13").Value = DutConnect
            worksheet.Range("G14").Value = SetPower
            worksheet.Range("G15").Value = UniflashProgramme
            worksheet.Range("G16").Value = FTDIProg
            worksheet.Range("G18").Value = InitialPower
            worksheet.Range("G19").Value = Can6705v
            worksheet.Range("G20").Value = Isolatedv12
            worksheet.Range("G21").Value = Sensorv12
            worksheet.Range("G22").Value = TT6205v
            worksheet.Range("G23").Value = IT6615v
            worksheet.Range("G24").Value = A0691
            worksheet.Range("G25").Value = D0691
            worksheet.Range("G27").Value = D0691ON
            worksheet.Range("G28").Value = D0661ON
            worksheet.Range("G29").Value = D0661OFF
            worksheet.Range("G30").Value = AUXENCMDON
            worksheet.Range("G31").Value = D0100ON
            worksheet.Range("G32").Value = D0100OFF
            worksheet.Range("G33").Value = D0101ON
            worksheet.Range("G34").Value = D0101OFF
            worksheet.Range("G35").Value = CMDOFF
            worksheet.Range("G36").Value = SW2ON
            worksheet.Range("G37").Value = SW3ON
            worksheet.Range("G38").Value = TOOGLECMDON
            Dim worksheet1 As Excel.Worksheet = workbook.Sheets.Add(After:=worksheet)
            worksheet1.Name = "1"
            Dim worksheet2 As Excel.Worksheet = workbook.Sheets.Add(After:=worksheet1)
            worksheet2.Name = "2"
            Dim worksheet3 As Excel.Worksheet = workbook.Sheets.Add(After:=worksheet2)
            worksheet3.Name = "3"
            Dim worksheet4 As Excel.Worksheet = workbook.Sheets.Add(After:=worksheet3)
            worksheet4.Name = "4"
            Dim worksheet5 As Excel.Worksheet = workbook.Sheets.Add(After:=worksheet4)
            worksheet5.Name = "5"
            Dim worksheet6 As Excel.Worksheet = workbook.Sheets.Add(After:=worksheet5)
            worksheet6.Name = "6"
            Dim worksheet7 As Excel.Worksheet = workbook.Sheets.Add(After:=worksheet6)
            worksheet7.Name = "7"
            Dim worksheet8 As Excel.Worksheet = workbook.Sheets.Add(After:=worksheet7)
            worksheet8.Name = "8"
            Dim worksheet9 As Excel.Worksheet = workbook.Sheets.Add(After:=worksheet8)
            worksheet9.Name = "9"
            Dim worksheet10 As Excel.Worksheet = workbook.Sheets.Add(After:=worksheet9)
            worksheet10.Name = "10"
            Dim worksheet11 As Excel.Worksheet = workbook.Sheets.Add(After:=worksheet10)
            worksheet11.Name = "11"
            Dim worksheet12 As Excel.Worksheet = workbook.Sheets.Add(After:=worksheet11)
            worksheet12.Name = "12"
            Dim worksheet13 As Excel.Worksheet = workbook.Sheets.Add(After:=worksheet12)
            worksheet13.Name = "13"
            Clipboard.Clear()
            Clipboard.SetImage(IMAGE_STR.PictureBox1.Image)
            worksheet1.Paste()
            Clipboard.Clear()
            Clipboard.SetImage(IMAGE_STR.PictureBox2.Image)
            worksheet2.Paste()
            Clipboard.Clear()
            Clipboard.SetImage(IMAGE_STR.PictureBox3.Image)
            worksheet3.Paste()
            Clipboard.Clear()
            Clipboard.SetImage(IMAGE_STR.PictureBox4.Image)
            worksheet4.Paste()
            Clipboard.Clear()
            Clipboard.SetImage(IMAGE_STR.PictureBox5.Image)
            worksheet5.Paste()
            Clipboard.Clear()
            Clipboard.SetImage(IMAGE_STR.PictureBox6.Image)
            worksheet6.Paste()
            Clipboard.Clear()
            Clipboard.SetImage(IMAGE_STR.PictureBox7.Image)
            worksheet7.Paste()
            Clipboard.Clear()
            Clipboard.SetImage(IMAGE_STR.PictureBox8.Image)
            worksheet8.Paste()
            Clipboard.Clear()
            Clipboard.SetImage(IMAGE_STR.PictureBox9.Image)
            worksheet9.Paste()
            Clipboard.Clear()
            Clipboard.SetImage(IMAGE_STR.PictureBox10.Image)
            worksheet10.Paste()
            Clipboard.Clear()
            Clipboard.SetImage(IMAGE_STR.PictureBox11.Image)
            worksheet11.Paste()
            Clipboard.Clear()
            Clipboard.SetImage(IMAGE_STR.PictureBox12.Image)
            worksheet12.Paste()
            Clipboard.Clear()
            Clipboard.SetImage(IMAGE_STR.PictureBox13.Image)
            worksheet13.Paste()
            Clipboard.Clear()
            With worksheet1.PageSetup
                .Orientation = Excel.XlPageOrientation.xlLandscape
                .Zoom = False
                .FitToPagesTall = 1
                .FitToPagesWide = 1
            End With
            With worksheet2.PageSetup
                .Orientation = Excel.XlPageOrientation.xlLandscape
                .Zoom = False
                .FitToPagesTall = 1
                .FitToPagesWide = 1
            End With
            With worksheet3.PageSetup
                .Orientation = Excel.XlPageOrientation.xlLandscape
                .Zoom = False
                .FitToPagesTall = 1
                .FitToPagesWide = 1
            End With
            With worksheet4.PageSetup
                .Orientation = Excel.XlPageOrientation.xlLandscape
                .Zoom = False
                .FitToPagesTall = 1
                .FitToPagesWide = 1
            End With
            With worksheet5.PageSetup
                .Orientation = Excel.XlPageOrientation.xlLandscape
                .Zoom = False
                .FitToPagesTall = 1
                .FitToPagesWide = 1
            End With
            With worksheet6.PageSetup
                .Orientation = Excel.XlPageOrientation.xlLandscape
                .Zoom = False
                .FitToPagesTall = 1
                .FitToPagesWide = 1
            End With
            With worksheet7.PageSetup
                .Orientation = Excel.XlPageOrientation.xlLandscape
                .Zoom = False
                .FitToPagesTall = 1
                .FitToPagesWide = 1
            End With
            With worksheet8.PageSetup
                .Orientation = Excel.XlPageOrientation.xlLandscape
                .Zoom = False
                .FitToPagesTall = 1
                .FitToPagesWide = 1
            End With
            With worksheet9.PageSetup
                .Orientation = Excel.XlPageOrientation.xlLandscape
                .Zoom = False
                .FitToPagesTall = 1
                .FitToPagesWide = 1
            End With
            With worksheet10.PageSetup
                .Orientation = Excel.XlPageOrientation.xlLandscape
                .Zoom = False
                .FitToPagesTall = 1
                .FitToPagesWide = 1
            End With
            With worksheet11.PageSetup
                .Orientation = Excel.XlPageOrientation.xlLandscape
                .Zoom = False
                .FitToPagesTall = 1
                .FitToPagesWide = 1
            End With
            With worksheet12.PageSetup
                .Orientation = Excel.XlPageOrientation.xlLandscape
                .Zoom = False
                .FitToPagesTall = 1
                .FitToPagesWide = 1
            End With
            With worksheet13.PageSetup
                .Orientation = Excel.XlPageOrientation.xlLandscape
                .Zoom = False
                .FitToPagesTall = 1
                .FitToPagesWide = 1
            End With
            workbook.ExportAsFixedFormat(Excel.XlFixedFormatType.xlTypePDF, CONFIG.TextBox83.Text.ToString & SerialNumber & "_" & TestMode & "_" & Status & Format(Now(), "hhmmssmmddyy") & ".pdf", vbNormal)
            workbook.Worksheets("1").Delete()
            workbook.Worksheets("2").Delete()
            workbook.Worksheets("3").Delete()
            workbook.Worksheets("4").Delete()
            workbook.Worksheets("5").Delete()
            workbook.Worksheets("6").Delete()
            workbook.Worksheets("7").Delete()
            workbook.Worksheets("8").Delete()
            workbook.Worksheets("9").Delete()
            workbook.Worksheets("10").Delete()
            workbook.Worksheets("11").Delete()
            workbook.Worksheets("12").Delete()
            workbook.Worksheets("13").Delete()
            workbook.Close(True)
            System.Runtime.InteropServices.Marshal.ReleaseComObject(workbook)
            workbook = Nothing
            xlapp.Quit()
            System.Runtime.InteropServices.Marshal.ReleaseComObject(xlapp)
            xlapp = Nothing
            Process.GetProcessesByName("EXCEL")(0).Kill()
            file1.Close()
            MAIN.ListBox1.Items.Add($"{DateTime.Now.ToString} Test Report Generated success : " & CONFIG.TextBox83.Text.ToString & SerialNumber & "_" & TestMode & "_" & Status & Format(Now(), "hhmmssmmddyy"))
            MAIN.ListBox1.SelectedIndex = MAIN.ListBox1.Items.Count - 1
            Dim SW As IO.StreamWriter = IO.File.CreateText(CONFIG.TextBox83.Text.ToString & SerialNumber)
            For Each S As String In MAIN.ListBox1.Items
                SW.WriteLine(S, True)
            Next
            SW.Close()
            file1.Close()
        Catch ex As Exception
            Dim iret As Object = MsgBox("Error generating Report : " & ex.Message, vbCritical + vbRetryCancel, "Error executing SMC.GenerateReportOnerrorInit")
            If iret = vbRetry Then
                file1.Close()
                GoTo line1
            ElseIf iret = vbCancel Then
                file1.WriteLine("Error generating Report :  " & ex.Message & " " & "@" & DateTime.Now.ToString("yyyy/MM/dd/HH:mm:ss"))
                file1.Close()
                MAIN.ListBox1.Items.Add($"{DateTime.Now.ToString} Error generating Report :" & ex.Message)
                MAIN.ListBox1.SelectedIndex = MAIN.ListBox1.Items.Count - 1
                Exit Function
            End If
        End Try
    End Function
    Function GenerateReportOnErrorSMCSwitchTestNode12(User As String, TestMode As String, SerialNumber As String, Status As String, CycleTime As String, TesterID As String, Mark As String, ErrorMsg As String, DutConnect As String, SetPower As String, UniflashProgramme As String, FTDIProg As String, InitialPower As String, Can6705v As String, Isolatedv12 As String, Sensorv12 As String, TT6205v As String, IT6615v As String, A0691 As String, D0691 As String, D0691ON As String, D0661ON As String, D0661OFF As String, AUXENCMDON As String, D0100ON As String, D0100OFF As String, D0101ON As String, D0101OFF As String, CMDOFF As String, SW2ON As String, SW3ON As String, TOOGLECMDON As String, CAN1SILENCEON As String)
        Dim line1 As Label
line1:
        Dim path1 As String = CONFIG.TextBox4.Text
        file1 = My.Computer.FileSystem.OpenTextFileWriter(path1, True)
        Try
            Dim xlapp = New Microsoft.Office.Interop.Excel.Application
            xlapp.DisplayAlerts = False
            Dim workbook = xlapp.Workbooks.Open(CONFIG.TextBox82.Text.ToString)
            Dim worksheet As Excel.Worksheet
            worksheet = workbook.Sheets("SMC")
            worksheet.Range("D6").Value = User
            worksheet.Range("D7").Value = TestMode
            worksheet.Range("D8").Value = SerialNumber
            worksheet.Range("D9").Value = DateTime.Now.ToString
            worksheet.Range("D10").Value = Status
            worksheet.Range("F6").Value = CycleTime
            worksheet.Range("F7").Value = TesterID
            worksheet.Range("F8").Value = Mark
            worksheet.Range("G7").Value = ErrorMsg
            worksheet.Range("G13").Value = DutConnect
            worksheet.Range("G14").Value = SetPower
            worksheet.Range("G15").Value = UniflashProgramme
            worksheet.Range("G16").Value = FTDIProg
            worksheet.Range("G18").Value = InitialPower
            worksheet.Range("G19").Value = Can6705v
            worksheet.Range("G20").Value = Isolatedv12
            worksheet.Range("G21").Value = Sensorv12
            worksheet.Range("G22").Value = TT6205v
            worksheet.Range("G23").Value = IT6615v
            worksheet.Range("G24").Value = A0691
            worksheet.Range("G25").Value = D0691
            worksheet.Range("G27").Value = D0691ON
            worksheet.Range("G28").Value = D0661ON
            worksheet.Range("G29").Value = D0661OFF
            worksheet.Range("G30").Value = AUXENCMDON
            worksheet.Range("G31").Value = D0100ON
            worksheet.Range("G32").Value = D0100OFF
            worksheet.Range("G33").Value = D0101ON
            worksheet.Range("G34").Value = D0101OFF
            worksheet.Range("G35").Value = CMDOFF
            worksheet.Range("G36").Value = SW2ON
            worksheet.Range("G37").Value = SW3ON
            worksheet.Range("G38").Value = TOOGLECMDON
            worksheet.Range("G39").Value = CAN1SILENCEON
            Dim worksheet1 As Excel.Worksheet = workbook.Sheets.Add(After:=worksheet)
            worksheet1.Name = "1"
            Dim worksheet2 As Excel.Worksheet = workbook.Sheets.Add(After:=worksheet1)
            worksheet2.Name = "2"
            Dim worksheet3 As Excel.Worksheet = workbook.Sheets.Add(After:=worksheet2)
            worksheet3.Name = "3"
            Dim worksheet4 As Excel.Worksheet = workbook.Sheets.Add(After:=worksheet3)
            worksheet4.Name = "4"
            Dim worksheet5 As Excel.Worksheet = workbook.Sheets.Add(After:=worksheet4)
            worksheet5.Name = "5"
            Dim worksheet6 As Excel.Worksheet = workbook.Sheets.Add(After:=worksheet5)
            worksheet6.Name = "6"
            Dim worksheet7 As Excel.Worksheet = workbook.Sheets.Add(After:=worksheet6)
            worksheet7.Name = "7"
            Dim worksheet8 As Excel.Worksheet = workbook.Sheets.Add(After:=worksheet7)
            worksheet8.Name = "8"
            Dim worksheet9 As Excel.Worksheet = workbook.Sheets.Add(After:=worksheet8)
            worksheet9.Name = "9"
            Dim worksheet10 As Excel.Worksheet = workbook.Sheets.Add(After:=worksheet9)
            worksheet10.Name = "10"
            Dim worksheet11 As Excel.Worksheet = workbook.Sheets.Add(After:=worksheet10)
            worksheet11.Name = "11"
            Dim worksheet12 As Excel.Worksheet = workbook.Sheets.Add(After:=worksheet11)
            worksheet12.Name = "12"
            Dim worksheet13 As Excel.Worksheet = workbook.Sheets.Add(After:=worksheet12)
            worksheet13.Name = "13"
            Dim worksheet14 As Excel.Worksheet = workbook.Sheets.Add(After:=worksheet13)
            worksheet14.Name = "14"
            Clipboard.Clear()
            Clipboard.SetImage(IMAGE_STR.PictureBox1.Image)
            worksheet1.Paste()
            Clipboard.Clear()
            Clipboard.SetImage(IMAGE_STR.PictureBox2.Image)
            worksheet2.Paste()
            Clipboard.Clear()
            Clipboard.SetImage(IMAGE_STR.PictureBox3.Image)
            worksheet3.Paste()
            Clipboard.Clear()
            Clipboard.SetImage(IMAGE_STR.PictureBox4.Image)
            worksheet4.Paste()
            Clipboard.Clear()
            Clipboard.SetImage(IMAGE_STR.PictureBox5.Image)
            worksheet5.Paste()
            Clipboard.Clear()
            Clipboard.SetImage(IMAGE_STR.PictureBox6.Image)
            worksheet6.Paste()
            Clipboard.Clear()
            Clipboard.SetImage(IMAGE_STR.PictureBox7.Image)
            worksheet7.Paste()
            Clipboard.Clear()
            Clipboard.SetImage(IMAGE_STR.PictureBox8.Image)
            worksheet8.Paste()
            Clipboard.Clear()
            Clipboard.SetImage(IMAGE_STR.PictureBox9.Image)
            worksheet9.Paste()
            Clipboard.Clear()
            Clipboard.SetImage(IMAGE_STR.PictureBox10.Image)
            worksheet10.Paste()
            Clipboard.Clear()
            Clipboard.SetImage(IMAGE_STR.PictureBox11.Image)
            worksheet11.Paste()
            Clipboard.Clear()
            Clipboard.SetImage(IMAGE_STR.PictureBox12.Image)
            worksheet12.Paste()
            Clipboard.Clear()
            Clipboard.SetImage(IMAGE_STR.PictureBox13.Image)
            worksheet13.Paste()
            Clipboard.Clear()
            Clipboard.SetImage(IMAGE_STR.PictureBox14.Image)
            worksheet14.Paste()
            Clipboard.Clear()
            With worksheet1.PageSetup
                .Orientation = Excel.XlPageOrientation.xlLandscape
                .Zoom = False
                .FitToPagesTall = 1
                .FitToPagesWide = 1
            End With
            With worksheet2.PageSetup
                .Orientation = Excel.XlPageOrientation.xlLandscape
                .Zoom = False
                .FitToPagesTall = 1
                .FitToPagesWide = 1
            End With
            With worksheet3.PageSetup
                .Orientation = Excel.XlPageOrientation.xlLandscape
                .Zoom = False
                .FitToPagesTall = 1
                .FitToPagesWide = 1
            End With
            With worksheet4.PageSetup
                .Orientation = Excel.XlPageOrientation.xlLandscape
                .Zoom = False
                .FitToPagesTall = 1
                .FitToPagesWide = 1
            End With
            With worksheet5.PageSetup
                .Orientation = Excel.XlPageOrientation.xlLandscape
                .Zoom = False
                .FitToPagesTall = 1
                .FitToPagesWide = 1
            End With
            With worksheet6.PageSetup
                .Orientation = Excel.XlPageOrientation.xlLandscape
                .Zoom = False
                .FitToPagesTall = 1
                .FitToPagesWide = 1
            End With
            With worksheet7.PageSetup
                .Orientation = Excel.XlPageOrientation.xlLandscape
                .Zoom = False
                .FitToPagesTall = 1
                .FitToPagesWide = 1
            End With
            With worksheet8.PageSetup
                .Orientation = Excel.XlPageOrientation.xlLandscape
                .Zoom = False
                .FitToPagesTall = 1
                .FitToPagesWide = 1
            End With
            With worksheet9.PageSetup
                .Orientation = Excel.XlPageOrientation.xlLandscape
                .Zoom = False
                .FitToPagesTall = 1
                .FitToPagesWide = 1
            End With
            With worksheet10.PageSetup
                .Orientation = Excel.XlPageOrientation.xlLandscape
                .Zoom = False
                .FitToPagesTall = 1
                .FitToPagesWide = 1
            End With
            With worksheet11.PageSetup
                .Orientation = Excel.XlPageOrientation.xlLandscape
                .Zoom = False
                .FitToPagesTall = 1
                .FitToPagesWide = 1
            End With
            With worksheet12.PageSetup
                .Orientation = Excel.XlPageOrientation.xlLandscape
                .Zoom = False
                .FitToPagesTall = 1
                .FitToPagesWide = 1
            End With
            With worksheet13.PageSetup
                .Orientation = Excel.XlPageOrientation.xlLandscape
                .Zoom = False
                .FitToPagesTall = 1
                .FitToPagesWide = 1
            End With
            With worksheet14.PageSetup
                .Orientation = Excel.XlPageOrientation.xlLandscape
                .Zoom = False
                .FitToPagesTall = 1
                .FitToPagesWide = 1
            End With
            workbook.ExportAsFixedFormat(Excel.XlFixedFormatType.xlTypePDF, CONFIG.TextBox83.Text.ToString & SerialNumber & "_" & TestMode & "_" & Status & Format(Now(), "hhmmssmmddyy") & ".pdf", vbNormal)
            workbook.Worksheets("1").Delete()
            workbook.Worksheets("2").Delete()
            workbook.Worksheets("3").Delete()
            workbook.Worksheets("4").Delete()
            workbook.Worksheets("5").Delete()
            workbook.Worksheets("6").Delete()
            workbook.Worksheets("7").Delete()
            workbook.Worksheets("8").Delete()
            workbook.Worksheets("9").Delete()
            workbook.Worksheets("10").Delete()
            workbook.Worksheets("11").Delete()
            workbook.Worksheets("12").Delete()
            workbook.Worksheets("14").Delete()
            workbook.Close(True)
            System.Runtime.InteropServices.Marshal.ReleaseComObject(workbook)
            workbook = Nothing
            xlapp.Quit()
            System.Runtime.InteropServices.Marshal.ReleaseComObject(xlapp)
            xlapp = Nothing
            Process.GetProcessesByName("EXCEL")(0).Kill()
            file1.Close()
            MAIN.ListBox1.Items.Add($"{DateTime.Now.ToString} Test Report Generated success : " & CONFIG.TextBox83.Text.ToString & SerialNumber & "_" & TestMode & "_" & Status & Format(Now(), "hhmmssmmddyy"))
            MAIN.ListBox1.SelectedIndex = MAIN.ListBox1.Items.Count - 1
            Dim SW As IO.StreamWriter = IO.File.CreateText(CONFIG.TextBox83.Text.ToString & SerialNumber)
            For Each S As String In MAIN.ListBox1.Items
                SW.WriteLine(S, True)
            Next
            SW.Close()
            file1.Close()
        Catch ex As Exception
            Dim iret As Object = MsgBox("Error generating Report : " & ex.Message, vbCritical + vbRetryCancel, "Error executing SMC.GenerateReportOnerrorInit")
            If iret = vbRetry Then
                file1.Close()
                GoTo line1
            ElseIf iret = vbCancel Then
                file1.WriteLine("Error generating Report :  " & ex.Message & " " & "@" & DateTime.Now.ToString("yyyy/MM/dd/HH:mm:ss"))
                file1.Close()
                MAIN.ListBox1.Items.Add($"{DateTime.Now.ToString} Error generating Report :" & ex.Message)
                MAIN.ListBox1.SelectedIndex = MAIN.ListBox1.Items.Count - 1
                Exit Function
            End If
        End Try
    End Function
    Function GenerateReportOnErrorSMCSwitchTestNode13(User As String, TestMode As String, SerialNumber As String, Status As String, CycleTime As String, TesterID As String, Mark As String, ErrorMsg As String, DutConnect As String, SetPower As String, UniflashProgramme As String, FTDIProg As String, InitialPower As String, Can6705v As String, Isolatedv12 As String, Sensorv12 As String, TT6205v As String, IT6615v As String, A0691 As String, D0691 As String, D0691ON As String, D0661ON As String, D0661OFF As String, AUXENCMDON As String, D0100ON As String, D0100OFF As String, D0101ON As String, D0101OFF As String, CMDOFF As String, SW2ON As String, SW3ON As String, TOOGLECMDON As String, CAN1SILENCEON As String, CAN2SILENCEON As String)
        Dim line1 As Label
line1:
        Dim path1 As String = CONFIG.TextBox4.Text
        file1 = My.Computer.FileSystem.OpenTextFileWriter(path1, True)
        Try
            Dim xlapp = New Microsoft.Office.Interop.Excel.Application
            xlapp.DisplayAlerts = False
            Dim workbook = xlapp.Workbooks.Open(CONFIG.TextBox82.Text.ToString)
            Dim worksheet As Excel.Worksheet
            worksheet = workbook.Sheets("SMC")
            worksheet.Range("D6").Value = User
            worksheet.Range("D7").Value = TestMode
            worksheet.Range("D8").Value = SerialNumber
            worksheet.Range("D9").Value = DateTime.Now.ToString
            worksheet.Range("D10").Value = Status
            worksheet.Range("F6").Value = CycleTime
            worksheet.Range("F7").Value = TesterID
            worksheet.Range("F8").Value = Mark
            worksheet.Range("G7").Value = ErrorMsg
            worksheet.Range("G13").Value = DutConnect
            worksheet.Range("G14").Value = SetPower
            worksheet.Range("G15").Value = UniflashProgramme
            worksheet.Range("G16").Value = FTDIProg
            worksheet.Range("G18").Value = InitialPower
            worksheet.Range("G19").Value = Can6705v
            worksheet.Range("G20").Value = Isolatedv12
            worksheet.Range("G21").Value = Sensorv12
            worksheet.Range("G22").Value = TT6205v
            worksheet.Range("G23").Value = IT6615v
            worksheet.Range("G24").Value = A0691
            worksheet.Range("G25").Value = D0691
            worksheet.Range("G27").Value = D0691ON
            worksheet.Range("G28").Value = D0661ON
            worksheet.Range("G29").Value = D0661OFF
            worksheet.Range("G30").Value = AUXENCMDON
            worksheet.Range("G31").Value = D0100ON
            worksheet.Range("G32").Value = D0100OFF
            worksheet.Range("G33").Value = D0101ON
            worksheet.Range("G34").Value = D0101OFF
            worksheet.Range("G35").Value = CMDOFF
            worksheet.Range("G36").Value = SW2ON
            worksheet.Range("G37").Value = SW3ON
            worksheet.Range("G38").Value = TOOGLECMDON
            worksheet.Range("G39").Value = CAN1SILENCEON
            worksheet.Range("G40").Value = CAN2SILENCEON
            Dim worksheet1 As Excel.Worksheet = workbook.Sheets.Add(After:=worksheet)
            worksheet1.Name = "1"
            Dim worksheet2 As Excel.Worksheet = workbook.Sheets.Add(After:=worksheet1)
            worksheet2.Name = "2"
            Dim worksheet3 As Excel.Worksheet = workbook.Sheets.Add(After:=worksheet2)
            worksheet3.Name = "3"
            Dim worksheet4 As Excel.Worksheet = workbook.Sheets.Add(After:=worksheet3)
            worksheet4.Name = "4"
            Dim worksheet5 As Excel.Worksheet = workbook.Sheets.Add(After:=worksheet4)
            worksheet5.Name = "5"
            Dim worksheet6 As Excel.Worksheet = workbook.Sheets.Add(After:=worksheet5)
            worksheet6.Name = "6"
            Dim worksheet7 As Excel.Worksheet = workbook.Sheets.Add(After:=worksheet6)
            worksheet7.Name = "7"
            Dim worksheet8 As Excel.Worksheet = workbook.Sheets.Add(After:=worksheet7)
            worksheet8.Name = "8"
            Dim worksheet9 As Excel.Worksheet = workbook.Sheets.Add(After:=worksheet8)
            worksheet9.Name = "9"
            Dim worksheet10 As Excel.Worksheet = workbook.Sheets.Add(After:=worksheet9)
            worksheet10.Name = "10"
            Dim worksheet11 As Excel.Worksheet = workbook.Sheets.Add(After:=worksheet10)
            worksheet11.Name = "11"
            Dim worksheet12 As Excel.Worksheet = workbook.Sheets.Add(After:=worksheet11)
            worksheet12.Name = "12"
            Dim worksheet13 As Excel.Worksheet = workbook.Sheets.Add(After:=worksheet12)
            worksheet13.Name = "13"
            Dim worksheet14 As Excel.Worksheet = workbook.Sheets.Add(After:=worksheet13)
            worksheet14.Name = "14"
            Dim worksheet15 As Excel.Worksheet = workbook.Sheets.Add(After:=worksheet14)
            worksheet15.Name = "15"
            Clipboard.Clear()
            Clipboard.SetImage(IMAGE_STR.PictureBox1.Image)
            worksheet1.Paste()
            Clipboard.Clear()
            Clipboard.SetImage(IMAGE_STR.PictureBox2.Image)
            worksheet2.Paste()
            Clipboard.Clear()
            Clipboard.SetImage(IMAGE_STR.PictureBox3.Image)
            worksheet3.Paste()
            Clipboard.Clear()
            Clipboard.SetImage(IMAGE_STR.PictureBox4.Image)
            worksheet4.Paste()
            Clipboard.Clear()
            Clipboard.SetImage(IMAGE_STR.PictureBox5.Image)
            worksheet5.Paste()
            Clipboard.Clear()
            Clipboard.SetImage(IMAGE_STR.PictureBox6.Image)
            worksheet6.Paste()
            Clipboard.Clear()
            Clipboard.SetImage(IMAGE_STR.PictureBox7.Image)
            worksheet7.Paste()
            Clipboard.Clear()
            Clipboard.SetImage(IMAGE_STR.PictureBox8.Image)
            worksheet8.Paste()
            Clipboard.Clear()
            Clipboard.SetImage(IMAGE_STR.PictureBox9.Image)
            worksheet9.Paste()
            Clipboard.Clear()
            Clipboard.SetImage(IMAGE_STR.PictureBox10.Image)
            worksheet10.Paste()
            Clipboard.Clear()
            Clipboard.SetImage(IMAGE_STR.PictureBox11.Image)
            worksheet11.Paste()
            Clipboard.Clear()
            Clipboard.SetImage(IMAGE_STR.PictureBox12.Image)
            worksheet12.Paste()
            Clipboard.Clear()
            Clipboard.SetImage(IMAGE_STR.PictureBox13.Image)
            worksheet13.Paste()
            Clipboard.Clear()
            Clipboard.SetImage(IMAGE_STR.PictureBox14.Image)
            worksheet14.Paste()
            Clipboard.Clear()
            Clipboard.SetImage(IMAGE_STR.PictureBox15.Image)
            worksheet15.Paste()
            Clipboard.Clear()
            With worksheet1.PageSetup
                .Orientation = Excel.XlPageOrientation.xlLandscape
                .Zoom = False
                .FitToPagesTall = 1
                .FitToPagesWide = 1
            End With
            With worksheet2.PageSetup
                .Orientation = Excel.XlPageOrientation.xlLandscape
                .Zoom = False
                .FitToPagesTall = 1
                .FitToPagesWide = 1
            End With
            With worksheet3.PageSetup
                .Orientation = Excel.XlPageOrientation.xlLandscape
                .Zoom = False
                .FitToPagesTall = 1
                .FitToPagesWide = 1
            End With
            With worksheet4.PageSetup
                .Orientation = Excel.XlPageOrientation.xlLandscape
                .Zoom = False
                .FitToPagesTall = 1
                .FitToPagesWide = 1
            End With
            With worksheet5.PageSetup
                .Orientation = Excel.XlPageOrientation.xlLandscape
                .Zoom = False
                .FitToPagesTall = 1
                .FitToPagesWide = 1
            End With
            With worksheet6.PageSetup
                .Orientation = Excel.XlPageOrientation.xlLandscape
                .Zoom = False
                .FitToPagesTall = 1
                .FitToPagesWide = 1
            End With
            With worksheet7.PageSetup
                .Orientation = Excel.XlPageOrientation.xlLandscape
                .Zoom = False
                .FitToPagesTall = 1
                .FitToPagesWide = 1
            End With
            With worksheet8.PageSetup
                .Orientation = Excel.XlPageOrientation.xlLandscape
                .Zoom = False
                .FitToPagesTall = 1
                .FitToPagesWide = 1
            End With
            With worksheet9.PageSetup
                .Orientation = Excel.XlPageOrientation.xlLandscape
                .Zoom = False
                .FitToPagesTall = 1
                .FitToPagesWide = 1
            End With
            With worksheet10.PageSetup
                .Orientation = Excel.XlPageOrientation.xlLandscape
                .Zoom = False
                .FitToPagesTall = 1
                .FitToPagesWide = 1
            End With
            With worksheet11.PageSetup
                .Orientation = Excel.XlPageOrientation.xlLandscape
                .Zoom = False
                .FitToPagesTall = 1
                .FitToPagesWide = 1
            End With
            With worksheet12.PageSetup
                .Orientation = Excel.XlPageOrientation.xlLandscape
                .Zoom = False
                .FitToPagesTall = 1
                .FitToPagesWide = 1
            End With
            With worksheet13.PageSetup
                .Orientation = Excel.XlPageOrientation.xlLandscape
                .Zoom = False
                .FitToPagesTall = 1
                .FitToPagesWide = 1
            End With
            With worksheet14.PageSetup
                .Orientation = Excel.XlPageOrientation.xlLandscape
                .Zoom = False
                .FitToPagesTall = 1
                .FitToPagesWide = 1
            End With
            With worksheet15.PageSetup
                .Orientation = Excel.XlPageOrientation.xlLandscape
                .Zoom = False
                .FitToPagesTall = 1
                .FitToPagesWide = 1
            End With
            workbook.ExportAsFixedFormat(Excel.XlFixedFormatType.xlTypePDF, CONFIG.TextBox83.Text.ToString & SerialNumber & "_" & TestMode & "_" & Status & Format(Now(), "hhmmssmmddyy") & ".pdf", vbNormal)
            workbook.Worksheets("1").Delete()
            workbook.Worksheets("2").Delete()
            workbook.Worksheets("3").Delete()
            workbook.Worksheets("4").Delete()
            workbook.Worksheets("5").Delete()
            workbook.Worksheets("6").Delete()
            workbook.Worksheets("7").Delete()
            workbook.Worksheets("8").Delete()
            workbook.Worksheets("9").Delete()
            workbook.Worksheets("10").Delete()
            workbook.Worksheets("11").Delete()
            workbook.Worksheets("12").Delete()
            workbook.Worksheets("14").Delete()
            workbook.Worksheets("15").Delete()
            workbook.Close(True)
            System.Runtime.InteropServices.Marshal.ReleaseComObject(workbook)
            workbook = Nothing
            xlapp.Quit()
            System.Runtime.InteropServices.Marshal.ReleaseComObject(xlapp)
            xlapp = Nothing
            Process.GetProcessesByName("EXCEL")(0).Kill()
            file1.Close()
            MAIN.ListBox1.Items.Add($"{DateTime.Now.ToString} Test Report Generated success : " & CONFIG.TextBox83.Text.ToString & SerialNumber & "_" & TestMode & "_" & Status & Format(Now(), "hhmmssmmddyy"))
            MAIN.ListBox1.SelectedIndex = MAIN.ListBox1.Items.Count - 1
            Dim SW As IO.StreamWriter = IO.File.CreateText(CONFIG.TextBox83.Text.ToString & SerialNumber)
            For Each S As String In MAIN.ListBox1.Items
                SW.WriteLine(S, True)
            Next
            SW.Close()
            file1.Close()
        Catch ex As Exception
            Dim iret As Object = MsgBox("Error generating Report : " & ex.Message, vbCritical + vbRetryCancel, "Error executing SMC.GenerateReportOnerrorInit")
            If iret = vbRetry Then
                file1.Close()
                GoTo line1
            ElseIf iret = vbCancel Then
                file1.WriteLine("Error generating Report :  " & ex.Message & " " & "@" & DateTime.Now.ToString("yyyy/MM/dd/HH:mm:ss"))
                file1.Close()
                MAIN.ListBox1.Items.Add($"{DateTime.Now.ToString} Error generating Report :" & ex.Message)
                MAIN.ListBox1.SelectedIndex = MAIN.ListBox1.Items.Count - 1
                Exit Function
            End If
        End Try
    End Function
    Function GenerateReportOnErrorSMCTimeNode0(User As String, TestMode As String, SerialNumber As String, Status As String, CycleTime As String, TesterID As String, Mark As String, ErrorMsg As String, DutConnect As String, SetPower As String, UniflashProgramme As String, FTDIProg As String, InitialPower As String, Can6705v As String, Isolatedv12 As String, Sensorv12 As String, TT6205v As String, IT6615v As String, A0691 As String, D0691 As String, D0691ON As String, D0661ON As String, D0661OFF As String, AUXENCMDON As String, D0100ON As String, D0100OFF As String, D0101ON As String, D0101OFF As String, CMDOFF As String, SW2ON As String, SW3ON As String, TOOGLECMDON As String, CAN1SILENCEON As String, CAN2SILENCEON As String, TimeTest As String)
        Dim line1 As Label
line1:
        Dim path1 As String = CONFIG.TextBox4.Text
        file1 = My.Computer.FileSystem.OpenTextFileWriter(path1, True)
        Try
            Dim xlapp = New Microsoft.Office.Interop.Excel.Application
            xlapp.DisplayAlerts = False
            Dim workbook = xlapp.Workbooks.Open(CONFIG.TextBox82.Text.ToString)
            Dim worksheet As Excel.Worksheet
            worksheet = workbook.Sheets("SMC")
            worksheet.Range("D6").Value = User
            worksheet.Range("D7").Value = TestMode
            worksheet.Range("D8").Value = SerialNumber
            worksheet.Range("D9").Value = DateTime.Now.ToString
            worksheet.Range("D10").Value = Status
            worksheet.Range("F6").Value = CycleTime
            worksheet.Range("F7").Value = TesterID
            worksheet.Range("F8").Value = Mark
            worksheet.Range("G7").Value = ErrorMsg
            worksheet.Range("G13").Value = DutConnect
            worksheet.Range("G14").Value = SetPower
            worksheet.Range("G15").Value = UniflashProgramme
            worksheet.Range("G16").Value = FTDIProg
            worksheet.Range("G18").Value = InitialPower
            worksheet.Range("G19").Value = Can6705v
            worksheet.Range("G20").Value = Isolatedv12
            worksheet.Range("G21").Value = Sensorv12
            worksheet.Range("G22").Value = TT6205v
            worksheet.Range("G23").Value = IT6615v
            worksheet.Range("G24").Value = A0691
            worksheet.Range("G25").Value = D0691
            worksheet.Range("G27").Value = D0691ON
            worksheet.Range("G28").Value = D0661ON
            worksheet.Range("G29").Value = D0661OFF
            worksheet.Range("G30").Value = AUXENCMDON
            worksheet.Range("G31").Value = D0100ON
            worksheet.Range("G32").Value = D0100OFF
            worksheet.Range("G33").Value = D0101ON
            worksheet.Range("G34").Value = D0101OFF
            worksheet.Range("G35").Value = CMDOFF
            worksheet.Range("G36").Value = SW2ON
            worksheet.Range("G37").Value = SW3ON
            worksheet.Range("G38").Value = TOOGLECMDON
            worksheet.Range("G39").Value = CAN1SILENCEON
            worksheet.Range("G40").Value = CAN2SILENCEON
            worksheet.Range("G42").Value = TimeTest
            Dim worksheet1 As Excel.Worksheet = workbook.Sheets.Add(After:=worksheet)
            worksheet1.Name = "1"
            Dim worksheet2 As Excel.Worksheet = workbook.Sheets.Add(After:=worksheet1)
            worksheet2.Name = "2"
            Dim worksheet3 As Excel.Worksheet = workbook.Sheets.Add(After:=worksheet2)
            worksheet3.Name = "3"
            Dim worksheet4 As Excel.Worksheet = workbook.Sheets.Add(After:=worksheet3)
            worksheet4.Name = "4"
            Dim worksheet5 As Excel.Worksheet = workbook.Sheets.Add(After:=worksheet4)
            worksheet5.Name = "5"
            Dim worksheet6 As Excel.Worksheet = workbook.Sheets.Add(After:=worksheet5)
            worksheet6.Name = "6"
            Dim worksheet7 As Excel.Worksheet = workbook.Sheets.Add(After:=worksheet6)
            worksheet7.Name = "7"
            Dim worksheet8 As Excel.Worksheet = workbook.Sheets.Add(After:=worksheet7)
            worksheet8.Name = "8"
            Dim worksheet9 As Excel.Worksheet = workbook.Sheets.Add(After:=worksheet8)
            worksheet9.Name = "9"
            Dim worksheet10 As Excel.Worksheet = workbook.Sheets.Add(After:=worksheet9)
            worksheet10.Name = "10"
            Dim worksheet11 As Excel.Worksheet = workbook.Sheets.Add(After:=worksheet10)
            worksheet11.Name = "11"
            Dim worksheet12 As Excel.Worksheet = workbook.Sheets.Add(After:=worksheet11)
            worksheet12.Name = "12"
            Dim worksheet13 As Excel.Worksheet = workbook.Sheets.Add(After:=worksheet12)
            worksheet13.Name = "13"
            Dim worksheet14 As Excel.Worksheet = workbook.Sheets.Add(After:=worksheet13)
            worksheet14.Name = "14"
            Dim worksheet15 As Excel.Worksheet = workbook.Sheets.Add(After:=worksheet14)
            worksheet15.Name = "15"
            Dim worksheet16 As Excel.Worksheet = workbook.Sheets.Add(After:=worksheet15)
            worksheet16.Name = "16"
            Clipboard.Clear()
            Clipboard.SetImage(IMAGE_STR.PictureBox1.Image)
            worksheet1.Paste()
            Clipboard.Clear()
            Clipboard.SetImage(IMAGE_STR.PictureBox2.Image)
            worksheet2.Paste()
            Clipboard.Clear()
            Clipboard.SetImage(IMAGE_STR.PictureBox3.Image)
            worksheet3.Paste()
            Clipboard.Clear()
            Clipboard.SetImage(IMAGE_STR.PictureBox4.Image)
            worksheet4.Paste()
            Clipboard.Clear()
            Clipboard.SetImage(IMAGE_STR.PictureBox5.Image)
            worksheet5.Paste()
            Clipboard.Clear()
            Clipboard.SetImage(IMAGE_STR.PictureBox6.Image)
            worksheet6.Paste()
            Clipboard.Clear()
            Clipboard.SetImage(IMAGE_STR.PictureBox7.Image)
            worksheet7.Paste()
            Clipboard.Clear()
            Clipboard.SetImage(IMAGE_STR.PictureBox8.Image)
            worksheet8.Paste()
            Clipboard.Clear()
            Clipboard.SetImage(IMAGE_STR.PictureBox9.Image)
            worksheet9.Paste()
            Clipboard.Clear()
            Clipboard.SetImage(IMAGE_STR.PictureBox10.Image)
            worksheet10.Paste()
            Clipboard.Clear()
            Clipboard.SetImage(IMAGE_STR.PictureBox11.Image)
            worksheet11.Paste()
            Clipboard.Clear()
            Clipboard.SetImage(IMAGE_STR.PictureBox12.Image)
            worksheet12.Paste()
            Clipboard.Clear()
            Clipboard.SetImage(IMAGE_STR.PictureBox13.Image)
            worksheet13.Paste()
            Clipboard.Clear()
            Clipboard.SetImage(IMAGE_STR.PictureBox14.Image)
            worksheet14.Paste()
            Clipboard.Clear()
            Clipboard.SetImage(IMAGE_STR.PictureBox15.Image)
            worksheet15.Paste()
            Clipboard.Clear()
            Clipboard.SetImage(IMAGE_STR.PictureBox16.Image)
            worksheet16.Paste()
            Clipboard.Clear()
            With worksheet1.PageSetup
                .Orientation = Excel.XlPageOrientation.xlLandscape
                .Zoom = False
                .FitToPagesTall = 1
                .FitToPagesWide = 1
            End With
            With worksheet2.PageSetup
                .Orientation = Excel.XlPageOrientation.xlLandscape
                .Zoom = False
                .FitToPagesTall = 1
                .FitToPagesWide = 1
            End With
            With worksheet3.PageSetup
                .Orientation = Excel.XlPageOrientation.xlLandscape
                .Zoom = False
                .FitToPagesTall = 1
                .FitToPagesWide = 1
            End With
            With worksheet4.PageSetup
                .Orientation = Excel.XlPageOrientation.xlLandscape
                .Zoom = False
                .FitToPagesTall = 1
                .FitToPagesWide = 1
            End With
            With worksheet5.PageSetup
                .Orientation = Excel.XlPageOrientation.xlLandscape
                .Zoom = False
                .FitToPagesTall = 1
                .FitToPagesWide = 1
            End With
            With worksheet6.PageSetup
                .Orientation = Excel.XlPageOrientation.xlLandscape
                .Zoom = False
                .FitToPagesTall = 1
                .FitToPagesWide = 1
            End With
            With worksheet7.PageSetup
                .Orientation = Excel.XlPageOrientation.xlLandscape
                .Zoom = False
                .FitToPagesTall = 1
                .FitToPagesWide = 1
            End With
            With worksheet8.PageSetup
                .Orientation = Excel.XlPageOrientation.xlLandscape
                .Zoom = False
                .FitToPagesTall = 1
                .FitToPagesWide = 1
            End With
            With worksheet9.PageSetup
                .Orientation = Excel.XlPageOrientation.xlLandscape
                .Zoom = False
                .FitToPagesTall = 1
                .FitToPagesWide = 1
            End With
            With worksheet10.PageSetup
                .Orientation = Excel.XlPageOrientation.xlLandscape
                .Zoom = False
                .FitToPagesTall = 1
                .FitToPagesWide = 1
            End With
            With worksheet11.PageSetup
                .Orientation = Excel.XlPageOrientation.xlLandscape
                .Zoom = False
                .FitToPagesTall = 1
                .FitToPagesWide = 1
            End With
            With worksheet12.PageSetup
                .Orientation = Excel.XlPageOrientation.xlLandscape
                .Zoom = False
                .FitToPagesTall = 1
                .FitToPagesWide = 1
            End With
            With worksheet13.PageSetup
                .Orientation = Excel.XlPageOrientation.xlLandscape
                .Zoom = False
                .FitToPagesTall = 1
                .FitToPagesWide = 1
            End With
            With worksheet14.PageSetup
                .Orientation = Excel.XlPageOrientation.xlLandscape
                .Zoom = False
                .FitToPagesTall = 1
                .FitToPagesWide = 1
            End With
            With worksheet15.PageSetup
                .Orientation = Excel.XlPageOrientation.xlLandscape
                .Zoom = False
                .FitToPagesTall = 1
                .FitToPagesWide = 1
            End With
            With worksheet16.PageSetup
                .Orientation = Excel.XlPageOrientation.xlLandscape
                .Zoom = False
                .FitToPagesTall = 1
                .FitToPagesWide = 1
            End With
            workbook.ExportAsFixedFormat(Excel.XlFixedFormatType.xlTypePDF, CONFIG.TextBox83.Text.ToString & SerialNumber & "_" & TestMode & "_" & Status & Format(Now(), "hhmmssmmddyy") & ".pdf", vbNormal)
            workbook.Worksheets("1").Delete()
            workbook.Worksheets("2").Delete()
            workbook.Worksheets("3").Delete()
            workbook.Worksheets("4").Delete()
            workbook.Worksheets("5").Delete()
            workbook.Worksheets("6").Delete()
            workbook.Worksheets("7").Delete()
            workbook.Worksheets("8").Delete()
            workbook.Worksheets("9").Delete()
            workbook.Worksheets("10").Delete()
            workbook.Worksheets("11").Delete()
            workbook.Worksheets("12").Delete()
            workbook.Worksheets("14").Delete()
            workbook.Worksheets("15").Delete()
            workbook.Worksheets("16").Delete()
            workbook.Close(True)
            System.Runtime.InteropServices.Marshal.ReleaseComObject(workbook)
            workbook = Nothing
            xlapp.Quit()
            System.Runtime.InteropServices.Marshal.ReleaseComObject(xlapp)
            xlapp = Nothing
            Process.GetProcessesByName("EXCEL")(0).Kill()
            file1.Close()
            MAIN.ListBox1.Items.Add($"{DateTime.Now.ToString} Test Report Generated success : " & CONFIG.TextBox83.Text.ToString & SerialNumber & "_" & TestMode & "_" & Status & Format(Now(), "hhmmssmmddyy"))
            MAIN.ListBox1.SelectedIndex = MAIN.ListBox1.Items.Count - 1
            Dim SW As IO.StreamWriter = IO.File.CreateText(CONFIG.TextBox83.Text.ToString & SerialNumber)
            For Each S As String In MAIN.ListBox1.Items
                SW.WriteLine(S, True)
            Next
            SW.Close()
            file1.Close()
        Catch ex As Exception
            Dim iret As Object = MsgBox("Error generating Report : " & ex.Message, vbCritical + vbRetryCancel, "Error executing SMC.GenerateReportOnerrorInit")
            If iret = vbRetry Then
                file1.Close()
                GoTo line1
            ElseIf iret = vbCancel Then
                file1.WriteLine("Error generating Report :  " & ex.Message & " " & "@" & DateTime.Now.ToString("yyyy/MM/dd/HH:mm:ss"))
                file1.Close()
                MAIN.ListBox1.Items.Add($"{DateTime.Now.ToString} Error generating Report :" & ex.Message)
                MAIN.ListBox1.SelectedIndex = MAIN.ListBox1.Items.Count - 1
                Exit Function
            End If
        End Try
    End Function
    Function GenerateReportOnErrorSMCStepperTestNode0(User As String, TestMode As String, SerialNumber As String, Status As String, CycleTime As String, TesterID As String, Mark As String, ErrorMsg As String, DutConnect As String, SetPower As String, UniflashProgramme As String, FTDIProg As String, InitialPower As String, Can6705v As String, Isolatedv12 As String, Sensorv12 As String, TT6205v As String, IT6615v As String, A0691 As String, D0691 As String, D0691ON As String, D0661ON As String, D0661OFF As String, AUXENCMDON As String, D0100ON As String, D0100OFF As String, D0101ON As String, D0101OFF As String, CMDOFF As String, SW2ON As String, SW3ON As String, TOOGLECMDON As String, CAN1SILENCEON As String, CAN2SILENCEON As String, TimeTest As String, STEPPERTEST As String)
        Dim line1 As Label
line1:
        Dim path1 As String = CONFIG.TextBox4.Text
        file1 = My.Computer.FileSystem.OpenTextFileWriter(path1, True)
        Try
            Dim xlapp = New Microsoft.Office.Interop.Excel.Application
            xlapp.DisplayAlerts = False
            Dim workbook = xlapp.Workbooks.Open(CONFIG.TextBox82.Text.ToString)
            Dim worksheet As Excel.Worksheet
            worksheet = workbook.Sheets("SMC")
            worksheet.Range("D6").Value = User
            worksheet.Range("D7").Value = TestMode
            worksheet.Range("D8").Value = SerialNumber
            worksheet.Range("D9").Value = DateTime.Now.ToString
            worksheet.Range("D10").Value = Status
            worksheet.Range("F6").Value = CycleTime
            worksheet.Range("F7").Value = TesterID
            worksheet.Range("F8").Value = Mark
            worksheet.Range("G7").Value = ErrorMsg
            worksheet.Range("G13").Value = DutConnect
            worksheet.Range("G14").Value = SetPower
            worksheet.Range("G15").Value = UniflashProgramme
            worksheet.Range("G16").Value = FTDIProg
            worksheet.Range("G18").Value = InitialPower
            worksheet.Range("G19").Value = Can6705v
            worksheet.Range("G20").Value = Isolatedv12
            worksheet.Range("G21").Value = Sensorv12
            worksheet.Range("G22").Value = TT6205v
            worksheet.Range("G23").Value = IT6615v
            worksheet.Range("G24").Value = A0691
            worksheet.Range("G25").Value = D0691
            worksheet.Range("G27").Value = D0691ON
            worksheet.Range("G28").Value = D0661ON
            worksheet.Range("G29").Value = D0661OFF
            worksheet.Range("G30").Value = AUXENCMDON
            worksheet.Range("G31").Value = D0100ON
            worksheet.Range("G32").Value = D0100OFF
            worksheet.Range("G33").Value = D0101ON
            worksheet.Range("G34").Value = D0101OFF
            worksheet.Range("G35").Value = CMDOFF
            worksheet.Range("G36").Value = SW2ON
            worksheet.Range("G37").Value = SW3ON
            worksheet.Range("G38").Value = TOOGLECMDON
            worksheet.Range("G39").Value = CAN1SILENCEON
            worksheet.Range("G40").Value = CAN2SILENCEON
            worksheet.Range("G42").Value = TimeTest
            worksheet.Range("G44").Value = STEPPERTEST
            Dim worksheet1 As Excel.Worksheet = workbook.Sheets.Add(After:=worksheet)
            worksheet1.Name = "1"
            Dim worksheet2 As Excel.Worksheet = workbook.Sheets.Add(After:=worksheet1)
            worksheet2.Name = "2"
            Dim worksheet3 As Excel.Worksheet = workbook.Sheets.Add(After:=worksheet2)
            worksheet3.Name = "3"
            Dim worksheet4 As Excel.Worksheet = workbook.Sheets.Add(After:=worksheet3)
            worksheet4.Name = "4"
            Dim worksheet5 As Excel.Worksheet = workbook.Sheets.Add(After:=worksheet4)
            worksheet5.Name = "5"
            Dim worksheet6 As Excel.Worksheet = workbook.Sheets.Add(After:=worksheet5)
            worksheet6.Name = "6"
            Dim worksheet7 As Excel.Worksheet = workbook.Sheets.Add(After:=worksheet6)
            worksheet7.Name = "7"
            Dim worksheet8 As Excel.Worksheet = workbook.Sheets.Add(After:=worksheet7)
            worksheet8.Name = "8"
            Dim worksheet9 As Excel.Worksheet = workbook.Sheets.Add(After:=worksheet8)
            worksheet9.Name = "9"
            Dim worksheet10 As Excel.Worksheet = workbook.Sheets.Add(After:=worksheet9)
            worksheet10.Name = "10"
            Dim worksheet11 As Excel.Worksheet = workbook.Sheets.Add(After:=worksheet10)
            worksheet11.Name = "11"
            Dim worksheet12 As Excel.Worksheet = workbook.Sheets.Add(After:=worksheet11)
            worksheet12.Name = "12"
            Dim worksheet13 As Excel.Worksheet = workbook.Sheets.Add(After:=worksheet12)
            worksheet13.Name = "13"
            Dim worksheet14 As Excel.Worksheet = workbook.Sheets.Add(After:=worksheet13)
            worksheet14.Name = "14"
            Dim worksheet15 As Excel.Worksheet = workbook.Sheets.Add(After:=worksheet14)
            worksheet15.Name = "15"
            Dim worksheet16 As Excel.Worksheet = workbook.Sheets.Add(After:=worksheet15)
            worksheet16.Name = "16"
            Dim worksheet17 As Excel.Worksheet = workbook.Sheets.Add(After:=worksheet16)
            worksheet17.Name = "17"
            Clipboard.Clear()
            Clipboard.SetImage(IMAGE_STR.PictureBox1.Image)
            worksheet1.Paste()
            Clipboard.Clear()
            Clipboard.SetImage(IMAGE_STR.PictureBox2.Image)
            worksheet2.Paste()
            Clipboard.Clear()
            Clipboard.SetImage(IMAGE_STR.PictureBox3.Image)
            worksheet3.Paste()
            Clipboard.Clear()
            Clipboard.SetImage(IMAGE_STR.PictureBox4.Image)
            worksheet4.Paste()
            Clipboard.Clear()
            Clipboard.SetImage(IMAGE_STR.PictureBox5.Image)
            worksheet5.Paste()
            Clipboard.Clear()
            Clipboard.SetImage(IMAGE_STR.PictureBox6.Image)
            worksheet6.Paste()
            Clipboard.Clear()
            Clipboard.SetImage(IMAGE_STR.PictureBox7.Image)
            worksheet7.Paste()
            Clipboard.Clear()
            Clipboard.SetImage(IMAGE_STR.PictureBox8.Image)
            worksheet8.Paste()
            Clipboard.Clear()
            Clipboard.SetImage(IMAGE_STR.PictureBox9.Image)
            worksheet9.Paste()
            Clipboard.Clear()
            Clipboard.SetImage(IMAGE_STR.PictureBox10.Image)
            worksheet10.Paste()
            Clipboard.Clear()
            Clipboard.SetImage(IMAGE_STR.PictureBox11.Image)
            worksheet11.Paste()
            Clipboard.Clear()
            Clipboard.SetImage(IMAGE_STR.PictureBox12.Image)
            worksheet12.Paste()
            Clipboard.Clear()
            Clipboard.SetImage(IMAGE_STR.PictureBox13.Image)
            worksheet13.Paste()
            Clipboard.Clear()
            Clipboard.SetImage(IMAGE_STR.PictureBox14.Image)
            worksheet14.Paste()
            Clipboard.Clear()
            Clipboard.SetImage(IMAGE_STR.PictureBox15.Image)
            worksheet15.Paste()
            Clipboard.Clear()
            Clipboard.SetImage(IMAGE_STR.PictureBox16.Image)
            worksheet16.Paste()
            Clipboard.Clear()
            Clipboard.SetImage(IMAGE_STR.PictureBox17.Image)
            worksheet17.Paste()
            Clipboard.Clear()
            With worksheet1.PageSetup
                .Orientation = Excel.XlPageOrientation.xlLandscape
                .Zoom = False
                .FitToPagesTall = 1
                .FitToPagesWide = 1
            End With
            With worksheet2.PageSetup
                .Orientation = Excel.XlPageOrientation.xlLandscape
                .Zoom = False
                .FitToPagesTall = 1
                .FitToPagesWide = 1
            End With
            With worksheet3.PageSetup
                .Orientation = Excel.XlPageOrientation.xlLandscape
                .Zoom = False
                .FitToPagesTall = 1
                .FitToPagesWide = 1
            End With
            With worksheet4.PageSetup
                .Orientation = Excel.XlPageOrientation.xlLandscape
                .Zoom = False
                .FitToPagesTall = 1
                .FitToPagesWide = 1
            End With
            With worksheet5.PageSetup
                .Orientation = Excel.XlPageOrientation.xlLandscape
                .Zoom = False
                .FitToPagesTall = 1
                .FitToPagesWide = 1
            End With
            With worksheet6.PageSetup
                .Orientation = Excel.XlPageOrientation.xlLandscape
                .Zoom = False
                .FitToPagesTall = 1
                .FitToPagesWide = 1
            End With
            With worksheet7.PageSetup
                .Orientation = Excel.XlPageOrientation.xlLandscape
                .Zoom = False
                .FitToPagesTall = 1
                .FitToPagesWide = 1
            End With
            With worksheet8.PageSetup
                .Orientation = Excel.XlPageOrientation.xlLandscape
                .Zoom = False
                .FitToPagesTall = 1
                .FitToPagesWide = 1
            End With
            With worksheet9.PageSetup
                .Orientation = Excel.XlPageOrientation.xlLandscape
                .Zoom = False
                .FitToPagesTall = 1
                .FitToPagesWide = 1
            End With
            With worksheet10.PageSetup
                .Orientation = Excel.XlPageOrientation.xlLandscape
                .Zoom = False
                .FitToPagesTall = 1
                .FitToPagesWide = 1
            End With
            With worksheet11.PageSetup
                .Orientation = Excel.XlPageOrientation.xlLandscape
                .Zoom = False
                .FitToPagesTall = 1
                .FitToPagesWide = 1
            End With
            With worksheet12.PageSetup
                .Orientation = Excel.XlPageOrientation.xlLandscape
                .Zoom = False
                .FitToPagesTall = 1
                .FitToPagesWide = 1
            End With
            With worksheet13.PageSetup
                .Orientation = Excel.XlPageOrientation.xlLandscape
                .Zoom = False
                .FitToPagesTall = 1
                .FitToPagesWide = 1
            End With
            With worksheet14.PageSetup
                .Orientation = Excel.XlPageOrientation.xlLandscape
                .Zoom = False
                .FitToPagesTall = 1
                .FitToPagesWide = 1
            End With
            With worksheet15.PageSetup
                .Orientation = Excel.XlPageOrientation.xlLandscape
                .Zoom = False
                .FitToPagesTall = 1
                .FitToPagesWide = 1
            End With
            With worksheet16.PageSetup
                .Orientation = Excel.XlPageOrientation.xlLandscape
                .Zoom = False
                .FitToPagesTall = 1
                .FitToPagesWide = 1
            End With
            With worksheet17.PageSetup
                .Orientation = Excel.XlPageOrientation.xlLandscape
                .Zoom = False
                .FitToPagesTall = 1
                .FitToPagesWide = 1
            End With
            workbook.ExportAsFixedFormat(Excel.XlFixedFormatType.xlTypePDF, CONFIG.TextBox83.Text.ToString & SerialNumber & "_" & TestMode & "_" & Status & Format(Now(), "hhmmssmmddyy") & ".pdf", vbNormal)
            workbook.Worksheets("1").Delete()
            workbook.Worksheets("2").Delete()
            workbook.Worksheets("3").Delete()
            workbook.Worksheets("4").Delete()
            workbook.Worksheets("5").Delete()
            workbook.Worksheets("6").Delete()
            workbook.Worksheets("7").Delete()
            workbook.Worksheets("8").Delete()
            workbook.Worksheets("9").Delete()
            workbook.Worksheets("10").Delete()
            workbook.Worksheets("11").Delete()
            workbook.Worksheets("12").Delete()
            workbook.Worksheets("14").Delete()
            workbook.Worksheets("15").Delete()
            workbook.Worksheets("16").Delete()
            workbook.Worksheets("17").Delete()
            workbook.Close(True)
            System.Runtime.InteropServices.Marshal.ReleaseComObject(workbook)
            workbook = Nothing
            xlapp.Quit()
            System.Runtime.InteropServices.Marshal.ReleaseComObject(xlapp)
            xlapp = Nothing
            Process.GetProcessesByName("EXCEL")(0).Kill()
            file1.Close()
            MAIN.ListBox1.Items.Add($"{DateTime.Now.ToString} Test Report Generated success : " & CONFIG.TextBox83.Text.ToString & SerialNumber & "_" & TestMode & "_" & Status & Format(Now(), "hhmmssmmddyy"))
            MAIN.ListBox1.SelectedIndex = MAIN.ListBox1.Items.Count - 1
            Dim SW As IO.StreamWriter = IO.File.CreateText(CONFIG.TextBox83.Text.ToString & SerialNumber)
            For Each S As String In MAIN.ListBox1.Items
                SW.WriteLine(S, True)
            Next
            SW.Close()
            file1.Close()
        Catch ex As Exception
            Dim iret As Object = MsgBox("Error generating Report : " & ex.Message, vbCritical + vbRetryCancel, "Error executing SMC.GenerateReportOnerrorInit")
            If iret = vbRetry Then
                file1.Close()
                GoTo line1
            ElseIf iret = vbCancel Then
                file1.WriteLine("Error generating Report :  " & ex.Message & " " & "@" & DateTime.Now.ToString("yyyy/MM/dd/HH:mm:ss"))
                file1.Close()
                MAIN.ListBox1.Items.Add($"{DateTime.Now.ToString} Error generating Report :" & ex.Message)
                MAIN.ListBox1.SelectedIndex = MAIN.ListBox1.Items.Count - 1
                Exit Function
            End If
        End Try
    End Function
    Function GenerateReportOnErrorSMCShutDownNod0(User As String, TestMode As String, SerialNumber As String, Status As String, CycleTime As String, TesterID As String, Mark As String, ErrorMsg As String, DutConnect As String, SetPower As String, UniflashProgramme As String, FTDIProg As String, InitialPower As String, Can6705v As String, Isolatedv12 As String, Sensorv12 As String, TT6205v As String, IT6615v As String, A0691 As String, D0691 As String, D0691ON As String, D0661ON As String, D0661OFF As String, AUXENCMDON As String, D0100ON As String, D0100OFF As String, D0101ON As String, D0101OFF As String, CMDOFF As String, SW2ON As String, SW3ON As String, TOOGLECMDON As String, CAN1SILENCEON As String, CAN2SILENCEON As String, TimeTest As String, STEPPERTEST As String, SHUTDOWNTEST As String)
        Dim line1 As Label
line1:
        Dim path1 As String = CONFIG.TextBox4.Text
        file1 = My.Computer.FileSystem.OpenTextFileWriter(path1, True)
        Try
            Dim xlapp = New Microsoft.Office.Interop.Excel.Application
            xlapp.DisplayAlerts = False
            Dim workbook = xlapp.Workbooks.Open(CONFIG.TextBox82.Text.ToString)
            Dim worksheet As Excel.Worksheet
            worksheet = workbook.Sheets("SMC")
            worksheet.Range("D6").Value = User
            worksheet.Range("D7").Value = TestMode
            worksheet.Range("D8").Value = SerialNumber
            worksheet.Range("D9").Value = DateTime.Now.ToString
            worksheet.Range("D10").Value = Status
            worksheet.Range("F6").Value = CycleTime
            worksheet.Range("F7").Value = TesterID
            worksheet.Range("F8").Value = Mark
            worksheet.Range("G7").Value = ErrorMsg
            worksheet.Range("G13").Value = DutConnect
            worksheet.Range("G14").Value = SetPower
            worksheet.Range("G15").Value = UniflashProgramme
            worksheet.Range("G16").Value = FTDIProg
            worksheet.Range("G18").Value = InitialPower
            worksheet.Range("G19").Value = Can6705v
            worksheet.Range("G20").Value = Isolatedv12
            worksheet.Range("G21").Value = Sensorv12
            worksheet.Range("G22").Value = TT6205v
            worksheet.Range("G23").Value = IT6615v
            worksheet.Range("G24").Value = A0691
            worksheet.Range("G25").Value = D0691
            worksheet.Range("G27").Value = D0691ON
            worksheet.Range("G28").Value = D0661ON
            worksheet.Range("G29").Value = D0661OFF
            worksheet.Range("G30").Value = AUXENCMDON
            worksheet.Range("G31").Value = D0100ON
            worksheet.Range("G32").Value = D0100OFF
            worksheet.Range("G33").Value = D0101ON
            worksheet.Range("G34").Value = D0101OFF
            worksheet.Range("G35").Value = CMDOFF
            worksheet.Range("G36").Value = SW2ON
            worksheet.Range("G37").Value = SW3ON
            worksheet.Range("G38").Value = TOOGLECMDON
            worksheet.Range("G39").Value = CAN1SILENCEON
            worksheet.Range("G40").Value = CAN2SILENCEON
            worksheet.Range("G42").Value = TimeTest
            worksheet.Range("G44").Value = STEPPERTEST
            worksheet.Range("G46").Value = SHUTDOWNTEST
            Dim worksheet1 As Excel.Worksheet = workbook.Sheets.Add(After:=worksheet)
            worksheet1.Name = "1"
            Dim worksheet2 As Excel.Worksheet = workbook.Sheets.Add(After:=worksheet1)
            worksheet2.Name = "2"
            Dim worksheet3 As Excel.Worksheet = workbook.Sheets.Add(After:=worksheet2)
            worksheet3.Name = "3"
            Dim worksheet4 As Excel.Worksheet = workbook.Sheets.Add(After:=worksheet3)
            worksheet4.Name = "4"
            Dim worksheet5 As Excel.Worksheet = workbook.Sheets.Add(After:=worksheet4)
            worksheet5.Name = "5"
            Dim worksheet6 As Excel.Worksheet = workbook.Sheets.Add(After:=worksheet5)
            worksheet6.Name = "6"
            Dim worksheet7 As Excel.Worksheet = workbook.Sheets.Add(After:=worksheet6)
            worksheet7.Name = "7"
            Dim worksheet8 As Excel.Worksheet = workbook.Sheets.Add(After:=worksheet7)
            worksheet8.Name = "8"
            Dim worksheet9 As Excel.Worksheet = workbook.Sheets.Add(After:=worksheet8)
            worksheet9.Name = "9"
            Dim worksheet10 As Excel.Worksheet = workbook.Sheets.Add(After:=worksheet9)
            worksheet10.Name = "10"
            Dim worksheet11 As Excel.Worksheet = workbook.Sheets.Add(After:=worksheet10)
            worksheet11.Name = "11"
            Dim worksheet12 As Excel.Worksheet = workbook.Sheets.Add(After:=worksheet11)
            worksheet12.Name = "12"
            Dim worksheet13 As Excel.Worksheet = workbook.Sheets.Add(After:=worksheet12)
            worksheet13.Name = "13"
            Dim worksheet14 As Excel.Worksheet = workbook.Sheets.Add(After:=worksheet13)
            worksheet14.Name = "14"
            Dim worksheet15 As Excel.Worksheet = workbook.Sheets.Add(After:=worksheet14)
            worksheet15.Name = "15"
            Dim worksheet16 As Excel.Worksheet = workbook.Sheets.Add(After:=worksheet15)
            worksheet16.Name = "16"
            Dim worksheet17 As Excel.Worksheet = workbook.Sheets.Add(After:=worksheet16)
            worksheet17.Name = "17"
            Dim worksheet18 As Excel.Worksheet = workbook.Sheets.Add(After:=worksheet17)
            worksheet18.Name = "18"
            Clipboard.Clear()
            Clipboard.SetImage(IMAGE_STR.PictureBox1.Image)
            worksheet1.Paste()
            Clipboard.Clear()
            Clipboard.SetImage(IMAGE_STR.PictureBox2.Image)
            worksheet2.Paste()
            Clipboard.Clear()
            Clipboard.SetImage(IMAGE_STR.PictureBox3.Image)
            worksheet3.Paste()
            Clipboard.Clear()
            Clipboard.SetImage(IMAGE_STR.PictureBox4.Image)
            worksheet4.Paste()
            Clipboard.Clear()
            Clipboard.SetImage(IMAGE_STR.PictureBox5.Image)
            worksheet5.Paste()
            Clipboard.Clear()
            Clipboard.SetImage(IMAGE_STR.PictureBox6.Image)
            worksheet6.Paste()
            Clipboard.Clear()
            Clipboard.SetImage(IMAGE_STR.PictureBox7.Image)
            worksheet7.Paste()
            Clipboard.Clear()
            Clipboard.SetImage(IMAGE_STR.PictureBox8.Image)
            worksheet8.Paste()
            Clipboard.Clear()
            Clipboard.SetImage(IMAGE_STR.PictureBox9.Image)
            worksheet9.Paste()
            Clipboard.Clear()
            Clipboard.SetImage(IMAGE_STR.PictureBox10.Image)
            worksheet10.Paste()
            Clipboard.Clear()
            Clipboard.SetImage(IMAGE_STR.PictureBox11.Image)
            worksheet11.Paste()
            Clipboard.Clear()
            Clipboard.SetImage(IMAGE_STR.PictureBox12.Image)
            worksheet12.Paste()
            Clipboard.Clear()
            Clipboard.SetImage(IMAGE_STR.PictureBox13.Image)
            worksheet13.Paste()
            Clipboard.Clear()
            Clipboard.SetImage(IMAGE_STR.PictureBox14.Image)
            worksheet14.Paste()
            Clipboard.Clear()
            Clipboard.SetImage(IMAGE_STR.PictureBox15.Image)
            worksheet15.Paste()
            Clipboard.Clear()
            Clipboard.SetImage(IMAGE_STR.PictureBox16.Image)
            worksheet16.Paste()
            Clipboard.Clear()
            Clipboard.SetImage(IMAGE_STR.PictureBox17.Image)
            worksheet17.Paste()
            Clipboard.Clear()
            Clipboard.SetImage(IMAGE_STR.PictureBox18.Image)
            worksheet18.Paste()
            Clipboard.Clear()
            With worksheet1.PageSetup
                .Orientation = Excel.XlPageOrientation.xlLandscape
                .Zoom = False
                .FitToPagesTall = 1
                .FitToPagesWide = 1
            End With
            With worksheet2.PageSetup
                .Orientation = Excel.XlPageOrientation.xlLandscape
                .Zoom = False
                .FitToPagesTall = 1
                .FitToPagesWide = 1
            End With
            With worksheet3.PageSetup
                .Orientation = Excel.XlPageOrientation.xlLandscape
                .Zoom = False
                .FitToPagesTall = 1
                .FitToPagesWide = 1
            End With
            With worksheet4.PageSetup
                .Orientation = Excel.XlPageOrientation.xlLandscape
                .Zoom = False
                .FitToPagesTall = 1
                .FitToPagesWide = 1
            End With
            With worksheet5.PageSetup
                .Orientation = Excel.XlPageOrientation.xlLandscape
                .Zoom = False
                .FitToPagesTall = 1
                .FitToPagesWide = 1
            End With
            With worksheet6.PageSetup
                .Orientation = Excel.XlPageOrientation.xlLandscape
                .Zoom = False
                .FitToPagesTall = 1
                .FitToPagesWide = 1
            End With
            With worksheet7.PageSetup
                .Orientation = Excel.XlPageOrientation.xlLandscape
                .Zoom = False
                .FitToPagesTall = 1
                .FitToPagesWide = 1
            End With
            With worksheet8.PageSetup
                .Orientation = Excel.XlPageOrientation.xlLandscape
                .Zoom = False
                .FitToPagesTall = 1
                .FitToPagesWide = 1
            End With
            With worksheet9.PageSetup
                .Orientation = Excel.XlPageOrientation.xlLandscape
                .Zoom = False
                .FitToPagesTall = 1
                .FitToPagesWide = 1
            End With
            With worksheet10.PageSetup
                .Orientation = Excel.XlPageOrientation.xlLandscape
                .Zoom = False
                .FitToPagesTall = 1
                .FitToPagesWide = 1
            End With
            With worksheet11.PageSetup
                .Orientation = Excel.XlPageOrientation.xlLandscape
                .Zoom = False
                .FitToPagesTall = 1
                .FitToPagesWide = 1
            End With
            With worksheet12.PageSetup
                .Orientation = Excel.XlPageOrientation.xlLandscape
                .Zoom = False
                .FitToPagesTall = 1
                .FitToPagesWide = 1
            End With
            With worksheet13.PageSetup
                .Orientation = Excel.XlPageOrientation.xlLandscape
                .Zoom = False
                .FitToPagesTall = 1
                .FitToPagesWide = 1
            End With
            With worksheet14.PageSetup
                .Orientation = Excel.XlPageOrientation.xlLandscape
                .Zoom = False
                .FitToPagesTall = 1
                .FitToPagesWide = 1
            End With
            With worksheet15.PageSetup
                .Orientation = Excel.XlPageOrientation.xlLandscape
                .Zoom = False
                .FitToPagesTall = 1
                .FitToPagesWide = 1
            End With
            With worksheet16.PageSetup
                .Orientation = Excel.XlPageOrientation.xlLandscape
                .Zoom = False
                .FitToPagesTall = 1
                .FitToPagesWide = 1
            End With
            With worksheet17.PageSetup
                .Orientation = Excel.XlPageOrientation.xlLandscape
                .Zoom = False
                .FitToPagesTall = 1
                .FitToPagesWide = 1
            End With
            With worksheet18.PageSetup
                .Orientation = Excel.XlPageOrientation.xlLandscape
                .Zoom = False
                .FitToPagesTall = 1
                .FitToPagesWide = 1
            End With
            workbook.ExportAsFixedFormat(Excel.XlFixedFormatType.xlTypePDF, CONFIG.TextBox83.Text.ToString & SerialNumber & "_" & TestMode & "_" & Status & Format(Now(), "hhmmssmmddyy") & ".pdf", vbNormal)
            workbook.Worksheets("1").Delete()
            workbook.Worksheets("2").Delete()
            workbook.Worksheets("3").Delete()
            workbook.Worksheets("4").Delete()
            workbook.Worksheets("5").Delete()
            workbook.Worksheets("6").Delete()
            workbook.Worksheets("7").Delete()
            workbook.Worksheets("8").Delete()
            workbook.Worksheets("9").Delete()
            workbook.Worksheets("10").Delete()
            workbook.Worksheets("11").Delete()
            workbook.Worksheets("12").Delete()
            workbook.Worksheets("14").Delete()
            workbook.Worksheets("15").Delete()
            workbook.Worksheets("16").Delete()
            workbook.Worksheets("17").Delete()
            workbook.Worksheets("18").Delete()
            workbook.Close(True)
            System.Runtime.InteropServices.Marshal.ReleaseComObject(workbook)
            workbook = Nothing
            xlapp.Quit()
            System.Runtime.InteropServices.Marshal.ReleaseComObject(xlapp)
            xlapp = Nothing
            Process.GetProcessesByName("EXCEL")(0).Kill()
            file1.Close()
            MAIN.ListBox1.Items.Add($"{DateTime.Now.ToString} Test Report Generated success : " & CONFIG.TextBox83.Text.ToString & SerialNumber & "_" & TestMode & "_" & Status & Format(Now(), "hhmmssmmddyy"))
            MAIN.ListBox1.SelectedIndex = MAIN.ListBox1.Items.Count - 1
            Dim SW As IO.StreamWriter = IO.File.CreateText(CONFIG.TextBox83.Text.ToString & SerialNumber)
            For Each S As String In MAIN.ListBox1.Items
                SW.WriteLine(S, True)
            Next
            SW.Close()
            file1.Close()
        Catch ex As Exception
            Dim iret As Object = MsgBox("Error generating Report : " & ex.Message, vbCritical + vbRetryCancel, "Error executing SMC.GenerateReportOnerrorInit")
            If iret = vbRetry Then
                file1.Close()
                GoTo line1
            ElseIf iret = vbCancel Then
                file1.WriteLine("Error generating Report :  " & ex.Message & " " & "@" & DateTime.Now.ToString("yyyy/MM/dd/HH:mm:ss"))
                file1.Close()
                MAIN.ListBox1.Items.Add($"{DateTime.Now.ToString} Error generating Report :" & ex.Message)
                MAIN.ListBox1.SelectedIndex = MAIN.ListBox1.Items.Count - 1
                Exit Function
            End If
        End Try
    End Function
    Function GenerateReport(User As String, TestMode As String, SerialNumber As String, Status As String, CycleTime As String, TesterID As String, Mark As String, ErrorMsg As String, DutConnect As String, SetPower As String, UniflashProgramme As String, FTDIProg As String, InitialPower As String, Can6705v As String, Isolatedv12 As String, Sensorv12 As String, TT6205v As String, IT6615v As String, A0691 As String, D0691 As String, D0691ON As String, D0661ON As String, D0661OFF As String, AUXENCMDON As String, D0100ON As String, D0100OFF As String, D0101ON As String, D0101OFF As String, CMDOFF As String, SW2ON As String, SW3ON As String, TOOGLECMDON As String, CAN1SILENCEON As String, CAN2SILENCEON As String, TimeTest As String, STEPPERTEST As String, SHUTDOWNTEST As String)
        Dim line1 As Label
line1:
        Dim path1 As String = CONFIG.TextBox4.Text
        file1 = My.Computer.FileSystem.OpenTextFileWriter(path1, True)
        Try
            Dim xlapp = New Microsoft.Office.Interop.Excel.Application
            xlapp.DisplayAlerts = False
            Dim workbook = xlapp.Workbooks.Open(CONFIG.TextBox82.Text.ToString)
            Dim worksheet As Excel.Worksheet
            worksheet = workbook.Sheets("SMC")
            worksheet.Range("D6").Value = User
            worksheet.Range("D7").Value = TestMode
            worksheet.Range("D8").Value = SerialNumber
            worksheet.Range("D9").Value = DateTime.Now.ToString
            worksheet.Range("D10").Value = Status
            worksheet.Range("F6").Value = CycleTime
            worksheet.Range("F7").Value = TesterID
            worksheet.Range("F8").Value = Mark
            worksheet.Range("G13").Value = DutConnect
            worksheet.Range("G14").Value = SetPower
            worksheet.Range("G15").Value = UniflashProgramme
            worksheet.Range("G16").Value = FTDIProg
            worksheet.Range("G18").Value = InitialPower
            worksheet.Range("G19").Value = Can6705v
            worksheet.Range("G20").Value = Isolatedv12
            worksheet.Range("G21").Value = Sensorv12
            worksheet.Range("G22").Value = TT6205v
            worksheet.Range("G23").Value = IT6615v
            worksheet.Range("G24").Value = A0691
            worksheet.Range("G25").Value = D0691
            worksheet.Range("G27").Value = D0691ON
            worksheet.Range("G28").Value = D0661ON
            worksheet.Range("G29").Value = D0661OFF
            worksheet.Range("G30").Value = AUXENCMDON
            worksheet.Range("G31").Value = D0100ON
            worksheet.Range("G32").Value = D0100OFF
            worksheet.Range("G33").Value = D0101ON
            worksheet.Range("G34").Value = D0101OFF
            worksheet.Range("G35").Value = CMDOFF
            worksheet.Range("G36").Value = SW2ON
            worksheet.Range("G37").Value = SW3ON
            worksheet.Range("G38").Value = TOOGLECMDON
            worksheet.Range("G39").Value = CAN1SILENCEON
            worksheet.Range("G40").Value = CAN2SILENCEON
            worksheet.Range("G42").Value = TimeTest
            worksheet.Range("G44").Value = STEPPERTEST
            worksheet.Range("G46").Value = SHUTDOWNTEST
            Dim worksheet1 As Excel.Worksheet = workbook.Sheets.Add(After:=worksheet)
            worksheet1.Name = "1"
            Dim worksheet2 As Excel.Worksheet = workbook.Sheets.Add(After:=worksheet1)
            worksheet2.Name = "2"
            Dim worksheet3 As Excel.Worksheet = workbook.Sheets.Add(After:=worksheet2)
            worksheet3.Name = "3"
            Dim worksheet4 As Excel.Worksheet = workbook.Sheets.Add(After:=worksheet3)
            worksheet4.Name = "4"
            Dim worksheet5 As Excel.Worksheet = workbook.Sheets.Add(After:=worksheet4)
            worksheet5.Name = "5"
            Dim worksheet6 As Excel.Worksheet = workbook.Sheets.Add(After:=worksheet5)
            worksheet6.Name = "6"
            Dim worksheet7 As Excel.Worksheet = workbook.Sheets.Add(After:=worksheet6)
            worksheet7.Name = "7"
            Dim worksheet8 As Excel.Worksheet = workbook.Sheets.Add(After:=worksheet7)
            worksheet8.Name = "8"
            Dim worksheet9 As Excel.Worksheet = workbook.Sheets.Add(After:=worksheet8)
            worksheet9.Name = "9"
            Dim worksheet10 As Excel.Worksheet = workbook.Sheets.Add(After:=worksheet9)
            worksheet10.Name = "10"
            Dim worksheet11 As Excel.Worksheet = workbook.Sheets.Add(After:=worksheet10)
            worksheet11.Name = "11"
            Dim worksheet12 As Excel.Worksheet = workbook.Sheets.Add(After:=worksheet11)
            worksheet12.Name = "12"
            Dim worksheet13 As Excel.Worksheet = workbook.Sheets.Add(After:=worksheet12)
            worksheet13.Name = "13"
            Dim worksheet14 As Excel.Worksheet = workbook.Sheets.Add(After:=worksheet13)
            worksheet14.Name = "14"
            Dim worksheet15 As Excel.Worksheet = workbook.Sheets.Add(After:=worksheet14)
            worksheet15.Name = "15"
            Dim worksheet16 As Excel.Worksheet = workbook.Sheets.Add(After:=worksheet15)
            worksheet16.Name = "16"
            Dim worksheet17 As Excel.Worksheet = workbook.Sheets.Add(After:=worksheet16)
            worksheet17.Name = "17"
            Dim worksheet18 As Excel.Worksheet = workbook.Sheets.Add(After:=worksheet17)
            worksheet18.Name = "18"
            Clipboard.Clear()
            Clipboard.SetImage(IMAGE_STR.PictureBox1.Image)
            worksheet1.Paste()
            Clipboard.Clear()
            Clipboard.SetImage(IMAGE_STR.PictureBox2.Image)
            worksheet2.Paste()
            Clipboard.Clear()
            Clipboard.SetImage(IMAGE_STR.PictureBox3.Image)
            worksheet3.Paste()
            Clipboard.Clear()
            Clipboard.SetImage(IMAGE_STR.PictureBox4.Image)
            worksheet4.Paste()
            Clipboard.Clear()
            Clipboard.SetImage(IMAGE_STR.PictureBox5.Image)
            worksheet5.Paste()
            Clipboard.Clear()
            Clipboard.SetImage(IMAGE_STR.PictureBox6.Image)
            worksheet6.Paste()
            Clipboard.Clear()
            Clipboard.SetImage(IMAGE_STR.PictureBox7.Image)
            worksheet7.Paste()
            Clipboard.Clear()
            Clipboard.SetImage(IMAGE_STR.PictureBox8.Image)
            worksheet8.Paste()
            Clipboard.Clear()
            Clipboard.SetImage(IMAGE_STR.PictureBox9.Image)
            worksheet9.Paste()
            Clipboard.Clear()
            Clipboard.SetImage(IMAGE_STR.PictureBox10.Image)
            worksheet10.Paste()
            Clipboard.Clear()
            Clipboard.SetImage(IMAGE_STR.PictureBox11.Image)
            worksheet11.Paste()
            Clipboard.Clear()
            Clipboard.SetImage(IMAGE_STR.PictureBox12.Image)
            worksheet12.Paste()
            Clipboard.Clear()
            Clipboard.SetImage(IMAGE_STR.PictureBox13.Image)
            worksheet13.Paste()
            Clipboard.Clear()
            Clipboard.SetImage(IMAGE_STR.PictureBox14.Image)
            worksheet14.Paste()
            Clipboard.Clear()
            Clipboard.SetImage(IMAGE_STR.PictureBox15.Image)
            worksheet15.Paste()
            Clipboard.Clear()
            Clipboard.SetImage(IMAGE_STR.PictureBox16.Image)
            worksheet16.Paste()
            Clipboard.Clear()
            Clipboard.SetImage(IMAGE_STR.PictureBox17.Image)
            worksheet17.Paste()
            Clipboard.Clear()
            Clipboard.SetImage(IMAGE_STR.PictureBox18.Image)
            worksheet18.Paste()
            Clipboard.Clear()
            With worksheet1.PageSetup
                .Orientation = Excel.XlPageOrientation.xlLandscape
                .Zoom = False
                .FitToPagesTall = 1
                .FitToPagesWide = 1
            End With
            With worksheet2.PageSetup
                .Orientation = Excel.XlPageOrientation.xlLandscape
                .Zoom = False
                .FitToPagesTall = 1
                .FitToPagesWide = 1
            End With
            With worksheet3.PageSetup
                .Orientation = Excel.XlPageOrientation.xlLandscape
                .Zoom = False
                .FitToPagesTall = 1
                .FitToPagesWide = 1
            End With
            With worksheet4.PageSetup
                .Orientation = Excel.XlPageOrientation.xlLandscape
                .Zoom = False
                .FitToPagesTall = 1
                .FitToPagesWide = 1
            End With
            With worksheet5.PageSetup
                .Orientation = Excel.XlPageOrientation.xlLandscape
                .Zoom = False
                .FitToPagesTall = 1
                .FitToPagesWide = 1
            End With
            With worksheet6.PageSetup
                .Orientation = Excel.XlPageOrientation.xlLandscape
                .Zoom = False
                .FitToPagesTall = 1
                .FitToPagesWide = 1
            End With
            With worksheet7.PageSetup
                .Orientation = Excel.XlPageOrientation.xlLandscape
                .Zoom = False
                .FitToPagesTall = 1
                .FitToPagesWide = 1
            End With
            With worksheet8.PageSetup
                .Orientation = Excel.XlPageOrientation.xlLandscape
                .Zoom = False
                .FitToPagesTall = 1
                .FitToPagesWide = 1
            End With
            With worksheet9.PageSetup
                .Orientation = Excel.XlPageOrientation.xlLandscape
                .Zoom = False
                .FitToPagesTall = 1
                .FitToPagesWide = 1
            End With
            With worksheet10.PageSetup
                .Orientation = Excel.XlPageOrientation.xlLandscape
                .Zoom = False
                .FitToPagesTall = 1
                .FitToPagesWide = 1
            End With
            With worksheet11.PageSetup
                .Orientation = Excel.XlPageOrientation.xlLandscape
                .Zoom = False
                .FitToPagesTall = 1
                .FitToPagesWide = 1
            End With
            With worksheet12.PageSetup
                .Orientation = Excel.XlPageOrientation.xlLandscape
                .Zoom = False
                .FitToPagesTall = 1
                .FitToPagesWide = 1
            End With
            With worksheet13.PageSetup
                .Orientation = Excel.XlPageOrientation.xlLandscape
                .Zoom = False
                .FitToPagesTall = 1
                .FitToPagesWide = 1
            End With
            With worksheet14.PageSetup
                .Orientation = Excel.XlPageOrientation.xlLandscape
                .Zoom = False
                .FitToPagesTall = 1
                .FitToPagesWide = 1
            End With
            With worksheet15.PageSetup
                .Orientation = Excel.XlPageOrientation.xlLandscape
                .Zoom = False
                .FitToPagesTall = 1
                .FitToPagesWide = 1
            End With
            With worksheet16.PageSetup
                .Orientation = Excel.XlPageOrientation.xlLandscape
                .Zoom = False
                .FitToPagesTall = 1
                .FitToPagesWide = 1
            End With
            With worksheet17.PageSetup
                .Orientation = Excel.XlPageOrientation.xlLandscape
                .Zoom = False
                .FitToPagesTall = 1
                .FitToPagesWide = 1
            End With
            With worksheet18.PageSetup
                .Orientation = Excel.XlPageOrientation.xlLandscape
                .Zoom = False
                .FitToPagesTall = 1
                .FitToPagesWide = 1
            End With
            workbook.ExportAsFixedFormat(Excel.XlFixedFormatType.xlTypePDF, CONFIG.TextBox83.Text.ToString & SerialNumber & "_" & TestMode & "_" & Status & Format(Now(), "hhmmssmmddyy") & ".pdf", vbNormal)
            workbook.Worksheets("1").Delete()
            workbook.Worksheets("2").Delete()
            workbook.Worksheets("3").Delete()
            workbook.Worksheets("4").Delete()
            workbook.Worksheets("5").Delete()
            workbook.Worksheets("6").Delete()
            workbook.Worksheets("7").Delete()
            workbook.Worksheets("8").Delete()
            workbook.Worksheets("9").Delete()
            workbook.Worksheets("10").Delete()
            workbook.Worksheets("11").Delete()
            workbook.Worksheets("12").Delete()
            workbook.Worksheets("14").Delete()
            workbook.Worksheets("15").Delete()
            workbook.Worksheets("16").Delete()
            workbook.Worksheets("17").Delete()
            workbook.Worksheets("18").Delete()
            workbook.Close(True)
            System.Runtime.InteropServices.Marshal.ReleaseComObject(workbook)
            workbook = Nothing
            xlapp.Quit()
            System.Runtime.InteropServices.Marshal.ReleaseComObject(xlapp)
            xlapp = Nothing
            Process.GetProcessesByName("EXCEL")(0).Kill()
            file1.Close()
            MAIN.ListBox1.Items.Add($"{DateTime.Now.ToString} Test Report Generated success : " & CONFIG.TextBox83.Text.ToString & SerialNumber & "_" & TestMode & "_" & Status & Format(Now(), "hhmmssmmddyy"))
            MAIN.ListBox1.SelectedIndex = MAIN.ListBox1.Items.Count - 1
            Dim SW As IO.StreamWriter = IO.File.CreateText(CONFIG.TextBox83.Text.ToString & SerialNumber)
            For Each S As String In MAIN.ListBox1.Items
                SW.WriteLine(S, True)
            Next
            SW.Close()
            file1.Close()
        Catch ex As Exception
            Dim iret As Object = MsgBox("Error generating Report : " & ex.Message, vbCritical + vbRetryCancel, "Error executing SMC.GenerateReportOnerrorInit")
            If iret = vbRetry Then
                file1.Close()
                GoTo line1
            ElseIf iret = vbCancel Then
                file1.WriteLine("Error generating Report :  " & ex.Message & " " & "@" & DateTime.Now.ToString("yyyy/MM/dd/HH:mm:ss"))
                file1.Close()
                MAIN.ListBox1.Items.Add($"{DateTime.Now.ToString} Error generating Report :" & ex.Message)
                MAIN.ListBox1.SelectedIndex = MAIN.ListBox1.Items.Count - 1
                Exit Function
            End If
        End Try
    End Function
    Function ClearReportBeforeGeneration()
        Dim line1 As Label
line1:
        Dim path1 As String = CONFIG.TextBox4.Text
        file1 = My.Computer.FileSystem.OpenTextFileWriter(path1, True)
        Try
            Dim xlapp = New Excel.Application
            Dim workbook = xlapp.Workbooks.Open(CONFIG.TextBox82.Text.ToString)
            xlapp.DisplayAlerts = False
            Dim worksheet As Excel.Worksheet
            worksheet = workbook.Sheets("SMC")
            worksheet.Range("D6", "D10").ClearContents()
            worksheet.Range("F6", "F8").ClearContents()
            worksheet.Range("G7").Value = ""
            worksheet.Range("G13", "G16").ClearContents()
            worksheet.Range("G18", "G25").ClearContents()
            worksheet.Range("G27", "G40").ClearContents()
            worksheet.Range("G42").ClearContents()
            worksheet.Range("G44").ClearContents()
            worksheet.Range("G46").ClearContents()
            workbook.Close(True)
            System.Runtime.InteropServices.Marshal.ReleaseComObject(workbook)
            workbook = Nothing
            xlapp.Quit()
            System.Runtime.InteropServices.Marshal.ReleaseComObject(xlapp)
            xlapp = Nothing
            Process.GetProcessesByName("EXCEL")(0).Kill()
            file1.Close()
        Catch ex As Exception
            Dim iret As Object = MsgBox("Error generating Report : " & ex.Message, vbCritical + vbRetryCancel, "Error executing SMC.ClearReportBeforeGeneration")
            If iret = vbRetry Then
                file1.Close()
                GoTo line1
            ElseIf iret = vbCancel Then
                file1.WriteLine("Error generating Report :  " & ex.Message & " " & "@" & DateTime.Now.ToString("yyyy/MM/dd/HH:mm:ss"))
                file1.Close()
                MAIN.ListBox1.Items.Add($"{DateTime.Now.ToString} Error generating Report :" & ex.Message)
                Exit Function
            End If
        End Try

    End Function
    Private Sub FindNode(ByVal tNode As TreeNode)
        Dim tn As TreeNode
        For Each tn In tNode.Nodes
            If tn.Text = MAIN.TextBox1.Text Then
                MAIN.TreeView1.SelectedNode = tn
                MAIN.TreeView1.SelectedNode.Expand()
                Exit For
            End If
            FindNode(tn)
        Next
    End Sub
    Private Function ImageToByteArray(image As Image) As Byte()
        Dim ms As New System.IO.MemoryStream()
        image.Save(ms, System.Drawing.Imaging.ImageFormat.Jpeg)
        Return ms.ToArray()
    End Function
    Function SMCCreateRecord(Provider As String, DataSource As String, Secqurity As String, SerialNumber As String, TestStep As String, User As String, DutID As String, Mark As String, ProgTest As String, FMWTest As String, InitPow As String, Can6705v As String, Isolatedv12 As String, sensorv12 As String, tt6205v As String, it6615v As String, a0691 As String, d0691 As String, d0691on As String, d0661on As String, d0661off As String, auxencmd As String, d0100on As String, d0100off As String, d0101on As String, d0101off As String, cmdoff As String, sw2on As String, sw3on As String, tooglecmd As String, can1 As String, can2 As String, timetest As String, steppertest As String, shutdown As String, testresult As String, Datetime As String, cycletime As String, testerid As String)
        Dim line1 As Label
line1:
        Dim path1 As String = CONFIG.TextBox4.Text
        file1 = My.Computer.FileSystem.OpenTextFileWriter(path1, True)
        Try
            cn = New OleDbConnection
            cn.ConnectionString = "Provider=" & Provider & ";Data Source=" & DataSource & "; Persist Security Info=" & Secqurity & ";"
            cn.Open()
            Dim SQL1 As String = "INSERT INTO SMC_1(SERIAL_NUMBER, TEST_STEP, DONE_BY, DUT_ID, MARK, PROGRAMMING_TEST, FIRMWARE_TEST, INITIAL_POWER, CAN670_5V, ISOLATED_V12, SENSOR_V12, TT620_5V, IT661_5V, A0691, D0691, D0691_CMD_ON, D0661_CMD_ON, D0661_CMD_OFF, AUX_EN_CMD_ON, D0100_CMD_ON, D0100_CMD_OFF, D0101_CMD_ON, D0101_CMD_OFF, CMD_SWITCH_OFF, SW2_ON, SW_3_ON, TOOGLE_CMD_ON, CAN1_ON, CAN2_ON, TIME_TEST, STEPPER_TEST, SHUTDOWN_TEST, TEST_RESULT, DATE_TIME,CYCLE_TIME, TESTER_ID ) VALUES ('" & SerialNumber & "','" & TestStep & "','" & User & "','" & DutID & "','" & Mark & "','" & ProgTest & "','" & FMWTest & "','" & InitPow & "','" & Can6705v & "','" & Isolatedv12 & "','" & sensorv12 & "','" & tt6205v & "','" & it6615v & "','" & a0691 & "','" & d0691 & "','" & d0691on & "','" & d0661on & "','" & d0661off & "','" & auxencmd & "','" & d0100on & "','" & d0100off & "','" & d0101on & "','" & d0101off & "','" & cmdoff & "','" & sw2on & "','" & sw3on & "','" & tooglecmd & "','" & can1 & "','" & can2 & "','" & timetest & "','" & steppertest & "','" & shutdown & "','" & testresult & "','" & Datetime & "','" & cycletime & "','" & testerid & "')"
            Dim command As New OleDbCommand(SQL1, cn)
            command.ExecuteNonQuery()
            MAIN.ListBox1.Items.Add("Test Record Generated Success : " & SerialNumber & TestStep & User & DutID & Mark & User & testresult & cycletime & testerid & MAIN.Label2.Text)
        Catch ex As Exception
            Dim iret1 As Object = MsgBox("ERROR WHILE CREATING TEST RECORD: " & ex.Message, vbRetryCancel, "SystemID.CreateRecord error")
            MAIN.ListBox1.Items.Add($"{MAIN.Label2.Text} Error Creating Test Report : " & ex.Message)
            If iret1 = vbRetry Then
                file1.Close()
                GoTo line1
            ElseIf iret1 = vbCancel Then
                file1.WriteLine("SystemID.CreateRecord FAILED DUE TO : " & ex.Message & " " & "@" & MAIN.Label2.Text)
                file1.Close()
                Exit Function
            End If
        End Try
    End Function
    Function ClearSequence()
        MAIN.Button3.Enabled = True
        MAIN.Button1.Enabled = False
        IMAGE_STR.PictureBox1.Image = Nothing
        IMAGE_STR.PictureBox2.Image = Nothing
        IMAGE_STR.PictureBox3.Image = Nothing
        IMAGE_STR.PictureBox4.Image = Nothing
        IMAGE_STR.PictureBox5.Image = Nothing
        IMAGE_STR.PictureBox6.Image = Nothing
        IMAGE_STR.PictureBox7.Image = Nothing
        IMAGE_STR.PictureBox8.Image = Nothing
        IMAGE_STR.PictureBox9.Image = Nothing
        IMAGE_STR.PictureBox10.Image = Nothing
        IMAGE_STR.PictureBox11.Image = Nothing
        IMAGE_STR.PictureBox12.Image = Nothing
        IMAGE_STR.PictureBox13.Image = Nothing
        IMAGE_STR.PictureBox14.Image = Nothing
        IMAGE_STR.PictureBox15.Image = Nothing
        IMAGE_STR.PictureBox16.Image = Nothing
        IMAGE_STR.PictureBox17.Image = Nothing
        IMAGE_STR.PictureBox18.Image = Nothing
        FUNNCTIONLIB.ProgressReset(True)
        MAIN.ListBox2.Items.Clear()
        For Each tb3 In SmcFUnctionButton
            tb3.visible = False
        Next
        For Each tb In SmcBox
            tb.enabled = False
        Next
        For Each tb1 In SmcTextBox
            tb1.enabled = True
            tb1.text = "IDLE"
        Next
        For Each tb2 In SmcButton
            tb2.text = "IDLE"
        Next
        MAIN.TabControl2.SelectedTab = MAIN.TabPage6
        MAIN.ListBox1.Items.Clear()
        MAIN.ListBox2.Items.Clear()
        MAIN.Button43.Text = "IDLE"
        MAIN.TimerFunction("RESET")
        MAIN.TextBox153.Clear()
        MAIN.TextBox4.Clear()
        MAIN.TextBox6.Clear()
        MAIN.GroupBox2.BackColor = Color.White
        MAIN.GroupBox3.BackColor = Color.White
        MAIN.TreeView1.BackColor = Color.White
        MAIN.TabControl1.BackColor = Color.White
        MAIN.TabControl2.BackColor = Color.White
        MAIN.BackColor = Color.White
        MAIN.GroupBox1.BackColor = Color.White
        MAIN.ListBox1.BackColor = Color.White
        MAIN.ListBox2.BackColor = Color.White
        ChangeAllNodesBackColor(MAIN.TreeView1, Color.White)
        MAIN.TreeView1.Nodes(0).Collapse()
        Dim myProcesses() As Process
        Dim myProcess As Process
        myProcesses = Process.GetProcessesByName(CONFIG.TextBox48.Text.ToString)
        If myProcesses.Length > 0 Then
            For Each myProcess In myProcesses
                If myProcess IsNot Nothing Then
                    myProcess.Kill()
                End If
            Next
        End If
    End Function
    Private Sub ChangeAllNodesBackColor(treeView As TreeView, backColor As Color)
        For Each node As TreeNode In treeView.Nodes
            ChangeNodeBackColor(node, backColor)
        Next
    End Sub
    Private Sub ChangeNodeBackColor(node As TreeNode, backColor As Color)
        node.BackColor = backColor
        For Each childNode As TreeNode In node.Nodes
            ChangeNodeBackColor(childNode, backColor)
        Next
    End Sub
    Function CalculateCount(TestResult As String)
        Dim TotalCount As Integer = MAIN.TextBox185.Text
        Dim PassCount As Integer = MAIN.TextBox184.Text
        Dim FailCount As Integer = MAIN.TextBox183.Text
        If TestResult = "PASS" Then
            TotalCount = TotalCount + 1
            PassCount = PassCount + 1
            FailCount = FailCount
        ElseIf TestResult = "FAIL" Then
            TotalCount = TotalCount + 1
            PassCount = PassCount
            FailCount = FailCount + 1
        End If
        MAIN.TextBox185.Text = TotalCount
        MAIN.TextBox184.Text = PassCount
        MAIN.TextBox183.Text = FailCount
        Dim Yield As Double = (TotalCount - FailCount) / TotalCount
        Dim formattedValue As String = (Yield * 100).ToString("0.00") & "%"
        MAIN.TextBox182.Text = formattedValue
    End Function
End Module
