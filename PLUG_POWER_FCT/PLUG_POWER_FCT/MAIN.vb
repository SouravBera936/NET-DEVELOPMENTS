Imports MadMilkman.Ini
Public Class MAIN
    Dim productversion = Application.ProductVersion.ToString
    Dim pcname = System.Net.Dns.GetHostName
    Dim StopWatch As New Diagnostics.Stopwatch
    Dim CapturedImg As Bitmap
    Dim ini As New IniFile
    Dim myProcesses() As Process
    Dim myProcess As Process

    Private Sub MAIN_Load(sender As Object, e As EventArgs) Handles MyBase.Load
        CheckBox1.Checked = False
        TextBox16.Clear()
        CheckBox2.Checked = False
        TextBox158.Text = 1
        PictureBox4.Visible = True
        GroupBox1.Visible = False
        GroupBox2.Visible = False
        GroupBox3.Visible = False
        GroupBox4.Visible = False
        GroupBox6.Visible = False
        LOGINToolStripMenuItem.Enabled = False
        ListBox1.Visible = True
        ListBox1.Items.Clear()
        LOGOUTToolStripMenuItem.Enabled = False
        EXITToolStripMenuItem.Enabled = False
        CREATEUSERToolStripMenuItem.Enabled = False
        REMOVEUSERToolStripMenuItem.Enabled = False
        CHANGEPASSWORDToolStripMenuItem.Enabled = False
        FIRSTPASSYIELDToolStripMenuItem.Enabled = False
        FINALPASSYIELDToolStripMenuItem.Enabled = False
        ConfrigurationsToolStripMenuItem.Enabled = False
        ABOUTSOFTWAREToolStripMenuItem.Enabled = False
        HELPToolStripMenuItem.Enabled = False
        SETUPToolStripMenuItem.Enabled = False
        UPDATEDESIGNATIONToolStripMenuItem.Enabled = False
        ENABLEMANUALMODEToolStripMenuItem.Enabled = False
        If FUNNCTIONLIB.LoadINI() = True Then
            ListBox1.Items.Add("LOADING CONFRIGURATION AND SETUPS Retuns : True")
            If FUNNCTIONLIB.CheckDBState(CONFIG.TextBox1.Text, CONFIG.TextBox2.Text, CONFIG.TextBox3.Text) = True Then
                ListBox1.Items.Add("CHECK DB CONNECTION STATE Returns  : True")
                If FUNNCTIONLIB.CheckRegistration(pcname, CONFIG.TextBox1.Text, CONFIG.TextBox2.Text, CONFIG.TextBox3.Text) = True Then
                    ListBox1.Items.Add("CHECK SYSTEM REGISTRATION STATUS  : True")
                    If FUNNCTIONLIB.CheckLiscence(pcname, CONFIG.TextBox1.Text, CONFIG.TextBox2.Text, CONFIG.TextBox3.Text) = True Then
                        LOGINToolStripMenuItem.Enabled = True
                        EXITToolStripMenuItem.Enabled = True
                        InitializeSoftware_beforeLogin()
                    Else
                        Dim iret1 As Object = MsgBox("SYSTEM LISCENCE EXPIRED. DO YOU WANT TO RENEW NOW?", vbQuestion + vbYesNo, "RENEW LISCENCE")
                        If iret1 = vbYes Then
                            SYSREG.ShowDialog()
                            Exit Sub
                        ElseIf iret1 = vbNo Then
                            MsgBox("CAN NOT RUN WIHOUT LISCENCE!")
                        End If
                    End If
                Else
                    ListBox1.Items.Add("CHECK SYSTEM REGISTRATION STATUS  : False [System Not Found in Database]")
                    Dim iret As Object = MsgBox("SYSTEM NOT FOUND IN REGISTRATION! DO YOU WANT TO REGISTER NOW?", vbQuestion + vbYesNo, "NEW SYSTEM REGISTRATION")
                    If iret = vbYes Then
                        ListBox1.Items.Add("Initializing New System Registration")
                        SYSREG.ShowDialog()
                        Exit Sub
                    ElseIf iret = vbNo Then
                        MsgBox("CAN NOT RUN WIHOUT REGISTRATION!")
                    End If
                End If
            Else
                ListBox1.Items.Add("CHECK DB CONNECTION STATE Retuns : " & FUNNCTIONLIB.CheckDBState(CONFIG.TextBox1.Text, CONFIG.TextBox2.Text, CONFIG.TextBox3.Text))
                ListBox1.Items.Add("EXPECTED RESULT : True")
            End If
        Else
            ListBox1.Items.Add("LOADING CONFRIGURATION AND SETPS Retuns : " & FUNNCTIONLIB.LoadINI())
            ListBox1.Items.Add("EXPECTED RESULT : True")
        End If
    End Sub
    Public Sub InitializeSoftware_beforeLogin()
        FUNNCTIONLIB.ResetYield(True)
        TextBox1.Clear()
        TextBox1.Enabled = False
        TextBox6.Clear()
        TextBox6.Enabled = False
        TextBox4.Clear()
        TextBox4.Enabled = False
        RadioButton1.Checked = False
        RadioButton1.Enabled = False
        RadioButton2.Checked = False
        RadioButton2.Enabled = False
        TextBox2.Clear()
        TextBox2.Enabled = False
        TextBox5.Clear()
        TextBox5.Text = "00:00:00"
        TextBox5.Enabled = False
        TextBox3.Enabled = False
        TextBox153.Clear()
        TextBox153.Enabled = False
        Button43.Enabled = False
        Button43.Text = "IDLE"
        Button3.Enabled = False
        Button1.Enabled = False
        FUNNCTIONLIB.ProgressReset(True)
        TreeView1.CollapseAll()
        PictureBox4.Visible = False
        Label3.Text = productversion
        Label10.Text = pcname
        Timer1.Enabled = True
        If SETUP.CheckBox5.Checked = True Then
            FUNNCTIONLIB.HideTaskbar(True)
        Else
            FUNNCTIONLIB.HideTaskbar(False)
        End If
        LOGIN.ShowDialog()

    End Sub
    Private Sub Timer1_Tick(sender As Object, e As EventArgs) Handles Timer1.Tick
        Label2.Text = Date.Now.ToString
    End Sub

    Private Sub EXITToolStripMenuItem_Click(sender As Object, e As EventArgs) Handles EXITToolStripMenuItem.Click
        Dim iret As Object = MsgBox("DO YOU WANT TO EXIT?", vbQuestion + vbYesNo, "SOFTWARE_EXIT")
        If iret = vbYes Then
            FUNNCTIONLIB.HideTaskbar(False)
            Application.Exit()
        ElseIf iret = vbNo Then
            'donothing
        End If
    End Sub

    Private Sub LOGINToolStripMenuItem_Click(sender As Object, e As EventArgs) Handles LOGINToolStripMenuItem.Click
        LOGIN.ShowDialog()
    End Sub

    Private Sub Button3_Click(sender As Object, e As EventArgs) Handles Button3.Click
        STARTUP.ShowDialog()
    End Sub

    Private Sub Button43_TextChanged(sender As Object, e As EventArgs) Handles Button43.TextChanged
        If Button43.Text = "IDLE" Then
            Button43.BackColor = Color.LightBlue
        ElseIf Button43.Text = "RUNNING" Then
            Button43.BackColor = Color.Yellow
        ElseIf Button43.Text = "PASS" Then
            Button43.BackColor = Color.LightGreen
        ElseIf Button43.Text = "FAIL" Then
            Button43.BackColor = Color.FromArgb(255, 127, 127)
        ElseIf Button43.Text = "ERROR" Then
            Button43.BackColor = Color.LightYellow
        End If
    End Sub

    Private Sub Timer2_Tick(sender As Object, e As EventArgs) Handles Timer2.Tick
        Dim elapsed As TimeSpan = Me.StopWatch.Elapsed
        TextBox5.Text = String.Format("{0:00}:{1:00}:{2:00}", Math.Floor(elapsed.TotalHours), elapsed.Minutes, elapsed.Seconds)
    End Sub
    Function TimerFunction(Operation As String)
        If Operation = "START" Then
            Timer2.Start()
            Me.StopWatch.Start()
        ElseIf Operation = "STOP" Then
            Timer2.Stop()
            Me.StopWatch.Stop()
        ElseIf Operation = "RESET" Then
            Timer2.Stop()
            Me.StopWatch.Reset()
            TextBox5.Text = "00:00:00"
        End If
    End Function
    Private Sub Button8_Click(sender As Object, e As EventArgs) Handles Button8.Click

    End Sub

    Private Sub Button8_TextChanged(sender As Object, e As EventArgs) Handles Button8.TextChanged
        If Button8.Text = "IDLE" Then
            Button8.BackColor = Color.LightBlue
        ElseIf Button8.Text = "RUNNING" Then
            Button8.BackColor = Color.Yellow
        ElseIf Button8.Text = "PASS" Then
            Button8.BackColor = Color.LightGreen
        ElseIf Button8.Text = "FAIL" Then
            Button8.BackColor = Color.FromArgb(255, 127, 127)
        ElseIf Button8.Text = "ERROR" Then
            Button8.BackColor = Color.LightYellow
        End If
    End Sub
    Private Sub TextBox158_TextChanged(sender As Object, e As EventArgs) Handles TextBox158.TextChanged
        ProgressBar1.Value = TextBox158.Text
    End Sub
    Private Sub Button15_Click(sender As Object, e As EventArgs) Handles Button15.Click
        ListBox1.Items.Add($"{DateTime.Now.ToString} RESPONSE RECIEVED FROM USER AFTER : " & TextBox5.Text)
        If RadioButton1.Checked = True Then
            Button15.Visible = False
            BatteryID.BatteryIDNode1()
        Else
            Exit Sub
        End If
    End Sub
    Private Sub FindNode(ByVal tNode As TreeNode)
        Dim tn As TreeNode
        For Each tn In tNode.Nodes
            If tn.Text = TextBox1.Text Then
                TreeView1.SelectedNode = tn
                Exit For
            End If
            FindNode(tn)
        Next
    End Sub

    Private Sub Button12_TextChanged(sender As Object, e As EventArgs) Handles Button12.TextChanged
        If Button12.Text = "IDLE" Then
            Button12.BackColor = Color.LightBlue
        ElseIf Button12.Text = "RUNNING" Then
            Button12.BackColor = Color.Yellow
        ElseIf Button12.Text = "PASS" Then
            Button12.BackColor = Color.LightGreen
        ElseIf Button12.Text = "FAIL" Then
            Button12.BackColor = Color.FromArgb(255, 127, 127)
        ElseIf Button12.Text = "ERROR" Then
            Button12.BackColor = Color.LightYellow
        End If
    End Sub

    Private Sub Button40_Click(sender As Object, e As EventArgs) Handles Button40.Click
        If RadioButton1.Checked = True Then
            ListBox1.Items.Add($"{DateTime.Now.ToString} : Response recieved from user for connecting the programmer after : {TextBox5.Text}")
            ListBox1.SelectedIndex = ListBox1.Items.Count - 1
            SMC.SMCProgrammingNode1()
        End If
    End Sub

    Private Sub Button16_TextChanged(sender As Object, e As EventArgs) Handles Button16.TextChanged
        If Button16.Text = "IDLE" Then
            Button16.BackColor = Color.LightBlue
        ElseIf Button16.Text = "RUNNING" Then
            Button16.BackColor = Color.Yellow
        ElseIf Button16.Text = "PASS" Then
            Button16.BackColor = Color.LightGreen
        ElseIf Button16.Text = "FAIL" Then
            Button16.BackColor = Color.FromArgb(255, 127, 127)
        ElseIf Button16.Text = "ERROR" Then
            Button16.BackColor = Color.LightYellow
        End If
    End Sub
    Private Sub Button17_TextChanged(sender As Object, e As EventArgs) Handles Button17.TextChanged
        If Button17.Text = "IDLE" Then
            Button17.BackColor = Color.LightBlue
        ElseIf Button17.Text = "RUNNING" Then
            Button17.BackColor = Color.Yellow
        ElseIf Button17.Text = "PASS" Then
            Button17.BackColor = Color.LightGreen
        ElseIf Button17.Text = "FAIL" Then
            Button17.BackColor = Color.FromArgb(255, 127, 127)
        ElseIf Button17.Text = "ERROR" Then
            Button17.BackColor = Color.LightYellow
        End If
    End Sub
    Private Sub Button19_TextChanged(sender As Object, e As EventArgs) Handles Button19.TextChanged
        If Button19.Text = "IDLE" Then
            Button19.BackColor = Color.LightBlue
        ElseIf Button19.Text = "RUNNING" Then
            Button19.BackColor = Color.Yellow
        ElseIf Button19.Text = "PASS" Then
            Button19.BackColor = Color.LightGreen
        ElseIf Button19.Text = "FAIL" Then
            Button19.BackColor = Color.FromArgb(255, 127, 127)
        ElseIf Button19.Text = "ERROR" Then
            Button19.BackColor = Color.LightYellow
        End If
    End Sub
    Private Sub Button18_TextChanged(sender As Object, e As EventArgs) Handles Button18.TextChanged
        If Button18.Text = "IDLE" Then
            Button18.BackColor = Color.LightBlue
        ElseIf Button18.Text = "RUNNING" Then
            Button18.BackColor = Color.Yellow
        ElseIf Button18.Text = "PASS" Then
            Button18.BackColor = Color.LightGreen
        ElseIf Button18.Text = "FAIL" Then
            Button18.BackColor = Color.FromArgb(255, 127, 127)
        ElseIf Button18.Text = "ERROR" Then
            Button18.BackColor = Color.LightYellow
        End If
    End Sub
    Private Sub Button20_TextChanged(sender As Object, e As EventArgs) Handles Button20.TextChanged
        If Button20.Text = "IDLE" Then
            Button20.BackColor = Color.LightBlue
        ElseIf Button20.Text = "RUNNING" Then
            Button20.BackColor = Color.Yellow
        ElseIf Button20.Text = "PASS" Then
            Button20.BackColor = Color.LightGreen
        ElseIf Button20.Text = "FAIL" Then
            Button20.BackColor = Color.FromArgb(255, 127, 127)
        ElseIf Button20.Text = "ERROR" Then
            Button20.BackColor = Color.LightYellow
        End If
    End Sub
    Private Sub Button21_TextChanged(sender As Object, e As EventArgs) Handles Button21.TextChanged
        If Button21.Text = "IDLE" Then
            Button21.BackColor = Color.LightBlue
        ElseIf Button21.Text = "RUNNING" Then
            Button21.BackColor = Color.Yellow
        ElseIf Button21.Text = "PASS" Then
            Button21.BackColor = Color.LightGreen
        ElseIf Button21.Text = "FAIL" Then
            Button21.BackColor = Color.FromArgb(255, 127, 127)
        ElseIf Button21.Text = "ERROR" Then
            Button21.BackColor = Color.LightYellow
        End If
    End Sub
    Private Sub Button22_TextChanged(sender As Object, e As EventArgs) Handles Button22.TextChanged
        If Button22.Text = "IDLE" Then
            Button22.BackColor = Color.LightBlue
        ElseIf Button22.Text = "RUNNING" Then
            Button22.BackColor = Color.Yellow
        ElseIf Button22.Text = "PASS" Then
            Button22.BackColor = Color.LightGreen
        ElseIf Button22.Text = "FAIL" Then
            Button22.BackColor = Color.FromArgb(255, 127, 127)
        ElseIf Button22.Text = "ERROR" Then
            Button22.BackColor = Color.LightYellow
        End If
    End Sub
    Private Sub Button23_TextChanged(sender As Object, e As EventArgs) Handles Button23.TextChanged
        If Button23.Text = "IDLE" Then
            Button23.BackColor = Color.LightBlue
        ElseIf Button23.Text = "RUNNING" Then
            Button23.BackColor = Color.Yellow
        ElseIf Button23.Text = "PASS" Then
            Button23.BackColor = Color.LightGreen
        ElseIf Button23.Text = "FAIL" Then
            Button23.BackColor = Color.FromArgb(255, 127, 127)
        ElseIf Button23.Text = "ERROR" Then
            Button23.BackColor = Color.LightYellow
        End If
    End Sub
    Private Sub Button24_TextChanged(sender As Object, e As EventArgs) Handles Button24.TextChanged
        If Button24.Text = "IDLE" Then
            Button24.BackColor = Color.LightBlue
        ElseIf Button24.Text = "RUNNING" Then
            Button24.BackColor = Color.Yellow
        ElseIf Button24.Text = "PASS" Then
            Button24.BackColor = Color.LightGreen
        ElseIf Button24.Text = "FAIL" Then
            Button24.BackColor = Color.FromArgb(255, 127, 127)
        ElseIf Button24.Text = "ERROR" Then
            Button24.BackColor = Color.LightYellow
        End If
    End Sub
    Private Sub Button25_TextChanged(sender As Object, e As EventArgs) Handles Button25.TextChanged
        If Button25.Text = "IDLE" Then
            Button25.BackColor = Color.LightBlue
        ElseIf Button25.Text = "RUNNING" Then
            Button25.BackColor = Color.Yellow
        ElseIf Button25.Text = "PASS" Then
            Button25.BackColor = Color.LightGreen
        ElseIf Button25.Text = "FAIL" Then
            Button25.BackColor = Color.FromArgb(255, 127, 127)
        ElseIf Button25.Text = "ERROR" Then
            Button25.BackColor = Color.LightYellow
        End If
    End Sub
    Private Sub Button4_TextChanged(sender As Object, e As EventArgs) Handles Button4.TextChanged
        If Button4.Text = "IDLE" Then
            Button4.BackColor = Color.LightBlue
        ElseIf Button4.Text = "RUNNING" Then
            Button4.BackColor = Color.Yellow
        ElseIf Button4.Text = "PASS" Then
            Button4.BackColor = Color.LightGreen
        ElseIf Button4.Text = "FAIL" Then
            Button4.BackColor = Color.FromArgb(255, 127, 127)
        ElseIf Button4.Text = "ERROR" Then
            Button4.BackColor = Color.LightYellow
        End If
    End Sub
    Private Sub Button26_TextChanged(sender As Object, e As EventArgs) Handles Button26.TextChanged
        If Button26.Text = "IDLE" Then
            Button26.BackColor = Color.LightBlue
        ElseIf Button26.Text = "RUNNING" Then
            Button26.BackColor = Color.Yellow
        ElseIf Button26.Text = "PASS" Then
            Button26.BackColor = Color.LightGreen
        ElseIf Button26.Text = "FAIL" Then
            Button26.BackColor = Color.FromArgb(255, 127, 127)
        ElseIf Button26.Text = "ERROR" Then
            Button26.BackColor = Color.LightYellow
        End If
    End Sub
    Private Sub Button27_TextChanged(sender As Object, e As EventArgs) Handles Button27.TextChanged
        If Button27.Text = "IDLE" Then
            Button27.BackColor = Color.LightBlue
        ElseIf Button27.Text = "RUNNING" Then
            Button27.BackColor = Color.Yellow
        ElseIf Button27.Text = "PASS" Then
            Button27.BackColor = Color.LightGreen
        ElseIf Button27.Text = "FAIL" Then
            Button27.BackColor = Color.FromArgb(255, 127, 127)
        ElseIf Button27.Text = "ERROR" Then
            Button27.BackColor = Color.LightYellow
        End If
    End Sub
    Private Sub Button28_TextChanged(sender As Object, e As EventArgs) Handles Button28.TextChanged
        If Button28.Text = "IDLE" Then
            Button28.BackColor = Color.LightBlue
        ElseIf Button28.Text = "RUNNING" Then
            Button28.BackColor = Color.Yellow
        ElseIf Button28.Text = "PASS" Then
            Button28.BackColor = Color.LightGreen
        ElseIf Button28.Text = "FAIL" Then
            Button28.BackColor = Color.FromArgb(255, 127, 127)
        ElseIf Button28.Text = "ERROR" Then
            Button28.BackColor = Color.LightYellow
        End If
    End Sub
    Private Sub Button29_TextChanged(sender As Object, e As EventArgs) Handles Button29.TextChanged
        If Button29.Text = "IDLE" Then
            Button29.BackColor = Color.LightBlue
        ElseIf Button29.Text = "RUNNING" Then
            Button29.BackColor = Color.Yellow
        ElseIf Button29.Text = "PASS" Then
            Button29.BackColor = Color.LightGreen
        ElseIf Button29.Text = "FAIL" Then
            Button29.BackColor = Color.FromArgb(255, 127, 127)
        ElseIf Button29.Text = "ERROR" Then
            Button29.BackColor = Color.LightYellow
        End If
    End Sub
    Private Sub Button30_TextChanged(sender As Object, e As EventArgs) Handles Button30.TextChanged
        If Button30.Text = "IDLE" Then
            Button30.BackColor = Color.LightBlue
        ElseIf Button30.Text = "RUNNING" Then
            Button30.BackColor = Color.Yellow
        ElseIf Button30.Text = "PASS" Then
            Button30.BackColor = Color.LightGreen
        ElseIf Button30.Text = "FAIL" Then
            Button30.BackColor = Color.FromArgb(255, 127, 127)
        ElseIf Button30.Text = "ERROR" Then
            Button30.BackColor = Color.LightYellow
        End If
    End Sub
    Private Sub Button31_TextChanged(sender As Object, e As EventArgs) Handles Button31.TextChanged
        If Button31.Text = "IDLE" Then
            Button31.BackColor = Color.LightBlue
        ElseIf Button31.Text = "RUNNING" Then
            Button31.BackColor = Color.Yellow
        ElseIf Button31.Text = "PASS" Then
            Button31.BackColor = Color.LightGreen
        ElseIf Button31.Text = "FAIL" Then
            Button31.BackColor = Color.FromArgb(255, 127, 127)
        ElseIf Button31.Text = "ERROR" Then
            Button31.BackColor = Color.LightYellow
        End If
    End Sub
    Private Sub Button32_TextChanged(sender As Object, e As EventArgs) Handles Button32.TextChanged
        If Button32.Text = "IDLE" Then
            Button32.BackColor = Color.LightBlue
        ElseIf Button32.Text = "RUNNING" Then
            Button32.BackColor = Color.Yellow
        ElseIf Button32.Text = "PASS" Then
            Button32.BackColor = Color.LightGreen
        ElseIf Button32.Text = "FAIL" Then
            Button32.BackColor = Color.FromArgb(255, 127, 127)
        ElseIf Button32.Text = "ERROR" Then
            Button32.BackColor = Color.LightYellow
        End If
    End Sub
    Private Sub Button33_TextChanged(sender As Object, e As EventArgs) Handles Button33.TextChanged
        If Button33.Text = "IDLE" Then
            Button33.BackColor = Color.LightBlue
        ElseIf Button33.Text = "RUNNING" Then
            Button33.BackColor = Color.Yellow
        ElseIf Button33.Text = "PASS" Then
            Button33.BackColor = Color.LightGreen
        ElseIf Button33.Text = "FAIL" Then
            Button33.BackColor = Color.FromArgb(255, 127, 127)
        ElseIf Button33.Text = "ERROR" Then
            Button33.BackColor = Color.LightYellow
        End If
    End Sub
    Private Sub Button5_TextChanged(sender As Object, e As EventArgs) Handles Button5.TextChanged
        If Button5.Text = "IDLE" Then
            Button5.BackColor = Color.LightBlue
        ElseIf Button5.Text = "RUNNING" Then
            Button5.BackColor = Color.Yellow
        ElseIf Button5.Text = "PASS" Then
            Button5.BackColor = Color.LightGreen
        ElseIf Button5.Text = "FAIL" Then
            Button5.BackColor = Color.FromArgb(255, 127, 127)
        ElseIf Button5.Text = "ERROR" Then
            Button5.BackColor = Color.LightYellow
        End If
    End Sub
    Private Sub Button6_TextChanged(sender As Object, e As EventArgs) Handles Button6.TextChanged
        If Button6.Text = "IDLE" Then
            Button6.BackColor = Color.LightBlue
        ElseIf Button6.Text = "RUNNING" Then
            Button6.BackColor = Color.Yellow
        ElseIf Button6.Text = "PASS" Then
            Button6.BackColor = Color.LightGreen
        ElseIf Button6.Text = "FAIL" Then
            Button6.BackColor = Color.FromArgb(255, 127, 127)
        ElseIf Button6.Text = "ERROR" Then
            Button6.BackColor = Color.LightYellow
        End If
    End Sub
    Private Sub Button34_TextChanged(sender As Object, e As EventArgs) Handles Button34.TextChanged
        If Button34.Text = "IDLE" Then
            Button34.BackColor = Color.LightBlue
        ElseIf Button34.Text = "RUNNING" Then
            Button34.BackColor = Color.Yellow
        ElseIf Button34.Text = "PASS" Then
            Button34.BackColor = Color.LightGreen
        ElseIf Button34.Text = "FAIL" Then
            Button34.BackColor = Color.FromArgb(255, 127, 127)
        ElseIf Button34.Text = "ERROR" Then
            Button34.BackColor = Color.LightYellow
        End If
    End Sub
    Private Sub Button35_TextChanged(sender As Object, e As EventArgs) Handles Button35.TextChanged
        If Button35.Text = "IDLE" Then
            Button35.BackColor = Color.LightBlue
        ElseIf Button35.Text = "RUNNING" Then
            Button35.BackColor = Color.Yellow
        ElseIf Button35.Text = "PASS" Then
            Button35.BackColor = Color.LightGreen
        ElseIf Button35.Text = "FAIL" Then
            Button35.BackColor = Color.FromArgb(255, 127, 127)
        ElseIf Button35.Text = "ERROR" Then
            Button35.BackColor = Color.LightYellow
        End If
    End Sub
    Private Sub Button36_TextChanged(sender As Object, e As EventArgs) Handles Button36.TextChanged
        If Button36.Text = "IDLE" Then
            Button36.BackColor = Color.LightBlue
        ElseIf Button36.Text = "RUNNING" Then
            Button36.BackColor = Color.Yellow
        ElseIf Button36.Text = "PASS" Then
            Button36.BackColor = Color.LightGreen
        ElseIf Button36.Text = "FAIL" Then
            Button36.BackColor = Color.FromArgb(255, 127, 127)
        ElseIf Button36.Text = "ERROR" Then
            Button36.BackColor = Color.LightYellow
        End If
    End Sub
    Private Sub Button37_TextChanged(sender As Object, e As EventArgs) Handles Button37.TextChanged
        If Button37.Text = "IDLE" Then
            Button37.BackColor = Color.LightBlue
        ElseIf Button37.Text = "RUNNING" Then
            Button37.BackColor = Color.Yellow
        ElseIf Button37.Text = "PASS" Then
            Button37.BackColor = Color.LightGreen
        ElseIf Button37.Text = "FAIL" Then
            Button37.BackColor = Color.FromArgb(255, 127, 127)
        ElseIf Button37.Text = "ERROR" Then
            Button37.BackColor = Color.LightYellow
        End If
    End Sub
    Private Sub Button38_TextChanged(sender As Object, e As EventArgs) Handles Button38.TextChanged
        If Button38.Text = "IDLE" Then
            Button38.BackColor = Color.LightBlue
        ElseIf Button38.Text = "RUNNING" Then
            Button38.BackColor = Color.Yellow
        ElseIf Button38.Text = "PASS" Then
            Button38.BackColor = Color.LightGreen
        ElseIf Button38.Text = "FAIL" Then
            Button38.BackColor = Color.FromArgb(255, 127, 127)
        ElseIf Button38.Text = "ERROR" Then
            Button38.BackColor = Color.LightYellow
        End If
    End Sub
    Private Sub Button39_TextChanged(sender As Object, e As EventArgs) Handles Button39.TextChanged
        If Button39.Text = "IDLE" Then
            Button39.BackColor = Color.LightBlue
        ElseIf Button39.Text = "RUNNING" Then
            Button39.BackColor = Color.Yellow
        ElseIf Button39.Text = "PASS" Then
            Button39.BackColor = Color.LightGreen
        ElseIf Button39.Text = "FAIL" Then
            Button39.BackColor = Color.FromArgb(255, 127, 127)
        ElseIf Button39.Text = "ERROR" Then
            Button39.BackColor = Color.LightYellow
        End If
    End Sub
    Private Sub Button43_Click(sender As Object, e As EventArgs) Handles Button43.Click

    End Sub

    Private Sub Button7_Click(sender As Object, e As EventArgs) Handles Button7.Click
        TextBox40.Text = DMM.MeasureDCVoltage()
        If TextBox40.Text.ToString >= TextBox42.Text.ToString And TextBox40.Text.ToString <= TextBox39.Text Then
            Button18.Text = "PASS"
            ListBox1.Items.Add($"{DateTime.Now.ToString} CAN670_5V Measured : {TextBox40.Text} Expected Value : {TextBox39.Text.ToString}<=X=>{TextBox42.Text.ToString} Result : PASS")
            ListBox1.SelectedIndex = ListBox1.Items.Count - 1
            Button43.Text = "RUNNING"
            If RadioButton1.Checked = True Then
                TreeView1.SelectedNode.Nodes(1).Nodes(1).BackColor = Color.Green
                PictureBox3.Visible = False
                Button7.Visible = False
                TextBox173.Visible = False
                SMC.SMCPowerTestNode2()
            Else
                Exit Sub
            End If
        Else
            Button18.Text = "FAIL"
            ListBox1.Items.Add($"{DateTime.Now.ToString} CAN670_5V Measured : {TextBox40.Text} Expected Value : {TextBox39.Text.ToString}<=X=>{TextBox42.Text.ToString} Result : FAIL")
            ListBox1.SelectedIndex = ListBox1.Items.Count - 1
            Button43.Text = "FAIL"
            If RadioButton1.Checked = True Then
                TreeView1.SelectedNode.Nodes(1).Nodes(1).BackColor = Color.FromArgb(255, 127, 127)
                PictureBox3.Visible = False
                Button7.Visible = False
                TextBox173.Visible = False
                TimerFunction("STOP")
                UUT_RESULT.RadioButton2.Checked = True
                UUT_RESULT.TextBox2.Text = "CAN670_5V"
                UUT_RESULT.ShowDialog()
            Else
                Exit Sub
            End If
        End If
    End Sub

    Private Sub Button13_Click(sender As Object, e As EventArgs) Handles Button13.Click
        TextBox54.Text = DMM.MeasureDCVoltage()
        If TextBox54.Text.ToString >= TextBox53.Text.ToString And TextBox54.Text.ToString <= TextBox55.Text Then
            Button20.Text = "PASS"
            ListBox1.Items.Add($"{DateTime.Now.ToString} ISOLATED_V+12 Measured : {TextBox54.Text} Expected Value : {TextBox55.Text.ToString}<=X=>{TextBox54.Text.ToString} Result : PASS")
            ListBox1.SelectedIndex = ListBox1.Items.Count - 1
            Button43.Text = "RUNNING"
            If RadioButton1.Checked = True Then
                TreeView1.SelectedNode.Nodes(1).Nodes(2).BackColor = Color.Green
                PictureBox3.Visible = False
                Button13.Visible = False
                TextBox173.Visible = False
                SMC.SMCPowerTestNode3()
            Else
                Exit Sub
            End If
        Else
            Button20.Text = "FAIL"
            ListBox1.Items.Add($"{DateTime.Now.ToString} ISOLATED_V+12 Measured : {TextBox54.Text} Expected Value : {TextBox55.Text.ToString}<=X=>{TextBox53.Text.ToString} Result : FAIL")
            ListBox1.SelectedIndex = ListBox1.Items.Count - 1
            Button43.Text = "FAIL"
            If RadioButton1.Checked = True Then
                TreeView1.SelectedNode.Nodes(1).Nodes(2).BackColor = Color.FromArgb(255, 127, 127)
                PictureBox3.Visible = False
                Button13.Visible = False
                TextBox173.Visible = False
                TimerFunction("STOP")
                UUT_RESULT.RadioButton2.Checked = True
                UUT_RESULT.TextBox2.Text = "ISOLATED_V+12"
                UUT_RESULT.ShowDialog()
            Else
                Exit Sub
            End If
        End If
    End Sub

    Private Sub Button14_Click(sender As Object, e As EventArgs) Handles Button14.Click
        TextBox58.Text = DMM.MeasureDCVoltage()
        If TextBox58.Text.ToString >= TextBox57.Text.ToString And TextBox59.Text.ToString <= TextBox55.Text Then
            Button21.Text = "PASS"
            ListBox1.Items.Add($"{DateTime.Now.ToString} SENSOR_V+12 Measured : {TextBox58.Text} Expected Value : {TextBox59.Text.ToString}<=X=>{TextBox57.Text.ToString} Result : PASS")
            ListBox1.SelectedIndex = ListBox1.Items.Count - 1
            Button43.Text = "RUNNING"
            If RadioButton1.Checked = True Then
                TreeView1.SelectedNode.Nodes(1).Nodes(3).BackColor = Color.Green
                PictureBox3.Visible = False
                Button14.Visible = False
                TextBox173.Visible = False
                SMC.SMCPowerTestNode4()
            Else
                Exit Sub
            End If
        Else
            Button21.Text = "FAIL"
            ListBox1.Items.Add($"{DateTime.Now.ToString} SENSOR_V+12 Measured : {TextBox58.Text} Expected Value : {TextBox59.Text.ToString}<=X=>{TextBox57.Text.ToString} Result : FAIL")
            ListBox1.SelectedIndex = ListBox1.Items.Count - 1
            Button43.Text = "FAIL"
            If RadioButton1.Checked = True Then
                TreeView1.SelectedNode.Nodes(1).Nodes(3).BackColor = Color.FromArgb(255, 127, 127)
                PictureBox3.Visible = False
                Button14.Visible = False
                TextBox173.Visible = False
                TimerFunction("STOP")
                UUT_RESULT.RadioButton2.Checked = True
                UUT_RESULT.TextBox2.Text = "ISOLATED_V+12"
                UUT_RESULT.ShowDialog()
            Else
                Exit Sub
            End If
        End If
    End Sub

    Private Sub Button44_Click(sender As Object, e As EventArgs) Handles Button44.Click
        TextBox62.Text = DMM.MeasureDCVoltage()
        If TextBox62.Text.ToString >= TextBox61.Text.ToString And TextBox62.Text.ToString <= TextBox63.Text Then
            Button22.Text = "PASS"
            ListBox1.Items.Add($"{DateTime.Now.ToString} TT620_5V Measured : {TextBox62.Text} Expected Value : {TextBox63.Text.ToString}<=X=>{TextBox61.Text.ToString} Result : PASS")
            ListBox1.SelectedIndex = ListBox1.Items.Count - 1
            Button43.Text = "RUNNING"
            If RadioButton1.Checked = True Then
                TreeView1.SelectedNode.Nodes(1).Nodes(4).BackColor = Color.Green
                PictureBox3.Visible = False
                Button44.Visible = False
                TextBox173.Visible = False
                SMC.SMCPowerTestNode5()
            Else
                Exit Sub
            End If
        Else
            Button22.Text = "FAIL"
            ListBox1.Items.Add($"{DateTime.Now.ToString} TT620_5V Measured : {TextBox62.Text} Expected Value : {TextBox63.Text.ToString}<=X=>{TextBox61.Text.ToString} Result : FAIL")
            ListBox1.SelectedIndex = ListBox1.Items.Count - 1
            Button43.Text = "FAIL"
            If RadioButton1.Checked = True Then
                TreeView1.SelectedNode.Nodes(1).Nodes(4).BackColor = Color.FromArgb(255, 127, 127)
                PictureBox3.Visible = False
                Button44.Visible = False
                TimerFunction("STOP")
                UUT_RESULT.RadioButton2.Checked = True
                UUT_RESULT.TextBox2.Text = "TT620_5V"
                UUT_RESULT.ShowDialog()
            Else
                Exit Sub
            End If
        End If
    End Sub

    Private Sub Button45_Click(sender As Object, e As EventArgs) Handles Button45.Click
        TextBox66.Text = DMM.MeasureDCVoltage()
        If TextBox66.Text.ToString >= TextBox65.Text.ToString And TextBox66.Text.ToString <= TextBox67.Text Then
            Button23.Text = "PASS"
            ListBox1.Items.Add($"{DateTime.Now.ToString} IT661_5V Measured : {TextBox66.Text} Expected Value : {TextBox67.Text.ToString}<=X=>{TextBox65.Text.ToString} Result : PASS")
            ListBox1.SelectedIndex = ListBox1.Items.Count - 1
            Button43.Text = "RUNNING"
            If RadioButton1.Checked = True Then
                TreeView1.SelectedNode.Nodes(1).Nodes(5).BackColor = Color.Green
                PictureBox3.Visible = False
                Button45.Visible = False
                TextBox173.Visible = False
                SMC.SMCPowerTestNode6()
            Else
                Exit Sub
            End If
        Else
            Button23.Text = "FAIL"
            ListBox1.Items.Add($"{DateTime.Now.ToString} IT661_5V Measured : {TextBox66.Text} Expected Value : {TextBox67.Text.ToString}<=X=>{TextBox65.Text.ToString} Result : FAIL")
            ListBox1.SelectedIndex = ListBox1.Items.Count - 1
            Button43.Text = "FAIL"
            If RadioButton1.Checked = True Then
                TreeView1.SelectedNode.Nodes(1).Nodes(5).BackColor = Color.FromArgb(255, 127, 127)
                PictureBox3.Visible = False
                Button45.Visible = False
                TextBox173.Visible = False
                TimerFunction("STOP")
                UUT_RESULT.RadioButton2.Checked = True
                UUT_RESULT.TextBox2.Text = "IT661_5V"
                UUT_RESULT.ShowDialog()
            Else
                Exit Sub
            End If
        End If
    End Sub

    Private Sub Button46_Click(sender As Object, e As EventArgs) Handles Button46.Click
        TextBox69.Text = DMM.MeasureDCVoltage()
        If TextBox69.Text.ToString >= TextBox70.Text.ToString And TextBox69.Text.ToString <= TextBox68.Text Then
            Button24.Text = "PASS"
            ListBox1.Items.Add($"{DateTime.Now.ToString} A0691 Measured : {TextBox69.Text} Expected Value : {TextBox68.Text.ToString}<=X=>{TextBox70.Text.ToString} Result : PASS")
            Button43.Text = "RUNNING"
            If RadioButton1.Checked = True Then
                TreeView1.SelectedNode.Nodes(1).Nodes(6).BackColor = Color.Green
                PictureBox3.Visible = False
                Button46.Visible = False
                TextBox173.Visible = False
                SMC.SMCPowerTestNode7()
            Else
                Exit Sub
            End If
        Else
            Button24.Text = "FAIL"
            ListBox1.Items.Add($"{DateTime.Now.ToString} A0691 Measured : {TextBox69.Text} Expected Value : {TextBox68.Text.ToString}<=X=>{TextBox70.Text.ToString} Result : FAIL")
            Button43.Text = "FAIL"
            If RadioButton1.Checked = True Then
                TreeView1.SelectedNode.Nodes(1).Nodes(6).BackColor = Color.FromArgb(255, 127, 127)
                PictureBox3.Visible = False
                Button46.Visible = False
                TextBox173.Visible = False
                TimerFunction("STOP")
                UUT_RESULT.RadioButton2.Checked = True
                UUT_RESULT.TextBox2.Text = "A0691"
                UUT_RESULT.ShowDialog()
            Else
                Exit Sub
            End If
        End If
    End Sub

    Private Sub Button47_Click(sender As Object, e As EventArgs) Handles Button47.Click
        TextBox74.Text = DMM.MeasureDCVoltage()
        If TextBox74.Text.ToString >= TextBox73.Text.ToString And TextBox74.Text.ToString <= TextBox75.Text Then
            Button25.Text = "PASS"
            ListBox1.Items.Add($"{DateTime.Now.ToString} D0691 measured : {TextBox74.Text} Expected Value : {TextBox75.Text.ToString}<=X=> {TextBox73.Text.ToString} :PASS ")
            Button43.Text = "RUNNING"
            If RadioButton1.Checked = True Then
                TreeView1.SelectedNode.Nodes(1).Nodes(7).BackColor = Color.Green
                PictureBox3.Visible = False
                Button47.Visible = False
                TextBox173.Visible = False
                SMC.SMCSwitchTestNode0()
            Else
                Exit Sub
            End If
        Else
            Button25.Text = "FAIL"
            ListBox1.Items.Add($"{DateTime.Now.ToString} D0691 measured : {TextBox74.Text} Expected Value : {TextBox75.Text.ToString}<=X=> {TextBox73.Text.ToString} : FAIL ")
            Button43.Text = "FAIL"
            If RadioButton1.Checked = True Then
                TreeView1.SelectedNode.Nodes(1).Nodes(7).BackColor = Color.FromArgb(255, 127, 127)
                PictureBox3.Visible = False
                Button46.Visible = False
                TextBox173.Visible = False
                TimerFunction("STOP")
                UUT_RESULT.RadioButton2.Checked = True
                UUT_RESULT.TextBox2.Text = "D0691"
                UUT_RESULT.ShowDialog()
            Else
                Exit Sub
            End If

        End If
    End Sub

    Private Sub ConfrigurationsToolStripMenuItem_Click(sender As Object, e As EventArgs) Handles ConfrigurationsToolStripMenuItem.Click
        CONFIG.ShowDialog()
    End Sub

    Private Sub Button48_Click(sender As Object, e As EventArgs) Handles Button48.Click
        PictureBox3.Visible = False
        TextBox173.Visible = False
        Button48.Visible = False
        SMC.SMCSwitchTestNode1()
    End Sub

    Private Sub Button49_Click(sender As Object, e As EventArgs) Handles Button49.Click
        PictureBox3.Visible = False
        TextBox173.Visible = False
        Button49.Visible = False
        SMC.SMCSwitchTestNode2()
    End Sub

    Private Sub Button50_Click(sender As Object, e As EventArgs) Handles Button50.Click
        PictureBox3.Visible = False
        TextBox173.Visible = False
        Button50.Visible = False
        SMC.SMCSwitchTestNode3()
    End Sub

    Private Sub Button51_Click(sender As Object, e As EventArgs) Handles Button51.Click
        PictureBox3.Visible = False
        TextBox173.Visible = False
        Button51.Visible = False
        SMC.SMCSwitchTestNode4()
    End Sub

    Private Sub Button53_Click(sender As Object, e As EventArgs) Handles Button53.Click
        PictureBox3.Visible = False
        TextBox173.Visible = False
        Button53.Visible = False
        SMC.SMCSwitchTestNode6()
    End Sub

    Private Sub Button52_Click(sender As Object, e As EventArgs) Handles Button52.Click
        PictureBox3.Visible = False
        TextBox173.Visible = False
        Button52.Visible = False
        SMC.SMCSwitchTestNode5()
    End Sub

    Private Sub PictureBox3_Click(sender As Object, e As EventArgs) Handles PictureBox3.Click

    End Sub

    Private Sub Button54_Click(sender As Object, e As EventArgs) Handles Button54.Click
        PictureBox3.Visible = False
        TextBox173.Visible = False
        Button54.Visible = False
        SMC.SMCSwitchTestNode7()
    End Sub

    Private Sub Button55_Click(sender As Object, e As EventArgs) Handles Button55.Click
        PictureBox3.Visible = False
        TextBox173.Visible = False
        Button55.Visible = False
        SMC.SMCSwitchTestNode8()
    End Sub

    Private Sub Button56_Click(sender As Object, e As EventArgs) Handles Button56.Click
        PictureBox3.Visible = False
        TextBox173.Visible = False
        Button56.Visible = False
        SMC.SMCSwitchTestNode9()
    End Sub

    Private Sub Button57_Click(sender As Object, e As EventArgs) Handles Button57.Click
        PictureBox3.Visible = False
        TextBox173.Visible = False
        Button57.Visible = False
        SMC.SMCSwitchTestNode10()
    End Sub

    Private Sub Button58_Click(sender As Object, e As EventArgs) Handles Button58.Click
        PictureBox3.Visible = False
        TextBox173.Visible = False
        Button58.Visible = False
        SMC.SMCSwitchTestNode11()
    End Sub

    Private Sub Button59_Click(sender As Object, e As EventArgs) Handles Button59.Click
        PictureBox3.Visible = False
        TextBox173.Visible = False
        Button59.Visible = False
        SMC.SMCSwitchTestNode12()
    End Sub

    Private Sub Button60_Click(sender As Object, e As EventArgs) Handles Button60.Click
        PictureBox3.Visible = False
        TextBox173.Visible = False
        Button60.Visible = False
        SMC.SMCSwitchTestNode13()
    End Sub

    Private Sub Button61_Click(sender As Object, e As EventArgs) Handles Button61.Click
        PictureBox3.Visible = False
        TextBox173.Visible = False
        Button61.Visible = False
        SMC.SMCTimeNode0()
    End Sub

    Private Sub Button37_Click(sender As Object, e As EventArgs) Handles Button37.Click

    End Sub

    Private Sub Button62_Click(sender As Object, e As EventArgs) Handles Button62.Click
        PictureBox3.Visible = False
        TextBox173.Visible = False
        Button62.Visible = False
        SMC.SMCStepperTestNode0()
    End Sub

    Private Sub Button63_Click(sender As Object, e As EventArgs) Handles Button63.Click
        PictureBox3.Visible = False
        TextBox173.Visible = False
        Button63.Visible = False
        SMC.SMCShutDownNod0()
    End Sub

    Private Sub TextBox93_TextChanged(sender As Object, e As EventArgs) Handles TextBox93.TextChanged

    End Sub

    Private Sub TabPage8_Click(sender As Object, e As EventArgs) Handles TabPage8.Click

    End Sub

    Private Sub ListBox1_SelectedIndexChanged(sender As Object, e As EventArgs) Handles ListBox1.SelectedIndexChanged

    End Sub

    Private Sub ListBox1_TextChanged(sender As Object, e As EventArgs) Handles ListBox1.TextChanged

    End Sub

    Private Sub CheckBox5_CheckedChanged(sender As Object, e As EventArgs)

    End Sub

    Private Sub TabPage3_Click(sender As Object, e As EventArgs) Handles TabPage3.Click

    End Sub

    Private Sub Button41_Click(sender As Object, e As EventArgs) Handles Button41.Click
        ListBox1.Items.Add($"{DateTime.Now.ToString} Response Recieved from user for connecting usb to programmer @ {TextBox5.Text.ToString}")
        PictureBox3.Visible = False
        Button41.Visible = False
        TextBox173.Visible = False
        SMC.SMCProgrammimgNode3()
    End Sub

    Private Sub Button42_Click(sender As Object, e As EventArgs) Handles Button42.Click
        ListBox1.Items.Add($"{DateTime.Now.ToString} Response Recieved from user for connecting usb to board @ {TextBox5.Text.ToString}")
        PictureBox3.Visible = False
        Button42.Visible = False
        TextBox173.Visible = False
        SMC.SMCProgrammingNode5()
    End Sub

    Private Sub Button10_Click_1(sender As Object, e As EventArgs) Handles Button10.Click
        ListBox1.Items.Add($"{DateTime.Now.ToString} Manual Response Recived for INITIAL POWER From User : True")
        ListBox1.SelectedIndex = ListBox1.Items.Count - 1
        PictureBox3.Visible = False
        Button10.Visible = False
        Button11.Visible = False
        TextBox173.Visible = False
        FUNNCTIONLIB.RestoreApp(CONFIG.TextBox57.Text)
        CapturedImg = FUNNCTIONLIB.TakeSnapofActiveApp()
        TextBox43.Text = "TRUE"
        Button19.Text = "PASS"
        Button43.Text = "RUNNING"
        IMAGE_STR.PictureBox1.Image = CapturedImg
        TreeView1.SelectedNode.Nodes(1).Nodes(0).BackColor = Color.Green
        SMC.SMCPowerTestNode1()
    End Sub
    Private Sub Button11_Click_1(sender As Object, e As EventArgs) Handles Button11.Click
        ListBox1.Items.Add($"{DateTime.Now.ToString} Manual Response Recived for INITIAL POWER From User : False")
        ListBox1.SelectedIndex = ListBox1.Items.Count - 1
        PictureBox3.Visible = False
        Button10.Visible = False
        Button11.Visible = False
        TextBox173.Visible = False
        FUNNCTIONLIB.RestoreApp(CONFIG.TextBox57.Text)
        CapturedImg = FUNNCTIONLIB.TakeSnapofActiveApp()
        TextBox43.Text = "FALSE"
        Button19.Text = "FAIL"
        Button43.Text = "FAIL"
        IMAGE_STR.PictureBox1.Image = CapturedImg
        TreeView1.SelectedNode.Nodes(1).Nodes(0).BackColor = Color.FromArgb(255, 127, 127)
        TimerFunction("STOP")
        UUT_RESULT.RadioButton2.Checked = True
        UUT_RESULT.TextBox2.Text = "INIT_POWER"
        UUT_RESULT.ShowDialog()
    End Sub

    Private Sub Button64_Click(sender As Object, e As EventArgs) Handles Button64.Click
        ListBox1.Items.Add($"{DateTime.Now.ToString} Manual Response Recived for D0691_CMD_ON From User : True")
        ListBox1.SelectedIndex = ListBox1.Items.Count - 1
        PictureBox3.Visible = False
        Button64.Visible = False
        Button65.Visible = False
        TextBox173.Visible = False
        FUNNCTIONLIB.RestoreApp(CONFIG.TextBox57.Text)
        CapturedImg = FUNNCTIONLIB.TakeSnapofActiveApp()
        IMAGE_STR.PictureBox2.Image = CapturedImg
        TextBox162.Text = "ON"
        Button4.Text = "PASS"
        Button43.Text = "RUNNING"
        TreeView1.SelectedNode.Nodes(2).Nodes(0).BackColor = Color.Green
        SMC.SMCSwitchTestNode1()
    End Sub

    Private Sub Button65_Click(sender As Object, e As EventArgs) Handles Button65.Click
        ListBox1.Items.Add($"{DateTime.Now.ToString} Manual Response Recived for D0691_CMD_ON From User : False")
        ListBox1.SelectedIndex = ListBox1.Items.Count - 1
        PictureBox3.Visible = False
        Button64.Visible = False
        Button65.Visible = False
        TextBox173.Visible = False
        FUNNCTIONLIB.RestoreApp(CONFIG.TextBox57.Text)
        CapturedImg = FUNNCTIONLIB.TakeSnapofActiveApp()
        IMAGE_STR.PictureBox2.Image = CapturedImg
        TextBox162.Text = "OFF"
        Button4.Text = "FAIL"
        Button43.Text = "FAIL"
        TreeView1.SelectedNode.Nodes(2).Nodes(0).BackColor = Color.FromArgb(255, 127, 127)
        TimerFunction("STOP")
        UUT_RESULT.RadioButton2.Checked = True
        UUT_RESULT.TextBox2.Text = "D0691_ON"
        UUT_RESULT.ShowDialog()
    End Sub

    Private Sub Button66_Click(sender As Object, e As EventArgs) Handles Button66.Click
        ListBox1.Items.Add($"{DateTime.Now.ToString} Manual Response Recived for D0661_CMD_ON From User : True")
        ListBox1.SelectedIndex = ListBox1.Items.Count - 1
        PictureBox3.Visible = False
        Button66.Visible = False
        Button67.Visible = False
        TextBox173.Visible = False
        FUNNCTIONLIB.RestoreApp(CONFIG.TextBox57.Text)
        CapturedImg = FUNNCTIONLIB.TakeSnapofActiveApp()
        IMAGE_STR.PictureBox3.Image = CapturedImg
        TextBox82.Text = "ON"
        Button26.Text = "PASS"
        Button43.Text = "RUNNING"
        TreeView1.SelectedNode.Nodes(2).Nodes(1).BackColor = Color.Green
        SMC.SMCSwitchTestNode2()
    End Sub

    Private Sub Button67_Click(sender As Object, e As EventArgs) Handles Button67.Click
        ListBox1.Items.Add($"{DateTime.Now.ToString} Manual Response Recived for D0661_CMD_ON From User : False")
        ListBox1.SelectedIndex = ListBox1.Items.Count - 1
        PictureBox3.Visible = False
        Button66.Visible = False
        Button67.Visible = False
        TextBox173.Visible = False
        FUNNCTIONLIB.RestoreApp(CONFIG.TextBox57.Text)
        CapturedImg = FUNNCTIONLIB.TakeSnapofActiveApp()
        IMAGE_STR.PictureBox3.Image = CapturedImg
        TextBox82.Text = "OFF"
        Button26.Text = "FAIL"
        Button43.Text = "FAIL"
        TreeView1.SelectedNode.Nodes(2).Nodes(1).BackColor = Color.FromArgb(255, 127, 127)
        TimerFunction("STOP")
        UUT_RESULT.RadioButton2.Checked = True
        UUT_RESULT.TextBox2.Text = "D0661_ON"
        UUT_RESULT.ShowDialog()
    End Sub

    Private Sub Button68_Click(sender As Object, e As EventArgs) Handles Button68.Click
        ListBox1.Items.Add($"{DateTime.Now.ToString} Manual Response Recived for D0661_CMD_OFF From User : True")
        ListBox1.SelectedIndex = ListBox1.Items.Count - 1
        PictureBox3.Visible = False
        Button68.Visible = False
        Button69.Visible = False
        TextBox173.Visible = False
        FUNNCTIONLIB.RestoreApp(CONFIG.TextBox57.Text)
        CapturedImg = FUNNCTIONLIB.TakeSnapofActiveApp()
        IMAGE_STR.PictureBox4.Image = CapturedImg
        TextBox86.Text = "OFF"
        Button27.Text = "PASS"
        Button43.Text = "RUNNING"
        TreeView1.SelectedNode.Nodes(2).Nodes(2).BackColor = Color.Green
        SMC.SMCSwitchTestNode3()
    End Sub

    Private Sub Button69_Click(sender As Object, e As EventArgs) Handles Button69.Click
        ListBox1.Items.Add($"{DateTime.Now.ToString} Manual Response Recived for D0661_CMD_OFF From User : False")
        ListBox1.SelectedIndex = ListBox1.Items.Count - 1
        PictureBox3.Visible = False
        Button68.Visible = False
        Button69.Visible = False
        TextBox173.Visible = False
        FUNNCTIONLIB.RestoreApp(CONFIG.TextBox57.Text)
        CapturedImg = FUNNCTIONLIB.TakeSnapofActiveApp()
        IMAGE_STR.PictureBox4.Image = CapturedImg
        TextBox86.Text = "ON"
        Button27.Text = "FAIL"
        Button43.Text = "FAIL"
        TreeView1.SelectedNode.Nodes(2).Nodes(1).BackColor = Color.FromArgb(255, 127, 127)
        TimerFunction("STOP")
        UUT_RESULT.RadioButton2.Checked = True
        UUT_RESULT.TextBox2.Text = "D0661_OFF"
        UUT_RESULT.ShowDialog()
    End Sub

    Private Sub Button70_Click(sender As Object, e As EventArgs) Handles Button70.Click
        ListBox1.Items.Add($"{DateTime.Now.ToString} Manual Response Recived for AUXEN_CMD_ON From User : True")
        ListBox1.SelectedIndex = ListBox1.Items.Count - 1
        PictureBox3.Visible = False
        Button70.Visible = False
        Button71.Visible = False
        TextBox173.Visible = False
        FUNNCTIONLIB.RestoreApp(CONFIG.TextBox57.Text)
        CapturedImg = FUNNCTIONLIB.TakeSnapofActiveApp()
        IMAGE_STR.PictureBox5.Image = CapturedImg
        TextBox90.Text = "ON"
        Button28.Text = "PASS"
        Button43.Text = "RUNNING"
        TreeView1.SelectedNode.Nodes(2).Nodes(3).BackColor = Color.Green
        SMC.SMCSwitchTestNode4()
    End Sub

    Private Sub Button71_Click(sender As Object, e As EventArgs) Handles Button71.Click
        ListBox1.Items.Add($"{DateTime.Now.ToString} Manual Response Recived for AUXEN_CMD_ON From User : False")
        ListBox1.SelectedIndex = ListBox1.Items.Count - 1
        PictureBox3.Visible = False
        Button70.Visible = False
        Button71.Visible = False
        TextBox173.Visible = False
        FUNNCTIONLIB.RestoreApp(CONFIG.TextBox57.Text)
        CapturedImg = FUNNCTIONLIB.TakeSnapofActiveApp()
        IMAGE_STR.PictureBox5.Image = CapturedImg
        TextBox90.Text = "OFF"
        Button28.Text = "FAIL"
        Button43.Text = "FAIL"
        TreeView1.SelectedNode.Nodes(2).Nodes(3).BackColor = Color.FromArgb(255, 127, 127)
        TimerFunction("STOP")
        UUT_RESULT.RadioButton2.Checked = True
        UUT_RESULT.TextBox2.Text = "AUXENCMD_ON"
        UUT_RESULT.ShowDialog()
    End Sub

    Private Sub Button72_Click(sender As Object, e As EventArgs) Handles Button72.Click
        ListBox1.Items.Add($"{DateTime.Now.ToString} Manual Response Recived for D0100_CMD_ON From User : True")
        ListBox1.SelectedIndex = ListBox1.Items.Count - 1
        PictureBox3.Visible = False
        Button72.Visible = False
        Button73.Visible = False
        TextBox173.Visible = False
        FUNNCTIONLIB.RestoreApp(CONFIG.TextBox57.Text)
        CapturedImg = FUNNCTIONLIB.TakeSnapofActiveApp()
        IMAGE_STR.PictureBox6.Image = CapturedImg
        TextBox95.Text = "ON"
        Button29.Text = "PASS"
        Button43.Text = "RUNNING"
        TreeView1.SelectedNode.Nodes(2).Nodes(4).BackColor = Color.Green
        SMC.SMCSwitchTestNode5()
    End Sub

    Private Sub Button73_Click(sender As Object, e As EventArgs) Handles Button73.Click
        ListBox1.Items.Add($"{DateTime.Now.ToString} Manual Response Recived for D0100_CMD_ON From User : False")
        ListBox1.SelectedIndex = ListBox1.Items.Count - 1
        PictureBox3.Visible = False
        Button72.Visible = False
        Button73.Visible = False
        TextBox173.Visible = False
        FUNNCTIONLIB.RestoreApp(CONFIG.TextBox57.Text)
        CapturedImg = FUNNCTIONLIB.TakeSnapofActiveApp()
        IMAGE_STR.PictureBox6.Image = CapturedImg
        TextBox95.Text = "OFF"
        Button29.Text = "FAIL"
        Button43.Text = "FAIL"
        TreeView1.SelectedNode.Nodes(2).Nodes(4).BackColor = Color.FromArgb(255, 127, 127)
        TimerFunction("STOP")
        UUT_RESULT.RadioButton2.Checked = True
        UUT_RESULT.TextBox2.Text = "D0100_ON"
        UUT_RESULT.ShowDialog()
    End Sub

    Private Sub Button74_Click(sender As Object, e As EventArgs) Handles Button74.Click
        ListBox1.Items.Add($"{DateTime.Now.ToString} Manual Response Recived for D0100_CMD_OFF From User : True")
        ListBox1.SelectedIndex = ListBox1.Items.Count - 1
        PictureBox3.Visible = False
        Button74.Visible = False
        Button75.Visible = False
        TextBox173.Visible = False
        FUNNCTIONLIB.RestoreApp(CONFIG.TextBox57.Text)
        CapturedImg = FUNNCTIONLIB.TakeSnapofActiveApp()
        IMAGE_STR.PictureBox7.Image = CapturedImg
        TextBox99.Text = "OFF"
        Button30.Text = "PASS"
        Button43.Text = "RUNNING"
        TreeView1.SelectedNode.Nodes(2).Nodes(5).BackColor = Color.Green
        SMC.SMCSwitchTestNode6()
    End Sub

    Private Sub Button75_Click(sender As Object, e As EventArgs) Handles Button75.Click
        ListBox1.Items.Add($"{DateTime.Now.ToString} Manual Response Recived for D0100_CMD_OFF From User : False")
        ListBox1.SelectedIndex = ListBox1.Items.Count - 1
        PictureBox3.Visible = False
        Button74.Visible = False
        Button75.Visible = False
        TextBox173.Visible = False
        FUNNCTIONLIB.RestoreApp(CONFIG.TextBox57.Text)
        CapturedImg = FUNNCTIONLIB.TakeSnapofActiveApp()
        IMAGE_STR.PictureBox7.Image = CapturedImg
        TextBox99.Text = "ON"
        Button30.Text = "FAIL"
        Button43.Text = "FAIL"
        TreeView1.SelectedNode.Nodes(2).Nodes(5).BackColor = Color.FromArgb(255, 127, 127)
        TimerFunction("STOP")
        UUT_RESULT.RadioButton2.Checked = True
        UUT_RESULT.TextBox2.Text = "D0100_OFF"
        UUT_RESULT.ShowDialog()
    End Sub

    Private Sub Button76_Click(sender As Object, e As EventArgs) Handles Button76.Click
        ListBox1.Items.Add($"{DateTime.Now.ToString} Manual Response Recived for D0101_CMD_ON From User : True")
        ListBox1.SelectedIndex = ListBox1.Items.Count - 1
        PictureBox3.Visible = False
        Button76.Visible = False
        Button77.Visible = False
        TextBox173.Visible = False
        FUNNCTIONLIB.RestoreApp(CONFIG.TextBox57.Text)
        CapturedImg = FUNNCTIONLIB.TakeSnapofActiveApp()
        IMAGE_STR.PictureBox8.Image = CapturedImg
        TextBox103.Text = "ON"
        Button31.Text = "PASS"
        Button43.Text = "RUNNING"
        TreeView1.SelectedNode.Nodes(2).Nodes(6).BackColor = Color.Green
        SMC.SMCSwitchTestNode7()
    End Sub

    Private Sub Button77_Click(sender As Object, e As EventArgs) Handles Button77.Click
        ListBox1.Items.Add($"{DateTime.Now.ToString} Manual Response Recived for D0101_CMD_ON From User : False")
        ListBox1.SelectedIndex = ListBox1.Items.Count - 1
        PictureBox3.Visible = False
        Button76.Visible = False
        Button77.Visible = False
        TextBox173.Visible = False
        FUNNCTIONLIB.RestoreApp(CONFIG.TextBox57.Text)
        CapturedImg = FUNNCTIONLIB.TakeSnapofActiveApp()
        IMAGE_STR.PictureBox8.Image = CapturedImg
        TextBox103.Text = "OFF"
        Button31.Text = "FAIL"
        Button43.Text = "FAIL"
        TreeView1.SelectedNode.Nodes(2).Nodes(6).BackColor = Color.FromArgb(255, 127, 127)
        TimerFunction("STOP")
        UUT_RESULT.RadioButton2.Checked = True
        UUT_RESULT.TextBox2.Text = "D0101_ON"
        UUT_RESULT.ShowDialog()
    End Sub

    Private Sub Button78_Click(sender As Object, e As EventArgs) Handles Button78.Click
        ListBox1.Items.Add($"{DateTime.Now.ToString} Manual Response Recived for D0101_CMD_OFF From User : True")
        ListBox1.SelectedIndex = ListBox1.Items.Count - 1
        PictureBox3.Visible = False
        Button78.Visible = False
        Button79.Visible = False
        TextBox173.Visible = False
        FUNNCTIONLIB.RestoreApp(CONFIG.TextBox57.Text)
        CapturedImg = FUNNCTIONLIB.TakeSnapofActiveApp()
        IMAGE_STR.PictureBox9.Image = CapturedImg
        TextBox107.Text = "OFF"
        Button32.Text = "PASS"
        Button43.Text = "RUNNING"
        TreeView1.SelectedNode.Nodes(2).Nodes(7).BackColor = Color.Green
        SMC.SMCSwitchTestNode8()
    End Sub

    Private Sub Button79_Click(sender As Object, e As EventArgs) Handles Button79.Click
        ListBox1.Items.Add($"{DateTime.Now.ToString} Manual Response Recived for D0101_CMD_OFF From User : False")
        ListBox1.SelectedIndex = ListBox1.Items.Count - 1
        PictureBox3.Visible = False
        Button78.Visible = False
        Button79.Visible = False
        TextBox173.Visible = False
        FUNNCTIONLIB.RestoreApp(CONFIG.TextBox57.Text)
        CapturedImg = FUNNCTIONLIB.TakeSnapofActiveApp()
        IMAGE_STR.PictureBox9.Image = CapturedImg
        TextBox107.Text = "ON"
        Button32.Text = "FAIL"
        Button43.Text = "FAIL"
        TreeView1.SelectedNode.Nodes(2).Nodes(7).BackColor = Color.FromArgb(255, 127, 127)
        TimerFunction("STOP")
        UUT_RESULT.RadioButton2.Checked = True
        UUT_RESULT.TextBox2.Text = "D0101_OFF"
        UUT_RESULT.ShowDialog()
    End Sub

    Private Sub Button80_Click(sender As Object, e As EventArgs) Handles Button80.Click
        ListBox1.Items.Add($"{DateTime.Now.ToString} Manual Response Recived for CMD_SWITCH_OFF From User : True")
        ListBox1.SelectedIndex = ListBox1.Items.Count - 1
        PictureBox3.Visible = False
        Button80.Visible = False
        Button81.Visible = False
        TextBox173.Visible = False
        FUNNCTIONLIB.RestoreApp(CONFIG.TextBox57.Text)
        CapturedImg = FUNNCTIONLIB.TakeSnapofActiveApp()
        IMAGE_STR.PictureBox10.Image = CapturedImg
        TextBox111.Text = "OFF"
        Button33.Text = "PASS"
        Button43.Text = "RUNNING"
        TreeView1.SelectedNode.Nodes(2).Nodes(8).BackColor = Color.Green
        SMC.SMCSwitchTestNode9()
    End Sub

    Private Sub Button81_Click(sender As Object, e As EventArgs) Handles Button81.Click
        ListBox1.Items.Add($"{DateTime.Now.ToString} Manual Response Recived for CMD_SWITCH_OFF From User : False")
        ListBox1.SelectedIndex = ListBox1.Items.Count - 1
        PictureBox3.Visible = False
        Button80.Visible = False
        Button81.Visible = False
        TextBox173.Visible = False
        FUNNCTIONLIB.RestoreApp(CONFIG.TextBox57.Text)
        CapturedImg = FUNNCTIONLIB.TakeSnapofActiveApp()
        IMAGE_STR.PictureBox10.Image = CapturedImg
        TextBox111.Text = "ON"
        Button33.Text = "FAIL"
        Button43.Text = "FAIL"
        TreeView1.SelectedNode.Nodes(2).Nodes(8).BackColor = Color.FromArgb(255, 127, 127)
        TimerFunction("STOP")
        UUT_RESULT.RadioButton2.Checked = True
        UUT_RESULT.TextBox2.Text = "CMD_OFF"
        UUT_RESULT.ShowDialog()
    End Sub

    Private Sub Button82_Click(sender As Object, e As EventArgs) Handles Button82.Click
        ListBox1.Items.Add($"{DateTime.Now.ToString} Manual Response Recived for SW2_ON From User : True")
        ListBox1.SelectedIndex = ListBox1.Items.Count - 1
        PictureBox3.Visible = False
        Button82.Visible = False
        Button83.Visible = False
        TextBox173.Visible = False
        FUNNCTIONLIB.RestoreApp(CONFIG.TextBox57.Text)
        CapturedImg = FUNNCTIONLIB.TakeSnapofActiveApp()
        IMAGE_STR.PictureBox11.Image = CapturedImg
        TextBox167.Text = "ON"
        Button5.Text = "PASS"
        Button43.Text = "RUNNING"
        TreeView1.SelectedNode.Nodes(2).Nodes(9).BackColor = Color.Green
        SMC.SMCSwitchTestNode10()
    End Sub

    Private Sub Button83_Click(sender As Object, e As EventArgs) Handles Button83.Click
        ListBox1.Items.Add($"{DateTime.Now.ToString} Manual Response Recived for SW2_ON From User : False")
        ListBox1.SelectedIndex = ListBox1.Items.Count - 1
        PictureBox3.Visible = False
        Button82.Visible = False
        Button83.Visible = False
        TextBox173.Visible = False
        FUNNCTIONLIB.RestoreApp(CONFIG.TextBox57.Text)
        CapturedImg = FUNNCTIONLIB.TakeSnapofActiveApp()
        IMAGE_STR.PictureBox11.Image = CapturedImg
        TextBox167.Text = "OFF"
        Button5.Text = "FAIL"
        Button43.Text = "FAIL"
        TreeView1.SelectedNode.Nodes(2).Nodes(9).BackColor = Color.FromArgb(255, 127, 127)
        TimerFunction("STOP")
        UUT_RESULT.RadioButton2.Checked = True
        UUT_RESULT.TextBox2.Text = "SW2_ON"
        UUT_RESULT.ShowDialog()
    End Sub

    Private Sub Button84_Click(sender As Object, e As EventArgs) Handles Button84.Click
        ListBox1.Items.Add($"{DateTime.Now.ToString} Manual Response Recived for SW3_ON From User : True")
        ListBox1.SelectedIndex = ListBox1.Items.Count - 1
        PictureBox3.Visible = False
        Button84.Visible = False
        Button85.Visible = False
        TextBox173.Visible = False
        FUNNCTIONLIB.RestoreApp(CONFIG.TextBox57.Text)
        CapturedImg = FUNNCTIONLIB.TakeSnapofActiveApp()
        IMAGE_STR.PictureBox12.Image = CapturedImg
        TextBox171.Text = "ON"
        Button6.Text = "PASS"
        Button43.Text = "RUNNING"
        TreeView1.SelectedNode.Nodes(2).Nodes(10).BackColor = Color.Green
        SMC.SMCSwitchTestNode11()
    End Sub

    Private Sub Button85_Click(sender As Object, e As EventArgs) Handles Button85.Click
        ListBox1.Items.Add($"{DateTime.Now.ToString} Manual Response Recived for SW3_ON From User : False")
        ListBox1.SelectedIndex = ListBox1.Items.Count - 1
        PictureBox3.Visible = False
        Button84.Visible = False
        Button85.Visible = False
        TextBox173.Visible = False
        FUNNCTIONLIB.RestoreApp(CONFIG.TextBox57.Text)
        CapturedImg = FUNNCTIONLIB.TakeSnapofActiveApp()
        IMAGE_STR.PictureBox12.Image = CapturedImg
        TextBox171.Text = "OFF"
        Button6.Text = "FAIL"
        Button43.Text = "FAIL"
        TreeView1.SelectedNode.Nodes(2).Nodes(10).BackColor = Color.FromArgb(255, 127, 127)
        TimerFunction("STOP")
        UUT_RESULT.RadioButton2.Checked = True
        UUT_RESULT.TextBox2.Text = "SW3_ON"
        UUT_RESULT.ShowDialog()
    End Sub

    Private Sub Button86_Click(sender As Object, e As EventArgs) Handles Button86.Click
        ListBox1.Items.Add($"{DateTime.Now.ToString} Manual Response Recived for TOOGLE_CMD_ON From User : True")
        ListBox1.SelectedIndex = ListBox1.Items.Count - 1
        PictureBox3.Visible = False
        Button86.Visible = False
        Button87.Visible = False
        TextBox173.Visible = False
        FUNNCTIONLIB.RestoreApp(CONFIG.TextBox57.Text)
        CapturedImg = FUNNCTIONLIB.TakeSnapofActiveApp()
        IMAGE_STR.PictureBox13.Image = CapturedImg
        TextBox115.Text = "ON"
        Button34.Text = "PASS"
        Button43.Text = "RUNNING"
        TreeView1.SelectedNode.Nodes(2).Nodes(11).BackColor = Color.Green
        SMC.SMCSwitchTestNode12()
    End Sub

    Private Sub Button87_Click(sender As Object, e As EventArgs) Handles Button87.Click
        ListBox1.Items.Add($"{DateTime.Now.ToString} Manual Response Recived for TOOGLE_CMD_ON From User : False")
        ListBox1.SelectedIndex = ListBox1.Items.Count - 1
        PictureBox3.Visible = False
        Button86.Visible = False
        Button87.Visible = False
        TextBox173.Visible = False
        FUNNCTIONLIB.RestoreApp(CONFIG.TextBox57.Text)
        CapturedImg = FUNNCTIONLIB.TakeSnapofActiveApp()
        IMAGE_STR.PictureBox13.Image = CapturedImg
        TextBox115.Text = "OFF"
        Button34.Text = "FAIL"
        Button43.Text = "FAIL"
        TreeView1.SelectedNode.Nodes(2).Nodes(11).BackColor = Color.FromArgb(255, 127, 127)
        TimerFunction("STOP")
        UUT_RESULT.RadioButton2.Checked = True
        UUT_RESULT.TextBox2.Text = "TOOGLECMD_ON"
        UUT_RESULT.ShowDialog()
    End Sub

    Private Sub Button88_Click(sender As Object, e As EventArgs) Handles Button88.Click
        ListBox1.Items.Add($"{DateTime.Now.ToString} Manual Response Recived for CAN1_SILENCE_CMD_ON From User : True")
        ListBox1.SelectedIndex = ListBox1.Items.Count - 1
        PictureBox3.Visible = False
        Button88.Visible = False
        Button89.Visible = False
        TextBox173.Visible = False
        FUNNCTIONLIB.RestoreApp(CONFIG.TextBox57.Text)
        CapturedImg = FUNNCTIONLIB.TakeSnapofActiveApp()
        IMAGE_STR.PictureBox14.Image = CapturedImg
        TextBox119.Text = "ON"
        Button35.Text = "PASS"
        Button43.Text = "RUNNING"
        TreeView1.SelectedNode.Nodes(2).Nodes(12).BackColor = Color.Green
        SMC.SMCSwitchTestNode13()
    End Sub

    Private Sub Button89_Click(sender As Object, e As EventArgs) Handles Button89.Click
        ListBox1.Items.Add($"{DateTime.Now.ToString} Manual Response Recived for CAN1_SILENCE_CMD_ON From User : False")
        ListBox1.SelectedIndex = ListBox1.Items.Count - 1
        PictureBox3.Visible = False
        Button88.Visible = False
        Button89.Visible = False
        TextBox173.Visible = False
        FUNNCTIONLIB.RestoreApp(CONFIG.TextBox57.Text)
        CapturedImg = FUNNCTIONLIB.TakeSnapofActiveApp()
        IMAGE_STR.PictureBox14.Image = CapturedImg
        TextBox119.Text = "OFF"
        Button35.Text = "FAIL"
        Button43.Text = "FAIL"
        TreeView1.SelectedNode.Nodes(2).Nodes(12).BackColor = Color.FromArgb(255, 127, 127)
        TimerFunction("STOP")
        UUT_RESULT.RadioButton2.Checked = True
        UUT_RESULT.TextBox2.Text = "CAN1SILENCE_ON"
        UUT_RESULT.ShowDialog()
    End Sub

    Private Sub Button90_Click(sender As Object, e As EventArgs) Handles Button90.Click
        ListBox1.Items.Add($"{DateTime.Now.ToString} Manual Response Recived for CAN2_SILENCE_CMD_ON From User : True")
        ListBox1.SelectedIndex = ListBox1.Items.Count - 1
        PictureBox3.Visible = False
        Button90.Visible = False
        Button91.Visible = False
        TextBox173.Visible = False
        FUNNCTIONLIB.RestoreApp(CONFIG.TextBox57.Text)
        CapturedImg = FUNNCTIONLIB.TakeSnapofActiveApp()
        IMAGE_STR.PictureBox15.Image = CapturedImg
        TextBox123.Text = "ON"
        Button36.Text = "PASS"
        Button43.Text = "RUNNING"
        TreeView1.SelectedNode.Nodes(2).Nodes(13).BackColor = Color.Green
        SMC.SMCTimeNode0()
    End Sub

    Private Sub Button91_Click(sender As Object, e As EventArgs) Handles Button91.Click
        ListBox1.Items.Add($"{DateTime.Now.ToString} Manual Response Recived for CAN2_SILENCE_CMD_ON From User : False")
        ListBox1.SelectedIndex = ListBox1.Items.Count - 1
        PictureBox3.Visible = False
        Button90.Visible = False
        Button91.Visible = False
        TextBox173.Visible = False
        FUNNCTIONLIB.RestoreApp(CONFIG.TextBox57.Text)
        CapturedImg = FUNNCTIONLIB.TakeSnapofActiveApp()
        IMAGE_STR.PictureBox15.Image = CapturedImg
        TextBox123.Text = "OFF"
        Button36.Text = "FAIL"
        Button43.Text = "FAIL"
        TreeView1.SelectedNode.Nodes(2).Nodes(13).BackColor = Color.FromArgb(255, 127, 127)
        TimerFunction("STOP")
        UUT_RESULT.RadioButton2.Checked = True
        UUT_RESULT.TextBox2.Text = "CAN2SILENCE_ON"
        UUT_RESULT.ShowDialog()
    End Sub

    Private Sub Button92_Click(sender As Object, e As EventArgs) Handles Button92.Click
        ListBox1.Items.Add($"{DateTime.Now.ToString} Manual Response Recived for TIME_TEST From User : True")
        ListBox1.SelectedIndex = ListBox1.Items.Count - 1
        PictureBox3.Visible = False
        Button92.Visible = False
        Button93.Visible = False
        TextBox173.Visible = False
        FUNNCTIONLIB.RestoreApp(CONFIG.TextBox57.Text)
        CapturedImg = FUNNCTIONLIB.TakeSnapofActiveApp()
        IMAGE_STR.PictureBox16.Image = CapturedImg
        SMC_TIMETEST.PictureBox1.Image = CapturedImg
        ini.Load("C:\PLUG_POWER_FCT\PLUG_POWER_FCT\Project Files\Confrigurations.ini")
        ListBox2.Items.Add(ini.Sections("SMC_TIMETEST").Keys("Coordinate1").Value)
        SMC_TIMETEST.ShowDialog()
        Button37.Text = "PASS"
        Button43.Text = "RUNNING"
        TreeView1.SelectedNode.Nodes(3).BackColor = Color.Green
        SMC.SMCStepperTestNode0()
    End Sub

    Private Sub Button93_Click(sender As Object, e As EventArgs) Handles Button93.Click
        ListBox1.Items.Add($"{DateTime.Now.ToString} Manual Response Recived for TIME_TEST From User : False")
        ListBox1.SelectedIndex = ListBox1.Items.Count - 1
        PictureBox3.Visible = False
        Button92.Visible = False
        Button93.Visible = False
        TextBox173.Visible = False
        FUNNCTIONLIB.RestoreApp(CONFIG.TextBox57.Text)
        CapturedImg = FUNNCTIONLIB.TakeSnapofActiveApp()
        IMAGE_STR.PictureBox16.Image = CapturedImg
        SMC_TIMETEST.PictureBox1.Image = CapturedImg
        ini.Load("C:\PLUG_POWER_FCT\PLUG_POWER_FCT\Project Files\Confrigurations.ini")
        ListBox2.Items.Add(ini.Sections("SMC_TIMETEST").Keys("Coordinate1").Value)
        SMC_TIMETEST.ShowDialog()
        Button37.Text = "FAIL"
        Button43.Text = "FAIL"
        TreeView1.SelectedNode.Nodes(3).BackColor = Color.FromArgb(255, 127, 127)
        TimerFunction("STOP")
        UUT_RESULT.RadioButton2.Checked = True
        UUT_RESULT.TextBox2.Text = "TIME_TEST"
        UUT_RESULT.ShowDialog()
    End Sub

    Private Sub Button94_Click(sender As Object, e As EventArgs) Handles Button94.Click
        ListBox1.Items.Add($"{DateTime.Now.ToString} Manual Response Recived for STEPPER_TEST From User : True")
        ListBox1.SelectedIndex = ListBox1.Items.Count - 1
        PictureBox3.Visible = False
        Button94.Visible = False
        Button95.Visible = False
        TextBox173.Visible = False
        FUNNCTIONLIB.RestoreApp(CONFIG.TextBox57.Text)
        CapturedImg = FUNNCTIONLIB.TakeSnapofActiveApp()
        IMAGE_STR.PictureBox17.Image = CapturedImg
        TextBox135.Text = "ON"
        Button38.Text = "PASS"
        Button43.Text = "RUNNING"
        TreeView1.SelectedNode.Nodes(4).BackColor = Color.Green
        SMC.SMCShutDownNod0()
    End Sub

    Private Sub Button95_Click(sender As Object, e As EventArgs) Handles Button95.Click
        ListBox1.Items.Add($"{DateTime.Now.ToString} Manual Response Recived for STEPPER_TEST From User : False")
        ListBox1.SelectedIndex = ListBox1.Items.Count - 1
        PictureBox3.Visible = False
        Button94.Visible = False
        Button95.Visible = False
        TextBox173.Visible = False
        FUNNCTIONLIB.RestoreApp(CONFIG.TextBox57.Text)
        CapturedImg = FUNNCTIONLIB.TakeSnapofActiveApp()
        IMAGE_STR.PictureBox17.Image = CapturedImg
        TextBox135.Text = "OFF"
        Button38.Text = "FAIL"
        Button43.Text = "FAIL"
        TreeView1.SelectedNode.Nodes(4).BackColor = Color.FromArgb(255, 127, 127)
        TimerFunction("STOP")
        UUT_RESULT.RadioButton2.Checked = True
        UUT_RESULT.TextBox2.Text = "STEPPER_TEST"
        UUT_RESULT.ShowDialog()
    End Sub

    Private Sub Button96_Click(sender As Object, e As EventArgs) Handles Button96.Click
        ListBox1.Items.Add($"{DateTime.Now.ToString} Manual Response Recived for SHUTDOWN_TEST From User : True")
        ListBox1.SelectedIndex = ListBox1.Items.Count - 1
        PictureBox3.Visible = False
        Button96.Visible = False
        Button97.Visible = False
        TextBox173.Visible = False
        FUNNCTIONLIB.RestoreApp(CONFIG.TextBox57.Text)
        CapturedImg = FUNNCTIONLIB.TakeSnapofActiveApp()
        IMAGE_STR.PictureBox18.Image = CapturedImg
        TextBox144.Text = "TRUE"
        Button39.Text = "PASS"
        Button43.Text = "PASS"
        TreeView1.SelectedNode.Nodes(5).BackColor = Color.Green
        TimerFunction("STOP")
        UUT_RESULT.RadioButton1.Checked = True
        UUT_RESULT.TextBox2.Text = "SHUTDOWN_TEST"
        UUT_RESULT.ShowDialog()
    End Sub

    Private Sub Button97_Click(sender As Object, e As EventArgs) Handles Button97.Click
        ListBox1.Items.Add($"{DateTime.Now.ToString} Manual Response Recived for SHUTDOWN_TEST From User : False")
        ListBox1.SelectedIndex = ListBox1.Items.Count - 1
        PictureBox3.Visible = False
        Button96.Visible = False
        Button97.Visible = False
        TextBox173.Visible = False
        FUNNCTIONLIB.RestoreApp(CONFIG.TextBox57.Text)
        CapturedImg = FUNNCTIONLIB.TakeSnapofActiveApp()
        IMAGE_STR.PictureBox18.Image = CapturedImg
        TextBox144.Text = "FALSE"
        Button39.Text = "FAIL"
        Button43.Text = "FAIL"
        TreeView1.SelectedNode.Nodes(5).BackColor = Color.FromArgb(255, 127, 127)
        TimerFunction("STOP")
        UUT_RESULT.RadioButton2.Checked = True
        UUT_RESULT.TextBox2.Text = "SHUTDOWN_TEST"
        UUT_RESULT.ShowDialog()
    End Sub

    Private Sub GroupBox2_Enter(sender As Object, e As EventArgs) Handles GroupBox2.Enter

    End Sub

    Private Sub Button99_Click(sender As Object, e As EventArgs) Handles Button99.Click
        ListBox1.Items.Add($"{DateTime.Now.ToString} Manual Response Recived for BatteryID Comparison From User : True")
        ListBox1.SelectedIndex = ListBox1.Items.Count - 1
        Button98.Visible = False
        Button99.Visible = False
        TextBox160.Visible = False
        PictureBox2.Visible = False
        FUNNCTIONLIB.RestoreApp(CONFIG.TextBox89.Text)
        System.Threading.Thread.Sleep(2000)
        CapturedImg = FUNNCTIONLIB.TakeSnapofActiveApp()
        IMAGE_STR.PictureBox21.Image = CapturedImg
        TreeView1.SelectedNode.Nodes(0).Nodes(2).BackColor = Color.Green
        TextBox18.Text = "TRUE"
        Button12.Text = "PASS"
        Button43.Text = "PASS"
        myProcesses = Process.GetProcessesByName(CONFIG.TextBox27.Text.ToString)
        If myProcesses.Length > 0 Then
            For Each myProcess In myProcesses
                If myProcess IsNot Nothing Then
                    myProcess.Kill()
                End If
            Next
        End If
        TreeView1.SelectedNode.Nodes(0).Nodes(3).BackColor = Color.Green
        UUT_RESULT.RadioButton1.Checked = True
        UUT_RESULT.TextBox2.Text = "BATTERYID_NODE3"
        UUT_RESULT.ShowDialog()
    End Sub

    Private Sub Button98_Click(sender As Object, e As EventArgs) Handles Button98.Click
        ListBox1.Items.Add($"{DateTime.Now.ToString} Manual Response Recived for BatteryID Comparison From User : False")
        ListBox1.SelectedIndex = ListBox1.Items.Count - 1
        Button98.Visible = False
        Button99.Visible = False
        TextBox160.Visible = False
        PictureBox2.Visible = False
        FUNNCTIONLIB.RestoreApp(CONFIG.TextBox89.Text)
        System.Threading.Thread.Sleep(2000)
        CapturedImg = FUNNCTIONLIB.TakeSnapofActiveApp()
        IMAGE_STR.PictureBox21.Image = CapturedImg
        TreeView1.SelectedNode.Nodes(0).Nodes(2).BackColor = Color.Green
        TextBox18.Text = "FALSE"
        Button12.Text = "FAIL"
        Button43.Text = "FAIL"
        myProcesses = Process.GetProcessesByName(CONFIG.TextBox27.Text.ToString)
        If myProcesses.Length > 0 Then
            For Each myProcess In myProcesses
                If myProcess IsNot Nothing Then
                    myProcess.Kill()
                End If
            Next
        End If
        TreeView1.SelectedNode.Nodes(0).Nodes(3).BackColor = Color.Red
        UUT_RESULT.RadioButton2.Checked = True
        UUT_RESULT.TextBox2.Text = "BATTERYID_NODE3"
        UUT_RESULT.ShowDialog()
    End Sub

    Private Sub Button9_Click(sender As Object, e As EventArgs) Handles Button9.Click
        Button9.Visible = False
        PictureBox1.Visible = False
        TextBox159.Visible = False
        SystemID.SystemIDNode1()
    End Sub

    Private Sub Button101_Click(sender As Object, e As EventArgs) Handles Button101.Click
        ListBox1.Items.Add($"{DateTime.Now.ToString} Manual Response Recived for SystemID Comparison From User : True")
        ListBox1.SelectedIndex = ListBox1.Items.Count - 1
        Button100.Visible = False
        Button101.Visible = False
        TextBox159.Visible = False
        PictureBox1.Visible = False
        FUNNCTIONLIB.RestoreApp(CONFIG.TextBox90.Text)
        System.Threading.Thread.Sleep(2000)
        CapturedImg = FUNNCTIONLIB.TakeSnap(True)
        IMAGE_STR.PictureBox22.Image = CapturedImg
        TreeView1.SelectedNode.Nodes(0).Nodes(2).BackColor = Color.Green
        TextBox13.Text = "TRUE"
        Button8.Text = "PASS"
        Button43.Text = "PASS"
        myProcesses = Process.GetProcessesByName(CONFIG.TextBox22.Text.ToString)
        If myProcesses.Length > 0 Then
            For Each myProcess In myProcesses
                If myProcess IsNot Nothing Then
                    myProcess.Kill()
                End If
            Next
        End If
        TreeView1.SelectedNode.Nodes(0).Nodes(3).BackColor = Color.Green
        UUT_RESULT.RadioButton1.Checked = True
        UUT_RESULT.TextBox2.Text = "SYSTEMID_NODE3"
        UUT_RESULT.ShowDialog()
    End Sub

    Private Sub Button100_Click(sender As Object, e As EventArgs) Handles Button100.Click
        ListBox1.Items.Add($"{DateTime.Now.ToString} Manual Response Recived for SystemID Comparison From User : False")
        ListBox1.SelectedIndex = ListBox1.Items.Count - 1
        Button100.Visible = False
        Button101.Visible = False
        TextBox159.Visible = False
        PictureBox1.Visible = False
        FUNNCTIONLIB.RestoreApp(CONFIG.TextBox90.Text)
        System.Threading.Thread.Sleep(2000)
        CapturedImg = FUNNCTIONLIB.TakeSnap(True)
        IMAGE_STR.PictureBox22.Image = CapturedImg
        TreeView1.SelectedNode.Nodes(0).Nodes(2).BackColor = Color.Green
        TextBox13.Text = "FALSE"
        Button8.Text = "FAIL"
        Button43.Text = "FAIL"
        myProcesses = Process.GetProcessesByName(CONFIG.TextBox22.Text.ToString)
        If myProcesses.Length > 0 Then
            For Each myProcess In myProcesses
                If myProcess IsNot Nothing Then
                    myProcess.Kill()
                End If
            Next
        End If
        TreeView1.SelectedNode.Nodes(0).Nodes(3).BackColor = Color.Red
        UUT_RESULT.RadioButton2.Checked = True
        UUT_RESULT.TextBox2.Text = "SYSTEMID_NODE3"
        UUT_RESULT.ShowDialog()
    End Sub

    Private Sub LOGOUTToolStripMenuItem_Click(sender As Object, e As EventArgs) Handles LOGOUTToolStripMenuItem.Click
        Dim iret As Object = MsgBox("DO YOU WANT TO LOG-OUT?", vbQuestion + vbYesNo, "LOG-OUT")
        If iret = vbYes Then
            TextBox1.Clear()
            TextBox2.Clear()
            TextBox153.Clear()
            TextBox6.Clear()
            TextBox5.Clear()
            TextBox4.Clear()
            TextBox3.Clear()
            FUNNCTIONLIB.ProgressReset(True)
            RadioButton1.Checked = False
            RadioButton2.Checked = False
            GroupBox1.Visible = False
            GroupBox2.Visible = False
            GroupBox3.Visible = False
            GroupBox4.Visible = False
            ListBox1.Visible = False
            ListBox2.Visible = False
            LOGINToolStripMenuItem.Enabled = True
            LOGOUTToolStripMenuItem.Enabled = False
            LOGIN.ShowDialog()
        ElseIf iret = vbNo Then
            'donothing
        End If
    End Sub

    Private Sub Button3_EnabledChanged(sender As Object, e As EventArgs) Handles Button3.EnabledChanged
        If Button3.Enabled = True Then
            If TextBox16.Text = "ADMINISTRATOR" Then
                LOGINToolStripMenuItem.Enabled = False
                LOGOUTToolStripMenuItem.Enabled = True
                EXITToolStripMenuItem.Enabled = True
                CREATEUSERToolStripMenuItem.Enabled = True
                REMOVEUSERToolStripMenuItem.Enabled = True
                CHANGEPASSWORDToolStripMenuItem.Enabled = False
                FIRSTPASSYIELDToolStripMenuItem.Enabled = True
                FINALPASSYIELDToolStripMenuItem.Enabled = True
                ConfrigurationsToolStripMenuItem.Enabled = True
                ABOUTSOFTWAREToolStripMenuItem.Enabled = True
                HELPToolStripMenuItem.Enabled = True
                SETUPToolStripMenuItem.Enabled = True
                UPDATEDESIGNATIONToolStripMenuItem.Enabled = True
                ENABLEMANUALMODEToolStripMenuItem.Enabled = True
            ElseIf TextBox16.Text = "OPERATOR" Then
                LOGINToolStripMenuItem.Enabled = False
                LOGOUTToolStripMenuItem.Enabled = True
                EXITToolStripMenuItem.Enabled = True
                CREATEUSERToolStripMenuItem.Enabled = False
                REMOVEUSERToolStripMenuItem.Enabled = False
                CHANGEPASSWORDToolStripMenuItem.Enabled = True
                FIRSTPASSYIELDToolStripMenuItem.Enabled = True
                FINALPASSYIELDToolStripMenuItem.Enabled = True
                ConfrigurationsToolStripMenuItem.Enabled = False
                ABOUTSOFTWAREToolStripMenuItem.Enabled = True
                HELPToolStripMenuItem.Enabled = True
                SETUPToolStripMenuItem.Enabled = False
                UPDATEDESIGNATIONToolStripMenuItem.Enabled = False
                ENABLEMANUALMODEToolStripMenuItem.Enabled = False
            ElseIf TextBox16.Text = "TEST_DEPT" Then
                LOGINToolStripMenuItem.Enabled = False
                LOGOUTToolStripMenuItem.Enabled = True
                EXITToolStripMenuItem.Enabled = True
                CREATEUSERToolStripMenuItem.Enabled = True
                REMOVEUSERToolStripMenuItem.Enabled = True
                CHANGEPASSWORDToolStripMenuItem.Enabled = True
                FIRSTPASSYIELDToolStripMenuItem.Enabled = True
                FINALPASSYIELDToolStripMenuItem.Enabled = True
                ConfrigurationsToolStripMenuItem.Enabled = False
                ABOUTSOFTWAREToolStripMenuItem.Enabled = True
                HELPToolStripMenuItem.Enabled = True
                SETUPToolStripMenuItem.Enabled = True
                UPDATEDESIGNATIONToolStripMenuItem.Enabled = True
                ENABLEMANUALMODEToolStripMenuItem.Enabled = False
            ElseIf TextBox16.Text = "DEBUG_DEPT" Then
                LOGINToolStripMenuItem.Enabled = False
                LOGOUTToolStripMenuItem.Enabled = True
                EXITToolStripMenuItem.Enabled = True
                CREATEUSERToolStripMenuItem.Enabled = False
                REMOVEUSERToolStripMenuItem.Enabled = False
                CHANGEPASSWORDToolStripMenuItem.Enabled = True
                FIRSTPASSYIELDToolStripMenuItem.Enabled = False
                FINALPASSYIELDToolStripMenuItem.Enabled = False
                ConfrigurationsToolStripMenuItem.Enabled = False
                ABOUTSOFTWAREToolStripMenuItem.Enabled = True
                HELPToolStripMenuItem.Enabled = True
                SETUPToolStripMenuItem.Enabled = False
                UPDATEDESIGNATIONToolStripMenuItem.Enabled = False
                ENABLEMANUALMODEToolStripMenuItem.Enabled = True
            End If
        ElseIf Button3.Enabled = False Then
            LOGINToolStripMenuItem.Enabled = False
            LOGOUTToolStripMenuItem.Enabled = False
            EXITToolStripMenuItem.Enabled = False
            CREATEUSERToolStripMenuItem.Enabled = False
            REMOVEUSERToolStripMenuItem.Enabled = False
            CHANGEPASSWORDToolStripMenuItem.Enabled = False
            FIRSTPASSYIELDToolStripMenuItem.Enabled = False
            FINALPASSYIELDToolStripMenuItem.Enabled = False
            ConfrigurationsToolStripMenuItem.Enabled = False
            ABOUTSOFTWAREToolStripMenuItem.Enabled = False
            HELPToolStripMenuItem.Enabled = False
            SETUPToolStripMenuItem.Enabled = False
            UPDATEDESIGNATIONToolStripMenuItem.Enabled = False
            ENABLEMANUALMODEToolStripMenuItem.Enabled = False
        End If
    End Sub

    Private Sub Button1_Click(sender As Object, e As EventArgs) Handles Button1.Click
        ListBox1.Items.Add($"{DateTime.Now.ToString} Test Sequence Aborted..")
        ListBox1.SelectedIndex = ListBox1.Items.Count - 1
        If TextBox1.Text = "SYSTEM ID" Then
            SystemID.ClearSequence()
            Button3.Enabled = True
            Button1.Enabled = False
        ElseIf TextBox1.Text = "BATTERY ID" Then
            BatteryID.ClearSequence()
            Button3.Enabled = True
            Button1.Enabled = False
        ElseIf TextBox1.Text = "SMC" Then
            SMC.ClearSequence()
            Button3.Enabled = True
            Button1.Enabled = False
        End If
    End Sub
End Class
