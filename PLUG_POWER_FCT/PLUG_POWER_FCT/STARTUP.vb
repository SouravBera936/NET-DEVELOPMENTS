Public Class STARTUP
    Dim productversion = Application.ProductVersion.ToString
    Dim Products = {"SYSTEM ID", "BATTERY ID", "SMC"}
    Dim Test_Methodes = {"PRE-COATING", "POST-COATING"}
    Dim Test_mode = {"AUTO", "MANUAL"}
    Dim SL As String
    Private Sub STARTUP_Load(sender As Object, e As EventArgs) Handles MyBase.Load
        MAIN.Button3.Enabled = False
        MAIN.Button1.Enabled = True
        If MAIN.CheckBox1.Checked = True Then
            MAIN.TimerFunction("RESET")
            MAIN.ListBox1.Items.Clear()
            fade_in()
            MAIN.CheckBox2.Checked = False 'RESET THE GOLDEN IDENTIFIER
            Label3.Text = productversion
            TextBox2.Text = MAIN.TextBox3.Text
            TextBox2.Enabled = False
            MAIN.CheckBox4.Checked = False
            MAIN.CheckBox2.Checked = False
            MAIN.CheckBox3.Checked = False
            TextBox4.Clear()
            TextBox4.Enabled = False
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
            MAIN.TextBox1.Clear()
            MAIN.TextBox2.Clear()
            MAIN.TextBox153.Clear()
            MAIN.TextBox4.Clear()
            MAIN.TextBox6.Clear()
        Else
            FUNNCTIONLIB.ProgressReset(True)
            MAIN.TimerFunction("RESET")
            MAIN.ListBox1.Items.Clear()
            fade_in()
            MAIN.CheckBox2.Checked = False 'RESET THE GOLDEN IDENTIFIER
            Label3.Text = productversion
            TextBox2.Text = MAIN.TextBox3.Text
            TextBox2.Enabled = False
            ComboBox1.Text = ""
            ComboBox1.Items.Clear()
            ComboBox1.Items.AddRange(Products)
            ComboBox2.Text = ""
            ComboBox2.Items.Clear()
            ComboBox2.Items.AddRange(Test_Methodes)
            ComboBox3.Text = ""
            ComboBox3.Items.Clear()
            ComboBox3.Items.AddRange(Test_mode)
            TextBox3.Clear()
            TextBox1.Clear()
            TextBox4.Clear()
            If MAIN.TextBox16.Text = "OPERATOR" Then
                ComboBox3.Enabled = False
                ComboBox3.SelectedIndex = 0
            Else
                ComboBox3.Enabled = True
            End If
            MAIN.CheckBox4.Checked = False
            MAIN.CheckBox2.Checked = False
            MAIN.CheckBox3.Checked = False
            TextBox4.Clear()
            TextBox4.Enabled = False
            MAIN.TextBox1.Clear()
            MAIN.TextBox2.Clear()
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
        End If
    End Sub
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
    Public Sub fade_out()
        For FadeOut = 90 To 10 Step -20
            Me.Opacity = FadeOut / 100
            Me.Refresh()
            Threading.Thread.Sleep(50)
        Next
    End Sub
    Public Sub fade_in()
        For FadeIn = 0.0 To 1.1 Step 0.2
            Me.Opacity = FadeIn
            Me.Refresh()
            Threading.Thread.Sleep(100)
        Next
    End Sub

    Private Sub Button1_Click(sender As Object, e As EventArgs) Handles Button1.Click
        MAIN.Button3.Enabled = True
        MAIN.Button1.Enabled = False
        MAIN.TextBox1.Clear()
        MAIN.TextBox2.Clear()
        MAIN.TextBox153.Clear()
        MAIN.TextBox4.Clear()
        MAIN.TextBox6.Clear()
        MAIN.CheckBox1.Checked = False
        fade_out()
        Me.Close()
    End Sub

    Private Sub Button2_Click(sender As Object, e As EventArgs) Handles Button2.Click
        If ComboBox1.SelectedItem = Nothing Then
            Dim iret As Object = MsgBox("NO MODELS SELECTED TO RUN!", vbOKOnly + vbCritical, "TEST EXECUTION error")
            If iret = vbOK Then
                ComboBox1.Focus()
            End If
        ElseIf ComboBox2.SelectedItem = Nothing Then
            Dim iret1 As Object = MsgBox("TEST MODE NOT SELECTED TO RUN!", vbOKOnly + vbCritical, "TEST EXECUTION error")
            If iret1 = vbOK Then
                ComboBox2.Focus()
            End If
        ElseIf ComboBox3.SelectedItem = Nothing Then
            Dim iret2 As Object = MsgBox("TEST MODE NOT SELECTED TO RUN!", vbOKOnly + vbCritical, "TEST EXECUTION error")
            If iret2 = vbOK Then
                ComboBox3.Focus()
            End If
        ElseIf TextBox1.Text = "" Then
            Dim iret3 As Object = MsgBox("NO SERIAL NUMBER FOUND TO RUN!", vbOKOnly + vbCritical, "TEST EXECUTION error")
            If iret3 = vbOK Then
                TextBox1.Focus()
            End If
        Else
            MAIN.ListBox1.Items.Clear()
            TestInitialization()
        End If
    End Sub
    Public Sub TestInitialization()
        If ComboBox1.SelectedItem = "SYSTEM ID" And ComboBox2.SelectedItem = "PRE-COATING" Then
            If ComboBox3.SelectedItem = "AUTO" Then
                MAIN.RadioButton1.Checked = True
                MAIN.TreeView1.Enabled = False
                MAIN.ListBox1.Items.Add($"{DateTime.Now.ToString}TEST SEQUENCE : SYSTEM ID > PRE-COATING > AUTO")
                MAIN.ListBox1.SelectedIndex = MAIN.ListBox1.Items.Count - 1
                If SETUP.CheckBox1.Checked = True Then
                    SL = FUNNCTIONLIB.SLtoUpper(TextBox1.Text, True)
                    If TextBox4.Text = "PRODUCTION" Then
                        MAIN.ListBox1.Items.Add($"{DateTime.Now.ToString} DUT ID : PRODUCTION")
                        MAIN.ListBox1.SelectedIndex = MAIN.ListBox1.Items.Count - 1
                        If FUNNCTIONLIB.Initialize("SYSTEM_ID", Me.TextBox1.Text, CONFIG.TextBox5.Text, ComboBox2.SelectedItem.ToString, CONFIG.TextBox1.Text, CONFIG.TextBox2.Text, CONFIG.TextBox3.Text, True) = True Then
                            MAIN.TextBox1.Text = Me.ComboBox1.SelectedItem
                            MAIN.TextBox2.Text = Me.ComboBox2.SelectedItem
                            MAIN.TextBox153.Text = Me.TextBox3.Text
                            MAIN.TextBox6.Text = Me.TextBox4.Text
                            MAIN.TextBox4.Text = SL
                            Me.Dispose()
                            SystemID.TestInitialize()
                        ElseIf FUNNCTIONLIB.Initialize("SYSTEM_ID", TextBox1.Text, CONFIG.TextBox5.Text, ComboBox2.SelectedItem.ToString, CONFIG.TextBox1.Text, CONFIG.TextBox2.Text, CONFIG.TextBox3.Text, True) = False Then
                            Exit Sub
                        End If
                    ElseIf TextBox4.Text = "GOLDEN" Then
                        MAIN.ListBox1.Items.Add($"{DateTime.Now.ToString} DUT ID : GOLDEN")
                        MAIN.ListBox1.SelectedIndex = MAIN.ListBox1.Items.Count - 1
                        MAIN.TextBox1.Text = Me.ComboBox1.SelectedItem
                        MAIN.TextBox2.Text = Me.ComboBox2.SelectedItem
                        MAIN.TextBox153.Text = Me.TextBox3.Text
                        MAIN.TextBox6.Text = Me.TextBox4.Text
                        MAIN.TextBox4.Text = SL
                        Me.Dispose()
                        SystemID.TestInitialize()
                    ElseIf TextBox4.Text = "INVALID" Then
                        MAIN.ListBox1.Items.Add($"{DateTime.Now.ToString} DUT ID : INVALID")
                        MAIN.ListBox1.SelectedIndex = MAIN.ListBox1.Items.Count - 1
                        Dim iret As Object = MsgBox("INVALID SERIAL NUMBER ENTERD", vbCritical + vbOKOnly, "FUNNLIB.Initialize")
                        If iret = vbOK Then
                            TextBox1.Clear()
                            Exit Sub
                        End If
                    End If
                ElseIf SETUP.CheckBox1.Checked = False Then
                    SL = FUNNCTIONLIB.SLtoUpper(TextBox1.Text, False)
                    If TextBox4.Text = "PRODUCTION" Then
                        MAIN.ListBox1.Items.Add($"{DateTime.Now.ToString} DUT ID : PRODUCTION")
                        MAIN.ListBox1.SelectedIndex = MAIN.ListBox1.Items.Count - 1
                        If FUNNCTIONLIB.Initialize("SYSTEM_ID", TextBox1.Text, CONFIG.TextBox5.Text, ComboBox2.SelectedItem.ToString, CONFIG.TextBox1.Text, CONFIG.TextBox2.Text, CONFIG.TextBox3.Text, True) = True Then
                            MAIN.TextBox1.Text = Me.ComboBox1.SelectedItem
                            MAIN.TextBox2.Text = Me.ComboBox2.SelectedItem
                            MAIN.TextBox153.Text = Me.TextBox3.Text
                            MAIN.TextBox6.Text = Me.TextBox4.Text
                            MAIN.TextBox4.Text = SL
                            Me.Dispose()
                            SystemID.TestInitialize()
                        ElseIf FUNNCTIONLIB.Initialize("SYSTEM_ID", TextBox1.Text, CONFIG.TextBox5.Text, ComboBox2.SelectedItem.ToString, CONFIG.TextBox1.Text, CONFIG.TextBox2.Text, CONFIG.TextBox3.Text, True) = False Then
                            Exit Sub
                        End If
                    ElseIf TextBox4.Text = "GOLDEN" Then
                        MAIN.ListBox1.Items.Add($"{DateTime.Now.ToString} DUT ID : GOLDEN")
                        MAIN.ListBox1.SelectedIndex = MAIN.ListBox1.Items.Count - 1
                        MAIN.TextBox1.Text = Me.ComboBox1.SelectedItem
                        MAIN.TextBox2.Text = Me.ComboBox2.SelectedItem
                        MAIN.TextBox153.Text = Me.TextBox3.Text
                        MAIN.TextBox6.Text = Me.TextBox4.Text
                        MAIN.TextBox4.Text = SL
                        Me.Dispose()
                        SystemID.TestInitialize()
                    ElseIf TextBox4.Text = "INVALID" Then
                        MAIN.ListBox1.Items.Add($"{DateTime.Now.ToString} DUT ID : INVALID")
                        MAIN.ListBox1.SelectedIndex = MAIN.ListBox1.Items.Count - 1
                        Dim iret As Object = MsgBox("INVALID SERIAL NUMBER ENTERD", vbCritical + vbOKOnly, "FUNNLIB.Initialize")
                        If iret = vbOK Then
                            TextBox1.Clear()
                            Exit Sub
                        End If
                    End If
                End If
            ElseIf ComboBox3.SelectedItem = "MANUAL" Then
                MAIN.RadioButton2.Checked = True
                MAIN.TreeView1.Enabled = True
                MAIN.ListBox1.Items.Add($"{DateTime.Now.ToString} TEST SEQUENCE : SYSTEM ID > PRE-COATING > MANUAL")
                MAIN.ListBox1.SelectedIndex = MAIN.ListBox1.Items.Count - 1
                If SETUP.CheckBox1.Checked = True Then
                    SL = FUNNCTIONLIB.SLtoUpper(TextBox1.Text, True)
                    If TextBox4.Text = "PRODUCTION" Then
                        MAIN.ListBox1.Items.Add($"{DateTime.Now.ToString} DUT ID : PRODUCTION")
                        MAIN.ListBox1.SelectedIndex = MAIN.ListBox1.Items.Count - 1
                        If FUNNCTIONLIB.Initialize("SYSTEM_ID", Me.TextBox1.Text, CONFIG.TextBox5.Text, ComboBox2.SelectedItem.ToString, CONFIG.TextBox1.Text, CONFIG.TextBox2.Text, CONFIG.TextBox3.Text, True) = True Then
                            MAIN.TextBox1.Text = Me.ComboBox1.SelectedItem
                            MAIN.TextBox2.Text = Me.ComboBox2.SelectedItem
                            MAIN.TextBox153.Text = Me.TextBox3.Text
                            MAIN.TextBox6.Text = Me.TextBox4.Text
                            MAIN.TextBox4.Text = SL
                            Me.Dispose()
                            SystemID.TestInitialize()
                        ElseIf FUNNCTIONLIB.Initialize("SYSTEM_ID", TextBox1.Text, CONFIG.TextBox5.Text, ComboBox2.SelectedItem.ToString, CONFIG.TextBox1.Text, CONFIG.TextBox2.Text, CONFIG.TextBox3.Text, True) = False Then
                            Exit Sub
                        End If
                    ElseIf TextBox4.Text = "GOLDEN" Then
                        MAIN.ListBox1.Items.Add($"{DateTime.Now.ToString} DUT ID : GOLDEN")
                        MAIN.ListBox1.SelectedIndex = MAIN.ListBox1.Items.Count - 1
                        MAIN.TextBox1.Text = Me.ComboBox1.SelectedItem
                        MAIN.TextBox2.Text = Me.ComboBox2.SelectedItem
                        MAIN.TextBox153.Text = Me.TextBox3.Text
                        MAIN.TextBox6.Text = Me.TextBox4.Text
                        MAIN.TextBox4.Text = SL
                        Me.Dispose()
                        SystemID.TestInitialize()
                    ElseIf TextBox4.Text = "INVALID" Then
                        MAIN.ListBox1.Items.Add($"{DateTime.Now.ToString} DUT ID : INVALID")
                        MAIN.ListBox1.SelectedIndex = MAIN.ListBox1.Items.Count - 1
                        Dim iret As Object = MsgBox("INVALID SERIAL NUMBER ENTERD", vbCritical + vbOKOnly, "FUNNLIB.Initialize")
                        If iret = vbOK Then
                            TextBox1.Clear()
                            Exit Sub
                        End If
                    End If
                ElseIf SETUP.CheckBox1.Checked = False Then
                    SL = FUNNCTIONLIB.SLtoUpper(TextBox1.Text, False)
                    If TextBox4.Text = "PRODUCTION" Then
                        MAIN.ListBox1.Items.Add($"{DateTime.Now.ToString} DUT ID : PRODUCTION")
                        MAIN.ListBox1.SelectedIndex = MAIN.ListBox1.Items.Count - 1
                        If FUNNCTIONLIB.Initialize("SYSTEM_ID", TextBox1.Text, CONFIG.TextBox5.Text, ComboBox2.SelectedItem.ToString, CONFIG.TextBox1.Text, CONFIG.TextBox2.Text, CONFIG.TextBox3.Text, True) = True Then
                            MAIN.TextBox1.Text = Me.ComboBox1.SelectedItem
                            MAIN.TextBox2.Text = Me.ComboBox2.SelectedItem
                            MAIN.TextBox153.Text = Me.TextBox3.Text
                            MAIN.TextBox6.Text = Me.TextBox4.Text
                            MAIN.TextBox4.Text = SL
                            Me.Dispose()
                            SystemID.TestInitialize()
                        ElseIf FUNNCTIONLIB.Initialize("SYSTEM_ID", TextBox1.Text, CONFIG.TextBox5.Text, ComboBox2.SelectedItem.ToString, CONFIG.TextBox1.Text, CONFIG.TextBox2.Text, CONFIG.TextBox3.Text, True) = False Then
                            Exit Sub
                        End If
                    ElseIf TextBox4.Text = "GOLDEN" Then
                        MAIN.ListBox1.Items.Add($"{DateTime.Now.ToString} DUT ID : GOLDEN")
                        MAIN.ListBox1.SelectedIndex = MAIN.ListBox1.Items.Count - 1
                        MAIN.TextBox1.Text = Me.ComboBox1.SelectedItem
                        MAIN.TextBox2.Text = Me.ComboBox2.SelectedItem
                        MAIN.TextBox153.Text = Me.TextBox3.Text
                        MAIN.TextBox6.Text = Me.TextBox4.Text
                        MAIN.TextBox4.Text = SL
                        Me.Dispose()
                        SystemID.TestInitialize()
                    ElseIf TextBox4.Text = "INVALID" Then
                        MAIN.ListBox1.Items.Add($"{DateTime.Now.ToString} DUT ID : INVALID")
                        MAIN.ListBox1.SelectedIndex = MAIN.ListBox1.Items.Count - 1
                        Dim iret As Object = MsgBox("INVALID SERIAL NUMBER ENTERD", vbCritical + vbOKOnly, "FUNNLIB.Initialize")
                        If iret = vbOK Then
                            TextBox1.Clear()
                            Exit Sub
                        End If
                    End If
                End If
            End If
        ElseIf ComboBox1.SelectedItem = "SYSTEM ID" And ComboBox2.SelectedItem = "POST-COATING" Then
            If ComboBox3.SelectedItem = "AUTO" Then
                MAIN.RadioButton1.Checked = True
                MAIN.TreeView1.Enabled = False
                MAIN.ListBox1.Items.Add($"{DateTime.Now.ToString} TEST SEQUENCE : SYSTEM ID > POST-COATING > AUTO")
                MAIN.ListBox1.SelectedIndex = MAIN.ListBox1.Items.Count - 1
                If SETUP.CheckBox1.Checked = True Then
                    SL = FUNNCTIONLIB.SLtoUpper(TextBox1.Text, True)
                    If TextBox4.Text = "PRODUCTION" Then
                        MAIN.ListBox1.Items.Add($"{DateTime.Now.ToString} DUT ID : PRODUCTION")
                        MAIN.ListBox1.SelectedIndex = MAIN.ListBox1.Items.Count - 1
                        If FUNNCTIONLIB.Initialize("SYSTEM_ID", Me.TextBox1.Text, CONFIG.TextBox5.Text, ComboBox2.SelectedItem.ToString, CONFIG.TextBox1.Text, CONFIG.TextBox2.Text, CONFIG.TextBox3.Text, True) = True Then
                            MAIN.TextBox1.Text = Me.ComboBox1.SelectedItem
                            MAIN.TextBox2.Text = Me.ComboBox2.SelectedItem
                            MAIN.TextBox153.Text = Me.TextBox3.Text
                            MAIN.TextBox6.Text = Me.TextBox4.Text
                            MAIN.TextBox4.Text = SL
                            Me.Dispose()
                            SystemID.TestInitialize()
                        ElseIf FUNNCTIONLIB.Initialize("SYSTEM_ID", TextBox1.Text, CONFIG.TextBox5.Text, ComboBox2.SelectedItem.ToString, CONFIG.TextBox1.Text, CONFIG.TextBox2.Text, CONFIG.TextBox3.Text, True) = False Then
                            Exit Sub
                        End If
                    ElseIf TextBox4.Text = "GOLDEN" Then
                        MAIN.ListBox1.Items.Add($"{DateTime.Now.ToString} DUT ID : GOLDEN")
                        MAIN.ListBox1.SelectedIndex = MAIN.ListBox1.Items.Count - 1
                        MAIN.TextBox1.Text = Me.ComboBox1.SelectedItem
                        MAIN.TextBox2.Text = Me.ComboBox2.SelectedItem
                        MAIN.TextBox153.Text = Me.TextBox3.Text
                        MAIN.TextBox6.Text = Me.TextBox4.Text
                        MAIN.TextBox4.Text = SL
                        Me.Dispose()
                        SystemID.TestInitialize()
                    ElseIf TextBox4.Text = "INVALID" Then
                        MAIN.ListBox1.Items.Add($"{DateTime.Now.ToString} DUT ID : INVALID")
                        MAIN.ListBox1.SelectedIndex = MAIN.ListBox1.Items.Count - 1
                        Dim iret As Object = MsgBox("INVALID SERIAL NUMBER ENTERD", vbCritical + vbOKOnly, "FUNNLIB.Initialize")
                        If iret = vbOK Then
                            TextBox1.Clear()
                            Exit Sub
                        End If
                    End If
                ElseIf SETUP.CheckBox1.Checked = False Then
                    SL = FUNNCTIONLIB.SLtoUpper(TextBox1.Text, False)
                    If TextBox4.Text = "PRODUCTION" Then
                        MAIN.ListBox1.Items.Add($"{DateTime.Now.ToString} DUT ID : PRODUCTION")
                        MAIN.ListBox1.SelectedIndex = MAIN.ListBox1.Items.Count - 1
                        If FUNNCTIONLIB.Initialize("SYSTEM_ID", TextBox1.Text, CONFIG.TextBox5.Text, ComboBox2.SelectedItem.ToString, CONFIG.TextBox1.Text, CONFIG.TextBox2.Text, CONFIG.TextBox3.Text, True) = True Then
                            MAIN.TextBox1.Text = Me.ComboBox1.SelectedItem
                            MAIN.TextBox2.Text = Me.ComboBox2.SelectedItem
                            MAIN.TextBox153.Text = Me.TextBox3.Text
                            MAIN.TextBox6.Text = Me.TextBox4.Text
                            MAIN.TextBox4.Text = SL
                            Me.Dispose()
                            SystemID.TestInitialize()
                        ElseIf FUNNCTIONLIB.Initialize("SYSTEM_ID", TextBox1.Text, CONFIG.TextBox5.Text, ComboBox2.SelectedItem.ToString, CONFIG.TextBox1.Text, CONFIG.TextBox2.Text, CONFIG.TextBox3.Text, True) = False Then
                            Exit Sub
                        End If
                    ElseIf TextBox4.Text = "GOLDEN" Then
                        MAIN.ListBox1.Items.Add($"{DateTime.Now.ToString} DUT ID : GOLDEN")
                        MAIN.ListBox1.SelectedIndex = MAIN.ListBox1.Items.Count - 1
                        MAIN.TextBox1.Text = Me.ComboBox1.SelectedItem
                        MAIN.TextBox2.Text = Me.ComboBox2.SelectedItem
                        MAIN.TextBox153.Text = Me.TextBox3.Text
                        MAIN.TextBox6.Text = Me.TextBox4.Text
                        MAIN.TextBox4.Text = SL
                        Me.Dispose()
                        SystemID.TestInitialize()
                    ElseIf TextBox4.Text = "INVALID" Then
                        MAIN.ListBox1.Items.Add($"{DateTime.Now.ToString} DUT ID : INVALID")
                        MAIN.ListBox1.SelectedIndex = MAIN.ListBox1.Items.Count - 1
                        Dim iret As Object = MsgBox("INVALID SERIAL NUMBER ENTERD", vbCritical + vbOKOnly, "FUNNLIB.Initialize")
                        If iret = vbOK Then
                            TextBox1.Clear()
                            Exit Sub
                        End If
                    End If
                End If
            ElseIf ComboBox3.SelectedItem = "MANUAL" Then
                MAIN.RadioButton2.Checked = True
                MAIN.TreeView1.Enabled = True
                MAIN.ListBox1.Items.Add($"{DateTime.Now.ToString} TEST SEQUENCE : SYSTEM ID > POST-COATING > MANUAL")
                MAIN.ListBox1.SelectedIndex = MAIN.ListBox1.Items.Count - 1
                If SETUP.CheckBox1.Checked = True Then
                    SL = FUNNCTIONLIB.SLtoUpper(TextBox1.Text, True)
                    If TextBox4.Text = "PRODUCTION" Then
                        MAIN.ListBox1.Items.Add($"{DateTime.Now.ToString} DUT ID : PRODUCTION")
                        MAIN.ListBox1.SelectedIndex = MAIN.ListBox1.Items.Count - 1
                        If FUNNCTIONLIB.Initialize("SYSTEM_ID", Me.TextBox1.Text, CONFIG.TextBox5.Text, ComboBox2.SelectedItem.ToString, CONFIG.TextBox1.Text, CONFIG.TextBox2.Text, CONFIG.TextBox3.Text, True) = True Then
                            MAIN.TextBox1.Text = Me.ComboBox1.SelectedItem
                            MAIN.TextBox2.Text = Me.ComboBox2.SelectedItem
                            MAIN.TextBox153.Text = Me.TextBox3.Text
                            MAIN.TextBox6.Text = Me.TextBox4.Text
                            MAIN.TextBox4.Text = SL
                            Me.Dispose()
                            SystemID.TestInitialize()
                        ElseIf FUNNCTIONLIB.Initialize("SYSTEM_ID", TextBox1.Text, CONFIG.TextBox5.Text, ComboBox2.SelectedItem.ToString, CONFIG.TextBox1.Text, CONFIG.TextBox2.Text, CONFIG.TextBox3.Text, True) = False Then
                            Exit Sub
                        End If
                    ElseIf TextBox4.Text = "GOLDEN" Then
                        MAIN.ListBox1.Items.Add($"{DateTime.Now.ToString} DUT ID : GOLDEN")
                        MAIN.ListBox1.SelectedIndex = MAIN.ListBox1.Items.Count - 1
                        MAIN.TextBox1.Text = Me.ComboBox1.SelectedItem
                        MAIN.TextBox2.Text = Me.ComboBox2.SelectedItem
                        MAIN.TextBox153.Text = Me.TextBox3.Text
                        MAIN.TextBox6.Text = Me.TextBox4.Text
                        MAIN.TextBox4.Text = SL
                        Me.Dispose()
                        SystemID.TestInitialize()
                    ElseIf TextBox4.Text = "INVALID" Then
                        MAIN.ListBox1.Items.Add($"{DateTime.Now.ToString} DUT ID : INVALID")
                        MAIN.ListBox1.SelectedIndex = MAIN.ListBox1.Items.Count - 1
                        Dim iret As Object = MsgBox("INVALID SERIAL NUMBER ENTERD", vbCritical + vbOKOnly, "FUNNLIB.Initialize")
                        If iret = vbOK Then
                            TextBox1.Clear()
                            Exit Sub
                        End If
                    End If
                ElseIf SETUP.CheckBox1.Checked = False Then
                    SL = FUNNCTIONLIB.SLtoUpper(TextBox1.Text, False)
                    If TextBox4.Text = "PRODUCTION" Then
                        MAIN.ListBox1.Items.Add($"{DateTime.Now.ToString} DUT ID : PRODUCTION")
                        MAIN.ListBox1.SelectedIndex = MAIN.ListBox1.Items.Count - 1
                        If FUNNCTIONLIB.Initialize("SYSTEM_ID", TextBox1.Text, CONFIG.TextBox5.Text, ComboBox2.SelectedItem.ToString, CONFIG.TextBox1.Text, CONFIG.TextBox2.Text, CONFIG.TextBox3.Text, True) = True Then
                            MAIN.TextBox1.Text = Me.ComboBox1.SelectedItem
                            MAIN.TextBox2.Text = Me.ComboBox2.SelectedItem
                            MAIN.TextBox153.Text = Me.TextBox3.Text
                            MAIN.TextBox6.Text = Me.TextBox4.Text
                            MAIN.TextBox4.Text = SL
                            Me.Dispose()
                            SystemID.TestInitialize()
                        ElseIf FUNNCTIONLIB.Initialize("SYSTEM_ID", TextBox1.Text, CONFIG.TextBox5.Text, ComboBox2.SelectedItem.ToString, CONFIG.TextBox1.Text, CONFIG.TextBox2.Text, CONFIG.TextBox3.Text, True) = False Then
                            Exit Sub
                        End If
                    ElseIf TextBox4.Text = "GOLDEN" Then
                        MAIN.ListBox1.Items.Add($"{DateTime.Now.ToString} DUT ID : GOLDEN")
                        MAIN.ListBox1.SelectedIndex = MAIN.ListBox1.Items.Count - 1
                        MAIN.TextBox1.Text = Me.ComboBox1.SelectedItem
                        MAIN.TextBox2.Text = Me.ComboBox2.SelectedItem
                        MAIN.TextBox153.Text = Me.TextBox3.Text
                        MAIN.TextBox6.Text = Me.TextBox4.Text
                        MAIN.TextBox4.Text = SL
                        Me.Dispose()
                        SystemID.TestInitialize()
                    ElseIf TextBox4.Text = "INVALID" Then
                        MAIN.ListBox1.Items.Add($"{DateTime.Now.ToString} DUT ID : INVALID")
                        MAIN.ListBox1.SelectedIndex = MAIN.ListBox1.Items.Count - 1
                        Dim iret As Object = MsgBox("INVALID SERIAL NUMBER ENTERD", vbCritical + vbOKOnly, "FUNNLIB.Initialize")
                        If iret = vbOK Then
                            TextBox1.Clear()
                            Exit Sub
                        End If
                    End If
                End If
            End If
        ElseIf ComboBox1.SelectedItem = "BATTERY ID" And ComboBox2.SelectedItem = "PRE-COATING" Then
            If ComboBox3.SelectedItem = "AUTO" Then
                MAIN.RadioButton1.Checked = True
                MAIN.TreeView1.Enabled = False
                MAIN.ListBox1.Items.Add($"{DateTime.Now.ToString} TEST SEQUENCE : BATTERY ID > PRE-COATING > AUTO")
                MAIN.ListBox1.SelectedIndex = MAIN.ListBox1.Items.Count - 1
                If SETUP.CheckBox1.Checked = True Then
                    SL = FUNNCTIONLIB.SLtoUpper(TextBox1.Text, True)
                    If TextBox4.Text = "PRODUCTION" Then
                        MAIN.ListBox1.Items.Add($"{DateTime.Now.ToString} DUT ID : PRODUCTION")
                        MAIN.ListBox1.SelectedIndex = MAIN.ListBox1.Items.Count - 1
                        If FUNNCTIONLIB.Initialize("BATTERY_ID", TextBox1.Text, CONFIG.TextBox6.Text, ComboBox2.SelectedItem.ToString, CONFIG.TextBox1.Text, CONFIG.TextBox2.Text, CONFIG.TextBox3.Text, True) = True Then
                            MAIN.TextBox1.Text = Me.ComboBox1.SelectedItem
                            MAIN.TextBox2.Text = Me.ComboBox2.SelectedItem
                            MAIN.TextBox153.Text = Me.TextBox3.Text
                            MAIN.TextBox6.Text = Me.TextBox4.Text
                            MAIN.TextBox4.Text = SL
                            Me.Dispose()
                            BatteryID.TestInitialize()
                        ElseIf FUNNCTIONLIB.Initialize("BATTERY_ID", TextBox1.text, CONFIG.TextBox6.Text, ComboBox2.SelectedItem.ToString, CONFIG.TextBox1.Text, CONFIG.TextBox2.Text, CONFIG.TextBox3.Text, True) = False Then
                            Exit Sub
                        End If
                    ElseIf TextBox4.Text = "GOLDEN" Then
                        MAIN.ListBox1.Items.Add($"{DateTime.Now.ToString} DUT ID : GOLDEN")
                        MAIN.ListBox1.SelectedIndex = MAIN.ListBox1.Items.Count - 1
                        MAIN.TextBox1.Text = Me.ComboBox1.SelectedItem
                        MAIN.TextBox2.Text = Me.ComboBox2.SelectedItem
                        MAIN.TextBox153.Text = Me.TextBox3.Text
                        MAIN.TextBox6.Text = Me.TextBox4.Text
                        MAIN.TextBox4.Text = SL
                        Me.Dispose()
                        BatteryID.TestInitialize()
                    ElseIf TextBox4.Text = "INVALID" Then
                        MAIN.ListBox1.Items.Add($"{DateTime.Now.ToString} DUT ID : INVALID")
                        MAIN.ListBox1.SelectedIndex = MAIN.ListBox1.Items.Count - 1
                        Dim iret As Object = MsgBox("INVALID SERIAL NUMBER ENTERD", vbCritical + vbOKOnly, "FUNNLIB.Initialize")
                        If iret = vbOK Then
                            TextBox1.Clear()
                            Exit Sub
                        End If
                    End If
                ElseIf SETUP.CheckBox1.Checked = False Then
                    SL = FUNNCTIONLIB.SLtoUpper(TextBox1.Text, False)
                    If TextBox4.Text = "PRODUCTION" Then
                        MAIN.ListBox1.Items.Add($"{DateTime.Now.ToString} DUT ID : PRODUCTION")
                        MAIN.ListBox1.SelectedIndex = MAIN.ListBox1.Items.Count - 1
                        If FUNNCTIONLIB.Initialize("BATTERY_ID", TextBox1.Text, CONFIG.TextBox6.Text, ComboBox2.SelectedItem.ToString, CONFIG.TextBox1.Text, CONFIG.TextBox2.Text, CONFIG.TextBox3.Text, True) = True Then
                            MAIN.TextBox1.Text = Me.ComboBox1.SelectedItem
                            MAIN.TextBox2.Text = Me.ComboBox2.SelectedItem
                            MAIN.TextBox153.Text = Me.TextBox3.Text
                            MAIN.TextBox6.Text = Me.TextBox4.Text
                            MAIN.TextBox4.Text = SL
                            Me.Dispose()
                            BatteryID.TestInitialize()
                        ElseIf FUNNCTIONLIB.Initialize("BATTERY_ID", TextBox1.Text, CONFIG.TextBox6.Text, ComboBox2.SelectedItem.ToString, CONFIG.TextBox1.Text, CONFIG.TextBox2.Text, CONFIG.TextBox3.Text, True) = False Then
                            Exit Sub
                        End If
                    ElseIf TextBox4.Text = "GOLDEN" Then
                        MAIN.ListBox1.Items.Add($"{DateTime.Now.ToString} DUT ID : GOLDEN")
                        MAIN.ListBox1.SelectedIndex = MAIN.ListBox1.Items.Count - 1
                        MAIN.TextBox1.Text = Me.ComboBox1.SelectedItem
                        MAIN.TextBox2.Text = Me.ComboBox2.SelectedItem
                        MAIN.TextBox153.Text = Me.TextBox3.Text
                        MAIN.TextBox6.Text = Me.TextBox4.Text
                        MAIN.TextBox4.Text = SL
                        Me.Dispose()
                        BatteryID.TestInitialize()
                    ElseIf TextBox4.Text = "INVALID" Then
                        MAIN.ListBox1.Items.Add($"{DateTime.Now.ToString} DUT ID : INVALID")
                        MAIN.ListBox1.SelectedIndex = MAIN.ListBox1.Items.Count - 1
                        Dim iret As Object = MsgBox("INVALID SERIAL NUMBER ENTERD", vbCritical + vbOKOnly, "FUNNLIB.Initialize")
                        If iret = vbOK Then
                            TextBox1.Clear()
                            Exit Sub
                        End If
                    End If
                End If
            ElseIf ComboBox3.SelectedItem = "MANUAL" Then
                MAIN.RadioButton2.Checked = True
                MAIN.TreeView1.Enabled = True
                MAIN.ListBox1.Items.Add($"{DateTime.Now.ToString} TEST SEQUENCE : BATTERY ID > PRE-COATING > MANUAL")
                MAIN.ListBox1.SelectedIndex = MAIN.ListBox1.Items.Count - 1
                If SETUP.CheckBox1.Checked = True Then
                    SL = FUNNCTIONLIB.SLtoUpper(TextBox1.Text, True)
                    If TextBox4.Text = "PRODUCTION" Then
                        MAIN.ListBox1.Items.Add($"{DateTime.Now.ToString} DUT ID : PRODUCTION")
                        MAIN.ListBox1.SelectedIndex = MAIN.ListBox1.Items.Count - 1
                        If FUNNCTIONLIB.Initialize("BATTERY_ID", Me.TextBox1.Text, CONFIG.TextBox6.Text, ComboBox2.SelectedItem.ToString, CONFIG.TextBox1.Text, CONFIG.TextBox2.Text, CONFIG.TextBox3.Text, True) = True Then
                            MAIN.TextBox1.Text = Me.ComboBox1.SelectedItem
                            MAIN.TextBox2.Text = Me.ComboBox2.SelectedItem
                            MAIN.TextBox153.Text = Me.TextBox3.Text
                            MAIN.TextBox6.Text = Me.TextBox4.Text
                            MAIN.TextBox4.Text = SL
                            Me.Dispose()
                            BatteryID.TestInitialize()
                        ElseIf FUNNCTIONLIB.Initialize("BATTERY_ID", TextBox1.Text, CONFIG.TextBox6.Text, ComboBox2.SelectedItem.ToString, CONFIG.TextBox1.Text, CONFIG.TextBox2.Text, CONFIG.TextBox3.Text, True) = False Then
                            Exit Sub
                        End If
                    ElseIf TextBox4.Text = "GOLDEN" Then
                        MAIN.ListBox1.Items.Add($"{DateTime.Now.ToString} DUT ID : GOLDEN")
                        MAIN.ListBox1.SelectedIndex = MAIN.ListBox1.Items.Count - 1
                        MAIN.TextBox1.Text = Me.ComboBox1.SelectedItem
                        MAIN.TextBox2.Text = Me.ComboBox2.SelectedItem
                        MAIN.TextBox153.Text = Me.TextBox3.Text
                        MAIN.TextBox6.Text = Me.TextBox4.Text
                        MAIN.TextBox4.Text = SL
                        Me.Dispose()
                        BatteryID.TestInitialize()
                    ElseIf TextBox4.Text = "INVALID" Then
                        MAIN.ListBox1.Items.Add($"{DateTime.Now.ToString} DUT ID : INVALID")
                        MAIN.ListBox1.SelectedIndex = MAIN.ListBox1.Items.Count - 1
                        Dim iret As Object = MsgBox("INVALID SERIAL NUMBER ENTERD", vbCritical + vbOKOnly, "FUNNLIB.Initialize")
                        If iret = vbOK Then
                            TextBox1.Clear()
                            Exit Sub
                        End If
                    End If
                ElseIf SETUP.CheckBox1.Checked = False Then
                    SL = FUNNCTIONLIB.SLtoUpper(TextBox1.Text, False)
                    If TextBox4.Text = "PRODUCTION" Then
                        MAIN.ListBox1.Items.Add($"{DateTime.Now.ToString} DUT ID : PRODUCTION")
                        MAIN.ListBox1.SelectedIndex = MAIN.ListBox1.Items.Count - 1
                        If FUNNCTIONLIB.Initialize("BATTERY_ID", TextBox1.Text, CONFIG.TextBox6.Text, ComboBox2.SelectedItem.ToString, CONFIG.TextBox1.Text, CONFIG.TextBox2.Text, CONFIG.TextBox3.Text, True) = True Then
                            MAIN.TextBox1.Text = Me.ComboBox1.SelectedItem
                            MAIN.TextBox2.Text = Me.ComboBox2.SelectedItem
                            MAIN.TextBox153.Text = Me.TextBox3.Text
                            MAIN.TextBox6.Text = Me.TextBox4.Text
                            MAIN.TextBox4.Text = SL
                            Me.Dispose()
                            BatteryID.TestInitialize()
                        ElseIf FUNNCTIONLIB.Initialize("BATTERY_ID", TextBox1.Text, CONFIG.TextBox6.Text, ComboBox2.SelectedItem.ToString, CONFIG.TextBox1.Text, CONFIG.TextBox2.Text, CONFIG.TextBox3.Text, True) = False Then
                            Exit Sub
                        End If
                    ElseIf TextBox4.Text = "GOLDEN" Then
                        MAIN.ListBox1.Items.Add($"{DateTime.Now.ToString} DUT ID : GOLDEN")
                        MAIN.ListBox1.SelectedIndex = MAIN.ListBox1.Items.Count - 1
                        MAIN.TextBox1.Text = Me.ComboBox1.SelectedItem
                        MAIN.TextBox2.Text = Me.ComboBox2.SelectedItem
                        MAIN.TextBox153.Text = Me.TextBox3.Text
                        MAIN.TextBox6.Text = Me.TextBox4.Text
                        MAIN.TextBox4.Text = SL
                        Me.Dispose()
                        BatteryID.TestInitialize()
                    ElseIf TextBox4.Text = "INVALID" Then
                        MAIN.ListBox1.Items.Add($"{DateTime.Now.ToString} DUT ID : INVALID")
                        MAIN.ListBox1.SelectedIndex = MAIN.ListBox1.Items.Count - 1
                        Dim iret As Object = MsgBox("INVALID SERIAL NUMBER ENTERD", vbCritical + vbOKOnly, "FUNNLIB.Initialize")
                        If iret = vbOK Then
                            TextBox1.Clear()
                            Exit Sub
                        End If
                    End If
                End If
            End If
        ElseIf ComboBox1.SelectedItem = "BATTERY ID" And ComboBox2.SelectedItem = "POST-COATING" Then
            If ComboBox3.SelectedItem = "AUTO" Then
                MAIN.RadioButton1.Checked = True
                MAIN.TreeView1.Enabled = False
                MAIN.ListBox1.Items.Add($"{DateTime.Now.ToString} TEST SEQUENCE : BATTERY ID > POST-COATING > AUTO")
                MAIN.ListBox1.SelectedIndex = MAIN.ListBox1.Items.Count - 1
                If SETUP.CheckBox1.Checked = True Then
                    SL = FUNNCTIONLIB.SLtoUpper(TextBox1.Text, True)
                    If TextBox4.Text = "PRODUCTION" Then
                        MAIN.ListBox1.Items.Add($"{DateTime.Now.ToString} DUT ID : PRODUCTION")
                        MAIN.ListBox1.SelectedIndex = MAIN.ListBox1.Items.Count - 1
                        If FUNNCTIONLIB.Initialize("BATTERY_ID", Me.TextBox1.Text, CONFIG.TextBox6.Text, ComboBox2.SelectedItem.ToString, CONFIG.TextBox1.Text, CONFIG.TextBox2.Text, CONFIG.TextBox3.Text, True) = True Then
                            MAIN.TextBox1.Text = Me.ComboBox1.SelectedItem
                            MAIN.TextBox2.Text = Me.ComboBox2.SelectedItem
                            MAIN.TextBox153.Text = Me.TextBox3.Text
                            MAIN.TextBox6.Text = Me.TextBox4.Text
                            MAIN.TextBox4.Text = SL
                            Me.Dispose()
                            BatteryID.TestInitialize()
                        ElseIf FUNNCTIONLIB.Initialize("BATTERY_ID", TextBox1.Text, CONFIG.TextBox6.Text, ComboBox2.SelectedItem.ToString, CONFIG.TextBox1.Text, CONFIG.TextBox2.Text, CONFIG.TextBox3.Text, True) = False Then
                            Exit Sub
                        End If
                    ElseIf TextBox4.Text = "GOLDEN" Then
                        MAIN.ListBox1.Items.Add($"{DateTime.Now.ToString} DUT ID : GOLDEN")
                        MAIN.ListBox1.SelectedIndex = MAIN.ListBox1.Items.Count - 1
                        MAIN.TextBox1.Text = Me.ComboBox1.SelectedItem
                        MAIN.TextBox2.Text = Me.ComboBox2.SelectedItem
                        MAIN.TextBox153.Text = Me.TextBox3.Text
                        MAIN.TextBox6.Text = Me.TextBox4.Text
                        MAIN.TextBox4.Text = SL
                        Me.Dispose()
                        BatteryID.TestInitialize()
                    ElseIf TextBox4.Text = "INVALID" Then
                        MAIN.ListBox1.Items.Add($"{DateTime.Now.ToString} DUT ID : INVALID")
                        MAIN.ListBox1.SelectedIndex = MAIN.ListBox1.Items.Count - 1
                        Dim iret As Object = MsgBox("INVALID SERIAL NUMBER ENTERD", vbCritical + vbOKOnly, "FUNNLIB.Initialize")
                        If iret = vbOK Then
                            TextBox1.Clear()
                            Exit Sub
                        End If
                    End If
                ElseIf SETUP.CheckBox1.Checked = False Then
                    SL = FUNNCTIONLIB.SLtoUpper(TextBox1.Text, False)
                    If TextBox4.Text = "PRODUCTION" Then
                        MAIN.ListBox1.Items.Add($"{DateTime.Now.ToString} DUT ID : PRODUCTION")
                        MAIN.ListBox1.SelectedIndex = MAIN.ListBox1.Items.Count - 1
                        If FUNNCTIONLIB.Initialize("BATTERY_ID", TextBox1.Text, CONFIG.TextBox6.Text, ComboBox2.SelectedItem.ToString, CONFIG.TextBox1.Text, CONFIG.TextBox2.Text, CONFIG.TextBox3.Text, True) = True Then
                            MAIN.TextBox1.Text = Me.ComboBox1.SelectedItem
                            MAIN.TextBox2.Text = Me.ComboBox2.SelectedItem
                            MAIN.TextBox153.Text = Me.TextBox3.Text
                            MAIN.TextBox6.Text = Me.TextBox4.Text
                            MAIN.TextBox4.Text = SL
                            Me.Dispose()
                            BatteryID.TestInitialize()
                        ElseIf FUNNCTIONLIB.Initialize("BATTERY_ID", TextBox1.Text, CONFIG.TextBox6.Text, ComboBox2.SelectedItem.ToString, CONFIG.TextBox1.Text, CONFIG.TextBox2.Text, CONFIG.TextBox3.Text, True) = False Then
                            Exit Sub
                        End If
                    ElseIf TextBox4.Text = "GOLDEN" Then
                        MAIN.ListBox1.Items.Add($"{DateTime.Now.ToString} DUT ID : GOLDEN")
                        MAIN.ListBox1.SelectedIndex = MAIN.ListBox1.Items.Count - 1
                        MAIN.TextBox1.Text = Me.ComboBox1.SelectedItem
                        MAIN.TextBox2.Text = Me.ComboBox2.SelectedItem
                        MAIN.TextBox153.Text = Me.TextBox3.Text
                        MAIN.TextBox6.Text = Me.TextBox4.Text
                        MAIN.TextBox4.Text = SL
                        Me.Dispose()
                        BatteryID.TestInitialize()
                    ElseIf TextBox4.Text = "INVALID" Then
                        MAIN.ListBox1.Items.Add($"{DateTime.Now.ToString} DUT ID : INVALID")
                        MAIN.ListBox1.SelectedIndex = MAIN.ListBox1.Items.Count - 1
                        Dim iret As Object = MsgBox("INVALID SERIAL NUMBER ENTERD", vbCritical + vbOKOnly, "FUNNLIB.Initialize")
                        If iret = vbOK Then
                            TextBox1.Clear()
                            Exit Sub
                        End If
                    End If
                End If
            ElseIf ComboBox3.SelectedItem = "MANUAL" Then
                MAIN.RadioButton2.Checked = True
                MAIN.TreeView1.Enabled = True
                MAIN.ListBox1.Items.Add($"{DateTime.Now.ToString} TEST SEQUENCE : BATTERY ID > POST-COATING > MANUAL")
                MAIN.ListBox1.SelectedIndex = MAIN.ListBox1.Items.Count - 1
                If SETUP.CheckBox1.Checked = True Then
                    SL = FUNNCTIONLIB.SLtoUpper(TextBox1.Text, True)
                    If TextBox4.Text = "PRODUCTION" Then
                        MAIN.ListBox1.Items.Add($"{DateTime.Now.ToString} DUT ID : PRODUCTION")
                        MAIN.ListBox1.SelectedIndex = MAIN.ListBox1.Items.Count - 1
                        If FUNNCTIONLIB.Initialize("BATTERY_ID", Me.TextBox1.Text, CONFIG.TextBox6.Text, ComboBox2.SelectedItem.ToString, CONFIG.TextBox1.Text, CONFIG.TextBox2.Text, CONFIG.TextBox3.Text, True) = True Then
                            MAIN.TextBox1.Text = Me.ComboBox1.SelectedItem
                            MAIN.TextBox2.Text = Me.ComboBox2.SelectedItem
                            MAIN.TextBox153.Text = Me.TextBox3.Text
                            MAIN.TextBox6.Text = Me.TextBox4.Text
                            MAIN.TextBox4.Text = SL
                            Me.Dispose()
                            BatteryID.TestInitialize()
                        ElseIf FUNNCTIONLIB.Initialize("BATTERY_ID", TextBox1.Text, CONFIG.TextBox6.Text, ComboBox2.SelectedItem.ToString, CONFIG.TextBox1.Text, CONFIG.TextBox2.Text, CONFIG.TextBox3.Text, True) = False Then
                            Exit Sub
                        End If
                    ElseIf TextBox4.Text = "GOLDEN" Then
                        MAIN.ListBox1.Items.Add($"{DateTime.Now.ToString} DUT ID : GOLDEN")
                        MAIN.ListBox1.SelectedIndex = MAIN.ListBox1.Items.Count - 1
                        MAIN.TextBox1.Text = Me.ComboBox1.SelectedItem
                        MAIN.TextBox2.Text = Me.ComboBox2.SelectedItem
                        MAIN.TextBox153.Text = Me.TextBox3.Text
                        MAIN.TextBox6.Text = Me.TextBox4.Text
                        MAIN.TextBox4.Text = SL
                        Me.Dispose()
                        BatteryID.TestInitialize()
                    ElseIf TextBox4.Text = "INVALID" Then
                        MAIN.ListBox1.Items.Add($"{DateTime.Now.ToString} DUT ID : INVALID")
                        MAIN.ListBox1.SelectedIndex = MAIN.ListBox1.Items.Count - 1
                        Dim iret As Object = MsgBox("INVALID SERIAL NUMBER ENTERD", vbCritical + vbOKOnly, "FUNNLIB.Initialize")
                        If iret = vbOK Then
                            TextBox1.Clear()
                            Exit Sub
                        End If
                    End If
                ElseIf SETUP.CheckBox1.Checked = False Then
                    SL = FUNNCTIONLIB.SLtoUpper(TextBox1.Text, False)
                    If TextBox4.Text = "PRODUCTION" Then
                        MAIN.ListBox1.Items.Add($"{DateTime.Now.ToString} DUT ID : PRODUCTION")
                        MAIN.ListBox1.SelectedIndex = MAIN.ListBox1.Items.Count - 1
                        If FUNNCTIONLIB.Initialize("BATTERY_ID", TextBox1.Text, CONFIG.TextBox6.Text, ComboBox2.SelectedItem.ToString, CONFIG.TextBox1.Text, CONFIG.TextBox2.Text, CONFIG.TextBox3.Text, True) = True Then
                            MAIN.TextBox1.Text = Me.ComboBox1.SelectedItem
                            MAIN.TextBox2.Text = Me.ComboBox2.SelectedItem
                            MAIN.TextBox153.Text = Me.TextBox3.Text
                            MAIN.TextBox6.Text = Me.TextBox4.Text
                            MAIN.TextBox4.Text = SL
                            Me.Dispose()
                            BatteryID.TestInitialize()
                        ElseIf FUNNCTIONLIB.Initialize("BATTERY_ID", TextBox1.Text, CONFIG.TextBox6.Text, ComboBox2.SelectedItem.ToString, CONFIG.TextBox1.Text, CONFIG.TextBox2.Text, CONFIG.TextBox3.Text, True) = False Then
                            Exit Sub
                        End If
                    ElseIf TextBox4.Text = "GOLDEN" Then
                        MAIN.ListBox1.Items.Add($"{DateTime.Now.ToString} DUT ID : GOLDEN")
                        MAIN.ListBox1.SelectedIndex = MAIN.ListBox1.Items.Count - 1
                        MAIN.TextBox1.Text = Me.ComboBox1.SelectedItem
                        MAIN.TextBox2.Text = Me.ComboBox2.SelectedItem
                        MAIN.TextBox153.Text = Me.TextBox3.Text
                        MAIN.TextBox6.Text = Me.TextBox4.Text
                        MAIN.TextBox4.Text = SL
                        Me.Dispose()
                        BatteryID.TestInitialize()
                    ElseIf TextBox4.Text = "INVALID" Then
                        MAIN.ListBox1.Items.Add($"{DateTime.Now.ToString} DUT ID : INVALID")
                        MAIN.ListBox1.SelectedIndex = MAIN.ListBox1.Items.Count - 1
                        Dim iret As Object = MsgBox("INVALID SERIAL NUMBER ENTERD", vbCritical + vbOKOnly, "FUNNLIB.Initialize")
                        If iret = vbOK Then
                            TextBox1.Clear()
                            Exit Sub
                        End If
                    End If
                End If
            End If
        ElseIf ComboBox1.SelectedItem = "SMC" And ComboBox2.SelectedItem = "PRE-COATING" Then
            If ComboBox3.SelectedItem = "AUTO" Then
                MAIN.RadioButton1.Checked = True
                MAIN.TreeView1.Enabled = False
                MAIN.ListBox1.Items.Add($"{DateTime.Now.ToString} TEST SEQUENCE : SMC > PRE-COATING > AUTO")
                MAIN.ListBox1.SelectedIndex = MAIN.ListBox1.Items.Count - 1
                If SETUP.CheckBox1.Checked = True Then
                    SL = FUNNCTIONLIB.SLtoUpper(TextBox1.Text, True)
                    If TextBox4.Text = "PRODUCTION" Then
                        MAIN.ListBox1.Items.Add($"{DateTime.Now.ToString} DUT ID : PRODUCTION")
                        MAIN.ListBox1.SelectedIndex = MAIN.ListBox1.Items.Count - 1
                        If FUNNCTIONLIB.Initialize("SMC_1", Me.TextBox1.Text, CONFIG.TextBox7.Text, ComboBox2.SelectedItem.ToString, CONFIG.TextBox1.Text, CONFIG.TextBox2.Text, CONFIG.TextBox3.Text, True) = True Then
                            MAIN.TextBox1.Text = Me.ComboBox1.SelectedItem
                            MAIN.TextBox2.Text = Me.ComboBox2.SelectedItem
                            MAIN.TextBox153.Text = Me.TextBox3.Text
                            MAIN.TextBox6.Text = Me.TextBox4.Text
                            MAIN.TextBox4.Text = SL
                            Me.Dispose()
                            SMC.TestInitialize()
                        ElseIf FUNNCTIONLIB.Initialize("SMC_1", TextBox1.Text, CONFIG.TextBox7.Text, ComboBox2.SelectedItem.ToString, CONFIG.TextBox1.Text, CONFIG.TextBox2.Text, CONFIG.TextBox3.Text, True) = False Then
                            Exit Sub
                        End If
                    ElseIf TextBox4.Text = "GOLDEN" Then
                        MAIN.ListBox1.Items.Add($"{DateTime.Now.ToString} DUT ID : GOLDEN")
                        MAIN.ListBox1.SelectedIndex = MAIN.ListBox1.Items.Count - 1
                        MAIN.TextBox1.Text = Me.ComboBox1.SelectedItem
                        MAIN.TextBox2.Text = Me.ComboBox2.SelectedItem
                        MAIN.TextBox153.Text = Me.TextBox3.Text
                        MAIN.TextBox6.Text = Me.TextBox4.Text
                        MAIN.TextBox4.Text = SL
                        Me.Dispose()
                        SMC.TestInitialize()
                    ElseIf TextBox4.Text = "INVALID" Then
                        MAIN.ListBox1.Items.Add($"{DateTime.Now.ToString} DUT ID : INVALID")
                        MAIN.ListBox1.SelectedIndex = MAIN.ListBox1.Items.Count - 1
                        Dim iret As Object = MsgBox("INVALID SERIAL NUMBER ENTERD", vbCritical + vbOKOnly, "FUNNLIB.Initialize")
                        If iret = vbOK Then
                            TextBox1.Clear()
                            Exit Sub
                        End If
                    End If
                ElseIf SETUP.CheckBox1.Checked = False Then
                    SL = FUNNCTIONLIB.SLtoUpper(TextBox1.Text, False)
                    If TextBox4.Text = "PRODUCTION" Then
                        MAIN.ListBox1.Items.Add($"{DateTime.Now.ToString} DUT ID : PRODUCTION")
                        MAIN.ListBox1.SelectedIndex = MAIN.ListBox1.Items.Count - 1
                        If FUNNCTIONLIB.Initialize("SMC_1", TextBox1.Text, CONFIG.TextBox7.Text, ComboBox2.SelectedItem.ToString, CONFIG.TextBox1.Text, CONFIG.TextBox2.Text, CONFIG.TextBox3.Text, True) = True Then
                            MAIN.TextBox1.Text = Me.ComboBox1.SelectedItem
                            MAIN.TextBox2.Text = Me.ComboBox2.SelectedItem
                            MAIN.TextBox153.Text = Me.TextBox3.Text
                            MAIN.TextBox6.Text = Me.TextBox4.Text
                            MAIN.TextBox4.Text = SL
                            Me.Dispose()
                            SMC.TestInitialize()
                        ElseIf FUNNCTIONLIB.Initialize("SMC_1", TextBox1.Text, CONFIG.TextBox7.Text, ComboBox2.SelectedItem.ToString, CONFIG.TextBox1.Text, CONFIG.TextBox2.Text, CONFIG.TextBox3.Text, True) = False Then
                            Exit Sub
                        End If
                    ElseIf TextBox4.Text = "GOLDEN" Then
                        MAIN.ListBox1.Items.Add($"{DateTime.Now.ToString} DUT ID : GOLDEN")
                        MAIN.ListBox1.SelectedIndex = MAIN.ListBox1.Items.Count - 1
                        MAIN.TextBox1.Text = Me.ComboBox1.SelectedItem
                        MAIN.TextBox2.Text = Me.ComboBox2.SelectedItem
                        MAIN.TextBox153.Text = Me.TextBox3.Text
                        MAIN.TextBox6.Text = Me.TextBox4.Text
                        MAIN.TextBox4.Text = SL
                        Me.Dispose()
                        SMC.TestInitialize()
                    ElseIf TextBox4.Text = "INVALID" Then
                        MAIN.ListBox1.Items.Add($"{DateTime.Now.ToString} DUT ID : INVALID")
                        MAIN.ListBox1.SelectedIndex = MAIN.ListBox1.Items.Count - 1
                        Dim iret As Object = MsgBox("INVALID SERIAL NUMBER ENTERD", vbCritical + vbOKOnly, "FUNNLIB.Initialize")
                        If iret = vbOK Then
                            TextBox1.Clear()
                            Exit Sub
                        End If
                    End If
                End If
            End If
        ElseIf ComboBox1.SelectedItem = "SMC" And ComboBox2.SelectedItem = "POST-COATING" Then
            If ComboBox3.SelectedItem = "AUTO" Then
                MAIN.RadioButton1.Checked = True
                MAIN.TreeView1.Enabled = False
                MAIN.ListBox1.Items.Add($"{DateTime.Now.ToString} TEST SEQUENCE : SMC > POST-COATING > AUTO")
                MAIN.ListBox1.SelectedIndex = MAIN.ListBox1.Items.Count - 1
                If SETUP.CheckBox1.Checked = True Then
                    SL = FUNNCTIONLIB.SLtoUpper(TextBox1.Text, True)
                    If TextBox4.Text = "PRODUCTION" Then
                        MAIN.ListBox1.Items.Add($"{DateTime.Now.ToString} DUT ID : PRODUCTION")
                        MAIN.ListBox1.SelectedIndex = MAIN.ListBox1.Items.Count - 1
                        If FUNNCTIONLIB.Initialize("SMC_1", Me.TextBox1.Text, CONFIG.TextBox7.Text, ComboBox2.SelectedItem.ToString, CONFIG.TextBox1.Text, CONFIG.TextBox2.Text, CONFIG.TextBox3.Text, True) = True Then
                            MAIN.TextBox1.Text = Me.ComboBox1.SelectedItem
                            MAIN.TextBox2.Text = Me.ComboBox2.SelectedItem
                            MAIN.TextBox153.Text = Me.TextBox3.Text
                            MAIN.TextBox6.Text = Me.TextBox4.Text
                            MAIN.TextBox4.Text = SL
                            Me.Dispose()
                            SMC.TestInitialize()
                        ElseIf FUNNCTIONLIB.Initialize("SMC_1", TextBox1.Text, CONFIG.TextBox7.Text, ComboBox2.SelectedItem.ToString, CONFIG.TextBox1.Text, CONFIG.TextBox2.Text, CONFIG.TextBox3.Text, True) = False Then
                            Exit Sub
                        End If
                    ElseIf TextBox4.Text = "GOLDEN" Then
                        MAIN.ListBox1.Items.Add($"{DateTime.Now.ToString} DUT ID : GOLDEN")
                        MAIN.ListBox1.SelectedIndex = MAIN.ListBox1.Items.Count - 1
                        MAIN.TextBox1.Text = Me.ComboBox1.SelectedItem
                        MAIN.TextBox2.Text = Me.ComboBox2.SelectedItem
                        MAIN.TextBox153.Text = Me.TextBox3.Text
                        MAIN.TextBox6.Text = Me.TextBox4.Text
                        MAIN.TextBox4.Text = SL
                        Me.Dispose()
                        SMC.TestInitialize()
                    ElseIf TextBox4.Text = "INVALID" Then
                        MAIN.ListBox1.Items.Add($"{DateTime.Now.ToString} DUT ID : INVALID")
                        MAIN.ListBox1.SelectedIndex = MAIN.ListBox1.Items.Count - 1
                        Dim iret As Object = MsgBox("INVALID SERIAL NUMBER ENTERD", vbCritical + vbOKOnly, "FUNNLIB.Initialize")
                        If iret = vbOK Then
                            TextBox1.Clear()
                            Exit Sub
                        End If
                    End If
                ElseIf SETUP.CheckBox1.Checked = False Then
                    SL = FUNNCTIONLIB.SLtoUpper(TextBox1.Text, False)
                    If TextBox4.Text = "PRODUCTION" Then
                        MAIN.ListBox1.Items.Add($"{DateTime.Now.ToString} DUT ID : PRODUCTION")
                        MAIN.ListBox1.SelectedIndex = MAIN.ListBox1.Items.Count - 1
                        If FUNNCTIONLIB.Initialize("SMC_1", TextBox1.Text, CONFIG.TextBox7.Text, ComboBox2.SelectedItem.ToString, CONFIG.TextBox1.Text, CONFIG.TextBox2.Text, CONFIG.TextBox3.Text, True) = True Then
                            MAIN.TextBox1.Text = Me.ComboBox1.SelectedItem
                            MAIN.TextBox2.Text = Me.ComboBox2.SelectedItem
                            MAIN.TextBox153.Text = Me.TextBox3.Text
                            MAIN.TextBox6.Text = Me.TextBox4.Text
                            MAIN.TextBox4.Text = SL
                            Me.Dispose()
                            SMC.TestInitialize()
                        ElseIf FUNNCTIONLIB.Initialize("SMC_1", TextBox1.Text, CONFIG.TextBox7.Text, ComboBox2.SelectedItem.ToString, CONFIG.TextBox1.Text, CONFIG.TextBox2.Text, CONFIG.TextBox3.Text, True) = False Then
                            Exit Sub
                        End If
                    ElseIf TextBox4.Text = "GOLDEN" Then
                        MAIN.ListBox1.Items.Add($"{DateTime.Now.ToString} DUT ID : GOLDEN")
                        MAIN.ListBox1.SelectedIndex = MAIN.ListBox1.Items.Count - 1
                        MAIN.TextBox1.Text = Me.ComboBox1.SelectedItem
                        MAIN.TextBox2.Text = Me.ComboBox2.SelectedItem
                        MAIN.TextBox153.Text = Me.TextBox3.Text
                        MAIN.TextBox6.Text = Me.TextBox4.Text
                        MAIN.TextBox4.Text = SL
                        Me.Dispose()
                        SMC.TestInitialize()
                    ElseIf TextBox4.Text = "INVALID" Then
                        MAIN.ListBox1.Items.Add($"{DateTime.Now.ToString} DUT ID : INVALID")
                        MAIN.ListBox1.SelectedIndex = MAIN.ListBox1.Items.Count - 1
                        Dim iret As Object = MsgBox("INVALID SERIAL NUMBER ENTERD", vbCritical + vbOKOnly, "FUNNLIB.Initialize")
                        If iret = vbOK Then
                            TextBox1.Clear()
                            Exit Sub
                        End If
                    End If
                End If
            End If
        End If
    End Sub

    Private Sub TextBox1_TextChanged(sender As Object, e As EventArgs) Handles TextBox1.TextChanged
        MAIN.TextBox152.Text = Me.TextBox1.Text.ToUpper
        If ComboBox1.SelectedItem = "SYSTEM ID" Then
            TextBox4.Text = FUNNCTIONLIB.CheckifGolden(MAIN.TextBox152.Text, CONFIG.TextBox10.Text, CONFIG.TextBox5.Text)
        ElseIf ComboBox1.SelectedItem = "BATTERY ID" Then
            TextBox4.Text = FUNNCTIONLIB.CheckifGolden(MAIN.TextBox152.Text, CONFIG.TextBox11.Text, CONFIG.TextBox6.Text)
        ElseIf ComboBox1.SelectedItem = "SMC" Then
            TextBox4.Text = FUNNCTIONLIB.CheckifGolden(MAIN.TextBox152.Text, CONFIG.TextBox12.Text, CONFIG.TextBox7.Text)
        End If
    End Sub
End Class