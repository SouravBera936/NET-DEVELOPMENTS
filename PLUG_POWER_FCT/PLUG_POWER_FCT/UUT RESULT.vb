
Public Class UUT_RESULT
    Dim productversion = Application.ProductVersion.ToString
    Dim Products = {"SYSTEM ID", "BATTERY ID", "SMC"}
    Dim Test_Methodes = {"PRE-COATING", "POST-COATING"}
    Dim Test_mode = {"AUTO", "MANUAL"}
    Private Sub UUT_RESULT_Load(sender As Object, e As EventArgs) Handles MyBase.Load
        Me.TextBox1.Text = MAIN.TextBox1.Text
        Me.TextBox3.Text = MAIN.TextBox6.Text
        Label3.Text = productversion
        If RadioButton1.Checked = True Then
            Me.BackColor = Color.LightGreen
            Label2.BackColor = Color.LightGreen
            Label2.Text = "TEST SEQUENCE PASSED"
            LinkLabel1.Visible = False
            Label1.BackColor = Color.LightGreen
            Label3.BackColor = Color.LightGreen
            MAIN.BackColor = Color.LightGreen
            MAIN.GroupBox1.BackColor = Color.LightGreen
            MAIN.GroupBox2.BackColor = Color.LightGreen
            MAIN.TreeView1.BackColor = Color.LightGreen
            MAIN.ListBox1.BackColor = Color.LightGreen
            MAIN.ListBox2.BackColor = Color.LightGreen
        ElseIf RadioButton2.Checked = True Then
            Me.BackColor = Color.FromArgb(255, 127, 127)
            Label2.BackColor = Color.FromArgb(255, 127, 127)
            Label2.Text = "TEST SEQUENCE FAILED"
            LinkLabel1.Text = "VIEW FAILURE"
            LinkLabel1.Visible = True
            Label1.BackColor = Color.FromArgb(255, 127, 127)
            Label3.BackColor = Color.FromArgb(255, 127, 127)
            MAIN.BackColor = Color.FromArgb(255, 127, 127)
            MAIN.GroupBox1.BackColor = Color.FromArgb(255, 127, 127)
            MAIN.GroupBox2.BackColor = Color.FromArgb(255, 127, 127)
            MAIN.TreeView1.BackColor = Color.FromArgb(255, 127, 127)
            MAIN.ListBox1.BackColor = Color.FromArgb(255, 127, 127)
            MAIN.ListBox2.BackColor = Color.FromArgb(255, 127, 127)
        ElseIf RadioButton3.Checked = True Then
            Me.BackColor = Color.LightYellow
            Label2.BackColor = Color.LightYellow
            Label2.Text = "TEST SEQUENCE ERROR"
            LinkLabel1.Visible = True
            LinkLabel1.Text = "VIEW ERROR CODE"
            Label1.BackColor = Color.LightYellow
            Label3.BackColor = Color.LightYellow
            MAIN.BackColor = Color.LightYellow
            MAIN.GroupBox1.BackColor = Color.LightYellow
            MAIN.GroupBox2.BackColor = Color.LightYellow
            MAIN.TreeView1.BackColor = Color.LightYellow
            MAIN.ListBox1.BackColor = Color.LightYellow
            MAIN.ListBox2.BackColor = Color.LightYellow
        End If
        If SETUP.CheckBox7.Checked = True Then
            MAIN.CheckBox1.Checked = True
        Else
            MAIN.CheckBox1.Checked = False
        End If
    End Sub

    Private Sub Button1_Click(sender As Object, e As EventArgs) Handles Button1.Click
        If TextBox1.Text = "SMC" And TextBox3.Text = "PRODUCTION" Then
            If TextBox2.Text = "SMC_INIT" And RadioButton2.Checked = True Then
                SMC.CalculateCount("FAIL")
                SMC.ClearReportBeforeGeneration()
                Dim lastItem As String = MAIN.ListBox1.Items(MAIN.ListBox1.Items.Count - 1).ToString()
                SMC.GenerateReportOnerrorInit(MAIN.TextBox3.Text.ToString, MAIN.TextBox2.Text.ToString, MAIN.TextBox4.Text.ToString, MAIN.Button43.Text.ToString, MAIN.TextBox5.Text.ToString, MAIN.Label10.Text.ToString, MAIN.TextBox153.Text.ToString, lastItem)
                SMC.SMCCreateRecord(CONFIG.TextBox1.Text, CONFIG.TextBox2.Text, CONFIG.TextBox3.Text, MAIN.TextBox4.Text.ToString, MAIN.TextBox2.Text.ToString, MAIN.TextBox3.Text.ToString, MAIN.TextBox6.Text.ToString, MAIN.TextBox153.Text.ToString, MAIN.TextBox27.Text.ToString, MAIN.TextBox37.Text.ToString, MAIN.TextBox43.Text.ToString, MAIN.TextBox40.Text.ToString, MAIN.TextBox54.Text.ToString,
                MAIN.TextBox58.Text.ToString, MAIN.TextBox62.Text.ToString, MAIN.TextBox66.Text.ToString, MAIN.TextBox69.Text.ToString, MAIN.TextBox74.Text.ToString, MAIN.TextBox162.Text.ToString, MAIN.TextBox82.Text.ToString, MAIN.TextBox86.Text.ToString, MAIN.TextBox90.Text.ToString, MAIN.TextBox95.Text.ToString, MAIN.TextBox99.Text.ToString, MAIN.TextBox103.Text.ToString, MAIN.TextBox107.Text.ToString,
                MAIN.TextBox111.Text.ToString, MAIN.TextBox167.Text.ToString, MAIN.TextBox171.Text.ToString, MAIN.TextBox115.Text.ToString, MAIN.TextBox119.Text.ToString, MAIN.TextBox123.Text.ToString, MAIN.TextBox126.Text.ToString, MAIN.TextBox135.Text.ToString, MAIN.TextBox144.Text.ToString, MAIN.Button43.Text.ToString, DateTime.Now.ToString, MAIN.TextBox5.Text.ToString, MAIN.Label10.Text.ToString)
                SMC.ClearSequence()
                TrckRecord()
                Me.Close()
            ElseIf TextBox2.Text = "SMC_INIT" And RadioButton3.Checked = True Then
                SMC.CalculateCount("FAIL")
                SMC.ClearReportBeforeGeneration()
                Dim lastItem As String = MAIN.ListBox1.Items(MAIN.ListBox1.Items.Count - 1).ToString()
                SMC.GenerateReportOnerrorInit(MAIN.TextBox3.Text, MAIN.TextBox2.Text, MAIN.TextBox4.Text, MAIN.Button43.Text, MAIN.TextBox5.Text, MAIN.Label10.Text, MAIN.TextBox153.Text, lastItem)
                SMC.SMCCreateRecord(CONFIG.TextBox1.Text, CONFIG.TextBox2.Text, CONFIG.TextBox3.Text, MAIN.TextBox4.Text.ToString, MAIN.TextBox2.Text.ToString, MAIN.TextBox3.Text.ToString, MAIN.TextBox6.Text.ToString, MAIN.TextBox153.Text.ToString, MAIN.TextBox27.Text.ToString, MAIN.TextBox37.Text.ToString, MAIN.TextBox43.Text.ToString, MAIN.TextBox40.Text.ToString, MAIN.TextBox54.Text.ToString,
                MAIN.TextBox58.Text.ToString, MAIN.TextBox62.Text.ToString, MAIN.TextBox66.Text.ToString, MAIN.TextBox69.Text.ToString, MAIN.TextBox74.Text.ToString, MAIN.TextBox162.Text.ToString, MAIN.TextBox82.Text.ToString, MAIN.TextBox86.Text.ToString, MAIN.TextBox90.Text.ToString, MAIN.TextBox95.Text.ToString, MAIN.TextBox99.Text.ToString, MAIN.TextBox103.Text.ToString, MAIN.TextBox107.Text.ToString,
                MAIN.TextBox111.Text.ToString, MAIN.TextBox167.Text.ToString, MAIN.TextBox171.Text.ToString, MAIN.TextBox115.Text.ToString, MAIN.TextBox119.Text.ToString, MAIN.TextBox123.Text.ToString, MAIN.TextBox126.Text.ToString, MAIN.TextBox135.Text.ToString, MAIN.TextBox144.Text.ToString, MAIN.Button43.Text.ToString, DateTime.Now.ToString, MAIN.TextBox5.Text.ToString, MAIN.Label10.Text.ToString)
                SMC.ClearSequence()
                Me.Close()
                TrckRecord()
            ElseIf TextBox2.Text = "SET_POWER" And RadioButton2.Checked = True Then
                SMC.CalculateCount("FAIL")
                SMC.ClearReportBeforeGeneration()
                Dim lastItem As String = MAIN.ListBox1.Items(MAIN.ListBox1.Items.Count - 1).ToString()
                SMC.GenerateReportOnerrorSMCProgrammingNode1(MAIN.TextBox3.Text, MAIN.TextBox2.Text, MAIN.TextBox4.Text, MAIN.Button43.Text, MAIN.TextBox5.Text, MAIN.Label10.Text, MAIN.TextBox153.Text, lastItem, "TRUE", "FALSE")
                SMC.SMCCreateRecord(CONFIG.TextBox1.Text, CONFIG.TextBox2.Text, CONFIG.TextBox3.Text, MAIN.TextBox4.Text.ToString, MAIN.TextBox2.Text.ToString, MAIN.TextBox3.Text.ToString, MAIN.TextBox6.Text.ToString, MAIN.TextBox153.Text.ToString, MAIN.TextBox27.Text.ToString, MAIN.TextBox37.Text.ToString, MAIN.TextBox43.Text.ToString, MAIN.TextBox40.Text.ToString, MAIN.TextBox54.Text.ToString,
                MAIN.TextBox58.Text.ToString, MAIN.TextBox62.Text.ToString, MAIN.TextBox66.Text.ToString, MAIN.TextBox69.Text.ToString, MAIN.TextBox74.Text.ToString, MAIN.TextBox162.Text.ToString, MAIN.TextBox82.Text.ToString, MAIN.TextBox86.Text.ToString, MAIN.TextBox90.Text.ToString, MAIN.TextBox95.Text.ToString, MAIN.TextBox99.Text.ToString, MAIN.TextBox103.Text.ToString, MAIN.TextBox107.Text.ToString,
                MAIN.TextBox111.Text.ToString, MAIN.TextBox167.Text.ToString, MAIN.TextBox171.Text.ToString, MAIN.TextBox115.Text.ToString, MAIN.TextBox119.Text.ToString, MAIN.TextBox123.Text.ToString, MAIN.TextBox126.Text.ToString, MAIN.TextBox135.Text.ToString, MAIN.TextBox144.Text.ToString, MAIN.Button43.Text.ToString, DateTime.Now.ToString, MAIN.TextBox5.Text.ToString, MAIN.Label10.Text.ToString)
                SMC.ClearSequence()
                Me.Close()
                TrckRecord()
            ElseIf TextBox2.Text = "SET_POWER" And RadioButton3.Checked = True Then
                SMC.CalculateCount("FAIL")
                SMC.ClearReportBeforeGeneration()
                Dim lastItem As String = MAIN.ListBox1.Items(MAIN.ListBox1.Items.Count - 1).ToString()
                SMC.GenerateReportOnerrorSMCProgrammingNode1(MAIN.TextBox3.Text, MAIN.TextBox2.Text, MAIN.TextBox4.Text, MAIN.Button43.Text, MAIN.TextBox5.Text, MAIN.Label10.Text, MAIN.TextBox153.Text, lastItem, "TRUE", "ERROR")
                SMC.SMCCreateRecord(CONFIG.TextBox1.Text, CONFIG.TextBox2.Text, CONFIG.TextBox3.Text, MAIN.TextBox4.Text.ToString, MAIN.TextBox2.Text.ToString, MAIN.TextBox3.Text.ToString, MAIN.TextBox6.Text.ToString, MAIN.TextBox153.Text.ToString, MAIN.TextBox27.Text.ToString, MAIN.TextBox37.Text.ToString, MAIN.TextBox43.Text.ToString, MAIN.TextBox40.Text.ToString, MAIN.TextBox54.Text.ToString,
                MAIN.TextBox58.Text.ToString, MAIN.TextBox62.Text.ToString, MAIN.TextBox66.Text.ToString, MAIN.TextBox69.Text.ToString, MAIN.TextBox74.Text.ToString, MAIN.TextBox162.Text.ToString, MAIN.TextBox82.Text.ToString, MAIN.TextBox86.Text.ToString, MAIN.TextBox90.Text.ToString, MAIN.TextBox95.Text.ToString, MAIN.TextBox99.Text.ToString, MAIN.TextBox103.Text.ToString, MAIN.TextBox107.Text.ToString,
                MAIN.TextBox111.Text.ToString, MAIN.TextBox167.Text.ToString, MAIN.TextBox171.Text.ToString, MAIN.TextBox115.Text.ToString, MAIN.TextBox119.Text.ToString, MAIN.TextBox123.Text.ToString, MAIN.TextBox126.Text.ToString, MAIN.TextBox135.Text.ToString, MAIN.TextBox144.Text.ToString, MAIN.Button43.Text.ToString, DateTime.Now.ToString, MAIN.TextBox5.Text.ToString, MAIN.Label10.Text.ToString)
                SMC.ClearSequence()
                Me.Close()
                TrckRecord()
            ElseIf TextBox2.Text = "POWER_ON" And RadioButton2.Checked = True Then
                SMC.CalculateCount("FAIL")
                SMC.ClearReportBeforeGeneration()
                Dim lastItem As String = MAIN.ListBox1.Items(MAIN.ListBox1.Items.Count - 1).ToString()
                SMC.GenerateReportOnerrorSMCProgrammingNode1(MAIN.TextBox3.Text, MAIN.TextBox2.Text, MAIN.TextBox4.Text, MAIN.Button43.Text, MAIN.TextBox5.Text, MAIN.Label10.Text, MAIN.TextBox153.Text, lastItem, "TRUE", "FALSE")
                SMC.SMCCreateRecord(CONFIG.TextBox1.Text, CONFIG.TextBox2.Text, CONFIG.TextBox3.Text, MAIN.TextBox4.Text.ToString, MAIN.TextBox2.Text.ToString, MAIN.TextBox3.Text.ToString, MAIN.TextBox6.Text.ToString, MAIN.TextBox153.Text.ToString, MAIN.TextBox27.Text.ToString, MAIN.TextBox37.Text.ToString, MAIN.TextBox43.Text.ToString, MAIN.TextBox40.Text.ToString, MAIN.TextBox54.Text.ToString,
                MAIN.TextBox58.Text.ToString, MAIN.TextBox62.Text.ToString, MAIN.TextBox66.Text.ToString, MAIN.TextBox69.Text.ToString, MAIN.TextBox74.Text.ToString, MAIN.TextBox162.Text.ToString, MAIN.TextBox82.Text.ToString, MAIN.TextBox86.Text.ToString, MAIN.TextBox90.Text.ToString, MAIN.TextBox95.Text.ToString, MAIN.TextBox99.Text.ToString, MAIN.TextBox103.Text.ToString, MAIN.TextBox107.Text.ToString,
                MAIN.TextBox111.Text.ToString, MAIN.TextBox167.Text.ToString, MAIN.TextBox171.Text.ToString, MAIN.TextBox115.Text.ToString, MAIN.TextBox119.Text.ToString, MAIN.TextBox123.Text.ToString, MAIN.TextBox126.Text.ToString, MAIN.TextBox135.Text.ToString, MAIN.TextBox144.Text.ToString, MAIN.Button43.Text.ToString, DateTime.Now.ToString, MAIN.TextBox5.Text.ToString, MAIN.Label10.Text.ToString)
                SMC.ClearSequence()
                Me.Close()
                TrckRecord()
            ElseIf TextBox2.Text = "POWER_ON" And RadioButton3.Checked = True Then
                SMC.CalculateCount("FAIL")
                SMC.ClearReportBeforeGeneration()
                Dim lastItem As String = MAIN.ListBox1.Items(MAIN.ListBox1.Items.Count - 1).ToString()
                SMC.GenerateReportOnerrorSMCProgrammingNode1(MAIN.TextBox3.Text, MAIN.TextBox2.Text, MAIN.TextBox4.Text, MAIN.Button43.Text, MAIN.TextBox5.Text, MAIN.Label10.Text, MAIN.TextBox153.Text, lastItem, "TRUE", "ERROR")
                SMC.SMCCreateRecord(CONFIG.TextBox1.Text, CONFIG.TextBox2.Text, CONFIG.TextBox3.Text, MAIN.TextBox4.Text.ToString, MAIN.TextBox2.Text.ToString, MAIN.TextBox3.Text.ToString, MAIN.TextBox6.Text.ToString, MAIN.TextBox153.Text.ToString, MAIN.TextBox27.Text.ToString, MAIN.TextBox37.Text.ToString, MAIN.TextBox43.Text.ToString, MAIN.TextBox40.Text.ToString, MAIN.TextBox54.Text.ToString,
                MAIN.TextBox58.Text.ToString, MAIN.TextBox62.Text.ToString, MAIN.TextBox66.Text.ToString, MAIN.TextBox69.Text.ToString, MAIN.TextBox74.Text.ToString, MAIN.TextBox162.Text.ToString, MAIN.TextBox82.Text.ToString, MAIN.TextBox86.Text.ToString, MAIN.TextBox90.Text.ToString, MAIN.TextBox95.Text.ToString, MAIN.TextBox99.Text.ToString, MAIN.TextBox103.Text.ToString, MAIN.TextBox107.Text.ToString,
                MAIN.TextBox111.Text.ToString, MAIN.TextBox167.Text.ToString, MAIN.TextBox171.Text.ToString, MAIN.TextBox115.Text.ToString, MAIN.TextBox119.Text.ToString, MAIN.TextBox123.Text.ToString, MAIN.TextBox126.Text.ToString, MAIN.TextBox135.Text.ToString, MAIN.TextBox144.Text.ToString, MAIN.Button43.Text.ToString, DateTime.Now.ToString, MAIN.TextBox5.Text.ToString, MAIN.Label10.Text.ToString)
                SMC.ClearSequence()
                Me.Close()
                TrckRecord()
            ElseIf TextBox2.Text = "UNIFLASH_PROG" And RadioButton2.Checked = True Then
                SMC.CalculateCount("FAIL")
                SMC.ClearReportBeforeGeneration()
                Dim lastItem As String = MAIN.ListBox1.Items(MAIN.ListBox1.Items.Count - 1).ToString()
                SMC.GenerateReportOnErrorUniflashProg(MAIN.TextBox3.Text, MAIN.TextBox2.Text, MAIN.TextBox4.Text, MAIN.Button43.Text, MAIN.TextBox5.Text, MAIN.Label10.Text, MAIN.TextBox153.Text, lastItem, "TRUE", "TRUE", MAIN.TextBox27.Text)
                SMC.SMCCreateRecord(CONFIG.TextBox1.Text, CONFIG.TextBox2.Text, CONFIG.TextBox3.Text, MAIN.TextBox4.Text.ToString, MAIN.TextBox2.Text.ToString, MAIN.TextBox3.Text.ToString, MAIN.TextBox6.Text.ToString, MAIN.TextBox153.Text.ToString, MAIN.TextBox27.Text.ToString, MAIN.TextBox37.Text.ToString, MAIN.TextBox43.Text.ToString, MAIN.TextBox40.Text.ToString, MAIN.TextBox54.Text.ToString,
                MAIN.TextBox58.Text.ToString, MAIN.TextBox62.Text.ToString, MAIN.TextBox66.Text.ToString, MAIN.TextBox69.Text.ToString, MAIN.TextBox74.Text.ToString, MAIN.TextBox162.Text.ToString, MAIN.TextBox82.Text.ToString, MAIN.TextBox86.Text.ToString, MAIN.TextBox90.Text.ToString, MAIN.TextBox95.Text.ToString, MAIN.TextBox99.Text.ToString, MAIN.TextBox103.Text.ToString, MAIN.TextBox107.Text.ToString,
                MAIN.TextBox111.Text.ToString, MAIN.TextBox167.Text.ToString, MAIN.TextBox171.Text.ToString, MAIN.TextBox115.Text.ToString, MAIN.TextBox119.Text.ToString, MAIN.TextBox123.Text.ToString, MAIN.TextBox126.Text.ToString, MAIN.TextBox135.Text.ToString, MAIN.TextBox144.Text.ToString, MAIN.Button43.Text.ToString, DateTime.Now.ToString, MAIN.TextBox5.Text.ToString, MAIN.Label10.Text.ToString)
                SMC.ClearSequence()
                Me.Close()
                TrckRecord()
            ElseIf TextBox2.Text = "UNIFLASH_PROG" And RadioButton3.Checked = True Then
                SMC.CalculateCount("FAIL")
                SMC.ClearReportBeforeGeneration()
                Dim lastItem As String = MAIN.ListBox1.Items(MAIN.ListBox1.Items.Count - 1).ToString()
                SMC.GenerateReportOnErrorUniflashProg(MAIN.TextBox3.Text, MAIN.TextBox2.Text, MAIN.TextBox4.Text, MAIN.Button43.Text, MAIN.TextBox5.Text, MAIN.Label10.Text, MAIN.TextBox153.Text, lastItem, "TRUE", "TRUE", MAIN.TextBox27.Text)
                SMC.SMCCreateRecord(CONFIG.TextBox1.Text, CONFIG.TextBox2.Text, CONFIG.TextBox3.Text, MAIN.TextBox4.Text.ToString, MAIN.TextBox2.Text.ToString, MAIN.TextBox3.Text.ToString, MAIN.TextBox6.Text.ToString, MAIN.TextBox153.Text.ToString, MAIN.TextBox27.Text.ToString, MAIN.TextBox37.Text.ToString, MAIN.TextBox43.Text.ToString, MAIN.TextBox40.Text.ToString, MAIN.TextBox54.Text.ToString,
                MAIN.TextBox58.Text.ToString, MAIN.TextBox62.Text.ToString, MAIN.TextBox66.Text.ToString, MAIN.TextBox69.Text.ToString, MAIN.TextBox74.Text.ToString, MAIN.TextBox162.Text.ToString, MAIN.TextBox82.Text.ToString, MAIN.TextBox86.Text.ToString, MAIN.TextBox90.Text.ToString, MAIN.TextBox95.Text.ToString, MAIN.TextBox99.Text.ToString, MAIN.TextBox103.Text.ToString, MAIN.TextBox107.Text.ToString,
                MAIN.TextBox111.Text.ToString, MAIN.TextBox167.Text.ToString, MAIN.TextBox171.Text.ToString, MAIN.TextBox115.Text.ToString, MAIN.TextBox119.Text.ToString, MAIN.TextBox123.Text.ToString, MAIN.TextBox126.Text.ToString, MAIN.TextBox135.Text.ToString, MAIN.TextBox144.Text.ToString, MAIN.Button43.Text.ToString, DateTime.Now.ToString, MAIN.TextBox5.Text.ToString, MAIN.Label10.Text.ToString)
                SMC.ClearSequence()
                Me.Close()
                TrckRecord()
            ElseIf TextBox2.Text = "POWER_REC" And RadioButton2.Checked = True Then
                SMC.CalculateCount("FAIL")
                SMC.ClearReportBeforeGeneration()
                Dim lastItem As String = MAIN.ListBox1.Items(MAIN.ListBox1.Items.Count - 1).ToString()
                SMC.GenerateReportOnErrorOfFTDIProg(MAIN.TextBox3.Text, MAIN.TextBox2.Text, MAIN.TextBox4.Text, MAIN.Button43.Text, MAIN.TextBox5.Text, MAIN.Label10.Text, MAIN.TextBox153.Text, lastItem, "TRUE", "TRUE", MAIN.TextBox27.Text, MAIN.TextBox37.Text)
                SMC.SMCCreateRecord(CONFIG.TextBox1.Text, CONFIG.TextBox2.Text, CONFIG.TextBox3.Text, MAIN.TextBox4.Text.ToString, MAIN.TextBox2.Text.ToString, MAIN.TextBox3.Text.ToString, MAIN.TextBox6.Text.ToString, MAIN.TextBox153.Text.ToString, MAIN.TextBox27.Text.ToString, MAIN.TextBox37.Text.ToString, MAIN.TextBox43.Text.ToString, MAIN.TextBox40.Text.ToString, MAIN.TextBox54.Text.ToString,
                MAIN.TextBox58.Text.ToString, MAIN.TextBox62.Text.ToString, MAIN.TextBox66.Text.ToString, MAIN.TextBox69.Text.ToString, MAIN.TextBox74.Text.ToString, MAIN.TextBox162.Text.ToString, MAIN.TextBox82.Text.ToString, MAIN.TextBox86.Text.ToString, MAIN.TextBox90.Text.ToString, MAIN.TextBox95.Text.ToString, MAIN.TextBox99.Text.ToString, MAIN.TextBox103.Text.ToString, MAIN.TextBox107.Text.ToString,
                MAIN.TextBox111.Text.ToString, MAIN.TextBox167.Text.ToString, MAIN.TextBox171.Text.ToString, MAIN.TextBox115.Text.ToString, MAIN.TextBox119.Text.ToString, MAIN.TextBox123.Text.ToString, MAIN.TextBox126.Text.ToString, MAIN.TextBox135.Text.ToString, MAIN.TextBox144.Text.ToString, MAIN.Button43.Text.ToString, DateTime.Now.ToString, MAIN.TextBox5.Text.ToString, MAIN.Label10.Text.ToString)
                SMC.ClearSequence()
                Me.Close()
                TrckRecord()
            ElseIf TextBox2.Text = "POWER_REC" And RadioButton3.Checked = True Then
                SMC.CalculateCount("FAIL")
                SMC.ClearReportBeforeGeneration()
                Dim lastItem As String = MAIN.ListBox1.Items(MAIN.ListBox1.Items.Count - 1).ToString()
                SMC.GenerateReportOnErrorOfFTDIProg(MAIN.TextBox3.Text, MAIN.TextBox2.Text, MAIN.TextBox4.Text, MAIN.Button43.Text, MAIN.TextBox5.Text, MAIN.Label10.Text, MAIN.TextBox153.Text, lastItem, "TRUE", "TRUE", MAIN.TextBox27.Text, MAIN.TextBox37.Text)
                SMC.SMCCreateRecord(CONFIG.TextBox1.Text, CONFIG.TextBox2.Text, CONFIG.TextBox3.Text, MAIN.TextBox4.Text.ToString, MAIN.TextBox2.Text.ToString, MAIN.TextBox3.Text.ToString, MAIN.TextBox6.Text.ToString, MAIN.TextBox153.Text.ToString, MAIN.TextBox27.Text.ToString, MAIN.TextBox37.Text.ToString, MAIN.TextBox43.Text.ToString, MAIN.TextBox40.Text.ToString, MAIN.TextBox54.Text.ToString,
                MAIN.TextBox58.Text.ToString, MAIN.TextBox62.Text.ToString, MAIN.TextBox66.Text.ToString, MAIN.TextBox69.Text.ToString, MAIN.TextBox74.Text.ToString, MAIN.TextBox162.Text.ToString, MAIN.TextBox82.Text.ToString, MAIN.TextBox86.Text.ToString, MAIN.TextBox90.Text.ToString, MAIN.TextBox95.Text.ToString, MAIN.TextBox99.Text.ToString, MAIN.TextBox103.Text.ToString, MAIN.TextBox107.Text.ToString,
                MAIN.TextBox111.Text.ToString, MAIN.TextBox167.Text.ToString, MAIN.TextBox171.Text.ToString, MAIN.TextBox115.Text.ToString, MAIN.TextBox119.Text.ToString, MAIN.TextBox123.Text.ToString, MAIN.TextBox126.Text.ToString, MAIN.TextBox135.Text.ToString, MAIN.TextBox144.Text.ToString, MAIN.Button43.Text.ToString, DateTime.Now.ToString, MAIN.TextBox5.Text.ToString, MAIN.Label10.Text.ToString)
                SMC.ClearSequence()
                Me.Close()
                TrckRecord()
            ElseIf TextBox2.Text = "FTDI_PROG" And RadioButton2.Checked = True Then
                SMC.CalculateCount("FAIL")
                SMC.ClearReportBeforeGeneration()
                Dim lastItem As String = MAIN.ListBox1.Items(MAIN.ListBox1.Items.Count - 1).ToString()
                SMC.GenerateReportOnErrorOfFTDIProg(MAIN.TextBox3.Text, MAIN.TextBox2.Text, MAIN.TextBox4.Text, MAIN.Button43.Text, MAIN.TextBox5.Text, MAIN.Label10.Text, MAIN.TextBox153.Text, lastItem, "TRUE", "TRUE", MAIN.TextBox27.Text, MAIN.TextBox37.Text)
                SMC.SMCCreateRecord(CONFIG.TextBox1.Text, CONFIG.TextBox2.Text, CONFIG.TextBox3.Text, MAIN.TextBox4.Text.ToString, MAIN.TextBox2.Text.ToString, MAIN.TextBox3.Text.ToString, MAIN.TextBox6.Text.ToString, MAIN.TextBox153.Text.ToString, MAIN.TextBox27.Text.ToString, MAIN.TextBox37.Text.ToString, MAIN.TextBox43.Text.ToString, MAIN.TextBox40.Text.ToString, MAIN.TextBox54.Text.ToString,
                MAIN.TextBox58.Text.ToString, MAIN.TextBox62.Text.ToString, MAIN.TextBox66.Text.ToString, MAIN.TextBox69.Text.ToString, MAIN.TextBox74.Text.ToString, MAIN.TextBox162.Text.ToString, MAIN.TextBox82.Text.ToString, MAIN.TextBox86.Text.ToString, MAIN.TextBox90.Text.ToString, MAIN.TextBox95.Text.ToString, MAIN.TextBox99.Text.ToString, MAIN.TextBox103.Text.ToString, MAIN.TextBox107.Text.ToString,
                MAIN.TextBox111.Text.ToString, MAIN.TextBox167.Text.ToString, MAIN.TextBox171.Text.ToString, MAIN.TextBox115.Text.ToString, MAIN.TextBox119.Text.ToString, MAIN.TextBox123.Text.ToString, MAIN.TextBox126.Text.ToString, MAIN.TextBox135.Text.ToString, MAIN.TextBox144.Text.ToString, MAIN.Button43.Text.ToString, DateTime.Now.ToString, MAIN.TextBox5.Text.ToString, MAIN.Label10.Text.ToString)
                SMC.ClearSequence()
                Me.Close()
                TrckRecord()
            ElseIf TextBox2.Text = "FTDI_PROG" And RadioButton3.Checked = True Then
                SMC.CalculateCount("FAIL")
                SMC.ClearReportBeforeGeneration()
                Dim lastItem As String = MAIN.ListBox1.Items(MAIN.ListBox1.Items.Count - 1).ToString()
                SMC.GenerateReportOnErrorOfFTDIProg(MAIN.TextBox3.Text, MAIN.TextBox2.Text, MAIN.TextBox4.Text, MAIN.Button43.Text, MAIN.TextBox5.Text, MAIN.Label10.Text, MAIN.TextBox153.Text, lastItem, "TRUE", "TRUE", MAIN.TextBox27.Text, MAIN.TextBox37.Text)
                SMC.SMCCreateRecord(CONFIG.TextBox1.Text, CONFIG.TextBox2.Text, CONFIG.TextBox3.Text, MAIN.TextBox4.Text.ToString, MAIN.TextBox2.Text.ToString, MAIN.TextBox3.Text.ToString, MAIN.TextBox6.Text.ToString, MAIN.TextBox153.Text.ToString, MAIN.TextBox27.Text.ToString, MAIN.TextBox37.Text.ToString, MAIN.TextBox43.Text.ToString, MAIN.TextBox40.Text.ToString, MAIN.TextBox54.Text.ToString,
                MAIN.TextBox58.Text.ToString, MAIN.TextBox62.Text.ToString, MAIN.TextBox66.Text.ToString, MAIN.TextBox69.Text.ToString, MAIN.TextBox74.Text.ToString, MAIN.TextBox162.Text.ToString, MAIN.TextBox82.Text.ToString, MAIN.TextBox86.Text.ToString, MAIN.TextBox90.Text.ToString, MAIN.TextBox95.Text.ToString, MAIN.TextBox99.Text.ToString, MAIN.TextBox103.Text.ToString, MAIN.TextBox107.Text.ToString,
                MAIN.TextBox111.Text.ToString, MAIN.TextBox167.Text.ToString, MAIN.TextBox171.Text.ToString, MAIN.TextBox115.Text.ToString, MAIN.TextBox119.Text.ToString, MAIN.TextBox123.Text.ToString, MAIN.TextBox126.Text.ToString, MAIN.TextBox135.Text.ToString, MAIN.TextBox144.Text.ToString, MAIN.Button43.Text.ToString, DateTime.Now.ToString, MAIN.TextBox5.Text.ToString, MAIN.Label10.Text.ToString)
                SMC.ClearSequence()
                Me.Close()
                TrckRecord()
            ElseIf TextBox2.Text = "INIT_POWER" And RadioButton2.Checked = True Then
                SMC.CalculateCount("FAIL")
                ClearReportBeforeGeneration()
                Dim lastItem As String = MAIN.ListBox1.Items(MAIN.ListBox1.Items.Count - 1).ToString()
                SMC.GenerateReportonErrorSMCPowerTestNode0(MAIN.TextBox3.Text, MAIN.TextBox2.Text, MAIN.TextBox4.Text, MAIN.Button43.Text, MAIN.TextBox5.Text, MAIN.Label10.Text, MAIN.TextBox153.Text, lastItem, "TRUE", "TRUE", MAIN.TextBox27.Text, MAIN.TextBox37.Text, MAIN.TextBox43.Text)
                SMC.SMCCreateRecord(CONFIG.TextBox1.Text, CONFIG.TextBox2.Text, CONFIG.TextBox3.Text, MAIN.TextBox4.Text.ToString, MAIN.TextBox2.Text.ToString, MAIN.TextBox3.Text.ToString, MAIN.TextBox6.Text.ToString, MAIN.TextBox153.Text.ToString, MAIN.TextBox27.Text.ToString, MAIN.TextBox37.Text.ToString, MAIN.TextBox43.Text.ToString, MAIN.TextBox40.Text.ToString, MAIN.TextBox54.Text.ToString,
                MAIN.TextBox58.Text.ToString, MAIN.TextBox62.Text.ToString, MAIN.TextBox66.Text.ToString, MAIN.TextBox69.Text.ToString, MAIN.TextBox74.Text.ToString, MAIN.TextBox162.Text.ToString, MAIN.TextBox82.Text.ToString, MAIN.TextBox86.Text.ToString, MAIN.TextBox90.Text.ToString, MAIN.TextBox95.Text.ToString, MAIN.TextBox99.Text.ToString, MAIN.TextBox103.Text.ToString, MAIN.TextBox107.Text.ToString,
                MAIN.TextBox111.Text.ToString, MAIN.TextBox167.Text.ToString, MAIN.TextBox171.Text.ToString, MAIN.TextBox115.Text.ToString, MAIN.TextBox119.Text.ToString, MAIN.TextBox123.Text.ToString, MAIN.TextBox126.Text.ToString, MAIN.TextBox135.Text.ToString, MAIN.TextBox144.Text.ToString, MAIN.Button43.Text.ToString, DateTime.Now.ToString, MAIN.TextBox5.Text.ToString, MAIN.Label10.Text.ToString)
                SMC.ClearSequence()
                Me.Close()
                TrckRecord()
            ElseIf TextBox2.Text = "INIT_POWER" And RadioButton3.Checked = True Then
                SMC.CalculateCount("FAIL")
                ClearReportBeforeGeneration()
                Dim lastItem As String = MAIN.ListBox1.Items(MAIN.ListBox1.Items.Count - 1).ToString()
                SMC.GenerateReportonErrorSMCPowerTestNode0(MAIN.TextBox3.Text, MAIN.TextBox2.Text, MAIN.TextBox4.Text, MAIN.Button43.Text, MAIN.TextBox5.Text, MAIN.Label10.Text, MAIN.TextBox153.Text, lastItem, "TRUE", "TRUE", MAIN.TextBox27.Text, MAIN.TextBox37.Text, MAIN.TextBox43.Text)
                SMC.SMCCreateRecord(CONFIG.TextBox1.Text, CONFIG.TextBox2.Text, CONFIG.TextBox3.Text, MAIN.TextBox4.Text.ToString, MAIN.TextBox2.Text.ToString, MAIN.TextBox3.Text.ToString, MAIN.TextBox6.Text.ToString, MAIN.TextBox153.Text.ToString, MAIN.TextBox27.Text.ToString, MAIN.TextBox37.Text.ToString, MAIN.TextBox43.Text.ToString, MAIN.TextBox40.Text.ToString, MAIN.TextBox54.Text.ToString,
                MAIN.TextBox58.Text.ToString, MAIN.TextBox62.Text.ToString, MAIN.TextBox66.Text.ToString, MAIN.TextBox69.Text.ToString, MAIN.TextBox74.Text.ToString, MAIN.TextBox162.Text.ToString, MAIN.TextBox82.Text.ToString, MAIN.TextBox86.Text.ToString, MAIN.TextBox90.Text.ToString, MAIN.TextBox95.Text.ToString, MAIN.TextBox99.Text.ToString, MAIN.TextBox103.Text.ToString, MAIN.TextBox107.Text.ToString,
                MAIN.TextBox111.Text.ToString, MAIN.TextBox167.Text.ToString, MAIN.TextBox171.Text.ToString, MAIN.TextBox115.Text.ToString, MAIN.TextBox119.Text.ToString, MAIN.TextBox123.Text.ToString, MAIN.TextBox126.Text.ToString, MAIN.TextBox135.Text.ToString, MAIN.TextBox144.Text.ToString, MAIN.Button43.Text.ToString, DateTime.Now.ToString, MAIN.TextBox5.Text.ToString, MAIN.Label10.Text.ToString)
                SMC.ClearSequence()
                Me.Close()
                TrckRecord()
            ElseIf TextBox2.Text = "CAN670_5V" And RadioButton2.Checked = True Then
                SMC.CalculateCount("FAIL")
                ClearReportBeforeGeneration()
                Dim lastItem As String = MAIN.ListBox1.Items(MAIN.ListBox1.Items.Count - 1).ToString()
                GenerateReportOnErrorSMCPowerTestNode1(MAIN.TextBox3.Text, MAIN.TextBox2.Text, MAIN.TextBox4.Text, MAIN.Button43.Text, MAIN.TextBox5.Text, MAIN.Label10.Text, MAIN.TextBox153.Text, lastItem, "TRUE", "TRUE", MAIN.TextBox27.Text, MAIN.TextBox37.Text, MAIN.TextBox43.Text, MAIN.TextBox40.Text)
                SMC.SMCCreateRecord(CONFIG.TextBox1.Text, CONFIG.TextBox2.Text, CONFIG.TextBox3.Text, MAIN.TextBox4.Text.ToString, MAIN.TextBox2.Text.ToString, MAIN.TextBox3.Text.ToString, MAIN.TextBox6.Text.ToString, MAIN.TextBox153.Text.ToString, MAIN.TextBox27.Text.ToString, MAIN.TextBox37.Text.ToString, MAIN.TextBox43.Text.ToString, MAIN.TextBox40.Text.ToString, MAIN.TextBox54.Text.ToString,
                MAIN.TextBox58.Text.ToString, MAIN.TextBox62.Text.ToString, MAIN.TextBox66.Text.ToString, MAIN.TextBox69.Text.ToString, MAIN.TextBox74.Text.ToString, MAIN.TextBox162.Text.ToString, MAIN.TextBox82.Text.ToString, MAIN.TextBox86.Text.ToString, MAIN.TextBox90.Text.ToString, MAIN.TextBox95.Text.ToString, MAIN.TextBox99.Text.ToString, MAIN.TextBox103.Text.ToString, MAIN.TextBox107.Text.ToString,
                MAIN.TextBox111.Text.ToString, MAIN.TextBox167.Text.ToString, MAIN.TextBox171.Text.ToString, MAIN.TextBox115.Text.ToString, MAIN.TextBox119.Text.ToString, MAIN.TextBox123.Text.ToString, MAIN.TextBox126.Text.ToString, MAIN.TextBox135.Text.ToString, MAIN.TextBox144.Text.ToString, MAIN.Button43.Text.ToString, DateTime.Now.ToString, MAIN.TextBox5.Text.ToString, MAIN.Label10.Text.ToString)
                SMC.ClearSequence()
                Me.Close()
                TrckRecord()
            ElseIf TextBox2.Text = "CAN670_5V" And RadioButton3.Checked = True Then
                SMC.CalculateCount("FAIL")
                ClearReportBeforeGeneration()
                Dim lastItem As String = MAIN.ListBox1.Items(MAIN.ListBox1.Items.Count - 1).ToString()
                GenerateReportOnErrorSMCPowerTestNode1(MAIN.TextBox3.Text, MAIN.TextBox2.Text, MAIN.TextBox4.Text, MAIN.Button43.Text, MAIN.TextBox5.Text, MAIN.Label10.Text, MAIN.TextBox153.Text, lastItem, "TRUE", "TRUE", MAIN.TextBox27.Text, MAIN.TextBox37.Text, MAIN.TextBox43.Text, MAIN.TextBox40.Text)
                SMC.SMCCreateRecord(CONFIG.TextBox1.Text, CONFIG.TextBox2.Text, CONFIG.TextBox3.Text, MAIN.TextBox4.Text.ToString, MAIN.TextBox2.Text.ToString, MAIN.TextBox3.Text.ToString, MAIN.TextBox6.Text.ToString, MAIN.TextBox153.Text.ToString, MAIN.TextBox27.Text.ToString, MAIN.TextBox37.Text.ToString, MAIN.TextBox43.Text.ToString, MAIN.TextBox40.Text.ToString, MAIN.TextBox54.Text.ToString,
                MAIN.TextBox58.Text.ToString, MAIN.TextBox62.Text.ToString, MAIN.TextBox66.Text.ToString, MAIN.TextBox69.Text.ToString, MAIN.TextBox74.Text.ToString, MAIN.TextBox162.Text.ToString, MAIN.TextBox82.Text.ToString, MAIN.TextBox86.Text.ToString, MAIN.TextBox90.Text.ToString, MAIN.TextBox95.Text.ToString, MAIN.TextBox99.Text.ToString, MAIN.TextBox103.Text.ToString, MAIN.TextBox107.Text.ToString,
                MAIN.TextBox111.Text.ToString, MAIN.TextBox167.Text.ToString, MAIN.TextBox171.Text.ToString, MAIN.TextBox115.Text.ToString, MAIN.TextBox119.Text.ToString, MAIN.TextBox123.Text.ToString, MAIN.TextBox126.Text.ToString, MAIN.TextBox135.Text.ToString, MAIN.TextBox144.Text.ToString, MAIN.Button43.Text.ToString, DateTime.Now.ToString, MAIN.TextBox5.Text.ToString, MAIN.Label10.Text.ToString)
                SMC.ClearSequence()
                Me.Close()
                TrckRecord()
            ElseIf TextBox2.Text = "ISOLATED_V+12" And RadioButton2.Checked = True Then
                SMC.CalculateCount("FAIL")
                ClearReportBeforeGeneration()
                Dim lastItem As String = MAIN.ListBox1.Items(MAIN.ListBox1.Items.Count - 1).ToString()
                SMC.GenerateReportOnErrorSMCPowerTestNode2(MAIN.TextBox3.Text, MAIN.TextBox2.Text, MAIN.TextBox4.Text, MAIN.Button43.Text, MAIN.TextBox5.Text, MAIN.Label10.Text, MAIN.TextBox153.Text, lastItem, "TRUE", "TRUE", MAIN.TextBox27.Text, MAIN.TextBox37.Text, MAIN.TextBox43.Text, MAIN.TextBox40.Text, MAIN.TextBox54.Text)
                SMC.SMCCreateRecord(CONFIG.TextBox1.Text, CONFIG.TextBox2.Text, CONFIG.TextBox3.Text, MAIN.TextBox4.Text.ToString, MAIN.TextBox2.Text.ToString, MAIN.TextBox3.Text.ToString, MAIN.TextBox6.Text.ToString, MAIN.TextBox153.Text.ToString, MAIN.TextBox27.Text.ToString, MAIN.TextBox37.Text.ToString, MAIN.TextBox43.Text.ToString, MAIN.TextBox40.Text.ToString, MAIN.TextBox54.Text.ToString,
                MAIN.TextBox58.Text.ToString, MAIN.TextBox62.Text.ToString, MAIN.TextBox66.Text.ToString, MAIN.TextBox69.Text.ToString, MAIN.TextBox74.Text.ToString, MAIN.TextBox162.Text.ToString, MAIN.TextBox82.Text.ToString, MAIN.TextBox86.Text.ToString, MAIN.TextBox90.Text.ToString, MAIN.TextBox95.Text.ToString, MAIN.TextBox99.Text.ToString, MAIN.TextBox103.Text.ToString, MAIN.TextBox107.Text.ToString,
                MAIN.TextBox111.Text.ToString, MAIN.TextBox167.Text.ToString, MAIN.TextBox171.Text.ToString, MAIN.TextBox115.Text.ToString, MAIN.TextBox119.Text.ToString, MAIN.TextBox123.Text.ToString, MAIN.TextBox126.Text.ToString, MAIN.TextBox135.Text.ToString, MAIN.TextBox144.Text.ToString, MAIN.Button43.Text.ToString, DateTime.Now.ToString, MAIN.TextBox5.Text.ToString, MAIN.Label10.Text.ToString)
                SMC.ClearSequence()
                Me.Close()
                TrckRecord()
            ElseIf TextBox2.Text = "ISOLATED_V+12" And RadioButton3.Checked = True Then
                SMC.CalculateCount("FAIL")
                ClearReportBeforeGeneration()
                Dim lastItem As String = MAIN.ListBox1.Items(MAIN.ListBox1.Items.Count - 1).ToString()
                SMC.GenerateReportOnErrorSMCPowerTestNode2(MAIN.TextBox3.Text, MAIN.TextBox2.Text, MAIN.TextBox4.Text, MAIN.Button43.Text, MAIN.TextBox5.Text, MAIN.Label10.Text, MAIN.TextBox153.Text, lastItem, "TRUE", "TRUE", MAIN.TextBox27.Text, MAIN.TextBox37.Text, MAIN.TextBox43.Text, MAIN.TextBox40.Text, MAIN.TextBox54.Text)
                SMC.SMCCreateRecord(CONFIG.TextBox1.Text, CONFIG.TextBox2.Text, CONFIG.TextBox3.Text, MAIN.TextBox4.Text.ToString, MAIN.TextBox2.Text.ToString, MAIN.TextBox3.Text.ToString, MAIN.TextBox6.Text.ToString, MAIN.TextBox153.Text.ToString, MAIN.TextBox27.Text.ToString, MAIN.TextBox37.Text.ToString, MAIN.TextBox43.Text.ToString, MAIN.TextBox40.Text.ToString, MAIN.TextBox54.Text.ToString,
                MAIN.TextBox58.Text.ToString, MAIN.TextBox62.Text.ToString, MAIN.TextBox66.Text.ToString, MAIN.TextBox69.Text.ToString, MAIN.TextBox74.Text.ToString, MAIN.TextBox162.Text.ToString, MAIN.TextBox82.Text.ToString, MAIN.TextBox86.Text.ToString, MAIN.TextBox90.Text.ToString, MAIN.TextBox95.Text.ToString, MAIN.TextBox99.Text.ToString, MAIN.TextBox103.Text.ToString, MAIN.TextBox107.Text.ToString,
                MAIN.TextBox111.Text.ToString, MAIN.TextBox167.Text.ToString, MAIN.TextBox171.Text.ToString, MAIN.TextBox115.Text.ToString, MAIN.TextBox119.Text.ToString, MAIN.TextBox123.Text.ToString, MAIN.TextBox126.Text.ToString, MAIN.TextBox135.Text.ToString, MAIN.TextBox144.Text.ToString, MAIN.Button43.Text.ToString, DateTime.Now.ToString, MAIN.TextBox5.Text.ToString, MAIN.Label10.Text.ToString)
                SMC.ClearSequence()
                Me.Close()
                TrckRecord()
            ElseIf TextBox2.Text = "SENSOR_V+12" And RadioButton2.Checked = True Then
                SMC.CalculateCount("FAIL")
                ClearReportBeforeGeneration()
                Dim lastItem As String = MAIN.ListBox1.Items(MAIN.ListBox1.Items.Count - 1).ToString()
                SMC.GenerateReportOnErrorSMCPowerTestNode3(MAIN.TextBox3.Text, MAIN.TextBox2.Text, MAIN.TextBox4.Text, MAIN.Button43.Text, MAIN.TextBox5.Text, MAIN.Label10.Text, MAIN.TextBox153.Text, lastItem, "TRUE", "TRUE", MAIN.TextBox27.Text, MAIN.TextBox37.Text, MAIN.TextBox43.Text, MAIN.TextBox40.Text, MAIN.TextBox54.Text, MAIN.TextBox58.Text)
                SMC.SMCCreateRecord(CONFIG.TextBox1.Text, CONFIG.TextBox2.Text, CONFIG.TextBox3.Text, MAIN.TextBox4.Text.ToString, MAIN.TextBox2.Text.ToString, MAIN.TextBox3.Text.ToString, MAIN.TextBox6.Text.ToString, MAIN.TextBox153.Text.ToString, MAIN.TextBox27.Text.ToString, MAIN.TextBox37.Text.ToString, MAIN.TextBox43.Text.ToString, MAIN.TextBox40.Text.ToString, MAIN.TextBox54.Text.ToString,
                MAIN.TextBox58.Text.ToString, MAIN.TextBox62.Text.ToString, MAIN.TextBox66.Text.ToString, MAIN.TextBox69.Text.ToString, MAIN.TextBox74.Text.ToString, MAIN.TextBox162.Text.ToString, MAIN.TextBox82.Text.ToString, MAIN.TextBox86.Text.ToString, MAIN.TextBox90.Text.ToString, MAIN.TextBox95.Text.ToString, MAIN.TextBox99.Text.ToString, MAIN.TextBox103.Text.ToString, MAIN.TextBox107.Text.ToString,
                MAIN.TextBox111.Text.ToString, MAIN.TextBox167.Text.ToString, MAIN.TextBox171.Text.ToString, MAIN.TextBox115.Text.ToString, MAIN.TextBox119.Text.ToString, MAIN.TextBox123.Text.ToString, MAIN.TextBox126.Text.ToString, MAIN.TextBox135.Text.ToString, MAIN.TextBox144.Text.ToString, MAIN.Button43.Text.ToString, DateTime.Now.ToString, MAIN.TextBox5.Text.ToString, MAIN.Label10.Text.ToString)
                SMC.ClearSequence()
                Me.Close()
                TrckRecord()
            ElseIf TextBox2.Text = "SENSOR_V+12" And RadioButton3.Checked = True Then
                SMC.CalculateCount("FAIL")
                ClearReportBeforeGeneration()
                Dim lastItem As String = MAIN.ListBox1.Items(MAIN.ListBox1.Items.Count - 1).ToString()
                SMC.GenerateReportOnErrorSMCPowerTestNode3(MAIN.TextBox3.Text, MAIN.TextBox2.Text, MAIN.TextBox4.Text, MAIN.Button43.Text, MAIN.TextBox5.Text, MAIN.Label10.Text, MAIN.TextBox153.Text, lastItem, "TRUE", "TRUE", MAIN.TextBox27.Text, MAIN.TextBox37.Text, MAIN.TextBox43.Text, MAIN.TextBox40.Text, MAIN.TextBox54.Text, MAIN.TextBox58.Text)
                SMC.SMCCreateRecord(CONFIG.TextBox1.Text, CONFIG.TextBox2.Text, CONFIG.TextBox3.Text, MAIN.TextBox4.Text.ToString, MAIN.TextBox2.Text.ToString, MAIN.TextBox3.Text.ToString, MAIN.TextBox6.Text.ToString, MAIN.TextBox153.Text.ToString, MAIN.TextBox27.Text.ToString, MAIN.TextBox37.Text.ToString, MAIN.TextBox43.Text.ToString, MAIN.TextBox40.Text.ToString, MAIN.TextBox54.Text.ToString,
                MAIN.TextBox58.Text.ToString, MAIN.TextBox62.Text.ToString, MAIN.TextBox66.Text.ToString, MAIN.TextBox69.Text.ToString, MAIN.TextBox74.Text.ToString, MAIN.TextBox162.Text.ToString, MAIN.TextBox82.Text.ToString, MAIN.TextBox86.Text.ToString, MAIN.TextBox90.Text.ToString, MAIN.TextBox95.Text.ToString, MAIN.TextBox99.Text.ToString, MAIN.TextBox103.Text.ToString, MAIN.TextBox107.Text.ToString,
                MAIN.TextBox111.Text.ToString, MAIN.TextBox167.Text.ToString, MAIN.TextBox171.Text.ToString, MAIN.TextBox115.Text.ToString, MAIN.TextBox119.Text.ToString, MAIN.TextBox123.Text.ToString, MAIN.TextBox126.Text.ToString, MAIN.TextBox135.Text.ToString, MAIN.TextBox144.Text.ToString, MAIN.Button43.Text.ToString, DateTime.Now.ToString, MAIN.TextBox5.Text.ToString, MAIN.Label10.Text.ToString)
                SMC.ClearSequence()
                Me.Close()
                TrckRecord()
            ElseIf TextBox2.Text = "TT620_5V" And RadioButton2.Checked = True Then
                SMC.CalculateCount("FAIL")
                ClearReportBeforeGeneration()
                Dim lastItem As String = MAIN.ListBox1.Items(MAIN.ListBox1.Items.Count - 1).ToString()
                SMC.GenerateReportOnErrorSMCPowerTestNode4(MAIN.TextBox3.Text, MAIN.TextBox2.Text, MAIN.TextBox4.Text, MAIN.Button43.Text, MAIN.TextBox5.Text, MAIN.Label10.Text, MAIN.TextBox153.Text, lastItem, "TRUE", "TRUE", MAIN.TextBox27.Text, MAIN.TextBox37.Text, MAIN.TextBox43.Text, MAIN.TextBox40.Text, MAIN.TextBox54.Text, MAIN.TextBox58.Text, MAIN.TextBox62.Text)
                SMC.SMCCreateRecord(CONFIG.TextBox1.Text, CONFIG.TextBox2.Text, CONFIG.TextBox3.Text, MAIN.TextBox4.Text.ToString, MAIN.TextBox2.Text.ToString, MAIN.TextBox3.Text.ToString, MAIN.TextBox6.Text.ToString, MAIN.TextBox153.Text.ToString, MAIN.TextBox27.Text.ToString, MAIN.TextBox37.Text.ToString, MAIN.TextBox43.Text.ToString, MAIN.TextBox40.Text.ToString, MAIN.TextBox54.Text.ToString,
                MAIN.TextBox58.Text.ToString, MAIN.TextBox62.Text.ToString, MAIN.TextBox66.Text.ToString, MAIN.TextBox69.Text.ToString, MAIN.TextBox74.Text.ToString, MAIN.TextBox162.Text.ToString, MAIN.TextBox82.Text.ToString, MAIN.TextBox86.Text.ToString, MAIN.TextBox90.Text.ToString, MAIN.TextBox95.Text.ToString, MAIN.TextBox99.Text.ToString, MAIN.TextBox103.Text.ToString, MAIN.TextBox107.Text.ToString,
                MAIN.TextBox111.Text.ToString, MAIN.TextBox167.Text.ToString, MAIN.TextBox171.Text.ToString, MAIN.TextBox115.Text.ToString, MAIN.TextBox119.Text.ToString, MAIN.TextBox123.Text.ToString, MAIN.TextBox126.Text.ToString, MAIN.TextBox135.Text.ToString, MAIN.TextBox144.Text.ToString, MAIN.Button43.Text.ToString, DateTime.Now.ToString, MAIN.TextBox5.Text.ToString, MAIN.Label10.Text.ToString)
                SMC.ClearSequence()
                Me.Close()
                TrckRecord()
            ElseIf TextBox2.Text = "TT620_5V" And RadioButton3.Checked = True Then
                SMC.CalculateCount("FAIL")
                ClearReportBeforeGeneration()
                Dim lastItem As String = MAIN.ListBox1.Items(MAIN.ListBox1.Items.Count - 1).ToString()
                SMC.GenerateReportOnErrorSMCPowerTestNode4(MAIN.TextBox3.Text, MAIN.TextBox2.Text, MAIN.TextBox4.Text, MAIN.Button43.Text, MAIN.TextBox5.Text, MAIN.Label10.Text, MAIN.TextBox153.Text, lastItem, "TRUE", "TRUE", MAIN.TextBox27.Text, MAIN.TextBox37.Text, MAIN.TextBox43.Text, MAIN.TextBox40.Text, MAIN.TextBox54.Text, MAIN.TextBox58.Text, MAIN.TextBox62.Text)
                SMC.SMCCreateRecord(CONFIG.TextBox1.Text, CONFIG.TextBox2.Text, CONFIG.TextBox3.Text, MAIN.TextBox4.Text.ToString, MAIN.TextBox2.Text.ToString, MAIN.TextBox3.Text.ToString, MAIN.TextBox6.Text.ToString, MAIN.TextBox153.Text.ToString, MAIN.TextBox27.Text.ToString, MAIN.TextBox37.Text.ToString, MAIN.TextBox43.Text.ToString, MAIN.TextBox40.Text.ToString, MAIN.TextBox54.Text.ToString,
                MAIN.TextBox58.Text.ToString, MAIN.TextBox62.Text.ToString, MAIN.TextBox66.Text.ToString, MAIN.TextBox69.Text.ToString, MAIN.TextBox74.Text.ToString, MAIN.TextBox162.Text.ToString, MAIN.TextBox82.Text.ToString, MAIN.TextBox86.Text.ToString, MAIN.TextBox90.Text.ToString, MAIN.TextBox95.Text.ToString, MAIN.TextBox99.Text.ToString, MAIN.TextBox103.Text.ToString, MAIN.TextBox107.Text.ToString,
                MAIN.TextBox111.Text.ToString, MAIN.TextBox167.Text.ToString, MAIN.TextBox171.Text.ToString, MAIN.TextBox115.Text.ToString, MAIN.TextBox119.Text.ToString, MAIN.TextBox123.Text.ToString, MAIN.TextBox126.Text.ToString, MAIN.TextBox135.Text.ToString, MAIN.TextBox144.Text.ToString, MAIN.Button43.Text.ToString, DateTime.Now.ToString, MAIN.TextBox5.Text.ToString, MAIN.Label10.Text.ToString)
                SMC.ClearSequence()
                Me.Close()
                TrckRecord()
            ElseIf TextBox2.Text = "IT661_5V" And RadioButton2.Checked = True Then
                SMC.CalculateCount("FAIL")
                ClearReportBeforeGeneration()
                Dim lastItem As String = MAIN.ListBox1.Items(MAIN.ListBox1.Items.Count - 1).ToString()
                SMC.GenerateReportOnErrorSMCPowerTestNode5(MAIN.TextBox3.Text, MAIN.TextBox2.Text, MAIN.TextBox4.Text, MAIN.Button43.Text, MAIN.TextBox5.Text, MAIN.Label10.Text, MAIN.TextBox153.Text, lastItem, "TRUE", "TRUE", MAIN.TextBox27.Text, MAIN.TextBox37.Text, MAIN.TextBox43.Text, MAIN.TextBox40.Text, MAIN.TextBox54.Text, MAIN.TextBox58.Text, MAIN.TextBox62.Text, MAIN.TextBox66.Text)
                SMC.SMCCreateRecord(CONFIG.TextBox1.Text, CONFIG.TextBox2.Text, CONFIG.TextBox3.Text, MAIN.TextBox4.Text.ToString, MAIN.TextBox2.Text.ToString, MAIN.TextBox3.Text.ToString, MAIN.TextBox6.Text.ToString, MAIN.TextBox153.Text.ToString, MAIN.TextBox27.Text.ToString, MAIN.TextBox37.Text.ToString, MAIN.TextBox43.Text.ToString, MAIN.TextBox40.Text.ToString, MAIN.TextBox54.Text.ToString,
                MAIN.TextBox58.Text.ToString, MAIN.TextBox62.Text.ToString, MAIN.TextBox66.Text.ToString, MAIN.TextBox69.Text.ToString, MAIN.TextBox74.Text.ToString, MAIN.TextBox162.Text.ToString, MAIN.TextBox82.Text.ToString, MAIN.TextBox86.Text.ToString, MAIN.TextBox90.Text.ToString, MAIN.TextBox95.Text.ToString, MAIN.TextBox99.Text.ToString, MAIN.TextBox103.Text.ToString, MAIN.TextBox107.Text.ToString,
                MAIN.TextBox111.Text.ToString, MAIN.TextBox167.Text.ToString, MAIN.TextBox171.Text.ToString, MAIN.TextBox115.Text.ToString, MAIN.TextBox119.Text.ToString, MAIN.TextBox123.Text.ToString, MAIN.TextBox126.Text.ToString, MAIN.TextBox135.Text.ToString, MAIN.TextBox144.Text.ToString, MAIN.Button43.Text.ToString, DateTime.Now.ToString, MAIN.TextBox5.Text.ToString, MAIN.Label10.Text.ToString)
                SMC.ClearSequence()
                Me.Close()
                TrckRecord()
            ElseIf TextBox2.Text = "IT661_5V" And RadioButton3.Checked = True Then
                SMC.CalculateCount("FAIL")
                ClearReportBeforeGeneration()
                Dim lastItem As String = MAIN.ListBox1.Items(MAIN.ListBox1.Items.Count - 1).ToString()
                SMC.GenerateReportOnErrorSMCPowerTestNode5(MAIN.TextBox3.Text, MAIN.TextBox2.Text, MAIN.TextBox4.Text, MAIN.Button43.Text, MAIN.TextBox5.Text, MAIN.Label10.Text, MAIN.TextBox153.Text, lastItem, "TRUE", "TRUE", MAIN.TextBox27.Text, MAIN.TextBox37.Text, MAIN.TextBox43.Text, MAIN.TextBox40.Text, MAIN.TextBox54.Text, MAIN.TextBox58.Text, MAIN.TextBox62.Text, MAIN.TextBox66.Text)
                SMC.SMCCreateRecord(CONFIG.TextBox1.Text, CONFIG.TextBox2.Text, CONFIG.TextBox3.Text, MAIN.TextBox4.Text.ToString, MAIN.TextBox2.Text.ToString, MAIN.TextBox3.Text.ToString, MAIN.TextBox6.Text.ToString, MAIN.TextBox153.Text.ToString, MAIN.TextBox27.Text.ToString, MAIN.TextBox37.Text.ToString, MAIN.TextBox43.Text.ToString, MAIN.TextBox40.Text.ToString, MAIN.TextBox54.Text.ToString,
                MAIN.TextBox58.Text.ToString, MAIN.TextBox62.Text.ToString, MAIN.TextBox66.Text.ToString, MAIN.TextBox69.Text.ToString, MAIN.TextBox74.Text.ToString, MAIN.TextBox162.Text.ToString, MAIN.TextBox82.Text.ToString, MAIN.TextBox86.Text.ToString, MAIN.TextBox90.Text.ToString, MAIN.TextBox95.Text.ToString, MAIN.TextBox99.Text.ToString, MAIN.TextBox103.Text.ToString, MAIN.TextBox107.Text.ToString,
                MAIN.TextBox111.Text.ToString, MAIN.TextBox167.Text.ToString, MAIN.TextBox171.Text.ToString, MAIN.TextBox115.Text.ToString, MAIN.TextBox119.Text.ToString, MAIN.TextBox123.Text.ToString, MAIN.TextBox126.Text.ToString, MAIN.TextBox135.Text.ToString, MAIN.TextBox144.Text.ToString, MAIN.Button43.Text.ToString, DateTime.Now.ToString, MAIN.TextBox5.Text.ToString, MAIN.Label10.Text.ToString)
                SMC.ClearSequence()
                Me.Close()
                TrckRecord()
            ElseIf TextBox2.Text = "A0691" And RadioButton2.Checked = True Then
                SMC.CalculateCount("FAIL")
                ClearReportBeforeGeneration()
                Dim lastItem As String = MAIN.ListBox1.Items(MAIN.ListBox1.Items.Count - 1).ToString()
                SMC.GenerateReportOnErrorSMCPowerTestNode6(MAIN.TextBox3.Text, MAIN.TextBox2.Text, MAIN.TextBox4.Text, MAIN.Button43.Text, MAIN.TextBox5.Text, MAIN.Label10.Text, MAIN.TextBox153.Text, lastItem, "TRUE", "TRUE", MAIN.TextBox27.Text, MAIN.TextBox37.Text, MAIN.TextBox43.Text, MAIN.TextBox40.Text, MAIN.TextBox54.Text, MAIN.TextBox58.Text, MAIN.TextBox62.Text, MAIN.TextBox66.Text, MAIN.TextBox69.Text)
                SMC.SMCCreateRecord(CONFIG.TextBox1.Text, CONFIG.TextBox2.Text, CONFIG.TextBox3.Text, MAIN.TextBox4.Text.ToString, MAIN.TextBox2.Text.ToString, MAIN.TextBox3.Text.ToString, MAIN.TextBox6.Text.ToString, MAIN.TextBox153.Text.ToString, MAIN.TextBox27.Text.ToString, MAIN.TextBox37.Text.ToString, MAIN.TextBox43.Text.ToString, MAIN.TextBox40.Text.ToString, MAIN.TextBox54.Text.ToString,
                MAIN.TextBox58.Text.ToString, MAIN.TextBox62.Text.ToString, MAIN.TextBox66.Text.ToString, MAIN.TextBox69.Text.ToString, MAIN.TextBox74.Text.ToString, MAIN.TextBox162.Text.ToString, MAIN.TextBox82.Text.ToString, MAIN.TextBox86.Text.ToString, MAIN.TextBox90.Text.ToString, MAIN.TextBox95.Text.ToString, MAIN.TextBox99.Text.ToString, MAIN.TextBox103.Text.ToString, MAIN.TextBox107.Text.ToString,
                MAIN.TextBox111.Text.ToString, MAIN.TextBox167.Text.ToString, MAIN.TextBox171.Text.ToString, MAIN.TextBox115.Text.ToString, MAIN.TextBox119.Text.ToString, MAIN.TextBox123.Text.ToString, MAIN.TextBox126.Text.ToString, MAIN.TextBox135.Text.ToString, MAIN.TextBox144.Text.ToString, MAIN.Button43.Text.ToString, DateTime.Now.ToString, MAIN.TextBox5.Text.ToString, MAIN.Label10.Text.ToString)
                SMC.ClearSequence()
                Me.Close()
                TrckRecord()
            ElseIf TextBox2.Text = "A0691" And RadioButton3.Checked = True Then
                SMC.CalculateCount("FAIL")
                ClearReportBeforeGeneration()
                Dim lastItem As String = MAIN.ListBox1.Items(MAIN.ListBox1.Items.Count - 1).ToString()
                SMC.GenerateReportOnErrorSMCPowerTestNode6(MAIN.TextBox3.Text, MAIN.TextBox2.Text, MAIN.TextBox4.Text, MAIN.Button43.Text, MAIN.TextBox5.Text, MAIN.Label10.Text, MAIN.TextBox153.Text, lastItem, "TRUE", "TRUE", MAIN.TextBox27.Text, MAIN.TextBox37.Text, MAIN.TextBox43.Text, MAIN.TextBox40.Text, MAIN.TextBox54.Text, MAIN.TextBox58.Text, MAIN.TextBox62.Text, MAIN.TextBox66.Text, MAIN.TextBox69.Text)
                SMC.SMCCreateRecord(CONFIG.TextBox1.Text, CONFIG.TextBox2.Text, CONFIG.TextBox3.Text, MAIN.TextBox4.Text.ToString, MAIN.TextBox2.Text.ToString, MAIN.TextBox3.Text.ToString, MAIN.TextBox6.Text.ToString, MAIN.TextBox153.Text.ToString, MAIN.TextBox27.Text.ToString, MAIN.TextBox37.Text.ToString, MAIN.TextBox43.Text.ToString, MAIN.TextBox40.Text.ToString, MAIN.TextBox54.Text.ToString,
                MAIN.TextBox58.Text.ToString, MAIN.TextBox62.Text.ToString, MAIN.TextBox66.Text.ToString, MAIN.TextBox69.Text.ToString, MAIN.TextBox74.Text.ToString, MAIN.TextBox162.Text.ToString, MAIN.TextBox82.Text.ToString, MAIN.TextBox86.Text.ToString, MAIN.TextBox90.Text.ToString, MAIN.TextBox95.Text.ToString, MAIN.TextBox99.Text.ToString, MAIN.TextBox103.Text.ToString, MAIN.TextBox107.Text.ToString,
                MAIN.TextBox111.Text.ToString, MAIN.TextBox167.Text.ToString, MAIN.TextBox171.Text.ToString, MAIN.TextBox115.Text.ToString, MAIN.TextBox119.Text.ToString, MAIN.TextBox123.Text.ToString, MAIN.TextBox126.Text.ToString, MAIN.TextBox135.Text.ToString, MAIN.TextBox144.Text.ToString, MAIN.Button43.Text.ToString, DateTime.Now.ToString, MAIN.TextBox5.Text.ToString, MAIN.Label10.Text.ToString)
                SMC.ClearSequence()
                Me.Close()
                TrckRecord()
            ElseIf TextBox2.Text = "D0691" And RadioButton2.Checked = True Then
                SMC.CalculateCount("FAIL")
                ClearReportBeforeGeneration()
                Dim lastItem As String = MAIN.ListBox1.Items(MAIN.ListBox1.Items.Count - 1).ToString()
                SMC.GenerateReportOnErrorSMCPowerTestNode7(MAIN.TextBox3.Text, MAIN.TextBox2.Text, MAIN.TextBox4.Text, MAIN.Button43.Text, MAIN.TextBox5.Text, MAIN.Label10.Text, MAIN.TextBox153.Text, lastItem, "TRUE", "TRUE", MAIN.TextBox27.Text, MAIN.TextBox37.Text, MAIN.TextBox43.Text, MAIN.TextBox40.Text, MAIN.TextBox54.Text, MAIN.TextBox58.Text, MAIN.TextBox62.Text, MAIN.TextBox66.Text, MAIN.TextBox69.Text, MAIN.TextBox74.Text)
                SMC.SMCCreateRecord(CONFIG.TextBox1.Text, CONFIG.TextBox2.Text, CONFIG.TextBox3.Text, MAIN.TextBox4.Text.ToString, MAIN.TextBox2.Text.ToString, MAIN.TextBox3.Text.ToString, MAIN.TextBox6.Text.ToString, MAIN.TextBox153.Text.ToString, MAIN.TextBox27.Text.ToString, MAIN.TextBox37.Text.ToString, MAIN.TextBox43.Text.ToString, MAIN.TextBox40.Text.ToString, MAIN.TextBox54.Text.ToString,
                MAIN.TextBox58.Text.ToString, MAIN.TextBox62.Text.ToString, MAIN.TextBox66.Text.ToString, MAIN.TextBox69.Text.ToString, MAIN.TextBox74.Text.ToString, MAIN.TextBox162.Text.ToString, MAIN.TextBox82.Text.ToString, MAIN.TextBox86.Text.ToString, MAIN.TextBox90.Text.ToString, MAIN.TextBox95.Text.ToString, MAIN.TextBox99.Text.ToString, MAIN.TextBox103.Text.ToString, MAIN.TextBox107.Text.ToString,
                MAIN.TextBox111.Text.ToString, MAIN.TextBox167.Text.ToString, MAIN.TextBox171.Text.ToString, MAIN.TextBox115.Text.ToString, MAIN.TextBox119.Text.ToString, MAIN.TextBox123.Text.ToString, MAIN.TextBox126.Text.ToString, MAIN.TextBox135.Text.ToString, MAIN.TextBox144.Text.ToString, MAIN.Button43.Text.ToString, DateTime.Now.ToString, MAIN.TextBox5.Text.ToString, MAIN.Label10.Text.ToString)
                SMC.ClearSequence()
                Me.Close()
                TrckRecord()
            ElseIf TextBox2.Text = "D0691" And RadioButton3.Checked = True Then
                ClearReportBeforeGeneration()
                Dim lastItem As String = MAIN.ListBox1.Items(MAIN.ListBox1.Items.Count - 1).ToString()
                SMC.GenerateReportOnErrorSMCPowerTestNode7(MAIN.TextBox3.Text, MAIN.TextBox2.Text, MAIN.TextBox4.Text, MAIN.Button43.Text, MAIN.TextBox5.Text, MAIN.Label10.Text, MAIN.TextBox153.Text, lastItem, "TRUE", "TRUE", MAIN.TextBox27.Text, MAIN.TextBox37.Text, MAIN.TextBox43.Text, MAIN.TextBox40.Text, MAIN.TextBox54.Text, MAIN.TextBox58.Text, MAIN.TextBox62.Text, MAIN.TextBox66.Text, MAIN.TextBox69.Text, MAIN.TextBox74.Text)
                SMC.SMCCreateRecord(CONFIG.TextBox1.Text, CONFIG.TextBox2.Text, CONFIG.TextBox3.Text, MAIN.TextBox4.Text.ToString, MAIN.TextBox2.Text.ToString, MAIN.TextBox3.Text.ToString, MAIN.TextBox6.Text.ToString, MAIN.TextBox153.Text.ToString, MAIN.TextBox27.Text.ToString, MAIN.TextBox37.Text.ToString, MAIN.TextBox43.Text.ToString, MAIN.TextBox40.Text.ToString, MAIN.TextBox54.Text.ToString,
                MAIN.TextBox58.Text.ToString, MAIN.TextBox62.Text.ToString, MAIN.TextBox66.Text.ToString, MAIN.TextBox69.Text.ToString, MAIN.TextBox74.Text.ToString, MAIN.TextBox162.Text.ToString, MAIN.TextBox82.Text.ToString, MAIN.TextBox86.Text.ToString, MAIN.TextBox90.Text.ToString, MAIN.TextBox95.Text.ToString, MAIN.TextBox99.Text.ToString, MAIN.TextBox103.Text.ToString, MAIN.TextBox107.Text.ToString,
                MAIN.TextBox111.Text.ToString, MAIN.TextBox167.Text.ToString, MAIN.TextBox171.Text.ToString, MAIN.TextBox115.Text.ToString, MAIN.TextBox119.Text.ToString, MAIN.TextBox123.Text.ToString, MAIN.TextBox126.Text.ToString, MAIN.TextBox135.Text.ToString, MAIN.TextBox144.Text.ToString, MAIN.Button43.Text.ToString, DateTime.Now.ToString, MAIN.TextBox5.Text.ToString, MAIN.Label10.Text.ToString)
                SMC.ClearSequence()
                Me.Close()
                TrckRecord()
            ElseIf TextBox2.Text = "D0691_ON" And RadioButton2.Checked = True Then
                SMC.CalculateCount("FAIL")
                ClearReportBeforeGeneration()
                Dim lastItem As String = MAIN.ListBox1.Items(MAIN.ListBox1.Items.Count - 1).ToString()
                SMC.GenerateReportOnErrorSMCSwitchTestNode0(MAIN.TextBox3.Text, MAIN.TextBox2.Text, MAIN.TextBox4.Text, MAIN.Button43.Text, MAIN.TextBox5.Text, MAIN.Label10.Text, MAIN.TextBox153.Text, lastItem, "TRUE", "TRUE", MAIN.TextBox27.Text, MAIN.TextBox37.Text, MAIN.TextBox43.Text, MAIN.TextBox40.Text, MAIN.TextBox54.Text, MAIN.TextBox58.Text, MAIN.TextBox62.Text, MAIN.TextBox66.Text, MAIN.TextBox69.Text, MAIN.TextBox74.Text, MAIN.TextBox162.Text)
                SMC.SMCCreateRecord(CONFIG.TextBox1.Text, CONFIG.TextBox2.Text, CONFIG.TextBox3.Text, MAIN.TextBox4.Text.ToString, MAIN.TextBox2.Text.ToString, MAIN.TextBox3.Text.ToString, MAIN.TextBox6.Text.ToString, MAIN.TextBox153.Text.ToString, MAIN.TextBox27.Text.ToString, MAIN.TextBox37.Text.ToString, MAIN.TextBox43.Text.ToString, MAIN.TextBox40.Text.ToString, MAIN.TextBox54.Text.ToString,
                MAIN.TextBox58.Text.ToString, MAIN.TextBox62.Text.ToString, MAIN.TextBox66.Text.ToString, MAIN.TextBox69.Text.ToString, MAIN.TextBox74.Text.ToString, MAIN.TextBox162.Text.ToString, MAIN.TextBox82.Text.ToString, MAIN.TextBox86.Text.ToString, MAIN.TextBox90.Text.ToString, MAIN.TextBox95.Text.ToString, MAIN.TextBox99.Text.ToString, MAIN.TextBox103.Text.ToString, MAIN.TextBox107.Text.ToString,
                MAIN.TextBox111.Text.ToString, MAIN.TextBox167.Text.ToString, MAIN.TextBox171.Text.ToString, MAIN.TextBox115.Text.ToString, MAIN.TextBox119.Text.ToString, MAIN.TextBox123.Text.ToString, MAIN.TextBox126.Text.ToString, MAIN.TextBox135.Text.ToString, MAIN.TextBox144.Text.ToString, MAIN.Button43.Text.ToString, DateTime.Now.ToString, MAIN.TextBox5.Text.ToString, MAIN.Label10.Text.ToString)
                SMC.ClearSequence()
                Me.Close()
                TrckRecord()
            ElseIf TextBox2.Text = "D0691_ON" And RadioButton3.Checked = True Then
                SMC.CalculateCount("FAIL")
                ClearReportBeforeGeneration()
                Dim lastItem As String = MAIN.ListBox1.Items(MAIN.ListBox1.Items.Count - 1).ToString()
                SMC.GenerateReportOnErrorSMCSwitchTestNode0(MAIN.TextBox3.Text, MAIN.TextBox2.Text, MAIN.TextBox4.Text, MAIN.Button43.Text, MAIN.TextBox5.Text, MAIN.Label10.Text, MAIN.TextBox153.Text, lastItem, "TRUE", "TRUE", MAIN.TextBox27.Text, MAIN.TextBox37.Text, MAIN.TextBox43.Text, MAIN.TextBox40.Text, MAIN.TextBox54.Text, MAIN.TextBox58.Text, MAIN.TextBox62.Text, MAIN.TextBox66.Text, MAIN.TextBox69.Text, MAIN.TextBox74.Text, MAIN.TextBox162.Text)
                SMC.SMCCreateRecord(CONFIG.TextBox1.Text, CONFIG.TextBox2.Text, CONFIG.TextBox3.Text, MAIN.TextBox4.Text.ToString, MAIN.TextBox2.Text.ToString, MAIN.TextBox3.Text.ToString, MAIN.TextBox6.Text.ToString, MAIN.TextBox153.Text.ToString, MAIN.TextBox27.Text.ToString, MAIN.TextBox37.Text.ToString, MAIN.TextBox43.Text.ToString, MAIN.TextBox40.Text.ToString, MAIN.TextBox54.Text.ToString,
                MAIN.TextBox58.Text.ToString, MAIN.TextBox62.Text.ToString, MAIN.TextBox66.Text.ToString, MAIN.TextBox69.Text.ToString, MAIN.TextBox74.Text.ToString, MAIN.TextBox162.Text.ToString, MAIN.TextBox82.Text.ToString, MAIN.TextBox86.Text.ToString, MAIN.TextBox90.Text.ToString, MAIN.TextBox95.Text.ToString, MAIN.TextBox99.Text.ToString, MAIN.TextBox103.Text.ToString, MAIN.TextBox107.Text.ToString,
                MAIN.TextBox111.Text.ToString, MAIN.TextBox167.Text.ToString, MAIN.TextBox171.Text.ToString, MAIN.TextBox115.Text.ToString, MAIN.TextBox119.Text.ToString, MAIN.TextBox123.Text.ToString, MAIN.TextBox126.Text.ToString, MAIN.TextBox135.Text.ToString, MAIN.TextBox144.Text.ToString, MAIN.Button43.Text.ToString, DateTime.Now.ToString, MAIN.TextBox5.Text.ToString, MAIN.Label10.Text.ToString)
                SMC.ClearSequence()
                Me.Close()
                TrckRecord()
            ElseIf TextBox2.Text = "D0661_ON" And RadioButton2.Checked = True Then
                SMC.CalculateCount("FAIL")
                ClearReportBeforeGeneration()
                Dim lastItem As String = MAIN.ListBox1.Items(MAIN.ListBox1.Items.Count - 1).ToString()
                SMC.GenerateReportOnErrorSMCSwitchTestNode1(MAIN.TextBox3.Text, MAIN.TextBox2.Text, MAIN.TextBox4.Text, MAIN.Button43.Text, MAIN.TextBox5.Text, MAIN.Label10.Text, MAIN.TextBox153.Text, lastItem, "TRUE", "TRUE", MAIN.TextBox27.Text, MAIN.TextBox37.Text, MAIN.TextBox43.Text, MAIN.TextBox40.Text, MAIN.TextBox54.Text, MAIN.TextBox58.Text, MAIN.TextBox62.Text, MAIN.TextBox66.Text, MAIN.TextBox69.Text, MAIN.TextBox74.Text, MAIN.TextBox162.Text, MAIN.TextBox82.Text)
                SMC.SMCCreateRecord(CONFIG.TextBox1.Text, CONFIG.TextBox2.Text, CONFIG.TextBox3.Text, MAIN.TextBox4.Text.ToString, MAIN.TextBox2.Text.ToString, MAIN.TextBox3.Text.ToString, MAIN.TextBox6.Text.ToString, MAIN.TextBox153.Text.ToString, MAIN.TextBox27.Text.ToString, MAIN.TextBox37.Text.ToString, MAIN.TextBox43.Text.ToString, MAIN.TextBox40.Text.ToString, MAIN.TextBox54.Text.ToString,
                MAIN.TextBox58.Text.ToString, MAIN.TextBox62.Text.ToString, MAIN.TextBox66.Text.ToString, MAIN.TextBox69.Text.ToString, MAIN.TextBox74.Text.ToString, MAIN.TextBox162.Text.ToString, MAIN.TextBox82.Text.ToString, MAIN.TextBox86.Text.ToString, MAIN.TextBox90.Text.ToString, MAIN.TextBox95.Text.ToString, MAIN.TextBox99.Text.ToString, MAIN.TextBox103.Text.ToString, MAIN.TextBox107.Text.ToString,
                MAIN.TextBox111.Text.ToString, MAIN.TextBox167.Text.ToString, MAIN.TextBox171.Text.ToString, MAIN.TextBox115.Text.ToString, MAIN.TextBox119.Text.ToString, MAIN.TextBox123.Text.ToString, MAIN.TextBox126.Text.ToString, MAIN.TextBox135.Text.ToString, MAIN.TextBox144.Text.ToString, MAIN.Button43.Text.ToString, DateTime.Now.ToString, MAIN.TextBox5.Text.ToString, MAIN.Label10.Text.ToString)
                SMC.ClearSequence()
                Me.Close()
                TrckRecord()
            ElseIf TextBox2.Text = "D0661_ON" And RadioButton3.Checked = True Then
                SMC.CalculateCount("FAIL")
                ClearReportBeforeGeneration()
                Dim lastItem As String = MAIN.ListBox1.Items(MAIN.ListBox1.Items.Count - 1).ToString()
                SMC.GenerateReportOnErrorSMCSwitchTestNode1(MAIN.TextBox3.Text, MAIN.TextBox2.Text, MAIN.TextBox4.Text, MAIN.Button43.Text, MAIN.TextBox5.Text, MAIN.Label10.Text, MAIN.TextBox153.Text, lastItem, "TRUE", "TRUE", MAIN.TextBox27.Text, MAIN.TextBox37.Text, MAIN.TextBox43.Text, MAIN.TextBox40.Text, MAIN.TextBox54.Text, MAIN.TextBox58.Text, MAIN.TextBox62.Text, MAIN.TextBox66.Text, MAIN.TextBox69.Text, MAIN.TextBox74.Text, MAIN.TextBox162.Text, MAIN.TextBox82.Text)
                SMC.SMCCreateRecord(CONFIG.TextBox1.Text, CONFIG.TextBox2.Text, CONFIG.TextBox3.Text, MAIN.TextBox4.Text.ToString, MAIN.TextBox2.Text.ToString, MAIN.TextBox3.Text.ToString, MAIN.TextBox6.Text.ToString, MAIN.TextBox153.Text.ToString, MAIN.TextBox27.Text.ToString, MAIN.TextBox37.Text.ToString, MAIN.TextBox43.Text.ToString, MAIN.TextBox40.Text.ToString, MAIN.TextBox54.Text.ToString,
                MAIN.TextBox58.Text.ToString, MAIN.TextBox62.Text.ToString, MAIN.TextBox66.Text.ToString, MAIN.TextBox69.Text.ToString, MAIN.TextBox74.Text.ToString, MAIN.TextBox162.Text.ToString, MAIN.TextBox82.Text.ToString, MAIN.TextBox86.Text.ToString, MAIN.TextBox90.Text.ToString, MAIN.TextBox95.Text.ToString, MAIN.TextBox99.Text.ToString, MAIN.TextBox103.Text.ToString, MAIN.TextBox107.Text.ToString,
                MAIN.TextBox111.Text.ToString, MAIN.TextBox167.Text.ToString, MAIN.TextBox171.Text.ToString, MAIN.TextBox115.Text.ToString, MAIN.TextBox119.Text.ToString, MAIN.TextBox123.Text.ToString, MAIN.TextBox126.Text.ToString, MAIN.TextBox135.Text.ToString, MAIN.TextBox144.Text.ToString, MAIN.Button43.Text.ToString, DateTime.Now.ToString, MAIN.TextBox5.Text.ToString, MAIN.Label10.Text.ToString)
                SMC.ClearSequence()
                Me.Close()
                TrckRecord()
            ElseIf TextBox2.Text = "D0661_OFF" And RadioButton2.Checked = True Then
                SMC.CalculateCount("FAIL")
                ClearReportBeforeGeneration()
                Dim lastItem As String = MAIN.ListBox1.Items(MAIN.ListBox1.Items.Count - 1).ToString()
                SMC.GenerateReportOnErrorSMCSwitchTestNode2(MAIN.TextBox3.Text, MAIN.TextBox2.Text, MAIN.TextBox4.Text, MAIN.Button43.Text, MAIN.TextBox5.Text, MAIN.Label10.Text, MAIN.TextBox153.Text, lastItem, "TRUE", "TRUE", MAIN.TextBox27.Text, MAIN.TextBox37.Text, MAIN.TextBox43.Text, MAIN.TextBox40.Text, MAIN.TextBox54.Text, MAIN.TextBox58.Text, MAIN.TextBox62.Text, MAIN.TextBox66.Text, MAIN.TextBox69.Text, MAIN.TextBox74.Text, MAIN.TextBox162.Text, MAIN.TextBox82.Text, MAIN.TextBox86.Text)
                SMC.SMCCreateRecord(CONFIG.TextBox1.Text, CONFIG.TextBox2.Text, CONFIG.TextBox3.Text, MAIN.TextBox4.Text.ToString, MAIN.TextBox2.Text.ToString, MAIN.TextBox3.Text.ToString, MAIN.TextBox6.Text.ToString, MAIN.TextBox153.Text.ToString, MAIN.TextBox27.Text.ToString, MAIN.TextBox37.Text.ToString, MAIN.TextBox43.Text.ToString, MAIN.TextBox40.Text.ToString, MAIN.TextBox54.Text.ToString,
                MAIN.TextBox58.Text.ToString, MAIN.TextBox62.Text.ToString, MAIN.TextBox66.Text.ToString, MAIN.TextBox69.Text.ToString, MAIN.TextBox74.Text.ToString, MAIN.TextBox162.Text.ToString, MAIN.TextBox82.Text.ToString, MAIN.TextBox86.Text.ToString, MAIN.TextBox90.Text.ToString, MAIN.TextBox95.Text.ToString, MAIN.TextBox99.Text.ToString, MAIN.TextBox103.Text.ToString, MAIN.TextBox107.Text.ToString,
                MAIN.TextBox111.Text.ToString, MAIN.TextBox167.Text.ToString, MAIN.TextBox171.Text.ToString, MAIN.TextBox115.Text.ToString, MAIN.TextBox119.Text.ToString, MAIN.TextBox123.Text.ToString, MAIN.TextBox126.Text.ToString, MAIN.TextBox135.Text.ToString, MAIN.TextBox144.Text.ToString, MAIN.Button43.Text.ToString, DateTime.Now.ToString, MAIN.TextBox5.Text.ToString, MAIN.Label10.Text.ToString)
                SMC.ClearSequence()
                Me.Close()
                TrckRecord()
            ElseIf TextBox2.Text = "D0661_OFF" And RadioButton3.Checked = True Then
                SMC.CalculateCount("FAIL")
                ClearReportBeforeGeneration()
                Dim lastItem As String = MAIN.ListBox1.Items(MAIN.ListBox1.Items.Count - 1).ToString()
                SMC.GenerateReportOnErrorSMCSwitchTestNode2(MAIN.TextBox3.Text, MAIN.TextBox2.Text, MAIN.TextBox4.Text, MAIN.Button43.Text, MAIN.TextBox5.Text, MAIN.Label10.Text, MAIN.TextBox153.Text, lastItem, "TRUE", "TRUE", MAIN.TextBox27.Text, MAIN.TextBox37.Text, MAIN.TextBox43.Text, MAIN.TextBox40.Text, MAIN.TextBox54.Text, MAIN.TextBox58.Text, MAIN.TextBox62.Text, MAIN.TextBox66.Text, MAIN.TextBox69.Text, MAIN.TextBox74.Text, MAIN.TextBox162.Text, MAIN.TextBox82.Text, MAIN.TextBox86.Text)
                SMC.SMCCreateRecord(CONFIG.TextBox1.Text, CONFIG.TextBox2.Text, CONFIG.TextBox3.Text, MAIN.TextBox4.Text.ToString, MAIN.TextBox2.Text.ToString, MAIN.TextBox3.Text.ToString, MAIN.TextBox6.Text.ToString, MAIN.TextBox153.Text.ToString, MAIN.TextBox27.Text.ToString, MAIN.TextBox37.Text.ToString, MAIN.TextBox43.Text.ToString, MAIN.TextBox40.Text.ToString, MAIN.TextBox54.Text.ToString,
                MAIN.TextBox58.Text.ToString, MAIN.TextBox62.Text.ToString, MAIN.TextBox66.Text.ToString, MAIN.TextBox69.Text.ToString, MAIN.TextBox74.Text.ToString, MAIN.TextBox162.Text.ToString, MAIN.TextBox82.Text.ToString, MAIN.TextBox86.Text.ToString, MAIN.TextBox90.Text.ToString, MAIN.TextBox95.Text.ToString, MAIN.TextBox99.Text.ToString, MAIN.TextBox103.Text.ToString, MAIN.TextBox107.Text.ToString,
                MAIN.TextBox111.Text.ToString, MAIN.TextBox167.Text.ToString, MAIN.TextBox171.Text.ToString, MAIN.TextBox115.Text.ToString, MAIN.TextBox119.Text.ToString, MAIN.TextBox123.Text.ToString, MAIN.TextBox126.Text.ToString, MAIN.TextBox135.Text.ToString, MAIN.TextBox144.Text.ToString, MAIN.Button43.Text.ToString, DateTime.Now.ToString, MAIN.TextBox5.Text.ToString, MAIN.Label10.Text.ToString)
                SMC.ClearSequence()
                Me.Close()
                TrckRecord()
            ElseIf TextBox2.Text = "AUXENCMD_ON" And RadioButton2.Checked = True Then
                SMC.CalculateCount("FAIL")
                ClearReportBeforeGeneration()
                Dim lastItem As String = MAIN.ListBox1.Items(MAIN.ListBox1.Items.Count - 1).ToString()
                SMC.GenerateReportOnErrorSMCSwitchTestNode3(MAIN.TextBox3.Text, MAIN.TextBox2.Text, MAIN.TextBox4.Text, MAIN.Button43.Text, MAIN.TextBox5.Text, MAIN.Label10.Text, MAIN.TextBox153.Text, lastItem, "TRUE", "TRUE", MAIN.TextBox27.Text, MAIN.TextBox37.Text, MAIN.TextBox43.Text, MAIN.TextBox40.Text, MAIN.TextBox54.Text, MAIN.TextBox58.Text, MAIN.TextBox62.Text, MAIN.TextBox66.Text, MAIN.TextBox69.Text, MAIN.TextBox74.Text, MAIN.TextBox162.Text, MAIN.TextBox82.Text, MAIN.TextBox86.Text, MAIN.TextBox90.Text)
                SMC.SMCCreateRecord(CONFIG.TextBox1.Text, CONFIG.TextBox2.Text, CONFIG.TextBox3.Text, MAIN.TextBox4.Text.ToString, MAIN.TextBox2.Text.ToString, MAIN.TextBox3.Text.ToString, MAIN.TextBox6.Text.ToString, MAIN.TextBox153.Text.ToString, MAIN.TextBox27.Text.ToString, MAIN.TextBox37.Text.ToString, MAIN.TextBox43.Text.ToString, MAIN.TextBox40.Text.ToString, MAIN.TextBox54.Text.ToString,
                MAIN.TextBox58.Text.ToString, MAIN.TextBox62.Text.ToString, MAIN.TextBox66.Text.ToString, MAIN.TextBox69.Text.ToString, MAIN.TextBox74.Text.ToString, MAIN.TextBox162.Text.ToString, MAIN.TextBox82.Text.ToString, MAIN.TextBox86.Text.ToString, MAIN.TextBox90.Text.ToString, MAIN.TextBox95.Text.ToString, MAIN.TextBox99.Text.ToString, MAIN.TextBox103.Text.ToString, MAIN.TextBox107.Text.ToString,
                MAIN.TextBox111.Text.ToString, MAIN.TextBox167.Text.ToString, MAIN.TextBox171.Text.ToString, MAIN.TextBox115.Text.ToString, MAIN.TextBox119.Text.ToString, MAIN.TextBox123.Text.ToString, MAIN.TextBox126.Text.ToString, MAIN.TextBox135.Text.ToString, MAIN.TextBox144.Text.ToString, MAIN.Button43.Text.ToString, DateTime.Now.ToString, MAIN.TextBox5.Text.ToString, MAIN.Label10.Text.ToString)
                SMC.ClearSequence()
                Me.Close()
                TrckRecord()
            ElseIf TextBox2.Text = "AUXENCMD_ON" And RadioButton3.Checked = True Then
                SMC.CalculateCount("FAIL")
                ClearReportBeforeGeneration()
                Dim lastItem As String = MAIN.ListBox1.Items(MAIN.ListBox1.Items.Count - 1).ToString()
                SMC.GenerateReportOnErrorSMCSwitchTestNode3(MAIN.TextBox3.Text, MAIN.TextBox2.Text, MAIN.TextBox4.Text, MAIN.Button43.Text, MAIN.TextBox5.Text, MAIN.Label10.Text, MAIN.TextBox153.Text, lastItem, "TRUE", "TRUE", MAIN.TextBox27.Text, MAIN.TextBox37.Text, MAIN.TextBox43.Text, MAIN.TextBox40.Text, MAIN.TextBox54.Text, MAIN.TextBox58.Text, MAIN.TextBox62.Text, MAIN.TextBox66.Text, MAIN.TextBox69.Text, MAIN.TextBox74.Text, MAIN.TextBox162.Text, MAIN.TextBox82.Text, MAIN.TextBox86.Text, MAIN.TextBox90.Text)
                SMC.SMCCreateRecord(CONFIG.TextBox1.Text, CONFIG.TextBox2.Text, CONFIG.TextBox3.Text, MAIN.TextBox4.Text.ToString, MAIN.TextBox2.Text.ToString, MAIN.TextBox3.Text.ToString, MAIN.TextBox6.Text.ToString, MAIN.TextBox153.Text.ToString, MAIN.TextBox27.Text.ToString, MAIN.TextBox37.Text.ToString, MAIN.TextBox43.Text.ToString, MAIN.TextBox40.Text.ToString, MAIN.TextBox54.Text.ToString,
                MAIN.TextBox58.Text.ToString, MAIN.TextBox62.Text.ToString, MAIN.TextBox66.Text.ToString, MAIN.TextBox69.Text.ToString, MAIN.TextBox74.Text.ToString, MAIN.TextBox162.Text.ToString, MAIN.TextBox82.Text.ToString, MAIN.TextBox86.Text.ToString, MAIN.TextBox90.Text.ToString, MAIN.TextBox95.Text.ToString, MAIN.TextBox99.Text.ToString, MAIN.TextBox103.Text.ToString, MAIN.TextBox107.Text.ToString,
                MAIN.TextBox111.Text.ToString, MAIN.TextBox167.Text.ToString, MAIN.TextBox171.Text.ToString, MAIN.TextBox115.Text.ToString, MAIN.TextBox119.Text.ToString, MAIN.TextBox123.Text.ToString, MAIN.TextBox126.Text.ToString, MAIN.TextBox135.Text.ToString, MAIN.TextBox144.Text.ToString, MAIN.Button43.Text.ToString, DateTime.Now.ToString, MAIN.TextBox5.Text.ToString, MAIN.Label10.Text.ToString)
                SMC.ClearSequence()
                Me.Close()
                TrckRecord()
            ElseIf TextBox2.Text = "D0100_ON" And RadioButton2.Checked = True Then
                SMC.CalculateCount("FAIL")
                ClearReportBeforeGeneration()
                Dim lastItem As String = MAIN.ListBox1.Items(MAIN.ListBox1.Items.Count - 1).ToString()
                SMC.GenerateReportOnErrorSMCSwitchTestNode4(MAIN.TextBox3.Text, MAIN.TextBox2.Text, MAIN.TextBox4.Text, MAIN.Button43.Text, MAIN.TextBox5.Text, MAIN.Label10.Text, MAIN.TextBox153.Text, lastItem, "TRUE", "TRUE", MAIN.TextBox27.Text, MAIN.TextBox37.Text, MAIN.TextBox43.Text, MAIN.TextBox40.Text, MAIN.TextBox54.Text, MAIN.TextBox58.Text, MAIN.TextBox62.Text, MAIN.TextBox66.Text, MAIN.TextBox69.Text, MAIN.TextBox74.Text, MAIN.TextBox162.Text, MAIN.TextBox82.Text, MAIN.TextBox86.Text, MAIN.TextBox90.Text, MAIN.TextBox95.Text)
                SMC.SMCCreateRecord(CONFIG.TextBox1.Text, CONFIG.TextBox2.Text, CONFIG.TextBox3.Text, MAIN.TextBox4.Text.ToString, MAIN.TextBox2.Text.ToString, MAIN.TextBox3.Text.ToString, MAIN.TextBox6.Text.ToString, MAIN.TextBox153.Text.ToString, MAIN.TextBox27.Text.ToString, MAIN.TextBox37.Text.ToString, MAIN.TextBox43.Text.ToString, MAIN.TextBox40.Text.ToString, MAIN.TextBox54.Text.ToString,
                MAIN.TextBox58.Text.ToString, MAIN.TextBox62.Text.ToString, MAIN.TextBox66.Text.ToString, MAIN.TextBox69.Text.ToString, MAIN.TextBox74.Text.ToString, MAIN.TextBox162.Text.ToString, MAIN.TextBox82.Text.ToString, MAIN.TextBox86.Text.ToString, MAIN.TextBox90.Text.ToString, MAIN.TextBox95.Text.ToString, MAIN.TextBox99.Text.ToString, MAIN.TextBox103.Text.ToString, MAIN.TextBox107.Text.ToString,
                MAIN.TextBox111.Text.ToString, MAIN.TextBox167.Text.ToString, MAIN.TextBox171.Text.ToString, MAIN.TextBox115.Text.ToString, MAIN.TextBox119.Text.ToString, MAIN.TextBox123.Text.ToString, MAIN.TextBox126.Text.ToString, MAIN.TextBox135.Text.ToString, MAIN.TextBox144.Text.ToString, MAIN.Button43.Text.ToString, DateTime.Now.ToString, MAIN.TextBox5.Text.ToString, MAIN.Label10.Text.ToString)
                SMC.ClearSequence()
                Me.Close()
                TrckRecord()
            ElseIf TextBox2.Text = "D0100_ON" And RadioButton3.Checked = True Then
                SMC.CalculateCount("FAIL")
                ClearReportBeforeGeneration()
                Dim lastItem As String = MAIN.ListBox1.Items(MAIN.ListBox1.Items.Count - 1).ToString()
                SMC.GenerateReportOnErrorSMCSwitchTestNode4(MAIN.TextBox3.Text, MAIN.TextBox2.Text, MAIN.TextBox4.Text, MAIN.Button43.Text, MAIN.TextBox5.Text, MAIN.Label10.Text, MAIN.TextBox153.Text, lastItem, "TRUE", "TRUE", MAIN.TextBox27.Text, MAIN.TextBox37.Text, MAIN.TextBox43.Text, MAIN.TextBox40.Text, MAIN.TextBox54.Text, MAIN.TextBox58.Text, MAIN.TextBox62.Text, MAIN.TextBox66.Text, MAIN.TextBox69.Text, MAIN.TextBox74.Text, MAIN.TextBox162.Text, MAIN.TextBox82.Text, MAIN.TextBox86.Text, MAIN.TextBox90.Text, MAIN.TextBox95.Text)
                SMC.SMCCreateRecord(CONFIG.TextBox1.Text, CONFIG.TextBox2.Text, CONFIG.TextBox3.Text, MAIN.TextBox4.Text.ToString, MAIN.TextBox2.Text.ToString, MAIN.TextBox3.Text.ToString, MAIN.TextBox6.Text.ToString, MAIN.TextBox153.Text.ToString, MAIN.TextBox27.Text.ToString, MAIN.TextBox37.Text.ToString, MAIN.TextBox43.Text.ToString, MAIN.TextBox40.Text.ToString, MAIN.TextBox54.Text.ToString,
                MAIN.TextBox58.Text.ToString, MAIN.TextBox62.Text.ToString, MAIN.TextBox66.Text.ToString, MAIN.TextBox69.Text.ToString, MAIN.TextBox74.Text.ToString, MAIN.TextBox162.Text.ToString, MAIN.TextBox82.Text.ToString, MAIN.TextBox86.Text.ToString, MAIN.TextBox90.Text.ToString, MAIN.TextBox95.Text.ToString, MAIN.TextBox99.Text.ToString, MAIN.TextBox103.Text.ToString, MAIN.TextBox107.Text.ToString,
                MAIN.TextBox111.Text.ToString, MAIN.TextBox167.Text.ToString, MAIN.TextBox171.Text.ToString, MAIN.TextBox115.Text.ToString, MAIN.TextBox119.Text.ToString, MAIN.TextBox123.Text.ToString, MAIN.TextBox126.Text.ToString, MAIN.TextBox135.Text.ToString, MAIN.TextBox144.Text.ToString, MAIN.Button43.Text.ToString, DateTime.Now.ToString, MAIN.TextBox5.Text.ToString, MAIN.Label10.Text.ToString)
                SMC.ClearSequence()
                Me.Close()
                TrckRecord()
            ElseIf TextBox2.Text = "D0100_OFF" And RadioButton2.Checked = True Then
                SMC.CalculateCount("FAIL")
                ClearReportBeforeGeneration()
                Dim lastItem As String = MAIN.ListBox1.Items(MAIN.ListBox1.Items.Count - 1).ToString()
                SMC.GenerateReportOnErrorSMCSwitchTestNode5(MAIN.TextBox3.Text, MAIN.TextBox2.Text, MAIN.TextBox4.Text, MAIN.Button43.Text, MAIN.TextBox5.Text, MAIN.Label10.Text, MAIN.TextBox153.Text, lastItem, "TRUE", "TRUE", MAIN.TextBox27.Text, MAIN.TextBox37.Text, MAIN.TextBox43.Text, MAIN.TextBox40.Text, MAIN.TextBox54.Text, MAIN.TextBox58.Text, MAIN.TextBox62.Text, MAIN.TextBox66.Text, MAIN.TextBox69.Text, MAIN.TextBox74.Text, MAIN.TextBox162.Text, MAIN.TextBox82.Text, MAIN.TextBox86.Text, MAIN.TextBox90.Text, MAIN.TextBox95.Text, MAIN.TextBox99.Text)
                SMC.SMCCreateRecord(CONFIG.TextBox1.Text, CONFIG.TextBox2.Text, CONFIG.TextBox3.Text, MAIN.TextBox4.Text.ToString, MAIN.TextBox2.Text.ToString, MAIN.TextBox3.Text.ToString, MAIN.TextBox6.Text.ToString, MAIN.TextBox153.Text.ToString, MAIN.TextBox27.Text.ToString, MAIN.TextBox37.Text.ToString, MAIN.TextBox43.Text.ToString, MAIN.TextBox40.Text.ToString, MAIN.TextBox54.Text.ToString,
                MAIN.TextBox58.Text.ToString, MAIN.TextBox62.Text.ToString, MAIN.TextBox66.Text.ToString, MAIN.TextBox69.Text.ToString, MAIN.TextBox74.Text.ToString, MAIN.TextBox162.Text.ToString, MAIN.TextBox82.Text.ToString, MAIN.TextBox86.Text.ToString, MAIN.TextBox90.Text.ToString, MAIN.TextBox95.Text.ToString, MAIN.TextBox99.Text.ToString, MAIN.TextBox103.Text.ToString, MAIN.TextBox107.Text.ToString,
                MAIN.TextBox111.Text.ToString, MAIN.TextBox167.Text.ToString, MAIN.TextBox171.Text.ToString, MAIN.TextBox115.Text.ToString, MAIN.TextBox119.Text.ToString, MAIN.TextBox123.Text.ToString, MAIN.TextBox126.Text.ToString, MAIN.TextBox135.Text.ToString, MAIN.TextBox144.Text.ToString, MAIN.Button43.Text.ToString, DateTime.Now.ToString, MAIN.TextBox5.Text.ToString, MAIN.Label10.Text.ToString)
                SMC.ClearSequence()
                Me.Close()
                TrckRecord()
            ElseIf TextBox2.Text = "D0100_OFF" And RadioButton3.Checked = True Then
                SMC.CalculateCount("FAIL")
                ClearReportBeforeGeneration()
                Dim lastItem As String = MAIN.ListBox1.Items(MAIN.ListBox1.Items.Count - 1).ToString()
                SMC.GenerateReportOnErrorSMCSwitchTestNode5(MAIN.TextBox3.Text, MAIN.TextBox2.Text, MAIN.TextBox4.Text, MAIN.Button43.Text, MAIN.TextBox5.Text, MAIN.Label10.Text, MAIN.TextBox153.Text, lastItem, "TRUE", "TRUE", MAIN.TextBox27.Text, MAIN.TextBox37.Text, MAIN.TextBox43.Text, MAIN.TextBox40.Text, MAIN.TextBox54.Text, MAIN.TextBox58.Text, MAIN.TextBox62.Text, MAIN.TextBox66.Text, MAIN.TextBox69.Text, MAIN.TextBox74.Text, MAIN.TextBox162.Text, MAIN.TextBox82.Text, MAIN.TextBox86.Text, MAIN.TextBox90.Text, MAIN.TextBox95.Text, MAIN.TextBox99.Text)
                SMC.SMCCreateRecord(CONFIG.TextBox1.Text, CONFIG.TextBox2.Text, CONFIG.TextBox3.Text, MAIN.TextBox4.Text.ToString, MAIN.TextBox2.Text.ToString, MAIN.TextBox3.Text.ToString, MAIN.TextBox6.Text.ToString, MAIN.TextBox153.Text.ToString, MAIN.TextBox27.Text.ToString, MAIN.TextBox37.Text.ToString, MAIN.TextBox43.Text.ToString, MAIN.TextBox40.Text.ToString, MAIN.TextBox54.Text.ToString,
                MAIN.TextBox58.Text.ToString, MAIN.TextBox62.Text.ToString, MAIN.TextBox66.Text.ToString, MAIN.TextBox69.Text.ToString, MAIN.TextBox74.Text.ToString, MAIN.TextBox162.Text.ToString, MAIN.TextBox82.Text.ToString, MAIN.TextBox86.Text.ToString, MAIN.TextBox90.Text.ToString, MAIN.TextBox95.Text.ToString, MAIN.TextBox99.Text.ToString, MAIN.TextBox103.Text.ToString, MAIN.TextBox107.Text.ToString,
                MAIN.TextBox111.Text.ToString, MAIN.TextBox167.Text.ToString, MAIN.TextBox171.Text.ToString, MAIN.TextBox115.Text.ToString, MAIN.TextBox119.Text.ToString, MAIN.TextBox123.Text.ToString, MAIN.TextBox126.Text.ToString, MAIN.TextBox135.Text.ToString, MAIN.TextBox144.Text.ToString, MAIN.Button43.Text.ToString, DateTime.Now.ToString, MAIN.TextBox5.Text.ToString, MAIN.Label10.Text.ToString)
                SMC.ClearSequence()
                Me.Close()
                TrckRecord()
            ElseIf TextBox2.Text = "D0101_ON" And RadioButton2.Checked = True Then
                SMC.CalculateCount("FAIL")
                ClearReportBeforeGeneration()
                Dim lastItem As String = MAIN.ListBox1.Items(MAIN.ListBox1.Items.Count - 1).ToString()
                SMC.GenerateReportOnErrorSMCSwitchTestNode6(MAIN.TextBox3.Text, MAIN.TextBox2.Text, MAIN.TextBox4.Text, MAIN.Button43.Text, MAIN.TextBox5.Text, MAIN.Label10.Text, MAIN.TextBox153.Text, lastItem, "TRUE", "TRUE", MAIN.TextBox27.Text, MAIN.TextBox37.Text, MAIN.TextBox43.Text, MAIN.TextBox40.Text, MAIN.TextBox54.Text, MAIN.TextBox58.Text, MAIN.TextBox62.Text, MAIN.TextBox66.Text, MAIN.TextBox69.Text, MAIN.TextBox74.Text, MAIN.TextBox162.Text, MAIN.TextBox82.Text, MAIN.TextBox86.Text, MAIN.TextBox90.Text, MAIN.TextBox95.Text, MAIN.TextBox99.Text, MAIN.TextBox103.Text)
                SMC.SMCCreateRecord(CONFIG.TextBox1.Text, CONFIG.TextBox2.Text, CONFIG.TextBox3.Text, MAIN.TextBox4.Text.ToString, MAIN.TextBox2.Text.ToString, MAIN.TextBox3.Text.ToString, MAIN.TextBox6.Text.ToString, MAIN.TextBox153.Text.ToString, MAIN.TextBox27.Text.ToString, MAIN.TextBox37.Text.ToString, MAIN.TextBox43.Text.ToString, MAIN.TextBox40.Text.ToString, MAIN.TextBox54.Text.ToString,
                MAIN.TextBox58.Text.ToString, MAIN.TextBox62.Text.ToString, MAIN.TextBox66.Text.ToString, MAIN.TextBox69.Text.ToString, MAIN.TextBox74.Text.ToString, MAIN.TextBox162.Text.ToString, MAIN.TextBox82.Text.ToString, MAIN.TextBox86.Text.ToString, MAIN.TextBox90.Text.ToString, MAIN.TextBox95.Text.ToString, MAIN.TextBox99.Text.ToString, MAIN.TextBox103.Text.ToString, MAIN.TextBox107.Text.ToString,
                MAIN.TextBox111.Text.ToString, MAIN.TextBox167.Text.ToString, MAIN.TextBox171.Text.ToString, MAIN.TextBox115.Text.ToString, MAIN.TextBox119.Text.ToString, MAIN.TextBox123.Text.ToString, MAIN.TextBox126.Text.ToString, MAIN.TextBox135.Text.ToString, MAIN.TextBox144.Text.ToString, MAIN.Button43.Text.ToString, DateTime.Now.ToString, MAIN.TextBox5.Text.ToString, MAIN.Label10.Text.ToString)
                SMC.ClearSequence()
                Me.Close()
                TrckRecord()
            ElseIf TextBox2.Text = "D0101_ON" And RadioButton3.Checked = True Then
                SMC.CalculateCount("FAIL")
                ClearReportBeforeGeneration()
                Dim lastItem As String = MAIN.ListBox1.Items(MAIN.ListBox1.Items.Count - 1).ToString()
                SMC.GenerateReportOnErrorSMCSwitchTestNode6(MAIN.TextBox3.Text, MAIN.TextBox2.Text, MAIN.TextBox4.Text, MAIN.Button43.Text, MAIN.TextBox5.Text, MAIN.Label10.Text, MAIN.TextBox153.Text, lastItem, "TRUE", "TRUE", MAIN.TextBox27.Text, MAIN.TextBox37.Text, MAIN.TextBox43.Text, MAIN.TextBox40.Text, MAIN.TextBox54.Text, MAIN.TextBox58.Text, MAIN.TextBox62.Text, MAIN.TextBox66.Text, MAIN.TextBox69.Text, MAIN.TextBox74.Text, MAIN.TextBox162.Text, MAIN.TextBox82.Text, MAIN.TextBox86.Text, MAIN.TextBox90.Text, MAIN.TextBox95.Text, MAIN.TextBox99.Text, MAIN.TextBox103.Text)
                SMC.SMCCreateRecord(CONFIG.TextBox1.Text, CONFIG.TextBox2.Text, CONFIG.TextBox3.Text, MAIN.TextBox4.Text.ToString, MAIN.TextBox2.Text.ToString, MAIN.TextBox3.Text.ToString, MAIN.TextBox6.Text.ToString, MAIN.TextBox153.Text.ToString, MAIN.TextBox27.Text.ToString, MAIN.TextBox37.Text.ToString, MAIN.TextBox43.Text.ToString, MAIN.TextBox40.Text.ToString, MAIN.TextBox54.Text.ToString,
                MAIN.TextBox58.Text.ToString, MAIN.TextBox62.Text.ToString, MAIN.TextBox66.Text.ToString, MAIN.TextBox69.Text.ToString, MAIN.TextBox74.Text.ToString, MAIN.TextBox162.Text.ToString, MAIN.TextBox82.Text.ToString, MAIN.TextBox86.Text.ToString, MAIN.TextBox90.Text.ToString, MAIN.TextBox95.Text.ToString, MAIN.TextBox99.Text.ToString, MAIN.TextBox103.Text.ToString, MAIN.TextBox107.Text.ToString,
                MAIN.TextBox111.Text.ToString, MAIN.TextBox167.Text.ToString, MAIN.TextBox171.Text.ToString, MAIN.TextBox115.Text.ToString, MAIN.TextBox119.Text.ToString, MAIN.TextBox123.Text.ToString, MAIN.TextBox126.Text.ToString, MAIN.TextBox135.Text.ToString, MAIN.TextBox144.Text.ToString, MAIN.Button43.Text.ToString, DateTime.Now.ToString, MAIN.TextBox5.Text.ToString, MAIN.Label10.Text.ToString)
                SMC.ClearSequence()
                Me.Close()
                TrckRecord()
            ElseIf TextBox2.Text = "D0101_OFF" And RadioButton2.Checked = True Then
                SMC.CalculateCount("FAIL")
                ClearReportBeforeGeneration()
                Dim lastItem As String = MAIN.ListBox1.Items(MAIN.ListBox1.Items.Count - 1).ToString()
                SMC.GenerateReportOnErrorSMCSwitchTestNode7(MAIN.TextBox3.Text, MAIN.TextBox2.Text, MAIN.TextBox4.Text, MAIN.Button43.Text, MAIN.TextBox5.Text, MAIN.Label10.Text, MAIN.TextBox153.Text, lastItem, "TRUE", "TRUE", MAIN.TextBox27.Text, MAIN.TextBox37.Text, MAIN.TextBox43.Text, MAIN.TextBox40.Text, MAIN.TextBox54.Text, MAIN.TextBox58.Text, MAIN.TextBox62.Text, MAIN.TextBox66.Text, MAIN.TextBox69.Text, MAIN.TextBox74.Text, MAIN.TextBox162.Text, MAIN.TextBox82.Text, MAIN.TextBox86.Text, MAIN.TextBox90.Text, MAIN.TextBox95.Text, MAIN.TextBox99.Text, MAIN.TextBox103.Text, MAIN.TextBox107.Text)
                SMC.SMCCreateRecord(CONFIG.TextBox1.Text, CONFIG.TextBox2.Text, CONFIG.TextBox3.Text, MAIN.TextBox4.Text.ToString, MAIN.TextBox2.Text.ToString, MAIN.TextBox3.Text.ToString, MAIN.TextBox6.Text.ToString, MAIN.TextBox153.Text.ToString, MAIN.TextBox27.Text.ToString, MAIN.TextBox37.Text.ToString, MAIN.TextBox43.Text.ToString, MAIN.TextBox40.Text.ToString, MAIN.TextBox54.Text.ToString,
                MAIN.TextBox58.Text.ToString, MAIN.TextBox62.Text.ToString, MAIN.TextBox66.Text.ToString, MAIN.TextBox69.Text.ToString, MAIN.TextBox74.Text.ToString, MAIN.TextBox162.Text.ToString, MAIN.TextBox82.Text.ToString, MAIN.TextBox86.Text.ToString, MAIN.TextBox90.Text.ToString, MAIN.TextBox95.Text.ToString, MAIN.TextBox99.Text.ToString, MAIN.TextBox103.Text.ToString, MAIN.TextBox107.Text.ToString,
                MAIN.TextBox111.Text.ToString, MAIN.TextBox167.Text.ToString, MAIN.TextBox171.Text.ToString, MAIN.TextBox115.Text.ToString, MAIN.TextBox119.Text.ToString, MAIN.TextBox123.Text.ToString, MAIN.TextBox126.Text.ToString, MAIN.TextBox135.Text.ToString, MAIN.TextBox144.Text.ToString, MAIN.Button43.Text.ToString, DateTime.Now.ToString, MAIN.TextBox5.Text.ToString, MAIN.Label10.Text.ToString)
                SMC.ClearSequence()
                Me.Close()
                TrckRecord()
            ElseIf TextBox2.Text = "D0101_OFF" And RadioButton3.Checked = True Then
                SMC.CalculateCount("FAIL")
                ClearReportBeforeGeneration()
                Dim lastItem As String = MAIN.ListBox1.Items(MAIN.ListBox1.Items.Count - 1).ToString()
                SMC.GenerateReportOnErrorSMCSwitchTestNode7(MAIN.TextBox3.Text, MAIN.TextBox2.Text, MAIN.TextBox4.Text, MAIN.Button43.Text, MAIN.TextBox5.Text, MAIN.Label10.Text, MAIN.TextBox153.Text, lastItem, "TRUE", "TRUE", MAIN.TextBox27.Text, MAIN.TextBox37.Text, MAIN.TextBox43.Text, MAIN.TextBox40.Text, MAIN.TextBox54.Text, MAIN.TextBox58.Text, MAIN.TextBox62.Text, MAIN.TextBox66.Text, MAIN.TextBox69.Text, MAIN.TextBox74.Text, MAIN.TextBox162.Text, MAIN.TextBox82.Text, MAIN.TextBox86.Text, MAIN.TextBox90.Text, MAIN.TextBox95.Text, MAIN.TextBox99.Text, MAIN.TextBox103.Text, MAIN.TextBox107.Text)
                SMC.SMCCreateRecord(CONFIG.TextBox1.Text, CONFIG.TextBox2.Text, CONFIG.TextBox3.Text, MAIN.TextBox4.Text.ToString, MAIN.TextBox2.Text.ToString, MAIN.TextBox3.Text.ToString, MAIN.TextBox6.Text.ToString, MAIN.TextBox153.Text.ToString, MAIN.TextBox27.Text.ToString, MAIN.TextBox37.Text.ToString, MAIN.TextBox43.Text.ToString, MAIN.TextBox40.Text.ToString, MAIN.TextBox54.Text.ToString,
                MAIN.TextBox58.Text.ToString, MAIN.TextBox62.Text.ToString, MAIN.TextBox66.Text.ToString, MAIN.TextBox69.Text.ToString, MAIN.TextBox74.Text.ToString, MAIN.TextBox162.Text.ToString, MAIN.TextBox82.Text.ToString, MAIN.TextBox86.Text.ToString, MAIN.TextBox90.Text.ToString, MAIN.TextBox95.Text.ToString, MAIN.TextBox99.Text.ToString, MAIN.TextBox103.Text.ToString, MAIN.TextBox107.Text.ToString,
                MAIN.TextBox111.Text.ToString, MAIN.TextBox167.Text.ToString, MAIN.TextBox171.Text.ToString, MAIN.TextBox115.Text.ToString, MAIN.TextBox119.Text.ToString, MAIN.TextBox123.Text.ToString, MAIN.TextBox126.Text.ToString, MAIN.TextBox135.Text.ToString, MAIN.TextBox144.Text.ToString, MAIN.Button43.Text.ToString, DateTime.Now.ToString, MAIN.TextBox5.Text.ToString, MAIN.Label10.Text.ToString)
                SMC.ClearSequence()
                Me.Close()
                TrckRecord()
            ElseIf TextBox2.Text = "CMD_OFF" And RadioButton2.Checked = True Then
                SMC.CalculateCount("FAIL")
                ClearReportBeforeGeneration()
                Dim lastItem As String = MAIN.ListBox1.Items(MAIN.ListBox1.Items.Count - 1).ToString()
                SMC.GenerateReportOnErrorSMCSwitchTestNode8(MAIN.TextBox3.Text, MAIN.TextBox2.Text, MAIN.TextBox4.Text, MAIN.Button43.Text, MAIN.TextBox5.Text, MAIN.Label10.Text, MAIN.TextBox153.Text, lastItem, "TRUE", "TRUE", MAIN.TextBox27.Text, MAIN.TextBox37.Text, MAIN.TextBox43.Text, MAIN.TextBox40.Text, MAIN.TextBox54.Text, MAIN.TextBox58.Text, MAIN.TextBox62.Text, MAIN.TextBox66.Text, MAIN.TextBox69.Text, MAIN.TextBox74.Text, MAIN.TextBox162.Text, MAIN.TextBox82.Text, MAIN.TextBox86.Text, MAIN.TextBox90.Text, MAIN.TextBox95.Text, MAIN.TextBox99.Text, MAIN.TextBox103.Text, MAIN.TextBox107.Text, MAIN.TextBox111.Text)
                SMC.SMCCreateRecord(CONFIG.TextBox1.Text, CONFIG.TextBox2.Text, CONFIG.TextBox3.Text, MAIN.TextBox4.Text.ToString, MAIN.TextBox2.Text.ToString, MAIN.TextBox3.Text.ToString, MAIN.TextBox6.Text.ToString, MAIN.TextBox153.Text.ToString, MAIN.TextBox27.Text.ToString, MAIN.TextBox37.Text.ToString, MAIN.TextBox43.Text.ToString, MAIN.TextBox40.Text.ToString, MAIN.TextBox54.Text.ToString,
                MAIN.TextBox58.Text.ToString, MAIN.TextBox62.Text.ToString, MAIN.TextBox66.Text.ToString, MAIN.TextBox69.Text.ToString, MAIN.TextBox74.Text.ToString, MAIN.TextBox162.Text.ToString, MAIN.TextBox82.Text.ToString, MAIN.TextBox86.Text.ToString, MAIN.TextBox90.Text.ToString, MAIN.TextBox95.Text.ToString, MAIN.TextBox99.Text.ToString, MAIN.TextBox103.Text.ToString, MAIN.TextBox107.Text.ToString,
                MAIN.TextBox111.Text.ToString, MAIN.TextBox167.Text.ToString, MAIN.TextBox171.Text.ToString, MAIN.TextBox115.Text.ToString, MAIN.TextBox119.Text.ToString, MAIN.TextBox123.Text.ToString, MAIN.TextBox126.Text.ToString, MAIN.TextBox135.Text.ToString, MAIN.TextBox144.Text.ToString, MAIN.Button43.Text.ToString, DateTime.Now.ToString, MAIN.TextBox5.Text.ToString, MAIN.Label10.Text.ToString)
                SMC.ClearSequence()
                Me.Close()
                TrckRecord()
            ElseIf TextBox2.Text = "CMD_OFF" And RadioButton3.Checked = True Then
                SMC.CalculateCount("FAIL")
                ClearReportBeforeGeneration()
                Dim lastItem As String = MAIN.ListBox1.Items(MAIN.ListBox1.Items.Count - 1).ToString()
                SMC.GenerateReportOnErrorSMCSwitchTestNode8(MAIN.TextBox3.Text, MAIN.TextBox2.Text, MAIN.TextBox4.Text, MAIN.Button43.Text, MAIN.TextBox5.Text, MAIN.Label10.Text, MAIN.TextBox153.Text, lastItem, "TRUE", "TRUE", MAIN.TextBox27.Text, MAIN.TextBox37.Text, MAIN.TextBox43.Text, MAIN.TextBox40.Text, MAIN.TextBox54.Text, MAIN.TextBox58.Text, MAIN.TextBox62.Text, MAIN.TextBox66.Text, MAIN.TextBox69.Text, MAIN.TextBox74.Text, MAIN.TextBox162.Text, MAIN.TextBox82.Text, MAIN.TextBox86.Text, MAIN.TextBox90.Text, MAIN.TextBox95.Text, MAIN.TextBox99.Text, MAIN.TextBox103.Text, MAIN.TextBox107.Text, MAIN.TextBox111.Text)
                SMC.SMCCreateRecord(CONFIG.TextBox1.Text, CONFIG.TextBox2.Text, CONFIG.TextBox3.Text, MAIN.TextBox4.Text.ToString, MAIN.TextBox2.Text.ToString, MAIN.TextBox3.Text.ToString, MAIN.TextBox6.Text.ToString, MAIN.TextBox153.Text.ToString, MAIN.TextBox27.Text.ToString, MAIN.TextBox37.Text.ToString, MAIN.TextBox43.Text.ToString, MAIN.TextBox40.Text.ToString, MAIN.TextBox54.Text.ToString,
                MAIN.TextBox58.Text.ToString, MAIN.TextBox62.Text.ToString, MAIN.TextBox66.Text.ToString, MAIN.TextBox69.Text.ToString, MAIN.TextBox74.Text.ToString, MAIN.TextBox162.Text.ToString, MAIN.TextBox82.Text.ToString, MAIN.TextBox86.Text.ToString, MAIN.TextBox90.Text.ToString, MAIN.TextBox95.Text.ToString, MAIN.TextBox99.Text.ToString, MAIN.TextBox103.Text.ToString, MAIN.TextBox107.Text.ToString,
                MAIN.TextBox111.Text.ToString, MAIN.TextBox167.Text.ToString, MAIN.TextBox171.Text.ToString, MAIN.TextBox115.Text.ToString, MAIN.TextBox119.Text.ToString, MAIN.TextBox123.Text.ToString, MAIN.TextBox126.Text.ToString, MAIN.TextBox135.Text.ToString, MAIN.TextBox144.Text.ToString, MAIN.Button43.Text.ToString, DateTime.Now.ToString, MAIN.TextBox5.Text.ToString, MAIN.Label10.Text.ToString)
                SMC.ClearSequence()
                Me.Close()
                TrckRecord()
            ElseIf TextBox2.Text = "SW2_ON" And RadioButton2.Checked = True Then
                SMC.CalculateCount("FAIL")
                ClearReportBeforeGeneration()
                Dim lastItem As String = MAIN.ListBox1.Items(MAIN.ListBox1.Items.Count - 1).ToString()
                SMC.GenerateReportOnErrorSMCSwitchTestNode9(MAIN.TextBox3.Text, MAIN.TextBox2.Text, MAIN.TextBox4.Text, MAIN.Button43.Text, MAIN.TextBox5.Text, MAIN.Label10.Text, MAIN.TextBox153.Text, lastItem, "TRUE", "TRUE", MAIN.TextBox27.Text, MAIN.TextBox37.Text, MAIN.TextBox43.Text, MAIN.TextBox40.Text, MAIN.TextBox54.Text, MAIN.TextBox58.Text, MAIN.TextBox62.Text, MAIN.TextBox66.Text, MAIN.TextBox69.Text, MAIN.TextBox74.Text, MAIN.TextBox162.Text, MAIN.TextBox82.Text, MAIN.TextBox86.Text, MAIN.TextBox90.Text, MAIN.TextBox95.Text, MAIN.TextBox99.Text, MAIN.TextBox103.Text, MAIN.TextBox107.Text, MAIN.TextBox111.Text, MAIN.TextBox167.Text)
                SMC.SMCCreateRecord(CONFIG.TextBox1.Text, CONFIG.TextBox2.Text, CONFIG.TextBox3.Text, MAIN.TextBox4.Text.ToString, MAIN.TextBox2.Text.ToString, MAIN.TextBox3.Text.ToString, MAIN.TextBox6.Text.ToString, MAIN.TextBox153.Text.ToString, MAIN.TextBox27.Text.ToString, MAIN.TextBox37.Text.ToString, MAIN.TextBox43.Text.ToString, MAIN.TextBox40.Text.ToString, MAIN.TextBox54.Text.ToString,
                MAIN.TextBox58.Text.ToString, MAIN.TextBox62.Text.ToString, MAIN.TextBox66.Text.ToString, MAIN.TextBox69.Text.ToString, MAIN.TextBox74.Text.ToString, MAIN.TextBox162.Text.ToString, MAIN.TextBox82.Text.ToString, MAIN.TextBox86.Text.ToString, MAIN.TextBox90.Text.ToString, MAIN.TextBox95.Text.ToString, MAIN.TextBox99.Text.ToString, MAIN.TextBox103.Text.ToString, MAIN.TextBox107.Text.ToString,
                MAIN.TextBox111.Text.ToString, MAIN.TextBox167.Text.ToString, MAIN.TextBox171.Text.ToString, MAIN.TextBox115.Text.ToString, MAIN.TextBox119.Text.ToString, MAIN.TextBox123.Text.ToString, MAIN.TextBox126.Text.ToString, MAIN.TextBox135.Text.ToString, MAIN.TextBox144.Text.ToString, MAIN.Button43.Text.ToString, DateTime.Now.ToString, MAIN.TextBox5.Text.ToString, MAIN.Label10.Text.ToString)
                SMC.ClearSequence()
                Me.Close()
                TrckRecord()
            ElseIf TextBox2.Text = "SW2_ON" And RadioButton3.Checked = True Then
                SMC.CalculateCount("FAIL")
                ClearReportBeforeGeneration()
                Dim lastItem As String = MAIN.ListBox1.Items(MAIN.ListBox1.Items.Count - 1).ToString()
                SMC.GenerateReportOnErrorSMCSwitchTestNode9(MAIN.TextBox3.Text, MAIN.TextBox2.Text, MAIN.TextBox4.Text, MAIN.Button43.Text, MAIN.TextBox5.Text, MAIN.Label10.Text, MAIN.TextBox153.Text, lastItem, "TRUE", "TRUE", MAIN.TextBox27.Text, MAIN.TextBox37.Text, MAIN.TextBox43.Text, MAIN.TextBox40.Text, MAIN.TextBox54.Text, MAIN.TextBox58.Text, MAIN.TextBox62.Text, MAIN.TextBox66.Text, MAIN.TextBox69.Text, MAIN.TextBox74.Text, MAIN.TextBox162.Text, MAIN.TextBox82.Text, MAIN.TextBox86.Text, MAIN.TextBox90.Text, MAIN.TextBox95.Text, MAIN.TextBox99.Text, MAIN.TextBox103.Text, MAIN.TextBox107.Text, MAIN.TextBox111.Text, MAIN.TextBox167.Text)
                SMC.SMCCreateRecord(CONFIG.TextBox1.Text, CONFIG.TextBox2.Text, CONFIG.TextBox3.Text, MAIN.TextBox4.Text.ToString, MAIN.TextBox2.Text.ToString, MAIN.TextBox3.Text.ToString, MAIN.TextBox6.Text.ToString, MAIN.TextBox153.Text.ToString, MAIN.TextBox27.Text.ToString, MAIN.TextBox37.Text.ToString, MAIN.TextBox43.Text.ToString, MAIN.TextBox40.Text.ToString, MAIN.TextBox54.Text.ToString,
                MAIN.TextBox58.Text.ToString, MAIN.TextBox62.Text.ToString, MAIN.TextBox66.Text.ToString, MAIN.TextBox69.Text.ToString, MAIN.TextBox74.Text.ToString, MAIN.TextBox162.Text.ToString, MAIN.TextBox82.Text.ToString, MAIN.TextBox86.Text.ToString, MAIN.TextBox90.Text.ToString, MAIN.TextBox95.Text.ToString, MAIN.TextBox99.Text.ToString, MAIN.TextBox103.Text.ToString, MAIN.TextBox107.Text.ToString,
                MAIN.TextBox111.Text.ToString, MAIN.TextBox167.Text.ToString, MAIN.TextBox171.Text.ToString, MAIN.TextBox115.Text.ToString, MAIN.TextBox119.Text.ToString, MAIN.TextBox123.Text.ToString, MAIN.TextBox126.Text.ToString, MAIN.TextBox135.Text.ToString, MAIN.TextBox144.Text.ToString, MAIN.Button43.Text.ToString, DateTime.Now.ToString, MAIN.TextBox5.Text.ToString, MAIN.Label10.Text.ToString)
                SMC.ClearSequence()
                Me.Close()
                TrckRecord()
            ElseIf TextBox2.Text = "SW3_ON" And RadioButton2.Checked = True Then
                SMC.CalculateCount("FAIL")
                ClearReportBeforeGeneration()
                Dim lastItem As String = MAIN.ListBox1.Items(MAIN.ListBox1.Items.Count - 1).ToString()
                SMC.GenerateReportOnErrorSMCSwitchTestNode10(MAIN.TextBox3.Text, MAIN.TextBox2.Text, MAIN.TextBox4.Text, MAIN.Button43.Text, MAIN.TextBox5.Text, MAIN.Label10.Text, MAIN.TextBox153.Text, lastItem, "TRUE", "TRUE", MAIN.TextBox27.Text, MAIN.TextBox37.Text, MAIN.TextBox43.Text, MAIN.TextBox40.Text, MAIN.TextBox54.Text, MAIN.TextBox58.Text, MAIN.TextBox62.Text, MAIN.TextBox66.Text, MAIN.TextBox69.Text, MAIN.TextBox74.Text, MAIN.TextBox162.Text, MAIN.TextBox82.Text, MAIN.TextBox86.Text, MAIN.TextBox90.Text, MAIN.TextBox95.Text, MAIN.TextBox99.Text, MAIN.TextBox103.Text, MAIN.TextBox107.Text, MAIN.TextBox111.Text, MAIN.TextBox167.Text, MAIN.TextBox171.Text)
                SMC.SMCCreateRecord(CONFIG.TextBox1.Text, CONFIG.TextBox2.Text, CONFIG.TextBox3.Text, MAIN.TextBox4.Text.ToString, MAIN.TextBox2.Text.ToString, MAIN.TextBox3.Text.ToString, MAIN.TextBox6.Text.ToString, MAIN.TextBox153.Text.ToString, MAIN.TextBox27.Text.ToString, MAIN.TextBox37.Text.ToString, MAIN.TextBox43.Text.ToString, MAIN.TextBox40.Text.ToString, MAIN.TextBox54.Text.ToString,
                MAIN.TextBox58.Text.ToString, MAIN.TextBox62.Text.ToString, MAIN.TextBox66.Text.ToString, MAIN.TextBox69.Text.ToString, MAIN.TextBox74.Text.ToString, MAIN.TextBox162.Text.ToString, MAIN.TextBox82.Text.ToString, MAIN.TextBox86.Text.ToString, MAIN.TextBox90.Text.ToString, MAIN.TextBox95.Text.ToString, MAIN.TextBox99.Text.ToString, MAIN.TextBox103.Text.ToString, MAIN.TextBox107.Text.ToString,
                MAIN.TextBox111.Text.ToString, MAIN.TextBox167.Text.ToString, MAIN.TextBox171.Text.ToString, MAIN.TextBox115.Text.ToString, MAIN.TextBox119.Text.ToString, MAIN.TextBox123.Text.ToString, MAIN.TextBox126.Text.ToString, MAIN.TextBox135.Text.ToString, MAIN.TextBox144.Text.ToString, MAIN.Button43.Text.ToString, DateTime.Now.ToString, MAIN.TextBox5.Text.ToString, MAIN.Label10.Text.ToString)
                SMC.ClearSequence()
                Me.Close()
                TrckRecord()
            ElseIf TextBox2.Text = "SW3_ON" And RadioButton3.Checked = True Then
                SMC.CalculateCount("FAIL")
                ClearReportBeforeGeneration()
                Dim lastItem As String = MAIN.ListBox1.Items(MAIN.ListBox1.Items.Count - 1).ToString()
                SMC.GenerateReportOnErrorSMCSwitchTestNode10(MAIN.TextBox3.Text, MAIN.TextBox2.Text, MAIN.TextBox4.Text, MAIN.Button43.Text, MAIN.TextBox5.Text, MAIN.Label10.Text, MAIN.TextBox153.Text, lastItem, "TRUE", "TRUE", MAIN.TextBox27.Text, MAIN.TextBox37.Text, MAIN.TextBox43.Text, MAIN.TextBox40.Text, MAIN.TextBox54.Text, MAIN.TextBox58.Text, MAIN.TextBox62.Text, MAIN.TextBox66.Text, MAIN.TextBox69.Text, MAIN.TextBox74.Text, MAIN.TextBox162.Text, MAIN.TextBox82.Text, MAIN.TextBox86.Text, MAIN.TextBox90.Text, MAIN.TextBox95.Text, MAIN.TextBox99.Text, MAIN.TextBox103.Text, MAIN.TextBox107.Text, MAIN.TextBox111.Text, MAIN.TextBox167.Text, MAIN.TextBox171.Text)
                SMC.SMCCreateRecord(CONFIG.TextBox1.Text, CONFIG.TextBox2.Text, CONFIG.TextBox3.Text, MAIN.TextBox4.Text.ToString, MAIN.TextBox2.Text.ToString, MAIN.TextBox3.Text.ToString, MAIN.TextBox6.Text.ToString, MAIN.TextBox153.Text.ToString, MAIN.TextBox27.Text.ToString, MAIN.TextBox37.Text.ToString, MAIN.TextBox43.Text.ToString, MAIN.TextBox40.Text.ToString, MAIN.TextBox54.Text.ToString,
                MAIN.TextBox58.Text.ToString, MAIN.TextBox62.Text.ToString, MAIN.TextBox66.Text.ToString, MAIN.TextBox69.Text.ToString, MAIN.TextBox74.Text.ToString, MAIN.TextBox162.Text.ToString, MAIN.TextBox82.Text.ToString, MAIN.TextBox86.Text.ToString, MAIN.TextBox90.Text.ToString, MAIN.TextBox95.Text.ToString, MAIN.TextBox99.Text.ToString, MAIN.TextBox103.Text.ToString, MAIN.TextBox107.Text.ToString,
                MAIN.TextBox111.Text.ToString, MAIN.TextBox167.Text.ToString, MAIN.TextBox171.Text.ToString, MAIN.TextBox115.Text.ToString, MAIN.TextBox119.Text.ToString, MAIN.TextBox123.Text.ToString, MAIN.TextBox126.Text.ToString, MAIN.TextBox135.Text.ToString, MAIN.TextBox144.Text.ToString, MAIN.Button43.Text.ToString, DateTime.Now.ToString, MAIN.TextBox5.Text.ToString, MAIN.Label10.Text.ToString)
                SMC.ClearSequence()
                Me.Close()
                TrckRecord()
            ElseIf TextBox2.Text = "TOOGLECMD_ON" And RadioButton2.Checked = True Then
                SMC.CalculateCount("FAIL")
                ClearReportBeforeGeneration()
                Dim lastItem As String = MAIN.ListBox1.Items(MAIN.ListBox1.Items.Count - 1).ToString()
                SMC.GenerateReportOnErrorSMCSwitchTestNode11(MAIN.TextBox3.Text, MAIN.TextBox2.Text, MAIN.TextBox4.Text, MAIN.Button43.Text, MAIN.TextBox5.Text, MAIN.Label10.Text, MAIN.TextBox153.Text, lastItem, "TRUE", "TRUE", MAIN.TextBox27.Text, MAIN.TextBox37.Text, MAIN.TextBox43.Text, MAIN.TextBox40.Text, MAIN.TextBox54.Text, MAIN.TextBox58.Text, MAIN.TextBox62.Text, MAIN.TextBox66.Text, MAIN.TextBox69.Text, MAIN.TextBox74.Text, MAIN.TextBox162.Text, MAIN.TextBox82.Text, MAIN.TextBox86.Text, MAIN.TextBox90.Text, MAIN.TextBox95.Text, MAIN.TextBox99.Text, MAIN.TextBox103.Text, MAIN.TextBox107.Text, MAIN.TextBox111.Text, MAIN.TextBox167.Text, MAIN.TextBox171.Text, MAIN.TextBox115.Text)
                SMC.SMCCreateRecord(CONFIG.TextBox1.Text, CONFIG.TextBox2.Text, CONFIG.TextBox3.Text, MAIN.TextBox4.Text.ToString, MAIN.TextBox2.Text.ToString, MAIN.TextBox3.Text.ToString, MAIN.TextBox6.Text.ToString, MAIN.TextBox153.Text.ToString, MAIN.TextBox27.Text.ToString, MAIN.TextBox37.Text.ToString, MAIN.TextBox43.Text.ToString, MAIN.TextBox40.Text.ToString, MAIN.TextBox54.Text.ToString,
                MAIN.TextBox58.Text.ToString, MAIN.TextBox62.Text.ToString, MAIN.TextBox66.Text.ToString, MAIN.TextBox69.Text.ToString, MAIN.TextBox74.Text.ToString, MAIN.TextBox162.Text.ToString, MAIN.TextBox82.Text.ToString, MAIN.TextBox86.Text.ToString, MAIN.TextBox90.Text.ToString, MAIN.TextBox95.Text.ToString, MAIN.TextBox99.Text.ToString, MAIN.TextBox103.Text.ToString, MAIN.TextBox107.Text.ToString,
                MAIN.TextBox111.Text.ToString, MAIN.TextBox167.Text.ToString, MAIN.TextBox171.Text.ToString, MAIN.TextBox115.Text.ToString, MAIN.TextBox119.Text.ToString, MAIN.TextBox123.Text.ToString, MAIN.TextBox126.Text.ToString, MAIN.TextBox135.Text.ToString, MAIN.TextBox144.Text.ToString, MAIN.Button43.Text.ToString, DateTime.Now.ToString, MAIN.TextBox5.Text.ToString, MAIN.Label10.Text.ToString)
                SMC.ClearSequence()
                Me.Close()
                TrckRecord()
            ElseIf TextBox2.Text = "TOOGLECMD_ON" And RadioButton3.Checked = True Then
                SMC.CalculateCount("FAIL")
                ClearReportBeforeGeneration()
                Dim lastItem As String = MAIN.ListBox1.Items(MAIN.ListBox1.Items.Count - 1).ToString()
                SMC.GenerateReportOnErrorSMCSwitchTestNode11(MAIN.TextBox3.Text, MAIN.TextBox2.Text, MAIN.TextBox4.Text, MAIN.Button43.Text, MAIN.TextBox5.Text, MAIN.Label10.Text, MAIN.TextBox153.Text, lastItem, "TRUE", "TRUE", MAIN.TextBox27.Text, MAIN.TextBox37.Text, MAIN.TextBox43.Text, MAIN.TextBox40.Text, MAIN.TextBox54.Text, MAIN.TextBox58.Text, MAIN.TextBox62.Text, MAIN.TextBox66.Text, MAIN.TextBox69.Text, MAIN.TextBox74.Text, MAIN.TextBox162.Text, MAIN.TextBox82.Text, MAIN.TextBox86.Text, MAIN.TextBox90.Text, MAIN.TextBox95.Text, MAIN.TextBox99.Text, MAIN.TextBox103.Text, MAIN.TextBox107.Text, MAIN.TextBox111.Text, MAIN.TextBox167.Text, MAIN.TextBox171.Text, MAIN.TextBox115.Text)
                SMC.SMCCreateRecord(CONFIG.TextBox1.Text, CONFIG.TextBox2.Text, CONFIG.TextBox3.Text, MAIN.TextBox4.Text.ToString, MAIN.TextBox2.Text.ToString, MAIN.TextBox3.Text.ToString, MAIN.TextBox6.Text.ToString, MAIN.TextBox153.Text.ToString, MAIN.TextBox27.Text.ToString, MAIN.TextBox37.Text.ToString, MAIN.TextBox43.Text.ToString, MAIN.TextBox40.Text.ToString, MAIN.TextBox54.Text.ToString,
                MAIN.TextBox58.Text.ToString, MAIN.TextBox62.Text.ToString, MAIN.TextBox66.Text.ToString, MAIN.TextBox69.Text.ToString, MAIN.TextBox74.Text.ToString, MAIN.TextBox162.Text.ToString, MAIN.TextBox82.Text.ToString, MAIN.TextBox86.Text.ToString, MAIN.TextBox90.Text.ToString, MAIN.TextBox95.Text.ToString, MAIN.TextBox99.Text.ToString, MAIN.TextBox103.Text.ToString, MAIN.TextBox107.Text.ToString,
                MAIN.TextBox111.Text.ToString, MAIN.TextBox167.Text.ToString, MAIN.TextBox171.Text.ToString, MAIN.TextBox115.Text.ToString, MAIN.TextBox119.Text.ToString, MAIN.TextBox123.Text.ToString, MAIN.TextBox126.Text.ToString, MAIN.TextBox135.Text.ToString, MAIN.TextBox144.Text.ToString, MAIN.Button43.Text.ToString, DateTime.Now.ToString, MAIN.TextBox5.Text.ToString, MAIN.Label10.Text.ToString)
                SMC.ClearSequence()
                Me.Close()
                TrckRecord()
            ElseIf TextBox2.Text = "CAN1SILENCE_ON" And RadioButton2.Checked = True Then
                SMC.CalculateCount("FAIL")
                ClearReportBeforeGeneration()
                Dim lastItem As String = MAIN.ListBox1.Items(MAIN.ListBox1.Items.Count - 1).ToString()
                SMC.GenerateReportOnErrorSMCSwitchTestNode12(MAIN.TextBox3.Text, MAIN.TextBox2.Text, MAIN.TextBox4.Text, MAIN.Button43.Text, MAIN.TextBox5.Text, MAIN.Label10.Text, MAIN.TextBox153.Text, lastItem, "TRUE", "TRUE", MAIN.TextBox27.Text, MAIN.TextBox37.Text, MAIN.TextBox43.Text, MAIN.TextBox40.Text, MAIN.TextBox54.Text, MAIN.TextBox58.Text, MAIN.TextBox62.Text, MAIN.TextBox66.Text, MAIN.TextBox69.Text, MAIN.TextBox74.Text, MAIN.TextBox162.Text, MAIN.TextBox82.Text, MAIN.TextBox86.Text, MAIN.TextBox90.Text, MAIN.TextBox95.Text, MAIN.TextBox99.Text, MAIN.TextBox103.Text, MAIN.TextBox107.Text, MAIN.TextBox111.Text, MAIN.TextBox167.Text, MAIN.TextBox171.Text, MAIN.TextBox115.Text, MAIN.TextBox119.Text)
                SMC.SMCCreateRecord(CONFIG.TextBox1.Text, CONFIG.TextBox2.Text, CONFIG.TextBox3.Text, MAIN.TextBox4.Text.ToString, MAIN.TextBox2.Text.ToString, MAIN.TextBox3.Text.ToString, MAIN.TextBox6.Text.ToString, MAIN.TextBox153.Text.ToString, MAIN.TextBox27.Text.ToString, MAIN.TextBox37.Text.ToString, MAIN.TextBox43.Text.ToString, MAIN.TextBox40.Text.ToString, MAIN.TextBox54.Text.ToString,
                MAIN.TextBox58.Text.ToString, MAIN.TextBox62.Text.ToString, MAIN.TextBox66.Text.ToString, MAIN.TextBox69.Text.ToString, MAIN.TextBox74.Text.ToString, MAIN.TextBox162.Text.ToString, MAIN.TextBox82.Text.ToString, MAIN.TextBox86.Text.ToString, MAIN.TextBox90.Text.ToString, MAIN.TextBox95.Text.ToString, MAIN.TextBox99.Text.ToString, MAIN.TextBox103.Text.ToString, MAIN.TextBox107.Text.ToString,
                MAIN.TextBox111.Text.ToString, MAIN.TextBox167.Text.ToString, MAIN.TextBox171.Text.ToString, MAIN.TextBox115.Text.ToString, MAIN.TextBox119.Text.ToString, MAIN.TextBox123.Text.ToString, MAIN.TextBox126.Text.ToString, MAIN.TextBox135.Text.ToString, MAIN.TextBox144.Text.ToString, MAIN.Button43.Text.ToString, DateTime.Now.ToString, MAIN.TextBox5.Text.ToString, MAIN.Label10.Text.ToString)
                SMC.ClearSequence()
                Me.Close()
                TrckRecord()
            ElseIf TextBox2.Text = "CAN1SILENCE_ON" And RadioButton3.Checked = True Then
                SMC.CalculateCount("FAIL")
                ClearReportBeforeGeneration()
                Dim lastItem As String = MAIN.ListBox1.Items(MAIN.ListBox1.Items.Count - 1).ToString()
                SMC.GenerateReportOnErrorSMCSwitchTestNode12(MAIN.TextBox3.Text, MAIN.TextBox2.Text, MAIN.TextBox4.Text, MAIN.Button43.Text, MAIN.TextBox5.Text, MAIN.Label10.Text, MAIN.TextBox153.Text, lastItem, "TRUE", "TRUE", MAIN.TextBox27.Text, MAIN.TextBox37.Text, MAIN.TextBox43.Text, MAIN.TextBox40.Text, MAIN.TextBox54.Text, MAIN.TextBox58.Text, MAIN.TextBox62.Text, MAIN.TextBox66.Text, MAIN.TextBox69.Text, MAIN.TextBox74.Text, MAIN.TextBox162.Text, MAIN.TextBox82.Text, MAIN.TextBox86.Text, MAIN.TextBox90.Text, MAIN.TextBox95.Text, MAIN.TextBox99.Text, MAIN.TextBox103.Text, MAIN.TextBox107.Text, MAIN.TextBox111.Text, MAIN.TextBox167.Text, MAIN.TextBox171.Text, MAIN.TextBox115.Text, MAIN.TextBox119.Text)
                SMC.SMCCreateRecord(CONFIG.TextBox1.Text, CONFIG.TextBox2.Text, CONFIG.TextBox3.Text, MAIN.TextBox4.Text.ToString, MAIN.TextBox2.Text.ToString, MAIN.TextBox3.Text.ToString, MAIN.TextBox6.Text.ToString, MAIN.TextBox153.Text.ToString, MAIN.TextBox27.Text.ToString, MAIN.TextBox37.Text.ToString, MAIN.TextBox43.Text.ToString, MAIN.TextBox40.Text.ToString, MAIN.TextBox54.Text.ToString,
                MAIN.TextBox58.Text.ToString, MAIN.TextBox62.Text.ToString, MAIN.TextBox66.Text.ToString, MAIN.TextBox69.Text.ToString, MAIN.TextBox74.Text.ToString, MAIN.TextBox162.Text.ToString, MAIN.TextBox82.Text.ToString, MAIN.TextBox86.Text.ToString, MAIN.TextBox90.Text.ToString, MAIN.TextBox95.Text.ToString, MAIN.TextBox99.Text.ToString, MAIN.TextBox103.Text.ToString, MAIN.TextBox107.Text.ToString,
                MAIN.TextBox111.Text.ToString, MAIN.TextBox167.Text.ToString, MAIN.TextBox171.Text.ToString, MAIN.TextBox115.Text.ToString, MAIN.TextBox119.Text.ToString, MAIN.TextBox123.Text.ToString, MAIN.TextBox126.Text.ToString, MAIN.TextBox135.Text.ToString, MAIN.TextBox144.Text.ToString, MAIN.Button43.Text.ToString, DateTime.Now.ToString, MAIN.TextBox5.Text.ToString, MAIN.Label10.Text.ToString)
                SMC.ClearSequence()
                Me.Close()
                TrckRecord()
            ElseIf TextBox2.Text = "CAN2SILENCE_ON" And RadioButton2.Checked = True Then
                SMC.CalculateCount("FAIL")
                ClearReportBeforeGeneration()
                Dim lastItem As String = MAIN.ListBox1.Items(MAIN.ListBox1.Items.Count - 1).ToString()
                SMC.GenerateReportOnErrorSMCSwitchTestNode13(MAIN.TextBox3.Text, MAIN.TextBox2.Text, MAIN.TextBox4.Text, MAIN.Button43.Text, MAIN.TextBox5.Text, MAIN.Label10.Text, MAIN.TextBox153.Text, lastItem, "TRUE", "TRUE", MAIN.TextBox27.Text, MAIN.TextBox37.Text, MAIN.TextBox43.Text, MAIN.TextBox40.Text, MAIN.TextBox54.Text, MAIN.TextBox58.Text, MAIN.TextBox62.Text, MAIN.TextBox66.Text, MAIN.TextBox69.Text, MAIN.TextBox74.Text, MAIN.TextBox162.Text, MAIN.TextBox82.Text, MAIN.TextBox86.Text, MAIN.TextBox90.Text, MAIN.TextBox95.Text, MAIN.TextBox99.Text, MAIN.TextBox103.Text, MAIN.TextBox107.Text, MAIN.TextBox111.Text, MAIN.TextBox167.Text, MAIN.TextBox171.Text, MAIN.TextBox115.Text, MAIN.TextBox119.Text, MAIN.TextBox123.Text)
                SMC.SMCCreateRecord(CONFIG.TextBox1.Text, CONFIG.TextBox2.Text, CONFIG.TextBox3.Text, MAIN.TextBox4.Text.ToString, MAIN.TextBox2.Text.ToString, MAIN.TextBox3.Text.ToString, MAIN.TextBox6.Text.ToString, MAIN.TextBox153.Text.ToString, MAIN.TextBox27.Text.ToString, MAIN.TextBox37.Text.ToString, MAIN.TextBox43.Text.ToString, MAIN.TextBox40.Text.ToString, MAIN.TextBox54.Text.ToString,
                MAIN.TextBox58.Text.ToString, MAIN.TextBox62.Text.ToString, MAIN.TextBox66.Text.ToString, MAIN.TextBox69.Text.ToString, MAIN.TextBox74.Text.ToString, MAIN.TextBox162.Text.ToString, MAIN.TextBox82.Text.ToString, MAIN.TextBox86.Text.ToString, MAIN.TextBox90.Text.ToString, MAIN.TextBox95.Text.ToString, MAIN.TextBox99.Text.ToString, MAIN.TextBox103.Text.ToString, MAIN.TextBox107.Text.ToString,
                MAIN.TextBox111.Text.ToString, MAIN.TextBox167.Text.ToString, MAIN.TextBox171.Text.ToString, MAIN.TextBox115.Text.ToString, MAIN.TextBox119.Text.ToString, MAIN.TextBox123.Text.ToString, MAIN.TextBox126.Text.ToString, MAIN.TextBox135.Text.ToString, MAIN.TextBox144.Text.ToString, MAIN.Button43.Text.ToString, DateTime.Now.ToString, MAIN.TextBox5.Text.ToString, MAIN.Label10.Text.ToString)
                SMC.ClearSequence()
                Me.Close()
                TrckRecord()
            ElseIf TextBox2.Text = "CAN2SILENCE_ON" And RadioButton3.Checked = True Then
                SMC.CalculateCount("FAIL")
                ClearReportBeforeGeneration()
                Dim lastItem As String = MAIN.ListBox1.Items(MAIN.ListBox1.Items.Count - 1).ToString()
                SMC.GenerateReportOnErrorSMCSwitchTestNode13(MAIN.TextBox3.Text, MAIN.TextBox2.Text, MAIN.TextBox4.Text, MAIN.Button43.Text, MAIN.TextBox5.Text, MAIN.Label10.Text, MAIN.TextBox153.Text, lastItem, "TRUE", "TRUE", MAIN.TextBox27.Text, MAIN.TextBox37.Text, MAIN.TextBox43.Text, MAIN.TextBox40.Text, MAIN.TextBox54.Text, MAIN.TextBox58.Text, MAIN.TextBox62.Text, MAIN.TextBox66.Text, MAIN.TextBox69.Text, MAIN.TextBox74.Text, MAIN.TextBox162.Text, MAIN.TextBox82.Text, MAIN.TextBox86.Text, MAIN.TextBox90.Text, MAIN.TextBox95.Text, MAIN.TextBox99.Text, MAIN.TextBox103.Text, MAIN.TextBox107.Text, MAIN.TextBox111.Text, MAIN.TextBox167.Text, MAIN.TextBox171.Text, MAIN.TextBox115.Text, MAIN.TextBox119.Text, MAIN.TextBox123.Text)
                SMC.SMCCreateRecord(CONFIG.TextBox1.Text, CONFIG.TextBox2.Text, CONFIG.TextBox3.Text, MAIN.TextBox4.Text.ToString, MAIN.TextBox2.Text.ToString, MAIN.TextBox3.Text.ToString, MAIN.TextBox6.Text.ToString, MAIN.TextBox153.Text.ToString, MAIN.TextBox27.Text.ToString, MAIN.TextBox37.Text.ToString, MAIN.TextBox43.Text.ToString, MAIN.TextBox40.Text.ToString, MAIN.TextBox54.Text.ToString,
                MAIN.TextBox58.Text.ToString, MAIN.TextBox62.Text.ToString, MAIN.TextBox66.Text.ToString, MAIN.TextBox69.Text.ToString, MAIN.TextBox74.Text.ToString, MAIN.TextBox162.Text.ToString, MAIN.TextBox82.Text.ToString, MAIN.TextBox86.Text.ToString, MAIN.TextBox90.Text.ToString, MAIN.TextBox95.Text.ToString, MAIN.TextBox99.Text.ToString, MAIN.TextBox103.Text.ToString, MAIN.TextBox107.Text.ToString,
                MAIN.TextBox111.Text.ToString, MAIN.TextBox167.Text.ToString, MAIN.TextBox171.Text.ToString, MAIN.TextBox115.Text.ToString, MAIN.TextBox119.Text.ToString, MAIN.TextBox123.Text.ToString, MAIN.TextBox126.Text.ToString, MAIN.TextBox135.Text.ToString, MAIN.TextBox144.Text.ToString, MAIN.Button43.Text.ToString, DateTime.Now.ToString, MAIN.TextBox5.Text.ToString, MAIN.Label10.Text.ToString)
                SMC.ClearSequence()
                Me.Close()
                TrckRecord()
            ElseIf TextBox2.Text = "TIME_TEST" And RadioButton2.Checked = True Then
                SMC.CalculateCount("FAIL")
                ClearReportBeforeGeneration()
                Dim lastItem As String = MAIN.ListBox1.Items(MAIN.ListBox1.Items.Count - 1).ToString()
                SMC.GenerateReportOnErrorSMCTimeNode0(MAIN.TextBox3.Text, MAIN.TextBox2.Text, MAIN.TextBox4.Text, MAIN.Button43.Text, MAIN.TextBox5.Text, MAIN.Label10.Text, MAIN.TextBox153.Text, lastItem, "TRUE", "TRUE", MAIN.TextBox27.Text, MAIN.TextBox37.Text, MAIN.TextBox43.Text, MAIN.TextBox40.Text, MAIN.TextBox54.Text, MAIN.TextBox58.Text, MAIN.TextBox62.Text, MAIN.TextBox66.Text, MAIN.TextBox69.Text, MAIN.TextBox74.Text, MAIN.TextBox162.Text, MAIN.TextBox82.Text, MAIN.TextBox86.Text, MAIN.TextBox90.Text, MAIN.TextBox95.Text, MAIN.TextBox99.Text, MAIN.TextBox103.Text, MAIN.TextBox107.Text, MAIN.TextBox111.Text, MAIN.TextBox167.Text, MAIN.TextBox171.Text, MAIN.TextBox115.Text, MAIN.TextBox119.Text, MAIN.TextBox123.Text, MAIN.TextBox126.Text)
                SMC.SMCCreateRecord(CONFIG.TextBox1.Text, CONFIG.TextBox2.Text, CONFIG.TextBox3.Text, MAIN.TextBox4.Text.ToString, MAIN.TextBox2.Text.ToString, MAIN.TextBox3.Text.ToString, MAIN.TextBox6.Text.ToString, MAIN.TextBox153.Text.ToString, MAIN.TextBox27.Text.ToString, MAIN.TextBox37.Text.ToString, MAIN.TextBox43.Text.ToString, MAIN.TextBox40.Text.ToString, MAIN.TextBox54.Text.ToString,
                MAIN.TextBox58.Text.ToString, MAIN.TextBox62.Text.ToString, MAIN.TextBox66.Text.ToString, MAIN.TextBox69.Text.ToString, MAIN.TextBox74.Text.ToString, MAIN.TextBox162.Text.ToString, MAIN.TextBox82.Text.ToString, MAIN.TextBox86.Text.ToString, MAIN.TextBox90.Text.ToString, MAIN.TextBox95.Text.ToString, MAIN.TextBox99.Text.ToString, MAIN.TextBox103.Text.ToString, MAIN.TextBox107.Text.ToString,
                MAIN.TextBox111.Text.ToString, MAIN.TextBox167.Text.ToString, MAIN.TextBox171.Text.ToString, MAIN.TextBox115.Text.ToString, MAIN.TextBox119.Text.ToString, MAIN.TextBox123.Text.ToString, MAIN.TextBox126.Text.ToString, MAIN.TextBox135.Text.ToString, MAIN.TextBox144.Text.ToString, MAIN.Button43.Text.ToString, DateTime.Now.ToString, MAIN.TextBox5.Text.ToString, MAIN.Label10.Text.ToString)
                SMC.ClearSequence()
                Me.Close()
                TrckRecord()
            ElseIf TextBox2.Text = "TIME_TEST" And RadioButton3.Checked = True Then
                SMC.CalculateCount("FAIL")
                ClearReportBeforeGeneration()
                Dim lastItem As String = MAIN.ListBox1.Items(MAIN.ListBox1.Items.Count - 1).ToString()
                SMC.GenerateReportOnErrorSMCTimeNode0(MAIN.TextBox3.Text, MAIN.TextBox2.Text, MAIN.TextBox4.Text, MAIN.Button43.Text, MAIN.TextBox5.Text, MAIN.Label10.Text, MAIN.TextBox153.Text, lastItem, "TRUE", "TRUE", MAIN.TextBox27.Text, MAIN.TextBox37.Text, MAIN.TextBox43.Text, MAIN.TextBox40.Text, MAIN.TextBox54.Text, MAIN.TextBox58.Text, MAIN.TextBox62.Text, MAIN.TextBox66.Text, MAIN.TextBox69.Text, MAIN.TextBox74.Text, MAIN.TextBox162.Text, MAIN.TextBox82.Text, MAIN.TextBox86.Text, MAIN.TextBox90.Text, MAIN.TextBox95.Text, MAIN.TextBox99.Text, MAIN.TextBox103.Text, MAIN.TextBox107.Text, MAIN.TextBox111.Text, MAIN.TextBox167.Text, MAIN.TextBox171.Text, MAIN.TextBox115.Text, MAIN.TextBox119.Text, MAIN.TextBox123.Text, MAIN.TextBox126.Text)
                SMC.SMCCreateRecord(CONFIG.TextBox1.Text, CONFIG.TextBox2.Text, CONFIG.TextBox3.Text, MAIN.TextBox4.Text.ToString, MAIN.TextBox2.Text.ToString, MAIN.TextBox3.Text.ToString, MAIN.TextBox6.Text.ToString, MAIN.TextBox153.Text.ToString, MAIN.TextBox27.Text.ToString, MAIN.TextBox37.Text.ToString, MAIN.TextBox43.Text.ToString, MAIN.TextBox40.Text.ToString, MAIN.TextBox54.Text.ToString,
                MAIN.TextBox58.Text.ToString, MAIN.TextBox62.Text.ToString, MAIN.TextBox66.Text.ToString, MAIN.TextBox69.Text.ToString, MAIN.TextBox74.Text.ToString, MAIN.TextBox162.Text.ToString, MAIN.TextBox82.Text.ToString, MAIN.TextBox86.Text.ToString, MAIN.TextBox90.Text.ToString, MAIN.TextBox95.Text.ToString, MAIN.TextBox99.Text.ToString, MAIN.TextBox103.Text.ToString, MAIN.TextBox107.Text.ToString,
                MAIN.TextBox111.Text.ToString, MAIN.TextBox167.Text.ToString, MAIN.TextBox171.Text.ToString, MAIN.TextBox115.Text.ToString, MAIN.TextBox119.Text.ToString, MAIN.TextBox123.Text.ToString, MAIN.TextBox126.Text.ToString, MAIN.TextBox135.Text.ToString, MAIN.TextBox144.Text.ToString, MAIN.Button43.Text.ToString, DateTime.Now.ToString, MAIN.TextBox5.Text.ToString, MAIN.Label10.Text.ToString)
                SMC.ClearSequence()
                Me.Close()
                TrckRecord()
            ElseIf TextBox2.Text = "STEPPER_TEST" And RadioButton2.Checked = True Then
                SMC.CalculateCount("FAIL")
                ClearReportBeforeGeneration()
                Dim lastItem As String = MAIN.ListBox1.Items(MAIN.ListBox1.Items.Count - 1).ToString()
                SMC.GenerateReportOnErrorSMCStepperTestNode0(MAIN.TextBox3.Text, MAIN.TextBox2.Text, MAIN.TextBox4.Text, MAIN.Button43.Text, MAIN.TextBox5.Text, MAIN.Label10.Text, MAIN.TextBox153.Text, lastItem, "TRUE", "TRUE", MAIN.TextBox27.Text, MAIN.TextBox37.Text, MAIN.TextBox43.Text, MAIN.TextBox40.Text, MAIN.TextBox54.Text, MAIN.TextBox58.Text, MAIN.TextBox62.Text, MAIN.TextBox66.Text, MAIN.TextBox69.Text, MAIN.TextBox74.Text, MAIN.TextBox162.Text, MAIN.TextBox82.Text, MAIN.TextBox86.Text, MAIN.TextBox90.Text, MAIN.TextBox95.Text, MAIN.TextBox99.Text, MAIN.TextBox103.Text, MAIN.TextBox107.Text, MAIN.TextBox111.Text, MAIN.TextBox167.Text, MAIN.TextBox171.Text, MAIN.TextBox115.Text, MAIN.TextBox119.Text, MAIN.TextBox123.Text, MAIN.TextBox126.Text, MAIN.TextBox135.Text)
                SMC.SMCCreateRecord(CONFIG.TextBox1.Text, CONFIG.TextBox2.Text, CONFIG.TextBox3.Text, MAIN.TextBox4.Text.ToString, MAIN.TextBox2.Text.ToString, MAIN.TextBox3.Text.ToString, MAIN.TextBox6.Text.ToString, MAIN.TextBox153.Text.ToString, MAIN.TextBox27.Text.ToString, MAIN.TextBox37.Text.ToString, MAIN.TextBox43.Text.ToString, MAIN.TextBox40.Text.ToString, MAIN.TextBox54.Text.ToString,
                MAIN.TextBox58.Text.ToString, MAIN.TextBox62.Text.ToString, MAIN.TextBox66.Text.ToString, MAIN.TextBox69.Text.ToString, MAIN.TextBox74.Text.ToString, MAIN.TextBox162.Text.ToString, MAIN.TextBox82.Text.ToString, MAIN.TextBox86.Text.ToString, MAIN.TextBox90.Text.ToString, MAIN.TextBox95.Text.ToString, MAIN.TextBox99.Text.ToString, MAIN.TextBox103.Text.ToString, MAIN.TextBox107.Text.ToString,
                MAIN.TextBox111.Text.ToString, MAIN.TextBox167.Text.ToString, MAIN.TextBox171.Text.ToString, MAIN.TextBox115.Text.ToString, MAIN.TextBox119.Text.ToString, MAIN.TextBox123.Text.ToString, MAIN.TextBox126.Text.ToString, MAIN.TextBox135.Text.ToString, MAIN.TextBox144.Text.ToString, MAIN.Button43.Text.ToString, DateTime.Now.ToString, MAIN.TextBox5.Text.ToString, MAIN.Label10.Text.ToString)
                SMC.ClearSequence()
                Me.Close()
                TrckRecord()
            ElseIf TextBox2.Text = "STEPPER_TEST" And RadioButton3.Checked = True Then
                SMC.CalculateCount("FAIL")
                ClearReportBeforeGeneration()
                Dim lastItem As String = MAIN.ListBox1.Items(MAIN.ListBox1.Items.Count - 1).ToString()
                SMC.GenerateReportOnErrorSMCStepperTestNode0(MAIN.TextBox3.Text, MAIN.TextBox2.Text, MAIN.TextBox4.Text, MAIN.Button43.Text, MAIN.TextBox5.Text, MAIN.Label10.Text, MAIN.TextBox153.Text, lastItem, "TRUE", "TRUE", MAIN.TextBox27.Text, MAIN.TextBox37.Text, MAIN.TextBox43.Text, MAIN.TextBox40.Text, MAIN.TextBox54.Text, MAIN.TextBox58.Text, MAIN.TextBox62.Text, MAIN.TextBox66.Text, MAIN.TextBox69.Text, MAIN.TextBox74.Text, MAIN.TextBox162.Text, MAIN.TextBox82.Text, MAIN.TextBox86.Text, MAIN.TextBox90.Text, MAIN.TextBox95.Text, MAIN.TextBox99.Text, MAIN.TextBox103.Text, MAIN.TextBox107.Text, MAIN.TextBox111.Text, MAIN.TextBox167.Text, MAIN.TextBox171.Text, MAIN.TextBox115.Text, MAIN.TextBox119.Text, MAIN.TextBox123.Text, MAIN.TextBox126.Text, MAIN.TextBox135.Text)
                SMC.SMCCreateRecord(CONFIG.TextBox1.Text, CONFIG.TextBox2.Text, CONFIG.TextBox3.Text, MAIN.TextBox4.Text.ToString, MAIN.TextBox2.Text.ToString, MAIN.TextBox3.Text.ToString, MAIN.TextBox6.Text.ToString, MAIN.TextBox153.Text.ToString, MAIN.TextBox27.Text.ToString, MAIN.TextBox37.Text.ToString, MAIN.TextBox43.Text.ToString, MAIN.TextBox40.Text.ToString, MAIN.TextBox54.Text.ToString,
                MAIN.TextBox58.Text.ToString, MAIN.TextBox62.Text.ToString, MAIN.TextBox66.Text.ToString, MAIN.TextBox69.Text.ToString, MAIN.TextBox74.Text.ToString, MAIN.TextBox162.Text.ToString, MAIN.TextBox82.Text.ToString, MAIN.TextBox86.Text.ToString, MAIN.TextBox90.Text.ToString, MAIN.TextBox95.Text.ToString, MAIN.TextBox99.Text.ToString, MAIN.TextBox103.Text.ToString, MAIN.TextBox107.Text.ToString,
                MAIN.TextBox111.Text.ToString, MAIN.TextBox167.Text.ToString, MAIN.TextBox171.Text.ToString, MAIN.TextBox115.Text.ToString, MAIN.TextBox119.Text.ToString, MAIN.TextBox123.Text.ToString, MAIN.TextBox126.Text.ToString, MAIN.TextBox135.Text.ToString, MAIN.TextBox144.Text.ToString, MAIN.Button43.Text.ToString, DateTime.Now.ToString, MAIN.TextBox5.Text.ToString, MAIN.Label10.Text.ToString)
                SMC.ClearSequence()
                Me.Close()
                TrckRecord()
            ElseIf TextBox2.Text = "SHUTDOWN_TEST" And RadioButton2.Checked = True Then
                SMC.CalculateCount("FAIL")
                ClearReportBeforeGeneration()
                Dim lastItem As String = MAIN.ListBox1.Items(MAIN.ListBox1.Items.Count - 1).ToString()
                SMC.GenerateReportOnErrorSMCShutDownNod0(MAIN.TextBox3.Text, MAIN.TextBox2.Text, MAIN.TextBox4.Text, MAIN.Button43.Text, MAIN.TextBox5.Text, MAIN.Label10.Text, MAIN.TextBox153.Text, lastItem, "TRUE", "TRUE", MAIN.TextBox27.Text, MAIN.TextBox37.Text, MAIN.TextBox43.Text, MAIN.TextBox40.Text, MAIN.TextBox54.Text, MAIN.TextBox58.Text, MAIN.TextBox62.Text, MAIN.TextBox66.Text, MAIN.TextBox69.Text, MAIN.TextBox74.Text, MAIN.TextBox162.Text, MAIN.TextBox82.Text, MAIN.TextBox86.Text, MAIN.TextBox90.Text, MAIN.TextBox95.Text, MAIN.TextBox99.Text, MAIN.TextBox103.Text, MAIN.TextBox107.Text, MAIN.TextBox111.Text, MAIN.TextBox167.Text, MAIN.TextBox171.Text, MAIN.TextBox115.Text, MAIN.TextBox119.Text, MAIN.TextBox123.Text, MAIN.TextBox126.Text, MAIN.TextBox135.Text, MAIN.TextBox144.Text)
                SMC.SMCCreateRecord(CONFIG.TextBox1.Text, CONFIG.TextBox2.Text, CONFIG.TextBox3.Text, MAIN.TextBox4.Text.ToString, MAIN.TextBox2.Text.ToString, MAIN.TextBox3.Text.ToString, MAIN.TextBox6.Text.ToString, MAIN.TextBox153.Text.ToString, MAIN.TextBox27.Text.ToString, MAIN.TextBox37.Text.ToString, MAIN.TextBox43.Text.ToString, MAIN.TextBox40.Text.ToString, MAIN.TextBox54.Text.ToString,
                MAIN.TextBox58.Text.ToString, MAIN.TextBox62.Text.ToString, MAIN.TextBox66.Text.ToString, MAIN.TextBox69.Text.ToString, MAIN.TextBox74.Text.ToString, MAIN.TextBox162.Text.ToString, MAIN.TextBox82.Text.ToString, MAIN.TextBox86.Text.ToString, MAIN.TextBox90.Text.ToString, MAIN.TextBox95.Text.ToString, MAIN.TextBox99.Text.ToString, MAIN.TextBox103.Text.ToString, MAIN.TextBox107.Text.ToString,
                MAIN.TextBox111.Text.ToString, MAIN.TextBox167.Text.ToString, MAIN.TextBox171.Text.ToString, MAIN.TextBox115.Text.ToString, MAIN.TextBox119.Text.ToString, MAIN.TextBox123.Text.ToString, MAIN.TextBox126.Text.ToString, MAIN.TextBox135.Text.ToString, MAIN.TextBox144.Text.ToString, MAIN.Button43.Text.ToString, DateTime.Now.ToString, MAIN.TextBox5.Text.ToString, MAIN.Label10.Text.ToString)
                SMC.ClearSequence()
                Me.Close()
                TrckRecord()
            ElseIf TextBox2.Text = "SHUTDOWN_TEST" And RadioButton3.Checked = True Then
                SMC.CalculateCount("FAIL")
                ClearReportBeforeGeneration()
                Dim lastItem As String = MAIN.ListBox1.Items(MAIN.ListBox1.Items.Count - 1).ToString()
                SMC.GenerateReportOnErrorSMCShutDownNod0(MAIN.TextBox3.Text, MAIN.TextBox2.Text, MAIN.TextBox4.Text, MAIN.Button43.Text, MAIN.TextBox5.Text, MAIN.Label10.Text, MAIN.TextBox153.Text, lastItem, "TRUE", "TRUE", MAIN.TextBox27.Text, MAIN.TextBox37.Text, MAIN.TextBox43.Text, MAIN.TextBox40.Text, MAIN.TextBox54.Text, MAIN.TextBox58.Text, MAIN.TextBox62.Text, MAIN.TextBox66.Text, MAIN.TextBox69.Text, MAIN.TextBox74.Text, MAIN.TextBox162.Text, MAIN.TextBox82.Text, MAIN.TextBox86.Text, MAIN.TextBox90.Text, MAIN.TextBox95.Text, MAIN.TextBox99.Text, MAIN.TextBox103.Text, MAIN.TextBox107.Text, MAIN.TextBox111.Text, MAIN.TextBox167.Text, MAIN.TextBox171.Text, MAIN.TextBox115.Text, MAIN.TextBox119.Text, MAIN.TextBox123.Text, MAIN.TextBox126.Text, MAIN.TextBox135.Text, MAIN.TextBox144.Text)
                SMC.SMCCreateRecord(CONFIG.TextBox1.Text, CONFIG.TextBox2.Text, CONFIG.TextBox3.Text, MAIN.TextBox4.Text.ToString, MAIN.TextBox2.Text.ToString, MAIN.TextBox3.Text.ToString, MAIN.TextBox6.Text.ToString, MAIN.TextBox153.Text.ToString, MAIN.TextBox27.Text.ToString, MAIN.TextBox37.Text.ToString, MAIN.TextBox43.Text.ToString, MAIN.TextBox40.Text.ToString, MAIN.TextBox54.Text.ToString,
                MAIN.TextBox58.Text.ToString, MAIN.TextBox62.Text.ToString, MAIN.TextBox66.Text.ToString, MAIN.TextBox69.Text.ToString, MAIN.TextBox74.Text.ToString, MAIN.TextBox162.Text.ToString, MAIN.TextBox82.Text.ToString, MAIN.TextBox86.Text.ToString, MAIN.TextBox90.Text.ToString, MAIN.TextBox95.Text.ToString, MAIN.TextBox99.Text.ToString, MAIN.TextBox103.Text.ToString, MAIN.TextBox107.Text.ToString,
                MAIN.TextBox111.Text.ToString, MAIN.TextBox167.Text.ToString, MAIN.TextBox171.Text.ToString, MAIN.TextBox115.Text.ToString, MAIN.TextBox119.Text.ToString, MAIN.TextBox123.Text.ToString, MAIN.TextBox126.Text.ToString, MAIN.TextBox135.Text.ToString, MAIN.TextBox144.Text.ToString, MAIN.Button43.Text.ToString, DateTime.Now.ToString, MAIN.TextBox5.Text.ToString, MAIN.Label10.Text.ToString)
                SMC.ClearSequence()
                Me.Close()
                TrckRecord()
            ElseIf TextBox2.Text = "SHUTDOWN_TEST" And RadioButton1.Checked = True Then
                SMC.CalculateCount("PASS")
                ClearReportBeforeGeneration()
                Dim lastItem As String = MAIN.ListBox1.Items(MAIN.ListBox1.Items.Count - 1).ToString()
                SMC.GenerateReport(MAIN.TextBox3.Text, MAIN.TextBox2.Text, MAIN.TextBox4.Text, MAIN.Button43.Text, MAIN.TextBox5.Text, MAIN.Label10.Text, MAIN.TextBox153.Text, lastItem, "TRUE", "TRUE", MAIN.TextBox27.Text, MAIN.TextBox37.Text, MAIN.TextBox43.Text, MAIN.TextBox40.Text, MAIN.TextBox54.Text, MAIN.TextBox58.Text, MAIN.TextBox62.Text, MAIN.TextBox66.Text, MAIN.TextBox69.Text, MAIN.TextBox74.Text, MAIN.TextBox162.Text, MAIN.TextBox82.Text, MAIN.TextBox86.Text, MAIN.TextBox90.Text, MAIN.TextBox95.Text, MAIN.TextBox99.Text, MAIN.TextBox103.Text, MAIN.TextBox107.Text, MAIN.TextBox111.Text, MAIN.TextBox167.Text, MAIN.TextBox171.Text, MAIN.TextBox115.Text, MAIN.TextBox119.Text, MAIN.TextBox123.Text, MAIN.TextBox126.Text, MAIN.TextBox135.Text, MAIN.TextBox144.Text)
                SMC.SMCCreateRecord(CONFIG.TextBox1.Text, CONFIG.TextBox2.Text, CONFIG.TextBox3.Text, MAIN.TextBox4.Text.ToString, MAIN.TextBox2.Text.ToString, MAIN.TextBox3.Text.ToString, MAIN.TextBox6.Text.ToString, MAIN.TextBox153.Text.ToString, MAIN.TextBox27.Text.ToString, MAIN.TextBox37.Text.ToString, MAIN.TextBox43.Text.ToString, MAIN.TextBox40.Text.ToString, MAIN.TextBox54.Text.ToString,
                MAIN.TextBox58.Text.ToString, MAIN.TextBox62.Text.ToString, MAIN.TextBox66.Text.ToString, MAIN.TextBox69.Text.ToString, MAIN.TextBox74.Text.ToString, MAIN.TextBox162.Text.ToString, MAIN.TextBox82.Text.ToString, MAIN.TextBox86.Text.ToString, MAIN.TextBox90.Text.ToString, MAIN.TextBox95.Text.ToString, MAIN.TextBox99.Text.ToString, MAIN.TextBox103.Text.ToString, MAIN.TextBox107.Text.ToString,
                MAIN.TextBox111.Text.ToString, MAIN.TextBox167.Text.ToString, MAIN.TextBox171.Text.ToString, MAIN.TextBox115.Text.ToString, MAIN.TextBox119.Text.ToString, MAIN.TextBox123.Text.ToString, MAIN.TextBox126.Text.ToString, MAIN.TextBox135.Text.ToString, MAIN.TextBox144.Text.ToString, MAIN.Button43.Text.ToString, DateTime.Now.ToString, MAIN.TextBox5.Text.ToString, MAIN.Label10.Text.ToString)
                SMC.ClearSequence()
                Me.Close()
                TrckRecord()
            End If
        ElseIf TextBox1.Text = "SMC" And TextBox3.Text = "GOLDEN" Then
            SMC.ClearSequence()
            Me.Close()
            TrckRecord()
        ElseIf TextBox1.Text = "BATTERY ID" And TextBox3.Text = "PRODUCTION" Then
            If TextBox2.Text = "BATTERYID_INIT" And RadioButton2.Checked = True Then
                MAIN.TimerFunction("STOP")
                BatteryID.CalculateCount("FAIL")
                BatteryID.ClearReportBeforeLogGeberate()
                Dim lastItem As String = MAIN.ListBox1.Items(MAIN.ListBox1.Items.Count - 1).ToString()
                BatteryID.GenerateReportOnErrorOfInit(MAIN.TextBox3.Text, MAIN.TextBox2.Text, MAIN.TextBox4.Text, MAIN.Button43.Text, MAIN.TextBox5.Text, MAIN.Label10.Text, MAIN.TextBox18.Text, lastItem)
                BatteryID.GenerateLogReport(CONFIG.TextBox28.Text, MAIN.TextBox4.Text)
                BatteryID.CreateRecord(CONFIG.TextBox1.Text, CONFIG.TextBox2.Text, CONFIG.TextBox3.Text, MAIN.TextBox4.Text, MAIN.TextBox2.Text, MAIN.TextBox3.Text, MAIN.TextBox6.Text, MAIN.TextBox153.Text, MAIN.TextBox18.Text, MAIN.Button43.Text, MAIN.TextBox5.Text, MAIN.Label10.Text)
                BatteryID.ClearSequence()
                TrckRecord()
                Me.Close()
            ElseIf TextBox2.Text = "BATTERYID_INIT" And RadioButton3.Checked = True Then
                MAIN.TimerFunction("STOP")
                BatteryID.CalculateCount("FAIL")
                BatteryID.ClearReportBeforeLogGeberate()
                Dim lastItem As String = MAIN.ListBox1.Items(MAIN.ListBox1.Items.Count - 1).ToString()
                BatteryID.GenerateReportOnErrorOfInit(MAIN.TextBox3.Text, MAIN.TextBox2.Text, MAIN.TextBox4.Text, MAIN.Button43.Text, MAIN.TextBox5.Text, MAIN.Label10.Text, MAIN.TextBox18.Text, lastItem)
                BatteryID.GenerateLogReport(CONFIG.TextBox28.Text, MAIN.TextBox4.Text)
                BatteryID.CreateRecord(CONFIG.TextBox1.Text, CONFIG.TextBox2.Text, CONFIG.TextBox3.Text, MAIN.TextBox4.Text, MAIN.TextBox2.Text, MAIN.TextBox3.Text, MAIN.TextBox6.Text, MAIN.TextBox153.Text, MAIN.TextBox18.Text, MAIN.Button43.Text, MAIN.TextBox5.Text, MAIN.Label10.Text)
                BatteryID.ClearSequence()
                TrckRecord()
                Me.Close()
            ElseIf TextBox2.Text = "BATTERYID_NODE1" And RadioButton2.Checked = True Then
                MAIN.TimerFunction("STOP")
                BatteryID.CalculateCount("FAIL")
                BatteryID.ClearReportBeforeLogGeberate()
                Dim lastItem As String = MAIN.ListBox1.Items(MAIN.ListBox1.Items.Count - 1).ToString()
                BatteryID.GenerateReportOnErrorOfBatteryIDNode1(MAIN.TextBox3.Text, MAIN.TextBox2.Text, MAIN.TextBox4.Text, MAIN.Button43.Text, MAIN.TextBox5.Text, MAIN.Label10.Text, MAIN.TextBox18.Text, lastItem)
                BatteryID.GenerateLogReport(CONFIG.TextBox28.Text, MAIN.TextBox4.Text)
                BatteryID.CreateRecord(CONFIG.TextBox1.Text, CONFIG.TextBox2.Text, CONFIG.TextBox3.Text, MAIN.TextBox4.Text, MAIN.TextBox2.Text, MAIN.TextBox3.Text, MAIN.TextBox6.Text, MAIN.TextBox153.Text, MAIN.TextBox18.Text, MAIN.Button43.Text, MAIN.TextBox5.Text, MAIN.Label10.Text)
                BatteryID.ClearSequence()
                TrckRecord()
                Me.Close()
            ElseIf TextBox2.Text = "BATTERYID_NODE1" And RadioButton3.Checked = True Then
                MAIN.TimerFunction("STOP")
                BatteryID.CalculateCount("FAIL")
                BatteryID.ClearReportBeforeLogGeberate()
                Dim lastItem As String = MAIN.ListBox1.Items(MAIN.ListBox1.Items.Count - 1).ToString()
                BatteryID.GenerateReportOnErrorOfBatteryIDNode1(MAIN.TextBox3.Text, MAIN.TextBox2.Text, MAIN.TextBox4.Text, MAIN.Button43.Text, MAIN.TextBox5.Text, MAIN.Label10.Text, MAIN.TextBox18.Text, lastItem)
                BatteryID.GenerateLogReport(CONFIG.TextBox28.Text, MAIN.TextBox4.Text)
                BatteryID.CreateRecord(CONFIG.TextBox1.Text, CONFIG.TextBox2.Text, CONFIG.TextBox3.Text, MAIN.TextBox4.Text, MAIN.TextBox2.Text, MAIN.TextBox3.Text, MAIN.TextBox6.Text, MAIN.TextBox153.Text, MAIN.TextBox18.Text, MAIN.Button43.Text, MAIN.TextBox5.Text, MAIN.Label10.Text)
                BatteryID.ClearSequence()
                TrckRecord()
                Me.Close()
            ElseIf TextBox2.Text = "BATTERYID_NODE2" And RadioButton2.Checked = True Then
                MAIN.TimerFunction("STOP")
                BatteryID.CalculateCount("FAIL")
                BatteryID.ClearReportBeforeLogGeberate()
                Dim lastItem As String = MAIN.ListBox1.Items(MAIN.ListBox1.Items.Count - 1).ToString()
                BatteryID.GenerateReportOnErrorOfBatteryIDNode2(MAIN.TextBox3.Text, MAIN.TextBox2.Text, MAIN.TextBox4.Text, MAIN.Button43.Text, MAIN.TextBox5.Text, MAIN.Label10.Text, MAIN.TextBox18.Text, lastItem)
                BatteryID.GenerateLogReport(CONFIG.TextBox28.Text, MAIN.TextBox4.Text)
                BatteryID.CreateRecord(CONFIG.TextBox1.Text, CONFIG.TextBox2.Text, CONFIG.TextBox3.Text, MAIN.TextBox4.Text, MAIN.TextBox2.Text, MAIN.TextBox3.Text, MAIN.TextBox6.Text, MAIN.TextBox153.Text, MAIN.TextBox18.Text, MAIN.Button43.Text, MAIN.TextBox5.Text, MAIN.Label10.Text)
                BatteryID.ClearSequence()
                TrckRecord()
                Me.Close()
            ElseIf TextBox2.Text = "BATTERYID_NODE2" And RadioButton3.Checked = True Then
                MAIN.TimerFunction("STOP")
                BatteryID.CalculateCount("FAIL")
                BatteryID.ClearReportBeforeLogGeberate()
                Dim lastItem As String = MAIN.ListBox1.Items(MAIN.ListBox1.Items.Count - 1).ToString()
                BatteryID.GenerateReportOnErrorOfBatteryIDNode2(MAIN.TextBox3.Text, MAIN.TextBox2.Text, MAIN.TextBox4.Text, MAIN.Button43.Text, MAIN.TextBox5.Text, MAIN.Label10.Text, MAIN.TextBox18.Text, lastItem)
                BatteryID.GenerateLogReport(CONFIG.TextBox28.Text, MAIN.TextBox4.Text)
                BatteryID.CreateRecord(CONFIG.TextBox1.Text, CONFIG.TextBox2.Text, CONFIG.TextBox3.Text, MAIN.TextBox4.Text, MAIN.TextBox2.Text, MAIN.TextBox3.Text, MAIN.TextBox6.Text, MAIN.TextBox153.Text, MAIN.TextBox18.Text, MAIN.Button43.Text, MAIN.TextBox5.Text, MAIN.Label10.Text)
                BatteryID.ClearSequence()
                TrckRecord()
                Me.Close()
            ElseIf TextBox2.Text = "BATTERYID_NODE3" And RadioButton2.Checked = True Then
                MAIN.TimerFunction("STOP")
                BatteryID.CalculateCount("FAIL")
                BatteryID.ClearReportBeforeLogGeberate()
                Dim lastItem As String = MAIN.ListBox1.Items(MAIN.ListBox1.Items.Count - 1).ToString()
                BatteryID.GenerateReport(MAIN.TextBox3.Text, MAIN.TextBox2.Text, MAIN.TextBox4.Text, MAIN.Button43.Text, MAIN.TextBox5.Text, MAIN.Label10.Text, MAIN.TextBox18.Text, lastItem)
                BatteryID.GenerateLogReport(CONFIG.TextBox28.Text, MAIN.TextBox4.Text)
                BatteryID.CreateRecord(CONFIG.TextBox1.Text, CONFIG.TextBox2.Text, CONFIG.TextBox3.Text, MAIN.TextBox4.Text, MAIN.TextBox2.Text, MAIN.TextBox3.Text, MAIN.TextBox6.Text, MAIN.TextBox153.Text, MAIN.TextBox18.Text, MAIN.Button43.Text, MAIN.TextBox5.Text, MAIN.Label10.Text)
                BatteryID.ClearSequence()
                TrckRecord()
                Me.Close()
            ElseIf TextBox2.Text = "BATTERYID_NODE3" And RadioButton3.Checked = True Then
                MAIN.TimerFunction("STOP")
                BatteryID.CalculateCount("FAIL")
                BatteryID.ClearReportBeforeLogGeberate()
                Dim lastItem As String = MAIN.ListBox1.Items(MAIN.ListBox1.Items.Count - 1).ToString()
                BatteryID.GenerateReportOnErrorOfBatteryIDNode2(MAIN.TextBox3.Text, MAIN.TextBox2.Text, MAIN.TextBox4.Text, MAIN.Button43.Text, MAIN.TextBox5.Text, MAIN.Label10.Text, MAIN.TextBox18.Text, lastItem)
                BatteryID.GenerateLogReport(CONFIG.TextBox28.Text, MAIN.TextBox4.Text)
                BatteryID.CreateRecord(CONFIG.TextBox1.Text, CONFIG.TextBox2.Text, CONFIG.TextBox3.Text, MAIN.TextBox4.Text, MAIN.TextBox2.Text, MAIN.TextBox3.Text, MAIN.TextBox6.Text, MAIN.TextBox153.Text, MAIN.TextBox18.Text, MAIN.Button43.Text, MAIN.TextBox5.Text, MAIN.Label10.Text)
                BatteryID.ClearSequence()
                TrckRecord()
                Me.Close()
            ElseIf TextBox2.Text = "BATTERYID_NODE3" And RadioButton1.Checked = True Then
                MAIN.TimerFunction("STOP")
                BatteryID.CalculateCount("PASS")
                BatteryID.ClearReportBeforeLogGeberate()
                Dim lastItem As String = MAIN.ListBox1.Items(MAIN.ListBox1.Items.Count - 1).ToString()
                BatteryID.GenerateReport(MAIN.TextBox3.Text, MAIN.TextBox2.Text, MAIN.TextBox4.Text, MAIN.Button43.Text, MAIN.TextBox5.Text, MAIN.Label10.Text, MAIN.TextBox18.Text, lastItem)
                BatteryID.GenerateLogReport(CONFIG.TextBox28.Text, MAIN.TextBox4.Text)
                BatteryID.CreateRecord(CONFIG.TextBox1.Text, CONFIG.TextBox2.Text, CONFIG.TextBox3.Text, MAIN.TextBox4.Text, MAIN.TextBox2.Text, MAIN.TextBox3.Text, MAIN.TextBox6.Text, MAIN.TextBox153.Text, MAIN.TextBox18.Text, MAIN.Button43.Text, MAIN.TextBox5.Text, MAIN.Label10.Text)
                BatteryID.ClearSequence()
                TrckRecord()
                Me.Close()
            End If
        ElseIf TextBox1.Text = "BATTERY ID" And TextBox3.Text = "GOLDEN" Then
            MAIN.TimerFunction("STOP")
            BatteryID.ClearSequence()
            Me.Close()
            TrckRecord()
        ElseIf TextBox1.Text = "SYSTEM ID" And TextBox3.Text = "PRODUCTION" Then
            If TextBox2.Text = "SYSTEMID_INIT" And RadioButton2.Checked = True Then
                MAIN.TimerFunction("STOP")
                SystemID.CalculateCount("FAIL")
                SystemID.ClearReportBeforeLogGeberate()
                Dim lastItem As String = MAIN.ListBox1.Items(MAIN.ListBox1.Items.Count - 1).ToString()
                SystemID.GenerateReportOnErrorOfInit(MAIN.TextBox3.Text, MAIN.TextBox2.Text, MAIN.TextBox4.Text, MAIN.TextBox13.Text, MAIN.TextBox5.Text, MAIN.Label10.Text, MAIN.Button43.Text, lastItem)
                SystemID.GenerateLogReport(CONFIG.TextBox21.Text, MAIN.TextBox4.Text)
                SystemID.CreateRecord(CONFIG.TextBox1.Text, CONFIG.TextBox2.Text, CONFIG.TextBox3.Text, MAIN.TextBox4.Text, MAIN.TextBox2.Text, MAIN.TextBox3.Text, MAIN.TextBox6.Text, MAIN.TextBox153.Text, MAIN.TextBox13.Text, MAIN.Button43.Text, MAIN.TextBox5.Text, MAIN.Label10.Text)
                SystemID.ClearSequence()
                TrckRecord()
                Me.Close()
            ElseIf TextBox2.Text = "SYSTEMID_INIT" And RadioButton3.Checked = True Then
                MAIN.TimerFunction("STOP")
                SystemID.CalculateCount("FAIL")
                SystemID.ClearReportBeforeLogGeberate()
                Dim lastItem As String = MAIN.ListBox1.Items(MAIN.ListBox1.Items.Count - 1).ToString()
                SystemID.GenerateReportOnErrorOfInit(MAIN.TextBox3.Text, MAIN.TextBox2.Text, MAIN.TextBox4.Text, MAIN.TextBox13.Text, MAIN.TextBox5.Text, MAIN.Label10.Text, MAIN.Button43.Text, lastItem)
                SystemID.GenerateLogReport(CONFIG.TextBox21.Text, MAIN.TextBox4.Text)
                SystemID.CreateRecord(CONFIG.TextBox1.Text, CONFIG.TextBox2.Text, CONFIG.TextBox3.Text, MAIN.TextBox4.Text, MAIN.TextBox2.Text, MAIN.TextBox3.Text, MAIN.TextBox6.Text, MAIN.TextBox153.Text, MAIN.TextBox13.Text, MAIN.Button43.Text, MAIN.TextBox5.Text, MAIN.Label10.Text)
                SystemID.ClearSequence()
                TrckRecord()
                Me.Close()
            ElseIf TextBox2.Text = "SYSTEMID_NODE1" And RadioButton2.Checked = True Then
                MAIN.TimerFunction("STOP")
                SystemID.CalculateCount("FAIL")
                SystemID.ClearReportBeforeLogGeberate()
                Dim lastItem As String = MAIN.ListBox1.Items(MAIN.ListBox1.Items.Count - 1).ToString()
                SystemID.GenerateReportOnErrorOfInit(MAIN.TextBox3.Text, MAIN.TextBox2.Text, MAIN.TextBox4.Text, MAIN.TextBox13.Text, MAIN.TextBox5.Text, MAIN.Label10.Text, MAIN.Button43.Text, lastItem)
                SystemID.GenerateLogReport(CONFIG.TextBox21.Text, MAIN.TextBox4.Text)
                SystemID.CreateRecord(CONFIG.TextBox1.Text, CONFIG.TextBox2.Text, CONFIG.TextBox3.Text, MAIN.TextBox4.Text, MAIN.TextBox2.Text, MAIN.TextBox3.Text, MAIN.TextBox6.Text, MAIN.TextBox153.Text, MAIN.TextBox13.Text, MAIN.Button43.Text, MAIN.TextBox5.Text, MAIN.Label10.Text)
                SystemID.ClearSequence()
                TrckRecord()
                Me.Close()
            ElseIf TextBox2.Text = "SYSTEMID_NODE1" And RadioButton3.Checked = True Then
                MAIN.TimerFunction("STOP")
                SystemID.CalculateCount("FAIL")
                SystemID.ClearReportBeforeLogGeberate()
                Dim lastItem As String = MAIN.ListBox1.Items(MAIN.ListBox1.Items.Count - 1).ToString()
                SystemID.GenerateReportOnErrorOfInit(MAIN.TextBox3.Text, MAIN.TextBox2.Text, MAIN.TextBox4.Text, MAIN.TextBox13.Text, MAIN.TextBox5.Text, MAIN.Label10.Text, MAIN.Button43.Text, lastItem)
                SystemID.GenerateLogReport(CONFIG.TextBox21.Text, MAIN.TextBox4.Text)
                SystemID.CreateRecord(CONFIG.TextBox1.Text, CONFIG.TextBox2.Text, CONFIG.TextBox3.Text, MAIN.TextBox4.Text, MAIN.TextBox2.Text, MAIN.TextBox3.Text, MAIN.TextBox6.Text, MAIN.TextBox153.Text, MAIN.TextBox13.Text, MAIN.Button43.Text, MAIN.TextBox5.Text, MAIN.Label10.Text)
                SystemID.ClearSequence()
                TrckRecord()
                Me.Close()
            ElseIf TextBox2.Text = "SYSTEMID_NODE2" And RadioButton2.Checked = True Then
                MAIN.TimerFunction("STOP")
                SystemID.CalculateCount("FAIL")
                SystemID.ClearReportBeforeLogGeberate()
                Dim lastItem As String = MAIN.ListBox1.Items(MAIN.ListBox1.Items.Count - 1).ToString()
                SystemID.GenerateReportOnErrorOfInit(MAIN.TextBox3.Text, MAIN.TextBox2.Text, MAIN.TextBox4.Text, MAIN.TextBox13.Text, MAIN.TextBox5.Text, MAIN.Label10.Text, MAIN.Button43.Text, lastItem)
                SystemID.GenerateLogReport(CONFIG.TextBox21.Text, MAIN.TextBox4.Text)
                SystemID.CreateRecord(CONFIG.TextBox1.Text, CONFIG.TextBox2.Text, CONFIG.TextBox3.Text, MAIN.TextBox4.Text, MAIN.TextBox2.Text, MAIN.TextBox3.Text, MAIN.TextBox6.Text, MAIN.TextBox153.Text, MAIN.TextBox13.Text, MAIN.Button43.Text, MAIN.TextBox5.Text, MAIN.Label10.Text)
                SystemID.ClearSequence()
                TrckRecord()
                Me.Close()
            ElseIf TextBox2.Text = "SYSTEMID_NODE2" And RadioButton3.Checked = True Then
                MAIN.TimerFunction("STOP")
                SystemID.CalculateCount("FAIL")
                SystemID.ClearReportBeforeLogGeberate()
                Dim lastItem As String = MAIN.ListBox1.Items(MAIN.ListBox1.Items.Count - 1).ToString()
                SystemID.GenerateReportOnErrorOfInit(MAIN.TextBox3.Text, MAIN.TextBox2.Text, MAIN.TextBox4.Text, MAIN.TextBox13.Text, MAIN.TextBox5.Text, MAIN.Label10.Text, MAIN.Button43.Text, lastItem)
                SystemID.GenerateLogReport(CONFIG.TextBox21.Text, MAIN.TextBox4.Text)
                SystemID.CreateRecord(CONFIG.TextBox1.Text, CONFIG.TextBox2.Text, CONFIG.TextBox3.Text, MAIN.TextBox4.Text, MAIN.TextBox2.Text, MAIN.TextBox3.Text, MAIN.TextBox6.Text, MAIN.TextBox153.Text, MAIN.TextBox13.Text, MAIN.Button43.Text, MAIN.TextBox5.Text, MAIN.Label10.Text)
                SystemID.ClearSequence()
                TrckRecord()
                Me.Close()
            ElseIf TextBox2.Text = "SYSTEMID_NODE3" And RadioButton1.Checked = True Then
                MAIN.TimerFunction("STOP")
                SystemID.CalculateCount("PASS")
                SystemID.ClearReportBeforeLogGeberate()
                Dim lastItem As String = MAIN.ListBox1.Items(MAIN.ListBox1.Items.Count - 1).ToString()
                SystemID.GenerateReport(MAIN.TextBox3.Text, MAIN.TextBox2.Text, MAIN.TextBox4.Text, MAIN.TextBox13.Text, MAIN.TextBox5.Text, MAIN.Label10.Text, MAIN.Button43.Text, lastItem)
                SystemID.GenerateLogReport(CONFIG.TextBox21.Text, MAIN.TextBox4.Text)
                SystemID.CreateRecord(CONFIG.TextBox1.Text, CONFIG.TextBox2.Text, CONFIG.TextBox3.Text, MAIN.TextBox4.Text, MAIN.TextBox2.Text, MAIN.TextBox3.Text, MAIN.TextBox6.Text, MAIN.TextBox153.Text, MAIN.TextBox13.Text, MAIN.Button43.Text, MAIN.TextBox5.Text, MAIN.Label10.Text)
                SystemID.ClearSequence()
                TrckRecord()
                Me.Close()
            ElseIf TextBox2.Text = "SYSTEMID_NODE3" And RadioButton2.Checked = True Then
                MAIN.TimerFunction("STOP")
                SystemID.CalculateCount("FAIL")
                SystemID.ClearReportBeforeLogGeberate()
                Dim lastItem As String = MAIN.ListBox1.Items(MAIN.ListBox1.Items.Count - 1).ToString()
                SystemID.GenerateReport(MAIN.TextBox3.Text, MAIN.TextBox2.Text, MAIN.TextBox4.Text, MAIN.TextBox13.Text, MAIN.TextBox5.Text, MAIN.Label10.Text, MAIN.Button43.Text, lastItem)
                SystemID.GenerateLogReport(CONFIG.TextBox21.Text, MAIN.TextBox4.Text)
                SystemID.CreateRecord(CONFIG.TextBox1.Text, CONFIG.TextBox2.Text, CONFIG.TextBox3.Text, MAIN.TextBox4.Text, MAIN.TextBox2.Text, MAIN.TextBox3.Text, MAIN.TextBox6.Text, MAIN.TextBox153.Text, MAIN.TextBox13.Text, MAIN.Button43.Text, MAIN.TextBox5.Text, MAIN.Label10.Text)
                SystemID.ClearSequence()
                TrckRecord()
                Me.Close()
            ElseIf TextBox2.Text = "SYSTEMID_NODE3" And RadioButton3.Checked = True Then
                MAIN.TimerFunction("STOP")
                SystemID.CalculateCount("FAIL")
                SystemID.ClearReportBeforeLogGeberate()
                Dim lastItem As String = MAIN.ListBox1.Items(MAIN.ListBox1.Items.Count - 1).ToString()
                SystemID.GenerateReport(MAIN.TextBox3.Text, MAIN.TextBox2.Text, MAIN.TextBox4.Text, MAIN.TextBox13.Text, MAIN.TextBox5.Text, MAIN.Label10.Text, MAIN.Button43.Text, lastItem)
                SystemID.GenerateLogReport(CONFIG.TextBox21.Text, MAIN.TextBox4.Text)
                SystemID.CreateRecord(CONFIG.TextBox1.Text, CONFIG.TextBox2.Text, CONFIG.TextBox3.Text, MAIN.TextBox4.Text, MAIN.TextBox2.Text, MAIN.TextBox3.Text, MAIN.TextBox6.Text, MAIN.TextBox153.Text, MAIN.TextBox13.Text, MAIN.Button43.Text, MAIN.TextBox5.Text, MAIN.Label10.Text)
                SystemID.ClearSequence()
                TrckRecord()
                Me.Close()
            End If
        ElseIf TextBox1.Text = "SYSTEM ID" And TextBox3.Text = "GOLDEN" Then
            MAIN.TimerFunction("STOP")
            SystemID.ClearSequence()
            Me.Close()
            TrckRecord()
        End If
    End Sub

    Private Sub LinkLabel1_LinkClicked(sender As Object, e As LinkLabelLinkClickedEventArgs) Handles LinkLabel1.LinkClicked
        If LinkLabel1.Text = "VIEW FAILURE" Then
            FAILUREVIEW.ShowDialog()
        Else
            'donothng
        End If
    End Sub
    Function TrckRecord()
        If MAIN.CheckBox1.Checked = True Then
            STARTUP.ComboBox1.Items.Clear()
            STARTUP.ComboBox1.Items.AddRange(Products)
            STARTUP.ComboBox2.Items.Clear()
            STARTUP.ComboBox2.Items.AddRange(Test_Methodes)
            STARTUP.ComboBox3.Items.Clear()
            STARTUP.ComboBox3.Items.AddRange(Test_mode)
            If MAIN.TextBox1.Text = "SYSTEM ID" And MAIN.TextBox2.Text = "PRE-COATING" Then
                STARTUP.ComboBox1.SelectedIndex = 0
                STARTUP.ComboBox2.SelectedIndex = 0
                If MAIN.TextBox16.Text = "OPERATOR" Then
                    STARTUP.ComboBox3.Enabled = False
                    STARTUP.ComboBox3.SelectedIndex = 0
                Else
                    STARTUP.ComboBox3.Enabled = True
                End If
            ElseIf MAIN.TextBox1.Text = "SYSTEM ID" And MAIN.TextBox2.Text = "POST-COATING" Then
                STARTUP.ComboBox1.SelectedIndex = 0
                STARTUP.ComboBox2.SelectedIndex = 1
                If MAIN.TextBox16.Text = "OPERATOR" Then
                    STARTUP.ComboBox3.Enabled = False
                    STARTUP.ComboBox3.SelectedIndex = 0
                Else
                    STARTUP.ComboBox3.Enabled = True
                End If
            ElseIf MAIN.TextBox1.Text = "BATTERY ID" And MAIN.TextBox2.Text = "PRE-COATING" Then
                STARTUP.ComboBox1.SelectedIndex = 1
                STARTUP.ComboBox2.SelectedIndex = 0
                If MAIN.TextBox16.Text = "OPERATOR" Then
                    STARTUP.ComboBox3.Enabled = False
                    STARTUP.ComboBox3.SelectedIndex = 0
                Else
                    STARTUP.ComboBox3.Enabled = True
                End If
            ElseIf MAIN.TextBox1.Text = "BATTERY ID" And MAIN.TextBox2.Text = "POST-COATING" Then
                STARTUP.ComboBox1.SelectedIndex = 1
                STARTUP.ComboBox2.SelectedIndex = 1
                If MAIN.TextBox16.Text = "OPERATOR" Then
                    STARTUP.ComboBox3.Enabled = False
                    STARTUP.ComboBox3.SelectedIndex = 0
                Else
                    STARTUP.ComboBox3.Enabled = True
                End If
            ElseIf MAIN.TextBox1.Text = "SMC" And MAIN.TextBox2.Text = "PRE-COATING" Then
                STARTUP.ComboBox1.SelectedIndex = 2
                STARTUP.ComboBox2.SelectedIndex = 0
                If MAIN.TextBox16.Text = "OPERATOR" Then
                    STARTUP.ComboBox3.Enabled = False
                    STARTUP.ComboBox3.SelectedIndex = 0
                Else
                    STARTUP.ComboBox3.Enabled = True
                End If
            ElseIf MAIN.TextBox1.Text = "SMC" And MAIN.TextBox2.Text = "POST-COATING" Then
                STARTUP.ComboBox1.SelectedIndex = 2
                STARTUP.ComboBox2.SelectedIndex = 1
                If MAIN.TextBox16.Text = "OPERATOR" Then
                    STARTUP.ComboBox3.Enabled = False
                    STARTUP.ComboBox3.SelectedIndex = 0
                Else
                    STARTUP.ComboBox3.Enabled = True
                End If
            End If
            Call STARTUP.Show()
            MAIN.TextBox1.Clear()
            MAIN.TextBox2.Clear()
        Else
            'donothing
        End If
    End Function
End Class