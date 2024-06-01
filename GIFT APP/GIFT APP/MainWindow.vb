Imports System.ComponentModel
Imports BunifuAnimatorNS
Imports System.IO.Ports
Imports System.Threading
Imports System.Text
Imports MadMilkman.Ini
Imports System.Configuration
Imports System.Windows.Forms.VisualStyles
Imports Org.BouncyCastle.Bcpg

Public Class MainWindow
    Private WithEvents serialPort As New SerialPort()
    Private readThread As Thread
    Dim AfterLogin = {"Log-Out"}
    Dim transition As New BunifuTransition
    Dim requiredColumns As New List(Of String)()
    Dim ini As New IniFile
    Dim inipath As String = ConfigurationManager.AppSettings("IniFile")
    Dim currentHostName As String = System.Net.Dns.GetHostName()
    Private Sub MainWindow_Load(sender As Object, e As EventArgs) Handles MyBase.Load
        BunifuLabel1.Visible = False
        BunifuPanel2.Visible = False
        BunifuDropdown1.Visible = False
        CheckBox1.Checked = False 'Reset Login
        ListBox1.Items.Clear() 'Clear All User Records
        Panel1.Visible = False
        TabControl1.Visible = False
        RadioButton1.Checked = False 'Reset Method For Gift Distribution
        RadioButton2.Checked = False 'Reset Method For Gift Distribution
        RadioButton3.Checked = False ' Reset Function Type
        RadioButton4.Checked = False 'Reset Function Type
        TextBox1.Text = Nothing 'Reset Variable Function Date
        TextBox2.Clear() 'Clear Comport Data
        RadioButton6.Checked = False 'Reset Emp Identification Methode
        RadioButton5.Checked = False 'Reset Emp Identification Methode
        BackgroundWorker1.RunWorkerAsync()
        BackgroundWorker1.WorkerSupportsCancellation = True
    End Sub
    Private Sub BackgroundWorker1_DoWork(sender As Object, e As System.ComponentModel.DoWorkEventArgs) Handles BackgroundWorker1.DoWork
        Dim line1, line2 As Label
line1:
        Dim LoadIniFile As String = FuncLib.LoadIniFile()
        If LoadIniFile IsNot Nothing Then
            Custommessage.Show($"Error Loading Confriguration File. {Environment.NewLine} Error : {LoadIniFile}", Color.IndianRed)
            e.Cancel = True
        Else
            GoTo line2
        End If
line2:
        Dim InitializeDB As String = FuncLib.DataBaseOperations.InitializeDB()
        If InitializeDB IsNot Nothing Then
            Custommessage.Show($"Error Connecting To DataBase. {Environment.NewLine} Error :{InitializeDB}", Color.IndianRed)
            FuncLib.WriteLog.WriteErrorLog($"Error Initializing DataBase : {InitializeDB}")
            e.Cancel = True
        Else
            'Will AutoTrigger backgroundworker1_RunworkerCompleted
        End If
    End Sub
    Private Sub BackgroundWorker1_RunWorkerCompleted(sender As Object, e As RunWorkerCompletedEventArgs) Handles BackgroundWorker1.RunWorkerCompleted
        If e.Error IsNot Nothing Then
            Custommessage.Show(e.Error.ToString, Color.IndianRed)
            FuncLib.WriteLog.WriteErrorLog(e.Error.ToString)
            Exit Sub
        ElseIf e.Cancelled = True Then
            Exit Sub
        Else
            transition.HideSync(BunifuPictureBox1, False, BunifuAnimatorNS.Animation.ScaleAndHorizSlide)
            transition.ShowSync(LoginForm, False, BunifuAnimatorNS.Animation.ScaleAndHorizSlide)
        End If
    End Sub
    Private Sub BunifuDropdown1_SelectedIndexChanged(sender As Object, e As EventArgs) Handles BunifuDropdown1.SelectedIndexChanged
        If BunifuDropdown1.SelectedItem = "Log-Out" Then
            BunifuSnackbar1.Show(Me, $"Log-Out Success From {ListBox1.Items(4)}", Bunifu.UI.WinForms.BunifuSnackbar.MessageTypes.Success, 2500)
            FuncLib.WriteLog.WriteAppLog($"Log-Out Succes For {ListBox1.Items(4)}")
            CheckBox1.Checked = False
            BunifuLabel1.Visible = False
            LoginForm.BunifuTextBox1.Clear()
            LoginForm.BunifuTextBox2.Clear()
            TabControl1.TabPages.Clear()
            BunifuPanel2.Visible = False
            ListBox1.Items.Clear()
            transition.ShowSync(LoginForm, False, BunifuAnimatorNS.Animation.ScaleAndHorizSlide)
        End If
    End Sub
    Private Sub BunifuDropdown1_Click(sender As Object, e As EventArgs) Handles BunifuDropdown1.Click
        If CheckBox1.Checked = True Then
            BunifuDropdown1.Items.Clear()
            BunifuDropdown1.Items.AddRange(AfterLogin)
        End If
    End Sub
    Private Sub CheckBox1_CheckedChanged(sender As Object, e As EventArgs) Handles CheckBox1.CheckedChanged
        If CheckBox1.Checked = False Then
            BunifuDropdown1.Visible = False
            ListBox1.Items.Clear()
            BunifuPanel3.Visible = False
        End If
    End Sub
    Private Sub BunifuButton8_Click(sender As Object, e As EventArgs) Handles BunifuButton8.Click
        transition.HideSync(Me, False, BunifuAnimatorNS.Animation.ScaleAndHorizSlide)
        Application.Exit()
    End Sub
    Private Sub BunifuDropdown2_SelectedIndexChanged(sender As Object, e As EventArgs) Handles BunifuDropdown2.SelectedIndexChanged
        If BunifuDropdown2.SelectedItem = "Home" Then
            BunifuDropdown2.Text = "Home"
            Dim BuildCardForm As String = FuncLib.BirthdayCardForm.InitializeBirthdayCards()
            If BuildCardForm = "True" Then
                If TabControl1.Visible = True Then
                    transition.HideSync(TabControl1, False, BunifuAnimatorNS.Animation.HorizBlind)
                    TabControl1.TabPages.Clear()
                    TabControl1.TabPages.Add(TabPage1)
                    transition.ShowSync(TabControl1, False, BunifuAnimatorNS.Animation.HorizBlind)
                Else
                    TabControl1.TabPages.Clear()
                    TabControl1.TabPages.Add(TabPage1)
                    transition.ShowSync(TabControl1, False, BunifuAnimatorNS.Animation.HorizBlind)
                End If
            Else
                TabControl1.TabPages.Clear()
                transition.HideSync(TabControl1, False, BunifuAnimatorNS.Animation.HorizBlind)
                BunifuSnackbar1.Show(Me, $"Error Occured While Building BirthdayCards.{Environment.NewLine}Error: {BuildCardForm}", Bunifu.UI.WinForms.BunifuSnackbar.MessageTypes.Error, 3000)
                FuncLib.WriteLog.WriteErrorLog($"Error While Building CardForm : {BuildCardForm}")
                Exit Sub
            End If
        End If
    End Sub
    Private Sub BunifuDropdown6_SelectedIndexChanged(sender As Object, e As EventArgs) Handles BunifuDropdown6.SelectedIndexChanged
        BunifuDropdown6.Text = "GIFTS"
        GiftClearance()
        If BunifuDropdown6.SelectedItem = "Issue Gift" Then
            If (ListBox1.Items(14) = "True" Or ListBox1.Items(16) = "True") Then
                If ListBox1.Items(4) = "ADMINISTRATOR" Then
                    GiftClearance()
                    TabControl1.TabPages.Clear()
                    transition.HideSync(TabControl1, False, BunifuAnimatorNS.Animation.HorizBlind)
                    BunifuSnackbar1.Show(Me, "Administrator Can Not Share Gifts", Bunifu.UI.WinForms.BunifuSnackbar.MessageTypes.Warning, 2500)
                    Exit Sub
                Else
                    Dim Occasion As List(Of String) = FuncLib.Gifts.GetOcaasioList()
                    If Occasion.Count <= 0 AndAlso TypeOf Occasion(0) Is String Then
                        Dim errorMessage As String = Occasion(0)
                        TabControl1.TabPages.Clear()
                        transition.HideSync(TabControl1, False, BunifuAnimatorNS.Animation.HorizBlind)
                        BunifuSnackbar1.Show(Me, $"Error While Fetching Occasion List From. {Environment.NewLine} Error : {errorMessage}", Bunifu.UI.WinForms.BunifuSnackbar.MessageTypes.Error, 2500)
                        FuncLib.WriteLog.WriteErrorLog($"Error occured While Trying To Fetch Occasion List : {errorMessage}")
                        Exit Sub
                    Else
                        BunifuDropdown7.Items.Clear()
                        BunifuDropdown7.Text = Nothing
                        GroupBox1.Visible = True
                        GroupBox1.Enabled = True
                        BunifuDropdown7.Items.AddRange(Occasion.ToArray())
                        If TabControl1.Visible = True Then
                            transition.HideSync(TabControl1, False, BunifuAnimatorNS.Animation.HorizBlind)
                            TabControl1.TabPages.Clear()
                            TabControl1.TabPages.Add(TabPage4)
                            transition.ShowSync(TabControl1, False, BunifuAnimatorNS.Animation.HorizBlind)
                        Else
                            TabControl1.TabPages.Clear()
                            TabControl1.TabPages.Add(TabPage4)
                            transition.ShowSync(TabControl1, False, BunifuAnimatorNS.Animation.HorizBlind)
                        End If
                    End If
                End If
            Else
                GiftClearance()
                TabControl1.TabPages.Clear()
                transition.HideSync(TabControl1, False, BunifuAnimatorNS.Animation.HorizBlind)
                BunifuSnackbar1.Show(Me, "You Are Not Authorised To Distribute Gifts", Bunifu.UI.WinForms.BunifuSnackbar.MessageTypes.Warning, 2500)
                Exit Sub
            End If
        ElseIf BunifuDropdown6.SelectedItem = "Your History" Then
            BunifuDataGridView3.DataSource = Nothing
            If ListBox1.Items(14) = "True" Or ListBox1.Items(16) = "True" Then
                If ListBox1.Items(4) = "ADMINISTRATOR" Then
                    GiftClearance()
                    TabControl1.TabPages.Clear()
                    transition.HideSync(TabControl1, False, BunifuAnimatorNS.Animation.HorizBlind)
                    BunifuSnackbar1.Show(Me, "Not Authorised", Bunifu.UI.WinForms.BunifuSnackbar.MessageTypes.Warning, 2500)
                    Exit Sub
                Else
                    Dim result As Tuple(Of DataTable, String) = FuncLib.Gifts.FetchSharedGiftsHistory(ListBox1.Items(0).ToString, ListBox1.Items(1).ToString)
                    If result IsNot Nothing AndAlso result.Item1 IsNot Nothing Then
                        Dim dataTable = result.Item1
                        If dataTable.Rows.Count > 0 Then
                            dataTable.Columns("Employee_ID").ColumnName = "EMP ID"
                            dataTable.Columns("NT_ID").ColumnName = "NT ID"
                            dataTable.Columns("Employee_Name").ColumnName = "NAME"
                            dataTable.Columns("Celebration_Date").ColumnName = "DATE"
                            dataTable.Columns("Buisness_Area").ColumnName = "BUISNESS"
                            dataTable.Columns("Functions").ColumnName = "DEPARTMENT"
                            dataTable.Columns("Leave_Approver").ColumnName = "MANAGER"
                            dataTable.Columns("Gift_Type").ColumnName = "GIFT SHARED"
                            dataTable.Columns("Shared_At").ColumnName = "TIME"
                            BunifuDataGridView3.DataSource = dataTable
                            If TabControl1.Visible = True Then
                                transition.HideSync(TabControl1, False, BunifuAnimatorNS.Animation.HorizBlind)
                                TabControl1.TabPages.Clear()
                                TabControl1.TabPages.Add(TabPage5)
                                transition.ShowSync(TabControl1, False, BunifuAnimatorNS.Animation.HorizBlind)
                            Else
                                TabControl1.TabPages.Clear()
                                TabControl1.TabPages.Add(TabPage5)
                                transition.ShowSync(TabControl1, False, BunifuAnimatorNS.Animation.HorizBlind)
                            End If
                        Else
                            BunifuSnackbar1.Show(Me, "You Dont Have Any Gift Sharing List", Bunifu.UI.WinForms.BunifuSnackbar.MessageTypes.Success, 2500)
                            GiftClearance()
                            Exit Sub
                        End If
                    Else
                        BunifuSnackbar1.Show(Me, $"Error Fetching Data From Gift history Database. {Environment.NewLine} Error : {result.Item2.ToString}", Bunifu.UI.WinForms.BunifuSnackbar.MessageTypes.Warning, 3000)
                        FuncLib.WriteLog.WriteErrorLog($"Error While Reading Data From Gift_History : {result.Item2.ToString}")
                        GiftClearance()
                        Exit Sub
                    End If
                End If
            Else
                GiftClearance()
                TabControl1.TabPages.Clear()
                transition.HideSync(TabControl1, False, BunifuAnimatorNS.Animation.HorizBlind)
                BunifuSnackbar1.Show(Me, "You Are Not Authorised To View history", Bunifu.UI.WinForms.BunifuSnackbar.MessageTypes.Warning, 2500)
                Exit Sub
            End If
        End If

    End Sub
    Private Sub GiftClearance()
        GroupBox1.Visible = False
        GroupBox2.Visible = False
        GroupBox4.Visible = False
        GroupBox5.Visible = False
        BunifuButton2.Visible = False
        BunifuTextBox1.Clear()
        BunifuTextBox2.Clear()
        BunifuTextBox3.Clear()
        BunifuTextBox4.Clear()
        BunifuTextBox5.Clear()
        BunifuTextBox6.Clear()
        BunifuTextBox7.Clear()
        BunifuTextBox8.Clear()
        BunifuDataGridView3.DataSource = Nothing
        TabControl1.TabPages.Clear()
        transition.ShowSync(TabControl1, False, BunifuAnimatorNS.Animation.HorizBlind)
    End Sub
    Private Sub BunifuDropdown7_SelectedIndexChanged(sender As Object, e As EventArgs) Handles BunifuDropdown7.SelectedIndexChanged
        ClearSequenceAfterOccasionChanged()
        Dim Occasiontype As String = FuncLib.Gifts.CheckTypeOfOcassion(BunifuDropdown7.SelectedItem.ToString)
        If Occasiontype = "Constant" Then
            RadioButton4.Checked = True 'constant Function
            RadioButton2.Checked = True 'Bulk Distribution Selector
            RadioButton6.Checked = True '10D SELCTOR
            RadioButton3.Checked = False 'variable function
            ini.Load(inipath)
            If RadioButton2.Checked = True Then
                Try
                    If serialPort.PortName <> ini.Sections("SERIAL").Keys($"{currentHostName}").Value AndAlso serialPort.IsOpen = False Then
                        serialPort.PortName = ini.Sections("SERIAL").Keys($"{currentHostName}").Value
                        serialPort.BaudRate = 9600
                        serialPort.Parity = Parity.None
                        serialPort.StopBits = StopBits.One
                        If serialPort.IsOpen = False Then
                            serialPort.Open()
                            serialPort.Write("START")
                            GroupBox4.Visible = True
                            BunifuTextBox1.Enabled = True
                            BunifuTextBox1.ReadOnly = True
                            BunifuTextBox1.PlaceholderText = "SCAN YOUR ID CARD"
                        Else
                            serialPort.Close()
                            Thread.Sleep(1000)
                            serialPort.Open()
                            serialPort.Write("START")
                            GroupBox4.Visible = True
                            BunifuTextBox1.Enabled = True
                            BunifuTextBox1.ReadOnly = True
                            BunifuTextBox1.PlaceholderText = "SCAN YOUR ID CARD"
                        End If
                    Else
                        serialPort.BaudRate = 9600
                        serialPort.Parity = Parity.None
                        serialPort.StopBits = StopBits.One
                        If serialPort.IsOpen = False Then
                            serialPort.Open()
                            serialPort.Write("START")
                            GroupBox4.Visible = True
                            BunifuTextBox1.Enabled = True
                            BunifuTextBox1.ReadOnly = True
                            BunifuTextBox1.PlaceholderText = "SCAN YOUR ID CARD"
                        Else
                            serialPort.Close()
                            System.Threading.Thread.Sleep(1000)
                            serialPort.Open()
                            serialPort.Write("START")
                            GroupBox4.Visible = True
                            BunifuTextBox1.Enabled = True
                            BunifuTextBox1.ReadOnly = True
                            BunifuTextBox1.PlaceholderText = "SCAN YOUR ID CARD"
                        End If
                    End If
                Catch ex As Exception
                    BunifuSnackbar1.Show(Me, $"Error Occured While Trying To Open Comport!{Environment.NewLine}Error:{ex.Message}", Bunifu.UI.WinForms.BunifuSnackbar.MessageTypes.Error, 2500)
                    FuncLib.WriteLog.WriteErrorLog($"Error Occured While Tring To Open COmport While Sahring Gift :{ex.Message}")
                    ClearSequenceAfterOccasionChanged()
                    Exit Sub
                End Try
            End If
        ElseIf Occasiontype = "Variable" Then
            RadioButton4.Checked = False
            RadioButton3.Checked = True
            Dim OccasionDate As List(Of Date) = FuncLib.Gifts.VarableDateQuery(BunifuDropdown7.SelectedItem)
            If OccasionDate.Count <= 0 AndAlso OccasionDate(0) = DateTime.MinValue Then
                BunifuSnackbar1.Show(Me, $"Error Occured While Fetching Occasion Dates!", Bunifu.UI.WinForms.BunifuSnackbar.MessageTypes.Error, 2500)
                FuncLib.WriteLog.WriteErrorLog($"Error occured While Executing FuncLib.Gifts.VarableDateQuery")
                ClearSequenceAfterOccasionChanged()
                Exit Sub
            Else
                BunifuDropdown8.Items.Clear()
                GroupBox2.Enabled = True
                GroupBox2.Visible = True
                BunifuDropdown8.Items.Clear()
                BunifuDropdown8.Text = Nothing
                For Each DateDetails In OccasionDate
                    BunifuDropdown8.Items.Add(DateDetails.ToString("MM/dd/yyyy"))
                Next
            End If
        ElseIf Occasiontype = "Undefined Type" Then
            BunifuSnackbar1.Show(Me, $"No OccasionType Mentioned For {BunifuDropdown7.SelectedItem}.{Environment.NewLine} Please Contact To Administrator or Developer!", Bunifu.UI.WinForms.BunifuSnackbar.MessageTypes.Error, 3000)
            FuncLib.WriteLog.WriteErrorLog($"{BunifuDropdown7.SelectedItem.ToString} Has Been Found Without Any OcassionType")
            ClearSequenceAfterOccasionChanged()
            Exit Sub
        Else
            BunifuSnackbar1.Show(Me, $"Error Occured While Reading The Type Of The Occasion {Environment.NewLine} Error : {Occasiontype}", Bunifu.UI.WinForms.BunifuSnackbar.MessageTypes.Error, 3000)
            FuncLib.WriteLog.WriteErrorLog($"Error While Trying To Read Type of Occasion : {Occasiontype}")
            ClearSequenceAfterOccasionChanged()
            Exit Sub
        End If
    End Sub
    Private Sub ClearSequenceAfterOccasionChanged()
        GroupBox2.Visible = False
        BunifuDropdown8.Items.Clear()
        BunifuDropdown8.Text = Nothing
        RadioButton4.Checked = False
        GroupBox4.Visible = False
        BunifuTextBox1.Clear()
        GroupBox5.Visible = False
        BunifuTextBox2.Clear()
        BunifuTextBox5.Clear()
        BunifuTextBox4.Clear()
        BunifuTextBox6.Clear()
        BunifuTextBox3.Clear()
        BunifuTextBox7.Clear()
        BunifuTextBox8.Clear()
        BunifuButton2.Visible = False
        BunifuButton15.Visible = False
        RadioButton2.Checked = False 'Bulk Distribution Selector
        RadioButton1.Checked = False 'Individual Distribution Selector
        RadioButton6.Checked = False '10D SELCTOR
        RadioButton5.Checked = False '5D SELECTOR
        RadioButton3.Checked = False 'variable function
    End Sub
    Private Sub BunifuDropdown8_SelectedIndexChanged(sender As Object, e As EventArgs) Handles BunifuDropdown8.SelectedIndexChanged
        ClearSequenceAfterSelectingVariableDate()
        RadioButton2.Checked = True 'Bulk Distribution Selector
        RadioButton6.Checked = True '10D SELCTOR
        If RadioButton2.Checked = True AndAlso RadioButton6.Checked = True AndAlso RadioButton3.Checked = True Then
            ini.Load(inipath)
            Try
                If serialPort.PortName <> ini.Sections("SERIAL").Keys($"{currentHostName}").Value AndAlso serialPort.IsOpen = False Then
                    serialPort.PortName = ini.Sections("SERIAL").Keys($"{currentHostName}").Value
                    serialPort.BaudRate = 9600
                    serialPort.Parity = Parity.None
                    serialPort.StopBits = StopBits.One
                    If serialPort.IsOpen = False Then
                        serialPort.Open()
                        serialPort.Write("START")
                        GroupBox4.Visible = True
                        BunifuTextBox1.Enabled = True
                        BunifuTextBox1.ReadOnly = True
                        BunifuTextBox1.PlaceholderText = "SCAN YOUR ID CARD"
                    Else
                        serialPort.Close()
                        Thread.Sleep(1000)
                        serialPort.Open()
                        serialPort.Write("START")
                        GroupBox4.Visible = True
                        BunifuTextBox1.Enabled = True
                        BunifuTextBox1.ReadOnly = True
                        BunifuTextBox1.PlaceholderText = "SCAN YOUR ID CARD"
                    End If
                Else
                    serialPort.BaudRate = 9600
                    serialPort.Parity = Parity.None
                    serialPort.StopBits = StopBits.One
                    If serialPort.IsOpen = False Then
                        serialPort.Open()
                        serialPort.Write("START")
                        GroupBox4.Visible = True
                        BunifuTextBox1.Enabled = True
                        BunifuTextBox1.ReadOnly = True
                        BunifuTextBox1.PlaceholderText = "SCAN YOUR ID CARD"
                    Else
                        serialPort.Close()
                        System.Threading.Thread.Sleep(1000)
                        serialPort.Open()
                        serialPort.Write("START")
                        GroupBox4.Visible = True
                        BunifuTextBox1.Enabled = True
                        BunifuTextBox1.ReadOnly = True
                        BunifuTextBox1.PlaceholderText = "SCAN YOUR ID CARD"
                    End If
                End If
            Catch ex As Exception
                BunifuSnackbar1.Show(Me, $"Error Occured While Trying To Open Comport!{Environment.NewLine}Error:{ex.Message}", Bunifu.UI.WinForms.BunifuSnackbar.MessageTypes.Error, 2500)
                FuncLib.WriteLog.WriteErrorLog($"Error Occured While Tring To Open COmport While Sahring Gift :{ex.Message}")
                ClearSequenceAfterOccasionChanged()
                Exit Sub
            End Try
        End If
    End Sub
    Private Sub ClearSequenceAfterSelectingVariableDate()
        GroupBox4.Visible = False
        BunifuTextBox1.Clear()
        GroupBox5.Visible = False
        BunifuTextBox2.Clear()
        BunifuTextBox5.Clear()
        BunifuTextBox4.Clear()
        BunifuTextBox6.Clear()
        BunifuTextBox3.Clear()
        BunifuTextBox7.Clear()
        BunifuTextBox8.Clear()
        BunifuButton2.Visible = False
        BunifuButton15.Visible = False
        RadioButton2.Checked = False 'Bulk Distribution Selector
        RadioButton1.Checked = False 'Individual Distribution Selector
        RadioButton6.Checked = False '10D SELCTOR
        RadioButton5.Checked = False '5D SELECTOR
    End Sub
    Private Sub BunifuButton2_Click(sender As Object, e As EventArgs) Handles BunifuButton2.Click
        If RadioButton2.Checked = True Then 'Bulk Distribution
            If RadioButton4.Checked = True Then 'Constant Function
                Dim DistributeoninbulkConstantFunction As String = FuncLib.Gifts.RecordEmployeeEligibilityForBulkDistributionConstant(BunifuDropdown7.Text.ToString, BunifuTextBox2.Text, BunifuTextBox3.Text, BunifuTextBox4.Text, BunifuTextBox5.Text, BunifuTextBox6.Text, BunifuTextBox7.Text, BunifuTextBox8.Text)
                If DistributeoninbulkConstantFunction = "True" Then
                    BunifuSnackbar1.Show(Me, $"{BunifuDropdown7.SelectedItem} Gift Shared With {BunifuTextBox4.Text}", Bunifu.UI.WinForms.BunifuSnackbar.MessageTypes.Success, 2500)
                    ClearsequenceaftersharingGift()
                Else
                    BunifuSnackbar1.Show(Me, $"Error Checking The Fuctionality Data : {DistributeoninbulkConstantFunction}", Bunifu.UI.WinForms.BunifuSnackbar.MessageTypes.Error, 2500)
                    FuncLib.WriteLog.WriteErrorLog($"Error Occured While Sahring Gift : {DistributeoninbulkConstantFunction}")
                    Exit Sub
                End If
            ElseIf RadioButton3.Checked = True Then 'Variable Function
                Dim DistributeoninbulkVariableFunction As String = FuncLib.Gifts.RecordEmployeeEligibilityForBulkDistributionvariable(BunifuDropdown7.Text.ToString, BunifuTextBox2.Text, BunifuTextBox3.Text, BunifuTextBox4.Text, BunifuTextBox5.Text, BunifuTextBox6.Text, BunifuTextBox7.Text, BunifuTextBox8.Text)
                If DistributeoninbulkVariableFunction = "True" Then
                    BunifuSnackbar1.Show(Me, $"{BunifuDropdown7.SelectedItem} Gift Shared With {BunifuTextBox4.Text}", Bunifu.UI.WinForms.BunifuSnackbar.MessageTypes.Success, 2500)
                    ClearsequenceaftersharingGift()
                Else
                    BunifuSnackbar1.Show(Me, $"Error Checking The Fuctionality Data : {DistributeoninbulkVariableFunction}", Bunifu.UI.WinForms.BunifuSnackbar.MessageTypes.Warning, 3000)
                    FuncLib.WriteLog.WriteErrorLog($"Error Occured While Sahring Gift : {DistributeoninbulkVariableFunction}")
                    Exit Sub
                End If
            End If
        End If
    End Sub
    Private Sub ClearsequenceaftersharingGift()
        BunifuTextBox1.Clear()
        BunifuTextBox2.Clear()
        BunifuTextBox3.Clear()
        BunifuTextBox4.Clear()
        BunifuTextBox5.Clear()
        BunifuTextBox6.Clear()
        BunifuTextBox7.Clear()
        BunifuTextBox8.Clear()
        BunifuButton2.Visible = False
        BunifuButton2.Enabled = False
        BunifuButton15.Visible = False
        BunifuButton15.Enabled = False
        GroupBox5.Visible = False
        BunifuTextBox1.Clear()
        BunifuTextBox1.Enabled = True
        BunifuTextBox1.ReadOnly = True
        BunifuTextBox1.PlaceholderText = "SCAN YOUR ID CARD"
    End Sub
    Private Sub BunifuButton15_Click(sender As Object, e As EventArgs) Handles BunifuButton15.Click
        BunifuSnackbar1.Show(Me, $"Gift Sharing With {BunifuTextBox4.Text} Aborted!", Bunifu.UI.WinForms.BunifuSnackbar.MessageTypes.Information, 1000)
        ClearsequenceaftersharingGift()
    End Sub
    Private Sub BunifuDropdown3_SelectedIndexChanged(sender As Object, e As EventArgs) Handles BunifuDropdown3.SelectedIndexChanged
        BunifuDropdown3.Text = "EVENTS"
        If BunifuDropdown3.SelectedItem = "Add Events" Then
            CLearSequenceAfterSelcetionOfEvents()
            If ListBox1.Items(15) = "True" Or ListBox1.Items(16) = "True" Then
                If ListBox1.Items(4) = "ADMINISTRATOR" Then
                    CLearSequenceAfterSelcetionOfEvents()
                    TabControl1.TabPages.Clear()
                    transition.HideSync(TabControl1, False, BunifuAnimatorNS.Animation.HorizBlind)
                    BunifuSnackbar1.Show(Me, "Administrator Can Not Create Events", Bunifu.UI.WinForms.BunifuSnackbar.MessageTypes.Warning, 2500)
                    Exit Sub
                Else
                    Try
                        BunifuDropdown9.Items.Clear()
                        BunifuDropdown9.Text = Nothing
                        Dim VariableOccasion As List(Of String) = FuncLib.Events.GetVariableocassions()
                        BunifuDropdown9.Items.AddRange(VariableOccasion.ToArray())
                        BunifuDropdown9.Items.Add("Add New Event")
                        If TabControl1.Visible = True Then
                            transition.HideSync(TabControl1, False, BunifuAnimatorNS.Animation.HorizBlind)
                            TabControl1.TabPages.Clear()
                            TabControl1.TabPages.Add(TabPage6)
                            transition.ShowSync(TabControl1, False, BunifuAnimatorNS.Animation.HorizBlind)
                        Else
                            TabControl1.TabPages.Clear()
                            TabControl1.TabPages.Add(TabPage6)
                            transition.ShowSync(TabControl1, False, BunifuAnimatorNS.Animation.HorizBlind)
                        End If
                    Catch ex As Exception
                        BunifuSnackbar1.Show(Me, $"Error Occured While Fetching Variable Occasion List From DB!{Environment.NewLine}Error: {ex.Message}", Bunifu.UI.WinForms.BunifuSnackbar.MessageTypes.Error, 2500)
                        FuncLib.WriteLog.WriteErrorLog($"Error Occured While Executing FuncLib.Events.GetVariableocassions:{ex.Message}")
                        CLearSequenceAfterSelcetionOfEvents()
                        Exit Sub
                    End Try
                End If
            Else
                CLearSequenceAfterSelcetionOfEvents()
                TabControl1.TabPages.Clear()
                transition.HideSync(TabControl1, False, BunifuAnimatorNS.Animation.HorizBlind)
                BunifuSnackbar1.Show(Me, "You Are Not Authorised To Add Events", Bunifu.UI.WinForms.BunifuSnackbar.MessageTypes.Warning, 2500)
                Exit Sub
            End If
        End If
    End Sub
    Private Sub CLearSequenceAfterSelcetionOfEvents()
        ProgressBar1.Visible = False
        ProgressBar1.Minimum = 0
        ProgressBar1.Value = 0
        GroupBox7.Enabled = True
        BunifuDropdown9.Enabled = True
        BunifuDropdown9.Items.Clear()
        BunifuDropdown9.Text = Nothing
        Dim VariableOccasion As List(Of String) = FuncLib.Events.GetVariableocassions()
        BunifuDropdown9.Items.AddRange(VariableOccasion.ToArray())
        BunifuDropdown9.Items.Add("Add New Event")
        GroupBox8.Visible = False
        GroupBox8.Enabled = False
        BunifuButton6.Visible = False
        BunifuButton6.Enabled = False
        GroupBox9.Visible = False
        GroupBox9.Enabled = False
        GroupBox11.Visible = False
        GroupBox11.Enabled = False
        BunifuDropdown10.Items.Clear()
        BunifuDropdown10.Text = Nothing
        GroupBox12.Visible = False
        GroupBox12.Enabled = False
        BunifuDropdown11.Items.Clear()
        BunifuDropdown11.Text = Nothing
        BunifuTextBox10.Clear()
        BunifuDataGridView4.DataSource = Nothing
        BunifuDataGridView4.Visible = False
        BunifuDataGridView4.Enabled = False
        BunifuButton5.Visible = False
        BunifuButton5.Enabled = False
        BunifuButton1.Visible = False
        BunifuButton1.Enabled = False
        CheckBox2.Checked = False
        CheckBox2.Visible = False
        CheckBox3.Checked = False
        CheckBox3.Enabled = False
        CheckBox3.Visible = False
    End Sub
    Private Sub BunifuDropdown9_SelectedIndexChanged(sender As Object, e As EventArgs) Handles BunifuDropdown9.SelectedIndexChanged
        ClearsequenceAfterSelectingOccasioninevent()
        If BunifuDropdown9.SelectedItem = "Add New Event" Then
            ClearsequenceAfterSelectingOccasioninevent()
            transition.ShowSync(Add_Events, False, BunifuAnimatorNS.Animation.ScaleAndHorizSlide)
        Else
            ClearsequenceAfterSelectingOccasioninevent()
            GroupBox8.Visible = True
            GroupBox8.Enabled = True
            BunifuButton6.Visible = True
            BunifuButton6.Enabled = True
            BunifuPanel4.Enabled = True
        End If
    End Sub
    Private Sub ClearsequenceAfterSelectingOccasioninevent()
        GroupBox8.Visible = False
        GroupBox8.Enabled = False
        BunifuButton6.Visible = False
        BunifuButton6.Enabled = False
        GroupBox9.Visible = False
        GroupBox9.Enabled = False
        GroupBox11.Visible = False
        GroupBox11.Enabled = False
        BunifuDropdown10.Items.Clear()
        BunifuDropdown10.Text = Nothing
        GroupBox12.Visible = False
        GroupBox12.Enabled = False
        BunifuDropdown11.Items.Clear()
        BunifuDropdown11.Text = Nothing
        BunifuTextBox10.Clear()
        BunifuDataGridView4.DataSource = Nothing
        BunifuDataGridView4.Visible = False
        BunifuDataGridView4.Enabled = False
        BunifuButton5.Visible = False
        BunifuButton5.Enabled = False
        BunifuButton1.Visible = False
        BunifuButton1.Enabled = False
        BunifuPanel4.Enabled = False
    End Sub
    Private Sub BunifuButton6_Click(sender As Object, e As EventArgs) Handles BunifuButton6.Click
        Dim Checkifdatealreadydeclared As String = FuncLib.Events.CompareDateMatcOfFunction(BunifuDropdown9.SelectedItem, BunifuDatePicker1.Value.Date)
        If Checkifdatealreadydeclared = "False" AndAlso BunifuDatePicker1.Value.Date >= DateTime.Now.Date Then
            Dim CreateFunctionDate As String = FuncLib.Events.addFunctionDateToFunction(BunifuDropdown9.SelectedItem, BunifuDatePicker1.Value)
            If CreateFunctionDate = "True" Then
                BunifuSnackbar1.Show(Me, $"New Function Date Created For {BunifuDropdown9.SelectedItem} On {BunifuDatePicker1.Value.Date}", Bunifu.UI.WinForms.BunifuSnackbar.MessageTypes.Success, 2500)
                CheckBox2.Checked = True
                FuncLib.WriteLog.WriteAppLog($"New Date[{BunifuDatePicker1.Value.Date}] Created For :{BunifuDropdown9.SelectedItem}")
                ClearsequenceAfterCreatingVariableOccasion()
            Else
                BunifuSnackbar1.Show(Me, $"Error Occured While Creating New Ocassion Date For : {BunifuDropdown9.SelectedItem} On {BunifuDatePicker1.Value.Date}{Environment.NewLine} Error :{CreateFunctionDate}", Bunifu.UI.WinForms.BunifuSnackbar.MessageTypes.Error, 3000)
                FuncLib.WriteLog.WriteErrorLog($"Error Occured WHile Executing FuncLib.Events.addFunctionDateToFunction :{CreateFunctionDate}")
                Exit Sub
            End If
        ElseIf Checkifdatealreadydeclared = "True" Then
            BunifuSnackbar1.Show(Me, $"{BunifuDropdown9.SelectedItem} Is Already Decalred On {BunifuDatePicker1.Value.Date}{Environment.NewLine}Please Select The Proper Date!", Bunifu.UI.WinForms.BunifuSnackbar.MessageTypes.Information, 2500)
            Exit Sub
        ElseIf BunifuDatePicker1.Value.Date < DateTime.Now.Date Then
            BunifuSnackbar1.Show(Me, $"Can Not Create Function Before Today", Bunifu.UI.WinForms.BunifuSnackbar.MessageTypes.Information, 2500)
            Exit Sub
        Else
            BunifuSnackbar1.Show(Me, $"Error Occured While Checking The Avialable Date Fro The Same Function!{Environment.NewLine}Error:{Checkifdatealreadydeclared}", Bunifu.UI.WinForms.BunifuSnackbar.MessageTypes.Error, 2500)
            FuncLib.WriteLog.WriteErrorLog($"Error Occured While Executing FuncLib.Events.CompareDateMatcOfFunction :{Checkifdatealreadydeclared}")
            Exit Sub
        End If
    End Sub
    Private Sub ClearsequenceAfterCreatingVariableOccasion()
        GroupBox7.Enabled = False
        GroupBox8.Enabled = False
        GroupBox11.Visible = True
        GroupBox11.Enabled = True
        BunifuDropdown10.Items.Clear()
        BunifuDropdown10.Text = Nothing
        BunifuPanel4.Enabled = False
        BunifuButton1.Visible = True
        BunifuButton1.Enabled = True
        BunifuDropdown10.Items.Add("From External")
        BunifuDropdown10.Items.Add("From Internal")
    End Sub
    Private UnidentifiedEmployees As New List(Of String)
    Private Sub BunifuDropdown10_SelectedIndexChanged(sender As Object, e As EventArgs) Handles BunifuDropdown10.SelectedIndexChanged
        ClearsequenceAfterSelectingSourceselection()
        If BunifuDropdown10.SelectedItem = "From Internal" Then
            CheckBox3.Enabled = True
            CheckBox3.Visible = True
            BackgroundWorker2.WorkerSupportsCancellation = True
            BackgroundWorker2.RunWorkerAsync()
        ElseIf BunifuDropdown10.SelectedItem = "From External" Then
            CheckBox3.Enabled = False
            CheckBox3.Visible = False
            BackgroundWorker5.WorkerSupportsCancellation = True
            BackgroundWorker5.RunWorkerAsync()
        End If
    End Sub
    Private Sub BackgroundWorker2_DoWork(sender As Object, e As DoWorkEventArgs) Handles BackgroundWorker2.DoWork
        Dim worker As BackgroundWorker = CType(sender, BackgroundWorker)
        GroupBox9.Invoke(Sub()
                             Try
                                 ClearsequenceAfterSelectingSourceselection()
                                 Dim result = FuncLib.Events.FetchInternalDbForNewVariableOccasion()
                                 If result.Item1 IsNot Nothing Then
                                     Dim grid = result.Item1
                                     Dim TotalRowsCount As Integer = grid.Rows.Count
                                     ProgressBar1.Maximum = TotalRowsCount
                                     Dim CurrentRow As Integer = 0
                                     Dim checkBoxColumn As New DataGridViewCheckBoxColumn()
                                     checkBoxColumn.HeaderText = "Select"
                                     checkBoxColumn.Width = 5
                                     checkBoxColumn.Name = "Select"
                                     If BunifuDataGridView4.Columns.Contains("Select") Then
                                         grid.Columns("Employee_ID").ColumnName = "EMP ID"
                                         grid.Columns("NT_ID").ColumnName = "NT ID"
                                         grid.Columns("Scanner_ID").ColumnName = "SCAN ID"
                                         grid.Columns("Scanner_5D").ColumnName = "SCAN 5D"
                                         grid.Columns("Employee_Name").ColumnName = "NAME"
                                         grid.Columns("Gender").ColumnName = "GENDER"
                                         grid.Columns.Add("DATE", GetType(Date))
                                         grid.Columns("Buisness_Area").ColumnName = "BUISNESS"
                                         grid.Columns("Functions").ColumnName = "DEPARTMENT"
                                         grid.Columns("Leave_Approver").ColumnName = "REPORTING MANAGER"
                                         grid.Columns.Add("TIME STAMP", GetType(DateTime))
                                         Dim fixedDate As New Date
                                         fixedDate = BunifuDatePicker1.Value.Date
                                         For Each row As DataRow In grid.Rows
                                             row("DATE") = fixedDate
                                             row("TIME STAMP") = DateTime.Now.ToString("G")
                                             CurrentRow += 1
                                             ProgressBar1.Value = CurrentRow
                                         Next
                                         BunifuDataGridView4.DataSource = grid
                                         GroupBox9.Visible = True
                                         GroupBox9.Enabled = True
                                         BunifuDataGridView4.Visible = True
                                         BunifuDataGridView4.Enabled = True
                                         BunifuButton5.Visible = True
                                         BunifuButton5.Enabled = True
                                         GroupBox12.Visible = True
                                         GroupBox12.Enabled = True
                                         For Each column As DataGridViewColumn In BunifuDataGridView4.Columns
                                             BunifuDropdown11.Items.Add(column.HeaderText)
                                         Next
                                     Else
                                         BunifuDataGridView4.Columns.Insert(0, checkBoxColumn)
                                         grid.Columns("Employee_ID").ColumnName = "EMP ID"
                                         grid.Columns("NT_ID").ColumnName = "NT ID"
                                         grid.Columns("Scanner_ID").ColumnName = "SCAN ID"
                                         grid.Columns("Scanner_5D").ColumnName = "SCAN 5D"
                                         grid.Columns("Employee_Name").ColumnName = "NAME"
                                         grid.Columns("Gender").ColumnName = "GENDER"
                                         grid.Columns.Add("DATE", GetType(Date))
                                         grid.Columns("Buisness_Area").ColumnName = "BUISNESS"
                                         grid.Columns("Functions").ColumnName = "DEPARTMENT"
                                         grid.Columns("Leave_Approver").ColumnName = "REPORTING MANAGER"
                                         grid.Columns.Add("TIME STAMP", GetType(DateTime))
                                         Dim fixedDate As New Date
                                         fixedDate = BunifuDatePicker1.Value.Date
                                         For Each row As DataRow In grid.Rows
                                             row("DATE") = fixedDate
                                             row("TIME STAMP") = DateTime.Now.ToString("G")
                                             CurrentRow += 1
                                             ProgressBar1.Value = CurrentRow
                                         Next
                                         BunifuDataGridView4.DataSource = grid
                                         GroupBox9.Visible = True
                                         GroupBox9.Enabled = True
                                         BunifuDataGridView4.Visible = True
                                         BunifuDataGridView4.Enabled = True
                                         BunifuButton5.Visible = True
                                         BunifuButton5.Enabled = True
                                         GroupBox12.Visible = True
                                         GroupBox12.Enabled = True
                                         For Each column As DataGridViewColumn In BunifuDataGridView4.Columns
                                             BunifuDropdown11.Items.Add(column.HeaderText)
                                         Next
                                     End If
                                 End If
                             Catch ex As Exception
                                 BunifuSnackbar1.Show(Me, $"Error Occured While Adding Data From Internal Database.{Environment.NewLine}Error :{ex.Message}", Bunifu.UI.WinForms.BunifuSnackbar.MessageTypes.Error, 2000)
                                 FuncLib.WriteLog.WriteErrorLog($"Error Occured While Executing FuncLib.Events.FetchInternalDbForNewVariableOccasion : {ex.Message}")
                                 e.Cancel = True
                             End Try
                         End Sub)
    End Sub
    Private Sub BackgroundWorker2_RunWorkerCompleted(sender As Object, e As RunWorkerCompletedEventArgs) Handles BackgroundWorker2.RunWorkerCompleted
        If e.Cancelled = True OrElse e.Error IsNot Nothing Then
            BunifuSnackbar1.Show(Me, $"Error Occured While Adding Data From Internal Database.{Environment.NewLine}Error :{e.Error.ToString}", Bunifu.UI.WinForms.BunifuSnackbar.MessageTypes.Error, 2000)
            FuncLib.WriteLog.WriteErrorLog($"Error Occured While Executing FuncLib.Events.FetchInternalDbForNewVariableOccasion : {e.Error.ToString}")
            FuncLib.Events.DeleteVariableFunctionDateFromIni(BunifuDropdown9.SelectedItem, BunifuDatePicker1.Value.Date)
            RebbotInitiaizer.TextBox1.Text = 5
            RebbotInitiaizer.ShowDialog()
        Else
            BunifuSnackbar1.Show(Me, $"All Avialable Data Succesfully Fetched!", Bunifu.UI.WinForms.BunifuSnackbar.MessageTypes.Success, 2000)
        End If
    End Sub
    Private Sub BackgroundWorker5_DoWork(sender As Object, e As DoWorkEventArgs) Handles BackgroundWorker5.DoWork
        Dim worker As BackgroundWorker = CType(sender, BackgroundWorker)
        GroupBox9.Invoke(Sub()
                             Try
                                 UnidentifiedEmployees.Clear()
                                 Dim openFileDialog As New OpenFileDialog()
                                 openFileDialog.Filter = "Text Files|*.txt"
                                 openFileDialog.Title = "Select Employee Data"
                                 openFileDialog.Multiselect = False
                                 If openFileDialog.ShowDialog() = DialogResult.OK Then
                                     Dim selectedFilePath As String = openFileDialog.FileName
                                     Dim lines As String() = System.IO.File.ReadAllLines(selectedFilePath)
                                     Dim CurrentRow As Integer = 0
                                     Dim FoundEmployees As Integer = 0
                                     Dim NotFoundEmployees As Integer = 0
                                     ProgressBar1.Maximum = lines.Length()
                                     ProgressBar1.Value = 0
                                     Dim dt As New DataTable()
                                     dt.Columns.AddRange({New DataColumn("EMP ID"), New DataColumn("NT ID"), New DataColumn("SCAN ID"), New DataColumn("SCAN 5D"), New DataColumn("NAME"), New DataColumn("GENDER"), New DataColumn("DATE"), New DataColumn("BUISNESS"), New DataColumn("DEPARTMENT"), New DataColumn("REPORTING MANAGER"), New DataColumn("TIME STAMP")})
                                     If BunifuDataGridView4.Columns.Contains("Select") Then
                                         BunifuDataGridView4.Columns.Remove("Select")
                                     End If
                                     GroupBox9.Visible = True
                                     GroupBox9.Enabled = True
                                     BunifuDataGridView4.Visible = True
                                     BunifuDataGridView4.Enabled = True
                                     BunifuButton5.Visible = True
                                     BunifuButton5.Enabled = True
                                     GroupBox12.Visible = True
                                     GroupBox12.Enabled = True
                                     For Each line As String In lines
                                         Dim result = FuncLib.Events.CheckAndReturnEmployeeDetailsByEmployeeID(line, BunifuDatePicker1.Value.Date)
                                         If result.Item1 = "True" AndAlso result.Item3 IsNot Nothing AndAlso result.Item2.Count < 1 Then
                                             Dim row = dt.NewRow()
                                             row("EMP ID") = result.Item3.Rows(0)("Employee_ID")
                                             row("NT ID") = result.Item3.Rows(0)("NT_ID")
                                             row("SCAN ID") = result.Item3.Rows(0)("Scanner_ID")
                                             row("SCAN 5D") = result.Item3.Rows(0)("Scanner_5D")
                                             row("NAME") = result.Item3.Rows(0)("Employee_Name")
                                             row("GENDER") = result.Item3.Rows(0)("Gender")
                                             row("DATE") = result.Item3.Rows(0)("Celebration_Date")
                                             row("BUISNESS") = result.Item3.Rows(0)("Buisness_Area")
                                             row("DEPARTMENT") = result.Item3.Rows(0)("Functions")
                                             row("REPORTING MANAGER") = result.Item3.Rows(0)("Leave_Approver")
                                             row("TIME STAMP") = result.Item3.Rows(0)("Created_At")
                                             dt.Rows.Add(row)
                                         ElseIf result.Item1 = "True" AndAlso result.Item3 IsNot Nothing AndAlso result.Item2.Count >= 1 Then
                                             UnidentifiedEmployees.Add(line)
                                         Else
                                             BunifuSnackbar1.Show(Me, $"Error Occured While Adding COmparing Employees!{Environment.NewLine}Error:{result.Item1}", Bunifu.UI.WinForms.BunifuSnackbar.MessageTypes.Error, 2500)
                                             FuncLib.WriteLog.WriteErrorLog($"Error Occured While executing FuncLib.Events.CheckAndReturnEmployeeDetailsByEmployeeID : {result.Item1}")
                                             e.Cancel = True
                                             Exit For
                                         End If
                                         CurrentRow += 1
                                         ProgressBar1.Value = CurrentRow
                                         BunifuDataGridView4.DataSource = dt
                                     Next
                                     For Each column As DataGridViewColumn In BunifuDataGridView4.Columns
                                         BunifuDropdown11.Items.Add(column.HeaderText)
                                     Next
                                     If UnidentifiedEmployees.Count > 0 Then
                                         EmpDet.TextBox1.Clear()
                                         EmpDet.ListBox1.Items.Clear()
                                         For Each str As String In UnidentifiedEmployees
                                             EmpDet.ListBox1.Items.Add(str)
                                         Next
                                         EmpDet.TextBox1.Text = UnidentifiedEmployees.Count
                                         EmpDet.TextBox1.Enabled = False
                                         EmpDet.ShowDialog()
                                     End If
                                 End If
                             Catch ex As Exception
                                 BunifuSnackbar1.Show(Me, $"Error Occured While Adding Data From Internal Database.{Environment.NewLine}Error :{ex.Message}", Bunifu.UI.WinForms.BunifuSnackbar.MessageTypes.Error, 2000)
                                 FuncLib.WriteLog.WriteErrorLog($"Error Occured While Executing FuncLib.Events.CheckAndReturnEmployeeDetailsByEmployeeID : {ex.Message}")
                                 e.Cancel = True
                             End Try
                         End Sub)
    End Sub
    Private Sub BackgroundWorker5_RunWorkerCompleted(sender As Object, e As RunWorkerCompletedEventArgs) Handles BackgroundWorker5.RunWorkerCompleted
        If e.Cancelled = True OrElse e.Error IsNot Nothing Then
            BunifuSnackbar1.Show(Me, $"Error Occured While Adding Data From Internal Database.{Environment.NewLine}Error :{e.Error.ToString}", Bunifu.UI.WinForms.BunifuSnackbar.MessageTypes.Error, 2000)
            FuncLib.WriteLog.WriteErrorLog($"Error Occured While Executing FuncLib.Events.FetchInternalDbForNewVariableOccasion : {e.Error.ToString}")
            FuncLib.Events.DeleteVariableFunctionDateFromIni(BunifuDropdown9.SelectedItem, BunifuDatePicker1.Value.Date)
            RebbotInitiaizer.TextBox1.Text = 5
            RebbotInitiaizer.ShowDialog()
        Else
            BunifuSnackbar1.Show(Me, $"All Avialable Data Succesfully Fetched!", Bunifu.UI.WinForms.BunifuSnackbar.MessageTypes.Success, 2000)
        End If
    End Sub
    Private Sub ClearsequenceAfterSelectingSourceselection()
        ProgressBar1.Visible = True
        ProgressBar1.Minimum = 0
        ProgressBar1.Value = 0
        GroupBox9.Visible = False
        GroupBox9.Enabled = False
        BunifuDataGridView4.Visible = False
        BunifuDataGridView4.Enabled = False
        BunifuDataGridView4.DataSource = Nothing
        BunifuButton5.Visible = False
        BunifuButton5.Enabled = False
        GroupBox12.Visible = False
        BunifuDropdown11.Items.Clear()
        BunifuDropdown11.Text = Nothing
        BunifuTextBox10.Clear()
        CheckBox3.Checked = False
    End Sub
    Private Sub CheckBox3_CheckedChanged(sender As Object, e As EventArgs) Handles CheckBox3.CheckedChanged
        Dim isChecked As Boolean = CheckBox3.Checked
        For i As Integer = 0 To BunifuDataGridView4.RowCount - 2
            Dim checkboxCell As DataGridViewCheckBoxCell = TryCast(BunifuDataGridView4.Rows(i).Cells("Select"), DataGridViewCheckBoxCell)
            If checkboxCell IsNot Nothing Then
                checkboxCell.Value = isChecked
            End If
            Dim rows As DataGridViewRow = BunifuDataGridView4.Rows(i)
            If CheckBox3.Checked = True Then
                selectedRows.Add(rows)
            Else
                selectedRows.Remove(rows)
            End If
        Next
    End Sub
    Private Sub BunifuTextBox10_TextChanged_1(sender As Object, e As EventArgs) Handles BunifuTextBox10.TextChanged
        Dim selectedColumnIndex As Integer = BunifuDropdown11.SelectedIndex
        Dim columnName As String = If(selectedColumnIndex >= 0, BunifuDropdown11.Items(selectedColumnIndex).ToString(), "")
        If BunifuDataGridView4.DataSource IsNot Nothing AndAlso Not String.IsNullOrEmpty(columnName) Then
            Dim bindingSource As New BindingSource()
            bindingSource.DataSource = BunifuDataGridView4.DataSource
            bindingSource.Filter = $"[{columnName}] LIKE '%{BunifuTextBox10.Text}%'"
            BunifuDataGridView4.DataSource = bindingSource
        End If
    End Sub
    Private selectedRows As New List(Of DataGridViewRow)
    Private Sub BunifuDataGridView4_CellContentClick(sender As Object, e As DataGridViewCellEventArgs) Handles BunifuDataGridView4.CellContentClick
        If BunifuDropdown10.SelectedItem = "From Internal" Then
            If e.ColumnIndex = BunifuDataGridView4.Columns("Select").Index AndAlso e.RowIndex >= 0 Then
                Dim checkboxCell As DataGridViewCheckBoxCell = TryCast(BunifuDataGridView4.Rows(e.RowIndex).Cells("Select"), DataGridViewCheckBoxCell)
                If checkboxCell IsNot Nothing Then
                    Dim row As DataGridViewRow = BunifuDataGridView4.Rows(e.RowIndex)
                    Dim isChecked As Boolean = Convert.ToBoolean(checkboxCell.Value)
                    If BunifuDataGridView4.Rows(e.RowIndex).Cells("Select").Value = isChecked Then
                        If selectedRows.Contains(row) Then
                            selectedRows.Remove(row)
                        Else
                            selectedRows.Add(row)
                        End If
                    End If
                End If
            End If
        End If
    End Sub
    Private Sub BunifuButton5_Click(sender As Object, e As EventArgs) Handles BunifuButton5.Click
        If BunifuDropdown10.SelectedItem = "From Internal" Then
            BackgroundWorker6.WorkerSupportsCancellation = True
            BackgroundWorker6.RunWorkerAsync()
        ElseIf BunifuDropdown10.SelectedItem = "From External" Then
            BackgroundWorker7.WorkerSupportsCancellation = True
            BackgroundWorker7.RunWorkerAsync()
        End If
    End Sub
    Private Sub BackgroundWorker6_DoWork(sender As Object, e As DoWorkEventArgs) Handles BackgroundWorker6.DoWork
        Dim worker As BackgroundWorker = CType(sender, BackgroundWorker)
        GroupBox9.Invoke(Sub()
                             If selectedRows.Count > 0 Then
                                 ProgressBar1.Minimum = 0
                                 ProgressBar1.Maximum = selectedRows.Count
                                 ProgressBar1.Value = 0
                                 Dim CurrentRow As Integer = 0
                                 For Each row As DataGridViewRow In selectedRows
                                     Try
                                         Dim Employee_ID As String = row.Cells("EMP ID").Value.ToString()
                                         Dim NT_ID As String = row.Cells("NT ID").Value.ToString()
                                         Dim Scanner_5D As String = row.Cells("SCAN 5D").Value
                                         Dim Scanner_ID As String = row.Cells("SCAN ID").Value
                                         Dim Employee_Name As String = row.Cells("NAME").Value.ToString()
                                         Dim Gender As String = row.Cells("GENDER").Value.ToString()
                                         Dim Celebration_Date As Date = DateTime.Parse(row.Cells("DATE").Value.ToString())
                                         Dim Buisness_Area As String = row.Cells("BUISNESS").Value.ToString()
                                         Dim Functions As String = row.Cells("DEPARTMENT").Value.ToString()
                                         Dim Leave_Approver As String = row.Cells("REPORTING MANAGER").Value.ToString()
                                         Dim Created_At As DateTime = row.Cells("TIME STAMP").Value
                                         FuncLib.Events.ImportDataIntoVariableFunction(BunifuDropdown9.SelectedItem, Employee_ID, NT_ID, Scanner_5D, Scanner_ID, Employee_Name, Gender, Celebration_Date, Buisness_Area, Functions, Leave_Approver, Created_At)
                                         CurrentRow += 1
                                         ProgressBar1.Value = CurrentRow
                                     Catch ex As Exception
                                         BunifuSnackbar1.Show(Me, $"Error Occured While Executing FuncLib.Events.ImportDataIntoVariableFunction. {Environment.NewLine} Error : {ex.Message}", Bunifu.UI.WinForms.BunifuSnackbar.MessageTypes.Error, 3000)
                                         FuncLib.WriteLog.WriteErrorLog($"Error Occured While Executing FuncLib.Events.ImportDataIntoVariableFunction :  {ex.Message}")
                                         FuncLib.Events.DeletevariableFunctionDate(BunifuDropdown9.SelectedItem, BunifuDatePicker1.Value.Date)
                                         e.Cancel = True
                                         Exit For
                                     End Try
                                 Next
                             Else
                                 BunifuSnackbar1.Show(Me, $"No Employee Selected For Import!", Bunifu.UI.WinForms.BunifuSnackbar.MessageTypes.Information, 3000)
                                 e.Cancel = True
                             End If
                         End Sub)
    End Sub
    Private Sub BackgroundWorker6_RunWorkerCompleted(sender As Object, e As RunWorkerCompletedEventArgs) Handles BackgroundWorker6.RunWorkerCompleted
        If e.Cancelled = True Then
            Exit Sub
        ElseIf e.Error IsNot Nothing Then
            BunifuSnackbar1.Show(Me, $"Error Occured While Executing FuncLib.Events.ImportDataIntoVariableFunction. {Environment.NewLine} Error : {e.Error}", Bunifu.UI.WinForms.BunifuSnackbar.MessageTypes.Error, 3000)
            FuncLib.WriteLog.WriteErrorLog($"Error Occured While Executing FuncLib.Events.ImportDataIntoVariableFunction :  {e.Error}")
            FuncLib.Events.DeletevariableFunctionDate(BunifuDropdown9.SelectedItem, BunifuDatePicker1.Value.Date)
            Exit Sub
        Else
            BunifuSnackbar1.Show(Me, $"All Avialable Data Succesfully Fetched!", Bunifu.UI.WinForms.BunifuSnackbar.MessageTypes.Success, 2000)
            CLearSequenceAfterSelcetionOfEvents()
            RebbotInitiaizer.TextBox1.Text = 5
            RebbotInitiaizer.BackColor = Color.LimeGreen
            RebbotInitiaizer.ShowDialog()
        End If
    End Sub
    Private Sub BackgroundWorker7_DoWork(sender As Object, e As DoWorkEventArgs) Handles BackgroundWorker7.DoWork
        Dim worker As BackgroundWorker = CType(sender, BackgroundWorker)
        GroupBox9.Invoke(Sub()
                             If BunifuDataGridView4.Rows.Count > 0 Then
                                 ProgressBar1.Minimum = 0
                                 ProgressBar1.Maximum = BunifuDataGridView4.Rows.Count - 1
                                 ProgressBar1.Value = 0
                                 Dim CurrentRow As Integer = 0
                                 For i As Integer = 0 To BunifuDataGridView4.RowCount - 2
                                     Dim row As DataGridViewRow = BunifuDataGridView4.Rows(i)
                                     Try
                                         Dim Employee_ID As String = row.Cells("EMP ID").Value.ToString()
                                         Dim NT_ID As String = row.Cells("NT ID").Value.ToString()
                                         Dim Scanner_5D As String = row.Cells("SCAN 5D").Value
                                         Dim Scanner_ID As String = row.Cells("SCAN ID").Value
                                         Dim Employee_Name As String = row.Cells("NAME").Value.ToString()
                                         Dim Gender As String = row.Cells("GENDER").Value.ToString()
                                         Dim Celebration_Date As Date = DateTime.Parse(row.Cells("DATE").Value.ToString())
                                         Dim Buisness_Area As String = row.Cells("BUISNESS").Value.ToString()
                                         Dim Functions As String = row.Cells("DEPARTMENT").Value.ToString()
                                         Dim Leave_Approver As String = row.Cells("REPORTING MANAGER").Value.ToString()
                                         Dim Created_At As DateTime = row.Cells("TIME STAMP").Value
                                         FuncLib.Events.ImportDataIntoVariableFunction(BunifuDropdown9.SelectedItem, Employee_ID, NT_ID, Scanner_5D, Scanner_ID, Employee_Name, Gender, Celebration_Date, Buisness_Area, Functions, Leave_Approver, Created_At)
                                         CurrentRow += 1
                                         ProgressBar1.Value = CurrentRow
                                     Catch ex As Exception
                                         BunifuSnackbar1.Show(Me, $"Error Occured While Executing FuncLib.Events.ImportDataIntoVariableFunction. {Environment.NewLine} Error : {ex.Message}", Bunifu.UI.WinForms.BunifuSnackbar.MessageTypes.Success, 3000)
                                         FuncLib.WriteLog.WriteErrorLog($"Error Occured While Executing FuncLib.Events.ImportDataIntoVariableFunction :  {ex.Message}")
                                         FuncLib.Events.DeletevariableFunctionDate(BunifuDropdown9.SelectedItem, BunifuDatePicker1.Value.Date)
                                         e.Cancel = True
                                         Exit For
                                     End Try
                                 Next
                             End If
                         End Sub)
    End Sub
    Private Sub BackgroundWorker7_RunWorkerCompleted(sender As Object, e As RunWorkerCompletedEventArgs) Handles BackgroundWorker7.RunWorkerCompleted
        If e.Cancelled = True Then
            Exit Sub
        ElseIf e.Error IsNot Nothing Then
            BunifuSnackbar1.Show(Me, $"Error Occured While Executing FuncLib.Events.ImportDataIntoVariableFunction. {Environment.NewLine} Error : {e.Error}", Bunifu.UI.WinForms.BunifuSnackbar.MessageTypes.Error, 3000)
            FuncLib.WriteLog.WriteErrorLog($"Error Occured While Executing FuncLib.Events.ImportDataIntoVariableFunction :  {e.Error}")
            FuncLib.Events.DeletevariableFunctionDate(BunifuDropdown9.SelectedItem, BunifuDatePicker1.Value.Date)
            Exit Sub
        Else
            BunifuSnackbar1.Show(Me, $"All Avialable Data Succesfully Fetched!", Bunifu.UI.WinForms.BunifuSnackbar.MessageTypes.Success, 2000)
            CLearSequenceAfterSelcetionOfEvents()
            RebbotInitiaizer.TextBox1.Text = 5
            RebbotInitiaizer.BackColor = Color.LimeGreen
            RebbotInitiaizer.ShowDialog()
        End If
    End Sub
    Private Sub BunifuButton1_Click(sender As Object, e As EventArgs) Handles BunifuButton1.Click
        If CheckBox2.Checked = True Then
            FuncLib.Events.DeleteVariableFunctionDateFromIni(BunifuDropdown9.SelectedItem, BunifuDatePicker1.Value.Date)
            CLearSequenceAfterSelcetionOfEvents()
        Else
            CLearSequenceAfterSelcetionOfEvents()
        End If
    End Sub
    Private Sub BunifuDropdown4_SelectedIndexChanged(sender As Object, e As EventArgs) Handles BunifuDropdown4.SelectedIndexChanged
        BunifuDropdown4.Text = "HISTORY"
        If BunifuDropdown4.SelectedItem = "Gift Histoiry" Then
            If ListBox1.Items(16) = "True" Then
                BunifuDropdown14.Items.Clear()
                BunifuDropdown14.Text = Nothing
                BunifuDropdown15.Items.Clear()
                BunifuDropdown15.Text = "SAVE AS"
                BunifuDropdown15.Items.Add(".XLS")
                BunifuDropdown15.Items.Add(".PDF")
                Dim result = FuncLib.History.FetchAllHistory()
                If result IsNot Nothing AndAlso result.Item1 IsNot Nothing Then
                    Dim dataTable = result.Item1
                    If dataTable.Rows.Count > 0 Then
                        If TabControl1.Visible = True Then
                            transition.HideSync(TabControl1, False, BunifuAnimatorNS.Animation.HorizBlind)
                            TabControl1.TabPages.Clear()
                            TabControl1.TabPages.Add(TabPage8)
                            transition.ShowSync(TabControl1, False, BunifuAnimatorNS.Animation.HorizBlind)
                        Else
                            TabControl1.TabPages.Clear()
                            TabControl1.TabPages.Add(TabPage8)
                            transition.ShowSync(TabControl1, False, BunifuAnimatorNS.Animation.HorizBlind)
                        End If
                        dataTable.Columns("Employee_ID").ColumnName = "EMP ID"
                        dataTable.Columns("NT_ID").ColumnName = "NT ID"
                        dataTable.Columns("Employee_Name").ColumnName = "NAME"
                        dataTable.Columns("Buisness_Area").ColumnName = "BUISNESS"
                        dataTable.Columns("Functions").ColumnName = "DEPARTMENT"
                        dataTable.Columns("Gift_Type").ColumnName = "GIFT"
                        dataTable.Columns("GivenBy_ID").ColumnName = "GIVENBY_ID"
                        dataTable.Columns("GivenBy_NT").ColumnName = "GIVENBY_NT"
                        dataTable.Columns("GivenBy_Name").ColumnName = "GIVENBY_NAME"
                        dataTable.Columns("Shared_At").ColumnName = "TIME"
                        BunifuDataGridView6.DataSource = dataTable
                        For Each column As DataGridViewColumn In BunifuDataGridView6.Columns
                            BunifuDropdown14.Items.Add(column.HeaderText)
                        Next
                    Else
                        BunifuSnackbar1.Show(Me, $"No History Found!", Bunifu.UI.WinForms.BunifuSnackbar.MessageTypes.Warning, 3000)
                        Exit Sub
                    End If
                Else
                    BunifuSnackbar1.Show(Me, $"Error Occured While Executing FuncLib.History.FetchAllHistory(). {Environment.NewLine} Error : {result.Item2}", Bunifu.UI.WinForms.BunifuSnackbar.MessageTypes.Error, 3000)
                    FuncLib.WriteLog.WriteErrorLog($"Error Occured While Executing FuncLib.History.FetchAllHistory() : {result.Item2}")
                    Exit Sub
                End If
            Else
                BunifuSnackbar1.Show(Me, "You Are Not Authorised To View History", Bunifu.UI.WinForms.BunifuSnackbar.MessageTypes.Warning, 3000)
                Exit Sub
            End If
        ElseIf BunifuDropdown4.SelectedItem = "Pending History" Then
            If ListBox1.Items(16) = "True" Then
                GroupBox13.Enabled = True
                GroupBox14.Visible = False
                GroupBox16.Visible = False
                Dim Occasion As List(Of String) = FuncLib.Gifts.GetOcaasioList()
                If Occasion.Count <= 0 AndAlso TypeOf Occasion(0) Is String Then
                    Dim errorMessage As String = Occasion(0)
                    BunifuSnackbar1.Show(Me, $"Error While Fetching Occasion List From. {Environment.NewLine} Error : {errorMessage}", Bunifu.UI.WinForms.BunifuSnackbar.MessageTypes.Error, 3000)
                    FuncLib.WriteLog.WriteErrorLog($"Error occured While Trying To Fetch Occasion List : {errorMessage}")
                    Exit Sub
                Else
                    BunifuDropdown12.Items.Clear()
                    BunifuDropdown12.Text = Nothing
                    BunifuDropdown12.Items.AddRange(Occasion.ToArray())
                    If TabControl1.Visible = True Then
                        transition.HideSync(TabControl1, False, BunifuAnimatorNS.Animation.HorizBlind)
                        TabControl1.TabPages.Clear()
                        TabControl1.TabPages.Add(TabPage7)
                        transition.ShowSync(TabControl1, False, BunifuAnimatorNS.Animation.HorizBlind)
                    Else
                        TabControl1.TabPages.Clear()
                        TabControl1.TabPages.Add(TabPage7)
                        transition.ShowSync(TabControl1, False, BunifuAnimatorNS.Animation.HorizBlind)
                    End If
                End If
            Else
                BunifuSnackbar1.Show(Me, "You Are Not Authorised To View History", Bunifu.UI.WinForms.BunifuSnackbar.MessageTypes.Warning, 3000)
                Exit Sub
            End If
        End If
    End Sub
    Private Sub BunifuTextBox11_TextChanged(sender As Object, e As EventArgs) Handles BunifuTextBox11.TextChanged
        Dim selectedColumnIndex As Integer = BunifuDropdown14.SelectedIndex
        Dim columnName As String = If(selectedColumnIndex >= 0, BunifuDropdown14.Items(selectedColumnIndex).ToString(), "")
        If BunifuDataGridView6.DataSource IsNot Nothing AndAlso Not String.IsNullOrEmpty(columnName) Then
            Dim bindingSource As New BindingSource()
            bindingSource.DataSource = BunifuDataGridView6.DataSource
            bindingSource.Filter = $"[{columnName}] LIKE '%{BunifuTextBox11.Text}%'"
            BunifuDataGridView6.DataSource = bindingSource
        End If
    End Sub
    Private Sub BunifuDatePicker2_ValueChanged(sender As Object, e As EventArgs) Handles BunifuDatePicker2.ValueChanged
        Dim startdate As Date = BunifuDatePicker2.Value.Date
        Dim enddate As Date = BunifuDatePicker3.Value.Date
        If startdate <= enddate Then
            Dim result = FuncLib.History.FilterDateWise(startdate, enddate)
            If result IsNot Nothing AndAlso result.Item1 IsNot Nothing Then
                BunifuDataGridView6.Visible = True
                Dim dataTable = result.Item1
                If dataTable.Rows.Count > 0 Then
                    dataTable.Columns("Employee_ID").ColumnName = "EMP ID"
                    dataTable.Columns("NT_ID").ColumnName = "NT ID"
                    dataTable.Columns("Employee_Name").ColumnName = "NAME"
                    dataTable.Columns("Buisness_Area").ColumnName = "BUISNESS"
                    dataTable.Columns("Functions").ColumnName = "DEPARTMENT"
                    dataTable.Columns("Gift_Type").ColumnName = "GIFT"
                    dataTable.Columns("GivenBy_ID").ColumnName = "GIVENBY_ID"
                    dataTable.Columns("GivenBy_NT").ColumnName = "GIVENBY_NT"
                    dataTable.Columns("GivenBy_Name").ColumnName = "GIVENBY_NAME"
                    dataTable.Columns("Shared_At").ColumnName = "TIME"
                    BunifuDataGridView6.DataSource = dataTable

                Else
                    BunifuDataGridView6.DataSource = Nothing
                End If
            End If
        End If
    End Sub
    Private Sub BunifuDatePicker3_ValueChanged(sender As Object, e As EventArgs) Handles BunifuDatePicker3.ValueChanged
        Dim startdate As Date = BunifuDatePicker2.Value.Date
        Dim enddate As Date = BunifuDatePicker3.Value.Date
        If enddate >= startdate Then
            Dim STDate As String = startdate.ToString("MM/dd/yyyy")
            Dim EndDt As String = enddate.ToString("MM/dd/yyyy")
            Dim result = FuncLib.History.FilterDateWise(STDate, EndDt)
            If result IsNot Nothing AndAlso result.Item1 IsNot Nothing Then
                BunifuDataGridView6.Visible = True
                Dim dataTable = result.Item1
                If dataTable.Rows.Count > 0 Then
                    dataTable.Columns("Employee_ID").ColumnName = "EMP ID"
                    dataTable.Columns("NT_ID").ColumnName = "NT ID"
                    dataTable.Columns("Employee_Name").ColumnName = "NAME"
                    dataTable.Columns("Buisness_Area").ColumnName = "BUISNESS"
                    dataTable.Columns("Functions").ColumnName = "DEPARTMENT"
                    dataTable.Columns("Gift_Type").ColumnName = "GIFT"
                    dataTable.Columns("GivenBy_ID").ColumnName = "GIVENBY_ID"
                    dataTable.Columns("GivenBy_NT").ColumnName = "GIVENBY_NT"
                    dataTable.Columns("GivenBy_Name").ColumnName = "GIVENBY_NAME"
                    dataTable.Columns("Shared_At").ColumnName = "TIME"
                    BunifuDataGridView6.DataSource = dataTable
                Else
                    BunifuDataGridView6.DataSource = Nothing
                End If
            End If
        End If
    End Sub
    Private Sub BunifuDropdown15_SelectedIndexChanged(sender As Object, e As EventArgs) Handles BunifuDropdown15.SelectedIndexChanged
        BunifuDropdown15.Text = "SAVE AS"
        If BunifuDropdown15.SelectedItem = ".XLS" Then
            Dim saveFileDialog As New SaveFileDialog()
            saveFileDialog.Filter = "Excel File (*.csv)|*.csv|All Files (*.*)|*.*"
            saveFileDialog.FilterIndex = 1
            saveFileDialog.RestoreDirectory = True
            If saveFileDialog.ShowDialog() = DialogResult.OK Then
                Dim fileName As String = saveFileDialog.FileName
                Dim SaveFileasExcel As String = FuncLib.History.SaveasExcel(BunifuDataGridView6, $"{fileName}")
                If SaveFileasExcel = "True" Then
                    BunifuSnackbar1.Show(Me, $"History Successfully saved as {fileName}", Bunifu.UI.WinForms.BunifuSnackbar.MessageTypes.Success, 3000)
                    Exit Sub
                Else
                    BunifuSnackbar1.Show(Me, $"Error Saving File as {fileName}. {Environment.NewLine} Error :{SaveFileasExcel}", Bunifu.UI.WinForms.BunifuSnackbar.MessageTypes.Error, 3000)
                    FuncLib.WriteLog.WriteErrorLog($"Error Occured While Executing FuncLib.History.SaveasExcel():{SaveFileasExcel}")
                    Exit Sub
                End If
            End If
        ElseIf BunifuDropdown15.SelectedItem = ".PDF" Then
            Dim saveFileDialog As New SaveFileDialog()
            saveFileDialog.Filter = "pdf File (*.pdf)|*.pdf|All Files (*.*)|*.*"
            saveFileDialog.FilterIndex = 1
            saveFileDialog.RestoreDirectory = True
            If saveFileDialog.ShowDialog() = DialogResult.OK Then
                Dim fileName As String = saveFileDialog.FileName
                Dim SaveFileaspdf As String = FuncLib.History.SaveAsPdf(BunifuDataGridView6, $"{fileName}")
                If SaveFileaspdf = "True" Then
                    BunifuSnackbar1.Show(Me, $"History Successfully saved as {fileName}", Bunifu.UI.WinForms.BunifuSnackbar.MessageTypes.Success, 3000)
                    Exit Sub
                Else
                    BunifuSnackbar1.Show(Me, $"Error Saving File as {fileName}. {Environment.NewLine} Error :{SaveFileaspdf}", Bunifu.UI.WinForms.BunifuSnackbar.MessageTypes.Error, 3000)
                    FuncLib.WriteLog.WriteErrorLog($"Error Occured While Executing FuncLib.History.SaveasExcel():{SaveFileaspdf}")
                    Exit Sub
                End If
            End If
        End If
    End Sub
    Private Sub SerialPort_DataReceived(sender As Object, e As SerialDataReceivedEventArgs) Handles serialPort.DataReceived
        Dim hexData As String = serialPort.ReadLine()
        Dim extractedString As String = hexData.Substring(1, hexData.Length - 2)
        BunifuTextBox1.Invoke(Sub()
                                  BunifuTextBox1.Text = FuncLib.Gifts.IdentifyEmployee(extractedString, "BY 10D")
                                  If RadioButton4.Checked = True And BunifuTextBox1.Text IsNot Nothing Then 'Constant Function
                                      Dim CheckEilgibilityforBulkConstant = FuncLib.Gifts.CheckEmployeeEligibilityForBulkDistributionConstant(BunifuDropdown7.SelectedItem.ToString, BunifuTextBox1.Text)
                                      If CheckEilgibilityforBulkConstant = "True" Then
                                          BunifuSnackbar1.Show(Me, $"Congratulations {BunifuTextBox4.Text} Is Eligible For {BunifuDropdown7.SelectedItem.ToString} Gift", Bunifu.UI.WinForms.BunifuSnackbar.MessageTypes.Success, 1000)
                                          GroupBox5.Visible = True
                                          BunifuTextBox1.Enabled = False
                                          BunifuButton2.Visible = True
                                          BunifuButton2.Enabled = True
                                          BunifuButton15.Visible = True
                                          BunifuButton15.Enabled = True
                                      ElseIf CheckEilgibilityforBulkConstant = "TLE" Then
                                          BunifuSnackbar1.Show(Me, $"Soory {BunifuTextBox4.Text} Is Not Eligible For {BunifuDropdown7.SelectedItem.ToString} Gift {Environment.NewLine} Reason : Time Limit Excedded", Bunifu.UI.WinForms.BunifuSnackbar.MessageTypes.Warning, 1000)
                                          ClearsequenceaftersharingGift()
                                      ElseIf CheckEilgibilityforBulkConstant Is Nothing Then
                                          BunifuSnackbar1.Show(Me, $"Soory {BunifuTextBox4.Text} Is Not Eligible For {BunifuDropdown7.SelectedItem.ToString} Gift {Environment.NewLine} Reason : Name Not Found On The Occasion List", Bunifu.UI.WinForms.BunifuSnackbar.MessageTypes.Warning, 1000)
                                          ClearsequenceaftersharingGift()
                                      ElseIf CheckEilgibilityforBulkConstant = "Same" Then
                                          BunifuSnackbar1.Show(Me, $"{BunifuTextBox4.Text} Can Not Be Shared.{Environment.NewLine} Reason : You Cannot Share Gift To Yourself ", Bunifu.UI.WinForms.BunifuSnackbar.MessageTypes.Warning, 1000)
                                          ClearsequenceaftersharingGift()
                                      Else
                                          BunifuSnackbar1.Show(Me, $"Error Occured While Trying To Fetch Data From DataBase {Environment.NewLine} Reason : {CheckEilgibilityforBulkConstant} {Environment.NewLine} For More Details See The Error Log", Bunifu.UI.WinForms.BunifuSnackbar.MessageTypes.Error, 1000)
                                          FuncLib.WriteLog.WriteErrorLog($"Error Occured While Trying To Fetch Data From Database :{CheckEilgibilityforBulkConstant}")
                                          ClearsequenceaftersharingGift()
                                      End If
                                  ElseIf RadioButton3.Checked = True And BunifuTextBox1.Text IsNot Nothing Then 'Variable Function
                                      Dim CheckEilgibilityforBulkvariable = FuncLib.Gifts.CheckEmployeeEligibilityForBulkDistributionvariable(BunifuDropdown7.SelectedItem.ToString, BunifuTextBox1.Text, BunifuDropdown8.SelectedItem)
                                      If CheckEilgibilityforBulkvariable = "True" Then
                                          BunifuSnackbar1.Show(Me, $"Congratulations {BunifuTextBox4.Text} Is Eligible For {BunifuDropdown7.SelectedItem.ToString} Gift", Bunifu.UI.WinForms.BunifuSnackbar.MessageTypes.Success, 1000)
                                          GroupBox5.Visible = True
                                          BunifuTextBox1.Enabled = False
                                          BunifuButton2.Visible = True
                                          BunifuButton2.Enabled = True
                                          BunifuButton15.Visible = True
                                          BunifuButton15.Enabled = True
                                      ElseIf CheckEilgibilityforBulkvariable = "TLE" Then
                                          BunifuSnackbar1.Show(Me, $"Soory {BunifuTextBox4.Text} Is Not Eligible For {BunifuDropdown7.SelectedItem.ToString} Gift {Environment.NewLine} Reason : Time Limit Excedded", Bunifu.UI.WinForms.BunifuSnackbar.MessageTypes.Warning, 1000)
                                          ClearsequenceaftersharingGift()
                                      ElseIf CheckEilgibilityforBulkvariable Is Nothing Then
                                          BunifuSnackbar1.Show(Me, $"Soory {BunifuTextBox4.Text} Is Not Eligible For {BunifuDropdown7.SelectedItem.ToString} Gift {Environment.NewLine} Reason : Name Not Found On The Occasion List", Bunifu.UI.WinForms.BunifuSnackbar.MessageTypes.Warning, 1000)
                                          ClearsequenceaftersharingGift()
                                      ElseIf CheckEilgibilityforBulkvariable = "Same" Then
                                          BunifuSnackbar1.Show(Me, $"{BunifuTextBox4.Text} Can Not Be Shared.{Environment.NewLine} Reason : You Cannot Share Gift To Yourself ", Bunifu.UI.WinForms.BunifuSnackbar.MessageTypes.Warning, 3000)
                                          ClearsequenceaftersharingGift()
                                      Else
                                          BunifuSnackbar1.Show(Me, $"Error Occured While Trying To Fetch Data From DataBase {Environment.NewLine} Reason : {CheckEilgibilityforBulkvariable} {Environment.NewLine} For More Details See The Error Log", Bunifu.UI.WinForms.BunifuSnackbar.MessageTypes.Error, 1000)
                                          FuncLib.WriteLog.WriteErrorLog($"Error Occured While Trying To Fetch Data From Database :{CheckEilgibilityforBulkvariable}")
                                          ClearsequenceaftersharingGift()
                                      End If
                                  Else
                                      BunifuSnackbar1.Show(Me, $"{BunifuTextBox1.Text} is Null or Invalid Retry By Entering Correct ID", Bunifu.UI.WinForms.BunifuSnackbar.MessageTypes.Warning, 1000)
                                      ClearsequenceaftersharingGift()
                                      Exit Sub
                                  End If
                              End Sub)
        Thread.Sleep(1000)
        serialPort.DiscardInBuffer()
        serialPort.DiscardOutBuffer()
    End Sub
    Private Sub BunifuDropdown5_SelectedIndexChanged(sender As Object, e As EventArgs) Handles BunifuDropdown5.SelectedIndexChanged
        BunifuDropdown5.Text = "SETUP"
        If ListBox1.Items(16) = "True" Then
            SETUP.StartPosition = FormStartPosition.CenterScreen
            SETUP.ShowDialog()
        Else
            BunifuSnackbar1.Show(Me, $"You Are Not Authorised For This.. Please Contact To Your Reporting Manager!", Bunifu.UI.WinForms.BunifuSnackbar.MessageTypes.Warning, 3000)
        End If

    End Sub
    Private Sub BunifuDropdown12_SelectedIndexChanged(sender As Object, e As EventArgs) Handles BunifuDropdown12.SelectedIndexChanged
        Dim Occasiontype As String = FuncLib.Gifts.CheckTypeOfOcassion(BunifuDropdown12.SelectedItem.ToString)
        If Occasiontype = "Constant" Then
            GroupBox13.Enabled = False
            GroupBox14.Visible = False
            GroupBox16.Visible = True
            BunifuDropdown17.Items.Clear()
            BunifuDropdown17.Text = "SAVE AS"
            BunifuDropdown17.Items.Add(".XLS")
            BunifuDropdown17.Items.Add(".PDF")
            Dim result = FuncLib.History.FetchAllPendingDataFromDBforConstant(BunifuDropdown12.SelectedItem.ToString)
            If result IsNot Nothing AndAlso result.Item1 IsNot Nothing Then
                Dim dataTable = result.Item1
                If dataTable.Rows.Count > 0 Then
                    GroupBox13.Enabled = False
                    dataTable.Columns("Employee_ID").ColumnName = "EMP ID"
                    dataTable.Columns("NT_ID").ColumnName = "NT ID"
                    dataTable.Columns("Employee_Name").ColumnName = "NAME"
                    dataTable.Columns("Buisness_Area").ColumnName = "BUISNESS"
                    dataTable.Columns("Functions").ColumnName = "DEPARTMENT"
                    dataTable.Columns("Created_At").ColumnName = "TIME"
                    BunifuDataGridView5.DataSource = dataTable
                    GroupBox16.Visible = True
                    BunifuDropdown16.Items.Clear()
                    BunifuDropdown16.Text = Nothing
                    BunifuTextBox12.Clear()
                    BunifuTextBox13.Clear()
                    BunifuTextBox13.Text = dataTable.Rows.Count
                    For Each column As DataGridViewColumn In BunifuDataGridView5.Columns
                        BunifuDropdown16.Items.Add(column.HeaderText)
                    Next
                Else
                    BunifuSnackbar1.Show(Me, $"No Pending Employees For {BunifuDropdown12.SelectedItem.ToString}", Bunifu.UI.WinForms.BunifuSnackbar.MessageTypes.Warning, 3000)
                    BunifuDataGridView5.DataSource = Nothing
                    BunifuDropdown16.Items.Clear()
                    BunifuDropdown16.Text = Nothing
                    BunifuTextBox12.Clear()
                    BunifuTextBox13.Clear()
                    GroupBox13.Enabled = True
                    GroupBox16.Visible = False
                    Exit Sub
                End If
            End If
        ElseIf Occasiontype = "Variable" Then
            Dim OccasionDate As List(Of Date) = FuncLib.Gifts.VarableDateQuery(BunifuDropdown12.SelectedItem)
            If OccasionDate.Count <= 0 AndAlso OccasionDate(0) = DateTime.MinValue Then
                BunifuSnackbar1.Show(Me, $"Error Occured While Fetching Occasion Dates!", Bunifu.UI.WinForms.BunifuSnackbar.MessageTypes.Error, 3000)
                FuncLib.WriteLog.WriteErrorLog($"Error occured While Fetching The Occasion Dates")
                BunifuDataGridView5.DataSource = Nothing
                BunifuDropdown16.Items.Clear()
                BunifuDropdown16.Text = Nothing
                BunifuTextBox12.Clear()
                BunifuTextBox13.Clear()
                GroupBox13.Enabled = True
                GroupBox16.Visible = False
                Exit Sub
            Else
                GroupBox13.Enabled = False
                GroupBox14.Visible = True
                GroupBox14.Enabled = True
                BunifuDropdown13.Items.Clear()
                BunifuDropdown13.Text = Nothing
                For Each DateDetails In OccasionDate
                    BunifuDropdown13.Items.Add(DateDetails.ToString("MM/dd/yyyy"))
                Next
            End If
        ElseIf Occasiontype = "Undefined Type" Then
            BunifuSnackbar1.Show(Me, $"No Occasion Type Mentioned For {BunifuDropdown12.SelectedItem.ToString}. Please Contact To Developer!", Bunifu.UI.WinForms.BunifuSnackbar.MessageTypes.Error, 3000)
            BunifuDataGridView5.DataSource = Nothing
            BunifuDropdown16.Items.Clear()
            BunifuDropdown16.Text = Nothing
            BunifuTextBox12.Clear()
            BunifuTextBox13.Clear()
            GroupBox13.Enabled = True
            GroupBox16.Visible = False
            Exit Sub
        Else
            BunifuSnackbar1.Show(Me, $"Error Occured While Executing FuncLib.Gifts.CheckTypeOfOcassion().{Environment.NewLine} Error :{Occasiontype}", Bunifu.UI.WinForms.BunifuSnackbar.MessageTypes.Error, 3000)
            FuncLib.WriteLog.WriteErrorLog($"Error Occured While Executing FuncLib.Gifts.CheckTypeOfOcassion() : {Occasiontype}")
            BunifuDataGridView5.DataSource = Nothing
            BunifuDropdown16.Items.Clear()
            BunifuDropdown16.Text = Nothing
            BunifuTextBox12.Clear()
            BunifuTextBox13.Clear()
            GroupBox13.Enabled = True
            GroupBox16.Visible = False
            Exit Sub
        End If
    End Sub
    Private Sub BunifuTextBox12_TextChanged(sender As Object, e As EventArgs) Handles BunifuTextBox12.TextChanged, BunifuTextBox13.TextChanged
        Dim selectedColumnIndex As Integer = BunifuDropdown16.SelectedIndex
        Dim columnName As String = If(selectedColumnIndex >= 0, BunifuDropdown16.Items(selectedColumnIndex).ToString(), "")
        If BunifuDataGridView5.DataSource IsNot Nothing AndAlso Not String.IsNullOrEmpty(columnName) Then
            Dim bindingSource As New BindingSource()
            bindingSource.DataSource = BunifuDataGridView5.DataSource
            bindingSource.Filter = $"[{columnName}] LIKE '%{BunifuTextBox12.Text}%'"
            BunifuDataGridView5.DataSource = bindingSource
        End If
    End Sub
    Private Sub BunifuDropdown13_SelectedIndexChanged(sender As Object, e As EventArgs) Handles BunifuDropdown13.SelectedIndexChanged
        Dim result = FuncLib.History.FetchAllPendingDataFromDBforVariable(BunifuDropdown12.SelectedItem.ToString, BunifuDropdown13.SelectedItem.ToString)
        If result IsNot Nothing AndAlso result.Item1 IsNot Nothing Then
            Dim dataTable = result.Item1
            If dataTable.Rows.Count > 0 Then
                dataTable.Columns("Employee_ID").ColumnName = "EMP ID"
                dataTable.Columns("NT_ID").ColumnName = "NT ID"
                dataTable.Columns("Employee_Name").ColumnName = "NAME"
                dataTable.Columns("Buisness_Area").ColumnName = "BUISNESS"
                dataTable.Columns("Functions").ColumnName = "DEPARTMENT"
                dataTable.Columns("Created_At").ColumnName = "TIME"
                BunifuDataGridView5.DataSource = dataTable
                GroupBox16.Visible = True
                BunifuDropdown16.Items.Clear()
                BunifuDropdown16.Text = Nothing
                BunifuTextBox12.Clear()
                BunifuTextBox13.Clear()
                BunifuTextBox13.Text = dataTable.Rows.Count
                For Each column As DataGridViewColumn In BunifuDataGridView5.Columns
                    BunifuDropdown16.Items.Add(column.HeaderText)
                Next
            Else
                BunifuSnackbar1.Show(Me, $"No Pending Employees For {BunifuDropdown12.SelectedItem.ToString}", Bunifu.UI.WinForms.BunifuSnackbar.MessageTypes.Warning, 3000)
                BunifuDataGridView5.DataSource = Nothing
                BunifuDropdown16.Items.Clear()
                BunifuDropdown16.Text = Nothing
                BunifuTextBox12.Clear()
                BunifuTextBox13.Clear()
                GroupBox13.Enabled = True
                GroupBox16.Visible = False
                Exit Sub
            End If
        Else
            BunifuSnackbar1.Show(Me, $"Error Occured While Executing FuncLib.Gifts.CheckTypeOfOcassion().{Environment.NewLine} Error :{result}", Bunifu.UI.WinForms.BunifuSnackbar.MessageTypes.Error, 3000)
            FuncLib.WriteLog.WriteErrorLog($"Error Occured While Executing FuncLib.Gifts.CheckTypeOfOcassion() : {result}")
            BunifuDataGridView5.DataSource = Nothing
            BunifuDropdown16.Items.Clear()
            BunifuDropdown16.Text = Nothing
            BunifuTextBox12.Clear()
            BunifuTextBox13.Clear()
            GroupBox13.Enabled = True
            GroupBox16.Visible = False
            Exit Sub
        End If
    End Sub
    Private Sub BunifuDropdown17_SelectedIndexChanged(sender As Object, e As EventArgs) Handles BunifuDropdown17.SelectedIndexChanged
        BunifuDropdown17.Text = "SAVE AS"
        If BunifuDropdown17.SelectedItem = ".XLS" Then
            Dim saveFileDialog As New SaveFileDialog()
            saveFileDialog.Filter = "Excel File (*.csv)|*.csv|All Files (*.*)|*.*"
            saveFileDialog.FilterIndex = 1
            saveFileDialog.RestoreDirectory = True
            If saveFileDialog.ShowDialog() = DialogResult.OK Then
                Dim fileName As String = saveFileDialog.FileName
                Dim SaveFileasExcel As String = FuncLib.History.SaveasExcel(BunifuDataGridView5, $"{fileName}")
                If SaveFileasExcel = "True" Then
                    BunifuSnackbar1.Show(Me, $"History Successfully saved as {fileName}", Bunifu.UI.WinForms.BunifuSnackbar.MessageTypes.Success, 3000)
                    Exit Sub
                Else
                    BunifuSnackbar1.Show(Me, $"Error Saving File as {fileName}. {Environment.NewLine} Error :{SaveFileasExcel}", Bunifu.UI.WinForms.BunifuSnackbar.MessageTypes.Error, 3000)
                    FuncLib.WriteLog.WriteErrorLog($"Error Occured While Executing FuncLib.History.SaveasExcel():{SaveFileasExcel}")
                    Exit Sub
                End If
            End If
        ElseIf BunifuDropdown17.SelectedItem = ".PDF" Then
            Dim saveFileDialog As New SaveFileDialog()
            saveFileDialog.Filter = "pdf File (*.pdf)|*.pdf|All Files (*.*)|*.*"
            saveFileDialog.FilterIndex = 1
            saveFileDialog.RestoreDirectory = True
            If saveFileDialog.ShowDialog() = DialogResult.OK Then
                Dim fileName As String = saveFileDialog.FileName
                Dim SaveFileaspdf As String = FuncLib.History.SaveAsPdf(BunifuDataGridView5, $"{fileName}")
                If SaveFileaspdf = "True" Then
                    BunifuSnackbar1.Show(Me, $"History Successfully saved as {fileName}", Bunifu.UI.WinForms.BunifuSnackbar.MessageTypes.Success, 3000)
                    Exit Sub
                Else
                    BunifuSnackbar1.Show(Me, $"Error Saving File as {fileName}. {Environment.NewLine} Error :{SaveFileaspdf}", Bunifu.UI.WinForms.BunifuSnackbar.MessageTypes.Error, 3000)
                    FuncLib.WriteLog.WriteErrorLog($"Error Occured While Executing FuncLib.History.SaveasExcel():{SaveFileaspdf}")
                    Exit Sub
                End If
            End If
        End If
    End Sub


End Class
