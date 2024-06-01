Imports System.Configuration
Imports MadMilkman.Ini
Imports BunifuAnimatorNS
Imports System.IO
Imports System.Data.OleDb
Imports System.Data.Common
Imports System.Windows.Forms.VisualStyles.VisualStyleElement.Button
Imports System.IO.Ports
Imports System.Text
Imports iTextSharp.text
Imports System.Threading
Imports OfficeOpenXml.ExcelErrorValue
Imports System.ComponentModel
Imports OfficeOpenXml
Public Class SETUP
    Private WithEvents serialPort As New SerialPort()
    Dim ports As String() = SerialPort.GetPortNames()
    Dim Ini As New IniFile
    Dim Inipath As String = ConfigurationManager.AppSettings("IniFile")
    Dim transition As New BunifuTransition
    Dim requiredColumns As New List(Of String)()
    Dim toolTip As New Bunifu.UI.WinForms.BunifuToolTip
    Dim additionalInfo As String = ""
    Dim Issue_Gift As Boolean
    Dim Events_Addition As Boolean
    Dim Admn_Access As Boolean
    Dim currentHostName As String = System.Net.Dns.GetHostName()
    Private selectedRows As New List(Of DataGridViewRow)
    Dim mergedDataTable As DataTable
    Dim totalRowCount As Integer = 0
    Dim duplicateRowCount As Integer = 0
    Dim importedRowCount As Integer = 0
    Dim notInsertedDataTable As DataTable = New DataTable()
    Dim AffectedEmployeeCOunt As Integer = 0
    Dim UnaffectedEmployeeCount As Integer = 0

    Private Sub SETUP_Load(sender As Object, e As EventArgs) Handles MyBase.Load
        BunifuPages1.TabPages.Clear()
        BunifuPages1.Visible = False
        BunifuDataGridView4.Visible = False
        BunifuDataGridView4.DataSource = Nothing
    End Sub
    Private Sub BunifuButton1_Click(sender As Object, e As EventArgs) Handles BunifuButton1.Click
        BunifuTextBox1.Clear()
        BunifuTextBox2.Clear()
        BunifuTextBox7.Clear()
        BunifuDropdown8.Items.Clear()
        BunifuDropdown8.Text = Nothing
        Ini.Load(Inipath)
        BunifuTextBox1.ReadOnly = True
        BunifuTextBox2.ReadOnly = True
        BunifuTextBox1.Text = Ini.Sections("BASIC").Keys("Application_Log").Value
        BunifuTextBox2.Text = Ini.Sections("BASIC").Keys("Error_Log").Value
        BunifuTextBox7.Text = $"{Ini.Sections("BASIC").Keys("Backup_Location").Value}\"
        For Each port As String In ports
            BunifuDropdown8.Items.Add(port)
        Next
        Dim currentHostName As String = System.Net.Dns.GetHostName()
        Dim serialSection As IniSection = Ini.Sections("SERIAL")
        Dim matchFound As Boolean = False
        For Each key As IniKey In serialSection.Keys
            If key.Name.Equals(currentHostName) Then
                BunifuDropdown8.SelectedItem = Ini.Sections("SERIAL").Keys($"{System.Net.Dns.GetHostName()}").Value
                matchFound = True
                Exit For
            End If
        Next
        If Not matchFound Then
            BunifuDropdown8.SelectedItem = Nothing
        End If
        If Ini.Sections("BASIC").Keys("AppLoging_State").Value = "True" Then
            BunifuToggleSwitch1.Checked = True
        Else
            BunifuToggleSwitch1.Checked = False
        End If
        If Ini.Sections("BASIC").Keys("ErrorLoging_State").Value = "True" Then
            BunifuToggleSwitch2.Checked = True
        Else
            BunifuToggleSwitch2.Checked = False
        End If
        ListBox1.Items.Clear()
        ListBox2.Items.Clear()
        If BunifuTextBox1.Text IsNot Nothing Then
            ListBox1.Items.Clear()
            Dim AppLogPath As String = BunifuTextBox1.Text.ToString
            Dim lines As String() = System.IO.File.ReadAllLines(AppLogPath)
            For Each line As String In lines
                ListBox1.Items.Add(line)
            Next
        End If
        If BunifuTextBox2.Text IsNot Nothing Then
            ListBox2.Items.Clear()
            Dim ErrorLogPath As String = BunifuTextBox2.Text.ToString
            Dim lines1 As String() = System.IO.File.ReadAllLines(ErrorLogPath)
            For Each line As String In lines1
                ListBox2.Items.Add(line)
            Next
        End If
        If BunifuPages1.Visible = False Then
            BunifuPages1.TabPages.Clear()
            transition.ShowSync(BunifuPages1, False, BunifuAnimatorNS.Animation.ScaleAndHorizSlide)
            BunifuPages1.TabPages.Add(TabPage1)
        Else
            transition.HideSync(BunifuPages1, False, BunifuAnimatorNS.Animation.ScaleAndHorizSlide)
            BunifuPages1.TabPages.Clear()
            transition.ShowSync(BunifuPages1, False, BunifuAnimatorNS.Animation.ScaleAndHorizSlide)
            BunifuPages1.TabPages.Add(TabPage1)
        End If
    End Sub
    Private Sub BunifuButton8_Click(sender As Object, e As EventArgs) Handles BunifuButton8.Click
        Dim openFileDialog As New OpenFileDialog()
        openFileDialog.Filter = "Txt File|*.txt"
        openFileDialog.Title = "Select Log File"
        openFileDialog.Multiselect = False
        If openFileDialog.ShowDialog() = DialogResult.OK Then
            Dim selectedFilePath As String = openFileDialog.FileName
            If Path.GetExtension(selectedFilePath).ToLower() = ".txt" Then
                If New FileInfo(selectedFilePath).Length = 0 Then
                    Dim newFilePath As String = Path.Combine(Path.GetDirectoryName(selectedFilePath), "Application_Log.txt")
                    File.Move(selectedFilePath, newFilePath)
                    BunifuTextBox1.Text = newFilePath
                    BunifuSnackbar1.Show(Me, $"{selectedFilePath} Is Selected SuccessFully.", Bunifu.UI.WinForms.BunifuSnackbar.MessageTypes.Success, 3000)
                Else
                    BunifuSnackbar1.Show(Me, $"{selectedFilePath} Is Not Empty! Please Select a New Txt File.", Bunifu.UI.WinForms.BunifuSnackbar.MessageTypes.Warning, 3000)
                    Exit Sub
                End If
            Else
                BunifuSnackbar1.Show(Me, $"{selectedFilePath} Is Not A Txt File. Please Select A Txt File Properly", Bunifu.UI.WinForms.BunifuSnackbar.MessageTypes.Warning, 3000)
                Exit Sub
            End If
        End If
    End Sub
    Private Sub BunifuButton9_Click(sender As Object, e As EventArgs) Handles BunifuButton9.Click
        Dim openFileDialog As New OpenFileDialog()
        openFileDialog.Filter = "Txt File|*.txt"
        openFileDialog.Title = "Select Log File"
        openFileDialog.Multiselect = False
        If openFileDialog.ShowDialog() = DialogResult.OK Then
            Dim selectedFilePath As String = openFileDialog.FileName
            If Path.GetExtension(selectedFilePath).ToLower() = ".txt" Then
                If New FileInfo(selectedFilePath).Length = 0 Then
                    Dim newFilePath As String = Path.Combine(Path.GetDirectoryName(selectedFilePath), "Error_Log.txt")
                    File.Move(selectedFilePath, newFilePath)
                    BunifuTextBox2.Text = newFilePath
                    BunifuSnackbar1.Show(Me, $"{selectedFilePath} Is Selected SuccessFully.", Bunifu.UI.WinForms.BunifuSnackbar.MessageTypes.Success, 3000)
                Else
                    BunifuSnackbar1.Show(Me, $"{selectedFilePath} Is Not Empty! Please Select a New Txt File.", Bunifu.UI.WinForms.BunifuSnackbar.MessageTypes.Warning, 3000)
                    Exit Sub
                End If
            Else
                BunifuSnackbar1.Show(Me, $"{selectedFilePath} Is Not A Txt File. Please Select A Txt File Properly", Bunifu.UI.WinForms.BunifuSnackbar.MessageTypes.Warning, 3000)
                Exit Sub
            End If
        End If
    End Sub
    Private Sub BunifuButton10_Click(sender As Object, e As EventArgs) Handles BunifuButton10.Click
        Try
            Dim AppLogPAth As String = BunifuTextBox1.Text
            If File.Exists(AppLogPAth) Then
                If New FileInfo(AppLogPAth).Length > 0 Then
                    File.WriteAllText(AppLogPAth, String.Empty)
                    ListBox1.Items.Clear()
                End If
            End If
            BunifuSnackbar1.Show(Me, $"Application Log Cleared Successfully", Bunifu.UI.WinForms.BunifuSnackbar.MessageTypes.Success, 2000)
            FuncLib.WriteLog.WriteAppLog($"Application_Log Cleared By {MainWindow.ListBox1.Items(1)}_{MainWindow.ListBox1.Items(4)}")
        Catch ex As Exception
            BunifuSnackbar1.Show(Me, $"Error Occured While Clearing Application Log File.{Environment.NewLine}Error:{ex.Message}", Bunifu.UI.WinForms.BunifuSnackbar.MessageTypes.Error, 2000)
            FuncLib.WriteLog.WriteErrorLog($"Error Occured While Clearing Application Log File :{ex.Message}")
            Exit Sub
        End Try

    End Sub
    Private Sub BunifuTextBox1_TextChanged(sender As Object, e As EventArgs) Handles BunifuTextBox1.TextChanged
        BunifuTextBox1.BorderColorIdle = Color.Yellow
        ListBox1.Items.Clear()
        If Not String.IsNullOrWhiteSpace(BunifuTextBox1.Text) Then
            Dim AppLogPAth As String = BunifuTextBox1.Text
            If File.Exists(AppLogPAth) Then
                If New FileInfo(AppLogPAth).Length > 0 Then
                    Dim lines As String() = File.ReadAllLines(AppLogPAth)
                    For Each line As String In lines
                        ListBox1.Items.Add(line)
                    Next
                End If
            End If
        End If
    End Sub
    Private Sub BunifuTextBox2_TextChanged(sender As Object, e As EventArgs) Handles BunifuTextBox2.TextChanged
        BunifuTextBox2.BorderColorIdle = Color.Yellow
        ListBox2.Items.Clear()
        If Not String.IsNullOrWhiteSpace(BunifuTextBox2.Text) Then
            Dim ErrorLogPath As String = BunifuTextBox2.Text
            If File.Exists(ErrorLogPath) Then
                If New FileInfo(ErrorLogPath).Length > 0 Then
                    Dim lines As String() = File.ReadAllLines(ErrorLogPath)
                    For Each line As String In lines
                        ListBox2.Items.Add(line)
                    Next
                End If
            End If
        End If
    End Sub
    Private Sub BunifuButton27_Click(sender As Object, e As EventArgs) Handles BunifuButton27.Click
        Try
            Dim ErrorLogPath As String = BunifuTextBox2.Text
            If File.Exists(ErrorLogPath) Then
                If New FileInfo(ErrorLogPath).Length > 0 Then
                    File.WriteAllText(ErrorLogPath, String.Empty)
                    ListBox2.Items.Clear()
                End If
            End If
            BunifuSnackbar1.Show(Me, $"Error Log Cleared Successfully", Bunifu.UI.WinForms.BunifuSnackbar.MessageTypes.Success, 2000)
            FuncLib.WriteLog.WriteAppLog($"Error Log Cleared By {MainWindow.ListBox1.Items(1)}_{MainWindow.ListBox1.Items(4)}")
        Catch ex As Exception
            BunifuSnackbar1.Show(Me, $"Error Occured While Clearing Error Log File.{Environment.NewLine}Error:{ex.Message}", Bunifu.UI.WinForms.BunifuSnackbar.MessageTypes.Error, 2000)
            FuncLib.WriteLog.WriteErrorLog($"Error Occured While Clearing Error Log File :{ex.Message}")
            Exit Sub
        End Try
    End Sub
    Private Sub BunifuButton25_Click(sender As Object, e As EventArgs) Handles BunifuButton25.Click
        Dim Ini1 As New IniFile
        Dim CheckedState As Boolean = False
        Try
            Ini1.Load(Inipath)
            Ini1.Sections("BASIC").Keys("Application_Log").Value = BunifuTextBox1.Text.ToString
            If BunifuToggleSwitch1.Checked = True Then
                Ini1.Sections("BASIC").Keys("AppLoging_State").Value = "True"
                CheckedState = True
            Else
                Ini1.Sections("BASIC").Keys("AppLoging_State").Value = "False"
                CheckedState = False
            End If
            Ini1.Save(Inipath)
            BunifuTextBox1.BorderColorIdle = Color.Green
            BunifuSnackbar1.Show(Me, $"COnfriguration Updated Successfully!", Bunifu.UI.WinForms.BunifuSnackbar.MessageTypes.Success, 2000)
            FuncLib.WriteLog.WriteAppLog($"Application Log Confriguration Changed By {MainWindow.ListBox1.Items(1)}_{MainWindow.ListBox1.Items(4)} Changes :{BunifuTextBox1.Text}{CheckedState}")
            RebbotInitiaizer.TextBox1.Text = 5
            RebbotInitiaizer.BackColor = Color.LimeGreen
            RebbotInitiaizer.ShowDialog()
        Catch ex As Exception
            BunifuSnackbar1.Show(Me, $"Error Occured While Saving the Confriguration File.{Environment.NewLine}Error:{ex.Message}", Bunifu.UI.WinForms.BunifuSnackbar.MessageTypes.Error, 2000)
            FuncLib.WriteLog.WriteErrorLog($"Error Occured While Updaing Confriguration :{ex.Message}")
            Exit Sub
        End Try
    End Sub
    Private Sub BunifuButton26_Click(sender As Object, e As EventArgs) Handles BunifuButton26.Click
        Dim Ini1 As New IniFile
        Dim CheckedState As Boolean = False
        Try
            Ini1.Load(Inipath)
            Ini1.Sections("BASIC").Keys("Error_Log").Value = BunifuTextBox2.Text.ToString
            If BunifuToggleSwitch2.Checked = True Then
                Ini1.Sections("BASIC").Keys("ErrorLoging_State").Value = "True"
                CheckedState = True
            Else
                Ini1.Sections("BASIC").Keys("ErrorLoging_State").Value = "False"
                CheckedState = False
            End If
            Ini1.Save(Inipath)
            BunifuTextBox2.BorderColorIdle = Color.Green
            BunifuSnackbar1.Show(Me, $"Confriguration Updated Successfully!", Bunifu.UI.WinForms.BunifuSnackbar.MessageTypes.Success, 2000)
            FuncLib.WriteLog.WriteAppLog($"Error Log Confriguration Changed By {MainWindow.ListBox1.Items(1)}_{MainWindow.ListBox1.Items(4)} Changes :{BunifuTextBox2.Text}{CheckedState}")
            RebbotInitiaizer.TextBox1.Text = 5
            RebbotInitiaizer.BackColor = Color.LimeGreen
            RebbotInitiaizer.ShowDialog()
        Catch ex As Exception
            BunifuSnackbar1.Show(Me, $"Error Occured While Saving the Confriguration File.{Environment.NewLine}Error:{ex.Message}", Bunifu.UI.WinForms.BunifuSnackbar.MessageTypes.Error, 2000)
            FuncLib.WriteLog.WriteErrorLog($"Error Occured While Updaing Confriguration :{ex.Message}")
            Exit Sub
        End Try

    End Sub
    Private Sub BunifuButton2_Click(sender As Object, e As EventArgs) Handles BunifuButton2.Click
        BunifuDropdown1.Items.Clear()
        BunifuDropdown1.Text = Nothing
        BunifuDropdown1.Items.Add("Microsoft.ACE.OLEDB.12.0")
        BunifuDropdown1.Items.Add("Microsoft.ACE.OLEDB.4.0")
        BunifuTextBox3.Clear()
        BunifuTextBox4.Clear()
        BunifuShadowPanel6.Enabled = False
        BunifuTextBox4.ReadOnly = True
        Ini.Load(Inipath)
        If Ini.Sections("DATABASE").Keys("Provider").Value = "Microsoft.ACE.OLEDB.12.0" Then
            BunifuDropdown1.SelectedItem = BunifuDropdown1.Items(0)
        ElseIf Ini.Sections("DATABASE").Keys("Provider").Value = "Microsoft.ACE.OLEDB.4.0" Then
            BunifuDropdown1.SelectedItem = BunifuDropdown1.Items(1)
        End If
        BunifuTextBox3.Text = Ini.Sections("DATABASE").Keys("Source").Value
        BunifuTextBox4.Text = Ini.Sections("DATABASE").Keys("Password").Value
        If BunifuPages1.Visible = False Then
            BunifuPages1.TabPages.Clear()
            transition.ShowSync(BunifuPages1, False, BunifuAnimatorNS.Animation.ScaleAndHorizSlide)
            BunifuPages1.TabPages.Add(TabPage2)
        Else
            transition.HideSync(BunifuPages1, False, BunifuAnimatorNS.Animation.ScaleAndHorizSlide)
            BunifuPages1.TabPages.Clear()
            transition.ShowSync(BunifuPages1, False, BunifuAnimatorNS.Animation.ScaleAndHorizSlide)
            BunifuPages1.TabPages.Add(TabPage2)
        End If
    End Sub
    Private Sub BunifuButton13_Click(sender As Object, e As EventArgs) Handles BunifuButton13.Click 'New DB Selection
        Dim saveFileDialog As New SaveFileDialog()
        saveFileDialog.Filter = "ACCESS DATABASE (*.accdb)|*.accdb|All Files (*.*)|*.*"
        saveFileDialog.FilterIndex = 1
        saveFileDialog.RestoreDirectory = True
        If saveFileDialog.ShowDialog() = DialogResult.OK Then
            Dim fileName As String = saveFileDialog.FileName
            SetDbPassword.BunifuTextBox1.Text = BunifuDropdown1.SelectedItem
            SetDbPassword.BunifuTextBox2.Text = fileName
            transition.ShowSync(SetDbPassword, False, BunifuAnimatorNS.Animation.ScaleAndHorizSlide)
        End If
    End Sub
    Private Sub BunifuTextBox4_TextChanged(sender As Object, e As EventArgs) Handles BunifuTextBox4.TextChanged
        BunifuTextBox4.PasswordChar = "*"
    End Sub
    Private Sub BunifuButton17_Click(sender As Object, e As EventArgs) Handles BunifuButton17.Click
        Dim SaveAndValidate As String = FuncLib.Setup.ValidateAndSaveDtaabseIntoIni(BunifuDropdown1.SelectedItem, BunifuTextBox3.Text, BunifuTextBox4.Text)
        If SaveAndValidate = "True" Then
            BunifuSnackbar1.Show(Me, $"New DataBase Has Been Succesfully Uploaded And Linked To Software. Please Restart The Software To Take Affect", Bunifu.UI.WinForms.BunifuSnackbar.MessageTypes.Success, 3000)
            RebbotInitiaizer.TextBox1.Text = 5
            RebbotInitiaizer.BackColor = Color.LimeGreen
            RebbotInitiaizer.ShowDialog()
        ElseIf SaveAndValidate = "False" Then
            BunifuSnackbar1.Show(Me, $"Failed To Verify The New Database", Bunifu.UI.WinForms.BunifuSnackbar.MessageTypes.Warning, 3000)
            Exit Sub
        Else
            BunifuSnackbar1.Show(Me, $"Error Occured While Validating New Database.{Environment.NewLine} Error :{SaveAndValidate}", Bunifu.UI.WinForms.BunifuSnackbar.MessageTypes.Error, 3000)
            FuncLib.WriteLog.WriteErrorLog($"Error Occured While Executing FuncLib.Setup.ValidateAndSaveDtaabseIntoIni() :{SaveAndValidate} ")
            Exit Sub
        End If
    End Sub
    Private Sub ClearUserAddition()
        BunifuShadowPanel6.Enabled = True
        BunifuShadowPanel6.Visible = True
        BunifuDropdown2.Items.Clear()
        BunifuDropdown2.Text = Nothing
        BunifuDropdown2.Items.Add("UPDATE MASTER DATABASE FILE")
        BunifuDropdown2.Items.Add("ADD NEW EMPLOYEE")
        BunifuDropdown2.Items.Add("REMOVE FROM MASTER DATABASE FILE")
        BunifuDropdown2.Items.Add("REPLACE MASTER DATABASE FILE")
        BunifuDropdown2.Items.Add("UPDATE/MODIFY EMPLOYEES")
        BunifuDropdown2.Enabled = True
        BunifuDropdown2.Visible = True
        GroupBox2.Enabled = False
        GroupBox2.Visible = False
        BunifuTextBox6.Clear()
        GroupBox1.Enabled = False
        GroupBox1.Visible = False
        BunifuDropdown5.Items.Clear()
        BunifuDropdown5.Text = Nothing
        BunifuTextBox21.Clear()
        BunifuDataGridView1.Enabled = False
        BunifuDataGridView1.Visible = False
        BunifuDataGridView1.DataSource = Nothing
        BunifuButton20.Enabled = False
        BunifuButton20.Visible = False
        BunifuButton29.Enabled = False
        BunifuButton29.Visible = False
        CheckBox1.Checked = False
        CheckBox1.Visible = False
        ProgressBar1.Minimum = 0
        ProgressBar1.Value = 0
        ProgressBar1.Visible = False
        BunifuButton11.Visible = False
        BunifuButton34.Visible = False
        CheckBox2.Checked = False
        CheckBox2.Enabled = False
        CheckBox2.Visible = False
    End Sub
    Private Sub BunifuButton3_Click(sender As Object, e As EventArgs) Handles BunifuButton3.Click
        ClearUserAddition()
        If BunifuPages1.Visible = False Then
            BunifuPages1.TabPages.Clear()
            transition.ShowSync(BunifuPages1, False, BunifuAnimatorNS.Animation.ScaleAndHorizSlide)
            BunifuPages1.TabPages.Add(TabPage3)
        Else
            transition.HideSync(BunifuPages1, False, BunifuAnimatorNS.Animation.ScaleAndHorizSlide)
            BunifuPages1.TabPages.Clear()
            transition.ShowSync(BunifuPages1, False, BunifuAnimatorNS.Animation.ScaleAndHorizSlide)
            BunifuPages1.TabPages.Add(TabPage3)
        End If
    End Sub
    Private Sub BunifuDropdown2_SelectedIndexChanged(sender As Object, e As EventArgs) Handles BunifuDropdown2.SelectedIndexChanged
        selectedRows.Clear()
        GroupBox2.Enabled = False
        GroupBox2.Visible = False
        BunifuTextBox6.Clear()
        GroupBox1.Enabled = False
        GroupBox1.Visible = False
        BunifuDropdown5.Items.Clear()
        BunifuDropdown5.Text = Nothing
        BunifuTextBox21.Clear()
        BunifuDataGridView1.Enabled = False
        BunifuDataGridView1.Visible = False
        BunifuDataGridView1.DataSource = Nothing
        BunifuButton20.Enabled = False
        BunifuButton20.Visible = False
        CheckBox1.Checked = False
        CheckBox1.Visible = False
        ProgressBar1.Minimum = 0
        ProgressBar1.Value = 0
        ProgressBar1.Visible = False
        If BunifuDropdown2.SelectedItem = "ADD NEW EMPLOYEE" Then
            CheckBox2.Checked = False
            CheckBox2.Visible = False
            CheckBox2.Enabled = False
            GroupBox2.Visible = False
            BunifuTextBox6.Clear()
            BunifuTextBox6.Enabled = False
            BunifuTextBox6.Visible = False
            BunifuButton19.Enabled = False
            BunifuButton19.Visible = False
            GroupBox1.Enabled = False
            GroupBox1.Visible = False
            BunifuDropdown5.Items.Clear()
            BunifuDropdown5.Text = Nothing
            BunifuDropdown5.Enabled = False
            BunifuDropdown5.Visible = False
            BunifuTextBox21.Clear()
            BunifuTextBox21.Enabled = False
            BunifuTextBox21.Visible = False
            BunifuButton11.Visible = False
            BunifuDataGridView1.DataSource = Nothing
            BunifuDataGridView1.Enabled = False
            BunifuDataGridView1.Visible = False
            BunifuButton20.Enabled = False
            BunifuButton20.Visible = False
            BunifuButton29.Enabled = False
            BunifuButton29.Visible = False
            ADD_NEW_EMPLOYEE.BackColor = Color.GhostWhite
            ADD_NEW_EMPLOYEE.TextBox1.Text = "ADD NEW EMPLOYEE"
            transition.ShowSync(ADD_NEW_EMPLOYEE, True, BunifuAnimatorNS.Animation.ScaleAndHorizSlide)
            ADD_NEW_EMPLOYEE.TopMost = True
            Me.Enabled = False
        ElseIf BunifuDropdown2.SelectedItem = "REMOVE FROM MASTER DATABASE FILE" Then
            CheckBox2.Checked = False
            CheckBox2.Visible = False
            CheckBox2.Enabled = False
            selectedRows.Clear()
            GroupBox2.Visible = False
            BunifuTextBox6.Clear()
            BunifuTextBox6.Visible = False
            BunifuButton19.Visible = False
            GroupBox1.Visible = False
            BunifuDropdown5.Items.Clear()
            BunifuDropdown5.Text = Nothing
            BunifuDropdown5.Visible = False
            BunifuTextBox21.Clear()
            BunifuTextBox21.Visible = False
            BunifuButton11.Visible = False
            BunifuDataGridView1.DataSource = Nothing
            BunifuDataGridView1.Visible = False
            BunifuButton20.Visible = False
            BunifuButton29.Visible = True
            ProgressBar1.Minimum = 0
            ProgressBar1.Value = 0
            Dim result = FuncLib.Setup.ExportDataFromEmployeeDB()
            If result.Item1 IsNot Nothing Then
                Dim grid = result.Item1
                If grid.Rows.Count > 0 Then
                    ProgressBar1.Visible = True
                    CheckBox1.Checked = False
                    CheckBox1.Visible = True
                    ProgressBar1.Minimum = 0
                    ProgressBar1.Value = 0
                    ProgressBar1.Maximum = grid.Rows.Count
                    Dim CurrentRowCount As Integer = 0
                    Dim checkBoxColumn As New DataGridViewCheckBoxColumn()
                    checkBoxColumn.HeaderText = "Select"
                    checkBoxColumn.Width = 5
                    checkBoxColumn.Name = "Select"
                    If BunifuDataGridView1.Columns.Contains("Select") Then
                        BunifuDataGridView1.Columns.Remove("Select")
                    End If
                    BunifuDataGridView1.Columns.Insert(0, checkBoxColumn)
                    grid.Columns("ID").ColumnName = "SR NO."
                    grid.Columns("Employee_ID").ColumnName = "EMP ID"
                    grid.Columns("NT_ID").ColumnName = "NT ID"
                    grid.Columns("Scanner_ID").ColumnName = "SCAN ID"
                    grid.Columns("Scanner_5D").ColumnName = "SCAN 5D"
                    grid.Columns("Employee_Name").ColumnName = "NAME"
                    grid.Columns("Gender").ColumnName = "GENDER"
                    grid.Columns("Celebration_Date").ColumnName = "DOB"
                    grid.Columns("Buisness_Area").ColumnName = "BUISNESS"
                    grid.Columns("Functions").ColumnName = "DEPARTMENT"
                    grid.Columns("Leave_Approver").ColumnName = "REPORTING MANAGER"
                    grid.Columns("Created_By").ColumnName = "CREATED BY"
                    grid.Columns("Created_At").ColumnName = "CREATED AT"
                    For Each row As DataRow In grid.Rows
                        CurrentRowCount += 1
                        ProgressBar1.Value = CurrentRowCount
                    Next
                    BunifuDataGridView1.DataSource = grid
                    BunifuDataGridView1.Visible = True
                    BunifuDataGridView1.Enabled = True
                    BunifuButton29.Visible = True
                    BunifuButton29.Enabled = True
                    GroupBox1.Visible = True
                    GroupBox1.Enabled = True
                    BunifuDropdown5.Visible = True
                    BunifuDropdown5.Enabled = True
                    BunifuTextBox21.Visible = True
                    BunifuTextBox21.Enabled = True
                    For Each column As DataGridViewColumn In BunifuDataGridView1.Columns
                        BunifuDropdown5.Items.Add(column.HeaderText)
                    Next
                Else
                    BunifuSnackbar1.Show(Me, $"No Employee Found To Show!", Bunifu.UI.WinForms.BunifuSnackbar.MessageTypes.Information, 3000)
                    Exit Sub
                End If
            Else
                BunifuSnackbar1.Show(Me, $"Error Occured While Reading Data From Employee DB.{Environment.NewLine} Error :{result.Item2}", Bunifu.UI.WinForms.BunifuSnackbar.MessageTypes.Error, 3000)
                FuncLib.WriteLog.WriteErrorLog($"Error Occured While Executing FuncLib.Setup.ExportDataFromEmployeeDB :{result.Item2} ")
                Exit Sub
            End If
        ElseIf BunifuDropdown2.SelectedItem = "UPDATE MASTER DATABASE FILE" Then
            CheckBox2.Checked = False
            CheckBox2.Visible = False
            CheckBox2.Enabled = False
            GroupBox2.Visible = True
            GroupBox2.Enabled = True
            BunifuTextBox6.Clear()
            BunifuTextBox6.Enabled = True
            BunifuTextBox6.Visible = True
            BunifuButton19.Enabled = True
            BunifuButton19.Visible = True
            GroupBox1.Visible = False
            BunifuDropdown5.Items.Clear()
            BunifuDropdown5.Text = Nothing
            BunifuDropdown5.Enabled = False
            BunifuDropdown5.Visible = False
            BunifuTextBox21.Clear()
            BunifuTextBox21.Enabled = False
            BunifuTextBox21.Visible = False
            CheckBox1.Checked = False
            CheckBox1.Visible = False
            BunifuButton11.Visible = False
            ProgressBar1.Minimum = 0
            ProgressBar1.Value = 0
            ProgressBar1.Visible = False
            BunifuDataGridView1.DataSource = Nothing
            BunifuDataGridView1.Visible = False
            BunifuButton20.Visible = False
            BunifuButton29.Enabled = True
            BunifuButton29.Visible = True
        ElseIf BunifuDropdown2.SelectedItem = "REPLACE MASTER DATABASE FILE" Then
            CheckBox2.Checked = False
            CheckBox2.Visible = False
            CheckBox2.Enabled = False
            GroupBox2.Visible = True
            GroupBox2.Enabled = True
            BunifuTextBox6.Clear()
            BunifuTextBox6.Enabled = True
            BunifuTextBox6.Visible = True
            BunifuButton19.Enabled = True
            BunifuButton19.Visible = True
            GroupBox1.Visible = False
            BunifuDropdown5.Items.Clear()
            BunifuDropdown5.Text = Nothing
            BunifuDropdown5.Enabled = False
            BunifuDropdown5.Visible = False
            BunifuTextBox21.Clear()
            BunifuTextBox21.Enabled = False
            BunifuTextBox21.Visible = False
            CheckBox1.Checked = False
            CheckBox1.Visible = False
            BunifuButton11.Visible = False
            ProgressBar1.Minimum = 0
            ProgressBar1.Value = 0
            ProgressBar1.Visible = False
            BunifuDataGridView1.DataSource = Nothing
            BunifuDataGridView1.Visible = False
            BunifuButton20.Visible = False
            BunifuButton29.Enabled = True
            BunifuButton29.Visible = True
        ElseIf BunifuDropdown2.SelectedItem = "UPDATE/MODIFY EMPLOYEES" Then
            CheckBox2.Checked = False
            CheckBox2.Visible = True
            CheckBox2.Enabled = False
            GroupBox2.Visible = False
            GroupBox2.Enabled = False
            BunifuTextBox6.Clear()
            BunifuTextBox6.Visible = False
            BunifuButton19.Visible = False
            GroupBox1.Visible = False
            GroupBox1.Enabled = False
            BunifuDropdown5.Items.Clear()
            BunifuDropdown5.Text = Nothing
            BunifuDropdown5.Visible = False
            BunifuDropdown5.Enabled = False
            BunifuTextBox21.Clear()
            BunifuTextBox21.Visible = False
            CheckBox1.Checked = False
            CheckBox1.Visible = False
            ProgressBar1.Minimum = 0
            ProgressBar1.Value = 0
            ProgressBar1.Visible = False
            BunifuDataGridView1.DataSource = Nothing
            BunifuDataGridView1.Visible = False
            BunifuButton20.Visible = False
            BunifuButton11.Visible = False
            Dim result = FuncLib.Setup.ExportDataFromEmployeeDB()
            If result.Item1 IsNot Nothing Then
                Dim grid = result.Item1
                If grid.Rows.Count > 0 Then
                    ProgressBar1.Visible = True
                    ProgressBar1.Minimum = 0
                    ProgressBar1.Value = 0
                    ProgressBar1.Maximum = grid.Rows.Count
                    Dim CurrentRowCount As Integer = 0
                    Dim checkBoxColumn As New DataGridViewCheckBoxColumn()
                    checkBoxColumn.HeaderText = "Select"
                    checkBoxColumn.Width = 5
                    checkBoxColumn.Name = "Select"
                    If BunifuDataGridView1.Columns.Contains("Select") Then
                        BunifuDataGridView1.Columns.Remove("Select")
                    End If
                    BunifuDataGridView1.Columns.Insert(0, checkBoxColumn)
                    grid.Columns("ID").ColumnName = "SR NO."
                    grid.Columns("Employee_ID").ColumnName = "EMP ID"
                    grid.Columns("NT_ID").ColumnName = "NT ID"
                    grid.Columns("Scanner_ID").ColumnName = "SCAN ID"
                    grid.Columns("Scanner_5D").ColumnName = "SCAN 5D"
                    grid.Columns("Employee_Name").ColumnName = "NAME"
                    grid.Columns("Gender").ColumnName = "GENDER"
                    grid.Columns("Celebration_Date").ColumnName = "DOB"
                    grid.Columns("Buisness_Area").ColumnName = "BUISNESS"
                    grid.Columns("Functions").ColumnName = "DEPARTMENT"
                    grid.Columns("Leave_Approver").ColumnName = "REPORTING MANAGER"
                    grid.Columns("Created_By").ColumnName = "CREATED BY"
                    grid.Columns("Created_At").ColumnName = "CREATED AT"
                    For Each row As DataRow In grid.Rows
                        CurrentRowCount += 1
                        ProgressBar1.Value = CurrentRowCount
                    Next
                    BunifuDataGridView1.DataSource = grid
                    BunifuDataGridView1.Visible = True
                    BunifuDataGridView1.Enabled = True
                    BunifuButton29.Visible = True
                    BunifuButton29.Enabled = True
                    GroupBox1.Visible = True
                    GroupBox1.Enabled = True
                    BunifuDropdown5.Visible = True
                    BunifuDropdown5.Enabled = True
                    BunifuTextBox21.Visible = True
                    BunifuTextBox21.Enabled = True
                    BunifuDropdown5.Items.Clear()
                    BunifuDropdown5.Text = Nothing
                    For Each column As DataGridViewColumn In BunifuDataGridView1.Columns
                        BunifuDropdown5.Items.Add(column.HeaderText)
                    Next
                Else
                    BunifuSnackbar1.Show(Me, $"No Employee Found To Show!", Bunifu.UI.WinForms.BunifuSnackbar.MessageTypes.Information, 3000)
                    Exit Sub
                End If
            Else
                BunifuSnackbar1.Show(Me, $"Error Occured While Reading Data From Employee DB.{Environment.NewLine} Error :{result.Item2}", Bunifu.UI.WinForms.BunifuSnackbar.MessageTypes.Error, 3000)
                FuncLib.WriteLog.WriteErrorLog($"Error Occured While Executing FuncLib.Setup.ExportDataFromEmployeeDB :{result.Item2} ")
                Exit Sub
            End If
        End If
    End Sub
    Private Sub CheckBox1_CheckedChanged(sender As Object, e As EventArgs) Handles CheckBox1.CheckedChanged, CheckBox2.CheckedChanged
        If BunifuDropdown2.SelectedItem = "REMOVE FROM MASTER DATABASE FILE" Then
            selectedRows.Clear()
            Dim isChecked As Boolean = CheckBox1.Checked
            ProgressBar1.Minimum = 0
            ProgressBar1.Maximum = BunifuDataGridView1.RowCount - 1
            ProgressBar1.Value = 0
            For i As Integer = 0 To BunifuDataGridView1.RowCount - 2
                Dim checkboxCell As DataGridViewCheckBoxCell = TryCast(BunifuDataGridView1.Rows(i).Cells("Select"), DataGridViewCheckBoxCell)
                If checkboxCell IsNot Nothing Then
                    checkboxCell.Value = isChecked
                End If
                Dim rows As DataGridViewRow = BunifuDataGridView1.Rows(i)
                If CheckBox1.Checked = True Then
                    selectedRows.Add(rows)
                Else
                    selectedRows.Remove(rows)
                End If
                ProgressBar1.Value += 1
            Next
            If selectedRows.Count > 0 Then
                BunifuButton20.Text = "REMOVE"
                BunifuButton20.Enabled = True
                BunifuButton20.Visible = True
            Else
                BunifuButton20.Enabled = False
                BunifuButton20.Visible = False
            End If
        End If
    End Sub
    Private Sub BunifuDataGridView1_CellContentClick(sender As Object, e As DataGridViewCellEventArgs) Handles BunifuDataGridView1.CellContentClick
        If BunifuDropdown2.SelectedItem = "REMOVE FROM MASTER DATABASE FILE" Then
            If e.ColumnIndex = BunifuDataGridView1.Columns("Select").Index AndAlso e.RowIndex >= 0 Then
                Dim checkboxCell As DataGridViewCheckBoxCell = TryCast(BunifuDataGridView1.Rows(e.RowIndex).Cells("Select"), DataGridViewCheckBoxCell)
                If checkboxCell IsNot Nothing Then
                    Dim row As DataGridViewRow = BunifuDataGridView1.Rows(e.RowIndex)
                    Dim isChecked As Boolean = Convert.ToBoolean(checkboxCell.Value)
                    If BunifuDataGridView1.Rows(e.RowIndex).Cells("Select").Value = isChecked Then
                        If row.Cells("EMP ID").Value Is Nothing OrElse IsDBNull(row.Cells("EMP ID").Value) Then
                            BunifuSnackbar1.Show(Me, $"Invalid Employee Selection!", Bunifu.UI.WinForms.BunifuSnackbar.MessageTypes.Information, 3000)
                            Dim result = FuncLib.Setup.ExportDataFromEmployeeDB()
                            If result.Item1 IsNot Nothing Then
                                Dim grid = result.Item1
                                If grid.Rows.Count > 0 Then
                                    Dim checkBoxColumn As New DataGridViewCheckBoxColumn()
                                    checkBoxColumn.HeaderText = "Select"
                                    checkBoxColumn.Width = 5
                                    checkBoxColumn.Name = "Select"
                                    If BunifuDataGridView1.Columns.Contains("Select") Then
                                        BunifuDataGridView1.Columns.Remove("Select")
                                    End If
                                    BunifuDataGridView1.Columns.Insert(0, checkBoxColumn)
                                    grid.Columns("ID").ColumnName = "SR NO."
                                    grid.Columns("Employee_ID").ColumnName = "EMP ID"
                                    grid.Columns("NT_ID").ColumnName = "NT ID"
                                    grid.Columns("Scanner_ID").ColumnName = "SCAN ID"
                                    grid.Columns("Scanner_5D").ColumnName = "SCAN 5D"
                                    grid.Columns("Employee_Name").ColumnName = "NAME"
                                    grid.Columns("Gender").ColumnName = "GENDER"
                                    grid.Columns("Celebration_Date").ColumnName = "DOB"
                                    grid.Columns("Buisness_Area").ColumnName = "BUISNESS"
                                    grid.Columns("Functions").ColumnName = "DEPARTMENT"
                                    grid.Columns("Leave_Approver").ColumnName = "REPORTING MANAGER"
                                    grid.Columns("Created_By").ColumnName = "CREATED BY"
                                    grid.Columns("Created_At").ColumnName = "CREATED AT"
                                    BunifuDataGridView1.DataSource = grid
                                    Exit Sub
                                End If
                            End If
                        Else
                            If selectedRows.Contains(row) Then
                                selectedRows.Remove(row)
                            Else
                                selectedRows.Add(row)
                            End If
                        End If
                    End If
                End If
            End If
            If selectedRows.Count > 0 Then
                BunifuButton20.Text = "REMOVE"
                BunifuButton20.Enabled = True
                BunifuButton20.Visible = True
            Else
                BunifuButton20.Enabled = False
                BunifuButton20.Visible = False
            End If
        ElseIf BunifuDropdown2.SelectedItem = "UPDATE/MODIFY EMPLOYEES" Then
            If e.ColumnIndex = BunifuDataGridView1.Columns("Select").Index AndAlso e.RowIndex >= 0 Then
                    Dim checkboxCell As DataGridViewCheckBoxCell = TryCast(BunifuDataGridView1.Rows(e.RowIndex).Cells("Select"), DataGridViewCheckBoxCell)
                    If checkboxCell IsNot Nothing Then
                        Dim row As DataGridViewRow = BunifuDataGridView1.Rows(e.RowIndex)
                        Dim isChecked As Boolean = Convert.ToBoolean(checkboxCell.Value)
                    If BunifuDataGridView1.Rows(e.RowIndex).Cells("Select").Value = isChecked Then
                        If CheckBox2.Checked = False Then
                            ADD_NEW_EMPLOYEE.BackColor = Color.GhostWhite
                            ADD_NEW_EMPLOYEE.TextBox1.Text = "UPDATE/MODIFY EMPLOYEES"
                            CheckBox2.Checked = True
                            ADD_NEW_EMPLOYEE.BunifuTextBox1.Text = row.Cells("SCAN ID").Value
                            ADD_NEW_EMPLOYEE.BunifuTextBox2.Text = row.Cells("SCAN 5D").Value
                            ADD_NEW_EMPLOYEE.BunifuTextBox3.Text = row.Cells("EMP ID").Value
                            If row.Cells("NT ID").Value Is Nothing OrElse IsDBNull(row.Cells("NT ID")) Then
                                ADD_NEW_EMPLOYEE.BunifuTextBox4.Text = Nothing
                            Else
                                ADD_NEW_EMPLOYEE.BunifuTextBox4.Text = row.Cells("NT ID").Value
                            End If
                            ADD_NEW_EMPLOYEE.BunifuTextBox5.Text = row.Cells("NAME").Value
                                ADD_NEW_EMPLOYEE.BunifuDropdown1.Items.Clear()
                                ADD_NEW_EMPLOYEE.BunifuDropdown1.Text = Nothing
                                ADD_NEW_EMPLOYEE.BunifuDropdown1.Items.Add("M")
                                ADD_NEW_EMPLOYEE.BunifuDropdown1.Items.Add("F")
                                If row.Cells("GENDER").Value = "M" Then
                                    ADD_NEW_EMPLOYEE.BunifuDropdown1.SelectedItem = ADD_NEW_EMPLOYEE.BunifuDropdown1.Items(0)
                                ElseIf row.Cells("GENDER").Value = "F" Then
                                    ADD_NEW_EMPLOYEE.BunifuDropdown1.SelectedItem = ADD_NEW_EMPLOYEE.BunifuDropdown1.Items(1)
                                End If
                                ADD_NEW_EMPLOYEE.BunifuDatePicker1.Value = row.Cells("DOB").Value
                                ADD_NEW_EMPLOYEE.BunifuTextBox8.Text = row.Cells("BUISNESS").Value
                                ADD_NEW_EMPLOYEE.BunifuTextBox9.Text = row.Cells("DEPARTMENT").Value
                                ADD_NEW_EMPLOYEE.BunifuTextBox10.Text = row.Cells("REPORTING MANAGER").Value
                                transition.ShowSync(ADD_NEW_EMPLOYEE, True, BunifuAnimatorNS.Animation.ScaleAndHorizSlide)
                                ADD_NEW_EMPLOYEE.TopMost = True
                                Me.Enabled = False
                            ElseIf CheckBox2.Checked = True Then
                                BunifuSnackbar1.Show(Me, $"Another Employee Updation Is Already Under Progress", Bunifu.UI.WinForms.BunifuSnackbar.MessageTypes.Information, 3000)
                                CheckBox2.Checked = False
                                Dim result = FuncLib.Setup.ExportDataFromEmployeeDB()
                                If result.Item1 IsNot Nothing Then
                                    Dim grid = result.Item1
                                    If grid.Rows.Count > 0 Then
                                        Dim checkBoxColumn As New DataGridViewCheckBoxColumn()
                                        checkBoxColumn.HeaderText = "Select"
                                        checkBoxColumn.Width = 5
                                        checkBoxColumn.Name = "Select"
                                        If BunifuDataGridView1.Columns.Contains("Select") Then
                                            BunifuDataGridView1.Columns.Remove("Select")
                                        End If
                                        BunifuDataGridView1.Columns.Insert(0, checkBoxColumn)
                                        grid.Columns("ID").ColumnName = "SR NO."
                                        grid.Columns("Employee_ID").ColumnName = "EMP ID"
                                        grid.Columns("NT_ID").ColumnName = "NT ID"
                                        grid.Columns("Scanner_ID").ColumnName = "SCAN ID"
                                        grid.Columns("Scanner_5D").ColumnName = "SCAN 5D"
                                        grid.Columns("Employee_Name").ColumnName = "NAME"
                                        grid.Columns("Gender").ColumnName = "GENDER"
                                        grid.Columns("Celebration_Date").ColumnName = "DOB"
                                        grid.Columns("Buisness_Area").ColumnName = "BUISNESS"
                                        grid.Columns("Functions").ColumnName = "DEPARTMENT"
                                        grid.Columns("Leave_Approver").ColumnName = "REPORTING MANAGER"
                                        grid.Columns("Created_By").ColumnName = "CREATED BY"
                                        grid.Columns("Created_At").ColumnName = "CREATED AT"
                                        BunifuDataGridView1.DataSource = grid
                                        Exit Sub
                                    End If
                                End If
                            ElseIf row.Cells("EMP ID").Value Is Nothing OrElse IsDBNull(row.Cells("EMP ID").Value) Then
                                BunifuSnackbar1.Show(Me, $"Invalid Employee Selection", Bunifu.UI.WinForms.BunifuSnackbar.MessageTypes.Information, 3000)
                            CheckBox2.Checked = False
                            Dim result = FuncLib.Setup.ExportDataFromEmployeeDB()
                            If result.Item1 IsNot Nothing Then
                                Dim grid = result.Item1
                                If grid.Rows.Count > 0 Then
                                    Dim checkBoxColumn As New DataGridViewCheckBoxColumn()
                                    checkBoxColumn.HeaderText = "Select"
                                    checkBoxColumn.Width = 5
                                    checkBoxColumn.Name = "Select"
                                    If BunifuDataGridView1.Columns.Contains("Select") Then
                                        BunifuDataGridView1.Columns.Remove("Select")
                                    End If
                                    BunifuDataGridView1.Columns.Insert(0, checkBoxColumn)
                                    grid.Columns("ID").ColumnName = "SR NO."
                                    grid.Columns("Employee_ID").ColumnName = "EMP ID"
                                    grid.Columns("NT_ID").ColumnName = "NT ID"
                                    grid.Columns("Scanner_ID").ColumnName = "SCAN ID"
                                    grid.Columns("Scanner_5D").ColumnName = "SCAN 5D"
                                    grid.Columns("Employee_Name").ColumnName = "NAME"
                                    grid.Columns("Gender").ColumnName = "GENDER"
                                    grid.Columns("Celebration_Date").ColumnName = "DOB"
                                    grid.Columns("Buisness_Area").ColumnName = "BUISNESS"
                                    grid.Columns("Functions").ColumnName = "DEPARTMENT"
                                    grid.Columns("Leave_Approver").ColumnName = "REPORTING MANAGER"
                                    grid.Columns("Created_By").ColumnName = "CREATED BY"
                                    grid.Columns("Created_At").ColumnName = "CREATED AT"
                                    BunifuDataGridView1.DataSource = grid
                                    Exit Sub
                                End If
                            End If
                        End If
                    End If
                End If
                End If
            End If
    End Sub
    Private Sub BunifuButton19_Click(sender As Object, e As EventArgs) Handles BunifuButton19.Click
        If BunifuDropdown2.SelectedItem = "UPDATE MASTER DATABASE FILE" Then
            GroupBox1.Visible = False
            GroupBox1.Enabled = False
            BunifuDropdown5.Items.Clear()
            BunifuDropdown5.Text = Nothing
            BunifuDropdown5.Visible = False
            BunifuTextBox21.Clear()
            BunifuTextBox21.Visible = False
            CheckBox1.Checked = False
            CheckBox1.Visible = False
            ProgressBar1.Minimum = 0
            ProgressBar1.Value = 0
            ProgressBar1.Visible = False
            BunifuDataGridView1.DataSource = Nothing
            BunifuDataGridView1.Visible = False
            BunifuButton20.Visible = False
            BackgroundWorker1.WorkerSupportsCancellation = True
            BackgroundWorker1.RunWorkerAsync()
        ElseIf BunifuDropdown2.SelectedItem = "REPLACE MASTER DATABASE FILE" Then
            GroupBox1.Visible = False
            GroupBox1.Enabled = False
            BunifuDropdown5.Items.Clear()
            BunifuDropdown5.Text = Nothing
            BunifuDropdown5.Visible = False
            BunifuTextBox21.Clear()
            BunifuTextBox21.Visible = False
            CheckBox1.Checked = False
            CheckBox1.Visible = False
            ProgressBar1.Minimum = 0
            ProgressBar1.Value = 0
            ProgressBar1.Visible = False
            BunifuDataGridView1.DataSource = Nothing
            BunifuDataGridView1.Visible = False
            BunifuButton20.Visible = False
            BackgroundWorker1.WorkerSupportsCancellation = True
            BackgroundWorker1.RunWorkerAsync()
        End If
    End Sub
    Private Sub BackgroundWorker1_DoWork(sender As Object, e As DoWorkEventArgs) Handles BackgroundWorker1.DoWork 'Copy Excel File To Backup Folder
        Dim worker As BackgroundWorker = CType(sender, BackgroundWorker)
        BunifuDataGridView1.Invoke(Sub()
                                       Ini.Load(Inipath)
                                       Dim openFileDialog As New OpenFileDialog()
                                       openFileDialog.Filter = "Excel Files|*.xls;*.xlsx;*.xlsm"
                                       openFileDialog.Title = "Select Database File"
                                       openFileDialog.Multiselect = False
                                       If openFileDialog.ShowDialog() = DialogResult.OK Then
                                           ProgressBar1.Minimum = 0
                                           ProgressBar1.Value = 0
                                           ProgressBar1.Visible = True
                                           ProgressBar1.Enabled = True
                                           Dim selectedFilePath As String = openFileDialog.FileName
                                           BunifuTextBox6.Text = selectedFilePath
                                           Dim destinationDirectory As String = Ini.Sections("BASIC").Keys("Backup_Location").Value.ToString
                                           Dim fileName As String = Path.GetFileName(selectedFilePath)
                                           Dim destinationFilePath As String = Path.Combine(destinationDirectory, fileName)
                                           If Not Directory.Exists(destinationDirectory) Then
                                               Directory.CreateDirectory(destinationDirectory)
                                           End If
                                           If File.Exists(destinationFilePath) Then
                                               File.Delete(destinationFilePath)
                                           End If
                                           Try
                                               ProgressBar1.Minimum = 0
                                               ProgressBar1.Maximum = CInt(New FileInfo(selectedFilePath).Length)
                                               ProgressBar1.Value = 0
                                               Using sourceStream As New FileStream(selectedFilePath, FileMode.Open, FileAccess.Read)
                                                   Using destinationStream As New FileStream(destinationFilePath, FileMode.CreateNew, FileAccess.Write)
                                                       Dim buffer(1024 * 1024) As Byte
                                                       Dim bytesRead As Integer
                                                       Do
                                                           bytesRead = sourceStream.Read(buffer, 0, buffer.Length)
                                                           If bytesRead > 0 Then
                                                               destinationStream.Write(buffer, 0, bytesRead)
                                                               ProgressBar1.Value += bytesRead
                                                           End If
                                                       Loop While bytesRead > 0
                                                   End Using
                                               End Using
                                               BunifuTextBox6.Text = destinationFilePath
                                           Catch ex As Exception
                                               BunifuSnackbar1.Show(Me, $"Error Occured While Saving Backup Copy Of {selectedFilePath} to{destinationDirectory}{Environment.NewLine}Error:{ex.Message}", Bunifu.UI.WinForms.BunifuSnackbar.MessageTypes.Error, 3000)
                                               FuncLib.WriteLog.WriteErrorLog($"Error Occured While Saving Backup Copy Of {selectedFilePath} has been saved to{destinationDirectory} :{ex.Message}")
                                               e.Cancel = True
                                           End Try
                                       ElseIf openFileDialog.ShowDialog() = DialogResult.Cancel Then
                                           e.Cancel = True
                                       End If
                                   End Sub)
    End Sub
    Private Sub BackgroundWorker1_RunWorkerCompleted(sender As Object, e As RunWorkerCompletedEventArgs) Handles BackgroundWorker1.RunWorkerCompleted
        If e.Cancelled = True Then
            Exit Sub
        ElseIf e.Error IsNot Nothing Then
            BunifuSnackbar1.Show(Me, $"Error Occured While Executing BackGroundWorker1_DoWork{Environment.NewLine}Error:{e.Error}", Bunifu.UI.WinForms.BunifuSnackbar.MessageTypes.Error, 3000)
            FuncLib.WriteLog.WriteErrorLog($"Error Occured While Executing BackGroundWorker1_DoWork :{e.Error}")
            Exit Sub
        Else
            Thread.Sleep(5000)
            BackgroundWorker2.WorkerSupportsCancellation = True
            BackgroundWorker2.RunWorkerAsync()
        End If
    End Sub
    Private Sub BackgroundWorker2_DoWork(sender As Object, e As DoWorkEventArgs) Handles BackgroundWorker2.DoWork ' Verify Columns Present in Excel File
        Dim worker As BackgroundWorker = CType(sender, BackgroundWorker)
        BunifuDataGridView1.Invoke(Sub()
                                       Try
                                           Dim destinationFilePath As String = BunifuTextBox6.Text
                                           Dim fileName As String = Path.GetFileName(destinationFilePath)
                                           ProgressBar1.Minimum = 0
                                           ProgressBar1.Value = 0
                                           Dim requiredColumn As New List(Of String) From {"Emp No", "Display Name", "Gender", "Birth Date", "Business Area", "Leave Approver Name", "Function", "NTID", "10 Digit Card Number"}
                                           ExcelPackage.LicenseContext = OfficeOpenXml.LicenseContext.NonCommercial
                                           Using package As New ExcelPackage(New System.IO.FileInfo(destinationFilePath))
                                               Dim worksheet As ExcelWorksheet = package.Workbook.Worksheets(0)
                                               For Each cell As ExcelRangeBase In worksheet.Cells
                                                   cell.Value = cell.Text
                                               Next
                                               Dim headerRow As ExcelRangeBase = worksheet.Cells(1, 1, 1, worksheet.Dimension.End.Column)
                                               Dim columnNames As New List(Of String)
                                               For Each cell As ExcelRangeBase In headerRow
                                                   columnNames.Add(cell.Text.Trim())
                                               Next
                                               ProgressBar1.Maximum = columnNames.Count
                                               Dim CurrentColumn As Integer = 0
                                               Dim extraColumns As New List(Of String)
                                               Dim PresentColumns As New List(Of String)
                                               Dim missingColumns As New List(Of String)
                                               For Each columnName As String In columnNames
                                                   CurrentColumn += 1
                                                   ProgressBar1.Value = CurrentColumn
                                                   If Not requiredColumn.Contains(columnName) Then
                                                       extraColumns.Add(columnName)
                                                   Else
                                                       PresentColumns.Add(columnName)
                                                   End If
                                               Next
                                               For Each requiredColumnName As String In requiredColumn
                                                   If Not columnNames.Contains(requiredColumnName) Then
                                                       missingColumns.Add(requiredColumnName)
                                                   End If
                                               Next
                                               If extraColumns.Count > 0 Then
                                                   BunifuSnackbar1.Show(Me, $"Failed To Verify Excel File {fileName} Reason :Found Extra Columns Inside The Excel File :{String.Join(", ", extraColumns)}", Bunifu.UI.WinForms.BunifuSnackbar.MessageTypes.Warning, 5000)
                                                   FuncLib.WriteLog.WriteAppLog($"Failed To Verify Excel While Updating master Database File : Found Extra Columns Inside The Excel File :{String.Join(", ", extraColumns)}")
                                                   e.Cancel = True
                                                   Exit Sub
                                               ElseIf missingColumns.Count > 0 Then
                                                   BunifuSnackbar1.Show(Me, $"Failed To Verify Excel File {fileName} Reason :Following Columns Are Missing From Excel File {String.Join(", ", missingColumns)}", Bunifu.UI.WinForms.BunifuSnackbar.MessageTypes.Warning, 5000)
                                                   FuncLib.WriteLog.WriteAppLog($"Failed To Verify Excel While Updating master Database File : Following Columns Are Missing From Excel File {String.Join(", ", missingColumns)}")
                                                   e.Cancel = True
                                                   Exit Sub
                                               Else
                                                   'do nothing
                                               End If
                                           End Using
                                       Catch ex As Exception
                                           BunifuSnackbar1.Show(Me, $"Error Occured While Verifying The Columns{Environment.NewLine}Error:{ex.Message}", Bunifu.UI.WinForms.BunifuSnackbar.MessageTypes.Error, 3000)
                                           FuncLib.WriteLog.WriteErrorLog($"Error Occured While Verifying The Columns:{ex.Message}")
                                           e.Cancel = True
                                       End Try
                                   End Sub)
    End Sub
    Private Sub BackgroundWorker2_RunWorkerCompleted(sender As Object, e As RunWorkerCompletedEventArgs) Handles BackgroundWorker2.RunWorkerCompleted
        If e.Cancelled = True Then
            Exit Sub
        ElseIf e.Error IsNot Nothing Then
            BunifuSnackbar1.Show(Me, $"Error Occured While Executing BackGroundWorker2_DoWork{Environment.NewLine}Error:{e.Error}", Bunifu.UI.WinForms.BunifuSnackbar.MessageTypes.Error, 3000)
            FuncLib.WriteLog.WriteErrorLog($"Error Occured While Executing BackGroundWorker2_DoWork :{e.Error}")
            Exit Sub
        Else
            Thread.Sleep(5000)
            BackgroundWorker3.WorkerSupportsCancellation = True
            BackgroundWorker3.RunWorkerAsync()
        End If
    End Sub
    Private Sub BackgroundWorker3_DoWork(sender As Object, e As DoWorkEventArgs) Handles BackgroundWorker3.DoWork
        Dim worker As BackgroundWorker = CType(sender, BackgroundWorker)
        BunifuDataGridView1.Invoke(Sub()
                                       Try
                                           Dim destinationFilePath As String = BunifuTextBox6.Text
                                           Dim fileName As String = Path.GetFileName(destinationFilePath)
                                           ProgressBar1.Minimum = 0
                                           ProgressBar1.Value = 0
                                           Dim columnNames As New List(Of String) From {"Emp No", "Display Name", "Gender", "Business Area", "Leave Approver Name", "Function", "NTID", "10 Digit Card Number"}
                                           ExcelPackage.LicenseContext = OfficeOpenXml.LicenseContext.NonCommercial
                                           Using package As New ExcelPackage(New System.IO.FileInfo(destinationFilePath))
                                               Dim worksheet As ExcelWorksheet = package.Workbook.Worksheets(0)
                                               For Each columnName As String In columnNames
                                                   Dim columnIndex As Integer = GetColumnIndex(worksheet, columnName)
                                                   If columnIndex <> -1 Then
                                                       Dim cellCount As Integer = worksheet.Cells(2, columnIndex, worksheet.Dimension.End.Row, columnIndex).Count(Function(c) Not String.IsNullOrEmpty(c.Text))
                                                       ProgressBar1.Minimum = 0
                                                       ProgressBar1.Value = 0
                                                       ProgressBar1.Maximum = cellCount
                                                       For Each cell As ExcelRangeBase In worksheet.Cells(2, columnIndex, worksheet.Dimension.End.Row, columnIndex)
                                                           If Not String.IsNullOrEmpty(cell.Text) Then
                                                               cell.Style.Numberformat.Format = "@"
                                                               ProgressBar1.Value += 1
                                                           End If
                                                       Next
                                                   End If
                                               Next
                                               package.Save()
                                           End Using
                                       Catch ex As Exception
                                           BunifuSnackbar1.Show(Me, $"Error Occured While Converting Required Columns To String{Environment.NewLine}Error:{ex.Message}", Bunifu.UI.WinForms.BunifuSnackbar.MessageTypes.Error, 3000)
                                           FuncLib.WriteLog.WriteErrorLog($"Error Occured While COnverting Required Columns To String:{ex.Message}")
                                           e.Cancel = True
                                           Exit Sub
                                       End Try
                                   End Sub)
    End Sub
    Private Sub BackgroundWorker3_RunWorkerCompleted(sender As Object, e As RunWorkerCompletedEventArgs) Handles BackgroundWorker3.RunWorkerCompleted
        If e.Cancelled = True Then
            Exit Sub
        ElseIf e.Error IsNot Nothing Then
            BunifuSnackbar1.Show(Me, $"Error Occured While Executing BackGroundWorker3_DoWork{Environment.NewLine}Error:{e.Error}", Bunifu.UI.WinForms.BunifuSnackbar.MessageTypes.Error, 3000)
            FuncLib.WriteLog.WriteErrorLog($"Error Occured While Executing BackGroundWorker3_DoWork :{e.Error}")
            Exit Sub
        Else
            Thread.Sleep(5000)
            BackgroundWorker4.WorkerSupportsCancellation = True
            BackgroundWorker4.RunWorkerAsync()
        End If
    End Sub
    Private Sub BackgroundWorker4_DoWork(sender As Object, e As DoWorkEventArgs) Handles BackgroundWorker4.DoWork
        Dim worker As BackgroundWorker = CType(sender, BackgroundWorker)
        BunifuDataGridView1.Invoke(Sub()
                                       Try
                                           Dim destinationFilePath As String = BunifuTextBox6.Text
                                           Dim fileName As String = Path.GetFileName(destinationFilePath)
                                           Dim sheetName As String = Nothing
                                           ProgressBar1.Minimum = 0
                                           ProgressBar1.Value = 0
                                           Dim connectionString As String = $"Provider=Microsoft.ACE.OLEDB.12.0;Data Source={destinationFilePath};Extended Properties='Excel 12.0;HDR=YES;'"
                                           Dim dataTable As New DataTable()
                                           Using connection As New OleDbConnection(connectionString)
                                               If String.IsNullOrEmpty(sheetName) Then
                                                   connection.Open()
                                                   Dim dtSheets As DataTable = connection.GetOleDbSchemaTable(OleDbSchemaGuid.Tables, Nothing)
                                                   sheetName = dtSheets.Rows(0)("TABLE_NAME").ToString()
                                                   connection.Close()
                                               End If
                                               Dim query As String = $"SELECT [Emp No],[Display Name],[Gender],[Birth Date],[Business Area],[Function],[10 Digit Card Number] FROM [{sheetName}]"
                                               Using adapter As New OleDbDataAdapter(query, connection)
                                                   adapter.Fill(dataTable)
                                               End Using
                                           End Using
                                           ProgressBar1.Maximum = dataTable.Rows.Count
                                           For Each row As DataRow In dataTable.Rows
                                               Dim IsNull As Boolean = False
                                               Dim rowIndex As Integer = dataTable.Rows.IndexOf(row) + 2
                                               For Each columnName In {"Emp No", "Display Name", "Gender", "Birth Date", "Business Area", "Function", "10 Digit Card Number"}
                                                   Dim cellValue As Object = row(columnName)
                                                   If cellValue Is Nothing OrElse IsDBNull(cellValue) OrElse String.IsNullOrEmpty(cellValue.ToString().Trim()) Then
                                                       IsNull = True
                                                   End If
                                               Next
                                               If IsNull Then
                                                   BunifuSnackbar1.Show(Me, $"Failed To Verify Excel File {fileName} Reason : One Null Data Found At {rowIndex}", Bunifu.UI.WinForms.BunifuSnackbar.MessageTypes.Warning, 5000)
                                                   FuncLib.WriteLog.WriteAppLog($"Failed To Verify Excel File {fileName} Reason : One Null Data Found At {rowIndex}")
                                                   e.Cancel = True
                                                   Exit For
                                               End If
                                               ProgressBar1.Value += 1
                                           Next
                                       Catch ex As Exception
                                           BunifuSnackbar1.Show(Me, $"Error Occured While Verifying Null Data{Environment.NewLine}Error:{ex.Message}", Bunifu.UI.WinForms.BunifuSnackbar.MessageTypes.Error, 3000)
                                           FuncLib.WriteLog.WriteErrorLog($"Error Occured While Verifying Null Data:{ex.Message}")
                                           e.Cancel = True
                                           Exit Sub
                                       End Try
                                   End Sub)
    End Sub
    Private Sub BackgroundWorker4_RunWorkerCompleted(sender As Object, e As RunWorkerCompletedEventArgs) Handles BackgroundWorker4.RunWorkerCompleted
        If e.Cancelled = True Then
            Exit Sub
        ElseIf e.Error IsNot Nothing Then
            BunifuSnackbar1.Show(Me, $"Error Occured While Executing BackGroundWorker4_DoWork{Environment.NewLine}Error:{e.Error}", Bunifu.UI.WinForms.BunifuSnackbar.MessageTypes.Error, 3000)
            FuncLib.WriteLog.WriteErrorLog($"Error Occured While Executing BackGroundWorker4_DoWork :{e.Error}")
            Exit Sub
        Else
            Thread.Sleep(5000)
            BackgroundWorker5.WorkerSupportsCancellation = True
            BackgroundWorker5.RunWorkerAsync()
        End If
    End Sub
    Private Sub BackgroundWorker5_DoWork(sender As Object, e As DoWorkEventArgs) Handles BackgroundWorker5.DoWork
        Dim worker As BackgroundWorker = CType(sender, BackgroundWorker)
        BunifuDataGridView1.Invoke(Sub()
                                       Try
                                           Dim destinationFilePath As String = BunifuTextBox6.Text
                                           Dim fileName As String = Path.GetFileName(destinationFilePath)

                                           ProgressBar1.Minimum = 0
                                           ProgressBar1.Value = 0
                                           Dim sheetName As String = Nothing
                                           Dim connectionString As String = $"Provider=Microsoft.ACE.OLEDB.12.0;Data Source={destinationFilePath};Extended Properties='Excel 12.0;HDR=YES;'"
                                           Dim dataTable As New DataTable()
                                           Using connection As New OleDbConnection(connectionString)
                                               If String.IsNullOrEmpty(sheetName) Then
                                                   connection.Open()
                                                   Dim dtSheets As DataTable = connection.GetOleDbSchemaTable(OleDbSchemaGuid.Tables, Nothing)
                                                   sheetName = dtSheets.Rows(0)("TABLE_NAME").ToString()
                                               End If
                                               Dim query As String = $"SELECT * FROM [{sheetName}]"
                                               Using adapter As New OleDbDataAdapter(query, connection)
                                                   adapter.Fill(dataTable)
                                               End Using
                                           End Using
                                           totalRowCount = dataTable.Rows.Count
                                           Dim dict As New Dictionary(Of String, DataRow)()
                                           Dim uniqueDataTable As DataTable = dataTable.Clone()
                                           Dim latestDuplicatesDataTable As DataTable = dataTable.Clone()
                                           For Each row As DataRow In dataTable.Rows
                                               Dim key1 As String = $"{row.Item("Emp No")}"
                                               Dim key2 As String = $"{row.Item("10 Digit Card Number")}"
                                               If dict.ContainsKey(key1) OrElse dict.ContainsKey(key2) Then
                                                   latestDuplicatesDataTable.ImportRow(row)
                                                   duplicateRowCount += 1
                                               Else
                                                   uniqueDataTable.ImportRow(row)
                                                   dict(key1) = row
                                                   dict(key2) = row
                                                   importedRowCount += 1
                                               End If
                                           Next
                                           mergedDataTable = uniqueDataTable.Clone()
                                           For Each column As DataColumn In latestDuplicatesDataTable.Columns
                                               If Not mergedDataTable.Columns.Contains(column.ColumnName) Then
                                                   mergedDataTable.Columns.Add(column.ColumnName, column.DataType)
                                               End If
                                           Next
                                           For Each row As DataRow In uniqueDataTable.Rows
                                               mergedDataTable.ImportRow(row)
                                           Next
                                           For Each row As DataRow In latestDuplicatesDataTable.Rows
                                               Dim key1 As String = $"{row.Item("Emp No")}"
                                               Dim key2 As String = $"{row.Item("10 Digit Card Number")}"
                                               If Not mergedDataTable.AsEnumerable().Any(Function(r) $"{r("Emp No")}" = key1 OrElse
                                                                                         $"{r("10 Digit Card Number")}" = key2) Then
                                                   mergedDataTable.ImportRow(row)
                                               End If
                                           Next
                                           Dim additionalColumnNames As String() = {"Scanner_5D", "Created_By", "Created_At"}
                                           For Each columnName As String In additionalColumnNames
                                               If Not mergedDataTable.Columns.Contains(columnName) Then
                                                   Select Case columnName
                                                       Case "Scanner_5D", "Created_By"
                                                           mergedDataTable.Columns.Add(columnName, GetType(String))
                                                       Case "Created_At"
                                                           mergedDataTable.Columns.Add(columnName, GetType(DateTime))
                                                   End Select
                                               End If
                                           Next
                                           ProgressBar1.Maximum = mergedDataTable.Rows.Count
                                           For Each row As DataRow In mergedDataTable.Rows
                                               row("Scanner_5D") = FuncLib.Gifts.ConvertESDNumberToCardNumber(row("10 Digit Card Number"))
                                               row("Created_By") = $"{Environment.UserName}"
                                               row("Created_At") = DateTime.Now.ToString("G")
                                               ProgressBar1.Value += 1
                                           Next
                                       Catch ex As Exception
                                           BunifuSnackbar1.Show(Me, $"Error Occured While Populating Data To DataGridView{Environment.NewLine}Error:{ex.Message}", Bunifu.UI.WinForms.BunifuSnackbar.MessageTypes.Error, 3000)
                                           FuncLib.WriteLog.WriteErrorLog($"Error Occured While Populating Data To Datagridview:{ex.Message}")
                                           e.Cancel = True
                                           Exit Sub
                                       End Try
                                   End Sub)
    End Sub
    Private Sub BackgroundWorker5_RunWorkerCompleted(sender As Object, e As RunWorkerCompletedEventArgs) Handles BackgroundWorker5.RunWorkerCompleted
        If e.Cancelled = True Then
            Exit Sub
        ElseIf e.Error IsNot Nothing Then
            BunifuSnackbar1.Show(Me, $"Error Occured While Executing BackGroundWorker5_DoWork{Environment.NewLine}Error:{e.Error}", Bunifu.UI.WinForms.BunifuSnackbar.MessageTypes.Error, 3000)
            FuncLib.WriteLog.WriteErrorLog($"Error Occured While Executing BackGroundWorker5_DoWork :{e.Error}")
            Exit Sub
        Else
            BunifuSnackbar1.Show(Me, $"All Possible Has Been Populated To The DataGrid SuccessFully!", Bunifu.UI.WinForms.BunifuSnackbar.MessageTypes.Success, 3000)
            BunifuDataGridView1.Invoke(Sub()
                                           If BunifuDataGridView1.Columns.Contains("Select") Then
                                               BunifuDataGridView1.Columns.Remove("Select")
                                           End If
                                           mergedDataTable.Columns("Emp No").ColumnName = "EMP ID"
                                           mergedDataTable.Columns("NTID").ColumnName = "NT ID"
                                           mergedDataTable.Columns("10 Digit Card Number").ColumnName = "SCAN ID"
                                           mergedDataTable.Columns("Scanner_5D").ColumnName = "SCAN 5D"
                                           mergedDataTable.Columns("Display Name").ColumnName = "NAME"
                                           mergedDataTable.Columns("Gender").ColumnName = "GENDER"
                                           mergedDataTable.Columns("Business Area").ColumnName = "BUISNESS"
                                           mergedDataTable.Columns("Function").ColumnName = "DEPARTMENT"
                                           mergedDataTable.Columns("Birth Date").ColumnName = "DOB"
                                           mergedDataTable.Columns("Leave Approver Name").ColumnName = "REPORTING MANAGER"
                                           mergedDataTable.Columns("Created_By").ColumnName = "CREATED BY"
                                           mergedDataTable.Columns("Created_At").ColumnName = "CREATED AT"
                                           GroupBox1.Visible = True
                                           GroupBox1.Enabled = True
                                           BunifuDropdown5.Items.Clear()
                                           BunifuDropdown5.Text = Nothing
                                           BunifuDropdown5.Visible = True
                                           BunifuDropdown5.Enabled = True
                                           BunifuTextBox21.Clear()
                                           BunifuTextBox21.Visible = True
                                           BunifuTextBox21.Enabled = True
                                           CheckBox1.Checked = False
                                           CheckBox1.Visible = False
                                           BunifuButton11.Enabled = True
                                           BunifuButton11.Visible = True
                                           BunifuDataGridView1.Visible = True
                                           BunifuDataGridView1.Enabled = True
                                           BunifuDataGridView1.DataSource = mergedDataTable
                                           If BunifuDataGridView1.Rows.Count > 0 Then
                                               BunifuButton20.Visible = True
                                               BunifuButton20.Enabled = True
                                           Else
                                               BunifuButton20.Visible = False
                                               BunifuButton20.Enabled = False
                                           End If
                                           For Each column As DataGridViewColumn In BunifuDataGridView1.Columns
                                               BunifuDropdown5.Items.Add(column.HeaderText)
                                           Next
                                           AddHandler BunifuButton11.LeftIcon.MouseEnter, Sub(sender1 As Object, e1 As EventArgs)
                                                                                              toolTip.SetToolTip(BunifuButton11, $"Total Row Fetched From Excel : {totalRowCount}{Environment.NewLine}Duplicate Rows Found:{duplicateRowCount}{Environment.NewLine}Imported Rows Count:{importedRowCount}{additionalInfo}")
                                                                                          End Sub
                                           AddHandler BunifuButton11.LeftIcon.MouseLeave, Sub(sender1 As Object, e1 As EventArgs)
                                                                                              toolTip.Hide()
                                                                                          End Sub
                                       End Sub)
        End If
    End Sub
    Private Shared Function GetColumnIndex(worksheet As ExcelWorksheet, columnName As String) As Integer
        For col As Integer = 1 To worksheet.Dimension.End.Column
            If worksheet.Cells(1, col).Text.Trim().Equals(columnName, StringComparison.OrdinalIgnoreCase) Then
                Return col
            End If
        Next
        Return -1
    End Function
    Private Sub BunifuButton29_Click(sender As Object, e As EventArgs) Handles BunifuButton29.Click
        ClearUserAddition()
    End Sub
    Private Sub BunifuButton20_Click(sender As Object, e As EventArgs) Handles BunifuButton20.Click
        If BunifuDropdown2.SelectedItem = "REMOVE FROM MASTER DATABASE FILE" Then
            If selectedRows.Count > 0 Then
                totalRowCount = 0
                duplicateRowCount = 0
                AffectedEmployeeCOunt = 0
                importedRowCount = 0
                UnaffectedEmployeeCount = 0
                BackgroundWorker8.WorkerSupportsCancellation = True
                BackgroundWorker8.RunWorkerAsync()
            Else
                BunifuSnackbar1.Show(Me, $"No Employees Found To Be Removed!", Bunifu.UI.WinForms.BunifuSnackbar.MessageTypes.Information, 3000)
                Exit Sub
            End If
        ElseIf BunifuDropdown2.SelectedItem = "REPLACE MASTER DATABASE FILE" Then
            totalRowCount = 0
            duplicateRowCount = 0
            AffectedEmployeeCOunt = 0
            importedRowCount = 0
            UnaffectedEmployeeCount = 0
            BackgroundWorker7.WorkerSupportsCancellation = True
            BackgroundWorker7.RunWorkerAsync()
        ElseIf BunifuDropdown2.SelectedItem = "UPDATE MASTER DATABASE FILE" Then
            totalRowCount = 0
            duplicateRowCount = 0
            AffectedEmployeeCOunt = 0
            importedRowCount = 0
            UnaffectedEmployeeCount = 0
            BackgroundWorker6.WorkerSupportsCancellation = True
            BackgroundWorker6.RunWorkerAsync()
        End If
    End Sub
    Private Sub BackgroundWorker8_DoWork(sender As Object, e As DoWorkEventArgs) Handles BackgroundWorker8.DoWork
        Dim worker As BackgroundWorker = CType(sender, BackgroundWorker)
        BunifuDataGridView1.Invoke(Sub()
                                       Try
                                           ProgressBar1.Minimum = 0
                                           ProgressBar1.Value = 0
                                           For Each row As DataGridViewRow In selectedRows
                                               Dim employeeID As String = row.Cells("EMP ID").Value.ToString()
                                               Dim RemoveEmployee As String = FuncLib.Setup.RemoveSelectedEmployessFromDB(employeeID)
                                               If RemoveEmployee = "True" Then
                                                   ProgressBar1.Value += 1
                                                   Continue For
                                               Else
                                                   BunifuSnackbar1.Show(Me, $"Error Occured While Removing Employees From Database.{Environment.NewLine}Error:{RemoveEmployee}", Bunifu.UI.WinForms.BunifuSnackbar.MessageTypes.Error, 3000)
                                                   FuncLib.WriteLog.WriteErrorLog($"Error Ocurred While Executing FuncLib.Setup.RemoveSelectedEmployessFromDB :{RemoveEmployee}")
                                                   e.Cancel = True
                                                   Exit For
                                               End If
                                           Next
                                       Catch ex As Exception
                                           BunifuSnackbar1.Show(Me, $"Error Occured While Removing The Employee Master Database!{Environment.NewLine}Error:{ex.Message}", Bunifu.UI.WinForms.BunifuSnackbar.MessageTypes.Error, 3000)
                                           FuncLib.WriteLog.WriteErrorLog($"Error Occured While Removing The Employee Master Database :{ex.Message}")
                                           e.Cancel = True
                                       End Try
                                   End Sub)
    End Sub
    Private Sub BackgroundWorker8_RunWorkerCompleted(sender As Object, e As RunWorkerCompletedEventArgs) Handles BackgroundWorker8.RunWorkerCompleted
        If e.Cancelled = True Then
            Exit Sub
        ElseIf e.Error IsNot Nothing Then
            BunifuSnackbar1.Show(Me, $"Error Occured While Executing BackgroundWorker6_DoWork!{Environment.NewLine}Error :{e.Error}", Bunifu.UI.WinForms.BunifuSnackbar.MessageTypes.Error, 5000)
            FuncLib.WriteLog.WriteErrorLog($"Error Occured While Executing backgroundWorker6_Dowork :{e.Error}")
            Exit Sub
        Else
            BunifuDataGridView1.Invoke(Sub()
                                           BunifuSnackbar1.Show(Me, $"All The Required action Are Taken And Executed Successfully.{Environment.NewLine}Employee Added :{AffectedEmployeeCOunt}{Environment.NewLine}Duplicate Employee Found :{UnaffectedEmployeeCount}")
                                           BunifuDataGridView1.DataSource = Nothing
                                           Thread.Sleep(1000)
                                           selectedRows.Clear()
                                           BunifuButton20.Visible = False
                                           Dim result = FuncLib.Setup.ExportDataFromEmployeeDB()
                                           If result.Item1 IsNot Nothing Then
                                               Dim grid = result.Item1
                                               If grid.Rows.Count > 0 Then
                                                   ProgressBar1.Visible = True
                                                   CheckBox1.Checked = False
                                                   CheckBox1.Visible = True
                                                   ProgressBar1.Minimum = 0
                                                   ProgressBar1.Value = 0
                                                   ProgressBar1.Maximum = grid.Rows.Count
                                                   Dim CurrentRowCount As Integer = 0
                                                   Dim checkBoxColumn As New DataGridViewCheckBoxColumn()
                                                   checkBoxColumn.HeaderText = "Select"
                                                   checkBoxColumn.Width = 5
                                                   checkBoxColumn.Name = "Select"
                                                   If BunifuDataGridView1.Columns.Contains("Select") Then
                                                       BunifuDataGridView1.Columns.Remove("Select")
                                                   End If
                                                   BunifuDataGridView1.Columns.Insert(0, checkBoxColumn)
                                                   grid.Columns("ID").ColumnName = "SR NO."
                                                   grid.Columns("Employee_ID").ColumnName = "EMP ID"
                                                   grid.Columns("NT_ID").ColumnName = "NT ID"
                                                   grid.Columns("Scanner_ID").ColumnName = "SCAN ID"
                                                   grid.Columns("Scanner_5D").ColumnName = "SCAN 5D"
                                                   grid.Columns("Employee_Name").ColumnName = "NAME"
                                                   grid.Columns("Gender").ColumnName = "GENDER"
                                                   grid.Columns("Celebration_Date").ColumnName = "DOB"
                                                   grid.Columns("Buisness_Area").ColumnName = "BUISNESS"
                                                   grid.Columns("Functions").ColumnName = "DEPARTMENT"
                                                   grid.Columns("Leave_Approver").ColumnName = "REPORTING MANAGER"
                                                   grid.Columns("Created_By").ColumnName = "CREATED BY"
                                                   grid.Columns("Created_At").ColumnName = "CREATED AT"
                                                   For Each row As DataRow In grid.Rows
                                                       CurrentRowCount += 1
                                                       ProgressBar1.Value = CurrentRowCount
                                                   Next
                                                   BunifuDataGridView1.DataSource = grid
                                                   BunifuDataGridView1.Visible = True
                                                   BunifuDataGridView1.Enabled = True
                                                   BunifuButton29.Visible = True
                                                   BunifuButton29.Enabled = True
                                                   GroupBox1.Visible = True
                                                   GroupBox1.Enabled = True
                                                   BunifuDropdown5.Visible = True
                                                   BunifuDropdown5.Enabled = True
                                                   BunifuDropdown5.Items.Clear()
                                                   BunifuDropdown5.Text = Nothing
                                                   BunifuTextBox21.Visible = True
                                                   BunifuTextBox21.Enabled = True
                                                   For Each column As DataGridViewColumn In BunifuDataGridView1.Columns
                                                       BunifuDropdown5.Items.Add(column.HeaderText)
                                                   Next
                                               End If
                                           End If
                                       End Sub)

        End If
    End Sub
    Private Sub BackgroundWorker6_DoWork(sender As Object, e As DoWorkEventArgs) Handles BackgroundWorker6.DoWork
        Dim worker As BackgroundWorker = CType(sender, BackgroundWorker)
        BunifuDataGridView1.Invoke(Sub()
                                       Try
                                           ProgressBar1.Minimum = 0
                                           ProgressBar1.Value = 0
                                           If BunifuDataGridView1.Rows.Count > 0 Then
                                               ProgressBar1.Maximum = BunifuDataGridView1.Rows.Count - 1
                                               notInsertedDataTable.Columns.Add("EMP ID", GetType(String))
                                               notInsertedDataTable.Columns.Add("NT ID", GetType(String))
                                               notInsertedDataTable.Columns.Add("SCAN ID", GetType(String))
                                               notInsertedDataTable.Columns.Add("SCAN 5D", GetType(String))
                                               notInsertedDataTable.Columns.Add("NAME", GetType(String))
                                               notInsertedDataTable.Columns.Add("GENDER", GetType(String))
                                               notInsertedDataTable.Columns.Add("DOB", GetType(Date))
                                               notInsertedDataTable.Columns.Add("BUISNESS", GetType(String))
                                               notInsertedDataTable.Columns.Add("DEPARTMENT", GetType(String))
                                               notInsertedDataTable.Columns.Add("REPORTING MANAGER", GetType(String))
                                               notInsertedDataTable.Columns.Add("CREATED BY", GetType(String))
                                               notInsertedDataTable.Columns.Add("CREATED AT", GetType(DateTime))
                                               For i As Integer = 0 To BunifuDataGridView1.RowCount - 2
                                                   Dim row As DataGridViewRow = BunifuDataGridView1.Rows(i)
                                                   Dim Employee_ID As String = row.Cells("EMP ID").Value.ToString()
                                                   Dim NT_ID As String = row.Cells("NT ID").Value
                                                   Dim Scanner_ID As String = row.Cells("SCAN ID").Value
                                                   Dim Scanner_5D As String = row.Cells("SCAN 5D").Value
                                                   Dim Employee_Name As String = row.Cells("NAME").Value
                                                   Dim Gender As String = row.Cells("GENDER").Value
                                                   Dim Celebration_Date As Date = DateTime.Parse(row.Cells("DOB").Value)
                                                   Dim Buisness_Area As String = row.Cells("BUISNESS").Value
                                                   Dim Functions As String = row.Cells("DEPARTMENT").Value
                                                   Dim Leave_Approver As String = row.Cells("REPORTING MANAGER").Value
                                                   Dim Created_By As String = row.Cells("CREATED BY").Value
                                                   Dim Created_At As DateTime = row.Cells("CREATED AT").Value
                                                   Dim ImportData As String = FuncLib.Setup.AddEmployeesToEmployeeDB(Employee_ID, NT_ID, Scanner_ID, Scanner_5D, Employee_Name, Gender, Celebration_Date, Buisness_Area, Functions, Leave_Approver, Created_By, Created_At)
                                                   If ImportData = "True" Then
                                                       AffectedEmployeeCOunt += 1
                                                   ElseIf ImportData = "False" Then
                                                       Dim newRow As DataRow = notInsertedDataTable.NewRow()
                                                       newRow("EMP ID") = Employee_ID
                                                       newRow("NT ID") = NT_ID
                                                       newRow("SCAN ID") = Scanner_ID
                                                       newRow("SCAN 5D") = Scanner_5D
                                                       newRow("NAME") = Employee_Name
                                                       newRow("GENDER") = Gender
                                                       newRow("DOB") = Celebration_Date
                                                       newRow("BUISNESS") = Buisness_Area
                                                       newRow("DEPARTMENT") = Functions
                                                       newRow("REPORTING MANAGER") = Leave_Approver
                                                       newRow("CREATED BY") = Created_By
                                                       newRow("CREATED AT") = Created_At
                                                       notInsertedDataTable.Rows.Add(newRow)
                                                       UnaffectedEmployeeCount += 1
                                                   Else
                                                       BunifuSnackbar1.Show(Me, $"Error Occured While Updating master Database File.{Environment.NewLine}Error:{ImportData}", Bunifu.UI.WinForms.BunifuSnackbar.MessageTypes.Error, 3000)
                                                       FuncLib.WriteLog.WriteErrorLog($"Error Ocurred While Executing FuncLib.Setup.AddEmployeesToEmployeeDB :{ImportData}")
                                                       e.Cancel = True
                                                       Exit For
                                                   End If
                                                   ProgressBar1.Value += 1
                                               Next
                                           Else
                                               BunifuSnackbar1.Show(Me, $"No Employees Found To Be Imported!", Bunifu.UI.WinForms.BunifuSnackbar.MessageTypes.Information, 3000)
                                               e.Cancel = True
                                           End If
                                       Catch ex As Exception
                                           BunifuSnackbar1.Show(Me, $"Error Occured While Updating The Employee Master Database!{Environment.NewLine}Error:{ex.Message}", Bunifu.UI.WinForms.BunifuSnackbar.MessageTypes.Error, 3000)
                                           FuncLib.WriteLog.WriteErrorLog($"Error Occured While Updating The Employee Master Database :{ex.Message}")
                                           e.Cancel = True
                                       End Try
                                   End Sub)
    End Sub
    Private Sub BackgroundWorker6_RunWorkerCompleted(sender As Object, e As RunWorkerCompletedEventArgs) Handles BackgroundWorker6.RunWorkerCompleted
        If e.Cancelled = True Then
            Exit Sub
        ElseIf e.Error IsNot Nothing Then
            BunifuSnackbar1.Show(Me, $"Error Occured While Executing BackgroundWorker6_DoWork!{Environment.NewLine}Error :{e.Error}", Bunifu.UI.WinForms.BunifuSnackbar.MessageTypes.Error, 5000)
            FuncLib.WriteLog.WriteErrorLog($"Error Occured While Executing backgroundWorker6_Dowork :{e.Error}")
            Exit Sub
        Else
            BunifuDataGridView1.Invoke(Sub()
                                           BunifuSnackbar1.Show(Me, $"All The Required action Are Taken And Executed Successfully.{Environment.NewLine}Employee Added :{AffectedEmployeeCOunt}{Environment.NewLine}Duplicate Employee Found :{UnaffectedEmployeeCount}")
                                           BunifuDataGridView1.DataSource = Nothing
                                           Thread.Sleep(1000)
                                           BunifuDataGridView1.DataSource = notInsertedDataTable
                                           BunifuButton29.Visible = False
                                           GroupBox2.Enabled = False
                                           BunifuShadowPanel6.Enabled = False
                                           BunifuButton20.Visible = False
                                           BunifuButton34.Visible = True
                                           BunifuButton34.Enabled = True
                                           AddHandler BunifuButton11.LeftIcon.MouseEnter, Sub(sender1 As Object, e1 As EventArgs)
                                                                                              toolTip.SetToolTip(BunifuButton11, $"Total Row Fetched From Excel : {totalRowCount}{Environment.NewLine}Duplicate Rows Found:{duplicateRowCount}{Environment.NewLine}Imported Rows Count:{importedRowCount}{Environment.NewLine}Employee Added :{AffectedEmployeeCOunt}{Environment.NewLine}Duplicate Employee Found :{UnaffectedEmployeeCount}{additionalInfo}")
                                                                                          End Sub
                                           AddHandler BunifuButton11.LeftIcon.MouseLeave, Sub(sender1 As Object, e1 As EventArgs)
                                                                                              toolTip.Hide()
                                                                                          End Sub
                                       End Sub)

        End If
    End Sub
    Private Sub BunifuButton34_Click(sender As Object, e As EventArgs) Handles BunifuButton34.Click
        ClearUserAddition()
    End Sub
    Private Sub BackgroundWorker7_DoWork(sender As Object, e As DoWorkEventArgs) Handles BackgroundWorker7.DoWork
        Dim worker As BackgroundWorker = CType(sender, BackgroundWorker)
        BunifuDataGridView1.Invoke(Sub()
                                       Try
                                           ProgressBar1.Minimum = 0
                                           ProgressBar1.Value = 0
                                           If BunifuDataGridView1.Rows.Count > 0 Then
                                               Dim DeleteEmployeeDB As String = FuncLib.Setup.ReplaceMaterFile()
                                               If DeleteEmployeeDB = "True" Then
                                                   ProgressBar1.Maximum = BunifuDataGridView1.Rows.Count - 1
                                                   notInsertedDataTable.Columns.Add("EMP ID", GetType(String))
                                                   notInsertedDataTable.Columns.Add("NT ID", GetType(String))
                                                   notInsertedDataTable.Columns.Add("SCAN ID", GetType(String))
                                                   notInsertedDataTable.Columns.Add("SCAN 5D", GetType(String))
                                                   notInsertedDataTable.Columns.Add("NAME", GetType(String))
                                                   notInsertedDataTable.Columns.Add("GENDER", GetType(String))
                                                   notInsertedDataTable.Columns.Add("DOB", GetType(Date))
                                                   notInsertedDataTable.Columns.Add("BUISNESS", GetType(String))
                                                   notInsertedDataTable.Columns.Add("DEPARTMENT", GetType(String))
                                                   notInsertedDataTable.Columns.Add("REPORTING MANAGER", GetType(String))
                                                   notInsertedDataTable.Columns.Add("CREATED BY", GetType(String))
                                                   notInsertedDataTable.Columns.Add("CREATED AT", GetType(DateTime))
                                                   For i As Integer = 0 To BunifuDataGridView1.RowCount - 2
                                                       Dim row As DataGridViewRow = BunifuDataGridView1.Rows(i)
                                                       Dim Employee_ID As String = row.Cells("EMP ID").Value.ToString()
                                                       Dim NT_ID As String = row.Cells("NT ID").Value
                                                       Dim Scanner_ID As String = row.Cells("SCAN ID").Value
                                                       Dim Scanner_5D As String = row.Cells("SCAN 5D").Value
                                                       Dim Employee_Name As String = row.Cells("NAME").Value
                                                       Dim Gender As String = row.Cells("GENDER").Value
                                                       Dim Celebration_Date As Date = DateTime.Parse(row.Cells("DOB").Value)
                                                       Dim Buisness_Area As String = row.Cells("BUISNESS").Value
                                                       Dim Functions As String = row.Cells("DEPARTMENT").Value
                                                       Dim Leave_Approver As String = row.Cells("REPORTING MANAGER").Value
                                                       Dim Created_By As String = row.Cells("CREATED BY").Value
                                                       Dim Created_At As DateTime = row.Cells("CREATED AT").Value
                                                       Dim ImportData As String = FuncLib.Setup.AddEmployeesToEmployeeDB(Employee_ID, NT_ID, Scanner_ID, Scanner_5D, Employee_Name, Gender, Celebration_Date, Buisness_Area, Functions, Leave_Approver, Created_By, Created_At)
                                                       If ImportData = "True" Then
                                                           AffectedEmployeeCOunt += 1
                                                       ElseIf ImportData = "False" Then
                                                           Dim newRow As DataRow = notInsertedDataTable.NewRow()
                                                           newRow("EMP ID") = Employee_ID
                                                           newRow("NT ID") = NT_ID
                                                           newRow("SCAN ID") = Scanner_ID
                                                           newRow("SCAN 5D") = Scanner_5D
                                                           newRow("NAME") = Employee_Name
                                                           newRow("GENDER") = Gender
                                                           newRow("DOB") = Celebration_Date
                                                           newRow("BUISNESS") = Buisness_Area
                                                           newRow("DEPARTMENT") = Functions
                                                           newRow("REPORTING MANAGER") = Leave_Approver
                                                           newRow("CREATED BY") = Created_By
                                                           newRow("CREATED AT") = Created_At
                                                           notInsertedDataTable.Rows.Add(newRow)
                                                           UnaffectedEmployeeCount += 1
                                                       Else
                                                           BunifuSnackbar1.Show(Me, $"Error Occured While Updating master Database File.{Environment.NewLine}Error:{ImportData}", Bunifu.UI.WinForms.BunifuSnackbar.MessageTypes.Error, 3000)
                                                           FuncLib.WriteLog.WriteErrorLog($"Error Ocurred While Executing FuncLib.Setup.AddEmployeesToEmployeeDB :{ImportData}")
                                                           e.Cancel = True
                                                           Exit For
                                                       End If
                                                       ProgressBar1.Value += 1
                                                   Next
                                               Else
                                                   BunifuSnackbar1.Show(Me, $"Error Occured While Erasing Employee Database.{Environment.NewLine}Error:{DeleteEmployeeDB}", Bunifu.UI.WinForms.BunifuSnackbar.MessageTypes.Error, 3000)
                                                   FuncLib.WriteLog.WriteErrorLog($"Error Occured While Executing FuncLib.Setup.ReplaceMaterFile:{DeleteEmployeeDB} ")
                                                   e.Cancel = True
                                               End If
                                           Else
                                               BunifuSnackbar1.Show(Me, $"No Employees Found To Be Imported!", Bunifu.UI.WinForms.BunifuSnackbar.MessageTypes.Information, 3000)
                                               e.Cancel = True
                                           End If
                                       Catch ex As Exception
                                           BunifuSnackbar1.Show(Me, $"Error Occured While Updating The Employee Master Database!{Environment.NewLine}Error:{ex.Message}", Bunifu.UI.WinForms.BunifuSnackbar.MessageTypes.Error, 3000)
                                           FuncLib.WriteLog.WriteErrorLog($"Error Occured While Updating The Employee Master Database :{ex.Message}")
                                           e.Cancel = True
                                       End Try
                                   End Sub)
    End Sub
    Private Sub BackgroundWorker7_RunWorkerCompleted(sender As Object, e As RunWorkerCompletedEventArgs) Handles BackgroundWorker7.RunWorkerCompleted
        If e.Cancelled = True Then
            Exit Sub
        ElseIf e.Error IsNot Nothing Then
            BunifuSnackbar1.Show(Me, $"Error Occured While Executing BackgroundWorker7_DoWork!{Environment.NewLine}Error :{e.Error}", Bunifu.UI.WinForms.BunifuSnackbar.MessageTypes.Error, 5000)
            FuncLib.WriteLog.WriteErrorLog($"Error Occured While Executing backgroundWorker7_Dowork :{e.Error}")
            Exit Sub
        Else
            BunifuDataGridView1.Invoke(Sub()
                                           BunifuSnackbar1.Show(Me, $"All The Required action Are Taken And Executed Successfully.{Environment.NewLine}Employee Added :{AffectedEmployeeCOunt}{Environment.NewLine}Duplicate Employee Found :{UnaffectedEmployeeCount}")
                                           BunifuDataGridView1.DataSource = Nothing
                                           Thread.Sleep(1000)
                                           BunifuDataGridView1.DataSource = notInsertedDataTable
                                           BunifuButton29.Visible = False
                                           GroupBox2.Enabled = False
                                           BunifuShadowPanel6.Enabled = False
                                           BunifuButton20.Visible = False
                                           BunifuButton34.Visible = True
                                           BunifuButton34.Enabled = True
                                           AddHandler BunifuButton11.LeftIcon.MouseEnter, Sub(sender1 As Object, e1 As EventArgs)
                                                                                              toolTip.SetToolTip(BunifuButton11, $"Total Row Fetched From Excel : {totalRowCount}{Environment.NewLine}Duplicate Rows Found:{duplicateRowCount}{Environment.NewLine}Imported Rows Count:{importedRowCount}{Environment.NewLine}Employee Added :{AffectedEmployeeCOunt}{Environment.NewLine}Duplicate Employee Found :{UnaffectedEmployeeCount}{additionalInfo}")
                                                                                          End Sub
                                           AddHandler BunifuButton11.LeftIcon.MouseLeave, Sub(sender1 As Object, e1 As EventArgs)
                                                                                              toolTip.Hide()
                                                                                          End Sub
                                       End Sub)

        End If
    End Sub
    Private Sub BunifuTextBox21_TextChanged(sender As Object, e As EventArgs) Handles BunifuTextBox21.TextChanged
        Dim selectedColumnIndex As Integer = BunifuDropdown5.SelectedIndex
        Dim columnName As String = If(selectedColumnIndex >= 0, BunifuDropdown5.Items(selectedColumnIndex).ToString(), "")
        If BunifuDataGridView1.DataSource IsNot Nothing AndAlso Not String.IsNullOrEmpty(columnName) Then
            Dim bindingSource As New BindingSource()
            bindingSource.DataSource = BunifuDataGridView1.DataSource
            bindingSource.Filter = $"[{columnName}] LIKE '%{BunifuTextBox21.Text}%'"
            BunifuDataGridView1.DataSource = bindingSource
        End If
    End Sub
    Private Sub BunifuButton6_Click(sender As Object, e As EventArgs) Handles BunifuButton6.Click
        BunifuDataGridView2.DataSource = Nothing
        GroupBox17.Visible = False
        GroupBox17.Enabled = False
        BunifuButton24.Enabled = False
        BunifuButton24.Visible = False
        BunifuPanel2.Visible = False
        BunifuPanel2.Enabled = False
        BunifuButton30.Enabled = False
        BunifuButton30.Visible = False
        BunifuButton31.Enabled = False
        BunifuButton31.Visible = False
        BunifuTextBox12.Clear()
        BunifuTextBox8.Clear()
        BunifuTextBox9.Clear()
        BunifuTextBox10.Clear()
        BunifuTextBox11.Clear()
        BunifuTextBox13.Clear()
        BunifuTextBox15.Clear()
        BunifuTextBox16.Clear()
        BunifuTextBox17.Clear()
        BunifuToggleSwitch3.Checked = False
        BunifuToggleSwitch4.Checked = False
        BunifuToggleSwitch5.Checked = False
        BunifuDropdown16.Items.Clear()
        BunifuDropdown16.Text = Nothing
        Dim result = FuncLib.Setup.ExportDataFromUserDB()
        If result.Item1 IsNot Nothing Then
            Dim grid = result.Item1
            grid.Columns("ID").ColumnName = "SR NO."
            grid.Columns("NT_ID").ColumnName = "NT ID"
            grid.Columns("Scanner_ID").ColumnName = "SCAN ID"
            grid.Columns("Scanner_5D").ColumnName = "SCAN 5D"
            grid.Columns("Employee_Name").ColumnName = "NAME"
            grid.Columns("Gender").ColumnName = "GENDER"
            grid.Columns("Buisness_Area").ColumnName = "BUISNESS"
            grid.Columns("Functions").ColumnName = "DEPARTMENT"
            grid.Columns("Leave_Approver").ColumnName = "REPORTING MANAGER"
            grid.Columns("Modified_By").ColumnName = "MODIFIED BY"
            grid.Columns("Modified_At").ColumnName = "MODIFIED AT"
            grid.Columns("Issue_Gift").ColumnName = "ISSUING"
            grid.Columns("Events_Addition").ColumnName = "EVENT ADDING"
            grid.Columns("Admn_Access").ColumnName = "ADMINISTRATIVE"
            BunifuDataGridView2.DataSource = grid
            GroupBox17.Visible = True
            GroupBox17.Enabled = True
            If BunifuPages1.Visible = False Then
                BunifuPages1.TabPages.Clear()
                transition.ShowSync(BunifuPages1, False, BunifuAnimatorNS.Animation.ScaleAndHorizSlide)
                BunifuPages1.TabPages.Add(TabPage4)
            Else
                transition.HideSync(BunifuPages1, False, BunifuAnimatorNS.Animation.ScaleAndHorizSlide)
                BunifuPages1.TabPages.Clear()
                transition.ShowSync(BunifuPages1, False, BunifuAnimatorNS.Animation.ScaleAndHorizSlide)
                BunifuPages1.TabPages.Add(TabPage4)
            End If
            For Each column As DataGridViewColumn In BunifuDataGridView2.Columns
                BunifuDropdown16.Items.Add(column.HeaderText)
            Next
            If BunifuDataGridView2.Rows.Count > 0 AndAlso BunifuDataGridView2.SelectedRows.Count > 0 Then
                BunifuButton31.Enabled = True
                BunifuButton31.Visible = True
            Else
                BunifuButton31.Enabled = False
                BunifuButton31.Visible = False
            End If
            BunifuButton30.Enabled = True
            BunifuButton30.Visible = True
        Else
            BunifuSnackbar1.Show(Me, $"Error Occured While Fetchin Dta From DataBase.{Environment.NewLine}Error:{result.Item2}", Bunifu.UI.WinForms.BunifuSnackbar.MessageTypes.Error, 2000)
            FuncLib.WriteLog.WriteErrorLog($"Error Occured While Executing FuncLib.Setup.ExportDataFromEmployeeDB:{result.Item2}")
            Exit Sub
        End If
    End Sub
    Private Sub BunifuTextBox12_TextChanged(sender As Object, e As EventArgs) Handles BunifuTextBox12.TextChanged
        Dim selectedColumnIndex As Integer = BunifuDropdown16.SelectedIndex
        Dim columnName As String = If(selectedColumnIndex >= 0, BunifuDropdown16.Items(selectedColumnIndex).ToString(), "")
        If BunifuDataGridView2.DataSource IsNot Nothing AndAlso Not String.IsNullOrEmpty(columnName) Then
            Dim bindingSource As New BindingSource()
            bindingSource.DataSource = BunifuDataGridView2.DataSource
            bindingSource.Filter = $"[{columnName}] LIKE '%{BunifuTextBox12.Text}%'"
            BunifuDataGridView2.DataSource = bindingSource
        End If
    End Sub
    Private Sub BunifuDataGridView2_SelectionChanged(sender As Object, e As EventArgs) Handles BunifuDataGridView2.SelectionChanged
        If BunifuDataGridView2.SelectedRows.Count > 0 AndAlso BunifuDataGridView2.SelectedRows.Count <= 1 Then
            Dim SelectedRow As DataGridViewRow = BunifuDataGridView2.SelectedRows(0)
            If IsDBNull(SelectedRow.Cells("NT ID").Value) Then
                BunifuTextBox8.Text = Nothing
            Else
                BunifuTextBox8.Text = SelectedRow.Cells("NT ID").Value
            End If
            BunifuTextBox9.Text = SelectedRow.Cells("SCAN ID").Value
            BunifuTextBox10.Text = SelectedRow.Cells("SCAN 5D").Value
            BunifuTextBox11.Text = SelectedRow.Cells("NAME").Value
            BunifuTextBox13.Text = SelectedRow.Cells("GENDER").Value
            BunifuTextBox15.Text = SelectedRow.Cells("BUISNESS").Value
            BunifuTextBox16.Text = SelectedRow.Cells("DEPARTMENT").Value
            BunifuTextBox17.Text = SelectedRow.Cells("REPORTING MANAGER").Value
            BunifuButton24.Visible = True
            BunifuButton24.Enabled = True
            If SelectedRow.Cells("ISSUING").Value = True Then
                BunifuToggleSwitch3.Checked = True
            Else
                BunifuToggleSwitch3.Checked = False
            End If
            If SelectedRow.Cells("EVENT ADDING").Value = True Then
                BunifuToggleSwitch4.Checked = True
            Else
                BunifuToggleSwitch4.Checked = False
            End If
            If SelectedRow.Cells("ADMINISTRATIVE").Value = True Then
                BunifuToggleSwitch5.Checked = True
            Else
                BunifuToggleSwitch5.Checked = False
            End If
            BunifuPanel2.Visible = True
            BunifuPanel2.Enabled = True
        Else
            BunifuTextBox8.Clear()
            BunifuTextBox9.Clear()
            BunifuTextBox10.Clear()
            BunifuTextBox11.Clear()
            BunifuTextBox13.Clear()
            BunifuTextBox15.Clear()
            BunifuTextBox16.Clear()
            BunifuTextBox17.Clear()
            BunifuToggleSwitch3.Checked = False
            BunifuToggleSwitch4.Checked = False
            BunifuToggleSwitch5.Checked = False
            BunifuPanel2.Visible = False
            BunifuPanel2.Enabled = False
            BunifuButton24.Visible = False
            BunifuButton24.Enabled = False
        End If
    End Sub
    Private Sub BunifuButton24_Click(sender As Object, e As EventArgs) Handles BunifuButton24.Click
        Dim NT_ID As String = BunifuTextBox8.Text
        Dim Buisness_Area As String = BunifuTextBox15.Text
        Dim Functions As String = BunifuTextBox16.Text
        Dim Leave_Approver As String = BunifuTextBox17.Text
        If BunifuToggleSwitch3.Checked = True Then
            Issue_Gift = True
        Else
            Issue_Gift = False
        End If
        If BunifuToggleSwitch4.Checked = True Then
            Events_Addition = True
        Else
            Events_Addition = False
        End If
        If BunifuToggleSwitch5.Checked = True Then
            Admn_Access = True
        Else
            Admn_Access = False
        End If
        Dim UpdateEmployee As String = FuncLib.Setup.UpdateEmployeeData(NT_ID, Buisness_Area, Functions, Leave_Approver, Issue_Gift, Events_Addition, Admn_Access)
        If UpdateEmployee = "True" Then
            BunifuSnackbar1.Show(Me, $"Update Successfull For {BunifuTextBox11.Text}", Bunifu.UI.WinForms.BunifuSnackbar.MessageTypes.Success, 2000)
            Dim result = FuncLib.Setup.ExportDataFromUserDB()
            If result.Item1 IsNot Nothing Then
                Dim grid = result.Item1
                grid.Columns("ID").ColumnName = "SR NO."
                grid.Columns("NT_ID").ColumnName = "NT ID"
                grid.Columns("Scanner_ID").ColumnName = "SCAN ID"
                grid.Columns("Scanner_5D").ColumnName = "SCAN 5D"
                grid.Columns("Employee_Name").ColumnName = "NAME"
                grid.Columns("Gender").ColumnName = "GENDER"
                grid.Columns("Buisness_Area").ColumnName = "BUISNESS"
                grid.Columns("Functions").ColumnName = "DEPARTMENT"
                grid.Columns("Leave_Approver").ColumnName = "REPORTING MANAGER"
                grid.Columns("Modified_By").ColumnName = "MODIFIED BY"
                grid.Columns("Modified_At").ColumnName = "MODIFIED AT"
                grid.Columns("Issue_Gift").ColumnName = "ISSUING"
                grid.Columns("Events_Addition").ColumnName = "EVENT ADDING"
                grid.Columns("Admn_Access").ColumnName = "ADMINISTRATIVE"
                BunifuDataGridView2.DataSource = grid
                BunifuDropdown16.Items.Clear()
                BunifuDropdown16.Text = Nothing
                For Each column As DataGridViewColumn In BunifuDataGridView2.Columns
                    BunifuDropdown16.Items.Add(column.HeaderText)
                Next
            End If
        Else
            BunifuSnackbar1.Show(Me, $"Error Occured While Updating The Data For {BunifuTextBox11.Text}.{Environment.NewLine}Error:{UpdateEmployee}", Bunifu.UI.WinForms.BunifuSnackbar.MessageTypes.Error, 2000)
            FuncLib.WriteLog.WriteErrorLog($"Error Occured While Executing FuncLib.Setup.UpdateEmployeeData:{UpdateEmployee}")
            Exit Sub
        End If
    End Sub
    Private Sub BunifuButton4_Click(sender As Object, e As EventArgs) Handles BunifuButton4.Click
        BunifuLabel26.Visible = False
        BunifuLabel26.Enabled = False
        BunifuDropdown7.Visible = False
        BunifuDropdown7.Enabled = False
        BunifuButton15.Visible = False
        BunifuShadowPanel7.Enabled = False
        BunifuShadowPanel7.Visible = False
        BunifuDropdown6.Items.Clear()
        BunifuDropdown6.Text = Nothing
        BunifuDropdown6.Items.Add("ADD TO EVENTS FROM INTERNAL")
        BunifuDropdown6.Items.Add("ADD TO EVENTS FROM EXTERNAL")
        BunifuDropdown6.Items.Add("REMOVE FROM EVENTS")
        BunifuDropdown6.Items.Add("DELETE THIS EVENTS")
        GroupBox3.Visible = False
        GroupBox3.Enabled = False
        BunifuTextBox5.Clear()
        GroupBox13.Visible = True
        GroupBox13.Enabled = True
        BunifuDropdown12.Items.Clear()
        BunifuDropdown12.Text = Nothing
        GroupBox14.Visible = False
        GroupBox14.Enabled = False
        BunifuDropdown13.Items.Clear()
        BunifuDropdown13.Text = Nothing
        BunifuDataGridView3.DataSource = Nothing
        BunifuDataGridView3.Visible = False
        BunifuDataGridView3.Enabled = False
        BunifuButton7.Visible = False
        BunifuButton7.Enabled = False
        Dim Occasion As List(Of String) = FuncLib.Gifts.GetOcaasioList()
        If Occasion.Count <= 0 AndAlso TypeOf Occasion(0) Is String Then
            Dim errorMessage As String = Occasion(0)
            BunifuSnackbar1.Show(Me, $"Error While Fetching Occasion List From. {Environment.NewLine} Error : {errorMessage}", Bunifu.UI.WinForms.BunifuSnackbar.MessageTypes.Warning, 3000)
            FuncLib.WriteLog.WriteErrorLog($"Error occured While Trying To Fetch Occasion List : {errorMessage}")
            Exit Sub
        Else
            GroupBox13.Visible = True
            GroupBox13.Enabled = True
            BunifuDropdown12.Items.Clear()
            BunifuDropdown12.Text = Nothing
            BunifuDropdown12.Items.AddRange(Occasion.ToArray())
            If BunifuPages1.Visible = False Then
                BunifuPages1.TabPages.Clear()
                transition.ShowSync(BunifuPages1, False, BunifuAnimatorNS.Animation.ScaleAndHorizSlide)
                BunifuPages1.TabPages.Add(TabPage5)
            Else
                transition.HideSync(BunifuPages1, False, BunifuAnimatorNS.Animation.ScaleAndHorizSlide)
                BunifuPages1.TabPages.Clear()
                transition.ShowSync(BunifuPages1, False, BunifuAnimatorNS.Animation.ScaleAndHorizSlide)
                BunifuPages1.TabPages.Add(TabPage5)
            End If
        End If
    End Sub
    Private Sub BunifuDropdown12_SelectedIndexChanged(sender As Object, e As EventArgs) Handles BunifuDropdown12.SelectedIndexChanged
        Dim Occasiontype As String = FuncLib.Gifts.CheckTypeOfOcassion(BunifuDropdown12.SelectedItem.ToString)
        If Occasiontype = "Constant" Then
            GroupBox13.Enabled = False
            GroupBox14.Visible = False
            GroupBox14.Enabled = False
            BunifuShadowPanel7.Visible = True
            BunifuShadowPanel7.Enabled = True
            BunifuDropdown6.Items.Clear()
            BunifuDropdown6.Text = Nothing
            BunifuButton15.Visible = True
            BunifuDropdown6.Items.Add("ADD TO EVENT")
            BunifuDropdown6.Items.Add("REMOVE FROM EVENTS")
            BunifuDropdown6.Items.Add("DELETE THIS EVENTS")
        ElseIf Occasiontype = "Variable" Then
            BunifuShadowPanel7.Visible = False
            BunifuShadowPanel7.Enabled = False
            GroupBox13.Enabled = False
            GroupBox14.Visible = True
            GroupBox14.Enabled = True
            BunifuDropdown13.Items.Clear()
            BunifuDropdown13.Text = Nothing
            BunifuButton15.Visible = True
            Dim OccasionDate As List(Of Date) = FuncLib.Gifts.VarableDateQuery(BunifuDropdown12.SelectedItem)
            If OccasionDate.Count <= 0 AndAlso OccasionDate(0) = DateTime.MinValue Then
                BunifuSnackbar1.Show(Me, $"Error Occured While Fetching Occasion Dates!", Bunifu.UI.WinForms.BunifuSnackbar.MessageTypes.Error, 3000)
                FuncLib.WriteLog.WriteErrorLog($"Error occured While Fetching The Occasion Dates")
                Exit Sub
            Else
                For Each DateDetails In OccasionDate
                    BunifuDropdown13.Items.Add(DateDetails.ToString("MM/dd/yyyy"))
                Next
            End If
        ElseIf Occasiontype = "Undefined Type" Then
            BunifuSnackbar1.Show(Me, $"No OccasionType Mentioned For {BunifuDropdown12.SelectedItem}.{Environment.NewLine} Please Contact To Administrator or Developer!", Bunifu.UI.WinForms.BunifuSnackbar.MessageTypes.Warning, 3000)
            FuncLib.WriteLog.WriteErrorLog($"{BunifuDropdown12.SelectedItem.ToString} Has Been Found Without Any OcassionType")
            Exit Sub
        Else
            BunifuSnackbar1.Show(Me, $"Error Occured While Reading The Type Of The Occasion {Environment.NewLine} Error : {Occasiontype}", Bunifu.UI.WinForms.BunifuSnackbar.MessageTypes.Error, 3000)
            FuncLib.WriteLog.WriteErrorLog($"Error While Trying To Read Type of Occasion : {Occasiontype}")
            Exit Sub
        End If

    End Sub
    Private Sub BunifuDropdown13_SelectedIndexChanged(sender As Object, e As EventArgs) Handles BunifuDropdown13.SelectedIndexChanged
        BunifuShadowPanel7.Visible = True
        BunifuShadowPanel7.Enabled = True
        BunifuDropdown6.Items.Clear()
        BunifuDropdown6.Text = Nothing
        BunifuDropdown6.Items.Add("ADD TO EVENT")
        BunifuDropdown6.Items.Add("REMOVE FROM EVENTS")
        BunifuDropdown6.Items.Add("DELETE THIS EVENTS")
        GroupBox14.Enabled = False
    End Sub
    Private Sub BunifuButton23_Click(sender As Object, e As EventArgs) Handles BunifuButton23.Click
        transition.HideSync(Me, False, BunifuAnimatorNS.Animation.ScaleAndHorizSlide)
        Me.Close()
    End Sub
    Private Sub BunifuDropdown6_SelectedIndexChanged(sender As Object, e As EventArgs) Handles BunifuDropdown6.SelectedIndexChanged
        If BunifuDropdown6.SelectedItem = "DELETE THIS EVENTS" Then
            If BunifuDropdown12.SelectedItem = "BirthDay" Then
                BunifuSnackbar1.Show(Me, $"BirtDay Is An Internal Function CanNot Be Deleted!", Bunifu.UI.WinForms.BunifuSnackbar.MessageTypes.Warning, 3000)
                GroupBox3.Visible = False
                BunifuDataGridView3.DataSource = Nothing
                BunifuDataGridView3.Visible = False
                BunifuDataGridView3.Enabled = False
                BunifuButton7.Visible = False
                BunifuButton7.Enabled = False
                BunifuLabel26.Visible = False
                BunifuLabel26.Enabled = False
                BunifuDropdown7.Visible = False
                BunifuDropdown7.Enabled = False
                Exit Sub
            Else
                GroupBox3.Visible = False
                BunifuDataGridView3.DataSource = Nothing
                BunifuDataGridView3.Visible = True
                BunifuDataGridView3.Enabled = True
                BunifuLabel26.Visible = False
                BunifuLabel26.Enabled = False
                BunifuDropdown7.Visible = False
                BunifuDropdown7.Enabled = False
                Dim result = FuncLib.Events.FetchAllDataForConstantFunctionDeletion(BunifuDropdown12.SelectedItem)
                If result IsNot Nothing AndAlso result.Item1 IsNot Nothing Then
                    Dim dataTable = result.Item1
                    If dataTable.Rows.Count > 0 Then
                        dataTable.Columns("Employee_ID").ColumnName = "EMP ID"
                        dataTable.Columns("NT_ID").ColumnName = "NT ID"
                        dataTable.Columns("Scanner_ID").ColumnName = "SCAN ID"
                        dataTable.Columns("Scanner_5D").ColumnName = "SCAN 5D"
                        dataTable.Columns("Employee_Name").ColumnName = "NAME"
                        dataTable.Columns("Gender").ColumnName = "GENDER"
                        dataTable.Columns("Celebration_Date").ColumnName = "FUNCTION DATE"
                        dataTable.Columns("Buisness_Area").ColumnName = "BUISNESS"
                        dataTable.Columns("Functions").ColumnName = "DEPARTMENT"
                        dataTable.Columns("Leave_Approver").ColumnName = "REPORTING MANAGER"
                        BunifuDataGridView3.DataSource = dataTable
                        BunifuButton7.Visible = True
                        BunifuButton7.Enabled = True
                        BunifuButton7.Text = "DELETE"
                    Else
                        dataTable.Columns("Employee_ID").ColumnName = "EMP ID"
                        dataTable.Columns("NT_ID").ColumnName = "NT ID"
                        dataTable.Columns("Scanner_ID").ColumnName = "SCAN ID"
                        dataTable.Columns("Scanner_5D").ColumnName = "SCAN 5D"
                        dataTable.Columns("Employee_Name").ColumnName = "NAME"
                        dataTable.Columns("Gender").ColumnName = "GENDER"
                        dataTable.Columns("Celebration_Date").ColumnName = "FUNCTION DATE"
                        dataTable.Columns("Buisness_Area").ColumnName = "BUISNESS"
                        dataTable.Columns("Functions").ColumnName = "DEPARTMENT"
                        dataTable.Columns("Leave_Approver").ColumnName = "REPORTING MANAGER"
                        BunifuDataGridView3.DataSource = dataTable
                        BunifuButton7.Visible = True
                        BunifuButton7.Enabled = True
                        BunifuButton7.Text = "DELETE"
                    End If
                Else
                    BunifuSnackbar1.Show(Me, $"Error Occured While Reading Data From Database {Environment.NewLine} Error : {result.Item2}", Bunifu.UI.WinForms.BunifuSnackbar.MessageTypes.Error, 3000)
                    FuncLib.WriteLog.WriteErrorLog($"Error While Executng FuncLib.Events.FetchAllDataForConstantFunctionDeletion: {result.Item2}")
                    GroupBox3.Visible = False
                    BunifuDataGridView3.DataSource = Nothing
                    BunifuDataGridView3.Visible = False
                    BunifuDataGridView3.Enabled = False
                    BunifuButton7.Visible = False
                    BunifuButton7.Enabled = False
                    BunifuLabel26.Visible = False
                    BunifuLabel26.Enabled = False
                    BunifuDropdown7.Visible = False
                    BunifuDropdown7.Enabled = False
                    Exit Sub
                End If
            End If
        ElseIf BunifuDropdown6.SelectedItem = "ADD TO EVENT" Then
            BunifuLabel26.Visible = False
            BunifuLabel26.Enabled = False
            BunifuDropdown7.Visible = False
            BunifuDropdown7.Enabled = False
            GroupBox3.Visible = True
            GroupBox3.Enabled = True
            BunifuTextBox5.Clear()
            BunifuDataGridView3.DataSource = Nothing
            BunifuDataGridView3.Visible = False
            BunifuDataGridView3.Enabled = False
            BunifuButton7.Visible = False
            BunifuButton7.Enabled = False
            BunifuTextBox5.ReadOnly = True
        ElseIf BunifuDropdown6.SelectedItem = "REMOVE FROM EVENTS" Then
            BunifuLabel26.Visible = False
            BunifuLabel26.Enabled = False
            BunifuDropdown7.Visible = False
            BunifuDropdown7.Enabled = False
            GroupBox3.Visible = False
            GroupBox3.Enabled = False
            BunifuDataGridView3.DataSource = Nothing
            BunifuDataGridView3.Visible = False
            BunifuDataGridView3.Enabled = False
            BunifuButton7.Visible = False
            BunifuButton7.Enabled = False
            BunifuTextBox5.Clear()
            BunifuDropdown7.Items.Clear()
            BunifuDropdown7.Text = Nothing
            BunifuDropdown7.Items.Add("REMOVE THIS DATE")
            BunifuDropdown7.Items.Add("REMOVE EMPLOYEES")
            If BunifuDropdown13.SelectedItem IsNot Nothing Then
                BunifuLabel26.Visible = True
                BunifuLabel26.Enabled = True
                BunifuDropdown7.Visible = True
                BunifuDropdown7.Enabled = True
                BunifuDropdown7.Items.Clear()
                BunifuDropdown7.Text = Nothing
                BunifuDropdown7.Items.Add("REMOVE THIS DATE")
                BunifuDropdown7.Items.Add("REMOVE EMPLOYEES")
            ElseIf BunifuDropdown13.SelectedItem Is Nothing Then
                BunifuLabel26.Visible = True
                BunifuLabel26.Enabled = True
                BunifuDropdown7.Visible = True
                BunifuDropdown7.Enabled = False
                BunifuDropdown7.Items.Clear()
                BunifuDropdown7.Text = Nothing
                BunifuDropdown7.Items.Clear()
                BunifuDropdown7.Text = Nothing
                BunifuDropdown7.Items.Add("REMOVE THIS DATE")
                BunifuDropdown7.Items.Add("REMOVE EMPLOYEES")
                BunifuDropdown7.SelectedItem = BunifuDropdown7.Items(0)
                Dim result = FuncLib.Events.FetchAllDataForConstantFunctionDeletion(BunifuDropdown12.SelectedItem)
                If result IsNot Nothing AndAlso result.Item1 IsNot Nothing Then
                    Dim dataTable = result.Item1
                    If dataTable.Rows.Count > 0 Then
                        dataTable.Columns("Employee_ID").ColumnName = "EMP ID"
                        dataTable.Columns("NT_ID").ColumnName = "NT ID"
                        dataTable.Columns("Scanner_ID").ColumnName = "SCAN ID"
                        dataTable.Columns("Scanner_5D").ColumnName = "SCAN 5D"
                        dataTable.Columns("Employee_Name").ColumnName = "NAME"
                        dataTable.Columns("Gender").ColumnName = "GENDER"
                        dataTable.Columns("Celebration_Date").ColumnName = "FUNCTION DATE"
                        dataTable.Columns("Buisness_Area").ColumnName = "BUISNESS"
                        dataTable.Columns("Functions").ColumnName = "DEPARTMENT"
                        dataTable.Columns("Leave_Approver").ColumnName = "REPORTING MANAGER"
                        BunifuDataGridView3.DataSource = dataTable
                        BunifuButton7.Visible = True
                        BunifuButton7.Enabled = True
                        BunifuButton7.Text = "REMOVE"
                        BunifuDataGridView3.Visible = True
                        BunifuDataGridView3.Enabled = True
                    Else
                        dataTable.Columns("Employee_ID").ColumnName = "EMP ID"
                        dataTable.Columns("NT_ID").ColumnName = "NT ID"
                        dataTable.Columns("Scanner_ID").ColumnName = "SCAN ID"
                        dataTable.Columns("Scanner_5D").ColumnName = "SCAN 5D"
                        dataTable.Columns("Employee_Name").ColumnName = "NAME"
                        dataTable.Columns("Gender").ColumnName = "GENDER"
                        dataTable.Columns("Celebration_Date").ColumnName = "FUNCTION DATE"
                        dataTable.Columns("Buisness_Area").ColumnName = "BUISNESS"
                        dataTable.Columns("Functions").ColumnName = "DEPARTMENT"
                        dataTable.Columns("Leave_Approver").ColumnName = "REPORTING MANAGER"
                        BunifuDataGridView3.DataSource = dataTable
                        BunifuButton7.Visible = True
                        BunifuButton7.Enabled = True
                        BunifuButton7.Text = "REMOVE"
                        BunifuDataGridView3.Visible = True
                        BunifuDataGridView3.Enabled = True
                    End If
                Else
                    BunifuSnackbar1.Show(Me, $"Error Occured While Reading Data From Database {Environment.NewLine} Error : {result.Item2}", Bunifu.UI.WinForms.BunifuSnackbar.MessageTypes.Error, 3000)
                    FuncLib.WriteLog.WriteErrorLog($"Error While Executng FuncLib.Events.FetchAllDataForConstantFunctionDeletion: {result.Item2}")
                    GroupBox3.Visible = False
                    BunifuDataGridView3.DataSource = Nothing
                    BunifuDataGridView3.Visible = False
                    BunifuDataGridView3.Enabled = False
                    BunifuButton7.Visible = False
                    BunifuButton7.Enabled = False
                    Exit Sub
                End If
            End If
        End If
    End Sub
    Private Sub BunifuButton7_Click(sender As Object, e As EventArgs) Handles BunifuButton7.Click
        If BunifuDropdown6.SelectedItem = "DELETE THIS EVENTS" Then
            Dim Deleteevents = FuncLib.Events.DeleteDataTableForConstantFunction(BunifuDropdown12.SelectedItem)
            If Deleteevents IsNot Nothing Then
                BunifuSnackbar1.Show(Me, $"Error Occured While Deleting The Datatable {Environment.NewLine} Error : {Deleteevents}", Bunifu.UI.WinForms.BunifuSnackbar.MessageTypes.Error, 3000)
                FuncLib.WriteLog.WriteErrorLog($"Error While ExecutngFuncLib.Events.DeleteDataTableForConstantFunction: {Deleteevents}")
                Exit Sub
            Else
                BunifuSnackbar1.Show(Me, $"{BunifuDropdown12.SelectedItem} Has been Removed Successfully", Bunifu.UI.WinForms.BunifuSnackbar.MessageTypes.Success, 3000)
                clearfunction()
            End If
        ElseIf BunifuDropdown6.SelectedItem = "ADD TO EVENT" Then
            Dim AddFunction As String = FuncLib.Events.AddEmployeeToConstantDatabase(BunifuDropdown12.SelectedItem, BunifuDataGridView3)
            If AddFunction = "True" Then
                BunifuSnackbar1.Show(Me, $"All Verified Employees Added To Function", Bunifu.UI.WinForms.BunifuSnackbar.MessageTypes.Success, 3000)
                clearfunction()
            Else
                BunifuSnackbar1.Show(Me, $"Error Occured While Adding Datatable {Environment.NewLine} Error : {AddFunction}", Bunifu.UI.WinForms.BunifuSnackbar.MessageTypes.Error, 3000)
                FuncLib.WriteLog.WriteErrorLog($"Error While Executng  FuncLib.Events.AddEmployeeToConstantDatabase {AddFunction}")
                Exit Sub
            End If
        ElseIf BunifuDropdown6.SelectedItem = "REMOVE FROM EVENTS" Then
            If BunifuDropdown13.SelectedItem Is Nothing AndAlso BunifuDropdown7.SelectedItem = "REMOVE EMPLOYEES" Then
                Dim RemoveEmployee As String
                If BunifuDataGridView3.SelectedRows.Count > 0 Then
                    Dim Count As Integer = 0
                    For Each row As DataGridViewRow In BunifuDataGridView3.SelectedRows
                        Dim Emp_ID As String = row.Cells("EMP ID").Value
                        Dim Scanner_ID As String = row.Cells("SCAN ID").Value
                        Dim CreationDate As Date = row.Cells("Created_At").Value
                        RemoveEmployee = FuncLib.Events.RemoveSelectedEmployeeFromConstantDB(BunifuDropdown12.SelectedItem, Emp_ID, Scanner_ID, CreationDate)
                        If RemoveEmployee <> "True" Then
                            Exit For
                        Else
                            Count += 1
                        End If
                    Next
                    If RemoveEmployee = "True" Then
                        BunifuSnackbar1.Show(Me, $"{Count} Has been Removed From {BunifuDropdown12.SelectedItem} Function", Bunifu.UI.WinForms.BunifuSnackbar.MessageTypes.Success, 3000)
                        clearfunction()
                        Exit Sub
                    Else
                        BunifuSnackbar1.Show(Me, $"Error Occured While Removing Data From DataBase.{Environment.NewLine} Error: {RemoveEmployee}", Bunifu.UI.WinForms.BunifuSnackbar.MessageTypes.Error, 3000)
                        FuncLib.WriteLog.WriteErrorLog($"Error Occured While Executing FuncLib.Events.RemoveSelectedEmployeeFromConstantDB :{RemoveEmployee}")
                        Exit Sub
                    End If
                End If
            ElseIf BunifuDropdown13.SelectedItem IsNot Nothing AndAlso BunifuDropdown7.SelectedItem = "REMOVE EMPLOYEES" Then
                Dim RemoveEmployee As String
                If BunifuDataGridView3.SelectedRows.Count > 0 Then
                    Dim Count As Integer = 0
                    For Each row As DataGridViewRow In BunifuDataGridView3.SelectedRows
                        Dim Emp_ID As String = row.Cells("EMP ID").Value
                        Dim Scanner_ID As String = row.Cells("SCAN ID").Value
                        Dim CreationDate As Date = row.Cells("Created_At").Value
                        Dim CelebrationDate As Date = row.Cells("FUNCTION DATE").Value
                        RemoveEmployee = FuncLib.Events.RemoveSelectedEmployeeFromVariableFunction(BunifuDropdown12.SelectedItem, Emp_ID, Scanner_ID, CreationDate, CelebrationDate)
                        If RemoveEmployee <> "True" Then
                            Exit For
                        Else
                            Count += 1
                        End If
                    Next
                    If RemoveEmployee = "True" Then
                        BunifuSnackbar1.Show(Me, $"{Count} Has been Removed From {BunifuDropdown12.SelectedItem} Function", Bunifu.UI.WinForms.BunifuSnackbar.MessageTypes.Success, 3000)
                        clearfunction()
                        Exit Sub
                    Else
                        BunifuSnackbar1.Show(Me, $"Error Occured While Removing Data From DataBase.{Environment.NewLine} Error: {RemoveEmployee}", Bunifu.UI.WinForms.BunifuSnackbar.MessageTypes.Error, 3000)
                        FuncLib.WriteLog.WriteErrorLog($"Error Occured While Executing FuncLib.Events.RemoveSelectedEmployeeFromConstantDB :{RemoveEmployee}")
                        Exit Sub
                    End If
                End If
            ElseIf BunifuDropdown13.SelectedItem IsNot Nothing AndAlso BunifuDropdown7.SelectedItem = "REMOVE THIS DATE" Then
                If BunifuDropdown13.Items.Count > 1 Then
                    Dim RemoveDate As String = FuncLib.Events.DeletevariableFunctionDate(BunifuDropdown12.SelectedItem, BunifuDropdown13.SelectedItem)
                    If RemoveDate = "True" Then
                        BunifuSnackbar1.Show(Me, $"{BunifuDropdown13.SelectedItem} has been removed from {BunifuDropdown12.SelectedItem}", Bunifu.UI.WinForms.BunifuSnackbar.MessageTypes.Success, 3000)
                        clearfunction()
                        Exit Sub
                    Else
                        BunifuSnackbar1.Show(Me, $"Error Occured While Removing Data From DataBase.{Environment.NewLine} Error: {RemoveDate}", Bunifu.UI.WinForms.BunifuSnackbar.MessageTypes.Error, 3000)
                        FuncLib.WriteLog.WriteErrorLog($"Error Occured While Executing FuncLib.Events.DeletevariableFunctionDate :{RemoveDate}")
                        Exit Sub
                    End If
                Else
                    BunifuSnackbar1.Show(Me, $"{BunifuDropdown12.SelectedItem} Contains Only One Date: {BunifuDropdown13.SelectedItem}", Bunifu.UI.WinForms.BunifuSnackbar.MessageTypes.Error, 3000)
                    Dim iret As Object = MsgBox($"{BunifuDropdown12.SelectedItem} Contains Only One Date : {BunifuDropdown13.SelectedItem}.{Environment.NewLine} Do You Want To Delete This Function?", vbQuestion + vbYesNo, "DeletionConfrimation()")
                    If iret = vbYes Then
                        Dim Deleteevents = FuncLib.Events.DeleteDataTableForConstantFunction(BunifuDropdown12.SelectedItem)
                        If Deleteevents IsNot Nothing Then
                            BunifuSnackbar1.Show(Me, $"Error Occured While Deleting The Datatable {Environment.NewLine} Error : {Deleteevents}", Bunifu.UI.WinForms.BunifuSnackbar.MessageTypes.Error, 3000)
                            FuncLib.WriteLog.WriteErrorLog($"Error While ExecutngFuncLib.Events.DeleteDataTableForConstantFunction: {Deleteevents}")
                            Exit Sub
                        Else
                            BunifuSnackbar1.Show(Me, $"{BunifuDropdown12.SelectedItem} Has been Removed Successfully", Bunifu.UI.WinForms.BunifuSnackbar.MessageTypes.Success, 3000)
                            clearfunction()
                        End If
                    ElseIf iret = vbNo Then
                        'donothing
                    End If
                End If
            End If
        End If
    End Sub
    Public Sub clearfunction()
        BunifuButton7.Visible = False
        BunifuButton7.Enabled = False
        BunifuDataGridView3.DataSource = Nothing
        BunifuDataGridView3.Visible = False
        BunifuDataGridView3.Enabled = False
        GroupBox3.Visible = False
        BunifuTextBox5.Clear()
        BunifuShadowPanel7.Visible = False
        BunifuShadowPanel7.Enabled = False
        BunifuDropdown6.Items.Clear()
        BunifuDropdown6.Text = Nothing
        GroupBox14.Visible = False
        BunifuDropdown13.Items.Clear()
        BunifuDropdown13.Text = Nothing
        GroupBox13.Visible = True
        GroupBox13.Enabled = True
        BunifuDropdown12.Items.Clear()
        BunifuDropdown12.Text = Nothing
        Dim Occasion As List(Of String) = FuncLib.Gifts.GetOcaasioList()
        If Occasion.Count <= 0 AndAlso TypeOf Occasion(0) Is String Then
            Dim errorMessage As String = Occasion(0)
            BunifuSnackbar1.Show(Me, $"Error While Fetching Occasion List From. {Environment.NewLine} Error : {errorMessage}", Bunifu.UI.WinForms.BunifuSnackbar.MessageTypes.Warning, 3000)
            FuncLib.WriteLog.WriteErrorLog($"Error occured While Trying To Fetch Occasion List : {errorMessage}")
            Exit Sub
        Else
            BunifuDropdown12.Items.AddRange(Occasion.ToArray())
        End If
        BunifuButton15.Visible = False
        BunifuLabel26.Visible = False
        BunifuLabel26.Enabled = False
        BunifuDropdown7.Visible = False
        BunifuDropdown7.Enabled = False
    End Sub
    Private Sub BunifuButton12_Click(sender As Object, e As EventArgs) Handles BunifuButton12.Click
        If BunifuDropdown12.SelectedItem = "BirthDay" Then
            requiredColumns.Clear()
            requiredColumns.Add("Employee_ID")
            requiredColumns.Add("NT_ID")
            requiredColumns.Add("Scanner_ID")
            requiredColumns.Add("Employee_Name")
            requiredColumns.Add("Gender")
            requiredColumns.Add("Birth_Date")
            requiredColumns.Add("Buisness_Area")
            requiredColumns.Add("Functions")
            requiredColumns.Add("Leave_Approver")
            Dim openFileDialog As New OpenFileDialog()
            openFileDialog.Filter = "Excel Files|*.xls;*.xlsx;*.xlsm"
            openFileDialog.Title = "Select Database File"
            openFileDialog.Multiselect = False
            If openFileDialog.ShowDialog() = DialogResult.OK Then
                Dim selectedFilePath As String = openFileDialog.FileName
                Dim VerifyExcel As Boolean = FuncLib.Events.VerifyExcel(selectedFilePath, requiredColumns)
                If VerifyExcel = True Then
                    BunifuSnackbar1.Show(Me, $"The File {selectedFilePath} has verified successfully", Bunifu.UI.WinForms.BunifuSnackbar.MessageTypes.Success, 2000)
                    BunifuTextBox5.Text = selectedFilePath
                    Dim result = FuncLib.Events.ImplementDataToGridforconstant(selectedFilePath)
                    If result.Item1 IsNot Nothing Then
                        Dim grid = result.Item1
                        grid.Columns("Employee_ID").ColumnName = "EMP ID"
                        grid.Columns("NT_ID").ColumnName = "NT ID"
                        grid.Columns("Scanner_ID").ColumnName = "SCAN ID"
                        grid.Columns("Scanner_5D").ColumnName = "SCAN 5D"
                        grid.Columns("Employee_Name").ColumnName = "NAME"
                        grid.Columns("Gender").ColumnName = "GENDER"
                        grid.Columns("Celebration_Date").ColumnName = "DATE"
                        grid.Columns("Buisness_Area").ColumnName = "BUISNESS"
                        grid.Columns("Functions").ColumnName = "DEPARTMENT"
                        grid.Columns("Leave_Approver").ColumnName = "REPORTING MANAGER"
                        grid.Columns("Created_At").ColumnName = "TIME STAMP"
                        BunifuDataGridView3.DataSource = grid
                        BunifuButton7.Visible = True
                        BunifuButton7.Enabled = True
                        BunifuButton7.Text = "UPDATE"
                        BunifuDataGridView3.Visible = True
                        BunifuDataGridView3.Enabled = True
                    End If
                Else
                    BunifuSnackbar1.Show(Me, $"File {selectedFilePath} has failed in Fileverification method!", Bunifu.UI.WinForms.BunifuSnackbar.MessageTypes.Error, 2000)
                    BunifuDataGridView3.DataSource = Nothing
                    BunifuDataGridView3.Visible = False
                    BunifuDataGridView3.Enabled = False
                    BunifuButton7.Visible = False
                    BunifuButton7.Enabled = False
                    BunifuTextBox5.Clear()
                    Exit Sub
                End If
            End If
        ElseIf BunifuDropdown12.SelectedItem IsNot Nothing AndAlso GroupBox14.Visible = False Then
            Ini.Load(Inipath)
            requiredColumns.Clear()
            requiredColumns.Add("Employee_ID")
            requiredColumns.Add("NT_ID")
            requiredColumns.Add("Scanner_ID")
            requiredColumns.Add("Employee_Name")
            requiredColumns.Add("Gender")
            requiredColumns.Add("Buisness_Area")
            requiredColumns.Add("Functions")
            requiredColumns.Add("Leave_Approver")
            Dim openFileDialog As New OpenFileDialog()
            openFileDialog.Filter = "Excel Files|*.xls;*.xlsx;*.xlsm"
            openFileDialog.Title = "Select Database File"
            openFileDialog.Multiselect = False
            If openFileDialog.ShowDialog() = DialogResult.OK Then
                Dim selectedFilePath As String = openFileDialog.FileName
                Dim VerifyExcel As Boolean = FuncLib.Events.VerifyExcel(selectedFilePath, requiredColumns)
                If VerifyExcel = True Then
                    BunifuSnackbar1.Show(Me, $"The File {selectedFilePath} has verified successfully", Bunifu.UI.WinForms.BunifuSnackbar.MessageTypes.Success, 2000)
                    BunifuTextBox5.Text = selectedFilePath
                    Dim result = FuncLib.Events.ImplementDataToGridforVariable(selectedFilePath, Ini.Sections("DATE").Keys(BunifuDropdown12.SelectedItem.ToString).Value)
                    If result.Item1 IsNot Nothing Then
                        Dim grid = result.Item1
                        grid.Columns("Employee_ID").ColumnName = "EMP ID"
                        grid.Columns("NT_ID").ColumnName = "NT ID"
                        grid.Columns("Scanner_ID").ColumnName = "SCAN ID"
                        grid.Columns("Scanner_5D").ColumnName = "SCAN 5D"
                        grid.Columns("Employee_Name").ColumnName = "NAME"
                        grid.Columns("Gender").ColumnName = "GENDER"
                        grid.Columns("Celebration_Date").ColumnName = "DATE"
                        grid.Columns("Buisness_Area").ColumnName = "BUISNESS"
                        grid.Columns("Functions").ColumnName = "DEPARTMENT"
                        grid.Columns("Leave_Approver").ColumnName = "REPORTING MANAGER"
                        grid.Columns("Created_At").ColumnName = "TIME STAMP"
                        BunifuDataGridView3.DataSource = grid
                        BunifuButton7.Visible = True
                        BunifuButton7.Enabled = True
                        BunifuButton7.Text = "UPDATE"
                        BunifuDataGridView3.Visible = True
                        BunifuDataGridView3.Enabled = True
                    End If
                Else
                    BunifuSnackbar1.Show(Me, $"File {selectedFilePath} has failed in Fileverification method!", Bunifu.UI.WinForms.BunifuSnackbar.MessageTypes.Error, 2000)
                    BunifuDataGridView3.DataSource = Nothing
                    BunifuDataGridView3.Visible = False
                    BunifuDataGridView3.Enabled = False
                    BunifuButton7.Visible = False
                    BunifuButton7.Enabled = False
                    BunifuTextBox5.Clear()
                    Exit Sub
                End If
            End If
        ElseIf BunifuDropdown12.SelectedItem IsNot Nothing AndAlso GroupBox14.Visible = True Then
            requiredColumns.Clear()
            requiredColumns.Add("Employee_ID")
            requiredColumns.Add("NT_ID")
            requiredColumns.Add("Scanner_ID")
            requiredColumns.Add("Employee_Name")
            requiredColumns.Add("Gender")
            requiredColumns.Add("Buisness_Area")
            requiredColumns.Add("Functions")
            requiredColumns.Add("Leave_Approver")
            Dim openFileDialog As New OpenFileDialog()
            openFileDialog.Filter = "Excel Files|*.xls;*.xlsx;*.xlsm"
            openFileDialog.Title = "Select Database File"
            openFileDialog.Multiselect = False
            If openFileDialog.ShowDialog() = DialogResult.OK Then
                Dim selectedFilePath As String = openFileDialog.FileName
                Dim VerifyExcel As Boolean = FuncLib.Events.VerifyExcel(selectedFilePath, requiredColumns)
                If VerifyExcel = True Then
                    BunifuSnackbar1.Show(Me, $"The File {selectedFilePath} has verified successfully", Bunifu.UI.WinForms.BunifuSnackbar.MessageTypes.Success, 2000)
                    BunifuTextBox5.Text = selectedFilePath
                    Dim result = FuncLib.Events.ImplementDataToGridforVariable(selectedFilePath, BunifuDropdown13.SelectedItem)
                    If result.Item1 IsNot Nothing Then
                        Dim grid = result.Item1
                        grid.Columns("Employee_ID").ColumnName = "EMP ID"
                        grid.Columns("NT_ID").ColumnName = "NT ID"
                        grid.Columns("Scanner_ID").ColumnName = "SCAN ID"
                        grid.Columns("Scanner_5D").ColumnName = "SCAN 5D"
                        grid.Columns("Employee_Name").ColumnName = "NAME"
                        grid.Columns("Gender").ColumnName = "GENDER"
                        grid.Columns("Celebration_Date").ColumnName = "DATE"
                        grid.Columns("Buisness_Area").ColumnName = "BUISNESS"
                        grid.Columns("Functions").ColumnName = "DEPARTMENT"
                        grid.Columns("Leave_Approver").ColumnName = "REPORTING MANAGER"
                        grid.Columns("Created_At").ColumnName = "TIME STAMP"
                        BunifuDataGridView3.DataSource = grid
                        BunifuButton7.Visible = True
                        BunifuButton7.Enabled = True
                        BunifuButton7.Text = "UPDATE"
                        BunifuDataGridView3.Visible = True
                        BunifuDataGridView3.Enabled = True
                    End If
                Else
                    BunifuSnackbar1.Show(Me, $"File {selectedFilePath} has failed in Fileverification method!", Bunifu.UI.WinForms.BunifuSnackbar.MessageTypes.Error, 2000)
                    BunifuDataGridView3.DataSource = Nothing
                    BunifuDataGridView3.Visible = False
                    BunifuDataGridView3.Enabled = False
                    BunifuButton7.Visible = False
                    BunifuButton7.Enabled = False
                    BunifuTextBox5.Clear()
                    Exit Sub
                End If
            End If
        End If
    End Sub
    Private Sub BunifuButton15_Click(sender As Object, e As EventArgs) Handles BunifuButton15.Click
        clearfunction()
    End Sub
    Private Sub BunifuDropdown7_SelectedIndexChanged(sender As Object, e As EventArgs) Handles BunifuDropdown7.SelectedIndexChanged
        BunifuDataGridView3.DataSource = Nothing
        BunifuDataGridView3.Visible = False
        BunifuDataGridView3.Enabled = False
        BunifuButton7.Visible = False
        BunifuButton7.Enabled = False
        If BunifuDropdown7.SelectedItem = "REMOVE THIS DATE" AndAlso BunifuDropdown13.SelectedItem IsNot Nothing Then
            Dim result = FuncLib.Events.FetchAllDataForvariableDeletion(BunifuDropdown12.SelectedItem, BunifuDropdown13.SelectedItem)
            If result IsNot Nothing AndAlso result.Item1 IsNot Nothing Then
                Dim dataTable = result.Item1
                If dataTable.Rows.Count > 0 Then
                    dataTable.Columns("Employee_ID").ColumnName = "EMP ID"
                    dataTable.Columns("NT_ID").ColumnName = "NT ID"
                    dataTable.Columns("Scanner_ID").ColumnName = "SCAN ID"
                    dataTable.Columns("Scanner_5D").ColumnName = "SCAN 5D"
                    dataTable.Columns("Employee_Name").ColumnName = "NAME"
                    dataTable.Columns("Gender").ColumnName = "GENDER"
                    dataTable.Columns("Celebration_Date").ColumnName = "FUNCTION DATE"
                    dataTable.Columns("Buisness_Area").ColumnName = "BUISNESS"
                    dataTable.Columns("Functions").ColumnName = "DEPARTMENT"
                    dataTable.Columns("Leave_Approver").ColumnName = "REPORTING MANAGER"
                    BunifuDataGridView3.DataSource = dataTable
                    BunifuButton7.Visible = True
                    BunifuButton7.Enabled = True
                    BunifuButton7.Text = "REMOVE"
                    BunifuDataGridView3.Visible = True
                    BunifuDataGridView3.Enabled = True
                Else
                    dataTable.Columns("Employee_ID").ColumnName = "EMP ID"
                    dataTable.Columns("NT_ID").ColumnName = "NT ID"
                    dataTable.Columns("Scanner_ID").ColumnName = "SCAN ID"
                    dataTable.Columns("Scanner_5D").ColumnName = "SCAN 5D"
                    dataTable.Columns("Employee_Name").ColumnName = "NAME"
                    dataTable.Columns("Gender").ColumnName = "GENDER"
                    dataTable.Columns("Celebration_Date").ColumnName = "FUNCTION DATE"
                    dataTable.Columns("Buisness_Area").ColumnName = "BUISNESS"
                    dataTable.Columns("Functions").ColumnName = "DEPARTMENT"
                    dataTable.Columns("Leave_Approver").ColumnName = "REPORTING MANAGER"
                    BunifuDataGridView3.DataSource = dataTable
                    BunifuButton7.Visible = True
                    BunifuButton7.Enabled = False
                    BunifuButton7.Text = "REMOVE"
                    BunifuDataGridView3.Visible = True
                    BunifuDataGridView3.Enabled = True
                End If
            End If
        ElseIf BunifuDropdown7.SelectedItem = "REMOVE EMPLOYEES" AndAlso BunifuDropdown13.SelectedItem Is Nothing Then 'Constant Function
            Dim result = FuncLib.Events.FetchAllDataForConstantFunctionDeletion(BunifuDropdown12.SelectedItem)
            If result IsNot Nothing AndAlso result.Item1 IsNot Nothing Then
                Dim dataTable = result.Item1
                If dataTable.Rows.Count > 0 Then
                    dataTable.Columns("Employee_ID").ColumnName = "EMP ID"
                    dataTable.Columns("NT_ID").ColumnName = "NT ID"
                    dataTable.Columns("Scanner_ID").ColumnName = "SCAN ID"
                    dataTable.Columns("Scanner_5D").ColumnName = "SCAN 5D"
                    dataTable.Columns("Employee_Name").ColumnName = "NAME"
                    dataTable.Columns("Gender").ColumnName = "GENDER"
                    dataTable.Columns("Celebration_Date").ColumnName = "FUNCTION DATE"
                    dataTable.Columns("Buisness_Area").ColumnName = "BUISNESS"
                    dataTable.Columns("Functions").ColumnName = "DEPARTMENT"
                    dataTable.Columns("Leave_Approver").ColumnName = "REPORTING MANAGER"
                    BunifuDataGridView3.DataSource = dataTable
                    BunifuButton7.Visible = True
                    BunifuButton7.Enabled = True
                    BunifuButton7.Text = "REMOVE"
                    BunifuDataGridView3.Visible = True
                    BunifuDataGridView3.Enabled = True
                Else
                    dataTable.Columns("Employee_ID").ColumnName = "EMP ID"
                    dataTable.Columns("NT_ID").ColumnName = "NT ID"
                    dataTable.Columns("Scanner_ID").ColumnName = "SCAN ID"
                    dataTable.Columns("Scanner_5D").ColumnName = "SCAN 5D"
                    dataTable.Columns("Employee_Name").ColumnName = "NAME"
                    dataTable.Columns("Gender").ColumnName = "GENDER"
                    dataTable.Columns("Celebration_Date").ColumnName = "FUNCTION DATE"
                    dataTable.Columns("Buisness_Area").ColumnName = "BUISNESS"
                    dataTable.Columns("Functions").ColumnName = "DEPARTMENT"
                    dataTable.Columns("Leave_Approver").ColumnName = "REPORTING MANAGER"
                    BunifuDataGridView3.DataSource = dataTable
                    BunifuButton7.Visible = True
                    BunifuButton7.Enabled = False
                    BunifuButton7.Text = "REMOVE"
                    BunifuDataGridView3.Visible = True
                    BunifuDataGridView3.Enabled = True
                End If
            End If
        ElseIf BunifuDropdown7.SelectedItem = "REMOVE EMPLOYEES" AndAlso BunifuDropdown13.SelectedItem IsNot Nothing Then 'variable functions
            Dim result = FuncLib.Events.FetchAllDataForvariableDeletion(BunifuDropdown12.SelectedItem, BunifuDropdown13.SelectedItem)
            If result IsNot Nothing AndAlso result.Item1 IsNot Nothing Then
                Dim dataTable = result.Item1
                If dataTable.Rows.Count > 0 Then
                    dataTable.Columns("Employee_ID").ColumnName = "EMP ID"
                    dataTable.Columns("NT_ID").ColumnName = "NT ID"
                    dataTable.Columns("Scanner_ID").ColumnName = "SCAN ID"
                    dataTable.Columns("Scanner_5D").ColumnName = "SCAN 5D"
                    dataTable.Columns("Employee_Name").ColumnName = "NAME"
                    dataTable.Columns("Gender").ColumnName = "GENDER"
                    dataTable.Columns("Celebration_Date").ColumnName = "FUNCTION DATE"
                    dataTable.Columns("Buisness_Area").ColumnName = "BUISNESS"
                    dataTable.Columns("Functions").ColumnName = "DEPARTMENT"
                    dataTable.Columns("Leave_Approver").ColumnName = "REPORTING MANAGER"
                    BunifuDataGridView3.DataSource = dataTable
                    BunifuButton7.Visible = True
                    BunifuButton7.Enabled = True
                    BunifuButton7.Text = "REMOVE"
                    BunifuDataGridView3.Visible = True
                    BunifuDataGridView3.Enabled = True
                Else
                    dataTable.Columns("Employee_ID").ColumnName = "EMP ID"
                    dataTable.Columns("NT_ID").ColumnName = "NT ID"
                    dataTable.Columns("Scanner_ID").ColumnName = "SCAN ID"
                    dataTable.Columns("Scanner_5D").ColumnName = "SCAN 5D"
                    dataTable.Columns("Employee_Name").ColumnName = "NAME"
                    dataTable.Columns("Gender").ColumnName = "GENDER"
                    dataTable.Columns("Celebration_Date").ColumnName = "FUNCTION DATE"
                    dataTable.Columns("Buisness_Area").ColumnName = "BUISNESS"
                    dataTable.Columns("Functions").ColumnName = "DEPARTMENT"
                    dataTable.Columns("Leave_Approver").ColumnName = "REPORTING MANAGER"
                    BunifuDataGridView3.DataSource = dataTable
                    BunifuButton7.Visible = True
                    BunifuButton7.Enabled = False
                    BunifuButton7.Text = "REMOVE"
                    BunifuDataGridView3.Visible = True
                    BunifuDataGridView3.Enabled = True
                End If
            End If
        End If
    End Sub
    Private Sub BunifuButton5_Click(sender As Object, e As EventArgs) Handles BunifuButton5.Click
        BunifuDropdown3.Items.Clear()
        BunifuDropdown3.Text = Nothing
        BunifuLabel24.Visible = False
        BunifuDropdown3.Visible = False
        BunifuLabel27.Visible = False
        BunifuButton21.Enabled = False
        BunifuButton21.Visible = False
        BunifuLabel28.Visible = False
        BunifuTextBox19.Clear()
        BunifuTextBox19.Visible = False
        BunifuTextBox19.Enabled = False
        BunifuTextBox20.Clear()
        BunifuTextBox20.Visible = False
        BunifuTextBox20.Enabled = False
        GroupBox4.Visible = False
        GroupBox4.Enabled = False
        BunifuTextBox29.Clear()
        BunifuTextBox29.Enabled = False
        GroupBox5.Visible = False
        GroupBox5.Enabled = False
        BunifuTextBox28.Clear()
        BunifuTextBox26.Clear()
        BunifuTextBox27.Clear()
        BunifuTextBox23.Clear()
        BunifuTextBox25.Clear()
        BunifuTextBox24.Clear()
        BunifuButton28.Visible = False
        BunifuButton28.Enabled = False
        BunifuButton22.Visible = False
        BunifuButton22.Enabled = False
        Dim Occasion As List(Of String) = FuncLib.Gifts.GetOcaasioList()
        If Occasion.Count <= 0 AndAlso TypeOf Occasion(0) Is String Then
            Dim errorMessage As String = Occasion(0)
            BunifuSnackbar1.Show(Me, $"Error While Fetching Occasion List From. {Environment.NewLine} Error : {errorMessage}", Bunifu.UI.WinForms.BunifuSnackbar.MessageTypes.Warning, 3000)
            FuncLib.WriteLog.WriteErrorLog($"Error occured While Trying To Fetch Occasion List : {errorMessage}")
            Exit Sub
        Else
            BunifuLabel24.Visible = True
            BunifuLabel24.Enabled = True
            BunifuDropdown3.Visible = True
            BunifuDropdown3.Enabled = True
            BunifuDropdown3.Items.AddRange(Occasion.ToArray())
            If BunifuPages1.Visible = False Then
                BunifuPages1.TabPages.Clear()
                transition.ShowSync(BunifuPages1, False, BunifuAnimatorNS.Animation.ScaleAndHorizSlide)
                BunifuPages1.TabPages.Add(TabPage6)
            Else
                transition.HideSync(BunifuPages1, False, BunifuAnimatorNS.Animation.ScaleAndHorizSlide)
                BunifuPages1.TabPages.Clear()
                transition.ShowSync(BunifuPages1, False, BunifuAnimatorNS.Animation.ScaleAndHorizSlide)
                BunifuPages1.TabPages.Add(TabPage6)
            End If
        End If
    End Sub
    Private Sub BunifuDropdown3_SelectedIndexChanged(sender As Object, e As EventArgs) Handles BunifuDropdown3.SelectedIndexChanged
        Dim CheckCriteriaDate = FuncLib.Setup.FetchCriterialCount(BunifuDropdown3.SelectedItem)
        If CheckCriteriaDate.Item3 <> "" Then 'Error Case
            BunifuSnackbar1.Show(Me, $"Error While Fetching Occasion List From. {Environment.NewLine} Error : {CheckCriteriaDate.Item3}", Bunifu.UI.WinForms.BunifuSnackbar.MessageTypes.Warning, 2500)
            FuncLib.WriteLog.WriteErrorLog($"Error occured While Trying To Fetch Occasion List : {CheckCriteriaDate.Item3}")
            Exit Sub
        ElseIf CheckCriteriaDate.Item1 > 0 AndAlso CheckCriteriaDate.Item2 > 0 Then 'Variable Functiuon
            BunifuLabel27.Visible = True
            BunifuLabel27.Enabled = True
            BunifuTextBox19.Text = CInt(CheckCriteriaDate.Item1)
            BunifuTextBox19.Visible = True
            BunifuTextBox20.Text = CInt(CheckCriteriaDate.Item2)
            BunifuTextBox20.Visible = True
            BunifuTextBox20.Enabled = False
            BunifuLabel28.Visible = True
            BunifuLabel28.Enabled = True
            BunifuButton21.Visible = True
            BunifuButton21.Enabled = True
            BunifuTextBox19.Enabled = True
            BunifuTextBox20.Enabled = True
        ElseIf CheckCriteriaDate.Item1 = 0 AndAlso CheckCriteriaDate.Item2 > 0 Then 'Constant Function
            BunifuLabel27.Visible = False
            BunifuLabel27.Enabled = False
            BunifuTextBox19.Clear()
            BunifuTextBox19.Visible = False
            BunifuTextBox20.Text = CInt(CheckCriteriaDate.Item2)
            BunifuTextBox20.Visible = True
            BunifuLabel28.Visible = True
            BunifuLabel28.Enabled = True
            BunifuButton21.Visible = True
            BunifuButton21.Enabled = True
            BunifuTextBox19.Enabled = True
            BunifuTextBox20.Enabled = True
        End If
    End Sub
    Private Sub BunifuButton21_Click(sender As Object, e As EventArgs) Handles BunifuButton21.Click
        Dim Occasiontype As String = FuncLib.Gifts.CheckTypeOfOcassion(BunifuDropdown3.SelectedItem)
        If Occasiontype = "Constant" Then
            Dim value1 As Integer
            If Integer.TryParse(BunifuTextBox20.Text, value1) Then
                Dim Updatevalue As String = FuncLib.Setup.UpdateCriteriaFunction(BunifuDropdown3.SelectedItem, 0, value1)
                If Updatevalue = "True" Then
                    BunifuSnackbar1.Show(Me, $"Repeat Span Time and Eligibility Time Has been updated for {BunifuDropdown3.SelectedItem}", Bunifu.UI.WinForms.BunifuSnackbar.MessageTypes.Success, 1000)
                    ClearAdvancedSetupSquence()
                Else
                    BunifuSnackbar1.Show(Me, $"Error While Fetching Occasion List From. {Environment.NewLine} Error : {Updatevalue}", Bunifu.UI.WinForms.BunifuSnackbar.MessageTypes.Warning, 2500)
                    FuncLib.WriteLog.WriteErrorLog($"Error occured While Trying To Fetch Occasion List : {Updatevalue}")
                    Exit Sub
                End If
            End If
        ElseIf Occasiontype = "Variable" Then
            Dim value2 As Integer
            If Integer.TryParse(BunifuTextBox19.Text, value2) Then
                Dim value3 As Integer
                If Integer.TryParse(BunifuTextBox20.Text, value3) Then
                    If value2 < value3 Then
                        BunifuSnackbar1.Show(Me, $"RepeatSpan Time can not be less than eligibility time!", Bunifu.UI.WinForms.BunifuSnackbar.MessageTypes.Warning, 1000)
                        Dim result = FuncLib.Setup.FetchCriterialCount(BunifuDropdown3.SelectedItem)
                        BunifuTextBox19.Text = CInt(result.Item1)
                        BunifuTextBox20.Text = CInt(result.Item2)
                    ElseIf value2 >= value3 Then
                        Dim Updatevalue As String = FuncLib.Setup.UpdateCriteriaFunction(BunifuDropdown3.SelectedItem, value2, value3)
                        If Updatevalue = "True" Then
                            BunifuSnackbar1.Show(Me, $"Repeat Span Time and Eligibility Time Has been updated for {BunifuDropdown3.SelectedItem}", Bunifu.UI.WinForms.BunifuSnackbar.MessageTypes.Success, 1000)
                            ClearAdvancedSetupSquence()
                        Else
                            BunifuSnackbar1.Show(Me, $"Error While Fetching Occasion List From. {Environment.NewLine} Error : {Updatevalue}", Bunifu.UI.WinForms.BunifuSnackbar.MessageTypes.Warning, 2500)
                            FuncLib.WriteLog.WriteErrorLog($"Error occured While Trying To Fetch Occasion List : {Updatevalue}")
                            Exit Sub
                        End If
                    End If
                End If
            End If
        End If
    End Sub
    Private Sub ClearAdvancedSetupSquence()
        BunifuDropdown3.Items.Clear()
        BunifuDropdown3.Text = Nothing
        BunifuLabel24.Visible = False
        BunifuDropdown3.Visible = False
        BunifuLabel27.Visible = False
        BunifuButton21.Enabled = False
        BunifuButton21.Visible = False
        BunifuLabel28.Visible = False
        BunifuTextBox19.Clear()
        BunifuTextBox19.Visible = False
        BunifuTextBox19.Enabled = False
        BunifuTextBox20.Clear()
        BunifuTextBox20.Visible = False
        BunifuTextBox20.Enabled = False
        Dim Occasion As List(Of String) = FuncLib.Gifts.GetOcaasioList()
        If Occasion.Count <= 0 AndAlso TypeOf Occasion(0) Is String Then
            Dim errorMessage As String = Occasion(0)
            BunifuSnackbar1.Show(Me, $"Error While Fetching Occasion List From. {Environment.NewLine} Error : {errorMessage}", Bunifu.UI.WinForms.BunifuSnackbar.MessageTypes.Warning, 3000)
            FuncLib.WriteLog.WriteErrorLog($"Error occured While Trying To Fetch Occasion List : {errorMessage}")
            Exit Sub
        Else
            BunifuLabel24.Visible = True
            BunifuLabel24.Enabled = True
            BunifuDropdown3.Visible = True
            BunifuDropdown3.Enabled = True
            BunifuDropdown3.Items.AddRange(Occasion.ToArray())
        End If
    End Sub
    Private Sub BunifuButton14_Click(sender As Object, e As EventArgs) Handles BunifuButton14.Click
        Dim iniFile As New IniFile()
        Try
            iniFile.Load(Inipath)
            If iniFile.Sections.Contains("SERIAL") Then
                Dim currentHostName As String = System.Net.Dns.GetHostName()
                Dim serialSection As IniSection = iniFile.Sections("SERIAL")
                Dim matchFound As Boolean = False
                For Each key As IniKey In serialSection.Keys
                    If key.Name.Equals(currentHostName) Then
                        key.Value = BunifuDropdown8.SelectedItem.ToString
                        iniFile.Save(Inipath)
                        matchFound = True
                        Exit For
                    End If
                Next
                If Not matchFound Then
                    serialSection.Keys.Add(currentHostName, BunifuDropdown8.SelectedItem.ToString)
                    iniFile.Save(Inipath)
                End If
            End If
            BunifuSnackbar1.Show(Me, $"Confriguration Updated Successfully!", Bunifu.UI.WinForms.BunifuSnackbar.MessageTypes.Success, 2000)
            FuncLib.WriteLog.WriteAppLog($"Comport Confriguration Changed By {MainWindow.ListBox1.Items(1)}_{MainWindow.ListBox1.Items(4)} Changes :{currentHostName}:{BunifuDropdown8.SelectedItem}")
            RebbotInitiaizer.TextBox1.Text = 5
            RebbotInitiaizer.BackColor = Color.LimeGreen
            RebbotInitiaizer.ShowDialog()
        Catch ex As Exception
            BunifuSnackbar1.Show(Me, $"Error Occured While Saving the Confriguration File.{Environment.NewLine}Error:{ex.Message}", Bunifu.UI.WinForms.BunifuSnackbar.MessageTypes.Error, 2000)
            FuncLib.WriteLog.WriteErrorLog($"Error Occured While Updaing Confriguration :{ex.Message}")
            Exit Sub
        End Try
    End Sub
    Private Sub BunifuButton16_Click(sender As Object, e As EventArgs) Handles BunifuButton16.Click
        Ini.Load(Inipath)
        Try
            If serialPort.PortName <> Ini.Sections("SERIAL").Keys($"{currentHostName}").Value AndAlso serialPort.IsOpen = False Then
                serialPort.PortName = Ini.Sections("SERIAL").Keys($"{currentHostName}").Value
                serialPort.BaudRate = 9600
                serialPort.Parity = Parity.None
                serialPort.StopBits = StopBits.One
                If serialPort.IsOpen = False Then
                    serialPort.Open()
                    serialPort.Write("START")
                    GroupBox4.Visible = True
                    GroupBox4.Enabled = True
                    BunifuTextBox29.Enabled = True
                    BunifuTextBox29.ReadOnly = True
                    BunifuTextBox29.Clear()
                    BunifuTextBox29.PlaceholderText = "SCAN YOUR ID CARD"
                Else
                    serialPort.Close()
                    Thread.Sleep(1000)
                    serialPort.Open()
                    GroupBox4.Visible = True
                    GroupBox4.Enabled = True
                    BunifuTextBox29.Enabled = True
                    BunifuTextBox29.ReadOnly = True
                    BunifuTextBox29.Clear()
                    BunifuTextBox29.PlaceholderText = "SCAN YOUR ID CARD"
                End If
            Else
                serialPort.BaudRate = 9600
                serialPort.Parity = Parity.None
                serialPort.StopBits = StopBits.One
                If serialPort.IsOpen = False Then
                    serialPort.Open()
                    GroupBox4.Visible = True
                    GroupBox4.Enabled = True
                    BunifuTextBox29.Clear()
                    BunifuTextBox29.Enabled = True
                    BunifuTextBox29.ReadOnly = True
                    BunifuTextBox29.PlaceholderText = "SCAN YOUR ID CARD"
                Else
                    serialPort.Close()
                    System.Threading.Thread.Sleep(1000)
                    serialPort.Open()
                    serialPort.Write("START")
                    GroupBox4.Visible = True
                    GroupBox4.Enabled = True
                    BunifuTextBox29.Clear()
                    BunifuTextBox29.Enabled = True
                    BunifuTextBox29.ReadOnly = True
                    BunifuTextBox29.PlaceholderText = "SCAN YOUR ID CARD"
                End If
            End If
        Catch ex As Exception
            BunifuSnackbar1.Show(Me, $"Error Occured While Trying To Open Comport!{Environment.NewLine}Error:{ex.Message}", Bunifu.UI.WinForms.BunifuSnackbar.MessageTypes.Error, 2500)
            FuncLib.WriteLog.WriteErrorLog($"Error Occured While Tring To Open COmport While Sahring Gift :{ex.Message}")

            Exit Sub
        End Try
    End Sub
    Private Sub SerialPort_DataReceived(sender As Object, e As SerialDataReceivedEventArgs) Handles serialPort.DataReceived
        Dim hexData As String = serialPort.ReadLine()
        Dim extractedString As String = hexData.Substring(1, hexData.Length - 2)
        BunifuTextBox29.Invoke(Sub()
                                   BunifuTextBox29.Text = FuncLib.Gifts.IdentifyEmployee(extractedString, "BY 10D")
                                   Dim result = FuncLib.Setup.FetchEmployeeDataForExceptionGiftSharing(BunifuTextBox29.Text)
                                   If result.Item1 = "" Then
                                       GroupBox5.Visible = True
                                       GroupBox5.Enabled = True
                                       BunifuButton28.Visible = True
                                       BunifuButton28.Enabled = True
                                       BunifuButton22.Visible = True
                                       BunifuButton22.Enabled = True
                                       BunifuTextBox28.Text = result.Item2.ToString
                                       BunifuTextBox28.ReadOnly = True
                                       BunifuTextBox26.Text = result.Item3.ToString
                                       BunifuTextBox26.ReadOnly = True
                                       BunifuTextBox27.Text = result.Item4.ToString
                                       BunifuTextBox27.ReadOnly = True
                                       BunifuTextBox23.Text = result.Item5.ToString
                                       BunifuTextBox23.ReadOnly = True
                                       BunifuTextBox24.Text = result.Item6.ToString
                                       BunifuTextBox24.ReadOnly = True
                                       BunifuTextBox25.Text = result.Item7.ToString
                                       BunifuTextBox25.ReadOnly = True
                                       BunifuButton28.Enabled = True
                                       BunifuButton28.Visible = True
                                       BunifuButton22.Enabled = True
                                       BunifuButton22.Visible = True
                                       BunifuDropdown4.Items.Clear()
                                       BunifuDropdown4.Items.Add("Exception Gift")
                                       BunifuDropdown4.SelectedItem = BunifuDropdown4.Items(0)
                                       BunifuDropdown4.Enabled = False
                                       BunifuTextBox29.Enabled = False
                                       BunifuSnackbar1.Show(Me, $"{BunifuTextBox26.Text} is selected for exceptional gift sharing!", Bunifu.UI.WinForms.BunifuSnackbar.MessageTypes.Success, 1000)
                                   ElseIf result.Item1 = "Invalid" Then
                                       GroupBox5.Visible = False
                                       GroupBox5.Enabled = False
                                       BunifuButton28.Visible = False
                                       BunifuButton28.Enabled = False
                                       BunifuButton22.Visible = False
                                       BunifuButton22.Enabled = False
                                       BunifuTextBox28.ReadOnly = True
                                       BunifuTextBox26.ReadOnly = True
                                       BunifuTextBox27.ReadOnly = True
                                       BunifuTextBox23.ReadOnly = True
                                       BunifuTextBox24.ReadOnly = True
                                       BunifuTextBox25.ReadOnly = True
                                       BunifuButton28.Enabled = True
                                       BunifuButton28.Visible = True
                                       BunifuButton22.Enabled = True
                                       BunifuButton22.Visible = True
                                       BunifuTextBox29.Clear()
                                       BunifuTextBox29.Enabled = True
                                       BunifuSnackbar1.Show(Me, $"Not A a valid Employee!", Bunifu.UI.WinForms.BunifuSnackbar.MessageTypes.Information, 1000)
                                   Else
                                       BunifuSnackbar1.Show(Me, $"Error Occured While Trying To Fetch Information FromDatabase!{Environment.NewLine}Error :{result.Item1}", Bunifu.UI.WinForms.BunifuSnackbar.MessageTypes.Error, 2500)
                                       FuncLib.WriteLog.WriteErrorLog($"Error Occured While executing FuncLib.Setup.FetchEmployeeDataForExceptionGiftSharing:{result.Item1}")
                                       Exit Sub
                                   End If
                               End Sub)
        Thread.Sleep(1000)
        serialPort.DiscardInBuffer()
        serialPort.DiscardOutBuffer()
    End Sub
    Private Sub SETUP_Closing(sender As Object, e As CancelEventArgs) Handles Me.Closing
        If serialPort.IsOpen Then
            serialPort.Close()
        End If
    End Sub
    Private Sub BunifuButton28_Click(sender As Object, e As EventArgs) Handles BunifuButton28.Click
        If BunifuDropdown4.SelectedItem IsNot Nothing Then
            Dim Employee_ID As String = BunifuTextBox28.Text
            Dim NT_ID As String = BunifuTextBox27.Text
            Dim Employee_Name As String = BunifuTextBox26.Text
            Dim Celebration_Date As Date = DateTime.Now.Date
            Dim Buisness_Area As String = BunifuTextBox25.Text
            Dim Functions As String = BunifuTextBox24.Text
            Dim Leave_Approver As String = BunifuTextBox23.Text
            Dim Gift_Type As String = BunifuDropdown4.SelectedItem
            Dim ShareGift As String = FuncLib.Setup.ShareExceptionGift(Employee_ID, NT_ID, Employee_Name, Celebration_Date, Buisness_Area, Functions, Leave_Approver, Gift_Type)
            If ShareGift = "True" Then
                GroupBox5.Visible = False
                GroupBox5.Enabled = False
                BunifuButton28.Visible = False
                BunifuButton28.Enabled = False
                BunifuButton22.Visible = False
                BunifuButton22.Enabled = False
                BunifuTextBox28.ReadOnly = True
                BunifuTextBox26.ReadOnly = True
                BunifuTextBox27.ReadOnly = True
                BunifuTextBox23.ReadOnly = True
                BunifuTextBox24.ReadOnly = True
                BunifuTextBox25.ReadOnly = True
                BunifuButton28.Enabled = True
                BunifuButton28.Visible = True
                BunifuButton22.Enabled = True
                BunifuButton22.Visible = True
                BunifuTextBox29.Clear()
                BunifuTextBox29.Enabled = True
                BunifuTextBox29.ReadOnly = True
                BunifuSnackbar1.Show(Me, $"{BunifuDropdown4.SelectedItem} Gift Shared With {BunifuTextBox26.Text}", Bunifu.UI.WinForms.BunifuSnackbar.MessageTypes.Success, 1000)
            Else
                BunifuSnackbar1.Show(Me, $"Error Occured While Share {BunifuDropdown4.SelectedItem} with {BunifuTextBox26.Text}{Environment.NewLine}Error :{ShareGift}", Bunifu.UI.WinForms.BunifuSnackbar.MessageTypes.Error, 2500)
                FuncLib.WriteLog.WriteErrorLog($"Error Occured While executing FuncLib.Setup.FetchEmployeeDataForExceptionGiftSharing:{ShareGift}")
                GroupBox5.Visible = False
                GroupBox5.Enabled = False
                BunifuButton28.Visible = False
                BunifuButton28.Enabled = False
                BunifuButton22.Visible = False
                BunifuButton22.Enabled = False
                BunifuTextBox28.ReadOnly = True
                BunifuTextBox26.ReadOnly = True
                BunifuTextBox27.ReadOnly = True
                BunifuTextBox23.ReadOnly = True
                BunifuTextBox24.ReadOnly = True
                BunifuTextBox25.ReadOnly = True
                BunifuButton28.Enabled = True
                BunifuButton28.Visible = True
                BunifuButton22.Enabled = True
                BunifuButton22.Visible = True
                BunifuTextBox29.Clear()
                BunifuTextBox29.Enabled = True
                BunifuTextBox29.ReadOnly = True
            End If
        Else
            BunifuSnackbar1.Show(Me, $"No Gift type Selected To Share", Bunifu.UI.WinForms.BunifuSnackbar.MessageTypes.Warning, 2500)
        End If
    End Sub

    Private Sub BunifuButton22_Click(sender As Object, e As EventArgs) Handles BunifuButton22.Click
        GroupBox5.Visible = False
        GroupBox5.Enabled = False
        BunifuButton28.Visible = False
        BunifuButton28.Enabled = False
        BunifuButton22.Visible = False
        BunifuButton22.Enabled = False
        BunifuTextBox28.ReadOnly = True
        BunifuTextBox26.ReadOnly = True
        BunifuTextBox27.ReadOnly = True
        BunifuTextBox23.ReadOnly = True
        BunifuTextBox24.ReadOnly = True
        BunifuTextBox25.ReadOnly = True
        BunifuButton28.Enabled = True
        BunifuButton28.Visible = True
        BunifuButton22.Enabled = True
        BunifuButton22.Visible = True
        BunifuTextBox29.Clear()
        BunifuTextBox29.Enabled = True
        BunifuTextBox29.ReadOnly = True
    End Sub
    Private Sub BunifuButton18_Click(sender As Object, e As EventArgs) Handles BunifuButton18.Click
        Ini.Load(Inipath)
        Dim iret As Object = MsgBox($"This Will Clear All The Gift Shared History. Make Sure To Take A backup before Erasing Data!{Environment.NewLine} Are You Sure Want To Clear HGistory?", vbQuestion + vbYesNo, "Clear_History")
        If iret = vbYes Then
            BunifuDataGridView4.DataSource = Nothing
            Dim HistoryData = FuncLib.History.FetchAllHistory
            Dim dataTable = HistoryData.Item1
            BunifuDataGridView4.DataSource = dataTable
            BunifuDataGridView4.Visible = False
            Dim backupFolderPath As String = $"{Ini.Sections("BASIC").Keys("Backup_Location").Value}"
            Dim backupFileName As String = $"Gift_History_Backup_{DateTime.Now.ToString("yyyyMMddHHmmss")}.csv"
            Dim backupFilePath As String = Path.Combine(backupFolderPath, backupFileName)
            Dim SaveFileasExcel As String = FuncLib.History.SaveasExcel(BunifuDataGridView4, $"{backupFilePath}")
            Dim ClearHistory As String = FuncLib.Setup.ClearHistory()
            If ClearHistory = "True" Then
                BunifuSnackbar1.Show(Me, $"Gift History Cleared Successfully.", Bunifu.UI.WinForms.BunifuSnackbar.MessageTypes.Success, 1000)
                FuncLib.WriteLog.WriteAppLog($"Gift_History Cleared")
            Else
                BunifuSnackbar1.Show(Me, $"Error Occured While Trying To Clear Gift history {Environment.NewLine} Error : {ClearHistory}", Bunifu.UI.WinForms.BunifuSnackbar.MessageTypes.Error, 2500)
                FuncLib.WriteLog.WriteErrorLog($"Error Occured While executing FuncLib.Setup.ClearHistory:{ClearHistory}")
                Exit Sub
            End If
        ElseIf iret = vbNo Then
            'donothing
        End If
    End Sub
    Private Sub BunifuButton33_Click(sender As Object, e As EventArgs) Handles BunifuButton33.Click
        Dim folderBrowserDialog As New FolderBrowserDialog()
        folderBrowserDialog.SelectedPath = "C:\"
        If folderBrowserDialog.ShowDialog() = DialogResult.OK Then
            BunifuTextBox7.Text = folderBrowserDialog.SelectedPath
        End If
    End Sub
    Private Sub BunifuButton32_Click(sender As Object, e As EventArgs) Handles BunifuButton32.Click
        Dim Ini1 As New IniFile
        Try
            Ini1.Load(Inipath)
            Ini1.Sections("BASIC").Keys("Backup_Location").Value = BunifuTextBox7.Text.ToString
            Ini1.Save(Inipath)
            BunifuTextBox7.BorderColorIdle = Color.Green
            BunifuSnackbar1.Show(Me, $"Confriguration Updated Successfully!", Bunifu.UI.WinForms.BunifuSnackbar.MessageTypes.Success, 2000)
            FuncLib.WriteLog.WriteAppLog($"Backup Location Confriguration Changed By {MainWindow.ListBox1.Items(1)}_{MainWindow.ListBox1.Items(4)} Changes:{BunifuTextBox7.Text}")
            RebbotInitiaizer.TextBox1.Text = 5
            RebbotInitiaizer.BackColor = Color.LimeGreen
            RebbotInitiaizer.ShowDialog()
        Catch ex As Exception
            BunifuSnackbar1.Show(Me, $"Error Occured While Saving the Confriguration File.{Environment.NewLine}Error:{ex.Message}", Bunifu.UI.WinForms.BunifuSnackbar.MessageTypes.Error, 2000)
            FuncLib.WriteLog.WriteErrorLog($"Error Occured While Updaing Confriguration :{ex.Message}")
            Exit Sub
        End Try

    End Sub


End Class