Imports System.ComponentModel
Imports System.Configuration
Imports System.Windows.Forms.VisualStyles
Imports BunifuAnimatorNS
Imports MadMilkman.Ini
Public Class Add_Events
    Dim requiredColumns As New List(Of String)()
    Dim transition As New BunifuTransition
    Dim Inipath As String = ConfigurationManager.AppSettings("IniFile")
    Private Sub Add_Events_Load(sender As Object, e As EventArgs) Handles MyBase.Load
        Me.StartPosition = FormStartPosition.CenterScreen
        ClearSequenceWhileInitializing()
    End Sub
    Private Sub ClearSequenceWhileInitializing()
        BunifuTextBox1.Clear()
        CheckBox3.Checked = False
        CheckBox3.Visible = False
        CheckBox3.Enabled = False
        ProgressBar1.Minimum = 0
        ProgressBar1.Maximum = 100
        ProgressBar1.Value = 0
        BunifuTextBox1.Enabled = True
        BunifuLabel1.Visible = True
        BunifuLabel1.Enabled = True
        BunifuPanel1.Visible = True
        BunifuPanel1.Enabled = True
        BunifuLabel3.Visible = False
        BunifuDropdown1.Visible = False
        BunifuDropdown1.Items.Clear()
        BunifuDropdown1.Text = -Nothing
        BunifuDropdown1.Enabled = False
        BunifuLabel2.Visible = False
        BunifuDatePicker1.Visible = False
        BunifuDatePicker1.Enabled = False
        BunifuButton1.Visible = False
        BunifuButton1.Enabled = False
        GroupBox1.Visible = False
        GroupBox1.Enabled = False
        GroupBox2.Visible = False
        GroupBox2.Enabled = False
        BunifuDropdown7.Items.Clear()
        BunifuDropdown7.Text = Nothing
        GroupBox12.Visible = False
        GroupBox12.Enabled = False
        BunifuDropdown11.Items.Clear()
        BunifuDropdown11.Text = Nothing
        BunifuTextBox10.Clear()
        BunifuButton5.Visible = False
        BunifuButton5.Enabled = False
        BunifuButton2.Visible = True
        BunifuButton2.Enabled = True
        BunifuDataGridView1.Visible = False
        BunifuDataGridView1.Enabled = False
        BunifuDataGridView1.DataSource = Nothing
        CheckBox1.Checked = False
        CheckBox1.Visible = False
    End Sub
    Private Sub BunifuTextBox1_KeyPress(sender As Object, e As KeyPressEventArgs) Handles BunifuTextBox1.KeyPress
        Dim invalidChars As String = "*\/[]!@#$%^&(){}|;:'""<>,.?+= "
        If invalidChars.Contains(e.KeyChar) Then
            e.Handled = True
            BunifuSnackbar1.Show(Me, $"{e.KeyChar} Is Not Allowed Please use only alphanumeric characters and underscores.", Bunifu.UI.WinForms.BunifuSnackbar.MessageTypes.Warning, 3000)
        End If
    End Sub
    Private Sub BunifuButton6_Click(sender As Object, e As EventArgs) Handles BunifuButton6.Click
        If BunifuTextBox1.Text IsNot Nothing Then
            Dim CheckTableexistance As String = FuncLib.Events.CheckIfOcassionAleareadyExist(BunifuTextBox1.Text)
            If CheckTableexistance = "True" Then
                BunifuSnackbar1.Show(Me, $"{BunifuTextBox1.Text}Is already present inside the Database.", Bunifu.UI.WinForms.BunifuSnackbar.MessageTypes.Information, 2500)
                Exit Sub
            ElseIf CheckTableexistance = "False" Then
                BunifuTextBox1.Enabled = False
                BunifuPanel1.Enabled = False
                BunifuLabel3.Visible = True
                BunifuLabel3.Enabled = True
                BunifuDropdown1.Visible = True
                BunifuDropdown1.Enabled = True
                BunifuDropdown1.Items.Clear()
                BunifuDropdown1.Text = Nothing
                BunifuDropdown1.Items.Add("Repeat Every Year")
                BunifuDropdown1.Items.Add("Only This Time")
            Else
                BunifuSnackbar1.Show(Me, $"Error Occured While Executing FuncLib.Events.CheckIfOcassionAleareadyExist(BunifuTextBox1.Text) : {CheckTableexistance} ", Bunifu.UI.WinForms.BunifuSnackbar.MessageTypes.Warning, 3000)
                FuncLib.WriteLog.WriteErrorLog($"Error Occured While Executing FuncLib.Events.CheckIfOcassionAleareadyExist(BunifuTextBox1.Text) : {CheckTableexistance}")
                ClearSequenceWhileInitializing()
                Exit Sub
            End If
        End If
    End Sub
    Private Sub BunifuDropdown1_SelectedIndexChanged(sender As Object, e As EventArgs) Handles BunifuDropdown1.SelectedIndexChanged
        ClearSequenceAfterSelectingFunctionProperty()
        If BunifuDropdown1.SelectedItem = "Repeat Every Year" Then
            BunifuLabel2.Visible = True
            BunifuLabel2.Enabled = True
            BunifuDatePicker1.Visible = True
            BunifuDatePicker1.Enabled = True
            BunifuButton1.Visible = True
            BunifuButton1.Enabled = True
        ElseIf BunifuDropdown1.SelectedItem = "Only This Time" Then
            BunifuLabel2.Visible = True
            BunifuLabel2.Enabled = True
            BunifuDatePicker1.Visible = True
            BunifuDatePicker1.Enabled = True
            BunifuButton1.Visible = True
            BunifuButton1.Enabled = True
        End If
    End Sub
    Private Sub ClearSequenceAfterSelectingFunctionProperty()
        BunifuLabel2.Visible = False
        BunifuDatePicker1.Visible = False
        BunifuDatePicker1.Enabled = False
        BunifuButton1.Visible = False
        BunifuButton1.Enabled = False
        GroupBox1.Visible = False
        GroupBox1.Enabled = False
        GroupBox2.Visible = False
        GroupBox2.Enabled = False
        BunifuDropdown7.Items.Clear()
        BunifuDropdown7.Text = Nothing
        GroupBox12.Visible = False
        GroupBox12.Enabled = False
        BunifuDropdown11.Items.Clear()
        BunifuDropdown11.Text = Nothing
        BunifuTextBox10.Clear()
        BunifuButton5.Visible = False
        BunifuButton5.Enabled = False
        BunifuButton2.Visible = False
        BunifuButton2.Enabled = False
        BunifuDataGridView1.Visible = False
        BunifuDataGridView1.Enabled = False
        BunifuDataGridView1.DataSource = Nothing
    End Sub
    Private Sub BunifuButton1_Click(sender As Object, e As EventArgs) Handles BunifuButton1.Click
        If BunifuDatePicker1.Value.Date >= DateTime.Now.Date Then
            Dim CreateINIRecord As String = FuncLib.Events.CreateNewOcassion(BunifuTextBox1.Text, BunifuDropdown1.SelectedItem.ToString, BunifuDatePicker1.Value.Date)
            If CreateINIRecord = "True" Then
                Dim CreateDBTable As String = FuncLib.Events.CreateNewoccasionTable(BunifuTextBox1.Text)
                If CreateDBTable = "True" Then
                    BunifuSnackbar1.Show(Me, $"{BunifuTextBox1.Text} has been Created Successfully", Bunifu.UI.WinForms.BunifuSnackbar.MessageTypes.Success, 2000)
                    BunifuLabel3.Enabled = False
                    BunifuDropdown1.Enabled = False
                    BunifuLabel2.Enabled = False
                    BunifuDatePicker1.Enabled = False
                    BunifuButton1.Enabled = False
                    GroupBox2.Visible = True
                    GroupBox2.Enabled = True
                    BunifuDropdown7.Items.Clear()
                    BunifuDropdown7.Text = Nothing
                    BunifuDropdown7.Enabled = True
                    BunifuDropdown7.Visible = True
                    BunifuDropdown7.Items.Add("From External")
                    BunifuDropdown7.Items.Add("From Internal")
                    BunifuButton2.Visible = True
                    BunifuButton2.Enabled = True
                    CheckBox1.Checked = True
                Else
                    BunifuSnackbar1.Show(Me, $"Error Occured While Executing FuncLib.Events.CreateNewoccasionTable(BunifuTextBox1.Text). {Environment.NewLine} Error : {CreateDBTable}", Bunifu.UI.WinForms.BunifuSnackbar.MessageTypes.Error, 3000)
                    FuncLib.WriteLog.WriteErrorLog($"Error Occured While Executing FuncLib.Events.CreateNewoccasionTable(BunifuTextBox1.Text) : {CreateDBTable}")
                    If BunifuDropdown1.SelectedItem = "Repeat Every Year" Then
                        Dim iniData As New IniFile()
                        iniData.Load(Inipath)
                        RemoveKeyFromSection(iniData.Sections("REPEATSPANTIME"), $"{BunifuTextBox1.Text}")
                        RemoveKeyFromSection(iniData.Sections("TYPE"), $"{BunifuTextBox1.Text}")
                        RemoveKeyFromSection(iniData.Sections("DATE"), $"{BunifuTextBox1.Text}")
                        RemoveKeyFromSection(iniData.Sections("ELIGIBILITY"), $"{BunifuTextBox1.Text}")
                        iniData.Save(Inipath)
                    Else
                        Dim iniData As New IniFile()
                        iniData.Load(Inipath)
                        RemoveKeyFromSection(iniData.Sections("TYPE"), $"{BunifuTextBox1.Text}")
                        RemoveKeyFromSection(iniData.Sections("DATE"), $"{BunifuTextBox1.Text}")
                        RemoveKeyFromSection(iniData.Sections("ELIGIBILITY"), $"{BunifuTextBox1.Text}")
                        iniData.Save(Inipath)
                    End If
                    Exit Sub
                End If
            Else
                BunifuSnackbar1.Show(Me, $"Error Executing FuncLib.Events.CreateNewOcassion(BunifuTextBox1.Text, BunifuDropdown1.SelectedItem.ToString, BunifuDatePicker1.Value.Date). {Environment.NewLine} Error : {CreateINIRecord}", Bunifu.UI.WinForms.BunifuSnackbar.MessageTypes.Error, 3000)
                FuncLib.WriteLog.WriteErrorLog($"Error Executing FuncLib.Events.CreateNewOcassion(BunifuTextBox1.Text, BunifuDropdown1.SelectedItem.ToString, BunifuDatePicker1.Value.Date) : {CreateINIRecord}")
                Exit Sub
            End If
        Else
            BunifuSnackbar1.Show(Me, $"You Cannot Create Function Before Today :{DateTime.Now.Date}", Bunifu.UI.WinForms.BunifuSnackbar.MessageTypes.Warning, 3000)
            Exit Sub
        End If
    End Sub
    Private Sub RemoveKeyFromSection(section As IniSection, keyName As String)
        If section IsNot Nothing Then
            Dim key As IniKey = section.Keys(keyName)
            If key IsNot Nothing Then
                section.Keys.Remove(key)
            End If
        End If
    End Sub
    Private Sub ClearsequenceAfterCreatingFunction()
        GroupBox12.Visible = False
        GroupBox12.Enabled = False
        GroupBox1.Visible = False
        GroupBox1.Enabled = False
        BunifuDropdown11.Items.Clear()
        BunifuDropdown11.Text = Nothing
        BunifuTextBox10.Clear()
        BunifuDataGridView1.Visible = False
        BunifuDataGridView1.Enabled = False
        BunifuDataGridView1.DataSource = Nothing
        BunifuButton5.Visible = False
        BunifuButton5.Enabled = False
        BunifuButton2.Visible = False
        BunifuButton2.Enabled = False
    End Sub
    Private Sub BunifuDropdown7_SelectedIndexChanged(sender As Object, e As EventArgs) Handles BunifuDropdown7.SelectedIndexChanged
        If BunifuDropdown7.SelectedItem = "From Internal" Then
            CheckBox3.Checked = False
            CheckBox3.Enabled = True
            CheckBox3.Visible = True
            BackgroundWorker1.WorkerSupportsCancellation = True
            BackgroundWorker1.RunWorkerAsync()
        ElseIf BunifuDropdown7.SelectedItem = "From External" Then
            CheckBox3.Checked = False
            CheckBox3.Enabled = False
            CheckBox3.Visible = False
            BackgroundWorker2.WorkerSupportsCancellation = True
            BackgroundWorker2.RunWorkerAsync()
        End If
    End Sub
    Private UnidentifiedEmployees As New List(Of String)
    Private Sub BackgroundWorker1_DoWork(sender As Object, e As DoWorkEventArgs) Handles BackgroundWorker1.DoWork
        Dim worker As BackgroundWorker = CType(sender, BackgroundWorker)
        GroupBox1.Invoke(Sub()
                             Try
                                 ClearSequenceaftersourceselected()
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
                                     If BunifuDataGridView1.Columns.Contains("Select") Then
                                         BunifuDataGridView1.Columns.Remove(checkBoxColumn)
                                     End If
                                     BunifuDataGridView1.Columns.Insert(0, checkBoxColumn)
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
                                         GroupBox12.Visible = True
                                         GroupBox12.Enabled = True
                                         BunifuDropdown11.Enabled = True
                                         BunifuDropdown11.Visible = True
                                         BunifuDropdown11.Items.Clear()
                                         BunifuDropdown11.Text = Nothing
                                         BunifuTextBox10.Clear()
                                         BunifuTextBox10.Visible = True
                                         BunifuTextBox10.Enabled = True
                                         GroupBox1.Visible = True
                                         GroupBox1.Enabled = True
                                         BunifuDataGridView1.Visible = True
                                         BunifuDataGridView1.Enabled = True
                                         BunifuDataGridView1.DataSource = grid
                                         BunifuButton5.Visible = True
                                         BunifuButton5.Enabled = True
                                         BunifuDropdown11.Items.Clear()
                                         BunifuDropdown11.Text = Nothing
                                     For Each column As DataGridViewColumn In BunifuDataGridView1.Columns
                                         BunifuDropdown11.Items.Add(column.HeaderText)
                                     Next
                                 End If
                             Catch ex As Exception
                                 BunifuSnackbar1.Show(Me, $"Error Occured While Adding Data From Internal Database.{Environment.NewLine}Error :{ex.Message}", Bunifu.UI.WinForms.BunifuSnackbar.MessageTypes.Error, 2000)
                                 FuncLib.WriteLog.WriteErrorLog($"Error Occured While Executing FuncLib.Events.FetchInternalDbForNewVariableOccasion : {ex.Message}")
                                 FuncLib.Events.DeleteDataTableForConstantFunction(BunifuTextBox1.Text)
                                 e.Cancel = True
                             End Try
                         End Sub)
    End Sub
    Private Sub BackgroundWorker1_RunWorkerCompleted(sender As Object, e As RunWorkerCompletedEventArgs) Handles BackgroundWorker1.RunWorkerCompleted
        If e.Cancelled = True Then
            Exit Sub
        ElseIf e.Error IsNot Nothing Then
            BunifuSnackbar1.Show(Me, $"Error Occured While Adding Data From Internal Database.{Environment.NewLine}Error :{e.Error.ToString}", Bunifu.UI.WinForms.BunifuSnackbar.MessageTypes.Error, 2000)
            FuncLib.WriteLog.WriteErrorLog($"Error Occured While Executing FuncLib.Events.FetchInternalDbForNewVariableOccasion : {e.Error.ToString}")
            FuncLib.Events.DeleteDataTableForConstantFunction(BunifuTextBox1.Text)
            ClearSequenceaftersourceselected()
        Else
            BunifuSnackbar1.Show(Me, $"All Avialable Data Succesfully Fetched!", Bunifu.UI.WinForms.BunifuSnackbar.MessageTypes.Success, 2000)
        End If
    End Sub
    Private Sub BackgroundWorker2_DoWork(sender As Object, e As DoWorkEventArgs) Handles BackgroundWorker2.DoWork
        Dim worker As BackgroundWorker = CType(sender, BackgroundWorker)
        GroupBox1.Invoke(Sub()
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
                                     If BunifuDataGridView1.Columns.Contains("Select") Then
                                         BunifuDataGridView1.Columns.Remove("Select")
                                     End If
                                     GroupBox12.Visible = True
                                     GroupBox12.Enabled = True
                                     BunifuDropdown11.Enabled = True
                                     BunifuDropdown11.Visible = True
                                     BunifuDropdown11.Items.Clear()
                                     BunifuDropdown11.Text = Nothing
                                     BunifuTextBox10.Clear()
                                     BunifuTextBox10.Visible = True
                                     BunifuTextBox10.Enabled = True
                                     GroupBox1.Visible = True
                                     GroupBox1.Enabled = True
                                     BunifuDataGridView1.Visible = True
                                     BunifuDataGridView1.Enabled = True
                                     BunifuDataGridView1.DataSource = Nothing
                                     BunifuButton5.Visible = True
                                     BunifuButton5.Enabled = True
                                     BunifuDropdown11.Items.Clear()
                                     BunifuDropdown11.Text = Nothing
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
                                         BunifuDataGridView1.DataSource = dt
                                     Next
                                     BunifuDropdown11.Items.Clear()
                                     BunifuDropdown11.Text = Nothing
                                     For Each column As DataGridViewColumn In BunifuDataGridView1.Columns
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
                                 FuncLib.Events.DeleteDataTableForConstantFunction(BunifuTextBox1.Text)
                                 e.Cancel = True
                             End Try
                         End Sub)
    End Sub
    Private Sub BackgroundWorker2_RunWorkerCompleted(sender As Object, e As RunWorkerCompletedEventArgs) Handles BackgroundWorker2.RunWorkerCompleted
        If e.Cancelled = True Then
            Exit Sub
        ElseIf e.Error IsNot Nothing Then
            BunifuSnackbar1.Show(Me, $"Error Occured While Adding Data From Internal Database.{Environment.NewLine}Error :{e.Error.ToString}", Bunifu.UI.WinForms.BunifuSnackbar.MessageTypes.Error, 2000)
            FuncLib.WriteLog.WriteErrorLog($"Error Occured While Executing FuncLib.Events.FetchInternalDbForNewVariableOccasion : {e.Error.ToString}")
            FuncLib.Events.DeleteDataTableForConstantFunction(BunifuTextBox1.Text)
        Else
            BunifuSnackbar1.Show(Me, $"All Avialable Data Succesfully Fetched!", Bunifu.UI.WinForms.BunifuSnackbar.MessageTypes.Success, 2000)
        End If
    End Sub
    Private Sub ClearSequenceaftersourceselected()
        GroupBox12.Visible = False
        GroupBox12.Enabled = False
        BunifuDropdown11.Items.Clear()
        BunifuDropdown11.Text = Nothing
        BunifuDropdown11.Visible = False
        BunifuDropdown11.Enabled = False
        BunifuTextBox10.Clear()
        BunifuTextBox10.Enabled = False
        GroupBox1.Visible = False
        GroupBox1.Enabled = False
        BunifuDataGridView1.Visible = False
        BunifuDataGridView1.Enabled = False
        BunifuDataGridView1.DataSource = Nothing
        BunifuButton5.Visible = False
        BunifuButton5.Enabled = False
    End Sub
    Private Sub BunifuTextBox10_TextChanged(sender As Object, e As EventArgs) Handles BunifuTextBox10.TextChanged
        Dim selectedColumnIndex As Integer = BunifuDropdown11.SelectedIndex
        Dim columnName As String = If(selectedColumnIndex >= 0, BunifuDropdown11.Items(selectedColumnIndex).ToString(), "")
        If BunifuDataGridView1.DataSource IsNot Nothing AndAlso Not String.IsNullOrEmpty(columnName) Then
            Dim bindingSource As New BindingSource()
            bindingSource.DataSource = BunifuDataGridView1.DataSource
            bindingSource.Filter = $"[{columnName}] LIKE '%{BunifuTextBox10.Text}%'"
            BunifuDataGridView1.DataSource = bindingSource
        End If
    End Sub
    Private selectedRows As New List(Of DataGridViewRow)
    Private Sub BunifuDataGridView1_CellContentClick(sender As Object, e As DataGridViewCellEventArgs) Handles BunifuDataGridView1.CellContentClick
        If BunifuDropdown7.SelectedItem = "From Internal" Then
            If e.ColumnIndex = BunifuDataGridView1.Columns("Select").Index AndAlso e.RowIndex >= 0 Then
                Dim checkboxCell As DataGridViewCheckBoxCell = TryCast(BunifuDataGridView1.Rows(e.RowIndex).Cells("Select"), DataGridViewCheckBoxCell)
                If checkboxCell IsNot Nothing Then
                    Dim row As DataGridViewRow = BunifuDataGridView1.Rows(e.RowIndex)
                    Dim isChecked As Boolean = Convert.ToBoolean(checkboxCell.Value)
                    If BunifuDataGridView1.Rows(e.RowIndex).Cells("Select").Value = isChecked Then
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
    Private Sub CheckBox3_CheckedChanged(sender As Object, e As EventArgs) Handles CheckBox3.CheckedChanged
        Dim isChecked As Boolean = CheckBox3.Checked
        For i As Integer = 0 To BunifuDataGridView1.RowCount - 2
            Dim checkboxCell As DataGridViewCheckBoxCell = TryCast(BunifuDataGridView1.Rows(i).Cells("Select"), DataGridViewCheckBoxCell)
            If checkboxCell IsNot Nothing Then
                checkboxCell.Value = isChecked
            End If
            Dim rows As DataGridViewRow = BunifuDataGridView1.Rows(i)
            If CheckBox3.Checked = True Then
                selectedRows.Add(rows)
            Else
                selectedRows.Remove(rows)
            End If
        Next
    End Sub
    Private Sub BunifuButton5_Click(sender As Object, e As EventArgs) Handles BunifuButton5.Click
        If BunifuDropdown7.SelectedItem = "From Internal" Then
            BackgroundWorker3.WorkerSupportsCancellation = True
            BackgroundWorker3.RunWorkerAsync()
        ElseIf BunifuDropdown7.SelectedItem = "From External" Then
            BackgroundWorker4.WorkerSupportsCancellation = True
            BackgroundWorker4.RunWorkerAsync()
        End If
    End Sub
    Private Sub BackgroundWorker3_DoWork(sender As Object, e As DoWorkEventArgs) Handles BackgroundWorker3.DoWork
        Dim worker As BackgroundWorker = CType(sender, BackgroundWorker)
        GroupBox1.Invoke(Sub()
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
                                         FuncLib.Events.ImportDataIntoVariableFunction(BunifuTextBox1.Text, Employee_ID, NT_ID, Scanner_5D, Scanner_ID, Employee_Name, Gender, Celebration_Date, Buisness_Area, Functions, Leave_Approver, Created_At)
                                         CurrentRow += 1
                                         ProgressBar1.Value = CurrentRow
                                     Catch ex As Exception
                                         BunifuSnackbar1.Show(Me, $"Error Occured While Executing FuncLib.Events.ImportDataIntoVariableFunction. {Environment.NewLine} Error : {ex.Message}", Bunifu.UI.WinForms.BunifuSnackbar.MessageTypes.Success, 3000)
                                         FuncLib.WriteLog.WriteErrorLog($"Error Occured While Executing FuncLib.Events.ImportDataIntoVariableFunction :  {ex.Message}")
                                         FuncLib.Events.DeleteDataTableForConstantFunction(BunifuTextBox1.Text)
                                         e.Cancel = True
                                         Exit For
                                     End Try
                                 Next
                             Else
                                 BunifuSnackbar1.Show(Me, $"No Employee Selected For Import!", Bunifu.UI.WinForms.BunifuSnackbar.MessageTypes.Information, 3000)
                                 e.Cancel = True
                                 Exit Sub
                             End If
                         End Sub)
    End Sub
    Private Sub BackgroundWorker3_RunWorkerCompleted(sender As Object, e As RunWorkerCompletedEventArgs) Handles BackgroundWorker3.RunWorkerCompleted
        If e.Cancelled = True Then
            Exit Sub
        ElseIf e.Error IsNot Nothing Then
            BunifuSnackbar1.Show(Me, $"Error Occured While Adding Data From Internal Database.{Environment.NewLine}Error :{e.Error.ToString}", Bunifu.UI.WinForms.BunifuSnackbar.MessageTypes.Error, 2000)
            FuncLib.WriteLog.WriteErrorLog($"Error Occured While Executing FuncLib.Events.FetchInternalDbForNewVariableOccasion : {e.Error.ToString}")
            FuncLib.Events.DeleteDataTableForConstantFunction(BunifuTextBox1.Text)
            Exit Sub
        Else
            BunifuSnackbar1.Show(Me, $"All Avialable Data Succesfully Fetched!", Bunifu.UI.WinForms.BunifuSnackbar.MessageTypes.Success, 2000)
            ClearSequenceaftersourceselected()
            RebbotInitiaizer.BackColor = Color.LimeGreen
            RebbotInitiaizer.TextBox1.Text = 5
            RebbotInitiaizer.ShowDialog()
        End If
    End Sub
    Private Sub BackgroundWorker4_DoWork(sender As Object, e As DoWorkEventArgs) Handles BackgroundWorker4.DoWork
        Dim worker As BackgroundWorker = CType(sender, BackgroundWorker)
        GroupBox1.Invoke(Sub()
                             If BunifuDataGridView1.Rows.Count > 0 Then
                                 ProgressBar1.Minimum = 0
                                 ProgressBar1.Maximum = BunifuDataGridView1.Rows.Count - 1
                                 ProgressBar1.Value = 0
                                 Dim CurrentRow As Integer = 0
                                 For i As Integer = 0 To BunifuDataGridView1.RowCount - 2
                                     Dim row As DataGridViewRow = BunifuDataGridView1.Rows(i)
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
                                         FuncLib.Events.ImportDataIntoVariableFunction(BunifuTextBox1.Text, Employee_ID, NT_ID, Scanner_5D, Scanner_ID, Employee_Name, Gender, Celebration_Date, Buisness_Area, Functions, Leave_Approver, Created_At)
                                         CurrentRow += 1
                                         ProgressBar1.Value = CurrentRow
                                     Catch ex As Exception
                                         BunifuSnackbar1.Show(Me, $"Error Occured While Executing FuncLib.Events.ImportDataIntoVariableFunction. {Environment.NewLine} Error : {ex.Message}", Bunifu.UI.WinForms.BunifuSnackbar.MessageTypes.Success, 3000)
                                         FuncLib.WriteLog.WriteErrorLog($"Error Occured While Executing FuncLib.Events.ImportDataIntoVariableFunction :  {ex.Message}")
                                         FuncLib.Events.DeleteDataTableForConstantFunction(BunifuTextBox1.Text)
                                         e.Cancel = True
                                         Exit For
                                     End Try
                                 Next
                             End If
                         End Sub)
    End Sub
    Private Sub BackgroundWorker4_RunWorkerCompleted(sender As Object, e As RunWorkerCompletedEventArgs) Handles BackgroundWorker4.RunWorkerCompleted
        If e.Cancelled = True Then
            Exit Sub
        ElseIf e.Error IsNot Nothing Then
            BunifuSnackbar1.Show(Me, $"Error Occured While Adding Data From Internal Database.{Environment.NewLine}Error :{e.Error.ToString}", Bunifu.UI.WinForms.BunifuSnackbar.MessageTypes.Error, 2000)
            FuncLib.WriteLog.WriteErrorLog($"Error Occured While Executing FuncLib.Events.FetchInternalDbForNewVariableOccasion : {e.Error.ToString}")
            FuncLib.Events.DeleteDataTableForConstantFunction(BunifuTextBox1.Text)
            Exit Sub
        Else
            BunifuSnackbar1.Show(Me, $"All Avialable Data Succesfully Fetched!", Bunifu.UI.WinForms.BunifuSnackbar.MessageTypes.Success, 2000)
            ClearSequenceaftersourceselected()
            RebbotInitiaizer.BackColor = Color.LimeGreen
            RebbotInitiaizer.TextBox1.Text = 5
            RebbotInitiaizer.ShowDialog()
        End If
    End Sub
    Private Sub BunifuButton2_Click(sender As Object, e As EventArgs) Handles BunifuButton2.Click
        If CheckBox1.Checked = True AndAlso BunifuDropdown1.SelectedItem = "Only This Time" Then
            FuncLib.Events.DeleteDataTableForConstantFunction(BunifuTextBox1.Text)
            transition.HideSync(Me, False, BunifuAnimatorNS.Animation.ScaleAndHorizSlide)
        ElseIf CheckBox1.Checked = True AndAlso BunifuDropdown1.SelectedItem = "Repeat Every Year" Then
            FuncLib.Events.DeleteDataTableForVariableFunction(BunifuTextBox1.Text)
            transition.HideSync(Me, False, BunifuAnimatorNS.Animation.ScaleAndHorizSlide)
        Else
            transition.HideSync(Me, False, BunifuAnimatorNS.Animation.ScaleAndHorizSlide)
        End If
        Me.Close()
    End Sub
End Class