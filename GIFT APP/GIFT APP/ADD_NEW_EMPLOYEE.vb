Imports System.ComponentModel
Imports System.Configuration
Imports System.IO.Ports
Imports System.Threading
Imports Bunifu.UI.WinForms
Imports BunifuAnimatorNS
Imports MadMilkman.Ini

Public Class ADD_NEW_EMPLOYEE
    Private WithEvents serialPort As New SerialPort()
    Dim ports As String() = SerialPort.GetPortNames()
    Dim Ini As New IniFile
    Dim Inipath As String = ConfigurationManager.AppSettings("IniFile")
    Dim currentHostName As String = System.Net.Dns.GetHostName()
    Dim IssueGiftAccess As Boolean = False
    Dim AddEventsAccess As Boolean = False
    Dim Administrativeaccess As Boolean = False
    Dim transition As New BunifuTransition
    Private Sub ClearSequenceonInitialize()
        If TextBox1.Text = "ADD NEW EMPLOYEE" Then
            BunifuTextBox1.Clear()
            BunifuTextBox1.ReadOnly = True
            BunifuTextBox2.Clear()
            BunifuTextBox2.ReadOnly = True
            BunifuLabel3.Visible = False
            BunifuTextBox3.Clear()
            BunifuTextBox3.Visible = False
            BunifuLabel4.Visible = False
            BunifuTextBox4.Clear()
            BunifuTextBox4.Visible = False
            BunifuLabel5.Visible = False
            BunifuTextBox5.Clear()
            BunifuTextBox5.Visible = False
            BunifuLabel6.Visible = False
            BunifuDropdown1.Items.Clear()
            BunifuDropdown1.Text = Nothing
            BunifuDropdown1.Visible = False
            BunifuLabel7.Visible = False
            BunifuDatePicker1.Visible = False
            BunifuLabel8.Visible = False
            BunifuTextBox8.Clear()
            BunifuTextBox8.Visible = False
            BunifuLabel9.Visible = False
            BunifuTextBox9.Clear()
            BunifuTextBox9.Visible = False
            BunifuLabel10.Visible = False
            BunifuTextBox10.Clear()
            BunifuTextBox10.Visible = False
            Guna2CheckBox1.Checked = False
            Guna2CheckBox1.Visible = False
            Guna2CheckBox2.Checked = False
            Guna2CheckBox2.Visible = False
            Guna2CheckBox3.Checked = False
            Guna2CheckBox3.Visible = False
            BunifuButton1.Visible = False
            BunifuButton2.Visible = False
        ElseIf TextBox1.Text = "UPDATE/MODIFY EMPLOYEES" Then
            BunifuTextBox1.ReadOnly = True
            BunifuTextBox2.ReadOnly = True
            BunifuTextBox3.ReadOnly = True
            BunifuTextBox4.ReadOnly = True
            BunifuTextBox5.ReadOnly = True
            BunifuDropdown1.Enabled = False
            Guna2CheckBox1.Visible = False
            Guna2CheckBox2.Visible = False
            Guna2CheckBox3.Visible = False
        End If
    End Sub
    Private Sub ADD_NEW_EMPLOYEE_Load(sender As Object, e As EventArgs) Handles MyBase.Load
        If TextBox1.Text = "ADD NEW EMPLOYEE" Then
            Ini.Load(Inipath)
            ClearSequenceonInitialize()
            Try
                If serialPort.PortName <> Ini.Sections("SERIAL").Keys($"{currentHostName}").Value AndAlso serialPort.IsOpen = False Then
                    serialPort.PortName = Ini.Sections("SERIAL").Keys($"{currentHostName}").Value
                    serialPort.BaudRate = 9600
                    serialPort.Parity = Parity.None
                    serialPort.StopBits = StopBits.One
                    If serialPort.IsOpen = False Then
                        serialPort.Open()
                        serialPort.Write("START")
                    Else
                        serialPort.Close()
                        Thread.Sleep(1000)
                        serialPort.Open()
                    End If
                Else
                    serialPort.BaudRate = 9600
                    serialPort.Parity = Parity.None
                    serialPort.StopBits = StopBits.One
                    If serialPort.IsOpen = False Then
                        serialPort.Open()
                    Else
                        serialPort.Close()
                        System.Threading.Thread.Sleep(1000)
                        serialPort.Open()
                        serialPort.Write("START")
                    End If
                End If
            Catch ex As Exception
                BunifuSnackbar1.Show(Me, $"Error Occured While Trying To Open Comport!{Environment.NewLine}Error:{ex.Message}", Bunifu.UI.WinForms.BunifuSnackbar.MessageTypes.Error, 2500)
                FuncLib.WriteLog.WriteErrorLog($"Error Occured While Tring To Open COmport While Sahring Gift :{ex.Message}")
                BunifuButton2.Enabled = True
                BunifuButton2.Visible = True
                Exit Sub
            End Try
        ElseIf TextBox1.Text = "UPDATE/MODIFY EMPLOYEES" Then
            ClearSequenceonInitialize()
        End If

    End Sub
    Private Sub SerialPort_DataReceived(sender As Object, e As SerialDataReceivedEventArgs) Handles serialPort.DataReceived
        If TextBox1.Text = "ADD NEW EMPLOYEE" Then
            Dim hexData As String = serialPort.ReadLine()
            Dim extractedString As String = hexData.Substring(1, hexData.Length - 2)
            BunifuTextBox1.Invoke(Sub()
                                      Dim ScannerID As String = FuncLib.Gifts.IdentifyEmployee(extractedString, "BY 10D")
                                      Dim CheckEmployeeexistance As String = FuncLib.Setup.CheckifEmployeeExistforaddnewemp(ScannerID)
                                      If CheckEmployeeexistance = "True" Then
                                          BunifuSnackbar1.Show(Me, $"Employee Is Already Present Inside The Database", Bunifu.UI.WinForms.BunifuSnackbar.MessageTypes.Information, 2500)
                                          ClearSequenceonInitialize()
                                          BunifuButton2.Enabled = True
                                          BunifuButton2.Visible = True
                                      ElseIf CheckEmployeeexistance = "False" Then
                                          BunifuTextBox2.Text = FuncLib.Gifts.ConvertESDNumberToCardNumber(ScannerID)
                                          BunifuTextBox1.Text = ScannerID
                                          ActionAfterEmployeeFound()
                                      Else
                                          BunifuSnackbar1.Show(Me, $"Error occured While Checking Employee Database.{Environment.NewLine}Error : {CheckEmployeeexistance}", Bunifu.UI.WinForms.BunifuSnackbar.MessageTypes.Error, 2500)
                                          FuncLib.WriteLog.WriteErrorLog($"Error Occureed While Executing FuncLib.Setup.CheckifEmployeeExistforaddnewemp:{CheckEmployeeexistance}")
                                          ClearSequenceonInitialize()
                                          BunifuButton2.Enabled = True
                                          BunifuButton2.Visible = True
                                      End If
                                  End Sub)
            Thread.Sleep(1000)
            serialPort.DiscardInBuffer()
            serialPort.DiscardOutBuffer()
        End If
    End Sub
    Private Sub ActionAfterEmployeeFound()
        If TextBox1.Text = "ADD NEW EMPLOYEE" Then
            BunifuLabel3.Visible = True
            BunifuTextBox3.Clear()
            BunifuTextBox3.Visible = True
            BunifuLabel4.Visible = True
            BunifuTextBox4.Clear()
            BunifuTextBox4.Visible = True
            BunifuLabel5.Visible = True
            BunifuTextBox5.Clear()
            BunifuTextBox5.Visible = True
            BunifuLabel6.Visible = True
            BunifuDropdown1.Items.Clear()
            BunifuDropdown1.Text = Nothing
            BunifuDropdown1.Items.Add("M")
            BunifuDropdown1.Items.Add("F")
            BunifuDropdown1.Visible = True
            BunifuLabel7.Visible = True
            BunifuDatePicker1.Visible = True
            BunifuLabel8.Visible = True
            BunifuTextBox8.Clear()
            BunifuTextBox8.Visible = True
            BunifuLabel9.Visible = True
            BunifuTextBox9.Clear()
            BunifuTextBox9.Visible = True
            BunifuLabel10.Visible = True
            BunifuTextBox10.Clear()
            BunifuTextBox10.Visible = True
            Guna2CheckBox1.Checked = False
            Guna2CheckBox1.Visible = True
            Guna2CheckBox2.Checked = False
            Guna2CheckBox2.Visible = True
            Guna2CheckBox3.Checked = False
            Guna2CheckBox3.Visible = True
            BunifuButton1.Visible = True
            BunifuButton2.Visible = True
        End If
    End Sub
    Private Sub Guna2CheckBox1_CheckedChanged(sender As Object, e As EventArgs) Handles Guna2CheckBox1.CheckedChanged
        If TextBox1.Text = "ADD NEW EMPLOYEE" Then
            If Guna2CheckBox1.Checked = True Then
                IssueGiftAccess = True
            ElseIf Guna2CheckBox1.Checked = False Then
                IssueGiftAccess = False
            End If
        End If
    End Sub
    Private Sub Guna2CheckBox2_CheckedChanged(sender As Object, e As EventArgs) Handles Guna2CheckBox2.CheckedChanged
        If TextBox1.Text = "ADD NEW EMPLOYEE" Then
            If Guna2CheckBox2.Checked = True Then
                AddEventsAccess = True
            ElseIf Guna2CheckBox2.Checked = False Then
                AddEventsAccess = False
            End If
        End If
    End Sub
    Private Sub Guna2CheckBox3_CheckedChanged(sender As Object, e As EventArgs) Handles Guna2CheckBox3.CheckedChanged
        If TextBox1.Text = "ADD NEW EMPLOYEE" Then
            If Guna2CheckBox3.Checked = True Then
                Administrativeaccess = True
            ElseIf Guna2CheckBox3.Checked = False Then
                Administrativeaccess = False
            End If
        End If
    End Sub
    Public Function CheckTextBoxesForEmpty(ByVal textBoxes As IEnumerable(Of BunifuTextBox)) As Boolean
        For Each textBox As BunifuTextBox In textBoxes
            If String.IsNullOrEmpty(textBox.Text.Trim()) Then
                Return True
            End If
        Next
        Return False
    End Function
    Private Sub BunifuButton1_Click(sender As Object, e As EventArgs) Handles BunifuButton1.Click
        If TextBox1.Text = "ADD NEW EMPLOYEE" Then
            Dim textBoxes As New List(Of BunifuTextBox) From {BunifuTextBox1, BunifuTextBox2, BunifuTextBox3, BunifuTextBox4, BunifuTextBox5, BunifuTextBox8, BunifuTextBox9, BunifuTextBox10}
            If CheckTextBoxesForEmpty(textBoxes) = True Then
                BunifuSnackbar1.Show(Me, $"One Of The Required Field Is Empty. Required Field Must Be Filled First!", Bunifu.UI.WinForms.BunifuSnackbar.MessageTypes.Error, 2500)
                Exit Sub
            ElseIf CheckTextBoxesForEmpty(textBoxes) = False Then
                If BunifuDropdown1.SelectedItem = Nothing Then
                    BunifuSnackbar1.Show(Me, $"Gender Not Selected For The Employee.{Environment.NewLine}. Please Select The Gender First!", Bunifu.UI.WinForms.BunifuSnackbar.MessageTypes.Error, 2500)
                    Exit Sub
                Else
                    Dim Employee_ID As String = BunifuTextBox3.Text
                    Dim NT_ID As String = BunifuTextBox4.Text
                    Dim Scanner_ID As String = BunifuTextBox1.Text
                    Dim Scanner_5D As String = BunifuTextBox2.Text
                    Dim Employee_Name As String = BunifuTextBox5.Text
                    Dim Gender As String = BunifuDropdown1.SelectedItem
                    Dim Celebration_Date As DateTime = BunifuDatePicker1.Value.Date
                    Dim Buisness_Area As String = BunifuTextBox8.Text
                    Dim Functions As String = BunifuTextBox9.Text
                    Dim Leave_Approver As String = BunifuTextBox10.Text
                    If IssueGiftAccess = True OrElse AddEventsAccess = True OrElse Administrativeaccess = True Then
                        Dim AddEmployeetoUserDatabase As String = FuncLib.Setup.AddNewEmployeeToUserDatabase(NT_ID, Scanner_ID, Scanner_5D, Employee_Name, Gender, Buisness_Area, Functions, Leave_Approver, IssueGiftAccess, AddEventsAccess, Administrativeaccess)
                        If AddEmployeetoUserDatabase = "True" Then
                            Dim AddEmployeeToDatabase As String = FuncLib.Setup.AddNewEmployeeToEmpDB(Employee_ID, NT_ID, Scanner_ID, Scanner_5D, Employee_Name, Gender, Celebration_Date, Buisness_Area, Functions, Leave_Approver)
                            If AddEmployeeToDatabase = "True" Then
                                BunifuSnackbar1.Show(Me, $"{Employee_Name} Has been successfull Added to Database!", Bunifu.UI.WinForms.BunifuSnackbar.MessageTypes.Success, 2500)
                                ClearSequenceonInitialize()
                                Exit Sub
                            Else
                                BunifuSnackbar1.Show(Me, $"Error Occured While Trying To Add The Person In User Database.{Environment.NewLine}Error:{AddEmployeeToDatabase}", Bunifu.UI.WinForms.BunifuSnackbar.MessageTypes.Error, 2500)
                                FuncLib.WriteLog.WriteErrorLog($"Error Occured While Executing FuncLib.Setup.AddNewEmployeeToEmpDB :{AddEmployeeToDatabase}")
                                Exit Sub
                            End If
                        Else
                            BunifuSnackbar1.Show(Me, $"Error Occured While Trying To Add The Person In User Database.{Environment.NewLine}Error:{AddEmployeetoUserDatabase}", Bunifu.UI.WinForms.BunifuSnackbar.MessageTypes.Error, 2500)
                            FuncLib.WriteLog.WriteErrorLog($"Error Occured While Executing FuncLib.Setup.AddNewEmployeeToUserDatabase :{AddEmployeetoUserDatabase}")
                            Exit Sub
                        End If
                    Else
                        Dim AddEmployeeToDatabase As String = FuncLib.Setup.AddNewEmployeeToEmpDB(Employee_ID, NT_ID, Scanner_ID, Scanner_5D, Employee_Name, Gender, Celebration_Date, Buisness_Area, Functions, Leave_Approver)
                        If AddEmployeeToDatabase = "True" Then
                            BunifuSnackbar1.Show(Me, $"{Employee_Name} Has been successfull Added to Database!", Bunifu.UI.WinForms.BunifuSnackbar.MessageTypes.Success, 2500)
                            Exit Sub
                        Else
                            BunifuSnackbar1.Show(Me, $"Error Occured While Trying To Add The Person In User Database.{Environment.NewLine}Error:{AddEmployeeToDatabase}", Bunifu.UI.WinForms.BunifuSnackbar.MessageTypes.Error, 2500)
                            FuncLib.WriteLog.WriteErrorLog($"Error Occured While Executing FuncLib.Setup.AddNewEmployeeToEmpDB :{AddEmployeeToDatabase}")
                            Exit Sub
                        End If
                    End If
                End If
            End If
        ElseIf TextBox1.Text = "UPDATE/MODIFY EMPLOYEES" Then
            Dim textBoxes As New List(Of BunifuTextBox) From {BunifuTextBox1, BunifuTextBox2, BunifuTextBox3, BunifuTextBox5, BunifuTextBox8, BunifuTextBox9, BunifuTextBox10}
            If CheckTextBoxesForEmpty(textBoxes) = True Then
                BunifuSnackbar1.Show(Me, $"One Of The Required Field Is Empty. Required Field Must Be Filled First!", Bunifu.UI.WinForms.BunifuSnackbar.MessageTypes.Error, 2500)
                Exit Sub
            Else

            End If
        End If
    End Sub
    Private Sub ADD_NEW_EMPLOYEE_Closing(sender As Object, e As CancelEventArgs) Handles Me.Closing
        If TextBox1.Text = "ADD NEW EMPLOYEE" Then
            If serialPort.IsOpen Then
                serialPort.Close()
            End If
        End If
    End Sub
    Private Sub BunifuButton23_Click(sender As Object, e As EventArgs) Handles BunifuButton23.Click
        transition.HideSync(Me, True, BunifuAnimatorNS.Animation.ScaleAndHorizSlide)
        SETUP.Enabled = True
        Me.Close()
    End Sub
    Private Sub BunifuButton2_Click(sender As Object, e As EventArgs) Handles BunifuButton2.Click
        ClearSequenceonInitialize()
    End Sub
End Class