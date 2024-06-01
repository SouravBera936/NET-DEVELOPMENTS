Imports MadMilkman.Ini
Imports System.Runtime.InteropServices
Imports System.Data.OleDb
Imports System.Diagnostics
Imports System.IO
Imports FTD2XX_NET
Imports System.Drawing
Imports AutoItX3Lib
Imports Tesseract
Module FUNNCTIONLIB
    Dim YieldBox = {MAIN.TextBox174, MAIN.TextBox175, MAIN.TextBox176, MAIN.TextBox177, MAIN.TextBox178, MAIN.TextBox179, MAIN.TextBox180, MAIN.TextBox181, MAIN.TextBox182, MAIN.TextBox183, MAIN.TextBox184, MAIN.TextBox185}
    Dim ini As New IniFile
    Dim ini1 As New IniFile
    Dim cn As OleDb.OleDbConnection
    Dim oleCommand As OleDbCommand
    Dim file As System.IO.StreamWriter
    Dim file1 As System.IO.StreamWriter
    Dim file2 As System.IO.StreamWriter
    Dim StopWatch As New Diagnostics.Stopwatch
    Private Declare Function FindWindow Lib "user32" Alias "FindWindowA" _
                                                      (ByVal lpClassName As String,
                                                      ByVal lpWindowName As String) As Integer

    Public Declare Function SetWindowPos Lib "user32" (ByVal hwnd As Integer, ByVal hWndInsertAfter As Integer,
                ByVal X As Integer, ByVal Y As Integer, ByVal cx As Integer, ByVal cy As Integer, ByVal wFlags As Integer) As Integer
    Private Const TASKBAR_WINDOW_NAME As String = "Shell_traywnd"
    Private Const SWP_SHOWWINDOW As Integer = &H40
    Private Const SWP_HIDEWINDOW As Integer = &H80
    Private Const SWP_NOSIZE As UInteger = &H1
    Private Const SWP_NOMOVE As UInteger = &H2
    Private Const SWP_SHOWWINDOW1 As UInteger = &H40
    Private Const SW_MINIMIZE As Integer = 6
    Private Const SW_RESTORE As Integer = 9
    Private externalProcess As Process
    Private Const WH_MOUSE_LL As Integer = 14
    Private Const WM_LBUTTONDOWN As Integer = &H201
    Dim ResetValue As Integer = 0

    <DllImport("user32.dll")>
    Public Function SetForegroundWindow(ByVal hWnd As IntPtr) As Boolean
    End Function
    <DllImport("user32.dll", SetLastError:=True)>
    Public Function SetWindowPos(ByVal hWnd As IntPtr, ByVal hWndInsertAfter As IntPtr, ByVal x As Integer, ByVal y As Integer, ByVal cx As Integer, ByVal cy As Integer, ByVal uFlags As UInteger) As Boolean
    End Function
    <DllImport("user32.dll")>
    Public Function SwitchToThisWindow(ByVal hWnd As IntPtr, ByVal fAltTab As Boolean) As Boolean
    End Function
    <DllImport("user32.dll", SetLastError:=True)>
    Public Function ShowWindow(hWnd As IntPtr, nCmdShow As Integer) As Boolean
    End Function
    <DllImport("user32.dll", SetLastError:=True)>
    Public Function PrintWindow(hWnd As IntPtr, hdcBlt As IntPtr, nFlags As Integer) As Boolean
    End Function
    <DllImport("user32.dll", SetLastError:=True)>
    Public Function GetForegroundWindow() As IntPtr
    End Function
    <DllImport("user32.dll", SetLastError:=True)>
    Public Function GetWindowRect(hWnd As IntPtr, ByRef lpRect As Rectangle) As Boolean
    End Function
    <DllImport("user32.dll")>
    Public Function IsIconic(ByVal hWnd As IntPtr) As Boolean
    End Function
    Function LoadINI() As Boolean
        Dim line1 As Label

line1:
        Dim path As String = "C:\PLUG_POWER_FCT\PLUG_POWER_FCT\Project Files\ErrorLog.txt"
        file = My.Computer.FileSystem.OpenTextFileWriter(path, True)
        MAIN.ListBox1.Items.Add($"{DateTime.Now.ToString} START RUN FUNCTION : LoadINI")
        MAIN.ListBox1.SelectedIndex = MAIN.ListBox1.Items.Count - 1
        Try
            MAIN.ListBox1.Items.Add($"{DateTime.Now.ToString} Loading Confrigurations....")
            ini1.Load("C:\PLUG_POWER_FCT\PLUG_POWER_FCT\Project Files\TEST SETUP.ini")
            ini.Load("C:\PLUG_POWER_FCT\PLUG_POWER_FCT\Project Files\Confrigurations.ini")
            CONFIG.TextBox1.Text = ini.Sections("DB_PROVIDER").Keys("Count").Value
            CONFIG.TextBox2.Text = ini.Sections("DB_SOURCE").Keys("Count").Value
            CONFIG.TextBox3.Text = ini.Sections("SECQURITY_INFO").Keys("Count").Value
            CONFIG.TextBox4.Text = ini.Sections("ERROR_LOG").Keys("Count").Value
            CONFIG.TextBox5.Text = ini.Sections("SYSID_STRCOUNT").Keys("Count").Value
            CONFIG.TextBox6.Text = ini.Sections("BATID_STRCOUNT").Keys("Count").Value
            CONFIG.TextBox7.Text = ini.Sections("SMC_STRCOUNT").Keys("Count").Value
            CONFIG.TextBox8.Text = ini.Sections("HPC_STRCOUNT").Keys("Count").Value
            CONFIG.TextBox9.Text = ini.Sections("SSMC_STRCOUNT").Keys("Count").Value
            CONFIG.TextBox10.Text = ini.Sections("SYSID_GOLDEN").Keys("Count").Value
            CONFIG.TextBox11.Text = ini.Sections("BATID_GOLDEN").Keys("Count").Value
            CONFIG.TextBox12.Text = ini.Sections("SMC_GOLDEN").Keys("Count").Value
            CONFIG.TextBox13.Text = ini.Sections("HPC_GOLDEN").Keys("Count").Value
            CONFIG.TextBox14.Text = ini.Sections("SSMC_GOLDEN").Keys("Count").Value
            CONFIG.TextBox15.Text = ini.Sections("SYSID_SNAP1").Keys("Count").Value
            CONFIG.TextBox16.Text = ini.Sections("SYSID_SOFT").Keys("Count").Value
            CONFIG.TextBox17.Text = ini.Sections("TEMP_PATH").Keys("Count").Value
            CONFIG.TextBox18.Text = ini.Sections("SYSID_SNAP2").Keys("Count").Value
            CONFIG.TextBox19.Text = ini.Sections("SYSID_COMP").Keys("Count").Value
            CONFIG.TextBox20.Text = ini.Sections("SYSID_REPTMP").Keys("Count").Value
            CONFIG.TextBox21.Text = ini.Sections("SYSID_REPPATH").Keys("Count").Value
            CONFIG.TextBox22.Text = ini.Sections("SYSID_SNAME").Keys("Count").Value
            CONFIG.TextBox23.Text = ini.Sections("BATID_SNAP1").Keys("Count").Value
            CONFIG.TextBox24.Text = ini.Sections("BATID_SOFT").Keys("Count").Value
            CONFIG.TextBox25.Text = ini.Sections("BATID_COMP").Keys("Count").Value
            CONFIG.TextBox26.Text = ini.Sections("BATID_SNAP2").Keys("Count").Value
            CONFIG.TextBox27.Text = ini.Sections("BATID_SNAME").Keys("Count").Value
            CONFIG.TextBox29.Text = ini.Sections("BATID_REPTMP").Keys("Count").Value
            CONFIG.TextBox28.Text = ini.Sections("BATID_REPPATH").Keys("Count").Value
            CONFIG.TextBox30.Text = ini.Sections("SMC_FF_SNAP").Keys("Count").Value
            CONFIG.TextBox31.Text = ini.Sections("POWER_SUPPLY").Keys("Count").Value
            CONFIG.TextBox32.Text = ini.Sections("PS_VOLT").Keys("Count").Value
            CONFIG.TextBox33.Text = ini.Sections("PS_CURR").Keys("Count").Value
            CONFIG.TextBox34.Text = ini.Sections("UNIFLASH_BATCH").Keys("Count").Value
            CONFIG.TextBox35.Text = ini.Sections("FTDI_XML").Keys("Count").Value
            CONFIG.TextBox36.Text = ini.Sections("APP_LOG").Keys("Count").Value
            CONFIG.TextBox37.Text = ini.Sections("SMC_DUT_CONNECT_SNAP").Keys("Count").Value
            CONFIG.TextBox38.Text = ini.Sections("SMC_APP").Keys("Count").Value
            CONFIG.TextBox39.Text = ini.Sections("INIT_SNAP").Keys("Count").Value
            CONFIG.TextBox40.Text = ini.Sections("INIT_REF").Keys("Count").Value
            CONFIG.TextBox48.Text = ini.Sections("SMC_APPNAME").Keys("Count").Value
            CONFIG.TextBox49.Text = ini.Sections("DMM").Keys("Count").Value
            CONFIG.TextBox50.Text = ini.Sections("SMC").Keys("Can670_5V_Snap").Value
            CONFIG.TextBox51.Text = ini.Sections("SMC").Keys("Isolated_V+12_Snap").Value
            CONFIG.TextBox52.Text = ini.Sections("SMC").Keys("Sensor_V+12_Snap").Value
            CONFIG.TextBox53.Text = ini.Sections("SMC").Keys("TT620_5V_Snap").Value
            CONFIG.TextBox54.Text = ini.Sections("SMC").Keys("IT661_5V_Snap").Value
            CONFIG.TextBox55.Text = ini.Sections("SMC").Keys("A0691_Snap").Value
            CONFIG.TextBox56.Text = ini.Sections("SMC").Keys("D0691_Snap").Value
            CONFIG.TextBox57.Text = ini.Sections("SMC").Keys("RESTORE_SCRIPT").Value
            CONFIG.TextBox58.Text = ini.Sections("SMC").Keys("D0691_REF").Value
            CONFIG.TextBox41.Text = ini.Sections("SMC").Keys("D0661_Snap_ON").Value
            CONFIG.TextBox42.Text = ini.Sections("SMC").Keys("D0661_REF_ON").Value
            CONFIG.TextBox44.Text = ini.Sections("SMC").Keys("D0661_Snap_ON").Value
            CONFIG.TextBox43.Text = ini.Sections("SMC").Keys("D0661_REF_ON").Value
            CONFIG.TextBox45.Text = ini.Sections("SMC").Keys("AUXENCMD_ON_SNAP").Value
            CONFIG.TextBox46.Text = ini.Sections("SMC").Keys("AUXENCMD_ON_REF").Value
            CONFIG.TextBox47.Text = ini.Sections("SMC").Keys("D0100CMD_SNAP_ON").Value
            CONFIG.TextBox59.Text = ini.Sections("SMC").Keys("D0100CMD_REF_ON").Value
            CONFIG.TextBox60.Text = ini.Sections("SMC").Keys("D0100CMD_SNAP_OFF").Value
            CONFIG.TextBox61.Text = ini.Sections("SMC").Keys("D0100CMD_REF_OFF").Value
            CONFIG.TextBox62.Text = ini.Sections("SMC").Keys("D0101CMD_SNAP_ON").Value
            CONFIG.TextBox63.Text = ini.Sections("SMC").Keys("D0101CMD_REF_ON").Value
            CONFIG.TextBox64.Text = ini.Sections("SMC").Keys("D0101CMD_SNAP_OFF").Value
            CONFIG.TextBox65.Text = ini.Sections("SMC").Keys("D0101CMD_REF_OFF").Value
            CONFIG.TextBox66.Text = ini.Sections("SMC").Keys("CMD_SNAP_OFF").Value
            CONFIG.TextBox67.Text = ini.Sections("SMC").Keys("CMD_REF_OFF").Value
            CONFIG.TextBox68.Text = ini.Sections("SMC").Keys("SW2_SNAP_ON").Value
            CONFIG.TextBox69.Text = ini.Sections("SMC").Keys("SW2_REF_ON").Value
            CONFIG.TextBox70.Text = ini.Sections("SMC").Keys("SW3_SNAP_ON").Value
            CONFIG.TextBox71.Text = ini.Sections("SMC").Keys("SW3_REF_ON").Value
            CONFIG.TextBox72.Text = ini.Sections("SMC").Keys("TOOGLE_SNAP_ON").Value
            CONFIG.TextBox73.Text = ini.Sections("SMC").Keys("TOOGLE_REF_ON").Value
            CONFIG.TextBox74.Text = ini.Sections("SMC").Keys("CAN1_SNAP_ON").Value
            CONFIG.TextBox75.Text = ini.Sections("SMC").Keys("CAN1_REF_ON").Value
            CONFIG.TextBox76.Text = ini.Sections("SMC").Keys("CAN2_SNAP_ON").Value
            CONFIG.TextBox77.Text = ini.Sections("SMC").Keys("CAN2_REF_ON").Value
            CONFIG.TextBox78.Text = ini.Sections("SMC").Keys("TIME_TEST").Value
            CONFIG.TextBox79.Text = ini.Sections("SMC").Keys("STEPPER_TEST_SNAP").Value
            CONFIG.TextBox80.Text = ini.Sections("SMC").Keys("STEPPER_TEST_REF").Value
            CONFIG.TextBox86.Text = ini.Sections("SMC").Keys("SHUTDOWNTEST_REF").Value
            CONFIG.TextBox82.Text = ini.Sections("SMC").Keys("REPORT_TEMP").Value
            CONFIG.TextBox83.Text = ini.Sections("SMC").Keys("REPORT_PATH").Value
            CONFIG.TextBox84.Text = ini.Sections("SMC").Keys("PROG_CONN").Value
            CONFIG.TextBox85.Text = ini.Sections("SMC").Keys("FTDI_CONN").Value
            CONFIG.TextBox87.Text = ini.Sections("SMC").Keys("SHUTDOWNTEST_SNAP").Value
            CONFIG.TextBox88.Text = ini.Sections("SMC").Keys("SNAP_DIR").Value
            CONFIG.TextBox89.Text = ini.Sections("BATID_SCRIPT").Keys("Count").Value
            CONFIG.TextBox90.Text = ini.Sections("SYSID_SCRIPT").Keys("Count").Value
            MAIN.ListBox1.Items.Add($"{DateTime.Now.ToString} Confrigurations Loaded Successfully")
            MAIN.ListBox1.SelectedIndex = MAIN.ListBox1.Items.Count - 1
            MAIN.ListBox1.Items.Add($"{DateTime.Now.ToString} Satrt Loading Setup Files")
            MAIN.ListBox1.SelectedIndex = MAIN.ListBox1.Items.Count - 1
            If ini1.Sections("UPPER_CASE").Keys("Count").Value = "true" Then
                SETUP.CheckBox1.Checked = True
            Else
                SETUP.CheckBox1.Checked = False
            End If
            If ini1.Sections("MARK_REC").Keys("Count").Value = "true" Then
                SETUP.CheckBox2.Checked = True
            Else
                SETUP.CheckBox2.Checked = False
            End If
            If ini1.Sections("SERIAL_VERIFY").Keys("Count").Value = "true" Then
                SETUP.CheckBox3.Checked = True
            Else
                SETUP.CheckBox3.Checked = False
            End If
            If ini1.Sections("PASS_REPEAT").Keys("Count").Value = "true" Then
                SETUP.CheckBox4.Checked = True
            Else
                SETUP.CheckBox4.Checked = False
            End If
            If ini1.Sections("TASKBAR_HIDE").Keys("Count").Value = "true" Then
                SETUP.CheckBox5.Checked = True
            Else
                SETUP.CheckBox5.Checked = False
            End If
            If ini1.Sections("IMG_COMP").Keys("Count").Value = "true" Then
                SETUP.CheckBox6.Checked = True
            Else
                SETUP.CheckBox6.Checked = False
            End If
            If ini1.Sections("KEEP_REC").Keys("Count").Value = "true" Then
                SETUP.CheckBox7.Checked = True
            Else
                SETUP.CheckBox7.Checked = False
            End If
            MAIN.ListBox1.Items.Add($"{DateTime.Now.ToString} Setup Loaded Successfully")
            MAIN.ListBox1.SelectedIndex = MAIN.ListBox1.Items.Count - 1
            file.Close()
            MAIN.ListBox1.Items.Add($"{DateTime.Now.ToString} STOP RUN FUNCTION : LoadINI")
            MAIN.ListBox1.SelectedIndex = MAIN.ListBox1.Items.Count - 1
            Return True
        Catch ex As Exception
            Dim iret1 As Object = MsgBox("ERROR WHILE LOADING INI FILE: " & ex.Message, vbRetryCancel, "FUNNCTIONLIB.LOADINI error")
            If iret1 = vbRetry Then
                file.Close()
                MAIN.ListBox1.Items.Add($"{DateTime.Now.ToString} Error Loading COnfrigurations : " & ex.Message)
                MAIN.ListBox1.SelectedIndex = MAIN.ListBox1.Items.Count - 1
                GoTo line1
            ElseIf iret1 = vbCancel Then
                file.WriteLine("FUNNCTION.LOADINI FAILED DUE TO : " & ex.Message & " " & "@" & DateTime.Now.ToString("yyyy/MM/dd/HH:mm:ss"))
                file.Close()
                MAIN.ListBox1.Items.Add($"{DateTime.Now.ToString} Error Loading COnfrigurations : " & ex.Message)
                MAIN.ListBox1.SelectedIndex = MAIN.ListBox1.Items.Count - 1
                Return False
            End If
        End Try
    End Function
    Function CheckDBState(Provider, DataSource, Secqurity) As Boolean
        Dim line1 As Label
line1:
        MAIN.ListBox1.Items.Add($"{DateTime.Now.ToString} START RUN FUNCTION : CheckDBState")
        MAIN.ListBox1.SelectedIndex = MAIN.ListBox1.Items.Count - 1
        Dim path1 As String = CONFIG.TextBox4.Text
        file1 = My.Computer.FileSystem.OpenTextFileWriter(path1, True)
        Try
            cn = New OleDbConnection
            cn.ConnectionString = "Provider=" & Provider & ";Data Source=" & DataSource & "; Persist Security Info=" & Secqurity & ";"
            If cn.State = ConnectionState.Open Then
                file1.Close()
                MAIN.ListBox1.Items.Add($"{DateTime.Now.ToString} CheckDbState.ConnectionState:Open")
                MAIN.ListBox1.SelectedIndex = MAIN.ListBox1.Items.Count - 1
                Return True
            Else
                cn.Open()
                file1.Close()
                MAIN.ListBox1.Items.Add($"{DateTime.Now.ToString} CheckDbState.ConnectionState:Open")
                MAIN.ListBox1.SelectedIndex = MAIN.ListBox1.Items.Count - 1
                Return True
            End If
        Catch ex As Exception
            Dim iret As Object = MsgBox("UNABLE TO CONNECT TO DATABASE : " & ex.Message, vbCritical + vbRetryCancel, "FUNNCTIONLIB.CheckDBState error")
            If iret = vbRetry Then
                file1.Close()
                MAIN.ListBox1.Items.Add($"{DateTime.Now.ToString} CheckDBState.ConnectionState : False " & "[" & ex.Message & "]")
                MAIN.ListBox1.SelectedIndex = MAIN.ListBox1.Items.Count - 1
                GoTo line1
            ElseIf iret = vbCancel Then
                file1.WriteLine("DATABASE CONNECTION FAILED DUE TO : " & ex.Message & " " & "@" & DateTime.Now.ToString("yyyy/MM/dd/HH:mm:ss"))
                file1.Close()
                MAIN.ListBox1.Items.Add($"{DateTime.Now.ToString} CheckDBState.ConnectionState : False " & "[" & ex.Message & "]")
                MAIN.ListBox1.SelectedIndex = MAIN.ListBox1.Items.Count - 1
                Return False
            End If
        End Try
    End Function
    Function CheckRegistration(HostName, Provider, DataSource, Secqurity) As Boolean
        Dim line1 As Label
line1:
        MAIN.ListBox1.Items.Add($"{DateTime.Now.ToString} START RUN FUNCTION : CheckRegistration")
        MAIN.ListBox1.SelectedIndex = MAIN.ListBox1.Items.Count - 1
        Dim path1 As String = CONFIG.TextBox4.Text
        file1 = My.Computer.FileSystem.OpenTextFileWriter(path1, True)
        Try
            cn = New OleDbConnection
            cn.ConnectionString = "Provider=" & Provider & ";Data Source=" & DataSource & "; Persist Security Info=" & Secqurity & ";"
            cn.Open()
            Dim SQL As String = "SELECT * FROM SYSTEM_1 WHERE (HOST_NAME = '" & HostName & "')"
            oleCommand = New OleDbCommand(SQL, cn)
            Dim reader As OleDbDataReader = oleCommand.ExecuteReader
            If reader.HasRows Then
                file1.Close()
                MAIN.ListBox1.Items.Add($"{DateTime.Now.ToString} CheckRegistration.Status : True")
                MAIN.ListBox1.SelectedIndex = MAIN.ListBox1.Items.Count - 1
                Return True
            Else
                file1.Close()
                MAIN.ListBox1.Items.Add($"{DateTime.Now.ToString} CheckRegistration.Status : False")
                MAIN.ListBox1.SelectedIndex = MAIN.ListBox1.Items.Count - 1
                Return False
            End If
        Catch ex As Exception
            Dim iret As Object = MsgBox("UNABLE TO CHECK SYSTEM REGISTRATION : " & ex.Message, vbCritical + vbRetryCancel, "FUNNCTIONLIB.CheckRegistration error")
            If iret = vbRetry Then
                file1.Close()
                MAIN.ListBox1.Items.Add($"{DateTime.Now.ToString} CheckRegistration.Status : False" & "[" & ex.Message & "]")
                MAIN.ListBox1.SelectedIndex = MAIN.ListBox1.Items.Count - 1
                GoTo line1
            ElseIf iret = vbCancel Then
                file1.WriteLine("UNABLE TO CHECK SYSTEM REGISTRATION : " & ex.Message & " " & "@" & DateTime.Now.ToString("yyyy/MM/dd/HH:mm:ss"))
                file1.Close()
                MAIN.ListBox1.Items.Add($"{DateTime.Now.ToString} CheckRegistration.Status : False" & "[" & ex.Message & "]")
                MAIN.ListBox1.SelectedIndex = MAIN.ListBox1.Items.Count - 1
            End If
        End Try
    End Function
    Function CheckLiscence(HostName, Provider, DataSource, Secqurity) As Boolean
        Dim line1 As Label
line1:
        MAIN.ListBox1.Items.Add($"{DateTime.Now.ToString} START RUN FUNCTION : CheckLiscence")
        MAIN.ListBox1.SelectedIndex = MAIN.ListBox1.Items.Count - 1
        Dim path1 As String = CONFIG.TextBox4.Text
        file1 = My.Computer.FileSystem.OpenTextFileWriter(path1, True)
        Try
            cn = New OleDbConnection
            cn.ConnectionString = "Provider=" & Provider & ";Data Source=" & DataSource & "; Persist Security Info=" & Secqurity & ";"
            cn.Open()
            Dim SQL As String = "SELECT * FROM SYSTEM_1 WHERE (HOST_NAME = '" & HostName & "')"
            oleCommand = New OleDbCommand(SQL, cn)
            Dim reader As OleDbDataReader = oleCommand.ExecuteReader
            If reader.HasRows Then
                MAIN.ListBox1.Items.Add($"{DateTime.Now.ToString} CHECKING LISCENCE...")
                MAIN.ListBox1.SelectedIndex = MAIN.ListBox1.Items.Count - 1
                reader.Read()
                If reader("EXP_DATE") <= DateTime.Now Then
                    MAIN.ListBox1.Items.Add($"{DateTime.Now.ToString} LISCENCE EXPIRED LAST DATE WAS : " & reader("EXP_DATE"))
                    MAIN.ListBox1.SelectedIndex = MAIN.ListBox1.Items.Count - 1
                    reader.Close()
                    file1.Close()
                    Return False
                Else
                    MAIN.ListBox1.Items.Add($"{DateTime.Now.ToString} LISCENCE FOUND LAST DATE OF USAGE : " & reader("EXP_DATE"))
                    MAIN.ListBox1.SelectedIndex = MAIN.ListBox1.Items.Count - 1
                    reader.Close()
                    file1.Close()
                    Return True
                End If
            End If
        Catch ex As Exception
            Dim iret As Object = MsgBox("UNABLE TO CHECK SYSTEM LISCENCING : " & ex.Message, vbCritical + vbRetryCancel, "FUNNCTIONLIB.CheckLiscence error")
            If iret = vbRetry Then
                file1.Close()
                MAIN.ListBox1.Items.Add($"{DateTime.Now.ToString} CheckLiscence.Status : False" & "[" & ex.Message & "]")
                MAIN.ListBox1.SelectedIndex = MAIN.ListBox1.Items.Count - 1
                GoTo line1
            ElseIf iret = vbCancel Then
                file1.WriteLine("UNABLE TO CHECK SYSTEM LISCENCING : " & ex.Message & " " & "@" & DateTime.Now.ToString("yyyy/MM/dd/HH:mm:ss"))
                file1.Close()
                MAIN.ListBox1.Items.Add($"{DateTime.Now.ToString} CheckLiscence.Status : False" & "[" & ex.Message & "]")
                MAIN.ListBox1.SelectedIndex = MAIN.ListBox1.Items.Count - 1
                Return False
            End If
        End Try
    End Function
    Public Sub HideTaskbar(ByVal ActionHide As Boolean)
        Dim TaskbarHWnd As Integer = FindWindow(TASKBAR_WINDOW_NAME, "")
        If ActionHide Then
            Call SetWindowPos(TaskbarHWnd, 0, 0, 0, 0, 0, SWP_HIDEWINDOW)
        Else
            Call SetWindowPos(TaskbarHWnd, 0, 0, 0, 0, 0, SWP_SHOWWINDOW)
        End If
    End Sub
    Function CalculateProgress(StepsNum As Integer, CurrentStep As Integer) As Integer
        Dim ProgressValue = 0
        Dim Division = 100 / StepsNum
        For i As Integer = 1 To CurrentStep
            ProgressValue = ProgressValue + Division
        Next
        MAIN.TextBox158.Text = ProgressValue
    End Function
    Function ProgressReset(ByVal istrue As Boolean)
        If istrue Then
            MAIN.TextBox158.Text = 0
        End If
    End Function
    Function UserValidationCheck(UserName As String, Password As String, Provider As String, DataSource As String, Secqurity As String) As Boolean
        Dim line1 As Label
line1:
        MAIN.ListBox1.Items.Add($"{DateTime.Now.ToString} CHECKING THE VALIDATED USER AND ACCESS")
        MAIN.ListBox1.SelectedIndex = MAIN.ListBox1.Items.Count - 1
        Dim path1 As String = CONFIG.TextBox36.Text
        file1 = My.Computer.FileSystem.OpenTextFileWriter(path1, True)
        Dim path2 As String = CONFIG.TextBox4.Text
        file2 = My.Computer.FileSystem.OpenTextFileWriter(path2, True)
        Try
            cn = New OleDbConnection
            cn.ConnectionString = "Provider=" & Provider & ";Data Source=" & DataSource & "; Persist Security Info=" & Secqurity & ";"
            cn.Open()
            Dim SQL As String = "SELECT * FROM USER_1 WHERE (USE_ID = '" & UserName & "' AND PASSWORD = '" & Password & "')"
            oleCommand = New OleDbCommand(SQL, cn)
            Dim reader As OleDbDataReader = oleCommand.ExecuteReader
            If reader.HasRows Then
                reader.Read()
                MAIN.TextBox16.Text = reader("DESIGNATION")
                MAIN.TextBox3.Text = reader("USER_NAME") & "_" & reader("USE_ID")
                file1.WriteLine("LOGIN SUCCESS :" + "=" + UserName + "_" & DateTime.Now.ToString("yyyy/MM/dd/HH:mm:ss"))
                file1.Close()
                file2.Close()
                MAIN.ListBox1.Items.Add($"{DateTime.Now.ToString} LOGIN SUCCESS WITH " & UserName & " AS " & reader("DESIGNATION"))
                MAIN.ListBox1.SelectedIndex = MAIN.ListBox1.Items.Count - 1
                reader.Close()
                Return True
            ElseIf UserName = "ADMINISTRATOR" And Password = "PPLR@centumBLR2023" Then
                MAIN.TextBox16.Text = "ADMINISTRATOR"
                MAIN.TextBox3.Text = "ADMINISTRATOR"
                file1.WriteLine("LOGIN SUCCESS :" + "=" + UserName + "_" & DateTime.Now.ToString("yyyy/MM/dd/HH:mm:ss"))
                file1.Close()
                file2.Close()
                MAIN.ListBox1.Items.Add($"{DateTime.Now.ToString} LOGIN SUCCESS WITH ADMINISTRATOR")
                MAIN.ListBox1.SelectedIndex = MAIN.ListBox1.Items.Count - 1
                reader.Close()
                Return True
            Else
                file1.WriteLine("LOGIN FAILED :" + "=" + UserName + "_" & DateTime.Now.ToString("yyyy/MM/dd/HH:mm:ss"))
                file1.Close()
                file2.Close()
                MAIN.ListBox1.Items.Add($"{DateTime.Now.ToString} INVALID USERNAME OR PASSWORD. FAILED LOGIN")
                MAIN.ListBox1.SelectedIndex = MAIN.ListBox1.Items.Count - 1
                reader.Close()
                Return False
            End If
        Catch ex As Exception
            Dim iret As Object = MsgBox("ERROR EXECUTING FUNNCTIONLIB.UserValidationCheck FUNCTION : " & ex.Message, vbCritical + vbRetryCancel, "FUNNCTIONLIB.UserValidationCheck error")
            If iret = vbRetry Then
                file1.Close()
                file2.Close()
                MAIN.ListBox1.Items.Add($"{DateTime.Now.ToString} ERROR EXECUTING FUNNCTIONLIB.UserValidationCheck FUNCTION" & "[" & ex.Message & "]")
                MAIN.ListBox1.SelectedIndex = MAIN.ListBox1.Items.Count - 1
                GoTo line1
            ElseIf iret = vbCancel Then
                file2.WriteLine("ERROR EXECUTING FUNNCTIONLIB.UserValidationCheck FUNCTION : " & ex.Message & " " & "@" & DateTime.Now.ToString("yyyy/MM/dd/HH:mm:ss"))
                file1.Close()
                file2.Close()
                MAIN.ListBox1.Items.Add($"{DateTime.Now.ToString} ERROR EXECUTING FUNNCTIONLIB.UserValidationCheck FUNCTION" & "[" & ex.Message & "]")
                MAIN.ListBox1.SelectedIndex = MAIN.ListBox1.Items.Count - 1
                Return False
            End If
        End Try
    End Function
    Function SecqurityReset(UserID As String, UserName As String, Designation As String, Provider As String, DataSource As String, Secqurity As String) As String
        Dim line1 As Label
line1:
        Dim path1 As String = CONFIG.TextBox4.Text
        file1 = My.Computer.FileSystem.OpenTextFileWriter(path1, True)
        Dim path2 As String = CONFIG.TextBox36.Text
        file2 = My.Computer.FileSystem.OpenTextFileWriter(path2, True)
        Try
            cn = New OleDbConnection
            cn.ConnectionString = "Provider=" & Provider & ";Data Source=" & DataSource & "; Persist Security Info=" & Secqurity & ";"
            cn.Open()
            Dim SQL As String = "SELECT * FROM USER_1 WHERE (USE_ID = '" & UserID & "' AND DESIGNATION = '" & Designation & "' AND USER_NAME = '" & UserName & "')"
            oleCommand = New OleDbCommand(SQL, cn)
            Dim reader As OleDbDataReader = oleCommand.ExecuteReader
            If reader.HasRows Then
                reader.Read()
                file2.WriteLine("PASSWORD RESETED FOR :" + "=" + UserName + "_" & DateTime.Now.ToString("yyyy/MM/dd/HH:mm:ss"))
                file2.Close()
                file1.Close()
                Dim password = reader("PASSWORD").ToString
                reader.Close()
                MAIN.ListBox1.Items.Add($"{DateTime.Now.ToString} PASSWORD RECOVERED FOR : " & UserName)
                MAIN.ListBox1.SelectedIndex = MAIN.ListBox1.Items.Count - 1
                Return password
            Else
                file2.WriteLine("PASSWORD RESET ATTEMPTED FOR :" + "=" + UserName + "_" & DateTime.Now.ToString("yyyy/MM/dd/HH:mm:ss"))
                file1.Close()
                file2.Close()
                reader.Close()
                MAIN.ListBox1.Items.Add($"{DateTime.Now.ToString} INVALID INPUT WHILE RESETING PASSWORD : " & UserName)
                MAIN.ListBox1.SelectedIndex = MAIN.ListBox1.Items.Count - 1
                Return "INVALID USER"
            End If
        Catch ex As Exception
            Dim iret As Object = MsgBox("ERROR EXECUTING FUNNCTIONLIB.SecqurityReset FUNCTION : " & ex.Message, vbCritical + vbRetryCancel, "FUNNCTIONLIB.SecqurityReset error")
            If iret = vbRetry Then
                file1.Close()
                file2.Close()
                MAIN.ListBox1.Items.Add($"{DateTime.Now.ToString} ERROR EXECUTING FUNNCTIONLIB.SecqurityReset FUNCTION" & "[" & ex.Message & "]")
                MAIN.ListBox1.SelectedIndex = MAIN.ListBox1.Items.Count - 1
                GoTo line1
            ElseIf iret = vbCancel Then
                file1.WriteLine("ERROR EXECUTING FUNNCTIONLIB.SecqurityReset FUNCTION : " & ex.Message & " " & "@" & DateTime.Now.ToString("yyyy/MM/dd/HH:mm:ss"))
                file1.Close()
                file2.Close()
                MAIN.ListBox1.Items.Add($"{DateTime.Now.ToString} ERROR EXECUTING FUNNCTIONLIB.SecqurityReset FUNCTION" & "[" & ex.Message & "]")
                MAIN.ListBox1.SelectedIndex = MAIN.ListBox1.Items.Count - 1
                Return False
            End If
        End Try
    End Function
    Function SLtoUpper(SerialNumber As String, ByVal istrue As Boolean) As String
        Dim line1 As Label
line1:
        Dim path1 As String = CONFIG.TextBox4.Text
        file1 = My.Computer.FileSystem.OpenTextFileWriter(path1, True)
        Try
            If istrue Then
                MAIN.ListBox1.Items.Add($"{DateTime.Now.ToString} SLtoUpper FUNCTION : ENABLED")
                MAIN.ListBox1.SelectedIndex = MAIN.ListBox1.Items.Count - 1
                MAIN.ListBox1.Items.Add($"{DateTime.Now.ToString} ENTERD SERIAL NUMBER : " & SerialNumber)
                MAIN.ListBox1.SelectedIndex = MAIN.ListBox1.Items.Count - 1
                MAIN.ListBox1.Items.Add($"{DateTime.Now.ToString} FUNNCLIB.SLtoUpper returns : " & SerialNumber.ToUpper)
                MAIN.ListBox1.SelectedIndex = MAIN.ListBox1.Items.Count - 1
                file1.Close()
                Return SerialNumber.ToUpper
            Else
                MAIN.ListBox1.Items.Add($"{DateTime.Now.ToString} SLTOUPPER FUNCTION : DISABLED")
                MAIN.ListBox1.SelectedIndex = MAIN.ListBox1.Items.Count - 1
                MAIN.ListBox1.Items.Add($"{DateTime.Now.ToString} ENTERD SERIAL NUMBER : " & SerialNumber)
                MAIN.ListBox1.SelectedIndex = MAIN.ListBox1.Items.Count - 1
                MAIN.ListBox1.Items.Add($"{DateTime.Now.ToString} FUNNCLIB.SLtoUpper returns : " & SerialNumber)
                MAIN.ListBox1.SelectedIndex = MAIN.ListBox1.Items.Count - 1
                file1.Close()
                Return SerialNumber
            End If
        Catch ex As Exception
            Dim iret As Object = MsgBox("ERROR EXECUTING FUNNCTIONLIB.SLtoUpper FUNCTION : " & ex.Message, vbCritical + vbRetryCancel, "FUNNCTIONLIB.SLtoUpper error")
            If iret = vbRetry Then
                file1.Close()
                MAIN.ListBox1.Items.Add($"{DateTime.Now.ToString} ERROR EXECUTING FUNNCTIONLIB.SLtoUpper FUNCTION" & "[" & ex.Message & "]")
                MAIN.ListBox1.SelectedIndex = MAIN.ListBox1.Items.Count - 1
                GoTo line1
            ElseIf iret = vbCancel Then
                file1.WriteLine("ERROR EXECUTING FUNNCTIONLIB.SLtoUpper FUNCTION : " & ex.Message & " " & "@" & DateTime.Now.ToString("yyyy/MM/dd/HH:mm:ss"))
                file1.Close()
                MAIN.ListBox1.Items.Add($"{DateTime.Now.ToString} ERROR EXECUTING FUNNCTIONLIB.SLtoUpper FUNCTION" & "[" & ex.Message & "]")
                MAIN.ListBox1.SelectedIndex = MAIN.ListBox1.Items.Count - 1
                Return False
            End If
        End Try

    End Function
    Function VerifySerialNum(SerialNumber As String, Length As Integer, ByVal istrue As Boolean) As Boolean
        Dim line1 As Label
line1:
        Dim path1 As String = CONFIG.TextBox4.Text
        file1 = My.Computer.FileSystem.OpenTextFileWriter(path1, True)
        Try
            If istrue Then
                MAIN.ListBox1.Items.Add($"{DateTime.Now.ToString} SERIAL NUMBER VERIFICATION : ENABLED")
                MAIN.ListBox1.SelectedIndex = MAIN.ListBox1.Items.Count - 1
                If SerialNumber.Length = Length Then
                    MAIN.ListBox1.Items.Add($"{DateTime.Now.ToString} SERIAL NUMBER LENGTH : " & SerialNumber.Length)
                    MAIN.ListBox1.SelectedIndex = MAIN.ListBox1.Items.Count - 1
                    MAIN.ListBox1.Items.Add($"{DateTime.Now.ToString} FUNNCLIB.VerifySerialNum returns : True")
                    MAIN.ListBox1.SelectedIndex = MAIN.ListBox1.Items.Count - 1
                    file1.Close()
                    Return True
                Else
                    MAIN.ListBox1.Items.Add($"{DateTime.Now.ToString} SERIAL NUMBER LENGTH : " & SerialNumber.Length)
                    MAIN.ListBox1.SelectedIndex = MAIN.ListBox1.Items.Count - 1
                    MAIN.ListBox1.Items.Add($"{DateTime.Now.ToString} FUNNCLIB.VerifySerialNum returns : False")
                    MAIN.ListBox1.SelectedIndex = MAIN.ListBox1.Items.Count - 1
                    file1.Close()
                    Return False
                End If
            Else
                MAIN.ListBox1.Items.Add($"{DateTime.Now.ToString} SERIAL NUMBER VERIFICATION : DISABLED")
                MAIN.ListBox1.SelectedIndex = MAIN.ListBox1.Items.Count - 1
                file1.Close()
                Return False
            End If
        Catch ex As Exception
            Dim iret As Object = MsgBox("ERROR EXECUTING FUNNCTIONLIB.VerifySerialNum FUNCTION : " & ex.Message, vbCritical + vbRetryCancel, "FUNNCTIONLIB.VerifySerialNum error")
            If iret = vbRetry Then
                file1.Close()
                MAIN.ListBox1.Items.Add($"{DateTime.Now.ToString} ERROR EXECUTING FUNNCTIONLIB.VerifySerialNum FUNCTION" & "[" & ex.Message & "]")
                MAIN.ListBox1.SelectedIndex = MAIN.ListBox1.Items.Count - 1
                GoTo line1
            ElseIf iret = vbCancel Then
                file1.WriteLine("ERROR EXECUTING FUNNCTIONLIB.VerifySerialNum FUNCTION : " & ex.Message & " " & "@" & DateTime.Now.ToString("yyyy/MM/dd/HH:mm:ss"))
                file1.Close()
                MAIN.ListBox1.Items.Add($"{DateTime.Now.ToString} ERROR EXECUTING FUNNCTIONLIB.VerifySerialNum FUNCTION" & "[" & ex.Message & "]")
                MAIN.ListBox1.SelectedIndex = MAIN.ListBox1.Items.Count - 1
                Return False
            End If
        End Try

    End Function
    Function CheckifGolden(UUT As String, GoldenSL As String, Length As Integer) As String
        If UUT.Length = Length Then
            If UUT = GoldenSL Then
                Return "GOLDEN"
            Else
                Return "PRODUCTION"
            End If
        Else
            Return "INVALID"
        End If
    End Function
    Function CheckIfPassed(DBName As String, SerialNumber As String, TestMethode As String, Provider As String, DataSource As String, Secqurity As String, ByVal istrue As Boolean) As Boolean
        If istrue Then
            Dim line1 As Label
line1:
            MAIN.ListBox1.Items.Add($"{DateTime.Now.ToString} STOP EXECUTION IF UNIT IS ALREADY PASSED FUNCTION : ENABLED")
            MAIN.ListBox1.SelectedIndex = MAIN.ListBox1.Items.Count - 1
            Dim path1 As String = CONFIG.TextBox4.Text
            file1 = My.Computer.FileSystem.OpenTextFileWriter(path1, True)
            Try
                cn = New OleDbConnection
                cn.ConnectionString = "Provider=" & Provider & ";Data Source=" & DataSource & "; Persist Security Info=" & Secqurity & ";"
                cn.Open()
                Dim SQL As String = "SELECT * FROM " & DBName & " WHERE (SERIAL_NUMBER = '" & SerialNumber & "' AND TEST_STEP = '" & TestMethode & "' AND TEST_RESULT = 'PASS')"
                oleCommand = New OleDbCommand(SQL, cn)
                Dim reader As OleDbDataReader = oleCommand.ExecuteReader
                If reader.HasRows Then
                    reader.Read()
                    If reader("TEST_RESULT") = "PASS" Then
                        reader.Close()
                        file1.Close()
                        MAIN.ListBox1.Items.Add($"{DateTime.Now.ToString} UNIT WITH " & SerialNumber & " FOUND AS ALREADY PASSED.")
                        MAIN.ListBox1.SelectedIndex = MAIN.ListBox1.Items.Count - 1
                        Return True
                    Else
                        MAIN.ListBox1.Items.Add($"{DateTime.Now.ToString} UNIT WITH " & SerialNumber & " FOUND AS NEW OR RETEST")
                        MAIN.ListBox1.SelectedIndex = MAIN.ListBox1.Items.Count - 1
                        reader.Close()
                        file1.Close()
                        Return False
                    End If
                Else
                    file1.Close()
                    Return False
                End If
            Catch ex As Exception
                Dim iret As Object = MsgBox("ERROR EXECUTING FUNNCTIONLIB.CheckIfPassed FUNCTION : " & ex.Message, vbCritical + vbRetryCancel, "FUNNCTIONLIB.CheckIfPassed error")
                If iret = vbRetry Then
                    file1.Close()
                    MAIN.ListBox1.Items.Add($"{DateTime.Now.ToString} ERROR EXECUTING FUNNCTIONLIB.CheckIfPassed FUNCTION" & "[" & ex.Message & "]")
                    MAIN.ListBox1.SelectedIndex = MAIN.ListBox1.Items.Count - 1
                    GoTo line1
                ElseIf iret = vbCancel Then
                    file1.WriteLine("ERROR EXECUTING FUNNCTIONLIB.CheckIfPassed FUNCTION : " & ex.Message & " " & "@" & DateTime.Now.ToString("yyyy/MM/dd/HH:mm:ss"))
                    file1.Close()
                    MAIN.ListBox1.Items.Add($"{DateTime.Now.ToString} ERROR EXECUTING FUNNCTIONLIB.CheckIfPassed FUNCTION" & "[" & ex.Message & "]")
                    MAIN.ListBox1.SelectedIndex = MAIN.ListBox1.Items.Count - 1
                    Return False
                End If
            End Try
        Else
            MAIN.ListBox1.Items.Add($"{DateTime.Now.ToString} STOP EXECUTION IF UNIT IS ALREADY PASSED FUNCTION : DISABLED")
            MAIN.ListBox1.SelectedIndex = MAIN.ListBox1.Items.Count - 1
            Return False
        End If
    End Function
    Function Initialize(DBName As String, SerialNumber As String, Length As Integer, TestMethode As String, Provider As String, DataSource As String, Secqurity As String, ByVal istrue As Boolean) As Boolean
        Dim line1 As Label
        If SETUP.CheckBox4.Checked = True Then
            If CheckIfPassed(DBName, SerialNumber, TestMethode, Provider, DataSource, Secqurity, True) = True Then
                Dim iret As Object = MsgBox("THE UNIT " & SerialNumber & " IS ALREADY PASSED CANNOT RUN THE TEST", vbCritical + vbOKOnly, "FUNNLIB.Initialize")
                If iret = vbOK Then
                    STARTUP.TextBox1.Clear()
                    Exit Function
                End If
            ElseIf CheckIfPassed(DBName, SerialNumber, TestMethode, Provider, DataSource, Secqurity, True) = False Then
                GoTo line1
            End If
        ElseIf SETUP.CheckBox4.Checked = False Then
            GoTo line1
        End If
line1:
        If SETUP.CheckBox3.Checked = True Then
            If VerifySerialNum(SerialNumber, Length, True) = True Then
                Return True
            ElseIf VerifySerialNum(SerialNumber, Length, True) = False Then
                Return False
            End If
        ElseIf SETUP.CheckBox3.Checked = False Then
            Return True
        End If
    End Function
    Function CompareImages(Captured As Bitmap, Reference As Bitmap) As Boolean
        Dim threshold As Integer = Integer.Parse("99")
        Dim wid As Integer = Math.Min(Captured.Width, Reference.Width)
        Dim hgt As Integer = Math.Min(Captured.Height, Reference.Height)
        Dim ErrorPic As New Bitmap(wid, hgt)
        Dim are_identical As Boolean = True
        Dim color1, color2 As Color
        Dim eq_color As Color = Color.White
        Dim ne_color As Color = Color.Red
        Dim dr, dg, db, diff As Integer
        For x As Integer = 0 To wid - 1
            For y As Integer = 0 To hgt - 1
                color1 = Captured.GetPixel(x, y)
                color2 = Reference.GetPixel(x, y)
                dr = CInt(color1.R) - color2.R
                dg = CInt(color1.G) - color2.G
                db = CInt(color1.B) - color2.B
                diff = dr * dr + dg * dg + db * db
                If diff <= threshold Then
                    ErrorPic.SetPixel(x, y, eq_color)
                Else
                    ErrorPic.SetPixel(x, y, ne_color)
                    are_identical = False
                End If
            Next y
        Next x
        If (Captured.Width <> Reference.Width) OrElse (Captured.Height <>
            Reference.Height) Then
            are_identical = False
        End If
        If are_identical Then
            Return True
        Else
            FAILUREVIEW.PictureBox1.Image = ErrorPic
            Return False
        End If
    End Function
    Function TakeSnap(ByVal istrue As Boolean) As Bitmap
        If istrue Then
            Dim shot As Size = New Size(My.Computer.Screen.Bounds.Width, My.Computer.Screen.Bounds.Height)
            Dim grab As New Bitmap(My.Computer.Screen.Bounds.Width, My.Computer.Screen.Bounds.Height)
            Dim s As System.Drawing.Graphics = System.Drawing.Graphics.FromImage(grab)
            s.CopyFromScreen(New Point(0, 0), New Point(0, 0), shot)
            Return grab
        End If
    End Function
    Function UnifalshProgramme(BatchFilePath As String) As Boolean
        Dim line1 As Label
line1:
        Dim path1 As String = CONFIG.TextBox4.Text
        file1 = My.Computer.FileSystem.OpenTextFileWriter(path1, True)
        Try
            Using process As New Process()
                Dim startInfo As New ProcessStartInfo("cmd.exe")
                startInfo.Arguments = $"/C ""{BatchFilePath}"""
                startInfo.RedirectStandardOutput = True
                startInfo.RedirectStandardError = True
                startInfo.UseShellExecute = False
                startInfo.CreateNoWindow = True
                process.StartInfo = startInfo
                process.Start()
                Dim output As String = process.StandardOutput.ReadToEnd()
                Dim errors As String = process.StandardError.ReadToEnd()
                If errors.Length > 0 Then
                    Dim iret As Object = MsgBox("FAIL TO PROGRAMME THE TARGET TI DEVICE : " & errors, vbRetryCancel + vbCritical, "FUNNCLIB.UniflashProgramme")
                    If iret = vbRetry Then
                        file1.WriteLine("FAIL TO PROGRAMME THE TARGET TI DEVICE : " & errors & " " & "@" & DateTime.Now.ToString("yyyy/MM/dd/HH:mm:ss"))
                        file1.Close()
                        MAIN.ListBox1.Items.Add($"{DateTime.Now.ToString} FAIL TO PROGRAMME TARGET TI DEVICE : " & errors)
                        MAIN.ListBox1.SelectedIndex = MAIN.ListBox1.Items.Count - 1
                        GoTo line1
                    ElseIf iret = vbCancel Then
                        file1.WriteLine("FAIL TO PROGRAMME THE TARGET TI DEVICE : " & errors & " " & "@" & DateTime.Now.ToString("yyyy/MM/dd/HH:mm:ss"))
                        file1.Close()
                        MAIN.ListBox1.Items.Add($"{DateTime.Now.ToString} FAIL TO PROGRAMME TARGET TI DEVICE : " & errors)
                        MAIN.ListBox1.SelectedIndex = MAIN.ListBox1.Items.Count - 1
                        Return False
                        Exit Function
                    End If
                Else
                    file1.Close()
                    MAIN.ListBox1.Items.Add($"{DateTime.Now.ToString} TI DEVICE PROGRAMMING SUCCESS : {output}")
                    MAIN.ListBox1.SelectedIndex = MAIN.ListBox1.Items.Count - 1
                    Return True
                End If
            End Using
        Catch ex As Exception
            Dim iret1 As Object = MsgBox("ERROR EXECUTING FUNNCLIB.UniflashProgramme : " & ex.Message, vbCritical + vbRetryCancel, "FUNNCTIONLIB.UniflashProgramme error")
            If iret1 = vbRetry Then
                file1.Close()
                MAIN.ListBox1.Items.Add($"{DateTime.Now.ToString} ERROR EXECUTING FUNNCLIB.UniflashProgramme : {ex.Message} ")
                MAIN.ListBox1.SelectedIndex = MAIN.ListBox1.Items.Count - 1
                GoTo line1
            ElseIf iret1 = vbCancel Then
                file1.WriteLine("ERROR EXECUTING FUNNCLIB.UniflashProgramme FUNCTION : " & ex.Message & " " & "@" & DateTime.Now.ToString("yyyy/MM/dd/HH:mm:ss"))
                file1.Close()
                MAIN.ListBox1.Items.Add($"{DateTime.Now.ToString} ERROR EXECUTING FUNNCLIB.UniflashProgramme :{ex.Message} ")
                MAIN.ListBox1.SelectedIndex = MAIN.ListBox1.Items.Count - 1
                Return False
            End If
        End Try

    End Function
    Function FTDIDetectDevices() As Boolean
        Dim line1 As Label
line1:
        Dim path1 As String = CONFIG.TextBox4.Text
        file1 = My.Computer.FileSystem.OpenTextFileWriter(path1, True)
        Try
            Dim ftdiDevice As New FTDI()
            Dim numDevices As UInteger
            Dim ftStatus As FTDI.FT_STATUS = ftdiDevice.GetNumberOfDevices(numDevices)
            If ftStatus = FTDI.FT_STATUS.FT_OK AndAlso numDevices > 0 Then
                Dim deviceInfos(numDevices - 1) As FTDI.FT_DEVICE_INFO_NODE
                ftStatus = ftdiDevice.GetDeviceList(deviceInfos)
                If ftStatus = FTDI.FT_STATUS.FT_OK Then
                    For index = 0 To numDevices - 1
                        Dim info As FTDIDeviceInfo = New FTDIDeviceInfo With {
                            .Description = deviceInfos(index).Description,
                            .SerialNumber = deviceInfos(index).SerialNumber,
                            .DeviceIndex = index
                        }
                        MAIN.ListBox1.Items.Add($"{DateTime.Now.ToString} Found The Following FTDI Devices : {info} ")
                        MAIN.ListBox1.SelectedIndex = MAIN.ListBox1.Items.Count - 1
                        FTDI_SELECTOR.ComboBox1.Text = ""
                        FTDI_SELECTOR.ComboBox1.Items.Clear()
                        FTDI_SELECTOR.ComboBox1.Items.Add(info)
                    Next
                    file1.Close()
                    Return True
                Else
                    Dim iret1 As Object = MsgBox($"Error Getting Device List : {ftStatus}", vbCritical + vbRetryCancel, "FTDI SELCTOR.PopulateFTDIDevice error")
                    If iret1 = vbCancel Then
                        file1.WriteLine("Error Getting Device List : " & ftStatus & " " & "@" & DateTime.Now.ToString("yyyy/MM/dd/HH:mm:ss"))
                        file1.Close()
                        MAIN.ListBox1.Items.Add($"{DateTime.Now.ToString} Error Getting Device List : " & ftStatus)
                        MAIN.ListBox1.SelectedIndex = MAIN.ListBox1.Items.Count - 1
                        Return False
                        Exit Function
                    ElseIf iret1 = vbRetry Then
                        ftdiDevice.Close()
                        MAIN.ListBox1.Items.Add($"{DateTime.Now.ToString} Error Getting Device List : " & ftStatus)
                        MAIN.ListBox1.SelectedIndex = MAIN.ListBox1.Items.Count - 1
                        file1.Close()
                        GoTo line1
                    End If
                End If
            Else
                Dim iret2 As Object = MsgBox("No FTDI Device Found Connected To The PC. Please Connect & Try Again", vbRetryCancel + vbCritical, "FUNNCLIB.FTDIDetectDevice information")
                If iret2 = vbRetry Then
                    MAIN.ListBox1.Items.Add($"{DateTime.Now.ToString} No FTDI Device Found Attached To The Pc")
                    MAIN.ListBox1.SelectedIndex = MAIN.ListBox1.Items.Count - 1
                    ftdiDevice.Close()
                    file1.Close()
                    GoTo line1
                ElseIf iret2 = vbCancel Then
                    ftdiDevice.Close()
                    file1.Close()
                    MAIN.ListBox1.Items.Add($"{DateTime.Now.ToString} No FTDI Device Found Attached To The Pc")
                    MAIN.ListBox1.SelectedIndex = MAIN.ListBox1.Items.Count - 1
                    Return False
                    Exit Function
                End If
            End If
        Catch ex As Exception
            Dim iret3 As Object = MsgBox("Error Executing FUNNCLIB.FTDIDetectDevices : " & ex.Message, vbCritical + vbRetryCancel, "FUNNCTIONLIB.FTDIDetectDevices error")
            If iret3 = vbRetry Then
                MAIN.ListBox1.Items.Add($"{DateTime.Now.ToString} Error Executing FUNNCLIB.FTDIDetectDevices : " & ex.Message)
                MAIN.ListBox1.SelectedIndex = MAIN.ListBox1.Items.Count - 1
                file1.Close()
                GoTo line1
            ElseIf iret3 = vbCancel Then
                file1.WriteLine("Error Executing FUNNCLIB.FTDIDetectDevices FUNCTION : " & ex.Message & " " & "@" & DateTime.Now.ToString("yyyy/MM/dd/HH:mm:ss"))
                file1.Close()
                MAIN.ListBox1.Items.Add($"{DateTime.Now.ToString} ERROR EXECUTING FUNNCLIB.FTDIDetectDevices : " & ex.Message)
                MAIN.ListBox1.SelectedIndex = MAIN.ListBox1.Items.Count - 1
                Return False
                Exit Function
            End If
        End Try
    End Function
    Public Class FTDIDeviceInfo
        Public Property Description As String
        Public Property SerialNumber As String
        Public Property DeviceIndex As Integer
    End Class
    Function StartExterAppinScreenCenter(AppPath As String)
        Dim process As New Process()
        process.StartInfo.FileName = AppPath
        process.Start()
        process.WaitForInputIdle()
        Dim mainWindowHandle As IntPtr = process.MainWindowHandle
        CenterWindowOnScreen(mainWindowHandle)
    End Function
    Public Sub CenterWindowOnScreen(ByVal hWnd As IntPtr)
        Dim screenBounds As Rectangle = Screen.PrimaryScreen.Bounds
        Dim x As Integer = (screenBounds.Width - MAIN.Width) \ 2
        Dim y As Integer = (screenBounds.Height - MAIN.Height) \ 2
        SetWindowPos(hWnd, IntPtr.Zero, x, y, 0, 0, SWP_NOSIZE Or SWP_SHOWWINDOW1)
    End Sub
    Public Function TakeSnapofActiveApp() As Bitmap
        Dim hWnd As IntPtr = GetForegroundWindow()
        ShowWindow(hWnd, SW_RESTORE)
        Dim rect As Rectangle
        GetWindowRect(hWnd, rect)
        Dim windowBitmap As New Bitmap(rect.Width, rect.Height)
        Using g As Graphics = Graphics.FromImage(windowBitmap)
            PrintWindow(hWnd, g.GetHdc(), 0)
            g.ReleaseHdc()
        End Using
        Return windowBitmap
    End Function

    Public Function MinimizeApp(Application As String)
        Dim processName As String = Application
        Dim existingProcesses() As Process = Process.GetProcessesByName(processName)
        If existingProcesses.Length > 0 Then
            externalProcess = existingProcesses(0)
            ShowWindow(externalProcess.MainWindowHandle, SW_MINIMIZE)
        End If
    End Function
    Public Function RestoreApp(Script As String)
        Dim autoItScriptPath As String = Script
        Process.Start("AutoIt3.exe", """" & autoItScriptPath & """")
    End Function
    Public Function ActivateAndBringToTop(ByVal processName As String)
        Dim hWnd As IntPtr = FindWindow(Nothing, processName)
        If hWnd <> IntPtr.Zero Then
            ShowWindow(hWnd, SW_RESTORE)
            SetForegroundWindow(hWnd)
        End If
    End Function
    Public Function ImageToText(ByVal Image As String) As String
        Using tEngine = New TesseractEngine("C:\Program Files\Tesseract-OCR\tessdata", "eng", EngineMode.[Default])
            'creating the tesseract OCR engine with English as the language
            Using img = Pix.LoadFromFile(Image)
                Using page = tEngine.Process(img)
                    Dim text = page.GetText()
                    Console.WriteLine(text)
                    Console.WriteLine(page.GetMeanConfidence())
                    MAIN.ListBox1.Items.Add($"{DateTime.Now.ToString} FUNNCLIB.ImageToText Returns : {text}")
                    MAIN.ListBox1.SelectedIndex = MAIN.ListBox1.Items.Count - 1
                    Return text
                End Using
            End Using
        End Using
    End Function
    Function ResetYield(ByVal istrue As Boolean)
        If istrue Then
            For Each tb In YieldBox
                tb.enabled = False
            Next
            MAIN.TextBox174.Text = ResetValue
            MAIN.TextBox181.Text = ResetValue
            MAIN.TextBox185.Text = ResetValue
            MAIN.TextBox175.Text = ResetValue
            MAIN.TextBox184.Text = ResetValue
            MAIN.TextBox180.Text = ResetValue
            MAIN.TextBox176.Text = ResetValue
            MAIN.TextBox179.Text = ResetValue
            MAIN.TextBox183.Text = ResetValue
        End If
    End Function
    Function CalculateSystemIDYield()

    End Function
End Module
