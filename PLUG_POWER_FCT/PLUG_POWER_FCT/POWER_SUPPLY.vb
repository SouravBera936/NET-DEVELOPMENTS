Imports Ivi.Visa.Interop
Module POWER_SUPPLY
    Dim Port As String = CONFIG.TextBox31.Text.ToString
    Dim SetVoltage As String = CONFIG.TextBox32.Text.ToString
    Dim SetCurrent As String = CONFIG.TextBox33.Text.ToString
    Dim iomgr As New Ivi.Visa.Interop.ResourceManager
    Dim instrany As New Ivi.Visa.Interop.FormattedIO488
    Dim file1 As System.IO.StreamWriter
    Function Initialize() As Boolean
        Dim line1 As Label
line1:
        Dim path1 As String = CONFIG.TextBox4.Text
        file1 = My.Computer.FileSystem.OpenTextFileWriter(path1, True)
        Try
            instrany.IO = iomgr.Open(Port)
            instrany.WriteString("*IDN?")
            Dim PSID As String = instrany.ReadString
            MAIN.ListBox1.Items.Add($"{DateTime.Now.ToString} POWER_SUPPLY.Initialize Returns IDN : " & PSID)
            MAIN.ListBox1.SelectedIndex = MAIN.ListBox1.Items.Count - 1
            file1.Close()
            instrany.IO.Close()
            Return True
        Catch ex As Exception
            Dim iret As Object = MsgBox("Error Initializing Power Supply : " & ex.Message, vbRetryCancel + vbCritical, "POWER_SUPPLY.Initialize()")
            If iret = vbCancel Then
                MAIN.ListBox1.Items.Add($"{DateTime.Now.ToString} POWER_SUPPLY.Initialize Returns Error : " & ex.Message)
                MAIN.ListBox1.SelectedIndex = MAIN.ListBox1.Items.Count - 1
                file1.WriteLine("POWER INITIALIZATION FAILED : " & ex.Message & " " & "@" & DateTime.Now.ToString("yyyy/MM/dd/HH:mm:ss"))
                file1.Close()
                Return False
            ElseIf iret = vbRetry Then
                MAIN.ListBox1.Items.Add($"{DateTime.Now.ToString} POWER_SUPPLY.Initialize Returns Error : " & ex.Message)
                MAIN.ListBox1.SelectedIndex = MAIN.ListBox1.Items.Count - 1
                file1.Close()
                GoTo line1
            End If
        End Try
    End Function
    Function SetLimit() As Boolean
        Dim line1 As Label
line1:
        Dim path1 As String = CONFIG.TextBox4.Text
        file1 = My.Computer.FileSystem.OpenTextFileWriter(path1, True)
        Try
            instrany.IO = iomgr.Open(Port)
            instrany.WriteString(SetVoltage)
            instrany.WriteString(SetCurrent)
            MAIN.ListBox1.Items.Add($"{DateTime.Now.ToString} Set Voltage : {SetVoltage}")
            MAIN.ListBox1.SelectedIndex = MAIN.ListBox1.Items.Count - 1
            MAIN.ListBox1.Items.Add($"{DateTime.Now.ToString} Set Current : {SetCurrent}")
            MAIN.ListBox1.SelectedIndex = MAIN.ListBox1.Items.Count - 1
            file1.Close()
            Return True
        Catch ex As Exception
            Dim iret As Object = MsgBox("Error Setting Power Supply Limit : " & ex.Message, vbRetryCancel + vbCritical, "POWER_SUPPLY.Initialize()")
            If iret = vbCancel Then
                MAIN.ListBox1.Items.Add($"{DateTime.Now.ToString} POWER_SUPPLY.SetLimit Returns Error : " & ex.Message)
                MAIN.ListBox1.SelectedIndex = MAIN.ListBox1.Items.Count - 1
                file1.WriteLine("Error Setting Power Supply Limit : " & ex.Message & " " & "@" & DateTime.Now.ToString("yyyy/MM/dd/HH:mm:ss"))
                file1.Close()
                Return False
            ElseIf iret = vbRetry Then
                MAIN.ListBox1.Items.Add($"{DateTime.Now.ToString} POWER_SUPPLY.SetLimit Returns Error : " & ex.Message)
                MAIN.ListBox1.SelectedIndex = MAIN.ListBox1.Items.Count - 1
                file1.Close()
                GoTo line1
            End If
        End Try

    End Function
    Function PS_ON() As Boolean
        Dim line1 As Label
line1:
        Dim path1 As String = CONFIG.TextBox4.Text
        file1 = My.Computer.FileSystem.OpenTextFileWriter(path1, True)
        Try
            instrany.IO = iomgr.Open(Port)
            instrany.WriteString(SetVoltage)
            instrany.WriteString(SetCurrent)
            instrany.WriteString("OUTP ON")
            file1.Close()
            Return True
        Catch ex As Exception
            Dim iret As Object = MsgBox("POWER_SUPPLY.PS_ON Returns Error : " & ex.Message, vbRetryCancel + vbCritical, "POWER_SUPPLY.Initialize()")
            If iret = vbCancel Then
                MAIN.ListBox1.Items.Add($"{DateTime.Now.ToString} POWER_SUPPLY.PS_ON Returns Error : " & ex.Message)
                MAIN.ListBox1.SelectedIndex = MAIN.ListBox1.Items.Count - 1
                file1.WriteLine("POWER_SUPPLY.PS_ON Returns Error : " & ex.Message & " " & "@" & DateTime.Now.ToString("yyyy/MM/dd/HH:mm:ss"))
                file1.Close()
                Return False
            ElseIf iret = vbRetry Then
                MAIN.ListBox1.Items.Add($"{DateTime.Now.ToString} POWER_SUPPLY.PS_ON Returns Error : " & ex.Message)
                MAIN.ListBox1.SelectedIndex = MAIN.ListBox1.Items.Count - 1
                file1.Close()
                GoTo line1
            End If
        End Try

    End Function
    Function PS_OFF() As Boolean
        Dim line1 As Label
line1:
        Dim path1 As String = CONFIG.TextBox4.Text
        file1 = My.Computer.FileSystem.OpenTextFileWriter(path1, True)
        Try
            instrany.IO = iomgr.Open(Port)
            instrany.WriteString("OUTP OFF")
            file1.Close()
            Return True
        Catch ex As Exception
            Dim iret As Object = MsgBox("ERROR TURNING OUTPUT OFF : " & ex.Message, vbOKOnly + vbCritical, "POWER_SUPPLY.Initialize()")
            If iret = vbOK Then
                MAIN.ListBox1.Items.Add("POWER_SUPPLY.PS_OFF Returns Error : " & ex.Message)
                MAIN.ListBox1.SelectedIndex = MAIN.ListBox1.Items.Count - 1
                file1.WriteLine("ERROR TURNING OUTPUT OFF : " & ex.Message & " " & "@" & DateTime.Now.ToString("yyyy/MM/dd/HH:mm:ss"))
                file1.Close()
                Return False
            End If
        End Try

    End Function
    Function PS_RECYCLE(Time As String) As Boolean
        Dim line1 As Label
line1:
        Try
            PS_OFF()
            MAIN.ListBox1.Items.Add("DelayTime : " & Time)
            MAIN.ListBox1.SelectedIndex = MAIN.ListBox1.Items.Count - 1
            System.Threading.Thread.Sleep(Time)
            PS_ON()
            Return True
        Catch ex As Exception
            Dim iret As Object = MsgBox("POWER_SUPPLY.PS_RECYCLE Returns Error : " & ex.Message, vbRetryCancel + vbCritical, "POWER_SUPPLY.Initialize()")
            If iret = vbCancel Then
                MAIN.ListBox1.Items.Add($"{DateTime.Now.ToString} POWER_SUPPLY.PS_RECYCLE Returns Error : " & ex.Message)
                MAIN.ListBox1.SelectedIndex = MAIN.ListBox1.Items.Count - 1
                Return False
            ElseIf iret = vbRetry Then
                MAIN.ListBox1.Items.Add($"{DateTime.Now.ToString} POWER_SUPPLY.PS_RECYCLE Returns Error : " & ex.Message)
                MAIN.ListBox1.SelectedIndex = MAIN.ListBox1.Items.Count - 1
                GoTo line1
            End If
        End Try

    End Function
End Module
