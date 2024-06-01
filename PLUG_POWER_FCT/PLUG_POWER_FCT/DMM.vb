Imports Ivi.Visa.Interop
Module DMM
    Dim Port As String = CONFIG.TextBox49.Text.ToString
    Dim SetVoltage As String = CONFIG.TextBox32.Text.ToString
    Dim SetCurrent As String = CONFIG.TextBox33.Text.ToString
    Dim iomgr As New Ivi.Visa.Interop.ResourceManager
    Dim instrany As New Ivi.Visa.Interop.FormattedIO488
    Dim file1 As System.IO.StreamWriter
    Dim MeasuredVoltage As Double
    Function Initialize() As Boolean
        Dim line1 As Label
line1:
        Dim path1 As String = CONFIG.TextBox4.Text
        file1 = My.Computer.FileSystem.OpenTextFileWriter(path1, True)
        Try
            MAIN.ListBox1.Items.Add("Initializing DMM")
            MAIN.ListBox1.SelectedIndex = MAIN.ListBox1.Items.Count - 1
            instrany.IO = iomgr.Open(Port)
            instrany.WriteString("*IDN?")
            Dim PSID As String = instrany.ReadString
            MAIN.ListBox1.Items.Add("DMM.Initialize Returns IDN : " & PSID)
            MAIN.ListBox1.SelectedIndex = MAIN.ListBox1.Items.Count - 1
            file1.Close()
            instrany.IO.Close()
            Return True
        Catch ex As Exception
            Dim iret As Object = MsgBox("ERROR INITIALIZING DMM : " & ex.Message, vbOKOnly + vbCritical, "DMM.Initialize()")
            If iret = vbOK Then
                MAIN.ListBox1.Items.Add("DMM.Initialize Returns Error : " & ex.Message)
                MAIN.ListBox1.SelectedIndex = MAIN.ListBox1.Items.Count - 1
                file1.WriteLine("DMM INITIALIZATION FAILED : " & ex.Message & " " & "@" & DateTime.Now.ToString("yyyy/MM/dd/HH:mm:ss"))
                file1.Close()
                Return False
            End If
        End Try
    End Function
    Function ErrorClear()
        Dim line1 As Label
line1:
        Dim path1 As String = CONFIG.TextBox4.Text
        file1 = My.Computer.FileSystem.OpenTextFileWriter(path1, True)
        Try
            MAIN.ListBox1.Items.Add("CLEARING DMM IF ANY ERROR OCCURS")
            MAIN.ListBox1.SelectedIndex = MAIN.ListBox1.Items.Count - 1
            instrany.IO = iomgr.Open(Port)
            instrany.WriteString("*CLS")
            instrany.IO.Close()
            file1.Close()
            Return True
        Catch ex As Exception
            Dim iret As Object = MsgBox("ERROR PERFORMING *CLS : " & ex.Message, vbOKOnly + vbCritical, "DMM.ErrorClear()")
            If iret = vbOK Then
                MAIN.ListBox1.Items.Add("DMM.ErrorClear Returns Error : " & ex.Message)
                MAIN.ListBox1.SelectedIndex = MAIN.ListBox1.Items.Count - 1
                file1.WriteLine("ERROR PERFORMING *CLS : " & ex.Message & " " & "@" & DateTime.Now.ToString("yyyy/MM/dd/HH:mm:ss"))
                file1.Close()
                Return False
            End If
        End Try
    End Function
    Function MeasureDCVoltage() As Double
        Dim line1 As Label
line1:
        Dim path1 As String = CONFIG.TextBox4.Text
        file1 = My.Computer.FileSystem.OpenTextFileWriter(path1, True)
        Try
            instrany.IO = iomgr.Open(Port)
            instrany.WriteString("CONF:DC")
            instrany.WriteString("MEAS?")
            MeasuredVoltage = instrany.ReadString
            instrany.IO.Close()
            file1.Close()
            MeasuredVoltage = Math.Round(MeasuredVoltage, 2)
            Return MeasuredVoltage
        Catch ex As Exception
            Dim iret As Object = MsgBox("ERROR MEASURING VOLTAGE : " & ex.Message, vbRetryCancel + vbCritical, "DMM.MeasureDCVoltage()")
            If iret = vbRetry Then
                file1.Close()
                GoTo line1
            ElseIf iret = vbCancel Then
                MAIN.ListBox1.Items.Add("DMM.MeasureDCVoltage Returns Error : " & ex.Message)
                MAIN.ListBox1.SelectedIndex = MAIN.ListBox1.Items.Count - 1
                file1.WriteLine("ERROR MEASURING DC VOLTAGE : " & ex.Message & " " & "@" & DateTime.Now.ToString("yyyy/MM/dd/HH:mm:ss"))
                file1.Close()
                Exit Function
            End If
        End Try
    End Function
End Module
