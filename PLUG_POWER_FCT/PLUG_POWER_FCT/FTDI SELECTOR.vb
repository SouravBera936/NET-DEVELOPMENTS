Imports System.IO
Public Class FTDI_SELECTOR
    Dim xmlFilePath As String = CONFIG.TextBox35.Text.ToString
    Dim file1 As System.IO.StreamWriter

    Private Sub Button1_Click(sender As Object, e As EventArgs) Handles Button1.Click
        Dim line1 As Label
line1:
        Dim path1 As String = CONFIG.TextBox4.Text
        file1 = My.Computer.FileSystem.OpenTextFileWriter(path1, True)
        ComboBox1.Items.Clear()
        ComboBox1.Text = ""
        If ComboBox1.SelectedItem IsNot Nothing AndAlso TypeOf ComboBox1.SelectedItem Is FTDIDeviceInfo Then
            Dim selectedDevice As FTDIDeviceInfo = DirectCast(ComboBox1.SelectedItem, FTDIDeviceInfo)
            Dim iret As Object = MsgBox($"Starting the FTDI Flash For The Following Device PLease Confirm ? : {selectedDevice.Description} - {selectedDevice.SerialNumber}", vbYesNo + vbQuestion, "FTDIDevice Confirmation")
            If iret = vbYes Then
                MAIN.ListBox1.Items.Add($"{DateTime.Now.ToString} : User Confirmation Recieved To Flash FTDI for The Device : {selectedDevice.Description} - {selectedDevice.SerialNumber}")
                MAIN.ListBox1.SelectedIndex = MAIN.ListBox1.Items.Count - 1
                Dim destinationPath As String = $"\\.\COM{selectedDevice.DeviceIndex + 1}"
                Try
                    File.Copy(xmlFilePath, destinationPath, True)
                    MAIN.ListBox1.Items.Add($"{DateTime.Now.ToString} FTDI Flashing Success For : {selectedDevice} ")
                    MAIN.ListBox1.SelectedIndex = MAIN.ListBox1.Items.Count - 1
                    MAIN.TextBox37.Text = "TRUE"
                    MAIN.Button17.Text = "PASS"
                    MAIN.TreeView1.SelectedNode.Nodes(0).Nodes(5).BackColor = Color.Green
                    SMC.SMCProgrammingNode6()
                Catch ex As Exception
                    Dim iret1 As Object = MsgBox("Failed to flash the FTDI Device : " & ex.Message, vbRetryCancel + vbCritical, "FTDI_SELECTOR.FTDIFlash")
                    If iret1 = vbRetry Then
                        MAIN.ListBox1.Items.Add($"{DateTime.Now.ToString} Failed to flash the FTDI Device : " & ex.Message)
                        MAIN.ListBox1.SelectedIndex = MAIN.ListBox1.Items.Count - 1
                        file1.Close()
                        GoTo line1
                    ElseIf iret1 = vbCancel Then
                        file1.WriteLine("Failed to flash the FTDI Device : " & ex.Message & " " & "@" & DateTime.Now.ToString("yyyy/MM/dd/HH:mm:ss"))
                        file1.Close()
                        MAIN.TextBox37.Text = "FALSE"
                        MAIN.Button17.Text = "FAIL"
                        MAIN.TreeView1.SelectedNode.Nodes(0).Nodes(5).BackColor = Color.Red
                        MAIN.ListBox1.Items.Add($"{DateTime.Now.ToString} Failed to flash the FTDI Device : " & ex.Message)
                        MAIN.ListBox1.SelectedIndex = MAIN.ListBox1.Items.Count - 1
                        MAIN.TimerFunction("STOP")
                        UUT_RESULT.RadioButton2.Checked = True
                        UUT_RESULT.TextBox2.Text = "FTDI_PROG"
                        UUT_RESULT.ShowDialog()
                    End If
                End Try
            ElseIf iret = vbNo Then
                'donothing
            End If
        End If
    End Sub
End Class