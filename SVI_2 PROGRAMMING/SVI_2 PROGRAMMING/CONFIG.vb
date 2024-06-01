Imports MadMilkman.Ini
Public Class CONFIG
    Private Sub Button5_Click(sender As Object, e As EventArgs) Handles Button5.Click
        Me.Hide()
        Call SETTING.Show()
    End Sub

    Private Sub Button4_Click(sender As Object, e As EventArgs) Handles Button4.Click
        OpenFileDialog1.InitialDirectory = "C:\"
        OpenFileDialog1.FileName = "USER TEMPALTE"
        OpenFileDialog1.Multiselect = False
        OpenFileDialog1.Filter = "excel file|*.xlsx"
        If OpenFileDialog1.ShowDialog = System.Windows.Forms.DialogResult.OK Then
            Dim sName As String = OpenFileDialog1.SafeFileName
            TextBox1.Text = OpenFileDialog1.FileName

        End If
    End Sub

    Private Sub OpenFileDialog1_FileOk(sender As Object, e As System.ComponentModel.CancelEventArgs) Handles OpenFileDialog1.FileOk
    End Sub

    Private Sub Button7_Click(sender As Object, e As EventArgs) Handles Button7.Click
        OpenFileDialog2.InitialDirectory = "C:\"
        OpenFileDialog2.FileName = "ACTIVITY TRACKER"
        OpenFileDialog2.Multiselect = False
        OpenFileDialog2.Filter = "text file|*.txt"
        If OpenFileDialog2.ShowDialog = System.Windows.Forms.DialogResult.OK Then
            Dim sName As String = OpenFileDialog2.SafeFileName
            TextBox2.Text = OpenFileDialog2.FileName
        End If
    End Sub

    Private Sub Button1_Click(sender As Object, e As EventArgs) Handles Button1.Click
        Call PORTS.Show()
    End Sub

    Private Sub CONFIG_Load(sender As Object, e As EventArgs) Handles MyBase.Load
        TextBox7.Enabled = False
        Button12.Enabled = False
    End Sub

    Private Sub Button2_Click(sender As Object, e As EventArgs) Handles Button2.Click
        OpenFileDialog3.InitialDirectory = "C:\"
        OpenFileDialog3.FileName = "MnMMATest"
        OpenFileDialog3.Multiselect = False
        OpenFileDialog3.Filter = "executable files|*.exe"
        If OpenFileDialog3.ShowDialog = System.Windows.Forms.DialogResult.OK Then
            Dim sName As String = OpenFileDialog3.SafeFileName
            TextBox5.Text = OpenFileDialog3.FileName
        End If
    End Sub

    Private Sub Button3_Click(sender As Object, e As EventArgs) Handles Button3.Click
        If (FolderBrowserDialog1.ShowDialog() = DialogResult.OK) Then
            TextBox6.Text = FolderBrowserDialog1.SelectedPath
        End If
    End Sub

    Private Sub Button9_Click(sender As Object, e As EventArgs) Handles Button9.Click
        Dim iret1 As Object = MsgBox("DO YOU WANT TO SAVE THE SETTINGS?", vbQuestion + vbYesNo, "SETTING UPDATE")
        If iret1 = vbYes Then
            Dim ini As New IniFile()
            ini.Load("C:\SVI 2 PROGRAMMING SOFTWARE\SVI_2 PROGRAMMING\SVI_2 PROGRAMMING\obj\CONFRIGURATIONS\CONFIG.ini")
            ini.Sections("USER_TEMPLATE").Keys("Count").Value = TextBox1.Text
            ini.Sections("ACTIVITY_LOG").Keys("Count").Value = TextBox2.Text
            ini.Sections("DB_PATH").Keys("Count").Value = TextBox7.Text
            ini.Sections("POWER_SUPPLY").Keys("Count").Value = TextBox4.Text
            ini.Sections("MNMMA_PATH").Keys("Count").Value = TextBox5.Text
            ini.Sections("LOG_PATH").Keys("Count").Value = TextBox6.Text
            ini.Save("C:\SVI 2 PROGRAMMING SOFTWARE\SVI_2 PROGRAMMING\SVI_2 PROGRAMMING\obj\CONFRIGURATIONS\CONFIG.ini")
            MsgBox("SETTINGS SAVED SUCCESSFULLY. PLEASE RESTART THE SOFTWARE")
            Me.Close()
            Application.Exit()
        Else
            If iret1 = vbNo Then
                'donothing
            End If
        End If
    End Sub
End Class