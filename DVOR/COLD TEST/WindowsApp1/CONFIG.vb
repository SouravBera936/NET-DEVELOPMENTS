Imports MadMilkman.Ini
Public Class CONFIG
    Private Sub CONFIG_Load(sender As Object, e As EventArgs) Handles MyBase.Load

    End Sub

    Private Sub Button6_Click(sender As Object, e As EventArgs) Handles Button6.Click
        Me.Hide()
        Call SETTING.Show()
    End Sub

    Private Sub Button1_Click(sender As Object, e As EventArgs) Handles Button1.Click
        OpenFileDialog1.InitialDirectory = "C:\"
        OpenFileDialog1.FileName = "USER FILE"
        OpenFileDialog1.Multiselect = False
        OpenFileDialog1.Filter = "Excel File|*.xlsx"
        If OpenFileDialog1.ShowDialog = System.Windows.Forms.DialogResult.OK Then
            Dim sName As String = OpenFileDialog1.SafeFileName
            TextBox1.Text = OpenFileDialog1.FileName
        End If
    End Sub

    Private Sub Button2_Click(sender As Object, e As EventArgs) Handles Button2.Click
        OpenFileDialog1.InitialDirectory = "C:\"
        OpenFileDialog1.FileName = "ACTIVITY TRACKER"
        OpenFileDialog1.Multiselect = False
        OpenFileDialog1.Filter = "Text File|*.txt"
        If OpenFileDialog1.ShowDialog = System.Windows.Forms.DialogResult.OK Then
            Dim sName As String = OpenFileDialog1.SafeFileName
            TextBox2.Text = OpenFileDialog1.FileName
        End If
    End Sub

    Private Sub Button3_Click(sender As Object, e As EventArgs) Handles Button3.Click
        OpenFileDialog1.InitialDirectory = "C:\"
        OpenFileDialog1.FileName = "COUPLER_TEMPLATE"
        OpenFileDialog1.Multiselect = False
        OpenFileDialog1.Filter = "Excel File|*.xlsx"
        If OpenFileDialog1.ShowDialog = System.Windows.Forms.DialogResult.OK Then
            Dim sName As String = OpenFileDialog1.SafeFileName
            TextBox3.Text = OpenFileDialog1.FileName
        End If
    End Sub

    Private Sub Button4_Click(sender As Object, e As EventArgs) Handles Button4.Click
        If (FolderBrowserDialog1.ShowDialog() = DialogResult.OK) Then
            TextBox4.Text = FolderBrowserDialog1.SelectedPath & "\"
        End If
    End Sub

    Private Sub Button7_Click(sender As Object, e As EventArgs) Handles Button7.Click
        OpenFileDialog1.InitialDirectory = "C:\"
        OpenFileDialog1.FileName = ""
        OpenFileDialog1.Multiselect = False
        OpenFileDialog1.Filter = "|*.jpg"
        If OpenFileDialog1.ShowDialog = System.Windows.Forms.DialogResult.OK Then
            Dim sName As String = OpenFileDialog1.SafeFileName
            TextBox5.Text = OpenFileDialog1.FileName
        End If
    End Sub

    Private Sub Button8_Click(sender As Object, e As EventArgs) Handles Button8.Click
        OpenFileDialog1.InitialDirectory = "C:\"
        OpenFileDialog1.FileName = ""
        OpenFileDialog1.Multiselect = False
        OpenFileDialog1.Filter = "|*.jpg"
        If OpenFileDialog1.ShowDialog = System.Windows.Forms.DialogResult.OK Then
            Dim sName As String = OpenFileDialog1.SafeFileName
            TextBox6.Text = OpenFileDialog1.FileName
        End If
    End Sub

    Private Sub Button9_Click(sender As Object, e As EventArgs) Handles Button9.Click
        OpenFileDialog1.InitialDirectory = "C:\"
        OpenFileDialog1.FileName = ""
        OpenFileDialog1.Multiselect = False
        OpenFileDialog1.Filter = "|*.jpg"
        If OpenFileDialog1.ShowDialog = System.Windows.Forms.DialogResult.OK Then
            Dim sName As String = OpenFileDialog1.SafeFileName
            TextBox7.Text = OpenFileDialog1.FileName
        End If
    End Sub

    Private Sub Button10_Click(sender As Object, e As EventArgs) Handles Button10.Click
        OpenFileDialog1.InitialDirectory = "C:\"
        OpenFileDialog1.FileName = ""
        OpenFileDialog1.Multiselect = False
        OpenFileDialog1.Filter = "|*.jpg"
        If OpenFileDialog1.ShowDialog = System.Windows.Forms.DialogResult.OK Then
            Dim sName As String = OpenFileDialog1.SafeFileName
            TextBox8.Text = OpenFileDialog1.FileName
        End If
    End Sub

    Private Sub Button11_Click(sender As Object, e As EventArgs) Handles Button11.Click
        OpenFileDialog1.InitialDirectory = "C:\"
        OpenFileDialog1.FileName = ""
        OpenFileDialog1.Multiselect = False
        OpenFileDialog1.Filter = "|*.jpg"
        If OpenFileDialog1.ShowDialog = System.Windows.Forms.DialogResult.OK Then
            Dim sName As String = OpenFileDialog1.SafeFileName
            TextBox9.Text = OpenFileDialog1.FileName
        End If
    End Sub

    Private Sub Button12_Click(sender As Object, e As EventArgs) Handles Button12.Click
        OpenFileDialog1.InitialDirectory = "C:\"
        OpenFileDialog1.FileName = ""
        OpenFileDialog1.Multiselect = False
        OpenFileDialog1.Filter = "|*.jpg"
        If OpenFileDialog1.ShowDialog = System.Windows.Forms.DialogResult.OK Then
            Dim sName As String = OpenFileDialog1.SafeFileName
            TextBox10.Text = OpenFileDialog1.FileName
        End If
    End Sub

    Private Sub Button13_Click(sender As Object, e As EventArgs) Handles Button13.Click
        OpenFileDialog1.InitialDirectory = "C:\"
        OpenFileDialog1.FileName = ""
        OpenFileDialog1.Multiselect = False
        OpenFileDialog1.Filter = "|*.jpg"
        If OpenFileDialog1.ShowDialog = System.Windows.Forms.DialogResult.OK Then
            Dim sName As String = OpenFileDialog1.SafeFileName
            TextBox11.Text = OpenFileDialog1.FileName
        End If
    End Sub

    Private Sub Button15_Click(sender As Object, e As EventArgs) Handles Button15.Click
        OpenFileDialog1.InitialDirectory = "C:\"
        OpenFileDialog1.FileName = "STATUS PANNEL_TEMPLATE"
        OpenFileDialog1.Multiselect = False
        OpenFileDialog1.Filter = "Excel File|*.xlsx"
        If OpenFileDialog1.ShowDialog = System.Windows.Forms.DialogResult.OK Then
            Dim sName As String = OpenFileDialog1.SafeFileName
            TextBox13.Text = OpenFileDialog1.FileName
        End If
    End Sub

    Private Sub Button14_Click(sender As Object, e As EventArgs) Handles Button14.Click
        If (FolderBrowserDialog1.ShowDialog() = DialogResult.OK) Then
            TextBox12.Text = FolderBrowserDialog1.SelectedPath & "\"
        End If
    End Sub

    Private Sub Button5_Click(sender As Object, e As EventArgs) Handles Button5.Click
        Dim iret As Object = MsgBox("DO YOU WANT TO UPDATE THE SETTINGS?", vbQuestion + vbYesNo, "SETTING UPDATE")
        If iret = vbNo Then
            'donothing
        Else
            If iret = vbYes Then
                Dim ini As New IniFile()
                ini.Load("C:\DVOR\COLD TEST\WindowsApp1\obj\CONFRIGURATIONS.ini")
                ini.Sections("USER_DATA").Keys("Count").Value = TextBox1.Text
                ini.Sections("TRACKER").Keys("Count").Value = TextBox2.Text
                ini.Sections("COUPLER_TEMPLATE").Keys("Count").Value = TextBox3.Text
                ini.Sections("COUPLER_LOG").Keys("Count").Value = TextBox4.Text
                ini.Sections("COUPLER_1").Keys("Count").Value = TextBox5.Text
                ini.Sections("COUPLER_2").Keys("Count").Value = TextBox6.Text
                ini.Sections("COUPLER_3").Keys("Count").Value = TextBox7.Text
                ini.Sections("COUPLER_4").Keys("Count").Value = TextBox8.Text
                ini.Sections("COUPLER_5").Keys("Count").Value = TextBox9.Text
                ini.Sections("COUPLER_6").Keys("Count").Value = TextBox10.Text
                ini.Sections("STATUS_TEMPLATE").Keys("Count").Value = TextBox13.Text
                ini.Sections("STATUS_LOG").Keys("Count").Value = TextBox12.Text
                ini.Sections("STATUS_1").Keys("Count").Value = TextBox11.Text
                ini.Save("C:\DVOR\COLD TEST\WindowsApp1\obj\CONFRIGURATIONS.ini")
                Dim iret1 As Object = MsgBox("SETTINGS UPDATED. PLEASE RESTART THE SOFTWARE", vbInformation + vbOKOnly, "INFORMATION")
                If iret1 = vbOK Then
                    Me.Close()
                    Application.Exit()
                End If
            End If
        End If
    End Sub
End Class