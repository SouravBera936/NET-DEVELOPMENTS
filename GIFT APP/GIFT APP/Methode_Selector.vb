Imports BunifuAnimatorNS
Imports System.IO.Ports
Imports System.Windows.Forms.VisualStyles.VisualStyleElement.Button
Public Class Methode_Selector
    Dim transition As New BunifuTransition
    Private Sub Methode_Selector_Load(sender As Object, e As EventArgs) Handles MyBase.Load
        Me.StartPosition = FormStartPosition.CenterScreen
        BunifuDropdown7.Text = Nothing
        BunifuDropdown7.Items.Clear()
        'BunifuDropdown7.Items.Add("Individual Distribution")
        BunifuDropdown7.Items.Add("Bulk Distribution")
        GroupBox2.Visible = False
        GroupBox1.Enabled = True
        MainWindow.TextBox2.Clear()
        MainWindow.RadioButton1.Checked = False
        MainWindow.RadioButton2.Checked = False
        MainWindow.RadioButton6.Checked = False 'Reset Emp Identification Methode
        MainWindow.RadioButton5.Checked = False 'Reset Emp Identification Methode
    End Sub

    Private Sub BunifuDropdown7_SelectedIndexChanged(sender As Object, e As EventArgs) Handles BunifuDropdown7.SelectedIndexChanged
        If BunifuDropdown7.SelectedItem = "Individual Distribution" Then
            MainWindow.RadioButton1.Checked = True
            MainWindow.RadioButton2.Checked = False
            transition.HideSync(Me, False, BunifuAnimatorNS.Animation.ScaleAndHorizSlide)
            MainWindow.GroupBox4.Visible = True
            MainWindow.BunifuTextBox1.Focus()
            MainWindow.BunifuTextBox1.PlaceholderText = "Enter Employee ID"
            MainWindow.BunifuTextBox1.Enabled = True
            MainWindow.BunifuTextBox1.ReadOnly = False
        ElseIf BunifuDropdown7.SelectedItem = "Bulk Distribution" Then
            MainWindow.RadioButton1.Checked = False
            MainWindow.RadioButton2.Checked = True
            GroupBox1.Enabled = False
            GroupBox2.Visible = True
            BunifuDropdown4.Items.Clear()
            BunifuDropdown4.Text = Nothing
            Dim ports As String() = SerialPort.GetPortNames()
            For Each port As String In ports
                BunifuDropdown4.Items.Add(port)
            Next
            BunifuDropdown4.Enabled = False
            BunifuDropdown1.Enabled = True
            BunifuDropdown1.Items.Clear()
            BunifuDropdown1.Text = Nothing
            BunifuDropdown1.Items.Add("BY 10D")
            BunifuDropdown1.Items.Add("BY 5D")
        End If
    End Sub
    Private Sub BunifuDropdown4_SelectedIndexChanged(sender As Object, e As EventArgs) Handles BunifuDropdown4.SelectedIndexChanged
        Dim selectedPort As String = BunifuDropdown4.SelectedItem.ToString()
        Dim serialPort As New SerialPort(selectedPort)
        Try
            If serialPort.IsOpen = True Then
                BunifuSnackbar1.Show(Me, $"{BunifuDropdown4.Text} has been Successfully validated", Bunifu.UI.WinForms.BunifuSnackbar.MessageTypes.Success, 3000)
                MainWindow.TextBox2.Text = BunifuDropdown4.SelectedItem.ToString
                MainWindow.BunifuTextBox1.Enabled = True
                MainWindow.BunifuTextBox1.Focus()
                serialPort.Close()
                MainWindow.BunifuTextBox1.PlaceholderText = "Scan ID Card"
                MainWindow.BunifuTextBox1.ReadOnly = True
                transition.HideSync(Me, False, BunifuAnimatorNS.Animation.ScaleAndHorizSlide)
                MainWindow.GroupBox4.Visible = True
            ElseIf serialPort.IsOpen = False Then
                serialPort.Open()
                BunifuSnackbar1.Show(Me, $"{BunifuDropdown4.Text} has been Successfully validated", Bunifu.UI.WinForms.BunifuSnackbar.MessageTypes.Success, 3000)
                MainWindow.TextBox2.Text = BunifuDropdown4.SelectedItem.ToString
                MainWindow.BunifuTextBox1.Enabled = True
                MainWindow.BunifuTextBox1.Focus()
                serialPort.Close()
                MainWindow.BunifuTextBox1.PlaceholderText = "Scan ID Card"
                MainWindow.BunifuTextBox1.ReadOnly = True
                transition.HideSync(Me, False, BunifuAnimatorNS.Animation.ScaleAndHorizSlide)
                MainWindow.GroupBox4.Visible = True
            End If
        Catch ex As Exception
            BunifuSnackbar1.Show(Me, $"Error While Initializing the {BunifuDropdown4.Text}{Environment.NewLine} Error : {ex.Message}", Bunifu.UI.WinForms.BunifuSnackbar.MessageTypes.Warning, 3000)
            FuncLib.WriteLog.WriteErrorLog($"Error While Initializing the {BunifuDropdown4.Text} : {ex.Message}")
            Exit Sub
        End Try
    End Sub

    Private Sub BunifuDropdown1_SelectedIndexChanged(sender As Object, e As EventArgs) Handles BunifuDropdown1.SelectedIndexChanged
        If BunifuDropdown1.SelectedItem = "BY 10D" Then
            MainWindow.RadioButton6.Checked = True
            MainWindow.RadioButton5.Checked = False
            BunifuDropdown4.Enabled = True
        ElseIf BunifuDropdown1.SelectedItem = "BY 5D" Then
            MainWindow.RadioButton6.Checked = False
            MainWindow.RadioButton5.Checked = True
            BunifuDropdown4.Enabled = True
        Else
            BunifuDropdown4.Enabled = False
        End If
    End Sub
End Class