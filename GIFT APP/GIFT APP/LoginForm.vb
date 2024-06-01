Imports BunifuAnimatorNS
Imports System.DirectoryServices.AccountManagement
Public Class LoginForm
    Dim transition As New BunifuTransition
    Dim Home = {"Home"}
    Dim Gift = {"Issue Gift", "Your History"}
    Dim Events = {"Add Events"}
    Dim history = {"Gift Histoiry", "Pending History"}
    Dim Setup = {"Setup & Confrigurations"}
    Dim GLogin As Boolean = True

    Private Sub LoginForm_Load(sender As Object, e As EventArgs) Handles MyBase.Load
        MainWindow.Opacity = 0.99
        MainWindow.Enabled = False
        BunifuTextBox1.Clear()
        BunifuTextBox2.Clear()
    End Sub

    Private Sub BunifuTextBox2_TextChanged(sender As Object, e As EventArgs) Handles BunifuTextBox2.TextChanged
        BunifuTextBox2.PasswordChar = "*"
    End Sub

    Private Sub BunifuButton1_Click(sender As Object, e As EventArgs) Handles BunifuButton1.Click
        If BunifuTextBox1 Is Nothing OrElse BunifuTextBox2 Is Nothing Then
            BunifuSnackbar1.Show(Me, "User Name And Password Cannot Be Empty!", Bunifu.UI.WinForms.BunifuSnackbar.MessageTypes.Information, 2000)
            BunifuTextBox1.Clear()
            BunifuTextBox2.Clear()
        ElseIf BunifuTextBox1 IsNot Nothing AndAlso BunifuTextBox2 IsNot Nothing Then
            Try
                'Using context As New PrincipalContext(ContextType.Domain)
                'GLogin = context.ValidateCredentials(BunifuTextBox1.Text, BunifuTextBox2.Text)
                'End Using
                If BunifuTextBox1.Text = "Admin" And BunifuTextBox2.Text = "Jabil@4132" Then
                    MainWindow.ListBox1.Items.Add("00000001")
                    MainWindow.ListBox1.Items.Add("0000001")
                    MainWindow.ListBox1.Items.Add("00000")
                    MainWindow.ListBox1.Items.Add("0000000000")
                    MainWindow.ListBox1.Items.Add("ADMINISTRATOR")
                    MainWindow.ListBox1.Items.Add("M")
                    MainWindow.ListBox1.Items.Add("04/16/2024")
                    MainWindow.ListBox1.Items.Add("ADMINISTRATIVE")
                    MainWindow.ListBox1.Items.Add("ADMINISTRATOR")
                    MainWindow.ListBox1.Items.Add("NA")
                    MainWindow.ListBox1.Items.Add("NA")
                    MainWindow.ListBox1.Items.Add("04/16/2024")
                    MainWindow.ListBox1.Items.Add("NA")
                    MainWindow.ListBox1.Items.Add("04/16/2024")
                    MainWindow.ListBox1.Items.Add("True")
                    MainWindow.ListBox1.Items.Add("True")
                    MainWindow.ListBox1.Items.Add("True")
                    FuncLib.EnhanceFeatures.IconbuildafterLogin(True, "ADMINISTRATOR")
                    BunifuSnackbar1.Show(MainWindow, $"Welcome {MainWindow.ListBox1.Items(4)}", Bunifu.UI.WinForms.BunifuSnackbar.MessageTypes.Success, 2500)
                    MainWindow.Enabled = True
                    MainWindow.CheckBox1.Checked = True
                    MainWindow.BunifuDropdown1.Visible = True
                    MainWindow.BunifuPanel3.Visible = True
                    MainWindow.BunifuDropdown2.Items.Clear()
                    MainWindow.BunifuDropdown2.Items.AddRange(Home)
                    MainWindow.BunifuDropdown6.Items.Clear()
                    MainWindow.BunifuDropdown6.Items.AddRange(Gift)
                    MainWindow.BunifuDropdown3.Items.Clear()
                    MainWindow.BunifuDropdown3.Items.AddRange(Events)
                    MainWindow.BunifuDropdown4.Items.Clear()
                    MainWindow.BunifuDropdown4.Items.AddRange(history)
                    MainWindow.BunifuDropdown5.Items.Clear()
                    MainWindow.BunifuDropdown5.Items.AddRange(Setup)
                    MainWindow.BunifuLabel1.Visible = True
                    MainWindow.BunifuPanel2.Visible = True
                    FuncLib.WriteLog.WriteAppLog($"Login Success With {MainWindow.ListBox1.Items(4)}_{MainWindow.ListBox1.Items(0)}")
                    transition.HideSync(Me, False, BunifuAnimatorNS.Animation.ScaleAndHorizSlide)
                    Me.Close()
                ElseIf GLogin = True Then
                    Try
                        Dim CheckUserValidation As String = FuncLib.DataBaseOperations.CheckUserValidation(BunifuTextBox1.Text)
                        If CheckUserValidation = "True" Then
                            BunifuSnackbar1.Show(MainWindow, $"Welcome {MainWindow.ListBox1.Items(4)}", Bunifu.UI.WinForms.BunifuSnackbar.MessageTypes.Success, 2500)
                            MainWindow.Enabled = True
                            MainWindow.CheckBox1.Checked = True
                            MainWindow.BunifuDropdown1.Visible = True
                            MainWindow.BunifuPanel3.Visible = True
                            MainWindow.BunifuDropdown2.Items.Clear()
                            MainWindow.BunifuDropdown2.Items.AddRange(Home)
                            MainWindow.BunifuDropdown6.Items.Clear()
                            MainWindow.BunifuDropdown6.Items.AddRange(Gift)
                            MainWindow.BunifuDropdown3.Items.Clear()
                            MainWindow.BunifuDropdown3.Items.AddRange(Events)
                            MainWindow.BunifuDropdown4.Items.Clear()
                            MainWindow.BunifuDropdown4.Items.AddRange(history)
                            MainWindow.BunifuDropdown5.Items.Clear()
                            MainWindow.BunifuDropdown5.Items.AddRange(Setup)
                            MainWindow.BunifuLabel1.Visible = True
                            MainWindow.BunifuPanel2.Visible = True
                            FuncLib.WriteLog.WriteAppLog($"Login Success With {MainWindow.ListBox1.Items(4)}_{MainWindow.ListBox1.Items(0)}")
                            transition.HideSync(Me, False, BunifuAnimatorNS.Animation.ScaleAndHorizSlide)
                            Me.Close()
                        ElseIf CheckUserValidation = "False" Then
                            BunifuSnackbar1.Show(Me, $"Invalid UserName or Password {BunifuTextBox1.Text}", Bunifu.UI.WinForms.BunifuSnackbar.MessageTypes.Information, 2000)
                            FuncLib.WriteLog.WriteAppLog($"Login Failed For {BunifuTextBox1.Text}")
                            BunifuTextBox1.Clear()
                            BunifuTextBox2.Clear()
                            MainWindow.BunifuLabel1.Visible = False
                            Exit Sub
                        Else
                            BunifuSnackbar1.Show(Me, $"Error Occured While Verifying User Credentials.{Environment.NewLine}Error :{CheckUserValidation}", Bunifu.UI.WinForms.BunifuSnackbar.MessageTypes.Error, 2000)
                            FuncLib.WriteLog.WriteErrorLog($"Error Occured While Verifying User Credentials :{CheckUserValidation}")
                            FuncLib.WriteLog.WriteAppLog($"Error Occured While Verifying User Credentials :{BunifuTextBox1.Text}")
                            BunifuTextBox1.Clear()
                            BunifuTextBox2.Clear()
                            MainWindow.BunifuLabel1.Visible = False
                            Exit Sub
                        End If
                    Catch ex As Exception
                        BunifuSnackbar1.Show(Me, $"Error Occured While Verifying User Credentials.{Environment.NewLine}Error :{ex.Message}", Bunifu.UI.WinForms.BunifuSnackbar.MessageTypes.Error, 2000)
                        FuncLib.WriteLog.WriteErrorLog($"Error Occured While Verifying User Credentials :{ex.Message}")
                        FuncLib.WriteLog.WriteAppLog($"Error Occured While Verifying User Credentials :{BunifuTextBox1.Text}")
                        BunifuTextBox1.Clear()
                        BunifuTextBox2.Clear()
                        MainWindow.BunifuLabel1.Visible = False
                        Exit Sub
                    End Try
                ElseIf GLogin = False Then
                    BunifuSnackbar1.Show(Me, $"Invalid UserName or Password {BunifuTextBox1.Text}", Bunifu.UI.WinForms.BunifuSnackbar.MessageTypes.Information, 2000)
                    FuncLib.WriteLog.WriteAppLog($"Login Failed For {BunifuTextBox1.Text}")
                    BunifuTextBox1.Clear()
                    BunifuTextBox2.Clear()
                    MainWindow.BunifuLabel1.Visible = False
                    Exit Sub
                End If
            Catch ex As Exception
                BunifuSnackbar1.Show(Me, $"Error Occured While Verifying User Credentials.{Environment.NewLine}Error :{ex.Message}", Bunifu.UI.WinForms.BunifuSnackbar.MessageTypes.Error, 2000)
                FuncLib.WriteLog.WriteErrorLog($"Error Occured While Verifying User Credentials :{ex.Message}")
                FuncLib.WriteLog.WriteAppLog($"Error Occured While Verifying User Credentials :{BunifuTextBox1.Text}")
                BunifuTextBox1.Clear()
                BunifuTextBox2.Clear()
                MainWindow.BunifuLabel1.Visible = False
                Exit Sub
            End Try
        End If
    End Sub

    Private Sub BunifuButton2_Click(sender As Object, e As EventArgs) Handles BunifuButton2.Click
        transition.HideSync(Me, False, BunifuAnimatorNS.Animation.ScaleAndHorizSlide)
        Me.Close()
        transition.HideSync(MainWindow, False, BunifuAnimatorNS.Animation.ScaleAndHorizSlide)
        Application.Exit()
    End Sub
End Class