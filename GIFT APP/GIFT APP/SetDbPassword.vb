Imports BunifuAnimatorNS
Public Class SetDbPassword
    Dim transition As New BunifuTransition
    Private Sub SetDbPassword_Load(sender As Object, e As EventArgs) Handles MyBase.Load
        BunifuTextBox3.Clear()
        BunifuTextBox5.Clear()
        BunifuTextBox1.Enabled = False
        BunifuTextBox2.Enabled = False
        BunifuButton1.Enabled = False
        BunifuTextBox5.Enabled = False
        BunifuTextBox3.BorderColorActive = Color.Yellow
        Me.StartPosition = FormStartPosition.CenterScreen
    End Sub
    Private Sub BunifuTextBox3_TextChanged(sender As Object, e As EventArgs) Handles BunifuTextBox3.TextChanged
        BunifuTextBox3.PasswordChar = "*"
        If BunifuTextBox3.Text = Nothing Then
            BunifuTextBox3.BorderColorActive = Color.Yellow
            BunifuTextBox5.Enabled = False
        Else
            BunifuTextBox3.BorderColorActive = Color.Green
            BunifuTextBox5.Enabled = True
            BunifuTextBox5.BorderColorActive = Color.Yellow
        End If
    End Sub
    Private Sub BunifuTextBox5_TextChanged(sender As Object, e As EventArgs) Handles BunifuTextBox5.TextChanged
        BunifuTextBox5.PasswordChar = "*"
        If BunifuTextBox3.Text = BunifuTextBox5.Text Then
            BunifuTextBox5.BorderColorActive = Color.Green
            BunifuButton1.Enabled = True
        Else
            BunifuTextBox5.BorderColorActive = Color.Red
            BunifuButton1.Enabled = False
        End If
    End Sub
    Private Sub BunifuButton1_Click(sender As Object, e As EventArgs) Handles BunifuButton1.Click
        Dim CreateNewDb As String = FuncLib.Setup.CreateNewDataBase(BunifuTextBox1.Text, BunifuTextBox2.Text, BunifuTextBox5.Text)
        If CreateNewDb = "True" Then
            Dim CreateDatabseTable As String = FuncLib.Setup.ForamttingOFDataabse(BunifuTextBox1.Text, BunifuTextBox2.Text, BunifuTextBox5.Text)
            If CreateDatabseTable = "True" Then
                Dim VerifyDb As String = FuncLib.Setup.VerifyNewDataBase(BunifuTextBox1.Text, BunifuTextBox2.Text, BunifuTextBox5.Text)
                If VerifyDb = "True" Then
                    BunifuSnackbar1.Show(Me, $"Database Created and Verified Successfully in {BunifuTextBox2.Text}", Bunifu.UI.WinForms.BunifuSnackbar.MessageTypes.Success, 2000)
                    FuncLib.WriteLog.WriteAppLog($"DataBase Source Changed To : {BunifuTextBox2.Text} By {MainWindow.ListBox1.Items(1)}_{MainWindow.ListBox1.Items(4)}")
                    SETUP.BunifuTextBox3.Text = BunifuTextBox2.Text
                    SETUP.BunifuTextBox4.Text = BunifuTextBox5.Text
                    BunifuSnackbar1.Show(Me, $"All Confrigurations Updated Successfully", Bunifu.UI.WinForms.BunifuSnackbar.MessageTypes.Success, 2000)
                    Threading.Thread.Sleep(1000)
                    transition.HideSync(Me, False, BunifuAnimatorNS.Animation.ScaleAndHorizSlide)
                    Me.Close()
                ElseIf VerifyDb = "False" Then
                    BunifuSnackbar1.Show(Me, $"Failed To Verify Database Structure Contact To Developer!", Bunifu.UI.WinForms.BunifuSnackbar.MessageTypes.Error, 2000)
                    Exit Sub
                Else
                    BunifuSnackbar1.Show(Me, $"Error Occured While Verifying The Database!.{Environment.NewLine} Error :{VerifyDb}", Bunifu.UI.WinForms.BunifuSnackbar.MessageTypes.Success, 2000)
                    FuncLib.WriteLog.WriteErrorLog($"Error Occured While Executing FuncLib.Setup.VerifyNewDataBase :{VerifyDb} ")
                    Exit Sub
                End If
            Else
                BunifuSnackbar1.Show(Me, $"Error Occured While Creating Tables Inside The Database!.{Environment.NewLine} Error :{CreateDatabseTable}", Bunifu.UI.WinForms.BunifuSnackbar.MessageTypes.Success, 2000)
                FuncLib.WriteLog.WriteErrorLog($"Error Occured While Executing FuncLib.Setup.ForamttingOFDataabse() :{CreateDatabseTable} ")
                Exit Sub
            End If
        Else
            BunifuSnackbar1.Show(Me, $"Error Occured While Creating New Database!.{Environment.NewLine} Error :{CreateNewDb}", Bunifu.UI.WinForms.BunifuSnackbar.MessageTypes.Success, 2000)
            FuncLib.WriteLog.WriteErrorLog($"Error Occured While Executing FuncLib.Setup.CreateNewDataBase() :{CreateNewDb} ")
            Exit Sub
        End If
    End Sub
    Private Sub BunifuButton2_Click(sender As Object, e As EventArgs) Handles BunifuButton2.Click
        transition.HideSync(Me, False, BunifuAnimatorNS.Animation.Scale)
        Me.Close()
    End Sub
End Class