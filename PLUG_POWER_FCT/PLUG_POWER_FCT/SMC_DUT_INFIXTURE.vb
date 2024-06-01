Public Class SMC_DUT_INFIXTURE
    Private Sub Button1_Click(sender As Object, e As EventArgs) Handles Button1.Click
        If MAIN.RadioButton1.Checked = True Then
            If MAIN.TextBox2.Text = "PRE-COATING" Then
                MAIN.TimerFunction("START")
                MAIN.ListBox1.Items.Add($"{DateTime.Now.ToString} DUT Is Placed In the Test Fixture Is Confirmed")
                MAIN.ListBox1.SelectedIndex = MAIN.ListBox1.Items.Count - 1
                Me.Close()
                SMC.SMCPowerTestNode0()
            ElseIf MAIN.TextBox2.Text = "POST-COATING" Then
                MAIN.TimerFunction("START")
                MAIN.ListBox1.Items.Add($"{DateTime.Now.ToString} DUT Is Placed In the Test Fixture Is Confirmed")
                MAIN.ListBox1.SelectedIndex = MAIN.ListBox1.Items.Count - 1
                Me.Close()
                POWER_SUPPLY.SetLimit()
                System.Threading.Thread.Sleep(2000)
                POWER_SUPPLY.PS_ON()
                SMC.SMCPowerTestNode0()
            End If

        Else
            Me.Close()
            Exit Sub
        End If
    End Sub
End Class