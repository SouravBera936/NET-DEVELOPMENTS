Public Class TESTFORM
    Private Sub TabPage2_Click(sender As Object, e As EventArgs) Handles TabPage1.Click

    End Sub

    Private Sub TESTFORM_Load(sender As Object, e As EventArgs) Handles MyBase.Load
        ComboBox1.Items.Clear()
        ComboBox1.Items.Add("4208-MOTHER BOARD")
        TextBox3.Enabled = False
        TextBox4.Text = ""
        Dim todaysdate As String = String.Format("{0:dd/MM/yyyy}", DateTime.Now)
        TextBox1.Text = todaysdate
        Timer1.Enabled = True
        TextBox1.Enabled = False
        TextBox2.Enabled = False
        Button1.Enabled = False
        Button1.Text = "IDLE"
        TextBox5.Enabled = False
        TextBox6.Enabled = False
        TextBox7.Enabled = False
        TextBox8.Enabled = False
        TextBox9.Enabled = False
        TextBox10.Enabled = False
        TextBox11.Enabled = False
        TextBox13.Enabled = False
        TextBox14.Enabled = False
        TextBox15.Enabled = False
        TextBox16.Enabled = False
        TextBox18.Enabled = False
        TextBox19.Enabled = False
        TextBox4.Enabled = False
        Button2.Enabled = False
    End Sub

    Private Sub Timer1_Tick(sender As Object, e As EventArgs) Handles Timer1.Tick
        TextBox2.Text = Date.Now.ToString("hh:mm:ss tt")
    End Sub

    Private Sub ComboBox1_SelectedIndexChanged(sender As Object, e As EventArgs) Handles ComboBox1.SelectedIndexChanged
        ComboBox2.Text = ""
        ComboBox2.Items.Clear()
        ComboBox2.Items.Add("PROGRAMMING TEST")
    End Sub

    Private Sub Button3_Click(sender As Object, e As EventArgs) Handles Button3.Click
        Me.Close()
        Call MAIN.Show()
    End Sub

    Private Sub ComboBox2_SelectedIndexChanged(sender As Object, e As EventArgs) Handles ComboBox2.SelectedIndexChanged
        If ComboBox2.SelectedItem = "PROGRAMMING TEST" Then
            TextBox4.Enabled = True
        Else
            TextBox4.Enabled = False
        End If
    End Sub

    Private Sub TextBox4_TextChanged(sender As Object, e As EventArgs) Handles TextBox4.TextChanged
        TextBox12.Text = ""
        TextBox10.Text = "IDLE"
        TextBox17.Text = ""
        TextBox15.Text = "IDLE"
        Button1.Text = "IDLE"
        If TextBox4.Text = "" Then
            Button2.Enabled = False
        Else
            Button2.Enabled = True
        End If
    End Sub

    Private Sub Button1_Click(sender As Object, e As EventArgs) Handles Button1.Click

    End Sub

    Private Sub Button1_TextChanged(sender As Object, e As EventArgs) Handles Button1.TextChanged
        If Button1.Text = "IDLE" Then
            Button1.BackColor = Color.Yellow
        Else
            If Button1.Text = "PASS" Then
                Button1.BackColor = Color.Green
            Else
                If Button1.Text = "FAIL" Then
                    Button1.BackColor = Color.Red
                Else
                    If Button1.Text = "RUNNING" Then
                        Button1.BackColor = Color.Indigo
                    End If
                End If
            End If
        End If
    End Sub

    Private Sub Button2_Click(sender As Object, e As EventArgs) Handles Button2.Click
        Button1.Text = "RUNNING"
        Dim sl As String
        sl = TextBox4.Text.ToUpper
        If sl.Length = 10 And sl.Substring(0, 2) = "C5" Then
            TextBox4.Text = sl
            TextBox4.Enabled = False
            TextBox12.Text = sl
            TextBox12.Enabled = False
            TextBox10.Text = "PASS"
        Else
            TextBox12.Text = sl
            TextBox12.Enabled = False
            TextBox10.Text = "FAIL"
        End If
    End Sub

    Private Sub TextBox4_EnabledChanged(sender As Object, e As EventArgs) Handles TextBox4.EnabledChanged
        If TextBox4.Enabled = True Then
            Button2.Enabled = True
        Else
            Button2.Enabled = False
        End If
    End Sub

    Private Sub TextBox10_TextChanged(sender As Object, e As EventArgs) Handles TextBox10.TextChanged
        If TextBox10.Text = "IDLE" Then
            TextBox10.BackColor = Color.Yellow
        Else
            If TextBox10.Text = "PASS" Then
                TextBox10.BackColor = Color.Green
                Call T2.Show()
            Else
                If TextBox10.Text = "FAIL" Then
                    TextBox10.BackColor = Color.Red
                    Button1.Text = "FAIL"
                    Dim iret As Object = MsgBox(TextBox12.Text & " " & "HAS FAILED FOR INCORRECT SERRIAL NUMBER", vbInformation + vbOKOnly)
                    If iret = vbOK Then
                        Dim name As String = TextBox4.Text
                        Dim file As System.IO.StreamWriter
                        Dim path As String = (CONFIG.TextBox6.Text.ToString) & name & "_" & "FAIL" + ".txt"
                        If System.IO.File.Exists(path) Then
                            file = My.Computer.FileSystem.OpenTextFileWriter(path, True)
                            file.WriteLine("")
                            file.WriteLine("1 -------USERNAME -------" & TextBox3.Text & " -------")
                            file.WriteLine("2 -------SERIAL NUMBER -------" & TextBox4.Text & " -------")
                            file.WriteLine("3 -------RESULT -------FAIL -------")
                            file.WriteLine("*****END OF DATA****" & DateAndTime.Now.ToString)
                            file.Close()
                            TextBox12.Text = ""
                            TextBox12.Enabled = True
                            TextBox10.Text = "IDLE"
                            TextBox4.Text = ""
                        Else
                            System.IO.File.CreateText(path).Dispose()
                            file = My.Computer.FileSystem.OpenTextFileWriter(path, True)
                            file.WriteLine("")
                            file.WriteLine("1 -------USERNAME -------" & TextBox3.Text & " -------")
                            file.WriteLine("2 -------SERIAL NUMBER -------" & TextBox4.Text & " -------")
                            file.WriteLine("3 -------RESULT -------FAIL -------")
                            file.WriteLine("*****END OF DATA****" & DateAndTime.Now.ToString)
                            file.Close()
                            TextBox12.Text = ""
                            TextBox12.Enabled = True
                            TextBox10.Text = "IDLE"
                            TextBox4.Text = ""
                        End If
                    End If
                End If
            End If
        End If
    End Sub

    Private Sub TextBox15_TextChanged(sender As Object, e As EventArgs)
        If TextBox15.Text = "IDLE" Then
            TextBox15.BackColor = Color.Yellow
        Else
            If TextBox15.Text = "PASS" Then
                TextBox15.BackColor = Color.Green
            Else
                If TextBox15.Text = "FAIL" Then
                    TextBox15.BackColor = Color.Red
                End If
            End If
        End If
    End Sub
    Private Sub TextBox15_TextChanged_1(sender As Object, e As EventArgs) Handles TextBox15.TextChanged
        If TextBox15.Text = "IDLE" Then
            TextBox15.BackColor = Color.Yellow
        Else
            If TextBox15.Text = "PASS" Then
                TextBox15.BackColor = Color.Green

            Else
                If TextBox15.Text = "FAIL" Then
                    TextBox15.BackColor = Color.Red
                End If
            End If
        End If
    End Sub

    Private Sub TextBox4_Enter(sender As Object, e As EventArgs) Handles TextBox4.Enter

    End Sub
End Class