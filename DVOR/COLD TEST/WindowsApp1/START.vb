Public Class START
    Private Sub Button2_Click(sender As Object, e As EventArgs) Handles Button2.Click
        Dim sl As String = TextBox2.Text.ToUpper
        If TextBox1.Text = "T11120920-0001-RF-COUPLER" And sl.Length = 14 AndAlso sl.Substring(0, 2) = "CE" Then
            TESTFORM.TextBox7.Text = sl
            TESTFORM.TabControl2.SelectedTab = TESTFORM.TabPage3
            TESTFORM.PictureBox2.Visible = True
            TESTFORM.PictureBox2.Load(CONFIG.TextBox5.Text)
            TESTFORM.Button26.Visible = True
            TESTFORM.Button27.Visible = True
            TESTFORM.Button1.Text = "RUNNING"
            TESTFORM.Button10.Text = "RUNNING"
            TESTFORM.Button2.Enabled = False
            TESTFORM.Button4.BackColor = Color.Yellow
            TESTFORM.ComboBox1.Enabled = False
            TESTFORM.TextBox5.Text = DateTime.Now.ToString
            Me.Close()
        Else
            If TextBox1.Text = "120888-0001 STATUS PANEL" And sl.Length = 14 AndAlso sl.Substring(0, 2) = "CE" Then
                TESTFORM.TextBox7.Text = sl
                TESTFORM.TabControl4.SelectedTab = TESTFORM.TabPage15
                TESTFORM.PictureBox8.Visible = True
                TESTFORM.PictureBox8.Load(CONFIG.TextBox11.Text)
                TESTFORM.Button40.Visible = True
                TESTFORM.Button41.Visible = True
                TESTFORM.Button1.Text = "RUNNING"
                TESTFORM.Button39.Text = "RUNNING"
                TESTFORM.Button2.Enabled = False
                TESTFORM.Button38.BackColor = Color.Yellow
                TESTFORM.ComboBox1.Enabled = False
                TESTFORM.TextBox5.Text = DateTime.Now.ToString
                Me.Close()
            Else
                Dim iret As Object = MsgBox("INVALID SERIAL NUMBER", vbCritical + vbOKOnly, "ERROR INITIALIZATION")
                If iret = vbOK Then
                    'donothing
                End If
            End If
        End If

    End Sub

    Private Sub START_Load(sender As Object, e As EventArgs) Handles MyBase.Load
        TextBox1.Enabled = False
    End Sub

    Private Sub Button1_Click(sender As Object, e As EventArgs) Handles Button1.Click
        Me.Close()
    End Sub
End Class