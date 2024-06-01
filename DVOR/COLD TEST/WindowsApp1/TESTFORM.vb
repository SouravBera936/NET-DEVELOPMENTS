Public Class TESTFORM
    Private Sub TESTFORM_Load(sender As Object, e As EventArgs) Handles MyBase.Load
        TextBox3.Enabled = False
        TextBox4.Enabled = False
        TextBox1.Enabled = False
        TextBox2.Enabled = False
        TextBox5.Enabled = False
        TextBox6.Enabled = False
        Button1.Text = "IDLE"
        ComboBox1.Items.Clear()
        ComboBox1.Items.Add("T11120920-0001-RF-COUPLER")
        ComboBox1.Items.Add("T11120839-0001-BATTERY PRESENCE")
        ComboBox1.Items.Add("120888-0001 STATUS PANEL")
        Me.TabControl1.Visible = False
        TextBox7.Clear()
        TextBox7.Enabled = False
        Button1.Enabled = False
        Button1.Text = "IDLE"
        Dim todaysdate As String = String.Format("{0:dd/MM/yyyy}", DateTime.Now)
        TextBox1.Text = todaysdate
        Timer1.Enabled = True
    End Sub

    Private Sub Timer1_Tick(sender As Object, e As EventArgs) Handles Timer1.Tick
        TextBox2.Text = Date.Now.ToString("hh:mm:ss tt")
    End Sub

    Private Sub ComboBox1_SelectedIndexChanged(sender As Object, e As EventArgs) Handles ComboBox1.SelectedIndexChanged
        TabControl1.TabPages.Remove(TabPage1)
        TabControl1.TabPages.Remove(TabPage2)
        TabControl1.TabPages.Remove(TabPage14)
        If ComboBox1.SelectedItem = "T11120920-0001-RF-COUPLER" Then
            TabControl1.Visible = True
            TabControl1.TabPages.Add(TabPage1)
            TabControl1.SelectedTab = TabPage1
            TabControl2.SelectedTab = TabPage3
            TabControl1.TabPages.Remove(TabPage2)
            TabControl1.TabPages.Remove(TabPage14)
            TextBox5.Clear()
            TextBox6.Clear()
            TextBox7.Clear()
            Button1.Text = "IDLE"
            Dim boxes = {TextBox8, TextBox9, TextBox10, TextBox11, TextBox12, TextBox13, TextBox14, TextBox15, TextBox16, TextBox17, TextBox18, TextBox19, TextBox20, TextBox21, TextBox22, TextBox23, TextBox24, TextBox25, TextBox26, TextBox27, TextBox28, TextBox29, TextBox30, TextBox31, TextBox32,
                TextBox33, TextBox34, TextBox35, TextBox36, TextBox37, TextBox38, TextBox39, TextBox40, TextBox41, TextBox42, TextBox43, TextBox44, TextBox45, TextBox46, TextBox47,
                TextBox48, TextBox49, TextBox50, TextBox51, TextBox52, TextBox53, TextBox54, TextBox55, TextBox56, TextBox57, TextBox58, TextBox59, TextBox60, TextBox61, TextBox62,
                TextBox63, TextBox64, TextBox65, TextBox66, TextBox67, TextBox68, TextBox69, TextBox70, TextBox71, TextBox72, TextBox73, TextBox74, TextBox75, TextBox76, TextBox77, TextBox78, TextBox79,
                TextBox80, TextBox81, TextBox82, TextBox83, TextBox84, TextBox85, Button10, Button11, Button12, Button13, Button14, Button15}
            For Each tb In boxes
                tb.Enabled = False
            Next
            Dim box = {PictureBox2, PictureBox3, PictureBox4, PictureBox5, PictureBox6, PictureBox7, Button30, Button26, Button27, Button28, Button29, Button30, Button31, Button32, Button33, Button34,
                Button35, Button36, Button37}
            For Each bt In box
                bt.visible = False
            Next
            Dim box1 = {Button10, Button11, Button12, Button13, Button14, Button15}
            For Each bt1 In box1
                bt1.Text = "IDLE"
                bt1.BackColor = Color.White
            Next
        Else
            If ComboBox1.SelectedItem = "T11120839-0001-BATTERY PRESENCE" Then
                TabControl1.Visible = True
                TabControl1.TabPages.Add(TabPage2)
                TabControl1.SelectedTab = TabPage2
                TabControl3.SelectedTab = TabPage9
                TabControl1.TabPages.Remove(TabPage1)
                TabControl1.TabPages.Remove(TabPage14)
                TextBox5.Clear()
                TextBox6.Clear()
                TextBox7.Clear()
                Button1.Text = "IDLE"
            Else
                If ComboBox1.SelectedItem = "120888-0001 STATUS PANEL" Then
                    TabControl1.Visible = True
                    TabControl1.TabPages.Add(TabPage14)
                    TabControl1.SelectedTab = TabPage14
                    TabControl4.SelectedTab = TabPage15
                    TabControl1.TabPages.Remove(TabPage1)
                    TabControl1.TabPages.Remove(TabPage2)
                    TextBox5.Clear()
                    TextBox6.Clear()
                    TextBox7.Clear()
                    Button1.Text = "IDLE"
                    TextBox154.Clear()
                    Button39.Text = "IDLE"
                    PictureBox8.Visible = False
                    Button40.Visible = False
                    Button41.Visible = False
                    Dim boxes1 = {TextBox151, TextBox152, TextBox153, TextBox154, TextBox155, TextBox156, TextBox157, TextBox158, TextBox159, TextBox160, TextBox161,
                        TextBox162, TextBox163}
                    For Each tb2 In boxes1
                        tb2.Enabled = False
                        Button38.BackColor = Color.White
                    Next
                End If
            End If
            End If
    End Sub

    Private Sub TabPage3_Click(sender As Object, e As EventArgs) Handles TabPage3.Click

    End Sub

    Private Sub Button4_Click(sender As Object, e As EventArgs) Handles Button4.Click
        TabControl2.SelectedTab = TabPage3
    End Sub

    Private Sub Button5_Click(sender As Object, e As EventArgs) Handles Button5.Click
        TabControl2.SelectedTab = TabPage4
    End Sub

    Private Sub Button6_Click(sender As Object, e As EventArgs) Handles Button6.Click
        TabControl2.SelectedTab = TabPage5
    End Sub

    Private Sub Button7_Click(sender As Object, e As EventArgs) Handles Button7.Click
        TabControl2.SelectedTab = TabPage6
    End Sub

    Private Sub Button8_Click(sender As Object, e As EventArgs) Handles Button8.Click
        TabControl2.SelectedTab = TabPage7
    End Sub

    Private Sub Button9_Click(sender As Object, e As EventArgs) Handles Button9.Click
        TabControl2.SelectedTab = TabPage8
    End Sub

    Private Sub TabPage14_Click(sender As Object, e As EventArgs) Handles TabPage14.Click

    End Sub

    Private Sub Button2_Click(sender As Object, e As EventArgs) Handles Button2.Click
        If ComboBox1.SelectedItem = "T11120920-0001-RF-COUPLER" Then
            START.TextBox1.Text = Me.ComboBox1.SelectedItem.ToString
            START.ShowDialog()
            START.TextBox2.Clear()
        Else
            If ComboBox1.SelectedItem = "120888-0001 STATUS PANEL" Then
                START.TextBox1.Text = Me.ComboBox1.SelectedItem.ToString
                START.ShowDialog()
                START.TextBox2.Clear()
            Else
                Dim iret As Object = MsgBox("NO MODELS SELECTED TO RUN", vbCritical + vbOKOnly, "ERROR INITIALILIZING TEST")
                If iret = vbOK Then
                    'donothing
                End If
            End If
        End If
    End Sub

    Private Sub Button26_Click(sender As Object, e As EventArgs) Handles Button26.Click
        TextBox19.Text = "SHORT"
        Button10.Text = "PASS"
        Button4.BackColor = Color.Green
        PictureBox2.Visible = False
        Button26.Visible = False
        Button27.Visible = False
        TabControl2.SelectedTab = TabPage4
        PictureBox3.Visible = True
        PictureBox3.Load(CONFIG.TextBox6.Text)
        Button28.Visible = True
        Button29.Visible = True
        Button11.Text = "RUNNING"
        Button5.BackColor = Color.Yellow
        Button1.Text = "RUNNING"
    End Sub

    Private Sub Button28_Click(sender As Object, e As EventArgs) Handles Button28.Click
        TextBox24.Text = "NOT OPEN"
        Button11.Text = "FAIL"
        Button5.BackColor = Color.Red
        PictureBox3.Visible = False
        Button28.Visible = False
        Button29.Visible = False
        TabControl2.SelectedTab = TabPage5
        PictureBox4.Visible = True
        PictureBox4.Load(CONFIG.TextBox7.Text)
        Button30.Visible = True
        Button31.Visible = True
        Button12.Text = "RUNNING"
        Button6.BackColor = Color.Yellow
        If Button1.Text = "RUNNING" Then
            Button1.Text = "FAIL"
        Else
            If Button1.Text = "FAIL" Then
                Button1.Text = "FAIL"
            End If
        End If
    End Sub

    Private Sub Button27_Click(sender As Object, e As EventArgs) Handles Button27.Click
        TextBox19.Text = "NOT SHORT"
        Button10.Text = "FAIL"
        Button4.BackColor = Color.Red
        PictureBox2.Visible = False
        Button26.Visible = False
        Button27.Visible = False
        TabControl2.SelectedTab = TabPage4
        PictureBox3.Visible = True
        PictureBox3.Load(CONFIG.TextBox6.Text)
        Button28.Visible = True
        Button29.Visible = True
        Button11.Text = "RUNNING"
        Button5.BackColor = Color.Yellow
        Button1.Text = "FAIL"
    End Sub

    Private Sub Button29_Click(sender As Object, e As EventArgs) Handles Button29.Click
        TextBox24.Text = "OPEN"
        Button11.Text = "PASS"
        Button5.BackColor = Color.Green
        PictureBox3.Visible = False
        Button28.Visible = False
        Button29.Visible = False
        TabControl2.SelectedTab = TabPage5
        PictureBox4.Visible = True
        PictureBox4.Load(CONFIG.TextBox7.Text)
        Button30.Visible = True
        Button31.Visible = True
        Button12.Text = "RUNNING"
        Button6.BackColor = Color.Yellow
        If Button1.Text = "RUNNING" Then
            Button1.Text = "RUNNING"
        Else
            If Button1.Text = "FAIL" Then
                Button1.Text = "FAIL"
            End If
        End If
    End Sub

    Private Sub Button30_Click(sender As Object, e As EventArgs) Handles Button30.Click
        TextBox37.Text = "NOT OPEN"
        Button12.Text = "FAIL"
        Button6.BackColor = Color.Red
        PictureBox4.Visible = False
        Button31.Visible = False
        Button30.Visible = False
        TabControl2.SelectedTab = TabPage6
        PictureBox5.Visible = True
        PictureBox5.Load(CONFIG.TextBox8.Text)
        Button33.Visible = True
        Button32.Visible = True
        Button13.Text = "RUNNING"
        Button7.BackColor = Color.Yellow
        If Button1.Text = "RUNNING" Then
            Button1.Text = "FAIL"
        Else
            If Button1.Text = "FAIL" Then
                Button1.Text = "FAIL"
            End If
        End If
    End Sub

    Private Sub Button31_Click(sender As Object, e As EventArgs) Handles Button31.Click
        TextBox37.Text = "OPEN"
        Button12.Text = "PASS"
        Button6.BackColor = Color.Green
        PictureBox4.Visible = False
        Button31.Visible = False
        Button30.Visible = False
        TabControl2.SelectedTab = TabPage6
        PictureBox5.Visible = True
        PictureBox5.Load(CONFIG.TextBox8.Text)
        Button33.Visible = True
        Button32.Visible = True
        Button13.Text = "RUNNING"
        Button7.BackColor = Color.Yellow
        If Button1.Text = "RUNNING" Then
            Button1.Text = "RUNNING"
        Else
            If Button1.Text = "FAIL" Then
                Button1.Text = "FAIL"
            End If
        End If
    End Sub

    Private Sub Button33_Click(sender As Object, e As EventArgs) Handles Button33.Click
        TextBox50.Text = "SHORT"
        Button13.Text = "PASS"
        Button7.BackColor = Color.Green
        PictureBox5.Visible = False
        Button32.Visible = False
        Button33.Visible = False
        TabControl2.SelectedTab = TabPage7
        PictureBox6.Visible = True
        PictureBox6.Load(CONFIG.TextBox9.Text)
        Button34.Visible = True
        Button35.Visible = True
        Button14.Text = "RUNNING"
        Button8.BackColor = Color.Yellow
        If Button1.Text = "RUNNING" Then
            Button1.Text = "RUNNING"
        Else
            If Button1.Text = "FAIL" Then
                Button1.Text = "FAIL"
            End If
        End If
    End Sub

    Private Sub Button32_Click(sender As Object, e As EventArgs) Handles Button32.Click
        TextBox50.Text = "NOT SHORT"
        Button13.Text = "FAIL"
        Button7.BackColor = Color.Red
        PictureBox5.Visible = False
        Button32.Visible = False
        Button33.Visible = False
        TabControl2.SelectedTab = TabPage7
        PictureBox6.Visible = True
        PictureBox6.Load(CONFIG.TextBox9.Text)
        Button34.Visible = True
        Button35.Visible = True
        Button14.Text = "RUNNING"
        Button8.BackColor = Color.Yellow
        If Button1.Text = "RUNNING" Then
            Button1.Text = "FAIL"
        Else
            If Button1.Text = "FAIL" Then
                Button1.Text = "FAIL"
            End If
        End If
    End Sub

    Private Sub TabPage7_Click(sender As Object, e As EventArgs) Handles TabPage7.Click

    End Sub

    Private Sub Button35_Click(sender As Object, e As EventArgs) Handles Button35.Click
        TextBox63.Text = "OPEN"
        Button14.Text = "PASS"
        Button8.BackColor = Color.Green
        PictureBox6.Visible = False
        Button35.Visible = False
        Button34.Visible = False
        TabControl2.SelectedTab = TabPage8
        PictureBox7.Visible = True
        PictureBox7.Load(CONFIG.TextBox10.Text)
        Button36.Visible = True
        Button37.Visible = True
        Button15.Text = "RUNNING"
        Button9.BackColor = Color.Yellow
        If Button1.Text = "RUNNING" Then
            Button1.Text = "RUNNING"
        Else
            If Button1.Text = "FAIL" Then
                Button1.Text = "FAIL"
            End If
        End If
    End Sub

    Private Sub Button34_Click(sender As Object, e As EventArgs) Handles Button34.Click
        TextBox63.Text = "NOT OPEN"
        Button14.Text = "FAIL"
        Button8.BackColor = Color.Red
        PictureBox6.Visible = False
        Button35.Visible = False
        Button34.Visible = False
        TabControl2.SelectedTab = TabPage8
        PictureBox7.Visible = True
        PictureBox7.Load(CONFIG.TextBox10.Text)
        Button36.Visible = True
        Button37.Visible = True
        Button15.Text = "RUNNING"
        Button9.BackColor = Color.Yellow
        If Button1.Text = "RUNNING" Then
            Button1.Text = "FAIL"
        Else
            If Button1.Text = "FAIL" Then
                Button1.Text = "FAIL"
            End If
        End If
    End Sub

    Private Sub Button37_Click(sender As Object, e As EventArgs) Handles Button37.Click
        TextBox76.Text = "OPEN"
        Button15.Text = "PASS"
        Button9.BackColor = Color.Green
        PictureBox7.Visible = False
        Button36.Visible = False
        Button37.Visible = False
        TextBox6.Text = DateTime.Now.ToString
        If Button1.Text = "RUNNING" Then
            Button1.Text = "PASS"
        Else
            If Button1.Text = "FAIL" Then
                Button1.Text = "FAIL"
            End If
        End If
        Call UUTRES.Show()
    End Sub

    Private Sub Button36_Click(sender As Object, e As EventArgs) Handles Button36.Click
        TextBox76.Text = "NOT OPEN"
        Button15.Text = "FAIL"
        Button9.BackColor = Color.Red
        PictureBox7.Visible = False
        Button36.Visible = False
        Button37.Visible = False
        TextBox6.Text = DateTime.Now.ToString
        If Button1.Text = "RUNNING" Then
            Button1.Text = "FAIL"
        Else
            If Button1.Text = "FAIL" Then
                Button1.Text = "FAIL"
            End If
        End If
        Call UUTRES.Show()
    End Sub

    Private Sub Button1_Click(sender As Object, e As EventArgs) Handles Button1.Click

    End Sub

    Private Sub Button1_TextChanged(sender As Object, e As EventArgs) Handles Button1.TextChanged
        If Button1.Text = "IDLE" Then
            Button1.BackColor = Color.Yellow
        Else
            If Button1.Text = "RUNNING" Then
                Button1.BackColor = Color.DarkOrange
            Else
                If Button1.Text = "PASS" Then
                    Button1.BackColor = Color.Green
                Else
                    If Button1.Text = "FAIL" Then
                        Button1.BackColor = Color.Red
                    End If
                End If
            End If
        End If
    End Sub

    Private Sub Button10_Click(sender As Object, e As EventArgs) Handles Button10.Click

    End Sub

    Private Sub Button10_TextChanged(sender As Object, e As EventArgs) Handles Button10.TextChanged
        If Button10.Text = "IDLE" Then
            Button10.BackColor = Color.Yellow
        Else
            If Button10.Text = "RUNNING" Then
                Button10.BackColor = Color.DarkOrange
            Else
                If Button10.Text = "PASS" Then
                    Button10.BackColor = Color.Green
                Else
                    If Button10.Text = "FAIL" Then
                        Button10.BackColor = Color.Red
                    End If
                End If
            End If
        End If
    End Sub

    Private Sub Button11_Click(sender As Object, e As EventArgs) Handles Button11.Click

    End Sub

    Private Sub Button11_TextChanged(sender As Object, e As EventArgs) Handles Button11.TextChanged
        If Button11.Text = "IDLE" Then
            Button11.BackColor = Color.Yellow
        Else
            If Button11.Text = "RUNNING" Then
                Button11.BackColor = Color.DarkOrange
            Else
                If Button11.Text = "PASS" Then
                    Button11.BackColor = Color.Green
                Else
                    If Button11.Text = "FAIL" Then
                        Button11.BackColor = Color.Red
                    End If
                End If
            End If
        End If
    End Sub

    Private Sub Button12_Click(sender As Object, e As EventArgs) Handles Button12.Click

    End Sub

    Private Sub Button12_TextChanged(sender As Object, e As EventArgs) Handles Button12.TextChanged
        If Button12.Text = "IDLE" Then
            Button12.BackColor = Color.Yellow
        Else
            If Button12.Text = "RUNNING" Then
                Button12.BackColor = Color.DarkOrange
            Else
                If Button12.Text = "PASS" Then
                    Button12.BackColor = Color.Green
                Else
                    If Button12.Text = "FAIL" Then
                        Button12.BackColor = Color.Red
                    End If
                End If
            End If
        End If
    End Sub

    Private Sub Button13_Click(sender As Object, e As EventArgs) Handles Button13.Click

    End Sub

    Private Sub Button13_TextChanged(sender As Object, e As EventArgs) Handles Button13.TextChanged
        If Button13.Text = "IDLE" Then
            Button13.BackColor = Color.Yellow
        Else
            If Button13.Text = "RUNNING" Then
                Button13.BackColor = Color.DarkOrange
            Else
                If Button13.Text = "PASS" Then
                    Button13.BackColor = Color.Green
                Else
                    If Button13.Text = "FAIL" Then
                        Button13.BackColor = Color.Red
                    End If
                End If
            End If
        End If
    End Sub

    Private Sub Button14_Click(sender As Object, e As EventArgs) Handles Button14.Click

    End Sub

    Private Sub Button14_TextChanged(sender As Object, e As EventArgs) Handles Button14.TextChanged
        If Button14.Text = "IDLE" Then
            Button14.BackColor = Color.Yellow
        Else
            If Button14.Text = "RUNNING" Then
                Button14.BackColor = Color.DarkOrange
            Else
                If Button14.Text = "PASS" Then
                    Button14.BackColor = Color.Green
                Else
                    If Button14.Text = "FAIL" Then
                        Button14.BackColor = Color.Red
                    End If
                End If
            End If
        End If
    End Sub

    Private Sub Button15_Click(sender As Object, e As EventArgs) Handles Button15.Click

    End Sub

    Private Sub Button15_TextChanged(sender As Object, e As EventArgs) Handles Button15.TextChanged
        If Button15.Text = "IDLE" Then
            Button15.BackColor = Color.Yellow
        Else
            If Button15.Text = "RUNNING" Then
                Button15.BackColor = Color.DarkOrange
            Else
                If Button15.Text = "PASS" Then
                    Button15.BackColor = Color.Green
                Else
                    If Button15.Text = "FAIL" Then
                        Button15.BackColor = Color.Red
                    End If
                End If
            End If
        End If
    End Sub

    Private Sub Button41_Click(sender As Object, e As EventArgs) Handles Button41.Click
        TextBox154.Text = "TRUE"
        Button39.Text = "PASS"
        Button38.BackColor = Color.Green
        PictureBox8.Visible = False
        Button40.Visible = False
        Button41.Visible = False
        Button1.Text = "PASS"
        TextBox6.Text = DateTime.Now.ToString
        Call UUTRES.Show()
    End Sub

    Private Sub Button40_Click(sender As Object, e As EventArgs) Handles Button40.Click
        TextBox154.Text = "FALSE"
        Button39.Text = "FAIL"
        Button38.BackColor = Color.Red
        PictureBox8.Visible = False
        Button40.Visible = False
        Button41.Visible = False
        Button1.Text = "FAIL"
        TextBox6.Text = DateTime.Now.ToString
        Call UUTRES.Show()
    End Sub

    Private Sub Button39_Click(sender As Object, e As EventArgs) Handles Button39.Click

    End Sub

    Private Sub Button39_TextChanged(sender As Object, e As EventArgs) Handles Button39.TextChanged
        If Button39.Text = "IDLE" Then
            Button39.BackColor = Color.Yellow
        Else
            If Button39.Text = "RUNNING" Then
                Button39.BackColor = Color.DarkOrange
            Else
                If Button39.Text = "PASS" Then
                    Button39.BackColor = Color.Green
                Else
                    If Button39.Text = "FAIL" Then
                        Button39.BackColor = Color.Red
                    End If
                End If
            End If
        End If
    End Sub

    Private Sub Button3_Click(sender As Object, e As EventArgs) Handles Button3.Click
        Me.Close()
        Call MAIN.Show()
    End Sub
End Class