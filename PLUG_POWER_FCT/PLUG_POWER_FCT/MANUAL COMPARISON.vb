Public Class MANUAL_COMPARISON
    Private Sub MANUAL_COMPARISON_Load(sender As Object, e As EventArgs) Handles MyBase.Load

    End Sub

    Private Sub Button2_Click(sender As Object, e As EventArgs) Handles Button2.Click
        MAIN.ListBox1.Items.Add("MANUAL RESONSE RECIEVED FOMM USER : TRUE")
        MAIN.ListBox1.SelectedIndex = MAIN.ListBox1.Items.Count - 1
        If MAIN.TextBox1.Text = "BATTERY ID" Then
            UUT_RESULT.Label2.Text = "PASS"
            MAIN.TextBox18.Text = "TRUE"
            MAIN.Button12.Text = "PASS"
            MAIN.Button43.Text = "PASS"
            Clipboard.Clear()
            Clipboard.SetImage(PictureBox2.Image)
            FUNNCTIONLIB.CalculateProgress(5, 4)
            MAIN.TreeView1.SelectedNode.Nodes(0).Nodes(3).BackColor = Color.Green
            Me.Close()
            UUT_RESULT.ShowDialog()
        End If
    End Sub

    Private Sub Button1_Click(sender As Object, e As EventArgs) Handles Button1.Click
        MAIN.ListBox1.Items.Add("MANUAL RESONSE RECIEVED FOMM USER : FALSE")
        MAIN.ListBox1.SelectedIndex = MAIN.ListBox1.Items.Count - 1
        If MAIN.TextBox1.Text = "BATTERY ID" Then
            UUT_RESULT.Label2.Text = "FAIL"
            MAIN.TextBox18.Text = "FALSE"
            MAIN.Button12.Text = "FAIL"
            MAIN.Button43.Text = "FAIL"
            Clipboard.Clear()
            Clipboard.SetImage(PictureBox2.Image)
            FUNNCTIONLIB.CalculateProgress(5, 4)
            MAIN.TreeView1.SelectedNode.Nodes(0).Nodes(3).BackColor = Color.Red
            Me.Close()
            UUT_RESULT.ShowDialog()
        End If
    End Sub
End Class