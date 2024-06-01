Imports BunifuAnimatorNS
Public Class CriteriUpdater
    Dim transition As New BunifuTransition
    Private Sub BunifuButton15_Click(sender As Object, e As EventArgs) Handles BunifuButton15.Click
        transition.HideSync(Me, False, BunifuAnimatorNS.Animation.ScaleAndHorizSlide)
        Me.Close()
    End Sub

    Private Sub CriteriUpdater_Load(sender As Object, e As EventArgs) Handles MyBase.Load
        NumericUpDown1.Enabled = False
        NumericUpDown2.Enabled = False
        BunifuDropdown7.Enabled = True
        BunifuDropdown7.Items.Clear()
        BunifuDropdown7.Text = Nothing
        BunifuDropdown7.Items.Add("Monthly")
        BunifuDropdown7.Items.Add("Quaterly")
        BunifuDropdown7.Items.Add("Half Yearly")
        BunifuDropdown7.Items.Add("Yearly")
    End Sub

    Private Sub BunifuDropdown7_SelectedIndexChanged(sender As Object, e As EventArgs) Handles BunifuDropdown7.SelectedIndexChanged
        If BunifuDropdown7.SelectedItem = "Monthly" Then
            NumericUpDown1.Enabled = False
            NumericUpDown1.ResetText()
            NumericUpDown1.Value = CInt(31)
            NumericUpDown2.Minimum = 1
            NumericUpDown2.Enabled = True
            NumericUpDown2.Maximum = NumericUpDown1.Value
        ElseIf BunifuDropdown7.SelectedItem = "Quaterly" Then
            NumericUpDown1.Enabled = False
            NumericUpDown1.ResetText()
            NumericUpDown1.Value = CInt(91)
            NumericUpDown2.Minimum = 1
            NumericUpDown2.Enabled = True
            NumericUpDown2.Maximum = NumericUpDown1.Value
        ElseIf BunifuDropdown7.SelectedItem = "Half Yearly" Then
            NumericUpDown1.Enabled = False
            NumericUpDown1.ResetText()
            NumericUpDown1.Value = CInt(181)
            NumericUpDown2.Minimum = 1
            NumericUpDown2.Enabled = True
            NumericUpDown2.Maximum = NumericUpDown1.Value
        ElseIf BunifuDropdown7.SelectedItem = "Yearly" Then
            NumericUpDown1.Enabled = False
            NumericUpDown1.ResetText()
            NumericUpDown1.Value = CInt(366)
            NumericUpDown2.Minimum = 1
            NumericUpDown2.Enabled = True
            NumericUpDown2.Maximum = NumericUpDown1.Value
        End If
    End Sub

    Private Sub NumericUpDown2_ValueChanged(sender As Object, e As EventArgs) Handles NumericUpDown2.ValueChanged

    End Sub
End Class