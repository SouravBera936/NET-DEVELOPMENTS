Public Class LOGIN
    Dim productversion = Application.ProductVersion.ToString
    Private Sub LOGIN_Load(sender As Object, e As EventArgs) Handles MyBase.Load
        fade_in()
        Label3.Text = productversion
        TextBox1.Clear()
        TextBox2.Clear()
        Button2.Enabled = False
    End Sub
    Public Sub fade_out()
        For FadeOut = 90 To 10 Step -10
            Me.Opacity = FadeOut / 100
            Me.Refresh()
            Threading.Thread.Sleep(50)
        Next
    End Sub
    Public Sub fade_in()
        For FadeIn = 0.0 To 1.1 Step 0.1
            Me.Opacity = FadeIn
            Me.Refresh()
            Threading.Thread.Sleep(100)
        Next
    End Sub

    Private Sub Button3_Click(sender As Object, e As EventArgs) Handles Button3.Click
        fade_out()
        Me.Close()
    End Sub

    Private Sub TextBox2_TextChanged(sender As Object, e As EventArgs) Handles TextBox2.TextChanged
        TextBox2.PasswordChar = "*"
    End Sub

    Private Sub Button1_Click(sender As Object, e As EventArgs) Handles Button1.Click
        Dim line1 As Label
        MAIN.ListBox1.Visible = True
        MAIN.ListBox2.Visible = True
line1:
        If FUNNCTIONLIB.UserValidationCheck(TextBox1.Text, TextBox2.Text, CONFIG.TextBox1.Text, CONFIG.TextBox2.Text, CONFIG.TextBox3.Text) = True Then
            Dim iret As Object = MsgBox("LOGIN SUCCESS WITH : " & TextBox1.Text, vbOKOnly + vbInformation, "UserValidationCheck")
            If iret = vbOK Then
                If MAIN.TextBox16.Text = "ADMINISTRATOR" Then
                    MAIN.LOGINToolStripMenuItem.Enabled = False
                    MAIN.LOGOUTToolStripMenuItem.Enabled = True
                    MAIN.EXITToolStripMenuItem.Enabled = True
                    MAIN.CREATEUSERToolStripMenuItem.Enabled = True
                    MAIN.REMOVEUSERToolStripMenuItem.Enabled = True
                    MAIN.CHANGEPASSWORDToolStripMenuItem.Enabled = False
                    MAIN.FIRSTPASSYIELDToolStripMenuItem.Enabled = True
                    MAIN.FINALPASSYIELDToolStripMenuItem.Enabled = True
                    MAIN.ConfrigurationsToolStripMenuItem.Enabled = True
                    MAIN.ABOUTSOFTWAREToolStripMenuItem.Enabled = True
                    MAIN.HELPToolStripMenuItem.Enabled = True
                    MAIN.SETUPToolStripMenuItem.Enabled = True
                    MAIN.UPDATEDESIGNATIONToolStripMenuItem.Enabled = True
                    MAIN.ENABLEMANUALMODEToolStripMenuItem.Enabled = True
                    MAIN.Button3.Enabled = True
                    MAIN.Button1.Enabled = False
                    fade_out()
                    Me.Close()
                    MAIN.GroupBox1.Visible = True
                ElseIf MAIN.TextBox16.Text = "OPERATOR" Then
                    MAIN.LOGINToolStripMenuItem.Enabled = False
                    MAIN.LOGOUTToolStripMenuItem.Enabled = True
                    MAIN.EXITToolStripMenuItem.Enabled = True
                    MAIN.CREATEUSERToolStripMenuItem.Enabled = False
                    MAIN.REMOVEUSERToolStripMenuItem.Enabled = False
                    MAIN.CHANGEPASSWORDToolStripMenuItem.Enabled = True
                    MAIN.FIRSTPASSYIELDToolStripMenuItem.Enabled = True
                    MAIN.FINALPASSYIELDToolStripMenuItem.Enabled = True
                    MAIN.ConfrigurationsToolStripMenuItem.Enabled = False
                    MAIN.ABOUTSOFTWAREToolStripMenuItem.Enabled = True
                    MAIN.HELPToolStripMenuItem.Enabled = True
                    MAIN.SETUPToolStripMenuItem.Enabled = False
                    MAIN.UPDATEDESIGNATIONToolStripMenuItem.Enabled = False
                    MAIN.ENABLEMANUALMODEToolStripMenuItem.Enabled = False
                    MAIN.Button3.Enabled = True
                    MAIN.Button1.Enabled = False
                    fade_out()
                    Me.Close()
                    MAIN.GroupBox1.Visible = True
                ElseIf MAIN.TextBox16.Text = "TEST_DEPT" Then
                    MAIN.LOGINToolStripMenuItem.Enabled = False
                    MAIN.LOGOUTToolStripMenuItem.Enabled = True
                    MAIN.EXITToolStripMenuItem.Enabled = True
                    MAIN.CREATEUSERToolStripMenuItem.Enabled = True
                    MAIN.REMOVEUSERToolStripMenuItem.Enabled = True
                    MAIN.CHANGEPASSWORDToolStripMenuItem.Enabled = True
                    MAIN.FIRSTPASSYIELDToolStripMenuItem.Enabled = True
                    MAIN.FINALPASSYIELDToolStripMenuItem.Enabled = True
                    MAIN.ConfrigurationsToolStripMenuItem.Enabled = False
                    MAIN.ABOUTSOFTWAREToolStripMenuItem.Enabled = True
                    MAIN.HELPToolStripMenuItem.Enabled = True
                    MAIN.SETUPToolStripMenuItem.Enabled = True
                    MAIN.UPDATEDESIGNATIONToolStripMenuItem.Enabled = True
                    MAIN.ENABLEMANUALMODEToolStripMenuItem.Enabled = False
                    MAIN.Button3.Enabled = True
                    MAIN.Button1.Enabled = False
                    fade_out()
                    Me.Close()
                    MAIN.GroupBox1.Visible = True
                ElseIf MAIN.TextBox16.Text = "DEBUG_DEPT" Then
                    MAIN.LOGINToolStripMenuItem.Enabled = False
                    MAIN.LOGOUTToolStripMenuItem.Enabled = True
                    MAIN.EXITToolStripMenuItem.Enabled = True
                    MAIN.CREATEUSERToolStripMenuItem.Enabled = False
                    MAIN.REMOVEUSERToolStripMenuItem.Enabled = False
                    MAIN.CHANGEPASSWORDToolStripMenuItem.Enabled = True
                    MAIN.FIRSTPASSYIELDToolStripMenuItem.Enabled = False
                    MAIN.FINALPASSYIELDToolStripMenuItem.Enabled = False
                    MAIN.ConfrigurationsToolStripMenuItem.Enabled = False
                    MAIN.ABOUTSOFTWAREToolStripMenuItem.Enabled = True
                    MAIN.HELPToolStripMenuItem.Enabled = True
                    MAIN.SETUPToolStripMenuItem.Enabled = False
                    MAIN.UPDATEDESIGNATIONToolStripMenuItem.Enabled = False
                    MAIN.ENABLEMANUALMODEToolStripMenuItem.Enabled = True
                    MAIN.Button3.Enabled = True
                    MAIN.Button1.Enabled = False
                    fade_out()
                    Me.Close()
                    MAIN.GroupBox1.Visible = True
                End If
            End If
        Else
            Dim iret1 As Object = MsgBox("INAVLID USER OR PASSWORD", vbRetryCancel + vbInformation, "INVALID ATTEMPT!")
            If iret1 = vbRetry Then
                GoTo line1
            ElseIf iret1 = vbCancel Then
                TextBox1.Clear()
                TextBox2.Clear()
            End If
        End If
    End Sub

    Private Sub TextBox1_TextChanged(sender As Object, e As EventArgs) Handles TextBox1.TextChanged
        If TextBox1.Text = "" Then
            Button2.Enabled = False
        Else
            Button2.Enabled = True
        End If
    End Sub

    Private Sub Button2_Click(sender As Object, e As EventArgs) Handles Button2.Click
        PASSWORD_RST.TextBox1.Text = Me.TextBox1.Text
        PASSWORD_RST.ShowDialog()
    End Sub

    Private Sub Button4_Click(sender As Object, e As EventArgs) 

    End Sub


End Class