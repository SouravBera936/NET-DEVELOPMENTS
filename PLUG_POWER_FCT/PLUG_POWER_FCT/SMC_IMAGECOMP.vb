Imports System.IO
Imports System.Text
Imports System.Runtime.InteropServices
Imports MadMilkman.Ini
Public Class SMC_IMAGECOMP
    Dim ini As New IniFile
    Private startPoint As Point
    Private endPoint As Point
    Private isDrawing As Boolean = True
    Private rectangles As New List(Of Rectangle)
    Dim CaptureComp As Bitmap
    Dim RefComp As Bitmap
    Dim x As Integer
    Dim y As Integer
    Dim width As Integer
    Dim height As Integer
    Private Sub SMC_IMAGECOMP_Load(sender As Object, e As EventArgs) Handles MyBase.Load
        Dim line1 As Label
        For Each item As String In MAIN.ListBox2.Items
            Dim coordinates As String() = item.Split(","c)
            If coordinates.Length = 4 Then
                x = Integer.Parse(coordinates(0))
                y = Integer.Parse(coordinates(1))
                width = Integer.Parse(coordinates(2))
                height = Integer.Parse(coordinates(3))
            End If
            PictureBox1.Refresh()
            Using g As Graphics = PictureBox1.CreateGraphics()
                Dim rect As New Rectangle(x, y, width, height)
                g.DrawRectangle(Pens.Red, rect)
                rectangles.Add(New Rectangle(x, y, width, height))
            End Using
            DrawRectangleOnPictureBox2()
            CaptureSnapshots()
            If FUNNCTIONLIB.CompareImages(CaptureComp, RefComp) = True Then
                Button4.BackColor = Color.Green
                Continue For
            Else
                Button4.BackColor = Color.Red
                GoTo line1
            End If
        Next
line1:
        updatetestparameter()
        Me.Close()
    End Sub
    Private Sub DrawRectangleOnPictureBox2()
        If rectangles.Count > 0 Then
            Dim drawRect As Rectangle = rectangles(rectangles.Count - 1)
            Using g As Graphics = PictureBox2.CreateGraphics()
                g.DrawRectangle(Pens.Red, drawRect)
            End Using
        End If
    End Sub
    Private Sub CaptureSnapshots()
        If rectangles.Count > 0 Then
            Dim captureRect As Rectangle = rectangles(rectangles.Count - 1)
            Dim screenshot1 As Bitmap = CapturePictureBoxSnapshot(PictureBox1, captureRect)
            PictureBox3.Image = screenshot1
            CaptureComp = screenshot1
            Dim screenshot2 As Bitmap = CapturePictureBoxSnapshot(PictureBox2, captureRect)
            PictureBox4.Image = screenshot2
            RefComp = screenshot2
        Else
            MessageBox.Show("Please draw a rectangle first.")
        End If
    End Sub
    Private Function CapturePictureBoxSnapshot(pictureBox As PictureBox, captureRect As Rectangle) As Bitmap
        Dim screenshot As New Bitmap(captureRect.Width, captureRect.Height)
        Using g As Graphics = Graphics.FromImage(screenshot)
            g.CopyFromScreen(pictureBox.PointToScreen(captureRect.Location), Point.Empty, captureRect.Size)
        End Using
        Return screenshot
    End Function
    Function updatetestparameter()
        If TextBox3.Text = "SMC_INITIAL POWER" Then
            If Button4.BackColor = Color.Green Then
                MAIN.TreeView1.SelectedNode.Nodes(1).Nodes(0).BackColor = Color.Green
                MAIN.TextBox43.Text = "TRUE"
                MAIN.Button19.Text = "PASS"
                MAIN.Button43.Text = "RUNNING"
                MAIN.ListBox1.Items.Add($"{DateTime.Now.ToString} SMC_IMAGECOMP Function for SMC_INITIAL POWER Return : true ")
                MAIN.ListBox1.SelectedIndex = MAIN.ListBox1.Items.Count - 1
            Else
                MAIN.TreeView1.SelectedNode.Nodes(1).Nodes(0).BackColor = Color.Red
                MAIN.TextBox43.Text = "FALSE"
                MAIN.Button19.Text = "FAIL"
                MAIN.Button43.Text = "FAIL"
                MAIN.ListBox1.Items.Add($"{DateTime.Now.ToString} SMC_IMAGECOMP Function for SMC_INITIAL POWER Return : false ")
                MAIN.ListBox1.SelectedIndex = MAIN.ListBox1.Items.Count - 1
            End If
        ElseIf TextBox3.Text = "SMC_D0691_CMD" Then
            If Button4.BackColor = Color.Green Then
                MAIN.TreeView1.SelectedNode.Nodes(2).Nodes(0).BackColor = Color.Green
                MAIN.TextBox162.Text = "ON"
                MAIN.Button4.Text = "PASS"
                MAIN.Button43.Text = "RUNNING"
                MAIN.ListBox1.Items.Add($"{DateTime.Now.ToString} SMC_IMAGECOMP Function for D0691_CMD_ON Return : true ")
                MAIN.ListBox1.SelectedIndex = MAIN.ListBox1.Items.Count - 1
            Else
                MAIN.TreeView1.SelectedNode.Nodes(2).Nodes(0).BackColor = Color.Red
                MAIN.TextBox162.Text = "OFF"
                MAIN.Button4.Text = "FAIL"
                MAIN.Button43.Text = "FAIL"
                MAIN.ListBox1.Items.Add($"{DateTime.Now.ToString} SMC_IMAGECOMP Function for D0691_CMD_ON Return : false ")
                MAIN.ListBox1.SelectedIndex = MAIN.ListBox1.Items.Count - 1
            End If
        ElseIf TextBox3.Text = "SMC_D0661_CMDON" Then
            If Button4.BackColor = Color.Green Then
                MAIN.TreeView1.SelectedNode.Nodes(2).Nodes(1).BackColor = Color.Green
                MAIN.TextBox82.Text = "ON"
                MAIN.Button26.Text = "PASS"
                MAIN.Button43.Text = "RUNNING"
                MAIN.ListBox1.Items.Add($"{DateTime.Now.ToString} SMC_IMAGECOMP Function for D0661_CMD_ON Return : true ")
                MAIN.ListBox1.SelectedIndex = MAIN.ListBox1.Items.Count - 1
            Else
                MAIN.TreeView1.SelectedNode.Nodes(2).Nodes(1).BackColor = Color.Red
                MAIN.TextBox82.Text = "OFF"
                MAIN.Button26.Text = "FAIL"
                MAIN.Button43.Text = "FAIL"
                MAIN.ListBox1.Items.Add($"{DateTime.Now.ToString} SMC_IMAGECOMP Function for D0691_CMD_ON Return : false ")
                MAIN.ListBox1.SelectedIndex = MAIN.ListBox1.Items.Count - 1
            End If
        ElseIf TextBox3.Text = "SMC_D0661_CMDOFF" Then
            If Button4.BackColor = Color.Green Then
                MAIN.TreeView1.SelectedNode.Nodes(2).Nodes(2).BackColor = Color.Green
                MAIN.TextBox86.Text = "OFF"
                MAIN.Button27.Text = "PASS"
                MAIN.Button43.Text = "RUNNING"
                MAIN.ListBox1.Items.Add($"{DateTime.Now.ToString} SMC_IMAGECOMP Function for D0661_CMD_OFF Return : true ")
                MAIN.ListBox1.SelectedIndex = MAIN.ListBox1.Items.Count - 1
            Else
                MAIN.TreeView1.SelectedNode.Nodes(2).Nodes(2).BackColor = Color.Red
                MAIN.TextBox86.Text = "ON"
                MAIN.Button27.Text = "FAIL"
                MAIN.Button43.Text = "FAIL"
                MAIN.ListBox1.Items.Add($"{DateTime.Now.ToString} SMC_IMAGECOMP Function for D0661_CMD_OFF Return : false ")
                MAIN.ListBox1.SelectedIndex = MAIN.ListBox1.Items.Count - 1
            End If
        ElseIf TextBox3.Text = "AUXENCMD_ON" Then
            If Button4.BackColor = Color.Green Then
                MAIN.TreeView1.SelectedNode.Nodes(2).Nodes(3).BackColor = Color.Green
                MAIN.TextBox90.Text = "ON"
                MAIN.Button28.Text = "PASS"
                MAIN.Button43.Text = "RUNNING"
                MAIN.ListBox1.Items.Add($"{DateTime.Now.ToString} SMC_IMAGECOMP Function for AUXEN_CMD_ON Return : true ")
                MAIN.ListBox1.SelectedIndex = MAIN.ListBox1.Items.Count - 1
            Else
                MAIN.TreeView1.SelectedNode.Nodes(2).Nodes(3).BackColor = Color.Red
                MAIN.TextBox90.Text = "OFF"
                MAIN.Button28.Text = "FAIL"
                MAIN.Button43.Text = "FAIL"
                MAIN.ListBox1.Items.Add($"{DateTime.Now.ToString} SMC_IMAGECOMP Function for AUXEN_CMD_ON Return : false ")
                MAIN.ListBox1.SelectedIndex = MAIN.ListBox1.Items.Count - 1
            End If
        ElseIf TextBox3.Text = "D0100_ON" Then
            If Button4.BackColor = Color.Green Then
                MAIN.TreeView1.SelectedNode.Nodes(2).Nodes(4).BackColor = Color.Green
                MAIN.TextBox95.Text = "ON"
                MAIN.Button29.Text = "PASS"
                MAIN.Button43.Text = "RUNNING"
                MAIN.ListBox1.Items.Add($"{DateTime.Now.ToString} SMC_IMAGECOMP Function for D0100_CMD_ON Return : true ")
                MAIN.ListBox1.SelectedIndex = MAIN.ListBox1.Items.Count - 1
            Else
                MAIN.TreeView1.SelectedNode.Nodes(2).Nodes(4).BackColor = Color.Red
                MAIN.TextBox95.Text = "OFF"
                MAIN.Button29.Text = "FAIL"
                MAIN.Button43.Text = "OFF"
                MAIN.ListBox1.Items.Add($"{DateTime.Now.ToString} SMC_IMAGECOMP Function for D0100_CMD_ON Return : false ")
                MAIN.ListBox1.SelectedIndex = MAIN.ListBox1.Items.Count - 1
            End If
        ElseIf TextBox3.Text = "D0100_OFF" Then
            If Button4.BackColor = Color.Green Then
                MAIN.TreeView1.SelectedNode.Nodes(2).Nodes(5).BackColor = Color.Green
                MAIN.TextBox99.Text = "OFF"
                MAIN.Button30.Text = "PASS"
                MAIN.Button43.Text = "RUNNING"
                MAIN.ListBox1.Items.Add($"{DateTime.Now.ToString} SMC_IMAGECOMP Function for D0100_CMD_OFF Return : true ")
                MAIN.ListBox1.SelectedIndex = MAIN.ListBox1.Items.Count - 1
            Else
                MAIN.TreeView1.SelectedNode.Nodes(2).Nodes(5).BackColor = Color.Red
                MAIN.TextBox99.Text = "ON"
                MAIN.Button30.Text = "FAIL"
                MAIN.Button43.Text = "FAIL"
                MAIN.ListBox1.Items.Add($"{DateTime.Now.ToString} SMC_IMAGECOMP Function for D0100_CMD_OFF Return : false ")
                MAIN.ListBox1.SelectedIndex = MAIN.ListBox1.Items.Count - 1
            End If
        ElseIf TextBox3.Text = "D0101 ON" Then
            If Button4.BackColor = Color.Green Then
                MAIN.TreeView1.SelectedNode.Nodes(2).Nodes(6).BackColor = Color.Green
                MAIN.TextBox103.Text = "ON"
                MAIN.Button31.Text = "PASS"
                MAIN.Button43.Text = "RUNNING"
                MAIN.ListBox1.Items.Add($"{DateTime.Now.ToString} SMC_IMAGECOMP Function for D0101_CMD_ON Return : true ")
                MAIN.ListBox1.SelectedIndex = MAIN.ListBox1.Items.Count - 1
            Else
                MAIN.TreeView1.SelectedNode.Nodes(2).Nodes(6).BackColor = Color.Red
                MAIN.TextBox103.Text = "OFF"
                MAIN.Button31.Text = "FAIL"
                MAIN.Button43.Text = "FAIL"
                MAIN.ListBox1.Items.Add($"{DateTime.Now.ToString} SMC_IMAGECOMP Function for D0101_CMD_ON Return : false ")
                MAIN.ListBox1.SelectedIndex = MAIN.ListBox1.Items.Count - 1
            End If
        ElseIf TextBox3.Text = "D0101 OFF" Then
            If Button4.BackColor = Color.Green Then
                MAIN.TreeView1.SelectedNode.Nodes(2).Nodes(7).BackColor = Color.Green
                MAIN.TextBox107.Text = "OFF"
                MAIN.Button32.Text = "PASS"
                MAIN.Button43.Text = "RUNNING"
                MAIN.ListBox1.Items.Add($"{DateTime.Now.ToString} SMC_IMAGECOMP Function for D0101_CMD_OFF Return : true ")
                MAIN.ListBox1.SelectedIndex = MAIN.ListBox1.Items.Count - 1
            Else
                MAIN.TreeView1.SelectedNode.Nodes(2).Nodes(7).BackColor = Color.Red
                MAIN.TextBox107.Text = "ON"
                MAIN.Button32.Text = "FAIL"
                MAIN.Button43.Text = "FAIL"
                MAIN.ListBox1.Items.Add($"{DateTime.Now.ToString} SMC_IMAGECOMP Function for D0101_CMD_OFF Return : false ")
                MAIN.ListBox1.SelectedIndex = MAIN.ListBox1.Items.Count - 1
            End If
        ElseIf TextBox3.Text = "CMD OFF" Then
            If Button4.BackColor = Color.Green Then
                MAIN.TreeView1.SelectedNode.Nodes(2).Nodes(8).BackColor = Color.Green
                MAIN.TextBox111.Text = "OFF"
                MAIN.Button33.Text = "PASS"
                MAIN.Button43.Text = "RUNNING"
                MAIN.ListBox1.Items.Add($"{DateTime.Now.ToString} SMC_IMAGECOMP Function for CMD_SWITCH_OFF Return : true ")
                MAIN.ListBox1.SelectedIndex = MAIN.ListBox1.Items.Count - 1
            Else
                MAIN.TreeView1.SelectedNode.Nodes(2).Nodes(8).BackColor = Color.Red
                MAIN.TextBox111.Text = "ON"
                MAIN.Button33.Text = "FAIL"
                MAIN.Button43.Text = "FAIL"
                MAIN.ListBox1.Items.Add($"{DateTime.Now.ToString} SMC_IMAGECOMP Function for CMD_SWITCH_OFF Return : false ")
                MAIN.ListBox1.SelectedIndex = MAIN.ListBox1.Items.Count - 1
            End If
        ElseIf TextBox3.Text = "SW2 ON" Then
            If Button4.BackColor = Color.Green Then
                MAIN.TreeView1.SelectedNode.Nodes(2).Nodes(9).BackColor = Color.Green
                MAIN.TextBox167.Text = "ON"
                MAIN.Button5.Text = "PASS"
                MAIN.Button43.Text = "RUNNING"
                MAIN.ListBox1.Items.Add($"{DateTime.Now.ToString} SMC_IMAGECOMP Function for SW2_ON Return : true ")
                MAIN.ListBox1.SelectedIndex = MAIN.ListBox1.Items.Count - 1
            Else
                MAIN.TreeView1.SelectedNode.Nodes(2).Nodes(9).BackColor = Color.Red
                MAIN.TextBox167.Text = "OFF"
                MAIN.Button5.Text = "FAIL"
                MAIN.Button43.Text = "FAIL"
                MAIN.ListBox1.Items.Add($"{DateTime.Now.ToString} SMC_IMAGECOMP Function for SW2_ON Return : false ")
                MAIN.ListBox1.SelectedIndex = MAIN.ListBox1.Items.Count - 1
            End If
        ElseIf TextBox3.Text = "SW3 ON" Then
            If Button4.BackColor = Color.Green Then
                MAIN.TreeView1.SelectedNode.Nodes(2).Nodes(10).BackColor = Color.Green
                MAIN.TextBox171.Text = "ON"
                MAIN.Button6.Text = "PASS"
                MAIN.Button43.Text = "RUNNING"
                MAIN.ListBox1.Items.Add($"{DateTime.Now.ToString} SMC_IMAGECOMP Function for SW3_ON Return : true ")
                MAIN.ListBox1.SelectedIndex = MAIN.ListBox1.Items.Count - 1
            Else
                MAIN.TreeView1.SelectedNode.Nodes(2).Nodes(10).BackColor = Color.Red
                MAIN.TextBox171.Text = "OFF"
                MAIN.Button6.Text = "FAIL"
                MAIN.Button43.Text = "FAIL"
                MAIN.ListBox1.Items.Add($"{DateTime.Now.ToString} SMC_IMAGECOMP Function for SW3_ON Return : false ")
                MAIN.ListBox1.SelectedIndex = MAIN.ListBox1.Items.Count - 1
            End If
        ElseIf TextBox3.Text = "TOOGLECMD ON" Then
            If Button4.BackColor = Color.Green Then
                MAIN.TreeView1.SelectedNode.Nodes(2).Nodes(11).BackColor = Color.Green
                MAIN.TextBox115.Text = "ON"
                MAIN.Button34.Text = "PASS"
                MAIN.Button43.Text = "RUNNING"
                MAIN.ListBox1.Items.Add($"{DateTime.Now.ToString} SMC_IMAGECOMP Function for TOOGLE_CMD_ON Return : true ")
                MAIN.ListBox1.SelectedIndex = MAIN.ListBox1.Items.Count - 1
            Else
                MAIN.TreeView1.SelectedNode.Nodes(2).Nodes(11).BackColor = Color.Red
                MAIN.TextBox115.Text = "OFF"
                MAIN.Button34.Text = "FAIL"
                MAIN.Button43.Text = "FAIL"
                MAIN.ListBox1.Items.Add($"{DateTime.Now.ToString} SMC_IMAGECOMP Function for TOOGLE_CMD_ON Return : false ")
                MAIN.ListBox1.SelectedIndex = MAIN.ListBox1.Items.Count - 1
            End If
        ElseIf TextBox3.Text = "CAN1SILENCE ON" Then
            If Button4.BackColor = Color.Green Then
                MAIN.TreeView1.SelectedNode.Nodes(2).Nodes(12).BackColor = Color.Green
                MAIN.TextBox119.Text = "ON"
                MAIN.Button35.Text = "PASS"
                MAIN.Button43.Text = "RUNNING"
                MAIN.ListBox1.Items.Add($"{DateTime.Now.ToString} SMC_IMAGECOMP Function for CAN1_SILENCE_CMD_ON Return : true ")
                MAIN.ListBox1.SelectedIndex = MAIN.ListBox1.Items.Count - 1
            Else
                MAIN.TreeView1.SelectedNode.Nodes(2).Nodes(12).BackColor = Color.Red
                MAIN.TextBox119.Text = "OFF"
                MAIN.Button35.Text = "FAIL"
                MAIN.Button43.Text = "FAIL"
                MAIN.ListBox1.Items.Add($"{DateTime.Now.ToString} SMC_IMAGECOMP Function for CAN1_SILENCE_CMD_ON Return : false ")
                MAIN.ListBox1.SelectedIndex = MAIN.ListBox1.Items.Count - 1
            End If
        ElseIf TextBox3.Text = "CAN2SILENCE ON" Then
            If Button4.BackColor = Color.Green Then
                MAIN.TreeView1.SelectedNode.Nodes(2).Nodes(13).BackColor = Color.Green
                MAIN.TextBox123.Text = "ON"
                MAIN.Button36.Text = "PASS"
                MAIN.Button43.Text = "RUNNING"
                MAIN.ListBox1.Items.Add($"{DateTime.Now.ToString} SMC_IMAGECOMP Function for CAN2_SILENCE_CMD_ON Return : true ")
                MAIN.ListBox1.SelectedIndex = MAIN.ListBox1.Items.Count - 1
            Else
                MAIN.TreeView1.SelectedNode.Nodes(2).Nodes(13).BackColor = Color.Red
                MAIN.TextBox123.Text = "OFF"
                MAIN.Button36.Text = "FAIL"
                MAIN.Button43.Text = "FAIL"
                MAIN.ListBox1.Items.Add($"{DateTime.Now.ToString} SMC_IMAGECOMP Function for CAN2_SILENCE_CMD_ON Return : false ")
                MAIN.ListBox1.SelectedIndex = MAIN.ListBox1.Items.Count - 1
            End If
        ElseIf TextBox3.Text = "STEPPER TEST" Then
            If Button4.BackColor = Color.Green Then
                MAIN.TreeView1.SelectedNode.Nodes(4).BackColor = Color.Green
                MAIN.TextBox135.Text = "ON"
                MAIN.Button38.Text = "PASS"
                MAIN.Button43.Text = "RUNNING"
                MAIN.ListBox1.Items.Add($"{DateTime.Now.ToString} SMC_IMAGECOMP Function for STEPPER_TEST Return : true ")
                MAIN.ListBox1.SelectedIndex = MAIN.ListBox1.Items.Count - 1
            Else
                MAIN.TreeView1.SelectedNode.Nodes(4).BackColor = Color.Red
                MAIN.TextBox135.Text = "OFF"
                MAIN.Button38.Text = "FAIL"
                MAIN.Button43.Text = "FAIL"
                MAIN.ListBox1.Items.Add($"{DateTime.Now.ToString} SMC_IMAGECOMP Function for STEPPER_TEST Return : false ")
                MAIN.ListBox1.SelectedIndex = MAIN.ListBox1.Items.Count - 1
            End If
        End If

    End Function

    Private Sub Button3_Click(sender As Object, e As EventArgs) Handles Button3.Click

    End Sub
End Class