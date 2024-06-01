Imports System.Drawing
Imports System.IO
Imports MadMilkman.Ini
Public Class SMC_TIMETEST
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

    Private Sub SMC_TIMETEST_Load(sender As Object, e As EventArgs) Handles MyBase.Load
        Try
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
                CaptureSnapshots()
                TextBox1.Text = FUNNCTIONLIB.ImageToText("C:\PLUG_POWER_FCT\PLUG_POWER_FCT\Project Files\Temp\1.jpeg")
                MAIN.TextBox126.Text = Me.TextBox1.Text
                If MAIN.TextBox127.Text = MAIN.TextBox126.Text And MAIN.TextBox125.Text = MAIN.TextBox126.Text Then
                    MAIN.Button37.Text = "PASS"
                Else
                    MAIN.Button37.Text = "FAIL"
                End If
            Next
            Me.Close()
        Catch ex As Exception
            MsgBox(ex.Message)
        End Try
    End Sub
    Private Sub CaptureSnapshots()
        If rectangles.Count > 0 Then
            Dim captureRect As Rectangle = rectangles(rectangles.Count - 1)
            Dim screenshot1 As Bitmap = CapturePictureBoxSnapshot(PictureBox1, captureRect)
            PictureBox2.Image = screenshot1
            PictureBox2.Image.Save("C:\PLUG_POWER_FCT\PLUG_POWER_FCT\Project Files\Temp\1.jpeg")
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

    Private Sub Button1_Click(sender As Object, e As EventArgs) Handles Button1.Click


    End Sub
End Class