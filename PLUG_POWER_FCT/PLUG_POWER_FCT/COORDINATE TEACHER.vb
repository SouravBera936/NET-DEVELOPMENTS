Imports System.IO
Imports System.Text
Imports System.Runtime.InteropServices
Imports MadMilkman.Ini
Public Class COORDINATE_TEACHER
    Private selectionStart As Point
    Private selectionEnd As Point
    Private selectedRectangle As Rectangle
    Private tempImage As Bitmap
    Private isSelecting As Boolean = False
    Private Sub Button1_Click(sender As Object, e As EventArgs) Handles Button1.Click
        Dim openFileDialog1 As New OpenFileDialog()
        openFileDialog1.Title = "Select a Picture"
        openFileDialog1.Filter = "Image Files|*.jpg;*.jpeg;*.png;*.gif;*.bmp|All Files|*.*"
        openFileDialog1.FilterIndex = 1
        openFileDialog1.RestoreDirectory = True
        If openFileDialog1.ShowDialog() = DialogResult.OK Then
            Dim selectedFilePath As String = openFileDialog1.FileName
            PictureBox1.Image = Image.FromFile(selectedFilePath)
        End If
    End Sub
    Private Sub PictureBox1_MouseDown(sender As Object, e As MouseEventArgs) Handles PictureBox1.MouseDown
        isSelecting = True
        selectionStart = e.Location
        selectionEnd = e.Location
        PictureBox1.Invalidate()
    End Sub
    Private Sub PictureBox1_MouseMove(sender As Object, e As MouseEventArgs) Handles PictureBox1.MouseMove
        If isSelecting Then
            selectionEnd = e.Location
            PictureBox1.Invalidate()
        End If
    End Sub
    Private Function GetSelectedArea() As Rectangle
        Dim x As Integer = Math.Min(selectionStart.X, selectionEnd.X)
        Dim y As Integer = Math.Min(selectionStart.Y, selectionEnd.Y)
        Dim width As Integer = Math.Abs(selectionEnd.X - selectionStart.X)
        Dim height As Integer = Math.Abs(selectionEnd.Y - selectionStart.Y)

        Return New Rectangle(x, y, width, height)
    End Function
    Private Sub PictureBox1_MouseUp(sender As Object, e As MouseEventArgs) Handles PictureBox1.MouseUp
        isSelecting = False
        selectedRectangle = GetSelectedArea()
        TextBox1.Text = selectedRectangle.X.ToString
        TextBox2.Text = selectedRectangle.Y.ToString
        TextBox3.Text = selectedRectangle.Width.ToString
        TextBox4.Text = selectedRectangle.Height.ToString
        'MessageBox.Show($"Selected Area Coordinates: X={selectedRectangle.X}, Y={selectedRectangle.Y}, Width={selectedRectangle.Width}, Height={selectedRectangle.Height}")
        PictureBox1.Invalidate()
    End Sub
    Private Sub PictureBox1_Paint(sender As Object, e As PaintEventArgs) Handles PictureBox1.Paint
        If isSelecting Then
            Using pen As New Pen(Color.Red, 2)
                e.Graphics.DrawRectangle(pen, selectedRectangle)
            End Using
        End If
    End Sub

    Private Sub Button3_Click(sender As Object, e As EventArgs) Handles Button3.Click
        TextBox1.Clear()
        TextBox2.Clear()
        TextBox3.Clear()
        TextBox4.Clear()
    End Sub
End Class