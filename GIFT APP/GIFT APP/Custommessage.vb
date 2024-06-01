Public Class Custommessage
    Private Shared msgBoxForm As Form
    Private Shared slideInTimer As Timer
    Private Shared slideOutTimer As Timer
    Private Shared closing As Boolean = False
    Public Shared Sub Show(message As String, Optional backColor As Color = Nothing, Optional radius As Integer = 10)
        Dim activeForm = Form.ActiveForm
        If activeForm IsNot Nothing Then
            Dim msgBoxForm As New Form()
            msgBoxForm.FormBorderStyle = FormBorderStyle.None
            msgBoxForm.StartPosition = FormStartPosition.Manual
            msgBoxForm.ShowInTaskbar = False
            Dim x = activeForm.Location.X + activeForm.Width - 250
            Dim y = activeForm.Location.Y + activeForm.Height \ 2 + 110
            msgBoxForm.Location = New Point(x, y)
            Dim label As New Windows.Forms.Label()
            label.Text = message
            label.AutoSize = True
            label.MaximumSize = New Size(200, 0)
            label.Dock = DockStyle.Fill
            label.TextAlign = ContentAlignment.MiddleCenter
            msgBoxForm.Controls.Add(label)
            msgBoxForm.Size = New Size(Math.Min(label.PreferredWidth + 50, 250), Math.Min(label.PreferredHeight + 10, 100))
            If backColor <> Nothing Then
                msgBoxForm.BackColor = backColor
            End If
            Dim path As New Drawing2D.GraphicsPath()
            Dim rect As New System.Drawing.Rectangle(0, 0, msgBoxForm.Width, msgBoxForm.Height)
            path.StartFigure()
            path.AddArc(rect.X, rect.Y, radius * 2, radius * 2, 180, 90)
            path.AddLine(rect.X + radius, rect.Y, rect.Right - radius * 2, rect.Y)
            path.AddArc(rect.Right - radius * 2, rect.Y, radius * 2, radius * 2, 270, 90)
            path.AddLine(rect.Right, rect.Y + radius * 2, rect.Right, rect.Bottom - radius * 2)
            path.AddArc(rect.Right - radius * 2, rect.Bottom - radius * 2, radius * 2, radius * 2, 0, 90)
            path.AddLine(rect.Right - radius * 2, rect.Bottom, rect.X + radius * 2, rect.Bottom)
            path.AddArc(rect.X, rect.Bottom - radius * 2, radius * 2, radius * 2, 90, 90)
            path.CloseFigure()
            msgBoxForm.Region = New Region(path)
            Dim closeButton As New PictureBox()
            closeButton.Image = System.Drawing.SystemIcons.Error.ToBitmap
            closeButton.Size = New Size(30, 30)
            closeButton.SizeMode = PictureBoxSizeMode.Zoom
            closeButton.Location = New Point(msgBoxForm.Width - closeButton.Width - 10, 10)
            AddHandler closeButton.Click, Sub(sender, e)
                                              msgBoxForm.Close()
                                          End Sub
            msgBoxForm.Controls.Add(closeButton)

            slideInTimer = New Timer()
            slideInTimer.Interval = 10
            AddHandler slideInTimer.Tick, AddressOf SlideInTimer_Tick
            slideInTimer.Start()

            activeForm.BeginInvoke(Sub()
                                       msgBoxForm.ShowDialog(activeForm)
                                   End Sub)
        End If
    End Sub
    Private Shared Sub SlideInTimer_Tick(sender As Object, e As EventArgs)
        If msgBoxForm.Location.X > Form.ActiveForm.Location.X + Form.ActiveForm.Width - msgBoxForm.Width Then
            msgBoxForm.Location = New Point(msgBoxForm.Location.X - 10, msgBoxForm.Location.Y)
        Else
            slideInTimer.Stop()
        End If
    End Sub
    Private Shared Sub SlideOutTimer_Tick(sender As Object, e As EventArgs)
        If msgBoxForm.Location.X < Form.ActiveForm.Location.X + Form.ActiveForm.Width Then
            msgBoxForm.Location = New Point(msgBoxForm.Location.X + 10, msgBoxForm.Location.Y)
        Else
            slideOutTimer.Stop()
            msgBoxForm.Close()
        End If
    End Sub
    Private Shared Sub CloseMessageBox()
        If Not closing Then
            closing = True
            slideInTimer.Stop()
            slideOutTimer = New Timer()
            slideOutTimer.Interval = 10
            AddHandler slideOutTimer.Tick, AddressOf SlideOutTimer_Tick
            slideOutTimer.Start()
        End If
    End Sub
End Class
