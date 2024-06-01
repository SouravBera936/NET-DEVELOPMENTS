<Global.Microsoft.VisualBasic.CompilerServices.DesignerGenerated()>
Partial Class EmpDet
    Inherits System.Windows.Forms.Form

    'Form overrides dispose to clean up the component list.
    <System.Diagnostics.DebuggerNonUserCode()>
    Protected Overrides Sub Dispose(ByVal disposing As Boolean)
        Try
            If disposing AndAlso components IsNot Nothing Then
                components.Dispose()
            End If
        Finally
            MyBase.Dispose(disposing)
        End Try
    End Sub

    'Required by the Windows Form Designer
    Private components As System.ComponentModel.IContainer

    'NOTE: The following procedure is required by the Windows Form Designer
    'It can be modified using the Windows Form Designer.  
    'Do not modify it using the code editor.
    <System.Diagnostics.DebuggerStepThrough()>
    Private Sub InitializeComponent()
        Me.components = New System.ComponentModel.Container()
        Dim resources As System.ComponentModel.ComponentResourceManager = New System.ComponentModel.ComponentResourceManager(GetType(EmpDet))
        Me.BunifuElipse1 = New Bunifu.Framework.UI.BunifuElipse(Me.components)
        Me.BunifuImageButton1 = New Bunifu.UI.WinForms.BunifuImageButton()
        Me.ListBox1 = New System.Windows.Forms.ListBox()
        Me.Button1 = New System.Windows.Forms.Button()
        Me.Label1 = New System.Windows.Forms.Label()
        Me.TextBox1 = New System.Windows.Forms.TextBox()
        Me.Button2 = New System.Windows.Forms.Button()
        Me.SuspendLayout()
        '
        'BunifuElipse1
        '
        Me.BunifuElipse1.ElipseRadius = 5
        Me.BunifuElipse1.TargetControl = Me
        '
        'BunifuImageButton1
        '
        Me.BunifuImageButton1.ActiveImage = Nothing
        Me.BunifuImageButton1.AllowAnimations = True
        Me.BunifuImageButton1.AllowBuffering = False
        Me.BunifuImageButton1.AllowToggling = False
        Me.BunifuImageButton1.AllowZooming = False
        Me.BunifuImageButton1.AllowZoomingOnFocus = False
        Me.BunifuImageButton1.BackColor = System.Drawing.Color.Transparent
        Me.BunifuImageButton1.DialogResult = System.Windows.Forms.DialogResult.None
        Me.BunifuImageButton1.ErrorImage = CType(resources.GetObject("BunifuImageButton1.ErrorImage"), System.Drawing.Image)
        Me.BunifuImageButton1.FadeWhenInactive = False
        Me.BunifuImageButton1.Flip = Bunifu.UI.WinForms.BunifuImageButton.FlipOrientation.Normal
        Me.BunifuImageButton1.Image = Global.GIFT_APP.My.Resources.Resources.Jabil_Logo_wine
        Me.BunifuImageButton1.ImageActive = Nothing
        Me.BunifuImageButton1.ImageLocation = Nothing
        Me.BunifuImageButton1.ImageMargin = 0
        Me.BunifuImageButton1.ImageSize = New System.Drawing.Size(229, 70)
        Me.BunifuImageButton1.ImageZoomSize = New System.Drawing.Size(229, 70)
        Me.BunifuImageButton1.InitialImage = CType(resources.GetObject("BunifuImageButton1.InitialImage"), System.Drawing.Image)
        Me.BunifuImageButton1.Location = New System.Drawing.Point(52, 3)
        Me.BunifuImageButton1.Name = "BunifuImageButton1"
        Me.BunifuImageButton1.Rotation = 0
        Me.BunifuImageButton1.ShowActiveImage = True
        Me.BunifuImageButton1.ShowCursorChanges = True
        Me.BunifuImageButton1.ShowImageBorders = True
        Me.BunifuImageButton1.ShowSizeMarkers = False
        Me.BunifuImageButton1.Size = New System.Drawing.Size(229, 70)
        Me.BunifuImageButton1.TabIndex = 10
        Me.BunifuImageButton1.ToolTipText = ""
        Me.BunifuImageButton1.WaitOnLoad = False
        Me.BunifuImageButton1.Zoom = 0
        Me.BunifuImageButton1.ZoomSpeed = 10
        '
        'ListBox1
        '
        Me.ListBox1.FormattingEnabled = True
        Me.ListBox1.Location = New System.Drawing.Point(12, 55)
        Me.ListBox1.Name = "ListBox1"
        Me.ListBox1.Size = New System.Drawing.Size(306, 225)
        Me.ListBox1.TabIndex = 11
        '
        'Button1
        '
        Me.Button1.Location = New System.Drawing.Point(12, 339)
        Me.Button1.Name = "Button1"
        Me.Button1.Size = New System.Drawing.Size(306, 33)
        Me.Button1.TabIndex = 12
        Me.Button1.Text = "OK"
        Me.Button1.UseVisualStyleBackColor = True
        '
        'Label1
        '
        Me.Label1.AutoSize = True
        Me.Label1.Location = New System.Drawing.Point(9, 284)
        Me.Label1.Name = "Label1"
        Me.Label1.Size = New System.Drawing.Size(148, 13)
        Me.Label1.TabIndex = 13
        Me.Label1.Text = "COUNT OF MISSING DATA :"
        '
        'TextBox1
        '
        Me.TextBox1.Location = New System.Drawing.Point(157, 281)
        Me.TextBox1.Name = "TextBox1"
        Me.TextBox1.Size = New System.Drawing.Size(161, 20)
        Me.TextBox1.TabIndex = 14
        '
        'Button2
        '
        Me.Button2.Location = New System.Drawing.Point(12, 303)
        Me.Button2.Name = "Button2"
        Me.Button2.Size = New System.Drawing.Size(306, 33)
        Me.Button2.TabIndex = 12
        Me.Button2.Text = "SAVE TO FILE"
        Me.Button2.UseVisualStyleBackColor = True
        '
        'EmpDet
        '
        Me.AutoScaleDimensions = New System.Drawing.SizeF(6.0!, 13.0!)
        Me.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font
        Me.ClientSize = New System.Drawing.Size(330, 375)
        Me.Controls.Add(Me.TextBox1)
        Me.Controls.Add(Me.Label1)
        Me.Controls.Add(Me.Button2)
        Me.Controls.Add(Me.Button1)
        Me.Controls.Add(Me.ListBox1)
        Me.Controls.Add(Me.BunifuImageButton1)
        Me.FormBorderStyle = System.Windows.Forms.FormBorderStyle.None
        Me.Name = "EmpDet"
        Me.StartPosition = System.Windows.Forms.FormStartPosition.CenterScreen
        Me.Text = "EmpDet"
        Me.ResumeLayout(False)
        Me.PerformLayout()

    End Sub

    Friend WithEvents BunifuElipse1 As Bunifu.Framework.UI.BunifuElipse
    Friend WithEvents ListBox1 As ListBox
    Friend WithEvents BunifuImageButton1 As Bunifu.UI.WinForms.BunifuImageButton
    Friend WithEvents TextBox1 As TextBox
    Friend WithEvents Label1 As Label
    Friend WithEvents Button1 As Button
    Friend WithEvents Button2 As Button
End Class
