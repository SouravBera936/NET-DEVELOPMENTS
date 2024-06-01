<Global.Microsoft.VisualBasic.CompilerServices.DesignerGenerated()> _
Partial Class T3
    Inherits System.Windows.Forms.Form

    'Form overrides dispose to clean up the component list.
    <System.Diagnostics.DebuggerNonUserCode()> _
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
    <System.Diagnostics.DebuggerStepThrough()> _
    Private Sub InitializeComponent()
        Dim resources As System.ComponentModel.ComponentResourceManager = New System.ComponentModel.ComponentResourceManager(GetType(T3))
        Me.Button1 = New System.Windows.Forms.Button()
        Me.Button2 = New System.Windows.Forms.Button()
        Me.pictureBox3 = New System.Windows.Forms.PictureBox()
        Me.textBox2 = New System.Windows.Forms.TextBox()
        CType(Me.pictureBox3, System.ComponentModel.ISupportInitialize).BeginInit()
        Me.SuspendLayout()
        '
        'Button1
        '
        Me.Button1.BackColor = System.Drawing.Color.Lime
        Me.Button1.Location = New System.Drawing.Point(313, 163)
        Me.Button1.Name = "Button1"
        Me.Button1.Size = New System.Drawing.Size(117, 35)
        Me.Button1.TabIndex = 52
        Me.Button1.Text = "PASS"
        Me.Button1.UseVisualStyleBackColor = False
        '
        'Button2
        '
        Me.Button2.BackColor = System.Drawing.Color.Red
        Me.Button2.Location = New System.Drawing.Point(313, 80)
        Me.Button2.Name = "Button2"
        Me.Button2.Size = New System.Drawing.Size(117, 35)
        Me.Button2.TabIndex = 53
        Me.Button2.Text = "FAIL"
        Me.Button2.UseVisualStyleBackColor = False
        '
        'pictureBox3
        '
        Me.pictureBox3.Image = CType(resources.GetObject("pictureBox3.Image"), System.Drawing.Image)
        Me.pictureBox3.Location = New System.Drawing.Point(5, 80)
        Me.pictureBox3.Name = "pictureBox3"
        Me.pictureBox3.Size = New System.Drawing.Size(302, 118)
        Me.pictureBox3.TabIndex = 50
        Me.pictureBox3.TabStop = False
        '
        'textBox2
        '
        Me.textBox2.BackColor = System.Drawing.SystemColors.ScrollBar
        Me.textBox2.Font = New System.Drawing.Font("Calibri", 11.25!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.textBox2.ForeColor = System.Drawing.SystemColors.HotTrack
        Me.textBox2.Location = New System.Drawing.Point(5, 0)
        Me.textBox2.Multiline = True
        Me.textBox2.Name = "textBox2"
        Me.textBox2.ReadOnly = True
        Me.textBox2.Size = New System.Drawing.Size(422, 74)
        Me.textBox2.TabIndex = 54
        Me.textBox2.Text = resources.GetString("textBox2.Text")
        '
        'T3
        '
        Me.AutoScaleDimensions = New System.Drawing.SizeF(6.0!, 13.0!)
        Me.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font
        Me.ClientSize = New System.Drawing.Size(435, 202)
        Me.ControlBox = False
        Me.Controls.Add(Me.textBox2)
        Me.Controls.Add(Me.Button2)
        Me.Controls.Add(Me.Button1)
        Me.Controls.Add(Me.pictureBox3)
        Me.Icon = CType(resources.GetObject("$this.Icon"), System.Drawing.Icon)
        Me.Name = "T3"
        Me.Text = "LOGIC CONFIRMATION"
        CType(Me.pictureBox3, System.ComponentModel.ISupportInitialize).EndInit()
        Me.ResumeLayout(False)
        Me.PerformLayout()

    End Sub
    Private WithEvents pictureBox3 As PictureBox
    Friend WithEvents Button1 As Button
    Friend WithEvents Button2 As Button
    Private WithEvents textBox2 As TextBox
End Class
