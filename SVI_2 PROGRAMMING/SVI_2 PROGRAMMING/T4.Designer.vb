<Global.Microsoft.VisualBasic.CompilerServices.DesignerGenerated()> _
Partial Class T4
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
        Dim resources As System.ComponentModel.ComponentResourceManager = New System.ComponentModel.ComponentResourceManager(GetType(T4))
        Me.textBox2 = New System.Windows.Forms.TextBox()
        Me.Button1 = New System.Windows.Forms.Button()
        Me.SuspendLayout()
        '
        'textBox2
        '
        Me.textBox2.BackColor = System.Drawing.SystemColors.ScrollBar
        Me.textBox2.Font = New System.Drawing.Font("Calibri", 11.25!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.textBox2.ForeColor = System.Drawing.SystemColors.HotTrack
        Me.textBox2.Location = New System.Drawing.Point(3, 2)
        Me.textBox2.Multiline = True
        Me.textBox2.Name = "textBox2"
        Me.textBox2.ReadOnly = True
        Me.textBox2.Size = New System.Drawing.Size(422, 74)
        Me.textBox2.TabIndex = 52
        Me.textBox2.Text = "1. CHANGE THE ""ComPort"" TO 9                                                     " &
    "                 2. CLICK ON THE ""Firmware"" & LOAD THE LATEST FIRMWARE PROVIDED " &
    "BY THE RESPECTIVE TEST ENGINNER  "
        '
        'Button1
        '
        Me.Button1.Location = New System.Drawing.Point(325, 82)
        Me.Button1.Name = "Button1"
        Me.Button1.Size = New System.Drawing.Size(100, 30)
        Me.Button1.TabIndex = 53
        Me.Button1.Text = "OK"
        Me.Button1.UseVisualStyleBackColor = True
        '
        'T4
        '
        Me.AutoScaleDimensions = New System.Drawing.SizeF(6.0!, 13.0!)
        Me.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font
        Me.ClientSize = New System.Drawing.Size(428, 113)
        Me.ControlBox = False
        Me.Controls.Add(Me.Button1)
        Me.Controls.Add(Me.textBox2)
        Me.Icon = CType(resources.GetObject("$this.Icon"), System.Drawing.Icon)
        Me.Name = "T4"
        Me.Text = "MnMMATest CONFRIGURATION"
        Me.ResumeLayout(False)
        Me.PerformLayout()

    End Sub

    Private WithEvents textBox2 As TextBox
    Friend WithEvents Button1 As Button
End Class
