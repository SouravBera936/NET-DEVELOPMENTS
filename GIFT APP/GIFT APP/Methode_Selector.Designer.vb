<Global.Microsoft.VisualBasic.CompilerServices.DesignerGenerated()>
Partial Class Methode_Selector
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
        Dim resources As System.ComponentModel.ComponentResourceManager = New System.ComponentModel.ComponentResourceManager(GetType(Methode_Selector))
        Me.BunifuElipse1 = New Bunifu.Framework.UI.BunifuElipse(Me.components)
        Me.BunifuImageButton1 = New Bunifu.UI.WinForms.BunifuImageButton()
        Me.GroupBox1 = New System.Windows.Forms.GroupBox()
        Me.BunifuDropdown7 = New Bunifu.UI.WinForms.BunifuDropdown()
        Me.GroupBox2 = New System.Windows.Forms.GroupBox()
        Me.Label2 = New System.Windows.Forms.Label()
        Me.Label1 = New System.Windows.Forms.Label()
        Me.BunifuDropdown1 = New Bunifu.UI.WinForms.BunifuDropdown()
        Me.BunifuDropdown4 = New Bunifu.UI.WinForms.BunifuDropdown()
        Me.BunifuSnackbar1 = New Bunifu.UI.WinForms.BunifuSnackbar(Me.components)
        Me.GroupBox1.SuspendLayout()
        Me.GroupBox2.SuspendLayout()
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
        Me.BunifuImageButton1.ImageSize = New System.Drawing.Size(229, 75)
        Me.BunifuImageButton1.ImageZoomSize = New System.Drawing.Size(229, 75)
        Me.BunifuImageButton1.InitialImage = CType(resources.GetObject("BunifuImageButton1.InitialImage"), System.Drawing.Image)
        Me.BunifuImageButton1.Location = New System.Drawing.Point(162, 3)
        Me.BunifuImageButton1.Name = "BunifuImageButton1"
        Me.BunifuImageButton1.Rotation = 0
        Me.BunifuImageButton1.ShowActiveImage = True
        Me.BunifuImageButton1.ShowCursorChanges = True
        Me.BunifuImageButton1.ShowImageBorders = True
        Me.BunifuImageButton1.ShowSizeMarkers = False
        Me.BunifuImageButton1.Size = New System.Drawing.Size(229, 75)
        Me.BunifuImageButton1.TabIndex = 10
        Me.BunifuImageButton1.ToolTipText = ""
        Me.BunifuImageButton1.WaitOnLoad = False
        Me.BunifuImageButton1.Zoom = 0
        Me.BunifuImageButton1.ZoomSpeed = 10
        '
        'GroupBox1
        '
        Me.GroupBox1.Controls.Add(Me.BunifuDropdown7)
        Me.GroupBox1.Location = New System.Drawing.Point(12, 75)
        Me.GroupBox1.Name = "GroupBox1"
        Me.GroupBox1.Size = New System.Drawing.Size(532, 61)
        Me.GroupBox1.TabIndex = 11
        Me.GroupBox1.TabStop = False
        Me.GroupBox1.Text = "Select Methode"
        '
        'BunifuDropdown7
        '
        Me.BunifuDropdown7.BackColor = System.Drawing.Color.Transparent
        Me.BunifuDropdown7.BackgroundColor = System.Drawing.Color.White
        Me.BunifuDropdown7.BorderColor = System.Drawing.Color.Silver
        Me.BunifuDropdown7.BorderRadius = 17
        Me.BunifuDropdown7.Color = System.Drawing.Color.Silver
        Me.BunifuDropdown7.Cursor = System.Windows.Forms.Cursors.Hand
        Me.BunifuDropdown7.Direction = Bunifu.UI.WinForms.BunifuDropdown.Directions.Down
        Me.BunifuDropdown7.DisabledBackColor = System.Drawing.Color.FromArgb(CType(CType(240, Byte), Integer), CType(CType(240, Byte), Integer), CType(CType(240, Byte), Integer))
        Me.BunifuDropdown7.DisabledBorderColor = System.Drawing.Color.FromArgb(CType(CType(204, Byte), Integer), CType(CType(204, Byte), Integer), CType(CType(204, Byte), Integer))
        Me.BunifuDropdown7.DisabledColor = System.Drawing.Color.FromArgb(CType(CType(240, Byte), Integer), CType(CType(240, Byte), Integer), CType(CType(240, Byte), Integer))
        Me.BunifuDropdown7.DisabledForeColor = System.Drawing.Color.FromArgb(CType(CType(109, Byte), Integer), CType(CType(109, Byte), Integer), CType(CType(109, Byte), Integer))
        Me.BunifuDropdown7.DisabledIndicatorColor = System.Drawing.Color.DarkGray
        Me.BunifuDropdown7.DrawMode = System.Windows.Forms.DrawMode.OwnerDrawFixed
        Me.BunifuDropdown7.DropdownBorderThickness = Bunifu.UI.WinForms.BunifuDropdown.BorderThickness.Thin
        Me.BunifuDropdown7.DropDownStyle = System.Windows.Forms.ComboBoxStyle.DropDownList
        Me.BunifuDropdown7.DropDownTextAlign = Bunifu.UI.WinForms.BunifuDropdown.TextAlign.Center
        Me.BunifuDropdown7.FillDropDown = True
        Me.BunifuDropdown7.FillIndicator = False
        Me.BunifuDropdown7.FlatStyle = System.Windows.Forms.FlatStyle.Flat
        Me.BunifuDropdown7.Font = New System.Drawing.Font("Segoe UI", 9.0!)
        Me.BunifuDropdown7.ForeColor = System.Drawing.Color.Black
        Me.BunifuDropdown7.FormattingEnabled = True
        Me.BunifuDropdown7.Icon = Nothing
        Me.BunifuDropdown7.IndicatorAlignment = Bunifu.UI.WinForms.BunifuDropdown.Indicator.Right
        Me.BunifuDropdown7.IndicatorColor = System.Drawing.Color.Gray
        Me.BunifuDropdown7.IndicatorLocation = Bunifu.UI.WinForms.BunifuDropdown.Indicator.Right
        Me.BunifuDropdown7.ItemBackColor = System.Drawing.Color.White
        Me.BunifuDropdown7.ItemBorderColor = System.Drawing.Color.White
        Me.BunifuDropdown7.ItemForeColor = System.Drawing.Color.Black
        Me.BunifuDropdown7.ItemHeight = 26
        Me.BunifuDropdown7.ItemHighLightColor = System.Drawing.Color.DodgerBlue
        Me.BunifuDropdown7.ItemHighLightForeColor = System.Drawing.Color.White
        Me.BunifuDropdown7.ItemTopMargin = 3
        Me.BunifuDropdown7.Location = New System.Drawing.Point(7, 19)
        Me.BunifuDropdown7.Name = "BunifuDropdown7"
        Me.BunifuDropdown7.Size = New System.Drawing.Size(519, 32)
        Me.BunifuDropdown7.TabIndex = 0
        Me.BunifuDropdown7.Text = Nothing
        Me.BunifuDropdown7.TextAlignment = Bunifu.UI.WinForms.BunifuDropdown.TextAlign.Center
        Me.BunifuDropdown7.TextLeftMargin = 5
        '
        'GroupBox2
        '
        Me.GroupBox2.Controls.Add(Me.Label2)
        Me.GroupBox2.Controls.Add(Me.Label1)
        Me.GroupBox2.Controls.Add(Me.BunifuDropdown1)
        Me.GroupBox2.Controls.Add(Me.BunifuDropdown4)
        Me.GroupBox2.Location = New System.Drawing.Point(12, 151)
        Me.GroupBox2.Name = "GroupBox2"
        Me.GroupBox2.Size = New System.Drawing.Size(532, 61)
        Me.GroupBox2.TabIndex = 12
        Me.GroupBox2.TabStop = False
        '
        'Label2
        '
        Me.Label2.AutoSize = True
        Me.Label2.Location = New System.Drawing.Point(18, 9)
        Me.Label2.Name = "Label2"
        Me.Label2.Size = New System.Drawing.Size(82, 13)
        Me.Label2.TabIndex = 6
        Me.Label2.Text = "Type To Read :"
        '
        'Label1
        '
        Me.Label1.AutoSize = True
        Me.Label1.Location = New System.Drawing.Point(284, 9)
        Me.Label1.Name = "Label1"
        Me.Label1.Size = New System.Drawing.Size(56, 13)
        Me.Label1.TabIndex = 5
        Me.Label1.Text = "Com Port :"
        '
        'BunifuDropdown1
        '
        Me.BunifuDropdown1.BackColor = System.Drawing.Color.Transparent
        Me.BunifuDropdown1.BackgroundColor = System.Drawing.Color.LightSkyBlue
        Me.BunifuDropdown1.BorderColor = System.Drawing.Color.Transparent
        Me.BunifuDropdown1.BorderRadius = 10
        Me.BunifuDropdown1.Color = System.Drawing.Color.Transparent
        Me.BunifuDropdown1.Cursor = System.Windows.Forms.Cursors.Hand
        Me.BunifuDropdown1.Direction = Bunifu.UI.WinForms.BunifuDropdown.Directions.Down
        Me.BunifuDropdown1.DisabledBackColor = System.Drawing.Color.FromArgb(CType(CType(240, Byte), Integer), CType(CType(240, Byte), Integer), CType(CType(240, Byte), Integer))
        Me.BunifuDropdown1.DisabledBorderColor = System.Drawing.Color.FromArgb(CType(CType(204, Byte), Integer), CType(CType(204, Byte), Integer), CType(CType(204, Byte), Integer))
        Me.BunifuDropdown1.DisabledColor = System.Drawing.Color.FromArgb(CType(CType(240, Byte), Integer), CType(CType(240, Byte), Integer), CType(CType(240, Byte), Integer))
        Me.BunifuDropdown1.DisabledForeColor = System.Drawing.Color.FromArgb(CType(CType(109, Byte), Integer), CType(CType(109, Byte), Integer), CType(CType(109, Byte), Integer))
        Me.BunifuDropdown1.DisabledIndicatorColor = System.Drawing.Color.DarkGray
        Me.BunifuDropdown1.DrawMode = System.Windows.Forms.DrawMode.OwnerDrawFixed
        Me.BunifuDropdown1.DropdownBorderThickness = Bunifu.UI.WinForms.BunifuDropdown.BorderThickness.Thick
        Me.BunifuDropdown1.DropDownStyle = System.Windows.Forms.ComboBoxStyle.DropDownList
        Me.BunifuDropdown1.DropDownTextAlign = Bunifu.UI.WinForms.BunifuDropdown.TextAlign.Center
        Me.BunifuDropdown1.FillDropDown = True
        Me.BunifuDropdown1.FillIndicator = True
        Me.BunifuDropdown1.FlatStyle = System.Windows.Forms.FlatStyle.Flat
        Me.BunifuDropdown1.Font = New System.Drawing.Font("Algerian", 12.0!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.BunifuDropdown1.ForeColor = System.Drawing.Color.Black
        Me.BunifuDropdown1.FormattingEnabled = True
        Me.BunifuDropdown1.Icon = Nothing
        Me.BunifuDropdown1.IndicatorAlignment = Bunifu.UI.WinForms.BunifuDropdown.Indicator.Right
        Me.BunifuDropdown1.IndicatorColor = System.Drawing.Color.Gray
        Me.BunifuDropdown1.IndicatorLocation = Bunifu.UI.WinForms.BunifuDropdown.Indicator.Right
        Me.BunifuDropdown1.ItemBackColor = System.Drawing.Color.White
        Me.BunifuDropdown1.ItemBorderColor = System.Drawing.Color.White
        Me.BunifuDropdown1.ItemForeColor = System.Drawing.Color.Black
        Me.BunifuDropdown1.ItemHeight = 24
        Me.BunifuDropdown1.ItemHighLightColor = System.Drawing.Color.DodgerBlue
        Me.BunifuDropdown1.ItemHighLightForeColor = System.Drawing.Color.White
        Me.BunifuDropdown1.ItemTopMargin = 3
        Me.BunifuDropdown1.Location = New System.Drawing.Point(6, 25)
        Me.BunifuDropdown1.Name = "BunifuDropdown1"
        Me.BunifuDropdown1.Size = New System.Drawing.Size(262, 30)
        Me.BunifuDropdown1.TabIndex = 4
        Me.BunifuDropdown1.Text = Nothing
        Me.BunifuDropdown1.TextAlignment = Bunifu.UI.WinForms.BunifuDropdown.TextAlign.Center
        Me.BunifuDropdown1.TextLeftMargin = 5
        '
        'BunifuDropdown4
        '
        Me.BunifuDropdown4.BackColor = System.Drawing.Color.Transparent
        Me.BunifuDropdown4.BackgroundColor = System.Drawing.Color.LightSkyBlue
        Me.BunifuDropdown4.BorderColor = System.Drawing.Color.Transparent
        Me.BunifuDropdown4.BorderRadius = 10
        Me.BunifuDropdown4.Color = System.Drawing.Color.Transparent
        Me.BunifuDropdown4.Cursor = System.Windows.Forms.Cursors.Hand
        Me.BunifuDropdown4.Direction = Bunifu.UI.WinForms.BunifuDropdown.Directions.Down
        Me.BunifuDropdown4.DisabledBackColor = System.Drawing.Color.FromArgb(CType(CType(240, Byte), Integer), CType(CType(240, Byte), Integer), CType(CType(240, Byte), Integer))
        Me.BunifuDropdown4.DisabledBorderColor = System.Drawing.Color.FromArgb(CType(CType(204, Byte), Integer), CType(CType(204, Byte), Integer), CType(CType(204, Byte), Integer))
        Me.BunifuDropdown4.DisabledColor = System.Drawing.Color.FromArgb(CType(CType(240, Byte), Integer), CType(CType(240, Byte), Integer), CType(CType(240, Byte), Integer))
        Me.BunifuDropdown4.DisabledForeColor = System.Drawing.Color.FromArgb(CType(CType(109, Byte), Integer), CType(CType(109, Byte), Integer), CType(CType(109, Byte), Integer))
        Me.BunifuDropdown4.DisabledIndicatorColor = System.Drawing.Color.DarkGray
        Me.BunifuDropdown4.DrawMode = System.Windows.Forms.DrawMode.OwnerDrawFixed
        Me.BunifuDropdown4.DropdownBorderThickness = Bunifu.UI.WinForms.BunifuDropdown.BorderThickness.Thick
        Me.BunifuDropdown4.DropDownStyle = System.Windows.Forms.ComboBoxStyle.DropDownList
        Me.BunifuDropdown4.DropDownTextAlign = Bunifu.UI.WinForms.BunifuDropdown.TextAlign.Center
        Me.BunifuDropdown4.FillDropDown = True
        Me.BunifuDropdown4.FillIndicator = True
        Me.BunifuDropdown4.FlatStyle = System.Windows.Forms.FlatStyle.Flat
        Me.BunifuDropdown4.Font = New System.Drawing.Font("Algerian", 12.0!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.BunifuDropdown4.ForeColor = System.Drawing.Color.Black
        Me.BunifuDropdown4.FormattingEnabled = True
        Me.BunifuDropdown4.Icon = Nothing
        Me.BunifuDropdown4.IndicatorAlignment = Bunifu.UI.WinForms.BunifuDropdown.Indicator.Right
        Me.BunifuDropdown4.IndicatorColor = System.Drawing.Color.Gray
        Me.BunifuDropdown4.IndicatorLocation = Bunifu.UI.WinForms.BunifuDropdown.Indicator.Right
        Me.BunifuDropdown4.ItemBackColor = System.Drawing.Color.White
        Me.BunifuDropdown4.ItemBorderColor = System.Drawing.Color.White
        Me.BunifuDropdown4.ItemForeColor = System.Drawing.Color.Black
        Me.BunifuDropdown4.ItemHeight = 24
        Me.BunifuDropdown4.ItemHighLightColor = System.Drawing.Color.DodgerBlue
        Me.BunifuDropdown4.ItemHighLightForeColor = System.Drawing.Color.White
        Me.BunifuDropdown4.ItemTopMargin = 3
        Me.BunifuDropdown4.Location = New System.Drawing.Point(274, 25)
        Me.BunifuDropdown4.Name = "BunifuDropdown4"
        Me.BunifuDropdown4.Size = New System.Drawing.Size(252, 30)
        Me.BunifuDropdown4.TabIndex = 4
        Me.BunifuDropdown4.Text = Nothing
        Me.BunifuDropdown4.TextAlignment = Bunifu.UI.WinForms.BunifuDropdown.TextAlign.Center
        Me.BunifuDropdown4.TextLeftMargin = 5
        '
        'BunifuSnackbar1
        '
        Me.BunifuSnackbar1.AllowDragging = False
        Me.BunifuSnackbar1.AllowMultipleViews = True
        Me.BunifuSnackbar1.ClickToClose = True
        Me.BunifuSnackbar1.DoubleClickToClose = True
        Me.BunifuSnackbar1.DurationAfterIdle = 3000
        Me.BunifuSnackbar1.ErrorOptions.ActionBackColor = System.Drawing.Color.FromArgb(CType(CType(255, Byte), Integer), CType(CType(255, Byte), Integer), CType(CType(255, Byte), Integer))
        Me.BunifuSnackbar1.ErrorOptions.ActionBorderColor = System.Drawing.Color.FromArgb(CType(CType(255, Byte), Integer), CType(CType(255, Byte), Integer), CType(CType(255, Byte), Integer))
        Me.BunifuSnackbar1.ErrorOptions.ActionBorderRadius = 1
        Me.BunifuSnackbar1.ErrorOptions.ActionFont = New System.Drawing.Font("Segoe UI", 8.25!, System.Drawing.FontStyle.Bold)
        Me.BunifuSnackbar1.ErrorOptions.ActionForeColor = System.Drawing.Color.Black
        Me.BunifuSnackbar1.ErrorOptions.BackColor = System.Drawing.Color.White
        Me.BunifuSnackbar1.ErrorOptions.BorderColor = System.Drawing.Color.White
        Me.BunifuSnackbar1.ErrorOptions.CloseIconColor = System.Drawing.Color.FromArgb(CType(CType(255, Byte), Integer), CType(CType(204, Byte), Integer), CType(CType(199, Byte), Integer))
        Me.BunifuSnackbar1.ErrorOptions.Font = New System.Drawing.Font("Segoe UI", 9.75!)
        Me.BunifuSnackbar1.ErrorOptions.ForeColor = System.Drawing.Color.Black
        Me.BunifuSnackbar1.ErrorOptions.Icon = CType(resources.GetObject("resource.Icon"), System.Drawing.Image)
        Me.BunifuSnackbar1.ErrorOptions.IconLeftMargin = 12
        Me.BunifuSnackbar1.FadeCloseIcon = False
        Me.BunifuSnackbar1.Host = Bunifu.UI.WinForms.BunifuSnackbar.Hosts.FormOwner
        Me.BunifuSnackbar1.InformationOptions.ActionBackColor = System.Drawing.Color.FromArgb(CType(CType(255, Byte), Integer), CType(CType(255, Byte), Integer), CType(CType(255, Byte), Integer))
        Me.BunifuSnackbar1.InformationOptions.ActionBorderColor = System.Drawing.Color.FromArgb(CType(CType(255, Byte), Integer), CType(CType(255, Byte), Integer), CType(CType(255, Byte), Integer))
        Me.BunifuSnackbar1.InformationOptions.ActionBorderRadius = 1
        Me.BunifuSnackbar1.InformationOptions.ActionFont = New System.Drawing.Font("Segoe UI", 8.25!, System.Drawing.FontStyle.Bold)
        Me.BunifuSnackbar1.InformationOptions.ActionForeColor = System.Drawing.Color.Black
        Me.BunifuSnackbar1.InformationOptions.BackColor = System.Drawing.Color.White
        Me.BunifuSnackbar1.InformationOptions.BorderColor = System.Drawing.Color.White
        Me.BunifuSnackbar1.InformationOptions.CloseIconColor = System.Drawing.Color.FromArgb(CType(CType(145, Byte), Integer), CType(CType(213, Byte), Integer), CType(CType(255, Byte), Integer))
        Me.BunifuSnackbar1.InformationOptions.Font = New System.Drawing.Font("Segoe UI", 9.75!)
        Me.BunifuSnackbar1.InformationOptions.ForeColor = System.Drawing.Color.Black
        Me.BunifuSnackbar1.InformationOptions.Icon = CType(resources.GetObject("resource.Icon1"), System.Drawing.Image)
        Me.BunifuSnackbar1.InformationOptions.IconLeftMargin = 12
        Me.BunifuSnackbar1.Margin = 10
        Me.BunifuSnackbar1.MaximumSize = New System.Drawing.Size(0, 0)
        Me.BunifuSnackbar1.MaximumViews = 7
        Me.BunifuSnackbar1.MessageRightMargin = 15
        Me.BunifuSnackbar1.MinimumSize = New System.Drawing.Size(0, 0)
        Me.BunifuSnackbar1.ShowBorders = False
        Me.BunifuSnackbar1.ShowCloseIcon = False
        Me.BunifuSnackbar1.ShowIcon = True
        Me.BunifuSnackbar1.ShowShadows = True
        Me.BunifuSnackbar1.SuccessOptions.ActionBackColor = System.Drawing.Color.FromArgb(CType(CType(255, Byte), Integer), CType(CType(255, Byte), Integer), CType(CType(255, Byte), Integer))
        Me.BunifuSnackbar1.SuccessOptions.ActionBorderColor = System.Drawing.Color.FromArgb(CType(CType(255, Byte), Integer), CType(CType(255, Byte), Integer), CType(CType(255, Byte), Integer))
        Me.BunifuSnackbar1.SuccessOptions.ActionBorderRadius = 1
        Me.BunifuSnackbar1.SuccessOptions.ActionFont = New System.Drawing.Font("Segoe UI", 8.25!, System.Drawing.FontStyle.Bold)
        Me.BunifuSnackbar1.SuccessOptions.ActionForeColor = System.Drawing.Color.Black
        Me.BunifuSnackbar1.SuccessOptions.BackColor = System.Drawing.Color.White
        Me.BunifuSnackbar1.SuccessOptions.BorderColor = System.Drawing.Color.White
        Me.BunifuSnackbar1.SuccessOptions.CloseIconColor = System.Drawing.Color.FromArgb(CType(CType(246, Byte), Integer), CType(CType(255, Byte), Integer), CType(CType(237, Byte), Integer))
        Me.BunifuSnackbar1.SuccessOptions.Font = New System.Drawing.Font("Segoe UI", 9.75!)
        Me.BunifuSnackbar1.SuccessOptions.ForeColor = System.Drawing.Color.Black
        Me.BunifuSnackbar1.SuccessOptions.Icon = CType(resources.GetObject("resource.Icon2"), System.Drawing.Image)
        Me.BunifuSnackbar1.SuccessOptions.IconLeftMargin = 12
        Me.BunifuSnackbar1.ViewsMargin = 7
        Me.BunifuSnackbar1.WarningOptions.ActionBackColor = System.Drawing.Color.FromArgb(CType(CType(255, Byte), Integer), CType(CType(255, Byte), Integer), CType(CType(255, Byte), Integer))
        Me.BunifuSnackbar1.WarningOptions.ActionBorderColor = System.Drawing.Color.FromArgb(CType(CType(255, Byte), Integer), CType(CType(255, Byte), Integer), CType(CType(255, Byte), Integer))
        Me.BunifuSnackbar1.WarningOptions.ActionBorderRadius = 1
        Me.BunifuSnackbar1.WarningOptions.ActionFont = New System.Drawing.Font("Segoe UI", 8.25!, System.Drawing.FontStyle.Bold)
        Me.BunifuSnackbar1.WarningOptions.ActionForeColor = System.Drawing.Color.Black
        Me.BunifuSnackbar1.WarningOptions.BackColor = System.Drawing.Color.White
        Me.BunifuSnackbar1.WarningOptions.BorderColor = System.Drawing.Color.White
        Me.BunifuSnackbar1.WarningOptions.CloseIconColor = System.Drawing.Color.FromArgb(CType(CType(255, Byte), Integer), CType(CType(229, Byte), Integer), CType(CType(143, Byte), Integer))
        Me.BunifuSnackbar1.WarningOptions.Font = New System.Drawing.Font("Segoe UI", 9.75!)
        Me.BunifuSnackbar1.WarningOptions.ForeColor = System.Drawing.Color.Black
        Me.BunifuSnackbar1.WarningOptions.Icon = CType(resources.GetObject("resource.Icon3"), System.Drawing.Image)
        Me.BunifuSnackbar1.WarningOptions.IconLeftMargin = 12
        Me.BunifuSnackbar1.ZoomCloseIcon = True
        '
        'Methode_Selector
        '
        Me.AutoScaleDimensions = New System.Drawing.SizeF(6.0!, 13.0!)
        Me.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font
        Me.ClientSize = New System.Drawing.Size(550, 218)
        Me.Controls.Add(Me.GroupBox2)
        Me.Controls.Add(Me.GroupBox1)
        Me.Controls.Add(Me.BunifuImageButton1)
        Me.FormBorderStyle = System.Windows.Forms.FormBorderStyle.None
        Me.Name = "Methode_Selector"
        Me.StartPosition = System.Windows.Forms.FormStartPosition.CenterScreen
        Me.Text = "Methode_Selector"
        Me.GroupBox1.ResumeLayout(False)
        Me.GroupBox2.ResumeLayout(False)
        Me.GroupBox2.PerformLayout()
        Me.ResumeLayout(False)

    End Sub

    Friend WithEvents BunifuElipse1 As Bunifu.Framework.UI.BunifuElipse
    Friend WithEvents BunifuImageButton1 As Bunifu.UI.WinForms.BunifuImageButton
    Friend WithEvents GroupBox1 As GroupBox
    Friend WithEvents BunifuDropdown7 As Bunifu.UI.WinForms.BunifuDropdown
    Friend WithEvents GroupBox2 As GroupBox
    Friend WithEvents BunifuDropdown4 As Bunifu.UI.WinForms.BunifuDropdown
    Friend WithEvents BunifuSnackbar1 As Bunifu.UI.WinForms.BunifuSnackbar
    Friend WithEvents Label2 As Label
    Friend WithEvents Label1 As Label
    Friend WithEvents BunifuDropdown1 As Bunifu.UI.WinForms.BunifuDropdown
End Class
