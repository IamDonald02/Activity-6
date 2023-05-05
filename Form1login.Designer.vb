<Global.Microsoft.VisualBasic.CompilerServices.DesignerGenerated()>
Partial Class Form1login
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
        Dim resources As System.ComponentModel.ComponentResourceManager = New System.ComponentModel.ComponentResourceManager(GetType(Form1login))
        Me.Guna2Elipse1 = New Guna.UI2.WinForms.Guna2Elipse(Me.components)
        Me.Guna2AnimateWindow1 = New Guna.UI2.WinForms.Guna2AnimateWindow(Me.components)
        Me.panelDragControl1 = New Guna.UI2.WinForms.Guna2DragControl(Me.components)
        Me.bgPictureBox = New Guna.UI2.WinForms.Guna2PictureBox()
        Me.Guna2ShadowForm1 = New Guna.UI2.WinForms.Guna2ShadowForm(Me.components)
        Me.bgDragControl1 = New Guna.UI2.WinForms.Guna2DragControl(Me.components)
        Me.loginPanel1 = New Guna.UI2.WinForms.Guna2Panel()
        Me.Guna2ControlBox3 = New Guna.UI2.WinForms.Guna2ControlBox()
        Me.Guna2ControlBox4 = New Guna.UI2.WinForms.Guna2ControlBox()
        Me.formDragControl1 = New Guna.UI2.WinForms.Guna2DragControl(Me.components)
        Me.registerButton = New Guna.UI2.WinForms.Guna2Button()
        Me.rememberMeToggleSwitch = New Guna.UI2.WinForms.Guna2ToggleSwitch()
        Me.forgotpassButton = New Guna.UI2.WinForms.Guna2Button()
        Me.password = New Guna.UI2.WinForms.Guna2TextBox()
        Me.username = New Guna.UI2.WinForms.Guna2TextBox()
        Me.rememLabel = New Guna.UI2.WinForms.Guna2HtmlLabel()
        Me.signinLabel = New Guna.UI2.WinForms.Guna2HtmlLabel()
        Me.helloLabel = New Guna.UI2.WinForms.Guna2HtmlLabel()
        Me.Guna2GradientButton1 = New Guna.UI2.WinForms.Guna2GradientButton()
        Me.Guna2CircleButton5 = New Guna.UI2.WinForms.Guna2CircleButton()
        Me.Guna2CircleButton4 = New Guna.UI2.WinForms.Guna2CircleButton()
        Me.Guna2CircleButton3 = New Guna.UI2.WinForms.Guna2CircleButton()
        Me.Guna2CircleButton2 = New Guna.UI2.WinForms.Guna2CircleButton()
        Me.Guna2CircleButton1 = New Guna.UI2.WinForms.Guna2CircleButton()
        Me.Guna2HtmlLabel1 = New Guna.UI2.WinForms.Guna2HtmlLabel()
        Me.Guna2PictureBox1 = New Guna.UI2.WinForms.Guna2PictureBox()
        CType(Me.bgPictureBox, System.ComponentModel.ISupportInitialize).BeginInit()
        Me.loginPanel1.SuspendLayout()
        CType(Me.Guna2PictureBox1, System.ComponentModel.ISupportInitialize).BeginInit()
        Me.SuspendLayout()
        '
        'Guna2Elipse1
        '
        Me.Guna2Elipse1.TargetControl = Me
        '
        'Guna2AnimateWindow1
        '
        Me.Guna2AnimateWindow1.TargetForm = Me
        '
        'panelDragControl1
        '
        Me.panelDragControl1.DockIndicatorTransparencyValue = 0.6R
        Me.panelDragControl1.DragStartTransparencyValue = 0.7R
        Me.panelDragControl1.TargetControl = Me.bgPictureBox
        Me.panelDragControl1.UseTransparentDrag = True
        '
        'bgPictureBox
        '
        Me.bgPictureBox.Image = CType(resources.GetObject("bgPictureBox.Image"), System.Drawing.Image)
        Me.bgPictureBox.ImageRotate = 0!
        Me.bgPictureBox.Location = New System.Drawing.Point(12, 0)
        Me.bgPictureBox.Name = "bgPictureBox"
        Me.bgPictureBox.Size = New System.Drawing.Size(400, 550)
        Me.bgPictureBox.TabIndex = 16
        Me.bgPictureBox.TabStop = False
        '
        'Guna2ShadowForm1
        '
        Me.Guna2ShadowForm1.ShadowColor = System.Drawing.Color.Transparent
        Me.Guna2ShadowForm1.TargetForm = Me
        '
        'bgDragControl1
        '
        Me.bgDragControl1.DockIndicatorTransparencyValue = 0.6R
        Me.bgDragControl1.DragStartTransparencyValue = 0.7R
        Me.bgDragControl1.TargetControl = Me.loginPanel1
        Me.bgDragControl1.UseTransparentDrag = True
        '
        'loginPanel1
        '
        Me.loginPanel1.Anchor = CType((System.Windows.Forms.AnchorStyles.Bottom Or System.Windows.Forms.AnchorStyles.Left), System.Windows.Forms.AnchorStyles)
        Me.loginPanel1.BackColor = System.Drawing.Color.Transparent
        Me.loginPanel1.Controls.Add(Me.Guna2ControlBox3)
        Me.loginPanel1.Controls.Add(Me.Guna2ControlBox4)
        Me.loginPanel1.ForeColor = System.Drawing.Color.Transparent
        Me.loginPanel1.Location = New System.Drawing.Point(-1, 0)
        Me.loginPanel1.Name = "loginPanel1"
        Me.loginPanel1.Size = New System.Drawing.Size(839, 30)
        Me.loginPanel1.TabIndex = 8
        '
        'Guna2ControlBox3
        '
        Me.Guna2ControlBox3.Anchor = CType((System.Windows.Forms.AnchorStyles.Top Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.Guna2ControlBox3.ControlBoxType = Guna.UI2.WinForms.Enums.ControlBoxType.MinimizeBox
        Me.Guna2ControlBox3.FillColor = System.Drawing.Color.Transparent
        Me.Guna2ControlBox3.IconColor = System.Drawing.Color.DimGray
        Me.Guna2ControlBox3.Location = New System.Drawing.Point(777, 3)
        Me.Guna2ControlBox3.Name = "Guna2ControlBox3"
        Me.Guna2ControlBox3.Size = New System.Drawing.Size(22, 22)
        Me.Guna2ControlBox3.TabIndex = 8
        '
        'Guna2ControlBox4
        '
        Me.Guna2ControlBox4.Anchor = CType((System.Windows.Forms.AnchorStyles.Top Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.Guna2ControlBox4.FillColor = System.Drawing.Color.Transparent
        Me.Guna2ControlBox4.IconColor = System.Drawing.Color.DimGray
        Me.Guna2ControlBox4.Location = New System.Drawing.Point(805, 3)
        Me.Guna2ControlBox4.Name = "Guna2ControlBox4"
        Me.Guna2ControlBox4.Size = New System.Drawing.Size(22, 22)
        Me.Guna2ControlBox4.TabIndex = 0
        '
        'formDragControl1
        '
        Me.formDragControl1.DockIndicatorTransparencyValue = 0.6R
        Me.formDragControl1.DragStartTransparencyValue = 0.7R
        Me.formDragControl1.TargetControl = Me
        Me.formDragControl1.UseTransparentDrag = True
        '
        'registerButton
        '
        Me.registerButton.Animated = True
        Me.registerButton.AutoRoundedCorners = True
        Me.registerButton.BackColor = System.Drawing.Color.Transparent
        Me.registerButton.BorderRadius = 17
        Me.registerButton.DisabledState.BorderColor = System.Drawing.Color.DarkGray
        Me.registerButton.DisabledState.CustomBorderColor = System.Drawing.Color.DarkGray
        Me.registerButton.DisabledState.FillColor = System.Drawing.Color.FromArgb(CType(CType(169, Byte), Integer), CType(CType(169, Byte), Integer), CType(CType(169, Byte), Integer))
        Me.registerButton.DisabledState.ForeColor = System.Drawing.Color.FromArgb(CType(CType(141, Byte), Integer), CType(CType(141, Byte), Integer), CType(CType(141, Byte), Integer))
        Me.registerButton.FillColor = System.Drawing.Color.Transparent
        Me.registerButton.Font = New System.Drawing.Font("Comic Sans MS", 9.0!)
        Me.registerButton.ForeColor = System.Drawing.SystemColors.GrayText
        Me.registerButton.ImageAlign = System.Windows.Forms.HorizontalAlignment.Right
        Me.registerButton.Location = New System.Drawing.Point(483, 501)
        Me.registerButton.Name = "registerButton"
        Me.registerButton.PressedColor = System.Drawing.Color.Transparent
        Me.registerButton.Size = New System.Drawing.Size(274, 37)
        Me.registerButton.TabIndex = 6
        Me.registerButton.Text = "Don't have an Account?  Create"
        '
        'rememberMeToggleSwitch
        '
        Me.rememberMeToggleSwitch.CheckedState.BorderColor = System.Drawing.Color.FromArgb(CType(CType(94, Byte), Integer), CType(CType(148, Byte), Integer), CType(CType(255, Byte), Integer))
        Me.rememberMeToggleSwitch.CheckedState.FillColor = System.Drawing.Color.FromArgb(CType(CType(94, Byte), Integer), CType(CType(148, Byte), Integer), CType(CType(255, Byte), Integer))
        Me.rememberMeToggleSwitch.CheckedState.InnerBorderColor = System.Drawing.Color.White
        Me.rememberMeToggleSwitch.CheckedState.InnerColor = System.Drawing.Color.White
        Me.rememberMeToggleSwitch.Location = New System.Drawing.Point(505, 338)
        Me.rememberMeToggleSwitch.Name = "rememberMeToggleSwitch"
        Me.rememberMeToggleSwitch.Size = New System.Drawing.Size(31, 19)
        Me.rememberMeToggleSwitch.TabIndex = 2
        Me.rememberMeToggleSwitch.UncheckedState.BorderColor = System.Drawing.Color.FromArgb(CType(CType(125, Byte), Integer), CType(CType(137, Byte), Integer), CType(CType(149, Byte), Integer))
        Me.rememberMeToggleSwitch.UncheckedState.FillColor = System.Drawing.Color.FromArgb(CType(CType(125, Byte), Integer), CType(CType(137, Byte), Integer), CType(CType(149, Byte), Integer))
        Me.rememberMeToggleSwitch.UncheckedState.InnerBorderColor = System.Drawing.Color.White
        Me.rememberMeToggleSwitch.UncheckedState.InnerColor = System.Drawing.Color.White
        '
        'forgotpassButton
        '
        Me.forgotpassButton.Animated = True
        Me.forgotpassButton.AutoRoundedCorners = True
        Me.forgotpassButton.BorderRadius = 9
        Me.forgotpassButton.DisabledState.BorderColor = System.Drawing.Color.DarkGray
        Me.forgotpassButton.DisabledState.CustomBorderColor = System.Drawing.Color.DarkGray
        Me.forgotpassButton.DisabledState.FillColor = System.Drawing.Color.FromArgb(CType(CType(169, Byte), Integer), CType(CType(169, Byte), Integer), CType(CType(169, Byte), Integer))
        Me.forgotpassButton.DisabledState.ForeColor = System.Drawing.Color.FromArgb(CType(CType(141, Byte), Integer), CType(CType(141, Byte), Integer), CType(CType(141, Byte), Integer))
        Me.forgotpassButton.FillColor = System.Drawing.Color.Transparent
        Me.forgotpassButton.Font = New System.Drawing.Font("Comic Sans MS", 9.0!)
        Me.forgotpassButton.ForeColor = System.Drawing.SystemColors.ButtonShadow
        Me.forgotpassButton.ImageAlign = System.Windows.Forms.HorizontalAlignment.Right
        Me.forgotpassButton.Location = New System.Drawing.Point(481, 442)
        Me.forgotpassButton.Name = "forgotpassButton"
        Me.forgotpassButton.PressedColor = System.Drawing.Color.Transparent
        Me.forgotpassButton.Size = New System.Drawing.Size(274, 20)
        Me.forgotpassButton.TabIndex = 6
        Me.forgotpassButton.Text = "Forgot Username or Password ?"
        '
        'password
        '
        Me.password.AutoRoundedCorners = True
        Me.password.BorderRadius = 17
        Me.password.Cursor = System.Windows.Forms.Cursors.IBeam
        Me.password.DefaultText = ""
        Me.password.DisabledState.BorderColor = System.Drawing.Color.FromArgb(CType(CType(208, Byte), Integer), CType(CType(208, Byte), Integer), CType(CType(208, Byte), Integer))
        Me.password.DisabledState.FillColor = System.Drawing.Color.FromArgb(CType(CType(226, Byte), Integer), CType(CType(226, Byte), Integer), CType(CType(226, Byte), Integer))
        Me.password.DisabledState.ForeColor = System.Drawing.Color.FromArgb(CType(CType(138, Byte), Integer), CType(CType(138, Byte), Integer), CType(CType(138, Byte), Integer))
        Me.password.DisabledState.PlaceholderForeColor = System.Drawing.Color.FromArgb(CType(CType(138, Byte), Integer), CType(CType(138, Byte), Integer), CType(CType(138, Byte), Integer))
        Me.password.FillColor = System.Drawing.Color.FromArgb(CType(CType(229, Byte), Integer), CType(CType(229, Byte), Integer), CType(CType(229, Byte), Integer))
        Me.password.FocusedState.BorderColor = System.Drawing.Color.FromArgb(CType(CType(94, Byte), Integer), CType(CType(148, Byte), Integer), CType(CType(255, Byte), Integer))
        Me.password.Font = New System.Drawing.Font("Comic Sans MS", 9.75!)
        Me.password.ForeColor = System.Drawing.Color.Black
        Me.password.HoverState.BorderColor = System.Drawing.Color.FromArgb(CType(CType(94, Byte), Integer), CType(CType(148, Byte), Integer), CType(CType(255, Byte), Integer))
        Me.password.IconLeft = CType(resources.GetObject("password.IconLeft"), System.Drawing.Image)
        Me.password.Location = New System.Drawing.Point(481, 285)
        Me.password.Margin = New System.Windows.Forms.Padding(3, 4, 3, 4)
        Me.password.Name = "password"
        Me.password.PasswordChar = Global.Microsoft.VisualBasic.ChrW(9679)
        Me.password.PlaceholderText = "Password"
        Me.password.SelectedText = ""
        Me.password.Size = New System.Drawing.Size(274, 37)
        Me.password.TabIndex = 1
        Me.password.UseSystemPasswordChar = True
        '
        'username
        '
        Me.username.AutoRoundedCorners = True
        Me.username.BackColor = System.Drawing.SystemColors.Control
        Me.username.BorderRadius = 17
        Me.username.Cursor = System.Windows.Forms.Cursors.IBeam
        Me.username.DefaultText = ""
        Me.username.DisabledState.BorderColor = System.Drawing.Color.FromArgb(CType(CType(208, Byte), Integer), CType(CType(208, Byte), Integer), CType(CType(208, Byte), Integer))
        Me.username.DisabledState.FillColor = System.Drawing.Color.FromArgb(CType(CType(226, Byte), Integer), CType(CType(226, Byte), Integer), CType(CType(226, Byte), Integer))
        Me.username.DisabledState.ForeColor = System.Drawing.Color.FromArgb(CType(CType(138, Byte), Integer), CType(CType(138, Byte), Integer), CType(CType(138, Byte), Integer))
        Me.username.DisabledState.PlaceholderForeColor = System.Drawing.Color.FromArgb(CType(CType(138, Byte), Integer), CType(CType(138, Byte), Integer), CType(CType(138, Byte), Integer))
        Me.username.FillColor = System.Drawing.Color.FromArgb(CType(CType(229, Byte), Integer), CType(CType(229, Byte), Integer), CType(CType(229, Byte), Integer))
        Me.username.FocusedState.BorderColor = System.Drawing.Color.FromArgb(CType(CType(94, Byte), Integer), CType(CType(148, Byte), Integer), CType(CType(255, Byte), Integer))
        Me.username.Font = New System.Drawing.Font("Comic Sans MS", 9.75!)
        Me.username.ForeColor = System.Drawing.Color.Black
        Me.username.HoverState.BorderColor = System.Drawing.Color.FromArgb(CType(CType(94, Byte), Integer), CType(CType(148, Byte), Integer), CType(CType(255, Byte), Integer))
        Me.username.IconLeft = CType(resources.GetObject("username.IconLeft"), System.Drawing.Image)
        Me.username.Location = New System.Drawing.Point(481, 231)
        Me.username.Margin = New System.Windows.Forms.Padding(3, 4, 3, 4)
        Me.username.Name = "username"
        Me.username.PasswordChar = Global.Microsoft.VisualBasic.ChrW(0)
        Me.username.PlaceholderText = "Username"
        Me.username.SelectedText = ""
        Me.username.Size = New System.Drawing.Size(274, 37)
        Me.username.TabIndex = 1
        '
        'rememLabel
        '
        Me.rememLabel.BackColor = System.Drawing.Color.Transparent
        Me.rememLabel.Font = New System.Drawing.Font("Comic Sans MS", 9.75!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.rememLabel.ForeColor = System.Drawing.SystemColors.ButtonShadow
        Me.rememLabel.Location = New System.Drawing.Point(541, 338)
        Me.rememLabel.Name = "rememLabel"
        Me.rememLabel.Size = New System.Drawing.Size(84, 20)
        Me.rememLabel.TabIndex = 3
        Me.rememLabel.Text = "Remember me"
        '
        'signinLabel
        '
        Me.signinLabel.BackColor = System.Drawing.Color.Transparent
        Me.signinLabel.Font = New System.Drawing.Font("Comic Sans MS", 9.75!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.signinLabel.ForeColor = System.Drawing.SystemColors.ButtonShadow
        Me.signinLabel.Location = New System.Drawing.Point(556, 101)
        Me.signinLabel.Name = "signinLabel"
        Me.signinLabel.Size = New System.Drawing.Size(131, 20)
        Me.signinLabel.TabIndex = 3
        Me.signinLabel.Text = "Sign in Your Account"
        '
        'helloLabel
        '
        Me.helloLabel.BackColor = System.Drawing.Color.Transparent
        Me.helloLabel.Font = New System.Drawing.Font("Comic Sans MS", 15.75!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.helloLabel.ForeColor = System.Drawing.SystemColors.ControlDarkDark
        Me.helloLabel.Location = New System.Drawing.Point(544, 63)
        Me.helloLabel.Name = "helloLabel"
        Me.helloLabel.Size = New System.Drawing.Size(154, 32)
        Me.helloLabel.TabIndex = 3
        Me.helloLabel.Text = "Welcome Back!"
        '
        'Guna2GradientButton1
        '
        Me.Guna2GradientButton1.Animated = True
        Me.Guna2GradientButton1.AutoRoundedCorners = True
        Me.Guna2GradientButton1.BackColor = System.Drawing.Color.Transparent
        Me.Guna2GradientButton1.BorderRadius = 20
        Me.Guna2GradientButton1.DisabledState.BorderColor = System.Drawing.Color.DarkGray
        Me.Guna2GradientButton1.DisabledState.CustomBorderColor = System.Drawing.Color.DarkGray
        Me.Guna2GradientButton1.DisabledState.FillColor = System.Drawing.Color.FromArgb(CType(CType(169, Byte), Integer), CType(CType(169, Byte), Integer), CType(CType(169, Byte), Integer))
        Me.Guna2GradientButton1.DisabledState.FillColor2 = System.Drawing.Color.FromArgb(CType(CType(169, Byte), Integer), CType(CType(169, Byte), Integer), CType(CType(169, Byte), Integer))
        Me.Guna2GradientButton1.DisabledState.ForeColor = System.Drawing.Color.FromArgb(CType(CType(141, Byte), Integer), CType(CType(141, Byte), Integer), CType(CType(141, Byte), Integer))
        Me.Guna2GradientButton1.FillColor = System.Drawing.Color.FromArgb(CType(CType(255, Byte), Integer), CType(CType(128, Byte), Integer), CType(CType(0, Byte), Integer))
        Me.Guna2GradientButton1.FillColor2 = System.Drawing.Color.FromArgb(CType(CType(128, Byte), Integer), CType(CType(64, Byte), Integer), CType(CType(0, Byte), Integer))
        Me.Guna2GradientButton1.Font = New System.Drawing.Font("Segoe UI Black", 9.75!, System.Drawing.FontStyle.Bold)
        Me.Guna2GradientButton1.ForeColor = System.Drawing.Color.FromArgb(CType(CType(229, Byte), Integer), CType(CType(229, Byte), Integer), CType(CType(229, Byte), Integer))
        Me.Guna2GradientButton1.GradientMode = System.Drawing.Drawing2D.LinearGradientMode.ForwardDiagonal
        Me.Guna2GradientButton1.Location = New System.Drawing.Point(534, 394)
        Me.Guna2GradientButton1.Name = "Guna2GradientButton1"
        Me.Guna2GradientButton1.Size = New System.Drawing.Size(164, 42)
        Me.Guna2GradientButton1.TabIndex = 9
        Me.Guna2GradientButton1.Text = "Sign In"
        '
        'Guna2CircleButton5
        '
        Me.Guna2CircleButton5.Animated = True
        Me.Guna2CircleButton5.BorderColor = System.Drawing.Color.FromArgb(CType(CType(213, Byte), Integer), CType(CType(218, Byte), Integer), CType(CType(223, Byte), Integer))
        Me.Guna2CircleButton5.DisabledState.BorderColor = System.Drawing.Color.DarkGray
        Me.Guna2CircleButton5.DisabledState.CustomBorderColor = System.Drawing.Color.DarkGray
        Me.Guna2CircleButton5.DisabledState.FillColor = System.Drawing.Color.FromArgb(CType(CType(169, Byte), Integer), CType(CType(169, Byte), Integer), CType(CType(169, Byte), Integer))
        Me.Guna2CircleButton5.DisabledState.ForeColor = System.Drawing.Color.FromArgb(CType(CType(141, Byte), Integer), CType(CType(141, Byte), Integer), CType(CType(141, Byte), Integer))
        Me.Guna2CircleButton5.FillColor = System.Drawing.Color.FromArgb(CType(CType(229, Byte), Integer), CType(CType(229, Byte), Integer), CType(CType(229, Byte), Integer))
        Me.Guna2CircleButton5.Font = New System.Drawing.Font("Segoe UI", 9.0!)
        Me.Guna2CircleButton5.ForeColor = System.Drawing.Color.FromArgb(CType(CType(229, Byte), Integer), CType(CType(229, Byte), Integer), CType(CType(229, Byte), Integer))
        Me.Guna2CircleButton5.Image = CType(resources.GetObject("Guna2CircleButton5.Image"), System.Drawing.Image)
        Me.Guna2CircleButton5.Location = New System.Drawing.Point(717, 141)
        Me.Guna2CircleButton5.Name = "Guna2CircleButton5"
        Me.Guna2CircleButton5.ShadowDecoration.Mode = Guna.UI2.WinForms.Enums.ShadowMode.Circle
        Me.Guna2CircleButton5.Size = New System.Drawing.Size(38, 37)
        Me.Guna2CircleButton5.TabIndex = 10
        '
        'Guna2CircleButton4
        '
        Me.Guna2CircleButton4.Animated = True
        Me.Guna2CircleButton4.BorderColor = System.Drawing.Color.FromArgb(CType(CType(213, Byte), Integer), CType(CType(218, Byte), Integer), CType(CType(223, Byte), Integer))
        Me.Guna2CircleButton4.DisabledState.BorderColor = System.Drawing.Color.DarkGray
        Me.Guna2CircleButton4.DisabledState.CustomBorderColor = System.Drawing.Color.DarkGray
        Me.Guna2CircleButton4.DisabledState.FillColor = System.Drawing.Color.FromArgb(CType(CType(169, Byte), Integer), CType(CType(169, Byte), Integer), CType(CType(169, Byte), Integer))
        Me.Guna2CircleButton4.DisabledState.ForeColor = System.Drawing.Color.FromArgb(CType(CType(141, Byte), Integer), CType(CType(141, Byte), Integer), CType(CType(141, Byte), Integer))
        Me.Guna2CircleButton4.FillColor = System.Drawing.Color.FromArgb(CType(CType(229, Byte), Integer), CType(CType(229, Byte), Integer), CType(CType(229, Byte), Integer))
        Me.Guna2CircleButton4.Font = New System.Drawing.Font("Segoe UI", 9.0!)
        Me.Guna2CircleButton4.ForeColor = System.Drawing.Color.FromArgb(CType(CType(229, Byte), Integer), CType(CType(229, Byte), Integer), CType(CType(229, Byte), Integer))
        Me.Guna2CircleButton4.Image = CType(resources.GetObject("Guna2CircleButton4.Image"), System.Drawing.Image)
        Me.Guna2CircleButton4.Location = New System.Drawing.Point(658, 141)
        Me.Guna2CircleButton4.Name = "Guna2CircleButton4"
        Me.Guna2CircleButton4.ShadowDecoration.Mode = Guna.UI2.WinForms.Enums.ShadowMode.Circle
        Me.Guna2CircleButton4.Size = New System.Drawing.Size(38, 37)
        Me.Guna2CircleButton4.TabIndex = 11
        '
        'Guna2CircleButton3
        '
        Me.Guna2CircleButton3.Animated = True
        Me.Guna2CircleButton3.BorderColor = System.Drawing.Color.FromArgb(CType(CType(213, Byte), Integer), CType(CType(218, Byte), Integer), CType(CType(223, Byte), Integer))
        Me.Guna2CircleButton3.DisabledState.BorderColor = System.Drawing.Color.DarkGray
        Me.Guna2CircleButton3.DisabledState.CustomBorderColor = System.Drawing.Color.DarkGray
        Me.Guna2CircleButton3.DisabledState.FillColor = System.Drawing.Color.FromArgb(CType(CType(169, Byte), Integer), CType(CType(169, Byte), Integer), CType(CType(169, Byte), Integer))
        Me.Guna2CircleButton3.DisabledState.ForeColor = System.Drawing.Color.FromArgb(CType(CType(141, Byte), Integer), CType(CType(141, Byte), Integer), CType(CType(141, Byte), Integer))
        Me.Guna2CircleButton3.FillColor = System.Drawing.Color.FromArgb(CType(CType(229, Byte), Integer), CType(CType(229, Byte), Integer), CType(CType(229, Byte), Integer))
        Me.Guna2CircleButton3.Font = New System.Drawing.Font("Segoe UI", 9.0!)
        Me.Guna2CircleButton3.ForeColor = System.Drawing.Color.FromArgb(CType(CType(229, Byte), Integer), CType(CType(229, Byte), Integer), CType(CType(229, Byte), Integer))
        Me.Guna2CircleButton3.Image = CType(resources.GetObject("Guna2CircleButton3.Image"), System.Drawing.Image)
        Me.Guna2CircleButton3.Location = New System.Drawing.Point(597, 141)
        Me.Guna2CircleButton3.Name = "Guna2CircleButton3"
        Me.Guna2CircleButton3.ShadowDecoration.Mode = Guna.UI2.WinForms.Enums.ShadowMode.Circle
        Me.Guna2CircleButton3.Size = New System.Drawing.Size(38, 37)
        Me.Guna2CircleButton3.TabIndex = 12
        '
        'Guna2CircleButton2
        '
        Me.Guna2CircleButton2.Animated = True
        Me.Guna2CircleButton2.BorderColor = System.Drawing.Color.FromArgb(CType(CType(213, Byte), Integer), CType(CType(218, Byte), Integer), CType(CType(223, Byte), Integer))
        Me.Guna2CircleButton2.DisabledState.BorderColor = System.Drawing.Color.DarkGray
        Me.Guna2CircleButton2.DisabledState.CustomBorderColor = System.Drawing.Color.DarkGray
        Me.Guna2CircleButton2.DisabledState.FillColor = System.Drawing.Color.FromArgb(CType(CType(169, Byte), Integer), CType(CType(169, Byte), Integer), CType(CType(169, Byte), Integer))
        Me.Guna2CircleButton2.DisabledState.ForeColor = System.Drawing.Color.FromArgb(CType(CType(141, Byte), Integer), CType(CType(141, Byte), Integer), CType(CType(141, Byte), Integer))
        Me.Guna2CircleButton2.FillColor = System.Drawing.Color.FromArgb(CType(CType(229, Byte), Integer), CType(CType(229, Byte), Integer), CType(CType(229, Byte), Integer))
        Me.Guna2CircleButton2.Font = New System.Drawing.Font("Segoe UI", 9.0!)
        Me.Guna2CircleButton2.ForeColor = System.Drawing.Color.FromArgb(CType(CType(229, Byte), Integer), CType(CType(229, Byte), Integer), CType(CType(229, Byte), Integer))
        Me.Guna2CircleButton2.Image = CType(resources.GetObject("Guna2CircleButton2.Image"), System.Drawing.Image)
        Me.Guna2CircleButton2.Location = New System.Drawing.Point(538, 141)
        Me.Guna2CircleButton2.Name = "Guna2CircleButton2"
        Me.Guna2CircleButton2.ShadowDecoration.Mode = Guna.UI2.WinForms.Enums.ShadowMode.Circle
        Me.Guna2CircleButton2.Size = New System.Drawing.Size(38, 37)
        Me.Guna2CircleButton2.TabIndex = 13
        '
        'Guna2CircleButton1
        '
        Me.Guna2CircleButton1.Animated = True
        Me.Guna2CircleButton1.BorderColor = System.Drawing.Color.FromArgb(CType(CType(213, Byte), Integer), CType(CType(218, Byte), Integer), CType(CType(223, Byte), Integer))
        Me.Guna2CircleButton1.DisabledState.BorderColor = System.Drawing.Color.DarkGray
        Me.Guna2CircleButton1.DisabledState.CustomBorderColor = System.Drawing.Color.DarkGray
        Me.Guna2CircleButton1.DisabledState.FillColor = System.Drawing.Color.FromArgb(CType(CType(169, Byte), Integer), CType(CType(169, Byte), Integer), CType(CType(169, Byte), Integer))
        Me.Guna2CircleButton1.DisabledState.ForeColor = System.Drawing.Color.FromArgb(CType(CType(141, Byte), Integer), CType(CType(141, Byte), Integer), CType(CType(141, Byte), Integer))
        Me.Guna2CircleButton1.FillColor = System.Drawing.Color.FromArgb(CType(CType(229, Byte), Integer), CType(CType(229, Byte), Integer), CType(CType(229, Byte), Integer))
        Me.Guna2CircleButton1.Font = New System.Drawing.Font("Segoe UI", 9.0!)
        Me.Guna2CircleButton1.ForeColor = System.Drawing.Color.FromArgb(CType(CType(229, Byte), Integer), CType(CType(229, Byte), Integer), CType(CType(229, Byte), Integer))
        Me.Guna2CircleButton1.Image = CType(resources.GetObject("Guna2CircleButton1.Image"), System.Drawing.Image)
        Me.Guna2CircleButton1.Location = New System.Drawing.Point(481, 141)
        Me.Guna2CircleButton1.Name = "Guna2CircleButton1"
        Me.Guna2CircleButton1.ShadowDecoration.Mode = Guna.UI2.WinForms.Enums.ShadowMode.Circle
        Me.Guna2CircleButton1.Size = New System.Drawing.Size(38, 37)
        Me.Guna2CircleButton1.TabIndex = 14
        '
        'Guna2HtmlLabel1
        '
        Me.Guna2HtmlLabel1.BackColor = System.Drawing.Color.Transparent
        Me.Guna2HtmlLabel1.Font = New System.Drawing.Font("Comic Sans MS", 9.75!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.Guna2HtmlLabel1.ForeColor = System.Drawing.SystemColors.ButtonShadow
        Me.Guna2HtmlLabel1.Location = New System.Drawing.Point(606, 193)
        Me.Guna2HtmlLabel1.Name = "Guna2HtmlLabel1"
        Me.Guna2HtmlLabel1.Size = New System.Drawing.Size(19, 21)
        Me.Guna2HtmlLabel1.TabIndex = 15
        Me.Guna2HtmlLabel1.Text = "Or"
        '
        'Guna2PictureBox1
        '
        Me.Guna2PictureBox1.BackColor = System.Drawing.Color.Transparent
        Me.Guna2PictureBox1.FillColor = System.Drawing.Color.Transparent
        Me.Guna2PictureBox1.Image = Global.CoffeeShop.My.Resources.Resources.bgcoflogo
        Me.Guna2PictureBox1.ImageRotate = 0!
        Me.Guna2PictureBox1.Location = New System.Drawing.Point(-37, 193)
        Me.Guna2PictureBox1.Name = "Guna2PictureBox1"
        Me.Guna2PictureBox1.Size = New System.Drawing.Size(449, 337)
        Me.Guna2PictureBox1.TabIndex = 17
        Me.Guna2PictureBox1.TabStop = False
        Me.Guna2PictureBox1.UseTransparentBackground = True
        '
        'Form1login
        '
        Me.AutoScaleDimensions = New System.Drawing.SizeF(6.0!, 13.0!)
        Me.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font
        Me.ClientSize = New System.Drawing.Size(840, 550)
        Me.Controls.Add(Me.Guna2PictureBox1)
        Me.Controls.Add(Me.bgPictureBox)
        Me.Controls.Add(Me.Guna2HtmlLabel1)
        Me.Controls.Add(Me.Guna2CircleButton5)
        Me.Controls.Add(Me.Guna2CircleButton4)
        Me.Controls.Add(Me.Guna2CircleButton3)
        Me.Controls.Add(Me.Guna2CircleButton2)
        Me.Controls.Add(Me.Guna2CircleButton1)
        Me.Controls.Add(Me.Guna2GradientButton1)
        Me.Controls.Add(Me.loginPanel1)
        Me.Controls.Add(Me.username)
        Me.Controls.Add(Me.password)
        Me.Controls.Add(Me.forgotpassButton)
        Me.Controls.Add(Me.rememberMeToggleSwitch)
        Me.Controls.Add(Me.registerButton)
        Me.Controls.Add(Me.rememLabel)
        Me.Controls.Add(Me.helloLabel)
        Me.Controls.Add(Me.signinLabel)
        Me.ForeColor = System.Drawing.SystemColors.AppWorkspace
        Me.FormBorderStyle = System.Windows.Forms.FormBorderStyle.None
        Me.Icon = CType(resources.GetObject("$this.Icon"), System.Drawing.Icon)
        Me.Name = "Form1login"
        Me.Text = "LoginForm1"
        CType(Me.bgPictureBox, System.ComponentModel.ISupportInitialize).EndInit()
        Me.loginPanel1.ResumeLayout(False)
        CType(Me.Guna2PictureBox1, System.ComponentModel.ISupportInitialize).EndInit()
        Me.ResumeLayout(False)
        Me.PerformLayout()

    End Sub

    Friend WithEvents Guna2Elipse1 As Guna.UI2.WinForms.Guna2Elipse
    Friend WithEvents Guna2AnimateWindow1 As Guna.UI2.WinForms.Guna2AnimateWindow
    Friend WithEvents panelDragControl1 As Guna.UI2.WinForms.Guna2DragControl
    Friend WithEvents Guna2ShadowForm1 As Guna.UI2.WinForms.Guna2ShadowForm
    Friend WithEvents bgDragControl1 As Guna.UI2.WinForms.Guna2DragControl
    Friend WithEvents formDragControl1 As Guna.UI2.WinForms.Guna2DragControl
    Friend WithEvents username As Guna.UI2.WinForms.Guna2TextBox
    Friend WithEvents password As Guna.UI2.WinForms.Guna2TextBox
    Friend WithEvents forgotpassButton As Guna.UI2.WinForms.Guna2Button
    Friend WithEvents rememberMeToggleSwitch As Guna.UI2.WinForms.Guna2ToggleSwitch
    Friend WithEvents registerButton As Guna.UI2.WinForms.Guna2Button
    Friend WithEvents rememLabel As Guna.UI2.WinForms.Guna2HtmlLabel
    Friend WithEvents helloLabel As Guna.UI2.WinForms.Guna2HtmlLabel
    Friend WithEvents signinLabel As Guna.UI2.WinForms.Guna2HtmlLabel
    Friend WithEvents loginPanel1 As Guna.UI2.WinForms.Guna2Panel
    Friend WithEvents Guna2ControlBox3 As Guna.UI2.WinForms.Guna2ControlBox
    Friend WithEvents Guna2ControlBox4 As Guna.UI2.WinForms.Guna2ControlBox
    Friend WithEvents Guna2GradientButton1 As Guna.UI2.WinForms.Guna2GradientButton
    Friend WithEvents Guna2CircleButton5 As Guna.UI2.WinForms.Guna2CircleButton
    Friend WithEvents Guna2CircleButton4 As Guna.UI2.WinForms.Guna2CircleButton
    Friend WithEvents Guna2CircleButton3 As Guna.UI2.WinForms.Guna2CircleButton
    Friend WithEvents Guna2CircleButton2 As Guna.UI2.WinForms.Guna2CircleButton
    Friend WithEvents Guna2CircleButton1 As Guna.UI2.WinForms.Guna2CircleButton
    Friend WithEvents Guna2HtmlLabel1 As Guna.UI2.WinForms.Guna2HtmlLabel
    Friend WithEvents bgPictureBox As Guna.UI2.WinForms.Guna2PictureBox
    Friend WithEvents Guna2PictureBox1 As Guna.UI2.WinForms.Guna2PictureBox
End Class
