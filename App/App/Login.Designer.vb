<Global.Microsoft.VisualBasic.CompilerServices.DesignerGenerated()>
Partial Class Login
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
        Dim resources As System.ComponentModel.ComponentResourceManager = New System.ComponentModel.ComponentResourceManager(GetType(Login))
        Me.SplitContainer1 = New System.Windows.Forms.SplitContainer()
        Me.lastname = New System.Windows.Forms.Label()
        Me.Label8 = New System.Windows.Forms.Label()
        Me.Label7 = New System.Windows.Forms.Label()
        Me.Label6 = New System.Windows.Forms.Label()
        Me.firstname = New System.Windows.Forms.Label()
        Me.jbTitle = New System.Windows.Forms.Label()
        Me.comDecrypt = New System.Windows.Forms.Label()
        Me.comsalt = New System.Windows.Forms.Label()
        Me.comUsername = New System.Windows.Forms.Label()
        Me.comId = New System.Windows.Forms.Label()
        Me.comUserpassword = New System.Windows.Forms.Label()
        Me.Guna2PictureBox1 = New Guna.UI2.WinForms.Guna2PictureBox()
        Me.btnExit = New Guna.UI2.WinForms.Guna2Button()
        Me.Label1 = New System.Windows.Forms.Label()
        Me.Guna2Panel2 = New Guna.UI2.WinForms.Guna2Panel()
        Me.Guna2Panel3 = New Guna.UI2.WinForms.Guna2Panel()
        Me.btnLogin = New Guna.UI2.WinForms.Guna2Button()
        Me.txtPassword = New Guna.UI2.WinForms.Guna2TextBox()
        Me.Label10 = New System.Windows.Forms.Label()
        Me.Guna2Panel1 = New Guna.UI2.WinForms.Guna2Panel()
        Me.txtUsername = New Guna.UI2.WinForms.Guna2TextBox()
        Me.Label5 = New System.Windows.Forms.Label()
        Me.Label3 = New System.Windows.Forms.Label()
        Me.Timer1 = New System.Windows.Forms.Timer(Me.components)
        CType(Me.SplitContainer1, System.ComponentModel.ISupportInitialize).BeginInit()
        Me.SplitContainer1.Panel1.SuspendLayout()
        Me.SplitContainer1.Panel2.SuspendLayout()
        Me.SplitContainer1.SuspendLayout()
        CType(Me.Guna2PictureBox1, System.ComponentModel.ISupportInitialize).BeginInit()
        Me.Guna2Panel2.SuspendLayout()
        Me.Guna2Panel3.SuspendLayout()
        Me.Guna2Panel1.SuspendLayout()
        Me.SuspendLayout()
        '
        'SplitContainer1
        '
        Me.SplitContainer1.Location = New System.Drawing.Point(0, -1)
        Me.SplitContainer1.Name = "SplitContainer1"
        '
        'SplitContainer1.Panel1
        '
        Me.SplitContainer1.Panel1.BackColor = System.Drawing.Color.FromArgb(CType(CType(23, Byte), Integer), CType(CType(18, Byte), Integer), CType(CType(59, Byte), Integer))
        Me.SplitContainer1.Panel1.Controls.Add(Me.lastname)
        Me.SplitContainer1.Panel1.Controls.Add(Me.Label8)
        Me.SplitContainer1.Panel1.Controls.Add(Me.Label7)
        Me.SplitContainer1.Panel1.Controls.Add(Me.Label6)
        Me.SplitContainer1.Panel1.Controls.Add(Me.firstname)
        Me.SplitContainer1.Panel1.Controls.Add(Me.jbTitle)
        Me.SplitContainer1.Panel1.Controls.Add(Me.comDecrypt)
        Me.SplitContainer1.Panel1.Controls.Add(Me.comsalt)
        Me.SplitContainer1.Panel1.Controls.Add(Me.comUsername)
        Me.SplitContainer1.Panel1.Controls.Add(Me.comId)
        Me.SplitContainer1.Panel1.Controls.Add(Me.comUserpassword)
        Me.SplitContainer1.Panel1.Controls.Add(Me.Guna2PictureBox1)
        '
        'SplitContainer1.Panel2
        '
        Me.SplitContainer1.Panel2.BackColor = System.Drawing.Color.FromArgb(CType(CType(23, Byte), Integer), CType(CType(18, Byte), Integer), CType(CType(59, Byte), Integer))
        Me.SplitContainer1.Panel2.Controls.Add(Me.btnExit)
        Me.SplitContainer1.Panel2.Controls.Add(Me.Label1)
        Me.SplitContainer1.Panel2.Controls.Add(Me.Guna2Panel2)
        Me.SplitContainer1.Panel2.Controls.Add(Me.Guna2Panel1)
        Me.SplitContainer1.Size = New System.Drawing.Size(1169, 637)
        Me.SplitContainer1.SplitterDistance = 388
        Me.SplitContainer1.TabIndex = 1
        '
        'lastname
        '
        Me.lastname.AutoSize = True
        Me.lastname.Font = New System.Drawing.Font("Microsoft Sans Serif", 12.0!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.lastname.ForeColor = System.Drawing.Color.White
        Me.lastname.Location = New System.Drawing.Point(160, 509)
        Me.lastname.Name = "lastname"
        Me.lastname.Size = New System.Drawing.Size(82, 20)
        Me.lastname.TabIndex = 17
        Me.lastname.Text = "lastname"
        '
        'Label8
        '
        Me.Label8.Dock = System.Windows.Forms.DockStyle.Top
        Me.Label8.Font = New System.Drawing.Font("Microsoft Sans Serif", 12.0!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.Label8.ForeColor = System.Drawing.Color.White
        Me.Label8.Location = New System.Drawing.Point(0, 313)
        Me.Label8.Name = "Label8"
        Me.Label8.Padding = New System.Windows.Forms.Padding(15, 5, 5, 5)
        Me.Label8.Size = New System.Drawing.Size(388, 48)
        Me.Label8.TabIndex = 8
        Me.Label8.Text = "Ticketing Booth Login Interface.."
        Me.Label8.TextAlign = System.Drawing.ContentAlignment.MiddleLeft
        '
        'Label7
        '
        Me.Label7.Dock = System.Windows.Forms.DockStyle.Top
        Me.Label7.Font = New System.Drawing.Font("Microsoft Sans Serif", 12.0!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.Label7.ForeColor = System.Drawing.Color.White
        Me.Label7.Location = New System.Drawing.Point(0, 267)
        Me.Label7.Name = "Label7"
        Me.Label7.Padding = New System.Windows.Forms.Padding(15, 5, 5, 5)
        Me.Label7.Size = New System.Drawing.Size(388, 46)
        Me.Label7.TabIndex = 7
        Me.Label7.Text = "This is Tabango Water Park"
        Me.Label7.TextAlign = System.Drawing.ContentAlignment.MiddleLeft
        '
        'Label6
        '
        Me.Label6.Dock = System.Windows.Forms.DockStyle.Top
        Me.Label6.Font = New System.Drawing.Font("Microsoft Sans Serif", 18.0!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.Label6.ForeColor = System.Drawing.Color.White
        Me.Label6.Location = New System.Drawing.Point(0, 182)
        Me.Label6.Margin = New System.Windows.Forms.Padding(5)
        Me.Label6.Name = "Label6"
        Me.Label6.Padding = New System.Windows.Forms.Padding(15, 5, 5, 5)
        Me.Label6.Size = New System.Drawing.Size(388, 85)
        Me.Label6.TabIndex = 6
        Me.Label6.Text = "WELCOME !!"
        Me.Label6.TextAlign = System.Drawing.ContentAlignment.MiddleLeft
        '
        'firstname
        '
        Me.firstname.AutoSize = True
        Me.firstname.Font = New System.Drawing.Font("Microsoft Sans Serif", 12.0!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.firstname.ForeColor = System.Drawing.Color.White
        Me.firstname.Location = New System.Drawing.Point(160, 469)
        Me.firstname.Name = "firstname"
        Me.firstname.Size = New System.Drawing.Size(53, 20)
        Me.firstname.TabIndex = 16
        Me.firstname.Text = "name"
        '
        'jbTitle
        '
        Me.jbTitle.AutoSize = True
        Me.jbTitle.Font = New System.Drawing.Font("Microsoft Sans Serif", 12.0!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.jbTitle.ForeColor = System.Drawing.Color.White
        Me.jbTitle.Location = New System.Drawing.Point(19, 573)
        Me.jbTitle.Name = "jbTitle"
        Me.jbTitle.Size = New System.Drawing.Size(53, 20)
        Me.jbTitle.TabIndex = 15
        Me.jbTitle.Text = "jbtitle"
        '
        'comDecrypt
        '
        Me.comDecrypt.AutoSize = True
        Me.comDecrypt.Font = New System.Drawing.Font("Microsoft Sans Serif", 12.0!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.comDecrypt.ForeColor = System.Drawing.Color.White
        Me.comDecrypt.Location = New System.Drawing.Point(19, 553)
        Me.comDecrypt.Name = "comDecrypt"
        Me.comDecrypt.Size = New System.Drawing.Size(68, 20)
        Me.comDecrypt.TabIndex = 14
        Me.comDecrypt.Text = "decrypt"
        '
        'comsalt
        '
        Me.comsalt.AutoSize = True
        Me.comsalt.Font = New System.Drawing.Font("Microsoft Sans Serif", 12.0!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.comsalt.ForeColor = System.Drawing.Color.White
        Me.comsalt.Location = New System.Drawing.Point(19, 529)
        Me.comsalt.Name = "comsalt"
        Me.comsalt.Size = New System.Drawing.Size(38, 20)
        Me.comsalt.TabIndex = 13
        Me.comsalt.Text = "salt"
        '
        'comUsername
        '
        Me.comUsername.AutoSize = True
        Me.comUsername.Font = New System.Drawing.Font("Microsoft Sans Serif", 12.0!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.comUsername.ForeColor = System.Drawing.Color.White
        Me.comUsername.Location = New System.Drawing.Point(19, 489)
        Me.comUsername.Name = "comUsername"
        Me.comUsername.Size = New System.Drawing.Size(88, 20)
        Me.comUsername.TabIndex = 12
        Me.comUsername.Text = "username"
        '
        'comId
        '
        Me.comId.AutoSize = True
        Me.comId.Font = New System.Drawing.Font("Microsoft Sans Serif", 12.0!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.comId.ForeColor = System.Drawing.Color.White
        Me.comId.Location = New System.Drawing.Point(19, 469)
        Me.comId.Name = "comId"
        Me.comId.Size = New System.Drawing.Size(23, 20)
        Me.comId.TabIndex = 11
        Me.comId.Text = "id"
        '
        'comUserpassword
        '
        Me.comUserpassword.AutoSize = True
        Me.comUserpassword.Font = New System.Drawing.Font("Microsoft Sans Serif", 12.0!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.comUserpassword.ForeColor = System.Drawing.Color.White
        Me.comUserpassword.Location = New System.Drawing.Point(19, 509)
        Me.comUserpassword.Name = "comUserpassword"
        Me.comUserpassword.Size = New System.Drawing.Size(85, 20)
        Me.comUserpassword.TabIndex = 10
        Me.comUserpassword.Text = "password"
        '
        'Guna2PictureBox1
        '
        Me.Guna2PictureBox1.Dock = System.Windows.Forms.DockStyle.Top
        Me.Guna2PictureBox1.Image = CType(resources.GetObject("Guna2PictureBox1.Image"), System.Drawing.Image)
        Me.Guna2PictureBox1.Location = New System.Drawing.Point(0, 0)
        Me.Guna2PictureBox1.Margin = New System.Windows.Forms.Padding(0)
        Me.Guna2PictureBox1.Name = "Guna2PictureBox1"
        Me.Guna2PictureBox1.Padding = New System.Windows.Forms.Padding(10)
        Me.Guna2PictureBox1.ShadowDecoration.BorderRadius = 0
        Me.Guna2PictureBox1.ShadowDecoration.Depth = 0
        Me.Guna2PictureBox1.ShadowDecoration.Parent = Me.Guna2PictureBox1
        Me.Guna2PictureBox1.Size = New System.Drawing.Size(388, 182)
        Me.Guna2PictureBox1.SizeMode = System.Windows.Forms.PictureBoxSizeMode.StretchImage
        Me.Guna2PictureBox1.TabIndex = 9
        Me.Guna2PictureBox1.TabStop = False
        '
        'btnExit
        '
        Me.btnExit.BackColor = System.Drawing.Color.Transparent
        Me.btnExit.BorderColor = System.Drawing.Color.White
        Me.btnExit.BorderRadius = 5
        Me.btnExit.BorderThickness = 2
        Me.btnExit.CheckedState.Parent = Me.btnExit
        Me.btnExit.Cursor = System.Windows.Forms.Cursors.Hand
        Me.btnExit.CustomImages.Parent = Me.btnExit
        Me.btnExit.FillColor = System.Drawing.Color.Transparent
        Me.btnExit.Font = New System.Drawing.Font("Segoe UI Semibold", 15.75!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.btnExit.ForeColor = System.Drawing.Color.White
        Me.btnExit.HoverState.BorderColor = System.Drawing.Color.Red
        Me.btnExit.HoverState.FillColor = System.Drawing.Color.LightSkyBlue
        Me.btnExit.HoverState.ForeColor = System.Drawing.Color.Red
        Me.btnExit.HoverState.Parent = Me.btnExit
        Me.btnExit.Location = New System.Drawing.Point(736, 0)
        Me.btnExit.Name = "btnExit"
        Me.btnExit.ShadowDecoration.Parent = Me.btnExit
        Me.btnExit.Size = New System.Drawing.Size(40, 35)
        Me.btnExit.TabIndex = 10
        Me.btnExit.Text = "X"
        '
        'Label1
        '
        Me.Label1.AutoSize = True
        Me.Label1.Font = New System.Drawing.Font("Microsoft Sans Serif", 30.0!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.Label1.ForeColor = System.Drawing.Color.White
        Me.Label1.Location = New System.Drawing.Point(113, 46)
        Me.Label1.Name = "Label1"
        Me.Label1.Size = New System.Drawing.Size(533, 46)
        Me.Label1.TabIndex = 0
        Me.Label1.Text = "Water Park Ticketing Booth"
        '
        'Guna2Panel2
        '
        Me.Guna2Panel2.BackColor = System.Drawing.Color.FromArgb(CType(CType(42, Byte), Integer), CType(CType(41, Byte), Integer), CType(CType(83, Byte), Integer))
        Me.Guna2Panel2.Controls.Add(Me.Guna2Panel3)
        Me.Guna2Panel2.Location = New System.Drawing.Point(-2, 340)
        Me.Guna2Panel2.Name = "Guna2Panel2"
        Me.Guna2Panel2.ShadowDecoration.Parent = Me.Guna2Panel2
        Me.Guna2Panel2.Size = New System.Drawing.Size(778, 299)
        Me.Guna2Panel2.TabIndex = 0
        '
        'Guna2Panel3
        '
        Me.Guna2Panel3.BackColor = System.Drawing.Color.Transparent
        Me.Guna2Panel3.BorderColor = System.Drawing.Color.Black
        Me.Guna2Panel3.BorderRadius = 30
        Me.Guna2Panel3.Controls.Add(Me.btnLogin)
        Me.Guna2Panel3.Controls.Add(Me.txtPassword)
        Me.Guna2Panel3.Controls.Add(Me.Label10)
        Me.Guna2Panel3.FillColor = System.Drawing.Color.FromArgb(CType(CType(23, Byte), Integer), CType(CType(18, Byte), Integer), CType(CType(59, Byte), Integer))
        Me.Guna2Panel3.Location = New System.Drawing.Point(141, -63)
        Me.Guna2Panel3.Name = "Guna2Panel3"
        Me.Guna2Panel3.ShadowDecoration.BorderRadius = 30
        Me.Guna2Panel3.ShadowDecoration.Enabled = True
        Me.Guna2Panel3.ShadowDecoration.Parent = Me.Guna2Panel3
        Me.Guna2Panel3.ShadowDecoration.Shadow = New System.Windows.Forms.Padding(2)
        Me.Guna2Panel3.Size = New System.Drawing.Size(476, 296)
        Me.Guna2Panel3.TabIndex = 17
        '
        'btnLogin
        '
        Me.btnLogin.Animated = True
        Me.btnLogin.BackColor = System.Drawing.Color.Transparent
        Me.btnLogin.BorderRadius = 7
        Me.btnLogin.CheckedState.Parent = Me.btnLogin
        Me.btnLogin.Cursor = System.Windows.Forms.Cursors.Hand
        Me.btnLogin.CustomImages.Parent = Me.btnLogin
        Me.btnLogin.FillColor = System.Drawing.Color.DodgerBlue
        Me.btnLogin.Font = New System.Drawing.Font("Century Gothic", 14.25!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.btnLogin.ForeColor = System.Drawing.Color.White
        Me.btnLogin.HoverState.Parent = Me.btnLogin
        Me.btnLogin.Location = New System.Drawing.Point(57, 167)
        Me.btnLogin.Name = "btnLogin"
        Me.btnLogin.ShadowDecoration.BorderRadius = 10
        Me.btnLogin.ShadowDecoration.Enabled = True
        Me.btnLogin.ShadowDecoration.Parent = Me.btnLogin
        Me.btnLogin.ShadowDecoration.Shadow = New System.Windows.Forms.Padding(3)
        Me.btnLogin.Size = New System.Drawing.Size(342, 45)
        Me.btnLogin.TabIndex = 16
        Me.btnLogin.Text = "Login"
        '
        'txtPassword
        '
        Me.txtPassword.Animated = True
        Me.txtPassword.BackColor = System.Drawing.Color.Transparent
        Me.txtPassword.BorderColor = System.Drawing.Color.FromArgb(CType(CType(23, Byte), Integer), CType(CType(18, Byte), Integer), CType(CType(59, Byte), Integer))
        Me.txtPassword.BorderRadius = 7
        Me.txtPassword.BorderThickness = 0
        Me.txtPassword.Cursor = System.Windows.Forms.Cursors.Hand
        Me.txtPassword.DefaultText = ""
        Me.txtPassword.DisabledState.BorderColor = System.Drawing.Color.FromArgb(CType(CType(208, Byte), Integer), CType(CType(208, Byte), Integer), CType(CType(208, Byte), Integer))
        Me.txtPassword.DisabledState.FillColor = System.Drawing.Color.FromArgb(CType(CType(226, Byte), Integer), CType(CType(226, Byte), Integer), CType(CType(226, Byte), Integer))
        Me.txtPassword.DisabledState.ForeColor = System.Drawing.Color.FromArgb(CType(CType(138, Byte), Integer), CType(CType(138, Byte), Integer), CType(CType(138, Byte), Integer))
        Me.txtPassword.DisabledState.Parent = Me.txtPassword
        Me.txtPassword.DisabledState.PlaceholderForeColor = System.Drawing.Color.FromArgb(CType(CType(138, Byte), Integer), CType(CType(138, Byte), Integer), CType(CType(138, Byte), Integer))
        Me.txtPassword.FillColor = System.Drawing.Color.FromArgb(CType(CType(42, Byte), Integer), CType(CType(41, Byte), Integer), CType(CType(83, Byte), Integer))
        Me.txtPassword.FocusedState.BorderColor = System.Drawing.Color.FromArgb(CType(CType(94, Byte), Integer), CType(CType(148, Byte), Integer), CType(CType(255, Byte), Integer))
        Me.txtPassword.FocusedState.Parent = Me.txtPassword
        Me.txtPassword.Font = New System.Drawing.Font("Century Gothic", 14.25!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.txtPassword.ForeColor = System.Drawing.Color.White
        Me.txtPassword.HoverState.BorderColor = System.Drawing.Color.FromArgb(CType(CType(94, Byte), Integer), CType(CType(148, Byte), Integer), CType(CType(255, Byte), Integer))
        Me.txtPassword.HoverState.Parent = Me.txtPassword
        Me.txtPassword.IconLeft = CType(resources.GetObject("txtPassword.IconLeft"), System.Drawing.Image)
        Me.txtPassword.IconRight = CType(resources.GetObject("txtPassword.IconRight"), System.Drawing.Image)
        Me.txtPassword.IconRightCursor = System.Windows.Forms.Cursors.Hand
        Me.txtPassword.Location = New System.Drawing.Point(57, 91)
        Me.txtPassword.Margin = New System.Windows.Forms.Padding(3, 4, 3, 4)
        Me.txtPassword.Name = "txtPassword"
        Me.txtPassword.PasswordChar = Global.Microsoft.VisualBasic.ChrW(9679)
        Me.txtPassword.PlaceholderForeColor = System.Drawing.Color.Gray
        Me.txtPassword.PlaceholderText = ""
        Me.txtPassword.SelectedText = ""
        Me.txtPassword.ShadowDecoration.BorderRadius = 10
        Me.txtPassword.ShadowDecoration.Enabled = True
        Me.txtPassword.ShadowDecoration.Parent = Me.txtPassword
        Me.txtPassword.ShadowDecoration.Shadow = New System.Windows.Forms.Padding(2)
        Me.txtPassword.Size = New System.Drawing.Size(342, 43)
        Me.txtPassword.TabIndex = 15
        Me.txtPassword.TextAlign = System.Windows.Forms.HorizontalAlignment.Center
        '
        'Label10
        '
        Me.Label10.AutoSize = True
        Me.Label10.Font = New System.Drawing.Font("Century Gothic", 15.75!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.Label10.ForeColor = System.Drawing.Color.White
        Me.Label10.Location = New System.Drawing.Point(53, 60)
        Me.Label10.Name = "Label10"
        Me.Label10.Size = New System.Drawing.Size(104, 24)
        Me.Label10.TabIndex = 12
        Me.Label10.Text = "Password"
        '
        'Guna2Panel1
        '
        Me.Guna2Panel1.BackColor = System.Drawing.Color.Transparent
        Me.Guna2Panel1.BorderColor = System.Drawing.Color.Black
        Me.Guna2Panel1.BorderRadius = 30
        Me.Guna2Panel1.Controls.Add(Me.txtUsername)
        Me.Guna2Panel1.Controls.Add(Me.Label5)
        Me.Guna2Panel1.Controls.Add(Me.Label3)
        Me.Guna2Panel1.FillColor = System.Drawing.Color.FromArgb(CType(CType(23, Byte), Integer), CType(CType(18, Byte), Integer), CType(CType(59, Byte), Integer))
        Me.Guna2Panel1.Location = New System.Drawing.Point(139, 140)
        Me.Guna2Panel1.Name = "Guna2Panel1"
        Me.Guna2Panel1.ShadowDecoration.BorderRadius = 30
        Me.Guna2Panel1.ShadowDecoration.Enabled = True
        Me.Guna2Panel1.ShadowDecoration.Parent = Me.Guna2Panel1
        Me.Guna2Panel1.ShadowDecoration.Shadow = New System.Windows.Forms.Padding(3)
        Me.Guna2Panel1.Size = New System.Drawing.Size(476, 257)
        Me.Guna2Panel1.TabIndex = 9
        '
        'txtUsername
        '
        Me.txtUsername.Animated = True
        Me.txtUsername.BackColor = System.Drawing.Color.Transparent
        Me.txtUsername.BorderColor = System.Drawing.Color.FromArgb(CType(CType(23, Byte), Integer), CType(CType(18, Byte), Integer), CType(CType(59, Byte), Integer))
        Me.txtUsername.BorderRadius = 7
        Me.txtUsername.BorderThickness = 0
        Me.txtUsername.Cursor = System.Windows.Forms.Cursors.Hand
        Me.txtUsername.DefaultText = ""
        Me.txtUsername.DisabledState.BorderColor = System.Drawing.Color.FromArgb(CType(CType(208, Byte), Integer), CType(CType(208, Byte), Integer), CType(CType(208, Byte), Integer))
        Me.txtUsername.DisabledState.FillColor = System.Drawing.Color.FromArgb(CType(CType(226, Byte), Integer), CType(CType(226, Byte), Integer), CType(CType(226, Byte), Integer))
        Me.txtUsername.DisabledState.ForeColor = System.Drawing.Color.FromArgb(CType(CType(138, Byte), Integer), CType(CType(138, Byte), Integer), CType(CType(138, Byte), Integer))
        Me.txtUsername.DisabledState.Parent = Me.txtUsername
        Me.txtUsername.DisabledState.PlaceholderForeColor = System.Drawing.Color.FromArgb(CType(CType(138, Byte), Integer), CType(CType(138, Byte), Integer), CType(CType(138, Byte), Integer))
        Me.txtUsername.FillColor = System.Drawing.Color.FromArgb(CType(CType(42, Byte), Integer), CType(CType(41, Byte), Integer), CType(CType(83, Byte), Integer))
        Me.txtUsername.FocusedState.BorderColor = System.Drawing.Color.FromArgb(CType(CType(94, Byte), Integer), CType(CType(148, Byte), Integer), CType(CType(255, Byte), Integer))
        Me.txtUsername.FocusedState.Parent = Me.txtUsername
        Me.txtUsername.Font = New System.Drawing.Font("Century Gothic", 14.25!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.txtUsername.ForeColor = System.Drawing.Color.White
        Me.txtUsername.HoverState.BorderColor = System.Drawing.Color.FromArgb(CType(CType(94, Byte), Integer), CType(CType(148, Byte), Integer), CType(CType(255, Byte), Integer))
        Me.txtUsername.HoverState.Parent = Me.txtUsername
        Me.txtUsername.IconLeft = CType(resources.GetObject("txtUsername.IconLeft"), System.Drawing.Image)
        Me.txtUsername.Location = New System.Drawing.Point(57, 111)
        Me.txtUsername.Margin = New System.Windows.Forms.Padding(3, 4, 3, 4)
        Me.txtUsername.Name = "txtUsername"
        Me.txtUsername.PasswordChar = Global.Microsoft.VisualBasic.ChrW(0)
        Me.txtUsername.PlaceholderForeColor = System.Drawing.Color.Black
        Me.txtUsername.PlaceholderText = ""
        Me.txtUsername.SelectedText = ""
        Me.txtUsername.ShadowDecoration.BorderRadius = 10
        Me.txtUsername.ShadowDecoration.Enabled = True
        Me.txtUsername.ShadowDecoration.Parent = Me.txtUsername
        Me.txtUsername.ShadowDecoration.Shadow = New System.Windows.Forms.Padding(2)
        Me.txtUsername.Size = New System.Drawing.Size(342, 44)
        Me.txtUsername.TabIndex = 14
        Me.txtUsername.TextAlign = System.Windows.Forms.HorizontalAlignment.Center
        '
        'Label5
        '
        Me.Label5.AutoSize = True
        Me.Label5.Font = New System.Drawing.Font("Century Gothic", 15.75!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.Label5.ForeColor = System.Drawing.Color.White
        Me.Label5.Location = New System.Drawing.Point(99, 17)
        Me.Label5.Name = "Label5"
        Me.Label5.Size = New System.Drawing.Size(251, 25)
        Me.Label5.TabIndex = 13
        Me.Label5.Text = "Login as admin or seller"
        '
        'Label3
        '
        Me.Label3.AutoSize = True
        Me.Label3.Font = New System.Drawing.Font("Century Gothic", 15.75!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.Label3.ForeColor = System.Drawing.Color.White
        Me.Label3.Location = New System.Drawing.Point(53, 79)
        Me.Label3.Name = "Label3"
        Me.Label3.Size = New System.Drawing.Size(202, 24)
        Me.Label3.TabIndex = 11
        Me.Label3.Text = "Username or email"
        '
        'Login
        '
        Me.AcceptButton = Me.btnLogin
        Me.AutoScaleDimensions = New System.Drawing.SizeF(6.0!, 13.0!)
        Me.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font
        Me.BackColor = System.Drawing.Color.White
        Me.ClientSize = New System.Drawing.Size(1169, 636)
        Me.Controls.Add(Me.SplitContainer1)
        Me.FormBorderStyle = System.Windows.Forms.FormBorderStyle.None
        Me.MaximizeBox = False
        Me.Name = "Login"
        Me.StartPosition = System.Windows.Forms.FormStartPosition.CenterScreen
        Me.Text = "7"
        Me.SplitContainer1.Panel1.ResumeLayout(False)
        Me.SplitContainer1.Panel1.PerformLayout()
        Me.SplitContainer1.Panel2.ResumeLayout(False)
        Me.SplitContainer1.Panel2.PerformLayout()
        CType(Me.SplitContainer1, System.ComponentModel.ISupportInitialize).EndInit()
        Me.SplitContainer1.ResumeLayout(False)
        CType(Me.Guna2PictureBox1, System.ComponentModel.ISupportInitialize).EndInit()
        Me.Guna2Panel2.ResumeLayout(False)
        Me.Guna2Panel3.ResumeLayout(False)
        Me.Guna2Panel3.PerformLayout()
        Me.Guna2Panel1.ResumeLayout(False)
        Me.Guna2Panel1.PerformLayout()
        Me.ResumeLayout(False)

    End Sub
    Friend WithEvents SplitContainer1 As SplitContainer
    Friend WithEvents Label8 As Label
    Friend WithEvents Label7 As Label
    Friend WithEvents Label6 As Label
    Friend WithEvents Guna2Panel1 As Guna.UI2.WinForms.Guna2Panel
    Friend WithEvents txtUsername As Guna.UI2.WinForms.Guna2TextBox
    Friend WithEvents Label5 As Label
    Friend WithEvents Label3 As Label
    Friend WithEvents Label1 As Label
    Friend WithEvents Guna2Panel2 As Guna.UI2.WinForms.Guna2Panel
    Friend WithEvents Guna2Panel3 As Guna.UI2.WinForms.Guna2Panel
    Friend WithEvents btnLogin As Guna.UI2.WinForms.Guna2Button
    Friend WithEvents txtPassword As Guna.UI2.WinForms.Guna2TextBox
    Friend WithEvents Label10 As Label
    Friend WithEvents Guna2PictureBox1 As Guna.UI2.WinForms.Guna2PictureBox
    Friend WithEvents btnExit As Guna.UI2.WinForms.Guna2Button
    Friend WithEvents comUsername As Label
    Friend WithEvents comId As Label
    Friend WithEvents comUserpassword As Label
    Friend WithEvents comsalt As Label
    Friend WithEvents comDecrypt As Label
    Friend WithEvents jbTitle As Label
    Friend WithEvents Timer1 As Timer
    Friend WithEvents lastname As Label
    Friend WithEvents firstname As Label
End Class
