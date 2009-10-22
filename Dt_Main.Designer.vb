<Global.Microsoft.VisualBasic.CompilerServices.DesignerGenerated()> Partial Class frmMain
#Region "Windows Form Designer generated code "
	<System.Diagnostics.DebuggerNonUserCode()> Public Sub New()
		MyBase.New()
		'This call is required by the Windows Form Designer.
		InitializeComponent()
	End Sub
	'Form overrides dispose to clean up the component list.
	<System.Diagnostics.DebuggerNonUserCode()> Protected Overloads Overrides Sub Dispose(ByVal Disposing As Boolean)
		If Disposing Then
			If Not components Is Nothing Then
				components.Dispose()
			End If
		End If
		MyBase.Dispose(Disposing)
	End Sub
	'Required by the Windows Form Designer
	Private components As System.ComponentModel.IContainer
	Public ToolTip1 As System.Windows.Forms.ToolTip
	Public WithEvents timerMusicLoop As System.Windows.Forms.Timer
	Public WithEvents tmrOnRealTime As System.Windows.Forms.Timer
	Public WithEvents picFaces As System.Windows.Forms.PictureBox
	Public WithEvents picSteps As System.Windows.Forms.PictureBox
	Public WithEvents picDice As System.Windows.Forms.PictureBox
	Public WithEvents tmrDialogBrief As System.Windows.Forms.Timer
	Public WithEvents picDialogBrief As System.Windows.Forms.Panel
	Public WithEvents picMicroMap As System.Windows.Forms.PictureBox
	Public WithEvents picMicroBox As System.Windows.Forms.Panel
	Public WithEvents picJournal As System.Windows.Forms.Panel
	Public WithEvents picContainer As System.Windows.Forms.PictureBox
	Public WithEvents picTalk As System.Windows.Forms.PictureBox
	Public WithEvents picSearch As System.Windows.Forms.PictureBox
	Public WithEvents picBuySell As System.Windows.Forms.PictureBox
	Public WithEvents picMenu As System.Windows.Forms.PictureBox
	Public WithEvents picConvoBottom As System.Windows.Forms.PictureBox
	Public WithEvents picBlack As System.Windows.Forms.PictureBox
	Public WithEvents picInventoryStatus As System.Windows.Forms.PictureBox
	Public WithEvents picItem As System.Windows.Forms.PictureBox
	Public WithEvents picInventory As System.Windows.Forms.PictureBox
	Public WithEvents picConvoEnter As System.Windows.Forms.PictureBox
	Public WithEvents picConvoList As System.Windows.Forms.PictureBox
	Public WithEvents lblHyperLink As System.Windows.Forms.Label
	Public WithEvents picConvo As System.Windows.Forms.Panel
	Public WithEvents picTPic As System.Windows.Forms.PictureBox
	Public WithEvents picCreateName As System.Windows.Forms.PictureBox
    Public WithEvents picTomeNew As System.Windows.Forms.PictureBox
	Public WithEvents picMainMenu As System.Windows.Forms.PictureBox
	Public WithEvents _picCMap_0 As System.Windows.Forms.PictureBox
	Public WithEvents _picCPic_0 As System.Windows.Forms.PictureBox
	Public WithEvents picRuneSet As System.Windows.Forms.PictureBox
	Public WithEvents picFonts As System.Windows.Forms.PictureBox
	Public WithEvents picMisc As System.Windows.Forms.PictureBox
	Public WithEvents picToHit As System.Windows.Forms.PictureBox
	Public WithEvents picInvDrag As System.Windows.Forms.PictureBox
	Public WithEvents picTmp As System.Windows.Forms.PictureBox
	Public WithEvents tmrEncounterName As System.Windows.Forms.Timer
	Public WithEvents tmrMoveParty As System.Windows.Forms.Timer
	Public WithEvents tmrMoveMap As System.Windows.Forms.Timer
	Public WithEvents _picSFXPic_0 As System.Windows.Forms.PictureBox
	Public WithEvents picTSmall As System.Windows.Forms.PictureBox
	Public WithEvents fraMediaPlayerMusic As System.Windows.Forms.Panel
	Public WithEvents picWallPaper As System.Windows.Forms.PictureBox
	Public WithEvents picGrid As System.Windows.Forms.PictureBox
	Public WithEvents fraMediaPlayerAVI As System.Windows.Forms.Panel
	Public WithEvents picMap As System.Windows.Forms.Panel
	Public WithEvents picBox As System.Windows.Forms.Panel
	Public WithEvents picCMap As Microsoft.VisualBasic.Compatibility.VB6.PictureBoxArray
	Public WithEvents picCPic As Microsoft.VisualBasic.Compatibility.VB6.PictureBoxArray
	Public WithEvents picSFXPic As Microsoft.VisualBasic.Compatibility.VB6.PictureBoxArray
	'NOTE: The following procedure is required by the Windows Form Designer
	'It can be modified using the Windows Form Designer.
	'Do not modify it using the code editor.
	<System.Diagnostics.DebuggerStepThrough()> Private Sub InitializeComponent()
		Dim resources As System.Resources.ResourceManager = New System.Resources.ResourceManager(GetType(frmMain))
		Me.components = New System.ComponentModel.Container()
		Me.ToolTip1 = New System.Windows.Forms.ToolTip(components)
		Me.timerMusicLoop = New System.Windows.Forms.Timer(components)
		Me.tmrOnRealTime = New System.Windows.Forms.Timer(components)
		Me.picFaces = New System.Windows.Forms.PictureBox
		Me.picSteps = New System.Windows.Forms.PictureBox
		Me.picDice = New System.Windows.Forms.PictureBox
		Me.picDialogBrief = New System.Windows.Forms.Panel
		Me.tmrDialogBrief = New System.Windows.Forms.Timer(components)
		Me.picJournal = New System.Windows.Forms.Panel
		Me.picMicroBox = New System.Windows.Forms.Panel
		Me.picMicroMap = New System.Windows.Forms.PictureBox
		Me.picContainer = New System.Windows.Forms.PictureBox
		Me.picTalk = New System.Windows.Forms.PictureBox
		Me.picSearch = New System.Windows.Forms.PictureBox
		Me.picBuySell = New System.Windows.Forms.PictureBox
		Me.picMenu = New System.Windows.Forms.PictureBox
		Me.picConvoBottom = New System.Windows.Forms.PictureBox
		Me.picBlack = New System.Windows.Forms.PictureBox
		Me.picInventoryStatus = New System.Windows.Forms.PictureBox
		Me.picItem = New System.Windows.Forms.PictureBox
		Me.picInventory = New System.Windows.Forms.PictureBox
		Me.picConvo = New System.Windows.Forms.Panel
		Me.picConvoEnter = New System.Windows.Forms.PictureBox
		Me.picConvoList = New System.Windows.Forms.PictureBox
		Me.lblHyperLink = New System.Windows.Forms.Label
		Me.picTPic = New System.Windows.Forms.PictureBox
		Me.picTomeNew = New System.Windows.Forms.Panel
		Me.picCreateName = New System.Windows.Forms.PictureBox
		Me.picMainMenu = New System.Windows.Forms.PictureBox
		Me._picCMap_0 = New System.Windows.Forms.PictureBox
		Me._picCPic_0 = New System.Windows.Forms.PictureBox
		Me.picRuneSet = New System.Windows.Forms.PictureBox
		Me.picFonts = New System.Windows.Forms.PictureBox
		Me.picMisc = New System.Windows.Forms.PictureBox
		Me.picToHit = New System.Windows.Forms.PictureBox
		Me.picInvDrag = New System.Windows.Forms.PictureBox
		Me.picTmp = New System.Windows.Forms.PictureBox
		Me.tmrEncounterName = New System.Windows.Forms.Timer(components)
		Me.tmrMoveParty = New System.Windows.Forms.Timer(components)
		Me.tmrMoveMap = New System.Windows.Forms.Timer(components)
		Me._picSFXPic_0 = New System.Windows.Forms.PictureBox
		Me.picTSmall = New System.Windows.Forms.PictureBox
		Me.fraMediaPlayerMusic = New System.Windows.Forms.Panel
		Me.picBox = New System.Windows.Forms.Panel
		Me.picMap = New System.Windows.Forms.Panel
		Me.picWallPaper = New System.Windows.Forms.PictureBox
		Me.picGrid = New System.Windows.Forms.PictureBox
		Me.fraMediaPlayerAVI = New System.Windows.Forms.Panel
		Me.picCMap = New Microsoft.VisualBasic.Compatibility.VB6.PictureBoxArray(components)
		Me.picCPic = New Microsoft.VisualBasic.Compatibility.VB6.PictureBoxArray(components)
		Me.picSFXPic = New Microsoft.VisualBasic.Compatibility.VB6.PictureBoxArray(components)
		Me.picDialogBrief.SuspendLayout()
		Me.picJournal.SuspendLayout()
		Me.picMicroBox.SuspendLayout()
		Me.picConvo.SuspendLayout()
		Me.picTomeNew.SuspendLayout()
		Me.picBox.SuspendLayout()
		Me.picMap.SuspendLayout()
		Me.SuspendLayout()
		Me.ToolTip1.Active = True
		CType(Me.picCMap, System.ComponentModel.ISupportInitialize).BeginInit()
		CType(Me.picCPic, System.ComponentModel.ISupportInitialize).BeginInit()
		CType(Me.picSFXPic, System.ComponentModel.ISupportInitialize).BeginInit()
		Me.ControlBox = False
		Me.BackColor = System.Drawing.Color.Black
		Me.FormBorderStyle = System.Windows.Forms.FormBorderStyle.None
		Me.ClientSize = New System.Drawing.Size(781, 606)
		Me.Location = New System.Drawing.Point(7, 7)
		Me.Icon = CType(resources.GetObject("frmMain.Icon"), System.Drawing.Icon)
		Me.MinimizeBox = False
		Me.ShowInTaskbar = False
		Me.StartPosition = System.Windows.Forms.FormStartPosition.WindowsDefaultLocation
		Me.WindowState = System.Windows.Forms.FormWindowState.Maximized
		Me.Font = New System.Drawing.Font("Arial", 8!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
		Me.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font
		Me.Enabled = True
		Me.KeyPreview = False
		Me.MaximizeBox = True
		Me.Cursor = System.Windows.Forms.Cursors.Default
		Me.RightToLeft = System.Windows.Forms.RightToLeft.No
		Me.HelpButton = False
		Me.Name = "frmMain"
		Me.timerMusicLoop.Interval = 5000
		Me.timerMusicLoop.Enabled = True
		Me.tmrOnRealTime.Interval = 100
		Me.tmrOnRealTime.Enabled = True
		Me.picFaces.BackColor = System.Drawing.Color.FromARGB(255, 255, 128)
		Me.picFaces.ForeColor = System.Drawing.SystemColors.WindowText
		Me.picFaces.Size = New System.Drawing.Size(55, 55)
		Me.picFaces.Location = New System.Drawing.Point(72, 416)
		Me.picFaces.TabIndex = 10
		Me.picFaces.Visible = False
		Me.picFaces.Font = New System.Drawing.Font("Arial", 8!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
		Me.picFaces.Dock = System.Windows.Forms.DockStyle.None
		Me.picFaces.CausesValidation = True
		Me.picFaces.Enabled = True
		Me.picFaces.Cursor = System.Windows.Forms.Cursors.Default
		Me.picFaces.RightToLeft = System.Windows.Forms.RightToLeft.No
		Me.picFaces.TabStop = True
		Me.picFaces.SizeMode = System.Windows.Forms.PictureBoxSizeMode.Normal
		Me.picFaces.BorderStyle = System.Windows.Forms.BorderStyle.None
		Me.picFaces.Name = "picFaces"
		Me.picSteps.BackColor = System.Drawing.Color.FromARGB(0, 64, 64)
		Me.picSteps.ForeColor = System.Drawing.SystemColors.Window
		Me.picSteps.Size = New System.Drawing.Size(55, 55)
		Me.picSteps.Location = New System.Drawing.Point(72, 96)
		Me.picSteps.TabIndex = 39
		Me.picSteps.Visible = False
		Me.picSteps.Font = New System.Drawing.Font("Arial", 8!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
		Me.picSteps.Dock = System.Windows.Forms.DockStyle.None
		Me.picSteps.CausesValidation = True
		Me.picSteps.Enabled = True
		Me.picSteps.Cursor = System.Windows.Forms.Cursors.Default
		Me.picSteps.RightToLeft = System.Windows.Forms.RightToLeft.No
		Me.picSteps.TabStop = True
		Me.picSteps.SizeMode = System.Windows.Forms.PictureBoxSizeMode.AutoSize
		Me.picSteps.BorderStyle = System.Windows.Forms.BorderStyle.None
		Me.picSteps.Name = "picSteps"
		Me.picDice.BackColor = System.Drawing.Color.FromARGB(192, 192, 0)
		Me.picDice.Size = New System.Drawing.Size(55, 55)
		Me.picDice.Location = New System.Drawing.Point(8, 544)
		Me.picDice.TabIndex = 4
		Me.picDice.Visible = False
		Me.picDice.Font = New System.Drawing.Font("Arial", 8!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
		Me.picDice.Dock = System.Windows.Forms.DockStyle.None
		Me.picDice.CausesValidation = True
		Me.picDice.Enabled = True
		Me.picDice.ForeColor = System.Drawing.SystemColors.ControlText
		Me.picDice.Cursor = System.Windows.Forms.Cursors.Default
		Me.picDice.RightToLeft = System.Windows.Forms.RightToLeft.No
		Me.picDice.TabStop = True
		Me.picDice.SizeMode = System.Windows.Forms.PictureBoxSizeMode.AutoSize
		Me.picDice.BorderStyle = System.Windows.Forms.BorderStyle.None
		Me.picDice.Name = "picDice"
		Me.picDialogBrief.BackColor = System.Drawing.Color.FromARGB(64, 64, 0)
		Me.picDialogBrief.ForeColor = System.Drawing.SystemColors.WindowText
		Me.picDialogBrief.Size = New System.Drawing.Size(55, 55)
		Me.picDialogBrief.Location = New System.Drawing.Point(136, 160)
		Me.picDialogBrief.TabIndex = 38
		Me.picDialogBrief.Visible = False
		Me.picDialogBrief.Font = New System.Drawing.Font("Arial", 8!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
		Me.picDialogBrief.Dock = System.Windows.Forms.DockStyle.None
		Me.picDialogBrief.CausesValidation = True
		Me.picDialogBrief.Enabled = True
		Me.picDialogBrief.Cursor = System.Windows.Forms.Cursors.Default
		Me.picDialogBrief.RightToLeft = System.Windows.Forms.RightToLeft.No
		Me.picDialogBrief.TabStop = True
		Me.picDialogBrief.BorderStyle = System.Windows.Forms.BorderStyle.None
		Me.picDialogBrief.Name = "picDialogBrief"
		Me.tmrDialogBrief.Enabled = False
		Me.tmrDialogBrief.Interval = 1
		Me.picJournal.BackColor = System.Drawing.SystemColors.Window
		Me.picJournal.ForeColor = System.Drawing.SystemColors.WindowText
		Me.picJournal.Size = New System.Drawing.Size(370, 380)
		Me.picJournal.Location = New System.Drawing.Point(384, 544)
		Me.picJournal.TabIndex = 20
		Me.picJournal.Visible = False
		Me.picJournal.Font = New System.Drawing.Font("Arial", 8!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
		Me.picJournal.Dock = System.Windows.Forms.DockStyle.None
		Me.picJournal.CausesValidation = True
		Me.picJournal.Enabled = True
		Me.picJournal.Cursor = System.Windows.Forms.Cursors.Default
		Me.picJournal.RightToLeft = System.Windows.Forms.RightToLeft.No
		Me.picJournal.TabStop = True
		Me.picJournal.BorderStyle = System.Windows.Forms.BorderStyle.None
		Me.picJournal.Name = "picJournal"
		Me.picMicroBox.BackColor = System.Drawing.Color.Black
		Me.picMicroBox.ForeColor = System.Drawing.SystemColors.WindowText
		Me.picMicroBox.Size = New System.Drawing.Size(306, 229)
		Me.picMicroBox.Location = New System.Drawing.Point(32, 76)
		Me.picMicroBox.TabIndex = 21
		Me.picMicroBox.Visible = False
		Me.picMicroBox.Font = New System.Drawing.Font("Arial", 8!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
		Me.picMicroBox.Dock = System.Windows.Forms.DockStyle.None
		Me.picMicroBox.CausesValidation = True
		Me.picMicroBox.Enabled = True
		Me.picMicroBox.Cursor = System.Windows.Forms.Cursors.Default
		Me.picMicroBox.RightToLeft = System.Windows.Forms.RightToLeft.No
		Me.picMicroBox.TabStop = True
		Me.picMicroBox.BorderStyle = System.Windows.Forms.BorderStyle.None
		Me.picMicroBox.Name = "picMicroBox"
		Me.picMicroMap.BackColor = System.Drawing.Color.Black
		Me.picMicroMap.ForeColor = System.Drawing.SystemColors.WindowText
		Me.picMicroMap.Size = New System.Drawing.Size(309, 604)
		Me.picMicroMap.Location = New System.Drawing.Point(-3, -2)
		Me.picMicroMap.TabIndex = 22
		Me.picMicroMap.Font = New System.Drawing.Font("Arial", 8!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
		Me.picMicroMap.Dock = System.Windows.Forms.DockStyle.None
		Me.picMicroMap.CausesValidation = True
		Me.picMicroMap.Enabled = True
		Me.picMicroMap.Cursor = System.Windows.Forms.Cursors.Default
		Me.picMicroMap.RightToLeft = System.Windows.Forms.RightToLeft.No
		Me.picMicroMap.TabStop = True
		Me.picMicroMap.Visible = True
		Me.picMicroMap.SizeMode = System.Windows.Forms.PictureBoxSizeMode.AutoSize
		Me.picMicroMap.BorderStyle = System.Windows.Forms.BorderStyle.None
		Me.picMicroMap.Name = "picMicroMap"
		Me.picContainer.BackColor = System.Drawing.Color.FromARGB(255, 255, 192)
		Me.picContainer.ForeColor = System.Drawing.SystemColors.WindowText
		Me.picContainer.Size = New System.Drawing.Size(55, 55)
		Me.picContainer.Location = New System.Drawing.Point(200, 224)
		Me.picContainer.TabIndex = 23
		Me.picContainer.Visible = False
		Me.picContainer.Font = New System.Drawing.Font("Arial", 8!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
		Me.picContainer.Dock = System.Windows.Forms.DockStyle.None
		Me.picContainer.CausesValidation = True
		Me.picContainer.Enabled = True
		Me.picContainer.Cursor = System.Windows.Forms.Cursors.Default
		Me.picContainer.RightToLeft = System.Windows.Forms.RightToLeft.No
		Me.picContainer.TabStop = True
		Me.picContainer.SizeMode = System.Windows.Forms.PictureBoxSizeMode.AutoSize
		Me.picContainer.BorderStyle = System.Windows.Forms.BorderStyle.None
		Me.picContainer.Name = "picContainer"
		Me.picTalk.BackColor = System.Drawing.Color.FromARGB(128, 0, 128)
		Me.picTalk.ForeColor = System.Drawing.SystemColors.WindowText
		Me.picTalk.Size = New System.Drawing.Size(55, 55)
		Me.picTalk.Location = New System.Drawing.Point(8, 224)
		Me.picTalk.TabIndex = 27
		Me.picTalk.Visible = False
		Me.picTalk.Font = New System.Drawing.Font("Arial", 8!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
		Me.picTalk.Dock = System.Windows.Forms.DockStyle.None
		Me.picTalk.CausesValidation = True
		Me.picTalk.Enabled = True
		Me.picTalk.Cursor = System.Windows.Forms.Cursors.Default
		Me.picTalk.RightToLeft = System.Windows.Forms.RightToLeft.No
		Me.picTalk.TabStop = True
		Me.picTalk.SizeMode = System.Windows.Forms.PictureBoxSizeMode.AutoSize
		Me.picTalk.BorderStyle = System.Windows.Forms.BorderStyle.None
		Me.picTalk.Name = "picTalk"
		Me.picSearch.BackColor = System.Drawing.Color.FromARGB(128, 0, 128)
		Me.picSearch.ForeColor = System.Drawing.SystemColors.WindowText
		Me.picSearch.Size = New System.Drawing.Size(55, 55)
		Me.picSearch.Location = New System.Drawing.Point(200, 288)
		Me.picSearch.TabIndex = 25
		Me.picSearch.Visible = False
		Me.picSearch.Font = New System.Drawing.Font("Arial", 8!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
		Me.picSearch.Dock = System.Windows.Forms.DockStyle.None
		Me.picSearch.CausesValidation = True
		Me.picSearch.Enabled = True
		Me.picSearch.Cursor = System.Windows.Forms.Cursors.Default
		Me.picSearch.RightToLeft = System.Windows.Forms.RightToLeft.No
		Me.picSearch.TabStop = True
		Me.picSearch.SizeMode = System.Windows.Forms.PictureBoxSizeMode.AutoSize
		Me.picSearch.BorderStyle = System.Windows.Forms.BorderStyle.None
		Me.picSearch.Name = "picSearch"
		Me.picBuySell.BackColor = System.Drawing.Color.FromARGB(0, 0, 192)
		Me.picBuySell.ForeColor = System.Drawing.SystemColors.WindowText
		Me.picBuySell.Size = New System.Drawing.Size(55, 55)
		Me.picBuySell.Location = New System.Drawing.Point(200, 160)
		Me.picBuySell.TabIndex = 28
		Me.picBuySell.Visible = False
		Me.picBuySell.Font = New System.Drawing.Font("Arial", 8!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
		Me.picBuySell.Dock = System.Windows.Forms.DockStyle.None
		Me.picBuySell.CausesValidation = True
		Me.picBuySell.Enabled = True
		Me.picBuySell.Cursor = System.Windows.Forms.Cursors.Default
		Me.picBuySell.RightToLeft = System.Windows.Forms.RightToLeft.No
		Me.picBuySell.TabStop = True
		Me.picBuySell.SizeMode = System.Windows.Forms.PictureBoxSizeMode.AutoSize
		Me.picBuySell.BorderStyle = System.Windows.Forms.BorderStyle.None
		Me.picBuySell.Name = "picBuySell"
		Me.picMenu.BackColor = System.Drawing.Color.FromARGB(128, 128, 128)
		Me.picMenu.ForeColor = System.Drawing.SystemColors.WindowText
		Me.picMenu.Size = New System.Drawing.Size(55, 55)
		Me.picMenu.Location = New System.Drawing.Point(136, 480)
		Me.picMenu.TabIndex = 3
		Me.picMenu.Visible = False
		Me.picMenu.Font = New System.Drawing.Font("Arial", 8!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
		Me.picMenu.Dock = System.Windows.Forms.DockStyle.None
		Me.picMenu.CausesValidation = True
		Me.picMenu.Enabled = True
		Me.picMenu.Cursor = System.Windows.Forms.Cursors.Default
		Me.picMenu.RightToLeft = System.Windows.Forms.RightToLeft.No
		Me.picMenu.TabStop = True
		Me.picMenu.SizeMode = System.Windows.Forms.PictureBoxSizeMode.AutoSize
		Me.picMenu.BorderStyle = System.Windows.Forms.BorderStyle.None
		Me.picMenu.Name = "picMenu"
		Me.picConvoBottom.BackColor = System.Drawing.Color.FromARGB(224, 224, 224)
		Me.picConvoBottom.ForeColor = System.Drawing.SystemColors.WindowText
		Me.picConvoBottom.Size = New System.Drawing.Size(55, 55)
		Me.picConvoBottom.Location = New System.Drawing.Point(136, 288)
		Me.picConvoBottom.TabIndex = 37
		Me.picConvoBottom.Visible = False
		Me.picConvoBottom.Font = New System.Drawing.Font("Arial", 8!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
		Me.picConvoBottom.Dock = System.Windows.Forms.DockStyle.None
		Me.picConvoBottom.CausesValidation = True
		Me.picConvoBottom.Enabled = True
		Me.picConvoBottom.Cursor = System.Windows.Forms.Cursors.Default
		Me.picConvoBottom.RightToLeft = System.Windows.Forms.RightToLeft.No
		Me.picConvoBottom.TabStop = True
		Me.picConvoBottom.SizeMode = System.Windows.Forms.PictureBoxSizeMode.AutoSize
		Me.picConvoBottom.BorderStyle = System.Windows.Forms.BorderStyle.None
		Me.picConvoBottom.Name = "picConvoBottom"
		Me.picBlack.BackColor = System.Drawing.Color.FromARGB(255, 128, 128)
		Me.picBlack.ForeColor = System.Drawing.SystemColors.Window
		Me.picBlack.Size = New System.Drawing.Size(55, 55)
		Me.picBlack.Location = New System.Drawing.Point(8, 160)
		Me.picBlack.TabIndex = 5
		Me.picBlack.Visible = False
		Me.picBlack.Font = New System.Drawing.Font("Arial", 8!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
		Me.picBlack.Dock = System.Windows.Forms.DockStyle.None
		Me.picBlack.CausesValidation = True
		Me.picBlack.Enabled = True
		Me.picBlack.Cursor = System.Windows.Forms.Cursors.Default
		Me.picBlack.RightToLeft = System.Windows.Forms.RightToLeft.No
		Me.picBlack.TabStop = True
		Me.picBlack.SizeMode = System.Windows.Forms.PictureBoxSizeMode.AutoSize
		Me.picBlack.BorderStyle = System.Windows.Forms.BorderStyle.None
		Me.picBlack.Name = "picBlack"
		Me.picInventoryStatus.BackColor = System.Drawing.Color.Yellow
		Me.picInventoryStatus.ForeColor = System.Drawing.SystemColors.WindowText
		Me.picInventoryStatus.Size = New System.Drawing.Size(55, 55)
		Me.picInventoryStatus.Location = New System.Drawing.Point(200, 544)
		Me.picInventoryStatus.TabIndex = 36
		Me.picInventoryStatus.Visible = False
		Me.picInventoryStatus.Font = New System.Drawing.Font("Arial", 8!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
		Me.picInventoryStatus.Dock = System.Windows.Forms.DockStyle.None
		Me.picInventoryStatus.CausesValidation = True
		Me.picInventoryStatus.Enabled = True
		Me.picInventoryStatus.Cursor = System.Windows.Forms.Cursors.Default
		Me.picInventoryStatus.RightToLeft = System.Windows.Forms.RightToLeft.No
		Me.picInventoryStatus.TabStop = True
		Me.picInventoryStatus.SizeMode = System.Windows.Forms.PictureBoxSizeMode.AutoSize
		Me.picInventoryStatus.BorderStyle = System.Windows.Forms.BorderStyle.None
		Me.picInventoryStatus.Name = "picInventoryStatus"
		Me.picItem.BackColor = System.Drawing.Color.FromARGB(192, 192, 0)
		Me.picItem.ForeColor = System.Drawing.SystemColors.WindowText
		Me.picItem.Size = New System.Drawing.Size(55, 55)
		Me.picItem.Location = New System.Drawing.Point(200, 480)
		Me.picItem.TabIndex = 12
		Me.picItem.Visible = False
		Me.picItem.Font = New System.Drawing.Font("Arial", 8!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
		Me.picItem.Dock = System.Windows.Forms.DockStyle.None
		Me.picItem.CausesValidation = True
		Me.picItem.Enabled = True
		Me.picItem.Cursor = System.Windows.Forms.Cursors.Default
		Me.picItem.RightToLeft = System.Windows.Forms.RightToLeft.No
		Me.picItem.TabStop = True
		Me.picItem.SizeMode = System.Windows.Forms.PictureBoxSizeMode.AutoSize
		Me.picItem.BorderStyle = System.Windows.Forms.BorderStyle.None
		Me.picItem.Name = "picItem"
		Me.picInventory.BackColor = System.Drawing.Color.Yellow
		Me.picInventory.ForeColor = System.Drawing.SystemColors.WindowText
		Me.picInventory.Size = New System.Drawing.Size(55, 55)
		Me.picInventory.Location = New System.Drawing.Point(200, 416)
		Me.picInventory.TabIndex = 35
		Me.picInventory.Visible = False
		Me.picInventory.Font = New System.Drawing.Font("Arial", 8!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
		Me.picInventory.Dock = System.Windows.Forms.DockStyle.None
		Me.picInventory.CausesValidation = True
		Me.picInventory.Enabled = True
		Me.picInventory.Cursor = System.Windows.Forms.Cursors.Default
		Me.picInventory.RightToLeft = System.Windows.Forms.RightToLeft.No
		Me.picInventory.TabStop = True
		Me.picInventory.SizeMode = System.Windows.Forms.PictureBoxSizeMode.AutoSize
		Me.picInventory.BorderStyle = System.Windows.Forms.BorderStyle.None
		Me.picInventory.Name = "picInventory"
		Me.picConvo.BackColor = System.Drawing.Color.FromARGB(255, 192, 128)
		Me.picConvo.ForeColor = System.Drawing.SystemColors.WindowText
		Me.picConvo.Size = New System.Drawing.Size(452, 498)
		Me.picConvo.Location = New System.Drawing.Point(284, 36)
		Me.picConvo.TabIndex = 19
		Me.picConvo.Visible = False
		Me.picConvo.Font = New System.Drawing.Font("Arial", 8!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
		Me.picConvo.Dock = System.Windows.Forms.DockStyle.None
		Me.picConvo.CausesValidation = True
		Me.picConvo.Enabled = True
		Me.picConvo.Cursor = System.Windows.Forms.Cursors.Default
		Me.picConvo.RightToLeft = System.Windows.Forms.RightToLeft.No
		Me.picConvo.TabStop = True
		Me.picConvo.BorderStyle = System.Windows.Forms.BorderStyle.None
		Me.picConvo.Name = "picConvo"
		Me.picConvoEnter.BackColor = System.Drawing.Color.FromARGB(255, 128, 0)
		Me.picConvoEnter.ForeColor = System.Drawing.SystemColors.WindowText
		Me.picConvoEnter.Size = New System.Drawing.Size(334, 28)
		Me.picConvoEnter.Location = New System.Drawing.Point(90, 314)
		Me.picConvoEnter.TabIndex = 26
		Me.picConvoEnter.Visible = False
		Me.picConvoEnter.Font = New System.Drawing.Font("Arial", 8!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
		Me.picConvoEnter.Dock = System.Windows.Forms.DockStyle.None
		Me.picConvoEnter.CausesValidation = True
		Me.picConvoEnter.Enabled = True
		Me.picConvoEnter.Cursor = System.Windows.Forms.Cursors.Default
		Me.picConvoEnter.RightToLeft = System.Windows.Forms.RightToLeft.No
		Me.picConvoEnter.TabStop = True
		Me.picConvoEnter.SizeMode = System.Windows.Forms.PictureBoxSizeMode.AutoSize
		Me.picConvoEnter.BorderStyle = System.Windows.Forms.BorderStyle.None
		Me.picConvoEnter.Name = "picConvoEnter"
		Me.picConvoList.BackColor = System.Drawing.Color.FromARGB(128, 128, 0)
		Me.picConvoList.ForeColor = System.Drawing.SystemColors.WindowText
		Me.picConvoList.Size = New System.Drawing.Size(351, 274)
		Me.picConvoList.Location = New System.Drawing.Point(90, 34)
		Me.picConvoList.TabIndex = 24
		Me.picConvoList.Visible = False
		Me.picConvoList.Font = New System.Drawing.Font("Arial", 8!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
		Me.picConvoList.Dock = System.Windows.Forms.DockStyle.None
		Me.picConvoList.CausesValidation = True
		Me.picConvoList.Enabled = True
		Me.picConvoList.Cursor = System.Windows.Forms.Cursors.Default
		Me.picConvoList.RightToLeft = System.Windows.Forms.RightToLeft.No
		Me.picConvoList.TabStop = True
		Me.picConvoList.SizeMode = System.Windows.Forms.PictureBoxSizeMode.AutoSize
		Me.picConvoList.BorderStyle = System.Windows.Forms.BorderStyle.None
		Me.picConvoList.Name = "picConvoList"
		Me.lblHyperLink.TextAlign = System.Drawing.ContentAlignment.TopCenter
		Me.lblHyperLink.Text = "Label1"
		Me.lblHyperLink.Font = New System.Drawing.Font("Arial", 8.25!, System.Drawing.FontStyle.Bold Or System.Drawing.FontStyle.Underline Or System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
		Me.lblHyperLink.ForeColor = System.Drawing.SystemColors.highlightText
		Me.lblHyperLink.Size = New System.Drawing.Size(49, 16)
		Me.lblHyperLink.Location = New System.Drawing.Point(196, 384)
		Me.lblHyperLink.TabIndex = 40
		Me.lblHyperLink.Visible = False
		Me.lblHyperLink.BackColor = System.Drawing.Color.Transparent
		Me.lblHyperLink.Enabled = True
		Me.lblHyperLink.Cursor = System.Windows.Forms.Cursors.Default
		Me.lblHyperLink.RightToLeft = System.Windows.Forms.RightToLeft.No
		Me.lblHyperLink.UseMnemonic = True
		Me.lblHyperLink.AutoSize = False
		Me.lblHyperLink.BorderStyle = System.Windows.Forms.BorderStyle.None
		Me.lblHyperLink.Name = "lblHyperLink"
		Me.picTPic.BackColor = System.Drawing.Color.Black
		Me.picTPic.ForeColor = System.Drawing.SystemColors.WindowText
		Me.picTPic.Size = New System.Drawing.Size(55, 55)
		Me.picTPic.Location = New System.Drawing.Point(136, 96)
		Me.picTPic.TabIndex = 0
		Me.picTPic.Visible = False
		Me.picTPic.Font = New System.Drawing.Font("Arial", 8!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
		Me.picTPic.Dock = System.Windows.Forms.DockStyle.None
		Me.picTPic.CausesValidation = True
		Me.picTPic.Enabled = True
		Me.picTPic.Cursor = System.Windows.Forms.Cursors.Default
		Me.picTPic.RightToLeft = System.Windows.Forms.RightToLeft.No
		Me.picTPic.TabStop = True
		Me.picTPic.SizeMode = System.Windows.Forms.PictureBoxSizeMode.AutoSize
		Me.picTPic.BorderStyle = System.Windows.Forms.BorderStyle.None
		Me.picTPic.Name = "picTPic"
		Me.picTomeNew.BackColor = System.Drawing.Color.FromARGB(255, 128, 255)
		Me.picTomeNew.ForeColor = System.Drawing.SystemColors.WindowText
		Me.picTomeNew.Size = New System.Drawing.Size(55, 55)
		Me.picTomeNew.Location = New System.Drawing.Point(200, 352)
		Me.picTomeNew.TabIndex = 18
		Me.picTomeNew.Visible = False
		Me.picTomeNew.Font = New System.Drawing.Font("Arial", 8!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
		Me.picTomeNew.Dock = System.Windows.Forms.DockStyle.None
		Me.picTomeNew.CausesValidation = True
		Me.picTomeNew.Enabled = True
		Me.picTomeNew.Cursor = System.Windows.Forms.Cursors.Default
		Me.picTomeNew.RightToLeft = System.Windows.Forms.RightToLeft.No
		Me.picTomeNew.TabStop = True
		Me.picTomeNew.BorderStyle = System.Windows.Forms.BorderStyle.None
		Me.picTomeNew.Name = "picTomeNew"
		Me.picCreateName.BackColor = System.Drawing.Color.FromARGB(0, 128, 128)
		Me.picCreateName.ForeColor = System.Drawing.SystemColors.WindowText
		Me.picCreateName.Size = New System.Drawing.Size(256, 28)
		Me.picCreateName.Location = New System.Drawing.Point(16, 352)
		Me.picCreateName.TabIndex = 29
		Me.picCreateName.Visible = False
		Me.picCreateName.Font = New System.Drawing.Font("Arial", 8!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
		Me.picCreateName.Dock = System.Windows.Forms.DockStyle.None
		Me.picCreateName.CausesValidation = True
		Me.picCreateName.Enabled = True
		Me.picCreateName.Cursor = System.Windows.Forms.Cursors.Default
		Me.picCreateName.RightToLeft = System.Windows.Forms.RightToLeft.No
		Me.picCreateName.TabStop = True
		Me.picCreateName.SizeMode = System.Windows.Forms.PictureBoxSizeMode.AutoSize
		Me.picCreateName.BorderStyle = System.Windows.Forms.BorderStyle.None
		Me.picCreateName.Name = "picCreateName"
		Me.picMainMenu.BackColor = System.Drawing.SystemColors.Window
		Me.picMainMenu.ForeColor = System.Drawing.SystemColors.WindowText
		Me.picMainMenu.Size = New System.Drawing.Size(55, 55)
		Me.picMainMenu.Location = New System.Drawing.Point(136, 352)
		Me.picMainMenu.TabIndex = 15
		Me.picMainMenu.Visible = False
		Me.picMainMenu.Font = New System.Drawing.Font("Arial", 8!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
		Me.picMainMenu.Dock = System.Windows.Forms.DockStyle.None
		Me.picMainMenu.CausesValidation = True
		Me.picMainMenu.Enabled = True
		Me.picMainMenu.Cursor = System.Windows.Forms.Cursors.Default
		Me.picMainMenu.RightToLeft = System.Windows.Forms.RightToLeft.No
		Me.picMainMenu.TabStop = True
		Me.picMainMenu.SizeMode = System.Windows.Forms.PictureBoxSizeMode.AutoSize
		Me.picMainMenu.BorderStyle = System.Windows.Forms.BorderStyle.None
		Me.picMainMenu.Name = "picMainMenu"
		Me._picCMap_0.BackColor = System.Drawing.Color.Cyan
		Me._picCMap_0.ForeColor = System.Drawing.SystemColors.WindowText
		Me._picCMap_0.Size = New System.Drawing.Size(55, 55)
		Me._picCMap_0.Location = New System.Drawing.Point(136, 544)
		Me._picCMap_0.TabIndex = 11
		Me._picCMap_0.Visible = False
		Me._picCMap_0.Font = New System.Drawing.Font("Arial", 8!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
		Me._picCMap_0.Dock = System.Windows.Forms.DockStyle.None
		Me._picCMap_0.CausesValidation = True
		Me._picCMap_0.Enabled = True
		Me._picCMap_0.Cursor = System.Windows.Forms.Cursors.Default
		Me._picCMap_0.RightToLeft = System.Windows.Forms.RightToLeft.No
		Me._picCMap_0.TabStop = True
		Me._picCMap_0.SizeMode = System.Windows.Forms.PictureBoxSizeMode.AutoSize
		Me._picCMap_0.BorderStyle = System.Windows.Forms.BorderStyle.None
		Me._picCMap_0.Name = "_picCMap_0"
		Me._picCPic_0.BackColor = System.Drawing.Color.FromARGB(192, 255, 192)
		Me._picCPic_0.ForeColor = System.Drawing.SystemColors.WindowText
		Me._picCPic_0.Size = New System.Drawing.Size(55, 55)
		Me._picCPic_0.Location = New System.Drawing.Point(72, 544)
		Me._picCPic_0.TabIndex = 7
		Me._picCPic_0.Visible = False
		Me._picCPic_0.Font = New System.Drawing.Font("Arial", 8!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
		Me._picCPic_0.Dock = System.Windows.Forms.DockStyle.None
		Me._picCPic_0.CausesValidation = True
		Me._picCPic_0.Enabled = True
		Me._picCPic_0.Cursor = System.Windows.Forms.Cursors.Default
		Me._picCPic_0.RightToLeft = System.Windows.Forms.RightToLeft.No
		Me._picCPic_0.TabStop = True
		Me._picCPic_0.SizeMode = System.Windows.Forms.PictureBoxSizeMode.AutoSize
		Me._picCPic_0.BorderStyle = System.Windows.Forms.BorderStyle.None
		Me._picCPic_0.Name = "_picCPic_0"
		Me.picRuneSet.BackColor = System.Drawing.Color.FromARGB(255, 192, 255)
		Me.picRuneSet.Size = New System.Drawing.Size(55, 55)
		Me.picRuneSet.Location = New System.Drawing.Point(8, 480)
		Me.picRuneSet.TabIndex = 34
		Me.picRuneSet.Visible = False
		Me.picRuneSet.Font = New System.Drawing.Font("Arial", 8!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
		Me.picRuneSet.Dock = System.Windows.Forms.DockStyle.None
		Me.picRuneSet.CausesValidation = True
		Me.picRuneSet.Enabled = True
		Me.picRuneSet.ForeColor = System.Drawing.SystemColors.ControlText
		Me.picRuneSet.Cursor = System.Windows.Forms.Cursors.Default
		Me.picRuneSet.RightToLeft = System.Windows.Forms.RightToLeft.No
		Me.picRuneSet.TabStop = True
		Me.picRuneSet.SizeMode = System.Windows.Forms.PictureBoxSizeMode.AutoSize
		Me.picRuneSet.BorderStyle = System.Windows.Forms.BorderStyle.None
		Me.picRuneSet.Name = "picRuneSet"
		Me.picFonts.Size = New System.Drawing.Size(55, 55)
		Me.picFonts.Location = New System.Drawing.Point(8, 416)
		Me.picFonts.TabIndex = 33
		Me.picFonts.Visible = False
		Me.picFonts.Font = New System.Drawing.Font("Arial", 8!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
		Me.picFonts.Dock = System.Windows.Forms.DockStyle.None
		Me.picFonts.BackColor = System.Drawing.SystemColors.Control
		Me.picFonts.CausesValidation = True
		Me.picFonts.Enabled = True
		Me.picFonts.ForeColor = System.Drawing.SystemColors.ControlText
		Me.picFonts.Cursor = System.Windows.Forms.Cursors.Default
		Me.picFonts.RightToLeft = System.Windows.Forms.RightToLeft.No
		Me.picFonts.TabStop = True
		Me.picFonts.SizeMode = System.Windows.Forms.PictureBoxSizeMode.AutoSize
		Me.picFonts.BorderStyle = System.Windows.Forms.BorderStyle.None
		Me.picFonts.Name = "picFonts"
		Me.picMisc.BackColor = System.Drawing.Color.FromARGB(192, 0, 0)
		Me.picMisc.ForeColor = System.Drawing.SystemColors.WindowText
		Me.picMisc.Size = New System.Drawing.Size(55, 55)
		Me.picMisc.Location = New System.Drawing.Point(8, 352)
		Me.picMisc.TabIndex = 32
		Me.picMisc.Visible = False
		Me.picMisc.Font = New System.Drawing.Font("Arial", 8!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
		Me.picMisc.Dock = System.Windows.Forms.DockStyle.None
		Me.picMisc.CausesValidation = True
		Me.picMisc.Enabled = True
		Me.picMisc.Cursor = System.Windows.Forms.Cursors.Default
		Me.picMisc.RightToLeft = System.Windows.Forms.RightToLeft.No
		Me.picMisc.TabStop = True
		Me.picMisc.SizeMode = System.Windows.Forms.PictureBoxSizeMode.AutoSize
		Me.picMisc.BorderStyle = System.Windows.Forms.BorderStyle.None
		Me.picMisc.Name = "picMisc"
		Me.picToHit.BackColor = System.Drawing.Color.FromARGB(128, 0, 128)
		Me.picToHit.ForeColor = System.Drawing.SystemColors.WindowText
		Me.picToHit.Size = New System.Drawing.Size(55, 55)
		Me.picToHit.Location = New System.Drawing.Point(136, 416)
		Me.picToHit.TabIndex = 14
		Me.picToHit.Visible = False
		Me.picToHit.Font = New System.Drawing.Font("Arial", 8!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
		Me.picToHit.Dock = System.Windows.Forms.DockStyle.None
		Me.picToHit.CausesValidation = True
		Me.picToHit.Enabled = True
		Me.picToHit.Cursor = System.Windows.Forms.Cursors.Default
		Me.picToHit.RightToLeft = System.Windows.Forms.RightToLeft.No
		Me.picToHit.TabStop = True
		Me.picToHit.SizeMode = System.Windows.Forms.PictureBoxSizeMode.AutoSize
		Me.picToHit.BorderStyle = System.Windows.Forms.BorderStyle.None
		Me.picToHit.Name = "picToHit"
		Me.picInvDrag.BackColor = System.Drawing.Color.White
		Me.picInvDrag.Size = New System.Drawing.Size(55, 55)
		Me.picInvDrag.Location = New System.Drawing.Point(72, 480)
		Me.picInvDrag.TabIndex = 17
		Me.picInvDrag.Visible = False
		Me.picInvDrag.Font = New System.Drawing.Font("Arial", 8!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
		Me.picInvDrag.Dock = System.Windows.Forms.DockStyle.None
		Me.picInvDrag.CausesValidation = True
		Me.picInvDrag.Enabled = True
		Me.picInvDrag.ForeColor = System.Drawing.SystemColors.ControlText
		Me.picInvDrag.Cursor = System.Windows.Forms.Cursors.Default
		Me.picInvDrag.RightToLeft = System.Windows.Forms.RightToLeft.No
		Me.picInvDrag.TabStop = True
		Me.picInvDrag.SizeMode = System.Windows.Forms.PictureBoxSizeMode.AutoSize
		Me.picInvDrag.BorderStyle = System.Windows.Forms.BorderStyle.None
		Me.picInvDrag.Name = "picInvDrag"
		Me.picTmp.BackColor = System.Drawing.Color.FromARGB(255, 128, 128)
		Me.picTmp.ForeColor = System.Drawing.SystemColors.WindowText
		Me.picTmp.Size = New System.Drawing.Size(55, 55)
		Me.picTmp.Location = New System.Drawing.Point(136, 224)
		Me.picTmp.TabIndex = 9
		Me.picTmp.Visible = False
		Me.picTmp.Font = New System.Drawing.Font("Arial", 8!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
		Me.picTmp.Dock = System.Windows.Forms.DockStyle.None
		Me.picTmp.CausesValidation = True
		Me.picTmp.Enabled = True
		Me.picTmp.Cursor = System.Windows.Forms.Cursors.Default
		Me.picTmp.RightToLeft = System.Windows.Forms.RightToLeft.No
		Me.picTmp.TabStop = True
		Me.picTmp.SizeMode = System.Windows.Forms.PictureBoxSizeMode.AutoSize
		Me.picTmp.BorderStyle = System.Windows.Forms.BorderStyle.None
		Me.picTmp.Name = "picTmp"
		Me.tmrEncounterName.Interval = 400
		Me.tmrEncounterName.Enabled = True
		Me.tmrMoveParty.Enabled = False
		Me.tmrMoveParty.Interval = 1
		Me.tmrMoveMap.Interval = 1
		Me.tmrMoveMap.Enabled = True
		Me._picSFXPic_0.BackColor = System.Drawing.Color.FromARGB(255, 192, 128)
		Me._picSFXPic_0.ForeColor = System.Drawing.SystemColors.WindowText
		Me._picSFXPic_0.Size = New System.Drawing.Size(55, 55)
		Me._picSFXPic_0.Location = New System.Drawing.Point(72, 288)
		Me._picSFXPic_0.TabIndex = 8
		Me._picSFXPic_0.Visible = False
		Me._picSFXPic_0.Font = New System.Drawing.Font("Arial", 8!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
		Me._picSFXPic_0.Dock = System.Windows.Forms.DockStyle.None
		Me._picSFXPic_0.CausesValidation = True
		Me._picSFXPic_0.Enabled = True
		Me._picSFXPic_0.Cursor = System.Windows.Forms.Cursors.Default
		Me._picSFXPic_0.RightToLeft = System.Windows.Forms.RightToLeft.No
		Me._picSFXPic_0.TabStop = True
		Me._picSFXPic_0.SizeMode = System.Windows.Forms.PictureBoxSizeMode.AutoSize
		Me._picSFXPic_0.BorderStyle = System.Windows.Forms.BorderStyle.None
		Me._picSFXPic_0.Name = "_picSFXPic_0"
		Me.picTSmall.BackColor = System.Drawing.Color.FromARGB(224, 224, 224)
		Me.picTSmall.ForeColor = System.Drawing.SystemColors.WindowText
		Me.picTSmall.Size = New System.Drawing.Size(55, 55)
		Me.picTSmall.Location = New System.Drawing.Point(72, 352)
		Me.picTSmall.TabIndex = 16
		Me.picTSmall.Visible = False
		Me.picTSmall.Font = New System.Drawing.Font("Arial", 8!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
		Me.picTSmall.Dock = System.Windows.Forms.DockStyle.None
		Me.picTSmall.CausesValidation = True
		Me.picTSmall.Enabled = True
		Me.picTSmall.Cursor = System.Windows.Forms.Cursors.Default
		Me.picTSmall.RightToLeft = System.Windows.Forms.RightToLeft.No
		Me.picTSmall.TabStop = True
		Me.picTSmall.SizeMode = System.Windows.Forms.PictureBoxSizeMode.AutoSize
		Me.picTSmall.BorderStyle = System.Windows.Forms.BorderStyle.None
		Me.picTSmall.Name = "picTSmall"
		Me.fraMediaPlayerMusic.BackColor = System.Drawing.Color.Blue
		Me.fraMediaPlayerMusic.BorderStyle = System.Windows.Forms.BorderStyle.None
		Me.fraMediaPlayerMusic.Text = "Frame1"
		Me.fraMediaPlayerMusic.Size = New System.Drawing.Size(55, 55)
		Me.fraMediaPlayerMusic.Location = New System.Drawing.Point(8, 288)
		Me.fraMediaPlayerMusic.TabIndex = 30
		Me.fraMediaPlayerMusic.Visible = False
		Me.fraMediaPlayerMusic.Font = New System.Drawing.Font("Arial", 8!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
		Me.fraMediaPlayerMusic.Enabled = True
		Me.fraMediaPlayerMusic.ForeColor = System.Drawing.SystemColors.ControlText
		Me.fraMediaPlayerMusic.Cursor = System.Windows.Forms.Cursors.Default
		Me.fraMediaPlayerMusic.RightToLeft = System.Windows.Forms.RightToLeft.No
		Me.fraMediaPlayerMusic.Name = "fraMediaPlayerMusic"
		Me.picBox.BackColor = System.Drawing.SystemColors.Window
		Me.picBox.ForeColor = System.Drawing.SystemColors.WindowText
		Me.picBox.Size = New System.Drawing.Size(692, 458)
		Me.picBox.Location = New System.Drawing.Point(64, 8)
		Me.picBox.TabIndex = 1
		Me.picBox.Font = New System.Drawing.Font("Arial", 8!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
		Me.picBox.Dock = System.Windows.Forms.DockStyle.None
		Me.picBox.CausesValidation = True
		Me.picBox.Enabled = True
		Me.picBox.Cursor = System.Windows.Forms.Cursors.Default
		Me.picBox.RightToLeft = System.Windows.Forms.RightToLeft.No
		Me.picBox.TabStop = True
		Me.picBox.Visible = True
		Me.picBox.BorderStyle = System.Windows.Forms.BorderStyle.None
		Me.picBox.Name = "picBox"
		Me.picMap.BackColor = System.Drawing.Color.Black
		Me.picMap.ForeColor = System.Drawing.SystemColors.WindowText
		Me.picMap.Size = New System.Drawing.Size(724, 548)
		Me.picMap.Location = New System.Drawing.Point(-1, 0)
		Me.picMap.TabIndex = 2
		Me.picMap.Font = New System.Drawing.Font("Arial", 8!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
		Me.picMap.Dock = System.Windows.Forms.DockStyle.None
		Me.picMap.CausesValidation = True
		Me.picMap.Enabled = True
		Me.picMap.Cursor = System.Windows.Forms.Cursors.Default
		Me.picMap.RightToLeft = System.Windows.Forms.RightToLeft.No
		Me.picMap.TabStop = True
		Me.picMap.Visible = True
		Me.picMap.BorderStyle = System.Windows.Forms.BorderStyle.None
		Me.picMap.Name = "picMap"
		Me.picWallPaper.BackColor = System.Drawing.Color.FromARGB(128, 128, 255)
		Me.picWallPaper.ForeColor = System.Drawing.SystemColors.WindowText
		Me.picWallPaper.Size = New System.Drawing.Size(55, 55)
		Me.picWallPaper.Location = New System.Drawing.Point(8, 152)
		Me.picWallPaper.TabIndex = 13
		Me.picWallPaper.Visible = False
		Me.picWallPaper.Font = New System.Drawing.Font("Arial", 8!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
		Me.picWallPaper.Dock = System.Windows.Forms.DockStyle.None
		Me.picWallPaper.CausesValidation = True
		Me.picWallPaper.Enabled = True
		Me.picWallPaper.Cursor = System.Windows.Forms.Cursors.Default
		Me.picWallPaper.RightToLeft = System.Windows.Forms.RightToLeft.No
		Me.picWallPaper.TabStop = True
		Me.picWallPaper.SizeMode = System.Windows.Forms.PictureBoxSizeMode.AutoSize
		Me.picWallPaper.BorderStyle = System.Windows.Forms.BorderStyle.None
		Me.picWallPaper.Name = "picWallPaper"
		Me.picGrid.BackColor = System.Drawing.Color.FromARGB(128, 128, 128)
		Me.picGrid.ForeColor = System.Drawing.SystemColors.WindowText
		Me.picGrid.Size = New System.Drawing.Size(55, 55)
		Me.picGrid.Location = New System.Drawing.Point(8, 216)
		Me.picGrid.TabIndex = 6
		Me.picGrid.Visible = False
		Me.picGrid.Font = New System.Drawing.Font("Arial", 8!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
		Me.picGrid.Dock = System.Windows.Forms.DockStyle.None
		Me.picGrid.CausesValidation = True
		Me.picGrid.Enabled = True
		Me.picGrid.RightToLeft = System.Windows.Forms.RightToLeft.No
		Me.picGrid.TabStop = True
		Me.picGrid.SizeMode = System.Windows.Forms.PictureBoxSizeMode.Normal
		Me.picGrid.BorderStyle = System.Windows.Forms.BorderStyle.None
		Me.picGrid.Name = "picGrid"
		Me.fraMediaPlayerAVI.BackColor = System.Drawing.Color.Black
		Me.fraMediaPlayerAVI.BorderStyle = System.Windows.Forms.BorderStyle.None
		Me.fraMediaPlayerAVI.Size = New System.Drawing.Size(601, 417)
		Me.fraMediaPlayerAVI.Location = New System.Drawing.Point(88, 40)
		Me.fraMediaPlayerAVI.TabIndex = 31
		Me.fraMediaPlayerAVI.Visible = False
		Me.fraMediaPlayerAVI.Font = New System.Drawing.Font("Arial", 8!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
		Me.fraMediaPlayerAVI.Enabled = True
		Me.fraMediaPlayerAVI.ForeColor = System.Drawing.SystemColors.ControlText
		Me.fraMediaPlayerAVI.Cursor = System.Windows.Forms.Cursors.Default
		Me.fraMediaPlayerAVI.RightToLeft = System.Windows.Forms.RightToLeft.No
		Me.fraMediaPlayerAVI.Name = "fraMediaPlayerAVI"
		Me.Controls.Add(picFaces)
		Me.Controls.Add(picSteps)
		Me.Controls.Add(picDice)
		Me.Controls.Add(picDialogBrief)
		Me.Controls.Add(picJournal)
		Me.Controls.Add(picContainer)
		Me.Controls.Add(picTalk)
		Me.Controls.Add(picSearch)
		Me.Controls.Add(picBuySell)
		Me.Controls.Add(picMenu)
		Me.Controls.Add(picConvoBottom)
		Me.Controls.Add(picBlack)
		Me.Controls.Add(picInventoryStatus)
		Me.Controls.Add(picItem)
		Me.Controls.Add(picInventory)
		Me.Controls.Add(picConvo)
		Me.Controls.Add(picTPic)
		Me.Controls.Add(picTomeNew)
		Me.Controls.Add(picMainMenu)
		Me.Controls.Add(_picCMap_0)
		Me.Controls.Add(_picCPic_0)
		Me.Controls.Add(picRuneSet)
		Me.Controls.Add(picFonts)
		Me.Controls.Add(picMisc)
		Me.Controls.Add(picToHit)
		Me.Controls.Add(picInvDrag)
		Me.Controls.Add(picTmp)
		Me.Controls.Add(_picSFXPic_0)
		Me.Controls.Add(picTSmall)
		Me.Controls.Add(fraMediaPlayerMusic)
		Me.Controls.Add(picBox)
		Me.picJournal.Controls.Add(picMicroBox)
		Me.picMicroBox.Controls.Add(picMicroMap)
		Me.picConvo.Controls.Add(picConvoEnter)
		Me.picConvo.Controls.Add(picConvoList)
		Me.picConvo.Controls.Add(lblHyperLink)
		Me.picTomeNew.Controls.Add(picCreateName)
		Me.picBox.Controls.Add(picMap)
		Me.picMap.Controls.Add(picWallPaper)
		Me.picMap.Controls.Add(picGrid)
		Me.picMap.Controls.Add(fraMediaPlayerAVI)
		Me.picCMap.SetIndex(_picCMap_0, CType(0, Short))
		Me.picCPic.SetIndex(_picCPic_0, CType(0, Short))
		Me.picSFXPic.SetIndex(_picSFXPic_0, CType(0, Short))
		CType(Me.picSFXPic, System.ComponentModel.ISupportInitialize).EndInit()
		CType(Me.picCPic, System.ComponentModel.ISupportInitialize).EndInit()
		CType(Me.picCMap, System.ComponentModel.ISupportInitialize).EndInit()
		Me.picDialogBrief.ResumeLayout(False)
		Me.picJournal.ResumeLayout(False)
		Me.picMicroBox.ResumeLayout(False)
		Me.picConvo.ResumeLayout(False)
		Me.picTomeNew.ResumeLayout(False)
		Me.picBox.ResumeLayout(False)
		Me.picMap.ResumeLayout(False)
		Me.ResumeLayout(False)
		Me.PerformLayout()
	End Sub
#End Region 
End Class