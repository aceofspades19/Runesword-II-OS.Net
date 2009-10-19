Option Strict Off
Option Explicit On
Imports VB = Microsoft.VisualBasic
Friend Class frmMain
	Inherits System.Windows.Forms.Form
	
	Dim isGameInit As Boolean
	
	'*************************************************************************
	' Combat related variables
	'*************************************************************************
	'Dim Turns(48) As Grid
	'UPGRADE_NOTE: Paint was upgraded to Paint_Renamed. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="A9E4979A-37FA-4718-9994-97DD76ED70A7"'
	Dim Paint_Renamed(48) As Short
	Dim CombatTurn As Short
	Dim TurnNow As Short
	Dim Target As Short
	Dim CombatMouseX, CombatMouseY As Short
	Dim LastTarget As Short
	Dim TargetCreatureRange As Short
	Dim CombatStepRow(50) As Short
	Dim CombatStepCol(50) As Short
	Dim CombatStepTop As Short
	Dim Combat20Bonus, Combat20, CombatDmgBonus As Short
	Dim CombatDiceType(5) As Short
    Dim CombatDiceValue(5) As Short
    Dim MainMap As Map
    Dim tome As Tome

	Const bdCombatGridWidth As Short = 48
	Const bdCombatGridHeight As Short = 24
	Const bdCombatTop As Short = 135
	Const bdCombatLeft As Short = 12
	Const bdContextDice As Short = 31
	
	Public Enum MOVEDIRECTION
		bdMoveToward = 0
		bdMoveAway = 1
		bdMoveClick = 2
		bdMoveFlee = 3
	End Enum
	
	Public Enum MOVEMOTIVE
		bdMoveClosest = 0
		bdMoveFarthest = 1
		bdMoveStrongest = 2
		bdMoveWeakest = 3
		bdMoveRandom = 4
		bdMoveTarget = 5
	End Enum
	
	Const bdSFXFromHead As Short = 0
	Const bdSFXFromCenter As Short = 1
	Const bdSFXFling As Short = 0
	Const bdSFXStream As Short = 1
	Const bdSFXBurstHere As Short = 2
	Const bdSFXBurstThere As Short = 3
	Const bdSFXDown As Short = 4
	Const bdSFXAround As Short = 5
	Const bdSFXArch As Short = 6
	Const bdSFXSlow As Short = 0
	Const bdSFXMedium As Short = 1
	Const bdSFXFast As Short = 2
	
	Const bdCombatAttitudeWait As Short = 1
	Const bdCombatAttitudeWait2 As Short = 2
	Const bdCombatAttitudeDefend As Short = 3
	
	' Variables to store picture file names and info
	Dim TileSetLoaded As String
	Dim SFXFile() As String
	Dim MaxSFX As Short
	Dim CommonScale As Double
	Dim PicSize As Double
	Dim LosDir(5, 5) As Short
	Dim picBlackHeightSmall As Short
	Const bdBlackWidth As Short = 288
	Dim MicroMapScale As Double
	
	' Variables for in game Menus
	Dim PartyLeft As Short
	Dim SplashNow, KillSplash As Short
	Dim CutSceneNow As Short
	Dim ActionNow As Short
	Dim MenuNow As Short
	Dim ClickY, ClickX, ClickZ As Short
	Dim HintX, HintY As Short
	Dim RuneMenuX(19) As Short
	Dim RuneMenuY(19) As Short
	Dim BtnClick As Short
	Dim TargetAllowDead As Short
	Const bdMenuActionDefault As Short = -1
	Const bdMenuDefault As Short = 0
	Const bdMenuOptions As Short = 1
	Const bdMenuInventory As Short = 3
	Const bdMenuSearch As Short = 4
	Const bdMenuTalkWith As Short = 7
	Const bdMenuEncounter As Short = 9
	Const bdMenuMapParty As Short = 10
	Const bdCombatMenuPC As Short = 13
	Const bdMenuTargetCreature As Short = 15
	Const bdMenuTargetItem As Short = 16
	Const bdMenuTargetAny As Short = 17
	Const bdMenuTargetParty As Short = 18
	
	' Variables for find path routine
	Dim DirX(7) As Short
	Dim DirY(7) As Short
	Dim TotalStepsTaken As Short
	
	' Constants for drawing the map
	Const bdMapBottom As Short = 0
	Const bdMapMiddle As Short = 1
	Const bdMapItems As Short = 2
	Const bdMapMonsters As Short = 3
	Const bdMapParty As Short = 4
	Const bdMapDoors As Short = 5
	Const bdMapTop As Short = 6
	Const bdMapDim As Short = 7
	
	Private WorldIndex As Short
	Private ScrollWorldDesc As Short
	
	Private FromEntryPoint As EntryPoint
	
	Private Const MM_TEXT As Short = 1
	Private Structure TEXTMETRIC
		Dim tmHeight As Integer
		Dim tmAscent As Integer
		Dim tmDescent As Integer
		Dim tmInternalLeading As Integer
		Dim tmExternalLeading As Integer
		Dim tmAveCharWidth As Integer
		Dim tmMaxCharWidth As Integer
		Dim tmWeight As Integer
		Dim tmOverhang As Integer
		Dim tmDigitizedAspectX As Integer
		Dim tmDigitizedAspectY As Integer
		Dim tmFirstChar As Byte
		Dim tmLastChar As Byte
		Dim tmDefaultChar As Byte
		Dim tmBreakChar As Byte
		Dim tmItalic As Byte
		Dim tmUnderlined As Byte
		Dim tmStruckOut As Byte
		Dim tmPitchAndFamily As Byte
		Dim tmCharSet As Byte
	End Structure
	'UPGRADE_WARNING: Structure TEXTMETRIC may require marshalling attributes to be passed as an argument in this Declare statement. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="C429C3A5-5D47-4CD9-8F51-74A1616405DC"'
	Private Declare Function GetTextMetrics Lib "gdi32"  Alias "GetTextMetricsA"(ByVal hdc As Integer, ByRef lpMetrics As TEXTMETRIC) As Integer
	Private Declare Function SetMapMode Lib "gdi32" (ByVal hdc As Integer, ByVal nMapMode As Integer) As Integer
	Private Declare Function GetDesktopWindow Lib "user32" () As Integer
	Private Declare Function GetWindowDC Lib "user32" (ByVal hWnd As Integer) As Integer
	Private Declare Function ReleaseDC Lib "user32" (ByVal hWnd As Integer, ByVal hdc As Integer) As Integer
	
	Function SmallFonts() As Boolean
		Dim hdc, hWnd As Integer
		Dim PrevMapMode As Integer
		Dim tm As TEXTMETRIC
		' Set the default return value to small fonts
		SmallFonts = True
		' Get the handle of the desktop window
		hWnd = GetDesktopWindow()
		' Get the device context for the desktop
		hdc = GetWindowDC(hWnd)
		If hdc Then
			' Set the mapping mode to pixels
			PrevMapMode = SetMapMode(hdc, MM_TEXT)
			' Get the size of the system font
			GetTextMetrics(hdc, tm)
			' Set the mapping mode back to what it was
			PrevMapMode = SetMapMode(hdc, PrevMapMode)
			' Release the device context
			ReleaseDC(hWnd, hdc)
			' If the system font is more than 16 pixels high, then large fonts are being used
			If tm.tmHeight > 16 Then SmallFonts = False
		End If
	End Function
	
	'UPGRADE_NOTE: Map was upgraded to Map_Renamed. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="A9E4979A-37FA-4718-9994-97DD76ED70A7"'
	Public Function GetRunePool(ByRef Index As Short) As Short
		Dim Map_Renamed As Object
		'UPGRADE_WARNING: Couldn't resolve default property of object Map.Runes. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
        GetRunePool = MainMap.Runes(Index)
	End Function
	
	'UPGRADE_NOTE: Map was upgraded to Map_Renamed. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="A9E4979A-37FA-4718-9994-97DD76ED70A7"'
	Public Sub LetRunePool(ByRef Index As Short, ByRef NewValue As Short)
		Dim Map_Renamed As Object
		'UPGRADE_WARNING: Couldn't resolve default property of object Map.Runes. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
        MainMap.Runes(Index) = NewValue
	End Sub
	
	Public Sub CutSceneVideo(ByRef VideoFile As String, Optional ByRef AlternatePath As Object = Nothing)
		Dim FileName As String
		Dim PosY As Short
		Dim OldTalk, OldInventory, OldFrozen, OldConvo, OldGrid As Short
		' CutScene trumps everything else
		OldFrozen = Frozen
		OldInventory = picInventory.Visible
		OldConvo = picConvo.Visible
		OldTalk = picTalk.Visible
		OldGrid = picGrid.Visible
		Frozen = True
		picInventory.Visible = False
		picConvo.Visible = False
		picTalk.Visible = False
		picGrid.Visible = False
		'UPGRADE_WARNING: Screen property Screen.MousePointer has a new behavior. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6BA9B8D2-2A32-4B6E-8D36-44949974A5B4"'
		System.Windows.Forms.Cursor.Current = System.Windows.Forms.Cursors.WaitCursor
		'UPGRADE_ISSUE: PictureBox method picMap.Cls was not upgraded. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="CC4C7EC0-C903-48FC-ACCC-81861D12DA4A"'
        picMap.Controls.Clear()

		' Find location of video
		'    FileName = Dir$(gAppPath & "\data\video\" & VideoFile)
		'UPGRADE_WARNING: Dir has a new behavior. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="9B7D5ADD-D8FE-4819-A36C-6DEDAF088CC7"'
		FileName = Dir(gDataPath & "\video\" & VideoFile)
		'UPGRADE_NOTE: IsMissing() was changed to IsNothing(). Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="8AE1CB93-37AB-439A-A4FF-BE3B6760BB23"'
		If FileName = "" And IsNothing(AlternatePath) = False Then
			'UPGRADE_WARNING: Couldn't resolve default property of object AlternatePath. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
			'UPGRADE_WARNING: Dir has a new behavior. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="9B7D5ADD-D8FE-4819-A36C-6DEDAF088CC7"'
			FileName = Dir(AlternatePath & "\" & VideoFile)
			If FileName = "" Then
				'            Exit Sub  <-- Bug:  exiting now will not set the parameters back to their previous state
				GoTo BackToPreviousState
			Else
				'ChDir AlternatePath
				'UPGRADE_WARNING: Couldn't resolve default property of object AlternatePath. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
				FileName = AlternatePath & "\" & VideoFile
			End If
		Else
			'    ChDir gAppPath & "\data\video"
			'        FileName = gAppPath & "\data\video\" & VideoFile
			FileName = gDataPath & "\video\" & VideoFile
		End If
		' Center Video display and play
		fraMediaPlayerAVI.Left = 0 : fraMediaPlayerAVI.Top = 0
		fraMediaPlayerAVI.Width = picBox.Width : fraMediaPlayerAVI.Height = picBox.Height
		'    MediaPlayerAVI.Filename = VideoFile
		'    MediaPlayerAVI.Left = 0
		'    MediaPlayerAVI.Top = 0
		'    MediaPlayerAVI.Height = fraMediaPlayerAVI.Height * Screen.TwipsPerPixelY
		'    MediaPlayerAVI.Width = fraMediaPlayerAVI.Width * Screen.TwipsPerPixelX
		'    MediaPlayerAVI.Play
		Dim mmAVI As New IMCI
		If mmAVI.InitAVI(FileName, fraMediaPlayerAVI.Handle.ToInt32, "CutScene") Then
			mmAVI.Left_Renamed = (fraMediaPlayerAVI.Width - mmAVI.Width) / 2
			mmAVI.Top = (fraMediaPlayerAVI.Height - mmAVI.Height) / 2
			mmAVI.Play()
			If mmAVI.GetError() = 0 Then
				fraMediaPlayerAVI.Visible = True
				picMap.Refresh()
				MessageShow("Click to Continue", 0)
				'UPGRADE_WARNING: Screen property Screen.MousePointer has a new behavior. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6BA9B8D2-2A32-4B6E-8D36-44949974A5B4"'
				System.Windows.Forms.Cursor.Current = System.Windows.Forms.Cursors.Default
				CutSceneNow = True
				Do Until CutSceneNow = False
					System.Windows.Forms.Application.DoEvents()
					If mmAVI.Status <> IMCI.VIDEOSTATE.vsPLAYING Then CutSceneNow = False
				Loop 
			Else
				Call oErr.LogText("Error CutSceneVideo: Video (#" & mmAVI.GetError() & ")")
			End If
			mmAVI.mciClose()
			fraMediaPlayerAVI.Visible = False
		Else
			Call oErr.LogText("Error CutSceneVideo: Video (#" & mmAVI.GetError() & ")")
		End If
BackToPreviousState: 
		' Put everything back
		Frozen = OldFrozen
		picInventory.Visible = OldInventory
		picConvo.Visible = OldConvo
		picTalk.Visible = OldTalk
		picGrid.Visible = OldGrid
		If picGrid.Visible = True Then
			CombatWallpaper()
			CombatDraw()
		Else
			DrawMapAll()
		End If
	End Sub
	
	'UPGRADE_NOTE: Tome was upgraded to Tome_Renamed. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="A9E4979A-37FA-4718-9994-97DD76ED70A7"'
	'UPGRADE_NOTE: Text was upgraded to Text_Renamed. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="A9E4979A-37FA-4718-9994-97DD76ED70A7"'
	Public Sub CutScene(ByRef Style As Short, ByRef Text_Renamed As String, ByRef PictureFile As String)
		Dim Tome_Renamed As Object
		Dim FileName As String
		Dim PosY As Short
		Dim OldTalk, OldInventory, OldFrozen, OldConvo, OldGrid As Short
		' CutScene trumps everything else
		OldFrozen = Frozen
		OldInventory = picInventory.Visible
		OldConvo = picConvo.Visible
		OldTalk = picTalk.Visible
		OldGrid = picGrid.Visible
		Frozen = True
		picInventory.Visible = False
		picConvo.Visible = False
		picTalk.Visible = False
		picGrid.Visible = False
		'UPGRADE_WARNING: Screen property Screen.MousePointer has a new behavior. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6BA9B8D2-2A32-4B6E-8D36-44949974A5B4"'
		System.Windows.Forms.Cursor.Current = System.Windows.Forms.Cursors.WaitCursor
		'UPGRADE_ISSUE: PictureBox method picMap.Cls was not upgraded. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="CC4C7EC0-C903-48FC-ACCC-81861D12DA4A"'
        picMap.Controls.Clear()
		'UPGRADE_ISSUE: PictureBox method picWallPaper.Cls was not upgraded. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="CC4C7EC0-C903-48FC-ACCC-81861D12DA4A"'
        picWallPaper.Dispose()
        picWallPaper = Nothing

		If PictureFile <> "" And Style < 5 Then
			'UPGRADE_WARNING: Couldn't resolve default property of object Tome.FullPath. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
			'UPGRADE_WARNING: Dir has a new behavior. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="9B7D5ADD-D8FE-4819-A36C-6DEDAF088CC7"'
            FileName = Dir(tome.FullPath & "\" & PictureFile)
			If FileName <> "" Then
				'UPGRADE_WARNING: Couldn't resolve default property of object Tome.FullPath. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
                picWallPaper.Image = System.Drawing.Image.FromFile(tome.FullPath & "\" & FileName)
			Else
				Style = 5 ' Couldn't load picture, so default to no picture
			End If
		End If
		Select Case Style
			Case 0 ' Picture Left, Text Right
				PosY = ShowText(picMap, picWallPaper.Width + 12 + (picBox.Width - 608) / 2, 0, Greatest(608 - picWallPaper.Width - 18, 1), picBox.Height, bdFontNoxiousWhite, Text_Renamed, False, True)
				'UPGRADE_ISSUE: PictureBox method picMap.Cls was not upgraded. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="CC4C7EC0-C903-48FC-ACCC-81861D12DA4A"'
                picMap.Controls.Clear()
				picWallPaper.Top = (picBox.Height - Greatest(PosY, (picWallPaper.Height))) / 2
				picWallPaper.Left = 6 + (picBox.Width - 608) / 2
				picWallPaper.Visible = True
				ShowText(picMap, picWallPaper.Width + 12 + (picBox.Width - 608) / 2, (picBox.Height - Greatest(PosY, (picWallPaper.Height))) / 2, Greatest(608 - picWallPaper.Width - 18, 1), picBox.Height, bdFontNoxiousWhite, Text_Renamed, False, False)
			Case 1 ' Picture Right, Text Left
				PosY = ShowText(picMap, 6 + (picBox.Width - 608) / 2, 0, Greatest(608 - picWallPaper.Width - 18, 1), picBox.Height, bdFontNoxiousWhite, Text_Renamed, False, True)
				'UPGRADE_ISSUE: PictureBox method picMap.Cls was not upgraded. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="CC4C7EC0-C903-48FC-ACCC-81861D12DA4A"'
                picMap.Controls.Clear()
				picWallPaper.Top = (picBox.Height - Greatest(PosY, (picWallPaper.Height))) / 2
				picWallPaper.Left = (picBox.Width - 608) / 2 + 12 + Greatest(608 - picWallPaper.Width - 18, 1)
				picWallPaper.Visible = True
				ShowText(picMap, 6 + (picBox.Width - 608) / 2, (picBox.Height - Greatest(PosY, (picWallPaper.Height))) / 2, Greatest(608 - picWallPaper.Width - 18, 1), picBox.Height, bdFontNoxiousWhite, Text_Renamed, False, False)
			Case 2 ' Picture Top, Text Bottom
				PosY = ShowText(picMap, (picBox.Width - 608) / 2, 0, 608, picBox.Height, bdFontNoxiousWhite, Text_Renamed, True, True)
				'UPGRADE_ISSUE: PictureBox method picMap.Cls was not upgraded. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="CC4C7EC0-C903-48FC-ACCC-81861D12DA4A"'
                picMap.Controls.Clear()
				picWallPaper.Top = (picBox.Height - PosY - picWallPaper.Height - 6) / 2
				picWallPaper.Left = (picBox.Width - picWallPaper.Width) / 2
				picWallPaper.Visible = True
				ShowText(picMap, (picBox.Width - 608) / 2, picWallPaper.Top + picWallPaper.Height + 6, 608, picBox.Height, bdFontNoxiousWhite, Text_Renamed, False, True)
			Case 3 ' Picture Bottom, Text Top
				PosY = ShowText(picMap, (picBox.Width - 608) / 2, 0, 608, picBox.Height, bdFontNoxiousWhite, Text_Renamed, True, True)
				'UPGRADE_ISSUE: PictureBox method picMap.Cls was not upgraded. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="CC4C7EC0-C903-48FC-ACCC-81861D12DA4A"'
                picMap.Controls.Clear()
				picWallPaper.Top = (picBox.Height - PosY - picWallPaper.Height - 6) / 2 + PosY + 6
				picWallPaper.Left = (picBox.Width - picWallPaper.Width) / 2
				picWallPaper.Visible = True
				ShowText(picMap, (picBox.Width - 608) / 2, (picBox.Height - PosY - picWallPaper.Height - 6) / 2, 608, picBox.Height, bdFontNoxiousWhite, Text_Renamed, False, True)
			Case 4 ' Picture Center
				picWallPaper.Left = (picBox.Width - picWallPaper.Width) / 2
				picWallPaper.Top = (picBox.Height - picWallPaper.Height) / 2
				picWallPaper.Visible = True
			Case 5 ' Text Center
				PosY = ShowText(picMap, (picBox.Width - 608) / 2, 0, 608, picBox.Height, bdFontNoxiousWhite, Text_Renamed, True, True)
				'UPGRADE_ISSUE: PictureBox method picMap.Cls was not upgraded. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="CC4C7EC0-C903-48FC-ACCC-81861D12DA4A"'
                picMap.Controls.Clear()
				ShowText(picMap, (picBox.Width - 608) / 2, (picBox.Height - PosY) / 2, 608, picBox.Height, bdFontNoxiousWhite, Text_Renamed, False, True)
		End Select
		picMap.Refresh()
		MessageShow("Click to Continue", 0)
		'UPGRADE_WARNING: Screen property Screen.MousePointer has a new behavior. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6BA9B8D2-2A32-4B6E-8D36-44949974A5B4"'
		System.Windows.Forms.Cursor.Current = System.Windows.Forms.Cursors.Default
		CutSceneNow = True
		Do Until CutSceneNow = False
			System.Windows.Forms.Application.DoEvents()
		Loop 
		picWallPaper.Visible = False
		' Put everything back
		Frozen = OldFrozen
		picInventory.Visible = OldInventory
		picConvo.Visible = OldConvo
		picTalk.Visible = OldTalk
		picGrid.Visible = OldGrid
		If picGrid.Visible = True Then
			CombatWallpaper()
			CombatDraw()
		End If
	End Sub
	
	Public Sub CutSceneHide()
		picWallPaper.Visible = False
		MessageClear()
		' Redraw Map
		DrawMapAll()
		MoveParty()
		picMap.Refresh()
		Me.Refresh()
		ActionNow = 0
	End Sub
	
	'UPGRADE_NOTE: Map was upgraded to Map_Renamed. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="A9E4979A-37FA-4718-9994-97DD76ED70A7"'
	Private Sub MapCenter(ByRef XMap As Short, ByRef YMap As Short)
		Dim Map_Renamed As Object
		'UPGRADE_WARNING: Couldn't resolve default property of object Map.Left. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
        MainMap.Left = XMap + 2
		'UPGRADE_WARNING: Couldn't resolve default property of object Map.Top. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
        MainMap.Top = YMap - Int(System.Math.Sqrt((picMap.ClientRectangle.Width ^ 2) + (picMap.ClientRectangle.Height ^ 2)) / 108) + 1
	End Sub
	
	'UPGRADE_NOTE: Map was upgraded to Map_Renamed. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="A9E4979A-37FA-4718-9994-97DD76ED70A7"'
	'UPGRADE_NOTE: Tome was upgraded to Tome_Renamed. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="A9E4979A-37FA-4718-9994-97DD76ED70A7"'
	Private Sub MovePartySet(ByRef AtClickX As Short, ByRef AtClickY As Short, Optional ByRef AtMapX As Object = Nothing, Optional ByRef AtMapY As Object = Nothing)
		Dim Tome_Renamed As Object
		Dim Map_Renamed As Object
		Dim Y, X, c As Short
		Dim MoveToX, MoveToY As Short
		Dim CreatureX As Creature
		'UPGRADE_NOTE: IsMissing() was changed to IsNothing(). Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="8AE1CB93-37AB-439A-A4FF-BE3B6760BB23"'
		If IsNothing(AtMapX) Or IsNothing(AtMapY) Then
			If Int(AtClickY / 16) Mod 3 = 1 Or (Int(AtClickY / 8) Mod 3 = 1 And Int((AtClickX - 24) / 48) Mod 2 = 0) Then
				'UPGRADE_WARNING: Couldn't resolve default property of object Map.Left. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
                MoveToX = MainMap.Left + Int(AtClickX / 96) - Int(AtClickY / 48)
				'UPGRADE_WARNING: Couldn't resolve default property of object Map.Top. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
                MoveToY = MainMap.Top + Int(AtClickX / 96) + Int(AtClickY / 48)
			Else
				'UPGRADE_WARNING: Couldn't resolve default property of object Map.Left. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
                MoveToX = MainMap.Left + Int((AtClickX - 48) / 96) - Int((AtClickY - 24) / 48)
				'UPGRADE_WARNING: Couldn't resolve default property of object Map.Top. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
                MoveToY = MainMap.Top + Int((AtClickX + 48) / 96) + Int((AtClickY - 24) / 48)
			End If
		Else
			'UPGRADE_WARNING: Couldn't resolve default property of object AtMapX. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
			MoveToX = CShort(AtMapX)
			'UPGRADE_WARNING: Couldn't resolve default property of object AtMapY. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
			MoveToY = CShort(AtMapY)
		End If
		'UPGRADE_WARNING: Couldn't resolve default property of object Tome.MoveToX. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
        tome.MoveToX = MoveToX
		'UPGRADE_WARNING: Couldn't resolve default property of object Tome.MoveToY. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
        tome.MoveToY = MoveToY
		'UPGRADE_WARNING: Couldn't resolve default property of object Tome. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
        FindPath(tome)
		' Play a sound (maybe)
		'    If GlobalWAVState = 1 Then
		'        PlaySound "step" & Int(Rnd * 4) + 1, False
		'    End If
		c = -1
		'UPGRADE_WARNING: Couldn't resolve default property of object Tome.Creatures. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
        For Each CreatureX In tome.Creatures
            c = LoopNumber(0, 8, c, 1)
            CreatureX.TileSpot = c
        Next CreatureX
		tmrMoveParty.Enabled = True
		If GlobalAutoCenter Then ' [Titi 2.4.9]
			MapCenter(MoveToX, MoveToY) ' Ephestion [2.4.9]
			DrawMapAll() ' Ephestion [2.4.9]
		End If
	End Sub
	
	Public Function CombatApplyDamage(ByRef Target As Creature, ByRef Damage As Short) As Short
		Dim c, NoFail As Short
		Dim ItemX As Item
		Dim CreatureX As Creature
		' If already dead, this does nothing more
		If Target.HPNow < 1 Then
			CombatApplyDamage = 0
			Exit Function
		End If
		' Fire Pre-ApplyDamage
		GlobalDamageRoll = Damage
		CreatureNow = Target
		ItemNow = Target.ItemInHand
		NoFail = FireTriggers(Target, bdPreApplyDamage)
		For	Each ItemNow In Target.Items
			If ItemNow.IsReady = True Then
				CreatureNow = Target
				NoFail = FireTriggers(ItemNow, bdPreApplyDamage)
			End If
		Next ItemNow
		' Apply Damage
		Target.HPNow = Target.HPNow - GlobalDamageRoll
		' Fire Post-ApplyDamage
		CreatureNow = Target
		ItemNow = Target.ItemInHand
        NoFail = FireTriggers(tome, bdPostApplyDamage)
		CreatureNow = Target
		ItemNow = Target.ItemInHand
		NoFail = FireTriggers(Target, bdPostApplyDamage)
		For	Each ItemNow In Target.Items
			If ItemNow.IsReady = True Then
				CreatureNow = Target
				NoFail = FireTriggers(ItemNow, bdPostApplyDamage)
			End If
		Next ItemNow
		' Determine if killed
		If Target.HPNow < 1 Then
			' Fire Pre-Death
			CreatureNow = Target
			ItemNow = Target.ItemInHand
			NoFail = FireTriggers(Target, bdPreDeath)
			For	Each ItemNow In Target.Items
				If ItemNow.IsReady = True Then
					CreatureNow = Target
					NoFail = FireTriggers(ItemNow, bdPreDeath)
				End If
			Next ItemNow
		End If
		' If still dead, then play dead sound
		If Target.HPNow < 1 Then
			If Target.DieWAV = "" Then
				If Target.Friendly = False Then
                    Call PlaySoundFile("MonsDie", tome, , 5)
				Else
					If Target.Male = True Then
                        Call PlaySoundFile("PCDie", tome, , 5)
					Else
                        Call PlaySoundFile("PCDie2", tome, , 5)
					End If
				End If
			Else
                Call PlaySoundFile(Target.DieWAV, tome, True)
				If Target.DieWAVOneTime = True Then
					Target.DieWAV = ""
					Target.DieWAVOneTime = False
				End If
			End If
			' Fire Post-Death
			CreatureNow = Target
			ItemNow = Target.ItemInHand
			NoFail = FireTriggers(Target, bdPostDeath)
			For	Each ItemNow In Target.Items
				If ItemNow.IsReady = True Then
					CreatureNow = Target
					NoFail = FireTriggers(ItemNow, bdPostDeath)
				End If
			Next ItemNow
			' Redraw Combat screen (if necessary)
			If picGrid.Visible = True Then
				CombatDraw()
			Else
				MenuDrawParty()
			End If
		Else
			If picGrid.Visible = True Then
				CombatDrawAttack(CreatureWithTurn, CreatureTarget, True)
				picToHit.Refresh()
				BorderDrawButtons(0)
				Me.Refresh()
			Else
				MenuDrawParty()
			End If
		End If
	End Function
	
	Private Sub CombatDrawAttack(ByRef FromCreature As Creature, ByRef ToCreature As Creature, ByRef ShowDice As Short)
		Dim c, HasRunes As Short
		Dim rc As Integer
		'UPGRADE_NOTE: Width was upgraded to Width_Renamed. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="A9E4979A-37FA-4718-9994-97DD76ED70A7"'
		Dim Y, X, Width_Renamed As Short
		' Redraw Bottom
		'UPGRADE_ISSUE: PictureBox method picToHit.Cls was not upgraded. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="CC4C7EC0-C903-48FC-ACCC-81861D12DA4A"'
        picToHit.Image.Dispose()
        picToHit.Image = Nothing

		BorderDrawBottom()
		' Location of Text
		Y = Me.ClientRectangle.Height - 14 : X = picToHit.Left : Width_Renamed = picToHit.ClientRectangle.Width
		' Draw the appropriate text
		If FromCreature.Friendly = True Then
			MenuDrawFace(FromCreature, 5)
			If FromCreature.ItemInHand.Name <> "Hand" Then
				ShowText(Me, X, Y, Width_Renamed, 14, bdFontNoxiousWhite, FromCreature.Name & " * " & FromCreature.ItemInHand.Name, False, False)
			Else
				ShowText(Me, X, Y, Width_Renamed, 14, bdFontNoxiousWhite, (FromCreature.Name), False, False)
			End If
		Else
			MenuDrawFace(FromCreature, 6)
			If FromCreature.ItemInHand.Name <> "Hand" Then
				ShowText(Me, X, Y, Width_Renamed, 14, bdFontNoxiousWhite, FromCreature.Name & " * " & FromCreature.ItemInHand.Name, 1, False)
			Else
				ShowText(Me, X, Y, Width_Renamed, 14, bdFontNoxiousWhite, (FromCreature.Name), 1, False)
			End If
		End If
		If FromCreature.Name <> ToCreature.Name Or FromCreature.Index <> ToCreature.Index Then
			' Show Defense Number
			ShowText(picToHit, 128, 8, 46, 28, bdFontLargeWhite, VB6.Format(ToCreature.Defense, "##0"), True, False)
			' Show Creature Face and ItemInHand
			If FromCreature.Friendly = True Then
				MenuDrawFace(ToCreature, 6)
				If ToCreature.ItemInHand.Name <> "Hand" Then
					ShowText(Me, X, Y, Width_Renamed, 14, bdFontNoxiousWhite, ToCreature.Name & " * " & ToCreature.ItemInHand.Name, 1, False)
				Else
					ShowText(Me, X, Y, Width_Renamed, 14, bdFontNoxiousWhite, (ToCreature.Name), 1, False)
				End If
			Else
				MenuDrawFace(ToCreature, 5)
				If ToCreature.ItemInHand.Name <> "Hand" Then
					ShowText(Me, X, Y, Width_Renamed, 14, bdFontNoxiousWhite, ToCreature.Name & " * " & ToCreature.ItemInHand.Name, False, False)
				Else
					ShowText(Me, X, Y, Width_Renamed, 14, bdFontNoxiousWhite, (ToCreature.Name), False, False)
				End If
			End If
			' ShowDice (if needed)
			If ShowDice = True Then
				For c = 0 To 5
					If CombatDiceType(c) <> 0 Then
						DiceDraw(c, CombatDiceType(c), CombatDiceValue(c), True)
					End If
				Next c
				If Combat20Bonus < 0 Then
					ShowText(picToHit, 132, 40, 40, 10, bdFontSmallWhite, VB6.Format(Combat20Bonus), 1, False)
				ElseIf Combat20Bonus > 0 Then 
					ShowText(picToHit, 132, 40, 40, 10, bdFontSmallWhite, "+" & VB6.Format(Combat20Bonus), 1, False)
				End If
				If CombatDmgBonus < 0 Then
					ShowText(picToHit, 136 + 1 * 46, 40, 40, 10, bdFontSmallWhite, VB6.Format(CombatDmgBonus), 1, False)
				ElseIf CombatDmgBonus > 0 Then 
					ShowText(picToHit, 136 + 1 * 46, 40, 40, 10, bdFontSmallWhite, "+" & VB6.Format(CombatDmgBonus), 1, False)
				End If
				' Show Hit Location
			End If
			' Show current Armor Chits
			c = CombatRollArmor(ToCreature, False, 0)
		End If
		picToHit.Refresh()
		BorderDrawButtons(0)
		Me.Refresh()
	End Sub
	
	'UPGRADE_NOTE: Tome was upgraded to Tome_Renamed. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="A9E4979A-37FA-4718-9994-97DD76ED70A7"'
	Public Sub MenuDrawParty()
		Dim Tome_Renamed As Object
		Dim c, i As Short
		Dim rc As Integer
		' Clear the faces
		For c = 0 To 5
            picMenu.Image = CType(CopyRect(picMenu, New RectangleF((c * 122), 0, 32, 84)), System.Drawing.Bitmap).Clone
            If c < 5 Then
                'UPGRADE_ISSUE: PictureBox property picMisc.hdc was not upgraded. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="CC4C7EC0-C903-48FC-ACCC-81861D12DA4A"'
                'UPGRADE_ISSUE: PictureBox property picMenu.hdc was not upgraded. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="CC4C7EC0-C903-48FC-ACCC-81861D12DA4A"'
                rc = BitBlt(picMenu.hdc, 32 + c * 122, 0, 90, 84, picMisc.hdc, 0, 36, SRCCOPY)
            End If
        Next c
		' Draw faces
		For c = 0 To 4
			'UPGRADE_WARNING: Couldn't resolve default property of object Tome.Creatures. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
            If PartyLeft + c < tome.Creatures.Count Then
                'UPGRADE_WARNING: Couldn't resolve default property of object Tome.Creatures. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
                MenuParty(c) = tome.Creatures(PartyLeft + c + 1)
                If picGrid.Visible = False Then
                    MenuDrawFace(MenuParty(c), c)
                End If
            Else
                'UPGRADE_ISSUE: PictureBox property picFaces.hdc was not upgraded. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="CC4C7EC0-C903-48FC-ACCC-81861D12DA4A"'
                'UPGRADE_ISSUE: PictureBox property picMenu.hdc was not upgraded. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="CC4C7EC0-C903-48FC-ACCC-81861D12DA4A"'
                rc = BitBlt(picMenu.hdc, 36 + c * 122, 4, 66, 76, picFaces.hdc, 66, 0, SRCAND)
                'UPGRADE_ISSUE: PictureBox property picFaces.hdc was not upgraded. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="CC4C7EC0-C903-48FC-ACCC-81861D12DA4A"'
                'UPGRADE_ISSUE: PictureBox property picMenu.hdc was not upgraded. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="CC4C7EC0-C903-48FC-ACCC-81861D12DA4A"'
                rc = BitBlt(picMenu.hdc, 36 + c * 122, 4, 66, 76, picFaces.hdc, 0, 0, SRCPAINT)
            End If
		Next c
		' Draw Scroll Arrows
		If PartyLeft > 0 Then
			' Left Arrow
			'UPGRADE_ISSUE: PictureBox property picMisc.hdc was not upgraded. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="CC4C7EC0-C903-48FC-ACCC-81861D12DA4A"'
			'UPGRADE_ISSUE: PictureBox property picMenu.hdc was not upgraded. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="CC4C7EC0-C903-48FC-ACCC-81861D12DA4A"'
			rc = BitBlt(picMenu.hdc, 8, 32, 18, 18, picMisc.hdc, 144, 18, SRCCOPY)
		End If
		'UPGRADE_WARNING: Couldn't resolve default property of object Tome.Creatures. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
        If PartyLeft + 5 < tome.Creatures.Count Then
            ' Right Arrow
            'UPGRADE_ISSUE: PictureBox property picMisc.hdc was not upgraded. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="CC4C7EC0-C903-48FC-ACCC-81861D12DA4A"'
            'UPGRADE_ISSUE: PictureBox property picMenu.hdc was not upgraded. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="CC4C7EC0-C903-48FC-ACCC-81861D12DA4A"'
            rc = BitBlt(picMenu.hdc, 618, 32, 18, 18, picMisc.hdc, 108, 18, SRCCOPY)
        End If
		' Draw select box
		'UPGRADE_ISSUE: PictureBox property picFaces.hdc was not upgraded. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="CC4C7EC0-C903-48FC-ACCC-81861D12DA4A"'
		'UPGRADE_ISSUE: PictureBox property picMenu.hdc was not upgraded. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="CC4C7EC0-C903-48FC-ACCC-81861D12DA4A"'
		rc = BitBlt(picMenu.hdc, 36 + MenuPartyIndex * 122, 4, 66, 76, picFaces.hdc, bdFaceSelect + 66, 0, SRCAND)
		'UPGRADE_ISSUE: PictureBox property picFaces.hdc was not upgraded. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="CC4C7EC0-C903-48FC-ACCC-81861D12DA4A"'
		'UPGRADE_ISSUE: PictureBox property picMenu.hdc was not upgraded. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="CC4C7EC0-C903-48FC-ACCC-81861D12DA4A"'
		rc = BitBlt(picMenu.hdc, 36 + MenuPartyIndex * 122, 4, 66, 76, picFaces.hdc, bdFaceSelect, 0, SRCPAINT)
		' Redraw buttons
		BorderDrawButtons(0)
		Me.Refresh()
		picMenu.Refresh()
	End Sub
	
	'UPGRADE_NOTE: Loc was upgraded to Loc_Renamed. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="A9E4979A-37FA-4718-9994-97DD76ED70A7"'
	Private Sub MenuDrawFace(ByRef CreatureX As Creature, ByRef Loc_Renamed As Short)
		Dim rc As Short
		Dim TriggerX As Trigger
		' Draw face
		LoadCreaturePic(CreatureX)
		If picGrid.Visible = True Then
			' In combat
			'UPGRADE_ISSUE: PictureBox property picFaces.hdc was not upgraded. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="CC4C7EC0-C903-48FC-ACCC-81861D12DA4A"'
			'UPGRADE_ISSUE: PictureBox property picToHit.hdc was not upgraded. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="CC4C7EC0-C903-48FC-ACCC-81861D12DA4A"'
			rc = BitBlt(picToHit.hdc, 36 + 388 * System.Math.Abs(CInt(Loc_Renamed = 6)), 4, 66, 76, picFaces.hdc, bdFaceMin + CreatureX.Pic * 66, 0, SRCCOPY)
			MenuDrawStats(CreatureX, Loc_Renamed)
		Else
			' On Menu
			'UPGRADE_ISSUE: PictureBox property picFaces.hdc was not upgraded. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="CC4C7EC0-C903-48FC-ACCC-81861D12DA4A"'
			'UPGRADE_ISSUE: PictureBox property picMenu.hdc was not upgraded. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="CC4C7EC0-C903-48FC-ACCC-81861D12DA4A"'
			rc = BitBlt(picMenu.hdc, 36 + Loc_Renamed * 122, 4, 66, 76, picFaces.hdc, bdFaceMin + CreatureX.Pic * 66, 0, SRCCOPY)
			If CreatureX.AllowedTurn = False Then
				'UPGRADE_ISSUE: PictureBox property picFaces.hdc was not upgraded. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="CC4C7EC0-C903-48FC-ACCC-81861D12DA4A"'
				'UPGRADE_ISSUE: PictureBox property picMenu.hdc was not upgraded. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="CC4C7EC0-C903-48FC-ACCC-81861D12DA4A"'
				rc = BitBlt(picMenu.hdc, 36 + Loc_Renamed * 122, 4, 66, 76, picFaces.hdc, 66, 0, SRCAND)
				'UPGRADE_ISSUE: PictureBox property picFaces.hdc was not upgraded. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="CC4C7EC0-C903-48FC-ACCC-81861D12DA4A"'
				'UPGRADE_ISSUE: PictureBox property picMenu.hdc was not upgraded. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="CC4C7EC0-C903-48FC-ACCC-81861D12DA4A"'
				rc = BitBlt(picMenu.hdc, 36 + Loc_Renamed * 122, 4, 66, 76, picFaces.hdc, 0, 0, SRCPAINT)
			End If
			MenuDrawStats(CreatureX, Loc_Renamed)
		End If
	End Sub
	
	'UPGRADE_NOTE: Loc was upgraded to Loc_Renamed. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="A9E4979A-37FA-4718-9994-97DD76ED70A7"'
	Private Sub MenuDrawStats(ByRef CreatureX As Creature, ByRef Loc_Renamed As Short)
		Dim c As Short
		Dim rc As Integer
		'UPGRADE_NOTE: Height was upgraded to Height_Renamed. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="A9E4979A-37FA-4718-9994-97DD76ED70A7"'
		Dim PicX As System.Windows.Forms.PictureBox
		Dim PosX, Height_Renamed As Short
		If picGrid.Visible = True Then
			PicX = picToHit
			PosX = 106 + 388 * System.Math.Abs(CInt(Loc_Renamed = 6))
		Else
			PicX = picMenu
			PosX = 106 + Loc_Renamed * 122
		End If
		' HitPoints
		If CreatureX.HPMax > 0 And CreatureX.HPNow > 0 Then
			Height_Renamed = 84 * (CreatureX.HPNow / CreatureX.HPMax)
			Select Case (CreatureX.HPNow / CreatureX.HPMax) * 100
				Case 0 To 50
					c = 32
				Case Else
					c = 16
			End Select
		Else
			c = 0
			Height_Renamed = 0
		End If
		'UPGRADE_ISSUE: PictureBox property picMisc.hdc was not upgraded. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="CC4C7EC0-C903-48FC-ACCC-81861D12DA4A"'
		'UPGRADE_ISSUE: PictureBox property PicX.hdc was not upgraded. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="CC4C7EC0-C903-48FC-ACCC-81861D12DA4A"'
		rc = BitBlt(PicX.hdc, PosX, 0, 16, 84 - Height_Renamed, picMisc.hdc, 90, 36, SRCCOPY)
		'UPGRADE_ISSUE: PictureBox property picMisc.hdc was not upgraded. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="CC4C7EC0-C903-48FC-ACCC-81861D12DA4A"'
		'UPGRADE_ISSUE: PictureBox property PicX.hdc was not upgraded. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="CC4C7EC0-C903-48FC-ACCC-81861D12DA4A"'
		rc = BitBlt(PicX.hdc, PosX, 0 + 84 - Height_Renamed, 16, Height_Renamed, picMisc.hdc, 90 + c, 36 + (84 - Height_Renamed), SRCCOPY)
	End Sub
	
	Private Sub InventoryContainerClose()
		Dim c As Short
		Dim rc As Integer
		' Close Container
		picContainer.Visible = False
		InvContainer.Selected = False
		' If Inventory open, then redraw
		If picSearch.Visible = True Then
			SearchShow()
		ElseIf picInventory.Visible = True Then 
			InventoryShow(bdInvItems)
		End If
	End Sub
	
	'UPGRADE_NOTE: Tome was upgraded to Tome_Renamed. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="A9E4979A-37FA-4718-9994-97DD76ED70A7"'
	Private Sub InventoryContainerDrag(ByRef AtX As Short, ByRef AtY As Short)
		Dim Tome_Renamed As Object
		Dim X, Found, c, i, Y As Short
		' No matter what, shut off dragging Item
		InvDragItem.Selected = False
		InvDragMode = False
		picInvDrag.Visible = False
		If PointIn(AtX, AtY, (picContainer.Left), (picContainer.Top), (picContainer.ClientRectangle.Width), (picContainer.ClientRectangle.Height)) = True Then
			' Drop back into Container
			For c = 1 To 80
				X = picContainer.Left + 23 + ((c - 1) Mod 8) * 34
				Y = picContainer.Top + 45 + Int((c - 1) / 8) * 34
				If PointIn(AtX, AtY, X, Y, 34, 34) Then
					TradeItem(CreatureWithTurn, CreatureWithTurn, bdInvContainer, bdInvContainer, InvC(c), c, InvDragItem)
					InventoryContainerShow(InvContainer)
					Exit For
				End If
			Next c
		ElseIf picInventory.Visible = True And PointIn(AtX, AtY, (picInventory.Left), (picInventory.Top), (picInventory.ClientRectangle.Width), (picInventory.ClientRectangle.Height)) = True Then 
			' Drop onto same creature
			TradeItem(CreatureWithTurn, CreatureWithTurn, bdInvContainer, bdInvObjects, 0, 0, InvDragItem)
			InventoryShow(bdInvItems)
		ElseIf AtY > picMenu.Top Then 
			' Drop to Party member
			c = Greatest(0, Least(Int((AtX - picMenu.Left) / 122), 4))
			'UPGRADE_WARNING: Couldn't resolve default property of object Tome.Creatures. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
            If IsBetween(c, 0, tome.Creatures.Count - 1) Then
                TradeItem(CreatureWithTurn, MenuParty(c), bdInvContainer, bdInvObjects, 0, 0, InvDragItem)
            End If
		Else
			TradeItem(CreatureWithTurn, CreatureWithTurn, bdInvContainer, bdInvEncounter, 0, 0, InvDragItem)
		End If
	End Sub
	
	Private Sub InventoryContainerClick(ByRef AtX As Short, ByRef AtY As Short, ByRef ButtonDown As Short, ByRef Button As Short)
		Dim c, i As Short
		Dim rc As Integer
		Dim ItemX As Item
		'UPGRADE_NOTE: Height was upgraded to Height_Renamed. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="A9E4979A-37FA-4718-9994-97DD76ED70A7"'
		'UPGRADE_NOTE: Width was upgraded to Width_Renamed. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="A9E4979A-37FA-4718-9994-97DD76ED70A7"'
		Dim Width_Renamed, Height_Renamed As Short
		If PointIn(AtX, AtY, picContainer.ClientRectangle.Width - 198, picContainer.ClientRectangle.Height - 30, 90, 18) Then
			' Empty
			ShowButton(picContainer, picContainer.ClientRectangle.Width - 198, picContainer.ClientRectangle.Height - 30, "Empty", ButtonDown)
			picContainer.Refresh()
			If ButtonDown = False Then
				Call PlayClickSnd(modIOFunc.ClickType.ifClick)
				If picSearch.Visible = True Then
					For	Each ItemX In InvContainer.Items
						TradeItem(CreatureWithTurn, CreatureWithTurn, bdInvContainer, bdInvEncounter, 0, 0, ItemX)
					Next ItemX
				ElseIf picInventory.Visible = True Then 
					For	Each ItemX In InvContainer.Items
						TradeItem(CreatureWithTurn, CreatureWithTurn, bdInvContainer, bdInvObjects, 0, 0, ItemX)
					Next ItemX
				End If
				InventoryContainerShow(InvContainer)
				If InvContainer.Items.Count() < 1 Then
					InventoryContainerClose()
				End If
			End If
		ElseIf PointIn(AtX, AtY, picContainer.ClientRectangle.Width - 102, picContainer.ClientRectangle.Height - 30, 90, 18) Then 
			' Done
			ShowButton(picContainer, picContainer.ClientRectangle.Width - 102, picContainer.ClientRectangle.Height - 30, "Done", ButtonDown)
			picContainer.Refresh()
			If ButtonDown = False Then
				Call PlayClickSnd(modIOFunc.ClickType.ifClick)
				InventoryContainerClose()
			End If
		ElseIf PointIn(AtX, AtY, picSearch.ClientRectangle.Width - 294, picSearch.ClientRectangle.Height - 30, 90, 18) And InvTooMany = True Then 
			' More
			ShowButton(picSearch, picSearch.ClientRectangle.Width - 294, picSearch.ClientRectangle.Height - 30, "More", ButtonDown)
			picSearch.Refresh()
			If ButtonDown = False Then
				Call PlayClickSnd(modIOFunc.ClickType.ifClick)
				' Shuffle the Items
				For	Each ItemX In InvContainer.Items
					If ItemX.InvSpot > 0 Then
						InvContainer.RemoveItem("I" & ItemX.Index)
						ItemX.InvSpot = 0
						InvContainer.AddItem.Copy(ItemX)
					End If
				Next ItemX
				InventoryContainerShow(InvContainer)
			End If
		ElseIf ButtonDown Then 
			For c = 1 To 80
				If PointIn(AtX, AtY, 23 + ((c - 1) Mod 8) * 34, 45 + Int((c - 1) / 8) * 34, 34, 34) And InvC(c) > 0 Then
					ItemX = InvContainer.Items.Item("I" & InvC(c))
					If Button = VB6.MouseButtonConstants.RightButton Then
						Call PlayClickSnd(modIOFunc.ClickType.ifClick)
						ExamineItem(ItemX)
						InventoryContainerShow(InvContainer)
					Else
						'UPGRADE_ISSUE: PictureBox method picInvDrag.Cls was not upgraded. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="CC4C7EC0-C903-48FC-ACCC-81861D12DA4A"'
						picInvDrag.Cls()
						LoadItemPic(ItemX)
						Select Case ItemPicWidth(ItemX.Pic)
							Case 32 : Width_Renamed = 34
							Case 64 : Width_Renamed = 68
						End Select
						Select Case ItemPicHeight(ItemX.Pic)
							Case 32 : Height_Renamed = 34
							Case 64 : Height_Renamed = 68
							Case 96 : Height_Renamed = 102
						End Select
						For i = 1 To 80
							If InvC(i) = ItemX.InvSpot Then InvC(i) = 0
						Next i
						picInvDrag.Width = Width_Renamed : picInvDrag.Height = Height_Renamed
						'UPGRADE_ISSUE: PictureBox property picContainer.hdc was not upgraded. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="CC4C7EC0-C903-48FC-ACCC-81861D12DA4A"'
						'UPGRADE_ISSUE: PictureBox property picInvDrag.hdc was not upgraded. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="CC4C7EC0-C903-48FC-ACCC-81861D12DA4A"'
						rc = BitBlt(picInvDrag.hdc, 0, 0, Width_Renamed, Height_Renamed, picContainer.hdc, 22 + ((ItemX.InvSpot - 1) Mod 8) * 34, 44 + Int((ItemX.InvSpot - 1) / 8) * 34, SRCCOPY)
						picInvDrag.Refresh()
						' Ready for Drag/Drop
						JournalX = AtX - 23 - ((ItemX.InvSpot - 1) Mod 8) * 34 : JournalY = AtY - 45 - Int((ItemX.InvSpot - 1) / 8) * 34
						InvDragItem = ItemX
						InvDragPos = c
						InvDragIndex = 0
						InvDragMode = True
						InvDragItem.Selected = True
					End If
					Exit For
				End If
			Next c
		End If
	End Sub
	
	'UPGRADE_NOTE: Tome was upgraded to Tome_Renamed. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="A9E4979A-37FA-4718-9994-97DD76ED70A7"'
	Private Sub InventoryDragItem(ByRef AtX As Short, ByRef AtY As Short)
		Dim Tome_Renamed As Object
		Dim ToList, Found, c, FromList, i As Short
		' No matter what, shut off dragging Item
		InvDragItem.Selected = False
		InvDragItem.InvSpot = 0
		InvDragMode = False
		picInvDrag.Visible = False
		' Drop where appropriate
		If PointIn(AtX, AtY, (picContainer.Left), (picContainer.Top), (picContainer.Width), (picContainer.Height)) And picContainer.Visible = True Then
			' Drop on Container
			TradeItem(CreatureWithTurn, CreatureWithTurn, bdInvObjects, bdInvContainer, 0, 0, InvDragItem)
			InventoryContainerShow(InvContainer)
		ElseIf PointIn(AtX, AtY, (picInventory.Left), (picInventory.Top), (picInventory.Width), (picInventory.Height)) And picSearch.Visible = False Then 
			' Drop on Inventory
			For c = 1 To 128
				If PointIn(AtX, AtY, picInventory.Left + InvX(c), picInventory.Top + InvY(c), 34, 34) Then
					If c < 49 Then
						ToList = bdInvWear
					Else
						ToList = bdInvObjects
					End If
					If InvDragItem.IsReady = True Then
						FromList = bdInvWear
					Else
						FromList = bdInvObjects
					End If
					TradeItem(CreatureWithTurn, CreatureWithTurn, FromList, ToList, InvZ(c), c, InvDragItem)
					InventoryShow(bdInvItems)
					Exit For
				End If
			Next c
		ElseIf AtY > picMenu.Top - 32 Then 
			' Drop to Party member
			c = Greatest(0, Least(Int((AtX - picMenu.Left) / 122), 4))
			'UPGRADE_WARNING: Couldn't resolve default property of object Tome.Creatures. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
            If IsBetween(c, 0, tome.Creatures.Count - 1) Then
                TradeItem(CreatureWithTurn, MenuParty(c), bdInvObjects, bdInvParty, 0, 0, InvDragItem)
            End If
		Else
			DropItem(CreatureWithTurn, 0, 0, InvDragItem)
			DrawMapAll()
		End If
	End Sub
	
	Public Sub InventoryClose()
		' Only release map if not in combat
		If picTomeNew.Visible = False Then
			Frozen = False
			picMap.Focus()
		End If
		' If Searching or a Container is open, close it
		If picSearch.Visible = True Then
			picSearch.Visible = False
		End If
		' Shut off Inventory
		picInventory.Visible = False
	End Sub
	
	Private Sub InventoryClick(ByRef AtX As Short, ByRef AtY As Short, ByRef ButtonDown As Short, ByRef Button As Short)
		Dim c, i As Short
		Dim rc As Integer
		Dim ItemX As Item
		'UPGRADE_NOTE: Height was upgraded to Height_Renamed. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="A9E4979A-37FA-4718-9994-97DD76ED70A7"'
		'UPGRADE_NOTE: Width was upgraded to Width_Renamed. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="A9E4979A-37FA-4718-9994-97DD76ED70A7"'
		Dim Width_Renamed, Height_Renamed As Short
		If PointIn(AtX, AtY, 482, 406, 90, 18) Then
			' Done
			ShowButton(picInventory, 482, 406, "Done", ButtonDown)
			picInventory.Refresh()
			If ButtonDown = False Then
				Call PlayClickSnd(modIOFunc.ClickType.ifClick)
				InventoryClose()
			End If
		ElseIf PointIn(AtX, AtY, 290, 406, 90, 18) And InvNowShow = bdInvItems And InvTooMany = True Then 
			' More
			ShowButton(picInventory, 290, 406, "More", ButtonDown)
			picInventory.Refresh()
			If ButtonDown = False Then
				Call PlayClickSnd(modIOFunc.ClickType.ifClick)
				For	Each ItemX In CreatureWithTurn.Items
					If ItemX.InvSpot > 0 Then
						CreatureWithTurn.RemoveItem("I" & ItemX.Index)
						ItemX.InvSpot = 0
						CreatureWithTurn.AddItem.Copy(ItemX)
					End If
				Next ItemX
				InventoryShow(bdInvItems)
			End If
		ElseIf PointIn(AtX, AtY, 386, 406, 90, 18) Then 
			' Status/Equip
			If InvNowShow = bdInvItems Then
				ShowButton(picInventory, 386, 406, "Status", ButtonDown)
			Else
				ShowButton(picInventory, 386, 406, "Equip", ButtonDown)
			End If
			picInventory.Refresh()
			If ButtonDown = False Then
				Call PlayClickSnd(modIOFunc.ClickType.ifClick)
				If InvNowShow = bdInvItems Then
					InventoryShow(bdInvStatus)
				Else
					InventoryShow(bdInvItems)
				End If
			End If
		ElseIf PointIn(AtX, AtY, 544, 45, 18, 338) And InvNowShow = bdInvStatus Then 
			' ScrollBar Click
			If ScrollBarClick(AtX, AtY, ButtonDown, picInventory, 544, 45, 338, ScrollTop, ScrollList.Count(), 17) = True Then
				InventoryShow(bdInvStatus)
			End If
		ElseIf ButtonDown And InvNowShow = bdInvItems Then 
			For c = 1 To 128
				If PointIn(AtX, AtY, InvX(c), InvY(c), 34, 34) And InvZ(c) > -1 Then
					ItemX = CreatureWithTurn.Items.Item("I" & InvZ(c))
					If ItemX.Selected = False Then
						If Button = VB6.MouseButtonConstants.RightButton Then
							Call PlayClickSnd(modIOFunc.ClickType.ifClick)
							ExamineItem(ItemX)
							InventoryShow(bdInvItems)
						Else
							'UPGRADE_ISSUE: PictureBox method picInvDrag.Cls was not upgraded. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="CC4C7EC0-C903-48FC-ACCC-81861D12DA4A"'
							picInvDrag.Cls()
							LoadItemPic(ItemX)
							Select Case ItemPicWidth(ItemX.Pic)
								Case 32 : Width_Renamed = 34
								Case 64 : Width_Renamed = 68
							End Select
							Select Case ItemPicHeight(ItemX.Pic)
								Case 32 : Height_Renamed = 34
								Case 64 : Height_Renamed = 68
								Case 96 : Height_Renamed = 102
							End Select
							For i = 1 To 128
								If InvZ(i) = ItemX.InvSpot Then InvZ(i) = -1
							Next i
							picInvDrag.Width = Width_Renamed : picInvDrag.Height = Height_Renamed
							'UPGRADE_ISSUE: PictureBox property picInventory.hdc was not upgraded. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="CC4C7EC0-C903-48FC-ACCC-81861D12DA4A"'
							'UPGRADE_ISSUE: PictureBox property picInvDrag.hdc was not upgraded. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="CC4C7EC0-C903-48FC-ACCC-81861D12DA4A"'
							rc = BitBlt(picInvDrag.hdc, 0, 0, Width_Renamed, Height_Renamed, picInventory.hdc, InvX(ItemX.InvSpot), InvY(ItemX.InvSpot), SRCCOPY)
							picInvDrag.Refresh()
							' Ready for Drag/Drop
							JournalX = AtX - InvX(ItemX.InvSpot) : JournalY = AtY - InvY(ItemX.InvSpot)
							InvDragItem = ItemX
							InvDragPos = c
							InvDragIndex = 0
							InvDragMode = True
							InvDragItem.Selected = True
						End If
					End If
					Exit For
				End If
			Next c
		End If
	End Sub
	
	Private Function InventoryStatus(ByRef Index As Short) As String
		If Index < 16 Then
			InventoryStatus = "Excellent"
		ElseIf Index > 20 Then 
			InventoryStatus = "Wild"
		Else
			Select Case Index
				Case 16
					InventoryStatus = "Low"
				Case 17
					InventoryStatus = "Mild"
				Case 18
					InventoryStatus = "Average"
				Case 19
					InventoryStatus = "Excessive"
				Case 20
					InventoryStatus = "High"
			End Select
		End If
	End Function
	
	'UPGRADE_NOTE: Width was upgraded to Width_Renamed. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="A9E4979A-37FA-4718-9994-97DD76ED70A7"'
	Private Sub InventoryFindSpot(ByRef ItemX As Item, ByRef AtSpot As Short, ByRef MaxSlot As Short, ByRef Width_Renamed As Short, ByRef Slot() As Short)
		Dim i, c, Found As Short
		' Find an appropriate spot for an Item
		Found = False
		For i = AtSpot To MaxSlot
			If ItemX.InvSpot < 1 Then
				c = i
			Else
				c = ItemX.InvSpot
			End If
			If Slot(c) < 0 Then
				Found = True
				If ItemPicHeight(ItemX.Pic) > 32 Then
					' Check for one spot down
					If c + Width_Renamed > MaxSlot Then
						Found = False
					ElseIf Slot(c + Width_Renamed) > -1 Then 
						Found = False
					End If
				End If
				If ItemPicHeight(ItemX.Pic) > 64 Then
					' Check for two spots down
					If c + Width_Renamed * 2 > MaxSlot Then
						Found = False
					ElseIf Slot(c + Width_Renamed * 2) > -1 Then 
						Found = False
					End If
				End If
				If ItemPicWidth(ItemX.Pic) > 32 Then
					' Check for one spot to the right
					If c + 1 > MaxSlot Or Int((c - 1) / Width_Renamed) <> Int(c / Width_Renamed) Then
						Found = False
					ElseIf Slot(c + 1) > -1 Then 
						Found = False
					End If
				End If
				If Found = True Then
					ItemX.InvSpot = c
					Exit For
				End If
			End If
		Next i
		If Found = True Then
			Slot(c) = ItemX.Index ' Mark current slot
			If ItemPicHeight(ItemX.Pic) > 32 Then Slot(c + Width_Renamed) = ItemX.Index
			If ItemPicHeight(ItemX.Pic) > 64 Then Slot(c + Width_Renamed * 2) = ItemX.Index
			If ItemPicWidth(ItemX.Pic) > 32 Then Slot(c + 1) = ItemX.Index
			If ItemPicHeight(ItemX.Pic) > 32 And ItemPicWidth(ItemX.Pic) > 32 Then
				Slot(c + Width_Renamed + 1) = ItemX.Index
			End If
			If ItemPicHeight(ItemX.Pic) > 64 And ItemPicWidth(ItemX.Pic) > 32 Then
				Slot(c + Width_Renamed * 2 + 1) = ItemX.Index
			End If
		Else
			ItemX.InvSpot = 0
		End If
	End Sub
	
	Private Function InventoryViceText(ByRef c As Short) As String
		Select Case c
			Case 0 To 14
				InventoryViceText = "Superb"
			Case 15 To 16
				InventoryViceText = "Great"
			Case 17
				InventoryViceText = "Good"
			Case 18
				InventoryViceText = "Average"
			Case 19
				InventoryViceText = "Weak"
			Case Else
				InventoryViceText = "Uncontrolled"
		End Select
	End Function
	
	Private Sub InventoryHint(ByRef Style As Short, ByRef AtX As Short, ByRef AtY As Short)
		'UPGRADE_NOTE: Text was upgraded to Text_Renamed. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="A9E4979A-37FA-4718-9994-97DD76ED70A7"'
		Dim c, Found As Short
		Dim ItemX As Item
		Dim Text_Renamed As String
		Found = False
		Select Case Style
			Case 0 ' Inventory
				For c = 1 To 128
					If PointIn(AtX, AtY, InvX(c), InvY(c), 34, 34) And InvZ(c) > -1 Then
						For	Each ItemX In CreatureWithTurn.Items
							If ItemX.Index = InvZ(c) Then
								Found = True
								Exit For
							End If
						Next ItemX
						Exit For
					End If
				Next c
			Case 1 ' Search
				For c = 1 To 80
					If PointIn(AtX, AtY, 23 + ((c - 1) Mod 8) * 34, 45 + Int((c - 1) / 8) * 34, 34, 34) And InvZ(c) > -1 Then
						ItemX = EncounterNow.Items.Item("I" & InvZ(c))
						Found = True
						Exit For
					End If
				Next c
			Case 2 ' Container
				For c = 1 To 80
					If PointIn(AtX, AtY, 23 + ((c - 1) Mod 8) * 34, 45 + Int((c - 1) / 8) * 34, 34, 34) And InvC(c) > 0 Then
						ItemX = InvContainer.Items.Item("I" & InvC(c))
						Found = True
						Exit For
					End If
				Next c
		End Select
		If Found = True Then
			If ItemX.Selected = False Then
				ItemNow = ItemX
				BorderDrawButtons(-1)
			Else
				BorderDrawButtons(0)
			End If
		Else
			BorderDrawButtons(0)
		End If
		Me.Refresh()
	End Sub
	
	'UPGRADE_NOTE: Tome was upgraded to Tome_Renamed. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="A9E4979A-37FA-4718-9994-97DD76ED70A7"'
	Public Sub InventoryShow(ByRef ListToShow As Short)
		Dim Tome_Renamed As Object
		'UPGRADE_NOTE: Text was upgraded to Text_Renamed. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="A9E4979A-37FA-4718-9994-97DD76ED70A7"'
		Dim k, c, i, Found As Short
		Dim rc As Integer
		Dim Text_Renamed As String
		'UPGRADE_NOTE: Size was upgraded to Size_Renamed. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="A9E4979A-37FA-4718-9994-97DD76ED70A7"'
		'UPGRADE_NOTE: Width was upgraded to Width_Renamed. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="A9E4979A-37FA-4718-9994-97DD76ED70A7"'
		'UPGRADE_NOTE: Height was upgraded to Height_Renamed. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="A9E4979A-37FA-4718-9994-97DD76ED70A7"'
		Dim Height_Renamed, Width_Renamed, NoFail As Short
		Dim Size_Renamed As Double
		Dim NextLevel As Double
		Dim OldFrozen As Short
		Dim ItemX As Item
		Dim TriggerX, TriggerZ As Trigger
		Dim X, Y As Short
		'UPGRADE_WARNING: Arrays in structure bmMons may need to be initialized before they can be used. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="814DF224-76BD-4BB4-BFFB-EA359CB9FC48"'
		Dim bmMons As BITMAPINFO
		Dim hMemMons As Integer
		'UPGRADE_WARNING: Arrays in structure bmBlack may need to be initialized before they can be used. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="814DF224-76BD-4BB4-BFFB-EA359CB9FC48"'
		Dim bmBlack As BITMAPINFO
		'UPGRADE_WARNING: Arrays in structure bmMask may need to be initialized before they can be used. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="814DF224-76BD-4BB4-BFFB-EA359CB9FC48"'
		Dim bmMask As BITMAPINFO
		Dim lpMem, TransparentRGB As Integer
		' Freeze out rest of screen
		'UPGRADE_WARNING: Screen property Screen.MousePointer has a new behavior. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6BA9B8D2-2A32-4B6E-8D36-44949974A5B4"'
		System.Windows.Forms.Cursor.Current = System.Windows.Forms.Cursors.WaitCursor
		Frozen = True
		' Show header
		'UPGRADE_ISSUE: PictureBox method picInventory.Cls was not upgraded. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="CC4C7EC0-C903-48FC-ACCC-81861D12DA4A"'
		picInventory.Cls()
		'    ShowText picInventory, 0, 12, 584, 14, bdFontElixirWhite, CreatureWithTurn.Name, True, False
		' [Titi 2.4.8] moved AFTER the picture is loaded ;)
		Dim sTmp As String
		Select Case ListToShow
			Case bdInvItems
				'            picInventory.Picture = LoadPicture(gAppPath & "\data\interface\" & GlobalInterfaceName & "\DialogInventory.bmp")
				picInventory.Image = System.Drawing.Image.FromFile(gDataPath & "\interface\" & GlobalInterfaceName & "\DialogInventory.bmp")
				ShowText(picInventory, 0, 12, 584, 14, bdFontElixirWhite, (CreatureWithTurn.Name), True, False)
				InvNowShow = bdInvItems
				' Show Bulk and Weight ([Titi 2.4.8] modified the bulk max)
				ShowText(picInventory, 16, 360, 82, 14, bdFontNoxiousWhite, "Bulk " & Int((CreatureWithTurn.Bulk / Greatest(CreatureWithTurn.Agility * 8, 1)) * 100) & "%", True, False)
				ShowText(picInventory, 186, 360, 82, 14, bdFontNoxiousWhite, "Wgt " & Int((CreatureWithTurn.Weight / Greatest((CreatureWithTurn.MaxWeight), 1)) * 100) & "%", True, False)
				' Clear current locations
				For c = 0 To 128
					InvZ(c) = -1
				Next c
				' Show all Items
				InvTooMany = False
				For c = 0 To 1
					For	Each ItemX In CreatureWithTurn.Items
						If (c = 0 And ItemX.InvSpot > 0) Or (c = 1 And ItemX.InvSpot = 0) Then
							LoadItemPic(ItemX)
							' Set Inventory Spot on screen
							If ItemX.IsReady = True Then
								Select Case ItemX.WearType
									Case 0 ' Body
										InventoryFindSpot(ItemX, 27, 32, 2, InvZ)
										If ItemX.InvSpot = 0 Then
											InventoryFindSpot(ItemX, 25, 32, 2, InvZ)
										End If
									Case 1 ' Helm
										InventoryFindSpot(ItemX, 13, 16, 2, InvZ)
									Case 2, 3, 9 ' Glove, Bracelet, Ring
										InventoryFindSpot(ItemX, 17, 20, 2, InvZ)
										If ItemX.InvSpot = 0 Then
											InventoryFindSpot(ItemX, 21, 24, 2, InvZ)
										End If
									Case 4 ' Backpack
										InventoryFindSpot(ItemX, 1, 6, 2, InvZ)
										If ItemX.InvSpot = 0 Then
											InventoryFindSpot(ItemX, 7, 12, 2, InvZ)
										End If
									Case 5 ' Shield
										InventoryFindSpot(ItemX, 39, 44, 2, InvZ)
									Case 6 ' Boots
										InventoryFindSpot(ItemX, 45, 48, 2, InvZ)
									Case 7 ' Necklace
										InventoryFindSpot(ItemX, 25, 26, 2, InvZ)
									Case 8 ' Belt
										InventoryFindSpot(ItemX, 31, 32, 2, InvZ)
									Case Else ' InHand
										InventoryFindSpot(ItemX, 33, 38, 2, InvZ)
								End Select
								' If no spot available, unready
								If ItemX.InvSpot = 0 Then
									ItemX.IsReady = False
									InventoryFindSpot(ItemX, 49, 128, 8, InvZ)
								End If
							Else
								InventoryFindSpot(ItemX, 49, 128, 8, InvZ)
							End If
							' Draw Item in Inventory (if have a spot to display)
							If ItemX.Selected = False Then
								If ItemX.InvSpot > 0 Then
									Width_Renamed = ItemPicWidth(ItemX.Pic)
									Height_Renamed = ItemPicHeight(ItemX.Pic)
									'UPGRADE_ISSUE: PictureBox property picItem.hdc was not upgraded. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="CC4C7EC0-C903-48FC-ACCC-81861D12DA4A"'
									'UPGRADE_ISSUE: PictureBox property picInventory.hdc was not upgraded. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="CC4C7EC0-C903-48FC-ACCC-81861D12DA4A"'
									rc = BitBlt(picInventory.hdc, InvX(ItemX.InvSpot), InvY(ItemX.InvSpot), Width_Renamed, Height_Renamed, picItem.hdc, 64 * ItemX.Pic - 64, 96, SRCAND)
									'UPGRADE_ISSUE: PictureBox property picItem.hdc was not upgraded. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="CC4C7EC0-C903-48FC-ACCC-81861D12DA4A"'
									'UPGRADE_ISSUE: PictureBox property picInventory.hdc was not upgraded. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="CC4C7EC0-C903-48FC-ACCC-81861D12DA4A"'
									rc = BitBlt(picInventory.hdc, InvX(ItemX.InvSpot), InvY(ItemX.InvSpot), Width_Renamed, Height_Renamed, picItem.hdc, 64 * ItemX.Pic - 64, 0, SRCPAINT)
									If ItemX.Capacity > 0 Then
										ShowText(picInventory, InvX(ItemX.InvSpot), InvY(ItemX.InvSpot), Width_Renamed, 10, bdFontSmallWhite, "(" & ItemX.Items.Count() & ")", 1, False)
									End If
									If ItemX.Count > 1 Then
										ShowText(picInventory, InvX(ItemX.InvSpot), InvY(ItemX.InvSpot) + Height_Renamed - 10, Width_Renamed, 10, bdFontSmallWhite, CStr(ItemX.Count), True, False)
									End If
								Else
									InvTooMany = True
								End If
							End If
						End If
					Next ItemX
				Next c
				' Show Buttons
				If InvTooMany = True Then
					ShowButton(picInventory, 290, 406, "More", False)
				End If
				ShowButton(picInventory, 386, 406, "Status", False)
				ShowButton(picInventory, 482, 406, "Done", False)
				picInventory.Refresh()
				picInventory.Left = 0
				picInventory.Top = Me.ClientRectangle.Height - 48 - picMenu.ClientRectangle.Height - picInventory.ClientRectangle.Height
				picInventory.Visible = True
			Case bdInvStatus
				'            picInventory.Picture = LoadPicture(gAppPath & "\data\interface\" & GlobalInterfaceName & "\DialogStatus.bmp")
				picInventory.Image = System.Drawing.Image.FromFile(gDataPath & "\interface\" & GlobalInterfaceName & "\DialogStatus.bmp")
				ShowText(picInventory, 0, 12, 584, 14, bdFontElixirWhite, (CreatureWithTurn.Name), True, False)
				InvNowShow = bdInvStatus
				' Load Clean Creature Picture
				'            sTmp = gAppPath & "\data\graphics\creatures\" & CreatureWithTurn.PictureFile
				sTmp = gDataPath & "\graphics\creatures\" & CreatureWithTurn.PictureFile
				If Not oFileSys.CheckExists(sTmp, clsInOut.IOActionType.File) Then
					'UPGRADE_WARNING: Couldn't resolve default property of object Tome.FullPath. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
                    sTmp = tome.FullPath & "\" & CreatureWithTurn.PictureFile
					If Not oFileSys.CheckExists(sTmp, clsInOut.IOActionType.File) Then
						'UPGRADE_WARNING: Couldn't resolve default property of object Tome.FullPath. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
                        sTmp = tome.FullPath & "\creatures\" & CreatureWithTurn.PictureFile
						If Not oFileSys.CheckExists(sTmp, clsInOut.IOActionType.File) Then
							sTmp = ""
						End If
					End If
				End If
				If sTmp <> "" Then
					ReadBitmapFile(sTmp, bmMons, hMemMons, TransparentRGB)
					' Make a copy of the current palette for the picture
					'UPGRADE_WARNING: Couldn't resolve default property of object bmBlack. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
					bmBlack = bmMons
					' Then change Pure Blue to Pure Black
					ChangeColor(bmBlack, TransparentRGB, 0, 0, 0)
					MakeMask(bmMons, bmMask, TransparentRGB)
					' Paint bitmap to picture box using converted palette
					lpMem = GlobalLock(hMemMons)
					Width_Renamed = bmMons.bmiHeader.biWidth
					Height_Renamed = bmMons.bmiHeader.biHeight
					Size_Renamed = 1
					If Width_Renamed * Size_Renamed > 228 Or Height_Renamed * Size_Renamed > 266 Then
						Size_Renamed = 228 / Width_Renamed
						If Height_Renamed * Size_Renamed > 266 Then
							Size_Renamed = 266 / Height_Renamed
						End If
					End If
					X = 35 + (228 - Width_Renamed * Size_Renamed) / 2 : Y = 81 + (266 - Height_Renamed * Size_Renamed) / 2
					'UPGRADE_ISSUE: PictureBox property picInventory.hdc was not upgraded. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="CC4C7EC0-C903-48FC-ACCC-81861D12DA4A"'
					rc = SetStretchBltMode(picInventory.hdc, 3)
					'UPGRADE_ISSUE: PictureBox property picInventory.hdc was not upgraded. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="CC4C7EC0-C903-48FC-ACCC-81861D12DA4A"'
					rc = StretchDIBits(picInventory.hdc, X, Y, Width_Renamed * Size_Renamed, Height_Renamed * Size_Renamed, 0, 0, Width_Renamed, Height_Renamed, lpMem, bmMask, DIB_RGB_COLORS, SRCAND)
					'UPGRADE_ISSUE: PictureBox property picInventory.hdc was not upgraded. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="CC4C7EC0-C903-48FC-ACCC-81861D12DA4A"'
					rc = StretchDIBits(picInventory.hdc, X, Y, Width_Renamed * Size_Renamed, Height_Renamed * Size_Renamed, 0, 0, Width_Renamed, Height_Renamed, lpMem, bmBlack, DIB_RGB_COLORS, SRCPAINT)
					' Release memory
					rc = GlobalUnlock(hMemMons)
					rc = GlobalFree(hMemMons)
				End If
				' Show Level and basic Stats
				ShowText(picInventory, 45, 52, 216, 14, bdFontNoxiousBlack, "Level: " & VB6.Format(CreatureWithTurn.Level), True, False)
				ShowText(picInventory, 45, 82, 216, 14, bdFontNoxiousBlack, "Int " & VB6.Format(CreatureWithTurn.Will), True, False)
				ShowText(picInventory, 45, 102, 216, 14, bdFontNoxiousBlack, "Str " & VB6.Format(CreatureWithTurn.Strength), 0, False)
				ShowText(picInventory, 45, 102, 216, 14, bdFontNoxiousBlack, "Agl " & VB6.Format(CreatureWithTurn.Agility), 1, False)
				If CreatureWithTurn.HPNow > 0 And CreatureWithTurn.HPMax > 0 Then
					ShowText(picInventory, 45, 335, 216, 14, bdFontNoxiousBlack, "Health " & VB6.Format(CreatureWithTurn.HPNow) & " of " & VB6.Format(CreatureWithTurn.HPMax) & " (" & Int((CreatureWithTurn.HPNow / CreatureWithTurn.HPMax) * 100) & "%)", True, False)
				ElseIf CreatureWithTurn.HPMax > 0 Then 
					ShowText(picInventory, 45, 335, 216, 14, bdFontNoxiousBlack, "Dead " & VB6.Format(CreatureWithTurn.HPNow) & " of " & VB6.Format(CreatureWithTurn.HPMax) & " (" & Int((CreatureWithTurn.HPNow / CreatureWithTurn.HPMax) * 100) & "%)", True, False)
				Else
					ShowText(picInventory, 45, 335, 216, 14, bdFontNoxiousBlack, "Stone Dead", True, False)
				End If
				ShowText(picInventory, 45, 365, 216, 14, bdFontNoxiousBlack, "Defense " & VB6.Format(CreatureWithTurn.Defense), True, False)
				' Set up Status
				ScrollList = New Collection
				' Determine if there are any Effects Triggers
				Found = False
				For	Each TriggerX In CreatureWithTurn.Triggers
					If (TriggerX.Styles And &H7C) > 0 And TriggerX.TriggerType <> bdOnCast Then
						Found = True
						Exit For
					End If
				Next TriggerX
				' If there are Effects Triggers (that are not On-Cast Triggers), show them
				If Found = True Then
					ScrollList.Add("`Effects")
					For	Each TriggerX In CreatureWithTurn.Triggers
						If (TriggerX.Styles And &H7C) > 0 And TriggerX.TriggerType <> bdOnCast Then
							ScrollList.Add(TriggerX.Name)
						End If
					Next TriggerX
					ScrollList.Add("")
				End If
				' Vices
				ScrollList.Add("`Vices")
				ScrollList.Add("Lunacy: " & InventoryViceText((CreatureWithTurn.Lunacy)))
				ScrollList.Add("Revelry: " & InventoryViceText((CreatureWithTurn.Revelry)))
				ScrollList.Add("Wrath: " & InventoryViceText((CreatureWithTurn.Wrath)))
				ScrollList.Add("Pride: " & InventoryViceText((CreatureWithTurn.Pride)))
				ScrollList.Add("Greed: " & InventoryViceText((CreatureWithTurn.Greed)))
				ScrollList.Add("Lust: " & InventoryViceText((CreatureWithTurn.Lust)))
				' Add any additional Special Status
				Found = False
				For	Each TriggerX In CreatureWithTurn.Triggers
					If TriggerX.TriggerType = bdOnStatus Then
						Found = True
						Exit For
					End If
				Next TriggerX
				If Found = True Then
					ScrollList.Add("")
					ScrollList.Add("`Stats")
					For	Each TriggerX In CreatureWithTurn.Triggers
						If TriggerX.TriggerType = bdOnStatus Then
							CreatureNow = CreatureWithTurn
							ItemNow = CreatureWithTurn.ItemInHand
							NoFail = FireTrigger(CreatureWithTurn, TriggerX)
							ScrollList.Add(TriggerX.Name & ": " & TriggerX.VarA)
						End If
					Next TriggerX
				End If
				' Experience Information
				ScrollList.Add("")
				ScrollList.Add("`Experience")
				ScrollList.Add("Current: " & VB6.Format(CreatureWithTurn.ExperiencePoints, "###,###,###,##0"))
				c = CreatureWithTurn.Level : NextLevel = (c * Int(c / 2 + 0.5) + (c / 2) * (1 - (c Mod 2))) * 1000
				ScrollList.Add("Next Level: " & VB6.Format(NextLevel, "###,###,###,##0"))
				ScrollList.Add("Skill Points: " & CreatureWithTurn.SkillPoints)
				' Resistances
				ScrollList.Add("")
				ScrollList.Add("`Resistances")
				ScrollList.Add("Sharp " & CreatureWithTurn.ResistanceTypeWithArmor(1) & "%")
				ScrollList.Add("Blunt " & CreatureWithTurn.ResistanceTypeWithArmor(2) & "%")
				ScrollList.Add("Cold " & CreatureWithTurn.ResistanceTypeWithArmor(3) & "%")
				ScrollList.Add("Heat " & CreatureWithTurn.ResistanceTypeWithArmor(4) & "%")
				ScrollList.Add("Evil " & CreatureWithTurn.ResistanceTypeWithArmor(5) & "%")
				ScrollList.Add("Good " & CreatureWithTurn.ResistanceTypeWithArmor(6) & "%")
				ScrollList.Add("Energy " & CreatureWithTurn.ResistanceTypeWithArmor(7) & "%")
				ScrollList.Add("Mind " & CreatureWithTurn.ResistanceTypeWithArmor(8) & "%")
				' Show all Stats
				ScrollTop = Greatest(Least(ScrollTop, ScrollList.Count() - 17), 1)
				For c = 0 To Least(17, ScrollList.Count() - ScrollTop)
					'UPGRADE_WARNING: Couldn't resolve default property of object ScrollList(). Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
					If VB.Left(ScrollList.Item(ScrollTop + c), 1) = "`" Then
						'UPGRADE_WARNING: Couldn't resolve default property of object ScrollList(). Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
						ShowText(picInventory, 305, 50 + c * 18, 235, 14, bdFontNoxiousGold, Mid(ScrollList.Item(ScrollTop + c), 2), True, False)
					Else
						'UPGRADE_WARNING: Couldn't resolve default property of object ScrollList(). Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
						ShowText(picInventory, 305, 50 + c * 18, 235, 14, bdFontNoxiousWhite, ScrollList.Item(ScrollTop + c), True, False)
					End If
				Next c
				' Show scroll bar
				ScrollBarShow(picInventory, 544, 45, 338, ScrollTop, ScrollList.Count() - 17, 0)
				' Show Buttons
				ShowButton(picInventory, 386, 406, "Equip", False)
				ShowButton(picInventory, 482, 406, "Done", False)
				picInventory.Refresh()
				picInventory.Visible = True
		End Select
		picInventory.Refresh()
		'UPGRADE_WARNING: Screen property Screen.MousePointer has a new behavior. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6BA9B8D2-2A32-4B6E-8D36-44949974A5B4"'
		System.Windows.Forms.Cursor.Current = System.Windows.Forms.Cursors.Default
	End Sub
	
	'UPGRADE_NOTE: Add was upgraded to Add_Renamed. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="A9E4979A-37FA-4718-9994-97DD76ED70A7"'
	Private Sub MenuSkillSpend(ByRef Add_Renamed As Short)
		Dim TriggerX As Trigger
		Dim NoFail As Short
		' Only allow spend if there is a positive cost for the Skill
		TriggerX = ScrollList.Item(CreateSkillsIndex)
		If TriggerX.Turns > 0 Then
			' Only add if have enough points to level. Else, only subtract to previous level
			If Add_Renamed = True And CreatureWithTurn.SkillPoints >= TriggerX.Turns Then
				CreatureNow = CreatureWithTurn
				GlobalSkillLevel = Int(TriggerX.SkillPoints / TriggerX.Turns)
				NoFail = FireTriggers(TriggerX, bdPreRaiseSkill)
				If NoFail Then
					TriggerX.SkillPoints = TriggerX.SkillPoints + TriggerX.Turns
					CreatureWithTurn.SkillPoints = CreatureWithTurn.SkillPoints - TriggerX.Turns
				End If
			ElseIf Add_Renamed = False And CShort(CShort(TriggerX.SkillPoints) - CShort(TriggerX.Turns)) >= CShort(TriggerX.TempSkill) Then 
				TriggerX.SkillPoints = TriggerX.SkillPoints - TriggerX.Turns
				CreatureWithTurn.SkillPoints = CreatureWithTurn.SkillPoints + TriggerX.Turns
			End If
			MenuSkillShow(CreateSkillsIndex)
		End If
	End Sub
	
	Private Sub MenuSkillSet()
	End Sub
	
	Private Sub DoorListen()
		Dim TriggerX As Trigger
		Dim Found, NoFail, OldFrozen As Short
		OldFrozen = Frozen
		Frozen = True
		If HasTrigger(EncounterNow, bdOnListen) Then
			CreatureNow = CreatureWithTurn
			ItemNow = CreatureWithTurn.ItemInHand
			NoFail = FireTriggers(EncounterNow, bdOnListen)
        ElseIf HasTrigger(tome, bdOnListen) Then
            CreatureNow = CreatureWithTurn
            ItemNow = CreatureWithTurn.ItemInHand
            NoFail = FireTriggers(tome, bdOnListen)
		Else
			DialogDM(CreatureWithTurn.Name & " listens, but doesn't hear anything.")
		End If
		Frozen = OldFrozen
	End Sub
	
	Private Sub SearchShow()
		Dim c, yOffSet As Short
		Dim rc As Integer
		Dim ItemX As Item
		'UPGRADE_NOTE: Height was upgraded to Height_Renamed. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="A9E4979A-37FA-4718-9994-97DD76ED70A7"'
		'UPGRADE_NOTE: Width was upgraded to Width_Renamed. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="A9E4979A-37FA-4718-9994-97DD76ED70A7"'
		Dim Width_Renamed, X, Y, Height_Renamed As Short
		'UPGRADE_WARNING: Screen property Screen.MousePointer has a new behavior. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6BA9B8D2-2A32-4B6E-8D36-44949974A5B4"'
		System.Windows.Forms.Cursor.Current = System.Windows.Forms.Cursors.WaitCursor
		'UPGRADE_ISSUE: PictureBox method picSearch.Cls was not upgraded. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="CC4C7EC0-C903-48FC-ACCC-81861D12DA4A"'
		picSearch.Cls()
		' Show header
		ShowText(picSearch, 0, 12, picSearch.Width, 14, bdFontElixirWhite, "On the Ground", True, False)
		' Clear current locations
		For c = 0 To 128
			InvZ(c) = -1
		Next c
		' Show all Items (those with a spot first, without a spot second)
		InvTooMany = False
		For	Each ItemX In EncounterNow.Items
			If ItemX.Selected = False Then
				LoadItemPic(ItemX)
				InventoryFindSpot(ItemX, 1, 80, 8, InvZ)
				' Draw Item in Inventory (if have a spot to display)
				If ItemX.InvSpot > 0 Then
					' Configure Width and Height
					Width_Renamed = ItemPicWidth(ItemX.Pic)
					Height_Renamed = ItemPicHeight(ItemX.Pic)
					' Set location on Search box
					X = 23 + ((ItemX.InvSpot - 1) Mod 8) * 34
					Y = 45 + Int((ItemX.InvSpot - 1) / 8) * 34
					' Paint the Item picture
					'UPGRADE_ISSUE: PictureBox property picItem.hdc was not upgraded. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="CC4C7EC0-C903-48FC-ACCC-81861D12DA4A"'
					'UPGRADE_ISSUE: PictureBox property picSearch.hdc was not upgraded. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="CC4C7EC0-C903-48FC-ACCC-81861D12DA4A"'
					rc = BitBlt(picSearch.hdc, X, Y, Width_Renamed, Height_Renamed - yOffSet, picItem.hdc, 64 * ItemX.Pic - 64, 96, SRCAND)
					'UPGRADE_ISSUE: PictureBox property picItem.hdc was not upgraded. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="CC4C7EC0-C903-48FC-ACCC-81861D12DA4A"'
					'UPGRADE_ISSUE: PictureBox property picSearch.hdc was not upgraded. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="CC4C7EC0-C903-48FC-ACCC-81861D12DA4A"'
					rc = BitBlt(picSearch.hdc, X, Y, Width_Renamed, Height_Renamed - yOffSet, picItem.hdc, 64 * ItemX.Pic - 64, 0, SRCPAINT)
					If ItemX.Capacity > 0 And yOffSet = 0 Then
						ShowText(picSearch, X, Y, Width_Renamed, 10, bdFontSmallWhite, "(" & ItemX.Items.Count() & ")", 1, False)
					End If
					If ItemX.Count > 1 Then
						ShowText(picSearch, X, Y + Height_Renamed - 10, Width_Renamed, 10, bdFontSmallWhite, CStr(ItemX.Count), True, False)
					End If
				Else
					InvTooMany = True
				End If
			End If
		Next ItemX
		' Show Buttons
		If InvTooMany = True Then
			ShowButton(picSearch, picSearch.ClientRectangle.Width - 294, picSearch.ClientRectangle.Height - 30, "More", False)
		End If
		ShowButton(picSearch, picSearch.ClientRectangle.Width - 198, picSearch.ClientRectangle.Height - 30, "Take All", False)
		ShowButton(picSearch, picSearch.ClientRectangle.Width - 102, picSearch.ClientRectangle.Height - 30, "Done", False)
		picSearch.Left = 0
		picSearch.Top = Me.ClientRectangle.Height - 48 - picMenu.ClientRectangle.Height - picSearch.ClientRectangle.Height
		' Search gets the focus however
		picSearch.Visible = True
		picSearch.BringToFront()
		picSearch.Focus()
		'UPGRADE_WARNING: Screen property Screen.MousePointer has a new behavior. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6BA9B8D2-2A32-4B6E-8D36-44949974A5B4"'
		System.Windows.Forms.Cursor.Current = System.Windows.Forms.Cursors.Default
	End Sub
	
	Private Sub SearchClose()
		Dim c As Short
		Dim rc As Integer
		' Close Search and fire Post-Search Triggers
		picSearch.Visible = False
		CreatureNow = CreatureWithTurn
		ItemNow = CreatureWithTurn.ItemInHand
		rc = FireTriggers(EncounterNow, bdPostSearch)
		' Redraw Map and release
		DrawMapAll()
		BorderDrawButtons(0)
		Frozen = False
	End Sub
	
	Private Sub SearchClick(ByRef AtX As Short, ByRef AtY As Short, ByRef ButtonDown As Short, ByRef Button As Short)
		Dim c, i As Short
		Dim rc As Integer
		Dim ItemX As Item
		'UPGRADE_NOTE: Height was upgraded to Height_Renamed. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="A9E4979A-37FA-4718-9994-97DD76ED70A7"'
		'UPGRADE_NOTE: Width was upgraded to Width_Renamed. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="A9E4979A-37FA-4718-9994-97DD76ED70A7"'
		Dim Width_Renamed, Height_Renamed As Short
		If PointIn(AtX, AtY, picSearch.ClientRectangle.Width - 198, picSearch.ClientRectangle.Height - 30, 90, 18) Then
			' Take All
			ShowButton(picSearch, picSearch.ClientRectangle.Width - 198, picSearch.ClientRectangle.Height - 30, "Take All", ButtonDown)
			picSearch.Refresh()
			If ButtonDown = False Then
				Call PlayClickSnd(modIOFunc.ClickType.ifClick)
				For	Each ItemX In EncounterNow.Items
					TradeItem(CreatureWithTurn, CreatureWithTurn, bdInvEncounter, bdInvObjects, 0, 0, ItemX)
				Next ItemX
				SearchShow()
				If EncounterNow.Items.Count() < 1 Then
					SearchClose()
				End If
			End If
		ElseIf PointIn(AtX, AtY, picSearch.ClientRectangle.Width - 102, picSearch.ClientRectangle.Height - 30, 90, 18) Then 
			' Done
			ShowButton(picSearch, picSearch.ClientRectangle.Width - 102, picSearch.ClientRectangle.Height - 30, "Done", ButtonDown)
			picSearch.Refresh()
			If ButtonDown = False Then
				Call PlayClickSnd(modIOFunc.ClickType.ifClick)
				SearchClose()
			End If
		ElseIf PointIn(AtX, AtY, picSearch.ClientRectangle.Width - 294, picSearch.ClientRectangle.Height - 30, 90, 18) And InvTooMany = True Then 
			' More
			ShowButton(picSearch, picSearch.ClientRectangle.Width - 294, picSearch.ClientRectangle.Height - 30, "More", ButtonDown)
			picSearch.Refresh()
			If ButtonDown = False Then
				Call PlayClickSnd(modIOFunc.ClickType.ifClick)
				' Shuffle the Items
				For	Each ItemX In EncounterNow.Items
					If ItemX.InvSpot > 0 Then
						EncounterNow.RemoveItem("I" & ItemX.Index)
						ItemX.InvSpot = 0
						EncounterNow.AddItem.Copy(ItemX)
					End If
				Next ItemX
				SearchShow()
			End If
		ElseIf ButtonDown Then 
			' Drag and drop some items
			For c = 1 To 80
				If PointIn(AtX, AtY, 23 + ((c - 1) Mod 8) * 34, 45 + Int((c - 1) / 8) * 34, 34, 34) And InvZ(c) > -1 Then
					ItemX = EncounterNow.Items.Item("I" & InvZ(c))
					If Button = VB6.MouseButtonConstants.RightButton Then
						Call PlayClickSnd(modIOFunc.ClickType.ifClick)
						ExamineItem(ItemX)
					Else
						'UPGRADE_ISSUE: PictureBox method picInvDrag.Cls was not upgraded. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="CC4C7EC0-C903-48FC-ACCC-81861D12DA4A"'
						picInvDrag.Cls()
						LoadItemPic(ItemX)
						Select Case ItemPicWidth(ItemX.Pic)
							Case 32 : Width_Renamed = 34
							Case 64 : Width_Renamed = 68
						End Select
						Select Case ItemPicHeight(ItemX.Pic)
							Case 32 : Height_Renamed = 34
							Case 64 : Height_Renamed = 68
							Case 96 : Height_Renamed = 102
						End Select
						For i = 1 To 80
							If InvZ(i) = ItemX.InvSpot Then InvZ(i) = -1
						Next i
						picInvDrag.Width = Width_Renamed : picInvDrag.Height = Height_Renamed
						'UPGRADE_ISSUE: PictureBox property picSearch.hdc was not upgraded. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="CC4C7EC0-C903-48FC-ACCC-81861D12DA4A"'
						'UPGRADE_ISSUE: PictureBox property picInvDrag.hdc was not upgraded. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="CC4C7EC0-C903-48FC-ACCC-81861D12DA4A"'
						rc = BitBlt(picInvDrag.hdc, 0, 0, Width_Renamed, Height_Renamed, picSearch.hdc, 22 + ((ItemX.InvSpot - 1) Mod 8) * 34, 44 + Int((ItemX.InvSpot - 1) / 8) * 34, SRCCOPY)
						picInvDrag.Refresh()
						' Ready for Drag/Drop
						JournalX = AtX - 23 - ((ItemX.InvSpot - 1) Mod 8) * 34 : JournalY = AtY - 45 - Int((ItemX.InvSpot - 1) / 8) * 34
						InvDragItem = ItemX
						InvDragPos = c
						InvDragIndex = 0
						InvDragMode = True
						InvDragItem.Selected = True
					End If
					Exit For
				End If
			Next c
		End If
	End Sub
	
	Private Sub Search()
		Dim Found, NoFail As Short
		Dim CreatureX As Creature
		Dim ItemX, ItemZ As Item
		For	Each CreatureX In EncounterNow.Creatures
			' [Titi 2.4.9] You cannot even try to search while in combat!
			'        If (EncounterNow.HaveEntered = False Or (CreatureX.Agressive = True And CreatureX.HPNow > 0)) And EncounterNow.Creatures.Count > 0 Then
			If (CreatureX.Agressive = True And CreatureX.HPNow > 0) And EncounterNow.CanFight = True Then
				DialogDM("You can't search while fighting...")
				DialogHide()
				CombatStart()
				Exit Sub
			End If
			If CreatureX.Guard = True And CreatureX.AllowedTurn = True Then
				Found = True
				Exit For
			End If
		Next CreatureX
		If Found Then
			DialogDM("You can't search while the " & CreatureX.Name & " is guarding.")
			DialogHide()
		Else
			CreatureNow = CreatureWithTurn
			ItemNow = CreatureWithTurn.ItemInHand
			NoFail = FireTriggers(EncounterNow, bdPreSearch)
			If NoFail Then
				' Dump Creatures Items into Encounter
				For	Each CreatureX In EncounterNow.Creatures
					If CreatureX.AllowedTurn = False Then
						For	Each ItemX In CreatureX.Items
							ItemX.IsReady = False
							EncounterNow.AddItem.Copy(ItemX)
							CreatureX.RemoveItem("I" & ItemX.Index)
						Next ItemX
					End If
				Next CreatureX
				' Generate any random items as required
				For	Each ItemX In EncounterNow.Items
					If IsBetween(ItemX.Family, 9, 13) Then
						MakeItem(ItemX)
					End If
				Next ItemX
				' Check for any Items to find
				If EncounterNow.Items.Count() < 1 Then
					DialogDM(CreatureWithTurn.Name & " searches, but doesn't find anything.")
					DialogHide()
				Else
					ScrollTop = 1
					SearchShow()
					Frozen = True
				End If
			End If
		End If
	End Sub
	
	'UPGRADE_NOTE: Tome was upgraded to Tome_Renamed. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="A9E4979A-37FA-4718-9994-97DD76ED70A7"'
	Private Sub SearchDrag(ByRef AtX As Short, ByRef AtY As Short)
		Dim Tome_Renamed As Object
		Dim X, c, Y As Short
		' No matter what, shut off dragging Item
		InvDragItem.Selected = False
		InvDragMode = False
		picInvDrag.Visible = False
		' First, check if droppping onto an open Container
		If PointIn(AtX, AtY, (picSearch.Left), (picSearch.Top), (picSearch.Width), (picSearch.Height)) Then
			' Drop on Search
			For c = 1 To 80
				X = picSearch.Left + 23 + ((c - 1) Mod 8) * 34
				Y = picSearch.Top + 45 + Int((c - 1) / 8) * 34
				If PointIn(AtX, AtY, X, Y, 34, 34) Then
					TradeItem(CreatureWithTurn, CreatureWithTurn, bdInvEncounter, bdInvEncounter, InvZ(c), c, InvDragItem)
					SearchShow()
					Exit For
				End If
			Next c
		ElseIf AtY > picMenu.Top Then 
			' Drop to Party member
			c = Greatest(Least(Int((AtX - picMenu.Left) / 122), 4), 0)
			'UPGRADE_WARNING: Couldn't resolve default property of object Tome.Creatures. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
            If IsBetween(c, 0, tome.Creatures.Count - 1) Then
                TradeItem(CreatureWithTurn, MenuParty(c), bdInvEncounter, bdInvObjects, 0, 0, InvDragItem)
            End If
		End If
	End Sub
	
	'UPGRADE_NOTE: Tome was upgraded to Tome_Renamed. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="A9E4979A-37FA-4718-9994-97DD76ED70A7"'
	Public Sub UseShow(ByRef ItemToUse As Item)
		Dim Tome_Renamed As Object
		Dim OldActionPoints, Found, NoFail, rc As Short
		Dim ItemX As Item
		Dim TriggerX As Trigger
		Dim StmtX As Statement
		Dim HasTarget As Short
		' Save actions points. If find use, deduct (if Trigger does not do it)
		OldActionPoints = CreatureWithTurn.ActionPoints
		Found = False
		' Fire Pre-UseInEncounter
		If HasTrigger(EncounterNow, bdPreUseInEncounter) = True Then
			CreatureNow = CreatureWithTurn
			ItemNow = ItemToUse
			NoFail = FireTriggers(EncounterNow, bdPreUseInEncounter)
			Found = True
		End If
		' Determine what type of item is being used and use it
		CreatureNow = CreatureWithTurn
		ItemNow = ItemToUse
		If HasTrigger(ItemToUse, bdOnUseOnItem) = True Then
			If TargetItem(ItemTarget) = True Then
				If FireTriggers(ItemTarget, bdPreUseOnItem) = True Then
					NoFail = FireTriggers(ItemToUse, bdOnUseOnItem)
					NoFail = FireTriggers(ItemTarget, bdPostUseOnItem)
				End If
			End If
			Found = True
		ElseIf HasTrigger(ItemToUse, bdOnUseOnCreature) = True Then 
			NoFail = True
			For	Each TriggerX In ItemToUse.Triggers
				If TriggerX.TriggerType = bdOnUseOnCreature Then
					' Determine if have a TargetCreature Statement
					HasTarget = False
					For	Each StmtX In TriggerX.Statements
						If StmtX.Statement = 31 Then ' TargetCreature
							HasTarget = True
							Exit For
						End If
					Next StmtX
					If HasTarget = False Then
						HasTarget = TargetCreature(CreatureTarget, 2, 0, False)
					End If
					If HasTarget = True Then
						NoFail = FireTrigger(ItemToUse, TriggerX)
						If NoFail = True Then
							NoFail = FireTriggers(CreatureTarget, bdPostUseOnCreature)
						End If
					End If
					Found = True
				End If
			Next TriggerX
			Found = True
		ElseIf HasTrigger(ItemToUse, bdOnUseInEncounter) = True Then 
			NoFail = FireTriggers(ItemToUse, bdOnUseInEncounter)
			Found = True
		ElseIf ItemToUse.Family = 1 Then 
			' DoorOpen with Key
			'UPGRADE_WARNING: Couldn't resolve default property of object Tome.MapY. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
			'UPGRADE_WARNING: Couldn't resolve default property of object Tome.MapX. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
            DoorOpen(tome.MapX, tome.MapY)
			Found = True
		ElseIf ItemToUse.Count > 1 Then 
			DialogDM("You split the " & ItemToUse.Name & " into two piles.")
			ItemX = CreatureWithTurn.AddItem
			ItemX.Copy(ItemToUse)
			ItemToUse.Count = Int(ItemToUse.Count / 2) + (ItemToUse.Count Mod 2)
			ItemToUse.Value = Int(ItemToUse.Value / 2) + (ItemToUse.Value Mod 2)
			ItemToUse.Weight = Int(ItemToUse.Weight / 2) + (ItemToUse.Weight Mod 2)
			ItemToUse.Bulk = Int(ItemToUse.Bulk / 2) + (ItemToUse.Bulk Mod 2)
			ItemX.Count = Int(ItemX.Count / 2)
			ItemX.Value = Int(ItemX.Value / 2)
			ItemX.Weight = Int(ItemX.Weight / 2)
			ItemX.Bulk = Int(ItemX.Bulk / 2)
			ItemX.IsReady = False
			ItemX.InvSpot = 0
			Found = True
		End If
		' Fire Post-UseInEncounter
		If HasTrigger(EncounterNow, bdPostUseInEncounter) Then
			CreatureNow = CreatureWithTurn
			ItemNow = ItemToUse
			NoFail = FireTriggers(EncounterNow, bdPostUseInEncounter)
			Found = True
		ElseIf Found = False Then 
			DialogDM("Nothing happens.")
		End If
		' If had a Use, deduct Action Points
		If Found = True And CreatureWithTurn.ActionPoints = OldActionPoints Then
			CreatureWithTurn.ActionPoints = CreatureWithTurn.ActionPoints - 5
		End If
	End Sub
	
	Private Sub SearchTraps()
		Dim NoFail, Found As Short
		Dim TriggerX As Trigger
		Dim ItemX As Item
		Found = False
		GlobalRemoveTrapChance = CreatureWithTurn.Agility + CreatureWithTurn.Will
		CreatureNow = CreatureWithTurn
		ItemNow = CreatureWithTurn.ItemInHand
		NoFail = FireTriggers(EncounterNow, bdPreSearchTraps)
		CreatureNow = CreatureWithTurn
		ItemNow = CreatureWithTurn.ItemInHand
		NoFail = FireTriggers(CreatureWithTurn, bdPreSearchTraps)
		If NoFail Then
			For	Each TriggerX In EncounterNow.Triggers
				If TriggerX.IsTrap Then
					If Int(Rnd() * 100) < GlobalRemoveTrapChance Then
						DialogDM("You found and removed the " & TriggerX.Name & " in this area.")
						EncounterNow.RemoveTrigger("T" & TriggerX.Index)
					End If
					Found = True
				End If
			Next TriggerX
			For	Each ItemX In EncounterNow.Items
				CreatureNow = CreatureWithTurn
				ItemNow = ItemX
				NoFail = FireTriggers(ItemX, bdPreSearchTraps)
				If NoFail Then
					For	Each TriggerX In ItemX.Triggers
						If TriggerX.IsTrap Then
							If Int(Rnd() * 100) < GlobalRemoveTrapChance Then
								DialogDM("You found and removed the " & TriggerX.Name & " on the " & ItemX.Name & ".")
								ItemX.RemoveTrigger("T" & TriggerX.Index)
							End If
							Found = True
						End If
					Next TriggerX
					CreatureNow = CreatureWithTurn
					ItemNow = ItemX
					NoFail = FireTriggers(ItemX, bdPostSearchTraps)
				End If
			Next ItemX
			CreatureNow = CreatureWithTurn
			ItemNow = CreatureWithTurn.ItemInHand
			NoFail = FireTriggers(CreatureNow, bdPostSearchTraps)
			CreatureNow = CreatureWithTurn
			ItemNow = CreatureWithTurn.ItemInHand
			NoFail = FireTriggers(EncounterNow, bdPostSearchTraps)
		End If
		If Not Found Then
			DialogDM("You don't find any traps.")
		End If
	End Sub
	
	Private Sub TargetCreatureClick(ByRef AtX As Short, ByRef AtY As Short, ByRef ButtonDown As Short)
		Dim c As Short
		Dim rc As Integer
		If PointIn(AtX, AtY, 331, 2, 18, 269) Then
			If ScrollBarClick(AtX, AtY, ButtonDown, picConvoList, 331, 2, 269, ScrollTop, ScrollList.Count(), 2) = True Then
				TargetCreatureShow()
			End If
		ElseIf PointIn(AtX, AtY, 12, 9, 275, 328) And ButtonDown = False Then 
			' Pick Creature in List
			c = Int((AtY - 9) / 87) + ScrollTop
			If c > 0 And c <= ScrollList.Count() Then
				ScrollSelect = c
				TargetCreatureShow()
				Call PlayClickSnd(modIOFunc.ClickType.ifClick)
			End If
		End If
	End Sub
	
	Private Sub TargetCreatureShow()
		Dim c, Pos, i As Short
		Dim rc As Integer
		Dim CreatureX As Creature
		Dim TriggerX As Trigger
		Pos = 0
		'UPGRADE_ISSUE: PictureBox method picConvoList.Cls was not upgraded. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="CC4C7EC0-C903-48FC-ACCC-81861D12DA4A"'
		picConvoList.Cls()
		'UPGRADE_ISSUE: PictureBox method picInvDrag.Cls was not upgraded. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="CC4C7EC0-C903-48FC-ACCC-81861D12DA4A"'
		picInvDrag.Cls()
		For c = ScrollTop To Least(ScrollTop + 2, ScrollList.Count())
			CreatureX = ScrollList.Item(c)
			' Show Picture
			LoadCreaturePic(CreatureX)
			'UPGRADE_ISSUE: PictureBox property picMisc.hdc was not upgraded. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="CC4C7EC0-C903-48FC-ACCC-81861D12DA4A"'
			'UPGRADE_ISSUE: PictureBox property picConvoList.hdc was not upgraded. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="CC4C7EC0-C903-48FC-ACCC-81861D12DA4A"'
			rc = BitBlt(picConvoList.hdc, 8, 8 + 87 * Pos, 74, 84, picMisc.hdc, 0, 36, SRCCOPY)
			'UPGRADE_ISSUE: PictureBox property picFaces.hdc was not upgraded. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="CC4C7EC0-C903-48FC-ACCC-81861D12DA4A"'
			'UPGRADE_ISSUE: PictureBox property picConvoList.hdc was not upgraded. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="CC4C7EC0-C903-48FC-ACCC-81861D12DA4A"'
			rc = BitBlt(picConvoList.hdc, 12, 12 + 87 * Pos, 66, 76, picFaces.hdc, bdFaceMin + CreatureX.Pic * 66, 0, SRCCOPY)
			' Show name and stats
			ShowText(picConvoList, 80, 12 + 87 * Pos, 250, 14, bdFontElixirWhite, (CreatureX.Name), True, False)
			If CreatureX.HPNow < 1 Then
				ShowText(picConvoList, 80, 28 + 87 * Pos, 250, 14, bdFontNoxiousWhite, "Lvl " & CreatureX.Level & " * Health DEAD", True, False)
			Else
				ShowText(picConvoList, 80, 28 + 87 * Pos, 250, 14, bdFontNoxiousWhite, "Lvl " & CreatureX.Level & " * Health " & CreatureX.HPNow & " of " & CreatureX.HPMax, True, False)
			End If
			' Show 3 skills
			i = 0
			For	Each TriggerX In CreatureX.Triggers
				If TriggerX.IsSkill = True Then
					ShowText(picConvoList, 80, 44 + 87 * Pos + i * 16, 250, 14, bdFontNoxiousWhite, (TriggerX.Name), True, False)
					i = i + 1
					If i > 2 Then
						Exit For
					End If
				End If
			Next TriggerX
			If i = 0 Then
				ShowText(picConvoList, 80, 44 + 87 * Pos, 250, 14, bdFontNoxiousWhite, "No Skills", True, False)
			End If
			' Show Select Box
			If ScrollSelect = c Then
				'UPGRADE_ISSUE: PictureBox property picFaces.hdc was not upgraded. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="CC4C7EC0-C903-48FC-ACCC-81861D12DA4A"'
				'UPGRADE_ISSUE: PictureBox property picConvoList.hdc was not upgraded. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="CC4C7EC0-C903-48FC-ACCC-81861D12DA4A"'
				rc = BitBlt(picConvoList.hdc, 12, 12 + 87 * Pos, 66, 76, picFaces.hdc, bdFaceSelect + 66, 0, SRCAND)
				'UPGRADE_ISSUE: PictureBox property picFaces.hdc was not upgraded. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="CC4C7EC0-C903-48FC-ACCC-81861D12DA4A"'
				'UPGRADE_ISSUE: PictureBox property picConvoList.hdc was not upgraded. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="CC4C7EC0-C903-48FC-ACCC-81861D12DA4A"'
				rc = BitBlt(picConvoList.hdc, 12, 12 + 87 * Pos, 66, 76, picFaces.hdc, bdFaceSelect, 0, SRCPAINT)
			End If
			Pos = Pos + 1
		Next c
		ScrollBarShow(picConvoList, 331, 2, 269, ScrollTop, ScrollList.Count() - 2, 0)
	End Sub
	
	Private Sub TargetCreatureRightClick()
		' Check for Skill or Spell
		If CreatureWithTurn.CombatAction2 > 0 Then
		Else
			
		End If
		'UPGRADE_WARNING: Couldn't resolve default property of object TargetCreatureSave(CreatureTarget). Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
		If TargetCreatureSave(CreatureTarget) = True Then
			
		End If
	End Sub
	
	Private Function TargetCreatureSave(ByRef CreatureTargeted As Creature) As Object
		'UPGRADE_NOTE: Text was upgraded to Text_Renamed. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="A9E4979A-37FA-4718-9994-97DD76ED70A7"'
		Dim i, c, SaveThrow As Short
		Dim Text_Renamed As String
		'UPGRADE_WARNING: Couldn't resolve default property of object TargetCreatureSave. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
		TargetCreatureSave = False
		If GlobalSaveStyle > 0 And CreatureTargeted.Family(0) = True Then
			For c = 8 To 13
				If (GlobalSaveStyle And (2 ^ c)) > 0 Then
					Select Case c
						Case 8 ' Lunacy
							SaveThrow = CreatureTargeted.Lunacy
							Text_Renamed = "Lunacy"
						Case 9 ' Revelry
							SaveThrow = CreatureTargeted.Revelry
							Text_Renamed = "Revelry"
						Case 10 ' Wrath
							SaveThrow = CreatureTargeted.Wrath
							Text_Renamed = "Wrath"
						Case 11 ' Pride
							SaveThrow = CreatureTargeted.Pride
							Text_Renamed = "Pride"
						Case 12 ' Greed
							SaveThrow = CreatureTargeted.Greed
							Text_Renamed = "Greed"
						Case 13 ' Lust
							SaveThrow = CreatureTargeted.Lust
							Text_Renamed = "Lust"
					End Select
					If CreatureTargeted.DMControlled = False Then
						DialogDice(CreatureTargeted, "Roll verses your " & Text_Renamed & " Vice.", 25, i)
						If i > SaveThrow Then
							DialogDM(CreatureTargeted.Name & " saves and is unaffected!")
							DialogHide()
							'UPGRADE_WARNING: Couldn't resolve default property of object TargetCreatureSave. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
							TargetCreatureSave = True
							Exit For
						End If
						DialogHide()
					ElseIf Int(Rnd() * 20) + 1 > SaveThrow Then 
						DialogDM(CreatureTargeted.Name & " saves and is unaffected!")
						DialogHide()
						'UPGRADE_WARNING: Couldn't resolve default property of object TargetCreatureSave. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
						TargetCreatureSave = True
						Exit For
					End If
				End If
			Next c
		End If
	End Function
	
	Public Function TargetCreature(ByRef CreatureTargeted As Creature, ByVal Range As Short, ByRef WithMotive As Short, ByRef AllowDead As Short) As Short
		'UPGRADE_NOTE: Text was upgraded to Text_Renamed. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="A9E4979A-37FA-4718-9994-97DD76ED70A7"'
		Dim SaveThrow, c, Found As Short
		Dim Text_Renamed As String
		Dim CreatureX As Creature
		Dim OldPointer As Short
		' Move 0 = Any, 1 = Party Only, 2 = Non-Party Only
		' If this is a DM controlled Creature, then Target with a Motive
		If CreatureWithTurn.DMControlled = True Then
			Found = False
			Select Case WithMotive
				Case 0 ' Any (Tome is first choice)
                    For Each CreatureX In tome.Creatures
                        'UPGRADE_WARNING: Couldn't resolve default property of object CombatRange(CreatureWithTurn, CreatureX.Col, CreatureX.Row). Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
                        If CombatRange(CreatureWithTurn, (CreatureX.Col), (CreatureX.Row)) <= Range And (Found = False Or Int(Rnd() * 2) < 1) And (CreatureX.HPNow > 0 Or AllowDead > 0) Then
                            CreatureTargeted = CreatureX
                            Found = True
                        End If
                    Next CreatureX
					' If can't find one in the Party, then try the Encounter
					If Found = False Then
						For	Each CreatureX In EncounterNow.Creatures
							'UPGRADE_WARNING: Couldn't resolve default property of object CombatRange(CreatureWithTurn, CreatureX.Col, CreatureX.Row). Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
							If CombatRange(CreatureWithTurn, (CreatureX.Col), (CreatureX.Row)) <= Range And (Found = False Or Int(Rnd() * 2) < 1) And (CreatureX.HPNow > 0 Or AllowDead > 0) Then
								CreatureTargeted = CreatureX
								Found = True
							End If
						Next CreatureX
					End If
				Case 1 ' PartyOnly
                    For Each CreatureX In tome.Creatures
                        'UPGRADE_WARNING: Couldn't resolve default property of object CombatRange(CreatureWithTurn, CreatureX.Col, CreatureX.Row). Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
                        If CombatRange(CreatureWithTurn, (CreatureX.Col), (CreatureX.Row)) <= Range And (CreatureX.HPNow > 0 Or AllowDead > 0) Then
                            CreatureTargeted = CreatureX
                            Found = True
                            Exit For
                        End If
                    Next CreatureX
				Case 2 ' NonPartyOnly
					For	Each CreatureX In EncounterNow.Creatures
						'UPGRADE_WARNING: Couldn't resolve default property of object CombatRange(CreatureWithTurn, CreatureX.Col, CreatureX.Row). Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
						If CombatRange(CreatureWithTurn, (CreatureX.Col), (CreatureX.Row)) <= Range And (CreatureX.HPNow > 0 Or AllowDead > 0) Then
							CreatureTargeted = CreatureX
							Found = True
							Exit For
						End If
					Next CreatureX
			End Select
			' Show who is being targeted if this is in combat
			TargetCreature = Found
			If picGrid.Visible = True And Found Then
				CombatDrawAttack(CreatureWithTurn, CreatureTargeted, False)
			End If
		Else
			'UPGRADE_WARNING: Screen property Screen.MousePointer has a new behavior. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6BA9B8D2-2A32-4B6E-8D36-44949974A5B4"'
			OldPointer = System.Windows.Forms.Cursor.Current
			'UPGRADE_WARNING: Screen property Screen.MousePointer has a new behavior. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6BA9B8D2-2A32-4B6E-8D36-44949974A5B4"'
			System.Windows.Forms.Cursor.Current = System.Windows.Forms.Cursors.Default
			Select Case WithMotive
				Case 0 ' Any
					MenuNow = bdMenuTargetAny
				Case 1 ' PartyOnly
					MenuNow = bdMenuTargetParty
				Case 2 ' NonPartyOnly
					MenuNow = bdMenuTargetCreature
			End Select
			' Else you can let the Player do the Targeting.
			If picGrid.Visible = True Then
				' In combat, do normal highlight to select
				TargetCreatureRange = Range
				TargetAllowDead = (AllowDead > 0)
				ConvoAction = -1
				MessageShow("Target a Creature (Esc to Cancel)", 0)
				picGrid.Focus()
				Do Until ConvoAction > -1
					System.Windows.Forms.Application.DoEvents()
				Loop 
				CombatDraw()
				TargetCreatureRange = 0
				TargetAllowDead = 0
				If ConvoAction > 0 Then
					CreatureTargeted = CreatureTarget
					TargetCreature = True
					' Show ToHit number
					CombatDrawAttack(CreatureWithTurn, CreatureTargeted, False)
				Else
					TargetCreature = False
				End If
				MenuNow = bdMenuDefault
			Else
				DialogSetUp(modGameGeneral.DLGTYPE.bdDlgCreatureList)
				DialogSetButton(1, "Done")
				DialogSetButton(2, "Cancel")
				ScrollList = New Collection
				If WithMotive = 0 Or WithMotive = 2 Then
					For	Each CreatureX In EncounterNow.Creatures
						If CreatureX.HPNow > 0 Or AllowDead > 0 Then
							ScrollList.Add(CreatureX)
						End If
					Next CreatureX
				End If
				If WithMotive = 0 Or WithMotive = 1 Then
                    For Each CreatureX In tome.Creatures
                        If CreatureX.HPNow > 0 Or AllowDead > 0 Then
                            ScrollList.Add(CreatureX)
                        End If
                    Next CreatureX
				End If
				ScrollTop = 1 : ScrollSelect = 1
				TargetCreatureShow()
				DialogShow("DM", "Target Creature")
				DialogHide()
				If ConvoAction = 1 And ScrollList.Count() >= ScrollSelect Then
					CreatureTargeted = ScrollList.Item(ScrollSelect)
					TargetCreature = True
				Else
					TargetCreature = False
				End If
			End If
			'UPGRADE_ISSUE: Screen property Screen.MousePointer does not support custom mousepointers. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="45116EAB-7060-405E-8ABE-9DBB40DC2E86"'
			'UPGRADE_WARNING: Screen property Screen.MousePointer has a new behavior. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6BA9B8D2-2A32-4B6E-8D36-44949974A5B4"'
			System.Windows.Forms.Cursor.Current = OldPointer
		End If
		' Check for a Saving Throw (if allowed and creature is sentient)
		'UPGRADE_WARNING: Couldn't resolve default property of object TargetCreatureSave(CreatureTargeted). Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
		If TargetCreature = True And TargetCreatureSave(CreatureTargeted) = True Then
			TargetCreature = False
		End If
	End Function
	
	Private Sub TargetItemClick(ByRef AtX As Short, ByRef AtY As Short, ByRef ButtonDown As Short)
		Dim c As Short
		Dim rc As Integer
		If PointIn(AtX, AtY, 331, 2, 18, 269) Then
			If ScrollBarClick(AtX, AtY, ButtonDown, picConvoList, 331, 2, 269, ScrollTop, ScrollList.Count(), 6) = True Then
				TargetItemShow()
			End If
		ElseIf PointIn(AtX, AtY, 12, 9, 275, 328) And ButtonDown = False Then 
			' Pick Item in List
			c = Int((AtY - 9) / 37) + ScrollTop
			If c > 0 And c <= ScrollList.Count() Then
				ScrollSelect = c
				TargetItemShow()
				Call PlayClickSnd(modIOFunc.ClickType.ifClick)
			End If
		End If
	End Sub
	
	Private Sub TargetItemShow()
		Dim Pos, c As Short
		Dim rc As Integer
		Dim ItemX As Item
		Pos = 0
		'UPGRADE_ISSUE: PictureBox method picConvoList.Cls was not upgraded. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="CC4C7EC0-C903-48FC-ACCC-81861D12DA4A"'
		picConvoList.Cls()
		'UPGRADE_ISSUE: PictureBox method picInvDrag.Cls was not upgraded. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="CC4C7EC0-C903-48FC-ACCC-81861D12DA4A"'
		picInvDrag.Cls()
		For c = ScrollTop To Least(ScrollTop + 6, ScrollList.Count())
			ItemX = ScrollList.Item(c)
			' Show Picture
			LoadItemPic(ItemX)
			'UPGRADE_ISSUE: PictureBox property picItem.hdc was not upgraded. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="CC4C7EC0-C903-48FC-ACCC-81861D12DA4A"'
			'UPGRADE_ISSUE: PictureBox property picConvoList.hdc was not upgraded. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="CC4C7EC0-C903-48FC-ACCC-81861D12DA4A"'
			rc = BitBlt(picConvoList.hdc, 12, 9 + 37 * Pos, ItemPicWidth(ItemX.Pic) / 3, ItemPicHeight(ItemX.Pic) / 3, picItem.hdc, 64 * ItemX.Pic - 32, 96 * 2, SRCAND)
			'UPGRADE_ISSUE: PictureBox property picItem.hdc was not upgraded. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="CC4C7EC0-C903-48FC-ACCC-81861D12DA4A"'
			'UPGRADE_ISSUE: PictureBox property picConvoList.hdc was not upgraded. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="CC4C7EC0-C903-48FC-ACCC-81861D12DA4A"'
			rc = BitBlt(picConvoList.hdc, 12, 9 + 37 * Pos, ItemPicWidth(ItemX.Pic) / 3, ItemPicHeight(ItemX.Pic) / 3, picItem.hdc, 64 * ItemX.Pic - 64, 96 * 2, SRCPAINT)
			' Show Description
			If ScrollSelect = c Then
				ShowText(picConvoList, 53, 20 + 37 * Pos, 275, 14, bdFontNoxiousGold, (ItemX.Name), False, False)
			Else
				ShowText(picConvoList, 53, 20 + 37 * Pos, 275, 14, bdFontNoxiousWhite, (ItemX.Name), False, False)
			End If
			Pos = Pos + 1
		Next c
		ScrollBarShow(picConvoList, 331, 2, 269, ScrollTop, ScrollList.Count() - 6, 0)
	End Sub
	
	Public Function TargetItem(ByRef IntoItem As Item) As Short
		Dim ItemX As Item
		Dim CreatureX As Creature
		Dim Found As Short
		' Setup with List
		DialogSetUp(modGameGeneral.DLGTYPE.bdDlgItemList)
		DialogSetButton(1, "Done")
		DialogSetButton(2, "Cancel")
		ScrollList = New Collection
		ScrollTop = 1
		For	Each ItemX In CreatureNow.Items
			ScrollList.Add(ItemX)
		Next ItemX
		' Check for Guarding creatures
		Found = False
		For	Each CreatureX In EncounterNow.Creatures
			If CreatureX.Guard = True And CreatureX.AllowedTurn = True Then
				Found = True
				Exit For
			End If
		Next CreatureX
		If Not Found Then
			For	Each ItemX In EncounterNow.Items
				ScrollList.Add(ItemX)
			Next ItemX
		End If
		ScrollTop = 1 : ScrollSelect = 1
		TargetItemShow()
		DialogShow("DM", "Pick an Item")
		DialogHide()
		If ConvoAction = 1 And ScrollList.Count() > 0 Then ' [Titi 2.4.7] crash if nothing to appraise!
			IntoItem = ScrollList.Item(ScrollSelect)
			TargetItem = True
		Else
			TargetItem = False
		End If
	End Function
	
	Public Sub MessageShow(ByRef InText As String, ByRef Index As Short)
		Dim c As Short
		Dim rc As Integer
		If picGrid.Visible = False Or Index <> 0 Then
			BorderDrawBottom()
			Select Case Index
				Case 0 ' Center
					ShowText(Me, 0, Me.ClientRectangle.Height - 14, Me.ClientRectangle.Width, 14, bdFontSmallWhite, InText, True, False)
				Case 1 ' Right
					ShowText(Me, 0, Me.ClientRectangle.Height - 14, Me.ClientRectangle.Width, 14, bdFontSmallWhite, InText, 1, False)
				Case 2 ' Left
					ShowText(Me, 0, Me.ClientRectangle.Height - 14, Me.ClientRectangle.Width, 14, bdFontSmallWhite, InText, False, False)
			End Select
			Me.Refresh()
		Else
			CombatMessage(InText)
		End If
	End Sub
	
	Private Sub MessageClear()
		Dim c As Short
		For c = 0 To 2
			MessageShow("", c)
		Next c
	End Sub
	
	'UPGRADE_NOTE: Map was upgraded to Map_Renamed. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="A9E4979A-37FA-4718-9994-97DD76ED70A7"'
	Private Sub MessageEncounterName(ByRef AtX As Short, ByRef AtY As Short)
		Dim Map_Renamed As Object
		'UPGRADE_NOTE: my was upgraded to my_Renamed. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="A9E4979A-37FA-4718-9994-97DD76ED70A7"'
		Dim Found, mx, my_Renamed, c As Short
		Dim rc As Integer
		Dim EncounterX As Encounter
		If Int(AtY / 16) Mod 3 = 1 Or (Int(AtY / 8) Mod 3 = 1 And Int((AtX - 24) / 48) Mod 2 = 0) Then
			'UPGRADE_WARNING: Couldn't resolve default property of object Map.Left. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
			mx = Map.Left + Int(AtX / 96) - Int(AtY / 48)
			'UPGRADE_WARNING: Couldn't resolve default property of object Map.Top. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
			my_Renamed = Map.Top + Int(AtX / 96) + Int(AtY / 48)
		Else
			'UPGRADE_WARNING: Couldn't resolve default property of object Map.Left. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
			mx = Map.Left + Int((AtX - 48) / 96) - Int((AtY - 24) / 48)
			'UPGRADE_WARNING: Couldn't resolve default property of object Map.Top. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
			my_Renamed = Map.Top + Int((AtX + 48) / 96) + Int((AtY - 24) / 48)
		End If
		Found = False
		'UPGRADE_WARNING: Couldn't resolve default property of object Map.Height. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
		'UPGRADE_WARNING: Couldn't resolve default property of object Map.Width. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
		If IsBetween(mx, 0, Map.Width) And IsBetween(my_Renamed, 0, Map.Height) Then
			'UPGRADE_WARNING: Couldn't resolve default property of object Map.EncPointer. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
			If Map.EncPointer(mx, my_Renamed) > 0 Then
				'UPGRADE_WARNING: Couldn't resolve default property of object Map.EncPointer. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
				'UPGRADE_WARNING: Couldn't resolve default property of object Map.Encounters. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
				EncounterX = Map.Encounters("E" & Map.EncPointer(mx, my_Renamed))
				If EncounterX.UseHint = True And EncounterX.HaveEntered = True Then
					BorderDrawBottom()
					ShowText(Me, 0, Me.ClientRectangle.Height - bdIntHeight, Me.ClientRectangle.Width, 14, bdFontNoxiousWhite, (EncounterX.Name), True, False)
					Me.Refresh()
				End If
			End If
		End If
	End Sub
	
	Private Function ReadyItem(ByRef ToChar As Creature, ByRef ItemToReady As Item) As Short
		Dim c, NoFail As Short
		Dim ItemX As Item
		Dim WearType(11) As Short
		' If in combat, then this Fails if equipping anything other than a OneHand Item.
		If picGrid.Visible = True And ItemToReady.WearType < 10 Then
			If ToChar.DMControlled = False Then
				DialogDM("You can only equip weapons during combat.")
			End If
			ReadyItem = False
			Exit Function
		End If
		' Fire Triggers on ReadyItem
		CreatureNow = ToChar
		ItemNow = ItemToReady
		NoFail = FireTriggers(ToChar, bdPreReady)
		If NoFail Then
			NoFail = FireTriggers(ItemToReady, bdPreReady)
			If NoFail Then
				' Count number of Items already worn
				For	Each ItemX In ToChar.Items
					If ItemX.IsReady = True Then
						WearType(ItemX.WearType) = WearType(ItemX.WearType) + ItemX.Count
					End If
				Next ItemX
				' Check if will be wearing too many
				ReadyItem = True
				Select Case ItemToReady.WearType
					Case 5, 10 ' Shield, OneHand
						If WearType(ItemToReady.WearType) > 0 Or WearType(11) > 0 Then
							ReadyItem = False
						End If
					Case 0, 1, 6, 7, 8 ' Body, Helm, Boots, Necklace, Belt
						' Can only wear one
						If WearType(ItemToReady.WearType) > 0 Then
							ReadyItem = False
						End If
					Case 2, 3, 4 ' Glove, Wrist, Backpack
						' Can wear only two
						If WearType(ItemToReady.WearType) > 1 Then
							ReadyItem = False
						End If
					Case 9 ' Rings
						' Can wear only eight
						If WearType(9) > 7 Then
							ReadyItem = False
						End If
					Case 11 ' TwoHand
						' Can only wear without OneHand, Shield or TwoHand
						If WearType(5) > 0 Or WearType(10) > 0 Or WearType(11) > 0 Then
							ReadyItem = False
						End If
				End Select
				' If conflicts then try to UnReady the offending item
				If ReadyItem = False Then
					' Loop through all ready itm
					For	Each ItemX In ToChar.Items
						If ItemX.IsReady = True And (ItemX.WearType = ItemToReady.WearType Or (ItemToReady.WearType = 11 And ItemX.WearType = 10 And WearType(5) = 0) Or (ItemToReady.WearType = 10 And ItemX.WearType = 11 And WearType(5) = 0)) Then
							ReadyItem = UnReadyItem(ToChar, ItemX)
							Exit For
						End If
					Next ItemX
				End If
				' Ready the Item
				If ReadyItem = True Then
					ItemToReady.IsReady = True
					CreatureNow = ToChar
					ItemNow = ItemToReady
					NoFail = FireTriggers(ToChar, bdPostReady)
					CreatureNow = ToChar
					ItemNow = ItemToReady
					NoFail = FireTriggers(ItemToReady, bdPostReady)
				End If
			Else
				ReadyItem = False
			End If
		Else
			ReadyItem = False
		End If
	End Function
	
	Public Function UnReadyItem(ByRef FromChar As Creature, ByRef ItemToUnReady As Item) As Short
		Dim NoFail As Short
		' This fails if in Combat and UnReady anything other than a OneHand or TwoHand Item
		If picGrid.Visible = True And ItemToUnReady.WearType < 10 Then
			DialogDM("You can only equip weapons during combat.")
			UnReadyItem = False
			Exit Function
		End If
		' Set up for Triggers
		If ItemToUnReady.IsReady Then
			CreatureNow = FromChar
			ItemNow = ItemToUnReady
			NoFail = FireTriggers(FromChar, bdPreUnReady)
			CreatureNow = FromChar
			ItemNow = ItemToUnReady
			NoFail = FireTriggers(ItemToUnReady, bdPreUnReady)
			If NoFail Then
				If FromChar.Bulk + ItemToUnReady.Bulk > FromChar.Agility * 8 Then 'was 100 - now depends on Agility
					' [Titi 2.4.8] added check bulk
					DialogDM("The " & ItemToUnReady.Name & " is too bulky for " & FromChar.Name & " to carry.")
					UnReadyItem = False
				Else
					ItemToUnReady.IsReady = False
					ItemToUnReady.InvSpot = 0
					CreatureNow = FromChar
					ItemNow = ItemToUnReady
					NoFail = FireTriggers(FromChar, bdPostUnReady)
					CreatureNow = FromChar
					ItemNow = ItemToUnReady
					NoFail = FireTriggers(ItemToUnReady, bdPostUnReady)
					UnReadyItem = True
				End If
			Else
				UnReadyItem = False
			End If
		Else
			UnReadyItem = True
		End If
	End Function
	
	Public Function CombineWithAnything(ByRef ObjectX As Object, ByRef ItemX As Item) As Short
		' This combines ItemX with some Item in ObjectX (if it can find one)
		Dim ItemZ As Item
		Dim rc3, rc1, rc2, rc4 As Integer
		CombineWithAnything = False
		If ItemX.CanCombine = True Then
			'UPGRADE_WARNING: Couldn't resolve default property of object ObjectX.Items. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
			For	Each ItemZ In ObjectX.Items
				If ItemZ.CanCombine = True And ItemX.NameText = ItemZ.NameText Then
					rc1 = CInt(ItemZ.Count) + CInt(ItemX.Count)
					rc2 = CInt(ItemZ.Value) + CInt(ItemX.Value)
					rc3 = CInt(ItemZ.Weight) + CInt(ItemX.Weight)
					rc4 = CInt(ItemZ.Bulk) + CInt(ItemX.Bulk)
					If rc1 < 32700 And rc2 < 32700 And rc3 < 32700 And rc4 < 32700 Then
						ItemZ.Count = rc1
						ItemZ.Value = rc2
						ItemZ.Weight = rc3
						ItemZ.Bulk = rc4
						CombineWithAnything = True
						Exit For
					End If
				End If
				If ItemZ.Capacity > 0 Then
					If CombineWithAnything(ItemZ, ItemX) = True Then
						CombineWithAnything = True
						Exit For
					End If
				End If
			Next ItemZ
		End If
	End Function
	
	Private Function CombineItemsTriggers(ByRef FromChar As Creature, ByRef FromItem As Item, ByRef ToItem As Item) As Short
		Dim c, NoFail As Short
		Dim ItemX As Item
		' Check if have any Combine Triggers to fire
		If HasTrigger(FromItem, bdPreCombine) Or HasTrigger(ToItem, bdPreCombine) Or HasTrigger(FromItem, bdOnCombine) Or HasTrigger(ToItem, bdOnCombine) Or HasTrigger(FromItem, bdPostCombine) Or HasTrigger(ToItem, bdPostCombine) Then
			CombineItemsTriggers = True
			' Create FromItem on FromCharacter
			ItemX = FromChar.AddItem
			ItemX.Copy(FromItem)
			' Fire PreCombine Triggers
			CreatureNow = FromChar
			ItemNow = ItemX : ItemTarget = ToItem
			NoFail = FireTriggers(ItemX, bdPreCombine)
			CreatureNow = FromChar
			ItemNow = ToItem : ItemTarget = ItemX
			c = FireTriggers(ToItem, bdPreCombine)
			If NoFail And c Then
				' Fire OnCombine Triggers
				CreatureNow = FromChar
				ItemNow = ItemX : ItemTarget = ToItem
				NoFail = FireTriggers(ItemX, bdOnCombine)
				CreatureNow = FromChar
				ItemNow = ToItem : ItemTarget = ItemX
				NoFail = FireTriggers(ToItem, bdOnCombine)
				If NoFail Then
					' Fire PostCombine Triggers
					CreatureNow = FromChar
					ItemNow = ItemX : ItemTarget = ToItem
					NoFail = FireTriggers(ItemX, bdPostCombine)
					CreatureNow = FromChar
					ItemNow = ToItem : ItemTarget = ItemX
					NoFail = FireTriggers(ToItem, bdPostCombine)
				End If
			End If
			' Refresh the Inventory screen
			InventoryShow(bdInvItems)
		End If
	End Function
	
	Private Function TradeItem(ByRef FromChar As Creature, ByRef ToChar As Creature, ByVal FromList As Short, ByVal ToList As Short, ByVal IntoIndex As Short, ByVal Spot As Short, ByRef ItemToTrade As Item) As Short
		Dim c, NoFail As Short
		Dim ItemX, ItemContainerX As Item
		Dim blnWasEquipped As Boolean
		blnWasEquipped = ItemToTrade.IsReady
		' If moving around on same list, just move the InvSpot
		If FromList = ToList And ToList <> bdInvWear And (IntoIndex < 1 Or IntoIndex = ItemToTrade.Index) Then
			ItemToTrade.InvSpot = Spot
			Call PlayClickSnd(modIOFunc.ClickType.ifClickDrop)
			TradeItem = True
			Exit Function
		Else
			ItemToTrade.InvSpot = 0
		End If
		' Check for Pre-Take if Search window is open
		If picSearch.Visible = True Then
			' Items in an Encounter are *always* not equipped
			ItemToTrade.IsReady = False
			' Fire Take Triggers
			CreatureTarget = ToChar
			CreatureNow = ToChar
			If FireTriggers(EncounterNow, bdPreTake) = False Then
				TradeItem = False
				Exit Function
			ElseIf FireTriggers(ItemToTrade, bdPreTake) = False Then 
				TradeItem = False
				Exit Function
			End If
		End If
		' If taking off, unready (if necessary)
		If FromList = bdInvWear Then
			If UnReadyItem(FromChar, ItemToTrade) = False Then
				TradeItem = False
				Exit Function
			End If
		End If
		' If putting on, ready the Item.
		If ToList = bdInvWear Then
			TradeItem = ReadyItem(ToChar, ItemToTrade)
			Call PlayClickSnd(modIOFunc.ClickType.ifClickDrop)
			Exit Function
		End If
		' Check if too heavy or bulky
		If ToList = bdInvObjects Or ToList = bdInvParty Then '[Titi 2.4.8] added the second part of the test
			' Check weight
			If ToChar.Weight + ItemToTrade.Weight > ToChar.MaxWeight Then
				If blnWasEquipped = False Or ToChar.Index <> FromChar.Index Then ' if the item was equipped, its weight is counted twice in the above test! [Titi 2.4.8]
					DialogDM("The " & ItemToTrade.Name & " weighs too much for " & ToChar.Name & " to carry.")
					TradeItem = False
					Exit Function
				End If
			End If
			' Check bulk
			If ToChar.Bulk + ItemToTrade.Bulk > ToChar.Agility * 8 Then ' was 100, now depends on Agility
				' [Titi 2.4.8] item cannot be unequipped if too bulky - check done in function UnReadyItem()
				If blnWasEquipped = False Or ToChar.Index <> FromChar.Index Then
					DialogDM("The " & ItemToTrade.Name & " is too bulky for " & ToChar.Name & " to carry.")
					TradeItem = False
					Exit Function
				End If
			End If
		End If
		' If moving into a container, check for capacity of target container
		If IntoIndex > 0 Then
			' Set up the target container
			If ToList = bdInvContainer Then
				ItemContainerX = InvContainer.Items.Item("I" & IntoIndex)
			ElseIf ToList = bdInvEncounter Then 
				ItemContainerX = EncounterNow.Items.Item("I" & IntoIndex)
			Else
				ItemContainerX = ToChar.Items.Item("I" & IntoIndex)
			End If
			' Check bulk
			If ItemContainerX.Capacity > 0 Then
				If ItemContainerX.Full + ItemToTrade.Bulk > ItemContainerX.Capacity Then
					DialogDM("The " & ItemToTrade.Name & " is too bulky to fit in the " & ItemContainerX.Name & ".")
					TradeItem = False
					Exit Function
				End If
			End If
			' Check if container is already open (as a container box)
			If ItemContainerX.Selected = True Then
				TradeItem = False
				Exit Function
			End If
		End If
		' If pass all checks, remove the Item from where it was: Container, Encounter or Creature
		If FromList = bdInvContainer Then
			InvContainer.RemoveItem("I" & ItemToTrade.Index)
		ElseIf FromList = bdInvEncounter Then 
			EncounterNow.RemoveItem("I" & ItemToTrade.Index)
		Else
			FromChar.RemoveItem("I" & ItemToTrade.Index)
		End If
		' Combine the Item with the target item
		If IntoIndex > 0 Then
			If ItemContainerX.Capacity > 0 Then
				If CombineWithAnything(ItemContainerX, ItemToTrade) = False Then
					ItemContainerX.AddItem.Copy(ItemToTrade)
				End If
				TradeItem = True
				Exit Function
			End If
			If CombineItemsTriggers(ToChar, ItemToTrade, ItemContainerX) = True Then
				TradeItem = True
				Exit Function
			End If
		End If
		' Create the item in the new location
		If ToList = bdInvContainer Then
			If CombineWithAnything(InvContainer, ItemToTrade) = False Then
				InvContainer.AddItem.Copy(ItemToTrade)
			End If
		ElseIf ToList = bdInvEncounter Then 
			If CombineWithAnything(EncounterNow, ItemToTrade) = False Then
				EncounterNow.AddItem.Copy(ItemToTrade)
			End If
		Else
			If CombineWithAnything(ToChar, ItemToTrade) = False Then
				ToChar.AddItem.Copy(ItemToTrade)
			End If
		End If
		Call PlayClickSnd(modIOFunc.ClickType.ifClickDrop)
		TradeItem = True
	End Function
	
	Private Sub CursorToGrid(ByVal CursorAtX As Short, ByVal CursorAtY As Short, ByRef Row As Short, ByRef Col As Short)
		' Returns number of char(0-4) or mons(5-14) at (x,y). Also sets x, y, Row, Col.
		Dim c, GridAdjust As Short
		Dim GridWidth As Double
		' Adjust where Cursor is now based on CommonScale
		CursorAtX = Int(CursorAtX / CommonScale)
		CursorAtY = Int(CursorAtY / CommonScale)
		' Calculate Grid Row
		Row = Least(bdCombatHeight - Int((CursorAtY - bdCombatTop) / bdCombatGridHeight) - 1, bdCombatHeight)
		' Set GridWidth and GridAdjust to X
		GridWidth = bdCombatGridWidth
		GridAdjust = bdCombatLeft + (bdCombatGridWidth / 2) * (Row Mod 2)
		' Calculate Grid Col
		Col = Least(Greatest(Int((CursorAtX - GridAdjust) / bdCombatGridWidth), 0), bdCombatWidth)
		' Do final adjustments to CursorAtX and CursorAtY (bound to grid)
		CursorAtY = (Int((bdCombatHeight - Row) * bdCombatGridHeight) + bdCombatTop) * CommonScale
		CursorAtX = Int(Col * GridWidth + GridAdjust) * CommonScale
	End Sub
	
	Private Sub SetMapCursor(ByVal ScreenX As Short, ByVal ScreenY As Short, ByRef MapLeft As Short, ByRef MapTop As Short, ByRef ToScreenX As Short, ByRef ToScreenY As Short, ByRef ToMapX As Short, ByRef ToMapY As Short)
		' Converts any ScreenX, ScreenY to true blue MapX and MapY
		If Int(ScreenY / 16) Mod 3 = 1 Or (Int(ScreenY / 8) Mod 3 = 1 And Int((ScreenX - 24) / 48) Mod 2 = 0) Then
			ToMapX = MapLeft + Int(ScreenX / 96) - Int(ScreenY / 48)
			ToMapY = MapTop + Int(ScreenX / 96) + Int(ScreenY / 48)
			ToScreenX = Int(ScreenX / 96) * 96
			ToScreenY = Int(ScreenY / 48) * 48
		Else
			ToMapX = MapLeft + Int((ScreenX - 48) / 96) - Int((ScreenY - 24) / 48)
			ToMapY = MapTop + Int((ScreenX + 48) / 96) + Int((ScreenY - 24) / 48)
			ToScreenX = Int((ScreenX - 48) / 96) * 96 + 48
			ToScreenY = Int((ScreenY - 24) / 48) * 48 + 24
		End If
	End Sub
	
	Public Function CombatMessage(ByRef InText As String) As Object
		Dim c As Short
		CreatureWithTurn.MsgQueTop = CreatureWithTurn.MsgQueTop + 1
		' If moved beyond top, scroll off
		If CreatureWithTurn.MsgQueTop > 7 Then
			For c = 0 To 6
				CreatureWithTurn.MsgQue(c) = CreatureWithTurn.MsgQue(c + 1)
			Next c
			CreatureWithTurn.MsgQueTop = CreatureWithTurn.MsgQueTop - 1
		End If
		CreatureWithTurn.MsgQue((CreatureWithTurn.MsgQueTop)) = InText
		If picGrid.Visible = True Then
			CombatDraw()
		End If
	End Function
	
	Public Function CombatTarget(ByRef CreatureX As Creature, ByRef Motive As Short) As Creature
		Dim i, c, Found As Short
		' Find a Target
		i = 0 : Found = False
		For c = 0 To CombatTurn
			If CreatureX.Friendly <> Turns(c).Ref.Friendly And Turns(c).Ref.HPNow > 0 Then
				Select Case Motive
					Case MOVEMOTIVE.bdMoveClosest
						'UPGRADE_WARNING: Couldn't resolve default property of object CombatRange(Turns(i).Ref, CreatureX.Col, CreatureX.Row). Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
						'UPGRADE_WARNING: Couldn't resolve default property of object CombatRange(Turns(c).Ref, CreatureX.Col, CreatureX.Row). Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
						If (CombatRange(Turns(c).Ref, (CreatureX.Col), (CreatureX.Row)) < CombatRange(Turns(i).Ref, (CreatureX.Col), (CreatureX.Row))) Or Found = False Then
							i = c
							Found = True
						End If
					Case MOVEMOTIVE.bdMoveFarthest
						'UPGRADE_WARNING: Couldn't resolve default property of object CombatRange(Turns(i).Ref, CreatureX.Col, CreatureX.Row). Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
						'UPGRADE_WARNING: Couldn't resolve default property of object CombatRange(Turns(c).Ref, CreatureX.Col, CreatureX.Row). Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
						If (CombatRange(Turns(c).Ref, (CreatureX.Col), (CreatureX.Row)) > CombatRange(Turns(i).Ref, (CreatureX.Col), (CreatureX.Row))) Or Found = False Then
							i = c
							Found = True
						End If
					Case MOVEMOTIVE.bdMoveStrongest
						If Turns(c).Ref.HPNow > Turns(i).Ref.HPNow Or Found = False Then
							i = c
							Found = True
						End If
					Case MOVEMOTIVE.bdMoveWeakest
						If Turns(c).Ref.HPNow < Turns(i).Ref.HPNow Or Found = False Then
							i = c
							Found = True
						End If
					Case MOVEMOTIVE.bdMoveRandom
						Found = True
						If i = 0 Then
							i = c
						ElseIf Int(Rnd() * 100) < 25 Then 
							i = c
							Exit For
						End If
				End Select
			End If
		Next c
		' Acquire target defaults to failed
		Target = -1
		If Found = True Then
			CombatTarget = Turns(i).Ref
			' Configure Target pointer to Turns() to new CreatureTarget
			CreatureTarget = CombatTarget
			For c = 0 To CombatTurn
				If Turns(c).Ref.Name = CreatureTarget.Name And Turns(c).Ref.Index = CreatureTarget.Index Then
					Target = c
					Exit For
				End If
			Next c
		Else
			CombatTarget = CreatureX
		End If
	End Function
	
	Private Sub CombatMonster()
		Dim i, c, OldActionPoints As Short
		Dim CreatureX As Creature
		Dim ItemX As Item
		Dim TriggerX As Trigger
		Dim Dmg, Attack, NoFail As Short
		Dim DiceCnt, DiceType As Short
		Dim Range As Short
		' Continue attacking until out of ActionPoints
		Do Until CreatureWithTurn.ActionPoints < 1
			System.Windows.Forms.Application.DoEvents()
			' Find the closest live Target
			CreatureTarget = CombatTarget(CreatureWithTurn, MOVEMOTIVE.bdMoveClosest)
			' If failed to find a Target, Monster stops attacking
			If Target < 0 Then
				Exit Do
			End If
			' Choose Monster attack by looping through Skills which are of Type On-Attack
			OldActionPoints = CreatureWithTurn.ActionPoints
			c = 0
			For	Each TriggerX In CreatureWithTurn.Triggers
				If TriggerX.TriggerType = bdOnAttack Then
					c = c + 1
				End If
			Next TriggerX
			If c > 0 Then
				c = Int(Rnd() * c) + 1 : i = 0
				For	Each TriggerX In CreatureWithTurn.Triggers
					If TriggerX.TriggerType = bdOnAttack Then
						i = i + 1
						If i = c Then
							Exit For
						End If
					End If
				Next TriggerX
				CreatureNow = CreatureWithTurn
				ItemNow = CreatureWithTurn.ItemInHand
				' [Titi 2.4.6] even without ammo, monsters continued to shoot! - case of the On-Attack triggers
				If CreatureWithTurn.ItemInHand.IsShooter = True Then
					If CreatureWithTurn.ItemAmmo.Index = CreatureWithTurn.ItemInHand.Index And CreatureWithTurn.ItemAmmo.Name = CreatureWithTurn.ItemInHand.Name Then
						CombatMessage(CreatureWithTurn.Name & " drops the " & CreatureWithTurn.ItemInHand.Name & " due to lack of " & CreatureWithTurn.ItemAmmo.Name)
						NoFail = DropItem(CreatureWithTurn, bdInvEncounter, 0, (CreatureWithTurn.ItemInHand))
					End If
				End If
				NoFail = FireTrigger(CreatureWithTurn, TriggerX)
			ElseIf CreatureWithTurn.IsInanimate = False Then 
				' [Titi 2.4.6] even without ammo, monsters continued to shoot! - case of the standard combat
				If CreatureWithTurn.ItemInHand.IsShooter = True Then
					If CreatureWithTurn.ItemAmmo.Index = CreatureWithTurn.ItemInHand.Index And CreatureWithTurn.ItemAmmo.Name = CreatureWithTurn.ItemInHand.Name Then
						CombatMessage(CreatureWithTurn.Name & " drops the " & CreatureWithTurn.ItemInHand.Name & " due to lack of " & CreatureWithTurn.ItemAmmo.Name)
						NoFail = DropItem(CreatureWithTurn, bdInvEncounter, 0, (CreatureWithTurn.ItemInHand))
					End If
				End If
				' Attack with a weapon
				CombatAttack(12, "", 0, 0)
			End If
			If CreatureWithTurn.ActionPoints = OldActionPoints Then
				CreatureWithTurn.ActionPoints = CreatureWithTurn.ActionPoints - 10
			End If
			System.Windows.Forms.Application.DoEvents()
			' Fire Post-Attack Triggers
			CreatureNow = CreatureWithTurn
			ItemNow = CreatureWithTurn.ItemInHand
			NoFail = FireTriggers(CreatureWithTurn, bdPostAttack)
			For	Each ItemX In CreatureWithTurn.Items
				If ItemX.IsReady = True Then
					CreatureNow = CreatureWithTurn
					ItemNow = ItemX
					NoFail = FireTriggers(ItemX, bdPostAttack)
				End If
			Next ItemX
		Loop 
	End Sub
	
	Public Sub CombatDefend()
		CreatureWithTurn.CombatAttitude = bdCombatAttitudeDefend
		CreatureWithTurn.DefenseBonus = Int(CreatureWithTurn.ActionPoints / 5)
		CreatureWithTurn.ActionPoints = 0
		CreatureWithTurn.OpportunityAttack = True
		CombatNextTurn()
	End Sub
	
	'UPGRADE_NOTE: Text was upgraded to Text_Renamed. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="A9E4979A-37FA-4718-9994-97DD76ED70A7"'
	Public Function CombatAttack(ByRef AttackType As Short, ByRef Text_Renamed As String, ByRef InContext As Short, ByRef InVar As Short) As Short
		Dim ItemX As Item
		Dim Found, OverSwing As Short
		Dim Attack, c, Dmg As Short
		Dim MaxRange, Distance, MinRange As Short
		Dim DiceType, DiceCnt, DmgBonus As Short
		Select Case AttackType
			Case 0, 1, 10, 63 ' AttackWith
				If Len(Text_Renamed) < 1 Then
					Text_Renamed = "Claw Attack"
				End If
				' Compute (DistanceToTarget - Range - Movement)
				'UPGRADE_WARNING: Couldn't resolve default property of object CombatRange(). Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
				Distance = CombatRange(CreatureWithTurn, (CreatureTarget.Col), (CreatureTarget.Row))
				If Distance > 1 Then
					CombatMove(MOVEDIRECTION.bdMoveToward, MOVEMOTIVE.bdMoveTarget)
					'UPGRADE_WARNING: Couldn't resolve default property of object CombatRange(). Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
					Distance = CombatRange(CreatureWithTurn, (CreatureTarget.Col), (CreatureTarget.Row))
				End If
				' Compute Overswing penalty or message
				OverSwing = 0
				If (CreatureWithTurn.ActionPoints - 10) < 0 Then
					If GlobalOverSwing = 1 Then
						OverSwing = CreatureWithTurn.ActionPoints - 10
					Else
						Distance = 2
					End If
				End If
				' If in range now, attack
				If Distance < 2 Then
					CombatMessage(Text_Renamed)
					If CombatRollAttack(0, OverSwing) = True Then
						c = CombatRollDamage(InContext, InVar, 0, True, 0)
					Else
						CombatMessage("Miss")
					End If
				End If
			Case 11 ' Say
				DialogSetUp(modGameGeneral.DLGTYPE.bdDlgNoReply)
				DialogShow(CreatureWithTurn, Text_Renamed)
				DialogHide()
			Case 12, 62 ' Weapon
				' Ready weapon if have nothing in hand
				If CreatureWithTurn.ItemInHand.Name = "Hand" Then
					Found = False
					For	Each ItemX In CreatureWithTurn.Items
						If InStr(UCase(ItemX.Name), UCase(Text_Renamed)) > 0 And Len(UCase(Text_Renamed)) > 0 Then
							If ReadyItem(CreatureWithTurn, ItemX) Then
								Found = True
								Exit For
							End If
						End If
					Next ItemX
				Else
					Found = True
				End If
				' If can't find matching weapon, find any damaging thing and ready it
				If Found = False Then
					For	Each ItemX In CreatureWithTurn.Items
						If ItemX.Damage > 0 And ItemX.IsAmmo = False Then
							If UnReadyItem(CreatureWithTurn, (CreatureWithTurn.ItemInHand)) Then
								If ReadyItem(CreatureWithTurn, ItemX) Then
									Found = True
									Exit For
								End If
							End If
						End If
					Next ItemX
				End If
				' If weapon is a shoot, but there's no ammo, you're done for
				If Found = True And CreatureWithTurn.ItemInHand.IsShooter = True Then
					If CreatureWithTurn.ItemAmmo.Index = CreatureWithTurn.ItemInHand.Index And CreatureWithTurn.ItemAmmo.Name = CreatureWithTurn.ItemInHand.Name Then
						Found = False
					End If
				End If
				' If now have a weapon, attack!
				If Found = True Or Int(Rnd() * 10) < 5 Then
					' Compute (DistanceToTarget - Range - Movement)
					MaxRange = CreatureWithTurn.ItemMaxRange
					MinRange = CreatureWithTurn.ItemMinRange
					'UPGRADE_WARNING: Couldn't resolve default property of object CombatRange(). Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
					Distance = CombatRange(CreatureWithTurn, (CreatureTarget.Col), (CreatureTarget.Row))
					If Distance > MaxRange Then
						CombatMove(MOVEDIRECTION.bdMoveToward, MOVEMOTIVE.bdMoveTarget)
					ElseIf Distance < MinRange Then 
						CombatMove(MOVEDIRECTION.bdMoveAway, MOVEMOTIVE.bdMoveTarget)
					End If
					' If in range now, attack
					'UPGRADE_WARNING: Couldn't resolve default property of object CombatRange(). Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
					Distance = CombatRange(CreatureWithTurn, (CreatureTarget.Col), (CreatureTarget.Row))
					' Compute Overswing penalty or message
					OverSwing = 0
					If (CreatureWithTurn.ActionPoints - CreatureWithTurn.ItemInHand.ActionPoints < 0) Or (CreatureWithTurn.ItemInHand.ActionPoints = 0 And CreatureWithTurn.ActionPoints < 10) Then
						If GlobalOverSwing = 1 Then
							If CreatureWithTurn.ItemInHand.ActionPoints = 0 Then
								OverSwing = CreatureWithTurn.ActionPoints - 10
							Else
								OverSwing = CreatureWithTurn.ActionPoints - CreatureWithTurn.ItemInHand.ActionPoints
							End If
						Else
							Distance = MaxRange + 1
						End If
					End If
					' If now in range, attack
					If Distance <= MaxRange And Distance >= MinRange Then
						' Remove ActionPoints and remove ability to Move
						If CreatureWithTurn.ItemInHand.ActionPoints = 0 Then
							CreatureWithTurn.ActionPoints = CreatureWithTurn.ActionPoints - 10
						Else
							CreatureWithTurn.ActionPoints = CreatureWithTurn.ActionPoints - CreatureWithTurn.ItemInHand.ActionPoints
						End If
						CombatDraw()
						ItemNow = CreatureWithTurn.ItemInHand
						' Begin Attack
						CombatMessage(ItemNow.NameText & " Attack")
						PlaySFXItem(CreatureWithTurn)
						If CombatRollAttack(Greatest(CreatureWithTurn.ItemAmmo.DamageType, 0), ItemNow.AttackBonus + OverSwing) = True Then
							c = CombatRollDamage(bdContextDice, CreatureWithTurn.ItemAmmo.Damage - 1, Greatest(CreatureNow.ItemAmmo.DamageType, 0), True, ItemNow.DamageBonus + CreatureNow.ItemAmmo.DamageBonus)
						Else
							CombatMessage("Miss")
						End If
						CreatureWithTurn.UseAmmo()
						CombatDraw()
					End If
				Else
					CombatMove(MOVEDIRECTION.bdMoveFlee, MOVEMOTIVE.bdMoveRandom)
				End If
		End Select
	End Function
	
	Public Function CombatRange(ByRef FromCreature As Creature, ByRef ToCol As Short, ByRef ToRow As Short) As Object
		' Compute corrected range from one creature to another in combat
		' Even: (-1,0), (+1,0), (0,+1), (1,+1), (0,-1), (1,-1)
		' Odd : (-1,0), (+1,0), (-1,+1), (0,+1), (-1,-1), (0,-1)
		' To find correct, keep traversing until within 1. Count how far have moved.
		Dim Row, Found, Col As Short
		Row = FromCreature.Row : Col = FromCreature.Col
		Found = False
		'UPGRADE_WARNING: Couldn't resolve default property of object CombatRange. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
		CombatRange = 0
		Do 
			If System.Math.Abs(Row - ToRow) < 2 And System.Math.Abs(Col - ToCol) < 2 Then
				If Row - ToRow = 0 Then
					Found = True
				Else
					Select Case Row Mod 2
						Case 0 ' Even
							If Col - ToCol > -1 Then
								Found = True
							End If
						Case 1 ' Odd
							If Col - ToCol < 1 Then
								Found = True
							End If
					End Select
				End If
			End If
			'UPGRADE_WARNING: Couldn't resolve default property of object CombatRange. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
			CombatRange = CombatRange + 1
			' Move toward ToCreature
			If Found = False Then
				Row = Greatest(Least(Row + System.Math.Sign(ToRow - Row), bdCombatHeight), 0)
				Col = Greatest(Least(Col + System.Math.Sign(ToCol - Col), bdCombatWidth), 0)
			End If
		Loop Until Found = True
	End Function
	
	Public Function CombatRollAttack(ByRef Style As Short, ByRef Bonus As Short) As Short
		' Returns True if Hit, False if Miss
		Dim c, NoFail As Short
		Dim ItemX As Item
		' Set GlobalDamageStyle for Triggers
		GlobalDamageStyle = Style
		' Draw the Attacker and Target in the menu bar
		CombatDrawAttack(CreatureWithTurn, CreatureTarget, False)
		' Clear global variables
		GlobalAttackRoll = Bonus + CreatureWithTurn.AttackBonus : GlobalArmorRoll = 0 : GlobalDamageRoll = 0
		For c = 1 To 5 : CombatDiceType(c) = 0 : CombatDiceValue(c) = 0 : Next c
		Combat20Bonus = 0 : CombatDmgBonus = 0
		' Fire Pre-Attacked Triggers
		CreatureNow = CreatureWithTurn
		ItemNow = CreatureWithTurn.ItemInHand
		NoFail = FireTriggers(CreatureTarget, bdPreAttacked)
		For	Each ItemNow In CreatureTarget.Items
			If ItemNow.IsReady = True Then
				NoFail = FireTriggers(ItemNow, bdPreAttacked)
			End If
		Next ItemNow
		' Fire Pre-RollAttack Triggers
		CreatureNow = CreatureWithTurn
		ItemNow = CreatureWithTurn.ItemInHand
		NoFail = FireTriggers(CreatureWithTurn, bdPreRollAttack)
		For	Each ItemNow In CreatureWithTurn.Items
			If ItemNow.IsReady = True Then
				NoFail = FireTriggers(ItemNow, bdPreRollAttack)
			End If
		Next ItemNow
		' Roll to Hit
		GlobalAttackRoll = DiceRoll(1, 20, GlobalAttackRoll, True, True)
		' Fire Post-RollAttack Triggers
		CreatureNow = CreatureWithTurn
		ItemNow = CreatureWithTurn.ItemInHand
		NoFail = FireTriggers(CreatureWithTurn, bdPostRollAttack)
		For	Each ItemNow In CreatureWithTurn.Items
			If ItemNow.IsReady = True Then
				NoFail = FireTriggers(ItemNow, bdPostRollAttack)
			End If
		Next ItemNow
		' Fire Post-Attacked Triggers
		CreatureNow = CreatureWithTurn
		ItemNow = CreatureWithTurn.ItemInHand
		NoFail = FireTriggers(CreatureTarget, bdPostAttacked)
		For	Each ItemNow In CreatureTarget.Items
			If ItemNow.IsReady = True Then
				CreatureNow = CreatureWithTurn
				ItemNow = CreatureTarget.ItemInHand
				NoFail = FireTriggers(ItemNow, bdPostAttacked)
			End If
		Next ItemNow
		' If roll a natural 1, chance to drop Item in hand
		If GlobalAttackRoll = 1 And CreatureWithTurn.ItemInHand.Name <> "Hand" Then
			If Int(Rnd() * 20) + 1 > CreatureWithTurn.Agility Then
				DialogDM(CreatureWithTurn.Name & " fumbles and drops the " & CreatureWithTurn.ItemInHand.Name & ".")
				NoFail = DropItem(CreatureWithTurn, bdInvEncounter, 0, (CreatureWithTurn.ItemInHand))
			End If
		End If
		' Play hit or miss sound
		If GlobalAttackRoll < CreatureTarget.Defense Then
            Call PlaySoundFile("miss" & Int(Rnd() * 4) + 1, tome, , 5)
			CombatRollAttack = False
		Else
			CombatRollAttack = True
		End If
		ItemNow = CreatureWithTurn.ItemInHand
	End Function
	
	Public Function CombatRollDamage(ByRef InContext As Short, ByRef InVar As Short, ByRef Style As Short, ByRef CheckArmor As Short, ByRef Bonus As Short) As Short
		Dim DiceType, c, DiceCnt, NoFail As Short
		' Style is 0 = None, 1 = Sharp, 2 = Blunt, etc. This applies to Resistance in Creatures.
		' If already dead, this does nothing
		If CreatureTarget.HPNow < 1 Then
			CombatRollDamage = 0
			Exit Function
		End If
		' Add Bonus for Creature, Weapon and Strength
		GlobalDamageRoll = GlobalDamageRoll + CreatureWithTurn.DamageBonus + Bonus
		' Add Strength ToDamage Bonus (Short Range Only)
		If CreatureWithTurn.ItemMinRange < 2 Then
			GlobalDamageRoll = GlobalDamageRoll + Fix((CreatureWithTurn.Strength - 10) / 3)
		End If
		' Set GlobalDamageStyle for Triggers to access
		GlobalDamageStyle = Style
		' Draw the Attacker and Target in the menu bar
		If picGrid.Visible = True Then
			CombatDrawAttack(CreatureWithTurn, CreatureTarget, True)
		End If
		' If allow checking Armor Then do so
		If CheckArmor = True Then
			' Fire Pre-RollArmorChit Triggers
			CreatureNow = CreatureWithTurn
			ItemNow = CreatureWithTurn.ItemInHand
			NoFail = FireTriggers(CreatureNow, bdPreRollArmorChit)
			For	Each ItemNow In CreatureNow.Items
				If ItemNow.IsReady = True Then
					NoFail = FireTriggers(ItemNow, bdPreRollArmorChit)
				End If
			Next ItemNow
			' Roll Armor
			GlobalArmorRoll = CombatRollArmor(CreatureTarget, True, GlobalArmorRoll)
			' Set HitLocation
			GlobalHitLocation = modBD.SetUpArmorType(CreatureTarget.BodyType(GlobalArmorRoll))
			' Fire Post-RollArmorChit Triggers
			CreatureNow = CreatureWithTurn
			ItemNow = CreatureWithTurn.ItemInHand
			NoFail = FireTriggers(CreatureWithTurn, bdPostRollArmorChit)
			For	Each ItemNow In CreatureWithTurn.Items
				If ItemNow.IsReady = True Then
					NoFail = FireTriggers(ItemNow, bdPostRollArmorChit)
				End If
			Next ItemNow
		Else
			' Roll Armor
			GlobalArmorRoll = CombatRollArmor(CreatureTarget, True, GlobalArmorRoll)
			' Set HitLocation
			GlobalHitLocation = modBD.SetUpArmorType(CreatureTarget.BodyType(GlobalArmorRoll))
		End If
		' Fire Pre-RollDamage Triggers
		CreatureNow = CreatureWithTurn
		ItemNow = CreatureWithTurn.ItemInHand
		NoFail = FireTriggers(CreatureWithTurn, bdPreRollDamage)
		For	Each ItemNow In CreatureWithTurn.Items
			If ItemNow.IsReady = True Then
				NoFail = FireTriggers(ItemNow, bdPreRollDamage)
			End If
		Next ItemNow
		' Roll for Damage (or Set Damage if not dice)
		If InContext = bdContextDice Then
			DiceCnt = InVar Mod 5 + 1
			DiceType = Int((InVar Mod 25) / 5) * 2 + 4
			If InVar > 24 Then
				GlobalDamageRoll = GlobalDamageRoll + Int(InVar / 25)
			End If
			GlobalDamageRoll = DiceRoll(DiceCnt, DiceType, GlobalDamageRoll, False, True)
		Else
			If IsNumeric(modEvents.GetVarContext(InContext, InVar)) Then
				'UPGRADE_WARNING: Couldn't resolve default property of object modEvents.GetVarContext(). Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
				GlobalDamageRoll = modEvents.GetVarContext(InContext, InVar)
			Else
				GlobalDamageRoll = 0
			End If
		End If
		' Fire Post-RollDamage Triggers
		CreatureNow = CreatureWithTurn
		ItemNow = CreatureWithTurn.ItemInHand
		NoFail = FireTriggers(CreatureWithTurn, bdPostRollDamage)
		For	Each ItemNow In CreatureWithTurn.Items
			If ItemNow.IsReady = True Then
				NoFail = FireTriggers(ItemNow, bdPostRollDamage)
			End If
		Next ItemNow
		' Modify Damage based on Great Roll
		If GlobalAttackRoll >= CreatureTarget.Defense * 2 Then
			GlobalDamageRoll = GlobalDamageRoll * 2
			CombatMessage("* Great Hit - Double Damage! *")
		End If
		' Modify Damage based on Vulnerable
		If IsBetween(Style, 1, 8) Then
			c = CreatureTarget.ResistanceTypeWithArmor(Style)
			GlobalDamageRoll = GlobalDamageRoll - Int(GlobalDamageRoll * (c / 100))
			If c < 0 Then
				If IsBetween(c, -100, -1) Then
					CombatMessage("* Weakness - Double Damage! *")
				ElseIf IsBetween(c, -200, -101) Then 
					CombatMessage("* Weakness - Triple Damage! *")
				Else
					CombatMessage("* Weakness - Crippling Damage! *")
				End If
			ElseIf c > 0 Then 
				CombatMessage("* " & modBD.SetUpResistanceType(Style) & " Resistance *")
			ElseIf c >= 100 Then 
				GlobalDamageRoll = 0
				CombatMessage("* " & modBD.SetUpResistanceType(Style) & " Immune - No Damage! *")
			End If
		End If
		' Normal Resistance
		If CheckArmor = True And GlobalDamageRoll > 0 Then
			c = CreatureTarget.ResistanceWithArmor(GlobalArmorRoll)
			If c > 0 Then
				CombatMessage("Armor Absorbs " & Int(GlobalDamageRoll * (c / 100)))
				GlobalDamageRoll = GlobalDamageRoll - Int(GlobalDamageRoll * (c / 100))
			End If
		End If
		' Never go less than 0
		If GlobalDamageRoll < 1 Then
			GlobalDamageRoll = 0
		End If
		' If display mode, show location hit
		If CheckArmor = True And picGrid.Visible = True Then
			CombatMessage(GlobalHitLocation & " hit for " & GlobalDamageRoll & " damage")
		Else
			CombatMessage("Hit for " & GlobalDamageRoll & " damage")
		End If
		' Damage or Not?
		If GlobalDamageRoll > 0 Then
			If CreatureTarget.HitWAV = "" Then
                Call PlaySoundFile("hit", tome, , 5)
			Else
                Call PlaySoundFile(CreatureTarget.HitWAV, tome, , 5)
				If CreatureTarget.HitWAVOneTime = True Then
					CreatureTarget.HitWAV = ""
					CreatureTarget.HitWAVOneTime = False
				End If
			End If
			' Apply Damage
			CombatApplyDamage(CreatureTarget, GlobalDamageRoll)
		Else
            Call PlaySoundFile("armor", tome, , 5)
		End If
		CombatRollDamage = GlobalDamageRoll
	End Function
	
	Public Sub CombatDraw(Optional ByRef ShowMsg As Object = Nothing)
		Dim n, c, i, Found As Short
		'UPGRADE_NOTE: Width was upgraded to Width_Renamed. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="A9E4979A-37FA-4718-9994-97DD76ED70A7"'
		'UPGRADE_NOTE: Height was upgraded to Height_Renamed. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="A9E4979A-37FA-4718-9994-97DD76ED70A7"'
		Dim rc As Integer
		Dim OffX, Height_Renamed, Width_Renamed, OffY As Short
		Dim HasRunes, X, Y, RuneX As Short
		Dim GridWidth As Double
		Dim hOldBrush, hOldPen, hNewPen, hNewBrush As Integer
		Dim CreatureX As Creature
		Dim Distance, MaxRange As Short
		' Sort all pictures in order
		CombatPaintSort()
		' Draw background wallpaper
		'UPGRADE_ISSUE: PictureBox property picWallPaper.hdc was not upgraded. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="CC4C7EC0-C903-48FC-ACCC-81861D12DA4A"'
		'UPGRADE_ISSUE: PictureBox property picGrid.hdc was not upgraded. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="CC4C7EC0-C903-48FC-ACCC-81861D12DA4A"'
		rc = BitBlt(picGrid.hdc, 0, 0, picGrid.Width, picGrid.Height, picWallPaper.hdc, 0, 0, SRCCOPY)
		' Draw Action Point globes at Top (500 Pixels is the total width of 20 ActionPoints)
		X = (picGrid.Width - 500) / 2
		For c = 0 To 19
			'UPGRADE_ISSUE: PictureBox property picMisc.hdc was not upgraded. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="CC4C7EC0-C903-48FC-ACCC-81861D12DA4A"'
			'UPGRADE_ISSUE: PictureBox property picGrid.hdc was not upgraded. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="CC4C7EC0-C903-48FC-ACCC-81861D12DA4A"'
			rc = BitBlt(picGrid.hdc, X + c * 25, 8, 25, 25, picMisc.hdc, 50, 121, SRCAND)
			If c < CreatureWithTurn.ActionPoints Then
				'UPGRADE_ISSUE: PictureBox property picMisc.hdc was not upgraded. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="CC4C7EC0-C903-48FC-ACCC-81861D12DA4A"'
				'UPGRADE_ISSUE: PictureBox property picGrid.hdc was not upgraded. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="CC4C7EC0-C903-48FC-ACCC-81861D12DA4A"'
				rc = BitBlt(picGrid.hdc, X + c * 25, 8, 25, 25, picMisc.hdc, 0, 121, SRCPAINT)
			Else
				'UPGRADE_ISSUE: PictureBox property picMisc.hdc was not upgraded. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="CC4C7EC0-C903-48FC-ACCC-81861D12DA4A"'
				'UPGRADE_ISSUE: PictureBox property picGrid.hdc was not upgraded. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="CC4C7EC0-C903-48FC-ACCC-81861D12DA4A"'
				rc = BitBlt(picGrid.hdc, X + c * 25, 8, 25, 25, picMisc.hdc, 25, 121, SRCPAINT)
			End If
		Next c
		' Draw path to Target
		If Target > -1 Then
			'UPGRADE_WARNING: Couldn't resolve default property of object CombatRange(). Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
			Distance = CombatRange(CreatureWithTurn, (Turns(Target).Ref.Col), (Turns(Target).Ref.Row))
			MaxRange = Greatest((CreatureWithTurn.ItemMaxRange), TargetCreatureRange)
		Else
			Distance = 1 : MaxRange = 0
		End If
		If CreatureWithTurn.DMControlled = False And MenuNow <> bdMenuTargetCreature And MenuNow <> bdMenuTargetAny And MenuNow <> bdMenuTargetParty And Distance > MaxRange Then
			For c = CombatStepTop To Greatest(CombatStepTop - CreatureWithTurn.MovementRange, 0) Step -1
				GridToCursor(X, Y, CombatStepRow(c), CombatStepCol(c), GridWidth)
				OffX = ((bdCombatGridWidth * CommonScale) - bdCombatGridWidth) / 2
				OffY = ((bdCombatGridHeight * CommonScale) - bdCombatGridHeight) / 2
				'UPGRADE_ISSUE: PictureBox property picSteps.hdc was not upgraded. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="CC4C7EC0-C903-48FC-ACCC-81861D12DA4A"'
				'UPGRADE_ISSUE: PictureBox property picGrid.hdc was not upgraded. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="CC4C7EC0-C903-48FC-ACCC-81861D12DA4A"'
				rc = BitBlt(picGrid.hdc, X + OffX, Y - bdCombatGridHeight + OffY, bdCombatGridWidth, bdCombatGridHeight, picSteps.hdc, 0, bdCombatGridHeight, SRCAND)
				'UPGRADE_ISSUE: PictureBox property picSteps.hdc was not upgraded. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="CC4C7EC0-C903-48FC-ACCC-81861D12DA4A"'
				'UPGRADE_ISSUE: PictureBox property picGrid.hdc was not upgraded. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="CC4C7EC0-C903-48FC-ACCC-81861D12DA4A"'
				rc = BitBlt(picGrid.hdc, X + OffX, Y - bdCombatGridHeight + OffY, bdCombatGridWidth, bdCombatGridHeight, picSteps.hdc, 0, 0, SRCPAINT)
			Next c
		End If
		' Determine color of rectangle based on distance from CreatureWithTurn
		For c = 0 To CombatTurn
			' Convert where CreatureWithTurn row and col to x and y coordinate
			GridToCursor(X, Y, Turns(Paint_Renamed(c)).Ref.Row, Turns(Paint_Renamed(c)).Ref.Col, GridWidth)
			If Turns(Paint_Renamed(c)).Ref.HPNow > 0 Then
				n = Turns(Paint_Renamed(c)).Ref.Pic
			Else
				n = 0 ' Default for bones
				Turns(Paint_Renamed(c)).Ref.Facing = 1
			End If
			' Determine size of picture to draw
			Height_Renamed = Int(picCPic(n).Height / 4)
			Width_Renamed = Int(picCPic(n).Width / 2)
			' Center the picture in the square
			Y = Int(Y - Height_Renamed)
			X = X - (Width_Renamed - GridWidth) / 2
			' Paint Yellow around TurnNow, Red outline around Target, others are Black
			If Paint_Renamed(c) = TurnNow Then
				' Yellow
				'UPGRADE_ISSUE: PictureBox property picCPic.hdc was not upgraded. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="CC4C7EC0-C903-48FC-ACCC-81861D12DA4A"'
				'UPGRADE_ISSUE: PictureBox property picGrid.hdc was not upgraded. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="CC4C7EC0-C903-48FC-ACCC-81861D12DA4A"'
				rc = BitBlt(picGrid.hdc, X, Y, Width_Renamed, Height_Renamed, picCPic(n).hdc, Turns(Paint_Renamed(c)).Ref.Facing * Width_Renamed, Height_Renamed, SRCAND)
				'UPGRADE_ISSUE: PictureBox property picCPic.hdc was not upgraded. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="CC4C7EC0-C903-48FC-ACCC-81861D12DA4A"'
				'UPGRADE_ISSUE: PictureBox property picGrid.hdc was not upgraded. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="CC4C7EC0-C903-48FC-ACCC-81861D12DA4A"'
				rc = BitBlt(picGrid.hdc, X, Y, Width_Renamed, Height_Renamed, picCPic(n).hdc, Turns(Paint_Renamed(c)).Ref.Facing * Width_Renamed, Height_Renamed * 2, SRCPAINT)
			ElseIf Paint_Renamed(c) = Target Then 
				'UPGRADE_WARNING: Couldn't resolve default property of object CombatRange(). Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
				If IsBetween(CombatRange(CreatureWithTurn, (CreatureTarget.Col), (CreatureTarget.Row)), CreatureWithTurn.ItemMinRange, CreatureWithTurn.ItemMaxRange) Then
					' Red
					'UPGRADE_ISSUE: PictureBox property picCPic.hdc was not upgraded. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="CC4C7EC0-C903-48FC-ACCC-81861D12DA4A"'
					'UPGRADE_ISSUE: PictureBox property picGrid.hdc was not upgraded. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="CC4C7EC0-C903-48FC-ACCC-81861D12DA4A"'
					rc = BitBlt(picGrid.hdc, X, Y, Width_Renamed, Height_Renamed, picCPic(n).hdc, Turns(Paint_Renamed(c)).Ref.Facing * Width_Renamed, Height_Renamed, SRCAND)
					'UPGRADE_ISSUE: PictureBox property picCPic.hdc was not upgraded. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="CC4C7EC0-C903-48FC-ACCC-81861D12DA4A"'
					'UPGRADE_ISSUE: PictureBox property picGrid.hdc was not upgraded. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="CC4C7EC0-C903-48FC-ACCC-81861D12DA4A"'
					rc = BitBlt(picGrid.hdc, X, Y, Width_Renamed, Height_Renamed, picCPic(n).hdc, Turns(Paint_Renamed(c)).Ref.Facing * Width_Renamed, Height_Renamed * 3, SRCPAINT)
				Else
					' Normal
					'UPGRADE_ISSUE: PictureBox property picCPic.hdc was not upgraded. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="CC4C7EC0-C903-48FC-ACCC-81861D12DA4A"'
					'UPGRADE_ISSUE: PictureBox property picGrid.hdc was not upgraded. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="CC4C7EC0-C903-48FC-ACCC-81861D12DA4A"'
					rc = BitBlt(picGrid.hdc, X, Y, Width_Renamed, Height_Renamed, picCPic(n).hdc, Turns(Paint_Renamed(c)).Ref.Facing * Width_Renamed, Height_Renamed, SRCAND)
					'UPGRADE_ISSUE: PictureBox property picCPic.hdc was not upgraded. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="CC4C7EC0-C903-48FC-ACCC-81861D12DA4A"'
					'UPGRADE_ISSUE: PictureBox property picGrid.hdc was not upgraded. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="CC4C7EC0-C903-48FC-ACCC-81861D12DA4A"'
					rc = BitBlt(picGrid.hdc, X, Y, Width_Renamed, Height_Renamed, picCPic(n).hdc, Turns(Paint_Renamed(c)).Ref.Facing * Width_Renamed, 0, SRCPAINT)
				End If
			Else
				' Normal
				'UPGRADE_ISSUE: PictureBox property picCPic.hdc was not upgraded. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="CC4C7EC0-C903-48FC-ACCC-81861D12DA4A"'
				'UPGRADE_ISSUE: PictureBox property picGrid.hdc was not upgraded. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="CC4C7EC0-C903-48FC-ACCC-81861D12DA4A"'
				rc = BitBlt(picGrid.hdc, X, Y, Width_Renamed, Height_Renamed, picCPic(n).hdc, Turns(Paint_Renamed(c)).Ref.Facing * Width_Renamed, Height_Renamed, SRCAND)
				'UPGRADE_ISSUE: PictureBox property picCPic.hdc was not upgraded. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="CC4C7EC0-C903-48FC-ACCC-81861D12DA4A"'
				'UPGRADE_ISSUE: PictureBox property picGrid.hdc was not upgraded. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="CC4C7EC0-C903-48FC-ACCC-81861D12DA4A"'
				rc = BitBlt(picGrid.hdc, X, Y, Width_Renamed, Height_Renamed, picCPic(n).hdc, Turns(Paint_Renamed(c)).Ref.Facing * Width_Renamed, 0, SRCPAINT)
			End If
			' Draw MsgQue
			'UPGRADE_NOTE: IsMissing() was changed to IsNothing(). Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="8AE1CB93-37AB-439A-A4FF-BE3B6760BB23"'
			If (Paint_Renamed(c) = Target And IsNothing(ShowMsg) = False) Or (Paint_Renamed(c) = TurnNow And (Target = -1 Or IsNothing(ShowMsg) = True)) Then
				GridToCursor(X, Y, Turns(Paint_Renamed(c)).Ref.Row, Turns(Paint_Renamed(c)).Ref.Col, GridWidth)
				Y = Greatest(Y - Int(picCPic(Turns(Paint_Renamed(c)).Ref.Pic).Height / 3) - Turns(Paint_Renamed(c)).Ref.MsgQueTop * 10, 32)
				Found = 0
				If X + 256 > picGrid.ClientRectangle.Width Then
					X = picGrid.ClientRectangle.Width - 264
					Found = 1
				End If
				For i = 0 To Turns(Paint_Renamed(c)).Ref.MsgQueTop
					Y = Y + ShowText(picGrid, X, Y, 256, 90, bdFontSmallWhite, Turns(Paint_Renamed(c)).Ref.MsgQue(i), Found, True)
				Next i
			End If
		Next c
		' Set focus to the combat screen
		picGrid.Focus()
		picGrid.Refresh()
	End Sub
	
	Private Sub GridToCursor(ByRef CursorAtX As Short, ByRef CursorAtY As Short, ByVal Row As Short, ByVal Col As Short, ByRef GridWidth As Double)
		Dim c, GridAdjust As Short
		' Set GridWidth and GridAdjust to X
		GridWidth = bdCombatGridWidth
		GridAdjust = bdCombatLeft + bdCombatGridHeight * (Row Mod 2)
		' Do final adjustments to CursorAtX and CursorAtY (bound to grid)
		CursorAtY = (Int((bdCombatHeight - Row) * bdCombatGridHeight) + bdCombatTop) * CommonScale
		CursorAtX = (Int(Col * GridWidth + GridAdjust)) * CommonScale
	End Sub
	
	'UPGRADE_NOTE: Map was upgraded to Map_Renamed. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="A9E4979A-37FA-4718-9994-97DD76ED70A7"'
	'UPGRADE_NOTE: Tome was upgraded to Tome_Renamed. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="A9E4979A-37FA-4718-9994-97DD76ED70A7"'
	Private Sub CombatWallpaper()
        Dim tome_Renamed As Object
		Dim Map_Renamed As Object
		Dim FileName As String
		Dim c, Found As Short
		Dim PictureFile As String
		'UPGRADE_WARNING: Arrays in structure bmBlack may need to be initialized before they can be used. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="814DF224-76BD-4BB4-BFFB-EA359CB9FC48"'
		Dim bmBlack As BITMAPINFO
		Dim hMem, TransparentRGB As Integer
		Dim rc As Short
		Dim lpMem As Integer
		' Default to a forest wall paper
		PictureFile = "forest1.bmp"
		' Set PictureFile based on EncounterNow
		If Len(EncounterNow.Wallpaper) > 0 Then
			PictureFile = EncounterNow.Wallpaper
		Else
			' Set PictureFile based on Tiles: Bottom, Middle and Top
			Found = False
			'UPGRADE_WARNING: Couldn't resolve default property of object Tome.MapY. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
			'UPGRADE_WARNING: Couldn't resolve default property of object Tome.MapX. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
			'UPGRADE_WARNING: Couldn't resolve default property of object Map.MiddleTile. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
            c = Map.MiddleTile(tome.MapX, tome.MapY)
			If c > 0 Then
				'UPGRADE_WARNING: Couldn't resolve default property of object Map.Tiles. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
				If Map.Tiles("L" & c).TileSet > 0 Then
					'UPGRADE_WARNING: Couldn't resolve default property of object Map.Tiles. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
					'UPGRADE_WARNING: Couldn't resolve default property of object Map.TileSets. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
					PictureFile = Map.TileSets(Map.Tiles("L" & c).TileSet).PictureFile
					Found = Len(PictureFile) > 0
				End If
			End If
			'UPGRADE_WARNING: Couldn't resolve default property of object Tome.MapY. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
			'UPGRADE_WARNING: Couldn't resolve default property of object Tome.MapX. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
			'UPGRADE_WARNING: Couldn't resolve default property of object Map.TopTile. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
            c = Map.TopTile(tome.MapX, tome.MapY)
			If c > 0 And Not Found Then
				'UPGRADE_WARNING: Couldn't resolve default property of object Map.Tiles. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
				If Map.Tiles("L" & c).TileSet > 0 Then
					'UPGRADE_WARNING: Couldn't resolve default property of object Map.Tiles. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
					'UPGRADE_WARNING: Couldn't resolve default property of object Map.TileSets. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
					PictureFile = Map.TileSets(Map.Tiles("L" & c).TileSet).PictureFile
					Found = Len(PictureFile) > 0
				End If
			End If
			'UPGRADE_WARNING: Couldn't resolve default property of object Tome.MapY. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
			'UPGRADE_WARNING: Couldn't resolve default property of object Tome.MapX. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
			'UPGRADE_WARNING: Couldn't resolve default property of object Map.BottomTile. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
            c = Map.BottomTile(tome.MapX, tome.MapY)
			If c > 0 And Not Found Then
				'UPGRADE_WARNING: Couldn't resolve default property of object Map.Tiles. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
				If Map.Tiles("L" & c).TileSet > 0 Then
					'UPGRADE_WARNING: Couldn't resolve default property of object Map.Tiles. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
					'UPGRADE_WARNING: Couldn't resolve default property of object Map.TileSets. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
					PictureFile = Map.TileSets(Map.Tiles("L" & c).TileSet).PictureFile
					Found = Len(PictureFile) > 0
				End If
			End If
			If Not Found Then
				'UPGRADE_WARNING: Couldn't resolve default property of object Map.TileSets. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
				If Map.TileSets.Count > 0 Then
					'UPGRADE_WARNING: Couldn't resolve default property of object Map.TileSets. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
					PictureFile = Map.TileSets(Int(Map.TileSets.Count * Rnd()) + 1).PictureFile
				End If
			End If
		End If
		' Load Bitmap
		'UPGRADE_WARNING: Couldn't resolve default property of object Tome.FullPath. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
		'UPGRADE_WARNING: Dir has a new behavior. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="9B7D5ADD-D8FE-4819-A36C-6DEDAF088CC7"'
        FileName = Dir(tome.FullPath & "\" & PictureFile)
		If FileName = "" Then
			'UPGRADE_WARNING: Couldn't resolve default property of object Tome.FullPath. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
			'UPGRADE_WARNING: Dir has a new behavior. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="9B7D5ADD-D8FE-4819-A36C-6DEDAF088CC7"'
            FileName = Dir(tome.FullPath & "\wallpapers\" & PictureFile)
			If FileName = "" Then
				'          FileName = gAppPath & "\data\graphics\wallpapers\" & PictureFile
				FileName = gDataPath & "\graphics\wallpapers\" & PictureFile
			Else
				'UPGRADE_WARNING: Couldn't resolve default property of object Tome.FullPath. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
                FileName = tome.FullPath & "\wallpapers\" & PictureFile
			End If
		Else
            'UPGRADE_WARNING: Couldn't resolve default property of object tome.FullPath. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
            FileName = tome.FullPath & "\" & PictureFile
		End If
		ReadBitmapFile(FileName, bmBlack, hMem, TransparentRGB)
		' Paint bitmap to picture box using converted palette
		lpMem = GlobalLock(hMem)
		picWallPaper.Width = picGrid.Width
		picWallPaper.Height = picGrid.Height
		' Draw Wallpaper
		'UPGRADE_ISSUE: PictureBox property picWallPaper.hdc was not upgraded. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="CC4C7EC0-C903-48FC-ACCC-81861D12DA4A"'
		rc = SetStretchBltMode(picWallPaper.hdc, 3)
		'UPGRADE_ISSUE: PictureBox property picWallPaper.hdc was not upgraded. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="CC4C7EC0-C903-48FC-ACCC-81861D12DA4A"'
		rc = StretchDIBits(picWallPaper.hdc, 0, 0, picWallPaper.Width, picWallPaper.Height, 0, 0, bmBlack.bmiHeader.biWidth, bmBlack.bmiHeader.biHeight, lpMem, bmBlack, DIB_RGB_COLORS, SRCCOPY)
		' Release memory
		rc = GlobalUnlock(hMem)
		rc = GlobalFree(hMem)
	End Sub
	
	'Public Sub CombatMove(InDirection As Integer, Motive As Integer)
	Public Sub CombatMove(ByRef InDirection As MOVEDIRECTION, ByRef Motive As MOVEMOTIVE)
		Dim Found, c, NoFail As Short
		Dim ItemX As Item
		Dim rc, Steps As Short
		Dim j, i, PlayedSound As Short
		Dim RowTo, ColTo As Short
		Dim MoveRow, MoveCol As Short
		Dim InDirectionCol, InvalidMove, InDirectionRow As Short
		Dim MaxTries As Short
		Dim MoveCnt, Near As Short
		Dim CreatureX As Creature
		Dim CreatureWithTurnHold As Creature
		Dim CreatureTargetHold As Creature
		Dim ItemNowHold As Item
		Dim CreatureNowHold As Creature
		' Locate Target to move toward
		i = -1
		Select Case Motive
			Case MOVEMOTIVE.bdMoveTarget, MOVEMOTIVE.bdMoveRandom
				' If Target set, use that Target
				If Motive = MOVEMOTIVE.bdMoveTarget And Target > -1 Then
					i = Target
				Else ' Move Random
					For c = 0 To CombatTurn
						If Turns(c).Ref.Friendly <> CreatureWithTurn.Friendly And (i = -1 Or Int(Rnd() * 100) < 50) Then
							i = c
						End If
					Next c
				End If
			Case MOVEDIRECTION.bdMoveClick
				' Don't need to choose a Target
			Case MOVEDIRECTION.bdMoveFlee
				' Nothing to do here
			Case Else
				For c = 0 To CombatTurn
					If Turns(c).Ref.Friendly <> CreatureWithTurn.Friendly Then
						i = 0
						Select Case Motive
							Case MOVEMOTIVE.bdMoveClosest
								If (Turns(c).Ref.Col - CreatureWithTurn.Col) ^ 2 + (Turns(c).Ref.Row - CreatureWithTurn.Row) ^ 2 < (Turns(i).Ref.Col - CreatureWithTurn.Col) ^ 2 + (Turns(i).Ref.Row - CreatureWithTurn.Row) ^ 2 Then
									i = c
								End If
							Case MOVEMOTIVE.bdMoveFarthest
								If (Turns(c).Ref.Col - CreatureWithTurn.Col) ^ 2 + (Turns(c).Ref.Row - CreatureWithTurn.Row) ^ 2 > (Turns(i).Ref.Col - CreatureWithTurn.Col) ^ 2 + (Turns(i).Ref.Row - CreatureWithTurn.Row) ^ 2 Then
									i = c
								End If
							Case MOVEMOTIVE.bdMoveStrongest
								If Turns(c).Ref.HPNow + Turns(c).Ref.ActionPoints > Turns(i).Ref.HPNow + Turns(i).Ref.ActionPoints Then
									i = c
								End If
							Case MOVEMOTIVE.bdMoveWeakest
								If Turns(c).Ref.HPNow + Turns(c).Ref.ActionPoints < Turns(i).Ref.HPNow + Turns(i).Ref.ActionPoints Then
									i = c
								End If
						End Select
					End If
				Next c
		End Select
		' When we get here, a Target has been chosen
		If i > -1 Then
			CreatureTarget = Turns(i).Ref
		End If
		' Fire Pre-CombatMove Triggers
		NoFail = True
		CreatureNow = CreatureWithTurn
		ItemNow = CreatureWithTurn.ItemInHand
		rc = FireTriggers(EncounterNow, bdPreCombatMove)
		If rc = False Then
			NoFail = False
		End If
		For	Each ItemNow In CreatureWithTurn.Items
			If ItemNow.IsReady = True Then
				CreatureNow = CreatureWithTurn
				rc = FireTriggers(ItemNow, bdPreCombatMove)
				If rc = False Then
					NoFail = False
				End If
			End If
		Next ItemNow
		' If NoFail on all Triggers, then you can Move
		If NoFail = True Then
			' Move the Creature
			If InDirection = MOVEDIRECTION.bdMoveClick Then
				' Move to cursor click
				CursorToGrid(ClickX, ClickY, RowTo, ColTo)
				Near = 1
			ElseIf InDirection = MOVEDIRECTION.bdMoveToward Then 
				ColTo = CreatureTarget.Col : RowTo = CreatureTarget.Row
				Near = 2
			ElseIf InDirection = MOVEDIRECTION.bdMoveFlee Then 
				If CreatureWithTurn.Col < 5 Then
					ColTo = 0
				Else
					ColTo = bdCombatWidth
				End If
				If CreatureWithTurn.Row < 3 Then
					RowTo = 0
				Else
					RowTo = bdCombatHeight
				End If
			Else ' bdMoveAway
				ColTo = 5 - CreatureTarget.Col : RowTo = 3 - CreatureTarget.Row
				Near = 1
			End If
			' Play Move Sound
			If CreatureWithTurn.MoveWAV <> "" Then
                Call PlaySoundFile(CreatureWithTurn.MoveWAV, tome)
				If CreatureWithTurn.MoveWAVOneTime = True Then
					CreatureWithTurn.MoveWAV = ""
					CreatureWithTurn.MoveWAVOneTime = False
				End If
			End If
			' Find path to Target
			CombatFindPath(CreatureWithTurn, ColTo, RowTo)
			' Play step sound and then move toward target
            Call PlaySoundFile("step" & Int(Rnd() * 4) + 1, tome)
			i = 0 : Steps = 0
			For c = CombatStepTop To 0 Step -1
				' Fire Pre-CombatStep Triggers
				NoFail = True
				CreatureNow = CreatureWithTurn
				ItemNow = CreatureWithTurn.ItemInHand
				rc = FireTriggers(CreatureNow, bdPreCombatStep)
				CreatureNow = CreatureWithTurn
				ItemNow = CreatureWithTurn.ItemInHand
				rc = FireTriggers(EncounterNow, bdPreCombatStep)
				If rc = False Then
					NoFail = False
				End If
				For	Each ItemNow In CreatureWithTurn.Items
					If ItemNow.IsReady = True Then
						CreatureNow = CreatureWithTurn
						rc = FireTriggers(ItemNow, bdPreCombatStep)
						If rc = False Then
							NoFail = False
						End If
					End If
				Next ItemNow
				' If no Trigger failure, move the Creature
				If NoFail = True And CreatureWithTurn.ActionPoints > 0 Then
					CreatureWithTurn.Row = CombatStepRow(c) : CreatureWithTurn.Col = CombatStepCol(c)
					If CombatStepCol(CombatStepTop) < CreatureWithTurn.Col Then
						CreatureWithTurn.Facing = 1
					Else
						CreatureWithTurn.Facing = 0
					End If
					i = i + 1
					' Refresh the screen and pause a brief moment
					CombatDraw()
					GameDelay(10)
					' Gains Opportunity Attack (if attack attitude is set)
					For j = 0 To CombatTurn
						If Turns(j).Ref.AllowedTurn = True And Turns(j).Ref.Friendly <> CreatureWithTurn.Friendly Then
							'UPGRADE_WARNING: Couldn't resolve default property of object CombatRange(CreatureWithTurn, Turns(j).Ref.Col, Turns(j).Ref.Row). Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
							If CombatRange(CreatureWithTurn, (Turns(j).Ref.Col), (Turns(j).Ref.Row)) < 2 And Turns(j).Ref.OpportunityAttack = True Then
								' Opportunity Attack
								CombatMessage(CreatureTarget.Name & " Defends")
								CreatureTargetHold = CreatureTarget
								CreatureTarget = CreatureWithTurn
								CreatureWithTurnHold = CreatureWithTurn
								CreatureWithTurn = Turns(j).Ref
								Turns(j).Ref.OpportunityAttack = False
								CombatAttack(62, CreatureWithTurn.ItemInHand.Name, 0, 0)
								If CreatureTarget.HPNow < 1 Then
									Exit Sub
								End If
								CreatureTarget = CreatureTargetHold
								CreatureWithTurn = CreatureWithTurnHold
							End If
						End If
					Next j
					' Fire Post-CombatStep Triggers
					NoFail = True
					CreatureNow = CreatureWithTurn
					ItemNow = CreatureWithTurn.ItemInHand
					rc = FireTriggers(CreatureNow, bdPostCombatStep)
					CreatureNow = CreatureWithTurn
					ItemNow = CreatureWithTurn.ItemInHand
					rc = FireTriggers(EncounterNow, bdPostCombatStep)
					If rc = False Then
						NoFail = False
					End If
					For	Each ItemNow In CreatureWithTurn.Items
						If ItemNow.IsReady = True Then
							CreatureNow = CreatureWithTurn
							rc = FireTriggers(ItemNow, bdPostCombatStep)
							If rc = False Then
								NoFail = False
							End If
						End If
					Next ItemNow
					' Take away ActionPoints for moving and end move if out of points
					Steps = Steps + 1
					If CreatureWithTurn.ActionPoints - CreatureWithTurn.MovementCost * Steps < 1 Then
						Exit For
					End If
				End If
			Next c
			' Spend Action Points
			CreatureWithTurn.ActionPoints = CreatureWithTurn.ActionPoints - CreatureWithTurn.MovementCost * Steps
			' Fire Post-CombatMove
			CreatureNow = CreatureWithTurn
			ItemNow = CreatureWithTurn.ItemInHand
			NoFail = FireTriggers(EncounterNow, bdPostCombatMove)
			CreatureNow = CreatureWithTurn
			For	Each ItemNow In CreatureWithTurn.Items
				If ItemNow.IsReady = True Then
					CreatureNow = CreatureWithTurn
					NoFail = FireTriggers(ItemNow, bdPostCombatMove)
				End If
			Next ItemNow
		End If
	End Sub
	
	Public Sub GameDelay(ByRef AtSpeed As Single)
		' Waits an amount of time.
		Dim SlowDown As Single
		SlowDown = VB.Timer() + ((21 - GlobalGameSpeed) / 3) / AtSpeed
		Do While VB.Timer() < SlowDown
		Loop 
	End Sub
	
	Private Sub GameSpeed(ByRef AddTo As Short)
		GlobalGameSpeed = LoopNumber(1, 20, GlobalGameSpeed, AddTo)
		MessageShow("Game Speed now set at " & 21 - GlobalGameSpeed, 0)
	End Sub
	
	Private Sub CombatWait()
		Dim c As Short
		Dim GridX As Grid
		'UPGRADE_WARNING: Couldn't resolve default property of object GridX. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
		GridX = Turns(TurnNow)
		For c = 0 To CombatTurn - 1
			If c >= TurnNow Then
				'UPGRADE_WARNING: Couldn't resolve default property of object Turns(c). Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
				Turns(c) = Turns(c + 1)
			End If
		Next c
		'UPGRADE_WARNING: Couldn't resolve default property of object Turns(CombatTurn). Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
		Turns(CombatTurn) = GridX
		CreatureWithTurn.CombatAttitude = bdCombatAttitudeWait
		CombatNextTurn()
	End Sub
	
	Public Sub CombatNextTurn()
		Dim PlayerTurn, c, n, NoFail As Short
		Dim rc As Integer
		Dim Y, X, Found As Short
		Dim CreatureX As Creature
		PlayerTurn = False
		Do 
			Frozen = True
			If CreatureWithTurn.CombatAttitude <> bdCombatAttitudeWait Then
				' Fire Post-Turn Trigger on CreatureWithTurn
				CreatureNow = CreatureWithTurn
				ItemNow = CreatureWithTurn.ItemInHand
				NoFail = FireTriggers(CreatureWithTurn, bdPostTurn)
				For	Each ItemNow In CreatureWithTurn.Items
					If ItemNow.IsReady = True Then
						CreatureNow = CreatureWithTurn
						NoFail = FireTriggers(ItemNow, bdPostTurn)
					End If
				Next ItemNow
				' Go to the next turn
				c = -1
				Do 
					TurnNow = LoopNumber(0, CombatTurn, TurnNow, 1)
					If TurnNow = 0 Then
						'UPGRADE_WARNING: Screen property Screen.MousePointer has a new behavior. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6BA9B8D2-2A32-4B6E-8D36-44949974A5B4"'
						System.Windows.Forms.Cursor.Current = System.Windows.Forms.Cursors.Default
						TurnCycle()
						CombatArrays()
						'UPGRADE_WARNING: Screen property Screen.MousePointer has a new behavior. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6BA9B8D2-2A32-4B6E-8D36-44949974A5B4"'
						System.Windows.Forms.Cursor.Current = System.Windows.Forms.Cursors.WaitCursor
					End If
					c = c + 1
				Loop Until (Turns(TurnNow).Ref.HPNow > 0 And Turns(TurnNow).Ref.IsInanimate = False) Or CombatTurn < 0 Or c > CombatTurn
			Else
				CreatureWithTurn.CombatAttitude = bdCombatAttitudeWait2
			End If
			picGrid.Refresh()
			' Check if everyone is dead
			If c > CombatTurn Then
				CombatEnd(False)
				Exit Sub
			End If
			GameDelay(5)
			System.Windows.Forms.Application.DoEvents()
			' Set up for next turn
			CreatureWithTurn = Turns(TurnNow).Ref
			Target = -1 : LastTarget = -1 : CreatureWithTurn.MsgQueTop = -1
			CombatDrawAttack(CreatureWithTurn, CreatureWithTurn, False)
			GlobalAttackRoll = 0 : GlobalArmorRoll = 0 : GlobalDamageRoll = 0 : GlobalHitLocation = ""
			For c = 0 To 5 : CombatDiceType(c) = 0 : CombatDiceValue(c) = 0 : Next c
			Combat20Bonus = 0 : CombatDmgBonus = 0
			CombatRectangle()
			' Play the Next Turn tone
			CombatDraw()
			
			If CreatureWithTurn.DMControlled = False Then
                '            Call PlaySoundFile(gAppPath & "\data\stock\next.wav", tome, True)
                Call PlaySoundFile(gDataPath & "\stock\next.wav", tome, True)
			Else
                '            Call PlaySoundFile(gAppPath & "\data\stock\next2.wav", tome, True)
                Call PlaySoundFile(gDataPath & "\stock\next2.wav", tome, True)
			End If
			System.Windows.Forms.Application.DoEvents()
			' Fire Pre-Turn Trigger on CreatureWithTurn
			If CreatureWithTurn.CombatAttitude <> bdCombatAttitudeWait2 Then
				' Clear all their stats
				TurnCycleCreatureClear(CreatureWithTurn)
				CreatureNow = CreatureWithTurn
				ItemNow = CreatureWithTurn.ItemInHand
				NoFail = FireTriggers(CreatureNow, bdPreTurn)
				For	Each ItemNow In CreatureNow.Items
					If ItemNow.IsReady = True Then
						CreatureNow = CreatureWithTurn
						NoFail = FireTriggers(ItemNow, bdPreTurn)
					End If
				Next ItemNow
			Else
				CreatureWithTurn.CombatAttitude = 0
			End If
			CombatDraw()
			' Check if no non-DMControlled Creatures are left
			For c = 0 To CombatTurn
				If Turns(c).Ref.DMControlled = False Then
					Exit For
				End If
			Next c
			If c > CombatTurn Then
				CombatEnd(False)
				Exit Do
			End If
			' Check to see if Frozen or Resting
			If CreatureWithTurn.AllowedTurn = False Then
				If CreatureWithTurn.Frozen = True Then
					CombatMessage(CreatureWithTurn.Name & " can't move!")
					GameDelay(2)
				End If
				If CreatureWithTurn.Unconscious = True Then
					CombatMessage(CreatureWithTurn.Name & " is unconscious!")
					GameDelay(2)
				End If
				If CreatureWithTurn.Afraid = True Then
					CombatMessage(CreatureWithTurn.Name & " is afraid!")
					GameDelay(2)
					CombatMove(MOVEDIRECTION.bdMoveFlee, MOVEDIRECTION.bdMoveFlee)
				End If
			Else
				' If no Unfriendly left alive, end combat (you won or lost)
				If CombatAllDead = True Then
					Found = True
					For	Each CreatureX In EncounterNow.Creatures
						If CreatureX.HPNow > 0 Then
							Found = False
							Exit For
						End If
					Next CreatureX
					CombatEnd(Found)
					Exit Do
				Else
					' Play the turn
					If CreatureWithTurn.DMControlled = True Then
						'UPGRADE_WARNING: Screen property Screen.MousePointer has a new behavior. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6BA9B8D2-2A32-4B6E-8D36-44949974A5B4"'
						System.Windows.Forms.Cursor.Current = System.Windows.Forms.Cursors.WaitCursor
						SorceryInvokeNext(CreatureWithTurn)
						Me.Refresh()
						CombatMonster()
					Else
						Frozen = False
						SorceryInvokeNext(CreatureWithTurn)
						Me.Refresh()
						PlayerTurn = True
						picMap.Focus()
					End If
				End If
			End If
			'UPGRADE_WARNING: Screen property Screen.MousePointer has a new behavior. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6BA9B8D2-2A32-4B6E-8D36-44949974A5B4"'
			System.Windows.Forms.Cursor.Current = System.Windows.Forms.Cursors.Default
		Loop Until PlayerTurn = True
	End Sub
	
	Private Sub CombatPaintSort()
		Dim i, c, k As Short
		' Sort CreatureWithTurn to place in stack
		For c = 0 To CombatTurn
			For i = c + 1 To CombatTurn
				If Turns(Paint_Renamed(c)).Ref.Row < Turns(Paint_Renamed(i)).Ref.Row - CShort(Turns(Paint_Renamed(i)).Ref.HPNow < 1) Then
					k = Paint_Renamed(c)
					Paint_Renamed(c) = Paint_Renamed(i)
					Paint_Renamed(i) = k
				End If
			Next i
		Next c
	End Sub
	
	Private Sub CombatRectangle()
		Dim c, n As Short
		Dim rc As Integer
		Dim GridWidth As Double
		'UPGRADE_NOTE: Width was upgraded to Width_Renamed. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="A9E4979A-37FA-4718-9994-97DD76ED70A7"'
		'UPGRADE_NOTE: Height was upgraded to Height_Renamed. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="A9E4979A-37FA-4718-9994-97DD76ED70A7"'
		Dim X, Height_Renamed, Width_Renamed, Y As Short
		Dim Distance2, Found, Distance1, MaxRange As Short
		Dim Movement As Short
		' Exit this if CreatureWithTurn is DM Controlled
		If CreatureWithTurn.DMControlled = True Then
			Exit Sub
		End If
		' Find Path
		CursorToGrid(CombatMouseX, CombatMouseY, Y, X)
		CombatFindPath(CreatureWithTurn, X, Y)
		' Determine if cursor on Target
		Found = False
		For c = CombatTurn To 0 Step -1
			If Turns(Paint_Renamed(c)).Ref.HPNow > 0 Or TargetAllowDead = True Then
				' Get coordinates
				GridToCursor(X, Y, Turns(Paint_Renamed(c)).Ref.Row, Turns(Paint_Renamed(c)).Ref.Col, GridWidth)
				' Get exact dimensions for picture
				n = Turns(Paint_Renamed(c)).Ref.Pic
				Height_Renamed = Int(picCPic(n).Height / 4)
				Width_Renamed = Int(picCPic(n).Width / 2)
				' Center the picture in the square
				Y = Int(Y - Height_Renamed)
				X = X - (Width_Renamed - GridWidth) / 2
				' Determine if cursor in picture
				If PointIn(CombatMouseX, CombatMouseY, X, Y, Width_Renamed, Height_Renamed) Then
					'UPGRADE_ISSUE: PictureBox property picCPic.hdc was not upgraded. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="CC4C7EC0-C903-48FC-ACCC-81861D12DA4A"'
					If GetPixel(picCPic(n).hdc, CombatMouseX - X + Turns(Paint_Renamed(c)).Ref.Facing * Width_Renamed, CombatMouseY - Y + Height_Renamed) = 0 Then
						CreatureTarget = Turns(Paint_Renamed(c)).Ref
						' Determine if valid to target this creature
						Found = False
						Select Case MenuNow
							Case bdMenuTargetCreature
								If CreatureWithTurn.Friendly <> CreatureTarget.Friendly Then
									Found = True
								End If
							Case bdMenuTargetAny
								Found = True
							Case bdMenuTargetParty
								If CreatureWithTurn.Friendly = CreatureTarget.Friendly Then
									Found = True
								End If
							Case Else
								Found = True
						End Select
						If Found = True Then
							Target = Paint_Renamed(c)
							If Target <> LastTarget Then
								Call PlayClickSnd(modIOFunc.ClickType.ifClickPass)
								LastTarget = Paint_Renamed(c)
								CombatDrawAttack(CreatureWithTurn, CreatureTarget, False)
								CombatFindPath(CreatureWithTurn, CreatureTarget.Col, CreatureTarget.Row)
								CombatDraw(1)
							End If
						Else
							Target = -1
						End If
						Exit Sub
					End If
				End If
			End If
		Next c
		If LastTarget <> -1 Then
			CombatDrawAttack(CreatureWithTurn, CreatureWithTurn, False)
		End If
		Target = -1 : LastTarget = -1
		CombatDraw()
	End Sub
	
	Public Sub CombatArrays()
		Dim i, c, Found As Short
		Dim GridNow(bdCombatWidth, bdCombatHeight) As Short
		Dim HereX(bdCombatWidth * bdCombatHeight) As Short
		Dim HereY(bdCombatWidth * bdCombatHeight) As Short
		'UPGRADE_NOTE: Top was upgraded to Top_Renamed. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="A9E4979A-37FA-4718-9994-97DD76ED70A7"'
		Dim Top_Renamed As Short
		Dim CreatureX As Creature
		Dim GridX As Grid
		' Clear combat pointers
		TurnNow = 0 : Target = -1 : CombatTurn = -1
		' Que all the spaces looking for at least one starting space
		Found = False
		For c = 0 To bdCombatHeight : For i = 0 To bdCombatWidth
				GridNow(i, c) = EncounterNow.CombatGrid(i, c)
				If GridNow(i, c) > 1 Then
					Found = True
				End If
			Next i : Next c
		' If no starting spaces, create some
		If Found = False Then
			Select Case Int(Rnd() * 3)
				Case 0 ' Mons Right, Chars Left
					For c = 0 To bdCombatHeight : For i = 0 To bdCombatWidth
							If i < 5 Then
								GridNow(i, c) = 2
							ElseIf i > 5 Then 
								GridNow(i, c) = 3
							End If
						Next i : Next c
				Case 1 ' Monst Top, Chars Bottom
					For c = 0 To bdCombatHeight : For i = 0 To bdCombatWidth
							If c < 3 Then
								GridNow(i, c) = 2
							ElseIf i > 3 Then 
								GridNow(i, c) = 3
							End If
						Next i : Next c
				Case 2 ' Chars Top, Mons Bottom
					For c = 0 To bdCombatHeight : For i = 0 To bdCombatWidth
							If c < 3 Then
								GridNow(i, c) = 3
							ElseIf i > 3 Then 
								GridNow(i, c) = 2
							End If
						Next i : Next c
			End Select
		End If
		' Find spot for all party members
        For Each CreatureX In tome.Creatures
            ' If not afraid or not on the edge of the combat area
            If (CreatureX.Afraid = False Or IsBetween(CreatureX.Col, 1, bdCombatWidth - 1)) And (CreatureX.IsInanimate = False Or CreatureX.HPNow > 0) Then
                If CreatureX.Initiative = 0 Then
                    ' Find a spot for it
                    Top_Renamed = 0
                    For c = 0 To bdCombatHeight : For i = 0 To bdCombatWidth
                            If GridNow(i, c) = 2 Then ' Party starting spot
                                HereX(Top_Renamed) = i : HereY(Top_Renamed) = c
                                Top_Renamed = Top_Renamed + 1
                                Found = True
                            End If
                        Next i : Next c
                    ' Find random spot
                    If Top_Renamed = 0 Then
                        ' Secondary look is to open spots
                        For c = 0 To bdCombatHeight : For i = 0 To bdCombatWidth
                                If GridNow(i, c) = 0 Then ' Open starting spot
                                    HereX(Top_Renamed) = i : HereY(Top_Renamed) = c
                                    Top_Renamed = Top_Renamed + 1
                                End If
                            Next i : Next c
                    End If
                    If Top_Renamed > 0 Then
                        CreatureX.Col = HereX(Int(Rnd() * Top_Renamed))
                        CreatureX.Row = HereY(Int(Rnd() * Top_Renamed))
                        ' Set initiative
                        CreatureX.Initiative = Int(Rnd() * 6) + 1
                    End If
                End If
                If CreatureX.Initiative > 0 And CombatTurn < 48 Then
                    ' Add to combat que
                    CombatTurn = Least(CombatTurn + 1, 48)
                    CreatureX.Friendly = True
                    Turns(CombatTurn).Ref = CreatureX
                    Paint_Renamed(CombatTurn) = CombatTurn
                    ' Mark grid space occupied
                    GridNow(CreatureX.Col, CreatureX.Row) = True
                End If
            End If
        Next CreatureX
		' Find spot for all Monsters
		For	Each CreatureX In EncounterNow.Creatures
			' If not afraid or not on the edge of the combat area
			If (CreatureX.Afraid = False Or IsBetween(CreatureX.Col, 1, bdCombatWidth - 1)) And (CreatureX.IsInanimate = False Or CreatureX.HPNow > 0) Then
				If CreatureX.Initiative = 0 Then
					' Find a spot for it
					Top_Renamed = 0
					For c = 0 To bdCombatHeight : For i = 0 To bdCombatWidth
							If GridNow(i, c) = 3 Then ' Monster starting spot
								HereX(Top_Renamed) = i : HereY(Top_Renamed) = c
								Top_Renamed = Top_Renamed + 1
								Found = True
							End If
						Next i : Next c
					' Find random spot
					If Top_Renamed = 0 Then
						' Secondary look is to open spots
						For c = 0 To bdCombatHeight : For i = 0 To bdCombatWidth
								If GridNow(i, c) = 0 Then ' Open starting spot
									HereX(Top_Renamed) = i : HereY(Top_Renamed) = c
									Top_Renamed = Top_Renamed + 1
								End If
							Next i : Next c
					End If
					If Top_Renamed > 0 Then
						CreatureX.Col = HereX(Int(Rnd() * Top_Renamed))
						CreatureX.Row = HereY(Int(Rnd() * Top_Renamed))
						' Set initiative
						CreatureX.Initiative = Int(Rnd() * 6) + 1
					End If
				End If
				If CreatureX.Initiative > 0 And CombatTurn < 48 Then
					' Add to combat que
					CombatTurn = Least(CombatTurn + 1, 48)
					CreatureX.Friendly = False
					Turns(CombatTurn).Ref = CreatureX
					Paint_Renamed(CombatTurn) = CombatTurn
					' Mark grid space occupied
					GridNow(CreatureX.Col, CreatureX.Row) = True
				End If
			ElseIf CreatureX.Afraid = True Then 
				EncounterNow.RemoveCreature("X" & CreatureX.Index)
			End If
		Next CreatureX
		' Sort by Initiative
		For c = 0 To CombatTurn
			For i = c + 1 To CombatTurn
				If Turns(c).Ref.Initiative < Turns(i).Ref.Initiative Then
					'UPGRADE_WARNING: Couldn't resolve default property of object GridX. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
					GridX = Turns(c)
					'UPGRADE_WARNING: Couldn't resolve default property of object Turns(c). Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
					Turns(c) = Turns(i)
					'UPGRADE_WARNING: Couldn't resolve default property of object Turns(i). Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
					Turns(i) = GridX
				End If
			Next i
		Next c
	End Sub
	
	
	Private Sub PlayMusicState()
		Dim rc As Integer
		GlobalMusicState = 1 - GlobalMusicState
		Select Case GlobalMusicState
			Case 0
				MessageShow("Music is now OFF", 0)
				'MediaPlayerMusic.Stop
				oGameMusic.StopPlay()
				oGameMusic.mciClose()
			Case 1
				MessageShow("Music is now ON", 0)
				Call PlayMusicRnd(IIf(picGrid.Visible, modIOFunc.RNDMUSICSTYLE.Combat, modIOFunc.RNDMUSICSTYLE.Adventure), Me)
		End Select
	End Sub
	
	Private Sub PlaySoundState()
		Dim rc As Integer
		GlobalWAVState = 1 - GlobalWAVState
		Select Case GlobalWAVState
			Case 0
				MessageShow("Sound Effects are now OFF", 0)
			Case 1
				MessageShow("Sound Effects are now ON", 0)
		End Select
	End Sub
	
	Public Sub CombatStart()
		Dim LastMons, c, i, Found As Short
		Dim NoFail As Short
		Dim CreatureX As Creature
		' Close down current menu, Convo and Talk (if any) and stop Party
		picConvo.Visible = False
		picTalk.Visible = False
		tmrMoveParty.Enabled = False
		' Check to see if Combat is allowed in this Encounter: at least one Creature must be
		' alive and the Encounter must allow combat.
		Found = False
		For	Each CreatureX In EncounterNow.Creatures
			If CreatureX.HPNow > 0 Then
				Found = True
				Exit For
			End If
		Next CreatureX
		If EncounterNow.CanFight = False Or Found = False Then
			DialogDM("There's no reason to fight here.")
			DialogHide()
			Exit Sub
		End If
		' Load Wallpaper for Combat
		'UPGRADE_WARNING: Screen property Screen.MousePointer has a new behavior. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6BA9B8D2-2A32-4B6E-8D36-44949974A5B4"'
		System.Windows.Forms.Cursor.Current = System.Windows.Forms.Cursors.WaitCursor
		CombatWallpaper()
		picToHit.Left = (Me.ClientRectangle.Width - picToHit.ClientRectangle.Width) / 2
		' Play combat music
		Call PlayMusicRnd(modIOFunc.RNDMUSICSTYLE.Combat, Me)
		' Clear Initiative for all Creatures
        For Each CreatureX In tome.Creatures
            CreatureX.Initiative = 0
        Next CreatureX
		For	Each CreatureX In EncounterNow.Creatures
			CreatureX.Initiative = 0
		Next CreatureX
		' Set up Combat Arrays with Creatures
		CombatTurn = -1
		CombatArrays()
		' If nobody is alive then cancel out
		If CombatTurn < 0 Then
			DialogDM("There's nobody alive to fight.")
			Exit Sub
		End If
		' Fire up Pre-Combat Trigger (if any)
        NoFail = FireTriggers(tome, bdPreCombat)
		NoFail = FireTriggers(EncounterNow, bdPreCombat)
		' Show menu, ToHit slab, Grid and Menu orbs
		picMenu.Visible = False
		picToHit.Visible = True
		picGrid.Visible = True
		picGrid.BringToFront()
		' Blank out CreatureWithTurn while starting combat
		CreatureWithTurn = New Creature
		' No turns or Message at combat start
		TurnNow = -1
		CombatDraw()
		BorderDrawButtons(0)
		CombatNextTurn()
		Me.Refresh()
		'UPGRADE_WARNING: Screen property Screen.MousePointer has a new behavior. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6BA9B8D2-2A32-4B6E-8D36-44949974A5B4"'
		System.Windows.Forms.Cursor.Current = System.Windows.Forms.Cursors.Default
	End Sub
	
	Public Function CombatRollArmor(ByRef Defender As Creature, ByVal RollIt As Short, ByVal SetValue As Object) As Short
		' This rolls armor to hit: flat display, roll and effect. It
		' returns the index of the ArmorChit hit.
		' NOTE: Armor Roll ranges from 0 to 7
		Dim n, rc, c, i, k As Short
		' If in combat, show rolling armor. Else, just set it
		If picGrid.Visible = True Then
			' Show all Armor Chits on Defender: Covered 0, 1-2, 3-4, 5-6 or 7-8
			' If 0-25 = Wooden, 26-74 = Silver, >75 = Golden.
			n = 0 : k = 0
			For c = 0 To 7
				If Defender.ResistanceWithArmor(c) > 0 Then
					n = n + 1
					k = k + Defender.ResistanceWithArmor(c)
				End If
			Next c
			If n > 0 Then
				If (k / n) < 1 Then
					k = -1 ' None
				ElseIf (k / n) < 26 Then 
					k = 0 ' Wooden
				ElseIf (k / n) < 76 Then 
					k = 1 ' Silver
				Else
					k = 2 ' Gold
				End If
			End If
			n = Int(n / 2 + 0.5)
			' Display (c = Number of Shields-1, k = Type of Shield)
			If k > -1 Then
				For c = 0 To n - 1
					'UPGRADE_ISSUE: PictureBox property picMisc.hdc was not upgraded. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="CC4C7EC0-C903-48FC-ACCC-81861D12DA4A"'
					'UPGRADE_ISSUE: PictureBox property picToHit.hdc was not upgraded. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="CC4C7EC0-C903-48FC-ACCC-81861D12DA4A"'
					rc = BitBlt(picToHit.hdc, 268 + c * 36, 2, 32, 32, picMisc.hdc, 96, 197, SRCAND)
					'UPGRADE_ISSUE: PictureBox property picMisc.hdc was not upgraded. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="CC4C7EC0-C903-48FC-ACCC-81861D12DA4A"'
					'UPGRADE_ISSUE: PictureBox property picToHit.hdc was not upgraded. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="CC4C7EC0-C903-48FC-ACCC-81861D12DA4A"'
					rc = BitBlt(picToHit.hdc, 268 + c * 36, 2, 32, 32, picMisc.hdc, k * 32, 197, SRCPAINT)
				Next c
			End If
			If RollIt Then
				n = Int(Rnd() * 5) + 5
				For i = 0 To n
					'UPGRADE_WARNING: Couldn't resolve default property of object SetValue. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
					If i = n And SetValue > 0 Then
						'UPGRADE_WARNING: Couldn't resolve default property of object SetValue. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
						c = Least(Greatest(SetValue - 1, 0), 7)
					Else
						c = Int(Rnd() * 8)
					End If
				Next i
			End If
			CombatRollArmor = c
			'UPGRADE_WARNING: Couldn't resolve default property of object SetValue. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
		ElseIf SetValue > 0 Then 
			'UPGRADE_WARNING: Couldn't resolve default property of object SetValue. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
			CombatRollArmor = Least(Greatest(SetValue - 1, 0), 7)
		Else
			CombatRollArmor = Int(Rnd() * 8)
		End If
	End Function
	
	Private Sub DiceDraw(ByRef DicePos As Short, ByRef DiceType As Short, ByRef DiceValue As Short, ByRef InCombat As Short)
		Dim rc As Integer
		Dim Y, X, c As Short
		c = DiceValue - 1
		If DiceValue > 0 Then
			Select Case DiceType
				Case 20
					X = (c Mod 8) * 40 : Y = Int(c / 8) * 40
				Case 12
					If DiceValue < 7 Then
						X = 40 * 2 + c * 40 : Y = 7 * 40
					Else
						X = (c - 6) * 40 : Y = 8 * 40
					End If
				Case 10
					If c < 8 Then
						X = c * 40 : Y = 6 * 40
					Else
						X = (c - 8) * 40 : Y = 7 * 40
					End If
				Case 8
					X = c * 40 : Y = 120
				Case 6
					X = c * 40 : Y = 200
				Case 4
					X = 120 + c * 40 : Y = 160
			End Select
		Else
			Select Case DiceType
				Case 20
					X = 40 * 4 + System.Math.Abs(c + 2) * 40 : Y = 2 * 40
				Case 12
					X = 40 * 6 + System.Math.Abs(c + 2) * 40 : Y = 8 * 40
				Case 10
					X = System.Math.Abs(c + 2) * 39 : Y = 4 * 40
				Case 8
					X = System.Math.Abs(c + 2) * 39 : Y = 4 * 40
				Case 6
					X = 40 * 6 + System.Math.Abs(c + 2) * 40 : Y = 5 * 40
				Case 4
					X = 40 * 4 : Y = 4 * 40
			End Select
		End If
		If InCombat = True Then
			If DicePos = 0 Then
				'UPGRADE_ISSUE: PictureBox property picDice.hdc was not upgraded. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="CC4C7EC0-C903-48FC-ACCC-81861D12DA4A"'
				'UPGRADE_ISSUE: PictureBox property picToHit.hdc was not upgraded. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="CC4C7EC0-C903-48FC-ACCC-81861D12DA4A"'
				rc = BitBlt(picToHit.hdc, 132, 40, 40, 40, picDice.hdc, X + picDice.ClientRectangle.Width / 2, Y, SRCAND)
				'UPGRADE_ISSUE: PictureBox property picDice.hdc was not upgraded. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="CC4C7EC0-C903-48FC-ACCC-81861D12DA4A"'
				'UPGRADE_ISSUE: PictureBox property picToHit.hdc was not upgraded. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="CC4C7EC0-C903-48FC-ACCC-81861D12DA4A"'
				rc = BitBlt(picToHit.hdc, 132, 40, 40, 40, picDice.hdc, X, Y, SRCPAINT)
			Else
				'UPGRADE_ISSUE: PictureBox property picDice.hdc was not upgraded. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="CC4C7EC0-C903-48FC-ACCC-81861D12DA4A"'
				'UPGRADE_ISSUE: PictureBox property picToHit.hdc was not upgraded. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="CC4C7EC0-C903-48FC-ACCC-81861D12DA4A"'
				rc = BitBlt(picToHit.hdc, 134 + DicePos * 46, 40, 40, 40, picDice.hdc, X + picDice.ClientRectangle.Width / 2, Y, SRCAND)
				'UPGRADE_ISSUE: PictureBox property picDice.hdc was not upgraded. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="CC4C7EC0-C903-48FC-ACCC-81861D12DA4A"'
				'UPGRADE_ISSUE: PictureBox property picToHit.hdc was not upgraded. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="CC4C7EC0-C903-48FC-ACCC-81861D12DA4A"'
				rc = BitBlt(picToHit.hdc, 134 + DicePos * 46, 40, 40, 40, picDice.hdc, X, Y, SRCPAINT)
			End If
		Else
			'UPGRADE_ISSUE: PictureBox property picDice.hdc was not upgraded. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="CC4C7EC0-C903-48FC-ACCC-81861D12DA4A"'
			'UPGRADE_ISSUE: PictureBox property picConvo.hdc was not upgraded. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="CC4C7EC0-C903-48FC-ACCC-81861D12DA4A"'
			rc = BitBlt(picConvo.hdc, 60 + DicePos * 46, picConvo.Height - 52, 40, 40, picDice.hdc, X + picDice.ClientRectangle.Width / 2, Y, SRCAND)
			'UPGRADE_ISSUE: PictureBox property picDice.hdc was not upgraded. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="CC4C7EC0-C903-48FC-ACCC-81861D12DA4A"'
			'UPGRADE_ISSUE: PictureBox property picConvo.hdc was not upgraded. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="CC4C7EC0-C903-48FC-ACCC-81861D12DA4A"'
			rc = BitBlt(picConvo.hdc, 60 + DicePos * 46, picConvo.Height - 52, 40, 40, picDice.hdc, X, Y, SRCPAINT)
		End If
	End Sub
	
	Public Function DiceRoll(ByRef DiceCnt As Short, ByRef DiceToRoll As Short, ByRef Bonus As Short, ByRef Attack As Short, ByRef InCombat As Short) As Short
		Dim X, c, n, Y As Short
		Dim FinalRoll, MaxTimes, RollNow, TotalDice As Short
		' Check for invalid dice and make valid
		DiceCnt = Greatest(Least(DiceCnt, 5), 1)
		If DiceToRoll <> 4 And DiceToRoll <> 6 And DiceToRoll <> 8 And DiceToRoll <> 10 And DiceToRoll <> 12 And DiceToRoll <> 20 Then
			DiceToRoll = 4
		End If
		' Set up number of times to roll
		If DiceCnt = 1 And Attack = True And InCombat = True Then
			c = CombatRollArmor(CreatureTarget, False, 0)
		End If
		MaxTimes = Int(Rnd() * 3) + 3
		' Clear out previous rolls if not in combat (in combat, we want to retain our rolls for display)
		If InCombat = False Then
			For n = 0 To 5
				CombatDiceType(n) = 0
			Next n
		End If
		' Bounce until get to final roll
		TotalDice = 0
		Do Until TotalDice > 0
			Sleep(10)
			For c = 1 To DiceCnt
				' Draw either a rolling die or a numbered die
				If Int(Rnd() * 100) < 50 Or MaxTimes < 1 Or GlobalDiceRolling = 0 Then
					FinalRoll = Int(Rnd() * DiceToRoll)
					RollNow = FinalRoll + 1
				Else
					' Roll a spinning die
					Select Case DiceToRoll
						Case 20
							RollNow = -Int(Rnd() * 4)
						Case 12
							RollNow = -Int(Rnd() * 2)
						Case 10
							RollNow = -Int(Rnd() * 3)
						Case 8
							RollNow = -Int(Rnd() * 3)
						Case 6
							RollNow = -Int(Rnd() * 2)
						Case 4
							RollNow = 0
					End Select
				End If
				If DiceCnt = 1 And Attack = True And InCombat = True Then
					CombatDiceType(0) = DiceToRoll : CombatDiceValue(0) = RollNow
				Else
					CombatDiceType(c) = DiceToRoll : CombatDiceValue(c) = RollNow
				End If
				If InCombat = True Then
					CombatDrawAttack(CreatureWithTurn, CreatureTarget, True)
				End If
				' Store away the final rolls
				If MaxTimes > 0 And GlobalDiceRolling = 1 Then
					If c = DiceCnt Then
						MaxTimes = MaxTimes - 1
					End If
				Else
					TotalDice = TotalDice + FinalRoll + 1
				End If
			Next c
		Loop 
		' Draw dice in DialogDice
		If InCombat = False Then
			For n = 0 To 5
				If CombatDiceType(n) <> 0 Then
					DiceDraw(n, CombatDiceType(n), CombatDiceValue(n), False)
				End If
			Next n
			picConvo.Refresh()
		Else
			CombatDrawAttack(CreatureWithTurn, CreatureTarget, True)
		End If
		' Play roll sound
		If GlobalDiceRolling = 1 Then
			If DiceCnt > 1 Then
                Call PlaySoundFile("manydice", tome, , IIf(InCombat = True, 5, 0))
			Else
                Call PlaySoundFile("diceroll", tome, , IIf(InCombat = True, 5, 0))
			End If
		End If
		' Add Bonus if necessary
		If Bonus <> 0 And InCombat = True Then
			If Attack = True Then
				' Attack Bonus
				If Bonus < 0 Then
					ShowText(picToHit, 132, 40, 40, 10, bdFontSmallWhite, VB6.Format(Bonus), 1, False)
				Else
					ShowText(picToHit, 132, 40, 40, 10, bdFontSmallWhite, "+" & VB6.Format(Bonus), 1, False)
				End If
			Else
				' Damage Bonus
				If Bonus < 0 Then
					ShowText(picToHit, 136 + DiceCnt * 46, 40, 40, 10, bdFontSmallWhite, VB6.Format(Bonus), 1, False)
				Else
					ShowText(picToHit, 136 + DiceCnt * 46, 40, 40, 10, bdFontSmallWhite, "+" & VB6.Format(Bonus), 1, False)
				End If
			End If
			picToHit.Refresh()
		End If
		' Set the global dice attributes
		GlobalDieTypeRoll = DiceToRoll
		GlobalDieCountRoll = DiceCnt
		If DiceToRoll = 20 Then
			Combat20 = TotalDice
			Combat20Bonus = Bonus
		End If
		' Return number rolled
		DiceRoll = TotalDice + Bonus
	End Function
	
	Public Sub AwardExperienceToParty(ByRef ExpPts As Integer)
		Dim c As Short
		Dim CreatureX As Creature
		' Find how many Creatures are still alive in the Party
		c = 0
        For Each CreatureX In tome.Creatures
            If CreatureX.HPNow > 0 Then
                c = c + 1
            End If
        Next CreatureX
		' If there are live, the divide up the Experience Points
		If c > 0 Then
			ExpPts = Fix(ExpPts / c)
			'UPGRADE_WARNING: Screen property Screen.MousePointer has a new behavior. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6BA9B8D2-2A32-4B6E-8D36-44949974A5B4"'
			System.Windows.Forms.Cursor.Current = System.Windows.Forms.Cursors.Default
			DialogDM("Each Party member receives " & ExpPts & " Experience Points!")
            For Each CreatureX In tome.Creatures
                AwardExperienceToCreature(CreatureX, ExpPts)
            Next CreatureX
		End If
	End Sub
	
	Public Sub AwardExperienceToCreature(ByRef CreatureX As Creature, ByRef ExpPts As Integer)
		Dim NextLevel As Integer
		Dim AwardSkill, c, AwardHP As Short
		' You can only award experience to live creatues
		If CreatureX.HPNow > 0 Then
			NextLevel = CreatureX.ExperiencePoints + ExpPts
			CreatureX.ExperiencePoints = NextLevel
			c = CreatureX.Level : AwardSkill = 0 : AwardHP = 0
			' Keep raising until correct level for ExperiencePoints
			Do 
				If NextLevel > (c * Int(c / 2 + 0.5) + (c / 2) * (1 - (c Mod 2))) * 1000 - 1 Then
					c = c + 1
					AwardSkill = AwardSkill + WorldNow.SkillPtsPerLevel ' 10
					AwardHP = AwardHP + CShort(WorldNow.HPPerLevel * 0.5) '10
					' [Titi 2.4.9] customize HP and SkillPoints depending on world settings
					If WorldNow.RandomizeHP Then
						AwardHP = AwardHP + CShort(Rnd(1) * 0.5 * WorldNow.HPPerLevel)
					Else
						AwardHP = AwardHP + CShort(WorldNow.HPPerLevel * 0.5)
					End If
				Else
					Exit Do
				End If
			Loop 
			If c > CreatureX.Level Then
				CreatureX.Level = c
				c = CreatureX.SkillPoints + AwardSkill
				CreatureX.SkillPoints = c
				c = CreatureX.HPNow + AwardHP
				CreatureX.HPNow = c
				c = CreatureX.HPMax + AwardHP
				CreatureX.HPMax = c
				'UPGRADE_WARNING: Screen property Screen.MousePointer has a new behavior. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6BA9B8D2-2A32-4B6E-8D36-44949974A5B4"'
				System.Windows.Forms.Cursor.Current = System.Windows.Forms.Cursors.Default
                Call PlaySoundFile("NewLevel", tome)
				DialogDM(CreatureX.Name & " is now Level " & CreatureX.Level & " and receives " & AwardSkill & " Skill Points and " & AwardHP & " Health Points!")
			End If
		End If
	End Sub
	
	Public Sub CombatEnd(ByRef GetExperience As Short)
		Dim NoFail, TotalLevels, AvgMonsLevel As Short
		Dim ExpPts As Integer
		Dim OldLevel, c, OldSkill As Short
		Dim CreatureX As Creature
		' Apply Experience (if didn't flee)
		ConvoAction = -1
		If GetExperience = True And EncounterNow.Creatures.Count() > 0 Then
			' Fire Post-Combat Triggers
            NoFail = FireTriggers(tome, bdPostCombat)
			NoFail = FireTriggers(EncounterNow, bdPostCombat)
			' Total up Experience Points and wipe Creatures
			ExpPts = 0
			For	Each CreatureX In EncounterNow.Creatures
				If CreatureX.HPNow < 1 Then
					ExpPts = ExpPts + CreatureX.ExperiencePoints
					CreatureX.ExperiencePoints = 0
				End If
			Next CreatureX
			' Award Experience Points
			If ExpPts > 0 Then
				AwardExperienceToParty(ExpPts)
			Else
				DialogDM("You receive NO Experience Points for winning this combat.")
				DialogHide()
			End If
			' If there are things to find, offer a search
			If EncounterNow.Items.Count() > 0 Then
				DialogSetUp(modGameGeneral.DLGTYPE.bdDlgNoReply)
				DialogSetButton(1, "Search")
				If EncounterNow.Items.Count() = 1 Then
					'UPGRADE_WARNING: Couldn't resolve default property of object EncounterNow.Items(1).Name. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
					'UPGRADE_WARNING: Couldn't resolve default property of object EncounterNow.Items().Name. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
					DialogShow("DM", "You notice " & AOrAn(EncounterNow.Items.Item(1).Name) & EncounterNow.Items.Item(1).Name & " in the area.")
				Else
					DialogShow("DM", "You notice several items in the area.")
				End If
				DialogHide()
			End If
		End If
		Frozen = False
		picToHit.Visible = False
		picMenu.Visible = True
		picGrid.Visible = False
		' Hide any stray DialogBrief and clear buffer
		picDialogBrief.Visible = False
		tmrDialogBrief.Enabled = False
		DialogBriefSet = New Collection
		' Refresh the screen
		Me.Refresh()
		'MediaPlayerMusic.Stop
		oGameMusic.StopPlay()
		oGameMusic.mciClose()
		' Clear all Creatures casting
        For Each CreatureX In tome.Creatures
            For c = 0 To 5
                CreatureX.Runes(c) = 0
            Next c
            CreatureX.RuneTop = 0
            CreatureX.RuneQueLimit = 0
        Next CreatureX
		For	Each CreatureX In EncounterNow.Creatures
			For c = 0 To 5
				CreatureX.Runes(c) = 0
			Next c
			CreatureX.RuneTop = 0
			CreatureX.RuneQueLimit = 0
		Next CreatureX
		' Refresh screen
		TurnCycle()
		MenuDrawParty()
		BorderDrawButtons(0)
		DrawMapAll()
		c = GameDeath
		If Not c And EncounterNow.Items.Count() > 0 And GetExperience = True Then
			MenuNow = bdMenuSearch
			DoAction(bdMenuActionDefault)
		End If
		Call PlayMusicRnd(modIOFunc.RNDMUSICSTYLE.Adventure, Me)
	End Sub
	
	Private Function GameDeath() As Short
		Dim CreatureX As Creature
		' Test if all Party is dead
		GameDeath = True
        For Each CreatureX In tome.Creatures
            If CreatureX.HPNow > 0 And CreatureX.DMControlled = False Then
                GameDeath = False
                Exit For
            End If
        Next CreatureX
		' If all dead then end game
		If GameDeath Then
            Call PlaySoundFile("partydead", tome)
			DialogSetUp(modGameGeneral.DLGTYPE.bdDlgNoReply)
			DialogSetButton(1, "Ok")
			DialogShow("DM", "Your Party has died.")
			DialogHide()
			GameEnd(False)
		End If
	End Function
	
	Private Sub DialogCredits()
		Dim PosY, LastY As Short
		Dim rc As Integer
		' Show the Close Credits
		DialogSetUp(modGameGeneral.DLGTYPE.bdDlgCredits)
		ShowText(picConvo, 0, 12, picConvo.Width, 14, bdFontElixirWhite, "Credits and Thanks", True, False)
		LastY = 32
		PosY = ShowText(picConvo, 0, LastY, picConvo.Width, 14, bdFontElixirWhite, "Original Concept and Design", True, True)
		LastY = LastY + PosY
		PosY = ShowText(picConvo, 0, LastY, picConvo.Width, 14, bdFontNoxiousWhite, "Dan Schnake and Adam West", True, True)
		LastY = LastY + PosY
		PosY = ShowText(picConvo, 0, LastY, picConvo.Width, 14, bdFontNoxiousWhite, "CrossCut Games, Inc.", True, True)
		LastY = LastY + PosY + 10
		PosY = ShowText(picConvo, 0, LastY, picConvo.Width, 14, bdFontElixirWhite, "Story Weaver and Documentation", True, True)
		LastY = LastY + PosY
		PosY = ShowText(picConvo, 0, LastY, picConvo.Width, 14, bdFontNoxiousWhite, "Dan Schnake", True, True)
		LastY = LastY + PosY + 10
		PosY = ShowText(picConvo, 0, LastY, picConvo.Width, 14, bdFontElixirWhite, "Programmer and Pixel Slinger", True, True)
		LastY = LastY + PosY
		PosY = ShowText(picConvo, 0, LastY, picConvo.Width, 14, bdFontNoxiousWhite, "Adam West", True, True)
		LastY = LastY + PosY + 10
		PosY = ShowText(picConvo, 0, LastY, picConvo.Width, 14, bdFontElixirWhite, "Additional Graphic Content", True, True)
		LastY = LastY + PosY
		PosY = ShowText(picConvo, 0, LastY, picConvo.Width, 14, bdFontNoxiousWhite, "Adam Krantz (Sarchimus)", True, True)
		LastY = LastY + PosY + 10
		PosY = ShowText(picConvo, 0, LastY, picConvo.Width, 14, bdFontElixirWhite, "RuneSword OpenSource Team", True, True)
		LastY = LastY + PosY
		PosY = ShowText(picConvo, 12, LastY, picConvo.Width - 24, 14, bdFontSmallWhite, "Project Manager :", 0, True)
		PosY = ShowText(picConvo, 12, LastY, picConvo.Width - 24, 14, bdFontSmallWhite, "VampiricDread (J.P. Carvalho)", 1, True)
		LastY = LastY + PosY
		PosY = ShowText(picConvo, 12, LastY, picConvo.Width - 24, 14, bdFontSmallWhite, "Content Management : ", 0, True)
		PosY = ShowText(picConvo, 12, LastY, picConvo.Width - 24, 14, bdFontSmallWhite, "Monty (Frank Montellano)    ", 1, True)
		LastY = LastY + PosY
		PosY = ShowText(picConvo, 12, LastY, picConvo.Width - 24, 14, bdFontSmallWhite, "Lead Developer :", 0, True)
		PosY = ShowText(picConvo, 12, LastY, picConvo.Width - 24, 14, bdFontSmallWhite, "Borfaux (Rob Blair)", 1, True)
		LastY = LastY + PosY
		PosY = ShowText(picConvo, 12, LastY, picConvo.Width - 24, 14, bdFontSmallWhite, "Developer & Documentation :", 0, True)
		PosY = ShowText(picConvo, 12, LastY, picConvo.Width - 24, 14, bdFontSmallWhite, "Titi (Thierry Petitjean)", 1, True)
		LastY = LastY + PosY
		PosY = ShowText(picConvo, 12, LastY, picConvo.Width - 24, 14, bdFontSmallWhite, "Developer : ", 0, True)
		PosY = ShowText(picConvo, 12, LastY, picConvo.Width - 24, 14, bdFontSmallWhite, "Phule (Robert Buchanan)", 1, True)
		LastY = LastY + PosY
		PosY = ShowText(picConvo, 12, LastY, picConvo.Width - 24, 14, bdFontSmallWhite, "Code & Testing :", 0, True)
		PosY = ShowText(picConvo, 12, LastY, picConvo.Width - 24, 14, bdFontSmallWhite, "Count0 (Todd Alexander)     ", 1, True)
		LastY = LastY + PosY
		PosY = ShowText(picConvo, 12, LastY, picConvo.Width - 24, 14, bdFontSmallWhite, "Documentation & Testing :", 0, True)
		PosY = ShowText(picConvo, 12, LastY, picConvo.Width - 24, 14, bdFontSmallWhite, "PaleoT (Travis Smith)     ", 1, True)
		LastY = LastY + PosY
		PosY = ShowText(picConvo, 12, LastY, picConvo.Width - 24, 14, bdFontSmallWhite, "Graphics Coordination :", 0, True)
		PosY = ShowText(picConvo, 12, LastY, picConvo.Width - 24, 14, bdFontSmallWhite, "Ephestion (Ephestion)", 1, True)
		LastY = LastY + PosY
		PosY = ShowText(picConvo, 12, LastY, picConvo.Width - 24, 14, bdFontSmallWhite, "Manuals :", 0, True)
		PosY = ShowText(picConvo, 12, LastY, picConvo.Width - 24, 14, bdFontSmallWhite, "Pat1 (Pat Roseinblume)", 1, True)
		LastY = LastY + PosY
		'PosY = ShowText(picConvo, 0, LastY + 10, picConvo.Width, 14, bdFontElixirWhite, "RuneSword on SourceForge : sourceforge.net-projects-runesword, True, True)
		'LastY = LastY + PosY
		LastY = picConvo.ClientRectangle.Height - 75
		PosY = ShowText(picConvo, 0, LastY, picConvo.Width, 14, bdFontNoxiousWhite, "CrossCut Games : www.crosscutgames.com", True, True)
		LastY = picConvo.ClientRectangle.Height - 48
		PosY = ShowText(picConvo, 0, LastY, picConvo.Width, 14, bdFontSmallWhite, "RuneSword Community Forums : www.runesword.com", True, True)
		DialogSetButton(1, "Wait")
		ShowButton(picConvo, picConvo.Width - 6 - 96, picConvo.Height - 30, "Wait", False)
		'UPGRADE_ISSUE: PictureBox property picConvoBottom.hdc was not upgraded. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="CC4C7EC0-C903-48FC-ACCC-81861D12DA4A"'
		'UPGRADE_ISSUE: PictureBox property picConvo.hdc was not upgraded. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="CC4C7EC0-C903-48FC-ACCC-81861D12DA4A"'
		rc = BitBlt(picConvo.hdc, 0, picConvo.Height - 6, picConvo.Width, 6, picConvoBottom.hdc, 0, 0, SRCCOPY)
		picConvo.Top = Greatest((picBox.ClientRectangle.Height - picConvo.ClientRectangle.Height) / 2, 0)
		picConvo.Refresh()
		picConvo.Visible = True
		picConvo.BringToFront()
		picConvo.Focus()
	End Sub
	
	Private Sub GameEndAll()
		Dim rc As Integer
		Call OptionsSave()
		'UPGRADE_WARNING: Screen property Screen.MousePointer has a new behavior. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6BA9B8D2-2A32-4B6E-8D36-44949974A5B4"'
		System.Windows.Forms.Cursor.Current = System.Windows.Forms.Cursors.WaitCursor
		'MediaPlayerMusic.Stop
		oGameMusic.StopPlay()
		Call oGameMusic.mciClose(True)
		'    rc = mciSendString("close songnow", 0&, 0, 0)
		picMainMenu.Visible = False
		' Free memory
		'UPGRADE_NOTE: Object Worlds may not be destroyed until it is garbage collected. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6E35BFF6-CD74-4B09-9689-3E1A43DF8969"'
		Worlds = Nothing
		'UPGRADE_NOTE: Object WorldNow may not be destroyed until it is garbage collected. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6E35BFF6-CD74-4B09-9689-3E1A43DF8969"'
		WorldNow = Nothing
		'UPGRADE_NOTE: Object Tome may not be destroyed until it is garbage collected. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6E35BFF6-CD74-4B09-9689-3E1A43DF8969"'
        tome = Nothing
		'UPGRADE_NOTE: Object UberWizMaps may not be destroyed until it is garbage collected. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6E35BFF6-CD74-4B09-9689-3E1A43DF8969"'
		UberWizMaps = Nothing
		'If (Val(GlobalCredits) = 1) Then
		DialogCredits()
		Me.Refresh()
		'Allow click to close
		DialogSetButton(1, "Done")
		ShowButton(picConvo, picConvo.Width - 6 - 96, picConvo.Height - 30, "Done", False)
		picConvo.Visible = True
		picConvo.BringToFront()
		picConvo.Focus()
		picConvo.Refresh()
		'UPGRADE_WARNING: Screen property Screen.MousePointer has a new behavior. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6BA9B8D2-2A32-4B6E-8D36-44949974A5B4"'
		System.Windows.Forms.Cursor.Current = System.Windows.Forms.Cursors.Default
		ConvoAction = -1
		Do Until ConvoAction > -1
			System.Windows.Forms.Application.DoEvents()
		Loop 
		ChangeScreenSettings(GlobalScreenX, GlobalScreenY, GlobalScreenColor)
		'End If
		'[borf] prevents PlaySoundFile error during about DoEvents
		'UPGRADE_NOTE: Object oFileSys may not be destroyed until it is garbage collected. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6E35BFF6-CD74-4B09-9689-3E1A43DF8969"'
		oFileSys = Nothing
		Me.Close()
	End Sub
	
	Private Sub TomeDo(ByRef Action As Short)
		Dim c As Short
		Dim rc As Integer
		Call PlayClickSnd(modIOFunc.ClickType.ifClickMenu)
		If TomeMenu = bdTomeOptions Then
			Select Case Action
				Case 0 ' Resume
					TomeMenu = bdNone
					picMainMenu.Visible = False
					picMap.Focus()
					Frozen = False
				Case 1 ' Saved Tomes
					TomeSavesLoad()
					TomeAction = bdTomeSaves
					TomeSavesList(1)
				Case 2 ' Save a new Tome
					TomeSavesLoad()
					TomeAction = bdTomeSaveAs
					TomeSavesList(0)
				Case 3 ' Options
					ScrollTop = 1
					OptionsShow(1)
				Case 4 ' Main Menu
					GameEnd(True)
			End Select
		Else
			Select Case Action
				Case 0 ' New Tome
					' Get Tome Listing
					TomeLoad()
					' Reset Tome Apprentice
					TomeWizardParty = 2
					TomeWizardLevel = 0
					TomeWizardStory = 0
					TomeWizardSize = 1
					' List Tomes
					TomeListTop = 1
					TomeNewList(1)
				Case 1 ' Saved Tomes
					TomeSavesLoad()
					TomeAction = bdTomeSaves
					TomeSavesList(1)
				Case 2 ' Characters
					CreatePCReturn = False
					CreatePCLoadMap(WorldNow)
					CreatePCKingdom(1)
					'                CreatePCKingdom worldindex
				Case 3 ' Options
					ScrollTop = 1
					OptionsShow(1)
				Case 4 ' Quit
					GameEndAll()
			End Select
		End If
	End Sub
	
	'UPGRADE_NOTE: Tome was upgraded to Tome_Renamed. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="A9E4979A-37FA-4718-9994-97DD76ED70A7"'
	Private Sub TomeSavesLoad()
		Dim Tome_Renamed As Object
		' Go to the /world directory and find all subdirectories that are "save*"
		Dim c As Short
		Dim tmpFullPath As String
		Dim FileName As String
		Dim TomeX As Tome
		Dim FactoidX As Factoid
		' Start new listing of Tomes on disk
		'UPGRADE_WARNING: Screen property Screen.MousePointer has a new behavior. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6BA9B8D2-2A32-4B6E-8D36-44949974A5B4"'
		System.Windows.Forms.Cursor.Current = System.Windows.Forms.Cursors.WaitCursor
		' Get list of directories where Tomes could be
		ScrollList = New Collection
		TomeSavePos = 1
		'UPGRADE_WARNING: Dir has a new behavior. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="9B7D5ADD-D8FE-4819-A36C-6DEDAF088CC7"'
		FileName = Dir(gAppPath & "\world\" & WorldNow.Name & "\save*", FileAttribute.Directory)
		Do Until FileName = ""
			If FileName <> "." And FileName <> ".." Then
				TomeX = New Tome
				TomeX.FullPath = gAppPath & "\world\" & WorldNow.Name & "\" & FileName
				' Increment the available slot (if can)
				If IsNumeric(Mid(FileName, 5)) = True Then
					If CShort(Mid(FileName, 5)) >= TomeSavePos Then
						TomeSavePos = CShort(Mid(FileName, 5)) + 1
					End If
				End If
				ScrollList.Add(TomeX)
			End If
			'UPGRADE_WARNING: Dir has a new behavior. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="9B7D5ADD-D8FE-4819-A36C-6DEDAF088CC7"'
			FileName = Dir()
		Loop 
		' Now go find Tomes in those directories
		For	Each TomeX In ScrollList
			'UPGRADE_WARNING: Dir has a new behavior. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="9B7D5ADD-D8FE-4819-A36C-6DEDAF088CC7"'
			FileName = Dir(TomeX.FullPath & "\*.tom")
			If FileName <> "" Then
				tmpFullPath = TomeX.FullPath
				c = FreeFile
				FileOpen(c, TomeX.FullPath & "\" & FileName, OpenMode.Binary)
				TomeX.ReadFromFileHeader(c)
				FileClose(c)
				TomeX.FileName = tmpFullPath & "\" & FileName
			End If
		Next TomeX
		' Set default save name
		'UPGRADE_WARNING: Couldn't resolve default property of object Tome.Name. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
		CreateNameNew = Tome.Name & " Save " & TomeSavePos
		ScrollTop = 1
		'UPGRADE_WARNING: Screen property Screen.MousePointer has a new behavior. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6BA9B8D2-2A32-4B6E-8D36-44949974A5B4"'
		System.Windows.Forms.Cursor.Current = System.Windows.Forms.Cursors.Default
	End Sub
	
	Private Sub TomeSavesDelete()
		Dim sMsg As String
		DialogSetUp(modGameGeneral.DLGTYPE.bdDlgNoReply)
		DialogSetButton(1, "No")
		DialogSetButton(2, "Yes")
		'UPGRADE_WARNING: Couldn't resolve default property of object ScrollList().Name. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
		DialogShow("DM", "Do you really wish to delete " & ScrollList.Item(ScrollSelect).Name & "?")
		picConvo.Visible = False
		' If Yes then Remove the Directory and Relist Saves
		If ConvoAction = 0 Then
			'UPGRADE_WARNING: Couldn't resolve default property of object ScrollList().FileName. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
			If oFileSys.Delete(ClipPath(ScrollList.Item(ScrollSelect).FileName), clsInOut.IOActionType.Folder, True) Then
				sMsg = "Save deleted."
			Else
				'UPGRADE_WARNING: Couldn't resolve default property of object ScrollList().FileName. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
				sMsg = "Failed to remove folder: " & ClipPath(ScrollList.Item(ScrollSelect).FileName)
			End If
			ScrollList.Remove(ScrollSelect)
			DialogSetUp(modGameGeneral.DLGTYPE.bdDlgNoReply)
			DialogShow("DM", sMsg)
			picConvo.Visible = False
			TomeSavesList(Least(ScrollSelect, ScrollList.Count()))
		End If
	End Sub
	
	Private Function LockedInTome(ByRef strName As String, Optional ByRef blnRelease As Boolean = False) As String
		Dim CreatureX As Creature
		Dim TomeX, TomeZ As Tome
		Dim i, c As Short
		Dim FileName As String
		Dim DirList As Collection
		Dim strTome As Object
		DirList = New Collection
		LockedInTome = "deleted savegame!"
		' create list of saves
		i = 0
		'UPGRADE_WARNING: Dir has a new behavior. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="9B7D5ADD-D8FE-4819-A36C-6DEDAF088CC7"'
		FileName = Dir(gAppPath & "\World\" & WorldNow.Name & "\*.*", FileAttribute.Directory)
		Do Until FileName = ""
			If FileName <> "." And FileName <> ".." Then
				If (GetAttr(gAppPath & "\World\" & WorldNow.Name & "\" & FileName) And FileAttribute.Directory) = FileAttribute.Directory Then
					DirList.Add(gAppPath & "\World\" & WorldNow.Name & "\" & FileName)
				End If
			End If
			'UPGRADE_WARNING: Dir has a new behavior. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="9B7D5ADD-D8FE-4819-A36C-6DEDAF088CC7"'
			FileName = Dir()
		Loop 
		For	Each strTome In DirList
			'UPGRADE_WARNING: Couldn't resolve default property of object strTome. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
			'UPGRADE_WARNING: Dir has a new behavior. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="9B7D5ADD-D8FE-4819-A36C-6DEDAF088CC7"'
			FileName = Dir(strTome & "\*.tom")
			Do Until FileName = ""
				' Read the Tome file
				c = FreeFile
				TomeX = New Tome
				'UPGRADE_WARNING: Couldn't resolve default property of object strTome. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
				FileOpen(c, strTome & "\" & FileName, OpenMode.Binary)
				TomeX.ReadFromFile(c)
				FileClose(c)
				' get the party members
				For	Each CreatureX In TomeX.Creatures
					If CreatureX.Name = strName Then
						i = i + 1 ' count how many saves of the same tome exist!
						LockedInTome = TomeX.Name
						If blnRelease Then
							ReleaseCharacter(CreatureX)
							' also, remove it from the tome!
							TomeX.RemoveCreature("X" & CreatureX.Index)
							c = FreeFile
							'UPGRADE_WARNING: Couldn't resolve default property of object strTome. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
							FileOpen(c, strTome & "\" & FileName, OpenMode.Binary)
							TomeX.SaveToFile(c)
							FileClose(c)
							i = i - 1
							' same for the \Worlds working version
							If i < 0 Then ' all saved games processed, we can now take care of the working version in \Worlds
								TomeZ = New Tome
								FileOpen(c, gAppPath & "\World\" & WorldNow.Name & "\" & FileName, OpenMode.Binary)
								TomeZ.ReadFromFile(c)
								FileClose(c)
								TomeZ.RemoveCreature("X" & CreatureX.Index)
								c = FreeFile
								FileOpen(c, gAppPath & "\World\" & WorldNow.Name & "\" & FileName, OpenMode.Binary)
								TomeZ.SaveToFile(c)
								FileClose(c)
							End If
						End If
					End If
				Next CreatureX
				'UPGRADE_WARNING: Dir has a new behavior. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="9B7D5ADD-D8FE-4819-A36C-6DEDAF088CC7"'
				FileName = Dir()
			Loop 
		Next strTome
	End Function
	
	Private Sub ReleaseCharacter(ByRef CreatureX As Creature)
		Dim FileName As String
		Dim c As Short
		CreatureX.OnAdventure = False
		'UPGRADE_WARNING: Dir has a new behavior. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="9B7D5ADD-D8FE-4819-A36C-6DEDAF088CC7"'
		FileName = Dir(gAppPath & "\roster\" & WorldNow.Name & "\" & CreatureX.Name & ".rsc")
		If FileName <> "" Then
			FileName = gAppPath & "\roster\" & WorldNow.Name & "\" & CreatureX.Name & ".rsc"
			Kill(FileName)
		Else
			FileName = gAppPath & "\roster\" & WorldNow.Name & "\" & CreatureX.Name & ".rsc"
		End If
		c = FreeFile
		FileOpen(c, FileName, OpenMode.Binary)
		CreatureX.SaveToFile(c)
		FileClose(c)
	End Sub
	
	'UPGRADE_NOTE: Tome was upgraded to Tome_Renamed. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="A9E4979A-37FA-4718-9994-97DD76ED70A7"'
	Private Sub TomeSavesList(ByRef Index As Short)
		Dim Tome_Renamed As Object
		Dim fp, c, n, i As Short
		Dim rc As Integer
		Dim SaveName As String
		' Set up picTmp to hold screen shot
		'UPGRADE_ISSUE: PictureBox method picTmp.Cls was not upgraded. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="CC4C7EC0-C903-48FC-ACCC-81861D12DA4A"'
		picTmp.Cls() : picTmp.Width = 220 : picTmp.Height = 165
		'UPGRADE_ISSUE: PictureBox method picTomeNew.Cls was not upgraded. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="CC4C7EC0-C903-48FC-ACCC-81861D12DA4A"'
		picTomeNew.Cls()
		' Set header correctly
		If TomeAction = bdTomeSaves Then
			ShowText(picTomeNew, 30, 12, 519, 14, bdFontElixirWhite, "Choose a Saved Tome to Play", True, False)
			ShowText(picTomeNew, 19, 45, 273, 14, bdFontElixirWhite, "Available Saves", True, False)
			picCreateName.Visible = False
		Else
			'UPGRADE_WARNING: Couldn't resolve default property of object Tome.Name. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
			ShowText(picTomeNew, 30, 12, 519, 14, bdFontElixirWhite, "Save Game " & Tome.Name, True, False)
			ShowText(picTomeNew, 19, 45, 273, 14, bdFontElixirWhite, "Current Saves", True, False)
			picCreateName.Visible = True
			If Index = 0 Then
				' Display an image of the Map
				'UPGRADE_ISSUE: PictureBox property picTomeNew.hdc was not upgraded. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="CC4C7EC0-C903-48FC-ACCC-81861D12DA4A"'
				rc = SetStretchBltMode(picTomeNew.hdc, 3)
				'UPGRADE_ISSUE: PictureBox property picMap.hdc was not upgraded. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="CC4C7EC0-C903-48FC-ACCC-81861D12DA4A"'
				'UPGRADE_ISSUE: PictureBox property picTomeNew.hdc was not upgraded. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="CC4C7EC0-C903-48FC-ACCC-81861D12DA4A"'
				rc = StretchBlt(picTomeNew.hdc, 327, 83, 220, 165, picMap.hdc, 0, 0, picBox.Width, picBox.Height, SRCCOPY)
				ShowText(picTomeNew, 324, 260, 230, 18, bdFontNoxiousBlack, VB6.Format(Now, "ddd mmm dd yyyy hh:mm AM/PM"), True, False)
			End If
		End If
		' List saved Tomes
		n = 0 : ScrollSelect = 0
		For c = ScrollTop To Least(ScrollTop + 11, ScrollList.Count())
			' Load Save Name
			fp = FreeFile
			'UPGRADE_WARNING: Couldn't resolve default property of object ScrollList().FileName. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
			FileOpen(fp, ClipPath(ScrollList.Item(c).FileName) & "\name.dat", OpenMode.Binary)
			'UPGRADE_WARNING: Get was upgraded to FileGet and has a new behavior. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="9B7D5ADD-D8FE-4819-A36C-6DEDAF088CC7"'
			FileGet(fp, rc)
			SaveName = ""
			For i = 1 To rc
				SaveName = SaveName & " "
			Next i
			'UPGRADE_WARNING: Get was upgraded to FileGet and has a new behavior. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="9B7D5ADD-D8FE-4819-A36C-6DEDAF088CC7"'
			FileGet(fp, SaveName)
			FileClose(fp)
			If c = Index Then
				If TomeAction = bdTomeSaves Then
					'UPGRADE_WARNING: Couldn't resolve default property of object ScrollList().Name. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
					ShowText(picTomeNew, 332, 48, 217, 15, bdFontElixirBlack, ScrollList.Item(c).Name, True, False)
				Else
					CreateNameNew = SaveName
					'UPGRADE_WARNING: Couldn't resolve default property of object ScrollList().FileName. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
					TomeSavePathName = ClipPath(ScrollList.Item(c).FileName)
				End If
				ShowText(picTomeNew, 24, 68 + n * 22, 245, 15, bdFontElixirGold, SaveName, True, False)
				' Show Screen Shot
				'UPGRADE_WARNING: Couldn't resolve default property of object ScrollList().FileName. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
				picTmp.Image = System.Drawing.Image.FromFile(ClipPath(ScrollList.Item(c).FileName) & "\screen.bmp")
				'UPGRADE_ISSUE: PictureBox property picTmp.hdc was not upgraded. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="CC4C7EC0-C903-48FC-ACCC-81861D12DA4A"'
				'UPGRADE_ISSUE: PictureBox property picTomeNew.hdc was not upgraded. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="CC4C7EC0-C903-48FC-ACCC-81861D12DA4A"'
				rc = BitBlt(picTomeNew.hdc, 327, 83, 220, 165, picTmp.hdc, 0, 0, SRCCOPY)
				' Show Date/Time Saved
				'UPGRADE_WARNING: Couldn't resolve default property of object ScrollList().FileName. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
				ShowText(picTomeNew, 324, 260, 230, 18, bdFontNoxiousBlack, VB6.Format(FileDateTime(ScrollList.Item(c).FileName), "ddd mmm dd yyyy hh:mm AM/PM"), True, False)
				ScrollSelect = Index
			Else
				ShowText(picTomeNew, 24, 68 + n * 22, 245, 15, bdFontElixirWhite, SaveName, True, False)
			End If
			n = n + 1
		Next c
		' Set up buttons
		'UPGRADE_ISSUE: picTomeNew was upgraded to a Panel, and cannot be coerced to a PictureBox. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="0FF1188E-27C0-4FED-842D-159C65894C9B"'
		ScrollBarShow(picTomeNew, 276, 62, 267, ScrollTop, ScrollList.Count() - 11, 0)
		If TomeAction = bdTomeSaves Then
			ShowButton(picTomeNew, 112, 331, "Delete", False)
			ShowButton(picTomeNew, 380, 364, "Cancel", False)
			ShowButton(picTomeNew, 476, 364, "Play", False)
		Else
			ShowButton(picTomeNew, 112, 331, "Delete", False)
			ShowButton(picTomeNew, 284, 364, "Cancel", False)
			ShowButton(picTomeNew, 380, 364, "Replace", False)
			ShowButton(picTomeNew, 476, 364, "Save", False)
		End If
		' Refresh the screen
		picTomeNew.Top = (Me.ClientRectangle.Height - picTomeNew.ClientRectangle.Height) / 2
		picTomeNew.Left = (Me.ClientRectangle.Width - picTomeNew.ClientRectangle.Width) / 2
		picTomeNew.Refresh()
		picTomeNew.BringToFront()
		picTomeNew.Visible = True
		If TomeAction = bdTomeSaveAs Then
			CreatePCNameEnter(0)
		End If
	End Sub
	
	Private Sub TomeNewList(ByRef Index As Short)
		Dim n, c, i As Short
		Dim rc As Integer
		Dim TomeX As Tome
		Dim FileName As String
		TomeAction = bdTomeNew
		'UPGRADE_ISSUE: PictureBox method picTomeNew.Cls was not upgraded. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="CC4C7EC0-C903-48FC-ACCC-81861D12DA4A"'
		picTomeNew.Cls()
		TomeIndex = Index
		ShowText(picTomeNew, 30, 12, 519, 14, bdFontElixirWhite, "Choose a Tome to Play", True, False)
		ShowText(picTomeNew, 19, 45, 273, 14, bdFontElixirWhite, "Available Tomes", True, False)
		n = 0
		For c = 1 To TomeList.Count()
			If IsBetween(c, TomeListTop, Least(TomeListTop + 11, TomeList.Count())) Then
				If c = Index Then
					'UPGRADE_WARNING: Couldn't resolve default property of object TomeList().Name. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
					ShowText(picTomeNew, 24, 68 + n * 22, 245, 22, bdFontElixirGold, TomeList.Item(c).Name, True, False)
				Else
					'UPGRADE_WARNING: Couldn't resolve default property of object TomeList().Name. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
					ShowText(picTomeNew, 24, 68 + n * 22, 245, 22, bdFontElixirWhite, TomeList.Item(c).Name, True, False)
				End If
				n = n + 1
			End If
			If c = Index Then
				'UPGRADE_WARNING: Couldn't resolve default property of object TomeList().Name. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
				ShowText(picTomeNew, 332, 48, 217, 15, bdFontElixirBlack, TomeList.Item(c).Name, True, False)
				If c = 1 Then
					' Special stuff for Tome Wizard
					ShowText(picTomeNew, 328, 80, 215, 14, bdFontElixirBlack, "Party Size", True, False)
					ShowText(picTomeNew, 328, 98, 215, 14, bdFontNoxiousBlack, "Solo     2-3     4-6     7-9", True, False)
					For i = 0 To 3
						If TomeWizardParty = i Then
							'UPGRADE_ISSUE: PictureBox property picMisc.hdc was not upgraded. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="CC4C7EC0-C903-48FC-ACCC-81861D12DA4A"'
							'UPGRADE_ISSUE: PictureBox property picTomeNew.hdc was not upgraded. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="CC4C7EC0-C903-48FC-ACCC-81861D12DA4A"'
							rc = BitBlt(picTomeNew.hdc, 347 + i * 55, 114, 18, 18, picMisc.hdc, 18, 18, SRCCOPY)
						Else
							'UPGRADE_ISSUE: PictureBox property picMisc.hdc was not upgraded. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="CC4C7EC0-C903-48FC-ACCC-81861D12DA4A"'
							'UPGRADE_ISSUE: PictureBox property picTomeNew.hdc was not upgraded. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="CC4C7EC0-C903-48FC-ACCC-81861D12DA4A"'
							rc = BitBlt(picTomeNew.hdc, 347 + i * 55, 114, 18, 18, picMisc.hdc, 0, 18, SRCCOPY)
						End If
					Next i
					ShowText(picTomeNew, 328, 146, 215, 14, bdFontElixirBlack, "Experience Level", True, False)
					ShowText(picTomeNew, 332, 164, 215, 14, bdFontNoxiousBlack, "1-3     4-6     7-10     10+", True, False)
					For i = 0 To 3
						If TomeWizardLevel = i Then
							'UPGRADE_ISSUE: PictureBox property picMisc.hdc was not upgraded. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="CC4C7EC0-C903-48FC-ACCC-81861D12DA4A"'
							'UPGRADE_ISSUE: PictureBox property picTomeNew.hdc was not upgraded. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="CC4C7EC0-C903-48FC-ACCC-81861D12DA4A"'
							rc = BitBlt(picTomeNew.hdc, 347 + i * 55, 180, 18, 18, picMisc.hdc, 18, 18, SRCCOPY)
						Else
							'UPGRADE_ISSUE: PictureBox property picMisc.hdc was not upgraded. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="CC4C7EC0-C903-48FC-ACCC-81861D12DA4A"'
							'UPGRADE_ISSUE: PictureBox property picTomeNew.hdc was not upgraded. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="CC4C7EC0-C903-48FC-ACCC-81861D12DA4A"'
							rc = BitBlt(picTomeNew.hdc, 347 + i * 55, 180, 18, 18, picMisc.hdc, 0, 18, SRCCOPY)
						End If
					Next i
					' Show Story Listing
					ShowText(picTomeNew, 328, 212, 215, 14, bdFontElixirBlack, "Story", True, False)
					'UPGRADE_ISSUE: PictureBox property picMisc.hdc was not upgraded. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="CC4C7EC0-C903-48FC-ACCC-81861D12DA4A"'
					'UPGRADE_ISSUE: PictureBox property picTomeNew.hdc was not upgraded. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="CC4C7EC0-C903-48FC-ACCC-81861D12DA4A"'
					rc = BitBlt(picTomeNew.hdc, 328, 230, 86, 18, picMisc.hdc, 0, 0, SRCCOPY)
					'UPGRADE_ISSUE: PictureBox property picMisc.hdc was not upgraded. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="CC4C7EC0-C903-48FC-ACCC-81861D12DA4A"'
					'UPGRADE_ISSUE: PictureBox property picTomeNew.hdc was not upgraded. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="CC4C7EC0-C903-48FC-ACCC-81861D12DA4A"'
					rc = BitBlt(picTomeNew.hdc, 414, 230, 82, 18, picMisc.hdc, 4, 0, SRCCOPY)
					'UPGRADE_ISSUE: PictureBox property picMisc.hdc was not upgraded. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="CC4C7EC0-C903-48FC-ACCC-81861D12DA4A"'
					'UPGRADE_ISSUE: PictureBox property picTomeNew.hdc was not upgraded. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="CC4C7EC0-C903-48FC-ACCC-81861D12DA4A"'
					rc = BitBlt(picTomeNew.hdc, 496, 230, 50, 18, picMisc.hdc, 40, 0, SRCCOPY)
					If UberWizMaps.MainMapIndex < 1 Then
						ShowText(picTomeNew, 328, 233, 220, 14, bdFontElixirWhite, "Random", True, False)
					Else
						'UPGRADE_WARNING: Couldn't resolve default property of object UberWizMaps.MapSketchs().MapName. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
						ShowText(picTomeNew, 328, 233, 220, 14, bdFontElixirWhite, UberWizMaps.MapSketchs.Item(UberWizMaps.MainMapIndex).MapName, True, False)
					End If
					' Show Size
					ShowText(picTomeNew, 328, 260, 215, 14, bdFontElixirBlack, "Tome Size", True, False)
					ShowText(picTomeNew, 324, 278, 220, 14, bdFontNoxiousBlack, "Small   Medium   Large   Huge", True, False)
					For i = 0 To 3
						If TomeWizardSize = i Then
							'UPGRADE_ISSUE: PictureBox property picMisc.hdc was not upgraded. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="CC4C7EC0-C903-48FC-ACCC-81861D12DA4A"'
							'UPGRADE_ISSUE: PictureBox property picTomeNew.hdc was not upgraded. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="CC4C7EC0-C903-48FC-ACCC-81861D12DA4A"'
							rc = BitBlt(picTomeNew.hdc, 338 + i * 61, 294, 18, 18, picMisc.hdc, 18, 18, SRCCOPY)
						Else
							'UPGRADE_ISSUE: PictureBox property picMisc.hdc was not upgraded. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="CC4C7EC0-C903-48FC-ACCC-81861D12DA4A"'
							'UPGRADE_ISSUE: PictureBox property picTomeNew.hdc was not upgraded. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="CC4C7EC0-C903-48FC-ACCC-81861D12DA4A"'
							rc = BitBlt(picTomeNew.hdc, 338 + i * 61, 294, 18, 18, picMisc.hdc, 0, 18, SRCCOPY)
						End If
					Next i
					' Show Build Message
					ShowText(picTomeNew, 328, 328, 220, 14, bdFontNoxiousBlack, BuildMessage, True, False)
				Else
					' Show Tome comments (if applicable)
					'UPGRADE_WARNING: Couldn't resolve default property of object TomeList().Comments. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
					If Len(TomeList.Item(c).Comments) > 0 Then
						'UPGRADE_WARNING: Couldn't resolve default property of object TomeList().Comments. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
						ShowText(picTomeNew, 328, 84, 215, 227, bdFontNoxiousBlack, TomeList.Item(c).Comments, False, False)
					Else
						ShowText(picTomeNew, 328, 84, 215, 227, bdFontNoxiousBlack, "No comments available for this Tome.", False, False)
					End If
					' Show Reset check box (if applicable)
					'UPGRADE_WARNING: Couldn't resolve default property of object TomeList(c).OnAdventure. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
					'UPGRADE_WARNING: Couldn't resolve default property of object TomeList(c).IsInPlay. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
					If TomeList.Item(c).IsInPlay = True And TomeList.Item(c).OnAdventure = True Then
						ShowText(picTomeNew, 358, 328, 220, 18, bdFontNoxiousBlack, "Reset the Tome?", False, False)
						'UPGRADE_WARNING: Couldn't resolve default property of object TomeList(c).IsReset. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
						If TomeList.Item(c).IsReset = True Then
							'UPGRADE_ISSUE: PictureBox property picMisc.hdc was not upgraded. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="CC4C7EC0-C903-48FC-ACCC-81861D12DA4A"'
							'UPGRADE_ISSUE: PictureBox property picTomeNew.hdc was not upgraded. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="CC4C7EC0-C903-48FC-ACCC-81861D12DA4A"'
							rc = BitBlt(picTomeNew.hdc, 500, 325, 18, 18, picMisc.hdc, 18, 18, SRCCOPY)
						Else
							'UPGRADE_ISSUE: PictureBox property picMisc.hdc was not upgraded. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="CC4C7EC0-C903-48FC-ACCC-81861D12DA4A"'
							'UPGRADE_ISSUE: PictureBox property picTomeNew.hdc was not upgraded. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="CC4C7EC0-C903-48FC-ACCC-81861D12DA4A"'
							rc = BitBlt(picTomeNew.hdc, 500, 325, 18, 18, picMisc.hdc, 0, 18, SRCCOPY)
						End If
					End If
				End If
			End If
		Next c
		' Set up buttons
		'UPGRADE_ISSUE: picTomeNew was upgraded to a Panel, and cannot be coerced to a PictureBox. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="0FF1188E-27C0-4FED-842D-159C65894C9B"'
		ScrollBarShow(picTomeNew, 276, 62, 267, TomeListTop, TomeList.Count() - 11, 0)
		ShowButton(picTomeNew, 62, 331, "Delete", False)
		ShowButton(picTomeNew, 158, 331, "Info", False)
		ShowButton(picTomeNew, 380, 364, "Cancel", False)
		ShowButton(picTomeNew, 476, 364, "Play", False)
		' Refresh the screen
		picTomeNew.Top = (Me.ClientRectangle.Height - picTomeNew.ClientRectangle.Height) / 2
		picTomeNew.Left = (Me.ClientRectangle.Width - picTomeNew.ClientRectangle.Width) / 2
		picTomeNew.Refresh()
		picTomeNew.BringToFront()
		picTomeNew.Visible = True
	End Sub
	
	Private Function TomeGatherCheck() As Short
		Dim CreatureX As Creature
		'Checks to see if there's at least one non-DM Controlled character in the Party.
		TomeGatherCheck = False
		For	Each CreatureX In Tome.Creatures
			If CreatureX.DMControlled = False Then
				TomeGatherCheck = True
				Exit For
			End If
		Next CreatureX
	End Function
	
	'Private Sub TomeGatherList()
	'UPGRADE_NOTE: Tome was upgraded to Tome_Renamed. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="A9E4979A-37FA-4718-9994-97DD76ED70A7"'
	Public Sub TomeGatherList()
		Dim Tome_Renamed As Object
		Dim rc As Integer
		Dim i, X, c, Y, n As Short
		Dim CreatureX As Creature
		Dim TriggerX As Trigger
		'UPGRADE_WARNING: Screen property Screen.MousePointer has a new behavior. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6BA9B8D2-2A32-4B6E-8D36-44949974A5B4"'
		System.Windows.Forms.Cursor.Current = System.Windows.Forms.Cursors.WaitCursor
		TomeAction = bdTomeGather
		'UPGRADE_ISSUE: PictureBox method picTomeNew.Cls was not upgraded. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="CC4C7EC0-C903-48FC-ACCC-81861D12DA4A"'
		picTomeNew.Cls()
		ShowText(picTomeNew, 30, 12, 519, 14, bdFontElixirWhite, "Gather your Party", True, False)
		ShowText(picTomeNew, 19, 45, 273, 14, bdFontElixirWhite, "Available Characters", True, False)
		ShowText(picTomeNew, 332, 48, 217, 15, bdFontElixirBlack, "Chosen Characters", True, False)
		' Show available slots in the Tome
		'UPGRADE_WARNING: Couldn't resolve default property of object Tome.PartySize. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
		If Tome.PartySize = 0 Then
			'UPGRADE_WARNING: Couldn't resolve default property of object Tome.PartySize. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
			Tome.PartySize = 5
		End If
		'UPGRADE_WARNING: Couldn't resolve default property of object Tome.PartySize. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
		For c = 1 To Least(Tome.PartySize, 9)
			X = 331 + ((c - 1) Mod 3) * 72
			Y = 79 + Int((c - 1) / 3) * 82
			'UPGRADE_ISSUE: PictureBox property picFaces.hdc was not upgraded. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="CC4C7EC0-C903-48FC-ACCC-81861D12DA4A"'
			'UPGRADE_ISSUE: PictureBox property picTomeNew.hdc was not upgraded. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="CC4C7EC0-C903-48FC-ACCC-81861D12DA4A"'
			rc = BitBlt(picTomeNew.hdc, X, Y, 66, 76, picFaces.hdc, bdFaceScroll, 0, SRCCOPY)
		Next c
		' Show Characters in the Tome (NPCs)
		c = 0
		'UPGRADE_WARNING: Couldn't resolve default property of object Tome.Creatures. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
		For	Each CreatureX In Tome.Creatures
			LoadCreaturePic(CreatureX)
			c = c + 1
			X = 331 + ((c - 1) Mod 3) * 74
			Y = 79 + Int((c - 1) / 3) * 84
			'UPGRADE_ISSUE: PictureBox property picMisc.hdc was not upgraded. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="CC4C7EC0-C903-48FC-ACCC-81861D12DA4A"'
			'UPGRADE_ISSUE: PictureBox property picTomeNew.hdc was not upgraded. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="CC4C7EC0-C903-48FC-ACCC-81861D12DA4A"'
			rc = BitBlt(picTomeNew.hdc, X - 4, Y - 4, 74, 84, picMisc.hdc, 0, 36, SRCCOPY)
			'UPGRADE_ISSUE: PictureBox property picFaces.hdc was not upgraded. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="CC4C7EC0-C903-48FC-ACCC-81861D12DA4A"'
			'UPGRADE_ISSUE: PictureBox property picTomeNew.hdc was not upgraded. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="CC4C7EC0-C903-48FC-ACCC-81861D12DA4A"'
			rc = BitBlt(picTomeNew.hdc, X, Y, 66, 76, picFaces.hdc, bdFaceMin + CreatureX.Pic * 66, 0, SRCCOPY)
		Next CreatureX
		' List available in Roster
		n = 0 : TomeRosterSelect = Least(Greatest(TomeRosterSelect, 1), TomeRoster.Count())
		For c = 0 To Least(2, TomeRoster.Count() - TomeRosterTop)
			n = n + 1
			' Load and show picture
			'UPGRADE_WARNING: Couldn't resolve default property of object TomeRoster(). Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
			LoadCreaturePic(TomeRoster.Item(TomeRosterTop + c))
			'UPGRADE_ISSUE: PictureBox property picMisc.hdc was not upgraded. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="CC4C7EC0-C903-48FC-ACCC-81861D12DA4A"'
			'UPGRADE_ISSUE: PictureBox property picTomeNew.hdc was not upgraded. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="CC4C7EC0-C903-48FC-ACCC-81861D12DA4A"'
			rc = BitBlt(picTomeNew.hdc, 24, n * 86 - 17, 74, 84, picMisc.hdc, 0, 36, SRCCOPY)
			'UPGRADE_WARNING: Couldn't resolve default property of object TomeRoster().Pic. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
			'UPGRADE_ISSUE: PictureBox property picFaces.hdc was not upgraded. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="CC4C7EC0-C903-48FC-ACCC-81861D12DA4A"'
			'UPGRADE_ISSUE: PictureBox property picTomeNew.hdc was not upgraded. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="CC4C7EC0-C903-48FC-ACCC-81861D12DA4A"'
			rc = BitBlt(picTomeNew.hdc, 28, n * 86 - 13, 66, 76, picFaces.hdc, bdFaceMin + TomeRoster.Item(TomeRosterTop + c).Pic * 66, 0, SRCCOPY)
			' If selected, show box
			If TomeRosterTop + c = TomeRosterSelect Then
				'UPGRADE_ISSUE: PictureBox property picFaces.hdc was not upgraded. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="CC4C7EC0-C903-48FC-ACCC-81861D12DA4A"'
				'UPGRADE_ISSUE: PictureBox property picTomeNew.hdc was not upgraded. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="CC4C7EC0-C903-48FC-ACCC-81861D12DA4A"'
				rc = BitBlt(picTomeNew.hdc, 28, n * 86 - 13, 66, 76, picFaces.hdc, bdFaceSelect + 66, 0, SRCAND)
				'UPGRADE_ISSUE: PictureBox property picFaces.hdc was not upgraded. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="CC4C7EC0-C903-48FC-ACCC-81861D12DA4A"'
				'UPGRADE_ISSUE: PictureBox property picTomeNew.hdc was not upgraded. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="CC4C7EC0-C903-48FC-ACCC-81861D12DA4A"'
				rc = BitBlt(picTomeNew.hdc, 28, n * 86 - 13, 66, 76, picFaces.hdc, bdFaceSelect, 0, SRCPAINT)
			End If
			' If unavailable, grey out
			'UPGRADE_WARNING: Couldn't resolve default property of object TomeRoster(TomeRosterTop + c).OnAdventure. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
			If TomeRoster.Item(TomeRosterTop + c).OnAdventure = True Then
				'UPGRADE_ISSUE: PictureBox property picFaces.hdc was not upgraded. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="CC4C7EC0-C903-48FC-ACCC-81861D12DA4A"'
				'UPGRADE_ISSUE: PictureBox property picTomeNew.hdc was not upgraded. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="CC4C7EC0-C903-48FC-ACCC-81861D12DA4A"'
				rc = BitBlt(picTomeNew.hdc, 28, n * 86 - 13, 66, 76, picFaces.hdc, 66, 0, SRCAND)
				'UPGRADE_ISSUE: PictureBox property picFaces.hdc was not upgraded. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="CC4C7EC0-C903-48FC-ACCC-81861D12DA4A"'
				'UPGRADE_ISSUE: PictureBox property picTomeNew.hdc was not upgraded. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="CC4C7EC0-C903-48FC-ACCC-81861D12DA4A"'
				rc = BitBlt(picTomeNew.hdc, 28, n * 86 - 13, 66, 76, picFaces.hdc, 0, 0, SRCPAINT)
				'            ShowText picTomeNew, 28, n * 86 + 50, 66, 76, bdFontSmallWhite, "In Tome", True, False
				'UPGRADE_WARNING: Couldn't resolve default property of object TomeRoster().Name. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
				ShowText(picTomeNew, 22, n * 86 + 35, 76, 76, bdFontSmallWhite, "In " & LockedInTome(TomeRoster.Item(TomeRosterTop + c).Name), True, False)
			End If
			' Show name and stats
			'UPGRADE_WARNING: Couldn't resolve default property of object TomeRoster().Name. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
			ShowText(picTomeNew, 102, n * 86 - 16, 166, 14, bdFontElixirWhite, TomeRoster.Item(TomeRosterTop + c).Name, True, False)
			'UPGRADE_WARNING: Couldn't resolve default property of object TomeRoster(TomeRosterTop + c).HPNow. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
			If TomeRoster.Item(TomeRosterTop + c).HPNow < 1 Then
				'UPGRADE_WARNING: Couldn't resolve default property of object TomeRoster().Level. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
				ShowText(picTomeNew, 102, n * 86, 166, 14, bdFontNoxiousWhite, "Lvl " & TomeRoster.Item(TomeRosterTop + c).Level & " * DEAD", True, False)
			Else
				'UPGRADE_WARNING: Couldn't resolve default property of object TomeRoster(TomeRosterTop + c).HPMax. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
				'UPGRADE_WARNING: Couldn't resolve default property of object TomeRoster(TomeRosterTop + c).HPNow. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
				'UPGRADE_WARNING: Couldn't resolve default property of object TomeRoster().Level. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
				ShowText(picTomeNew, 102, n * 86, 166, 14, bdFontNoxiousWhite, "Lvl " & TomeRoster.Item(TomeRosterTop + c).Level & " * " & TomeRoster.Item(TomeRosterTop + c).HPNow & " of " & TomeRoster.Item(TomeRosterTop + c).HPMax, True, False)
			End If
			' Show 3 skills
			i = 1
			'UPGRADE_WARNING: Couldn't resolve default property of object TomeRoster(TomeRosterTop + c).Triggers. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
			For	Each TriggerX In TomeRoster.Item(TomeRosterTop + c).Triggers
				If TriggerX.IsSkill = True Then
					ShowText(picTomeNew, 102, n * 86 + i * 16, 166, 14, bdFontNoxiousWhite, (TriggerX.Name), True, False)
					i = i + 1
					If i > 3 Then
						Exit For
					End If
				End If
			Next TriggerX
			If i = 1 Then
				ShowText(picTomeNew, 102, n * 86 + 16, 166, 14, bdFontNoxiousWhite, "No Skills", True, False)
			End If
		Next c
		' Set up buttons
		'UPGRADE_ISSUE: picTomeNew was upgraded to a Panel, and cannot be coerced to a PictureBox. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="0FF1188E-27C0-4FED-842D-159C65894C9B"'
		ScrollBarShow(picTomeNew, 276, 62, 267, TomeRosterTop, TomeRoster.Count() - 2, 0)
		' [Titi 2.4.6] If character is already adventuring, delete button shouldn't show
		If TomeRoster.Count() > 0 Then
			'UPGRADE_WARNING: Couldn't resolve default property of object TomeRoster(TomeRosterSelect).OnAdventure. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
			If TomeRoster.Item(TomeRosterSelect).OnAdventure = False Then
				ShowButton(picTomeNew, 62, 331, "Delete", False)
			Else
				'            ShowButton picTomeNew, 62, 331, "Delete", False, True
				' [Titi 2.4.9] propose to release the character instead
				ShowButton(picTomeNew, 62, 331, "Release", False)
			End If
		Else
			ShowButton(picTomeNew, 62, 331, "Delete", False, True)
		End If
		ShowButton(picTomeNew, 158, 331, "Create", False)
		ShowButton(picTomeNew, 380, 364, "Cancel", False)
		ShowButton(picTomeNew, 476, 364, "Play", False)
		' Refresh the screen
		picTomeNew.Refresh()
		picTomeNew.BringToFront()
		'UPGRADE_WARNING: Screen property Screen.MousePointer has a new behavior. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6BA9B8D2-2A32-4B6E-8D36-44949974A5B4"'
		System.Windows.Forms.Cursor.Current = System.Windows.Forms.Cursors.Default
	End Sub
	
	Private Sub TomeNewBuild()
		Dim UberWiz As UberWizard
		Dim FileName As String
		Dim k, c, i, j As Short
		Dim Found As Short
		Dim MapSketchX As MapSketch
		Dim TomeX As Tome
		Dim AreaX As Area
		Dim MapX As Map
		Dim TomeZ As Tome
		Dim TomeSketchX As TomeSketch
		Dim CreatureX As Creature
		Dim TriggerX As Trigger
		Dim JournalX, FactoidX As Factoid
		' Generate an UberWiz
		'UPGRADE_WARNING: Screen property Screen.MousePointer has a new behavior. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6BA9B8D2-2A32-4B6E-8D36-44949974A5B4"'
		System.Windows.Forms.Cursor.Current = System.Windows.Forms.Cursors.WaitCursor
		BuildMessage = "Scanning Library"
		TomeNewList(1)
		UberWiz = New UberWizard
		' Set Size, Party Level and Party Size
		UberWiz.TomePartySize = TomeWizardParty
		UberWiz.TomePartyAvgLevel = TomeWizardLevel
		Select Case TomeWizardSize
			Case 0 ' Small / 8,000 Sq
				UberWiz.TotalSize = 10000 - Int(Rnd() * 4000)
			Case 1 ' Medium / 20,000 Sq
				UberWiz.TotalSize = 25000 - Int(Rnd() * 10000)
			Case 2 ' Large / 40,000 Sq
				UberWiz.TotalSize = 45000 - Int(Rnd() * 10000)
			Case 3 ' Huge / 60,000+
				UberWiz.TotalSize = 70000 - Int(Rnd() * 20000)
		End Select
		' Finish off UberWiz parts and pieces
		If UberWizMaps.MainMapIndex > 0 Then
			modDungeonMaker.InitUberWizMaps(UberWiz)
			modDungeonMaker.InitUberWizThemes(UberWiz)
			modDungeonMaker.InitUberWizCreatures(UberWiz)
			modDungeonMaker.InitUberWizItems(UberWiz)
			'UPGRADE_WARNING: Couldn't resolve default property of object UberWizMaps.MapSketchs().Index. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
			modDungeonMaker.UberWizMainMap(UberWiz, UberWizMaps.MapSketchs.Item(UberWizMaps.MainMapIndex).Index)
			modDungeonMaker.UberWizFinish(UberWiz, 2, False)
		Else
			modDungeonMaker.UberWizFinish(UberWiz, 0, False)
		End If
		' If no maps, then cannot continue
		If UberWiz.MapSketchs.Count() < 1 Then
			DialogDM("There are no maps available in the library.")
			BuildMessage = "No Library Available"
			TomeNewList(1)
			'UPGRADE_WARNING: Screen property Screen.MousePointer has a new behavior. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6BA9B8D2-2A32-4B6E-8D36-44949974A5B4"'
			System.Windows.Forms.Cursor.Current = System.Windows.Forms.Cursors.Default
			Exit Sub
		End If
		'UPGRADE_WARNING: Screen property Screen.MousePointer has a new behavior. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6BA9B8D2-2A32-4B6E-8D36-44949974A5B4"'
		System.Windows.Forms.Cursor.Current = System.Windows.Forms.Cursors.WaitCursor
		' Configure new Tome
		TomeX = New Tome
		TomeX.Name = UberWiz.TomeName
		TomeX.Comments = UberWiz.TomeComments
		TomeX.AreaIndex = UberWiz.AreaIndex
		TomeX.MapIndex = UberWiz.MapIndex
		TomeX.EntryIndex = UberWiz.EntryIndex
		TomeX.PartySize = 9
		' Copy Tome Creatures, Triggers, Journals and Factoids (if have them)
		If UberWiz.MainMap.TomeIndex > 0 Then
			TomeSketchX = UberWiz.TomeSketchs.Item("T" & UberWiz.MainMap.TomeIndex)
			For	Each CreatureX In TomeSketchX.Creatures
				TomeX.AddCreature.Copy(CreatureX)
			Next CreatureX
			For	Each TriggerX In TomeSketchX.Triggers
				TomeX.AddTrigger.Copy(TriggerX)
			Next TriggerX
			For	Each JournalX In TomeSketchX.Journals
				'UPGRADE_WARNING: Couldn't resolve default property of object JournalX. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
				TomeX.AddJournal.Copy(JournalX)
			Next JournalX
			For	Each FactoidX In TomeSketchX.Factoids
				TomeX.AddFactoid.Copy(FactoidX)
			Next FactoidX
		End If
		' Set up real average level based on selected average level
		Select Case UberWiz.TomePartyAvgLevel
			Case 0
				TomeX.PartyAvgLevel = 2
			Case 1
				TomeX.PartyAvgLevel = 5
			Case 2
				TomeX.PartyAvgLevel = 9
			Case 3
				TomeX.PartyAvgLevel = 12
		End Select
		UberWiz.MainMap.PartySize = UberWiz.TomePartySize
		UberWiz.MainMap.PartyAvgLevel = UberWiz.TomePartyAvgLevel
		' Make Maps
		modDungeonMaker.UberWizConnectMap(UberWiz, (UberWiz.MainMap), (UberWiz.TotalSize), 0, 0)
		modDungeonMaker.UberWizScatterQuests(UberWiz)
		AreaX = TomeX.AddArea
		For	Each MapSketchX In UberWiz.MapSketchs
			If MapSketchX.IsUsed = True Then
				BuildMessage = "Making " & MapSketchX.MapName
				TomeNewList(1)
				modDungeonMaker.UberWizMakeMap(UberWiz, MapSketchX, AreaX)
			End If
		Next MapSketchX
		' Fill Maps
		For	Each MapX In AreaX.Plot.Maps
			BuildMessage = "Filling " & MapX.Name
			TomeNewList(1)
			modDungeonMaker.FillMap(MapX)
		Next MapX
		' Save to a Folder
		BuildMessage = "Saving Tome"
		TomeNewList(1)
		FileName = "TomeWizard" & VB6.Format(Today, "YYMMDD") & VB6.Format(TimeOfDay, "HHMMSS")
		'UPGRADE_WARNING: Couldn't resolve default property of object TomeX.Areas().FileName. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
		TomeX.Areas.Item(1).FileName = FileName & "1.rsa"
		'TomeX.FullPath = gAppPath & "\tomes"
		TomeX.FullPath = gAppPath & "\tomes\" & WorldNow.Name & "\"
		TomeX.FileName = FileName & ".tom"
		c = FreeFile
		'Open gAppPath & "\tomes\" & FileName & ".tom" For Binary As c
		FileOpen(c, gAppPath & "\tomes\" & WorldNow.Name & "\" & FileName & ".tom", OpenMode.Binary)
		TomeX.SaveToFile(c)
		FileClose(c)
		c = FreeFile
		'Open gAppPath & "\tomes\" & FileName & "1.rsa" For Binary As c
		FileOpen(c, gAppPath & "\tomes\" & WorldNow.Name & "\" & FileName & "1.rsa", OpenMode.Binary)
		AreaX.Plot.SaveToFile(c)
		FileClose(c)
		' Add Tome to the List of Tomes
		i = 1
		For c = 1 To TomeList.Count()
			'UPGRADE_WARNING: Couldn't resolve default property of object TomeList(c).Index. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
			If TomeList.Item(c).Index >= i Then
				'UPGRADE_WARNING: Couldn't resolve default property of object TomeList().Index. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
				i = TomeList.Item(c).Index + 1
			End If
		Next c
		TomeX.Index = i
		TomeList.Add(TomeX, "T" & TomeX.Index)
		TomeIndex = TomeList.Count()
		BuildMessage = ""
		'UPGRADE_WARNING: Screen property Screen.MousePointer has a new behavior. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6BA9B8D2-2A32-4B6E-8D36-44949974A5B4"'
		System.Windows.Forms.Cursor.Current = System.Windows.Forms.Cursors.Default
	End Sub
	
	Private Sub TomeSavesClick(ByRef AtX As Short, ByRef AtY As Short, ByRef ButtonDown As Short)
		Dim c As Short
		Dim rc As Integer
		Dim CreatureX As Creature
		If PointIn(AtX, AtY, 380, 364, 90, 18) Then
			' Cancel/Replace
			If ButtonDown Then
				Select Case TomeAction
					Case bdTomeSaves
						ShowButton(picTomeNew, 380, 364, "Cancel", True)
					Case bdTomeSaveAs
						ShowButton(picTomeNew, 380, 364, "Replace", True)
				End Select
				picTomeNew.Refresh()
			Else
				Select Case TomeAction
					Case bdTomeSaves
						ShowButton(picTomeNew, 380, 364, "Cancel", False)
						picTomeNew.Refresh()
						Call PlayClickSnd(modIOFunc.ClickType.ifClick)
						picCreateName.Visible = False
						picTomeNew.Visible = False
						If picMap.Visible = True Then
							picMap.Focus()
						End If
					Case bdTomeSaveAs
						ShowButton(picTomeNew, 380, 364, "Replace", False)
						picTomeNew.Refresh()
						Call PlayClickSnd(modIOFunc.ClickType.ifClick)
						If ScrollSelect > 0 Then
							TomeSave(False)
							picCreateName.Visible = False
							picTomeNew.Visible = False
						End If
				End Select
			End If
		ElseIf PointIn(AtX, AtY, 476, 364, 90, 18) Then 
			' Play
			If ButtonDown Then
				Select Case TomeAction
					Case bdTomeSaves
						ShowButton(picTomeNew, 476, 364, "Play", True)
					Case bdTomeSaveAs
						ShowButton(picTomeNew, 476, 364, "Save", True)
				End Select
				picTomeNew.Refresh()
			Else
				Select Case TomeAction
					Case bdTomeSaves
						ShowButton(picTomeNew, 476, 364, "Play", False)
						picTomeNew.Refresh()
						Call PlayClickSnd(modIOFunc.ClickType.ifClick)
						If ScrollSelect > 0 Then
							picTomeNew.Visible = False
							picMainMenu.Visible = False
							'UPGRADE_WARNING: Couldn't resolve default property of object ScrollList().FileName. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
							TomeRestore(ScrollList.Item(ScrollSelect).FileName)
						End If
					Case bdTomeSaveAs
						ShowButton(picTomeNew, 476, 364, "Save", False)
						picTomeNew.Refresh()
						Call PlayClickSnd(modIOFunc.ClickType.ifClick)
						TomeSave(True)
						picCreateName.Visible = False
						picTomeNew.Visible = False
				End Select
			End If
		ElseIf PointIn(AtX, AtY, 284, 364, 90, 18) Then 
			' Cancel/Replace
			If ButtonDown Then
				Select Case TomeAction
					Case bdTomeSaves
					Case bdTomeSaveAs
						ShowButton(picTomeNew, 284, 364, "Cancel", True)
				End Select
				picTomeNew.Refresh()
			Else
				Select Case TomeAction
					Case bdTomeSaves
					Case bdTomeSaveAs
						Call PlayClickSnd(modIOFunc.ClickType.ifClick)
						ShowButton(picTomeNew, 284, 364, "Cancel", False)
						picTomeNew.Refresh()
						picCreateName.Visible = False
						picTomeNew.Visible = False
						If picMap.Visible = True Then
							picMap.Focus()
						End If
				End Select
			End If
		ElseIf PointIn(AtX, AtY, 276, 62, 18, 267) Then 
			'UPGRADE_ISSUE: picTomeNew was upgraded to a Panel, and cannot be coerced to a PictureBox. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="0FF1188E-27C0-4FED-842D-159C65894C9B"'
			If ScrollBarClick(AtX, AtY, ButtonDown, picTomeNew, 276, 62, 267, ScrollTop, ScrollList.Count(), 11) = True Then
				TomeSavesList(TomeIndex)
			End If
		ElseIf PointIn(AtX, AtY, 112, 331, 90, 18) And ScrollSelect > 0 Then 
			' Delete Save
			If ButtonDown Then
				ShowButton(picTomeNew, 112, 331, "Delete", True)
				picTomeNew.Refresh()
			Else
				ShowButton(picTomeNew, 112, 331, "Delete", False)
				picTomeNew.Refresh()
				Call PlayClickSnd(modIOFunc.ClickType.ifClick)
				TomeSavesDelete()
			End If
		ElseIf PointIn(AtX, AtY, 24, 68, 245, 266) And ButtonDown = False Then 
			' Saves Listed
			c = CShort((AtY - 76) / 22) + ScrollTop
			If IsBetween(c, 1, ScrollList.Count()) Then
				TomeSavesList(c)
				Call PlayClickSnd(modIOFunc.ClickType.ifClick)
			End If
		End If
	End Sub
	
	'UPGRADE_NOTE: Tome was upgraded to Tome_Renamed. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="A9E4979A-37FA-4718-9994-97DD76ED70A7"'
	Private Sub TomeNewClick(ByRef AtX As Short, ByRef AtY As Short, ByRef ButtonDown As Short, Optional ByRef btnRight As Boolean = False)
		Dim Tome_Renamed As Object
		Dim c As Short
		Dim rc As Integer
		Dim CreatureX As Creature
		Dim TomeName As String
		Dim strInTome As String
		If TomeRosterDrag > 0 And ButtonDown = False Then
			' Dropping into the Tome
			If PointIn(AtX, AtY, 331, 79, 210, 240) Then
				c = Int((AtX - 331) / 72) + Int((AtY - 79) / 82) * 3 + 1
				'UPGRADE_WARNING: Couldn't resolve default property of object Tome.Creatures. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
				If c <= Tome.Creatures.Count Then
					' Replacing a Creature (check if required in Tome)
					'UPGRADE_WARNING: Couldn't resolve default property of object Tome.Creatures. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
					If Tome.Creatures(c).RequiredInTome = False Then
						TomeRoster.Remove(TomeRosterDrag)
						CreatureX = New Creature
						'UPGRADE_WARNING: Couldn't resolve default property of object Tome.Creatures. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
						CreatureX.Copy(Tome.Creatures(c))
						TomeAddRoster(CreatureX)
						'UPGRADE_WARNING: Couldn't resolve default property of object Tome.Creatures. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
						Tome.Creatures(c).Copy(CreatureNow)
						If TomeRosterTop + 2 > TomeRoster.Count() Then
							TomeRosterTop = Greatest(1, TomeRoster.Count() - 2)
						End If
						TomeGatherList()
						Call PlayClickSnd(modIOFunc.ClickType.ifClickDrop)
						picTomeNew.Refresh()
					End If
					'UPGRADE_WARNING: Couldn't resolve default property of object Tome.PartySize. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
				ElseIf c <= Tome.PartySize Then 
					' Adding a new character
					TomeRoster.Remove(TomeRosterDrag)
					'UPGRADE_WARNING: Couldn't resolve default property of object Tome.AppendCreature. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
					Tome.AppendCreature(CreatureNow)
					If TomeRosterTop + 2 > TomeRoster.Count() Then
						TomeRosterTop = Greatest(1, TomeRoster.Count() - 2)
					End If
					TomeGatherList()
					Call PlayClickSnd(modIOFunc.ClickType.ifClickDrop)
					picTomeNew.Refresh()
				End If
			End If
			picInvDrag.Visible = False
			TomeRosterDrag = 0
		ElseIf TomeRosterDrag < 0 And ButtonDown = False Then 
			' Dropping back into the Roster
			If PointIn(AtX, AtY, 20, 63, 254, 266) Then
				'UPGRADE_WARNING: Couldn't resolve default property of object Tome.Creatures. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
				Tome.Creatures.Remove(System.Math.Abs(TomeRosterDrag))
				TomeAddRoster(CreatureNow)
				Call PlayClickSnd(modIOFunc.ClickType.ifClickDrop)
				TomeGatherList()
				picTomeNew.Refresh()
			End If
			picInvDrag.Visible = False
			TomeRosterDrag = 0
		ElseIf PointIn(AtX, AtY, 380, 364, 90, 18) Then 
			' Cancel
			If ButtonDown Then
				ShowButton(picTomeNew, 380, 364, "Cancel", True)
				picTomeNew.Refresh()
			Else
				Call PlayClickSnd(modIOFunc.ClickType.ifClick)
				ShowButton(picTomeNew, 380, 364, "Cancel", False)
				picTomeNew.Refresh()
				' Remove characters from current tome
				If TomeAction = bdTomeGather Then
					'UPGRADE_WARNING: Couldn't resolve default property of object TomeList(TomeIndex).Creatures. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
					For	Each CreatureX In TomeList.Item(TomeIndex).Creatures
						If CreatureX.RequiredInTome = False Then
							'UPGRADE_WARNING: Couldn't resolve default property of object TomeList().RemoveCreature. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
							TomeList.Item(TomeIndex).RemoveCreature("X" & CreatureX.Index)
						End If
					Next CreatureX
					TomeNewList(1)
				Else
					picTomeNew.Visible = False
				End If
			End If
		ElseIf PointIn(AtX, AtY, 476, 364, 90, 18) Then 
			' Play
			If ButtonDown Then
				ShowButton(picTomeNew, 476, 364, "Play", True)
				picTomeNew.Refresh()
			Else
				ShowButton(picTomeNew, 476, 364, "Play", False)
				picTomeNew.Refresh()
				Call PlayClickSnd(modIOFunc.ClickType.ifClick)
				Select Case TomeAction
					Case bdTomeNew
						If TomeIndex = 1 Then
							TomeNewBuild()
							If BuildMessage = "No Library Available" Then
								Exit Sub
							End If
						End If
						' Load the Tome file and Party members
						'UPGRADE_WARNING: Screen property Screen.MousePointer has a new behavior. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6BA9B8D2-2A32-4B6E-8D36-44949974A5B4"'
						System.Windows.Forms.Cursor.Current = System.Windows.Forms.Cursors.WaitCursor
						'UPGRADE_WARNING: Couldn't resolve default property of object TomeList(). Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
						TomeLoadParty(TomeList.Item(TomeIndex))
						' If has party members then skip to the end
						'UPGRADE_WARNING: Couldn't resolve default property of object Tome.OnAdventure. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
						If Tome.OnAdventure = True Then
							picMainMenu.Visible = False
							picTomeNew.Visible = False
							'UPGRADE_WARNING: Couldn't resolve default property of object TomeList(). Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
							TomeCopy(TomeList.Item(TomeIndex))
							'UPGRADE_WARNING: Couldn't resolve default property of object Tome. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
							TomeStart(Tome)
						Else
							TomeLoadRoster()
							TomeGatherList()
						End If
						'UPGRADE_WARNING: Screen property Screen.MousePointer has a new behavior. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6BA9B8D2-2A32-4B6E-8D36-44949974A5B4"'
						System.Windows.Forms.Cursor.Current = System.Windows.Forms.Cursors.Default
					Case bdTomeGather
						If TomeGatherCheck = True Then
							picMainMenu.Visible = False
							picTomeNew.Visible = False
							'UPGRADE_WARNING: Couldn't resolve default property of object Tome. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
							TomeCopy(Tome)
							'UPGRADE_WARNING: Couldn't resolve default property of object Tome. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
							TomeStart(Tome)
						Else
							DialogSetUp(modGameGeneral.DLGTYPE.bdDlgNoReply)
							DialogSetButton(1, "Done")
							DialogShow("DM", "You must have at least one human controlled Character in the Party.")
							picConvo.Visible = False
						End If
					Case bdTomeSaves
						If ScrollSelect > 0 Then
							picTomeNew.Visible = False
							picMainMenu.Visible = False
							'UPGRADE_WARNING: Couldn't resolve default property of object ScrollList().FileName. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
							TomeRestore(ScrollList.Item(ScrollSelect).FileName)
						End If
					Case bdTomeSaveAs
						If ScrollSelect > 0 Then
						End If
				End Select
			End If
		ElseIf PointIn(AtX, AtY, 276, 62, 18, 267) Then 
			' ScrollBar Click
			Select Case TomeAction
				Case bdTomeGather
					'UPGRADE_ISSUE: picTomeNew was upgraded to a Panel, and cannot be coerced to a PictureBox. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="0FF1188E-27C0-4FED-842D-159C65894C9B"'
					If ScrollBarClick(AtX, AtY, ButtonDown, picTomeNew, 276, 62, 267, TomeRosterTop, TomeRoster.Count(), 2) = True Then
						TomeGatherList()
					End If
				Case bdTomeNew
					'UPGRADE_ISSUE: picTomeNew was upgraded to a Panel, and cannot be coerced to a PictureBox. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="0FF1188E-27C0-4FED-842D-159C65894C9B"'
					If ScrollBarClick(AtX, AtY, ButtonDown, picTomeNew, 276, 62, 267, TomeListTop, TomeList.Count(), 11) = True Then
						TomeNewList(TomeIndex)
					End If
				Case bdTomeSaves, bdTomeSaveAs
					'UPGRADE_ISSUE: picTomeNew was upgraded to a Panel, and cannot be coerced to a PictureBox. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="0FF1188E-27C0-4FED-842D-159C65894C9B"'
					If ScrollBarClick(AtX, AtY, ButtonDown, picTomeNew, 276, 62, 267, ScrollTop, ScrollList.Count(), 11) = True Then
						TomeSavesList(TomeIndex)
					End If
			End Select
		ElseIf PointIn(AtX, AtY, 62, 331, 90, 18) And TomeAction = bdTomeGather Then 
			' Delete Character
			' [Titi 2.4.6] Not allowed if character is already in tome!
			If TomeRoster.Count() > 0 Then
				'UPGRADE_WARNING: Couldn't resolve default property of object TomeRoster(TomeRosterSelect).OnAdventure. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
				If Not TomeRoster.Item(TomeRosterSelect).OnAdventure Then
					ShowButton(picTomeNew, 62, 331, "Delete", ButtonDown)
					picTomeNew.Refresh()
					If ButtonDown = False And TomeRosterSelect > 0 Then
						Call PlayClickSnd(modIOFunc.ClickType.ifClick)
						CreatePCDelete()
					End If
				Else
					' Put Character back into roster
					' [Titi 2.4.9] only if character is already in tome!
					ShowButton(picTomeNew, 62, 331, "Release", ButtonDown)
					picTomeNew.Refresh()
					If ButtonDown = False And TomeRosterSelect > 0 Then
						Call PlayClickSnd(modIOFunc.ClickType.ifClick)
						DialogSetUp(modGameGeneral.DLGTYPE.bdDlgNoReply)
						DialogSetButton(1, "No")
						DialogSetButton(2, "Yes")
						'UPGRADE_WARNING: Couldn't resolve default property of object TomeRoster().Name. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
						DialogShow("DM", "You're about to free " & TomeRoster.Item(TomeRosterSelect).Name & ".")
						DialogHide()
						If ConvoAction = 0 Then
							' confirmation to release given!
							'UPGRADE_WARNING: Couldn't resolve default property of object TomeRoster().Name. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
							strInTome = LockedInTome(TomeRoster.Item(TomeRosterSelect).Name)
							If strInTome <> "deleted savegame!" Then
								DialogSetUp(modGameGeneral.DLGTYPE.bdDlgNoReply)
								DialogSetButton(1, "Tome")
								DialogSetButton(2, "Roster")
								'UPGRADE_WARNING: Couldn't resolve default property of object TomeRoster(TomeRosterSelect).Male. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
								'UPGRADE_WARNING: Couldn't resolve default property of object TomeRoster().Name. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
								DialogShow("DM", TomeRoster.Item(TomeRosterSelect).Name & " may have gained XP and items already." & Chr(13) & "Do you want to release " & IIf(TomeRoster.Item(TomeRosterSelect).Male, "him", "her") & " as " & IIf(TomeRoster.Item(TomeRosterSelect).Male, "he", "she") & " was before entering " & strInTome & " (click " & Chr(34) & "Roster" & Chr(34) & ") or do you wish " & IIf(TomeRoster.Item(TomeRosterSelect).Male, "him", "her") & " to keep " & IIf(TomeRoster.Item(TomeRosterSelect).Male, "his", "her") & " present status (click " & Chr(34) & "Tome" & Chr(34) & ")?")
								DialogHide()
								If ConvoAction = 0 Then
									' roster precedes
									' but first, remove character from the tome
									'UPGRADE_WARNING: Couldn't resolve default property of object TomeRoster().Name. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
									strInTome = LockedInTome(TomeRoster.Item(TomeRosterSelect).Name, True)
									' now, overwrite with the roster version
									'UPGRADE_WARNING: Couldn't resolve default property of object TomeRoster(). Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
									ReleaseCharacter(TomeRoster.Item(TomeRosterSelect))
								Else
									' tome version precedes
									'UPGRADE_WARNING: Couldn't resolve default property of object TomeRoster().Name. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
									strInTome = LockedInTome(TomeRoster.Item(TomeRosterSelect).Name, True)
								End If
								DialogSetUp(modGameGeneral.DLGTYPE.bdDlgNoReply)
								'UPGRADE_WARNING: Couldn't resolve default property of object TomeRoster().Name. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
								DialogShow("DM", TomeRoster.Item(TomeRosterSelect).Name & " has now left " & strInTome & "...")
								DialogHide()
							Else
								' no more saved game, roster version only
								'UPGRADE_WARNING: Couldn't resolve default property of object TomeRoster(). Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
								ReleaseCharacter(TomeRoster.Item(TomeRosterSelect))
							End If
							DialogSetUp(modGameGeneral.DLGTYPE.bdDlgNoReply)
							'UPGRADE_WARNING: Couldn't resolve default property of object TomeRoster().Name. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
							DialogShow("DM", TomeRoster.Item(TomeRosterSelect).Name & " is now available for other adventures!")
							DialogHide()
							'UPGRADE_WARNING: Couldn't resolve default property of object TomeRoster().OnAdventure. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
							TomeRoster.Item(TomeRosterSelect).OnAdventure = False
						End If
						TomeGatherList()
					End If
				End If
			End If
		ElseIf PointIn(AtX, AtY, 158, 331, 90, 18) And TomeAction = bdTomeGather Then 
			' Create Character
			ShowButton(picTomeNew, 158, 331, "Create", ButtonDown)
			picTomeNew.Refresh()
			If ButtonDown = False Then
				Call PlayClickSnd(modIOFunc.ClickType.ifClick)
				CreatePCReturn = True
				CreatePCLoadMap(WorldNow)
				CreatePCKingdom(1)
			End If
		ElseIf PointIn(AtX, AtY, 500, 325, 18, 18) And TomeAction = bdTomeNew And TomeIndex > 1 And ButtonDown = True Then 
			' Click to reset the Tome
			'UPGRADE_WARNING: Couldn't resolve default property of object TomeList(TomeIndex).OnAdventure. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
			'UPGRADE_WARNING: Couldn't resolve default property of object TomeList(TomeIndex).IsInPlay. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
			If TomeList.Item(TomeIndex).IsInPlay = True And TomeList.Item(TomeIndex).OnAdventure = True Then
				picTomeNew.Refresh()
				Call PlayClickSnd(modIOFunc.ClickType.ifClick)
				'UPGRADE_WARNING: Couldn't resolve default property of object TomeList(TomeIndex).IsReset. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
				'UPGRADE_WARNING: Couldn't resolve default property of object TomeList().IsReset. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
				TomeList.Item(TomeIndex).IsReset = Not TomeList.Item(TomeIndex).IsReset
				TomeNewList(TomeIndex)
			End If
		ElseIf PointIn(AtX, AtY, 62, 331, 90, 18) And TomeAction = bdTomeNew And TomeIndex > 1 Then 
			' Delete Tome
			If ButtonDown Then
				ShowButton(picTomeNew, 62, 331, "Delete", True)
				picTomeNew.Refresh()
			Else
				ShowButton(picTomeNew, 62, 331, "Delete", False)
				picTomeNew.Refresh()
				Call PlayClickSnd(modIOFunc.ClickType.ifClick)
				'UPGRADE_WARNING: Couldn't resolve default property of object TomeList(). Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
				TomeDelete(TomeList.Item(TomeIndex))
			End If
		ElseIf PointIn(AtX, AtY, 158, 331, 90, 18) And TomeAction = bdTomeNew And TomeIndex > 1 Then 
			' Info Tome
			If ButtonDown Then
				ShowButton(picTomeNew, 158, 331, "Info", True)
				picTomeNew.Refresh()
			Else
				ShowButton(picTomeNew, 158, 331, "Info", False)
				picTomeNew.Refresh()
				Call PlayClickSnd(modIOFunc.ClickType.ifClick)
				'UPGRADE_WARNING: Couldn't resolve default property of object TomeList(). Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
				TomeInfo(TomeList.Item(TomeIndex))
			End If
		ElseIf PointIn(AtX, AtY, 25, 70, 72, 250) And TomeAction = bdTomeGather Then 
			' Drag from Roster
			c = Int((AtY - 70) / 86) + TomeRosterTop
			If c <= TomeRoster.Count() And c > 0 And ButtonDown = True Then
				TomeRosterSelect = c
				TomeGatherList()
				Call PlayClickSnd(modIOFunc.ClickType.ifClick)
				' [Titi 2.4.9] display character's stats and equipment
				If btnRight Then
					CreatureNow = TomeRoster.Item(c)
					frmMonsExplorerPlayer.Show()
				End If
				'UPGRADE_WARNING: Couldn't resolve default property of object TomeRoster(c).OnAdventure. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
				If TomeRoster.Item(c).OnAdventure = False Then
					CreatureNow = TomeRoster.Item(c)
					'UPGRADE_ISSUE: PictureBox method picInvDrag.Cls was not upgraded. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="CC4C7EC0-C903-48FC-ACCC-81861D12DA4A"'
					picInvDrag.Cls() : picInvDrag.Width = 72 : picInvDrag.Height = 84
					'UPGRADE_ISSUE: PictureBox property picFaces.hdc was not upgraded. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="CC4C7EC0-C903-48FC-ACCC-81861D12DA4A"'
					'UPGRADE_ISSUE: PictureBox property picInvDrag.hdc was not upgraded. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="CC4C7EC0-C903-48FC-ACCC-81861D12DA4A"'
					rc = BitBlt(picInvDrag.hdc, 3, 3, 66, 76, picFaces.hdc, bdFaceMin + CreatureNow.Pic * 66, 0, SRCCOPY)
					picInvDrag.Refresh()
					picInvDrag.BringToFront()
					TomeRosterDrag = c
					TomeRosterDragX = AtX - 25
					TomeRosterDragY = (AtY - 70) Mod 86
				End If
			End If
		ElseIf PointIn(AtX, AtY, 331, 79, 210, 240) And TomeAction = bdTomeGather And ButtonDown = True Then 
			' Drag from Tome
			c = Int((AtX - 331) / 72) + Int((AtY - 79) / 82) * 3 + 1
			'UPGRADE_WARNING: Couldn't resolve default property of object Tome.Creatures. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
			If c <= Tome.Creatures.Count Then
				'UPGRADE_WARNING: Couldn't resolve default property of object Tome.Creatures. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
				CreatureNow = Tome.Creatures(c)
				If CreatureNow.RequiredInTome = False Then
					'UPGRADE_ISSUE: PictureBox method picInvDrag.Cls was not upgraded. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="CC4C7EC0-C903-48FC-ACCC-81861D12DA4A"'
					picInvDrag.Cls() : picInvDrag.Width = 72 : picInvDrag.Height = 84
					'UPGRADE_ISSUE: PictureBox property picFaces.hdc was not upgraded. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="CC4C7EC0-C903-48FC-ACCC-81861D12DA4A"'
					'UPGRADE_ISSUE: PictureBox property picInvDrag.hdc was not upgraded. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="CC4C7EC0-C903-48FC-ACCC-81861D12DA4A"'
					rc = BitBlt(picInvDrag.hdc, 3, 3, 66, 76, picFaces.hdc, bdFaceMin + CreatureNow.Pic * 66, 0, SRCCOPY)
					picInvDrag.Refresh()
					picInvDrag.BringToFront()
					TomeRosterDrag = -c
					TomeRosterDragX = (AtX - 331) Mod 72
					TomeRosterDragY = (AtY - 79) Mod 82
				End If
			End If
		ElseIf PointIn(AtX, AtY, 326, 97, 220, 36) And TomeAction = bdTomeNew And TomeIndex = 1 And ButtonDown = True Then 
			' Click on Wizard Party
			TomeWizardParty = Int((AtX - 326) / 55)
			TomeNewList(1)
			Call PlayClickSnd(modIOFunc.ClickType.ifClick)
		ElseIf PointIn(AtX, AtY, 326, 163, 220, 36) And TomeAction = bdTomeNew And TomeIndex = 1 And ButtonDown = True Then 
			' Click on Wizard Level
			TomeWizardLevel = Int((AtX - 326) / 55)
			TomeNewList(1)
			Call PlayClickSnd(modIOFunc.ClickType.ifClick)
		ElseIf PointIn(AtX, AtY, 326, 278, 220, 36) And TomeAction = bdTomeNew And TomeIndex = 1 And ButtonDown = True Then 
			' Click on Wizard Size
			TomeWizardSize = Int((AtX - 326) / 55)
			TomeNewList(1)
			Call PlayClickSnd(modIOFunc.ClickType.ifClick)
		ElseIf PointIn(AtX, AtY, 328, 233, 220, 14) And TomeAction = bdTomeNew And TomeIndex = 1 And ButtonDown = True Then 
			' Click on Wizard Story
			If UberWizMaps.MapSketchs.Count() > 0 Then
				UberWizMaps.MainMapIndex = LoopNumber(1, UberWizMaps.MapSketchs.Count(), UberWizMaps.MainMapIndex, 1)
				'UPGRADE_WARNING: Couldn't resolve default property of object UberWizMaps.MapSketchs().TomeIndex. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
				'UPGRADE_WARNING: Couldn't resolve default property of object UberWizMaps.TomeSketchs(T & UberWizMaps.MapSketchs(UberWizMaps.MainMapIndex).TomeIndex).PartySize. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
				Select Case UberWizMaps.TomeSketchs.Item("T" & UberWizMaps.MapSketchs.Item(UberWizMaps.MainMapIndex).TomeIndex).PartySize
					Case 1
						TomeWizardParty = 0
					Case 2 To 3
						TomeWizardParty = 1
					Case 4 To 6
						TomeWizardParty = 2
					Case Else
						TomeWizardParty = 3
				End Select
				'UPGRADE_WARNING: Couldn't resolve default property of object UberWizMaps.MapSketchs().TomeIndex. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
				'UPGRADE_WARNING: Couldn't resolve default property of object UberWizMaps.TomeSketchs(T & UberWizMaps.MapSketchs(UberWizMaps.MainMapIndex).TomeIndex).PartyAvgLevel. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
				Select Case UberWizMaps.TomeSketchs.Item("T" & UberWizMaps.MapSketchs.Item(UberWizMaps.MainMapIndex).TomeIndex).PartyAvgLevel
					Case 1 To 3
						TomeWizardLevel = 0
					Case 4 To 6
						TomeWizardLevel = 1
					Case 7 To 10
						TomeWizardLevel = 2
					Case Else
						TomeWizardLevel = 3
				End Select
			Else
				TomeWizardParty = 2
				TomeWizardLevel = 0
			End If
			TomeNewList(1)
			Call PlayClickSnd(modIOFunc.ClickType.ifClick)
		ElseIf PointIn(AtX, AtY, 24, 68, 245, 266) And TomeAction = bdTomeNew And ButtonDown = False Then 
			' Item in ListBox
			c = CShort((AtY - 76) / 22) + TomeListTop
			If IsBetween(c, 1, TomeList.Count()) Then
				TomeNewList(c)
				Call PlayClickSnd(modIOFunc.ClickType.ifClick)
			End If
		ElseIf PointIn(AtX, AtY, 24, 68, 245, 266) And (TomeAction = bdTomeSaves Or TomeAction = bdTomeSaveAs) And ButtonDown = False Then 
			' Saves Listed
			c = CShort((AtY - 76) / 22) + ScrollTop
			If IsBetween(c, 1, ScrollList.Count()) Then
				TomeSavesList(c)
				Call PlayClickSnd(modIOFunc.ClickType.ifClick)
			End If
		End If
	End Sub
	
	'UPGRADE_NOTE: Text was upgraded to Text_Renamed. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="A9E4979A-37FA-4718-9994-97DD76ED70A7"'
	Public Sub ShowButton(ByRef picBox As Object, ByRef X As Short, ByRef Y As Short, ByRef Text_Renamed As String, ByRef Down As Short, Optional ByRef NotAvailable As Boolean = False)
		Dim rc As Integer
		' [Titi 2.4.6] added the greyed out display
		If NotAvailable = True Then
			'UPGRADE_ISSUE: PictureBox property picMisc.hdc was not upgraded. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="CC4C7EC0-C903-48FC-ACCC-81861D12DA4A"'
			'UPGRADE_WARNING: Couldn't resolve default property of object picBox.hdc. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
			rc = BitBlt(picBox.hdc, X, Y, 90, 18, picMisc.hdc, 0, 0, SRCCOPY)
			ShowText(picBox, X, Y + 2, 90, 18, bdFontElixirGrey, Text_Renamed, True, False)
		Else
			If Down = True Then
				'UPGRADE_ISSUE: PictureBox property picMisc.hdc was not upgraded. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="CC4C7EC0-C903-48FC-ACCC-81861D12DA4A"'
				'UPGRADE_WARNING: Couldn't resolve default property of object picBox.hdc. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
				rc = BitBlt(picBox.hdc, X, Y, 90, 18, picMisc.hdc, 90, 0, SRCCOPY)
				ShowText(picBox, X + 2, Y + 4, 90, 18, bdFontElixirWhite, Text_Renamed, True, False)
			Else
				'UPGRADE_ISSUE: PictureBox property picMisc.hdc was not upgraded. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="CC4C7EC0-C903-48FC-ACCC-81861D12DA4A"'
				'UPGRADE_WARNING: Couldn't resolve default property of object picBox.hdc. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
				rc = BitBlt(picBox.hdc, X, Y, 90, 18, picMisc.hdc, 0, 0, SRCCOPY)
				ShowText(picBox, X, Y + 2, 90, 18, bdFontElixirWhite, Text_Renamed, True, False)
			End If
		End If
	End Sub
	
	'UPGRADE_NOTE: Height was upgraded to Height_Renamed. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="A9E4979A-37FA-4718-9994-97DD76ED70A7"'
	Public Function ScrollBarClick(ByRef AtX As Short, ByRef AtY As Short, ByRef ButtonDown As Short, ByRef picBox As System.Windows.Forms.PictureBox, ByRef X As Short, ByRef Y As Short, ByRef Height_Renamed As Short, ByRef Index As Short, ByRef Count As Short, ByRef ScrollSize As Short) As Short
		Dim ThumbY As Short
		' Find scroll thumb position (distance in pixels down from scroll top button)
		ThumbY = Int(((Height_Renamed - 54) / Greatest(Count - ScrollSize - 1, 1)) * (Index - 1))
		' Check where click
		If PointIn(AtX, AtY, X, Y, 18, 18) And Index > 1 Then
			' Scroll Up
			If ButtonDown = True Then
				ScrollBarShow(picBox, X, Y, Height_Renamed, Index, Count - ScrollSize, 1)
				picBox.Refresh()
				ScrollBarClick = False
			Else
				Index = Index - 1
				Call PlayClickSnd(modIOFunc.ClickType.ifClick)
				ScrollBarClick = True
			End If
		ElseIf PointIn(AtX, AtY, X, Y + Height_Renamed - 18, 18, 18) And Index < Count - ScrollSize Then 
			' Scroll Down
			If ButtonDown = True Then
				ScrollBarShow(picBox, X, Y, Height_Renamed, Index, Count - ScrollSize, 2)
				picBox.Refresh()
				ScrollBarClick = False
			Else
				Index = Index + 1
				Call PlayClickSnd(modIOFunc.ClickType.ifClick)
				ScrollBarClick = True
			End If
		ElseIf PointIn(AtX, AtY, X, Y + 18, 18, ThumbY) And ButtonDown And Index > 1 Then 
			' Big Scroll Up
			Index = Greatest(Index - ScrollSize, 1)
			Call PlayClickSnd(modIOFunc.ClickType.ifClick)
			ScrollBarClick = True
		ElseIf PointIn(AtX, AtY, X, Y + 36 + ThumbY, 18, Height_Renamed - ThumbY - 54) And ButtonDown And Index < Count - ScrollSize Then 
			' Big Scroll Down
			Index = Least(Index + ScrollSize, Count - ScrollSize)
			Call PlayClickSnd(modIOFunc.ClickType.ifClick)
			ScrollBarClick = True
		End If
	End Function
	
	'UPGRADE_NOTE: Height was upgraded to Height_Renamed. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="A9E4979A-37FA-4718-9994-97DD76ED70A7"'
	Public Function ScrollBarShow(ByRef picBox As System.Windows.Forms.PictureBox, ByRef X As Short, ByRef Y As Short, ByRef Height_Renamed As Short, ByRef Index As Short, ByRef Count As Short, ByRef ButtonDown As Short) As Short
		' ButtonDown = 1 for top scroll, ButtonDown = 2 for bottom scroll, 0 is no scroll
		Dim rc As Integer
		Dim ThumbY As Short
		' Show top and bottom scroll buttons
		Select Case ButtonDown
			Case 0 ' No Scroll
				'UPGRADE_ISSUE: PictureBox property picMisc.hdc was not upgraded. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="CC4C7EC0-C903-48FC-ACCC-81861D12DA4A"'
				'UPGRADE_ISSUE: PictureBox property picBox.hdc was not upgraded. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="CC4C7EC0-C903-48FC-ACCC-81861D12DA4A"'
				rc = BitBlt(picBox.hdc, X, Y, 18, 18, picMisc.hdc, 72, 18, SRCCOPY)
				'UPGRADE_ISSUE: PictureBox property picMisc.hdc was not upgraded. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="CC4C7EC0-C903-48FC-ACCC-81861D12DA4A"'
				'UPGRADE_ISSUE: PictureBox property picBox.hdc was not upgraded. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="CC4C7EC0-C903-48FC-ACCC-81861D12DA4A"'
				rc = BitBlt(picBox.hdc, X, Y + Height_Renamed - 18, 18, 18, picMisc.hdc, 36, 18, SRCCOPY)
			Case 1 ' Scroll Up
				'UPGRADE_ISSUE: PictureBox property picMisc.hdc was not upgraded. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="CC4C7EC0-C903-48FC-ACCC-81861D12DA4A"'
				'UPGRADE_ISSUE: PictureBox property picBox.hdc was not upgraded. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="CC4C7EC0-C903-48FC-ACCC-81861D12DA4A"'
				rc = BitBlt(picBox.hdc, X, Y, 18, 18, picMisc.hdc, 90, 18, SRCCOPY)
				'UPGRADE_ISSUE: PictureBox property picMisc.hdc was not upgraded. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="CC4C7EC0-C903-48FC-ACCC-81861D12DA4A"'
				'UPGRADE_ISSUE: PictureBox property picBox.hdc was not upgraded. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="CC4C7EC0-C903-48FC-ACCC-81861D12DA4A"'
				rc = BitBlt(picBox.hdc, X, Y + Height_Renamed - 18, 18, 18, picMisc.hdc, 36, 18, SRCCOPY)
			Case 2 ' Scroll Down
				'UPGRADE_ISSUE: PictureBox property picMisc.hdc was not upgraded. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="CC4C7EC0-C903-48FC-ACCC-81861D12DA4A"'
				'UPGRADE_ISSUE: PictureBox property picBox.hdc was not upgraded. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="CC4C7EC0-C903-48FC-ACCC-81861D12DA4A"'
				rc = BitBlt(picBox.hdc, X, Y, 18, 18, picMisc.hdc, 72, 18, SRCCOPY)
				'UPGRADE_ISSUE: PictureBox property picMisc.hdc was not upgraded. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="CC4C7EC0-C903-48FC-ACCC-81861D12DA4A"'
				'UPGRADE_ISSUE: PictureBox property picBox.hdc was not upgraded. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="CC4C7EC0-C903-48FC-ACCC-81861D12DA4A"'
				rc = BitBlt(picBox.hdc, X, Y + Height_Renamed - 18, 18, 18, picMisc.hdc, 54, 18, SRCCOPY)
		End Select
		' Position Thumb Button
		ThumbY = Int(((Height_Renamed - 54) / Greatest(Count - 1, 1)) * (Index - 1))
		' [Titi 2.4.9] negative offset if last item removed in the calling proc
		If ThumbY < 0 Then ThumbY = 0
		If ThumbY > Height_Renamed - 54 Then ThumbY = Height_Renamed - 54
		ScrollBarShow = Y + 18 + ThumbY
		'UPGRADE_ISSUE: PictureBox property picMisc.hdc was not upgraded. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="CC4C7EC0-C903-48FC-ACCC-81861D12DA4A"'
		'UPGRADE_ISSUE: PictureBox property picBox.hdc was not upgraded. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="CC4C7EC0-C903-48FC-ACCC-81861D12DA4A"'
		rc = BitBlt(picBox.hdc, X, ScrollBarShow, 18, 18, picMisc.hdc, 0, 18, SRCCOPY)
	End Function
	
	'UPGRADE_NOTE: Height was upgraded to Height_Renamed. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="A9E4979A-37FA-4718-9994-97DD76ED70A7"'
	Public Function WorldScrollBarClick(ByRef AtX As Short, ByRef AtY As Short, ByRef ButtonDown As Short, ByRef picBox As System.Windows.Forms.PictureBox, ByRef X As Short, ByRef Y As Short, ByRef Height_Renamed As Short, ByRef Index As Short, ByRef Count As Short, ByRef ScrollSize As Short) As Short
		Dim ThumbY As Short
		' Find scroll thumb position (distance in pixels down from scroll top button)
		ThumbY = Int((Height_Renamed - 54) / Count) * (Index - 1)
		' Check where click
		If PointIn(AtX, AtY, X, Y, 18, 18) And Index > 0 Then
			' Scroll Up
			If ButtonDown = True Then
				WorldScrollBarShow(picBox, X, Y, Height_Renamed, Index, Count - ScrollSize, 1)
				picBox.Refresh()
				WorldScrollBarClick = False
			Else
				Index = Index - 1
				Call PlayClickSnd(modIOFunc.ClickType.ifClick)
				WorldScrollBarClick = True
			End If
		ElseIf PointIn(AtX, AtY, X, Y + Height_Renamed - 18, 18, 18) And Index < Count - ScrollSize Then 
			' Scroll Down
			If ButtonDown = True Then
				WorldScrollBarShow(picBox, X, Y, Height_Renamed, Index, Count - ScrollSize, 2)
				picBox.Refresh()
				WorldScrollBarClick = False
			Else
				Index = Index + 1
				Call PlayClickSnd(modIOFunc.ClickType.ifClick)
				WorldScrollBarClick = True
			End If
		ElseIf PointIn(AtX, AtY, X, Y + 18, 18, ThumbY) And ButtonDown And Index > 0 Then 
			' Big Scroll Up
			Index = Greatest(Index - ScrollSize, 1)
			Call PlayClickSnd(modIOFunc.ClickType.ifClick)
			WorldScrollBarClick = True
		ElseIf PointIn(AtX, AtY, X, Y + 36 + ThumbY, 18, Height_Renamed - ThumbY - 54) And ButtonDown And Index < Count - ScrollSize Then 
			' Big Scroll Down
			Index = Least(Index + ScrollSize, Count - ScrollSize)
			Call PlayClickSnd(modIOFunc.ClickType.ifClick)
			WorldScrollBarClick = True
		End If
	End Function
	
	'UPGRADE_NOTE: Height was upgraded to Height_Renamed. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="A9E4979A-37FA-4718-9994-97DD76ED70A7"'
	Public Function WorldScrollBarShow(ByRef picBox As System.Windows.Forms.PictureBox, ByRef X As Short, ByRef Y As Short, ByRef Height_Renamed As Short, ByRef Index As Short, ByRef Count As Short, ByRef ButtonDown As Short) As Short
		' ButtonDown = 1 for top scroll, ButtonDown = 2 for bottom scroll, 0 is no scroll
		Dim rc As Integer
		Dim ThumbY As Short
		' Show top and bottom scroll buttons
		Select Case ButtonDown
			Case 0 ' No Scroll
				'UPGRADE_ISSUE: PictureBox property picMisc.hdc was not upgraded. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="CC4C7EC0-C903-48FC-ACCC-81861D12DA4A"'
				'UPGRADE_ISSUE: PictureBox property picBox.hdc was not upgraded. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="CC4C7EC0-C903-48FC-ACCC-81861D12DA4A"'
				rc = BitBlt(picBox.hdc, X, Y, 18, 18, picMisc.hdc, 72, 18, SRCCOPY)
				'UPGRADE_ISSUE: PictureBox property picMisc.hdc was not upgraded. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="CC4C7EC0-C903-48FC-ACCC-81861D12DA4A"'
				'UPGRADE_ISSUE: PictureBox property picBox.hdc was not upgraded. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="CC4C7EC0-C903-48FC-ACCC-81861D12DA4A"'
				rc = BitBlt(picBox.hdc, X, Y + Height_Renamed - 18, 18, 18, picMisc.hdc, 36, 18, SRCCOPY)
			Case 1 ' Scroll Up
				'UPGRADE_ISSUE: PictureBox property picMisc.hdc was not upgraded. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="CC4C7EC0-C903-48FC-ACCC-81861D12DA4A"'
				'UPGRADE_ISSUE: PictureBox property picBox.hdc was not upgraded. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="CC4C7EC0-C903-48FC-ACCC-81861D12DA4A"'
				rc = BitBlt(picBox.hdc, X, Y, 18, 18, picMisc.hdc, 90, 18, SRCCOPY)
				'UPGRADE_ISSUE: PictureBox property picMisc.hdc was not upgraded. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="CC4C7EC0-C903-48FC-ACCC-81861D12DA4A"'
				'UPGRADE_ISSUE: PictureBox property picBox.hdc was not upgraded. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="CC4C7EC0-C903-48FC-ACCC-81861D12DA4A"'
				rc = BitBlt(picBox.hdc, X, Y + Height_Renamed - 18, 18, 18, picMisc.hdc, 36, 18, SRCCOPY)
			Case 2 ' Scroll Down
				'UPGRADE_ISSUE: PictureBox property picMisc.hdc was not upgraded. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="CC4C7EC0-C903-48FC-ACCC-81861D12DA4A"'
				'UPGRADE_ISSUE: PictureBox property picBox.hdc was not upgraded. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="CC4C7EC0-C903-48FC-ACCC-81861D12DA4A"'
				rc = BitBlt(picBox.hdc, X, Y, 18, 18, picMisc.hdc, 72, 18, SRCCOPY)
				'UPGRADE_ISSUE: PictureBox property picMisc.hdc was not upgraded. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="CC4C7EC0-C903-48FC-ACCC-81861D12DA4A"'
				'UPGRADE_ISSUE: PictureBox property picBox.hdc was not upgraded. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="CC4C7EC0-C903-48FC-ACCC-81861D12DA4A"'
				rc = BitBlt(picBox.hdc, X, Y + Height_Renamed - 18, 18, 18, picMisc.hdc, 54, 18, SRCCOPY)
		End Select
		' Position Thumb Button
		ThumbY = Int((Height_Renamed - 54) / Greatest(1, Count - 1)) * (Index - 1)
		WorldScrollBarShow = Y + 18 + ThumbY
		'UPGRADE_ISSUE: PictureBox property picMisc.hdc was not upgraded. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="CC4C7EC0-C903-48FC-ACCC-81861D12DA4A"'
		'UPGRADE_ISSUE: PictureBox property picBox.hdc was not upgraded. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="CC4C7EC0-C903-48FC-ACCC-81861D12DA4A"'
		rc = BitBlt(picBox.hdc, X, WorldScrollBarShow, 18, 18, picMisc.hdc, 0, 18, SRCCOPY)
	End Function
	
	'UPGRADE_NOTE: Text was upgraded to Text_Renamed. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="A9E4979A-37FA-4718-9994-97DD76ED70A7"'
	'UPGRADE_NOTE: Height was upgraded to Height_Renamed. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="A9E4979A-37FA-4718-9994-97DD76ED70A7"'
	'UPGRADE_NOTE: Width was upgraded to Width_Renamed. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="A9E4979A-37FA-4718-9994-97DD76ED70A7"'
	Public Function ShowText(ByRef picBox As Object, ByVal X As Short, ByVal Y As Short, ByVal Width_Renamed As Short, ByVal Height_Renamed As Short, ByRef FontNum As Short, ByRef Text_Renamed As String, ByRef Center As Short, ByRef ReturnRows As Short) As Short
		' IF Center = True (-1), then Center Text
		' IF Center = False (0), then Left Justify Text
		' IF Center = 1, the Right Justify Text
		Dim LastSpace, PosX, PosY, LastPixels As Short
		'UPGRADE_NOTE: FontHeight was upgraded to FontHeight_Renamed. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="A9E4979A-37FA-4718-9994-97DD76ED70A7"'
		Dim FontIndex, FontSpace, FontTop, FontMask, FontHeight_Renamed, FontLine As Short
		Dim TextLen, TextPos, TextPixels As Short
		Dim c, n As Short
		Dim rc As Integer
		' Set up Font
		Select Case FontNum
			Case bdFontSmallWhite
				FontIndex = 0 : FontTop = 0 : FontMask = 20 : FontSpace = 6 : FontHeight_Renamed = 10 : FontLine = 1
			Case bdFontNoxiousWhite
				FontIndex = 1 : FontTop = 40 : FontMask = 208 : FontSpace = 6 : FontHeight_Renamed = 14 : FontLine = 2
			Case bdFontNoxiousGold
				FontIndex = 1 : FontTop = 82 : FontMask = 208 : FontSpace = 6 : FontHeight_Renamed = 14 : FontLine = 2
			Case bdFontNoxiousGrey
				FontIndex = 1 : FontTop = 124 : FontMask = 208 : FontSpace = 6 : FontHeight_Renamed = 14 : FontLine = 2
			Case bdFontNoxiousBlack
				FontIndex = 1 : FontTop = 166 : FontMask = 208 : FontSpace = 6 : FontHeight_Renamed = 14 : FontLine = 2
			Case bdFontElixirWhite
				FontIndex = 2 : FontTop = 250 : FontMask = 362 : FontSpace = 8 : FontHeight_Renamed = 14 : FontLine = 2
			Case bdFontElixirGold
				FontIndex = 2 : FontTop = 278 : FontMask = 362 : FontSpace = 8 : FontHeight_Renamed = 14 : FontLine = 2
			Case bdFontElixirGrey
				FontIndex = 2 : FontTop = 306 : FontMask = 362 : FontSpace = 8 : FontHeight_Renamed = 14 : FontLine = 2
			Case bdFontElixirBlack
				FontIndex = 2 : FontTop = 334 : FontMask = 362 : FontSpace = 8 : FontHeight_Renamed = 14 : FontLine = 2
			Case bdFontLargeWhite
				FontIndex = 3 : FontTop = 390 : FontMask = 408 : FontSpace = 17 : FontHeight_Renamed = 18 : FontLine = 2
		End Select
		' Initialize positions
		TextLen = LenStr(Text_Renamed) : TextPos = 1 : PosX = X : PosY = Y
		' Loop until reach end of Text
		Do Until TextPos > TextLen Or PosY > Y + Height_Renamed
			' Length of Next Line (in Pixels)
			TextPixels = 0 : c = TextPos : LastSpace = TextPos
			Do 
				If c > TextLen Then
					LastSpace = c
					LastPixels = TextPixels
					Exit Do
				End If
				n = Least(Asc(Mid(Text_Renamed, c, 1)) - 33, 89)
				If n = 59 Then
					TextPixels = TextPixels + 18
				ElseIf n > -1 Then 
					TextPixels = TextPixels + aW(FontIndex, n)
				Else
					LastSpace = c
					LastPixels = TextPixels
					TextPixels = TextPixels + FontSpace
				End If
				c = c + 1
			Loop Until (TextPixels > Width_Renamed Or n = -20)
			' If centering, then adjust PosX
			If Center = True Then
				PosX = PosX + (Width_Renamed - LastPixels) / 2
			ElseIf Center = 1 Then 
				PosX = PosX + (Width_Renamed - LastPixels)
			End If
			' Print Next Line
			For c = TextPos To LastSpace - 1
				n = Asc(Mid(Text_Renamed, c, 1)) - 33
				If n = 59 Then ' Special Box as back slash
					'UPGRADE_ISSUE: PictureBox property picMisc.hdc was not upgraded. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="CC4C7EC0-C903-48FC-ACCC-81861D12DA4A"'
					'UPGRADE_WARNING: Couldn't resolve default property of object picBox.hdc. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
					rc = BitBlt(picBox.hdc, PosX + 2, PosY - 1, 16, 16, picMisc.hdc, 1, 19, SRCCOPY)
					PosX = PosX + 18
				ElseIf n > -1 And n < 90 Then 
					'UPGRADE_ISSUE: PictureBox property picFonts.hdc was not upgraded. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="CC4C7EC0-C903-48FC-ACCC-81861D12DA4A"'
					'UPGRADE_WARNING: Couldn't resolve default property of object picBox.hdc. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
					rc = BitBlt(picBox.hdc, PosX, PosY, aW(FontIndex, n), FontHeight_Renamed, picFonts.hdc, aX(FontIndex, n), aY(FontIndex, n) + FontMask, SRCAND)
					'UPGRADE_ISSUE: PictureBox property picFonts.hdc was not upgraded. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="CC4C7EC0-C903-48FC-ACCC-81861D12DA4A"'
					'UPGRADE_WARNING: Couldn't resolve default property of object picBox.hdc. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
					rc = BitBlt(picBox.hdc, PosX, PosY, aW(FontIndex, n), FontHeight_Renamed, picFonts.hdc, aX(FontIndex, n), aY(FontIndex, n) + FontTop, SRCPAINT)
					PosX = PosX + aW(FontIndex, n)
				Else
					PosX = PosX + FontSpace
				End If
			Next c
			' Bump a space
			PosY = PosY + FontHeight_Renamed + FontLine
			PosX = X
			TextPos = LastSpace + 1
		Loop 
		If ReturnRows = True Then
			ShowText = PosY - Y
		Else
			ShowText = TextLen - TextPos
		End If
	End Function
	
	'UPGRADE_NOTE: Tome was upgraded to Tome_Renamed. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="A9E4979A-37FA-4718-9994-97DD76ED70A7"'
	'UPGRADE_NOTE: Area was upgraded to Area_Renamed. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="A9E4979A-37FA-4718-9994-97DD76ED70A7"'
	Public Function TomeSaveArea(ByRef sPath As String) As Boolean
		Dim Area_Renamed As Object
		Dim Tome_Renamed As Object
		Dim c As Short
		'UPGRADE_WARNING: Screen property Screen.MousePointer has a new behavior. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6BA9B8D2-2A32-4B6E-8D36-44949974A5B4"'
		System.Windows.Forms.Cursor.Current = System.Windows.Forms.Cursors.WaitCursor
		On Error GoTo Err_Handler
		' Save Tome
		c = FreeFile
		FileOpen(c, sPath & "\" & Tome.FileName, OpenMode.Binary)
		'UPGRADE_WARNING: Couldn't resolve default property of object Tome.SaveToFile. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
		Tome.SaveToFile(c)
		FileClose(c)
		' Save Area
		c = FreeFile
		FileOpen(c, sPath & "\" & Area.FileName, OpenMode.Binary)
		'UPGRADE_WARNING: Couldn't resolve default property of object Area.Plot. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
		Area.Plot.SaveToFile(c)
		FileClose(c)
		TomeSaveArea = True
Exit_Function: 
		'UPGRADE_WARNING: Screen property Screen.MousePointer has a new behavior. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6BA9B8D2-2A32-4B6E-8D36-44949974A5B4"'
		System.Windows.Forms.Cursor.Current = System.Windows.Forms.Cursors.Default
		Exit Function
Err_Handler: 
		TomeSaveArea = False
		oErr.logError("TomeSaveArea")
		Resume Exit_Function
	End Function
	
	Public Sub TomeSave(ByRef SaveAsNew As Short, Optional ByRef ShowMessage As Boolean = True)
		Dim FileName, PathName As String
		Dim c As Short
		Dim rc As Integer
		Dim sMsg As String
		On Error GoTo ErrorHandler ' [Titi 2.4.9]
		'UPGRADE_WARNING: Screen property Screen.MousePointer has a new behavior. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6BA9B8D2-2A32-4B6E-8D36-44949974A5B4"'
		System.Windows.Forms.Cursor.Current = System.Windows.Forms.Cursors.WaitCursor
		' Save as new folder or over existing
		If SaveAsNew = True Then
			PathName = gAppPath & "\world\" & WorldNow.Name & "\save" & VB6.Format(TomeSavePos)
			MkDir(PathName)
		Else
			PathName = TomeSavePathName & "\" & oFileSys.AssignTmpName(clsInOut.IOActionType.Folder)
			Call oFileSys.CheckExists(PathName, clsInOut.IOActionType.Folder, True)
		End If
		' Setup an image of the Map
		'UPGRADE_ISSUE: PictureBox method picTmp.Cls was not upgraded. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="CC4C7EC0-C903-48FC-ACCC-81861D12DA4A"'
		picTmp.Cls() : picTmp.Width = 220 : picTmp.Height = 165
		'UPGRADE_ISSUE: PictureBox property picTmp.hdc was not upgraded. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="CC4C7EC0-C903-48FC-ACCC-81861D12DA4A"'
		rc = SetStretchBltMode(picTmp.hdc, 3)
		'UPGRADE_ISSUE: PictureBox property picMap.hdc was not upgraded. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="CC4C7EC0-C903-48FC-ACCC-81861D12DA4A"'
		'UPGRADE_ISSUE: PictureBox property picTmp.hdc was not upgraded. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="CC4C7EC0-C903-48FC-ACCC-81861D12DA4A"'
		rc = StretchBlt(picTmp.hdc, 0, 0, 220, 165, picMap.hdc, 0, 0, picBox.Width, picBox.Height, SRCCOPY)
		' Save screen shot
		'UPGRADE_WARNING: SavePicture was upgraded to System.Drawing.Image.Save and has a new behavior. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="9B7D5ADD-D8FE-4819-A36C-6DEDAF088CC7"'
		picTmp.Image.Save(PathName & "\screen.bmp")
		' Save description
		c = FreeFile
		FileOpen(c, PathName & "\name.dat", OpenMode.Binary)
		'UPGRADE_WARNING: Put was upgraded to FilePut and has a new behavior. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="9B7D5ADD-D8FE-4819-A36C-6DEDAF088CC7"'
		FilePut(c, Len(CreateNameNew))
		'UPGRADE_WARNING: Put was upgraded to FilePut and has a new behavior. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="9B7D5ADD-D8FE-4819-A36C-6DEDAF088CC7"'
		FilePut(c, CreateNameNew)
		FileClose(c)
		' Save Area
		' [borf] Copy all files from saved as folder or temp folder
		If TomeSaveArea(PathName) Then
			Call oFileSys.Copy(PathName & "\*.*", gAppPath & "\world\" & WorldNow.Name, clsInOut.IOActionType.File, True)
		End If
		If Not SaveAsNew Then
			Call oFileSys.Copy(PathName & "\*.*", TomeSavePathName, clsInOut.IOActionType.File, True)
			Call oFileSys.Delete(PathName, clsInOut.IOActionType.Folder, True)
		End If
		sMsg = "Tome saved."
Exit_Sub: 
		'UPGRADE_WARNING: Screen property Screen.MousePointer has a new behavior. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6BA9B8D2-2A32-4B6E-8D36-44949974A5B4"'
		System.Windows.Forms.Cursor.Current = System.Windows.Forms.Cursors.Default
		If ShowMessage Then
			DialogDM(sMsg)
		End If
		Exit Sub
ErrorHandler: 
		If Not SaveAsNew Then
			Call oFileSys.Delete(PathName, clsInOut.IOActionType.Folder, True)
		End If
		sMsg = "Failed to save tome."
		Resume Exit_Sub
	End Sub
	
	'UPGRADE_NOTE: Tome was upgraded to Tome_Renamed. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="A9E4979A-37FA-4718-9994-97DD76ED70A7"'
	Private Sub TomeRestore(ByRef FromFile As String, Optional ByRef blnHome As Boolean = False)
		Dim Tome_Renamed As Object
		Dim FileName As String
		Dim c As Short
		Dim FromDir As String
		' End current Combat (if any)
		If picGrid.Visible = True Then
			CombatEnd(False)
		End If
		FromDir = ClipPath(FromFile)
		' Copy all the files from the Save to World
		Call oFileSys.Copy(FromDir & "\*.tom", gAppPath & "\world\" & WorldNow.Name, clsInOut.IOActionType.File, True)
		Call oFileSys.Copy(FromDir & "\*.rsa", gAppPath & "\world\" & WorldNow.Name, clsInOut.IOActionType.File, True)
		' Clear current game
		GameClear()
		' Read in the new Tome
		If oFileSys.CheckExists(FromFile, clsInOut.IOActionType.File) Then
			c = FreeFile
			FileOpen(c, FromFile, OpenMode.Binary)
			Tome = New Tome
			'UPGRADE_WARNING: Couldn't resolve default property of object Tome.ReadFromFile. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
			Tome.ReadFromFile(c)
			'UPGRADE_WARNING: Couldn't resolve default property of object Tome.FileName. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
			Tome.FileName = Mid(FromFile, InStrRev(FromFile, "\") + 1)
			FileClose(c)
			' Start the tome
			DialogDM("Game Restored.")
			'UPGRADE_WARNING: Couldn't resolve default property of object Tome. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
			If blnHome = False Then TomeStart(Tome)
		Else
			DialogDM("ERROR: I can't find the saved Tome file in " & FromDir & ".")
		End If
	End Sub
	
	Public Sub TomeAddRoster(ByRef CreatureX As Creature)
		Dim c As Short
		Dim CreatureZ As Creature
		' Find new available unused index identifier
		c = 1
		For	Each CreatureZ In TomeRoster
			If CreatureZ.Index >= c Then
				c = CreatureZ.Index + 1
			End If
		Next CreatureZ
		CreatureX.Index = c
		' Find alphabetical spot in list
		For c = 1 To TomeRoster.Count()
			'UPGRADE_WARNING: Couldn't resolve default property of object TomeRoster().Name. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
			If UCase(TomeRoster.Item(c).Name) > UCase(CreatureX.Name) Then
				TomeRoster.Add(CreatureX, "X" & CreatureX.Index, c)
				Exit For
			End If
		Next c
		If c > TomeRoster.Count() Then
			TomeRoster.Add(CreatureX, "X" & CreatureX.Index)
		End If
	End Sub
	
	Public Sub TomeLoadRoster()
		Dim CreatureX, CreatureZ As Creature
		Dim FileName, sPath As String
		Dim c As Short
		'UPGRADE_WARNING: Screen property Screen.MousePointer has a new behavior. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6BA9B8D2-2A32-4B6E-8D36-44949974A5B4"'
		System.Windows.Forms.Cursor.Current = System.Windows.Forms.Cursors.WaitCursor
		TomeRoster = New Collection
		sPath = gAppPath & "\roster\" & WorldNow.Name & "\"
		'UPGRADE_WARNING: Dir has a new behavior. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="9B7D5ADD-D8FE-4819-A36C-6DEDAF088CC7"'
		FileName = Dir(sPath & "*.rsc")
		Do Until FileName = ""
			CreatureX = New Creature
			c = FreeFile
			FileOpen(c, sPath & FileName, OpenMode.Binary)
			CreatureX.ReadFromFile(c)
			FileClose(c)
			TomeAddRoster(CreatureX)
			'UPGRADE_WARNING: Dir has a new behavior. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="9B7D5ADD-D8FE-4819-A36C-6DEDAF088CC7"'
			FileName = Dir()
		Loop 
		' Remove from Roster any Creature already in the Tome
		For	Each CreatureX In Tome.Creatures
			For	Each CreatureZ In TomeRoster
				If CreatureX.Name = CreatureZ.Name Then
					TomeRoster.Remove("X" & CreatureZ.Index)
				End If
			Next CreatureZ
		Next CreatureX
		' Set the top of the Roster
		TomeRosterTop = 1
		'UPGRADE_WARNING: Screen property Screen.MousePointer has a new behavior. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6BA9B8D2-2A32-4B6E-8D36-44949974A5B4"'
		System.Windows.Forms.Cursor.Current = System.Windows.Forms.Cursors.Default
	End Sub
	
	Public Sub TomeLoad()
		Dim FileName As String
		Dim c, i As Short
		Dim TomeX, TomeZ As Tome
		Dim DirList As Collection
		Dim FactoidX As Factoid
		On Error GoTo ErrorHandler
		' Start new listing of Tomes on disk
		'UPGRADE_WARNING: Screen property Screen.MousePointer has a new behavior. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6BA9B8D2-2A32-4B6E-8D36-44949974A5B4"'
		System.Windows.Forms.Cursor.Current = System.Windows.Forms.Cursors.WaitCursor
		' Add default \tomes folder
		DirList = New Collection
		FactoidX = New Factoid
		FactoidX.Text = gAppPath & "\tomes\" & WorldNow.Name
		i = i + 1
		FactoidX.Index = 0
		DirList.Add(FactoidX, "F" & FactoidX.Index)
		' Get list of directories where Tomes could be
		i = 0
		'UPGRADE_WARNING: Dir has a new behavior. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="9B7D5ADD-D8FE-4819-A36C-6DEDAF088CC7"'
		FileName = Dir(gAppPath & "\tomes\" & WorldNow.Name & "\*.*", FileAttribute.Directory)
		Do Until FileName = ""
			If FileName <> "." And FileName <> ".." Then
				FactoidX = New Factoid
				FactoidX.Text = gAppPath & "\tomes\" & WorldNow.Name & "\" & FileName
				i = i + 1
				FactoidX.Index = i
				DirList.Add(FactoidX, "F" & FactoidX.Index)
			End If
			'UPGRADE_WARNING: Dir has a new behavior. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="9B7D5ADD-D8FE-4819-A36C-6DEDAF088CC7"'
			FileName = Dir()
		Loop 
		' Now go find Tomes in those directories
		i = 0
		TomeList = New Collection
		For	Each FactoidX In DirList
			'UPGRADE_WARNING: Dir has a new behavior. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="9B7D5ADD-D8FE-4819-A36C-6DEDAF088CC7"'
			FileName = Dir(FactoidX.Text & "\*.tom")
			Do Until FileName = ""
				' Read the Tome file
				c = FreeFile
				TomeX = New Tome
				FileOpen(c, FactoidX.Text & "\" & FileName, OpenMode.Binary)
				TomeX.ReadFromFileHeader(c)
				FileClose(c)
				' Set to directory and file name pulled, set to Not on Adventure
				TomeX.FullPath = FactoidX.Text
				TomeX.FileName = FileName
				TomeX.OnAdventure = False
				' Index the Tome
				i = i + 1
				TomeX.Index = i
				' Que the Tome into the List alphabetically
				For c = 1 To TomeList.Count()
					'UPGRADE_WARNING: Couldn't resolve default property of object TomeList(c).Name. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
					If TomeList.Item(c).Name > TomeX.Name Then
						'UPGRADE_WARNING: Couldn't resolve default property of object TomeList().Index. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
						TomeList.Add(TomeX, "T" & TomeX.Index, "T" & TomeList.Item(c).Index)
						Exit For
					End If
				Next c
				If c > TomeList.Count() Then
					TomeList.Add(TomeX, "T" & TomeX.Index)
				End If
				'UPGRADE_WARNING: Dir has a new behavior. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="9B7D5ADD-D8FE-4819-A36C-6DEDAF088CC7"'
				FileName = Dir()
			Loop 
		Next FactoidX
		' Check if Tomes exists in \World folder. If so, get correct OnAdventure indicator.
		For	Each TomeX In TomeList
			If oFileSys.CheckExists(gAppPath & "\World\" & WorldNow.Name, clsInOut.IOActionType.Folder, True) Then
				'UPGRADE_WARNING: Dir has a new behavior. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="9B7D5ADD-D8FE-4819-A36C-6DEDAF088CC7"'
				FileName = Dir(gAppPath & "\World\" & WorldNow.Name & "\" & TomeX.FileName)
				If FileName <> "" Then
					TomeX.IsInPlay = True
					' Read the Tome file
					c = FreeFile
					TomeZ = New Tome
					FileOpen(c, gAppPath & "\World\" & WorldNow.Name & "\" & FileName, OpenMode.Binary)
					TomeZ.ReadFromFileHeader(c)
					FileClose(c)
					TomeX.OnAdventure = TomeZ.OnAdventure
				End If
			End If
		Next TomeX
		' Add Tome Wizard to top
		TomeX = New Tome
		i = i + 1
		TomeX.Name = "Create a new Tome"
		TomeX.Index = i
		If UberWizMaps.MapSketchs.Count() > 0 Then
			UberWizMaps.MainMapIndex = 1
		End If
		If TomeList.Count() > 0 Then
			TomeList.Add(TomeX, "T" & TomeX.Index, 1)
		Else
			TomeList.Add(TomeX, "T" & TomeX.Index)
		End If
		' Set to top
		TomeListTop = 1
		'UPGRADE_WARNING: Screen property Screen.MousePointer has a new behavior. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6BA9B8D2-2A32-4B6E-8D36-44949974A5B4"'
		System.Windows.Forms.Cursor.Current = System.Windows.Forms.Cursors.Default
		Exit Sub
ErrorHandler: 
		If VB.Right(FactoidX.Text, 4) = ".tom" Or VB.Right(FactoidX.Text, 4) = ".rsa" Then
			' found a tome directly in the world tomes directory - skip it
			Resume Next
		Else
			oErr.logError("TomeLoad")
			Resume Next
		End If
	End Sub
	
	'UPGRADE_NOTE: Map was upgraded to Map_Renamed. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="A9E4979A-37FA-4718-9994-97DD76ED70A7"'
	Private Sub LayerTileMicro(ByRef Layer As Short, ByRef XMap As Short, ByRef YMap As Short, ByRef AtX As Short, ByRef AtY As Short, ByRef AtWidth As Short, ByRef AtHeight As Short, ByRef AtOffX As Short, ByRef AtOffY As Short, ByRef Gray As Short)
		Dim Map_Renamed As Object
		Dim c, Flip As Short
		Select Case Layer
			Case bdMapBottom
				'UPGRADE_WARNING: Couldn't resolve default property of object Map.BottomTile. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
				c = Map.BottomTile(XMap, YMap)
				'UPGRADE_WARNING: Couldn't resolve default property of object Map.Hidden. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
				If c > 0 And Map.Hidden(XMap, YMap) = False Then
					'UPGRADE_WARNING: Couldn't resolve default property of object Map.Tiles. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
					c = Map.Tiles("L" & c).Pic
				Else
					c = -1
				End If
				'UPGRADE_WARNING: Couldn't resolve default property of object Map.BottomFlip. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
				Flip = Map.BottomFlip(XMap, YMap)
				PlotTileMicro(c, AtX, AtY, AtWidth, AtHeight, AtOffX, AtOffY, Flip)
			Case bdMapMiddle
				'UPGRADE_WARNING: Couldn't resolve default property of object Map.MiddleTile. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
				c = Map.MiddleTile(XMap, YMap)
				'UPGRADE_WARNING: Couldn't resolve default property of object Map.Hidden. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
				'UPGRADE_WARNING: Couldn't resolve default property of object Map.Width. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
				If c > 0 And (Map.Hidden(XMap, YMap) = False Or Map.Hidden(Least(XMap + 1, Map.Width), YMap) = False Or Map.Hidden(XMap, Greatest(YMap - 1, 0)) = False) Then
					'UPGRADE_WARNING: Couldn't resolve default property of object Map.Tiles. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
					c = Map.Tiles("L" & c).Pic
					'UPGRADE_WARNING: Couldn't resolve default property of object Map.MiddleFlip. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
					Flip = Map.MiddleFlip(XMap, YMap)
					PlotTileMicro(c, AtX, AtY, AtWidth, AtHeight, AtOffX, AtOffY, Flip)
				End If
			Case bdMapTop
				'UPGRADE_WARNING: Couldn't resolve default property of object Map.TopTile. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
				c = Map.TopTile(XMap, YMap)
				'UPGRADE_WARNING: Couldn't resolve default property of object Map.Hidden. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
				'UPGRADE_WARNING: Couldn't resolve default property of object Map.Width. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
				If c > 0 And (Map.Hidden(XMap, YMap) = False Or Map.Hidden(Least(XMap + 1, Map.Width), YMap) = False Or Map.Hidden(XMap, Greatest(YMap - 1, 0)) = False) Then
					'UPGRADE_WARNING: Couldn't resolve default property of object Map.Tiles. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
					c = Map.Tiles("L" & c).Pic
					'UPGRADE_WARNING: Couldn't resolve default property of object Map.TopFlip. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
					Flip = Map.TopFlip(XMap, YMap)
					PlotTileMicro(c, AtX, AtY, AtWidth, AtHeight, AtOffX, AtOffY, Flip)
				End If
		End Select
	End Sub
	
	Public Sub DrawTileMicro(ByRef AtX As Short, ByRef AtY As Short, ByRef Side As Short)
		' AtX is a Pixel / (bdTileWidth/ 2), AtY is Pixel / (bdTileHeight / 3)
		'UPGRADE_NOTE: my was upgraded to my_Renamed. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="A9E4979A-37FA-4718-9994-97DD76ED70A7"'
		Dim mx, c, my_Renamed As Short
		Dim X, Y As Short
		Dim rc As Integer
		Dim OffX, OffY As Short
		'UPGRADE_NOTE: Height was upgraded to Height_Renamed. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="A9E4979A-37FA-4718-9994-97DD76ED70A7"'
		'UPGRADE_NOTE: Width was upgraded to Width_Renamed. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="A9E4979A-37FA-4718-9994-97DD76ED70A7"'
		Dim Width_Renamed, Height_Renamed As Short
		Dim XMap, YMap As Short
		' Convert AtX and AtY to Map Coordinates
		XMap = MicroMapLeft + AtX - Int(AtY / 2)
		YMap = MicroMapTop + AtX + Int((AtY + 1) / 2) - 1
		X = (AtX * 2 - 1 + (AtY Mod 2)) * (bdTileWidth / 2)
		Y = (AtY - 2) * (bdTileHeight / 3)
		' The calc above is spot on. I've double checked it.
		OffX = 0 : OffY = 0
		Height_Renamed = bdTileHeight : Width_Renamed = bdTileWidth
		' Set Offsets and Width/Height
		If (Side And bdSideTop) > 0 Then
			OffY = (bdTileHeight / 3) * 2 : Height_Renamed = (bdTileHeight / 3)
		End If
		If (Side And bdSideLeft) > 0 Then
			OffX = (bdTileWidth / 2) : Width_Renamed = (bdTileWidth / 2)
		End If
		If (Side And bdSideFirstBottom) > 0 Then
			Height_Renamed = (bdTileHeight / 3) * 2
		End If
		If (Side And bdSideSecondBottom) > 0 Then
			Height_Renamed = (bdTileHeight / 3)
		End If
		If (Side And bdSideRight) > 0 Then
			Width_Renamed = (bdTileWidth / 2)
		End If
		' LayerTile: XMap, YMap, x, y, width, height, offx, offy
		' Plot blank tile if out of bounds
		If XMap < 0 Or YMap < 0 Or XMap > Map.Width Or YMap > Map.Height Then
			PlotTileMicro(bdTileBlack, X + OffX, Y + OffY, Width_Renamed, Height_Renamed, OffX, OffY, 0)
		Else
			LayerTileMicro(bdMapBottom, XMap, YMap, X + OffX, Y + OffY, Width_Renamed, Height_Renamed, OffX, OffY, 0)
			LayerTileMicro(bdMapMiddle, XMap, YMap, X + OffX, Y + OffY, Width_Renamed, Height_Renamed, OffX, OffY, 0)
			LayerTileMicro(bdMapTop, XMap, YMap, X + OffX, Y + OffY, Width_Renamed, Height_Renamed, OffX, OffY, 0)
			DrawTileAnnotate(XMap, YMap)
		End If
	End Sub
	
	Private Sub PlotTileMicro(ByRef TileToPlot As Short, ByRef X As Short, ByRef Y As Short, ByRef XWidth As Short, ByRef YWidth As Short, ByRef XSrcOff As Short, ByRef YSrcOff As Short, ByRef XFlip As Short)
		Dim rc As Short
		Dim MaxHeight, MaxWidth As Short
		Dim TileX, TileY As Short
		' Draw tile
		If TileToPlot < 0 Then
			'UPGRADE_ISSUE: PictureBox property picBlack.hdc was not upgraded. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="CC4C7EC0-C903-48FC-ACCC-81861D12DA4A"'
			'UPGRADE_ISSUE: PictureBox property picMicroMap.hdc was not upgraded. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="CC4C7EC0-C903-48FC-ACCC-81861D12DA4A"'
			rc = BitBlt(picMicroMap.hdc, X, Y, XWidth, YWidth, picBlack.hdc, bdBlackWidth + XSrcOff, YSrcOff, SRCAND)
			'UPGRADE_ISSUE: PictureBox property picBlack.hdc was not upgraded. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="CC4C7EC0-C903-48FC-ACCC-81861D12DA4A"'
			'UPGRADE_ISSUE: PictureBox property picMicroMap.hdc was not upgraded. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="CC4C7EC0-C903-48FC-ACCC-81861D12DA4A"'
			rc = BitBlt(picMicroMap.hdc, X, Y, XWidth, YWidth, picBlack.hdc, bdBlackWidth + XSrcOff, picBlackHeightSmall + YSrcOff, SRCPAINT)
		Else
			' Set height, width and location of tile in bmp
			MaxHeight = (picTSmall.Height / (bdTileHeight * 2))
			MaxWidth = (picTSmall.Width / (bdTileWidth * 2))
			TileX = bdTileWidth * Int(TileToPlot / MaxHeight)
			TileY = bdTileHeight * (TileToPlot Mod MaxHeight)
			If XFlip = True Then
				TileX = picTSmall.Width - bdTileWidth * (TileX / bdTileWidth) - bdTileWidth
			End If
			'UPGRADE_ISSUE: PictureBox property picTSmall.hdc was not upgraded. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="CC4C7EC0-C903-48FC-ACCC-81861D12DA4A"'
			'UPGRADE_ISSUE: PictureBox property picMicroMap.hdc was not upgraded. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="CC4C7EC0-C903-48FC-ACCC-81861D12DA4A"'
			rc = BitBlt(picMicroMap.hdc, X, Y, XWidth, YWidth, picTSmall.hdc, TileX + XSrcOff, TileY + YSrcOff + (picTSmall.Height / 2), SRCAND)
			'UPGRADE_ISSUE: PictureBox property picTSmall.hdc was not upgraded. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="CC4C7EC0-C903-48FC-ACCC-81861D12DA4A"'
			'UPGRADE_ISSUE: PictureBox property picMicroMap.hdc was not upgraded. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="CC4C7EC0-C903-48FC-ACCC-81861D12DA4A"'
			rc = BitBlt(picMicroMap.hdc, X, Y, XWidth, YWidth, picTSmall.hdc, TileX + XSrcOff, TileY + YSrcOff, SRCPAINT)
		End If
	End Sub
	
	Private Sub MenuSkillPostRaise()
		Dim TriggerX As Trigger
		Dim c As Short
		Dim NoFail As Short
		For	Each TriggerX In CreatureWithTurn.Triggers
			' For each Skill which has gone up a level, fire Triggers
			If TriggerX.Style(0) = True And TriggerX.SkillPoints > TriggerX.TempSkill Then
				For c = Int(TriggerX.TempSkill / TriggerX.Turns) + 1 To Int(TriggerX.SkillPoints / TriggerX.Turns)
					GlobalSkillLevel = c
					CreatureNow = CreatureWithTurn
					NoFail = FireTriggers(TriggerX, bdPostRaiseSkill)
				Next c
			End If
		Next TriggerX
	End Sub
	
	Private Sub MenuSkillClose()
		MenuSkillPostRaise()
		picTomeNew.Visible = False
		Frozen = False
		If picGrid.Visible = True Then
			picGrid.Focus()
		Else
			picMap.Focus()
		End If
	End Sub
	
	Private Sub MenuSkillClick(ByRef AtX As Short, ByRef AtY As Short, ByRef ButtonDown As Short)
		Dim c As Short
		Dim rc As Integer
		Dim TriggerX As Trigger
		If PointIn(AtX, AtY, picTomeNew.ClientRectangle.Width - 102, 364, 90, 18) Then
			' Done
			If ButtonDown Then
				ShowButton(picTomeNew, picTomeNew.ClientRectangle.Width - 102, 364, "Done", True)
				picTomeNew.Refresh()
			Else
				ShowButton(picTomeNew, picTomeNew.ClientRectangle.Width - 102, 364, "Done", False)
				picTomeNew.Refresh()
				Call PlayClickSnd(modIOFunc.ClickType.ifClick)
				MenuSkillClose()
			End If
		ElseIf PointIn(AtX, AtY, picTomeNew.ClientRectangle.Width - 198, 364, 90, 18) And ScrollList.Count() > 0 And CreateSkillsIndex > 0 Then 
			' Use
			'UPGRADE_WARNING: Couldn't resolve default property of object ScrollList().TriggerType. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
			c = ScrollList.Item(CreateSkillsIndex).TriggerType
			Select Case c
				Case bdOnCast, bdOnSkillUse, bdOnCastTarget, bdOnSkillTarget
					If ButtonDown Then
						ShowButton(picTomeNew, picTomeNew.ClientRectangle.Width - 198, 364, "Use", True)
						picTomeNew.Refresh()
					Else
						ShowButton(picTomeNew, picTomeNew.ClientRectangle.Width - 198, 364, "Use", False)
						picTomeNew.Refresh()
						Call PlayClickSnd(modIOFunc.ClickType.ifClick)
						picTomeNew.Visible = False
						' If targeting, acquire target
						Select Case c
							Case bdOnCastTarget, bdOnSkillTarget
								' Range is 99, you can target anything, yes-dead too
								If TargetCreature(CreatureTarget, 99, 0, 1) = False Then
									Frozen = False
									Exit Sub
								End If
						End Select
						Frozen = False
						MenuSkillUse(False)
					End If
			End Select
		ElseIf PointIn(AtX, AtY, 60, 331, 90, 18) And CreateSkillsIndex > 0 Then 
			' Lower Skill
			'UPGRADE_WARNING: Couldn't resolve default property of object ScrollList(CreateSkillsIndex).TempSkill. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
			'UPGRADE_WARNING: Couldn't resolve default property of object ScrollList(CreateSkillsIndex).SkillPoints. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
			'UPGRADE_WARNING: Couldn't resolve default property of object ScrollList(CreateSkillsIndex).Turns. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
			If CreatureWithTurn.SkillPoints >= ScrollList.Item(CreateSkillsIndex).Turns Or ScrollList.Item(CreateSkillsIndex).SkillPoints <> ScrollList.Item(CreateSkillsIndex).TempSkill Then
				If ButtonDown Then
					ShowButton(picTomeNew, 60, 331, "Lower", True)
					picTomeNew.Refresh()
				Else
					ShowButton(picTomeNew, 60, 331, "Lower", False)
					picTomeNew.Refresh()
					Call PlayClickSnd(modIOFunc.ClickType.ifClick)
					MenuSkillSpend(False)
				End If
			End If
		ElseIf PointIn(AtX, AtY, 156, 331, 90, 18) And CreateSkillsIndex > 0 Then 
			' Raise Skill
			'UPGRADE_WARNING: Couldn't resolve default property of object ScrollList(CreateSkillsIndex).TempSkill. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
			'UPGRADE_WARNING: Couldn't resolve default property of object ScrollList(CreateSkillsIndex).SkillPoints. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
			'UPGRADE_WARNING: Couldn't resolve default property of object ScrollList(CreateSkillsIndex).Turns. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
			If CreatureWithTurn.SkillPoints >= ScrollList.Item(CreateSkillsIndex).Turns Or ScrollList.Item(CreateSkillsIndex).SkillPoints <> ScrollList.Item(CreateSkillsIndex).TempSkill Then
				If ButtonDown Then
					ShowButton(picTomeNew, 156, 331, "Raise", True)
					picTomeNew.Refresh()
				Else
					ShowButton(picTomeNew, 156, 331, "Raise", False)
					picTomeNew.Refresh()
					Call PlayClickSnd(modIOFunc.ClickType.ifClick)
					MenuSkillSpend(True)
				End If
			End If
		ElseIf PointIn(AtX, AtY, 276, 62, 18, 267) Then 
			' ScrollBar Click
			'UPGRADE_ISSUE: picTomeNew was upgraded to a Panel, and cannot be coerced to a PictureBox. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="0FF1188E-27C0-4FED-842D-159C65894C9B"'
			If ScrollBarClick(AtX, AtY, ButtonDown, picTomeNew, 276, 62, 267, ScrollTop, ScrollList.Count(), 9) = True Then
				MenuSkillShow(CreateSkillsIndex)
			End If
		ElseIf PointIn(AtX, AtY, 26, 73, 276, 250) And ButtonDown = True Then 
			' Describe Skills
			c = Int((AtY - 73) / 25) + ScrollTop
			If c <= ScrollList.Count() Then
				'UPGRADE_WARNING: Couldn't resolve default property of object ScrollList().Index. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
				CombatAction1 = ScrollList.Item(c).Index
				'UPGRADE_WARNING: Couldn't resolve default property of object ScrollList2(). Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
				CombatAction2 = ScrollList2.Item(c)
				Call PlayClickSnd(modIOFunc.ClickType.ifClick)
				MenuSkillShow(c)
			End If
		ElseIf PointIn(AtX, AtY, 358, 325, 160, 18) And ButtonDown = True Then 
			' Click the Right-Click Action check box
			'UPGRADE_WARNING: Couldn't resolve default property of object ScrollList(CreateSkillsIndex).TriggerType. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
			Select Case ScrollList.Item(CreateSkillsIndex).TriggerType
				Case bdOnCastTarget, bdOnCast, bdOnSkillUse, bdOnSkillTarget
					' Set Right-Click Action or clear
					'UPGRADE_WARNING: Couldn't resolve default property of object ScrollList2(CreateSkillsIndex). Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
					'UPGRADE_WARNING: Couldn't resolve default property of object ScrollList(CreateSkillsIndex).Index. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
					If CreatureWithTurn.CombatAction1 = ScrollList.Item(CreateSkillsIndex).Index And CreatureWithTurn.CombatAction2 = ScrollList2.Item(CreateSkillsIndex) Then
						CreatureWithTurn.CombatAction1 = 0
						CreatureWithTurn.CombatAction2 = 0
					Else
						'UPGRADE_WARNING: Couldn't resolve default property of object ScrollList().Index. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
						CreatureWithTurn.CombatAction1 = ScrollList.Item(CreateSkillsIndex).Index
						'UPGRADE_WARNING: Couldn't resolve default property of object ScrollList2(). Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
						CreatureWithTurn.CombatAction2 = ScrollList2.Item(CreateSkillsIndex)
					End If
			End Select
			Call PlayClickSnd(modIOFunc.ClickType.ifClick)
			MenuSkillShow(CreateSkillsIndex)
		End If
	End Sub
	
	Private Sub MenuSkillLoad()
		Dim TriggerX, TriggerZ As Trigger
		' Load up the current creature's skills into our ScrollList.
		' ScrollList2 holds the Trigger's Index. TempSkill = -1 if the "Skill" is really a Spell.
		ScrollList = New Collection
		ScrollList2 = New Collection
		ScrollTop = 1
		For	Each TriggerX In CreatureWithTurn.Triggers
			If TriggerX.TriggerType = bdOnSkillUse Or TriggerX.TriggerType = bdOnSkillTarget Or TriggerX.IsSkill = True Then
				ScrollList.Add(TriggerX)
				ScrollList2.Add(0)
				TriggerX.TempSkill = TriggerX.SkillPoints
			End If
		Next TriggerX
		' Load up the current creature's spells
		For	Each TriggerX In CreatureWithTurn.Triggers
			' Show Creature based spells
			If TriggerX.TriggerType = bdOnCast Or TriggerX.TriggerType = bdOnCastTarget Then
				ScrollList.Add(TriggerX)
				ScrollList2.Add(0)
				TriggerX.TempSkill = 0
			End If
			' Show Skill based spells
			If TriggerX.IsSkill = True Then
				For	Each TriggerZ In TriggerX.Triggers
					If TriggerZ.TriggerType = bdOnCast Or TriggerZ.TriggerType = bdOnCastTarget Then
						ScrollList.Add(TriggerZ)
						ScrollList2.Add(TriggerX.Index)
						TriggerZ.TempSkill = 0
					End If
				Next TriggerZ
			End If
		Next TriggerX
		' Default to first in list
		If ScrollList.Count() > 0 Then
			'UPGRADE_WARNING: Couldn't resolve default property of object ScrollList().Index. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
			CombatAction1 = ScrollList.Item(1).Index
			'UPGRADE_WARNING: Couldn't resolve default property of object ScrollList2(). Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
			CombatAction2 = ScrollList2.Item(1)
		End If
	End Sub
	
	Private Sub MenuSkillUse(ByRef WithTarget As Short)
		Dim Found, Action1, Action2, c As Short
		Dim TriggerX, TriggerZ As Trigger
		Dim StatementX As Statement
		Const bdRuneStatement As Short = 23
		' Check if CreatureWithTurn is dead
		If CreatureWithTurn.AllowedTurn = False Then
			DialogDM(CreatureWithTurn.Name & " can't use any skills!")
			Exit Sub
		End If
		' Use either Right-Click Action or last selection
		If CreatureWithTurn.CombatAction1 > 0 And WithTarget = True Then
			If CreatureWithTurn.CombatAction2 > 0 Then
				Action1 = CreatureWithTurn.CombatAction2
				Action2 = CreatureWithTurn.CombatAction1
			Else
				Action1 = CreatureWithTurn.CombatAction1
				Action2 = CreatureWithTurn.CombatAction2
			End If
		ElseIf CombatAction1 > 0 And WithTarget = False Then 
			If CombatAction2 > 0 Then
				Action1 = CombatAction2
				Action2 = CombatAction1
			Else
				Action1 = CombatAction1
				Action2 = CombatAction2
			End If
		Else
			Action1 = 0
			Action2 = 0
		End If
		' If not doing default then execute Trigger
		Dim Runes(5) As Byte
		If Action1 > 0 Then
			Found = False
			For	Each TriggerX In CreatureWithTurn.Triggers
				If TriggerX.Index = Action1 Then
					Found = True
					Exit For
				End If
			Next TriggerX
			If Found = True Then
				If Action2 > 0 Then
					Found = False
					For	Each TriggerZ In TriggerX.Triggers
						If TriggerZ.Index = Action2 Then
							TriggerX = TriggerZ
							Found = True
							Exit For
						End If
					Next TriggerZ
				End If
				' Fire Skill if non-targeting type
				Select Case TriggerX.TriggerType
					Case bdOnSkillUse, bdOnSkillTarget
						CreatureNow = CreatureWithTurn
						ItemNow = CreatureWithTurn.ItemInHand
						FireTrigger(CreatureWithTurn, TriggerX)
					Case bdOnCast, bdOnCastTarget
						' Acquire Runes to match
						Found = False
						For	Each StatementX In TriggerX.Statements
							If StatementX.Statement = bdRuneStatement Then
								For c = 0 To 5
									Runes(c) = StatementX.B(c)
								Next c
								Found = True
								Exit For
							End If
						Next StatementX
						' If found Runes, then invoke it
						If Found = True Then
							' If Targeting, check for Saving Throw
							If TriggerX.TriggerType = bdOnCastTarget And WithTarget = False Then
								' Set GlobalSaveStyle value if the Trigger has the Rune Statement with Save Checked
								If (StatementX.B(6) And &H2) > 0 Then
									GlobalSaveStyle = TriggerX.Styles
								Else
									GlobalSaveStyle = 0
								End If
								Found = Not TargetCreatureSave(CreatureTarget)
							End If
							If Found = True Then
								' If have Runes, then Cast it. Else declare no match.
								If SorceryMatchRunes(CreatureWithTurn, Runes, False) = 0 Then
									SorceryMatchRunes(CreatureWithTurn, Runes, True)
								Else
									MessageShow("No Runes!", 0)
								End If
							End If
						Else
							' This is a non-Rune Spell Trigger
							CreatureNow = CreatureWithTurn
							ItemNow = CreatureWithTurn.ItemInHand
							FireTrigger(CreatureWithTurn, TriggerX)
						End If
				End Select
			Else
				' Previous default no longer exists
				DialogDM("No Skill Selected! Click the Skill button to choose a Skill.")
			End If
		Else
			' Do default action (none was selected)
		End If
		' Clear last selection
		CombatAction1 = 0
		CombatAction2 = 0
		' If out of Action Points in Combat, next turn
		If picGrid.Visible = True And CreatureWithTurn.ActionPoints < 1 Then
			CombatNextTurn()
		End If
	End Sub
	
	Private Sub MenuSkillShow(ByRef Index As Short)
		Dim i, c, n, SkillLevel As Short
		Dim SkillLevelName As String
		Dim StatementX As Statement
		Dim TriggerX As Trigger
		Dim rc As Integer
		Const bdRuneStatement As Short = 23
		Frozen = True
		'UPGRADE_ISSUE: PictureBox method picTomeNew.Cls was not upgraded. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="CC4C7EC0-C903-48FC-ACCC-81861D12DA4A"'
		picTomeNew.Cls()
		TomeAction = bdMenuSkill
		CreateSkillsIndex = Index
		'UPGRADE_ISSUE: PictureBox method picTomeNew.Cls was not upgraded. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="CC4C7EC0-C903-48FC-ACCC-81861D12DA4A"'
		picTomeNew.Cls()
		ShowText(picTomeNew, 30, 12, 519, 14, bdFontElixirWhite, CreatureWithTurn.Name & " Skills", True, False)
		ShowText(picTomeNew, 28, 45, 255, 14, bdFontElixirWhite, "Skill Points " & CreatureWithTurn.SkillPoints, False, False)
		ShowText(picTomeNew, 212, 45, 64, 14, bdFontElixirWhite, "Level", True, False)
		' List available and selected skills
		n = 0
		For c = 1 To ScrollList.Count()
			' Show selected Skill
			If c = Index Then
				'UPGRADE_WARNING: Couldn't resolve default property of object ScrollList().Name. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
				ShowText(picTomeNew, 332, 48, 217, 15, bdFontElixirBlack, ScrollList.Item(c).Name, True, False)
				'UPGRADE_WARNING: Couldn't resolve default property of object ScrollList().Comments. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
				ShowText(picTomeNew, 328, 80, 220, 190, bdFontNoxiousBlack, ScrollList.Item(c).Comments, False, False)
				'UPGRADE_WARNING: Couldn't resolve default property of object ScrollList(c).IsSkill. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
				If ScrollList.Item(c).IsSkill = True Then
					'UPGRADE_WARNING: Couldn't resolve default property of object ScrollList().Turns. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
					ShowText(picTomeNew, 328, 300, 220, 14, bdFontNoxiousBlack, "Cost: " & ScrollList.Item(c).Turns, True, False)
				End If
				' Show Right-Click Action check box
				'UPGRADE_WARNING: Couldn't resolve default property of object ScrollList(c).TriggerType. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
				Select Case ScrollList.Item(c).TriggerType
					Case bdOnCast, bdOnSkillUse, bdOnCastTarget, bdOnSkillTarget
						ShowText(picTomeNew, 358, 328, 220, 18, bdFontNoxiousBlack, "Right-Click Action?", False, False)
						'UPGRADE_WARNING: Couldn't resolve default property of object ScrollList2(c). Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
						'UPGRADE_WARNING: Couldn't resolve default property of object ScrollList(c).Index. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
						If CreatureWithTurn.CombatAction1 = ScrollList.Item(c).Index And CreatureWithTurn.CombatAction2 = ScrollList2.Item(c) Then
							'UPGRADE_ISSUE: PictureBox property picMisc.hdc was not upgraded. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="CC4C7EC0-C903-48FC-ACCC-81861D12DA4A"'
							'UPGRADE_ISSUE: PictureBox property picTomeNew.hdc was not upgraded. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="CC4C7EC0-C903-48FC-ACCC-81861D12DA4A"'
							rc = BitBlt(picTomeNew.hdc, 500, 325, 18, 18, picMisc.hdc, 18, 18, SRCCOPY)
						Else
							'UPGRADE_ISSUE: PictureBox property picMisc.hdc was not upgraded. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="CC4C7EC0-C903-48FC-ACCC-81861D12DA4A"'
							'UPGRADE_ISSUE: PictureBox property picTomeNew.hdc was not upgraded. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="CC4C7EC0-C903-48FC-ACCC-81861D12DA4A"'
							rc = BitBlt(picTomeNew.hdc, 500, 325, 18, 18, picMisc.hdc, 0, 18, SRCCOPY)
						End If
						ShowButton(picTomeNew, picTomeNew.ClientRectangle.Width - 198, 364, "Use", False)
					Case Else
						ShowText(picTomeNew, 332, 328, 217, 18, bdFontNoxiousBlack, "Always On", True, False)
				End Select
				' Show Runes (if a Spell)
				'UPGRADE_WARNING: Couldn't resolve default property of object ScrollList(c).TriggerType. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
				If ScrollList.Item(c).TriggerType = bdOnCast Or ScrollList.Item(c).TriggerType = bdOnCastTarget Then
					' Find a Rune Statement (if exists)
					'UPGRADE_WARNING: Couldn't resolve default property of object ScrollList(c).Statements. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
					For	Each StatementX In ScrollList.Item(c).Statements
						If StatementX.Statement = bdRuneStatement Then
							For i = 0 To 5
								If StatementX.B(i) > 0 Then
									' If available in Map.Runes, show it. Else, show it grey
									'UPGRADE_WARNING: Untranslated statement in MenuSkillShow. Please check source code.
								End If
							Next i
							Exit For
						End If
					Next StatementX
					' If out of Action Points, say so
					If CreatureWithTurn.ActionPoints < 10 Then
						ShowText(picTomeNew, 328, 308, 220, 15, bdFontNoxiousBlack, "Not Enough Action Points!", True, False)
					End If
				Else
					' Show Raise/Lower if applicable
					'UPGRADE_WARNING: Couldn't resolve default property of object ScrollList(c).TempSkill. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
					'UPGRADE_WARNING: Couldn't resolve default property of object ScrollList(c).SkillPoints. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
					'UPGRADE_WARNING: Couldn't resolve default property of object ScrollList(c).Turns. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
					If CreatureWithTurn.SkillPoints >= ScrollList.Item(c).Turns Or ScrollList.Item(c).SkillPoints <> ScrollList.Item(c).TempSkill Then
						ShowButton(picTomeNew, 60, 331, "Lower", False)
						ShowButton(picTomeNew, 156, 331, "Raise", False)
					End If
				End If
			End If
			If IsBetween(c, ScrollTop, Least(ScrollTop + 9, ScrollList.Count())) Then
				' Set Skill Level
				'UPGRADE_WARNING: Couldn't resolve default property of object ScrollList(c).Turns. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
				If ScrollList.Item(c).Turns > 0 Then
					'UPGRADE_WARNING: Couldn't resolve default property of object ScrollList(c).Turns. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
					'UPGRADE_WARNING: Couldn't resolve default property of object ScrollList().SkillPoints. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
					SkillLevel = Int(ScrollList.Item(c).SkillPoints / ScrollList.Item(c).Turns)
					Select Case SkillLevel
						Case 0, 1 : SkillLevelName = "Novice"
						Case 2 : SkillLevelName = "Adept"
						Case 3 : SkillLevelName = "Expert"
						Case 4 : SkillLevelName = "Master"
						Case Else : SkillLevelName = "Master+" & SkillLevel - 4
					End Select
					'UPGRADE_WARNING: Couldn't resolve default property of object ScrollList(c).TriggerType. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
				ElseIf ScrollList.Item(c).TriggerType = bdOnCast Or ScrollList.Item(c).TriggerType = bdOnCastTarget Then 
					SkillLevelName = "*"
				Else
					SkillLevelName = "Weakness"
				End If
				If c = Index Then
					'UPGRADE_WARNING: Couldn't resolve default property of object ScrollList().Name. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
					ShowText(picTomeNew, 26, 76 + n * 25, 240, 14, bdFontNoxiousGold, ScrollList.Item(c).Name, False, False)
					ShowText(picTomeNew, 26, 76 + n * 25, 240, 14, bdFontNoxiousGold, SkillLevelName, 1, False)
				Else
					'UPGRADE_WARNING: Couldn't resolve default property of object ScrollList().Name. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
					ShowText(picTomeNew, 26, 76 + n * 25, 240, 14, bdFontNoxiousWhite, ScrollList.Item(c).Name, False, False)
					ShowText(picTomeNew, 26, 76 + n * 25, 240, 14, bdFontNoxiousWhite, SkillLevelName, 1, False)
				End If
				n = n + 1
			End If
		Next c
		' Set up buttons
		'UPGRADE_ISSUE: picTomeNew was upgraded to a Panel, and cannot be coerced to a PictureBox. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="0FF1188E-27C0-4FED-842D-159C65894C9B"'
		ScrollBarShow(picTomeNew, 276, 62, 267, ScrollTop, ScrollList.Count() - 9, 0)
		ShowButton(picTomeNew, picTomeNew.ClientRectangle.Width - 102, 364, "Done", False)
		' Refresh the screen
		picTomeNew.Refresh()
		picTomeNew.BringToFront()
		picTomeNew.Top = Me.ClientRectangle.Height - picMenu.ClientRectangle.Height - picTomeNew.ClientRectangle.Height - 48
		picTomeNew.Left = Me.ClientRectangle.Width - picTomeNew.ClientRectangle.Width
		picTomeNew.Visible = True
	End Sub
	
	'UPGRADE_NOTE: Tome was upgraded to Tome_Renamed. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="A9E4979A-37FA-4718-9994-97DD76ED70A7"'
	Private Sub MenuAction(ByRef X As Short, ByRef Y As Short, ByRef IsDown As Short)
		Dim Tome_Renamed As Object
		Dim c, Button As Short
		' If Frozen or not in game, then exit
		If Frozen = True Or picMap.Visible = False Then
			Exit Sub
		End If
		' Else you can try to click
		c = Me.ClientRectangle.Height - picMenu.ClientRectangle.Height - bdIntHeight * 3 + 7
		If IsBetween(Y, c, c + 32) Then
			Select Case X
				Case 0 To 24 ' Menu
					Button = 7
				Case 40 To 130 ' Inventory
					Button = 1
				Case 138 To 228 ' Talk
					Button = 2
				Case 236 To 326 ' Search/Defend
					Button = 3
				Case 334 To 424 ' Listen/Wait
					Button = 4
				Case 432 To 522 ' Open/Flee
					Button = 5
				Case Me.ClientRectangle.Width - 130 To Me.ClientRectangle.Width - 40 ' Skills
					Button = 6
				Case Me.ClientRectangle.Width - 24 To Me.ClientRectangle.Width ' Journal
					Button = 8
				Case Else
					Button = 0
			End Select
			If IsDown = True Then
				BorderDrawButtons(Button)
				Me.Refresh()
			ElseIf Button > 0 Then 
				BorderDrawButtons(0)
				Me.Refresh()
				Call PlayClickSnd(modIOFunc.ClickType.ifClick)
				Select Case Button
					Case 1 ' Inventory
						InvTopIndex = 0
						InventoryShow(InvNowShow)
						picInventory.Visible = True
						picInventory.BringToFront()
						picInventory.Focus()
					Case 2 ' Talk
						TalkSetup()
					Case 3 ' Search/Defend
						If picGrid.Visible = True Then
							CombatDefend()
						Else
							Search()
						End If
					Case 4 ' Listen/Wait
						If picGrid.Visible = True Then
							CombatWait()
						Else
							DoorListen()
						End If
					Case 5 ' Open/Flee
						If picGrid.Visible = True Then
							CombatFlee()
						Else
							'UPGRADE_WARNING: Couldn't resolve default property of object Tome.MapY. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
							'UPGRADE_WARNING: Couldn't resolve default property of object Tome.MapX. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
							DoorOpen(Tome.MapX, Tome.MapY)
						End If
					Case 6 ' Skills
						If CreatureWithTurn.HPNow > 0 Then
							MenuSkillLoad()
							MenuSkillShow(1)
						Else
							DialogDM(CreatureWithTurn.Name & " is dead.")
						End If
					Case 7 ' Menu
						MenuShow()
					Case 8 ' Journal
						Select Case JournalMode
							Case 1 ' Quests
								JournalQuestsLoad()
							Case 2 ' Map
							Case Else ' Journal Entries
								JournalEntryLoad()
						End Select
						JournalShow()
				End Select
			End If
		Else
			BorderDrawButtons(0)
			Me.Refresh()
		End If
	End Sub
	
	Public Sub DoAction(ByRef Action As Short)
		Dim c, ButtonCnt As Short
		Dim TriggerX As Trigger
		Dim SkipTurn As Short
		SkipTurn = False
		Select Case MenuNow
			Case bdMenuInventory
				If Action = bdMenuActionDefault Then
					InvTopIndex = 0
					InventoryShow(bdInvItems)
					picInventory.Visible = True
					picInventory.BringToFront()
					picInventory.Focus()
				End If
			Case bdMenuSearch
				Search()
			Case bdMenuEncounter
				For c = 0 To 4
					If ButtonList(c) <> "" Then
						ButtonCnt = c
					End If
				Next c
				Select Case ButtonList(ButtonCnt - Action)
					Case "Talk"
						TalkSetup()
					Case "Flee"
						CombatFlee()
					Case "Fight"
						CombatStart()
					Case "Ignore"
						EncounterIgnore()
					Case "Done"
				End Select
		End Select
		' Show action menu if count is there
		If CreatureWithTurn.ActionPoints < 1 Or SkipTurn = True Then
			' If in combat then take a turn
			If picGrid.Visible = True And Frozen = False Then
				CombatNextTurn()
			End If
			' If outside of combat, count as a step
			If picGrid.Visible = False Then
				TotalStepsTaken = TotalStepsTaken + 1
				If TotalStepsTaken > 9 Then
					Do Until TotalStepsTaken < 10
						TurnCycle()
						TotalStepsTaken = TotalStepsTaken - 10
					Loop 
				End If
			End If
		End If
	End Sub
	
	Private Sub MenuShow()
		Dim rc As Integer
		Frozen = True
		'UPGRADE_ISSUE: PictureBox method picMainMenu.Cls was not upgraded. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="CC4C7EC0-C903-48FC-ACCC-81861D12DA4A"'
		picMainMenu.Cls()
		TomeMenu = bdTomeOptions
		'UPGRADE_ISSUE: PictureBox property picBlack.hdc was not upgraded. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="CC4C7EC0-C903-48FC-ACCC-81861D12DA4A"'
		'UPGRADE_ISSUE: PictureBox property picMainMenu.hdc was not upgraded. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="CC4C7EC0-C903-48FC-ACCC-81861D12DA4A"'
		rc = BitBlt(picMainMenu.hdc, 122, 103, 288, 195, picBlack.hdc, 0, 340, SRCCOPY)
		ShowText(picMainMenu, 30, 12, 519, 14, bdFontElixirWhite, "Show which game World is active.", True, False)
		picMainMenu.Visible = True
		picMainMenu.BringToFront()
		TomeAction = 1
		picMainMenu_MouseMove(picMainMenu, New System.Windows.Forms.MouseEventArgs(0 * &H100000, 0, 0, 128, 0))
	End Sub
	
	Private Sub OptionsShow(ByRef Index As Short)
		Dim n, c, Status As Short
		Dim rc As Integer
		'UPGRADE_NOTE: Text was upgraded to Text_Renamed. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="A9E4979A-37FA-4718-9994-97DD76ED70A7"'
		Dim Text_Renamed, Explain As String
		Dim WorldTemp As World
		Dim k As Short
		TomeAction = bdTomeOptions
		ScrollSelect = Index
		' Center picTome picture
		picTomeNew.Top = (Me.ClientRectangle.Height - picTomeNew.ClientRectangle.Height) / 2
		picTomeNew.Left = (Me.ClientRectangle.Width - picTomeNew.ClientRectangle.Width) / 2
		'UPGRADE_ISSUE: PictureBox method picTomeNew.Cls was not upgraded. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="CC4C7EC0-C903-48FC-ACCC-81861D12DA4A"'
		picTomeNew.Cls()
		ShowText(picTomeNew, 30, 12, 519, 14, bdFontElixirWhite, "Set Game Options", True, False)
		ShowText(picTomeNew, 19, 45, 273, 14, bdFontElixirWhite, "Options", True, False)
		n = 0
		If TomeMenu = bdTomeOptions Then
			k = bdOptionsCount - 2
		Else
			k = bdOptionsCount
		End If
		For c = 1 To k
			Select Case c
				Case 1 ' Music
					Text_Renamed = "Play Music"
					Explain = "Turn this on if you want to hear music play during the adventure and combat. The default mode is ON." & Chr(13) & Chr(13) & "Music Playing Now: " & Chr(13) & GlobalMusicName
					Status = GlobalMusicState
				Case 2 ' Random Music
					Text_Renamed = "Play Random Music"
					Explain = "Turn this on if you want the system to randomly pick music to play during adventure and combat. The default mode is ON."
					Status = GlobalMusicRandom
				Case 3 ' Right-Click Menu
					Text_Renamed = "Right-Click Skills"
					Explain = "Turn this on if you want to be able to right-click and hold to select an option. OFF means you have to right click, release and then choose an option. The default mode is ON."
					Status = GlobalRightClick
				Case 4 ' Game Speed
					Text_Renamed = "Set Game Speed"
					Explain = "Click this option to cycle through the game speeds. Slow is 1, Fast is 20. The default speed is 12." & Chr(13) & Chr(13) & "Game Speed now is " & GlobalGameSpeed
					Status = GlobalGameSpeed
				Case 5 ' Dice Rolling
					Text_Renamed = "Show Dice Rolling"
					Explain = "Turn this option on if you want to see dice rolling during combat. Turning it off will speed up combat slightly. Either way you'll still see the resulting amount of damage. The default is ON."
					Status = GlobalDiceRolling
				Case 6 ' Sound Effects
					Text_Renamed = "Play Sound Effects"
					Explain = "Turn this on to hear sound effects in combat and while adventuring. This does not effect the audible click when using the mouse. The default is ON."
					Status = GlobalWAVState
				Case 7 ' Debug Mode
					Text_Renamed = "Use Debug Mode"
					Explain = "Turn this on if you want to see the Trigger Statements right after they fire. This helps in debugging what is happening in your Tome. The default is OFF."
					Status = GlobalDebugMode
				Case 8 ' Overswing
					Text_Renamed = "Over swing Penalty"
					Explain = "Turning this on means your characters get a parting shot in combat which ends their turn even if they don't have enough Action Points. There is a penalty to their ToHit roll, however. The default is ON."
					Status = GlobalOverSwing
				Case 9 ' Mouse Click
					Text_Renamed = "Play MouseClick Sound"
					Explain = "Turning this on means you'll hear a sound when you click the mouse. The default is ON."
					Status = GlobalMouseClick
				Case 10 ' Auto End Turn
					Text_Renamed = "Auto End Turn"
					Explain = "Turning this on means a Character's turn will automatically end when there is no one left in weapon range. The default is OFF."
					Status = GlobalAutoEndTurn
				Case 11 ' Fast Move
					Text_Renamed = "Fast Move"
					Explain = "Fast Move ON means the party on the map will move faster than normal (skipping frames of animation). Default is OFF."
					Status = GlobalFastMove
				Case 12 ' Auto Center Map
					Text_Renamed = "Auto Center Map"
					Explain = "Auto Center Map ON means the map will be centered on the party after their move. Default is OFF."
					Status = GlobalAutoCenter
				Case 13 ' Interface
					Text_Renamed = "Interface Settings"
					Explain = "You can change the look of the interface to one of several selections. Default Interface is WOOD WEAVE." & Chr(13) & Chr(13) & "Interface: " & UCase(GlobalInterfaceName)
					Status = 0
				Case 14 ' Dice Set
					Text_Renamed = "Dice Set"
					'Explain = "You can change the look of the dice to one of several selections. Default Interface is BONE WHITE." & Chr(13) & Chr(13) & "Dice Set: " & UCase$(Left$(GlobalDiceName, Len(GlobalDiceName) - 4))
					Explain = "You can change the look of the dice to one of several selections. Default Interface is BONE WHITE." & Chr(13) & Chr(13) & "Dice Set: " & VB.Left(GlobalDiceName, Len(GlobalDiceName) - 4)
					Status = 0
				Case 15 ' change keyboard shortcuts [Titi 2.4.9]
					Text_Renamed = "Keyboard Shortcuts"
					Explain = "You can change the assignment of the keys used as shortcuts: "
				Case 16 ' Change World
					Text_Renamed = "Change World"
					'UPGRADE_WARNING: Couldn't resolve default property of object Worlds(WorldIndex).Name. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
					Explain = "You can change entire worlds settings. " & Chr(13) & "Current: " & Worlds.Item(WorldIndex).Name
					Status = 0
				Case 17 ' toggle splash and credits screen
					Text_Renamed = "Toggle Splash screen"
					'Explain = "Display splash screens on startup and credits on closing. Default is OFF."
					Explain = "Display splash screens on startup. Default is OFF."
					If Val(GlobalCredits) = 1 Then
						Status = 1
					Else
						Status = 0
					End If
			End Select
			If c = Index Then
				' Show text
				ShowText(picTomeNew, 332, 48, 217, 15, bdFontElixirBlack, Text_Renamed, True, False)
				ShowText(picTomeNew, 328, 80, 220, 190, bdFontNoxiousBlack, Explain, False, False)
				'            If Index = 12 Or Index = 13 Or Index = 14 Then
				If Index >= 13 And Index <= 16 Then
					'                ShowButton picTomeNew, 340, 200, "Back", False
					'                ShowButton picTomeNew, 435, 200, "Next", False
					' [Titi 2.4.9] moved to the bottom to give more space to show the world's details
					ShowButton(picTomeNew, 340, 298, "Back", False)
					ShowButton(picTomeNew, 435, 298, "Next", False)
				End If
			End If
			If IsBetween(c, ScrollTop, Least(ScrollTop + 9, bdOptionsCount)) Then
				If c = Index Then
					ShowText(picTomeNew, 53, 76 + n * 25, 176, 14, bdFontNoxiousGold, Text_Renamed, False, False)
				Else
					ShowText(picTomeNew, 53, 76 + n * 25, 176, 14, bdFontNoxiousWhite, Text_Renamed, False, False)
				End If
				' Show clicked if ON
				If Status > 0 Then
					'UPGRADE_ISSUE: PictureBox property picMisc.hdc was not upgraded. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="CC4C7EC0-C903-48FC-ACCC-81861D12DA4A"'
					'UPGRADE_ISSUE: PictureBox property picTomeNew.hdc was not upgraded. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="CC4C7EC0-C903-48FC-ACCC-81861D12DA4A"'
					rc = BitBlt(picTomeNew.hdc, 26, 73 + n * 25, 18, 18, picMisc.hdc, 18, 18, SRCCOPY)
				Else
					'UPGRADE_ISSUE: PictureBox property picMisc.hdc was not upgraded. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="CC4C7EC0-C903-48FC-ACCC-81861D12DA4A"'
					'UPGRADE_ISSUE: PictureBox property picTomeNew.hdc was not upgraded. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="CC4C7EC0-C903-48FC-ACCC-81861D12DA4A"'
					rc = BitBlt(picTomeNew.hdc, 26, 73 + n * 25, 18, 18, picMisc.hdc, 0, 18, SRCCOPY)
				End If
				n = n + 1
			End If
		Next c
		Dim FileName As String
		Dim X2, Y2 As Short
		Dim Size2 As Single
		Dim blnShow2 As Boolean
		Dim X1, Y1 As Short
		Dim Size1 As Single
		Dim blnShow1 As Boolean
		If Index = 14 Then ' [Titi 2.4.9] show an example of the dice set
			'UPGRADE_ISSUE: PictureBox property picDice.hdc was not upgraded. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="CC4C7EC0-C903-48FC-ACCC-81861D12DA4A"'
			'UPGRADE_ISSUE: PictureBox property picTomeNew.hdc was not upgraded. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="CC4C7EC0-C903-48FC-ACCC-81861D12DA4A"'
			rc = BitBlt(picTomeNew.hdc, 365, 210, 40, 40, picDice.hdc, picDice.ClientRectangle.Width / 2, 0, SRCAND)
			'UPGRADE_ISSUE: PictureBox property picDice.hdc was not upgraded. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="CC4C7EC0-C903-48FC-ACCC-81861D12DA4A"'
			'UPGRADE_ISSUE: PictureBox property picTomeNew.hdc was not upgraded. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="CC4C7EC0-C903-48FC-ACCC-81861D12DA4A"'
			rc = BitBlt(picTomeNew.hdc, 365, 210, 40, 40, picDice.hdc, 0, 0, SRCPAINT)
			'UPGRADE_ISSUE: PictureBox property picDice.hdc was not upgraded. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="CC4C7EC0-C903-48FC-ACCC-81861D12DA4A"'
			'UPGRADE_ISSUE: PictureBox property picTomeNew.hdc was not upgraded. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="CC4C7EC0-C903-48FC-ACCC-81861D12DA4A"'
			rc = BitBlt(picTomeNew.hdc, 460, 210, 40, 40, picDice.hdc, picDice.ClientRectangle.Width / 2, 200, SRCAND)
			'UPGRADE_ISSUE: PictureBox property picDice.hdc was not upgraded. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="CC4C7EC0-C903-48FC-ACCC-81861D12DA4A"'
			'UPGRADE_ISSUE: PictureBox property picTomeNew.hdc was not upgraded. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="CC4C7EC0-C903-48FC-ACCC-81861D12DA4A"'
			rc = BitBlt(picTomeNew.hdc, 460, 210, 40, 40, picDice.hdc, 0, 200, SRCPAINT)
			'        rc = BitBlt(picTomeNew.hdc, 365, 230, picDice.Width / 16, picDice.Height / 9, picDice.hdc, 0, 0, SRCCOPY)
			'        rc = BitBlt(picTomeNew.hdc, 460, 230, picDice.Width / 16, picDice.Height / 9, picDice.hdc, 0, picDice.Height / 9 * 5, SRCCOPY)
		ElseIf Index = 15 Then  ' [Titi 2.4.9] display current shortcuts keys
			Select Case KingdomIndex
				Case 0
					ShowText(picTomeNew, 328, 143, 200, 15, bdFontNoxiousBlack, "In & Out keys", False, False)
					ShowText(picTomeNew, 328, 150, 200, 15, bdFontNoxiousBlack, "---------------------", False, False)
					ShowText(picTomeNew, 328, 169, 80, 15, bdFontNoxiousBlack, "Debug:", False, False)
					ShowText(picTomeNew, 398, 169, 80, 15, IIf(Len(ShowShortcuts(bdKey(1))) > 2, bdFontSmallWhite, bdFontNoxiousBlack), ShowShortcuts(bdKey(1)), False, False)
					ShowText(picTomeNew, 444, 169, 80, 15, bdFontNoxiousBlack, "Quit:", False, False)
					ShowText(picTomeNew, 510, 169, 80, 15, IIf(Len(ShowShortcuts(bdKey(2))) > 2, bdFontSmallWhite, bdFontNoxiousBlack), ShowShortcuts(bdKey(2)), False, False)
					ShowText(picTomeNew, 328, 186, 80, 15, bdFontNoxiousBlack, "Exit:", False, False)
					ShowText(picTomeNew, 398, 186, 80, 15, IIf(Len(ShowShortcuts(bdKey(3))) > 2, bdFontSmallWhite, bdFontNoxiousBlack), ShowShortcuts(bdKey(3)), False, False)
					ShowText(picTomeNew, 444, 186, 80, 15, bdFontNoxiousBlack, "Speed:", False, False)
					ShowText(picTomeNew, 510, 186, 80, 15, IIf(Len(ShowShortcuts(bdKey(4))) > 2, bdFontSmallWhite, bdFontNoxiousBlack), ShowShortcuts(bdKey(4)), False, False)
					ShowText(picTomeNew, 328, 203, 80, 15, bdFontNoxiousBlack, "Load:", False, False)
					ShowText(picTomeNew, 398, 203, 80, 15, IIf(Len(ShowShortcuts(bdKey(5))) > 2, bdFontSmallWhite, bdFontNoxiousBlack), ShowShortcuts(bdKey(5)), False, False)
					ShowText(picTomeNew, 444, 203, 80, 15, bdFontNoxiousBlack, "Save:", False, False)
					ShowText(picTomeNew, 510, 203, 80, 15, IIf(Len(ShowShortcuts(bdKey(6))) > 2, bdFontSmallWhite, bdFontNoxiousBlack), ShowShortcuts(bdKey(6)), False, False)
					ShowText(picTomeNew, 328, 220, 80, 15, bdFontNoxiousBlack, "Music:", False, False)
					ShowText(picTomeNew, 398, 220, 80, 15, IIf(Len(ShowShortcuts(bdKey(7))) > 2, bdFontSmallWhite, bdFontNoxiousBlack), ShowShortcuts(bdKey(7)), False, False)
					ShowText(picTomeNew, 444, 220, 80, 15, bdFontNoxiousBlack, "Sound:", False, False)
					ShowText(picTomeNew, 510, 220, 80, 15, IIf(Len(ShowShortcuts(bdKey(8))) > 2, bdFontSmallWhite, bdFontNoxiousBlack), ShowShortcuts(bdKey(8)), False, False)
				Case 1
					ShowText(picTomeNew, 328, 143, 200, 15, bdFontNoxiousBlack, "Map Scrolling", False, False)
					ShowText(picTomeNew, 328, 150, 200, 15, bdFontNoxiousBlack, "-------------------", False, False)
					ShowText(picTomeNew, 328, 169, 80, 15, bdFontNoxiousBlack, "Left:", False, False)
					ShowText(picTomeNew, 398, 169, 80, 15, IIf(Len(ShowShortcuts(bdKey(9))) > 2, bdFontSmallWhite, bdFontNoxiousBlack), ShowShortcuts(bdKey(9)), False, False)
					ShowText(picTomeNew, 444, 169, 80, 15, bdFontNoxiousBlack, "Right:", False, False)
					ShowText(picTomeNew, 510, 169, 80, 15, IIf(Len(ShowShortcuts(bdKey(10))) > 2, bdFontSmallWhite, bdFontNoxiousBlack), ShowShortcuts(bdKey(10)), False, False)
					ShowText(picTomeNew, 328, 186, 80, 15, bdFontNoxiousBlack, "Up:", False, False)
					ShowText(picTomeNew, 398, 186, 80, 15, IIf(Len(ShowShortcuts(bdKey(11))) > 2, bdFontSmallWhite, bdFontNoxiousBlack), ShowShortcuts(bdKey(11)), False, False)
					ShowText(picTomeNew, 444, 186, 80, 15, bdFontNoxiousBlack, "Down:", False, False)
					ShowText(picTomeNew, 510, 186, 80, 15, IIf(Len(ShowShortcuts(bdKey(12))) > 2, bdFontSmallWhite, bdFontNoxiousBlack), ShowShortcuts(bdKey(12)), False, False)
				Case 2
					ShowText(picTomeNew, 328, 143, 200, 15, bdFontNoxiousBlack, "Party Actions", False, False)
					ShowText(picTomeNew, 328, 150, 200, 15, bdFontNoxiousBlack, "-------------------", False, False)
					ShowText(picTomeNew, 328, 169, 80, 15, bdFontNoxiousBlack, "Left:", False, False)
					ShowText(picTomeNew, 398, 169, 80, 15, IIf(Len(ShowShortcuts(bdKey(13))) > 2, bdFontSmallWhite, bdFontNoxiousBlack), ShowShortcuts(bdKey(13)), False, False)
					ShowText(picTomeNew, 444, 169, 80, 15, bdFontNoxiousBlack, "Right:", False, False)
					ShowText(picTomeNew, 510, 169, 80, 15, IIf(Len(ShowShortcuts(bdKey(14))) > 2, bdFontSmallWhite, bdFontNoxiousBlack), ShowShortcuts(bdKey(14)), False, False)
					ShowText(picTomeNew, 328, 186, 80, 15, bdFontNoxiousBlack, "Up:", False, False)
					ShowText(picTomeNew, 398, 186, 80, 15, IIf(Len(ShowShortcuts(bdKey(15))) > 2, bdFontSmallWhite, bdFontNoxiousBlack), ShowShortcuts(bdKey(15)), False, False)
					ShowText(picTomeNew, 444, 186, 80, 15, bdFontNoxiousBlack, "Down:", False, False)
					ShowText(picTomeNew, 510, 186, 80, 15, IIf(Len(ShowShortcuts(bdKey(16))) > 2, bdFontSmallWhite, bdFontNoxiousBlack), ShowShortcuts(bdKey(16)), False, False)
					ShowText(picTomeNew, 328, 211, 160, 15, bdFontNoxiousBlack, "Journal & Quests:", False, False)
					ShowText(picTomeNew, 510, 211, 80, 15, IIf(Len(ShowShortcuts(bdKey(17))) > 2, bdFontSmallWhite, bdFontNoxiousBlack), ShowShortcuts(bdKey(17)), False, False)
					ShowText(picTomeNew, 328, 228, 80, 15, bdFontNoxiousBlack, "Listen:", False, False)
					ShowText(picTomeNew, 398, 228, 80, 15, IIf(Len(ShowShortcuts(bdKey(18))) > 2, bdFontSmallWhite, bdFontNoxiousBlack), ShowShortcuts(bdKey(18)), False, False)
					ShowText(picTomeNew, 444, 228, 80, 15, bdFontNoxiousBlack, "Talk:", False, False)
					ShowText(picTomeNew, 510, 228, 80, 15, IIf(Len(ShowShortcuts(bdKey(19))) > 2, bdFontSmallWhite, bdFontNoxiousBlack), ShowShortcuts(bdKey(19)), False, False)
					ShowText(picTomeNew, 328, 245, 80, 15, bdFontNoxiousBlack, "Skills:", False, False)
					ShowText(picTomeNew, 398, 245, 80, 15, IIf(Len(ShowShortcuts(bdKey(20))) > 2, bdFontSmallWhite, bdFontNoxiousBlack), ShowShortcuts(bdKey(20)), False, False)
					ShowText(picTomeNew, 444, 245, 80, 15, bdFontNoxiousBlack, "Status:", False, False)
					ShowText(picTomeNew, 510, 245, 80, 15, IIf(Len(ShowShortcuts(bdKey(21))) > 2, bdFontSmallWhite, bdFontNoxiousBlack), ShowShortcuts(bdKey(21)), False, False)
					ShowText(picTomeNew, 328, 262, 80, 15, bdFontNoxiousBlack, "Search:", False, False)
					ShowText(picTomeNew, 398, 262, 80, 15, IIf(Len(ShowShortcuts(bdKey(22))) > 2, bdFontSmallWhite, bdFontNoxiousBlack), ShowShortcuts(bdKey(22)), False, False)
					ShowText(picTomeNew, 444, 262, 80, 15, bdFontNoxiousBlack, "Open:", False, False)
					ShowText(picTomeNew, 510, 262, 80, 15, IIf(Len(ShowShortcuts(bdKey(23))) > 2, bdFontSmallWhite, bdFontNoxiousBlack), ShowShortcuts(bdKey(23)), False, False)
					ShowText(picTomeNew, 328, 279, 80, 15, bdFontNoxiousBlack, "Inventory:", False, False)
					ShowText(picTomeNew, 398, 279, 80, 15, IIf(Len(ShowShortcuts(bdKey(24))) > 2, bdFontSmallWhite, bdFontNoxiousBlack), ShowShortcuts(bdKey(24)), False, False)
					ShowText(picTomeNew, 444, 279, 80, 15, bdFontNoxiousBlack, "Equip:", False, False)
					ShowText(picTomeNew, 510, 279, 80, 15, IIf(Len(ShowShortcuts(bdKey(25))) > 2, bdFontSmallWhite, bdFontNoxiousBlack), ShowShortcuts(bdKey(25)), False, False)
			End Select
		ElseIf Index = 16 Then  ' [Titi 2.4.9] display world details
			blnShow1 = True : blnShow2 = True
			WorldTemp = Worlds.Item(WorldIndex)
			' Call CreatePCLoadWorlds(WorldTemp) too long delay!
			FileName = gAppPath & "\roster\" & WorldTemp.Name & "\" & WorldTemp.Name & ".ini"
			' [Titi 2.4.9] Get the World Description from the *.ini file
			rc = fReadValue(FileName, "World", "DescriptionLines", "S", "1", Text_Renamed)
			X1 = Val(Text_Renamed)
			rc = fReadValue(FileName, "World", "Description" & Trim(Str(ScrollWorldDesc + 1)), "S", "No description available", Text_Renamed)
			WorldTemp.Description = Text_Renamed
			'        WorldTemp.Description = Mid$(Text & Space(100), 100 * ScrollWorldDesc + 1, 100)
			'UPGRADE_ISSUE: picTomeNew was upgraded to a Panel, and cannot be coerced to a PictureBox. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="0FF1188E-27C0-4FED-842D-159C65894C9B"'
			If X1 > 1 Then WorldScrollBarShow(picTomeNew, 533, 135, 72, ScrollWorldDesc + 1, X1, 1)
			rc = fReadValue(FileName, "World", "PicDesc", "S", gDataPath & "\Stock\blankmap.bmp|" & gDataPath & "\Stock\blankmap.bmp", Text_Renamed)
			If InStr(Text_Renamed, "|") = 0 Then Text_Renamed = Text_Renamed & "|" & gDataPath & "\Stock\blankmap.bmp"
			WorldTemp.PicDesc = Text_Renamed
			' Call PlayMusic(WorldTemp.IntroMusic, frmMain, WorldTemp.MusicFolder)
			ShowText(picTomeNew, 328, 135, 207, 72, bdFontNoxiousBlack, (WorldTemp.Description), False, False)
			' picTmp and picToHit are "borrowed" to store the world's pictures (if any)
			If oFileSys.CheckExists(gAppPath & "\Roster\" & WorldTemp.Name & "\" & VB.Left(WorldTemp.PicDesc, InStr(WorldTemp.PicDesc, "|") - 1), clsInOut.IOActionType.File) Then
				picTmp.Image = System.Drawing.Image.FromFile(gAppPath & "\Roster\" & WorldTemp.Name & "\" & VB.Left(WorldTemp.PicDesc, InStr(WorldTemp.PicDesc, "|") - 1))
			Else
				blnShow1 = False
			End If
			' Resize picture to fit in a 100 x 100 area
			If blnShow1 Then ResizePicture(picTmp, Size1, X1, Y1)
			If oFileSys.CheckExists(gAppPath & "\Roster\" & WorldTemp.Name & "\" & VB.Right(WorldTemp.PicDesc, Len(WorldTemp.PicDesc) - InStr(WorldTemp.PicDesc, "|")), clsInOut.IOActionType.File) Then
				picToHit.Image = System.Drawing.Image.FromFile(gAppPath & "\Roster\" & WorldTemp.Name & "\" & VB.Right(WorldTemp.PicDesc, Len(WorldTemp.PicDesc) - InStr(WorldTemp.PicDesc, "|")))
			Else
				blnShow2 = False
			End If
			' Resize picture to fit in a 100 x 100 area
			If blnShow2 Then ResizePicture(picToHit, Size2, X2, Y2)
			If blnShow1 And blnShow2 Then
				' two pics available
				'UPGRADE_ISSUE: PictureBox method picTomeNew.PaintPicture was not upgraded. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="CC4C7EC0-C903-48FC-ACCC-81861D12DA4A"'
				picTomeNew.PaintPicture(picTmp, 330 + X1, 202 + Y1, picTmp.Width * Size1, picTmp.Height * Size1)
				'UPGRADE_ISSUE: PictureBox method picTomeNew.PaintPicture was not upgraded. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="CC4C7EC0-C903-48FC-ACCC-81861D12DA4A"'
				picTomeNew.PaintPicture(picToHit, 434 + X2, 202 + Y2, picToHit.Width * Size2, picToHit.Height * Size2)
			ElseIf blnShow1 Then 
				' only the first pic
				'UPGRADE_ISSUE: PictureBox method picTomeNew.PaintPicture was not upgraded. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="CC4C7EC0-C903-48FC-ACCC-81861D12DA4A"'
				picTomeNew.PaintPicture(picTmp, 382 + X1, 202 + Y1, picTmp.Width * Size1, picTmp.Height * Size1)
			ElseIf blnShow2 Then 
				' only the second pic
				'UPGRADE_ISSUE: PictureBox method picTomeNew.PaintPicture was not upgraded. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="CC4C7EC0-C903-48FC-ACCC-81861D12DA4A"'
				picTomeNew.PaintPicture(picToHit, 382 + X2, 202 + Y2, picToHit.Width * Size2, picToHit.Height * Size2)
			Else
				' nothing, in that case default to the world map
				If oFileSys.CheckExists(gAppPath & "\Roster\" & WorldTemp.Name & "\" & WorldTemp.PictureFile, clsInOut.IOActionType.File) = False Then
					' [Titi 2.4.9] not found - either deleted or renamed - anyway default to current world!
					MessageShow(WorldTemp.PictureFile & " not found! Defaulting to blank map...", 0)
					DialogSetUp(modGameGeneral.DLGTYPE.bdDlgNoReply)
					DialogSetButton(1, "OK")
					DialogShow("Dm", "Error in " & WorldTemp.Name & "!" & Chr(13) & WorldTemp.PictureFile & " not found!" & Chr(13) & Chr(13) & WorldNow.Name & " is still the current world.")
					picTmp.Image = System.Drawing.Image.FromFile(gDataPath & "\Stock\NoSuchFile_2.bmp")
					DialogHide()
					WorldIndex = WorldNow.Index
				Else
					picTmp.Image = System.Drawing.Image.FromFile(gAppPath & "\Roster\" & WorldTemp.Name & "\" & WorldTemp.PictureFile)
				End If
				ResizePicture(picTmp, Size1, X1, Y1)
				'UPGRADE_ISSUE: PictureBox method picTomeNew.PaintPicture was not upgraded. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="CC4C7EC0-C903-48FC-ACCC-81861D12DA4A"'
				picTomeNew.PaintPicture(picTmp, 382 + X1, 202 + Y1, picTmp.Width * Size1, picTmp.Height * Size1)
			End If
			' now, reset picToHit
			picToHit.Image = System.Drawing.Image.FromFile(gDataPath & "\Interface\" & GlobalInterfaceName & "\" & "CombatRoll.bmp")
		End If
		' Set up buttons
		'UPGRADE_ISSUE: picTomeNew was upgraded to a Panel, and cannot be coerced to a PictureBox. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="0FF1188E-27C0-4FED-842D-159C65894C9B"'
		ScrollBarShow(picTomeNew, 276, 62, 267, ScrollTop, bdOptionsCount - 9, 0)
		ShowButton(picTomeNew, 380, 364, "Default", False)
		ShowButton(picTomeNew, 476, 364, "Done", False)
		' Refresh the screen
		picTomeNew.Refresh()
		picTomeNew.BringToFront()
		picTomeNew.Visible = True
	End Sub
	
	Private Function ShowShortcuts(ByRef intKey As Short) As String
		Dim code As Short
		code = IIf(intKey > 200, intKey - 200, IIf(intKey > 100, intKey - 100, intKey + 200 * CShort(intKey > 200)))
		Select Case code
			Case System.Windows.Forms.Keys.Left
				ShowShortcuts = IIf(intKey > 200, "+", IIf(intKey > 100, "*", "")) & "left"
			Case System.Windows.Forms.Keys.Right
				ShowShortcuts = IIf(intKey > 200, "+", IIf(intKey > 100, "*", "")) & "right"
			Case System.Windows.Forms.Keys.Up
				ShowShortcuts = IIf(intKey > 200, "+", IIf(intKey > 100, "*", "")) & "up "
			Case System.Windows.Forms.Keys.Down
				ShowShortcuts = IIf(intKey > 200, "+", IIf(intKey > 100, "*", "")) & "down"
			Case Else
				ShowShortcuts = IIf(intKey > 200, "+" & Chr(intKey + 200 * CShort(intKey > 200)), IIf(intKey > 100, "*" & Chr(intKey + 100 * CShort(intKey > 100)), Chr(intKey + 200 * CShort(intKey > 200))))
		End Select
	End Function
	
	'UPGRADE_NOTE: Size was upgraded to Size_Renamed. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="A9E4979A-37FA-4718-9994-97DD76ED70A7"'
	Private Sub ResizePicture(ByRef picPic As System.Windows.Forms.PictureBox, ByRef Size_Renamed As Single, ByRef X As Short, ByRef Y As Short)
		' Resize picture to fit in a 100 x 100 area
		Size_Renamed = VB6.PixelsToTwipsY(picPic.Height) / VB6.PixelsToTwipsX(picPic.Width)
		If VB6.PixelsToTwipsY(picPic.Height) * Size_Renamed > 100 Then
			Size_Renamed = 100 / VB6.PixelsToTwipsY(picPic.Height)
		End If
		If VB6.PixelsToTwipsX(picPic.Width) * Size_Renamed > 100 Then
			Size_Renamed = 100 / VB6.PixelsToTwipsX(picPic.Width)
		End If
		' Center resized picture in the area
		X = CShort(100 - (VB6.PixelsToTwipsX(picPic.Width) * Size_Renamed)) / 2
		Y = CShort(100 - (VB6.PixelsToTwipsY(picPic.Height) * Size_Renamed)) / 2
	End Sub
	
	Private Sub OptionsChange(ByRef Index As Short, ByRef Direction As Short)
		Dim rc As Integer
		Dim Found As Short
		Dim FileName As String
		Dim CreatureX As Creature
		Select Case Index
			Case 1 ' Music
				GlobalMusicState = LoopNumber(0, 1, GlobalMusicState, 1)
				If GlobalMusicState = 0 Then
					'MediaPlayerMusic.Stop
					oGameMusic.StopPlay()
					oGameMusic.mciClose()
					'                rc = mciSendString("close songnow", 0&, 0, 0)
				Else
					Call PlayMusicRnd(modIOFunc.RNDMUSICSTYLE.Adventure, Me)
				End If
			Case 2 ' Music Random
				'GlobalMusicRandom = LoopNumber(0, 1, GlobalMusicRandom, 1)
				GlobalMusicRandom = IIf(GlobalMusicRandom, 0, 1)
			Case 3 ' Right-Click Menu
				GlobalRightClick = LoopNumber(0, 1, GlobalRightClick, 1)
			Case 4 ' Game Speed
				GlobalGameSpeed = LoopNumber(1, 20, GlobalGameSpeed, 1)
			Case 5 ' Dice Rolling
				GlobalDiceRolling = LoopNumber(0, 1, GlobalDiceRolling, 1)
			Case 6 ' Sound Effects
				GlobalWAVState = LoopNumber(0, 1, GlobalWAVState, 1)
			Case 7 ' Debug Mode
				GlobalDebugMode = LoopNumber(0, 1, GlobalDebugMode, 1)
			Case 8 ' Overswing
				GlobalOverSwing = LoopNumber(0, 1, GlobalOverSwing, 1)
			Case 9 ' Mouse Click
				GlobalMouseClick = LoopNumber(0, 1, GlobalMouseClick, 1)
			Case 10 ' Auto End Turn
				GlobalAutoEndTurn = LoopNumber(0, 1, GlobalAutoEndTurn, 1)
			Case 11 ' Fast Move
				GlobalFastMove = LoopNumber(0, 1, GlobalFastMove, 1)
			Case 12 ' Automatic centering of the map on the party
				GlobalAutoCenter = LoopNumber(0, 1, GlobalAutoCenter, 1)
			Case 13 ' Interface
				'UPGRADE_WARNING: Screen property Screen.MousePointer has a new behavior. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6BA9B8D2-2A32-4B6E-8D36-44949974A5B4"'
				System.Windows.Forms.Cursor.Current = System.Windows.Forms.Cursors.WaitCursor
				GlobalInterfaceIndex = LoopNumber(1, GlobalInterfaceList.Count(), GlobalInterfaceIndex, Direction)
				'UPGRADE_WARNING: Couldn't resolve default property of object GlobalInterfaceList(). Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
				GlobalInterfaceName = GlobalInterfaceList.Item(GlobalInterfaceIndex)
				' Reload and refesh the entire screen
				GameInterfaceLoad(True)
				OptionsShow(ScrollSelect)
				BorderDrawAll(False)
				If picMenu.Visible = True Or picToHit.Visible = True Then
					BorderDrawSides(True, -1)
					BorderDrawButtons(0)
					Me.Refresh()
					'UPGRADE_ISSUE: PictureBox method picMainMenu.Cls was not upgraded. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="CC4C7EC0-C903-48FC-ACCC-81861D12DA4A"'
					picMainMenu.Cls()
					'UPGRADE_ISSUE: PictureBox property picBlack.hdc was not upgraded. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="CC4C7EC0-C903-48FC-ACCC-81861D12DA4A"'
					'UPGRADE_ISSUE: PictureBox property picMainMenu.hdc was not upgraded. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="CC4C7EC0-C903-48FC-ACCC-81861D12DA4A"'
					rc = BitBlt(picMainMenu.hdc, 122, 103, 288, 195, picBlack.hdc, 0, 340, SRCCOPY)
					'UPGRADE_ISSUE: PictureBox property picBlack.hdc was not upgraded. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="CC4C7EC0-C903-48FC-ACCC-81861D12DA4A"'
					'UPGRADE_ISSUE: PictureBox property picMainMenu.hdc was not upgraded. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="CC4C7EC0-C903-48FC-ACCC-81861D12DA4A"'
					rc = BitBlt(picMainMenu.hdc, 122, 103 + 0 * 39, 288, 39, picBlack.hdc, 0, 536 + 0 * 39, SRCCOPY)
					picMainMenu.Refresh()
					If picGrid.Visible = False Then
						MenuDrawParty()
					Else
						CombatDrawAttack(CreatureWithTurn, CreatureWithTurn, False)
						For	Each CreatureX In EncounterNow.Creatures
							LoadCreaturePic(CreatureX)
						Next CreatureX
					End If
					picMenu.Refresh()
				End If
				Me.Refresh()
				'UPGRADE_WARNING: Screen property Screen.MousePointer has a new behavior. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6BA9B8D2-2A32-4B6E-8D36-44949974A5B4"'
				System.Windows.Forms.Cursor.Current = System.Windows.Forms.Cursors.Default
			Case 14 ' Dice
				GlobalDiceIndex = LoopNumber(1, GlobalDiceList.Count(), GlobalDiceIndex, Direction)
				'UPGRADE_WARNING: Couldn't resolve default property of object GlobalDiceList(). Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
				GlobalDiceName = GlobalDiceList.Item(GlobalDiceIndex)
				LoadDicePic(GlobalDiceName)
				'            Set picDice = LoadPicture(gDataPath & "\Interface\Dice\" & GlobalDiceName)
				'            picTomeNew.PaintPicture picDice.Picture, 320, 220, picDice.Width, picDice.Height
				OptionsShow(ScrollSelect)
			Case 15 ' Keyboard shortcuts
				' [Titi 2.4.9] KingdomIndex exists, and is used only in Creator, so I borrow it
				KingdomIndex = LoopNumber(0, 2, KingdomIndex, Direction)
				OptionsShow(ScrollSelect)
			Case 16 ' Current active WORLD
				If TomeMenu <> bdTomeOptions Then
					WorldIndex = WorldIndex + Direction
					If WorldIndex > Worlds.Count() Then
						WorldIndex = 1
					ElseIf WorldIndex < 1 Then 
						WorldIndex = Worlds.Count()
					End If
				End If
				OptionsShow(ScrollSelect)
			Case 17 ' show splash and credits screens
				If Val(GlobalCredits) = 1 Then
					GlobalCredits = "0"
					'UPGRADE_ISSUE: PictureBox property picMisc.hdc was not upgraded. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="CC4C7EC0-C903-48FC-ACCC-81861D12DA4A"'
					'UPGRADE_ISSUE: PictureBox property picTomeNew.hdc was not upgraded. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="CC4C7EC0-C903-48FC-ACCC-81861D12DA4A"'
					rc = BitBlt(picTomeNew.hdc, 26, 73 + 14 * 25, 18, 18, picMisc.hdc, 0, 18, SRCCOPY)
				Else
					GlobalCredits = "1"
					'UPGRADE_ISSUE: PictureBox property picMisc.hdc was not upgraded. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="CC4C7EC0-C903-48FC-ACCC-81861D12DA4A"'
					'UPGRADE_ISSUE: PictureBox property picTomeNew.hdc was not upgraded. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="CC4C7EC0-C903-48FC-ACCC-81861D12DA4A"'
					rc = BitBlt(picTomeNew.hdc, 26, 73 + 14 * 25, 18, 18, picMisc.hdc, 18, 18, SRCCOPY)
				End If
				Call oErr.LogText("Splash option changed. GlobalCredits = " & GlobalCredits)
		End Select
	End Sub
	
	Private Sub OptionsClick(ByRef AtX As Short, ByRef AtY As Short, ByRef ButtonDown As Short)
		Dim lResult As Integer
		Dim code, c, X As Short
		Dim sText As String
		Dim WorldTemp As World '* 8192
		'UPGRADE_NOTE: Text was upgraded to Text_Renamed. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="A9E4979A-37FA-4718-9994-97DD76ED70A7"'
		Dim Text_Renamed As String
		If PointIn(AtX, AtY, 380, 364, 90, 18) Then
			' Default
			If ButtonDown Then
				ShowButton(picTomeNew, 380, 364, "Default", True)
			Else
				Call InitKeyboardShortCuts() ' [Titi 2.4.9]
				Call PlayClickSnd(modIOFunc.ClickType.ifClick)
				ShowButton(picTomeNew, 380, 364, "Default", False)
				GlobalMusicState = 1
				GlobalMusicRandom = 1
				GlobalRightClick = 1
				GlobalGameSpeed = 12
				GlobalDiceRolling = 1
				GlobalWAVState = 1
				GlobalDebugMode = 0
				GlobalOverSwing = 1
				GlobalMouseClick = 1
				GlobalAutoEndTurn = 0
				GlobalFastMove = 0
				GlobalAutoCenter = 0
				ScrollTop = 1
				OptionsShow(1)
				GlobalInterfaceName = "WoodWeave"
				GlobalDiceName = "Bone White.bmp"
				' updated by count0 for v2-4-6
				GlobalCredits = "0"
				GameInterfaceLoad(True)
				OptionsShow(ScrollSelect)
			End If
			picTomeNew.Refresh()
		ElseIf PointIn(AtX, AtY, 476, 364, 90, 18) Then 
			' Done
			If ButtonDown Then
				ShowButton(picTomeNew, 476, 364, "Done", True)
				picTomeNew.Refresh()
			Else
				ShowButton(picTomeNew, 476, 364, "Done", False)
				picTomeNew.Refresh()
				Call PlayClickSnd(modIOFunc.ClickType.ifClick)
				picTomeNew.Visible = False
				' Save the new Options
				If WorldIndex <> WorldNow.Index Then
					'If MediaPlayerMusic.PlayState = 2 Then
					If oGameMusic.Status = IMCI.VIDEOSTATE.vsPLAYING Then
						'MediaPlayerMusic.Stop
						oGameMusic.StopPlay()
						oGameMusic.mciClose()
					End If
					WorldNow.IsCurrent = False
					WorldNow = Worlds.Item(WorldIndex)
					WorldNow.IsCurrent = True
					Call PlayMusic(WorldNow.IntroMusic, Me, WorldNow.MusicFolder)
				End If
				OptionsSave()
				' save world changes
				Call oFileSys.CheckExists(gAppPath & "\Settings.ini", clsInOut.IOActionType.File, True)
				lResult = fWriteValue(gAppPath & "\Settings.ini", "World", "Current", "S", WorldNow.Name)
				'Frozen = False
			End If
		ElseIf PointIn(AtX, AtY, 276, 62, 18, 267) Then 
			'UPGRADE_ISSUE: picTomeNew was upgraded to a Panel, and cannot be coerced to a PictureBox. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="0FF1188E-27C0-4FED-842D-159C65894C9B"'
			If ScrollBarClick(AtX, AtY, ButtonDown, picTomeNew, 276, 62, 267, ScrollTop, bdOptionsCount, 9) = True Then
				OptionsShow(ScrollSelect)
			End If
		ElseIf PointIn(AtX, AtY, 26, 73, 18, 250) And ButtonDown = True Then 
			c = CShort((AtY - 76) / 25)
			If IsBetween(c, 0, 9) Then
				OptionsChange(ScrollTop + c, 0)
				OptionsShow(ScrollTop + c)
				Call PlayClickSnd(modIOFunc.ClickType.ifClick)
			End If
		ElseIf PointIn(AtX, AtY, 53, 73, 276, 250) And ButtonDown = True Then 
			' Item in ListBox
			c = CShort((AtY - 76) / 25)
			' If IsBetween(c, 0, 9) Or c = 14 Then
			If IsBetween(c, 0, 9) Then
				OptionsShow(ScrollTop + c)
				Call PlayClickSnd(modIOFunc.ClickType.ifClick)
			End If
			'    ElseIf PointIn(AtX, AtY, 340, 200, 90, 18) And (ScrollSelect = 11 Or ScrollSelect = 12 Or ScrollSelect = 13) Then
		ElseIf PointIn(AtX, AtY, 340, 298, 90, 18) And (ScrollSelect >= 12 And ScrollSelect <= 16) Then  ' World change forgotten in 2.4.7 [Titi]
			' Back Interface or Dice Set or Change World or Keyboard Shortcuts
			If ButtonDown Then
				'            ShowButton picTomeNew, 340, 200, "Back", True
				ShowButton(picTomeNew, 340, 298, "Back", True) ' [Titi 2.4.9] moved to bottom to give space to display the world's details
				picTomeNew.Refresh()
			Else
				'            ShowButton picTomeNew, 340, 200, "Back", False
				ShowButton(picTomeNew, 340, 298, "Back", False) ' [Titi 2.4.9] moved to bottom to give space to display the world's details
				picTomeNew.Refresh()
				Call PlayClickSnd(modIOFunc.ClickType.ifClick)
				OptionsChange(ScrollSelect, -1)
			End If
			'    ElseIf PointIn(AtX, AtY, 435, 200, 90, 18) And (ScrollSelect = 11 Or ScrollSelect = 12 Or ScrollSelect = 13) Then
		ElseIf PointIn(AtX, AtY, 435, 298, 90, 18) And (ScrollSelect >= 12 And ScrollSelect <= 16) Then  ' World change forgotten in 2.4.7 [Titi]
			' Next Interface or Dice Set or Change World or Keyboard Shortcuts
			If ButtonDown Then
				'            ShowButton picTomeNew, 435, 200, "Next", True
				ShowButton(picTomeNew, 435, 298, "Next", True) ' [Titi 2.4.9] moved to bottom to give space to display the world's details
				picTomeNew.Refresh()
			Else
				'            ShowButton picTomeNew, 435, 200, "Next", False
				ShowButton(picTomeNew, 435, 298, "Next", False) ' [Titi 2.4.9] moved to bottom to give space to display the world's details
				picTomeNew.Refresh()
				Call PlayClickSnd(modIOFunc.ClickType.ifClick)
				OptionsChange(ScrollSelect, 1)
			End If
			'    ElseIf PointIn(AtX, AtY, 53, 73, 276, 250) And (ScrollSelect = 14) Then
		ElseIf PointIn(AtX, AtY, 53, 73, 276, 250) And (ScrollSelect = 17) Then  ' 14 is now used by the world change option
			' Splash/credit screen toggle
			'    End If
		ElseIf PointIn(AtX, AtY, 533, 135, 15, 72) And ScrollSelect = 16 Then 
			' scrollbar next to the world description
			WorldTemp = Worlds.Item(WorldIndex)
			'        lResult = fReadValue(gAppPath & "\roster\" & WorldTemp.Name & "\" & WorldTemp.Name & ".ini", "World", "Description", "S", "No description available", sText)
			lResult = fReadValue(gAppPath & "\roster\" & WorldTemp.Name & "\" & WorldTemp.Name & ".ini", "World", "DescriptionLines", "S", "1", sText)
			'        rc = fReadValue(FileName, "World", "Description" & Trim$(Str(ScrollWorldDesc + 1)), "S", "No description available", Text)
			'        WorldTemp.Description = Text
			'UPGRADE_ISSUE: picTomeNew was upgraded to a Panel, and cannot be coerced to a PictureBox. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="0FF1188E-27C0-4FED-842D-159C65894C9B"'
			If WorldScrollBarClick(AtX, AtY, ButtonDown, picTomeNew, 533, 135, 72, ScrollWorldDesc, Val(sText), 1) Then
				OptionsShow(ScrollSelect)
			End If
		ElseIf ScrollSelect = 15 Then 
			X = AtX
			' change keyboard shortcuts
			If PointIn(AtX, AtY, 328, 169, 80, 127) Then ' first column
				c = CShort(AtY / 17) - 9
				Select Case c
					Case 1
						sText = IIf(KingdomIndex = 0, "Debug", IIf(KingdomIndex = 1, "Left", "Left"))
					Case 2
						sText = IIf(KingdomIndex = 0, "Exit", IIf(KingdomIndex = 1, "Up", "Up"))
					Case 3
						sText = IIf(KingdomIndex = 0, "Load", IIf(KingdomIndex = 1, "", ""))
					Case 4
						sText = IIf(KingdomIndex = 0, "Music", IIf(KingdomIndex = 1, "", "Journals & Quests"))
					Case 5
						sText = IIf(KingdomIndex = 0, "", IIf(KingdomIndex = 1, "", "Listen"))
					Case 6
						sText = IIf(KingdomIndex = 0, "", IIf(KingdomIndex = 1, "", "Skills"))
					Case 7
						sText = IIf(KingdomIndex = 0, "", IIf(KingdomIndex = 1, "", "Search"))
					Case 8
						sText = IIf(KingdomIndex = 0, "", IIf(KingdomIndex = 1, "", "Inventory"))
					Case Else
						sText = vbNullString
				End Select
			ElseIf PointIn(AtX, AtY, 444, 169, 80, 127) Then  ' second column
				c = CShort(AtY / 17) - 9
				Select Case c
					Case 1
						sText = IIf(KingdomIndex = 0, "Quit", IIf(KingdomIndex = 1, "Right", "Right"))
					Case 2
						sText = IIf(KingdomIndex = 0, "Speed", IIf(KingdomIndex = 1, "Down", "Down"))
					Case 3
						sText = IIf(KingdomIndex = 0, "Save", IIf(KingdomIndex = 1, "", ""))
					Case 4
						sText = IIf(KingdomIndex = 0, "Sound", IIf(KingdomIndex = 1, "", "Journals & Quests"))
					Case 5
						sText = IIf(KingdomIndex = 0, "", IIf(KingdomIndex = 1, "", "Talk"))
					Case 6
						sText = IIf(KingdomIndex = 0, "", IIf(KingdomIndex = 1, "", "Status"))
					Case 7
						sText = IIf(KingdomIndex = 0, "", IIf(KingdomIndex = 1, "", "Open"))
					Case 8
						sText = IIf(KingdomIndex = 0, "", IIf(KingdomIndex = 1, "", "Equip"))
					Case Else
						sText = vbNullString
				End Select
			End If
			DialogSetUp(modGameGeneral.DLGTYPE.bdDlgReplyText)
			If sText <> "" Then
				DialogShow("", "Shortcut for " & sText)
				DialogHide()
				If DialogText <> vbNullString Then
					If Len(DialogText) > 1 Then
						If UCase(VB.Left(DialogText & Space(4), 4)) = "CTRL" Then
							If VB.Right(Space(4) & DialogText, 4) = "left" Then
								code = System.Windows.Forms.Keys.Left + 200
							ElseIf VB.Right(Space(5) & DialogText, 5) = "right" Then 
								code = System.Windows.Forms.Keys.Right + 200
							ElseIf VB.Right(Space(2) & DialogText, 2) = "up" Then 
								code = System.Windows.Forms.Keys.Up + 200
							ElseIf VB.Right(Space(4) & DialogText, 4) = "down" Then 
								code = System.Windows.Forms.Keys.Down + 200
							Else
								code = Asc(VB.Right(DialogText, 1)) + 200 - 32
							End If
						ElseIf UCase(VB.Left(DialogText & Space(5), 5)) = "SHIFT" Then 
							If VB.Right(Space(4) & DialogText, 4) = "left" Then
								code = System.Windows.Forms.Keys.Left + 100
							ElseIf VB.Right(Space(5) & DialogText, 5) = "right" Then 
								code = System.Windows.Forms.Keys.Right + 100
							ElseIf VB.Right(Space(2) & DialogText, 2) = "up" Then 
								code = System.Windows.Forms.Keys.Up + 100
							ElseIf VB.Right(Space(4) & DialogText, 4) = "down" Then 
								code = System.Windows.Forms.Keys.Down + 100
							Else
								code = Asc(VB.Right(DialogText, 1)) + 100 - 32
							End If
						Else
							If VB.Right(Space(4) & DialogText, 4) = "left" Then
								code = System.Windows.Forms.Keys.Left
							ElseIf VB.Right(Space(5) & DialogText, 5) = "right" Then 
								code = System.Windows.Forms.Keys.Right
							ElseIf VB.Right(Space(2) & DialogText, 2) = "up" Then 
								code = System.Windows.Forms.Keys.Up
							ElseIf VB.Right(Space(4) & DialogText, 4) = "down" Then 
								code = System.Windows.Forms.Keys.Down
							End If
						End If
					Else
						code = Asc(DialogText) - 32
					End If
					Select Case KingdomIndex
						Case 0
							If X < 444 Then
								If NotTwice(2 * c - 1, code) Then bdKey(2 * c - 1) = code
							Else
								If NotTwice(2 * c, code) Then bdKey(2 * c) = code
							End If
						Case 1
							If X < 444 Then
								If NotTwice(8 + 2 * c - 1, code) Then bdKey(8 + 2 * c - 1) = code
							Else
								If NotTwice(8 + 2 * c, code) Then bdKey(8 + 2 * c) = code
							End If
						Case 2
							If c = 4 Then ' Journals and Quests takes a full line
								If NotTwice(17, code) Then bdKey(17) = code
							ElseIf c < 3 Then  ' still the odd/even counting
								If X < 444 Then
									If NotTwice(12 + 2 * c - 1, code) Then bdKey(12 + 2 * c - 1) = code
								Else
									If NotTwice(12 + 2 * c, code) Then bdKey(12 + 2 * c) = code
								End If
							Else ' shortcuts AFTER J&Q become even/odd PLUS a line is skipped
								If X < 444 Then
									If NotTwice(12 + 2 * c - 4, code) Then bdKey(12 + 2 * c - 4) = code
								Else
									If NotTwice(12 + 2 * c - 3, code) Then bdKey(12 + 2 * c - 3) = code
								End If
							End If
					End Select
				End If
			End If
			OptionsShow(ScrollSelect)
		End If
	End Sub
	
	Private Sub TurnCycleCreatureClear(ByRef CreatureX As Creature)
		Dim c As Short
		' If Frozen, then Frozen no longer
		CreatureX.Frozen = False
		CreatureX.Afraid = False
		' Rest the Creature
		CreatureX.ActionPoints = CreatureX.ActionPointsMax
		' Clear all Bonuses
		CreatureX.StrengthBonus = 0
		CreatureX.AgilityBonus = 0
		CreatureX.WillBonus = 0
		CreatureX.AttackBonus = 0
		CreatureX.DamageBonus = 0
		CreatureX.DefenseBonus = 0
		CreatureX.MovementCostBonus = 0
		CreatureX.CombatAttitude = 0
		CreatureX.OpportunityAttack = False
		For c = 0 To 8
			CreatureX.ResistanceBonus(c) = 0
		Next c
	End Sub
	
	Private Sub TurnCycleCreature(ByRef CreatureX As Creature)
		Dim NoFail, c As Short
		If picGrid.Visible = False Then
			' Fire Post-Turn Trigger on CreatureWithTurn
			If CreatureX.HPNow > 0 Then
				CreatureNow = CreatureX
				ItemNow = CreatureX.ItemInHand
				NoFail = FireTriggers(CreatureX, bdPostTurn)
				For	Each ItemNow In CreatureX.Items
					If ItemNow.IsReady = True Then
						CreatureNow = CreatureX
						NoFail = FireTriggers(ItemNow, bdPostTurn)
					End If
				Next ItemNow
			End If
			' Clear all their statuses
			TurnCycleCreatureClear(CreatureX)
			' Fire Pre-Turn Trigger on CreatureWithTurn
			If CreatureX.HPNow > 0 Then
				CreatureNow = CreatureX
				ItemNow = CreatureX.ItemInHand
				NoFail = FireTriggers(CreatureX, bdPreTurn)
				For	Each ItemNow In CreatureX.Items
					If ItemNow.IsReady = True Then
						CreatureNow = CreatureX
						NoFail = FireTriggers(ItemNow, bdPreTurn)
					End If
				Next ItemNow
			End If
		End If
		' Fire up Post-TurnCycle Triggers on CreatureWithTurn
		If CreatureX.HPNow > 0 Then
			CreatureNow = CreatureX
			ItemNow = CreatureX.ItemInHand
			NoFail = FireTriggers(CreatureX, bdPostTurnCycle)
			For	Each ItemNow In CreatureX.Items
				If ItemNow.IsReady Then
					CreatureNow = CreatureX
					NoFail = FireTriggers(ItemNow, bdPostTurnCycle)
				End If
			Next ItemNow
		End If
		CreatureX.HadTurn = False
		' Remove Timed Triggers
		TurnCycleTimedTriggers(CreatureX)
		For	Each ItemNow In CreatureX.Items
			TurnCycleTimedTriggers(ItemNow)
		Next ItemNow
	End Sub
	
	Private Sub TurnCycle()
		Dim NoFail, c As Short
		Dim CreatureX As Creature
		' Change the time first
		TurnCycleTime()
		' Cycle the Rune Pool
		TurnCycleRunes()
		' Remove Timed Triggers
		TurnCycleTimedTriggers(Tome)
		' Cycle through entire Party
		For	Each CreatureX In Tome.Creatures
			TurnCycleCreature(CreatureX)
		Next CreatureX
		' Fire Post-TurnCycle Triggers on Encounter
		NoFail = FireTriggers(EncounterNow, bdPostTurnCycle)
		' Remove Timed Triggers
		TurnCycleTimedTriggers(EncounterNow)
		' Cycle through all Creatures in the Encounter
		For	Each CreatureX In EncounterNow.Creatures
			TurnCycleCreature(CreatureX)
		Next CreatureX
		If picGrid.Visible = False Then
			MessageShow("", 0)
		End If
		
	End Sub
	
	Private Sub TurnCycleDemo()
		'more garbage
	End Sub
	
	'UPGRADE_NOTE: Map was upgraded to Map_Renamed. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="A9E4979A-37FA-4718-9994-97DD76ED70A7"'
	Private Sub TurnCycleRunes()
		Dim Map_Renamed As Object
		Dim c, Found As Short
		If Int(Rnd() * 10) > 8 Then
			For c = 0 To 19
				Found = False
				Select Case c Mod 4
					Case 0 ' Very Rare
						'UPGRADE_WARNING: Untranslated statement in TurnCycleRunes. Please check source code.
					Case 1, 2 ' Rare
						'UPGRADE_WARNING: Untranslated statement in TurnCycleRunes. Please check source code.
					Case 3 ' Common
						'UPGRADE_WARNING: Untranslated statement in TurnCycleRunes. Please check source code.
				End Select
				If Found = True Then
					'UPGRADE_WARNING: Screen property Screen.MousePointer has a new behavior. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6BA9B8D2-2A32-4B6E-8D36-44949974A5B4"'
					System.Windows.Forms.Cursor.Current = System.Windows.Forms.Cursors.WaitCursor
					BorderDrawSides(True, c)
					BorderDrawButtons(0)
					Me.Refresh()
					Call PlayClickSnd(modIOFunc.ClickType.ifClickCast)
					'UPGRADE_WARNING: Couldn't resolve default property of object Map.Runes. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
					Map.Runes(c) = Map.Runes(c) + 1
				End If
			Next c
			BorderDrawSides(True, -1)
			BorderDrawButtons(0)
			Me.Refresh()
			'UPGRADE_WARNING: Screen property Screen.MousePointer has a new behavior. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6BA9B8D2-2A32-4B6E-8D36-44949974A5B4"'
			System.Windows.Forms.Cursor.Current = System.Windows.Forms.Cursors.Default
		End If
	End Sub
	
	'UPGRADE_NOTE: Tome was upgraded to Tome_Renamed. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="A9E4979A-37FA-4718-9994-97DD76ED70A7"'
	Private Sub TurnCycleTime()
		Dim Tome_Renamed As Object
		Dim c As Short
		Dim rc As Integer
		Dim SunChange, NoFail As Short
		' Update the Turn, Cycle, Moon, and Year
		'UPGRADE_WARNING: Couldn't resolve default property of object Tome.Turn. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
		Tome.Turn = Tome.Turn + 1
		'UPGRADE_WARNING: Couldn't resolve default property of object Tome.Turn. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
		If Tome.Turn > 200 Then
			'UPGRADE_WARNING: Couldn't resolve default property of object Tome.Turn. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
			Tome.Turn = 1
			'UPGRADE_WARNING: Couldn't resolve default property of object Tome.Cycle. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
			Tome.Cycle = Tome.Cycle + 1
			'UPGRADE_WARNING: Couldn't resolve default property of object Tome.Cycle. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
			If Tome.Cycle > 50 Then
				'UPGRADE_WARNING: Couldn't resolve default property of object Tome.Cycle. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
				Tome.Cycle = 1
				'UPGRADE_WARNING: Couldn't resolve default property of object Tome.Moon. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
				Tome.Moon = Tome.Moon + 1
				'UPGRADE_WARNING: Couldn't resolve default property of object Tome.Moon. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
				If Tome.Moon > 10 Then
					'UPGRADE_WARNING: Couldn't resolve default property of object Tome.Moon. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
					Tome.Moon = 1
					'UPGRADE_WARNING: Couldn't resolve default property of object Tome.Year. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
					Tome.Year = Tome.Year + 1
					'UPGRADE_WARNING: Couldn't resolve default property of object Tome.Year. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
					If Tome.Year > 10000 Then
						'UPGRADE_WARNING: Couldn't resolve default property of object Tome.Year. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
						Tome.Year = 1
					End If
				End If
			End If
		End If
		' Update the darkness
		SunChange = 0
		'UPGRADE_WARNING: Couldn't resolve default property of object Tome.Turn. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
		If (Tome.Turn > 169 Or Tome.Turn < 31) Then
			'UPGRADE_WARNING: Couldn't resolve default property of object Tome.Turn. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
			Select Case Tome.Turn
				Case 30 ' Daylight
					SunChange = 0
				Case 23 To 29
					SunChange = 20
					'UPGRADE_WARNING: Couldn't resolve default property of object Tome.Turn. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
					If Tome.Turn = 23 Then
						MessageShow("It's dawn....", 0)
					End If
				Case 170 To 176
					SunChange = 20
					'UPGRADE_WARNING: Couldn't resolve default property of object Tome.Turn. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
					If Tome.Turn = 170 Then
						MessageShow("It's dusk....", 0)
					End If
				Case 177 To 185, 18 To 22
					SunChange = 40
				Case 186 To 199, 10 To 17
					SunChange = 60
				Case 200, 0 To 9
					SunChange = 80
					'UPGRADE_WARNING: Couldn't resolve default property of object Tome.Turn. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
					If Tome.Turn = 200 Then
						MessageShow("It's midnight....", 0)
					End If
			End Select
			If SunChange <> Darker Then
				Darker = SunChange
				LoadTileSet(Darker, True)
				DrawMapAll()
				MessageClear()
			End If
		ElseIf Darker <> 0 Then 
			Darker = 0
			LoadTileSet(Darker, True)
			DrawMapAll()
			MessageClear()
		End If
		' Show current Date at the top of the screen
		'UPGRADE_WARNING: Couldn't resolve default property of object Tome.Turn. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
		GlobalTurnName = "Turn " & Tome.Turn
		'UPGRADE_WARNING: Couldn't resolve default property of object Tome.Cycle. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
		GlobalDayName = "Day " & Tome.Cycle
		'UPGRADE_WARNING: Couldn't resolve default property of object Tome.Moon. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
		GlobalMoonName = "of Moon " & Tome.Moon
		'UPGRADE_WARNING: Couldn't resolve default property of object Tome.Year. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
		GlobalYearName = "in the Year " & Tome.Year
		' Fire Post-TurnCycle Triggers on Party
		NoFail = FireTriggers(Tome, bdPostTurnCycle)
		BorderDrawTop()
		ShowText(Me, 0, 2, Me.ClientRectangle.Width, 14, bdFontSmallWhite, " ( " & GlobalTurnName & " ) ( " & GlobalDayName & " " & GlobalMoonName & " " & GlobalYearName & " ) ", True, False)
		Me.Refresh()
	End Sub
	
	Private Sub TurnCycleTimedTriggers(ByRef ObjectX As Object)
		Dim TriggerX As Trigger
		' Subtract Charges from Timed Triggers (and Destroy if < 1)
		'UPGRADE_WARNING: Couldn't resolve default property of object ObjectX.Triggers. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
		For	Each TriggerX In ObjectX.Triggers
			If TriggerX.IsTimed And TriggerX.Turns > 0 And Not TriggerX.IsSkill Then
				TriggerX.Turns = TriggerX.Turns - 1
				If TriggerX.Turns < 1 Then
					'UPGRADE_WARNING: Couldn't resolve default property of object ObjectX.RemoveTrigger. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
					ObjectX.RemoveTrigger("T" & TriggerX.Index)
				End If
			End If
		Next TriggerX
	End Sub
	
	'UPGRADE_NOTE: Map was upgraded to Map_Renamed. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="A9E4979A-37FA-4718-9994-97DD76ED70A7"'
	Private Function DoorNearBy(ByRef AtX As Short, ByRef AtY As Short) As Object
		Dim Map_Renamed As Object
		Dim X, c, i, Y As Short
		' Look in all directions for a locked/door Tile
		'UPGRADE_WARNING: Couldn't resolve default property of object DoorNearBy. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
		DoorNearBy = False
		For c = 1 To 3
			If c > 1 Then
				X = Greatest(Least(AtX + DirX(c), (Map.Width)), 0)
				Y = Greatest(Least(AtY + DirY(c), (Map.Height)), 0)
			Else
				X = AtX
				Y = AtY
			End If
			' Check all layers for a door
			For i = 0 To 2
				'UPGRADE_WARNING: Couldn't resolve default property of object Map.Tile. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
				If DoorNearByTile(Map.Tile(i, X, Y)) = True Then
					'UPGRADE_WARNING: Couldn't resolve default property of object DoorNearBy. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
					DoorNearBy = True
					Exit For
				End If
			Next i
			'UPGRADE_WARNING: Couldn't resolve default property of object DoorNearBy. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
			If DoorNearBy = True Then
				Exit For
			End If
		Next c
	End Function
	
	Private Function DoorNearByTile(ByRef Index As Short) As Short
		DoorNearByTile = False
		If Index > 0 Then
			'UPGRADE_WARNING: Untranslated statement in DoorNearByTile. Please check source code.
		End If
	End Function
	
	'UPGRADE_NOTE: Map was upgraded to Map_Renamed. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="A9E4979A-37FA-4718-9994-97DD76ED70A7"'
	'UPGRADE_NOTE: Tome was upgraded to Tome_Renamed. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="A9E4979A-37FA-4718-9994-97DD76ED70A7"'
	Private Sub DoorOpen(ByRef AtX As Short, ByRef AtY As Short)
		Dim Tome_Renamed As Object
		Dim Map_Renamed As Object
		Dim Y, i, c, X, Found As Short
		Dim NoFail, OpenIt As Short
		Dim EntryPointX As EntryPoint
		Dim TileX As Tile
		Dim ItemX As Item
		'UPGRADE_WARNING: Screen property Screen.MousePointer has a new behavior. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6BA9B8D2-2A32-4B6E-8D36-44949974A5B4"'
		System.Windows.Forms.Cursor.Current = System.Windows.Forms.Cursors.WaitCursor
		' Look in all directions for a locked/door Tile
		Found = False
		For c = 1 To 3
			If c > 1 Then
				X = Greatest(Least(AtX + DirX(c), (Map.Width)), 0)
				Y = Greatest(Least(AtY + DirY(c), (Map.Height)), 0)
			Else
				X = AtX
				Y = AtY
			End If
			' Look at all levels
			For i = 0 To 2
				'UPGRADE_WARNING: Couldn't resolve default property of object Map.Tile. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
				If DoorNearByTile(Map.Tile(i, X, Y)) = True Then
					'UPGRADE_WARNING: Couldn't resolve default property of object Map.Tile. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
					'UPGRADE_WARNING: Couldn't resolve default property of object Map.Tiles. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
					TileX = Map.Tiles("L" & Map.Tile(i, X, Y))
					Found = True
					Exit For
				End If
			Next i
			If Found = True Then
				Exit For
			End If
		Next c
		' If Found a door, then attempt to Open/Unlock it
		If Found = True Then
			' If Locked, then attempt to unlock. Else, attempt to unstick it. Else open it.
			OpenIt = False
			If TileX.KeyBits > 0 Then
				' Try all keys on the current Creature
				Found = False
				For	Each ItemX In CreatureWithTurn.Items
					If ItemX.KeyBits > 0 Then
						If (ItemX.KeyBits And TileX.KeyBits) = ItemX.KeyBits Then
							Found = True
							Exit For
						End If
					End If
				Next ItemX
				' If found a key, then unlock it
				If Found = True Then
					Call PlaySoundFile("DoorUnlock", Tome)
					DialogSetUp(modGameGeneral.DLGTYPE.bdDlgNoReply)
					DialogShow(CreatureWithTurn, CreatureWithTurn.Name & " unlocked the door with the " & ItemX.Name & ".")
					DialogHide()
					OpenIt = True
				Else
					' Attempt a lock pick
					GlobalPickLockChance = 0
					CreatureNow = CreatureWithTurn
					NoFail = FireTriggers(CreatureWithTurn, bdPrePickLock)
					For	Each ItemNow In CreatureWithTurn.Items
						If ItemNow.IsReady = True Then
							NoFail = FireTriggers(ItemNow, bdPrePickLock)
						End If
					Next ItemNow
					NoFail = FireTriggers(EncounterNow, bdPreUnlock)
					If Int(Rnd() * 100) < GlobalPickLockChance Then
						Call PlaySoundFile("DoorUnlock", Tome)
						DialogSetUp(modGameGeneral.DLGTYPE.bdDlgNoReply)
						DialogShow(CreatureWithTurn, CreatureWithTurn.Name & " picked the lock on the door.")
						DialogHide()
						OpenIt = True
					ElseIf GlobalPickLockChance > 0 Then 
						If Int(Rnd() * 100) > GlobalPickLockChance Then
							DialogDM(CreatureWithTurn.Name & " attempted to pick the lock, but jammed it shut instead!")
							'                        TileX.Chance = 0  <-- [Titi 2.4.9] will not update the "chance to open" on the map!
							' [Titi 2.4.9] if door is jammed, make it appear so!
							'                        Map.Tiles("L" & Map.Tile(i, X, Y)).Chance = 0
							'UPGRADE_WARNING: Couldn't resolve default property of object Map.Tile. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
							'UPGRADE_WARNING: Couldn't resolve default property of object Map.Tiles. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
							Map.Tiles("L" & Map.Tile(i, X, Y)).Key = 0
						Else
							DialogDM(CreatureWithTurn.Name & " failed to pick the lock on the door.")
							' [Titi 2.4.9] failure - it is now more difficult to pick the lock!
							'UPGRADE_WARNING: Couldn't resolve default property of object Map.Tile. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
							'UPGRADE_WARNING: Couldn't resolve default property of object Map.Tiles. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
							Map.Tiles("L" & Map.Tile(i, X, Y)).Chance = Map.Tiles("L" & Map.Tile(i, X, Y)).Chance - 5
							'UPGRADE_WARNING: Couldn't resolve default property of object Map.Tile. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
							'UPGRADE_WARNING: Couldn't resolve default property of object Map.Tiles. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
							If Map.Tiles("L" & Map.Tile(i, X, Y)).Chance <= 0 Then
								' too many failures - only option now is to bash the door open...
								'UPGRADE_WARNING: Couldn't resolve default property of object Map.Tile. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
								'UPGRADE_WARNING: Couldn't resolve default property of object Map.Tiles. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
								Map.Tiles("L" & Map.Tile(i, X, Y)).Chance = Int(Rnd() * 100)
								'UPGRADE_WARNING: Couldn't resolve default property of object Map.Tile. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
								'UPGRADE_WARNING: Couldn't resolve default property of object Map.Tiles. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
								Map.Tiles("L" & Map.Tile(i, X, Y)).Key = 0
							End If
						End If
					Else
						DialogDM("The door is locked.")
					End If
				End If
			ElseIf IsBetween(TileX.Chance, 1, 99) Then 
				' Try to open a stuck door
				If Int(Rnd() * 100) - CreatureWithTurn.Strength < TileX.Chance Then
					Call PlaySoundFile("Door", Tome)
					DialogSetUp(modGameGeneral.DLGTYPE.bdDlgNoReply)
					DialogShow(CreatureWithTurn, CreatureWithTurn.Name & " bashed the door open.")
					DialogHide()
					OpenIt = True
				Else
					DialogDM("The door is stuck shut.")
				End If
			Else
				' Open a plain door (not locked or stuck)
				Call PlaySoundFile("Door", Tome)
				OpenIt = True
			End If
			If OpenIt = True Then
				NoFail = FireTriggers(EncounterNow, bdPreOpen)
				If NoFail Then
					'UPGRADE_WARNING: Couldn't resolve default property of object Map.Tile. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
					Map.Tile(i, X, Y) = TileX.SwapTile
					'UPGRADE_WARNING: Couldn't resolve default property of object Tome.MapY. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
					'UPGRADE_WARNING: Couldn't resolve default property of object Tome.MapX. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
					MovePartySet(0, 0, Tome.MapX, Tome.MapY)
				End If
			End If
		Else ' Look for EntryPoint to Exit
			MovePartyExit()
		End If
		'UPGRADE_WARNING: Screen property Screen.MousePointer has a new behavior. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6BA9B8D2-2A32-4B6E-8D36-44949974A5B4"'
		System.Windows.Forms.Cursor.Current = System.Windows.Forms.Cursors.Default
	End Sub
	
	Private Sub ParseShortCuts(ByRef KeyCode As Short, ByRef Shift As Short)
		' [Titi 2.4.9] added for user-defined Shortcuts
		Dim i As Short
		For i = 1 To 25
			If bdKey(i) = KeyCode + 100 * Shift Then
				Select Case i
					Case 1
						KeyBoardShortCuts(System.Windows.Forms.Keys.D, 2)
					Case 2
						KeyBoardShortCuts(System.Windows.Forms.Keys.Q, 2)
					Case 3
						KeyBoardShortCuts(System.Windows.Forms.Keys.X, 2)
					Case 4
						KeyBoardShortCuts(System.Windows.Forms.Keys.Z, 2)
					Case 5
						KeyBoardShortCuts(System.Windows.Forms.Keys.L, 2)
					Case 6
						KeyBoardShortCuts(System.Windows.Forms.Keys.S, 2)
					Case 7
						KeyBoardShortCuts(System.Windows.Forms.Keys.M, 2)
					Case 8
						KeyBoardShortCuts(System.Windows.Forms.Keys.W, 2)
					Case 9
						KeyBoardShortCuts(System.Windows.Forms.Keys.Left, 1)
					Case 10
						KeyBoardShortCuts(System.Windows.Forms.Keys.Right, 1)
					Case 11
						KeyBoardShortCuts(System.Windows.Forms.Keys.Up, 1)
					Case 12
						KeyBoardShortCuts(System.Windows.Forms.Keys.Down, 1)
					Case 13
						KeyBoardShortCuts(System.Windows.Forms.Keys.Left, 0)
					Case 14
						KeyBoardShortCuts(System.Windows.Forms.Keys.Right, 0)
					Case 15
						KeyBoardShortCuts(System.Windows.Forms.Keys.Up, 0)
					Case 16
						KeyBoardShortCuts(System.Windows.Forms.Keys.Down, 0)
					Case 17
						KeyBoardShortCuts(System.Windows.Forms.Keys.J, 0)
					Case 18
						KeyBoardShortCuts(System.Windows.Forms.Keys.L, 0)
					Case 19
						KeyBoardShortCuts(System.Windows.Forms.Keys.T, 0)
					Case 20
						KeyBoardShortCuts(System.Windows.Forms.Keys.K, 0)
					Case 21
						KeyBoardShortCuts(System.Windows.Forms.Keys.Z, 0)
					Case 22
						KeyBoardShortCuts(System.Windows.Forms.Keys.S, 0)
					Case 23
						KeyBoardShortCuts(System.Windows.Forms.Keys.O, 0)
					Case 24
						KeyBoardShortCuts(System.Windows.Forms.Keys.I, 0)
					Case 25
						KeyBoardShortCuts(System.Windows.Forms.Keys.E, 0)
				End Select
			End If
		Next i
		If KeyCode = System.Windows.Forms.Keys.Return Or KeyCode = System.Windows.Forms.Keys.Escape Then KeyBoardShortCuts(KeyCode, 0)
	End Sub
	
	'UPGRADE_NOTE: Tome was upgraded to Tome_Renamed. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="A9E4979A-37FA-4718-9994-97DD76ED70A7"'
	Private Sub KeyBoardShortCuts(ByRef KeyCode As Short, ByRef Shift As Short)
		Dim Tome_Renamed As Object
		If (Shift And VB6.ShiftConstants.CtrlMask) > 0 Then
			Select Case KeyCode
				Case System.Windows.Forms.Keys.D
					GlobalDebugMode = 1
				Case System.Windows.Forms.Keys.Q
					GameEnd(True)
				Case System.Windows.Forms.Keys.S
					TomeSavesLoad()
					TomeAction = bdTomeSaveAs
					TomeSavesList(0)
				Case System.Windows.Forms.Keys.L
					TomeSavesLoad()
					TomeAction = bdTomeSaves
					TomeSavesList(1)
				Case System.Windows.Forms.Keys.M
					PlayMusicState()
				Case System.Windows.Forms.Keys.W
					PlaySoundState()
				Case System.Windows.Forms.Keys.X
					GameEnd(True)
				Case System.Windows.Forms.Keys.Z
					GameSpeed(1)
			End Select
		ElseIf Shift Then 
			Select Case KeyCode
				Case System.Windows.Forms.Keys.Left
					ScrollMap(-1, -1)
				Case System.Windows.Forms.Keys.Right
					ScrollMap(1, 1)
				Case System.Windows.Forms.Keys.Up
					ScrollMap(1, -1)
				Case System.Windows.Forms.Keys.Down
					ScrollMap(-1, 1)
			End Select
		Else
			Select Case KeyCode
				Case System.Windows.Forms.Keys.Left
					'UPGRADE_WARNING: Couldn't resolve default property of object Tome.MapY. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
					'UPGRADE_WARNING: Couldn't resolve default property of object Tome.MapX. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
					MovePartySet(0, 0, Tome.MapX - 1, Tome.MapY)
				Case System.Windows.Forms.Keys.Right
					'UPGRADE_WARNING: Couldn't resolve default property of object Tome.MapY. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
					'UPGRADE_WARNING: Couldn't resolve default property of object Tome.MapX. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
					MovePartySet(0, 0, Tome.MapX + 1, Tome.MapY)
				Case System.Windows.Forms.Keys.Up
					'UPGRADE_WARNING: Couldn't resolve default property of object Tome.MapY. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
					'UPGRADE_WARNING: Couldn't resolve default property of object Tome.MapX. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
					MovePartySet(0, 0, Tome.MapX, Tome.MapY - 1)
				Case System.Windows.Forms.Keys.Down
					'UPGRADE_WARNING: Couldn't resolve default property of object Tome.MapY. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
					'UPGRADE_WARNING: Couldn't resolve default property of object Tome.MapX. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
					MovePartySet(0, 0, Tome.MapX, Tome.MapY + 1)
				Case System.Windows.Forms.Keys.W ' Wait
					If picGrid.Visible = True Then
						CombatWait()
					End If
				Case System.Windows.Forms.Keys.L ' Listen
					If picGrid.Visible = False Then
						DoorListen()
					End If
				Case System.Windows.Forms.Keys.T ' Talk
					TalkSetup()
				Case System.Windows.Forms.Keys.K ' Skills
					If CreatureWithTurn.HPNow > 0 Then
						MenuSkillLoad()
						MenuSkillShow(1)
					Else
						DialogDM(CreatureWithTurn.Name & " is dead.")
					End If
				Case System.Windows.Forms.Keys.Z ' Status
					InvTopIndex = 0
					InventoryShow(bdInvStatus)
					picInventory.Visible = True
					picInventory.BringToFront()
					picInventory.Focus()
				Case System.Windows.Forms.Keys.E ' Equip
					InvTopIndex = 0
					InventoryShow(bdInvItems)
					picInventory.Visible = True
					picInventory.BringToFront()
					picInventory.Focus()
				Case System.Windows.Forms.Keys.S
					MenuNow = bdMenuSearch
					DoAction(bdMenuActionDefault)
				Case System.Windows.Forms.Keys.I
					MenuNow = bdMenuInventory
					DoAction(bdMenuActionDefault)
				Case System.Windows.Forms.Keys.J
					Select Case JournalMode
						Case 1 ' Quests
							JournalQuestsLoad()
						Case 2 ' Map
						Case Else ' Journal Entries
							JournalEntryLoad()
					End Select
					JournalShow()
				Case System.Windows.Forms.Keys.O
					'UPGRADE_WARNING: Couldn't resolve default property of object Tome.MapY. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
					'UPGRADE_WARNING: Couldn't resolve default property of object Tome.MapX. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
					DoorOpen(Tome.MapX, Tome.MapY)
					'            Case vbKeyD
				Case System.Windows.Forms.Keys.Return
					If ButtonList(0) = "Save" Then
						ConvoAction = 1
					Else
						ConvoAction = 0
					End If
				Case System.Windows.Forms.Keys.Escape
					If picGrid.Visible = True And (MenuNow = bdMenuTargetCreature Or MenuNow = bdMenuTargetAny Or MenuNow = bdMenuTargetParty) Then
						ConvoAction = 0
						MenuNow = bdMenuDefault
						CombatDraw()
					Else
						MenuShow()
					End If
			End Select
		End If
	End Sub
	
	Private Sub ExamineItem(ByRef ItemX As Item)
		Dim NoFail, Found As Short
		Dim TriggerX As Trigger
		Dim ItemZ As Item
		'UPGRADE_NOTE: Text was upgraded to Text_Renamed. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="A9E4979A-37FA-4718-9994-97DD76ED70A7"'
		Dim Text_Renamed As String
		Dim c As Short
		CreatureNow = CreatureWithTurn
		ItemNow = ItemX
		NoFail = FireTriggers(CreatureNow, bdPreExamine)
		CreatureNow = CreatureWithTurn
		ItemNow = ItemX
		NoFail = FireTriggers(ItemX, bdPreExamine)
		If NoFail Then
			If ItemX.UseDescription = True Then
				DialogItem(ItemX, (ItemX.Comments))
			ElseIf ItemX.WearType < 10 Then  ' Potential Armor
				Select Case ItemX.WearType
					Case 0 ' Body
						Text_Renamed = "on the body"
					Case 1 ' Helm
						Text_Renamed = "on the head"
					Case 2 ' Glove
						Text_Renamed = "on the hand"
					Case 3 ' Bracelet
						Text_Renamed = "on the wrist"
					Case 4 ' BackPack
						Text_Renamed = "on the back"
					Case 5 ' Shield
						Text_Renamed = "on the arm"
					Case 6 ' Boots
						Text_Renamed = "on the feet"
					Case 7 ' Necklace
						Text_Renamed = "around the neck"
					Case 8 ' Belt
						Text_Renamed = "around the waist"
					Case 9 ' Ring
						Text_Renamed = "on a finger"
				End Select
				DialogItem(ItemX, "This " & ItemX.Name & " is worn " & Text_Renamed & " and absorbs up to " & ItemX.Resistance & "% damage.")
			ElseIf ItemX.Damage > 0 Then  ' Potential Weapon
				Text_Renamed = (ItemX.Damage - 1) Mod 5 + 1 & "d" & Int(((ItemX.Damage - 1) Mod 25) / 5) * 2 + 4
				If ItemX.Damage - 1 > 24 Then
					Text_Renamed = Text_Renamed & "+" & Int((ItemX.Damage - 1) / 25)
				End If
				If ItemX.WearType = 11 Then ' Two-Handed
					DialogItem(ItemX, "This " & ItemX.NameText & " requires two hands to use and does " & Text_Renamed & " in damage.")
				Else
					DialogItem(ItemX, "This " & ItemX.NameText & " does " & Text_Renamed & " in damage.")
				End If
			ElseIf ItemX.Capacity = 0 Then 
				DialogItem(ItemX, CreatureWithTurn.Name & " sees nothing unusual about the " & ItemX.Name & ".")
			End If
			' Open the Item (if applicable)
			If ItemX.Capacity > 0 Then
				' Unlock the Item if Locked
				If SetUpItemFamily((ItemX.Family)) = "Locked" Or SetUpItemFamily((ItemX.Family)) = "Jammed Shut" Then
					' Try to find a key on the CreatureWithTurn
					Found = False
					For	Each ItemZ In CreatureWithTurn.Items
						If ItemZ.KeyBits > 0 Then
							' Make sure KeyBits match, not the same Item and the Key does not have Capacity (so is not another chest with same lock).
							If (ItemZ.KeyBits And ItemX.KeyBits) = ItemZ.KeyBits And (ItemZ.Index <> ItemX.Index Or ItemZ.Name <> ItemX.Name) And ItemZ.Capacity = 0 Then
								Found = True
								Exit For
							End If
						End If
					Next ItemZ
					' If found a key, open the Item
					If Found = True Then
						ItemX.Family = bdNone
						DialogDM(CreatureWithTurn.Name & " unlocked the " & ItemX.Name & " with the " & ItemZ.Name & ".")
					Else
						' Else try to pick the lock or bash it open
						GlobalPickLockChance = 0
						' If Jammed Shut, then chance is -100
						If SetUpItemFamily((ItemX.Family)) = "Jammed Shut" Then
							GlobalPickLockChance = -100
						End If
						CreatureNow = CreatureWithTurn
						ItemNow = CreatureWithTurn.ItemInHand
						ItemTarget = ItemX
						NoFail = FireTriggers(CreatureNow, bdPrePickLock)
						For	Each ItemNow In CreatureWithTurn.Items
							If ItemNow.IsReady = True Then
								NoFail = FireTriggers(ItemNow, bdPrePickLock)
							End If
						Next ItemNow
						CreatureNow = CreatureWithTurn
						ItemNow = ItemX
						NoFail = FireTriggers(ItemNow, bdPreUnlock)
						If GlobalPickLockChance > 0 Then
							DialogSetUp(modGameGeneral.DLGTYPE.bdDlgNoReply)
							DialogSetButton(3, "Bash")
							DialogSetButton(2, "Pick")
							DialogSetButton(1, "Done")
						Else
							DialogSetUp(modGameGeneral.DLGTYPE.bdDlgNoReply)
							DialogSetButton(2, "Bash")
							DialogSetButton(1, "Done")
						End If
						DialogShow("DM", "The " & ItemX.Name & " is locked.")
						If ConvoAction = 0 Then
							' If you attempt to bash, then you could damage the contents
							For	Each ItemZ In ItemX.Items
								If Mid(ItemX.Name, 1, 7) = "Damaged" Then
									ItemX.RemoveItem("I" & ItemZ.Index)
									ItemX.AddItem.Name = "Damaged Junk"
								ElseIf SetUpItemFamily((ItemZ.Family)) = "Fragile" Then 
									If Int(Rnd() * 100) < 50 Then
										ItemZ.Value = 1
										ItemZ.Name = "Damaged " & ItemZ.Name
									End If
								ElseIf SetUpItemFamily((ItemZ.Family)) <> "Durable" And ItemZ.IsMoney = False Then 
									If Int(Rnd() * 100) < 10 Then
										ItemZ.Value = 1
										ItemZ.Name = "Damaged " & ItemZ.NameText
									End If
								End If
							Next ItemZ
							If Int(Rnd() * 100) < CreatureWithTurn.Strength * 2 Then
								DialogDM(CreatureWithTurn.Name & " successfully bashed open the " & ItemX.Name & ".")
								ItemX.Family = bdNone
							Else
								DialogDM(CreatureWithTurn.Name & " failed to bash open the " & ItemX.Name & ".")
							End If
						ElseIf GlobalPickLockChance > 0 And ConvoAction = 1 Then 
							If Int(Rnd() * 100) < GlobalPickLockChance Then
								DialogDM(CreatureWithTurn.Name & " successfully picked the lock on the " & ItemX.Name & ".")
								ItemX.Family = bdNone
							ElseIf Int(Rnd() * 100) > GlobalPickLockChance Then 
								DialogDM(CreatureWithTurn.Name & " failed to pick the lock and jams shut the " & ItemX.Name & "!")
								ItemX.Family = 16 ' Jammed Shut
							Else
								DialogDM(CreatureWithTurn.Name & " failed to pick the lock on the " & ItemX.Name & ".")
							End If
						End If
					End If
				End If
				' Open the Item if not Locked
				If SetUpItemFamily((ItemX.Family)) <> "Locked" And SetUpItemFamily((ItemX.Family)) <> "Jammed Shut" Then
					NoFail = FireTriggers(ItemX, bdPreOpen)
					If NoFail Then
						InventoryContainerShow(ItemX)
					End If
				End If
			End If
			DialogHide()
			CreatureNow = CreatureWithTurn
			ItemNow = ItemX
			NoFail = FireTriggers(ItemX, bdPostExamine)
			CreatureNow = CreatureWithTurn
			ItemNow = ItemX
			NoFail = FireTriggers(CreatureNow, bdPostExamine)
		End If
	End Sub
	
	Private Sub InventoryContainerShow(ByRef ContainerX As Item)
		Dim c, yOffSet As Short
		Dim rc As Integer
		Dim ItemX As Item
		'UPGRADE_NOTE: Height was upgraded to Height_Renamed. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="A9E4979A-37FA-4718-9994-97DD76ED70A7"'
		'UPGRADE_NOTE: Width was upgraded to Width_Renamed. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="A9E4979A-37FA-4718-9994-97DD76ED70A7"'
		Dim Width_Renamed, X, Y, Height_Renamed As Short
		'UPGRADE_WARNING: Screen property Screen.MousePointer has a new behavior. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6BA9B8D2-2A32-4B6E-8D36-44949974A5B4"'
		System.Windows.Forms.Cursor.Current = System.Windows.Forms.Cursors.WaitCursor
		' If a Container is already open, close it first
		If picContainer.Visible = True Then
			InventoryContainerClose()
		End If
		' Set up the new container
		InvContainer = ContainerX
		InvContainer.Selected = True
		'UPGRADE_ISSUE: PictureBox method picContainer.Cls was not upgraded. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="CC4C7EC0-C903-48FC-ACCC-81861D12DA4A"'
		picContainer.Cls()
		' Show header
		ShowText(picContainer, 0, 12, picContainer.Width, 14, bdFontElixirWhite, (InvContainer.Name), True, False)
		' Clear current locations
		For c = 0 To 80
			InvC(c) = -1
		Next c
		' Show all Items (those with a spot first, without a spot second)
		InvTooMany = False
		For	Each ItemX In InvContainer.Items
			If ItemX.Selected = False Then
				LoadItemPic(ItemX)
				InventoryFindSpot(ItemX, 1, 80, 8, InvC)
				' Draw Item in Inventory (if have a spot to display)
				If ItemX.InvSpot > 0 Then
					' Configure Width and Height
					Width_Renamed = ItemPicWidth(ItemX.Pic)
					Height_Renamed = ItemPicHeight(ItemX.Pic)
					' Set location on Search box
					X = 23 + ((ItemX.InvSpot - 1) Mod 8) * 34
					Y = 45 + Int((ItemX.InvSpot - 1) / 8) * 34
					' Paint the Item picture
					'UPGRADE_ISSUE: PictureBox property picItem.hdc was not upgraded. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="CC4C7EC0-C903-48FC-ACCC-81861D12DA4A"'
					'UPGRADE_ISSUE: PictureBox property picContainer.hdc was not upgraded. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="CC4C7EC0-C903-48FC-ACCC-81861D12DA4A"'
					rc = BitBlt(picContainer.hdc, X, Y, Width_Renamed, Height_Renamed - yOffSet, picItem.hdc, 64 * ItemX.Pic - 64, 96, SRCAND)
					'UPGRADE_ISSUE: PictureBox property picItem.hdc was not upgraded. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="CC4C7EC0-C903-48FC-ACCC-81861D12DA4A"'
					'UPGRADE_ISSUE: PictureBox property picContainer.hdc was not upgraded. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="CC4C7EC0-C903-48FC-ACCC-81861D12DA4A"'
					rc = BitBlt(picContainer.hdc, X, Y, Width_Renamed, Height_Renamed - yOffSet, picItem.hdc, 64 * ItemX.Pic - 64, 0, SRCPAINT)
					If ItemX.Capacity > 0 And yOffSet = 0 Then
						ShowText(picContainer, X, Y, Width_Renamed, 10, bdFontSmallWhite, "(" & ItemX.Items.Count() & ")", 1, False)
					End If
					If ItemX.Count > 1 Then
						ShowText(picContainer, X, Y + Height_Renamed - 10, Width_Renamed, 10, bdFontSmallWhite, CStr(ItemX.Count), True, False)
					End If
				Else
					InvTooMany = True
				End If
			End If
		Next ItemX
		' Show Buttons
		If InvTooMany = True Then
			ShowButton(picContainer, picSearch.ClientRectangle.Width - 294, picSearch.ClientRectangle.Height - 30, "More", False)
		End If
		ShowButton(picContainer, picSearch.ClientRectangle.Width - 198, picSearch.ClientRectangle.Height - 30, "Empty", False)
		ShowButton(picContainer, picSearch.ClientRectangle.Width - 102, picSearch.ClientRectangle.Height - 30, "Done", False)
		picContainer.Left = 0
		picContainer.Top = Me.ClientRectangle.Height - 48 - picMenu.ClientRectangle.Height - picContainer.ClientRectangle.Height
		' Search gets the focus however
		picContainer.Visible = True
		picContainer.BringToFront()
		picContainer.Focus()
		'UPGRADE_WARNING: Screen property Screen.MousePointer has a new behavior. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6BA9B8D2-2A32-4B6E-8D36-44949974A5B4"'
		System.Windows.Forms.Cursor.Current = System.Windows.Forms.Cursors.Default
	End Sub
	
	'UPGRADE_NOTE: Map was upgraded to Map_Renamed. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="A9E4979A-37FA-4718-9994-97DD76ED70A7"'
	'UPGRADE_NOTE: Tome was upgraded to Tome_Renamed. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="A9E4979A-37FA-4718-9994-97DD76ED70A7"'
	Private Function DropItem(ByRef FromChar As Creature, ByRef FromList As Short, ByRef FromContainer As Short, ByRef ItemToDrop As Item) As Short
		Dim Tome_Renamed As Object
		Dim Map_Renamed As Object
		Dim c, NoFail As Short
		Dim ItemX As Item
		' Set up for Triggers
		CreatureNow = FromChar
		ItemNow = ItemToDrop
		' Unready (if necessary)
		DropItem = UnReadyItem(FromChar, ItemToDrop)
		If DropItem = False Then
			Exit Function
		End If
		CreatureNow = FromChar
		ItemNow = ItemToDrop
		NoFail = FireTriggers(ItemToDrop, bdPreDropItem)
		If NoFail Then
			' If above checks ok, drop item between characters/lists
			If FromList = bdInvContainer Then
				'            Containers("I" & FromContainer - 1).ItemX.RemoveItem "I" & ItemToDrop.Index
			Else
				FromChar.RemoveItem("I" & ItemToDrop.Index)
			End If
			' Create encounter on map if necessary
			'UPGRADE_WARNING: Couldn't resolve default property of object Tome.MapY. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
			'UPGRADE_WARNING: Couldn't resolve default property of object Tome.MapX. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
			'UPGRADE_WARNING: Couldn't resolve default property of object Map.EncPointer. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
			If Map.EncPointer(Tome.MapX, Tome.MapY) = 0 Then
				'UPGRADE_WARNING: Couldn't resolve default property of object Map.AddEncounter. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
				EncounterNow = Map.AddEncounter
				'UPGRADE_WARNING: Couldn't resolve default property of object Tome.MapY. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
				'UPGRADE_WARNING: Couldn't resolve default property of object Tome.MapX. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
				'UPGRADE_WARNING: Couldn't resolve default property of object Map.EncPointer. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
				Map.EncPointer(Tome.MapX, Tome.MapY) = EncounterNow.Index
			End If
			ItemX = EncounterNow.AddItem
			ItemX.Copy(ItemToDrop)
			'UPGRADE_WARNING: Couldn't resolve default property of object Map. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
			PositionItem(Map, EncounterNow, ItemX)
			NoFail = FireTriggers(ItemToDrop, bdPostDropItem)
			DropItem = True
		Else
			DropItem = False
		End If
	End Function
	
	Private Function DestroyItem(ByRef FromChar As Creature, ByRef ItemToDestroy As Item) As Short
		' Unready (if necessary)
		If UnReadyItem(FromChar, ItemToDestroy) Then
			' If above checks ok, drop item between characters/lists
			FromChar.RemoveItem("I" & ItemToDestroy.Index)
			DestroyItem = True
		Else
			DestroyItem = False
		End If
	End Function
	
	Private Sub GameEnd(ByRef AskToEnd As Short)
		Dim c As Short
		' Ask if really wish to stop
		picMainMenu.Visible = False
		If AskToEnd = True Then
			DialogSetUp(modGameGeneral.DLGTYPE.bdDlgNoReply)
			DialogSetButton(1, "No")
			DialogSetButton(2, "Yes")
			DialogShow("DM", "Do you really wish to quit this game?")
			DialogHide()
		End If
		If ConvoAction = 0 Or AskToEnd = False Then
			GameClear()
			' Back to Main Menu
			'UPGRADE_WARNING: Screen property Screen.MousePointer has a new behavior. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6BA9B8D2-2A32-4B6E-8D36-44949974A5B4"'
			System.Windows.Forms.Cursor.Current = System.Windows.Forms.Cursors.Default
			picMenu.Visible = False
			Frozen = True
			'UPGRADE_ISSUE: Form method frmMain.Cls was not upgraded. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="CC4C7EC0-C903-48FC-ACCC-81861D12DA4A"'
			frmMain.Cls()
			BorderDrawAll(False)
			TomeMenu = bdNone
			picMainMenu.Visible = True
			Call PlayMusic(WorldNow.IntroMusic, Me, WorldNow.MusicFolder)
			Me.Refresh()
		Else
			Frozen = False
		End If
	End Sub
	
	Private Sub GameClear()
		Dim c As Short
		' If in Combat, then End it
		If picGrid.Visible = True Then
			CombatEnd(False)
		End If
		' Stop current song that is playing
		'UPGRADE_WARNING: Screen property Screen.MousePointer has a new behavior. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6BA9B8D2-2A32-4B6E-8D36-44949974A5B4"'
		System.Windows.Forms.Cursor.Current = System.Windows.Forms.Cursors.WaitCursor
		MessageShow("Clearing Map.", 0)
		'MediaPlayerMusic.Stop
		oGameMusic.StopPlay()
		Call oGameMusic.mciClose(True)
		'    c = mciSendString("close songnow", 0&, 0, 0)
		' Clear Map and CreatureWithTurn Pics
		'UPGRADE_ISSUE: PictureBox method picMap.Cls was not upgraded. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="CC4C7EC0-C903-48FC-ACCC-81861D12DA4A"'
        picMap.Controls.Clear()
		' Hide any stray DialogBrief
		picDialogBrief.Visible = False
		tmrDialogBrief.Enabled = False
		' Reset Tome, Area and Map
		Tome = New Tome
		MessageShow("Clearing Map..", 0)
		Area = New Area
		MessageShow("Clearing Map...", 0)
		Map = New Map
		MessageShow("Clearing Map....", 0)
		EncounterNow = New Encounter
	End Sub
	
	Public Sub PlaySFXItem(ByRef CreatureX As Creature)
		Dim ItemX As Item
		ItemX = CreatureX.ItemInHand
		If ItemX.IsShooter = True And CreatureX.ItemMaxRange > 1 Then
			If ItemX.ShootType = 0 Or ItemX.ShootType = 1 Then ' Arrow (Short/Long) for Ammo
				Call PlaySoundFile("arrow.wav", Tome)
				PlaySFX("arrow.bmp", bdSFXFromCenter, bdSFXFling, bdSFXFast, 0)
			ElseIf ItemX.ShootType = 2 Or ItemX.ShootType = 3 Then  ' Sling/Illusion
			ElseIf ItemX.ShootType = 4 Or ItemX.ShootType = 5 Then  ' Gun/Bullet
			ElseIf ItemX.ShootType = 6 Or ItemX.ShootType = 7 Then  ' Energy High/Low
			End If
		End If
	End Sub
	
	Public Sub PlaySFX(ByRef PictureFile As String, ByRef FromLoc As Short, ByRef WithStyle As Short, ByRef Speed As Short, ByRef Frames As Short)
		' Plays SFX in Combat. Flip = 0 if left to right, Or 1 if right to left
		' Speed is 0 (Fast), 1 (Medium) or 2 (Slow)
		Dim rc, c, i, Pic As Short
		'UPGRADE_NOTE: Width was upgraded to Width_Renamed. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="A9E4979A-37FA-4718-9994-97DD76ED70A7"'
		'UPGRADE_NOTE: Height was upgraded to Height_Renamed. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="A9E4979A-37FA-4718-9994-97DD76ED70A7"'
		Dim Height_Renamed, Width_Renamed, Flip As Short
		Dim ToX, FromX, FromY, ToY As Short
		Dim FromWidth, ToWidth As Double
		Dim RateX, Cels, RateY As Short
		' If not in Combat, then return immediately
		If picGrid.Visible = False Then
			Exit Sub
		End If
		' Load SFX picture file. If can't find it, then exit.
		Pic = LoadSFXPic(PictureFile)
		If Pic < 0 Then
			Exit Sub
		End If
		' Set animate parameters
		Frames = Greatest(Frames, 1)
		Width_Renamed = (picSFXPic(Pic).Width / Frames) / 2 : Height_Renamed = picSFXPic(Pic).Height / 2
		picTmp.Width = Width_Renamed : picTmp.Height = Height_Renamed
		' Compute FromX,Y and ToX,Y
		GridToCursor(FromX, FromY, CreatureWithTurn.Row, CreatureWithTurn.Col, FromWidth)
		GridToCursor(ToX, ToY, CreatureTarget.Row, CreatureTarget.Col, ToWidth)
		ToY = ToY - (picCPic(CreatureTarget.Pic).Height + Height_Renamed) / 4
		ToX = ToX + (ToWidth - Width_Renamed) / 2
		If FromLoc = bdSFXFromHead Then
			FromY = FromY - (picCPic(CreatureWithTurn.Pic).Height / 2) + CreatureWithTurn.FaceTop * PicSize
			FromX = FromX + CreatureWithTurn.FaceLeft * PicSize - Width_Renamed
		Else
			FromY = FromY - Int(picCPic(CreatureWithTurn.Pic).Height / 4)
			FromX = FromX + (FromWidth - Width_Renamed) / 2
		End If
		Cels = Greatest(CShort(System.Math.Sqrt((ToX - FromX) ^ 2 + (ToY - FromY) ^ 2) / Width_Renamed) * 2, 1) * 2
		RateX = (ToX - FromX) / Cels
		RateY = (ToY - FromY) / Cels
		' If CreatureTarget to the left of CreatureWithTurn, then Flip
		If CreatureTarget.Col < CreatureWithTurn.Col Then
			Flip = picSFXPic(Pic).Width / 2
		Else
			Flip = 0
		End If
		' Animate to Combat screen
		Select Case WithStyle
			Case bdSFXFling
				i = 0
				Cels = Cels / 2 : RateX = RateX * 2 : RateY = RateY * 2
				For c = 0 To Cels
					'UPGRADE_ISSUE: PictureBox property picGrid.hdc was not upgraded. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="CC4C7EC0-C903-48FC-ACCC-81861D12DA4A"'
					'UPGRADE_ISSUE: PictureBox property picTmp.hdc was not upgraded. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="CC4C7EC0-C903-48FC-ACCC-81861D12DA4A"'
					rc = BitBlt(picTmp.hdc, 0, 0, Width_Renamed, Height_Renamed, picGrid.hdc, FromX + c * RateX, FromY + c * RateY, SRCCOPY)
					picTmp.Refresh()
					i = LoopNumber(0, Frames - 1, i, 1)
					'UPGRADE_ISSUE: PictureBox property picSFXPic.hdc was not upgraded. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="CC4C7EC0-C903-48FC-ACCC-81861D12DA4A"'
					'UPGRADE_ISSUE: PictureBox property picGrid.hdc was not upgraded. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="CC4C7EC0-C903-48FC-ACCC-81861D12DA4A"'
					rc = BitBlt(picGrid.hdc, FromX + c * RateX, FromY + c * RateY, Width_Renamed, Height_Renamed, picSFXPic(Pic).hdc, i * Width_Renamed + Flip, Height_Renamed, SRCAND)
					'UPGRADE_ISSUE: PictureBox property picSFXPic.hdc was not upgraded. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="CC4C7EC0-C903-48FC-ACCC-81861D12DA4A"'
					'UPGRADE_ISSUE: PictureBox property picGrid.hdc was not upgraded. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="CC4C7EC0-C903-48FC-ACCC-81861D12DA4A"'
					rc = BitBlt(picGrid.hdc, FromX + c * RateX, FromY + c * RateY, Width_Renamed, Height_Renamed, picSFXPic(Pic).hdc, i * Width_Renamed + Flip, 0, SRCPAINT)
					picGrid.Refresh()
					Select Case Speed
						Case 0 ' Fast
							GameDelay(5000)
						Case 1 ' Medium
							GameDelay(1700)
						Case 2 ' Slow
							GameDelay(500)
					End Select
					'UPGRADE_ISSUE: PictureBox property picTmp.hdc was not upgraded. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="CC4C7EC0-C903-48FC-ACCC-81861D12DA4A"'
					'UPGRADE_ISSUE: PictureBox property picGrid.hdc was not upgraded. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="CC4C7EC0-C903-48FC-ACCC-81861D12DA4A"'
					rc = BitBlt(picGrid.hdc, FromX + c * RateX, FromY + c * RateY, Width_Renamed, Height_Renamed, picTmp.hdc, 0, 0, SRCCOPY)
				Next c
			Case bdSFXStream
				i = 0
				For c = 0 To Cels
					i = LoopNumber(0, Frames - 1, i, 1)
					'UPGRADE_ISSUE: PictureBox property picSFXPic.hdc was not upgraded. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="CC4C7EC0-C903-48FC-ACCC-81861D12DA4A"'
					'UPGRADE_ISSUE: PictureBox property picGrid.hdc was not upgraded. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="CC4C7EC0-C903-48FC-ACCC-81861D12DA4A"'
					rc = BitBlt(picGrid.hdc, FromX + c * RateX, FromY + c * RateY, Width_Renamed, Height_Renamed, picSFXPic(Pic).hdc, i * Width_Renamed + Flip, Height_Renamed, SRCAND)
					'UPGRADE_ISSUE: PictureBox property picSFXPic.hdc was not upgraded. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="CC4C7EC0-C903-48FC-ACCC-81861D12DA4A"'
					'UPGRADE_ISSUE: PictureBox property picGrid.hdc was not upgraded. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="CC4C7EC0-C903-48FC-ACCC-81861D12DA4A"'
					rc = BitBlt(picGrid.hdc, FromX + c * RateX, FromY + c * RateY, Width_Renamed, Height_Renamed, picSFXPic(Pic).hdc, i * Width_Renamed + Flip, 0, SRCPAINT)
					picGrid.Refresh()
					Select Case Speed
						Case 0 ' Fast
							GameDelay(3000)
						Case 1 ' Medium
							GameDelay(1000)
						Case 2 ' Slow
							GameDelay(300)
					End Select
				Next c
				CombatDraw()
			Case bdSFXDown
				FromX = ToX : FromY = 0 : Flip = 0
				Cels = Greatest(CShort(System.Math.Sqrt((ToX - FromX) ^ 2 + (ToY - FromY) ^ 2) / Width_Renamed) * 2, 1)
				Cels = Cels * (Speed + 1)
				RateX = (ToX - FromX) / Cels
				RateY = (ToY - FromY) / Cels
				i = 0
				For c = 0 To Cels
					i = LoopNumber(0, Frames - 1, i, 1)
					'UPGRADE_ISSUE: PictureBox property picSFXPic.hdc was not upgraded. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="CC4C7EC0-C903-48FC-ACCC-81861D12DA4A"'
					'UPGRADE_ISSUE: PictureBox property picGrid.hdc was not upgraded. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="CC4C7EC0-C903-48FC-ACCC-81861D12DA4A"'
					rc = BitBlt(picGrid.hdc, FromX + c * RateX, FromY + c * RateY, Width_Renamed, Height_Renamed, picSFXPic(Pic).hdc, i * Width_Renamed + Flip, Height_Renamed, SRCAND)
					'UPGRADE_ISSUE: PictureBox property picSFXPic.hdc was not upgraded. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="CC4C7EC0-C903-48FC-ACCC-81861D12DA4A"'
					'UPGRADE_ISSUE: PictureBox property picGrid.hdc was not upgraded. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="CC4C7EC0-C903-48FC-ACCC-81861D12DA4A"'
					rc = BitBlt(picGrid.hdc, FromX + c * RateX, FromY + c * RateY, Width_Renamed, Height_Renamed, picSFXPic(Pic).hdc, i * Width_Renamed + Flip, 0, SRCPAINT)
					picGrid.Refresh()
					Select Case Speed
						Case 0 ' Fast
							GameDelay(1000)
						Case 1 ' Medium
							GameDelay(300)
						Case 2 ' Slow
							GameDelay(100)
					End Select
				Next c
				CombatDraw()
			Case bdSFXBurstHere
				For i = 0 To Frames - 1
					'UPGRADE_ISSUE: PictureBox property picGrid.hdc was not upgraded. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="CC4C7EC0-C903-48FC-ACCC-81861D12DA4A"'
					'UPGRADE_ISSUE: PictureBox property picTmp.hdc was not upgraded. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="CC4C7EC0-C903-48FC-ACCC-81861D12DA4A"'
					rc = BitBlt(picTmp.hdc, 0, 0, Width_Renamed, Height_Renamed, picGrid.hdc, FromX, FromY, SRCCOPY)
					picTmp.Refresh()
					'UPGRADE_ISSUE: PictureBox property picSFXPic.hdc was not upgraded. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="CC4C7EC0-C903-48FC-ACCC-81861D12DA4A"'
					'UPGRADE_ISSUE: PictureBox property picGrid.hdc was not upgraded. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="CC4C7EC0-C903-48FC-ACCC-81861D12DA4A"'
					rc = BitBlt(picGrid.hdc, FromX, FromY, Width_Renamed, Height_Renamed, picSFXPic(Pic).hdc, i * Width_Renamed + Flip, Height_Renamed, SRCAND)
					'UPGRADE_ISSUE: PictureBox property picSFXPic.hdc was not upgraded. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="CC4C7EC0-C903-48FC-ACCC-81861D12DA4A"'
					'UPGRADE_ISSUE: PictureBox property picGrid.hdc was not upgraded. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="CC4C7EC0-C903-48FC-ACCC-81861D12DA4A"'
					rc = BitBlt(picGrid.hdc, FromX, FromY, Width_Renamed, Height_Renamed, picSFXPic(Pic).hdc, i * Width_Renamed + Flip, 0, SRCPAINT)
					picGrid.Refresh()
					Select Case Speed
						Case 0 ' Fast
							GameDelay(1000)
						Case 1 ' Medium
							GameDelay(300)
						Case 2 ' Slow
							GameDelay(100)
					End Select
					'UPGRADE_ISSUE: PictureBox property picTmp.hdc was not upgraded. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="CC4C7EC0-C903-48FC-ACCC-81861D12DA4A"'
					'UPGRADE_ISSUE: PictureBox property picGrid.hdc was not upgraded. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="CC4C7EC0-C903-48FC-ACCC-81861D12DA4A"'
					rc = BitBlt(picGrid.hdc, FromX, FromY, Width_Renamed, Height_Renamed, picTmp.hdc, 0, 0, SRCCOPY)
				Next i
			Case bdSFXBurstThere
				For i = 0 To Frames - 1
					'UPGRADE_ISSUE: PictureBox property picGrid.hdc was not upgraded. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="CC4C7EC0-C903-48FC-ACCC-81861D12DA4A"'
					'UPGRADE_ISSUE: PictureBox property picTmp.hdc was not upgraded. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="CC4C7EC0-C903-48FC-ACCC-81861D12DA4A"'
					rc = BitBlt(picTmp.hdc, 0, 0, Width_Renamed, Height_Renamed, picGrid.hdc, ToX, ToY, SRCCOPY)
					picTmp.Refresh()
					'UPGRADE_ISSUE: PictureBox property picSFXPic.hdc was not upgraded. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="CC4C7EC0-C903-48FC-ACCC-81861D12DA4A"'
					'UPGRADE_ISSUE: PictureBox property picGrid.hdc was not upgraded. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="CC4C7EC0-C903-48FC-ACCC-81861D12DA4A"'
					rc = BitBlt(picGrid.hdc, ToX, ToY, Width_Renamed, Height_Renamed, picSFXPic(Pic).hdc, i * Width_Renamed + Flip, Height_Renamed, SRCAND)
					'UPGRADE_ISSUE: PictureBox property picSFXPic.hdc was not upgraded. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="CC4C7EC0-C903-48FC-ACCC-81861D12DA4A"'
					'UPGRADE_ISSUE: PictureBox property picGrid.hdc was not upgraded. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="CC4C7EC0-C903-48FC-ACCC-81861D12DA4A"'
					rc = BitBlt(picGrid.hdc, ToX, ToY, Width_Renamed, Height_Renamed, picSFXPic(Pic).hdc, i * Width_Renamed + Flip, 0, SRCPAINT)
					picGrid.Refresh()
					Select Case Speed
						Case 0 ' Fast
							GameDelay(1000)
						Case 1 ' Medium
							GameDelay(300)
						Case 2 ' Slow
							GameDelay(100)
					End Select
					'UPGRADE_ISSUE: PictureBox property picTmp.hdc was not upgraded. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="CC4C7EC0-C903-48FC-ACCC-81861D12DA4A"'
					'UPGRADE_ISSUE: PictureBox property picGrid.hdc was not upgraded. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="CC4C7EC0-C903-48FC-ACCC-81861D12DA4A"'
					rc = BitBlt(picGrid.hdc, ToX, ToY, Width_Renamed, Height_Renamed, picTmp.hdc, 0, 0, SRCCOPY)
				Next i
		End Select
		picGrid.Refresh()
		'UPGRADE_WARNING: Screen property Screen.MousePointer has a new behavior. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6BA9B8D2-2A32-4B6E-8D36-44949974A5B4"'
		System.Windows.Forms.Cursor.Current = System.Windows.Forms.Cursors.Default
	End Sub
	
	Private Sub ResizeForm()
		Dim X, Y As Short
		' Size and Place Menu
		picMenu.Width = 32 + 5 * (90 + 32) : picMenu.Height = 84
		picMenu.Top = Me.ClientRectangle.Height - picMenu.Height - 16
		picMenu.Left = (Me.ClientRectangle.Width - picMenu.Width) / 2
		picMenu.Visible = False
		' Load the Interface
		' OptionsLoad True
		GameInterfaceLoad(True)
		' Place and Size Borders (Not in Game)
		BorderDrawAll(False)
		' Place and Size Map
		picBox.Top = 16 : picBox.Left = 32
		picBox.Height = Me.ClientRectangle.Height - picMenu.Height - 64
		picBox.Width = Me.ClientRectangle.Width - 64
		picMap.Top = 0 : picMap.Left = 0
		picMap.Height = Int(picBox.Height / 48) * 48 + 48
		picMap.Width = Int(picBox.Width / 96) * 96 + 96
		' Place Combat Boxes
		picGrid.Top = 0 : picGrid.Left = 0
		picGrid.Height = picBox.Height : picGrid.Width = picBox.Width
		picToHit.Top = Me.ClientRectangle.Height - picToHit.Height - 16
		picToHit.Left = (Me.ClientRectangle.Width - picToHit.Width) / 2
		' Place Inventory Box
		picInventory.Top = (picBox.ClientRectangle.Height - picInventory.ClientRectangle.Height) / 2
		picInventory.Left = (Me.ClientRectangle.Width - picInventory.ClientRectangle.Width) / 2
		' Place Search box
		picSearch.Top = (picBox.ClientRectangle.Height - picSearch.ClientRectangle.Height) / 2
		picSearch.Left = (Me.ClientRectangle.Width - picSearch.ClientRectangle.Width) / 2
		' Place Talk box
		picTalk.Top = (picBox.ClientRectangle.Height - picTalk.ClientRectangle.Height) / 2
		picTalk.Left = (Me.ClientRectangle.Width - picTalk.ClientRectangle.Width) / 2
		' Place Convo box
		picConvo.Left = (Me.ClientRectangle.Width - picConvo.ClientRectangle.Width) / 2
		' Place BuySell box
		picBuySell.Top = (picBox.ClientRectangle.Height - picBuySell.ClientRectangle.Height) / 2
		picBuySell.Left = (Me.ClientRectangle.Width - picBuySell.ClientRectangle.Width) / 2
		' Place Journal Box
		picJournal.Top = (picBox.ClientRectangle.Height - picJournal.ClientRectangle.Height) / 2
		picJournal.Left = (Me.ClientRectangle.Width - picJournal.ClientRectangle.Width) / 2
		' Place Main Menu
		picMainMenu.Top = (Me.ClientRectangle.Height - picMainMenu.Height) / 2
		picMainMenu.Left = (Me.ClientRectangle.Width - picMainMenu.Width) / 2
		picMainMenu.BringToFront()
		' Place New Tome
		picTomeNew.Top = (Me.ClientRectangle.Height - picTomeNew.ClientRectangle.Height) / 2
		picTomeNew.Left = (Me.ClientRectangle.Width - picTomeNew.Width) / 2
		picTomeNew.BringToFront()
		' Refresh
		Me.Refresh()
	End Sub
	
	'UPGRADE_NOTE: Map was upgraded to Map_Renamed. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="A9E4979A-37FA-4718-9994-97DD76ED70A7"'
	Public Sub DrawMapAll()
		Dim Map_Renamed As Object
		' If TileSet is not loaded, load it now
		'UPGRADE_WARNING: Couldn't resolve default property of object Map.PictureFile. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
		If TileSetLoaded <> Map.PictureFile Then
			LoadTileSet(Darker, False)
			LoadTileSetMicro()
			'UPGRADE_WARNING: Couldn't resolve default property of object Map.Left. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
			MicroMapLeft = Map.Left
			'UPGRADE_WARNING: Couldn't resolve default property of object Map.Top. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
			MicroMapTop = Map.Top
		End If
		' Draw map
		'UPGRADE_WARNING: Screen property Screen.MousePointer has a new behavior. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6BA9B8D2-2A32-4B6E-8D36-44949974A5B4"'
		System.Windows.Forms.Cursor.Current = System.Windows.Forms.Cursors.WaitCursor
		DrawMapRegion(0, 0, picMap.ClientRectangle.Height, picMap.ClientRectangle.Width)
		picMap.Focus()
		'UPGRADE_WARNING: Screen property Screen.MousePointer has a new behavior. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6BA9B8D2-2A32-4B6E-8D36-44949974A5B4"'
		System.Windows.Forms.Cursor.Current = System.Windows.Forms.Cursors.Default
	End Sub
	
	Private Sub BorderDrawTop()
		Dim rc As Integer
		Dim c As Short
		' Top
		For c = 0 To Int(Me.ClientRectangle.Width / bdIntWidth)
			'UPGRADE_ISSUE: PictureBox property picMisc.hdc was not upgraded. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="CC4C7EC0-C903-48FC-ACCC-81861D12DA4A"'
			'UPGRADE_ISSUE: Form property frmMain.hdc was not upgraded. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="CC4C7EC0-C903-48FC-ACCC-81861D12DA4A"'
			rc = BitBlt(Me.hdc, c * bdIntWidth, 0, bdIntWidth, bdIntHeight, picMisc.hdc, 38, 180, SRCCOPY)
		Next c
		' Top Corners
		'UPGRADE_ISSUE: PictureBox property picMisc.hdc was not upgraded. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="CC4C7EC0-C903-48FC-ACCC-81861D12DA4A"'
		'UPGRADE_ISSUE: Form property frmMain.hdc was not upgraded. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="CC4C7EC0-C903-48FC-ACCC-81861D12DA4A"'
		rc = BitBlt(Me.hdc, 0, 0, bdIntHeight * 2, bdIntHeight, picMisc.hdc, 139, 180, SRCCOPY)
		'UPGRADE_ISSUE: PictureBox property picMisc.hdc was not upgraded. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="CC4C7EC0-C903-48FC-ACCC-81861D12DA4A"'
		'UPGRADE_ISSUE: Form property frmMain.hdc was not upgraded. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="CC4C7EC0-C903-48FC-ACCC-81861D12DA4A"'
		rc = BitBlt(Me.hdc, Me.ClientRectangle.Width - bdIntHeight * 2, 0, bdIntHeight * 2, bdIntHeight, picMisc.hdc, 139, 180, SRCCOPY)
	End Sub
	
	Private Sub BorderDrawBottom()
		Dim rc As Integer
		Dim c As Short
		' Bottom
		For c = 0 To Int(Me.ClientRectangle.Width / bdIntWidth)
			'UPGRADE_ISSUE: PictureBox property picMisc.hdc was not upgraded. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="CC4C7EC0-C903-48FC-ACCC-81861D12DA4A"'
			'UPGRADE_ISSUE: Form property frmMain.hdc was not upgraded. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="CC4C7EC0-C903-48FC-ACCC-81861D12DA4A"'
			rc = BitBlt(Me.hdc, c * bdIntWidth, Me.ClientRectangle.Height - bdIntHeight, bdIntWidth, bdIntHeight, picMisc.hdc, 38, 180, SRCCOPY)
		Next c
		' Bottom Corners
		'UPGRADE_ISSUE: PictureBox property picMisc.hdc was not upgraded. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="CC4C7EC0-C903-48FC-ACCC-81861D12DA4A"'
		'UPGRADE_ISSUE: Form property frmMain.hdc was not upgraded. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="CC4C7EC0-C903-48FC-ACCC-81861D12DA4A"'
		rc = BitBlt(Me.hdc, 0, Me.ClientRectangle.Height - bdIntHeight, bdIntHeight * 2, bdIntHeight, picMisc.hdc, 139, 180, SRCCOPY)
		'UPGRADE_ISSUE: PictureBox property picMisc.hdc was not upgraded. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="CC4C7EC0-C903-48FC-ACCC-81861D12DA4A"'
		'UPGRADE_ISSUE: Form property frmMain.hdc was not upgraded. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="CC4C7EC0-C903-48FC-ACCC-81861D12DA4A"'
		rc = BitBlt(Me.hdc, Me.ClientRectangle.Width - bdIntHeight * 2, Me.ClientRectangle.Height - bdIntHeight, bdIntHeight * 2, bdIntHeight, picMisc.hdc, 139, 180, SRCCOPY)
	End Sub
	
	'UPGRADE_NOTE: Map was upgraded to Map_Renamed. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="A9E4979A-37FA-4718-9994-97DD76ED70A7"'
	Private Sub BorderDrawSides(ByRef InGame As Short, ByRef ShowRune As Short)
		Dim Map_Renamed As Object
		Dim rc As Integer
		Dim fx, X, c, Y, fy As Short
		' Left and Right
		For c = 0 To Int(Me.ClientRectangle.Height / bdIntWidth)
			'UPGRADE_ISSUE: PictureBox property picMisc.hdc was not upgraded. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="CC4C7EC0-C903-48FC-ACCC-81861D12DA4A"'
			'UPGRADE_ISSUE: Form property frmMain.hdc was not upgraded. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="CC4C7EC0-C903-48FC-ACCC-81861D12DA4A"'
			rc = BitBlt(Me.hdc, 0, 16 + c * bdIntWidth, bdIntHeight * 2, VB6.PixelsToTwipsX(Width), picMisc.hdc, 139, 89, SRCCOPY)
			'UPGRADE_ISSUE: PictureBox property picMisc.hdc was not upgraded. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="CC4C7EC0-C903-48FC-ACCC-81861D12DA4A"'
			'UPGRADE_ISSUE: Form property frmMain.hdc was not upgraded. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="CC4C7EC0-C903-48FC-ACCC-81861D12DA4A"'
			rc = BitBlt(Me.hdc, Me.ClientRectangle.Width - bdIntHeight * 2, 16 + c * bdIntWidth, bdIntHeight * 2, bdIntWidth, picMisc.hdc, 139, 89, SRCCOPY)
		Next c
		' If InGame then Draw Rune Pool
		'UPGRADE_WARNING: Couldn't resolve default property of object Map.IsNoRunes. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
		If InGame = True And Map.IsNoRunes = False Then
			' Draw number of Runes available
			For c = 0 To 19
				If c > 9 Then
					X = Me.ClientRectangle.Width - 32 : Y = (c - 10) * 32 + 16
					fx = c * 32 - 320
					fy = 32
				Else
					X = 0 : Y = c * 32 + 16
					fx = c * 32
					fy = 0
				End If
				If c = ShowRune Then
					'UPGRADE_ISSUE: PictureBox property picRuneSet.hdc was not upgraded. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="CC4C7EC0-C903-48FC-ACCC-81861D12DA4A"'
					'UPGRADE_ISSUE: Form property frmMain.hdc was not upgraded. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="CC4C7EC0-C903-48FC-ACCC-81861D12DA4A"'
					rc = BitBlt(Me.hdc, X, Y, 32, 32, picRuneSet.hdc, fx, fy + 128, SRCCOPY)
					'UPGRADE_WARNING: Couldn't resolve default property of object Map.Runes. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
					ShowText(Me, X, Y + 22, 32, 10, bdFontSmallWhite, VB6.Format(Map.Runes(c)), True, False)
					'UPGRADE_WARNING: Couldn't resolve default property of object Map.Runes. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
				ElseIf Map.Runes(c) > 0 Then 
					'UPGRADE_ISSUE: PictureBox property picRuneSet.hdc was not upgraded. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="CC4C7EC0-C903-48FC-ACCC-81861D12DA4A"'
					'UPGRADE_ISSUE: Form property frmMain.hdc was not upgraded. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="CC4C7EC0-C903-48FC-ACCC-81861D12DA4A"'
					rc = BitBlt(Me.hdc, X, Y, 32, 32, picRuneSet.hdc, fx, fy, SRCCOPY)
					'UPGRADE_WARNING: Couldn't resolve default property of object Map.Runes. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
					ShowText(Me, X, Y + 22, 32, 10, bdFontSmallWhite, VB6.Format(Map.Runes(c)), True, False)
				Else
					'UPGRADE_ISSUE: PictureBox property picRuneSet.hdc was not upgraded. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="CC4C7EC0-C903-48FC-ACCC-81861D12DA4A"'
					'UPGRADE_ISSUE: Form property frmMain.hdc was not upgraded. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="CC4C7EC0-C903-48FC-ACCC-81861D12DA4A"'
					rc = BitBlt(Me.hdc, X, Y, 32, 32, picRuneSet.hdc, fx, fy + 64, SRCCOPY)
					ShowText(Me, X, Y + 22, 32, 10, bdFontSmallWhite, "-", True, False)
				End If
			Next c
		End If
		' Bottom Corners
		'UPGRADE_ISSUE: PictureBox property picMisc.hdc was not upgraded. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="CC4C7EC0-C903-48FC-ACCC-81861D12DA4A"'
		'UPGRADE_ISSUE: Form property frmMain.hdc was not upgraded. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="CC4C7EC0-C903-48FC-ACCC-81861D12DA4A"'
		rc = BitBlt(Me.hdc, 0, Me.ClientRectangle.Height - bdIntHeight, bdIntHeight * 2, bdIntHeight, picMisc.hdc, 139, 180, SRCCOPY)
		'UPGRADE_ISSUE: PictureBox property picMisc.hdc was not upgraded. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="CC4C7EC0-C903-48FC-ACCC-81861D12DA4A"'
		'UPGRADE_ISSUE: Form property frmMain.hdc was not upgraded. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="CC4C7EC0-C903-48FC-ACCC-81861D12DA4A"'
		rc = BitBlt(Me.hdc, Me.ClientRectangle.Width - bdIntHeight * 2, Me.ClientRectangle.Height - bdIntHeight, bdIntHeight * 2, bdIntHeight, picMisc.hdc, 139, 180, SRCCOPY)
	End Sub
	
	Private Sub BorderDrawButtons(ByRef ButtonDown As Short, Optional ByVal Ymouse As Short = 0)
		'UPGRADE_NOTE: Text was upgraded to Text_Renamed. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="A9E4979A-37FA-4718-9994-97DD76ED70A7"'
		Dim rc As Integer
		Dim Y, c, X As Short
		Dim Text_Renamed As String
		Dim sRune() As String
		Dim PauseTime, Start As Single
		' Middle
		For c = 0 To Int(Me.ClientRectangle.Width / bdIntWidth)
			'UPGRADE_ISSUE: PictureBox property picMisc.hdc was not upgraded. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="CC4C7EC0-C903-48FC-ACCC-81861D12DA4A"'
			'UPGRADE_ISSUE: Form property frmMain.hdc was not upgraded. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="CC4C7EC0-C903-48FC-ACCC-81861D12DA4A"'
			rc = BitBlt(Me.hdc, c * bdIntWidth, Me.ClientRectangle.Height - picMenu.ClientRectangle.Height - bdIntHeight * 3, bdIntWidth, bdIntHeight, picMisc.hdc, 38, 180, SRCCOPY)
			'UPGRADE_ISSUE: PictureBox property picMisc.hdc was not upgraded. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="CC4C7EC0-C903-48FC-ACCC-81861D12DA4A"'
			'UPGRADE_ISSUE: Form property frmMain.hdc was not upgraded. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="CC4C7EC0-C903-48FC-ACCC-81861D12DA4A"'
			rc = BitBlt(Me.hdc, c * bdIntWidth, Me.ClientRectangle.Height - picMenu.ClientRectangle.Height - bdIntHeight * 2, bdIntWidth, bdIntHeight, picMisc.hdc, 38, 180, SRCCOPY)
		Next c
		' Draw Buttons
		Y = Me.ClientRectangle.Height - picMenu.ClientRectangle.Height - bdIntHeight * 3 + 7
		' Display an Item Hint if needed
		If ButtonDown = -1 Then
			ShowText(Me, 40, Y, picInventory.ClientRectangle.Width - 80, 14, bdFontNoxiousWhite, (ItemNow.Hint), True, False)
		Else
			If picGrid.Visible = True Then
				ShowButton(Me, 40, Y, "Equip", (ButtonDown = 1))
				ShowButton(Me, 138, Y, "Talk", (ButtonDown = 2))
				ShowButton(Me, 236, Y, "Defend", (ButtonDown = 3))
				ShowButton(Me, 334, Y, "Wait", (ButtonDown = 4))
				ShowButton(Me, 432, Y, "Flee", (ButtonDown = 5))
			Else
				ShowButton(Me, 40, Y, "Equip", (ButtonDown = 1))
				ShowButton(Me, 138, Y, "Talk", (ButtonDown = 2))
				ShowButton(Me, 236, Y, "Search", (ButtonDown = 3))
				ShowButton(Me, 334, Y, "Listen", (ButtonDown = 4))
				ShowButton(Me, 432, Y, "Open", (ButtonDown = 5))
			End If
			' Draw skill buttons
			ShowText(Me, 528, Y, Me.ClientRectangle.Width - 136 - 528, 14, bdFontElixirWhite, (CreatureWithTurn.Name), 1, False)
			ShowButton(Me, Me.ClientRectangle.Width - 130, Y, "Skills", (ButtonDown = 6))
		End If
		'    If ButtonDown < -1 Then
		'        ' [Titi 2.4.8] get the runes names
		'        modBD.InitializeRunes (WorldNow.Name)
		'        ReDim sRune(intNbRunes)
		'        Text = Right$(strRunesList, Len(strRunesList) - 5) ' get rid of "List="
		'        For c = 1 To intNbRunes - 1
		'            sRune(c) = Left$(Text, InStr(Text, ",") - 1)
		'            Text = Right$(Text, Len(Text) - Len(sRune(c)) - 1)
		'        Next
		'        sRune(c) = Text
		''        Select Case Abs(ButtonDown + 2)
		''            Case 0: Text = "Blood"
		''            Case 1: Text = "Bile"
		''            Case 2: Text = "Oil"
		''            Case 3: Text = "Nectar"
		''            Case 4: Text = "Fire"
		''            Case 5: Text = "Earth"
		''            Case 6: Text = "Water"
		''            Case 7: Text = "Air"
		''            Case 8: Text = "Time"
		''            Case 9: Text = "Moon"
		''            Case 10: Text = "Sun"
		''            Case 11: Text = "Space"
		''            Case 12: Text = "Insect"
		''            Case 13: Text = "Man"
		''            Case 14: Text = "Animal"
		''            Case 15: Text = "Fish"
		''            Case 16: Text = "Twilight"
		''            Case 17: Text = "Abyss"
		''            Case 18: Text = "Eternium"
		''            Case 19: Text = "Dreams"
		''        End Select
		''        ShowText frmMain, 40, y, picInventory.ScaleWidth - 80, 14, bdFontNoxiousWhite, Map.Runes(Abs(ButtonDown + 2)) & " " & Text & " Rune for Spells", True, False
		'        x = 16
		'        ' first half of the runes list
		'        y = (Abs(ButtonDown) - 2) * 32 + 16
		'        ' second part of the list --> right of the screen
		'        If ButtonDown < -11 Then y = (Abs(ButtonDown) - 12) * 32 + 16
		'        ShowText picMap, x, y, picMap.ScaleWidth - 80, 14, bdFontNoxiousWhite, sRune(Abs(ButtonDown + 1)), IIf(ButtonDown < -11, 1, 0), False
		'        picMap.Refresh
		'        PauseTime = 1   ' duration of display
		'        Start = Timer
		'        Do While Timer < Start + PauseTime
		'           DoEvents
		'        Loop
		'        ' ShowText picMap, x, y, picMap.ScaleWidth - 80, 14, bdFontNoxiousWhite, String(Len(sRune(Abs(ButtonDown + 1))), "*"), IIf(ButtonDown < -11, 1, 0), False
		'        DrawMapRegion 0, 0, 340 + HintX, Len(sRune(Abs(ButtonDown + 1))) * 8 + 16 + HintY
		'        DrawMapRegion 0, picMap.ScaleWidth - Len(sRune(Abs(ButtonDown + 1))) * 8 - 16 - HintY, 340 + HintX, picMap.ScaleWidth
		'    End If
		' Draw the Menu (Button 7) and Journal (Button 8) Orbs
		If ButtonDown = 7 Then
			'UPGRADE_ISSUE: PictureBox property picMisc.hdc was not upgraded. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="CC4C7EC0-C903-48FC-ACCC-81861D12DA4A"'
			'UPGRADE_ISSUE: Form property frmMain.hdc was not upgraded. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="CC4C7EC0-C903-48FC-ACCC-81861D12DA4A"'
			rc = BitBlt(Me.hdc, 0, Me.ClientRectangle.Height - picMenu.ClientRectangle.Height - bdIntHeight * 3, 32, 32, picMisc.hdc, 32, 147, SRCCOPY)
			'UPGRADE_ISSUE: PictureBox property picMisc.hdc was not upgraded. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="CC4C7EC0-C903-48FC-ACCC-81861D12DA4A"'
			'UPGRADE_ISSUE: Form property frmMain.hdc was not upgraded. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="CC4C7EC0-C903-48FC-ACCC-81861D12DA4A"'
			rc = BitBlt(Me.hdc, Me.ClientRectangle.Width - 32, Me.ClientRectangle.Height - picMenu.ClientRectangle.Height - bdIntHeight * 3, 32, 32, picMisc.hdc, 64, 147, SRCCOPY)
		ElseIf ButtonDown = 8 Then 
			'UPGRADE_ISSUE: PictureBox property picMisc.hdc was not upgraded. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="CC4C7EC0-C903-48FC-ACCC-81861D12DA4A"'
			'UPGRADE_ISSUE: Form property frmMain.hdc was not upgraded. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="CC4C7EC0-C903-48FC-ACCC-81861D12DA4A"'
			rc = BitBlt(Me.hdc, 0, Me.ClientRectangle.Height - picMenu.ClientRectangle.Height - bdIntHeight * 3, 32, 32, picMisc.hdc, 0, 147, SRCCOPY)
			'UPGRADE_ISSUE: PictureBox property picMisc.hdc was not upgraded. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="CC4C7EC0-C903-48FC-ACCC-81861D12DA4A"'
			'UPGRADE_ISSUE: Form property frmMain.hdc was not upgraded. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="CC4C7EC0-C903-48FC-ACCC-81861D12DA4A"'
			rc = BitBlt(Me.hdc, Me.ClientRectangle.Width - 32, Me.ClientRectangle.Height - picMenu.ClientRectangle.Height - bdIntHeight * 3, 32, 32, picMisc.hdc, 96, 147, SRCCOPY)
		Else
			'UPGRADE_ISSUE: PictureBox property picMisc.hdc was not upgraded. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="CC4C7EC0-C903-48FC-ACCC-81861D12DA4A"'
			'UPGRADE_ISSUE: Form property frmMain.hdc was not upgraded. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="CC4C7EC0-C903-48FC-ACCC-81861D12DA4A"'
			rc = BitBlt(Me.hdc, 0, Me.ClientRectangle.Height - picMenu.ClientRectangle.Height - bdIntHeight * 3, 32, 32, picMisc.hdc, 0, 147, SRCCOPY)
			'UPGRADE_ISSUE: PictureBox property picMisc.hdc was not upgraded. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="CC4C7EC0-C903-48FC-ACCC-81861D12DA4A"'
			'UPGRADE_ISSUE: Form property frmMain.hdc was not upgraded. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="CC4C7EC0-C903-48FC-ACCC-81861D12DA4A"'
			rc = BitBlt(Me.hdc, Me.ClientRectangle.Width - 32, Me.ClientRectangle.Height - picMenu.ClientRectangle.Height - bdIntHeight * 3, 32, 32, picMisc.hdc, 64, 147, SRCCOPY)
		End If
		' Fill in around menu sides
		If picGrid.Visible = False Then
			'UPGRADE_ISSUE: PictureBox property picConvoBottom.hdc was not upgraded. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="CC4C7EC0-C903-48FC-ACCC-81861D12DA4A"'
			'UPGRADE_ISSUE: Form property frmMain.hdc was not upgraded. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="CC4C7EC0-C903-48FC-ACCC-81861D12DA4A"'
			rc = BitBlt(Me.hdc, 32, Me.ClientRectangle.Height - picMenu.ClientRectangle.Height - 16, (Me.ClientRectangle.Width - picMenu.ClientRectangle.Width - 64) / 2, picMenu.ClientRectangle.Height, picConvoBottom.hdc, 0, 6, SRCCOPY)
			'UPGRADE_ISSUE: PictureBox property picConvoBottom.hdc was not upgraded. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="CC4C7EC0-C903-48FC-ACCC-81861D12DA4A"'
			'UPGRADE_ISSUE: Form property frmMain.hdc was not upgraded. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="CC4C7EC0-C903-48FC-ACCC-81861D12DA4A"'
			rc = BitBlt(Me.hdc, 32 + (Me.ClientRectangle.Width - picMenu.ClientRectangle.Width - 64) / 2 + picMenu.ClientRectangle.Width, Me.ClientRectangle.Height - picMenu.ClientRectangle.Height - 16, (Me.ClientRectangle.Width - picMenu.ClientRectangle.Width - 64) / 2, picMenu.ClientRectangle.Height, picConvoBottom.hdc, 0, 6, SRCCOPY)
		Else
			'UPGRADE_ISSUE: PictureBox property picConvoBottom.hdc was not upgraded. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="CC4C7EC0-C903-48FC-ACCC-81861D12DA4A"'
			'UPGRADE_ISSUE: Form property frmMain.hdc was not upgraded. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="CC4C7EC0-C903-48FC-ACCC-81861D12DA4A"'
			rc = BitBlt(Me.hdc, 32, Me.ClientRectangle.Height - picToHit.ClientRectangle.Height - 16, (Me.ClientRectangle.Width - picToHit.ClientRectangle.Width - 64) / 2, picToHit.ClientRectangle.Height, picConvoBottom.hdc, 0, 6, SRCCOPY)
			'UPGRADE_ISSUE: PictureBox property picConvoBottom.hdc was not upgraded. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="CC4C7EC0-C903-48FC-ACCC-81861D12DA4A"'
			'UPGRADE_ISSUE: Form property frmMain.hdc was not upgraded. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="CC4C7EC0-C903-48FC-ACCC-81861D12DA4A"'
			rc = BitBlt(Me.hdc, 32 + (Me.ClientRectangle.Width - picToHit.ClientRectangle.Width - 64) / 2 + picToHit.ClientRectangle.Width, Me.ClientRectangle.Height - picToHit.ClientRectangle.Height - 16, (Me.ClientRectangle.Width - picToHit.ClientRectangle.Width - 64) / 2, picToHit.ClientRectangle.Height, picConvoBottom.hdc, 0, 6, SRCCOPY)
		End If
	End Sub
	
	Private Sub BorderDrawAll(ByRef InGame As Short)
		Dim X, i, Y As Short
		Dim rc As Integer
		'UPGRADE_NOTE: Height was upgraded to Height_Renamed. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="A9E4979A-37FA-4718-9994-97DD76ED70A7"'
		'UPGRADE_NOTE: Width was upgraded to Width_Renamed. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="A9E4979A-37FA-4718-9994-97DD76ED70A7"'
		Dim Width_Renamed, Height_Renamed As Short
		BorderDrawTop()
		BorderDrawBottom()
		BorderDrawSides(InGame, -1)
		Width_Renamed = 90 : Height_Renamed = 16
		If InGame = True Then
			BorderDrawButtons(0)
			picMenu.Visible = True
		End If
		Me.Refresh()
	End Sub
	
	'UPGRADE_NOTE: Refresh was upgraded to Refresh_Renamed. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="A9E4979A-37FA-4718-9994-97DD76ED70A7"'
	Private Sub GameInterfaceLoad(ByRef Refresh_Renamed As Short)
		Dim c As Short
		Dim rc As Integer
		Dim CreatureX As Creature
		Dim sPath As String
		'UPGRADE_NOTE: Text was upgraded to Text_Renamed. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="A9E4979A-37FA-4718-9994-97DD76ED70A7"'
		Dim hndFile As Short
		Dim Text_Renamed, sWorld As String
		Dim lResult As Integer
		On Error GoTo ErrorHandler
		'    sPath = gAppPath & "\data\Interface\" & GlobalInterfaceName & "\"
		sPath = gDataPath & "\Interface\" & GlobalInterfaceName & "\"
		' Misc has borders and face stuff
		picSteps.Image = System.Drawing.Image.FromFile(sPath & "Steps.bmp")
		picMisc.Image = System.Drawing.Image.FromFile(sPath & "Misc.bmp")
		picFonts.Image = System.Drawing.Image.FromFile(sPath & "Fonts.bmp")
		' [Titi 2.4.8] since we can now change the name, why not change the picture of the runes?
		lResult = fReadValue(gAppPath & "\Settings.ini", "World", "Current", "S", "Eternia", sWorld)
		InitializeRunes(sWorld)
		' Combat stuff
		picToHit.Image = System.Drawing.Image.FromFile(sPath & "CombatRoll.bmp")
		' Load dice
		LoadDicePic(GlobalDiceName)
		' Clear out any existing Creature Faces
		For c = 1 To bdMaxMonsPics
			PicFile(c) = ""
			If PicFileTime(c) > 0 Then
				picCPic.Unload(c)
				picCMap.Unload(c)
			End If
			PicFileTime(c) = 0
		Next c
		For	Each CreatureX In EncounterNow.Creatures
			CreatureX.Pic = 0
		Next CreatureX
		For	Each CreatureX In Tome.Creatures
			CreatureX.Pic = 0
		Next CreatureX
		picFaces.Image = System.Drawing.Image.FromFile(sPath & "DialogFace.bmp")
		picFaces.SizeMode = System.Windows.Forms.PictureBoxSizeMode.AutoSize : picFaces.SizeMode = System.Windows.Forms.PictureBoxSizeMode.Normal
		picMainMenu.Image = System.Drawing.Image.FromFile(sPath & "MainMenu.bmp")
		picBlack.Image = System.Drawing.Image.FromFile(sPath & "MainMenuRollOver.bmp")
		picTomeNew.BackgroundImageLayout = System.Windows.Forms.ImageLayout.None
		'UPGRADE_WARNING: Picture has a new behavior. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="9B7D5ADD-D8FE-4819-A36C-6DEDAF088CC7"'
		picTomeNew.BackgroundImage = System.Drawing.Image.FromFile(sPath & "TomeMenu.bmp")
		picCreateName.Image = System.Drawing.Image.FromFile(sPath & "DialogAcceptText.bmp")
		picInventory.Image = System.Drawing.Image.FromFile(sPath & "DialogInventory.bmp")
		' DialogShow
		picConvo.BackgroundImageLayout = System.Windows.Forms.ImageLayout.None
		'UPGRADE_WARNING: Picture has a new behavior. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="9B7D5ADD-D8FE-4819-A36C-6DEDAF088CC7"'
		picConvo.BackgroundImage = System.Drawing.Image.FromFile(sPath & "DialogShow.bmp")
		picConvoBottom.Width = picConvo.Width : picConvoBottom.Height = 6 + picMenu.ClientRectangle.Height
		'UPGRADE_ISSUE: PictureBox property picConvo.hdc was not upgraded. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="CC4C7EC0-C903-48FC-ACCC-81861D12DA4A"'
		'UPGRADE_ISSUE: PictureBox property picConvoBottom.hdc was not upgraded. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="CC4C7EC0-C903-48FC-ACCC-81861D12DA4A"'
		rc = BitBlt(picConvoBottom.hdc, 0, 0, picConvo.ClientRectangle.Width, picConvo.ClientRectangle.Height, picConvo.hdc, 0, picConvo.Height - 6, SRCCOPY)
		'UPGRADE_ISSUE: PictureBox property picConvo.hdc was not upgraded. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="CC4C7EC0-C903-48FC-ACCC-81861D12DA4A"'
		'UPGRADE_ISSUE: PictureBox property picConvoBottom.hdc was not upgraded. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="CC4C7EC0-C903-48FC-ACCC-81861D12DA4A"'
		rc = BitBlt(picConvoBottom.hdc, 0, 6, picConvo.ClientRectangle.Width, picConvo.ClientRectangle.Height, picConvo.hdc, 6, 64, SRCCOPY)
		picConvoBottom.Refresh()
		picConvoEnter.Image = System.Drawing.Image.FromFile(sPath & "DialogAcceptText.bmp")
		picConvoList.Image = System.Drawing.Image.FromFile(sPath & "DialogListBox.bmp")
		picTalk.Image = System.Drawing.Image.FromFile(sPath & "DialogTalk.bmp")
		picJournal.BackgroundImageLayout = System.Windows.Forms.ImageLayout.None
		'UPGRADE_WARNING: Picture has a new behavior. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="9B7D5ADD-D8FE-4819-A36C-6DEDAF088CC7"'
		picJournal.BackgroundImage = System.Drawing.Image.FromFile(sPath & "DialogJournal.bmp")
		picSearch.Image = System.Drawing.Image.FromFile(sPath & "DialogSearch.bmp")
		picContainer.Image = System.Drawing.Image.FromFile(sPath & "DialogSearch.bmp")
		picBuySell.Image = System.Drawing.Image.FromFile(sPath & "DialogBuySell.bmp")
		picDialogBrief.BackgroundImageLayout = System.Windows.Forms.ImageLayout.None
		'UPGRADE_WARNING: Picture has a new behavior. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="9B7D5ADD-D8FE-4819-A36C-6DEDAF088CC7"'
		picDialogBrief.BackgroundImage = System.Drawing.Image.FromFile(sPath & "DialogBrief.bmp")
		Exit Sub
ErrorHandler: 
		oErr.logError("GameInterfaceLoad()")
		End '  this is an unclean end to the program [borfaux]
		Resume Next
	End Sub
	
	'UPGRADE_NOTE: Width was upgraded to Width_Renamed. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="A9E4979A-37FA-4718-9994-97DD76ED70A7"'
	Private Sub GameInitInventory(ByRef StartAt As Short, ByRef EndAt As Short, ByRef Width_Renamed As Short)
		Dim c As Short
		For c = 1 To EndAt - StartAt
			InvX(StartAt + c) = InvX(StartAt) + 34 * (c Mod Width_Renamed)
			InvY(StartAt + c) = InvY(StartAt) + 34 * Int(c / Width_Renamed)
		Next c
	End Sub
	
	Private Sub GameInitVariables()
		' Initialize Inventory Tile Positions
		InvX(1) = 22 : InvY(1) = 44 ' LeftBack
		GameInitInventory(1, 6, 2)
		InvX(7) = 190 : InvY(7) = 44 ' RightBack
		GameInitInventory(7, 12, 2)
		InvX(13) = 106 : InvY(13) = 44 ' Head
		GameInitInventory(13, 16, 2)
		InvX(17) = 22 : InvY(17) = 160 ' LeftHandOther
		GameInitInventory(17, 20, 2)
		InvX(21) = 190 : InvY(21) = 160 ' RightHandOther
		GameInitInventory(21, 24, 2)
		InvX(25) = 106 : InvY(25) = 120 ' Necklace
		GameInitInventory(25, 26, 2)
		InvX(27) = 106 : InvY(27) = 154 ' Body
		GameInitInventory(27, 30, 2)
		InvX(31) = 106 : InvY(31) = 222 ' Belt
		GameInitInventory(31, 32, 2)
		InvX(33) = 22 : InvY(33) = 244 ' LeftHand
		GameInitInventory(33, 38, 2)
		InvX(39) = 190 : InvY(39) = 244 ' RightHand
		GameInitInventory(39, 44, 2)
		InvX(45) = 106 : InvY(45) = 328 ' Boots
		GameInitInventory(45, 48, 2)
		InvX(49) = 290 : InvY(49) = 44 ' All Other Items
		GameInitInventory(49, 128, 8)
		' Initialize directions for find path routine
		DirX(0) = 0 : DirY(0) = -1 ' up / Upper Left
		DirX(1) = 1 : DirY(1) = 0 ' right / Upper Right
		DirX(2) = -1 : DirY(2) = 0 ' left / Lower Left
		DirX(3) = 0 : DirY(3) = 1 ' down / Lower Right
		DirX(4) = -1 : DirY(4) = -1 ' up left
		DirX(5) = -1 : DirY(5) = 1 ' down left
		DirX(6) = 1 : DirY(6) = -1 ' up right
		DirX(7) = 1 : DirY(7) = 1 ' down right
		' Initialize directors for LineOfSite (LOS)
		' 0 - up, 1 - right, 2 - left, 3 - down
		LosDir(0, 0) = 0 : LosDir(0, 1) = 0 : LosDir(0, 2) = 0 : LosDir(0, 3) = 0 : LosDir(0, 4) = -1 : LosDir(0, 5) = -1
		LosDir(1, 0) = 0 : LosDir(1, 1) = 0 : LosDir(1, 2) = 1 : LosDir(1, 3) = 0 : LosDir(1, 4) = 0 : LosDir(1, 5) = -1
		LosDir(2, 0) = 0 : LosDir(2, 1) = 1 : LosDir(2, 2) = 0 : LosDir(2, 3) = 1 : LosDir(2, 4) = 0 : LosDir(2, 5) = 1
		LosDir(3, 0) = 1 : LosDir(3, 1) = 0 : LosDir(3, 2) = 1 : LosDir(3, 3) = 0 : LosDir(3, 4) = 1 : LosDir(3, 5) = 0
		LosDir(4, 0) = 1 : LosDir(4, 1) = 1 : LosDir(4, 2) = 0 : LosDir(4, 3) = 1 : LosDir(4, 4) = 1 : LosDir(4, 5) = -1
		LosDir(5, 0) = 1 : LosDir(5, 1) = 1 : LosDir(5, 2) = 1 : LosDir(5, 3) = 1 : LosDir(5, 4) = -1 : LosDir(5, 5) = -1
		' Initialize RuneMenuX and RuneMenuY
		RuneMenuX(0) = 12 : RuneMenuY(0) = 66
		RuneMenuX(1) = 12 : RuneMenuY(1) = 98
		RuneMenuX(2) = 44 : RuneMenuY(2) = 66
		RuneMenuX(3) = 44 : RuneMenuY(3) = 98
		RuneMenuX(4) = 82 : RuneMenuY(4) = 50
		RuneMenuX(6) = 114 : RuneMenuY(6) = 50
		RuneMenuX(5) = 82 : RuneMenuY(5) = 88
		RuneMenuX(7) = 114 : RuneMenuY(7) = 88
		RuneMenuX(8) = 152 : RuneMenuY(8) = 34
		RuneMenuX(9) = 152 : RuneMenuY(9) = 66
		RuneMenuX(10) = 184 : RuneMenuY(10) = 34
		RuneMenuX(11) = 184 : RuneMenuY(11) = 66
		RuneMenuX(12) = 222 : RuneMenuY(12) = 50
		RuneMenuX(13) = 222 : RuneMenuY(13) = 88
		RuneMenuX(14) = 254 : RuneMenuY(14) = 50
		RuneMenuX(15) = 254 : RuneMenuY(15) = 88
		RuneMenuX(16) = 292 : RuneMenuY(16) = 66
		RuneMenuX(17) = 292 : RuneMenuY(17) = 98
		RuneMenuX(18) = 324 : RuneMenuY(18) = 66
		RuneMenuX(19) = 324 : RuneMenuY(19) = 98
		' Init Alpha (small font)
		aW(0, 32) = 8 : aX(0, 32) = 0 : aY(0, 32) = 0
		aW(0, 33) = 7 : aX(0, 33) = 8 : aY(0, 33) = 0
		aW(0, 34) = 7 : aX(0, 34) = 15 : aY(0, 34) = 0
		aW(0, 35) = 8 : aX(0, 35) = 22 : aY(0, 35) = 0
		aW(0, 36) = 8 : aX(0, 36) = 30 : aY(0, 36) = 0
		aW(0, 37) = 8 : aX(0, 37) = 38 : aY(0, 37) = 0
		aW(0, 38) = 8 : aX(0, 38) = 46 : aY(0, 38) = 0
		aW(0, 39) = 9 : aX(0, 39) = 54 : aY(0, 39) = 0
		aW(0, 40) = 8 : aX(0, 40) = 63 : aY(0, 40) = 0
		aW(0, 41) = 8 : aX(0, 41) = 71 : aY(0, 41) = 0
		aW(0, 42) = 7 : aX(0, 42) = 79 : aY(0, 42) = 0
		aW(0, 43) = 7 : aX(0, 43) = 86 : aY(0, 43) = 0
		aW(0, 44) = 10 : aX(0, 44) = 93 : aY(0, 44) = 0
		aW(0, 45) = 10 : aX(0, 45) = 103 : aY(0, 45) = 0
		aW(0, 46) = 9 : aX(0, 46) = 113 : aY(0, 46) = 0
		aW(0, 47) = 6 : aX(0, 47) = 122 : aY(0, 47) = 0
		aW(0, 48) = 10 : aX(0, 48) = 128 : aY(0, 48) = 0
		aW(0, 49) = 8 : aX(0, 49) = 138 : aY(0, 49) = 0
		aW(0, 50) = 9 : aX(0, 50) = 146 : aY(0, 50) = 0
		aW(0, 51) = 8 : aX(0, 51) = 155 : aY(0, 51) = 0
		aW(0, 52) = 9 : aX(0, 52) = 163 : aY(0, 52) = 0
		aW(0, 53) = 9 : aX(0, 53) = 172 : aY(0, 53) = 0
		aW(0, 54) = 12 : aX(0, 54) = 181 : aY(0, 54) = 0
		aW(0, 55) = 8 : aX(0, 55) = 193 : aY(0, 55) = 0
		aW(0, 56) = 8 : aX(0, 56) = 201 : aY(0, 56) = 0
		aW(0, 57) = 9 : aX(0, 57) = 209 : aY(0, 57) = 0 ' Z
		aW(0, 16) = 6 : aX(0, 16) = 0 : aY(0, 16) = 10 ' A
		aW(0, 17) = 8 : aX(0, 17) = 6 : aY(0, 17) = 10
		aW(0, 18) = 7 : aX(0, 18) = 14 : aY(0, 18) = 10
		aW(0, 19) = 8 : aX(0, 19) = 21 : aY(0, 19) = 10
		aW(0, 20) = 8 : aX(0, 20) = 29 : aY(0, 20) = 10
		aW(0, 21) = 8 : aX(0, 21) = 37 : aY(0, 21) = 10
		aW(0, 22) = 8 : aX(0, 22) = 45 : aY(0, 22) = 10
		aW(0, 23) = 8 : aX(0, 23) = 53 : aY(0, 23) = 10
		aW(0, 24) = 8 : aX(0, 24) = 61 : aY(0, 24) = 10
		aW(0, 15) = 8 : aX(0, 15) = 69 : aY(0, 15) = 10
		aW(0, 10) = 6 : aX(0, 10) = 77 : aY(0, 10) = 10
		aW(0, 12) = 5 : aX(0, 12) = 83 : aY(0, 12) = 10
		aW(0, 0) = 4 : aX(0, 0) = 88 : aY(0, 0) = 10
		aW(0, 13) = 3 : aX(0, 13) = 92 : aY(0, 13) = 10
		aW(0, 11) = 4 : aX(0, 11) = 95 : aY(0, 11) = 10
		aW(0, 30) = 7 : aX(0, 30) = 99 : aY(0, 30) = 10
		aW(0, 1) = 6 : aX(0, 1) = 106 : aY(0, 1) = 10
		aW(0, 6) = 4 : aX(0, 6) = 112 : aY(0, 6) = 10
		aW(0, 7) = 6 : aX(0, 7) = 116 : aY(0, 7) = 10
		aW(0, 8) = 5 : aX(0, 8) = 122 : aY(0, 8) = 10
		aW(0, 5) = 8 : aX(0, 5) = 127 : aY(0, 5) = 10
		aW(0, 9) = 11 : aX(0, 9) = 135 : aY(0, 9) = 10
		aW(0, 27) = 0 : aX(0, 27) = 146 : aY(0, 27) = 10
		aW(0, 29) = 0 : aX(0, 29) = 146 : aY(0, 29) = 10
		aW(0, 4) = 0 : aX(0, 4) = 146 : aY(0, 4) = 10
		aW(0, 3) = 0 : aX(0, 3) = 146 : aY(0, 3) = 10
		aW(0, 25) = 0 : aX(0, 25) = 146 : aY(0, 25) = 10
		aW(0, 26) = 0 : aX(0, 26) = 146 : aY(0, 26) = 10 ' Z
		aW(0, 64) = 8 : aX(0, 64) = 0 : aY(0, 64) = 0
		aW(0, 65) = 7 : aX(0, 65) = 8 : aY(0, 65) = 0
		aW(0, 66) = 7 : aX(0, 66) = 15 : aY(0, 66) = 0
		aW(0, 67) = 8 : aX(0, 67) = 22 : aY(0, 67) = 0
		aW(0, 68) = 8 : aX(0, 68) = 30 : aY(0, 68) = 0
		aW(0, 69) = 8 : aX(0, 69) = 38 : aY(0, 69) = 0
		aW(0, 70) = 8 : aX(0, 70) = 46 : aY(0, 70) = 0
		aW(0, 71) = 9 : aX(0, 71) = 54 : aY(0, 71) = 0
		aW(0, 72) = 8 : aX(0, 72) = 63 : aY(0, 72) = 0
		aW(0, 73) = 8 : aX(0, 73) = 71 : aY(0, 73) = 0
		aW(0, 74) = 7 : aX(0, 74) = 79 : aY(0, 74) = 0
		aW(0, 75) = 7 : aX(0, 75) = 86 : aY(0, 75) = 0
		aW(0, 76) = 10 : aX(0, 76) = 93 : aY(0, 76) = 0
		aW(0, 77) = 10 : aX(0, 77) = 103 : aY(0, 77) = 0
		aW(0, 78) = 9 : aX(0, 78) = 113 : aY(0, 78) = 0
		aW(0, 79) = 6 : aX(0, 79) = 122 : aY(0, 79) = 0
		aW(0, 80) = 10 : aX(0, 80) = 128 : aY(0, 80) = 0
		aW(0, 81) = 8 : aX(0, 81) = 138 : aY(0, 81) = 0
		aW(0, 82) = 9 : aX(0, 82) = 146 : aY(0, 82) = 0
		aW(0, 83) = 8 : aX(0, 83) = 155 : aY(0, 83) = 0
		aW(0, 84) = 9 : aX(0, 84) = 163 : aY(0, 84) = 0
		aW(0, 85) = 9 : aX(0, 85) = 172 : aY(0, 85) = 0
		aW(0, 86) = 12 : aX(0, 86) = 181 : aY(0, 86) = 0
		aW(0, 87) = 8 : aX(0, 87) = 193 : aY(0, 87) = 0
		aW(0, 88) = 8 : aX(0, 88) = 201 : aY(0, 88) = 0
		aW(0, 89) = 9 : aX(0, 89) = 209 : aY(0, 89) = 0
		aW(0, 14) = 6 : aX(0, 14) = 209 : aY(0, 14) = 10 ' [Titi 2.4.9]
		' Noxious
		aW(1, 32) = 11 : aX(1, 32) = 0 : aY(1, 32) = 0 ' A
		aW(1, 33) = 9 : aX(1, 33) = 11 : aY(1, 33) = 0 ' B
		aW(1, 34) = 10 : aX(1, 34) = 20 : aY(1, 34) = 0 ' C
		aW(1, 35) = 11 : aX(1, 35) = 30 : aY(1, 35) = 0 ' D
		aW(1, 36) = 8 : aX(1, 36) = 41 : aY(1, 36) = 0 ' E
		aW(1, 37) = 8 : aX(1, 37) = 49 : aY(1, 37) = 0 ' F
		aW(1, 38) = 10 : aX(1, 38) = 57 : aY(1, 38) = 0 ' G
		aW(1, 39) = 11 : aX(1, 39) = 67 : aY(1, 39) = 0 ' H
		aW(1, 40) = 5 : aX(1, 40) = 78 : aY(1, 40) = 0 ' I
		aW(1, 41) = 8 : aX(1, 41) = 83 : aY(1, 41) = 0 ' J
		aW(1, 42) = 10 : aX(1, 42) = 91 : aY(1, 42) = 0 ' K
		aW(1, 43) = 8 : aX(1, 43) = 101 : aY(1, 43) = 0 ' L
		aW(1, 44) = 12 : aX(1, 44) = 109 : aY(1, 44) = 0 ' M
		aW(1, 45) = 10 : aX(1, 45) = 121 : aY(1, 45) = 0 ' N
		aW(1, 46) = 11 : aX(1, 46) = 131 : aY(1, 46) = 0 ' O
		aW(1, 47) = 8 : aX(1, 47) = 142 : aY(1, 47) = 0 ' P
		aW(1, 48) = 13 : aX(1, 48) = 150 : aY(1, 48) = 0 ' Q
		aW(1, 49) = 9 : aX(1, 49) = 163 : aY(1, 49) = 0 ' R
		aW(1, 50) = 8 : aX(1, 50) = 172 : aY(1, 50) = 0 ' S
		aW(1, 51) = 9 : aX(1, 51) = 181 : aY(1, 51) = 0 ' T
		aW(1, 52) = 10 : aX(1, 52) = 190 : aY(1, 52) = 0 ' U
		aW(1, 53) = 10 : aX(1, 53) = 200 : aY(1, 53) = 0 ' V
		aW(1, 54) = 13 : aX(1, 54) = 210 : aY(1, 54) = 0 ' W
		aW(1, 55) = 11 : aX(1, 55) = 223 : aY(1, 55) = 0 ' X
		aW(1, 56) = 9 : aX(1, 56) = 234 : aY(1, 56) = 0 ' Y
		aW(1, 57) = 8 : aX(1, 57) = 243 : aY(1, 57) = 0 ' Z
		aW(1, 16) = 6 : aX(1, 16) = 0 : aY(1, 16) = 14
		aW(1, 17) = 10 : aX(1, 17) = 6 : aY(1, 17) = 14
		aW(1, 18) = 10 : aX(1, 18) = 16 : aY(1, 18) = 14
		aW(1, 19) = 9 : aX(1, 19) = 26 : aY(1, 19) = 14
		aW(1, 20) = 10 : aX(1, 20) = 35 : aY(1, 20) = 14
		aW(1, 21) = 9 : aX(1, 21) = 45 : aY(1, 21) = 14
		aW(1, 22) = 10 : aX(1, 22) = 54 : aY(1, 22) = 14
		aW(1, 23) = 9 : aX(1, 23) = 64 : aY(1, 23) = 14
		aW(1, 24) = 9 : aX(1, 24) = 73 : aY(1, 24) = 14
		aW(1, 15) = 10 : aX(1, 15) = 82 : aY(1, 15) = 14
		aW(1, 10) = 7 : aX(1, 10) = 92 : aY(1, 10) = 14
		aW(1, 12) = 5 : aX(1, 12) = 99 : aY(1, 12) = 14
		aW(1, 0) = 4 : aX(1, 0) = 104 : aY(1, 0) = 14
		aW(1, 13) = 4 : aX(1, 13) = 108 : aY(1, 13) = 14
		aW(1, 11) = 4 : aX(1, 11) = 112 : aY(1, 11) = 14
		aW(1, 30) = 7 : aX(1, 30) = 116 : aY(1, 30) = 14
		aW(1, 1) = 4 : aX(1, 1) = 123 : aY(1, 1) = 14
		aW(1, 6) = 5 : aX(1, 6) = 127 : aY(1, 6) = 14
		aW(1, 7) = 4 : aX(1, 7) = 132 : aY(1, 7) = 14
		aW(1, 8) = 4 : aX(1, 8) = 136 : aY(1, 8) = 14
		aW(1, 5) = 14 : aX(1, 5) = 140 : aY(1, 5) = 14
		aW(1, 9) = 11 : aX(1, 9) = 154 : aY(1, 9) = 14
		aW(1, 27) = 9 : aX(1, 27) = 165 : aY(1, 27) = 14
		aW(1, 29) = 9 : aX(1, 29) = 174 : aY(1, 29) = 14
		aW(1, 4) = 13 : aX(1, 4) = 183 : aY(1, 4) = 14
		aW(1, 3) = 9 : aX(1, 3) = 196 : aY(1, 3) = 14
		aW(1, 25) = 4 : aX(1, 25) = 205 : aY(1, 25) = 14
		aW(1, 26) = 6 : aX(1, 26) = 209 : aY(1, 26) = 14 ' Z
		aW(1, 64) = 7 : aX(1, 64) = 0 : aY(1, 64) = 28
		aW(1, 65) = 10 : aX(1, 65) = 7 : aY(1, 65) = 28
		aW(1, 66) = 7 : aX(1, 66) = 17 : aY(1, 66) = 28
		aW(1, 67) = 10 : aX(1, 67) = 24 : aY(1, 67) = 28
		aW(1, 68) = 8 : aX(1, 68) = 34 : aY(1, 68) = 28
		aW(1, 69) = 6 : aX(1, 69) = 42 : aY(1, 69) = 28
		aW(1, 70) = 9 : aX(1, 70) = 48 : aY(1, 70) = 28
		aW(1, 71) = 9 : aX(1, 71) = 57 : aY(1, 71) = 28
		aW(1, 72) = 4 : aX(1, 72) = 66 : aY(1, 72) = 28
		aW(1, 73) = 6 : aX(1, 73) = 70 : aY(1, 73) = 28
		aW(1, 74) = 9 : aX(1, 74) = 76 : aY(1, 74) = 28
		aW(1, 75) = 4 : aX(1, 75) = 85 : aY(1, 75) = 28
		aW(1, 76) = 14 : aX(1, 76) = 89 : aY(1, 76) = 28
		aW(1, 77) = 9 : aX(1, 77) = 103 : aY(1, 77) = 28
		aW(1, 78) = 9 : aX(1, 78) = 112 : aY(1, 78) = 28
		aW(1, 79) = 10 : aX(1, 79) = 121 : aY(1, 79) = 28
		aW(1, 80) = 9 : aX(1, 80) = 131 : aY(1, 80) = 28
		aW(1, 81) = 5 : aX(1, 81) = 140 : aY(1, 81) = 28
		aW(1, 82) = 7 : aX(1, 82) = 145 : aY(1, 82) = 28
		aW(1, 83) = 6 : aX(1, 83) = 152 : aY(1, 83) = 28
		aW(1, 84) = 9 : aX(1, 84) = 158 : aY(1, 84) = 28
		aW(1, 85) = 8 : aX(1, 85) = 167 : aY(1, 85) = 28
		aW(1, 86) = 11 : aX(1, 86) = 175 : aY(1, 86) = 28
		aW(1, 87) = 9 : aX(1, 87) = 186 : aY(1, 87) = 28
		aW(1, 88) = 7 : aX(1, 88) = 195 : aY(1, 88) = 28
		aW(1, 89) = 7 : aX(1, 89) = 202 : aY(1, 89) = 28 ' Z
		aW(1, 14) = 9 : aX(1, 14) = 216 : aY(1, 14) = 14 ' [Titi 2.4.9]
		' Elixir
		aW(2, 32) = 10 : aX(2, 32) = 0 : aY(2, 32) = 0
		aW(2, 33) = 11 : aX(2, 33) = 10 : aY(2, 33) = 0
		aW(2, 34) = 10 : aX(2, 34) = 21 : aY(2, 34) = 0
		aW(2, 35) = 11 : aX(2, 35) = 31 : aY(2, 35) = 0
		aW(2, 36) = 10 : aX(2, 36) = 42 : aY(2, 36) = 0
		aW(2, 37) = 11 : aX(2, 37) = 52 : aY(2, 37) = 0
		aW(2, 38) = 11 : aX(2, 38) = 63 : aY(2, 38) = 0
		aW(2, 39) = 11 : aX(2, 39) = 74 : aY(2, 39) = 0
		aW(2, 40) = 6 : aX(2, 40) = 85 : aY(2, 40) = 0
		aW(2, 41) = 6 : aX(2, 41) = 91 : aY(2, 41) = 0
		aW(2, 42) = 12 : aX(2, 42) = 97 : aY(2, 42) = 0
		aW(2, 43) = 10 : aX(2, 43) = 109 : aY(2, 43) = 0
		aW(2, 44) = 17 : aX(2, 44) = 119 : aY(2, 44) = 0
		aW(2, 45) = 14 : aX(2, 45) = 136 : aY(2, 45) = 0
		aW(2, 46) = 13 : aX(2, 46) = 150 : aY(2, 46) = 0
		aW(2, 47) = 13 : aX(2, 47) = 163 : aY(2, 47) = 0
		aW(2, 48) = 11 : aX(2, 48) = 176 : aY(2, 48) = 0
		aW(2, 49) = 12 : aX(2, 49) = 187 : aY(2, 49) = 0
		aW(2, 50) = 10 : aX(2, 50) = 199 : aY(2, 50) = 0
		aW(2, 51) = 9 : aX(2, 51) = 209 : aY(2, 51) = 0
		aW(2, 52) = 13 : aX(2, 52) = 218 : aY(2, 52) = 0
		aW(2, 53) = 11 : aX(2, 53) = 231 : aY(2, 53) = 0
		aW(2, 54) = 14 : aX(2, 54) = 242 : aY(2, 54) = 0
		aW(2, 55) = 13 : aX(2, 55) = 256 : aY(2, 55) = 0
		aW(2, 56) = 11 : aX(2, 56) = 269 : aY(2, 56) = 0
		aW(2, 57) = 9 : aX(2, 57) = 280 : aY(2, 57) = 0 ' Z
		aW(2, 16) = 6 : aX(2, 16) = 0 : aY(2, 16) = 14
		aW(2, 17) = 9 : aX(2, 17) = 6 : aY(2, 17) = 14
		aW(2, 18) = 9 : aX(2, 18) = 15 : aY(2, 18) = 14
		aW(2, 19) = 8 : aX(2, 19) = 24 : aY(2, 19) = 14
		aW(2, 20) = 10 : aX(2, 20) = 32 : aY(2, 20) = 14
		aW(2, 21) = 8 : aX(2, 21) = 42 : aY(2, 21) = 14
		aW(2, 22) = 7 : aX(2, 22) = 50 : aY(2, 22) = 14
		aW(2, 23) = 8 : aX(2, 23) = 57 : aY(2, 23) = 14
		aW(2, 24) = 9 : aX(2, 24) = 65 : aY(2, 24) = 14
		aW(2, 15) = 11 : aX(2, 15) = 74 : aY(2, 15) = 14
		aW(2, 10) = 8 : aX(2, 10) = 85 : aY(2, 10) = 14
		aW(2, 12) = 7 : aX(2, 12) = 93 : aY(2, 12) = 14
		aW(2, 0) = 4 : aX(2, 0) = 100 : aY(2, 0) = 14
		aW(2, 13) = 4 : aX(2, 13) = 104 : aY(2, 13) = 14
		aW(2, 11) = 5 : aX(2, 11) = 108 : aY(2, 11) = 14
		aW(2, 30) = 7 : aX(2, 30) = 113 : aY(2, 30) = 14
		aW(2, 1) = 9 : aX(2, 1) = 120 : aY(2, 1) = 14
		aW(2, 6) = 8 : aX(2, 6) = 129 : aY(2, 6) = 14
		aW(2, 7) = 4 : aX(2, 7) = 137 : aY(2, 7) = 14
		aW(2, 8) = 4 : aX(2, 8) = 141 : aY(2, 8) = 14
		aW(2, 5) = 12 : aX(2, 5) = 145 : aY(2, 5) = 14
		aW(2, 9) = 10 : aX(2, 9) = 157 : aY(2, 9) = 14
		aW(2, 27) = 1 : aX(2, 27) = 167 : aY(2, 27) = 14
		aW(2, 29) = 1 : aX(2, 29) = 168 : aY(2, 29) = 14
		aW(2, 4) = 1 : aX(2, 4) = 169 : aY(2, 4) = 14
		aW(2, 3) = 1 : aX(2, 3) = 170 : aY(2, 3) = 14
		aW(2, 25) = 1 : aX(2, 25) = 171 : aY(2, 25) = 14
		aW(2, 26) = 1 : aX(2, 26) = 172 : aY(2, 26) = 14 ' Z
		aW(2, 64) = 10 : aX(2, 64) = 0 : aY(2, 64) = 0
		aW(2, 65) = 11 : aX(2, 65) = 10 : aY(2, 65) = 0
		aW(2, 66) = 10 : aX(2, 66) = 21 : aY(2, 66) = 0
		aW(2, 67) = 11 : aX(2, 67) = 31 : aY(2, 67) = 0
		aW(2, 68) = 10 : aX(2, 68) = 42 : aY(2, 68) = 0
		aW(2, 69) = 11 : aX(2, 69) = 52 : aY(2, 69) = 0
		aW(2, 70) = 11 : aX(2, 70) = 63 : aY(2, 70) = 0
		aW(2, 71) = 11 : aX(2, 71) = 74 : aY(2, 71) = 0
		aW(2, 72) = 6 : aX(2, 72) = 85 : aY(2, 72) = 0
		aW(2, 73) = 6 : aX(2, 73) = 91 : aY(2, 73) = 0
		aW(2, 74) = 12 : aX(2, 74) = 97 : aY(2, 74) = 0
		aW(2, 75) = 10 : aX(2, 75) = 109 : aY(2, 75) = 0
		aW(2, 76) = 17 : aX(2, 76) = 119 : aY(2, 76) = 0
		aW(2, 77) = 14 : aX(2, 77) = 136 : aY(2, 77) = 0
		aW(2, 78) = 13 : aX(2, 78) = 150 : aY(2, 78) = 0
		aW(2, 79) = 13 : aX(2, 79) = 163 : aY(2, 79) = 0
		aW(2, 80) = 11 : aX(2, 80) = 176 : aY(2, 80) = 0
		aW(2, 81) = 12 : aX(2, 81) = 187 : aY(2, 81) = 0
		aW(2, 82) = 10 : aX(2, 82) = 199 : aY(2, 82) = 0
		aW(2, 83) = 9 : aX(2, 83) = 209 : aY(2, 83) = 0
		aW(2, 84) = 13 : aX(2, 84) = 218 : aY(2, 84) = 0
		aW(2, 85) = 11 : aX(2, 85) = 231 : aY(2, 85) = 0
		aW(2, 86) = 14 : aX(2, 86) = 242 : aY(2, 86) = 0
		aW(2, 87) = 13 : aX(2, 87) = 256 : aY(2, 87) = 0
		aW(2, 88) = 11 : aX(2, 88) = 269 : aY(2, 88) = 0
		aW(2, 89) = 9 : aX(2, 89) = 280 : aY(2, 89) = 0
		aW(2, 14) = 7 : aX(2, 14) = 280 : aY(2, 14) = 14 ' [Titi 2.4.9]
		' Large White Numbers
		aW(3, 16) = 13 : aX(3, 16) = 0 : aY(3, 16) = 0
		aW(3, 17) = 14 : aX(3, 17) = 16 : aY(3, 17) = 0
		aW(3, 18) = 16 : aX(3, 18) = 30 : aY(3, 18) = 0
		aW(3, 19) = 16 : aX(3, 19) = 46 : aY(3, 19) = 0
		aW(3, 20) = 18 : aX(3, 20) = 62 : aY(3, 20) = 0
		aW(3, 21) = 16 : aX(3, 21) = 80 : aY(3, 21) = 0
		aW(3, 22) = 14 : aX(3, 22) = 96 : aY(3, 22) = 0
		aW(3, 23) = 15 : aX(3, 23) = 110 : aY(3, 23) = 0
		aW(3, 24) = 16 : aX(3, 24) = 125 : aY(3, 24) = 0
		aW(3, 15) = 18 : aX(3, 15) = 142 : aY(3, 15) = 0
	End Sub
	
	Private Sub GameInit()
		On Error GoTo Err_Handler
		Dim rc, hCursor As Integer
		Dim CreatureX As Creature
		Dim MapSketchX As MapSketch
		Dim FileName, StringX As String
		Dim c As Short
		Dim bm As BITMAP
		Dim CommandArgs() As String
		Dim ArgText As String
		Dim tmp As String
		Dim PauseTime, Start As Integer
		Dim logError As Short
		Dim errLevel As Short
		Dim lResult As Integer
		' [rb] Added for 2.4.6
		'Call ParseCommandLineArgs
		' Initialize Error Handler Object
		'Set oErr = New CErrorHandler
		' [rb] I borrow the GlobalDebugMode variable just to set this
		'Call oErr.Initialize("gengine.txt", GlobalDebugMode)
		gAppPath = My.Application.Info.DirectoryPath
		oErr = New CErrorHandler
		lResult = fReadValue(gAppPath & "\Settings.ini", "System", "ErrorLog", "B", "0", logError)
		lResult = fReadValue(gAppPath & "\Settings.ini", "System", "ErrorLvl", "B", "3", errLevel)
		Call oErr.Initialize(CShort(logError), CShort(errLevel), "EngineLog.txt")
		'Call oErr.Initialize(LOG_DEBUG Or LOG_TOFILE, "creatorlog.txt")
		oErr.logError("GameInit: Started", CErrorHandler.ErrorLevel.ERR_DEBUG)
		On Error GoTo Err_Handler
		' Initialize our new IO Class
		oFileSys = New clsInOut
		oGameMusic = New IMCI
		' [Titi 2.4.9] Borfaux's addition
		lResult = fReadValue(gAppPath & "\Settings.ini", "System", "DataFolder", "S", gAppPath, tmp)
		gDataPath = tmp
		If oFileSys.CheckExists(gDataPath & "\Data", clsInOut.IOActionType.Folder) Then
			gDataPath = gDataPath & "\Data"
		End If
		'UPGRADE_WARNING: Screen property Screen.MousePointer has a new behavior. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6BA9B8D2-2A32-4B6E-8D36-44949974A5B4"'
		System.Windows.Forms.Cursor.Current = System.Windows.Forms.Cursors.WaitCursor
		'UPGRADE_ISSUE: Form property frmMain.Image was not upgraded. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="CC4C7EC0-C903-48FC-ACCC-81861D12DA4A"'
		GetObject_Renamed(Image, Len(bm), bm)
		GlobalScreenX = GetSystemMetrics(SM_CXSCREEN)
		GlobalScreenY = GetSystemMetrics(SM_CYSCREEN)
		GlobalScreenColor = CShort(VB6.Format(bm.bmBitsPixel))
		' ChangeScreenSettings 800, 600, 16
		' Determine CommonScale for Pictures (default to 800x600)
		CommonScale = GlobalScreenX / 800
		PicSize = CommonScale * 0.6
		DarkDir = 1
		Randomize()
		' Set some variables
		Call SetUpBits()
		MaxSFX = -1
		TotalStepsTaken = 0
		JournalMode = 0
		' Load the Options for the Game
		'UPGRADE_ISSUE: PictureBox method picTPic.Cls was not upgraded. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="CC4C7EC0-C903-48FC-ACCC-81861D12DA4A"'
		picTPic.Cls()
		ShowText(picTPic, 0, 0, picTPic.ClientRectangle.Width, 0, bdFontElixirWhite, "Loading game options", True, True)
		picTPic.Refresh()
		oErr.logError("GameInit: " & "Loading game options - started")
		Call OptionsLoad(True)
		oErr.logError("GameInit: " & "Loading game options - finished")
		' Set TomeWizard
		TomeWizardParty = 2
		TomeWizardLevel = 0
		TomeWizardStory = 0
		TomeWizardSize = 1
		' Set up default Item picture size (bdMaxItemPics Pictures)
		picItem.Width = 64 * bdMaxItemPics
		picItem.Height = 96 * 2 + 32
		' Init all the game variables
		'UPGRADE_ISSUE: PictureBox method picTPic.Cls was not upgraded. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="CC4C7EC0-C903-48FC-ACCC-81861D12DA4A"'
		picTPic.Cls()
		ShowText(picTPic, 0, 0, picTPic.ClientRectangle.Width, 0, bdFontElixirWhite, "Loading game variables", True, True)
		picTPic.Refresh()
		Call GameInitVariables()
		oErr.logError("GameInit: " & "Loading game variables - started")
		Me.Show()
		picTPic.Width = Me.ClientRectangle.Width - 64 : picTPic.Height = 14
		picTPic.Left = 32 : picTPic.Top = Me.ClientRectangle.Height / 2
		picTPic.Visible = True
		oErr.logError("GameInit: " & "Loading game variables - finished")
		' Load bones picture for Creature
		'UPGRADE_ISSUE: PictureBox method picTPic.Cls was not upgraded. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="CC4C7EC0-C903-48FC-ACCC-81861D12DA4A"'
		picTPic.Cls()
		ShowText(picTPic, 0, 0, picTPic.ClientRectangle.Width, 0, bdFontElixirWhite, "Loading bones picture", True, True)
		picTPic.Refresh()
		oErr.logError("GameInit: " & "Loading bones picture - Started")
		CreatureX = New Creature
		CreatureX.PictureFile = "bones.bmp"
		LoadCreaturePic(CreatureX)
		' Freeze Map and Menu actions
		Frozen = True
		' Load up the tomes for selection
		TomeAction = -1
		' Load up the Roster Defaults
		oErr.logError("GameInit: " & "Loading bones picture - Finished")
		'UPGRADE_ISSUE: PictureBox method picTPic.Cls was not upgraded. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="CC4C7EC0-C903-48FC-ACCC-81861D12DA4A"'
		picTPic.Cls()
		ShowText(picTPic, 0, 0, picTPic.ClientRectangle.Width, 0, bdFontElixirWhite, "Loading Roster Defaults", True, True)
		picTPic.Refresh()
		Call CreatePCLoadWorlds()
		WorldIndex = WorldNow.Index
		' [borfaux] Moved here.  Let the Worlds INI file dictate which file to play as the Worlds
		' intro music
		' Play the Intro music
		Call PlayMusic(WorldNow.IntroMusic, Me, WorldNow.MusicFolder)
		oErr.logError("GameInit: " & "Loading Roster Defaults")
		'UPGRADE_WARNING: Screen property Screen.MousePointer has a new behavior. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6BA9B8D2-2A32-4B6E-8D36-44949974A5B4"'
		System.Windows.Forms.Cursor.Current = System.Windows.Forms.Cursors.WaitCursor
		' Now load the Story Tomes for the Tome Wizard
		'UPGRADE_ISSUE: PictureBox method picTPic.Cls was not upgraded. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="CC4C7EC0-C903-48FC-ACCC-81861D12DA4A"'
		picTPic.Cls()
		ShowText(picTPic, 0, 0, picTPic.ClientRectangle.Width, 0, bdFontElixirWhite, "Loading Story Tomes for the Tome Wizard", True, True)
		picTPic.Refresh()
		UberWizMaps = New UberWizard
		modDungeonMaker.InitUberWizMaps(UberWizMaps)
		oErr.logError("GameInit: " & "Loading Story Tomes for the Tome Wizard - Started")
		'UPGRADE_ISSUE: PictureBox method picTPic.Cls was not upgraded. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="CC4C7EC0-C903-48FC-ACCC-81861D12DA4A"'
		picTPic.Cls()
		ShowText(picTPic, 0, 0, picTPic.ClientRectangle.Width, 0, bdFontElixirWhite, "Loading Interface List", True, True)
		picTPic.Refresh()
		' Load Interface List
		oErr.logError("GameInit: " & "Loading Interface List - Started")
		GlobalInterfaceList = New Collection
		'    FileName = Dir$(gAppPath & "/Data/Interface/*", vbDirectory)
		'UPGRADE_WARNING: Dir has a new behavior. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="9B7D5ADD-D8FE-4819-A36C-6DEDAF088CC7"'
		FileName = Dir(gDataPath & "/Interface/*", FileAttribute.Directory)
		If FileName = "" Then oErr.logError("GameInit: " & vbTab & "No Interface List!!!", CErrorHandler.ErrorLevel.ERR_CRITICAL)
		Do While FileName <> ""
			If FileName <> "." And FileName <> ".." And FileName <> "Dice" Then
				' Find alphabetical spot in list
				For c = 1 To GlobalInterfaceList.Count()
					'UPGRADE_WARNING: Couldn't resolve default property of object GlobalInterfaceList(). Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
					If UCase(GlobalInterfaceList.Item(c)) > UCase(FileName) Then
						GlobalInterfaceList.Add(FileName,  , c)
						Exit For
					End If
				Next c
				If c > GlobalInterfaceList.Count() Then
					GlobalInterfaceList.Add(FileName)
				End If
			End If
			'UPGRADE_WARNING: Dir has a new behavior. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="9B7D5ADD-D8FE-4819-A36C-6DEDAF088CC7"'
			FileName = Dir()
		Loop 
		GlobalInterfaceIndex = 1
		oErr.logError("GameInit: " & "Loading Interface List - Finished")
		' Load Dice List
		'UPGRADE_ISSUE: PictureBox method picTPic.Cls was not upgraded. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="CC4C7EC0-C903-48FC-ACCC-81861D12DA4A"'
		picTPic.Cls()
		ShowText(picTPic, 0, 0, picTPic.ClientRectangle.Width, 0, bdFontElixirWhite, "Loading Dice List", True, True)
		picTPic.Refresh()
		oErr.logError("GameInit: " & "Loading Dice List - Started")
		GlobalDiceList = New Collection
		'    FileName = Dir$(gAppPath & "/Data/Interface/Dice/*.bmp", vbDirectory)
		'UPGRADE_WARNING: Dir has a new behavior. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="9B7D5ADD-D8FE-4819-A36C-6DEDAF088CC7"'
		FileName = Dir(gDataPath & "/Interface/Dice/*.bmp", FileAttribute.Directory)
		If FileName = "" Then oErr.logError("GameInit: " & vbTab & "No Dice List!!!", CErrorHandler.ErrorLevel.ERR_CRITICAL)
		Do While FileName <> ""
			If FileName <> "." And FileName <> ".." Then
				' Find alphabetical spot in list
				For c = 1 To GlobalDiceList.Count()
					'UPGRADE_WARNING: Couldn't resolve default property of object GlobalDiceList(). Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
					If UCase(GlobalDiceList.Item(c)) > UCase(FileName) Then
						GlobalDiceList.Add(FileName,  , c)
						Exit For
					End If
				Next c
				If c > GlobalDiceList.Count() Then
					GlobalDiceList.Add(FileName)
				End If
			End If
			'UPGRADE_WARNING: Dir has a new behavior. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="9B7D5ADD-D8FE-4819-A36C-6DEDAF088CC7"'
			FileName = Dir()
		Loop 
		GlobalDiceIndex = 1
		oErr.logError("GameInit: " & "Loading Dice List - Finished")
		'UPGRADE_ISSUE: PictureBox method picTPic.Cls was not upgraded. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="CC4C7EC0-C903-48FC-ACCC-81861D12DA4A"'
		picTPic.Cls()
		ShowText(picTPic, 0, 0, picTPic.ClientRectangle.Width, 0, bdFontElixirWhite, "Finishing Game Init . . .", True, True)
		picTPic.Refresh()
		oErr.logError("GameInit: " & "Finishing Game Init . . .")
		' Clean up any non-MainMap MapSketchs
		For	Each MapSketchX In UberWizMaps.MapSketchs
			If MapSketchX.IsMainMap = False Or MapSketchX.TomeIndex = 0 Then
				UberWizMaps.MapSketchs.Remove("M" & MapSketchX.Index)
			End If
		Next MapSketchX
		UberWizMaps.MainMapIndex = 0
		' Load the splash logo
		'UPGRADE_ISSUE: PictureBox method picTPic.Cls was not upgraded. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="CC4C7EC0-C903-48FC-ACCC-81861D12DA4A"'
		picTPic.Cls()
		picTPic.Visible = False
		oErr.logError("GameInit: " & "Splash Screen")
		If (Val(GlobalCredits) = 1) Then
			myLoadSplash()
		End If
		'UPGRADE_WARNING: Screen property Screen.MousePointer has a new behavior. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6BA9B8D2-2A32-4B6E-8D36-44949974A5B4"'
		System.Windows.Forms.Cursor.Current = System.Windows.Forms.Cursors.Default
		Call UpdateMainMenu()
		oErr.logError("GameInit: Loaded", CErrorHandler.ErrorLevel.ERR_DEBUG)
		picMainMenu.Visible = True
Exit_Sub: 
		Exit Sub
Err_Handler: 
		If Err.Number = 91 Then
			' not a problem with a block "With" (hehe), but here we probably have two instances of Creator running
			Call oErr.Initialize(CShort(logError), CShort(errLevel), "EngineLog2.txt")
			Resume Next
		Else
			oErr.logError("GameInit has failed", CErrorHandler.ErrorLevel.ERR_CRITICAL)
			Resume Exit_Sub
		End If
	End Sub
	
	Private Sub myLoadSplash()
		Dim aImages(2, 1) As Object
		'Dim KillSplash As Boolean
		Dim SplashNow As Object
		Dim bCheckFile As Boolean
		Dim iImage As Short
		'UPGRADE_WARNING: Couldn't resolve default property of object SplashNow. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
		SplashNow = True : KillSplash = False
		'    bCheckFile = oFileSys.CheckExists(gAppPath & "\data\interface\" & GlobalInterfaceName & "\crosscut.bmp", File, False)
		bCheckFile = oFileSys.CheckExists(gDataPath & "\interface\" & GlobalInterfaceName & "\crosscut.bmp", clsInOut.IOActionType.File, False)
		If bCheckFile = True Then
			'     aImages(0, 0) = gAppPath & "\data\interface\" & GlobalInterfaceName & "\crosscut.bmp": aImages(0, 1) = 1
			'UPGRADE_WARNING: Couldn't resolve default property of object aImages(0, 0). Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
			aImages(0, 0) = gDataPath & "\interface\" & GlobalInterfaceName & "\crosscut.bmp"
			'UPGRADE_WARNING: Couldn't resolve default property of object aImages(0, 1). Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
			aImages(0, 1) = 1
		Else
			'     aImages(0, 0) = gAppPath & "\data\stock\crosscut.bmp": aImages(0, 1) = 1
			'UPGRADE_WARNING: Couldn't resolve default property of object aImages(0, 0). Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
			aImages(0, 0) = gDataPath & "\stock\crosscut.bmp"
			'UPGRADE_WARNING: Couldn't resolve default property of object aImages(0, 1). Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
			aImages(0, 1) = 1
		End If
		'    bCheckFile = oFileSys.CheckExists(gAppPath & "\data\interface\" & GlobalInterfaceName & "\splash.bmp", File, False)
		bCheckFile = oFileSys.CheckExists(gDataPath & "\interface\" & GlobalInterfaceName & "\splash.bmp", clsInOut.IOActionType.File, False)
		If bCheckFile = True Then
			'     aImages(1, 0) = gAppPath & "\data\interface\" & GlobalInterfaceName & "\splash.bmp": aImages(0, 1) = 1
			'UPGRADE_WARNING: Couldn't resolve default property of object aImages(1, 0). Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
			aImages(1, 0) = gDataPath & "\interface\" & GlobalInterfaceName & "\splash.bmp"
			'UPGRADE_WARNING: Couldn't resolve default property of object aImages(0, 1). Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
			aImages(0, 1) = 1
		Else
			'     aImages(1, 0) = gAppPath & "\data\stock\splash.bmp": aImages(0, 1) = 1
			'UPGRADE_WARNING: Couldn't resolve default property of object aImages(1, 0). Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
			aImages(1, 0) = gDataPath & "\stock\splash.bmp"
			'UPGRADE_WARNING: Couldn't resolve default property of object aImages(0, 1). Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
			aImages(0, 1) = 1
		End If
		'UPGRADE_WARNING: Screen property Screen.MousePointer has a new behavior. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6BA9B8D2-2A32-4B6E-8D36-44949974A5B4"'
		System.Windows.Forms.Cursor.Current = System.Windows.Forms.Cursors.Default
		iImage = 0
		Do While KillSplash = False And iImage < 2
			'UPGRADE_WARNING: Couldn't resolve default property of object aImages(iImage, 1). Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
			'UPGRADE_WARNING: Couldn't resolve default property of object aImages(). Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
			Call FadeToBlack(aImages(iImage, 0), aImages(iImage, 1))
			iImage = iImage + 1
		Loop 
		'UPGRADE_WARNING: Couldn't resolve default property of object SplashNow. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
		SplashNow = False
		picMainMenu.Visible = True
	End Sub
	
	Private Sub FadeToBlack(ByVal FileName As String, ByVal bFade As Short)
		'UPGRADE_NOTE: Height was upgraded to Height_Renamed. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="A9E4979A-37FA-4718-9994-97DD76ED70A7"'
		'UPGRADE_NOTE: Width was upgraded to Width_Renamed. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="A9E4979A-37FA-4718-9994-97DD76ED70A7"'
		Dim Width_Renamed, c, Height_Renamed As Short
		'UPGRADE_NOTE: Size was upgraded to Size_Renamed. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="A9E4979A-37FA-4718-9994-97DD76ED70A7"'
		Dim Y, X, i As Short
		Dim Size_Renamed As Double
		Dim hMemSplash As Integer ', filename As String
		'UPGRADE_WARNING: Arrays in structure bmSplash may need to be initialized before they can be used. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="814DF224-76BD-4BB4-BFFB-EA359CB9FC48"'
		Dim bmSplash As BITMAPINFO
		'UPGRADE_WARNING: Array bmFadeToBlack may need to have individual elements initialized. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="B97B714D-9338-48AC-B03F-345B617E2B02"'
		Dim bmFadeToBlack(255) As BITMAPINFO
		Dim rc, lpMem, TransparentRGB As Integer
		
		If KillSplash = False Then
			ReadBitmapFile(FileName, bmSplash, hMemSplash, TransparentRGB)
			MakeFadeToBlack(bmSplash, bmFadeToBlack)
			lpMem = GlobalLock(hMemSplash)
			'UPGRADE_ISSUE: PictureBox property picMap.hdc was not upgraded. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="CC4C7EC0-C903-48FC-ACCC-81861D12DA4A"'
			rc = SetStretchBltMode(picMap.hdc, 3)
			Width_Renamed = bmSplash.bmiHeader.biWidth
			Height_Renamed = bmSplash.bmiHeader.biHeight
			Y = (Me.ClientRectangle.Height - Height_Renamed - picMenu.Height) / 2
			X = (Me.ClientRectangle.Width - Width_Renamed - 64) / 2
			'UPGRADE_ISSUE: PictureBox property picMap.hdc was not upgraded. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="CC4C7EC0-C903-48FC-ACCC-81861D12DA4A"'
			rc = StretchDIBits(picMap.hdc, X, Y, Width_Renamed, Height_Renamed, 0, 0, Width_Renamed, Height_Renamed, lpMem, bmSplash, DIB_RGB_COLORS, SRCCOPY)
			
			
			If bFade = 0 Then ' Fade In
				For c = 0 To 255
					'UPGRADE_ISSUE: PictureBox property picMap.hdc was not upgraded. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="CC4C7EC0-C903-48FC-ACCC-81861D12DA4A"'
					rc = StretchDIBits(picMap.hdc, X, Y, Width_Renamed, Height_Renamed, 0, 0, Width_Renamed, Height_Renamed, lpMem, bmFadeToBlack(c), DIB_RGB_COLORS, SRCCOPY)
					picMap.Refresh()
					Sleep(10)
					System.Windows.Forms.Application.DoEvents()
					If KillSplash = True Then
						Exit For
					End If
				Next c
			ElseIf bFade = 1 Then  ' Fade Out
				picMap.Refresh()
				For c = 1 To 250
					Sleep(10)
					System.Windows.Forms.Application.DoEvents()
					If KillSplash = True Then
						Exit For
					End If
				Next c
				
				For c = 255 To 0 Step -5
					'UPGRADE_ISSUE: PictureBox property picMap.hdc was not upgraded. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="CC4C7EC0-C903-48FC-ACCC-81861D12DA4A"'
					rc = StretchDIBits(picMap.hdc, X, Y, Width_Renamed, Height_Renamed, 0, 0, Width_Renamed, Height_Renamed, lpMem, bmFadeToBlack(c), DIB_RGB_COLORS, SRCCOPY)
					picMap.Refresh()
					Sleep(10)
					System.Windows.Forms.Application.DoEvents()
					If KillSplash = True Then
						Exit For
					End If
				Next c
			End If
			'UPGRADE_ISSUE: PictureBox method picMap.Cls was not upgraded. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="CC4C7EC0-C903-48FC-ACCC-81861D12DA4A"'
            picMap.Controls.Clear()
			rc = GlobalUnlock(hMemSplash)
			rc = GlobalFree(hMemSplash)
		End If
		
	End Sub
	
	'UPGRADE_NOTE: Map was upgraded to Map_Renamed. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="A9E4979A-37FA-4718-9994-97DD76ED70A7"'
	Private Function SorceryQueRune(ByRef CreatureX As Creature, ByRef RuneToQue As Short) As Short
		Dim Map_Renamed As Object
		Dim rc, c, Cost As Short
		' If blank rune to que then bad
		If RuneToQue < 0 Then
			SorceryQueRune = False
			Exit Function
		End If
		' Que Rune to next available slot.
		For c = CreatureX.RuneQueLimit To 5
			If CreatureX.Runes(c) = 0 Then
				Exit For
			End If
		Next c
		' If RunePool is drained or there's no room left in the que, do nothing.
		'UPGRADE_WARNING: Untranslated statement in SorceryQueRune. Please check source code.
		' Show before take Rune
		BorderDrawSides(True, RuneToQue)
		BorderDrawButtons(0)
		Me.Refresh()
		Call PlayClickSnd(modIOFunc.ClickType.ifClickCast)
		' Que the Rune
		CreatureX.Runes(c) = RuneToQue + 1
		'UPGRADE_WARNING: Couldn't resolve default property of object Map.Runes. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
		Map.Runes(RuneToQue) = Map.Runes(RuneToQue) - 1
		' Show after
		BorderDrawSides(True, -1)
		BorderDrawButtons(0)
		Me.Refresh()
		SorceryQueRune = True
	End Function
	
	Private Sub SorceryTrigger(ByRef CreatureX As Creature, ByRef OwnerObjectX As Object, ByRef TriggerX As Trigger)
		Dim Found, c, i, NoFail As Short
		Dim SpellRunes, QuedRunes As String
		Dim EterniumCost As Short
		QuedRunes = ""
		For c = 5 To 0 Step -1
			QuedRunes = QuedRunes & Chr(CreatureX.Runes(c) + 64)
		Next c
		' Check for Rune match
		'UPGRADE_WARNING: Couldn't resolve default property of object TriggerX.Statements(1).Statement. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
		If TriggerX.Statements.Item(1).Statement = 23 Then ' Rune Statement
			SpellRunes = ""
			For c = 5 To 0 Step -1
				'UPGRADE_WARNING: Couldn't resolve default property of object TriggerX.Statements().B. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
				If TriggerX.Statements.Item(1).B(c) > 0 Then
					'UPGRADE_WARNING: Couldn't resolve default property of object TriggerX.Statements().B. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
					SpellRunes = SpellRunes & Chr(TriggerX.Statements.Item(1).B(c) + 64)
				End If
			Next c
			For i = 6 To CreatureX.RuneTop Step -2
				If IsBetween(InStr(Mid(QuedRunes, 7 - i), SpellRunes), 1, 2) Then
					If i - CreatureX.RuneTop < 2 Then
						MessageShow("Casting " & TriggerX.Name, 0)
						' Check if Fizzle
						'UPGRADE_WARNING: Couldn't resolve default property of object TriggerX.Statements().B. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
						If Int(Rnd() * 100) + 1 < GlobalSpellFizzleChance And ((TriggerX.Statements.Item(1).B(6) And &H1) > 0) Then
							If picGrid.Visible = True Then
								MessageShow(TriggerX.Name & " fizzles!", 0)
							Else
								DialogDM(TriggerX.Name & " fizzles!")
							End If
						Else
							' Set GlobalSaveStyle value if the Trigger has the Rune Statement with Save Checked
							'UPGRADE_WARNING: Couldn't resolve default property of object TriggerX.Statements().B. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
							If (TriggerX.Statements.Item(1).B(6) And &H2) > 0 Then
								GlobalSaveStyle = TriggerX.Styles
							Else
								GlobalSaveStyle = 0
							End If
							' Cast the Spell
							CreatureNow = CreatureWithTurn
							ItemNow = CreatureWithTurn.ItemInHand
							If picGrid.Visible = True Then
								CreatureTarget = CreatureWithTurn.CreatureTarget
							End If
							NoFail = FireTrigger(OwnerObjectX, TriggerX)
							' Turn off the GlobalSaveStyle
							GlobalSaveStyle = 0
						End If
					Else
						MessageShow("Invoking " & TriggerX.Name, 0)
					End If
				End If
			Next i
		End If
	End Sub
	
	Public Function SorceryMatchRunes(ByRef CreatureX As Creature, ByRef Runes() As Byte, ByRef CastIt As Short) As Short
		Dim k, n, c, i, Found As Short
		Dim OldPointer As Short
		' Returns 0= Ok, 1 = Not Enough ActionPoints, 2 = Undefined, 3 = Too Many Runes, 4 = Not Enough Runes, 99 = Undefined
		'UPGRADE_WARNING: Screen property Screen.MousePointer has a new behavior. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6BA9B8D2-2A32-4B6E-8D36-44949974A5B4"'
		OldPointer = System.Windows.Forms.Cursor.Current
		'UPGRADE_WARNING: Screen property Screen.MousePointer has a new behavior. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6BA9B8D2-2A32-4B6E-8D36-44949974A5B4"'
		System.Windows.Forms.Cursor.Current = System.Windows.Forms.Cursors.WaitCursor
		SorceryMatchRunes = 0
		' Count the Runes (0 to 5)
		For c = 0 To 5
			If Runes(c) > 0 Then
				n = c
			End If
		Next c
		' Match the Runes
		Found = False
		For c = 0 To CreatureX.RuneQueLimit - 1
			' Check if beyond limit of RuneQue
			If c + n > 5 Then
				Exit For
			End If
			' Check for matching Runes
			For i = 0 To n
				If CreatureX.Runes(c + i) <> Runes(i) And CreatureX.Runes(c + i) <> 0 Then
					Exit For
				End If
			Next i
			If i > n Then
				Found = True
				Exit For
			End If
		Next c
		' If piggy-back, then set starting Rune and new length
		If Found = True Then
			c = CreatureX.RuneQueLimit - c
			n = n - c
		Else
			c = 0
		End If
		' Check if have enough ActionPoints,  room in RuneQue
		If picGrid.Visible = True And CreatureX.ActionPoints < 10 Then
			SorceryMatchRunes = 1
		ElseIf c + n > 5 Then 
			SorceryMatchRunes = 3
		ElseIf CastIt = True Then 
			CreatureX.ActionPoints = 0
			If picGrid.Visible = True Then
				CombatDraw()
			End If
			For i = c To c + n
				k = SorceryQueRune(CreatureX, Runes(i) - 1)
				If k = False Then
					SorceryMatchRunes = 4
					Exit For
				End If
			Next i
			' Lock in the RuneQueLimit
			For c = 5 To 0 Step -1
				If CreatureX.Runes(c) > 0 Then
					CreatureX.RuneQueLimit = c + 1
					Exit For
				End If
			Next c
			' Invoke (if possible)
			If CreatureX.RuneTop = 0 Then
				If picGrid.Visible = True Then
					CreatureX.RuneTop = 2
				Else
					CreatureX.RuneTop = Least((CreatureX.RuneQueLimit), 6)
				End If
				SorceryCast(CreatureX)
			End If
		Else
			' Check for available in RunePool
			Found = False
			For i = c To c + n
				'UPGRADE_WARNING: Untranslated statement in SorceryMatchRunes. Please check source code.
			Next i
			If Found = True Then
				SorceryMatchRunes = 4
			Else
				SorceryMatchRunes = 0
			End If
		End If
		'UPGRADE_ISSUE: Screen property Screen.MousePointer does not support custom mousepointers. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="45116EAB-7060-405E-8ABE-9DBB40DC2E86"'
		'UPGRADE_WARNING: Screen property Screen.MousePointer has a new behavior. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6BA9B8D2-2A32-4B6E-8D36-44949974A5B4"'
		System.Windows.Forms.Cursor.Current = OldPointer
	End Function
	
	Public Sub SorceryInvokeNext(ByRef CreatureX As Creature)
		If CreatureX.RuneTop > 0 Then
			CreatureX.RuneTop = CreatureX.RuneTop + 2
			SorceryCast(CreatureX)
		End If
	End Sub
	
	Private Sub SorceryCastTrigger(ByRef CreatureX As Creature, ByRef ParentX As Object)
		Dim TriggerX As Trigger
		'UPGRADE_WARNING: Couldn't resolve default property of object ParentX.Triggers. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
		For	Each TriggerX In ParentX.Triggers
			If TriggerX.TriggerType = bdOnCast Or TriggerX.TriggerType = bdOnCastTarget Then
				SorceryTrigger(CreatureX, ParentX, TriggerX)
			End If
			SorceryCastTrigger(CreatureX, TriggerX)
		Next TriggerX
	End Sub
	
	Public Sub SorceryCast(ByRef CreatureX As Creature)
		Dim i, c, Found As Short
		Dim TriggerX, TriggerZ As Trigger
		Dim ItemX As Item
		' Find OnCast Triggers on CreatureX
		SorceryCastTrigger(CreatureX, CreatureX)
		' Find OnCast Triggers on ready items
		For	Each ItemX In CreatureX.Items
			If ItemX.IsReady = True Then
				SorceryCastTrigger(CreatureX, ItemX)
			End If
		Next ItemX
		' Clear out Runes if not in Combat or at end of Que in combat
		If picGrid.Visible = False Or (CreatureX.RuneTop >= CreatureX.RuneQueLimit And CreatureX.RuneTop > 2) Then
			For c = 0 To 5
				CreatureX.Runes(c) = 0
			Next 
			CreatureX.RuneTop = 0
			CreatureX.RuneQueLimit = 0
		End If
	End Sub
	
	Private Sub TalkMove(ByRef AtX As Short, ByRef AtY As Short)
		Dim c, PosY As Short
		PosY = ReplyHeight(0) - 4
		For c = 1 To ReplyTop
			If PointIn(AtX, AtY, 119, PosY, 306, ReplyHeight(c)) And c <> ReplySelect Then
				ReplySelect = c
				TalkList()
				Exit For
			End If
			PosY = PosY + ReplyHeight(c) + 4
		Next c
	End Sub
	
	Private Sub TalkReply(ByRef Index As Short)
		Dim Found, NoFail, c, Abort, FactFound As Short
		Dim ConvoX As Conversation
		Dim TopicX As Topic
		Dim FactoidX, FactoidZ As Factoid
		' Set up Creature and Conversation
		CreatureTarget = ScrollList.Item(ScrollSelect)
		CreatureNow = CreatureWithTurn
		If CreatureTarget.CurrentConvo = 0 Then
			ReplyList(0) = CreatureTarget.Name & " doesn't want to talk right now."
			ConvoX = New Conversation
		ElseIf Index = 0 And CreatureTarget.Conversations.Count() > 0 Then 
			' Establish default Replies
			Found = False
			For	Each ConvoX In CreatureTarget.Conversations
				If ConvoX.Index = CreatureTarget.CurrentConvo Then
					ConvoX = CreatureTarget.Conversations.Item("C" & CreatureTarget.CurrentConvo)
					Found = True
					Exit For
				End If
			Next ConvoX
			' If default convo is misplaced, can't talk.
			If Found = False Then
				ReplyList(0) = CreatureTarget.Name & " doesn't want to talk right now."
				ConvoX = New Conversation
			End If
			If ConvoX.HaveTalked = True And Len(ConvoX.SecondTalk) > 0 Then
				ReplyList(0) = ConvoX.SecondTalk
			Else
				ConvoX.HaveTalked = True
				ReplyList(0) = ConvoX.FirstTalk
			End If
		Else
			ConvoX = CreatureTarget.Conversations.Item("C" & CreatureTarget.CurrentConvo)
			TopicX = ConvoX.Topics.Item("Q" & ReplyIndex(Index))
			ReplyList(0) = TopicX.Reply
			'        Abort = False
			Select Case TopicX.Action
				Case 0 ' Do Nothing
				Case 1 ' SetConvo
					CreatureTarget.CurrentConvo = TopicX.ActionRef
					ConvoX = CreatureTarget.Conversations.Item("C" & CreatureTarget.CurrentConvo)
				Case 2 ' Remove
					TopicX.Default_Renamed = False
				Case 3 ' Add
					'UPGRADE_WARNING: Couldn't resolve default property of object CreatureTarget.Conversations().Topics. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
					CreatureTarget.Conversations.Item("C" & CreatureTarget.CurrentConvo).Topics("Q" & TopicX.ActionRef).Default = True
				Case 4 ' Replace
					TopicX.Default_Renamed = False
					'UPGRADE_WARNING: Couldn't resolve default property of object CreatureTarget.Conversations().Topics. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
					CreatureTarget.Conversations.Item("C" & CreatureTarget.CurrentConvo).Topics("Q" & TopicX.ActionRef).Default = True
				Case 5 ' Close
					' [Titi 2.4.9] show the final reply if there's one
					'                picTalk.Cls
					'                ShowText picTalk, 119, 42, 312, 306, bdFontNoxiousWhite, TopicX.Reply, False, False
					'                TopicX.Say = vbNullString
					If ReplyList(0) <> "He/She replies...." Then
						DialogBrief(CreatureTarget, ReplyList(0))
					End If
					Abort = True
			End Select
			' Fire any Post-Topic Triggers
			NoFail = FireTriggers(TopicX, bdPostTopic)
			If NoFail = False Or Abort = True Then
				picTalk.Visible = False
				Frozen = False
				picMap.Focus()
				Exit Sub
			End If
		End If
		' List Topics
		c = 1
		For	Each TopicX In ConvoX.Topics
			' Check for Factoid existence
			Found = True
			For	Each FactoidX In TopicX.Factoids
				FactFound = False
				For	Each FactoidZ In Tome.Factoids
					If FactoidX.Text = FactoidZ.Text Then
						FactFound = True
						Exit For
					End If
				Next FactoidZ
				If FactFound = False Then
					Found = False
					Exit For
				End If
			Next FactoidX
			' Check for PreTopic Triggers
			NoFail = FireTriggers(TopicX, bdPreTopic)
			If NoFail And Found = True Then
				If TopicX.Default_Renamed = True And c < 11 Then
					ReplyList(c) = TopicX.Say
					ReplyIndex(c) = TopicX.Index
					c = c + 1
				End If
			End If
		Next TopicX
		ReplyTop = c - 1
		TalkList()
	End Sub
	
	Private Sub TalkClick(ByRef AtX As Short, ByRef AtY As Short, ByRef ButtonDown As Short)
		Dim c, PosY As Short
		If PointIn(AtX, AtY, 350, 363, 90, 18) Then
			ShowButton(picTalk, 350, 363, "Done", ButtonDown)
			picTalk.Refresh()
			If ButtonDown = False Then
				Call PlayClickSnd(modIOFunc.ClickType.ifClick)
				picTalk.Visible = False
				'[borf] 2.4.6a Freeze issue - Phule's example #1
				' If a Dialog has been covered by the talk form
				' the Dialog is not present but is still caught in a DoEvent
				' If this is the case we don't want to return to the map
				If Not bDialog Then
					Frozen = False
					picMap.Focus()
				End If
			End If
		ElseIf PointIn(AtX, AtY, 18, 38, 66, 316) And ButtonDown = False Then 
			' Pick a person to talk with
			ScrollSelect = Least(ScrollTop + Greatest(Least(Int((AtY - 38) / 80), 3), 0), ScrollList.Count())
			TalkReply(0)
			Call PlayClickSnd(modIOFunc.ClickType.ifClick)
		ElseIf PointIn(AtX, AtY, 85, 37 + 18, 18, ScrollThumbY - (37 + 18)) And ButtonDown Then 
			' Big Scroll Up
			ScrollTop = ScrollTop - 3
			TalkList()
			Call PlayClickSnd(modIOFunc.ClickType.ifClick)
		ElseIf PointIn(AtX, AtY, 85, 37, 18, 18) Then 
			' Scroll Up
			If ButtonDown Then
				ScrollThumbY = ScrollBarShow(picConvoList, 85, 37, 317, ScrollTop, ScrollList.Count() - 3, 1)
				picConvoList.Refresh()
			Else
				ScrollTop = ScrollTop - 1
				TalkList()
				Call PlayClickSnd(modIOFunc.ClickType.ifClick)
			End If
		ElseIf PointIn(AtX, AtY, 85, ScrollThumbY + 18, 18, (37 + 317) - 36 - ScrollThumbY) And ButtonDown Then 
			' Big Scroll Down
			ScrollTop = ScrollTop + 3
			TalkList()
			Call PlayClickSnd(modIOFunc.ClickType.ifClick)
		ElseIf PointIn(AtX, AtY, 85, (37 + 317) - 18, 18, 18) Then 
			' Scroll Down
			If ButtonDown Then
				ScrollThumbY = ScrollBarShow(picConvoList, 85, 37, 317, ScrollTop, ScrollList.Count() - 3, 2)
				picConvoList.Refresh()
			Else
				ScrollTop = Greatest(Least(ScrollTop + 1, ScrollList.Count() - 3), 1)
				TalkList()
				Call PlayClickSnd(modIOFunc.ClickType.ifClick)
			End If
		ElseIf PointIn(AtX, AtY, 119, 42, 312, 306) And ButtonDown = False Then 
			' Choose a Reply
			PosY = ReplyHeight(0) - 4
			For c = 1 To ReplyTop
				If PointIn(AtX, AtY, 119, PosY, 306, ReplyHeight(c)) Then
					Call PlayClickSnd(modIOFunc.ClickType.ifClick)
					TalkReply(c)
					Exit For
				End If
				PosY = PosY + ReplyHeight(c) + 4
			Next c
		End If
	End Sub
	
	Private Sub TalkList()
		Dim c, PosY As Short
		Dim rc As Integer
		Dim CreatureX As Creature
		' List Portraits
		'UPGRADE_ISSUE: PictureBox method picTalk.Cls was not upgraded. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="CC4C7EC0-C903-48FC-ACCC-81861D12DA4A"'
		picTalk.Cls()
		ShowText(picTalk, 0, 12, picTalk.Width, 14, bdFontElixirWhite, "Talk with " & CreatureTarget.Name, True, False)
		ScrollTop = Greatest(Least(ScrollTop, ScrollList.Count() - 3), 1)
		ScrollSelect = Least(Greatest(ScrollSelect, 1), ScrollList.Count())
		For c = ScrollTop To Least(ScrollTop + 3, ScrollList.Count())
			' Load and show picture
			CreatureX = ScrollList.Item(c)
			LoadCreaturePic(CreatureX)
			'UPGRADE_ISSUE: PictureBox property picFaces.hdc was not upgraded. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="CC4C7EC0-C903-48FC-ACCC-81861D12DA4A"'
			'UPGRADE_ISSUE: PictureBox property picTalk.hdc was not upgraded. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="CC4C7EC0-C903-48FC-ACCC-81861D12DA4A"'
			rc = BitBlt(picTalk.hdc, 18, 38 + 80 * (c - ScrollTop), 66, 76, picFaces.hdc, bdFaceMin + CreatureX.Pic * 66, 0, SRCCOPY)
			' Show select indicator
			If c = ScrollSelect Then
				'UPGRADE_ISSUE: PictureBox property picFaces.hdc was not upgraded. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="CC4C7EC0-C903-48FC-ACCC-81861D12DA4A"'
				'UPGRADE_ISSUE: PictureBox property picTalk.hdc was not upgraded. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="CC4C7EC0-C903-48FC-ACCC-81861D12DA4A"'
				rc = BitBlt(picTalk.hdc, 18, 38 + 80 * (c - ScrollTop), 66, 76, picFaces.hdc, bdFaceSelect + 66, 0, SRCAND)
				'UPGRADE_ISSUE: PictureBox property picFaces.hdc was not upgraded. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="CC4C7EC0-C903-48FC-ACCC-81861D12DA4A"'
				'UPGRADE_ISSUE: PictureBox property picTalk.hdc was not upgraded. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="CC4C7EC0-C903-48FC-ACCC-81861D12DA4A"'
				rc = BitBlt(picTalk.hdc, 18, 38 + 80 * (c - ScrollTop), 66, 76, picFaces.hdc, bdFaceSelect, 0, SRCPAINT)
			End If
		Next c
		' Show Current Reply and Topics
		If CreatureTarget.CurrentConvo = 0 Then
			ShowText(picTalk, 119, 42, 312, 306, bdFontNoxiousWhite, CreatureTarget.Name & " has nothing to say.", False, False)
		Else
			' List Reply or First/Second Talk
			PosY = 42 + 16 + ShowText(picTalk, 119, 42, 312, 306, bdFontNoxiousWhite, ReplyList(0), False, True)
			ReplyHeight(0) = PosY
			' List Topics
			For c = 1 To ReplyTop
				If ReplySelect = c Then
					ReplyHeight(c) = ShowText(picTalk, 119, PosY, 312, 152, bdFontNoxiousGold, ReplyList(c), False, True)
				Else
					ReplyHeight(c) = ShowText(picTalk, 119, PosY, 312, 152, bdFontNoxiousGrey, ReplyList(c), False, True)
				End If
				PosY = PosY + ReplyHeight(c) + 4
			Next c
		End If
		' Show Scrollbar
		If ScrollList.Count() > 4 Then
			ScrollThumbY = ScrollBarShow(picTalk, 85, 37, 317, Least(ScrollTop, ScrollList.Count() - 3), ScrollList.Count() - 3, 0)
		End If
		ShowButton(picTalk, 350, 363, "Done", False)
		' Refresh the screen
		'    If picConvo.Visible = True Then
		'        picConvo.Enabled = False
		'    End If
		picTalk.Refresh()
		picTalk.BringToFront()
		picTalk.Visible = True
	End Sub
	
	Private Sub TalkSetup()
		Dim CreatureX As Creature
		Dim OldFrozen As Short
		MenuNow = bdMenuTalkWith
		ScrollList = New Collection
		For	Each CreatureX In EncounterNow.Creatures
			If CreatureX.Conversations.Count() > 0 And CreatureX.HPNow > 0 Then
				ScrollList.Add(CreatureX)
			End If
		Next CreatureX
		For	Each CreatureX In Tome.Creatures
			If CreatureX.CurrentConvo > 0 And CreatureX.HPNow > 0 And CreatureX.Index <> CreatureWithTurn.Index Then
				ScrollList.Add(CreatureX)
			End If
		Next CreatureX
		If ScrollList.Count() < 1 Then
			DialogDM("There's nobody to talk with here.")
		Else
			ScrollTop = 1
			ScrollSelect = 1
			TalkReply(0)
			Frozen = True
		End If
	End Sub
	
	Private Sub TomeDelete(ByRef TomeX As Tome)
		Dim FileName As String
		Dim c As Short
		Dim CreatureX As Creature
		Dim TomeZ As Tome
		Dim CreatureZ As Creature
		Dim sPath As String
		On Error GoTo ErrorHandler
		DialogSetUp(modGameGeneral.DLGTYPE.bdDlgNoReply)
		DialogSetButton(1, "No")
		DialogSetButton(2, "Yes")
		DialogShow("DM", "Do you really wish to delete " & TomeX.Name & "?")
		picConvo.Visible = False
		' If Yes then Remove the Tome Files
		If ConvoAction = 0 Then
			sPath = gAppPath & "\World\" & WorldNow.Name & "\"
			' Read in the Tome (if you can find it)
			FileName = sPath & TomeX.FileName
			If oFileSys.CheckExists(FileName, clsInOut.IOActionType.File) Then
				c = FreeFile
				TomeZ = New Tome
				FileOpen(c, FileName, OpenMode.Binary)
				TomeZ.ReadFromFile(c)
				FileClose(c)
				' Free up any characters in the Tome
				For	Each CreatureX In TomeZ.Creatures
					FileName = gAppPath & "\Roster\" & WorldNow.Name & "\"
					If CreatureX.RequiredInTome = False And oFileSys.CheckExists(FileName & CreatureX.Name & ".rsc", clsInOut.IOActionType.File) Then
						c = FreeFile
						CreatureZ = New Creature
						FileOpen(c, FileName & CreatureX.Name & ".rsc", OpenMode.Binary)
						CreatureZ.ReadFromFile(c)
						FileClose(c)
						CreatureZ.OnAdventure = False
						c = FreeFile
						FileOpen(c, FileName & CreatureZ.Name & ".rsc", OpenMode.Binary)
						CreatureZ.SaveToFile(c)
						FileClose(c)
					End If
				Next CreatureX
			End If
			' Clean out folder of origin under Tomes
			'UPGRADE_WARNING: Dir has a new behavior. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="9B7D5ADD-D8FE-4819-A36C-6DEDAF088CC7"'
			FileName = Dir(TomeX.FullPath & "\" & WorldNow.Name & "\" & VB.Left(TomeX.FileName, Len(TomeX.FileName) - 4) & "*.*")
			Do Until FileName = ""
				Kill(TomeX.FullPath & "\" & FileName)
				'UPGRADE_WARNING: Dir has a new behavior. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="9B7D5ADD-D8FE-4819-A36C-6DEDAF088CC7"'
				FileName = Dir()
			Loop 
			' Clean out World folder
			'UPGRADE_WARNING: Dir has a new behavior. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="9B7D5ADD-D8FE-4819-A36C-6DEDAF088CC7"'
			FileName = Dir(gAppPath & "\World\" & WorldNow.Name & "\" & VB.Left(TomeX.FileName, Len(TomeX.FileName) - 5) & "*.*")
			Do Until FileName = ""
				Kill(gAppPath & "\World\" & WorldNow.Name & "\" & FileName)
				'UPGRADE_WARNING: Dir has a new behavior. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="9B7D5ADD-D8FE-4819-A36C-6DEDAF088CC7"'
				FileName = Dir()
			Loop 
			' Remove it from the TomeList
			TomeList.Remove("T" & TomeX.Index)
			TomeNewList(TomeIndex)
		End If
ErrorHandler: 
		oErr.logError("TomeDelete")
		Resume Next
	End Sub
	
	'UPGRADE_NOTE: Tome was upgraded to Tome_Renamed. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="A9E4979A-37FA-4718-9994-97DD76ED70A7"'
	Private Sub TomeLoadParty(ByRef TomeX As Tome)
		Dim Tome_Renamed As Object
		' When this routine finishes, Tome contains the "right" version
		Dim c As Short
		Dim CreatureX, CreatureZ As Creature
		Dim FileName As String
		' If Tome is InPlay, then load from \World
		If TomeX.IsInPlay = True Then
			c = FreeFile
			'Open gAppPath & "\World\" & TomeX.FileName For Binary As c
			FileOpen(c, gAppPath & "\World\" & WorldNow.Name & "\" & TomeX.FileName, OpenMode.Binary)
			Tome = New Tome
			'UPGRADE_WARNING: Couldn't resolve default property of object Tome.ReadFromFile. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
			Tome.ReadFromFile(c)
			FileClose(c)
			' If reset, then restore the Party
			If TomeX.IsReset = True Then
				'UPGRADE_WARNING: Couldn't resolve default property of object Tome.Creatures. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
				For	Each CreatureX In Tome.Creatures
					'FileName = Dir(gAppPath & "\Roster\" & CreatureX.Name & ".rsc")
					'UPGRADE_WARNING: Dir has a new behavior. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="9B7D5ADD-D8FE-4819-A36C-6DEDAF088CC7"'
					FileName = Dir(gAppPath & "\Roster\" & WorldNow.Name & "\" & CreatureX.Name & ".rsc")
					If FileName <> "" Then
						' Read the existing file
						c = FreeFile
						FileOpen(c, gAppPath & "\Roster\" & WorldNow.Name & "\" & CreatureX.Name & ".rsc", OpenMode.Binary)
						CreatureZ = New Creature
						CreatureZ.ReadFromFile(c)
						FileClose(c)
						' Set to not on adventure and save back
						CreatureZ.OnAdventure = False
						c = FreeFile
						FileOpen(c, gAppPath & "\Roster\" & WorldNow.Name & "\" & CreatureX.Name & ".rsc", OpenMode.Binary)
						CreatureZ.SaveToFile(c)
						FileClose(c)
					End If
				Next CreatureX
			End If
		End If
		' If InPlay and Reset or Not InPlay, load the Tome from it's FullPath
		If (TomeX.IsInPlay And TomeX.IsReset = True) Or TomeX.IsInPlay = False Then
			c = FreeFile
			FileOpen(c, TomeX.FullPath & "\" & TomeX.FileName, OpenMode.Binary)
			Tome = New Tome
			'UPGRADE_WARNING: Couldn't resolve default property of object Tome.ReadFromFile. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
			Tome.ReadFromFile(c)
			FileClose(c)
			'UPGRADE_WARNING: Couldn't resolve default property of object Tome.OnAdventure. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
			Tome.OnAdventure = False
		End If
		' In anycase, make certain the FileName and FullPath are correct
		'UPGRADE_WARNING: Couldn't resolve default property of object Tome.FileName. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
		Tome.FileName = TomeX.FileName
		'UPGRADE_WARNING: Couldn't resolve default property of object Tome.FullPath. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
		Tome.FullPath = TomeX.FullPath
	End Sub
	
	Private Function TomeInPlayCheck() As Object
		Dim c As Short
		Dim TomeX As Tome
		Dim CreatureX As Creature
		'UPGRADE_WARNING: Couldn't resolve default property of object TomeInPlayCheck. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
		TomeInPlayCheck = False
		' If InPlay then check for existing Party
		'UPGRADE_WARNING: Couldn't resolve default property of object TomeList(TomeIndex).IsInPlay. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
		If TomeList.Item(TomeIndex).IsInPlay = True Then
			' Get the Tome in the World folder and check for OnAdventure
			c = FreeFile
			'Open gAppPath & "\World\" & TomeList(TomeIndex).FileName For Binary As c
			'UPGRADE_WARNING: Couldn't resolve default property of object TomeList(TomeIndex).FileName. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
			FileOpen(c, gAppPath & "\World\" & WorldNow.Name & "\" & TomeList.Item(TomeIndex).FileName, OpenMode.Binary)
			TomeX = New Tome
			TomeX.ReadFromFileHeader(c)
			FileClose(c)
			If TomeX.OnAdventure = True Then
				'UPGRADE_WARNING: Couldn't resolve default property of object TomeInPlayCheck. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
				TomeInPlayCheck = True
			End If
		End If
	End Function
	
	Private Sub TomeInfo(ByRef TomeX As Tome)
		DialogSetUp(modGameGeneral.DLGTYPE.bdDlgNoReply)
		DialogSetButton(1, "Done")
		DialogDM("Folder Name: " & Mid(TomeX.FullPath, Len(gAppPath & "\Tomes\") + 1) & Chr(13) & "Tome File: " & TomeX.FileName & Chr(13) & "Modified: " & VB6.Format(FileDateTime(TomeX.FullPath & "\" & TomeX.FileName), "dddd, mmm d yyyy hh:mm AMPM") & Chr(13) & "File Size: " & FileLen(TomeX.FullPath & "\" & TomeX.FileName) & " Bytes")
		picConvo.Visible = False
	End Sub
	
	Private Sub TomeCopy(ByRef TomeX As Tome)
		Dim FileName As String
		Dim c As Short
		' If InPlay, read in the Tome
		' If not InPlay or resetting the Tome, clean up and copy the Tome and Area files
		If TomeX.IsInPlay = False Or TomeX.IsReset = True Then
			' Delete any copies already in the \world directory
			'UPGRADE_WARNING: Dir has a new behavior. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="9B7D5ADD-D8FE-4819-A36C-6DEDAF088CC7"'
			FileName = Dir(gAppPath & "\world\" & WorldNow.Name & "\" & Mid(TomeX.FileName, 1, Len(TomeX.FileName) - 4) & "*.*")
			If FileName <> "" Then
				Kill(gAppPath & "\world\" & WorldNow.Name & "\" & Mid(TomeX.FileName, 1, Len(TomeX.FileName) - 4) & "*.*")
			End If
			' Save the Tome to the \world directory
			c = FreeFile
			FileOpen(c, gAppPath & "\world\" & WorldNow.Name & "\" & TomeX.FileName, OpenMode.Binary)
			TomeX.OnAdventure = False
			TomeX.SaveToFile(c)
			FileClose(c)
			' Copy Area files to \world directory
			'UPGRADE_WARNING: Dir has a new behavior. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="9B7D5ADD-D8FE-4819-A36C-6DEDAF088CC7"'
			FileName = Dir(TomeX.FullPath & "\*.rsa")
			Do Until FileName = ""
				FileCopy(TomeX.FullPath & "\" & FileName, gAppPath & "\world\" & WorldNow.Name & "\" & FileName)
				'UPGRADE_WARNING: Dir has a new behavior. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="9B7D5ADD-D8FE-4819-A36C-6DEDAF088CC7"'
				FileName = Dir()
			Loop 
		End If
	End Sub
	
	Private Sub TomeStartClearGlobals()
		GlobalAttackRoll = 0
		GlobalArmorRoll = 0
		GlobalHitLocation = ""
		GlobalDamageRoll = 0
		GlobalDieTypeRoll = 0
		GlobalDieCountRoll = 0
		GlobalOffer = 0
		GlobalSkillLevel = 0
		GlobalDamageStyle = 0
		GlobalPickLockChance = 0
		GlobalRemoveTrapChance = 0
		GlobalIntegerA = 0
		GlobalIntegerB = 0
		GlobalIntegerC = 0
		GlobalTextA = ""
		GlobalTextB = ""
		GlobalTextC = ""
		GlobalDayName = ""
		GlobalMoonName = ""
		GlobalYearName = ""
		GlobalTurnName = ""
		GlobalTicks = 0
	End Sub
	
	'UPGRADE_NOTE: Tome was upgraded to Tome_Renamed. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="A9E4979A-37FA-4718-9994-97DD76ED70A7"'
	'UPGRADE_NOTE: Map was upgraded to Map_Renamed. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="A9E4979A-37FA-4718-9994-97DD76ED70A7"'
	Public Sub TomeStart(ByRef TomeX As Tome)
		Dim Map_Renamed As Object
		Dim Tome_Renamed As Object
		' This presumes that TomeX is already in the /world directory
		Dim c As Short
		Dim AreaX As Area
		Dim MapX As Map
		Dim CreatureX As Creature
		Dim FileName As String
		Dim NoFail, Found As Short
		Dim Creatures As Collection
		' [Titi 2.4.9] if tome is a home tome, load version in progress, unless a saved game is selected
		If TomeAction = bdTomeNew Or TomeAction = bdTomeGather Then
			FileName = gAppPath & "\Roster\" & WorldNow.Name & "\Home\" & TomeX.FileName
			Creatures = New Collection
			For	Each CreatureX In TomeX.Creatures
				Creatures.Add(CreatureX)
				If CreatureX.Home = WorldNow.Name & "/" & VB.Right(TomeX.FullPath, Len(TomeX.FullPath) - InStrRev(TomeX.FullPath, "\")) & "/" & TomeX.FileName Then
					' found a character whose home this tome is
					If oFileSys.CheckExists(FileName, clsInOut.IOActionType.File) Then
						' and this tome has already been lived in
						Found = True
					End If
				End If
			Next CreatureX
			If Found Then
				TomeRestore(FileName, True)
				For	Each CreatureX In TomeX.Creatures
					TomeX.RemoveCreature("X" & CreatureX.Index)
				Next CreatureX
				Do While Creatures.Count() > 0
					CreatureX = Creatures.Item(Creatures.Count())
					TomeX.AddCreature.Copy(CreatureX)
					Creatures.Remove((Creatures.Count()))
				Loop 
			End If
		End If
		Frozen = True
		' Create Tome Object
		Tome = TomeX
		' [Titi 2.4.7] Tome.LoadPath not set?
		'UPGRADE_WARNING: Couldn't resolve default property of object Tome.LoadPath. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
		'UPGRADE_WARNING: Couldn't resolve default property of object Tome.FullPath. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
		Tome.LoadPath = Tome.FullPath
		CreatureWithTurn = New Creature
		' Check the obvious errors
		If TomeX.Creatures.Count() < 1 Then
			DialogDM("Nobody is on this adventure.")
			GameEnd(False)
			Exit Sub
		End If
		If TomeX.Areas.Count() < 1 Then
			DialogDM("There are no Maps in this Tome.")
			GameEnd(False)
			Exit Sub
		End If
		If TomeX.AreaIndex < 1 Or TomeX.MapIndex < 1 Then
			DialogDM("There is no connecting EntryPoint in this Tome.")
			GameEnd(False)
			Exit Sub
		End If
		' Check for live Party
		Found = False
		For	Each CreatureX In TomeX.Creatures
			If CreatureX.HPNow > 0 Then
				Found = True
				Exit For
			End If
		Next CreatureX
		If Not Found Then
			DialogDM("Everybody on this adventure is dead.")
			GameEnd(False)
			Exit Sub
		End If
		' [Titi 2.4.9] moved this block out of the following For/Next CreatureX loop
		If oGameMusic.Status = IMCI.VIDEOSTATE.vsPLAYING Then
			Call oGameMusic.mciClose()
		End If
		' Display Party
		LoadPartyPics()
		MenuDrawParty()
		' Set CreatureWithTurn to first party member
		SetTurn(0)
		' Clear Global Variables
		TomeStartClearGlobals()
		' Load Current Plot
		If TomeX.OnAdventure = True Then
			TomeStartArea((TomeX.AreaIndex), (TomeX.MapIndex), 0, (TomeX.MapX), (TomeX.MapY))
		Else
			TomeStartArea((TomeX.AreaIndex), (TomeX.MapIndex), (TomeX.EntryIndex), (TomeX.MapX), (TomeX.MapY))
		End If
		' [borf] 2.4.7 moved this here so that users can assign a song to their tome
		' in the Enter_Tome trigger without it being jacked by this call below.
		Call PlayMusicRnd(modIOFunc.RNDMUSICSTYLE.Adventure, Me)
		' Save all the Creatures in this Tome (that are not locked NPCs)
		For	Each CreatureX In TomeX.Creatures
			If CreatureX.RequiredInTome = False And CreatureX.OnAdventure = False Then
				CreatureX.OnAdventure = True
				'UPGRADE_WARNING: Dir has a new behavior. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="9B7D5ADD-D8FE-4819-A36C-6DEDAF088CC7"'
				FileName = Dir(gAppPath & "\roster\" & WorldNow.Name & "\" & CreatureX.Name & ".rsc")
				If FileName <> "" Then
					FileName = gAppPath & "\roster\" & WorldNow.Name & "\" & CreatureX.Name & ".rsc"
					' RT 76 error! was Filename instead of CreatureX.Name & ".rsc"
					Kill(FileName)
				Else
					' RT 76 error! was Filename instead of CreatureX.Name & ".rsc"
					FileName = gAppPath & "\roster\" & WorldNow.Name & "\" & CreatureX.Name & ".rsc" ' error! was Filename
				End If
				c = FreeFile
				FileOpen(c, FileName, OpenMode.Binary)
				CreatureX.SaveToFile(c)
				FileClose(c)
			End If
			' [Titi 2.4.9] check if current tome is the home of the active character
			' note the "/" instead of "\" - this is due to the use of "/" or "/n" as carriage return
			'UPGRADE_WARNING: Couldn't resolve default property of object Tome.FileName. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
			'UPGRADE_WARNING: Couldn't resolve default property of object Tome.LoadPath. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
			If CreatureX.Home = WorldNow.Name & "/" & VB.Right(Tome.LoadPath, Len(Tome.LoadPath) - InStrRev(Tome.LoadPath, "\")) & "/" & Tome.FileName Then
				' The character has elected this tome as his living place.
				' Fire 'Home' triggers to put things back as they were when he left
				NoFail = FireTriggers(CreatureX, bdOnEnterHome)
				If NoFail = False Then
					GameEnd(False)
					Exit Sub
				End If
			End If
		Next CreatureX
		' Draw Borders (turn off drawing runes for a moment) [Titi 2.4.9] moved here to have the borders free of runes when tomes titles/splashscreens show
		'UPGRADE_WARNING: Couldn't resolve default property of object Map.IsNoRunes. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
		Found = Map.IsNoRunes '[Titi 2.4.9]
		'UPGRADE_WARNING: Couldn't resolve default property of object Map.IsNoRunes. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
		Map.IsNoRunes = True
		BorderDrawAll(True)
		' Fire Post-EnterTome Triggers
		If TomeX.OnAdventure = False Then
			ItemNow = CreatureWithTurn.ItemInHand
			NoFail = FireTriggers(Tome, bdPostEnterTome)
			If NoFail = False Then
				GameEnd(False)
				Exit Sub
			End If
			' Save Tome to \Worlds folder
			c = FreeFile
			FileOpen(c, gAppPath & "\World\" & WorldNow.Name & "\" & TomeX.FileName, OpenMode.Binary)
			TomeX.OnAdventure = True
			TomeX.SaveToFile(c)
			FileClose(c)
		End If
		' fire On-EnterHome triggers (even if tome is on adventure!)
		NoFail = FireTriggers(Tome, bdOnEnterHome)
		If NoFail = False Then
			GameEnd(False)
			Exit Sub
		End If
		' Run a TurnCycle
		TurnCycle()
		' Stop Party moving and Play Map Music
		tmrMoveParty.Enabled = False
		'    Call PlayMusicRnd(Adventure, Me)
		Frozen = False
		' Draw Borders (turn off drawing runes for a moment)
		'[Titi 2.4.9] set the runes back to the correct status
		'UPGRADE_WARNING: Couldn't resolve default property of object Map.IsNoRunes. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
		Map.IsNoRunes = Found
		BorderDrawAll(True)
		picMap.Focus()
	End Sub
	
	'UPGRADE_NOTE: Tome was upgraded to Tome_Renamed. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="A9E4979A-37FA-4718-9994-97DD76ED70A7"'
	'UPGRADE_NOTE: Area was upgraded to Area_Renamed. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="A9E4979A-37FA-4718-9994-97DD76ED70A7"'
	'UPGRADE_NOTE: Map was upgraded to Map_Renamed. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="A9E4979A-37FA-4718-9994-97DD76ED70A7"'
	Public Sub TomeStartArea(ByRef AreaIndex As Short, ByRef MapIndex As Short, ByRef EntryIndex As Short, ByRef AtX As Short, ByRef AtY As Short, Optional ByRef MapName As String = "")
		Dim Map_Renamed As Object
		Dim Area_Renamed As Object
		Dim Tome_Renamed As Object
		Dim i, c, Found As Short
		Dim CreatureX As Creature
		Dim FileName As String
		Dim NoFail As Short
		Dim AreaX As Area
		Dim MapX As Map
		Dim EntryX As EntryPoint
		' If going to AreaIndex = 0 then Exit to Menu
		If AreaIndex = 0 Then
			' Fire Pre-ExitTome Triggers
			ItemNow = CreatureWithTurn.ItemInHand
			NoFail = FireTriggers(EncounterNow, bdPreExitTome) And FireTriggers(Tome, bdPreExitTome)
			If NoFail Then
				' Save out the Party
				For	Each CreatureX In Tome.Creatures
					' [Titi 2.4.9] store 'home' info
					'UPGRADE_WARNING: Couldn't resolve default property of object Tome.FileName. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
					'UPGRADE_WARNING: Couldn't resolve default property of object Tome.LoadPath. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
					If CreatureX.Home = WorldNow.Name & "/" & VB.Right(Tome.LoadPath, Len(Tome.LoadPath) - InStrRev(Tome.LoadPath, "\")) & "/" & Tome.FileName Then
						c = oFileSys.CheckExists(gAppPath & "\Roster\" & WorldNow.Name & "\Home", clsInOut.IOActionType.Folder, True)
						TomeSavePathName = gAppPath & "\Roster\" & WorldNow.Name & "\Home"
						'UPGRADE_WARNING: Couldn't resolve default property of object Tome.FileName. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
						FileName = TomeSavePathName & "\" & Tome.FileName
						TomeSave(False)
					End If
					If CreatureX.RequiredInTome = False Then
						CreatureX.OnAdventure = False
						FileName = oFileSys.Move(gAppPath & "\roster\" & WorldNow.Name & "\" & CreatureX.Name & ".rsc", clsInOut.IOActionType.File,  , True)
						On Error Resume Next
						c = FreeFile
						FileOpen(c, gAppPath & "\roster\" & WorldNow.Name & "\" & CreatureX.Name & ".rsc", OpenMode.Binary)
						CreatureX.SaveToFile(c)
						FileClose(c)
						If Err.Number <> 0 Then ' error occured lets put rollback
							Call oFileSys.Move(FileName, clsInOut.IOActionType.File, gAppPath & "\roster\" & WorldNow.Name & "\" & CreatureX.Name & ".rsc")
							DialogDM("Could not save character " & CreatureX.Name & ".")
							Exit Sub
						Else
							Call oFileSys.Delete(FileName, clsInOut.IOActionType.File, True)
						End If
						' Clear out the Tome [Titi 2.4.9] no need for another loop!
						'UPGRADE_WARNING: Couldn't resolve default property of object Tome.RemoveCreature. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
						Tome.RemoveCreature("X" & CreatureX.Index)
					End If
				Next CreatureX
				'            ' Clear out the Tome
				'            For Each CreatureX In Tome.Creatures
				'                If CreatureX.RequiredInTome = False Then
				'                    Tome.RemoveCreature "X" & CreatureX.Index
				'                End If
				'            Next CreatureX
				' Save Tome to Worlds folder
				'UPGRADE_WARNING: Couldn't resolve default property of object Tome.FileName. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
				FileName = oFileSys.Move(gAppPath & "\World\" & WorldNow.Name & "\" & Tome.FileName, clsInOut.IOActionType.File,  , True)
				On Error Resume Next
				c = FreeFile
				'UPGRADE_WARNING: Couldn't resolve default property of object Tome.FileName. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
				FileOpen(c, gAppPath & "\World\" & WorldNow.Name & "\" & Tome.FileName, OpenMode.Binary)
				'UPGRADE_WARNING: Couldn't resolve default property of object Tome.OnAdventure. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
				Tome.OnAdventure = False
				'UPGRADE_WARNING: Couldn't resolve default property of object Tome.SaveToFile. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
				Tome.SaveToFile(c)
				FileClose(c)
				If Err.Number <> 0 Then
					'                Call oFileSys.Move(Filename, File, gAppPath & "\roster\" & WorldNow.Name & "\" & CreatureX.Name & ".rsc")
					'               [Titi] shouldn't we save the tome here?
					'               Not a problem if one area, but will prevent the "Reset the Tome?" box from showing if more than one area
					'               since the OnAdventure property is not set into the correct "tome" file
					'UPGRADE_WARNING: Couldn't resolve default property of object Tome.FileName. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
					Call oFileSys.Move(FileName, clsInOut.IOActionType.File, gAppPath & "\World\" & WorldNow.Name & "\" & Tome.FileName & ".tom")
					DialogDM("Could not save tome.")
					Exit Sub
				Else
					Call oFileSys.Delete(FileName, clsInOut.IOActionType.File, True)
				End If
				'[borfaux] Still need to remove all the saved tomes so characters don't get fubar
				' [Titi 2.4.9] Delete any copies already in the \world directory
				'UPGRADE_WARNING: Couldn't resolve default property of object Tome.FileName. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
				'UPGRADE_WARNING: Dir has a new behavior. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="9B7D5ADD-D8FE-4819-A36C-6DEDAF088CC7"'
				FileName = Dir(gAppPath & "\world\" & WorldNow.Name & "\" & Mid(Tome.FileName, 1, Len(Tome.FileName) - 4) & "*.*")
				If FileName <> "" Then
					'UPGRADE_WARNING: Couldn't resolve default property of object Tome.FileName. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
					Kill(gAppPath & "\world\" & WorldNow.Name & "\" & Mid(Tome.FileName, 1, Len(Tome.FileName) - 4) & "*.*")
				End If
				' Back to the Main Menu
				GameEnd(False)
			End If
			Exit Sub
		End If
		' Find Plot to load
		Found = False
		'UPGRADE_WARNING: Couldn't resolve default property of object Tome.Areas. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
		For	Each AreaX In Tome.Areas
			If AreaX.Index = AreaIndex Then
				FileName = AreaX.FileName
				Found = True
				Exit For
			End If
		Next AreaX
		If Not Found Then
			DialogDM("Could not locate Area #" & AreaIndex & " for this Tome.")
			GameEnd(False)
			Exit Sub
		End If
		' If new Area, then save current and load new
		'UPGRADE_WARNING: Screen property Screen.MousePointer has a new behavior. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6BA9B8D2-2A32-4B6E-8D36-44949974A5B4"'
		System.Windows.Forms.Cursor.Current = System.Windows.Forms.Cursors.WaitCursor
		'UPGRADE_WARNING: Couldn't resolve default property of object Area.Index. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
		If AreaX.Index <> Area.Index Then
			If Area.Index <> 0 Then
				Call TomeSaveArea(gAppPath & "\World\" & WorldNow.Name)
				'UPGRADE_WARNING: Screen property Screen.MousePointer has a new behavior. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6BA9B8D2-2A32-4B6E-8D36-44949974A5B4"'
				System.Windows.Forms.Cursor.Current = System.Windows.Forms.Cursors.WaitCursor
			End If
			c = FreeFile
			'Open gAppPath & "\world\" & FileName For Binary As c
			FileOpen(c, gAppPath & "\world\" & WorldNow.Name & "\" & FileName, OpenMode.Binary)
			Area = New Area
			'UPGRADE_WARNING: Couldn't resolve default property of object Area.Plot. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
			Area.Plot.ReadFromFile(c)
			'UPGRADE_WARNING: Couldn't resolve default property of object Area.Index. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
			Area.Index = AreaIndex
			'UPGRADE_WARNING: Couldn't resolve default property of object Area.FileName. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
			Area.FileName = FileName
			FileClose(c)
		End If
		' Check if missing Maps
		'UPGRADE_WARNING: Couldn't resolve default property of object Area.Plot. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
		If Area.Plot.Maps.Count < 1 Then
			DialogDM("There are no Maps in the starting Area.")
			GameEnd(False)
			Exit Sub
		End If
		' Set current Map (will autodraw when window sizes)
		Found = False
		'UPGRADE_WARNING: Couldn't resolve default property of object Area.Plot. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
		For	Each MapX In Area.Plot.Maps
			If MapX.Index = MapIndex Or MapX.Name = MapName Then
				Map = MapX
				Found = True
				Exit For
			End If
		Next MapX
		If Not Found Then
			'UPGRADE_NOTE: IsMissing() was changed to IsNothing(). Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="8AE1CB93-37AB-439A-A4FF-BE3B6760BB23"'
			If Not IsNothing(MapName) Then
				DialogDM("Could not locate Map '" & MapName & "' for Area.")
			Else
				DialogDM("Could not locate Map #" & MapIndex & " for Area.")
			End If
			'UPGRADE_WARNING: Couldn't resolve default property of object Tome.MapY. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
			'UPGRADE_WARNING: Couldn't resolve default property of object Tome.MapX. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
			MapCenter(Tome.MapX, Tome.MapY)
			DrawMapAll()
			'UPGRADE_WARNING: Screen property Screen.MousePointer has a new behavior. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6BA9B8D2-2A32-4B6E-8D36-44949974A5B4"'
			System.Windows.Forms.Cursor.Current = System.Windows.Forms.Cursors.Default
			Exit Sub
		End If
		' ReGen Map if set
		'UPGRADE_WARNING: Couldn't resolve default property of object Map.GenerateUponEntry. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
		If Map.GenerateUponEntry = True Then
			modDungeonMaker.MakeMap(Map, True)
			'UPGRADE_WARNING: Couldn't resolve default property of object Map.GenerateUponEntry. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
			Map.GenerateUponEntry = False
		End If
		' Locate EntryPoint and set (if first time in)
		Found = False
		For	Each EntryX In MapX.EntryPoints
			If EntryX.Index = EntryIndex Then
				'UPGRADE_WARNING: Couldn't resolve default property of object Tome.MapX. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
				Tome.MapX = EntryX.MapX
				'UPGRADE_WARNING: Couldn't resolve default property of object Tome.MapY. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
				Tome.MapY = EntryX.MapY
				Found = True
				Exit For
			End If
		Next EntryX
		If Found = False Then
			'UPGRADE_WARNING: Couldn't resolve default property of object Tome.MapX. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
			Tome.MapX = AtX
			'UPGRADE_WARNING: Couldn't resolve default property of object Tome.MapY. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
			Tome.MapY = AtY
		End If
		'UPGRADE_WARNING: Couldn't resolve default property of object Tome.MoveToX. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
		'UPGRADE_WARNING: Couldn't resolve default property of object Tome.MapX. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
		Tome.MoveToX = Tome.MapX
		'UPGRADE_WARNING: Couldn't resolve default property of object Tome.MoveToY. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
		'UPGRADE_WARNING: Couldn't resolve default property of object Tome.MapY. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
		Tome.MoveToY = Tome.MapY
		'UPGRADE_WARNING: Couldn't resolve default property of object Tome.AreaIndex. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
		Tome.AreaIndex = AreaIndex
		'UPGRADE_WARNING: Couldn't resolve default property of object Tome.MapIndex. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
		Tome.MapIndex = MapIndex
		'UPGRADE_WARNING: Couldn't resolve default property of object Tome.EntryIndex. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
		Tome.EntryIndex = EntryIndex
		' Set all CreatureWithTurns in Party to Map's starting point
		c = -1
		'UPGRADE_WARNING: Couldn't resolve default property of object Tome.Creatures. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
		For	Each CreatureX In Tome.Creatures
			c = LoopNumber(0, 4, c, 1)
			CreatureX.TileSpot = c
			'UPGRADE_WARNING: Couldn't resolve default property of object Tome.MapX. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
			CreatureX.MapX = Tome.MapX
			'UPGRADE_WARNING: Couldn't resolve default property of object Tome.MapY. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
			CreatureX.MapY = Tome.MapY
		Next CreatureX
		' Draw Map
		'UPGRADE_WARNING: Couldn't resolve default property of object Tome.MapY. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
		'UPGRADE_WARNING: Couldn't resolve default property of object Tome.MapX. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
		MapCenter(Tome.MapX, Tome.MapY)
		DrawMapAll()
		MoveParty()
		BorderDrawAll(True)
		' Set up current Encounter
		'UPGRADE_WARNING: Couldn't resolve default property of object Tome.MapY. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
		'UPGRADE_WARNING: Couldn't resolve default property of object Tome.MapX. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
		'UPGRADE_WARNING: Couldn't resolve default property of object Map.EncPointer. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
		If Map.EncPointer(Tome.MapX, Tome.MapY) > 0 Then
			'UPGRADE_WARNING: Couldn't resolve default property of object Tome.MapY. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
			'UPGRADE_WARNING: Couldn't resolve default property of object Tome.MapX. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
			'UPGRADE_WARNING: Couldn't resolve default property of object Map.EncPointer. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
			'UPGRADE_WARNING: Couldn't resolve default property of object Map.Encounters. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
			EncounterNow = Map.Encounters("E" & Map.EncPointer(Tome.MapX, Tome.MapY))
		Else
			EncounterNow = New Encounter
		End If
		'UPGRADE_WARNING: Screen property Screen.MousePointer has a new behavior. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6BA9B8D2-2A32-4B6E-8D36-44949974A5B4"'
		System.Windows.Forms.Cursor.Current = System.Windows.Forms.Cursors.Default
		' Enter Encounter
		If EncounterNow.IsActive = True Then
			EncounterEnter()
		End If
		Me.Refresh()
	End Sub
	
	'UPGRADE_NOTE: Tome was upgraded to Tome_Renamed. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="A9E4979A-37FA-4718-9994-97DD76ED70A7"'
	'UPGRADE_NOTE: Map was upgraded to Map_Renamed. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="A9E4979A-37FA-4718-9994-97DD76ED70A7"'
	Public Sub LoadTileSetMicro()
		Dim Map_Renamed As Object
		Dim Tome_Renamed As Object
		' Load and sets Tiles for Map
		Dim FileName As String
		'UPGRADE_NOTE: Height was upgraded to Height_Renamed. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="A9E4979A-37FA-4718-9994-97DD76ED70A7"'
		'UPGRADE_NOTE: Width was upgraded to Width_Renamed. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="A9E4979A-37FA-4718-9994-97DD76ED70A7"'
		Dim Width_Renamed, Height_Renamed, OldWidth As Short
		Dim hMemTiles As Integer
		'UPGRADE_WARNING: Arrays in structure bmTiles may need to be initialized before they can be used. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="814DF224-76BD-4BB4-BFFB-EA359CB9FC48"'
		Dim bmTiles As BITMAPINFO
		'UPGRADE_WARNING: Arrays in structure bmBlack may need to be initialized before they can be used. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="814DF224-76BD-4BB4-BFFB-EA359CB9FC48"'
		Dim bmBlack As BITMAPINFO
		'UPGRADE_WARNING: Arrays in structure bmMask may need to be initialized before they can be used. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="814DF224-76BD-4BB4-BFFB-EA359CB9FC48"'
		Dim bmMask As BITMAPINFO
		Dim rc, lpMem, TransparentRGB As Integer
		'UPGRADE_WARNING: Screen property Screen.MousePointer has a new behavior. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6BA9B8D2-2A32-4B6E-8D36-44949974A5B4"'
		System.Windows.Forms.Cursor.Current = System.Windows.Forms.Cursors.WaitCursor ' Hourglass
		' Load TileSet bitmap
		'UPGRADE_WARNING: Couldn't resolve default property of object Map.PictureFile. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
		'UPGRADE_WARNING: Couldn't resolve default property of object Tome.FullPath. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
		'UPGRADE_WARNING: Dir has a new behavior. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="9B7D5ADD-D8FE-4819-A36C-6DEDAF088CC7"'
		FileName = Dir(Tome.FullPath & "\" & Map.PictureFile)
		If FileName = "" Then
			'UPGRADE_WARNING: Couldn't resolve default property of object Map.PictureFile. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
			'UPGRADE_WARNING: Couldn't resolve default property of object Tome.FullPath. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
			'UPGRADE_WARNING: Dir has a new behavior. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="9B7D5ADD-D8FE-4819-A36C-6DEDAF088CC7"'
			FileName = Dir(Tome.FullPath & "\tiles\" & Map.PictureFile)
			If FileName = "" Then
				'FileName = Dir(gAppPath & "\data\graphics\tiles\" & Map.PictureFile)
				'        FileName = gAppPath & "\data\graphics\tiles\" & Map.PictureFile
				'UPGRADE_WARNING: Couldn't resolve default property of object Map.PictureFile. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
				FileName = gDataPath & "\graphics\tiles\" & Map.PictureFile
			Else
				'UPGRADE_WARNING: Couldn't resolve default property of object Map.PictureFile. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
				'UPGRADE_WARNING: Couldn't resolve default property of object Tome.FullPath. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
				FileName = Tome.FullPath & "\tiles\" & Map.PictureFile
			End If
		Else
			'UPGRADE_WARNING: Couldn't resolve default property of object Map.PictureFile. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
			'UPGRADE_WARNING: Couldn't resolve default property of object Tome.FullPath. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
			FileName = Tome.FullPath & "\" & Map.PictureFile
		End If
		ReadBitmapFile(FileName, bmTiles, hMemTiles, TransparentRGB)
		' Make a copy of the current palette for the picture
		'UPGRADE_WARNING: Couldn't resolve default property of object bmBlack. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
		bmBlack = bmTiles
		' Build palettes: first, change from blue to black
		ChangeColor(bmBlack, TransparentRGB, 0, 0, 0)
		' Convert to Mask (pure Blue is the mask color)
		MakeMask(bmTiles, bmMask, TransparentRGB)
		' Paint bitmap to picture box using converted palette
		lpMem = GlobalLock(hMemTiles)
		Width_Renamed = bmTiles.bmiHeader.biWidth
		Height_Renamed = bmTiles.bmiHeader.biHeight
		picTSmall.Width = Width_Renamed * 2 * bdMicroSize
		picTSmall.Height = Height_Renamed * 2 * bdMicroSize
		'UPGRADE_ISSUE: PictureBox property picTSmall.hdc was not upgraded. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="CC4C7EC0-C903-48FC-ACCC-81861D12DA4A"'
		rc = SetStretchBltMode(picTSmall.hdc, 3)
		' Upper left quad: normal picture
		'UPGRADE_ISSUE: PictureBox property picTSmall.hdc was not upgraded. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="CC4C7EC0-C903-48FC-ACCC-81861D12DA4A"'
		rc = StretchDIBits(picTSmall.hdc, 0, 0, Width_Renamed * bdMicroSize, Height_Renamed * bdMicroSize, 0, 0, Width_Renamed, Height_Renamed, lpMem, bmTiles, DIB_RGB_COLORS, SRCCOPY)
		' Upper right quad: flip normal picture
		'UPGRADE_ISSUE: PictureBox property picTSmall.hdc was not upgraded. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="CC4C7EC0-C903-48FC-ACCC-81861D12DA4A"'
		rc = StretchBlt(picTSmall.hdc, Width_Renamed * bdMicroSize, 0, Width_Renamed * bdMicroSize, Height_Renamed * bdMicroSize, picTSmall.hdc, Width_Renamed * bdMicroSize - 1, 0, -Width_Renamed * bdMicroSize, Height_Renamed * bdMicroSize, SRCCOPY)
		' Lower left quad: mask
		'UPGRADE_ISSUE: PictureBox property picTSmall.hdc was not upgraded. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="CC4C7EC0-C903-48FC-ACCC-81861D12DA4A"'
		rc = StretchDIBits(picTSmall.hdc, 0, Height_Renamed * bdMicroSize, Width_Renamed * bdMicroSize, Height_Renamed * bdMicroSize, 0, 0, Width_Renamed, Height_Renamed, lpMem, bmMask, DIB_RGB_COLORS, SRCCOPY)
		' Lower right quadrant: flip mask
		'UPGRADE_ISSUE: PictureBox property picTSmall.hdc was not upgraded. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="CC4C7EC0-C903-48FC-ACCC-81861D12DA4A"'
		rc = StretchBlt(picTSmall.hdc, Width_Renamed * bdMicroSize, Height_Renamed * bdMicroSize, Width_Renamed * bdMicroSize, Height_Renamed * bdMicroSize, picTSmall.hdc, Width_Renamed * bdMicroSize - 1, Height_Renamed * bdMicroSize, -Width_Renamed * bdMicroSize, Height_Renamed * bdMicroSize, SRCCOPY)
		' Set Current TileSet Name and release memory
		rc = GlobalUnlock(hMemTiles)
		rc = GlobalFree(hMemTiles)
		picTSmall.Refresh()
		' Load DarkTile bitmap
		'    ReadBitmapFile gAppPath & "\data\stock\darkmicro.bmp", bmTiles, hMemTiles, TransparentRGB
		ReadBitmapFile(gDataPath & "\stock\darkmicro.bmp", bmTiles, hMemTiles, TransparentRGB)
		' Paint bitmap to picture box using converted palette
		lpMem = GlobalLock(hMemTiles)
		Width_Renamed = bmTiles.bmiHeader.biWidth
		Height_Renamed = bmTiles.bmiHeader.biHeight
		picBlackHeightSmall = (Height_Renamed * bdMicroSize) / 2
		picBlack.Width = bdBlackWidth + Width_Renamed * bdMicroSize
		' Writes out normal tiles (black border) to upper left.
		'UPGRADE_ISSUE: PictureBox property picBlack.hdc was not upgraded. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="CC4C7EC0-C903-48FC-ACCC-81861D12DA4A"'
		rc = SetStretchBltMode(picBlack.hdc, 3)
		'UPGRADE_ISSUE: PictureBox property picBlack.hdc was not upgraded. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="CC4C7EC0-C903-48FC-ACCC-81861D12DA4A"'
		rc = StretchDIBits(picBlack.hdc, bdBlackWidth, 0, Width_Renamed * bdMicroSize, Height_Renamed * bdMicroSize, 0, 0, Width_Renamed, Height_Renamed, lpMem, bmTiles, DIB_RGB_COLORS, SRCCOPY)
		' Release memory
		rc = GlobalUnlock(hMemTiles)
		rc = GlobalFree(hMemTiles)
		picBlack.Refresh()
		'UPGRADE_WARNING: Screen property Screen.MousePointer has a new behavior. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6BA9B8D2-2A32-4B6E-8D36-44949974A5B4"'
		System.Windows.Forms.Cursor.Current = System.Windows.Forms.Cursors.Default ' Default
	End Sub
	
	'UPGRADE_NOTE: Tome was upgraded to Tome_Renamed. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="A9E4979A-37FA-4718-9994-97DD76ED70A7"'
	'UPGRADE_NOTE: Map was upgraded to Map_Renamed. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="A9E4979A-37FA-4718-9994-97DD76ED70A7"'
	Public Sub LoadTileSet(ByRef AsDark As Short, ByRef SupressMsg As Short)
		Dim Map_Renamed As Object
		Dim Tome_Renamed As Object
		' Load and sets Tiles for Map
		Dim FileName As String
		'UPGRADE_NOTE: Width was upgraded to Width_Renamed. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="A9E4979A-37FA-4718-9994-97DD76ED70A7"'
		'UPGRADE_NOTE: Height was upgraded to Height_Renamed. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="A9E4979A-37FA-4718-9994-97DD76ED70A7"'
		Dim Height_Renamed, Width_Renamed As Short
		Dim hMemTiles As Integer
		'UPGRADE_WARNING: Arrays in structure bmTiles may need to be initialized before they can be used. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="814DF224-76BD-4BB4-BFFB-EA359CB9FC48"'
		Dim bmTiles As BITMAPINFO
		'UPGRADE_WARNING: Arrays in structure bmDark may need to be initialized before they can be used. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="814DF224-76BD-4BB4-BFFB-EA359CB9FC48"'
		'UPGRADE_WARNING: Arrays in structure bmBlack may need to be initialized before they can be used. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="814DF224-76BD-4BB4-BFFB-EA359CB9FC48"'
		Dim bmBlack, bmDark As BITMAPINFO
		'UPGRADE_WARNING: Arrays in structure bmMask may need to be initialized before they can be used. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="814DF224-76BD-4BB4-BFFB-EA359CB9FC48"'
		Dim bmMask As BITMAPINFO
		Dim rc, lpMem, TransparentRGB As Integer
		'UPGRADE_WARNING: Screen property Screen.MousePointer has a new behavior. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6BA9B8D2-2A32-4B6E-8D36-44949974A5B4"'
		System.Windows.Forms.Cursor.Current = System.Windows.Forms.Cursors.WaitCursor ' Hourglass
		If Not SupressMsg Then
			MessageShow("Loading Map..", 0)
		End If
		' Load TileSet bitmap
		'UPGRADE_WARNING: Couldn't resolve default property of object Map.PictureFile. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
		'UPGRADE_WARNING: Couldn't resolve default property of object Tome.FullPath. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
		'UPGRADE_WARNING: Dir has a new behavior. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="9B7D5ADD-D8FE-4819-A36C-6DEDAF088CC7"'
		FileName = Dir(Tome.FullPath & "\" & Map.PictureFile)
		If FileName = "" Then
			'UPGRADE_WARNING: Couldn't resolve default property of object Map.PictureFile. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
			'UPGRADE_WARNING: Couldn't resolve default property of object Tome.FullPath. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
			'UPGRADE_WARNING: Dir has a new behavior. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="9B7D5ADD-D8FE-4819-A36C-6DEDAF088CC7"'
			FileName = Dir(Tome.FullPath & "\tiles\" & Map.PictureFile)
			If FileName = "" Then
				'    FileName = gAppPath & "\data\graphics\tiles\" & Map.PictureFile
				'UPGRADE_WARNING: Couldn't resolve default property of object Map.PictureFile. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
				FileName = gDataPath & "\graphics\tiles\" & Map.PictureFile
			Else
				'UPGRADE_WARNING: Couldn't resolve default property of object Map.PictureFile. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
				'UPGRADE_WARNING: Couldn't resolve default property of object Tome.FullPath. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
				FileName = Tome.FullPath & "\tiles\" & Map.PictureFile
			End If
		Else
			'UPGRADE_WARNING: Couldn't resolve default property of object Map.PictureFile. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
			'UPGRADE_WARNING: Couldn't resolve default property of object Tome.FullPath. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
			FileName = Tome.FullPath & "\" & Map.PictureFile
		End If
		ReadBitmapFile(FileName, bmTiles, hMemTiles, TransparentRGB)
		' Make a copy of the current palette for the picture
		If Not SupressMsg Then
			MessageShow("Loading Map....", 0)
		End If
		'UPGRADE_WARNING: Couldn't resolve default property of object bmBlack. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
		bmBlack = bmTiles
		' Build palettes: first, change from blue to black
		ChangeColor(bmBlack, TransparentRGB, 0, 0, 0)
		' Convert to Mask (pure Blue is the mask color)
		MakeMask(bmTiles, bmMask, TransparentRGB)
		' Darken a bit
		'UPGRADE_WARNING: Couldn't resolve default property of object Map.IsOutside. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
		If Map.IsOutside = True Then
			MakeDark(bmBlack, bmDark, AsDark)
		Else
			MakeDark(bmBlack, bmDark, 0)
		End If
		' Paint bitmap to picture box using converted palette
		lpMem = GlobalLock(hMemTiles)
		Width_Renamed = bmTiles.bmiHeader.biWidth
		Height_Renamed = bmTiles.bmiHeader.biHeight
		'UPGRADE_ISSUE: PictureBox method picTPic.Cls was not upgraded. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="CC4C7EC0-C903-48FC-ACCC-81861D12DA4A"'
		picTPic.Cls()
		picTPic.Width = Width_Renamed * 2
		picTPic.Height = Height_Renamed * 2
		If Not SupressMsg Then
			MessageShow("Loading Map.....", 0)
		End If
		' Upper left quad: normal picture
		'UPGRADE_ISSUE: PictureBox property picTPic.hdc was not upgraded. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="CC4C7EC0-C903-48FC-ACCC-81861D12DA4A"'
		rc = SetStretchBltMode(picTPic.hdc, 3)
		'UPGRADE_ISSUE: PictureBox property picTPic.hdc was not upgraded. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="CC4C7EC0-C903-48FC-ACCC-81861D12DA4A"'
		rc = StretchDIBits(picTPic.hdc, 0, 0, Width_Renamed, Height_Renamed, 0, 0, Width_Renamed, Height_Renamed, lpMem, bmDark, DIB_RGB_COLORS, SRCCOPY)
		If Not SupressMsg Then
			MessageShow("Loading Map......", 0)
		End If
		' Upper right quad: flip normal picture
		'UPGRADE_ISSUE: PictureBox property picTPic.hdc was not upgraded. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="CC4C7EC0-C903-48FC-ACCC-81861D12DA4A"'
		rc = StretchBlt(picTPic.hdc, Width_Renamed, 0, Width_Renamed, Height_Renamed, picTPic.hdc, Width_Renamed - 1, 0, -Width_Renamed, Height_Renamed, SRCCOPY)
		' Lower left quad: mask
		'UPGRADE_ISSUE: PictureBox property picTPic.hdc was not upgraded. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="CC4C7EC0-C903-48FC-ACCC-81861D12DA4A"'
		rc = StretchDIBits(picTPic.hdc, 0, Height_Renamed, Width_Renamed, Height_Renamed, 0, 0, Width_Renamed, Height_Renamed, lpMem, bmMask, DIB_RGB_COLORS, SRCCOPY)
		If Not SupressMsg Then
			MessageShow("Loading Map........", 0)
		End If
		' Lower right quadrant: flip mask
		'UPGRADE_ISSUE: PictureBox property picTPic.hdc was not upgraded. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="CC4C7EC0-C903-48FC-ACCC-81861D12DA4A"'
		rc = StretchBlt(picTPic.hdc, Width_Renamed, Height_Renamed, Width_Renamed, Height_Renamed, picTPic.hdc, Width_Renamed - 1, Height_Renamed, -Width_Renamed, Height_Renamed, SRCCOPY)
		picTPic.Refresh()
		' Release memory
		If Not SupressMsg Then
			MessageShow("Loading Map..........", 0)
		End If
		' Set Current TileSet Name
		'UPGRADE_WARNING: Couldn't resolve default property of object Map.PictureFile. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
		TileSetLoaded = Map.PictureFile
		rc = GlobalUnlock(hMemTiles)
		rc = GlobalFree(hMemTiles)
		If Not SupressMsg Then
			MessageClear()
		End If
		'UPGRADE_WARNING: Screen property Screen.MousePointer has a new behavior. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6BA9B8D2-2A32-4B6E-8D36-44949974A5B4"'
		System.Windows.Forms.Cursor.Current = System.Windows.Forms.Cursors.Default ' Default
	End Sub
	
	'UPGRADE_NOTE: Map was upgraded to Map_Renamed. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="A9E4979A-37FA-4718-9994-97DD76ED70A7"'
	'UPGRADE_NOTE: Tome was upgraded to Tome_Renamed. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="A9E4979A-37FA-4718-9994-97DD76ED70A7"'
	Private Sub PlotTile(ByRef TileToPlot As Short, ByRef X As Short, ByRef Y As Short, ByRef XWidth As Short, ByRef YWidth As Short, ByRef XSrcOff As Short, ByRef YSrcOff As Short, ByRef XFlip As Short)
		Dim Tome_Renamed As Object
		Dim Map_Renamed As Object
		Dim rc As Integer
		Dim MaxHeight, MaxWidth As Short
		Dim TileX, TileY As Short
		' [Titi 2.4.6] case of the missing tileset
		Dim save As String
		On Error GoTo ErrorHandler
		' Draw tile
		If TileToPlot < 0 Then
			Select Case TileToPlot
				Case bdTileBlack ' Pure Black
					'UPGRADE_ISSUE: PictureBox property picBlack.hdc was not upgraded. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="CC4C7EC0-C903-48FC-ACCC-81861D12DA4A"'
					'UPGRADE_ISSUE: PictureBox property picMap.hdc was not upgraded. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="CC4C7EC0-C903-48FC-ACCC-81861D12DA4A"'
					rc = BitBlt(picMap.hdc, X, Y, XWidth, YWidth, picBlack.hdc, XSrcOff, 72 + YSrcOff, SRCAND)
					'UPGRADE_ISSUE: PictureBox property picBlack.hdc was not upgraded. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="CC4C7EC0-C903-48FC-ACCC-81861D12DA4A"'
					'UPGRADE_ISSUE: PictureBox property picMap.hdc was not upgraded. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="CC4C7EC0-C903-48FC-ACCC-81861D12DA4A"'
					rc = BitBlt(picMap.hdc, X, Y, XWidth, YWidth, picBlack.hdc, XSrcOff, YSrcOff, SRCPAINT)
				Case bdTileGrey ' Light Dim
					'UPGRADE_ISSUE: PictureBox property picBlack.hdc was not upgraded. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="CC4C7EC0-C903-48FC-ACCC-81861D12DA4A"'
					'UPGRADE_ISSUE: PictureBox property picMap.hdc was not upgraded. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="CC4C7EC0-C903-48FC-ACCC-81861D12DA4A"'
					rc = BitBlt(picMap.hdc, X, Y, XWidth, YWidth, picBlack.hdc, 96 + XSrcOff, YSrcOff, SRCAND)
					'UPGRADE_ISSUE: PictureBox property picBlack.hdc was not upgraded. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="CC4C7EC0-C903-48FC-ACCC-81861D12DA4A"'
					'UPGRADE_ISSUE: PictureBox property picMap.hdc was not upgraded. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="CC4C7EC0-C903-48FC-ACCC-81861D12DA4A"'
					rc = BitBlt(picMap.hdc, X, Y, XWidth, YWidth, picBlack.hdc, XSrcOff, YSrcOff, SRCPAINT)
				Case bdTileDark ' Dark Dim
					'UPGRADE_ISSUE: PictureBox property picBlack.hdc was not upgraded. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="CC4C7EC0-C903-48FC-ACCC-81861D12DA4A"'
					'UPGRADE_ISSUE: PictureBox property picMap.hdc was not upgraded. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="CC4C7EC0-C903-48FC-ACCC-81861D12DA4A"'
					rc = BitBlt(picMap.hdc, X, Y, XWidth, YWidth, picBlack.hdc, 96 + XSrcOff, YSrcOff, SRCAND)
					'UPGRADE_ISSUE: PictureBox property picBlack.hdc was not upgraded. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="CC4C7EC0-C903-48FC-ACCC-81861D12DA4A"'
					'UPGRADE_ISSUE: PictureBox property picMap.hdc was not upgraded. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="CC4C7EC0-C903-48FC-ACCC-81861D12DA4A"'
					rc = BitBlt(picMap.hdc, X, Y, XWidth, YWidth, picBlack.hdc, XSrcOff, YSrcOff, SRCPAINT)
			End Select
		Else
			' Set height, width and location of tile in bmp
			MaxHeight = (picTPic.Height / 144)
			MaxWidth = (picTPic.Width / 192)
			TileX = 96 * Int(TileToPlot / MaxHeight)
			TileY = 72 * (TileToPlot Mod MaxHeight)
			If XFlip = True Then
				TileX = picTPic.Width - 96 * (TileX / 96) - 96
			End If
			'UPGRADE_ISSUE: PictureBox property picTPic.hdc was not upgraded. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="CC4C7EC0-C903-48FC-ACCC-81861D12DA4A"'
			'UPGRADE_ISSUE: PictureBox property picMap.hdc was not upgraded. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="CC4C7EC0-C903-48FC-ACCC-81861D12DA4A"'
			rc = BitBlt(picMap.hdc, X, Y, XWidth, YWidth, picTPic.hdc, TileX + XSrcOff, TileY + YSrcOff + (picTPic.Height / 2), SRCAND)
			'UPGRADE_ISSUE: PictureBox property picTPic.hdc was not upgraded. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="CC4C7EC0-C903-48FC-ACCC-81861D12DA4A"'
			'UPGRADE_ISSUE: PictureBox property picMap.hdc was not upgraded. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="CC4C7EC0-C903-48FC-ACCC-81861D12DA4A"'
			rc = BitBlt(picMap.hdc, X, Y, XWidth, YWidth, picTPic.hdc, TileX + XSrcOff, TileY + YSrcOff, SRCPAINT)
		End If
		Exit Sub
ErrorHandler: 
		oErr.logError("PlotTile")
		If Err.Number = 11 Then
			'UPGRADE_WARNING: Couldn't resolve default property of object Map.PictureFile. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
			DialogDM("Sorry, you cannot enter this area without the tileset named " & Map.PictureFile)
			'UPGRADE_WARNING: Couldn't resolve default property of object Tome.MoveToX. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
			Tome.MoveToX = FromEntryPoint.MapX
			'UPGRADE_WARNING: Couldn't resolve default property of object Tome.MoveToY. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
			Tome.MoveToY = FromEntryPoint.MapY
			'UPGRADE_NOTE: Object FromEntryPoint may not be destroyed until it is garbage collected. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6E35BFF6-CD74-4B09-9689-3E1A43DF8969"'
			FromEntryPoint = Nothing
			MovePartyExit(True)
			'        save = MsgBox("Save the game now?", vbYesNo)
			'        If save = vbYes Then
			'            TomeSavesLoad
			'            CreateNameNew = "Missing " & Map.PictureFile & " (" & CreateNameNew & ")"
			'            TomeSave True, False
			'        End If
			'        End
		End If
	End Sub
	
	'UPGRADE_NOTE: Tome was upgraded to Tome_Renamed. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="A9E4979A-37FA-4718-9994-97DD76ED70A7"'
	Private Sub CombatFlee()
		Dim Tome_Renamed As Object
		' Check for chance to flee
		If picGrid.Visible = True Then
			CreatureWithTurn.Afraid = True
			CombatMove(MOVEDIRECTION.bdMoveFlee, MOVEDIRECTION.bdMoveFlee)
			'UPGRADE_WARNING: Couldn't resolve default property of object Tome.FleeToY. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
			'UPGRADE_WARNING: Couldn't resolve default property of object Tome.FleeToX. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
			MovePartySet(0, 0, Tome.FleeToX, Tome.FleeToY)
			CreatureWithTurn.ActionPoints = 0
			CombatNextTurn()
		ElseIf Int(Rnd() * 100) < EncounterNow.ChanceToFlee Then 
			'UPGRADE_WARNING: Couldn't resolve default property of object Tome.FleeToY. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
			'UPGRADE_WARNING: Couldn't resolve default property of object Tome.FleeToX. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
			MovePartySet(0, 0, Tome.FleeToX, Tome.FleeToY)
		Else
			DialogDM("You try to flee, but fail.")
		End If
	End Sub
	
	'UPGRADE_NOTE: Tome was upgraded to Tome_Renamed. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="A9E4979A-37FA-4718-9994-97DD76ED70A7"'
	'UPGRADE_NOTE: Map was upgraded to Map_Renamed. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="A9E4979A-37FA-4718-9994-97DD76ED70A7"'
	Private Sub MoveParty()
		Dim Map_Renamed As Object
		Dim Tome_Renamed As Object
		'UPGRADE_NOTE: my was upgraded to my_Renamed. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="A9E4979A-37FA-4718-9994-97DD76ED70A7"'
		Dim X, mx, my_Renamed, Y As Short
		Dim cy, cx, Found As Short
		Dim Friendly, i, c, k, NoFail As Short
		Dim MaxX, MinX, MinY, MaxY As Short
		Dim rc, cdir As Short
		Dim CreatureX As Creature
		Dim TileX As Tile
		Dim NewEncounter As Short
		Dim EntryPointX As EntryPoint
		Dim Start, sinTooHeavy, PauseTime As Single
		Dim CreatureY As Creature
		' Find minimum and maximum coordinates of Party
		MinX = 32700 : MinY = 32700
		' If Party standing in destination, then wipe destination. Disable movement.
		'UPGRADE_WARNING: Couldn't resolve default property of object Tome.MoveToY. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
		'UPGRADE_WARNING: Couldn't resolve default property of object Tome.MapY. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
		'UPGRADE_WARNING: Couldn't resolve default property of object Tome.MoveToX. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
		'UPGRADE_WARNING: Couldn't resolve default property of object Tome.MapX. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
		If Tome.MapX = Tome.MoveToX And Tome.MapY = Tome.MoveToY Then
			'UPGRADE_WARNING: Couldn't resolve default property of object Tome.MoveToX. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
			Tome.MoveToX = 0
			'UPGRADE_WARNING: Couldn't resolve default property of object Tome.MoveToY. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
			Tome.MoveToY = 0
			tmrMoveParty.Enabled = False
		End If
		' Move Tome
		NewEncounter = False
		If tmrMoveParty.Enabled = True Then
			' [Titi 2.4.8] slow movement if too heavy!
			'UPGRADE_WARNING: Couldn't resolve default property of object Tome.Creatures. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
			For	Each CreatureX In Tome.Creatures
				If (CreatureX.Weight > CreatureX.MaxWeight * 0.75 Or CreatureX.Bulk > 0.75 * CreatureX.Agility * 8) And sinTooHeavy <> 2 Then
					sinTooHeavy = 1
					If Not CreatureY Is Nothing Then
						If CreatureY.Weight / CreatureY.MaxWeight < CreatureX.Weight / CreatureX.MaxWeight Then
							CreatureY = CreatureX
						End If
					Else
						CreatureY = CreatureX
					End If
				End If
				If CreatureX.Weight > CreatureX.MaxWeight * 0.9 Or CreatureX.Bulk > 0.9 * CreatureX.Agility * 8 Then
					sinTooHeavy = 2
					If Not CreatureY Is Nothing Then
						If CreatureY.Weight / CreatureY.MaxWeight < CreatureX.Weight / CreatureX.MaxWeight Then
							CreatureY = CreatureX
						End If
					Else
						CreatureY = CreatureX
					End If
				End If
			Next CreatureX
			If sinTooHeavy <> 0 Then
				DialogSetUp(modGameGeneral.DLGTYPE.bdDlgNoReply)
				DialogBrief(CreatureY, CreatureY.Name & IIf(sinTooHeavy = 1, " is heavily loaded", " carries too much") & " and slows the party down.")
				' pause...
				PauseTime = sinTooHeavy
				Start = VB.Timer()
				Do While VB.Timer() < Start + PauseTime
					System.Windows.Forms.Application.DoEvents()
				Loop 
				' advance time
				If GlobalFastMove > 0 Then
					'UPGRADE_WARNING: Couldn't resolve default property of object Map.MovementCost. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
					TotalStepsTaken = TotalStepsTaken + (Map.MovementCost(mx, my_Renamed) + Int(sinTooHeavy)) * 3
				Else
					'UPGRADE_WARNING: Couldn't resolve default property of object Map.MovementCost. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
					TotalStepsTaken = TotalStepsTaken + Map.MovementCost(mx, my_Renamed) + Int(sinTooHeavy)
				End If
			End If
			' Party takes a step
			If GlobalFastMove > 0 Then
				'UPGRADE_WARNING: Couldn't resolve default property of object Tome.NextMaxStep. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
				Tome.NextMaxStep(mx, my_Renamed)
				'            If GlobalAutoCenter Then ' [Titi 2.4.9]
				'                MapCenter mx, my ' Ephestion [2.4.9]
				'                DrawMapAll ' Ephestion [2.4.9]
				'            End If
			Else
				'UPGRADE_WARNING: Couldn't resolve default property of object Tome.NextStep. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
				Tome.NextStep(mx, my_Renamed)
				'            If GlobalAutoCenter Then ' [Titi 2.4.9]
				'                MapCenter mx, my ' Ephestion [2.4.9]
				'                DrawMapAll ' Ephestion [2.4.9]
				'            End If
			End If
			' If moving to a new encounter, then save laststep information to FleeTo
			'UPGRADE_WARNING: Couldn't resolve default property of object Map.EncPointer. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
			If Map.EncPointer(mx, my_Renamed) <> EncounterNow.Index Then
				' Check to see if the new Encounter is truly Active
				'UPGRADE_WARNING: Couldn't resolve default property of object Map.EncPointer. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
				If Map.EncPointer(mx, my_Renamed) > 0 Then
					'UPGRADE_WARNING: Couldn't resolve default property of object Map.EncPointer. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
					'UPGRADE_WARNING: Couldn't resolve default property of object Map.Encounters. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
					If Map.Encounters("E" & Map.EncPointer(mx, my_Renamed)).IsActive = True Then
						NewEncounter = True
					End If
				Else
					NewEncounter = True
				End If
				'UPGRADE_WARNING: Couldn't resolve default property of object Tome.FleeToX. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
				'UPGRADE_WARNING: Couldn't resolve default property of object Tome.MapX. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
				Tome.FleeToX = Tome.MapX
				'UPGRADE_WARNING: Couldn't resolve default property of object Tome.FleeToY. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
				'UPGRADE_WARNING: Couldn't resolve default property of object Tome.MapY. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
				Tome.FleeToY = Tome.MapY
			End If
			'UPGRADE_WARNING: Couldn't resolve default property of object Tome.MapX. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
			Tome.MapX = mx
			'UPGRADE_WARNING: Couldn't resolve default property of object Tome.MapY. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
			Tome.MapY = my_Renamed
			' Add one to steps. If exceed 9, then take a turn
			If GlobalFastMove > 0 Then
				'UPGRADE_WARNING: Couldn't resolve default property of object Map.MovementCost. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
				TotalStepsTaken = TotalStepsTaken + Map.MovementCost(mx, my_Renamed) * 3
			Else
				'UPGRADE_WARNING: Couldn't resolve default property of object Map.MovementCost. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
				TotalStepsTaken = TotalStepsTaken + Map.MovementCost(mx, my_Renamed)
			End If
			If TotalStepsTaken > 9 Then
				Do Until TotalStepsTaken < 10
					TurnCycle()
					TotalStepsTaken = TotalStepsTaken - 10
				Loop 
			End If
			If NewEncounter = False Then
				ItemNow = CreatureWithTurn.ItemInHand
				NoFail = FireTriggers(EncounterNow, bdPostStepEncounter)
			End If
		End If
		'UPGRADE_WARNING: Couldn't resolve default property of object Tome.MapX. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
		mx = Tome.MapX
		'UPGRADE_WARNING: Couldn't resolve default property of object Tome.MapY. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
		my_Renamed = Tome.MapY
		' Reset back to a human controlled player instead of an NPC
		Found = False
		If CreatureWithTurn.DMControlled = True Then
			'UPGRADE_WARNING: Couldn't resolve default property of object Tome.Creatures. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
			For c = 1 To Tome.Creatures.Count
				'UPGRADE_WARNING: Couldn't resolve default property of object Tome.Creatures. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
				If Tome.Creatures(c).DMControlled = False Then
					Found = True
					Exit For
				End If
			Next c
		End If
		If Found = True Then
			If c < PartyLeft Or c > PartyLeft + 4 Then
				PartyLeft = c - 1
				MenuDrawParty()
			End If
			SetTurn(c - 1)
		End If
		' Line of Site fill to show any hidden tiles
		'UPGRADE_WARNING: Couldn't resolve default property of object Map.Hidden. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
		Map.Hidden(mx, my_Renamed) = False
		For k = 0 To 3
			For c = 0 To 5 : cx = mx : cy = my_Renamed : For i = 0 To 5
					cdir = LosDir(c, i)
					Select Case k
						Case 0 ' Normal
						Case 1 ' Flip X
							If LosDir(c, i) = 1 Then
								cdir = 2
							End If
						Case 2 ' Flip Y
							If LosDir(c, i) = 0 Then
								cdir = 3
							End If
						Case 3 ' Flip X and Y
							If LosDir(c, i) = 1 Then
								cdir = 2
							End If
							If LosDir(c, i) = 0 Then
								cdir = 3
							End If
					End Select
					'UPGRADE_WARNING: Couldn't resolve default property of object Map.See. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
					If Map.See(cx, cy, cdir) = False Then
						Select Case cdir
							Case 0 ' Up
								cy = Greatest(cy - 1, 0)
							Case 1 ' Right
								'UPGRADE_WARNING: Couldn't resolve default property of object Map.Width. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
								cx = Least(cx + 1, Map.Width)
							Case 2 ' Left
								cx = Greatest(cx - 1, 0)
							Case 3 ' Down
								'UPGRADE_WARNING: Couldn't resolve default property of object Map.Height. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
								cy = Least(cy + 1, Map.Height)
						End Select
						'UPGRADE_WARNING: Couldn't resolve default property of object Map.Hidden. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
						If Map.Hidden(cx, cy) = True Then
							'UPGRADE_WARNING: Couldn't resolve default property of object Map.Hidden. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
							Map.Hidden(cx, cy) = False
							'UPGRADE_WARNING: Couldn't resolve default property of object Map.Left. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
							'UPGRADE_WARNING: Couldn't resolve default property of object Map.Top. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
							MinX = Least(((cy - Map.Top) + (cx - Map.Left)) * 48, MinX)
							'UPGRADE_WARNING: Couldn't resolve default property of object Map.Left. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
							'UPGRADE_WARNING: Couldn't resolve default property of object Map.Top. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
							MinY = Least(((cy - Map.Top) - (cx - Map.Left)) * 24 - 24, MinY)
							'UPGRADE_WARNING: Couldn't resolve default property of object Map.Left. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
							'UPGRADE_WARNING: Couldn't resolve default property of object Map.Top. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
							MaxX = Greatest(((cy - Map.Top) + (cx - Map.Left)) * 48 + 96, MaxX)
							'UPGRADE_WARNING: Couldn't resolve default property of object Map.Left. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
							'UPGRADE_WARNING: Couldn't resolve default property of object Map.Top. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
							MaxY = Greatest(((cy - Map.Top) - (cx - Map.Left)) * 24 + 48, MaxY)
						End If
					Else
						Exit For
					End If
				Next i : Next c
		Next k
		' Determine bounding rectangle for party
		'UPGRADE_WARNING: Couldn't resolve default property of object Tome.Creatures. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
		For	Each CreatureX In Tome.Creatures
			' Anchor at corners of picture
			'UPGRADE_WARNING: Couldn't resolve default property of object Map.Left. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
			'UPGRADE_WARNING: Couldn't resolve default property of object Map.Top. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
			X = (CreatureX.MapY + CreatureX.MapX - Map.Top - Map.Left) * 48 + CreatureX.TileSpotX
			'UPGRADE_WARNING: Couldn't resolve default property of object Map.Left. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
			'UPGRADE_WARNING: Couldn't resolve default property of object Map.Top. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
			Y = (CreatureX.MapY - CreatureX.MapX - Map.Top + Map.Left) * 24 + CreatureX.TileSpotY - 24
			MinX = Least(X, MinX)
			MinY = Least(Y - picCMap(CreatureX.Pic).Height / 2, MinY)
			MaxX = Greatest(X + picCMap(CreatureX.Pic).Width / 2, MaxX)
			MaxY = Greatest(Y, MaxY)
		Next CreatureX
		DrawMapRegion(MinY - 72, MinX - 48, MaxY + 96, MaxX + 144)
		'    DrawMapRegion MinY - 24, MinX - 24, MaxY + 48, MaxX + 96
		' If NewEncounter found, then deal with it
		If NewEncounter = True Then
			'UPGRADE_WARNING: Couldn't resolve default property of object Map.EncPointer. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
			If Map.EncPointer(mx, my_Renamed) < 1 Then
				EncounterNow = New Encounter
			Else
				'UPGRADE_WARNING: Couldn't resolve default property of object Map.EncPointer. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
				'UPGRADE_WARNING: Couldn't resolve default property of object Map.Encounters. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
				EncounterNow = Map.Encounters("E" & Map.EncPointer(mx, my_Renamed))
				EncounterEnter()
			End If
		End If
		'UPGRADE_WARNING: Couldn't resolve default property of object Tome.MapY. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
		'UPGRADE_WARNING: Couldn't resolve default property of object Tome.MoveToY. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
		'UPGRADE_WARNING: Couldn't resolve default property of object Tome.MapX. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
		'UPGRADE_WARNING: Couldn't resolve default property of object Tome.MoveToX. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
		If Tome.MoveToX = Tome.MapX And Tome.MoveToY = Tome.MapY Then
			' save coordinates in case the tileset for the next area is missing
			'UPGRADE_WARNING: Couldn't resolve default property of object Map.EntryPoints. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
			For	Each EntryPointX In Map.EntryPoints
				'UPGRADE_WARNING: Couldn't resolve default property of object Tome.MapY. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
				'UPGRADE_WARNING: Couldn't resolve default property of object Tome.MapX. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
				If EntryPointX.MapX = Tome.MapX And EntryPointX.MapY = Tome.MapY Then
					FromEntryPoint = EntryPointX
					Exit For
				End If
			Next EntryPointX
			' If have arrived, then stop the party
			tmrMoveParty.Enabled = False
			MovePartyExit()
		End If
		If picJournal.Visible = True And JournalMode = 2 And tmrMoveParty.Enabled = False Then
			JournalShowMap()
		End If
	End Sub
	
	'UPGRADE_NOTE: Tome was upgraded to Tome_Renamed. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="A9E4979A-37FA-4718-9994-97DD76ED70A7"'
	'UPGRADE_NOTE: Map was upgraded to Map_Renamed. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="A9E4979A-37FA-4718-9994-97DD76ED70A7"'
	Private Sub MovePartyExit(Optional ByRef blnDontShowMsg As Boolean = False)
		Dim Map_Renamed As Object
		Dim Tome_Renamed As Object
		'UPGRADE_NOTE: my was upgraded to my_Renamed. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="A9E4979A-37FA-4718-9994-97DD76ED70A7"'
		Dim mx, my_Renamed As Short
		Dim cy, cx, Found As Short
		Dim i, c, k As Short
		Dim cdir As Short
		Dim EntryPointX As EntryPoint
		' Determine if stopped on an EntryPoint
		For	Each EntryPointX In Map.EntryPoints
			'UPGRADE_WARNING: Couldn't resolve default property of object Tome.MapY. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
			'UPGRADE_WARNING: Couldn't resolve default property of object Tome.MapX. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
			If Tome.MapX = EntryPointX.MapX And Tome.MapY = EntryPointX.MapY Then
				tmrMoveParty.Enabled = False
				' Ask Yes/No question to leave through exit
				ConvoAction = 0
				If Not (blnDontShowMsg And FromEntryPoint Is Nothing) Then ' but only if we're not coming back from an error
					DialogSetUp(modGameGeneral.DLGTYPE.bdDlgNoReply)
					DialogSetButton(1, "No")
					DialogSetButton(2, "Yes")
					If Len(EntryPointX.Description) > 0 Then
						DialogShow("DM", (EntryPointX.Description))
					ElseIf EntryPointX.AreaIndex = 0 Then 
						DialogShow("DM", "Do you want to exit this Tome and return to the Main Menu?")
					Else
						DialogShow("DM", "Do you want to exit to the next area?")
					End If
				End If
				If ConvoAction = 0 Then
					' If leaving to the MainMenu
					If EntryPointX.AreaIndex = 0 Then
						TomeStartArea(0, 0, 0, 0, 0)
					Else
						TomeStartArea((EntryPointX.AreaIndex), (EntryPointX.MapIndex), (EntryPointX.EntryIndex), (EntryPointX.MapX), (EntryPointX.MapY))
						'UPGRADE_WARNING: Couldn't resolve default property of object Tome.MapX. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
						mx = Tome.MapX
						'UPGRADE_WARNING: Couldn't resolve default property of object Tome.MapY. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
						my_Renamed = Tome.MapY
						'UPGRADE_WARNING: Couldn't resolve default property of object Tome.MoveToX. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
						Tome.MoveToX = mx
						'UPGRADE_WARNING: Couldn't resolve default property of object Tome.MoveToY. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
						Tome.MoveToY = my_Renamed
						MapCenter(mx, my_Renamed)
						'UPGRADE_WARNING: Couldn't resolve default property of object Map.Hidden. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
						Map.Hidden(mx, my_Renamed) = False
						For k = 0 To 3
							For c = 0 To 5 : cx = mx : cy = my_Renamed : For i = 0 To 5
									cdir = LosDir(c, i)
									Select Case k
										Case 0 ' Normal
										Case 1 ' Flip X
											If LosDir(c, i) = 1 Then
												cdir = 2
											End If
										Case 2 ' Flip Y
											If LosDir(c, i) = 0 Then
												cdir = 3
											End If
										Case 3 ' Flip X and Y
											If LosDir(c, i) = 1 Then
												cdir = 2
											End If
											If LosDir(c, i) = 0 Then
												cdir = 3
											End If
									End Select
									'UPGRADE_WARNING: Couldn't resolve default property of object Map.See. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
									If Map.See(cx, cy, cdir) = False Then
										Select Case cdir
											Case 0 ' Up
												cy = Greatest(cy - 1, 0)
											Case 1 ' Right
												'UPGRADE_WARNING: Couldn't resolve default property of object Map.Width. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
												cx = Least(cx + 1, Map.Width)
											Case 2 ' Left
												cx = Greatest(cx - 1, 0)
											Case 3 ' Down
												'UPGRADE_WARNING: Couldn't resolve default property of object Map.Height. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
												cy = Least(cy + 1, Map.Height)
										End Select
										'UPGRADE_WARNING: Couldn't resolve default property of object Map.Hidden. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
										If Map.Hidden(cx, cy) = True Then
											'UPGRADE_WARNING: Couldn't resolve default property of object Map.Hidden. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
											Map.Hidden(cx, cy) = False
										End If
									Else
										Exit For
									End If
								Next i : Next c
						Next k
						' Cycle a turn to check map lighting
						TurnCycleTime()
					End If
				End If
				DialogHide()
				Exit For
			End If
		Next EntryPointX
	End Sub
	
	'UPGRADE_NOTE: Tome was upgraded to Tome_Renamed. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="A9E4979A-37FA-4718-9994-97DD76ED70A7"'
	Public Sub EncounterEnter()
		Dim Tome_Renamed As Object
		Dim CreatureX As Creature
		Dim ItemX As Item
		Dim i, j As Short
		Dim Agressive, CreatureCount, Friendly, NoFail As Short
		' Fire Pre-Enter Triggers
		NoFail = FireTriggers(EncounterNow, bdPreEnterEncounter)
		' If ReGen set, do it now.
		If EncounterNow.GenerateUponEntry = True Then
			MakeEncounter(Map, EncounterNow)
		End If
		' Load up Creature pictures if missing
		For	Each CreatureX In EncounterNow.Creatures
			If CreatureX.Pic = 0 Then
				LoadCreaturePic(CreatureX)
			End If
		Next CreatureX
		' Load up Item pictures if missing
		For	Each ItemX In EncounterNow.Items
			If ItemX.Pic = 0 Then
				LoadItemPic(ItemX)
			End If
		Next ItemX
		' Fire Post-Enter Triggers
		ItemNow = CreatureWithTurn.ItemInHand
		NoFail = FireTriggers(EncounterNow, bdPostStepEncounter)
		' If nobody is in here, then just display intro stuff
		If EncounterNow.Creatures.Count() = 0 Then
			If EncounterNow.HaveEntered = False Then
				If Len(EncounterNow.FirstEntry) > 0 Then
					DialogDM((EncounterNow.FirstEntry))
				End If
			ElseIf Len(BreakText(EncounterNow.SecondEntry, 1)) > 0 Then 
				DialogDM(BreakText(EncounterNow.SecondEntry, 1))
			End If
		Else
			' Look and see if unfriendlies and agressives in the Encounter
			Friendly = True : CreatureCount = 0 : Agressive = False
			For	Each CreatureX In EncounterNow.Creatures
				If CreatureX.AllowedTurn = True Then
					CreatureCount = CreatureCount + 1
					If CreatureX.Friendly = False Then
						Friendly = False
					End If
					If CreatureX.Agressive = True Then
						Agressive = True
					End If
				End If
			Next CreatureX
			' Set up dialog buttons for Encounter
			If CreatureCount > 0 Then
				If Friendly = True Then
					' If friendly then you might be able to talk with them
					If EncounterNow.CanTalk = True Then
						DialogSetUp(modGameGeneral.DLGTYPE.bdDlgNoReply)
						DialogSetButton(1, "Talk")
						DialogSetButton(2, "Ignore")
					Else
						DialogSetUp(modGameGeneral.DLGTYPE.bdDlgNoReply)
						DialogSetButton(1, "Done")
					End If
				ElseIf Agressive = True Then 
					If EncounterNow.CanTalk = True Then
						DialogSetUp(modGameGeneral.DLGTYPE.bdDlgNoReply)
						DialogSetButton(1, "Talk")
						DialogSetButton(2, "Fight")
					Else
						DialogSetUp(modGameGeneral.DLGTYPE.bdDlgNoReply)
						DialogSetButton(1, "Fight")
					End If
				Else
					If EncounterNow.CanFlee = True And EncounterNow.CanTalk = True Then
						DialogSetUp(modGameGeneral.DLGTYPE.bdDlgNoReply)
						DialogSetButton(4, "Ignore")
						DialogSetButton(3, "Fight")
						DialogSetButton(2, "Flee")
						DialogSetButton(1, "Talk")
					ElseIf EncounterNow.CanTalk = True Then 
						DialogSetUp(modGameGeneral.DLGTYPE.bdDlgNoReply)
						DialogSetButton(3, "Ignore")
						DialogSetButton(2, "Fight")
						DialogSetButton(1, "Talk")
					Else
						DialogSetUp(modGameGeneral.DLGTYPE.bdDlgNoReply)
						DialogSetButton(3, "Ignore")
						DialogSetButton(2, "Fight")
						DialogSetButton(1, "Flee")
					End If
				End If
			Else
				DialogSetUp(modGameGeneral.DLGTYPE.bdDlgNoReply)
			End If
			' Now display it
			If Len(EncounterNow.FirstEntry) > 0 And EncounterNow.HaveEntered = False Then
				DialogShow(EncounterNow.Creatures.Item(1), (EncounterNow.FirstEntry))
				DialogHide()
				MenuNow = bdMenuEncounter
				DoAction(ConvoAction)
			ElseIf Len(BreakText(EncounterNow.SecondEntry, 1)) > 0 And EncounterNow.HaveEntered = True Then 
				DialogShow(EncounterNow.Creatures.Item(1), BreakText(EncounterNow.SecondEntry, 1))
				DialogHide()
				MenuNow = bdMenuEncounter
				DoAction(ConvoAction)
			ElseIf (EncounterNow.HaveEntered = False Or Agressive = True) And CreatureCount > 0 Then 
				If EncounterNow.Creatures.Count() = 1 Then
					'UPGRADE_WARNING: Couldn't resolve default property of object EncounterNow.Creatures(1).Name. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
					DialogShow(EncounterNow.Creatures.Item(1), "A single " & EncounterNow.Creatures.Item(1).Name & " is here.")
				Else
					' [Titi 2.4.9] check how many different creatures there are
					Dim arrCreatures(CreatureCount) As String
					Dim intNb(CreatureCount) As Short
					i = 1
					For	Each CreatureX In EncounterNow.Creatures
						If CreatureX.Friendly = False Then
							arrCreatures(i) = CreatureX.Name
							i = i + 1
						End If
					Next CreatureX
					' sort alphabetically
					For i = 1 To CreatureCount - 1
						For j = i + 1 To CreatureCount
							If arrCreatures(i) > arrCreatures(j) Then
								arrCreatures(0) = arrCreatures(i)
								arrCreatures(i) = arrCreatures(j)
								arrCreatures(j) = arrCreatures(0)
							End If
						Next j
					Next i
					' now, count how many of them
					j = 1
					For i = 1 To CreatureCount - 1
						intNb(j) = intNb(j) + 1
						If arrCreatures(i) <> arrCreatures(i + 1) Then
							j = j + 1
						End If
					Next 
					' don't forget to count the last one!
					intNb(j) = intNb(j) + 1
					' OK, now we have j-1 types, and intNb(1..j-1) each
					arrCreatures(0) = "" : i = 0 : intNb(0) = 1
					While i < j
						i = i + 1
						arrCreatures(0) = arrCreatures(0) & Str(intNb(i)) & " " & arrCreatures(intNb(0)) & IIf(intNb(i) > 1, "s, ", ", ")
						intNb(0) = intNb(0) + intNb(i)
					End While
					' erase the last ", "
					arrCreatures(0) = VB.Left(arrCreatures(0), Len(arrCreatures(0)) - 2)
					'                DialogShow EncounterNow.Creatures(1), "There are " & EncounterNow.Creatures.Count & " " & EncounterNow.Creatures(1).Name & "s here."
					DialogShow(EncounterNow.Creatures.Item(1), "There are " & arrCreatures(0) & " here.")
				End If
				DialogHide()
				MenuNow = bdMenuEncounter
				DoAction(ConvoAction)
			End If
		End If
		EncounterNow.HaveEntered = True
		ItemNow = CreatureWithTurn.ItemInHand
		NoFail = FireTriggers(EncounterNow, bdPostEnterEncounter)
		If Not NoFail Then
			'UPGRADE_WARNING: Couldn't resolve default property of object Tome.MoveToX. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
			'UPGRADE_WARNING: Couldn't resolve default property of object Tome.MapX. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
			Tome.MoveToX = Tome.MapX
			'UPGRADE_WARNING: Couldn't resolve default property of object Tome.MoveToY. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
			'UPGRADE_WARNING: Couldn't resolve default property of object Tome.MapY. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
			Tome.MoveToY = Tome.MapY
			tmrMoveParty.Enabled = False
		End If
	End Sub
	
	Private Sub EncounterIgnore()
		Dim NoFail As Short
		ItemNow = CreatureWithTurn.ItemInHand
		NoFail = FireTriggers(EncounterNow, bdOnIgnore)
	End Sub
	
	Private Function CombatFindPath(ByRef CreatureX As Creature, ByVal ToCol As Short, ByVal ToRow As Short) As Short
		Dim pxlist(500) As Short
		Dim pylist(500) As Short
		Dim llist(500) As Short
		Dim lMap(bdCombatWidth + 1, bdCombatHeight + 1) As Short
		
		Dim c, X, Y, i As Short
		Dim shortest, NumPoints, PathFound As Short
		Dim dy, dx, StickySpot As Short
		'UPGRADE_NOTE: Dir was upgraded to Dir_Renamed. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="A9E4979A-37FA-4718-9994-97DD76ED70A7"'
		Dim Dir_Renamed, TotalTiles As Short
		'UPGRADE_NOTE: my was upgraded to my_Renamed. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="A9E4979A-37FA-4718-9994-97DD76ED70A7"'
		Dim my_Renamed, sy, sx, mx, mv As Short
		Dim ny, ey, ex, nx, tx As Short
		Dim CloseY, CloseX, CloseRange As Short
		Dim oldmv, Failed As Short
		' If don't need to move anywhere, then pop out of this
		If CreatureX.Row = ToRow And CreatureX.Col = ToCol Then
			CombatStepTop = -1
			CombatFindPath = True
			Exit Function
		End If
		' Load starting point in my list
		pxlist(0) = CreatureX.Col : pylist(0) = CreatureX.Row
		sx = CreatureX.Col : sy = CreatureX.Row
		NumPoints = 1
		' Compute distance from starting point to ending point
		dx = sx - ToCol : dy = sy - ToRow
		llist(0) = (dx * dx) + (dy * dy)
		' Initialize our move cost where we are starting
		mv = 1
		' Set all coordinates in our traveled map to large value (we're not moving there)
		For X = 0 To bdCombatWidth : For Y = 0 To bdCombatHeight
				lMap(X, Y) = 32700
			Next Y : Next X
		' But plop a zero where we're starting from (we are moving there)
		lMap(sx, sy) = 0
		TotalTiles = 0
		CloseRange = 32700
		CombatStepTop = -1
		Do 
			TotalTiles = TotalTiles + 1
			' We only try 50 times to find a path (arbitrary number)
			If TotalTiles > 50 Then
				PathFound = 0
				Exit Do
			End If
			' Find the coordinate that has the shortest distance
			shortest = 0
			For c = 0 To NumPoints - 1
				If llist(c) < llist(shortest) Then
					shortest = c
				End If
			Next c
			' That's where we start from
			sx = pxlist(shortest)
			sy = pylist(shortest)
			' The shortest is swaped with the top of the stack
			pxlist(shortest) = pxlist(NumPoints - 1)
			pylist(shortest) = pylist(NumPoints - 1)
			llist(shortest) = llist(NumPoints - 1)
			' And we reduce the stack size by one
			NumPoints = NumPoints - 1
			' Move values will range from 0 (where we start) plus one for every square away
			oldmv = lMap(sx, sy)
			' For every direction, see if can move that way
			For Dir_Renamed = 0 To 5
				' Set a funky x direction based on the offset of the grid
				If Dir_Renamed > 3 Then
					tx = DirX(Dir_Renamed) + 2 * (sy Mod 2)
				Else
					tx = DirX(Dir_Renamed)
				End If
				' See if blocked by any Creature
				Failed = False : StickySpot = 0
				For i = 0 To CombatTurn
					If i <> TurnNow And Turns(i).Ref.HPNow > 0 Then
						' If Creature is standing in the block
						If Turns(i).Ref.Col = sx + tx And Turns(i).Ref.Row = sy + DirY(Dir_Renamed) Then
							Failed = True
						End If
					End If
				Next i
				' See if blocked by a CombatGrid space
				If EncounterNow.CombatGrid(sx + tx, sy + DirY(Dir_Renamed)) = 1 Then
					Failed = True
				End If
				' If not then add that space to the que
				If Not Failed And IsBetween(sx + tx, 0, bdCombatWidth) And IsBetween(sy + DirY(Dir_Renamed), 0, bdCombatHeight) Then
					' If can, look at the movement cost compare to where we are
					mx = sx + tx : my_Renamed = sy + DirY(Dir_Renamed)
					If lMap(sx, sy) < lMap(mx, my_Renamed) Then
						' If where we are is less than where we're moving...
						' (which means we haven't tried there yet)
						' Compute the distance from our destination point
						dx = mx - ToCol
						dy = my_Renamed - ToRow
						' Add a cost of one to the tile we were just in
						lMap(mx, my_Renamed) = oldmv + 1
						' Place on the top of the stack this new coordinate and distance
						pxlist(NumPoints) = mx
						pylist(NumPoints) = my_Renamed
						llist(NumPoints) = (dx * dx) + (dy * dy)
						' Store CloseX and CloseY
						If llist(NumPoints) < CloseRange Then
							CloseX = mx : CloseY = my_Renamed
							CloseRange = llist(NumPoints)
						End If
						NumPoints = NumPoints + 1
						' If we just moved to our final destination, you're done
						If (mx = ToCol) And (my_Renamed = ToRow) Then
							PathFound = 1
							Exit For
						End If
					End If
				End If
			Next Dir_Renamed
			If NumPoints = 0 Then
				PathFound = 0
				Exit Function
			End If
		Loop Until PathFound = 1
		' If could not find a path, find the closest point.
		If PathFound = 0 Then
			ToCol = CloseX
			ToRow = CloseY
		End If
		' Now back track best path
		sx = CreatureX.Col : sy = CreatureX.Row
		ex = ToCol : ey = ToRow
		Do While (ex <> sx) Or (ey <> sy)
			c = 32700
			For Dir_Renamed = 0 To 5
				' Have to set a funky x based on offset grid
				If Dir_Renamed > 3 Then
					tx = DirX(Dir_Renamed) + 2 * (ey Mod 2)
				Else
					tx = DirX(Dir_Renamed)
				End If
				mx = Greatest(Least(ex + tx, bdCombatWidth), 0)
				my_Renamed = Greatest(Least(ey + DirY(Dir_Renamed), bdCombatHeight), 0)
				If lMap(mx, my_Renamed) < c Then
					c = lMap(mx, my_Renamed)
					nx = mx : ny = my_Renamed
				End If
			Next Dir_Renamed
			CombatStepTop = CombatStepTop + 1
			CombatStepCol(CombatStepTop) = ex
			CombatStepRow(CombatStepTop) = ey
			ex = nx : ey = ny
		Loop 
	End Function
	
	'UPGRADE_NOTE: Map was upgraded to Map_Renamed. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="A9E4979A-37FA-4718-9994-97DD76ED70A7"'
	Private Sub FindPath(ByRef TomeX As Tome)
		Dim pxlist As Object
		Dim Map_Renamed As Object
		'UPGRADE_WARNING: Untranslated statement in FindPath. Please check source code.
		'UPGRADE_WARNING: Couldn't resolve default property of object Map.Width. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
		'UPGRADE_WARNING: Couldn't resolve default property of object Map.Height. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
		Dim pylist(CDbl(Map.Height + 1) * CDbl(Map.Width + 1)) As Short
		'UPGRADE_WARNING: Couldn't resolve default property of object Map.Width. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
		'UPGRADE_WARNING: Couldn't resolve default property of object Map.Height. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
		Dim llist(CDbl(Map.Height + 1) * CDbl(Map.Width + 1)) As Short
		'UPGRADE_WARNING: Couldn't resolve default property of object Map.Height. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
		'UPGRADE_WARNING: Couldn't resolve default property of object Map.Width. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
		Dim lMap(Map.Width, Map.Height) As Short
		Dim Y, X, c As Short
		Dim shortest, NumPoints, PathFound As Short
		Dim dx, dy As Short
		'UPGRADE_NOTE: Dir was upgraded to Dir_Renamed. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="A9E4979A-37FA-4718-9994-97DD76ED70A7"'
		Dim Dir_Renamed, TotalTiles As Short
		'UPGRADE_NOTE: my was upgraded to my_Renamed. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="A9E4979A-37FA-4718-9994-97DD76ED70A7"'
		Dim my_Renamed, sy, sx, mx, mv As Short
		Dim nx, ex, ey, ny As Short
		Dim CloseY, CloseX, CloseRange As Short
		Dim oldmv As Short
		' If don't need to move anywhere, then pop out of this
		If TomeX.MapX = TomeX.MoveToX And TomeX.MapY = TomeX.MoveToY Then
			Exit Sub
		End If
		' Load starting point in my list
		'UPGRADE_WARNING: Couldn't resolve default property of object pxlist(). Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
		pxlist(0) = TomeX.MapX : pylist(0) = TomeX.MapY
		sx = TomeX.MapX : sy = TomeX.MapY
		NumPoints = 1
		' Compute distance from starting point to ending point
		dx = sx - TomeX.MoveToX : dy = sy - TomeX.MoveToY
		llist(0) = (dx * dx) + (dy * dy)
		' Initialize our move cost where we are starting
		mv = 1
		' Set all coordinates in our traveled map to large value (we're not moving there)
		'UPGRADE_WARNING: Couldn't resolve default property of object Map.Width. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
		For X = 0 To Map.Width
			'UPGRADE_WARNING: Couldn't resolve default property of object Map.Height. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
			For Y = 0 To Map.Height
				lMap(X, Y) = 32700
			Next Y : Next X
		' But plop a zero where we're starting from (we are moving there)
		lMap(sx, sy) = 0
		TotalTiles = 0
		CloseRange = 32700
		Do 
			TotalTiles = TotalTiles + 1
			' We only try 250 times to find a path (arbitrary number)
			If TotalTiles > 250 Then
				PathFound = 0
				Exit Do
			End If
			' Find the coordinate that has the shortest distance
			shortest = 0
			For c = 0 To NumPoints - 1
				If llist(c) < llist(shortest) Then
					shortest = c
				End If
			Next c
			' That's where we start from
			'UPGRADE_WARNING: Couldn't resolve default property of object pxlist(). Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
			sx = pxlist(shortest)
			sy = pylist(shortest)
			' The shortest which was is swaped with the top of the stack
			'UPGRADE_WARNING: Couldn't resolve default property of object pxlist(). Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
			pxlist(shortest) = pxlist(NumPoints - 1)
			pylist(shortest) = pylist(NumPoints - 1)
			llist(shortest) = llist(NumPoints - 1)
			' And we reduce the stack size by one
			NumPoints = NumPoints - 1
			' Move values will range from 0 (where we start) plus one for every square away
			oldmv = lMap(sx, sy)
			' For every direction, see if can move that way
			For Dir_Renamed = 0 To 3
				'UPGRADE_WARNING: Couldn't resolve default property of object Map.Blocked. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
				If Not Map.Blocked(sx, sy, Dir_Renamed) Then
					' If can, look at the movement cost compare to where we are
					mx = sx + DirX(Dir_Renamed) : my_Renamed = sy + DirY(Dir_Renamed)
					If lMap(sx, sy) < lMap(mx, my_Renamed) Then
						' If where we are is less than where we're moving...
						' (which means we haven't tried there yet)
						' Compute the distance from our destination point
						dx = mx - TomeX.MoveToX
						dy = my_Renamed - TomeX.MoveToY
						' Add a cost of one to the tile we were just in
						lMap(mx, my_Renamed) = oldmv + 1
						' Place on the top of the stack this new coordinate and distance
						'UPGRADE_WARNING: Couldn't resolve default property of object pxlist(). Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
						pxlist(NumPoints) = mx
						pylist(NumPoints) = my_Renamed
						llist(NumPoints) = (dx * dx) + (dy * dy)
						' Store CloseX and CloseY
						If llist(NumPoints) < CloseRange Then
							CloseX = mx : CloseY = my_Renamed
							CloseRange = llist(NumPoints)
						End If
						NumPoints = NumPoints + 1
						' If we just moved to our final destination, you're done
						If (mx = TomeX.MoveToX) And (my_Renamed = TomeX.MoveToY) Then
							PathFound = 1
							Exit Do
						End If
						'UPGRADE_WARNING: Couldn't resolve default property of object Map.Width. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
						'UPGRADE_WARNING: Couldn't resolve default property of object Map.Height. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
						If NumPoints > CDbl(Map.Height + 1) * CDbl(Map.Width + 1) Then
							PathFound = 0
							Exit Do
						End If
					End If
				End If
			Next Dir_Renamed
			If NumPoints = 0 Then
				PathFound = 0
				Exit Do
			End If
		Loop 
		' If could not find a path, find the closest point.
		If PathFound = 0 Then
			TomeX.MoveToX = CloseX
			TomeX.MoveToY = CloseY
		End If
		' Now back track best path
		sx = TomeX.MapX : sy = TomeX.MapY
		ex = TomeX.MoveToX : ey = TomeX.MoveToY
		Do While (ex <> sx) Or (ey <> sy)
			c = 32700
			For Dir_Renamed = 0 To 3
				'UPGRADE_WARNING: Couldn't resolve default property of object Map.Blocked. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
				If Not Map.Blocked(ex, ey, Dir_Renamed) Then
					'UPGRADE_WARNING: Couldn't resolve default property of object Map.Width. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
					mx = Least(Greatest(ex + DirX(Dir_Renamed), 0), Map.Width)
					'UPGRADE_WARNING: Couldn't resolve default property of object Map.Height. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
					my_Renamed = Least(Greatest(ey + DirY(Dir_Renamed), 0), Map.Height)
					If lMap(mx, my_Renamed) < c Then
						c = lMap(mx, my_Renamed)
						nx = mx : ny = my_Renamed
					End If
				End If
			Next Dir_Renamed
			TomeX.AddStep(ex, ey)
			ex = nx : ey = ny
		Loop 
	End Sub
	
	Private Function HasTrigger(ByRef ObjectX As Object, ByRef AsTriggerType As Short) As Short
		Dim rc As Short
		Dim TriggerX As Trigger
		HasTrigger = False
		'UPGRADE_WARNING: Couldn't resolve default property of object ObjectX.Triggers. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
		For	Each TriggerX In ObjectX.Triggers
			If TriggerX.TriggerType = AsTriggerType Then
				HasTrigger = True
				Exit Function
			End If
		Next TriggerX
	End Function
	
	'UPGRADE_NOTE: Map was upgraded to Map_Renamed. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="A9E4979A-37FA-4718-9994-97DD76ED70A7"'
	'UPGRADE_NOTE: Right was upgraded to Right_Renamed. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="A9E4979A-37FA-4718-9994-97DD76ED70A7"'
	'UPGRADE_NOTE: Bottom was upgraded to Bottom_Renamed. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="A9E4979A-37FA-4718-9994-97DD76ED70A7"'
	'UPGRADE_NOTE: Left was upgraded to Left_Renamed. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="A9E4979A-37FA-4718-9994-97DD76ED70A7"'
	'UPGRADE_NOTE: Top was upgraded to Top_Renamed. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="A9E4979A-37FA-4718-9994-97DD76ED70A7"'
	Private Sub DrawMapRegion(ByRef Top_Renamed As Short, ByRef Left_Renamed As Short, ByRef Bottom_Renamed As Short, ByRef Right_Renamed As Short)
		Dim Map_Renamed As Object
		' Based on bounding coordinates in Pixels, draw the correct number
		' of tiles the correct way. Only draw the overlaps necessary
		' (only on the outer most region of the rectangle).
		' Draw Bottom Layer of entire region, Items, Creatures, Party
		' then Middle, Top.
		Dim Side, X, Y, i As Short
		Dim rc As Short
		'UPGRADE_NOTE: my was upgraded to my_Renamed. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="A9E4979A-37FA-4718-9994-97DD76ED70A7"'
		Dim XMap, mx, my_Renamed, YMap As Short
		Dim sx, sy As Short
		Dim FromY, ToY As Short
		Dim FromX, ToX As Short
		'UPGRADE_WARNING: Couldn't resolve default property of object Map.Top. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
		'UPGRADE_WARNING: Couldn't resolve default property of object Map.Left. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
		SetMapCursor(Left_Renamed, Top_Renamed, Map.Left, Map.Top, FromX, FromY, XMap, YMap)
		'UPGRADE_WARNING: Couldn't resolve default property of object Map.Top. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
		'UPGRADE_WARNING: Couldn't resolve default property of object Map.Left. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
		SetMapCursor(Right_Renamed, Bottom_Renamed, Map.Left, Map.Top, ToX, ToY, XMap, YMap)
		FromX = Int(FromX / 96) : FromY = Int(FromY / 24) + (XMap + YMap) Mod 2
		ToX = Int(ToX / 96) + 1 : ToY = Int(ToY / 24)
		' Draw Top of Region
		Y = FromY
		For X = FromX To ToX - 1
			Side = bdSideTop
			If X = FromX Then
				Side = Side Or bdSideLeft
			End If
			If X = ToX Then
				Side = Side Or bdSideRight
			End If
			DrawTile(X, Y, Side)
		Next X
		' Draw Middle of Region
		For Y = FromY + 1 To ToY
			For X = FromX To ToX - (Y Mod 2)
				Side = 0
				If X = FromX And (Y Mod 2) = 0 Then
					Side = Side Or bdSideLeft
				End If
				If X = ToX And (Y Mod 2) = 0 Then
					Side = Side Or bdSideRight
				End If
				DrawTile(X, Y, Side)
			Next X
		Next Y
		' Draw Bottom of Region
		For Y = ToY + 1 To ToY + 2
			For X = FromX To ToX - (Y Mod 2)
				If Y = ToY + 1 Then
					Side = bdSideFirstBottom
				Else
					Side = bdSideSecondBottom
				End If
				If X = FromX And (Y Mod 2) = 0 Then
					Side = Side Or bdSideLeft
				End If
				If X = ToX And (Y Mod 2) = 0 Then
					Side = Side Or bdSideRight
				End If
				DrawTile(X, Y, Side)
			Next X
		Next Y
		picMap.Refresh()
	End Sub
	
	'UPGRADE_NOTE: Map was upgraded to Map_Renamed. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="A9E4979A-37FA-4718-9994-97DD76ED70A7"'
	Private Sub ScrollMap(Optional ByRef MoveMapX As Object = Nothing, Optional ByRef MoveMapY As Object = Nothing)
		Dim Map_Renamed As Object
		Dim PosNow As POINTAPI
		Dim rc As Integer
		Dim X, Y As Short
		Static LastMove As Short
		' Calculate if mouse in borders of Screen
		'UPGRADE_NOTE: IsMissing() was changed to IsNothing(). Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="8AE1CB93-37AB-439A-A4FF-BE3B6760BB23"'
		If IsNothing(MoveMapX) And IsNothing(MoveMapY) Then
			rc = GetCursorPos(PosNow)
			If PosNow.Y < 0 And PosNow.X > 0 Then
				
			ElseIf PosNow.Y < (VB6.PixelsToTwipsY(System.Windows.Forms.Screen.PrimaryScreen.Bounds.Height) / VB6.TwipsPerPixelY) - picMenu.ClientRectangle.Height - 128 And (PosNow.X > VB6.PixelsToTwipsX(System.Windows.Forms.Screen.PrimaryScreen.Bounds.Width) / VB6.TwipsPerPixelX - 2 Or PosNow.X < 2) Then 
				'UPGRADE_WARNING: Couldn't resolve default property of object MoveMapX. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
				MoveMapX = System.Math.Sign(PosNow.X - 3)
				'UPGRADE_WARNING: Couldn't resolve default property of object MoveMapY. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
				MoveMapY = System.Math.Sign(PosNow.X - 3)
			ElseIf PosNow.Y > VB6.PixelsToTwipsY(System.Windows.Forms.Screen.PrimaryScreen.Bounds.Height) / VB6.TwipsPerPixelY - 16 Or PosNow.Y < 16 Then 
				'UPGRADE_WARNING: Couldn't resolve default property of object MoveMapX. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
				MoveMapX = -System.Math.Sign(PosNow.Y - 17)
				'UPGRADE_WARNING: Couldn't resolve default property of object MoveMapY. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
				MoveMapY = System.Math.Sign(PosNow.Y - 17)
			ElseIf LastMove = True And picJournal.Visible = True And JournalMode = 2 Then 
				LastMove = False
				'UPGRADE_WARNING: Couldn't resolve default property of object Map.Left. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
				MicroMapLeft = Map.Left
				'UPGRADE_WARNING: Couldn't resolve default property of object Map.Top. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
				MicroMapTop = Map.Top
				JournalShowMap()
				Exit Sub
			Else
				Exit Sub
			End If
		End If
		' Check if moving will move the Party off the Map
		'UPGRADE_WARNING: Couldn't resolve default property of object Tome.MapX. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
		'UPGRADE_WARNING: Couldn't resolve default property of object MoveMapX. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
		'UPGRADE_WARNING: Couldn't resolve default property of object Map.Left. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
		'UPGRADE_WARNING: Couldn't resolve default property of object Tome.MapY. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
		'UPGRADE_WARNING: Couldn't resolve default property of object MoveMapY. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
		'UPGRADE_WARNING: Couldn't resolve default property of object Map.Top. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
		If System.Math.Sqrt((Map.Top + MoveMapY + Int(System.Math.Sqrt((picMap.ClientRectangle.Width ^ 2) + (picMap.ClientRectangle.Height ^ 2)) / 108) - 1 - Tome.MapY) ^ 2 + (Map.Left - 2 + MoveMapX - Tome.MapX) ^ 2) > (picMap.ClientRectangle.Height / 54) Then
			Exit Sub
		End If
		'UPGRADE_WARNING: Couldn't resolve default property of object Map.Top. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
		'UPGRADE_WARNING: Couldn't resolve default property of object MoveMapY. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
		Map.Top = Map.Top + MoveMapY
		'UPGRADE_WARNING: Couldn't resolve default property of object Map.Left. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
		'UPGRADE_WARNING: Couldn't resolve default property of object MoveMapX. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
		Map.Left = Map.Left + MoveMapX
		LastMove = True
		' Scroll Screen
		'UPGRADE_WARNING: Couldn't resolve default property of object MoveMapY. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
		'UPGRADE_WARNING: Couldn't resolve default property of object MoveMapX. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
		If MoveMapX = -1 And MoveMapY = -1 Then
			' Mouse Left (Shift Screen Right)
			'UPGRADE_ISSUE: PictureBox property picMap.hdc was not upgraded. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="CC4C7EC0-C903-48FC-ACCC-81861D12DA4A"'
			rc = BitBlt(picMap.hdc, 96, 0, picMap.ClientRectangle.Width - 96, picMap.ClientRectangle.Height, picMap.hdc, 0, 0, SRCCOPY)
			DrawMapRegion(-48, 0, picMap.ClientRectangle.Height, 96)
			'UPGRADE_WARNING: Couldn't resolve default property of object MoveMapY. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
			'UPGRADE_WARNING: Couldn't resolve default property of object MoveMapX. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
		ElseIf MoveMapX = 1 And MoveMapY = 1 Then 
			' Mouse Right (Shift Screen Left)
			'UPGRADE_ISSUE: PictureBox property picMap.hdc was not upgraded. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="CC4C7EC0-C903-48FC-ACCC-81861D12DA4A"'
			rc = BitBlt(picMap.hdc, 0, 0, picMap.ClientRectangle.Width - 96, picMap.ClientRectangle.Height, picMap.hdc, 96, 0, SRCCOPY)
			DrawMapRegion(0, picMap.ClientRectangle.Width - 96, picMap.ClientRectangle.Height, picMap.ClientRectangle.Width)
			'UPGRADE_WARNING: Couldn't resolve default property of object MoveMapY. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
			'UPGRADE_WARNING: Couldn't resolve default property of object MoveMapX. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
		ElseIf MoveMapX = 1 And MoveMapY = -1 Then 
			' Mouse Up (Shift Screen Down)
			'UPGRADE_ISSUE: PictureBox property picMap.hdc was not upgraded. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="CC4C7EC0-C903-48FC-ACCC-81861D12DA4A"'
			rc = BitBlt(picMap.hdc, 0, 48, picMap.ClientRectangle.Width, picMap.ClientRectangle.Height - 48, picMap.hdc, 0, 0, SRCCOPY)
			DrawMapRegion(0, 0, 48, picMap.ClientRectangle.Width)
			'UPGRADE_WARNING: Couldn't resolve default property of object MoveMapY. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
			'UPGRADE_WARNING: Couldn't resolve default property of object MoveMapX. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
		ElseIf MoveMapX = -1 And MoveMapY = 1 Then 
			' Mouse Down (Shift Screen Up)
			'UPGRADE_ISSUE: PictureBox property picMap.hdc was not upgraded. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="CC4C7EC0-C903-48FC-ACCC-81861D12DA4A"'
			rc = BitBlt(picMap.hdc, 0, 0, picMap.ClientRectangle.Width, picMap.ClientRectangle.Height - 48, picMap.hdc, 0, 48, SRCCOPY)
			DrawMapRegion(picMap.ClientRectangle.Height - 48, 0, picMap.ClientRectangle.Height, picMap.ClientRectangle.Width)
		End If
	End Sub
	
	
	Private Sub LoadPartyPics()
		' Load Party Pictures
		Dim c As Short
		Dim CreatureX As Creature
		'UPGRADE_WARNING: Screen property Screen.MousePointer has a new behavior. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6BA9B8D2-2A32-4B6E-8D36-44949974A5B4"'
		System.Windows.Forms.Cursor.Current = System.Windows.Forms.Cursors.WaitCursor ' Hourglass
		' Loop through Party and load pictures
		For	Each CreatureX In Tome.Creatures
			' Load Creature bitmap
			LoadCreaturePic(CreatureX)
		Next CreatureX
		'UPGRADE_WARNING: Screen property Screen.MousePointer has a new behavior. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6BA9B8D2-2A32-4B6E-8D36-44949974A5B4"'
		System.Windows.Forms.Cursor.Current = System.Windows.Forms.Cursors.Default ' Default
	End Sub
	
	'UPGRADE_NOTE: Tome was upgraded to Tome_Renamed. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="A9E4979A-37FA-4718-9994-97DD76ED70A7"'
	Public Function LoadSFXPic(ByRef PictureFile As String) As Short
		Dim Tome_Renamed As Object
		' Loads SFX picture and returns pointer to array
		Dim FileName As String
		Dim c, PicToLoad As Short
		'UPGRADE_WARNING: Arrays in structure bmBlack may need to be initialized before they can be used. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="814DF224-76BD-4BB4-BFFB-EA359CB9FC48"'
		Dim bmBlack As BITMAPINFO
		'UPGRADE_WARNING: Arrays in structure bmMask may need to be initialized before they can be used. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="814DF224-76BD-4BB4-BFFB-EA359CB9FC48"'
		Dim bmMask As BITMAPINFO
		Dim rc, lpMem, hMem, TransparentRGB As Integer
		'UPGRADE_NOTE: Height was upgraded to Height_Renamed. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="A9E4979A-37FA-4718-9994-97DD76ED70A7"'
		'UPGRADE_NOTE: Width was upgraded to Width_Renamed. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="A9E4979A-37FA-4718-9994-97DD76ED70A7"'
		Dim Width_Renamed, Height_Renamed As Short
		' If PictureFile already loaded, then return that reference
		For c = 0 To MaxSFX
			If SFXFile(c) = PictureFile Then
				LoadSFXPic = c
				Exit Function
			End If
		Next c
		' Find an empty picture box
		For c = 0 To MaxSFX
			If SFXFile(c) = "" Then
				Exit For
			End If
		Next c
		' If found one, use that. Else, load a new one.
		If c < MaxSFX Then
			PicToLoad = c
		Else
			MaxSFX = MaxSFX + 1
			PicToLoad = MaxSFX
			ReDim Preserve SFXFile(PicToLoad)
			If PicToLoad > 0 Then
				picSFXPic.Load(PicToLoad)
			End If
		End If
		SFXFile(PicToLoad) = PictureFile
		' Load Bitmap (try Tome path first). Then try other directories.
		'UPGRADE_WARNING: Couldn't resolve default property of object Tome.FullPath. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
		'UPGRADE_WARNING: Dir has a new behavior. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="9B7D5ADD-D8FE-4819-A36C-6DEDAF088CC7"'
		FileName = Dir(Tome.FullPath & "\" & PictureFile)
		If FileName = "" Then
			'UPGRADE_WARNING: Couldn't resolve default property of object Tome.FullPath. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
			'UPGRADE_WARNING: Dir has a new behavior. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="9B7D5ADD-D8FE-4819-A36C-6DEDAF088CC7"'
			FileName = Dir(Tome.FullPath & "\effects\" & PictureFile)
			If FileName = "" Then
				'UPGRADE_WARNING: Couldn't resolve default property of object Tome.FullPath. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
				'UPGRADE_WARNING: Dir has a new behavior. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="9B7D5ADD-D8FE-4819-A36C-6DEDAF088CC7"'
				FileName = Dir(Tome.FullPath & "\creatures\" & PictureFile)
				If FileName = "" Then
					'UPGRADE_WARNING: Couldn't resolve default property of object Tome.FullPath. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
					'UPGRADE_WARNING: Dir has a new behavior. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="9B7D5ADD-D8FE-4819-A36C-6DEDAF088CC7"'
					FileName = Dir(Tome.FullPath & "\tiles\" & PictureFile)
					'                If FileName = "" Then
					'                    LoadSFXPic = -1
					'                    Exit Function
					'                'Else
					'                '    FileName = gAppPath & "\data\graphics\tiles\" & PictureFile
					'                End If
				Else
					'UPGRADE_WARNING: Couldn't resolve default property of object Tome.FullPath. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
					FileName = Tome.FullPath & "\creatures\" & PictureFile
				End If
			Else
				'UPGRADE_WARNING: Couldn't resolve default property of object Tome.FullPath. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
				FileName = Tome.FullPath & "\effects\" & PictureFile
			End If
		Else
			'UPGRADE_WARNING: Couldn't resolve default property of object Tome.FullPath. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
			FileName = Tome.FullPath & "\" & PictureFile
		End If
		If FileName = "" Then
			'        FileName = Dir$(gAppPath & "\data\graphics\effects\" & PictureFile)
			'UPGRADE_WARNING: Dir has a new behavior. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="9B7D5ADD-D8FE-4819-A36C-6DEDAF088CC7"'
			FileName = Dir(gDataPath & "\graphics\effects\" & PictureFile)
			If FileName = "" Then
				'            FileName = Dir$(gAppPath & "\data\graphics\creatures\" & PictureFile)
				'UPGRADE_WARNING: Dir has a new behavior. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="9B7D5ADD-D8FE-4819-A36C-6DEDAF088CC7"'
				FileName = Dir(gDataPath & "\graphics\creatures\" & PictureFile)
				If FileName = "" Then
					'                FileName = Dir$(gAppPath & "\data\graphics\tiles\" & PictureFile)
					'UPGRADE_WARNING: Dir has a new behavior. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="9B7D5ADD-D8FE-4819-A36C-6DEDAF088CC7"'
					FileName = Dir(gDataPath & "\graphics\tiles\" & PictureFile)
					If FileName = "" Then
						LoadSFXPic = -1
						Exit Function
					End If
				Else
					'                FileName = gAppPath & "\data\graphics\creatures\" & PictureFile
					FileName = gDataPath & "\graphics\creatures\" & PictureFile
				End If
			Else
				'            FileName = gAppPath & "\data\graphics\effects\" & PictureFile
				FileName = gDataPath & "\graphics\effects\" & PictureFile
			End If
		End If
		ReadBitmapFile(FileName, bmBlack, hMem, TransparentRGB)
		' Convert to Mask (pure Blue is the mask color)
		MakeMask(bmBlack, bmMask, TransparentRGB)
		' Paint bitmap to picture box using converted palette
		lpMem = GlobalLock(hMem)
		Width_Renamed = Int(bmBlack.bmiHeader.biWidth * PicSize)
		Height_Renamed = Int(bmBlack.bmiHeader.biHeight * PicSize)
		picSFXPic(PicToLoad).Width = Width_Renamed * 2
		picSFXPic(PicToLoad).Height = Height_Renamed * 2
		' Draw Normal and Flip (X = 0, X = Width respectively)
		'UPGRADE_ISSUE: PictureBox property picSFXPic.hdc was not upgraded. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="CC4C7EC0-C903-48FC-ACCC-81861D12DA4A"'
		rc = SetStretchBltMode(picSFXPic(PicToLoad).hdc, 3)
		'UPGRADE_ISSUE: PictureBox property picSFXPic.hdc was not upgraded. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="CC4C7EC0-C903-48FC-ACCC-81861D12DA4A"'
		rc = StretchDIBits(picSFXPic(PicToLoad).hdc, 0, 0, Width_Renamed, Height_Renamed, 0, 0, bmBlack.bmiHeader.biWidth, bmBlack.bmiHeader.biHeight, lpMem, bmBlack, DIB_RGB_COLORS, SRCCOPY)
		'UPGRADE_ISSUE: PictureBox property picSFXPic.hdc was not upgraded. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="CC4C7EC0-C903-48FC-ACCC-81861D12DA4A"'
		rc = StretchBlt(picSFXPic(PicToLoad).hdc, Width_Renamed, 0, Width_Renamed, Height_Renamed, picSFXPic(PicToLoad).hdc, Width_Renamed - 1, 0, -Width_Renamed, Height_Renamed, SRCCOPY)
		' Draw Mask Normal and Flip (Y = Height, X = 0 and X = Width respectively)
		'UPGRADE_ISSUE: PictureBox property picSFXPic.hdc was not upgraded. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="CC4C7EC0-C903-48FC-ACCC-81861D12DA4A"'
		rc = StretchDIBits(picSFXPic(PicToLoad).hdc, 0, Height_Renamed, Width_Renamed, Height_Renamed, 0, 0, bmMask.bmiHeader.biWidth, bmMask.bmiHeader.biHeight, lpMem, bmMask, DIB_RGB_COLORS, SRCCOPY)
		'UPGRADE_ISSUE: PictureBox property picSFXPic.hdc was not upgraded. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="CC4C7EC0-C903-48FC-ACCC-81861D12DA4A"'
		rc = StretchBlt(picSFXPic(PicToLoad).hdc, Width_Renamed, Height_Renamed, Width_Renamed, Height_Renamed, picSFXPic(PicToLoad).hdc, Width_Renamed - 1, Height_Renamed, -Width_Renamed, Height_Renamed, SRCCOPY)
		picSFXPic(PicToLoad).Refresh()
		' Release memory
		rc = GlobalUnlock(hMem)
		rc = GlobalFree(hMem)
		' Return pointer to PictureBox
		LoadSFXPic = PicToLoad
	End Function
	
	Public Sub LoadCreaturePicForce(ByRef CreatureX As Creature)
		If CreatureX.Pic > -1 Then
			PicFile(CreatureX.Pic) = ""
		End If
	End Sub
	
	'UPGRADE_NOTE: Tome was upgraded to Tome_Renamed. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="A9E4979A-37FA-4718-9994-97DD76ED70A7"'
	Public Sub LoadCreaturePic(ByRef CreatureX As Creature)
		Dim Tome_Renamed As Object
		Dim FileName As String
		'UPGRADE_NOTE: Size was upgraded to Size_Renamed. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="A9E4979A-37FA-4718-9994-97DD76ED70A7"'
		Dim c, i As Short
		Dim k, Size_Renamed As Double
		Dim PicToLoad As Short
		Dim PartyX As Creature
		Dim Y, X, Keep As Short
		'UPGRADE_WARNING: Arrays in structure bmBlack may need to be initialized before they can be used. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="814DF224-76BD-4BB4-BFFB-EA359CB9FC48"'
		Dim bmBlack As BITMAPINFO
		'UPGRADE_WARNING: Arrays in structure bmMask may need to be initialized before they can be used. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="814DF224-76BD-4BB4-BFFB-EA359CB9FC48"'
		Dim bmMask As BITMAPINFO
		Dim rc, lpMem, hMem, TransparentRGB As Integer
		'UPGRADE_NOTE: Height was upgraded to Height_Renamed. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="A9E4979A-37FA-4718-9994-97DD76ED70A7"'
		'UPGRADE_NOTE: Width was upgraded to Width_Renamed. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="A9E4979A-37FA-4718-9994-97DD76ED70A7"'
		Dim Width_Renamed, Height_Renamed As Short
		Dim PartyWidth, PartyHeight As Short
		' If loading the bones.bmp picture, skip all this.
		If CreatureX.PictureFile = "bones.bmp" Or CreatureX.PictureFile = "" Then
			PicToLoad = 0
		Else
			' If PictureFile already loaded, then return that reference
			For c = 1 To bdMaxMonsPics
				If PicFile(c) = CreatureX.PictureFile And (PicPortrait(c) = CreatureX.PortraitFile Or Len(CreatureX.PortraitFile) = 0) Then
					CreatureX.Pic = c
					PicFileTime(c) = TimeOfDay.ToOADate()
					Exit Sub
				End If
			Next c
			' Sort to find oldest picture box free
			k = 1 : i = 0
			For c = 1 To bdMaxMonsPics
				Keep = False
				If PicFileTime(c) < k Then
					For	Each PartyX In Tome.Creatures
						If PartyX.PictureFile = PicFile(c) Then
							Keep = True
						End If
					Next PartyX
					If Not Keep Then
						i = c
						k = PicFileTime(c)
					End If
				End If
			Next c
			PicToLoad = i
			If PicFileTime(PicToLoad) = 0 Then
				picFaces.Width = 66 * PicToLoad + bdFaceMin + 66
				picFaces.Height = 76
				If PicToLoad > 0 Then
					picCPic.Load(PicToLoad)
					picCMap.Load(PicToLoad)
				End If
			End If
		End If
		PicFile(PicToLoad) = CreatureX.PictureFile
		PicPortrait(PicToLoad) = CreatureX.PortraitFile
		PicFileTime(PicToLoad) = TimeOfDay.ToOADate()
		' Load Bitmap
		'UPGRADE_WARNING: Couldn't resolve default property of object Tome.FullPath. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
		'UPGRADE_WARNING: Dir has a new behavior. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="9B7D5ADD-D8FE-4819-A36C-6DEDAF088CC7"'
		FileName = Dir(Tome.FullPath & "\" & CreatureX.PictureFile)
		If FileName = "" Then
			'UPGRADE_WARNING: Couldn't resolve default property of object Tome.FullPath. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
			'UPGRADE_WARNING: Dir has a new behavior. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="9B7D5ADD-D8FE-4819-A36C-6DEDAF088CC7"'
			FileName = Dir(Tome.FullPath & "\creatures\" & CreatureX.PictureFile)
			If FileName = "" Then
				'            FileName = gAppPath & "\data\graphics\creatures\" & CreatureX.PictureFile
				FileName = gDataPath & "\graphics\creatures\" & CreatureX.PictureFile
			Else
				'UPGRADE_WARNING: Couldn't resolve default property of object Tome.FullPath. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
				FileName = Tome.FullPath & "\creatures\" & CreatureX.PictureFile
			End If
		Else
			'UPGRADE_WARNING: Couldn't resolve default property of object Tome.FullPath. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
			FileName = Tome.FullPath & "\" & CreatureX.PictureFile
		End If
		ReadBitmapFile(FileName, bmBlack, hMem, TransparentRGB)
		' Convert to Mask (pure Blue is the mask color)
		MakeMask(bmBlack, bmMask, TransparentRGB)
		' Paint bitmap to picture box using converted palette
		lpMem = GlobalLock(hMem)
		Width_Renamed = Int(bmBlack.bmiHeader.biWidth * PicSize)
		Height_Renamed = Int(bmBlack.bmiHeader.biHeight * PicSize)
		picCPic(PicToLoad).Width = Width_Renamed * 2
		picCPic(PicToLoad).Height = Height_Renamed * 4
		'UPGRADE_ISSUE: PictureBox property picCPic.hdc was not upgraded. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="CC4C7EC0-C903-48FC-ACCC-81861D12DA4A"'
		rc = SetStretchBltMode(picCPic(PicToLoad).hdc, 3)
		' Draw Normal and Flip
		'UPGRADE_ISSUE: PictureBox property picCPic.hdc was not upgraded. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="CC4C7EC0-C903-48FC-ACCC-81861D12DA4A"'
		rc = StretchDIBits(picCPic(PicToLoad).hdc, 0, 0, Width_Renamed, Height_Renamed, 0, 0, bmBlack.bmiHeader.biWidth, bmBlack.bmiHeader.biHeight, lpMem, bmBlack, DIB_RGB_COLORS, SRCCOPY)
		'UPGRADE_ISSUE: PictureBox property picCPic.hdc was not upgraded. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="CC4C7EC0-C903-48FC-ACCC-81861D12DA4A"'
		rc = StretchBlt(picCPic(PicToLoad).hdc, Width_Renamed, 0, Width_Renamed, Height_Renamed, picCPic(PicToLoad).hdc, Width_Renamed - 1, 0, -Width_Renamed, Height_Renamed, SRCCOPY)
		' Paint Mask Normal and Flip
		'UPGRADE_ISSUE: PictureBox property picCPic.hdc was not upgraded. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="CC4C7EC0-C903-48FC-ACCC-81861D12DA4A"'
		rc = StretchDIBits(picCPic(PicToLoad).hdc, 0, Height_Renamed, Width_Renamed, Height_Renamed, 0, 0, bmMask.bmiHeader.biWidth, bmMask.bmiHeader.biHeight, lpMem, bmMask, DIB_RGB_COLORS, SRCCOPY)
		'UPGRADE_ISSUE: PictureBox property picCPic.hdc was not upgraded. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="CC4C7EC0-C903-48FC-ACCC-81861D12DA4A"'
		rc = StretchBlt(picCPic(PicToLoad).hdc, Width_Renamed, Height_Renamed, Width_Renamed, Height_Renamed, picCPic(PicToLoad).hdc, Width_Renamed - 1, Height_Renamed, -Width_Renamed, Height_Renamed, SRCCOPY)
		' Draw Normal and Flip (Below for Yellow Outline)
		'UPGRADE_ISSUE: PictureBox property picCPic.hdc was not upgraded. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="CC4C7EC0-C903-48FC-ACCC-81861D12DA4A"'
		rc = StretchDIBits(picCPic(PicToLoad).hdc, 0, Height_Renamed * 2, Width_Renamed, Height_Renamed, 0, 0, bmBlack.bmiHeader.biWidth, bmBlack.bmiHeader.biHeight, lpMem, bmBlack, DIB_RGB_COLORS, SRCCOPY)
		'UPGRADE_ISSUE: PictureBox property picCPic.hdc was not upgraded. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="CC4C7EC0-C903-48FC-ACCC-81861D12DA4A"'
		rc = StretchBlt(picCPic(PicToLoad).hdc, Width_Renamed, Height_Renamed * 2, Width_Renamed, Height_Renamed, picCPic(PicToLoad).hdc, Width_Renamed - 1, 0, -Width_Renamed, Height_Renamed, SRCCOPY)
		' Draw Normal and Flip (Below for Red Outline)
		'UPGRADE_ISSUE: PictureBox property picCPic.hdc was not upgraded. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="CC4C7EC0-C903-48FC-ACCC-81861D12DA4A"'
		rc = StretchDIBits(picCPic(PicToLoad).hdc, 0, Height_Renamed * 3, Width_Renamed, Height_Renamed, 0, 0, bmBlack.bmiHeader.biWidth, bmBlack.bmiHeader.biHeight, lpMem, bmBlack, DIB_RGB_COLORS, SRCCOPY)
		'UPGRADE_ISSUE: PictureBox property picCPic.hdc was not upgraded. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="CC4C7EC0-C903-48FC-ACCC-81861D12DA4A"'
		rc = StretchBlt(picCPic(PicToLoad).hdc, Width_Renamed, Height_Renamed * 3, Width_Renamed, Height_Renamed, picCPic(PicToLoad).hdc, Width_Renamed - 1, 0, -Width_Renamed, Height_Renamed, SRCCOPY)
		' Find outline and paint it yellow and black
		For Y = Height_Renamed To Height_Renamed * 2 - 1 : For X = 0 To Width_Renamed - 1
				'UPGRADE_ISSUE: PictureBox property picCPic.hdc was not upgraded. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="CC4C7EC0-C903-48FC-ACCC-81861D12DA4A"'
				rc = GetPixel(picCPic(PicToLoad).hdc, X, Y)
				If rc = System.Drawing.ColorTranslator.ToOle(System.Drawing.Color.Black) Then
					For c = 0 To 3
						Select Case c
							Case 0
								If X < Width_Renamed - 1 Then
									'UPGRADE_ISSUE: PictureBox property picCPic.hdc was not upgraded. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="CC4C7EC0-C903-48FC-ACCC-81861D12DA4A"'
									rc = GetPixel(picCPic(PicToLoad).hdc, X + 1, Y)
								End If
							Case 1
								If Y > 0 Then
									'UPGRADE_ISSUE: PictureBox property picCPic.hdc was not upgraded. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="CC4C7EC0-C903-48FC-ACCC-81861D12DA4A"'
									rc = GetPixel(picCPic(PicToLoad).hdc, X, Y - 1)
								End If
							Case 2
								If Y < Height_Renamed * 2 - 1 Then
									'UPGRADE_ISSUE: PictureBox property picCPic.hdc was not upgraded. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="CC4C7EC0-C903-48FC-ACCC-81861D12DA4A"'
									rc = GetPixel(picCPic(PicToLoad).hdc, X, Y + 1)
								End If
							Case 3
								If X > 0 Then
									'UPGRADE_ISSUE: PictureBox property picCPic.hdc was not upgraded. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="CC4C7EC0-C903-48FC-ACCC-81861D12DA4A"'
									rc = GetPixel(picCPic(PicToLoad).hdc, X - 1, Y)
								End If
						End Select
						If rc > System.Drawing.ColorTranslator.ToOle(System.Drawing.Color.Black) Then
							' Black Outline
							'UPGRADE_ISSUE: PictureBox property picCPic.hdc was not upgraded. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="CC4C7EC0-C903-48FC-ACCC-81861D12DA4A"'
							rc = SetPixel(picCPic(PicToLoad).hdc, X, Y - Height_Renamed, System.Drawing.ColorTranslator.ToOle(System.Drawing.Color.Black))
							'UPGRADE_ISSUE: PictureBox property picCPic.hdc was not upgraded. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="CC4C7EC0-C903-48FC-ACCC-81861D12DA4A"'
							rc = SetPixel(picCPic(PicToLoad).hdc, System.Math.Abs(X - Width_Renamed) + Width_Renamed, Y - Height_Renamed, System.Drawing.ColorTranslator.ToOle(System.Drawing.Color.Black))
							' Yellow Outline
							'UPGRADE_ISSUE: PictureBox property picCPic.hdc was not upgraded. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="CC4C7EC0-C903-48FC-ACCC-81861D12DA4A"'
							rc = SetPixel(picCPic(PicToLoad).hdc, X, Y + Height_Renamed, System.Drawing.ColorTranslator.ToOle(System.Drawing.Color.Yellow))
							'UPGRADE_ISSUE: PictureBox property picCPic.hdc was not upgraded. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="CC4C7EC0-C903-48FC-ACCC-81861D12DA4A"'
							rc = SetPixel(picCPic(PicToLoad).hdc, System.Math.Abs(X - Width_Renamed) + Width_Renamed, Y + Height_Renamed, System.Drawing.ColorTranslator.ToOle(System.Drawing.Color.Yellow))
							' Red Outline
							'UPGRADE_ISSUE: PictureBox property picCPic.hdc was not upgraded. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="CC4C7EC0-C903-48FC-ACCC-81861D12DA4A"'
							rc = SetPixel(picCPic(PicToLoad).hdc, X, Y + Height_Renamed * 2, System.Drawing.ColorTranslator.ToOle(System.Drawing.Color.Red))
							'UPGRADE_ISSUE: PictureBox property picCPic.hdc was not upgraded. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="CC4C7EC0-C903-48FC-ACCC-81861D12DA4A"'
							rc = SetPixel(picCPic(PicToLoad).hdc, System.Math.Abs(X - Width_Renamed) + Width_Renamed, Y + Height_Renamed * 2, System.Drawing.ColorTranslator.ToOle(System.Drawing.Color.Red))
							Exit For
						End If
					Next c
				End If
			Next X : Next Y
		picCPic(PicToLoad).Refresh()
		' Paint party size picture box
		PartyWidth = CShort(bmBlack.bmiHeader.biWidth * bdMapPartySize)
		PartyHeight = CShort(bmBlack.bmiHeader.biHeight * bdMapPartySize)
		picCMap(PicToLoad).Width = PartyWidth
		picCMap(PicToLoad).Height = PartyHeight * 2
		'UPGRADE_ISSUE: PictureBox property picCMap.hdc was not upgraded. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="CC4C7EC0-C903-48FC-ACCC-81861D12DA4A"'
		rc = SetStretchBltMode(picCMap(PicToLoad).hdc, 3)
		'UPGRADE_ISSUE: PictureBox property picCMap.hdc was not upgraded. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="CC4C7EC0-C903-48FC-ACCC-81861D12DA4A"'
		rc = StretchDIBits(picCMap(PicToLoad).hdc, 0, 0, PartyWidth, PartyHeight, 0, 0, bmBlack.bmiHeader.biWidth, bmBlack.bmiHeader.biHeight, lpMem, bmBlack, DIB_RGB_COLORS, SRCCOPY)
		'UPGRADE_ISSUE: PictureBox property picCMap.hdc was not upgraded. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="CC4C7EC0-C903-48FC-ACCC-81861D12DA4A"'
		rc = StretchDIBits(picCMap(PicToLoad).hdc, 0, PartyHeight, PartyWidth, PartyHeight, 0, 0, bmMask.bmiHeader.biWidth, bmMask.bmiHeader.biHeight, lpMem, bmMask, DIB_RGB_COLORS, SRCCOPY)
		' Draw line around party mask picture
		For Y = 0 To PartyHeight - 1 : For X = 0 To PartyWidth - 1
				'UPGRADE_ISSUE: PictureBox property picCMap.hdc was not upgraded. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="CC4C7EC0-C903-48FC-ACCC-81861D12DA4A"'
				rc = GetPixel(picCMap(PicToLoad).hdc, X, Y)
				If rc = System.Drawing.ColorTranslator.ToOle(System.Drawing.Color.Black) Then
					For c = 0 To 3
						Select Case c
							Case 0
								If X < PartyWidth - 1 Then
									'UPGRADE_ISSUE: PictureBox property picCMap.hdc was not upgraded. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="CC4C7EC0-C903-48FC-ACCC-81861D12DA4A"'
									rc = GetPixel(picCMap(PicToLoad).hdc, X + 1, Y)
								End If
							Case 1
								If Y > 0 Then
									'UPGRADE_ISSUE: PictureBox property picCMap.hdc was not upgraded. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="CC4C7EC0-C903-48FC-ACCC-81861D12DA4A"'
									rc = GetPixel(picCMap(PicToLoad).hdc, X, Y - 1)
								End If
							Case 2
								If Y < PartyHeight - 1 Then
									'UPGRADE_ISSUE: PictureBox property picCMap.hdc was not upgraded. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="CC4C7EC0-C903-48FC-ACCC-81861D12DA4A"'
									rc = GetPixel(picCMap(PicToLoad).hdc, X, Y + 1)
								End If
							Case 3
								If X > 0 Then
									'UPGRADE_ISSUE: PictureBox property picCMap.hdc was not upgraded. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="CC4C7EC0-C903-48FC-ACCC-81861D12DA4A"'
									rc = GetPixel(picCMap(PicToLoad).hdc, X - 1, Y)
								End If
						End Select
						If rc > System.Drawing.ColorTranslator.ToOle(System.Drawing.Color.Black) Then
							'UPGRADE_ISSUE: PictureBox property picCMap.hdc was not upgraded. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="CC4C7EC0-C903-48FC-ACCC-81861D12DA4A"'
							rc = SetPixel(picCMap(PicToLoad).hdc, X, Y + PartyHeight, System.Drawing.ColorTranslator.ToOle(System.Drawing.Color.Black))
							Exit For
						End If
					Next c
				End If
			Next X : Next Y
		picCMap(PicToLoad).Refresh()
		' If have portrait, fetch that and paint. Else, scale face from picture.
		'UPGRADE_ISSUE: PictureBox property picFaces.hdc was not upgraded. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="CC4C7EC0-C903-48FC-ACCC-81861D12DA4A"'
		rc = SetStretchBltMode(picFaces.hdc, 3)
		'UPGRADE_ISSUE: PictureBox property picFaces.hdc was not upgraded. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="CC4C7EC0-C903-48FC-ACCC-81861D12DA4A"'
		rc = BitBlt(picFaces.hdc, bdFaceMin + 66 * PicToLoad, 0, 66, 76, picFaces.hdc, bdFaceBlank, 0, SRCCOPY)
		If Len(CreatureX.PortraitFile) > 0 Then
			' Release memory
			rc = GlobalUnlock(hMem)
			rc = GlobalFree(hMem)
			' Load Bitmap
			'        Filename = Dir$(Tome.FullPath & "\" & CreatureX.PortraitFile)
			'        If Filename = "" Then
			'            Filename = Dir$(Tome.FullPath & "\creatures\" & CreatureX.PortraitFile)
			'            If Filename = "" Then
			'                Filename = Dir$(gAppPath & "\Data\Graphics\Creatures\" & CreatureX.PortraitFile)
			'            Else
			'                Filename = Tome.FullPath & "\creatures\" & CreatureX.PortraitFile
			'            End If
			'        Else
			'            Filename = Tome.FullPath & "\" & CreatureX.PortraitFile
			'        End If
			' [Titi 2.4.8] added the alternative portraits folders
			'UPGRADE_WARNING: Couldn't resolve default property of object Tome.FullPath. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
			'UPGRADE_WARNING: Dir has a new behavior. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="9B7D5ADD-D8FE-4819-A36C-6DEDAF088CC7"'
			If Dir(Tome.FullPath & "\" & CreatureX.PortraitFile) = CreatureX.PortraitFile Then
				'UPGRADE_WARNING: Couldn't resolve default property of object Tome.FullPath. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
				FileName = Tome.FullPath & "\" & CreatureX.PortraitFile
				'UPGRADE_WARNING: Couldn't resolve default property of object Tome.FullPath. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
				'UPGRADE_WARNING: Dir has a new behavior. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="9B7D5ADD-D8FE-4819-A36C-6DEDAF088CC7"'
			ElseIf Dir(Tome.FullPath & "\Creatures\" & CreatureX.PortraitFile) = CreatureX.PortraitFile Then 
				'UPGRADE_WARNING: Couldn't resolve default property of object Tome.FullPath. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
				FileName = Tome.FullPath & "\Creatures\" & CreatureX.PortraitFile
				'        ElseIf Dir$(gAppPath & "\Data\Graphics\Creatures\" & CreatureX.PortraitFile) = CreatureX.PortraitFile Then
				'UPGRADE_WARNING: Dir has a new behavior. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="9B7D5ADD-D8FE-4819-A36C-6DEDAF088CC7"'
			ElseIf Dir(gDataPath & "\Graphics\Creatures\" & CreatureX.PortraitFile) = CreatureX.PortraitFile Then 
				'            FileName = gAppPath & "\Data\Graphics\Creatures\" & CreatureX.PortraitFile
				FileName = gDataPath & "\Graphics\Creatures\" & CreatureX.PortraitFile
				'        ElseIf Dir$(gAppPath & "\Data\Graphics\Portraits\" & CreatureX.PortraitFile) = CreatureX.PortraitFile Then
				'UPGRADE_WARNING: Dir has a new behavior. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="9B7D5ADD-D8FE-4819-A36C-6DEDAF088CC7"'
			ElseIf Dir(gDataPath & "\Graphics\Portraits\" & CreatureX.PortraitFile) = CreatureX.PortraitFile Then 
				'            FileName = gAppPath & "\Data\Graphics\Portraits\" & CreatureX.PortraitFile
				FileName = gDataPath & "\Graphics\Portraits\" & CreatureX.PortraitFile
				'UPGRADE_WARNING: Couldn't resolve default property of object Tome.FullPath. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
				'UPGRADE_WARNING: Dir has a new behavior. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="9B7D5ADD-D8FE-4819-A36C-6DEDAF088CC7"'
			ElseIf Dir(Tome.FullPath & "\Portraits\" & CreatureX.PortraitFile) = CreatureX.PortraitFile Then 
				'UPGRADE_WARNING: Couldn't resolve default property of object Tome.FullPath. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
				FileName = Tome.FullPath & "\Portraits\" & CreatureX.PortraitFile
			Else
				'            FileName = gAppPath & "\Data\Graphics\Creatures\NoSuchFile.bmp"
				FileName = gDataPath & "\Graphics\Creatures\NoSuchFile.bmp"
			End If
			ReadBitmapFile(FileName, bmBlack, hMem, TransparentRGB)
			' Convert to Mask (pure Blue is the mask color)
			MakeMask(bmBlack, bmMask, TransparentRGB)
			' Paint bitmap to picture box using converted palette
			lpMem = GlobalLock(hMem)
			Width_Renamed = Int(bmBlack.bmiHeader.biWidth)
			Height_Renamed = Int(bmBlack.bmiHeader.biHeight)
			' Resize Creature to fit in space available
			Size_Renamed = 1
			If Width_Renamed > 66 Then
				Size_Renamed = 66 / Width_Renamed
			End If
			If Height_Renamed * Size_Renamed > 76 Then
				Size_Renamed = 76 / Height_Renamed
			End If
			' Center Creature in frame
			X = CShort(66 - (Width_Renamed * Size_Renamed)) / 2
			Y = CShort(76 - (Height_Renamed * Size_Renamed)) / 2
			Height_Renamed = CShort(Height_Renamed * Size_Renamed)
			Width_Renamed = CShort(Width_Renamed * Size_Renamed)
			'UPGRADE_ISSUE: PictureBox property picFaces.hdc was not upgraded. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="CC4C7EC0-C903-48FC-ACCC-81861D12DA4A"'
			rc = StretchDIBits(picFaces.hdc, bdFaceMin + 66 * PicToLoad + X, Y, 66, 76, 0, 0, bmBlack.bmiHeader.biWidth, bmBlack.bmiHeader.biHeight, lpMem, bmMask, DIB_RGB_COLORS, SRCAND)
			'UPGRADE_ISSUE: PictureBox property picFaces.hdc was not upgraded. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="CC4C7EC0-C903-48FC-ACCC-81861D12DA4A"'
			rc = StretchDIBits(picFaces.hdc, bdFaceMin + 66 * PicToLoad + X, Y, 66, 76, 0, 0, bmBlack.bmiHeader.biWidth, bmBlack.bmiHeader.biHeight, lpMem, bmBlack, DIB_RGB_COLORS, SRCPAINT)
			' Release memory
			rc = GlobalUnlock(hMem)
			rc = GlobalFree(hMem)
		Else
			'UPGRADE_ISSUE: PictureBox property picFaces.hdc was not upgraded. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="CC4C7EC0-C903-48FC-ACCC-81861D12DA4A"'
			rc = StretchDIBits(picFaces.hdc, bdFaceMin + 66 * PicToLoad, Greatest(-CreatureX.FaceTop - 15, 0), 66, 76, Greatest(CreatureX.FaceLeft - 6, 0), Greatest(bmBlack.bmiHeader.biHeight - Greatest(CreatureX.FaceTop - 15, 0) - 61, 0), Least(bmBlack.bmiHeader.biWidth - Greatest(CreatureX.FaceLeft - 6, 0), 53), Least(bmBlack.bmiHeader.biHeight - Greatest(CreatureX.FaceTop - 15, 0), 61), lpMem, bmMask, DIB_RGB_COLORS, SRCAND)
			'UPGRADE_ISSUE: PictureBox property picFaces.hdc was not upgraded. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="CC4C7EC0-C903-48FC-ACCC-81861D12DA4A"'
			rc = StretchDIBits(picFaces.hdc, bdFaceMin + 66 * PicToLoad, Greatest(-CreatureX.FaceTop - 15, 0), 66, 76, Greatest(CreatureX.FaceLeft - 6, 0), Greatest(bmBlack.bmiHeader.biHeight - Greatest(CreatureX.FaceTop - 15, 0) - 61, 0), Least(bmBlack.bmiHeader.biWidth - Greatest(CreatureX.FaceLeft - 6, 0), 53), Least(bmBlack.bmiHeader.biHeight - Greatest(CreatureX.FaceTop - 15, 0), 61), lpMem, bmBlack, DIB_RGB_COLORS, SRCPAINT)
			' Release memory
			rc = GlobalUnlock(hMem)
			rc = GlobalFree(hMem)
		End If
		picFaces.Refresh()
		' Return pointer to PictureBox
		CreatureX.Pic = PicToLoad
	End Sub
	
	Public Sub LoadDicePic(ByRef FileName As String)
		Dim c, i As Short
		Dim LoadFile As String
		Dim YY, X, Y, XX As Short
		Dim k As Double
		'UPGRADE_WARNING: Arrays in structure bmBlack may need to be initialized before they can be used. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="814DF224-76BD-4BB4-BFFB-EA359CB9FC48"'
		Dim bmBlack As BITMAPINFO
		'UPGRADE_WARNING: Arrays in structure bmMask may need to be initialized before they can be used. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="814DF224-76BD-4BB4-BFFB-EA359CB9FC48"'
		Dim bmMask As BITMAPINFO
		Dim rc, lpMem, hMem, TransparentRGB As Integer
		'UPGRADE_NOTE: Height was upgraded to Height_Renamed. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="A9E4979A-37FA-4718-9994-97DD76ED70A7"'
		'UPGRADE_NOTE: Width was upgraded to Width_Renamed. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="A9E4979A-37FA-4718-9994-97DD76ED70A7"'
		Dim Width_Renamed, Height_Renamed As Short
		' Load Bitmap
		'    LoadFile = gAppPath & "\data\interface\Dice\" & FileName
		LoadFile = gDataPath & "\interface\Dice\" & FileName
		ReadBitmapFile(LoadFile, bmBlack, hMem, TransparentRGB)
		' Convert to Mask
		MakeMask(bmBlack, bmMask, TransparentRGB)
		' Paint bitmap to picture box using converted palette
		lpMem = GlobalLock(hMem)
		Width_Renamed = Int(bmBlack.bmiHeader.biWidth)
		Height_Renamed = Int(bmBlack.bmiHeader.biHeight)
		picDice.Height = Height_Renamed : picDice.Width = Width_Renamed * 2
		' Paint the Item PIcture
		'UPGRADE_ISSUE: PictureBox property picDice.hdc was not upgraded. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="CC4C7EC0-C903-48FC-ACCC-81861D12DA4A"'
		rc = SetStretchBltMode(picDice.hdc, 3)
		'UPGRADE_ISSUE: PictureBox property picDice.hdc was not upgraded. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="CC4C7EC0-C903-48FC-ACCC-81861D12DA4A"'
		rc = StretchDIBits(picDice.hdc, 0, 0, Width_Renamed, Height_Renamed, 0, 0, bmBlack.bmiHeader.biWidth, bmBlack.bmiHeader.biHeight, lpMem, bmBlack, DIB_RGB_COLORS, SRCCOPY)
		'UPGRADE_ISSUE: PictureBox property picDice.hdc was not upgraded. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="CC4C7EC0-C903-48FC-ACCC-81861D12DA4A"'
		rc = StretchDIBits(picDice.hdc, Width_Renamed, 0, Width_Renamed, Height_Renamed, 0, 0, bmMask.bmiHeader.biWidth, bmMask.bmiHeader.biHeight, lpMem, bmMask, DIB_RGB_COLORS, SRCCOPY)
		picDice.Refresh()
		' Release memory
		rc = GlobalUnlock(hMem)
		rc = GlobalFree(hMem)
	End Sub
	
	'UPGRADE_NOTE: Tome was upgraded to Tome_Renamed. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="A9E4979A-37FA-4718-9994-97DD76ED70A7"'
	Public Sub LoadItemPic(ByRef ItemX As Item)
		Dim Tome_Renamed As Object
		Dim FileName As String
		Dim i, c, PicToLoad As Short
		Dim YY, X, Y, XX As Short
		Dim k As Double
		'UPGRADE_WARNING: Arrays in structure bmBlack may need to be initialized before they can be used. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="814DF224-76BD-4BB4-BFFB-EA359CB9FC48"'
		Dim bmBlack As BITMAPINFO
		'UPGRADE_WARNING: Arrays in structure bmMask may need to be initialized before they can be used. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="814DF224-76BD-4BB4-BFFB-EA359CB9FC48"'
		Dim bmMask As BITMAPINFO
		Dim rc, lpMem, hMem, TransparentRGB As Integer
		'UPGRADE_NOTE: Height was upgraded to Height_Renamed. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="A9E4979A-37FA-4718-9994-97DD76ED70A7"'
		'UPGRADE_NOTE: Width was upgraded to Width_Renamed. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="A9E4979A-37FA-4718-9994-97DD76ED70A7"'
		Dim Width_Renamed, Height_Renamed As Short
		' If PictureFile already loaded, then return that reference
		For c = 1 To bdMaxItemPics
			If ItemPicFile(c) = ItemX.PictureFile Then
				ItemX.Pic = c
				ItemPicTime(c) = TimeOfDay.ToOADate()
				Exit Sub
			End If
		Next c
		' Sort to find oldest picture box free
		k = 1 : i = 0
		For c = 1 To bdMaxItemPics
			If ItemPicTime(c) < k Then
				i = c
				k = ItemPicTime(c)
			End If
		Next c
		PicToLoad = i
		ItemPicFile(PicToLoad) = ItemX.PictureFile
		ItemPicTime(PicToLoad) = TimeOfDay.ToOADate()
		' Load Bitmap
		'UPGRADE_WARNING: Couldn't resolve default property of object Tome.FullPath. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
		'UPGRADE_WARNING: Dir has a new behavior. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="9B7D5ADD-D8FE-4819-A36C-6DEDAF088CC7"'
		FileName = Dir(Tome.FullPath & "\" & ItemX.PictureFile)
		If FileName = "" Then
			'UPGRADE_WARNING: Couldn't resolve default property of object Tome.FullPath. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
			'UPGRADE_WARNING: Dir has a new behavior. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="9B7D5ADD-D8FE-4819-A36C-6DEDAF088CC7"'
			FileName = Dir(Tome.FullPath & "\items\" & ItemX.PictureFile)
			If FileName = "" Then
				'            FileName = gAppPath & "\data\graphics\items\" & ItemX.PictureFile
				FileName = gDataPath & "\graphics\items\" & ItemX.PictureFile
			Else
				'UPGRADE_WARNING: Couldn't resolve default property of object Tome.FullPath. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
				FileName = Tome.FullPath & "\items\" & ItemX.PictureFile
			End If
		Else
			'UPGRADE_WARNING: Couldn't resolve default property of object Tome.FullPath. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
			FileName = Tome.FullPath & "\" & ItemX.PictureFile
		End If
		ReadBitmapFile(FileName, bmBlack, hMem, TransparentRGB)
		' Convert to Mask (pure Blue is the mask color)
		MakeMask(bmBlack, bmMask, TransparentRGB)
		' Paint bitmap to picture box using converted palette
		lpMem = GlobalLock(hMem)
		Width_Renamed = Int(bmBlack.bmiHeader.biWidth)
		Height_Renamed = Int(bmBlack.bmiHeader.biHeight)
		' Store Height/Width of Item Picture
		Select Case Height_Renamed
			Case 0 To 32 : ItemPicHeight(PicToLoad) = 32
			Case 33 To 64 : ItemPicHeight(PicToLoad) = 64
			Case Else : ItemPicHeight(PicToLoad) = 96
		End Select
		Select Case Width_Renamed
			Case 0 To 32 : ItemPicWidth(PicToLoad) = 32
			Case Else : ItemPicWidth(PicToLoad) = 64
		End Select
		' Center the Item Picture
		X = 64 * PicToLoad - 64 + (ItemPicWidth(PicToLoad) - Width_Renamed) / 2
		Y = (ItemPicHeight(PicToLoad) - Height_Renamed) / 2
		' Paint the Item PIcture (normal and mask)
		'UPGRADE_ISSUE: PictureBox property picItem.hdc was not upgraded. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="CC4C7EC0-C903-48FC-ACCC-81861D12DA4A"'
		rc = SetStretchBltMode(picItem.hdc, 3)
		'UPGRADE_ISSUE: PictureBox property picItem.hdc was not upgraded. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="CC4C7EC0-C903-48FC-ACCC-81861D12DA4A"'
		rc = StretchDIBits(picItem.hdc, X, Y, Width_Renamed, Height_Renamed, 0, 0, bmBlack.bmiHeader.biWidth, bmBlack.bmiHeader.biHeight, lpMem, bmBlack, DIB_RGB_COLORS, SRCCOPY)
		'UPGRADE_ISSUE: PictureBox property picItem.hdc was not upgraded. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="CC4C7EC0-C903-48FC-ACCC-81861D12DA4A"'
		rc = StretchDIBits(picItem.hdc, X, Y + 96, Width_Renamed, Height_Renamed, 0, 0, bmMask.bmiHeader.biWidth, bmMask.bmiHeader.biHeight, lpMem, bmMask, DIB_RGB_COLORS, SRCCOPY)
		' Paint micro size for map object
		'UPGRADE_ISSUE: PictureBox property picItem.hdc was not upgraded. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="CC4C7EC0-C903-48FC-ACCC-81861D12DA4A"'
		rc = StretchDIBits(picItem.hdc, X, Y + 96 * 2, ItemPicWidth(PicToLoad) / 3, ItemPicHeight(PicToLoad) / 3, 0, 0, bmBlack.bmiHeader.biWidth, bmBlack.bmiHeader.biHeight, lpMem, bmBlack, DIB_RGB_COLORS, SRCCOPY)
		'UPGRADE_ISSUE: PictureBox property picItem.hdc was not upgraded. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="CC4C7EC0-C903-48FC-ACCC-81861D12DA4A"'
		rc = StretchDIBits(picItem.hdc, X + 32, Y + 96 * 2, ItemPicWidth(PicToLoad) / 3, ItemPicHeight(PicToLoad) / 3, 0, 0, bmMask.bmiHeader.biWidth, bmMask.bmiHeader.biHeight, lpMem, bmMask, DIB_RGB_COLORS, SRCCOPY)
		' Find outline and paint it black
		For YY = Y + 96 To Y + 96 * 2 - 1 : For XX = X To X + Width_Renamed - 1
				'UPGRADE_ISSUE: PictureBox property picItem.hdc was not upgraded. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="CC4C7EC0-C903-48FC-ACCC-81861D12DA4A"'
				rc = GetPixel(picItem.hdc, XX, YY)
				If rc = System.Drawing.ColorTranslator.ToOle(System.Drawing.Color.Black) Then
					For c = 0 To 3
						Select Case c
							Case 0
								If XX < X + Width_Renamed - 1 Then
									'UPGRADE_ISSUE: PictureBox property picItem.hdc was not upgraded. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="CC4C7EC0-C903-48FC-ACCC-81861D12DA4A"'
									rc = GetPixel(picItem.hdc, XX + 1, YY)
								End If
							Case 1
								If YY > Y Then
									'UPGRADE_ISSUE: PictureBox property picItem.hdc was not upgraded. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="CC4C7EC0-C903-48FC-ACCC-81861D12DA4A"'
									rc = GetPixel(picItem.hdc, XX, YY - 1)
								End If
							Case 2
								If YY < Y + Height_Renamed * 2 - 1 Then
									'UPGRADE_ISSUE: PictureBox property picItem.hdc was not upgraded. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="CC4C7EC0-C903-48FC-ACCC-81861D12DA4A"'
									rc = GetPixel(picItem.hdc, XX, YY + 1)
								End If
							Case 3
								If XX > X Then
									'UPGRADE_ISSUE: PictureBox property picItem.hdc was not upgraded. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="CC4C7EC0-C903-48FC-ACCC-81861D12DA4A"'
									rc = GetPixel(picItem.hdc, XX - 1, YY)
								End If
						End Select
						If rc > System.Drawing.ColorTranslator.ToOle(System.Drawing.Color.Black) Then
							' Black Outline
							'UPGRADE_ISSUE: PictureBox property picItem.hdc was not upgraded. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="CC4C7EC0-C903-48FC-ACCC-81861D12DA4A"'
							rc = SetPixel(picItem.hdc, XX, YY - 96, System.Drawing.ColorTranslator.ToOle(System.Drawing.Color.Black))
							'UPGRADE_ISSUE: PictureBox property picItem.hdc was not upgraded. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="CC4C7EC0-C903-48FC-ACCC-81861D12DA4A"'
							rc = SetPixel(picItem.hdc, System.Math.Abs(XX - Width_Renamed) + Width_Renamed, YY - 96, System.Drawing.ColorTranslator.ToOle(System.Drawing.Color.Black))
							Exit For
						End If
					Next c
				End If
			Next XX : Next YY
		picItem.Refresh()
		' Release memory
		rc = GlobalUnlock(hMem)
		rc = GlobalFree(hMem)
		' Return pointer to PictureBox
		ItemX.Pic = PicToLoad
	End Sub
	
	'UPGRADE_NOTE: Map was upgraded to Map_Renamed. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="A9E4979A-37FA-4718-9994-97DD76ED70A7"'
	Private Sub LayerTile(ByRef Layer As Short, ByRef XMap As Short, ByRef YMap As Short, ByRef AtX As Short, ByRef AtY As Short, ByRef AtWidth As Short, ByRef AtHeight As Short, ByRef AtOffX As Short, ByRef AtOffY As Short, ByRef Gray As Short)
		Dim Map_Renamed As Object
		Dim c, Flip As Short
		Select Case Layer
			Case bdMapBottom
				'UPGRADE_WARNING: Couldn't resolve default property of object Map.BottomTile. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
				c = Map.BottomTile(XMap, YMap)
				'UPGRADE_WARNING: Couldn't resolve default property of object Map.Hidden. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
				If c > 0 And Map.Hidden(XMap, YMap) = False Then
					'UPGRADE_WARNING: Couldn't resolve default property of object Map.Tiles. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
					c = Map.Tiles("L" & c).Pic
				Else
					c = bdTileBlack
				End If
				'UPGRADE_WARNING: Couldn't resolve default property of object Map.BottomFlip. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
				Flip = Map.BottomFlip(XMap, YMap)
				PlotTile(c, AtX, AtY, AtWidth, AtHeight, AtOffX, AtOffY, Flip)
			Case bdMapMiddle
				'UPGRADE_WARNING: Couldn't resolve default property of object Map.MiddleTile. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
				c = Map.MiddleTile(XMap, YMap)
				'UPGRADE_WARNING: Couldn't resolve default property of object Map.Hidden. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
				'UPGRADE_WARNING: Couldn't resolve default property of object Map.Width. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
				If c > 0 And (Map.Hidden(XMap, YMap) = False Or Map.Hidden(Least(XMap + 1, Map.Width), YMap) = False Or Map.Hidden(XMap, Greatest(YMap - 1, 0)) = False) Then
					'UPGRADE_WARNING: Couldn't resolve default property of object Map.Tiles. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
					c = Map.Tiles("L" & c).Pic
					'UPGRADE_WARNING: Couldn't resolve default property of object Map.MiddleFlip. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
					Flip = Map.MiddleFlip(XMap, YMap)
					PlotTile(c, AtX, AtY, AtWidth, AtHeight, AtOffX, AtOffY, Flip)
				End If
			Case bdMapTop
				'UPGRADE_WARNING: Couldn't resolve default property of object Map.TopTile. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
				c = Map.TopTile(XMap, YMap)
				'UPGRADE_WARNING: Couldn't resolve default property of object Map.Hidden. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
				'UPGRADE_WARNING: Couldn't resolve default property of object Map.Width. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
				If c > 0 And (Map.Hidden(XMap, YMap) = False Or Map.Hidden(Least(XMap + 1, Map.Width), YMap) = False Or Map.Hidden(XMap, Greatest(YMap - 1, 0)) = False) Then
					'UPGRADE_WARNING: Couldn't resolve default property of object Map.Tiles. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
					c = Map.Tiles("L" & c).Pic
					'UPGRADE_WARNING: Couldn't resolve default property of object Map.TopFlip. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
					Flip = Map.TopFlip(XMap, YMap)
					PlotTile(c, AtX, AtY, AtWidth, AtHeight, AtOffX, AtOffY, Flip)
				End If
			Case bdMapDim
				'UPGRADE_WARNING: Couldn't resolve default property of object Map.TopTile. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
				c = Map.TopTile(XMap, YMap)
				'UPGRADE_WARNING: Couldn't resolve default property of object Map.Hidden. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
				'UPGRADE_WARNING: Couldn't resolve default property of object Map.Width. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
				If Map.Hidden(XMap, YMap) = True And (Map.Hidden(Least(XMap + 1, Map.Width), YMap) = False Or Map.Hidden(XMap, Greatest(YMap - 1, 0)) = False) Then
					PlotTile(bdTileDark, AtX, AtY, AtWidth, AtHeight, AtOffX, AtOffY, Flip)
				End If
		End Select
	End Sub
	
	'UPGRADE_NOTE: Map was upgraded to Map_Renamed. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="A9E4979A-37FA-4718-9994-97DD76ED70A7"'
	Private Sub DrawTile(ByRef AtX As Short, ByRef AtY As Short, ByRef Side As Short)
		Dim Map_Renamed As Object
		' AtX is a Pixel / 48, AtY is Pixel / 24
		'UPGRADE_NOTE: my was upgraded to my_Renamed. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="A9E4979A-37FA-4718-9994-97DD76ED70A7"'
		Dim mx, c, my_Renamed As Short
		Dim X, Y As Short
		Dim rc As Integer
		Dim OffX, OffY As Short
		'UPGRADE_NOTE: Height was upgraded to Height_Renamed. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="A9E4979A-37FA-4718-9994-97DD76ED70A7"'
		'UPGRADE_NOTE: Width was upgraded to Width_Renamed. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="A9E4979A-37FA-4718-9994-97DD76ED70A7"'
		Dim Width_Renamed, Height_Renamed As Short
		Dim XMap, YMap As Short
		' Convert AtX and AtY to Map Coordinates
		'UPGRADE_WARNING: Couldn't resolve default property of object Map.Left. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
		XMap = Map.Left + AtX - Int(AtY / 2)
		'UPGRADE_WARNING: Couldn't resolve default property of object Map.Top. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
		YMap = Map.Top + AtX + Int((AtY + 1) / 2) - 1
		X = (AtX * 2 - 1 + (AtY Mod 2)) * 48
		Y = (AtY - 2) * 24
		' The calc above is spot on. I've double checked it.
		OffX = 0 : OffY = 0
		Height_Renamed = 72 : Width_Renamed = 96
		' Set Offsets and Width/Height
		If (Side And bdSideTop) > 0 Then
			OffY = 48 : Height_Renamed = 24
		End If
		If (Side And bdSideLeft) > 0 Then
			OffX = 48 : Width_Renamed = 48
		End If
		If (Side And bdSideFirstBottom) > 0 Then
			Height_Renamed = 48
		End If
		If (Side And bdSideSecondBottom) > 0 Then
			Height_Renamed = 24
		End If
		If (Side And bdSideRight) > 0 Then
			Width_Renamed = 48
		End If
		' LayerTile: XMap, YMap, x, y, width, height, offx, offy
		' Plot blank tile if out of bounds
		'UPGRADE_WARNING: Couldn't resolve default property of object Map.Height. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
		'UPGRADE_WARNING: Couldn't resolve default property of object Map.Width. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
		If XMap < 0 Or YMap < 0 Or XMap > Map.Width Or YMap > Map.Height Then
			PlotTile(bdTileBlack, X + OffX, Y + OffY, Width_Renamed, Height_Renamed, OffX, OffY, 0)
		Else
			LayerTile(bdMapBottom, XMap, YMap, X + OffX, Y + OffY, Width_Renamed, Height_Renamed, OffX, OffY, 0)
			LayerTile(bdMapMiddle, XMap, YMap, X + OffX, Y + OffY, Width_Renamed, Height_Renamed, OffX, OffY, 0)
			LayerTile(bdMapTop, XMap, YMap, X + OffX, Y + OffY, Width_Renamed, Height_Renamed, OffX, OffY, 0)
			DrawTileObjects(XMap, YMap)
			DrawTileExitSign(XMap, YMap)
			LayerTile(bdMapDim, XMap, YMap, X + OffX, Y + OffY, Width_Renamed, Height_Renamed, OffX, OffY, 0)
		End If
	End Sub
	
	Private Sub DrawTileAnnotate(ByRef XMap As Short, ByRef YMap As Short)
		Dim cy, cx, n As Short
		Dim rc As Integer
		' Draw Party
		If Tome.MapX = XMap And Tome.MapY = YMap Then
			cx = ((YMap - MicroMapTop) + (XMap - MicroMapLeft)) * (bdTileWidth / 2) + 6
			cy = ((YMap - MicroMapTop) - (XMap - MicroMapLeft)) * (bdTileHeight / 3) - 12
			'UPGRADE_ISSUE: PictureBox property picMisc.hdc was not upgraded. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="CC4C7EC0-C903-48FC-ACCC-81861D12DA4A"'
			'UPGRADE_ISSUE: PictureBox property picMicroMap.hdc was not upgraded. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="CC4C7EC0-C903-48FC-ACCC-81861D12DA4A"'
			rc = BitBlt(picMicroMap.hdc, cx, cy, 15, 15, picMisc.hdc, 36, 18, SRCCOPY)
		End If
	End Sub
	
	'UPGRADE_NOTE: Map was upgraded to Map_Renamed. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="A9E4979A-37FA-4718-9994-97DD76ED70A7"'
	Private Sub DrawTileObjects(ByRef XMap As Short, ByRef YMap As Short)
		Dim Map_Renamed As Object
		Dim EncounterX As Encounter
		Dim CreatureX As Creature
		Dim ItemX As Item
		Dim cy, cx, n As Short
		Dim rc As Integer
		' Draw Encounter
		'UPGRADE_WARNING: Couldn't resolve default property of object Map.Hidden. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
		'UPGRADE_WARNING: Couldn't resolve default property of object Map.EncPointer. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
		If Map.EncPointer(XMap, YMap) > 0 And Map.Hidden(XMap, YMap) = False Then
			'UPGRADE_WARNING: Couldn't resolve default property of object Map.EncPointer. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
			'UPGRADE_WARNING: Couldn't resolve default property of object Map.Encounters. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
			EncounterX = Map.Encounters("E" & Map.EncPointer(XMap, YMap))
			If EncounterX.IsActive = True Then
				' Encounter Items
				For	Each ItemX In EncounterX.Items
					If ItemX.MapX = 0 And ItemX.MapY = 0 Then
						'UPGRADE_WARNING: Couldn't resolve default property of object Map. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
						PositionItem(Map, EncounterX, ItemX)
					End If
					If ItemX.MapX = XMap And ItemX.MapY = YMap And Len(ItemX.PictureFile) > 0 Then
						LoadItemPic(ItemX)
						'UPGRADE_WARNING: Couldn't resolve default property of object Map.Left. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
						'UPGRADE_WARNING: Couldn't resolve default property of object Map.Top. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
						cx = ((YMap - Map.Top) + (XMap - Map.Left)) * 48 + ItemX.TileSpotX
						'UPGRADE_WARNING: Couldn't resolve default property of object Map.Left. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
						'UPGRADE_WARNING: Couldn't resolve default property of object Map.Top. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
						cy = ((YMap - Map.Top) - (XMap - Map.Left)) * 24 - 24 + ItemX.TileSpotY - 24
						'UPGRADE_ISSUE: PictureBox property picItem.hdc was not upgraded. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="CC4C7EC0-C903-48FC-ACCC-81861D12DA4A"'
						'UPGRADE_ISSUE: PictureBox property picMap.hdc was not upgraded. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="CC4C7EC0-C903-48FC-ACCC-81861D12DA4A"'
						rc = BitBlt(picMap.hdc, cx, cy, ItemPicWidth(ItemX.Pic) / 3, ItemPicHeight(ItemX.Pic) / 3, picItem.hdc, 64 * ItemX.Pic - 32, 96 * 2, SRCAND)
						'UPGRADE_ISSUE: PictureBox property picItem.hdc was not upgraded. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="CC4C7EC0-C903-48FC-ACCC-81861D12DA4A"'
						'UPGRADE_ISSUE: PictureBox property picMap.hdc was not upgraded. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="CC4C7EC0-C903-48FC-ACCC-81861D12DA4A"'
						rc = BitBlt(picMap.hdc, cx, cy, ItemPicWidth(ItemX.Pic) / 3, ItemPicHeight(ItemX.Pic) / 3, picItem.hdc, 64 * ItemX.Pic - 64, 96 * 2, SRCPAINT)
					End If
				Next ItemX
				' Encounter Creatures
				For	Each CreatureX In EncounterX.Creatures
					If CreatureX.MapX = XMap And CreatureX.MapY = YMap And CreatureX.IsInanimate = False And Len(CreatureX.PictureFile) > 0 Then
						LoadCreaturePic(CreatureX)
						If CreatureX.HPNow > 0 Then
							n = CreatureX.Pic
						Else
							n = 0
						End If
						'UPGRADE_WARNING: Couldn't resolve default property of object Map.Left. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
						'UPGRADE_WARNING: Couldn't resolve default property of object Map.Top. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
						cx = ((YMap - Map.Top) + (XMap - Map.Left)) * 48 + CreatureX.TileSpotX
						'UPGRADE_WARNING: Couldn't resolve default property of object Map.Left. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
						'UPGRADE_WARNING: Couldn't resolve default property of object Map.Top. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
						cy = ((YMap - Map.Top) - (XMap - Map.Left)) * 24 - 24 + CreatureX.TileSpotY - (picCMap(n).Height / 2)
						'UPGRADE_ISSUE: PictureBox property picCMap.hdc was not upgraded. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="CC4C7EC0-C903-48FC-ACCC-81861D12DA4A"'
						'UPGRADE_ISSUE: PictureBox property picMap.hdc was not upgraded. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="CC4C7EC0-C903-48FC-ACCC-81861D12DA4A"'
						rc = BitBlt(picMap.hdc, cx, cy, picCMap(n).Width, picCMap(n).Height / 2, picCMap(n).hdc, 0, picCMap(n).Height / 2, SRCAND)
						'UPGRADE_ISSUE: PictureBox property picCMap.hdc was not upgraded. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="CC4C7EC0-C903-48FC-ACCC-81861D12DA4A"'
						'UPGRADE_ISSUE: PictureBox property picMap.hdc was not upgraded. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="CC4C7EC0-C903-48FC-ACCC-81861D12DA4A"'
						rc = BitBlt(picMap.hdc, cx, cy, picCMap(n).Width, picCMap(n).Height / 2, picCMap(n).hdc, 0, 0, SRCPAINT)
					End If
				Next CreatureX
			End If
		End If
		' Draw Party
		For	Each CreatureX In Tome.Creatures
			'        If CreatureX.MapX = XMap And CreatureX.MapY = YMap And CreatureX.IsInanimate = False Then
			' [Titi 2.4.9] added a check (creature position = party position?) to avoid "ghosts" creatures after MoveCreature or CopyCreature statements
			If CreatureX.MapX = XMap And CreatureX.MapY = YMap And CreatureX.MapX = Tome.MapX And CreatureX.MapY = Tome.MapY And CreatureX.IsInanimate = False Then
				If CreatureX.HPNow > 0 Then
					LoadCreaturePic(CreatureX)
					'UPGRADE_WARNING: Couldn't resolve default property of object Map.Left. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
					'UPGRADE_WARNING: Couldn't resolve default property of object Map.Top. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
					cx = ((YMap - Map.Top) + (XMap - Map.Left)) * 48 + CreatureX.TileSpotX
					'UPGRADE_WARNING: Couldn't resolve default property of object Map.Left. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
					'UPGRADE_WARNING: Couldn't resolve default property of object Map.Top. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
					cy = ((YMap - Map.Top) - (XMap - Map.Left)) * 24 - 24 + CreatureX.TileSpotY - (picCMap(CreatureX.Pic).Height / 2)
					'UPGRADE_ISSUE: PictureBox property picCMap.hdc was not upgraded. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="CC4C7EC0-C903-48FC-ACCC-81861D12DA4A"'
					'UPGRADE_ISSUE: PictureBox property picMap.hdc was not upgraded. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="CC4C7EC0-C903-48FC-ACCC-81861D12DA4A"'
					rc = BitBlt(picMap.hdc, cx, cy, picCMap(CreatureX.Pic).Width, picCMap(CreatureX.Pic).Height / 2, picCMap(CreatureX.Pic).hdc, 0, picCMap(CreatureX.Pic).Height / 2, SRCAND)
					'UPGRADE_ISSUE: PictureBox property picCMap.hdc was not upgraded. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="CC4C7EC0-C903-48FC-ACCC-81861D12DA4A"'
					'UPGRADE_ISSUE: PictureBox property picMap.hdc was not upgraded. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="CC4C7EC0-C903-48FC-ACCC-81861D12DA4A"'
					rc = BitBlt(picMap.hdc, cx, cy, picCMap(CreatureX.Pic).Width, picCMap(CreatureX.Pic).Height / 2, picCMap(CreatureX.Pic).hdc, 0, 0, SRCPAINT)
				End If
			End If
		Next CreatureX
	End Sub
	
	'UPGRADE_NOTE: Map was upgraded to Map_Renamed. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="A9E4979A-37FA-4718-9994-97DD76ED70A7"'
	Private Sub DrawTileExitSign(ByRef XMap As Short, ByRef YMap As Short)
		Dim Map_Renamed As Object
		Dim EntryPointX As EntryPoint
		Dim cx, cy As Short
		'UPGRADE_WARNING: Couldn't resolve default property of object Map.Hidden. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
		If Map.Hidden(XMap, YMap) = False Then
			For	Each EntryPointX In Map.EntryPoints
				If EntryPointX.MapX = XMap And EntryPointX.MapY = YMap And EntryPointX.IsNoExitSign = False Then
					'UPGRADE_WARNING: Couldn't resolve default property of object Map.Left. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
					'UPGRADE_WARNING: Couldn't resolve default property of object Map.Top. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
					cx = ((YMap - Map.Top) + (XMap - Map.Left)) * 48
					'UPGRADE_WARNING: Couldn't resolve default property of object Map.Left. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
					'UPGRADE_WARNING: Couldn't resolve default property of object Map.Top. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
					cy = ((YMap - Map.Top) - (XMap - Map.Left)) * 24
					If EntryPointX.AreaIndex = 0 Then
						ShowText(picMap, cx + 24, cy - 12, 48, 72, bdFontSmallWhite, "Main Exit", True, False)
					Else
						ShowText(picMap, cx, cy, 96, 72, bdFontSmallWhite, "Exit", True, False)
					End If
				End If
			Next EntryPointX
		End If
	End Sub
	
	Private Sub frmMain_Load(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles MyBase.Load
		Dim ctl As System.Windows.Forms.Control
		GameInit()
		' Determine if large fonts are used
		If Not SmallFonts Then
			picMainMenu.Visible = False
			DialogSetUp(modGameGeneral.DLGTYPE.bdDlgNoReply)
			DialogDM("Large fonts are not fully supported. Some windows may look strange unless you change your display settings.")
			'        GameEndAll ' no need to continue...
			picMainMenu.Visible = True
			For	Each ctl In Me.Controls
				' Adjust forms where large fonts cause display problems
				If ctl.Name = "picCreateName" Then
					ctl.Top = VB6.TwipsToPixelsY(VB6.PixelsToTwipsY(ctl.Top) / 1.25)
				End If
			Next ctl
		End If
		'    DoEvents
	End Sub
	
	Private Sub frmMain_MouseDown(ByVal eventSender As System.Object, ByVal eventArgs As System.Windows.Forms.MouseEventArgs) Handles MyBase.MouseDown
		Dim Button As Short = eventArgs.Button \ &H100000
		Dim Shift As Short = System.Windows.Forms.Control.ModifierKeys \ &H10000
		Dim X As Single = eventArgs.X
		Dim Y As Single = eventArgs.Y
		MenuAction(CShort(X), CShort(Y), True)
	End Sub
	
	Private Sub frmMain_MouseUp(ByVal eventSender As System.Object, ByVal eventArgs As System.Windows.Forms.MouseEventArgs) Handles MyBase.MouseUp
		Dim Button As Short = eventArgs.Button \ &H100000
		Dim Shift As Short = System.Windows.Forms.Control.ModifierKeys \ &H10000
		Dim X As Single = eventArgs.X
		Dim Y As Single = eventArgs.Y
		MenuAction(CShort(X), CShort(Y), False)
	End Sub
	
	Private Sub frmMain_MouseMove(ByVal eventSender As System.Object, ByVal eventArgs As System.Windows.Forms.MouseEventArgs) Handles MyBase.MouseMove
		Dim Button As Short = eventArgs.Button \ &H100000
		Dim Shift As Short = System.Windows.Forms.Control.ModifierKeys \ &H10000
		Dim X As Single = eventArgs.X
		Dim Y As Single = eventArgs.Y
		'UPGRADE_NOTE: Text was upgraded to Text_Renamed. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="A9E4979A-37FA-4718-9994-97DD76ED70A7"'
		Dim c As Short
		Dim blnShow As Boolean
		Dim Text_Renamed As String
		Dim sRune() As String
		Dim PauseTime, Start As Single
		' [Titi 2.4.9] no display when map doesn't have any runes!
		If Map.IsNoRunes = True Then Exit Sub
		' [Titi 2.4.8] get the runes names
		modBD.InitializeRunes((WorldNow.Name))
		ReDim sRune(intNbRunes + 1)
		Text_Renamed = VB.Right(strRunesList, Len(strRunesList) - 5) ' get rid of "List="
		For c = 1 To intNbRunes - 1
			sRune(c) = VB.Left(Text_Renamed, InStr(Text_Renamed, ",") - 1)
			Text_Renamed = VB.Right(Text_Renamed, Len(Text_Renamed) - Len(sRune(c)) - 1)
		Next 
		sRune(c) = Text_Renamed
		If IsBetween(X, Me.ClientRectangle.Width - 32, Me.ClientRectangle.Width) And IsBetween(Y, 16, 336) Then
			' hovering over second half of runes list
			c = Int((Y - 16) / 32 + 10)
			Y = (c - 10) * 32 + 16
			blnShow = True
		ElseIf IsBetween(X, 0, 32) And IsBetween(Y, 16, 336) Then 
			' hovering on first half of runes list
			c = Int((Y - 16) / 32)
			Y = c * 32 + 16
			blnShow = True
		End If
		X = 16
		If blnShow Then
			ShowText(picMap, X, Y, picMap.ClientRectangle.Width - 80, 14, bdFontNoxiousWhite, sRune(c + 1), IIf(c > 9, 1, 0), False)
			picMap.Refresh()
			PauseTime = 0.25 ' duration of display
			Start = VB.Timer()
			Do While VB.Timer() < Start + PauseTime
				'            DoEvents
			Loop 
			' clear message
			DrawMapRegion(0, 0, 340 + HintX, Len(sRune(c)) * 8 + 16 + HintY)
			DrawMapRegion(0, picMap.ClientRectangle.Width - Len(sRune(c)) * 8 - 16 - HintY, 340 + HintX, picMap.ClientRectangle.Width)
		End If
	End Sub
	
	'UPGRADE_WARNING: Event frmMain.Resize may fire when form is initialized. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="88B12AE1-6DE0-48A0-86F1-60C0686C026A"'
	Private Sub frmMain_Resize(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles MyBase.Resize
		ResizeForm()
	End Sub
	
	Private Sub fraMediaPlayerAVI_Click(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles fraMediaPlayerAVI.Click
		fraMediaPlayerAVI.Visible = False
		fraMediaPlayerAVI.Refresh()
		'MediaPlayerAVI.Stop
		CutSceneNow = False
	End Sub
	
	Private Sub lblHyperLink_Click(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles lblHyperLink.Click
		'Dim rc As Long
		'rc = HyperJump(bdOrderURL)
	End Sub
	
	'Private Sub MediaPlayerAVI_Click(Button As Integer, ShiftState As Integer, x As Single, y As Single)
	'    fraMediaPlayerAVI.Visible = False
	'    fraMediaPlayerAVI.Refresh
	'    MediaPlayerAVI.Stop
	'    CutSceneNow = False
	'End Sub
	'
	'Private Sub MediaPlayerAVI_EndOfStream(ByVal Result As Long)
	'    fraMediaPlayerAVI.Visible = False
	'    MediaPlayerAVI.Stop
	'    CutSceneNow = False
	'End Sub
	
	Private Sub picBuySell_MouseDown(ByVal eventSender As System.Object, ByVal eventArgs As System.Windows.Forms.MouseEventArgs) Handles picBuySell.MouseDown
		Dim Button As Short = eventArgs.Button \ &H100000
		Dim Shift As Short = System.Windows.Forms.Control.ModifierKeys \ &H10000
		Dim X As Single = eventArgs.X
		Dim Y As Single = eventArgs.Y
		If picConvo.Visible = False Then
			ClickX = X : ClickY = Y
			DialogBuySellClick(ClickX, ClickY, True)
		End If
	End Sub
	
	Private Sub picBuySell_MouseUp(ByVal eventSender As System.Object, ByVal eventArgs As System.Windows.Forms.MouseEventArgs) Handles picBuySell.MouseUp
		Dim Button As Short = eventArgs.Button \ &H100000
		Dim Shift As Short = System.Windows.Forms.Control.ModifierKeys \ &H10000
		Dim X As Single = eventArgs.X
		Dim Y As Single = eventArgs.Y
		If picConvo.Visible = False Then
			DialogBuySellClick(ClickX, ClickY, False)
		End If
	End Sub
	
	Private Sub picContainer_MouseDown(ByVal eventSender As System.Object, ByVal eventArgs As System.Windows.Forms.MouseEventArgs) Handles picContainer.MouseDown
		Dim Button As Short = eventArgs.Button \ &H100000
		Dim Shift As Short = System.Windows.Forms.Control.ModifierKeys \ &H10000
		Dim X As Single = eventArgs.X
		Dim Y As Single = eventArgs.Y
		If picConvo.Visible = False Then
			ClickX = X : ClickY = Y
			InventoryContainerClick(ClickX, ClickY, True, Button)
		End If
	End Sub
	
	Private Sub picContainer_MouseMove(ByVal eventSender As System.Object, ByVal eventArgs As System.Windows.Forms.MouseEventArgs) Handles picContainer.MouseMove
		Dim Button As Short = eventArgs.Button \ &H100000
		Dim Shift As Short = System.Windows.Forms.Control.ModifierKeys \ &H10000
		Dim X As Single = eventArgs.X
		Dim Y As Single = eventArgs.Y
		If InvDragMode = True Then
			If picInvDrag.Visible = False Then
				InvDragItem.Selected = 1
				InventoryContainerShow(InvContainer)
				picInvDrag.BringToFront()
				picInvDrag.Visible = True
			End If
			picInvDrag.Top = picContainer.Top + Y - JournalY
			picInvDrag.Left = picContainer.Left + X - JournalX
		Else
			InventoryHint(2, Int(X), Int(Y))
		End If
	End Sub
	
	Private Sub picContainer_MouseUp(ByVal eventSender As System.Object, ByVal eventArgs As System.Windows.Forms.MouseEventArgs) Handles picContainer.MouseUp
		Dim Button As Short = eventArgs.Button \ &H100000
		Dim Shift As Short = System.Windows.Forms.Control.ModifierKeys \ &H10000
		Dim X As Single = eventArgs.X
		Dim Y As Single = eventArgs.Y
		Dim FromList, c, ToList As Short
		If picConvo.Visible = False Then
			If InvDragMode = True Then
				InventoryContainerDrag(picContainer.Left + CShort(X), picContainer.Top + CShort(Y))
				InventoryContainerShow(InvContainer)
			Else
				InventoryContainerClick(ClickX, ClickY, False, Button)
			End If
		End If
	End Sub
	
	'UPGRADE_ISSUE: PictureBox event picConvo.KeyDown was not upgraded. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="ABD9AF39-7E24-4AFF-AD8D-3675C1AA3054"'
	Private Sub picConvo_KeyDown(ByRef KeyCode As Short, ByRef Shift As Short)
		'    KeyBoardShortCuts KeyCode, Shift
		ParseShortCuts(KeyCode, Shift)
	End Sub
	
	Private Sub picConvo_MouseDown(ByVal eventSender As System.Object, ByVal eventArgs As System.Windows.Forms.MouseEventArgs) Handles picConvo.MouseDown
		Dim Button As Short = eventArgs.Button \ &H100000
		Dim Shift As Short = System.Windows.Forms.Control.ModifierKeys \ &H10000
		Dim X As Single = eventArgs.X
		Dim Y As Single = eventArgs.Y
		ClickX = X : ClickY = Y
		DialogClick(ClickX, ClickY, True)
	End Sub
	
	Private Sub picConvo_MouseMove(ByVal eventSender As System.Object, ByVal eventArgs As System.Windows.Forms.MouseEventArgs) Handles picConvo.MouseMove
		Dim Button As Short = eventArgs.Button \ &H100000
		Dim Shift As Short = System.Windows.Forms.Control.ModifierKeys \ &H10000
		Dim X As Single = eventArgs.X
		Dim Y As Single = eventArgs.Y
		If modDialog.DialogStyle = modGameGeneral.DLGTYPE.bdDlgWithReply Then
			DialogReplyMove(CShort(X), CShort(Y))
		End If
	End Sub
	
	Private Sub picConvo_MouseUp(ByVal eventSender As System.Object, ByVal eventArgs As System.Windows.Forms.MouseEventArgs) Handles picConvo.MouseUp
		Dim Button As Short = eventArgs.Button \ &H100000
		Dim Shift As Short = System.Windows.Forms.Control.ModifierKeys \ &H10000
		Dim X As Single = eventArgs.X
		Dim Y As Single = eventArgs.Y
		DialogClick(ClickX, ClickY, False)
	End Sub
	
	'UPGRADE_ISSUE: PictureBox event picConvoEnter.KeyDown was not upgraded. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="ABD9AF39-7E24-4AFF-AD8D-3675C1AA3054"'
	Private Sub picConvoEnter_KeyDown(ByRef KeyCode As Short, ByRef Shift As Short)
		' [Titi 2.4.9] get CTRL, SHIFT and arrows for shortcuts
		Dim oldDialogText As String
		oldDialogText = DialogText
		Select Case KeyCode
			Case System.Windows.Forms.Keys.Left, System.Windows.Forms.Keys.Right, System.Windows.Forms.Keys.Up, System.Windows.Forms.Keys.Down
				If TomeAction = bdTomeOptions Then
					' options shown, then we're defining shortcuts
					DialogText = IIf(KeyCode = System.Windows.Forms.Keys.Left, "left", IIf(KeyCode = System.Windows.Forms.Keys.Right, "right", IIf(KeyCode = System.Windows.Forms.Keys.Up, "up", "down")))
					If VB.Left(oldDialogText & Space(4), 4) = "shif" Or VB.Left(oldDialogText & Space(4), 4) = "ctrl" Then DialogText = Trim(VB.Left(oldDialogText & Space(5), 5)) & DialogText
				End If
			Case System.Windows.Forms.Keys.ControlKey, System.Windows.Forms.Keys.ShiftKey
				If TomeAction = bdTomeOptions Then
					' options shown, then we're defining shortcuts
					DialogText = IIf(KeyCode = System.Windows.Forms.Keys.ShiftKey, "shift", "ctrl")
					If Shift = 3 Then DialogText = "ctrl" ' if both shift and control pressed, default to ctrl
				End If
		End Select
		DialogEnter(DialogText)
	End Sub
	
	'UPGRADE_ISSUE: PictureBox event picConvoEnter.KeyPress was not upgraded. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="ABD9AF39-7E24-4AFF-AD8D-3675C1AA3054"'
	Private Sub picConvoEnter_KeyPress(ByRef KeyAscii As Short)
		Select Case KeyAscii
			Case System.Windows.Forms.Keys.Back ', vbKeyLeft   ' BackSpace
				If Len(DialogText) > 0 Then
					DialogText = VB.Left(DialogText, Len(DialogText) - 1)
				End If
			Case 32 ' Space
				DialogText = DialogText & " "
			Case 47 To 57, 65 To 90, 97 To 122 ' [Titi 2.4.9] allowed "/" (ascii code # 47)
				If VB.Left(DialogText & Space(5), 5) = "shift" And KeyAscii > 96 Then KeyAscii = KeyAscii - 32
				DialogText = DialogText & Chr(KeyAscii)
		End Select
		DialogEnter(DialogText)
	End Sub
	
	Private Sub picConvoList_MouseDown(ByVal eventSender As System.Object, ByVal eventArgs As System.Windows.Forms.MouseEventArgs) Handles picConvoList.MouseDown
		Dim Button As Short = eventArgs.Button \ &H100000
		Dim Shift As Short = System.Windows.Forms.Control.ModifierKeys \ &H10000
		Dim X As Single = eventArgs.X
		Dim Y As Single = eventArgs.Y
		ClickX = X : ClickY = Y
		If modDialog.DialogStyle = modGameGeneral.DLGTYPE.bdDlgItemList Then
			TargetItemClick(ClickX, ClickY, True)
		ElseIf modDialog.DialogStyle = modGameGeneral.DLGTYPE.bdDlgCreatureList Then 
			TargetCreatureClick(ClickX, ClickY, True)
		ElseIf modDialog.DialogStyle = modGameGeneral.DLGTYPE.bdDlgDebug Then 
			DebugClick(ClickX, ClickY, True)
		End If
	End Sub
	
	Private Sub picConvoList_MouseUp(ByVal eventSender As System.Object, ByVal eventArgs As System.Windows.Forms.MouseEventArgs) Handles picConvoList.MouseUp
		Dim Button As Short = eventArgs.Button \ &H100000
		Dim Shift As Short = System.Windows.Forms.Control.ModifierKeys \ &H10000
		Dim X As Single = eventArgs.X
		Dim Y As Single = eventArgs.Y
		If modDialog.DialogStyle = modGameGeneral.DLGTYPE.bdDlgItemList Then
			TargetItemClick(ClickX, ClickY, False)
		ElseIf modDialog.DialogStyle = modGameGeneral.DLGTYPE.bdDlgCreatureList Then 
			TargetCreatureClick(ClickX, ClickY, False)
		ElseIf modDialog.DialogStyle = modGameGeneral.DLGTYPE.bdDlgDebug Then 
			DebugClick(ClickX, ClickY, False)
		End If
	End Sub
	
	Private Sub picDialogBrief_Click(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles picDialogBrief.Click
		picDialogBrief.Visible = False
	End Sub
	
	Private Sub picInventory_MouseDown(ByVal eventSender As System.Object, ByVal eventArgs As System.Windows.Forms.MouseEventArgs) Handles picInventory.MouseDown
		Dim Button As Short = eventArgs.Button \ &H100000
		Dim Shift As Short = System.Windows.Forms.Control.ModifierKeys \ &H10000
		Dim X As Single = eventArgs.X
		Dim Y As Single = eventArgs.Y
		If picConvo.Visible = False Then
			ClickX = X : ClickY = Y
			InventoryClick(ClickX, ClickY, True, Button)
		End If
	End Sub
	
	Private Sub picInventory_MouseMove(ByVal eventSender As System.Object, ByVal eventArgs As System.Windows.Forms.MouseEventArgs) Handles picInventory.MouseMove
		Dim Button As Short = eventArgs.Button \ &H100000
		Dim Shift As Short = System.Windows.Forms.Control.ModifierKeys \ &H10000
		Dim X As Single = eventArgs.X
		Dim Y As Single = eventArgs.Y
		If InvDragMode = True Then
			If picInvDrag.Visible = False Then
				picInvDrag.BringToFront()
				picInvDrag.Visible = True
				InvDragItem.Selected = 1
				InventoryShow(bdInvItems)
			End If
			picInvDrag.Top = picInventory.Top + Y - JournalY
			picInvDrag.Left = picInventory.Left + X - JournalX
		ElseIf InvNowShow = bdInvItems Then 
			InventoryHint(0, Int(X), Int(Y))
		End If
	End Sub
	
	Private Sub picInventory_MouseUp(ByVal eventSender As System.Object, ByVal eventArgs As System.Windows.Forms.MouseEventArgs) Handles picInventory.MouseUp
		Dim Button As Short = eventArgs.Button \ &H100000
		Dim Shift As Short = System.Windows.Forms.Control.ModifierKeys \ &H10000
		Dim X As Single = eventArgs.X
		Dim Y As Single = eventArgs.Y
		If picConvo.Visible = False Then
			If InvDragMode = True Then
				InventoryDragItem(picInventory.Left + CShort(X) - (JournalX / 2), picInventory.Top + CShort(Y) - (JournalY / 2))
				InventoryShow(bdInvItems)
			Else
				InventoryClick(ClickX, ClickY, False, Button)
			End If
		End If
	End Sub
	
	'UPGRADE_ISSUE: PictureBox event picJournal.KeyPress was not upgraded. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="ABD9AF39-7E24-4AFF-AD8D-3675C1AA3054"'
	Private Sub picJournal_KeyPress(ByRef KeyAscii As Short)
		Call AddJournalText(KeyAscii)
	End Sub
	
	Private Sub picJournal_MouseDown(ByVal eventSender As System.Object, ByVal eventArgs As System.Windows.Forms.MouseEventArgs) Handles picJournal.MouseDown
		Dim Button As Short = eventArgs.Button \ &H100000
		Dim Shift As Short = System.Windows.Forms.Control.ModifierKeys \ &H10000
		Dim X As Single = eventArgs.X
		Dim Y As Single = eventArgs.Y
		ClickX = X : ClickY = Y
		If Y < 24 Then
			JournalX = X : JournalY = Y
			JournalDragDrop = True
		Else
			JournalClick(ClickX, ClickY, True)
		End If
	End Sub
	
	Private Sub picJournal_MouseMove(ByVal eventSender As System.Object, ByVal eventArgs As System.Windows.Forms.MouseEventArgs) Handles picJournal.MouseMove
		Dim Button As Short = eventArgs.Button \ &H100000
		Dim Shift As Short = System.Windows.Forms.Control.ModifierKeys \ &H10000
		Dim X As Single = eventArgs.X
		Dim Y As Single = eventArgs.Y
		If JournalDragDrop = True Then
			picJournal.Top = picJournal.Top + Y - JournalY : picJournal.Left = picJournal.Left + X - JournalX
		End If
	End Sub
	
	Private Sub picJournal_MouseUp(ByVal eventSender As System.Object, ByVal eventArgs As System.Windows.Forms.MouseEventArgs) Handles picJournal.MouseUp
		Dim Button As Short = eventArgs.Button \ &H100000
		Dim Shift As Short = System.Windows.Forms.Control.ModifierKeys \ &H10000
		Dim X As Single = eventArgs.X
		Dim Y As Single = eventArgs.Y
		JournalDragDrop = False
		JournalClick(ClickX, ClickY, False)
	End Sub
	
	Private Sub picMainMenu_Click(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles picMainMenu.Click
		TomeDo(TomeAction)
	End Sub
	
	Private Sub picMainMenu_MouseMove(ByVal eventSender As System.Object, ByVal eventArgs As System.Windows.Forms.MouseEventArgs) Handles picMainMenu.MouseMove
		Dim Button As Short = eventArgs.Button \ &H100000
		Dim Shift As Short = System.Windows.Forms.Control.ModifierKeys \ &H10000
		Dim X As Single = eventArgs.X
		Dim Y As Single = eventArgs.Y
		Dim rc As Integer
		Dim c As Short
		If IsBetween(Y, 98, 289) Then
			c = Int((Y - 98) / 39)
			If TomeAction <> c Then
				'UPGRADE_ISSUE: PictureBox method picMainMenu.Cls was not upgraded. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="CC4C7EC0-C903-48FC-ACCC-81861D12DA4A"'
				picMainMenu.Cls()
				If TomeMenu = bdTomeOptions Then
					'UPGRADE_ISSUE: PictureBox property picBlack.hdc was not upgraded. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="CC4C7EC0-C903-48FC-ACCC-81861D12DA4A"'
					'UPGRADE_ISSUE: PictureBox property picMainMenu.hdc was not upgraded. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="CC4C7EC0-C903-48FC-ACCC-81861D12DA4A"'
					rc = BitBlt(picMainMenu.hdc, 122, 103, 288, 195, picBlack.hdc, 0, 340, SRCCOPY)
					'UPGRADE_ISSUE: PictureBox property picBlack.hdc was not upgraded. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="CC4C7EC0-C903-48FC-ACCC-81861D12DA4A"'
					'UPGRADE_ISSUE: PictureBox property picMainMenu.hdc was not upgraded. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="CC4C7EC0-C903-48FC-ACCC-81861D12DA4A"'
					rc = BitBlt(picMainMenu.hdc, 122, 103 + c * 39, 288, 39, picBlack.hdc, 0, 536 + c * 39, SRCCOPY)
				Else
					'UPGRADE_ISSUE: PictureBox property picBlack.hdc was not upgraded. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="CC4C7EC0-C903-48FC-ACCC-81861D12DA4A"'
					'UPGRADE_ISSUE: PictureBox property picMainMenu.hdc was not upgraded. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="CC4C7EC0-C903-48FC-ACCC-81861D12DA4A"'
					rc = BitBlt(picMainMenu.hdc, 122, 103 + c * 39, 288, 39, picBlack.hdc, 0, 144 + c * 39, SRCCOPY)
				End If
				Call UpdateMainMenu()
				TomeAction = c
				picMainMenu.Refresh()
			End If
		End If
	End Sub
	
	Private Sub UpdateMainMenu()
		ShowText(picMainMenu, 1, picMainMenu.ClientRectangle.Height - 15, picMainMenu.ClientRectangle.Width, 14, bdFontNoxiousGold, "Current World: " & WorldNow.Name, True, False)
		ShowText(picMainMenu, 12, picMainMenu.ClientRectangle.Height - 25, picMainMenu.ClientRectangle.Width - 24, 14, bdFontSmallWhite, "CrossCut Games, Inc.", False, False)
		ShowText(picMainMenu, 12, picMainMenu.ClientRectangle.Height - 25, picMainMenu.ClientRectangle.Width - 24, 14, bdFontSmallWhite, "v" & My.Application.Info.Version.Major & "." & My.Application.Info.Version.Minor & " Build " & My.Application.Info.Version.Revision, 1, False)
	End Sub
	
	Private Sub picMap_Click(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles picMap.Click
		If Not Frozen And picGrid.Visible = False Then
			If BtnClick = VB6.MouseButtonConstants.LeftButton Then
				MovePartySet(ClickX, ClickY)
			End If
		ElseIf SplashNow = True Then 
			KillSplash = True
		ElseIf CutSceneNow = True Then 
			CutSceneNow = False
		End If
	End Sub
	
	'UPGRADE_ISSUE: PictureBox event picMap.KeyDown was not upgraded. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="ABD9AF39-7E24-4AFF-AD8D-3675C1AA3054"'
	Private Sub picMap_KeyDown(ByRef KeyCode As Short, ByRef Shift As Short)
		If Not Frozen And picGrid.Visible = False Then
			' KeyBoardShortCuts KeyCode, Shift
			ParseShortCuts(KeyCode, Shift)
		End If
	End Sub
	
	Private Sub picMap_MouseDown(ByVal eventSender As System.Object, ByVal eventArgs As System.Windows.Forms.MouseEventArgs) Handles picMap.MouseDown
		Dim Button As Short = eventArgs.Button \ &H100000
		Dim Shift As Short = System.Windows.Forms.Control.ModifierKeys \ &H10000
		Dim X As Single = eventArgs.X
		Dim Y As Single = eventArgs.Y
		If Not Frozen And picGrid.Visible = False Then
			ClickX = CShort(X)
			ClickY = CShort(Y)
			BtnClick = Button
			If BtnClick = VB6.MouseButtonConstants.RightButton Then
				MenuSkillUse(False)
			End If
		End If
	End Sub
	
	Private Sub picMap_MouseMove(ByVal eventSender As System.Object, ByVal eventArgs As System.Windows.Forms.MouseEventArgs) Handles picMap.MouseMove
		Dim Button As Short = eventArgs.Button \ &H100000
		Dim Shift As Short = System.Windows.Forms.Control.ModifierKeys \ &H10000
		Dim X As Single = eventArgs.X
		Dim Y As Single = eventArgs.Y
		If Not Frozen And picGrid.Visible = False Then
			HintX = CShort(X) : HintY = CShort(Y)
		End If
	End Sub
	
	Private Sub SetTurn(ByRef Index As Short)
		' Need to scroll if more than 5 away from far left CreatureWithTurn
		MenuPartyIndex = Index
		MenuDrawParty()
		CreatureWithTurn = MenuParty(MenuPartyIndex)
		MenuDrawParty()
		' If Inventory is open, show new CreatureWithTurn's inventory
		If picInventory.Visible = True Then
			InvTopIndex = 0
			InventoryShow(InvNowShow)
		ElseIf picBuySell.Visible = True Then 
			DialogBuySellPC()
			DialogBuySellShow()
		ElseIf picTomeNew.Visible = True And TomeAction = bdMenuSkill Then 
			'        MenuSkillPostRaise
			' [Titi 2.4.9] this was causing the "Raise Skill" bug
			MenuSkillLoad()
			MenuSkillShow(1)
		End If
	End Sub
	
	'UPGRADE_NOTE: Tome was upgraded to Tome_Renamed. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="A9E4979A-37FA-4718-9994-97DD76ED70A7"'
	Private Sub picMenu_DoubleClick(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles picMenu.DoubleClick
		Dim Tome_Renamed As Object
		'UPGRADE_WARNING: Couldn't resolve default property of object Tome.Creatures. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
		If ClickX > 32 And Least(CShort((ClickX - 32) / 122), 4) < Tome.Creatures.Count Then
			'UPGRADE_WARNING: Couldn't resolve default property of object Tome.MapY. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
			'UPGRADE_WARNING: Couldn't resolve default property of object Tome.MapX. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
			MapCenter(Tome.MapX, Tome.MapY)
			DrawMapAll()
		End If
	End Sub
	
	'UPGRADE_ISSUE: PictureBox event picMenu.KeyDown was not upgraded. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="ABD9AF39-7E24-4AFF-AD8D-3675C1AA3054"'
	Private Sub picMenu_KeyDown(ByRef KeyCode As Short, ByRef Shift As Short)
		If Not Frozen Or picInventory.Visible = True Or picBuySell.Visible = True Or picTalk.Visible = True Then
			'UPGRADE_ISSUE: PictureBox event picMap.KeyDown was not upgraded. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="ABD9AF39-7E24-4AFF-AD8D-3675C1AA3054"'
			picMap_KeyDown(KeyCode, Shift)
		End If
	End Sub
	
	Private Sub picMenu_MouseDown(ByVal eventSender As System.Object, ByVal eventArgs As System.Windows.Forms.MouseEventArgs) Handles picMenu.MouseDown
		Dim Button As Short = eventArgs.Button \ &H100000
		Dim Shift As Short = System.Windows.Forms.Control.ModifierKeys \ &H10000
		Dim X As Single = eventArgs.X
		Dim Y As Single = eventArgs.Y
		ClickX = CShort(X)
		ClickY = CShort(Y)
		' [Titi 2.4.9] moved from SetTurn to fix the "Gained new Spells" bug
		If (picTomeNew.Visible = True And TomeAction = bdMenuSkill) Then MenuSkillPostRaise()
		' Scroll Party Members (if more than 5)
		If PartyLeft > 0 And ClickX < 32 Then
			Call PlayClickSnd(modIOFunc.ClickType.ifClick)
			PartyLeft = PartyLeft - 1
			SetTurn(Least(MenuPartyIndex + 1, 4))
		ElseIf PartyLeft + 5 < Tome.Creatures.Count() And ClickX > 610 Then 
			Call PlayClickSnd(modIOFunc.ClickType.ifClick)
			PartyLeft = PartyLeft + 1
			SetTurn(Greatest(MenuPartyIndex - 1, 0))
		ElseIf Not Frozen Or picInventory.Visible = True Or picBuySell.Visible = True Or picTalk.Visible = True Or picSearch.Visible = True Or (picTomeNew.Visible = True And TomeAction = bdMenuSkill) Then 
			If Least(CShort((ClickX - 32) / 122), 4) < Tome.Creatures.Count() And ClickX > 32 And ClickX < 610 Then
				If Button = VB6.MouseButtonConstants.RightButton Then
					CreatureTarget = MenuParty(Least(CShort((ClickX - 32) / 122), 4))
					MenuSkillUse(True)
				Else
					SetTurn(Least(CShort((ClickX - 32) / 122), 4))
					Call PlayClickSnd(modIOFunc.ClickType.ifClick)
				End If
			End If
		End If
	End Sub
	
	Private Sub picMicroMap_MouseDown(ByVal eventSender As System.Object, ByVal eventArgs As System.Windows.Forms.MouseEventArgs) Handles picMicroMap.MouseDown
		Dim Button As Short = eventArgs.Button \ &H100000
		Dim Shift As Short = System.Windows.Forms.Control.ModifierKeys \ &H10000
		Dim X As Single = VB6.FromPixelsUserX(eventArgs.X, 0, 309, 309)
		Dim Y As Single = VB6.FromPixelsUserY(eventArgs.Y, 0, 225, 604)
		JournalDragDrop = True
		JournalX = X : JournalY = Y
		'UPGRADE_WARNING: PictureBox property picMicroMap.MousePointer has a new behavior. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6BA9B8D2-2A32-4B6E-8D36-44949974A5B4"'
		picMicroMap.Cursor = System.Windows.Forms.Cursors.SizeAll
	End Sub
	
	Private Sub picMicroMap_MouseMove(ByVal eventSender As System.Object, ByVal eventArgs As System.Windows.Forms.MouseEventArgs) Handles picMicroMap.MouseMove
		Dim Button As Short = eventArgs.Button \ &H100000
		Dim Shift As Short = System.Windows.Forms.Control.ModifierKeys \ &H10000
		Dim X As Single = VB6.FromPixelsUserX(eventArgs.X, 0, 309, 309)
		Dim Y As Single = VB6.FromPixelsUserY(eventArgs.Y, 0, 225, 604)
		If JournalDragDrop = True Then
			picMicroMap.Top = picMicroMap.Top + Y - JournalY : picMicroMap.Left = picMicroMap.Left + X - JournalX
		End If
	End Sub
	
	Private Sub picMicroMap_MouseUp(ByVal eventSender As System.Object, ByVal eventArgs As System.Windows.Forms.MouseEventArgs) Handles picMicroMap.MouseUp
		Dim Button As Short = eventArgs.Button \ &H100000
		Dim Shift As Short = System.Windows.Forms.Control.ModifierKeys \ &H10000
		Dim X As Single = VB6.FromPixelsUserX(eventArgs.X, 0, 309, 309)
		Dim Y As Single = VB6.FromPixelsUserY(eventArgs.Y, 0, 225, 604)
		JournalDragDrop = False
		JournalMoveMap()
		picMicroMap.Cursor = System.Windows.Forms.Cursors.Default
	End Sub
	
	Private Sub picSearch_MouseDown(ByVal eventSender As System.Object, ByVal eventArgs As System.Windows.Forms.MouseEventArgs) Handles picSearch.MouseDown
		Dim Button As Short = eventArgs.Button \ &H100000
		Dim Shift As Short = System.Windows.Forms.Control.ModifierKeys \ &H10000
		Dim X As Single = eventArgs.X
		Dim Y As Single = eventArgs.Y
		If picConvo.Visible = False Then
			ClickX = X : ClickY = Y
			SearchClick(ClickX, ClickY, True, Button)
		End If
	End Sub
	
	Private Sub picSearch_MouseMove(ByVal eventSender As System.Object, ByVal eventArgs As System.Windows.Forms.MouseEventArgs) Handles picSearch.MouseMove
		Dim Button As Short = eventArgs.Button \ &H100000
		Dim Shift As Short = System.Windows.Forms.Control.ModifierKeys \ &H10000
		Dim X As Single = eventArgs.X
		Dim Y As Single = eventArgs.Y
		If InvDragMode = True Then
			If picInvDrag.Visible = False Then
				InvDragItem.Selected = 1
				SearchShow()
				picInvDrag.BringToFront()
				picInvDrag.Visible = True
			End If
			picInvDrag.Top = picSearch.Top + Y - JournalY
			picInvDrag.Left = picSearch.Left + X - JournalX
		Else
			InventoryHint(1, Int(X), Int(Y))
		End If
	End Sub
	
	Private Sub picSearch_MouseUp(ByVal eventSender As System.Object, ByVal eventArgs As System.Windows.Forms.MouseEventArgs) Handles picSearch.MouseUp
		Dim Button As Short = eventArgs.Button \ &H100000
		Dim Shift As Short = System.Windows.Forms.Control.ModifierKeys \ &H10000
		Dim X As Single = eventArgs.X
		Dim Y As Single = eventArgs.Y
		Dim FromList, c, ToList As Short
		If picConvo.Visible = False Then
			If InvDragMode = True Then
				SearchDrag(picSearch.Left + CShort(X), picSearch.Top + CShort(Y))
				If EncounterNow.Items.Count() < 1 Then
					SearchClose()
				Else
					SearchShow()
				End If
			Else
				SearchClick(ClickX, ClickY, False, Button)
			End If
		End If
	End Sub
	
	Private Sub picTalk_MouseDown(ByVal eventSender As System.Object, ByVal eventArgs As System.Windows.Forms.MouseEventArgs) Handles picTalk.MouseDown
		Dim Button As Short = eventArgs.Button \ &H100000
		Dim Shift As Short = System.Windows.Forms.Control.ModifierKeys \ &H10000
		Dim X As Single = eventArgs.X
		Dim Y As Single = eventArgs.Y
		ClickX = X : ClickY = Y
		TalkClick(ClickX, ClickY, True)
	End Sub
	
	Private Sub picTalk_MouseMove(ByVal eventSender As System.Object, ByVal eventArgs As System.Windows.Forms.MouseEventArgs) Handles picTalk.MouseMove
		Dim Button As Short = eventArgs.Button \ &H100000
		Dim Shift As Short = System.Windows.Forms.Control.ModifierKeys \ &H10000
		Dim X As Single = eventArgs.X
		Dim Y As Single = eventArgs.Y
		TalkMove(CShort(X), CShort(Y))
	End Sub
	
	Private Sub picTalk_MouseUp(ByVal eventSender As System.Object, ByVal eventArgs As System.Windows.Forms.MouseEventArgs) Handles picTalk.MouseUp
		Dim Button As Short = eventArgs.Button \ &H100000
		Dim Shift As Short = System.Windows.Forms.Control.ModifierKeys \ &H10000
		Dim X As Single = eventArgs.X
		Dim Y As Single = eventArgs.Y
		TalkClick(ClickX, ClickY, False)
	End Sub
	
	'UPGRADE_ISSUE: PictureBox event picCreateName.KeyPress was not upgraded. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="ABD9AF39-7E24-4AFF-AD8D-3675C1AA3054"'
	Private Sub picCreateName_KeyPress(ByRef KeyAscii As Short)
		CreatePCNameEnter(KeyAscii)
	End Sub
	
	Private Sub picTomeNew_MouseDown(ByVal eventSender As System.Object, ByVal eventArgs As System.Windows.Forms.MouseEventArgs) Handles picTomeNew.MouseDown
		Dim Button As Short = eventArgs.Button \ &H100000
		Dim Shift As Short = System.Windows.Forms.Control.ModifierKeys \ &H10000
		Dim X As Single = eventArgs.X
		Dim Y As Single = eventArgs.Y
		ClickX = X : ClickY = Y
		Select Case TomeAction
			Case bdTomeNew, bdTomeGather
				TomeNewClick(ClickX, ClickY, True, (Button = VB6.MouseButtonConstants.RightButton)) ' [Titi 2.4.9] added right-click
			Case bdTomeSaves, bdTomeSaveAs
				TomeSavesClick(ClickX, ClickY, True)
			Case bdCreatePCKingdom, bdCreatePCPicture, bdCreatePCSkills, bdCreatePCName
				CreatePCClick(ClickX, ClickY, True)
			Case bdTomeOptions
				OptionsClick(ClickX, ClickY, True)
			Case bdMenuSkill
				MenuSkillClick(ClickX, ClickY, True)
		End Select
	End Sub
	
	Private Sub picTomeNew_MouseMove(ByVal eventSender As System.Object, ByVal eventArgs As System.Windows.Forms.MouseEventArgs) Handles picTomeNew.MouseMove
		Dim Button As Short = eventArgs.Button \ &H100000
		Dim Shift As Short = System.Windows.Forms.Control.ModifierKeys \ &H10000
		Dim X As Single = eventArgs.X
		Dim Y As Single = eventArgs.Y
		ClickX = X : ClickY = Y
		Select Case TomeAction
			Case bdTomeNew, bdTomeGather
				If TomeRosterDrag <> 0 Then
					picInvDrag.Visible = True
					picInvDrag.Top = picTomeNew.Top - TomeRosterDragY + Y
					picInvDrag.Left = picTomeNew.Left - TomeRosterDragX + X
				End If
		End Select
	End Sub
	
	Private Sub picTomeNew_MouseUp(ByVal eventSender As System.Object, ByVal eventArgs As System.Windows.Forms.MouseEventArgs) Handles picTomeNew.MouseUp
		Dim Button As Short = eventArgs.Button \ &H100000
		Dim Shift As Short = System.Windows.Forms.Control.ModifierKeys \ &H10000
		Dim X As Single = eventArgs.X
		Dim Y As Single = eventArgs.Y
		Select Case TomeAction
			Case bdTomeNew, bdTomeGather
				TomeNewClick(ClickX, ClickY, False)
			Case bdTomeSaves, bdTomeSaveAs
				TomeSavesClick(ClickX, ClickY, False)
			Case bdCreatePCKingdom, bdCreatePCPicture, bdCreatePCSkills, bdCreatePCName
				CreatePCClick(ClickX, ClickY, False)
			Case bdTomeOptions
				OptionsClick(ClickX, ClickY, False)
			Case bdMenuSkill
				MenuSkillClick(ClickX, ClickY, False)
		End Select
	End Sub
	
	Private Sub picWallPaper_Click(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles picWallPaper.Click
		If CutSceneNow = True Then
			CutSceneNow = False
		End If
	End Sub
	
	Private Sub timerMusicLoop_Tick(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles timerMusicLoop.Tick
		' check status of background music
		If GlobalMusicState = 1 Then
			' see if the music is still playing
			If oGameMusic.Status = IMCI.VIDEOSTATE.vsSTOPPED Then
				oGameMusic.Play()
			End If
		End If
	End Sub
	
	Private Sub tmrDialogBrief_Tick(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles tmrDialogBrief.Tick
		Static TimeToRead As Double
		If DialogBriefSet.Count() > 0 And TimeToRead = 0 Then
			'UPGRADE_ISSUE: PictureBox method picDialogBrief.Cls was not upgraded. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="CC4C7EC0-C903-48FC-ACCC-81861D12DA4A"'
			picDialogBrief.Cls()
			'UPGRADE_WARNING: Couldn't resolve default property of object DialogBriefSet().CreatureTalking. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
			DialogShowFaceBrief(DialogBriefSet.Item(1).CreatureTalking)
			'UPGRADE_WARNING: Couldn't resolve default property of object DialogBriefSet().Text. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
			ShowText(picDialogBrief, 90, 37, picDialogBrief.ClientRectangle.Width - 90, picDialogBrief.ClientRectangle.Height, bdFontNoxiousWhite, DialogBriefSet.Item(1).Text, False, False)
			' Position DialogBrief box
			picDialogBrief.Top = 0
			'UPGRADE_WARNING: Couldn't resolve default property of object DialogBriefSet(1).CreatureTalking. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
			If DialogBriefSet.Item(1).CreatureTalking.Friendly = True Then
				picDialogBrief.Left = 0
				'UPGRADE_WARNING: Couldn't resolve default property of object DialogBriefSet(1).CreatureTalking. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
			ElseIf DialogBriefSet.Item(1).CreatureTalking.Name = "DM" Then 
				picDialogBrief.Left = Me.ClientRectangle.Width - picDialogBrief.Width
			Else
				picDialogBrief.Left = Me.ClientRectangle.Width - picDialogBrief.Width
			End If
			' Make it visible
			picDialogBrief.Refresh()
			picDialogBrief.Visible = True
			picDialogBrief.BringToFront()
			TimeToRead = VB.Timer() + 5
			' [Titi 2.4.8] change duration if message for too heavy/too bulky
			'UPGRADE_WARNING: Couldn't resolve default property of object DialogBriefSet().Text. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
			If VB.Right(DialogBriefSet.Item(1).Text, 21) = "slows the party down." Then TimeToRead = VB.Timer() + 1
		End If
		If VB.Timer() > TimeToRead Then
			If DialogBriefSet.Count() > 0 Then
				DialogBriefSet.Remove(1)
			End If
			TimeToRead = 0
			If DialogBriefSet.Count() = 0 Then
				picDialogBrief.Visible = False
				tmrDialogBrief.Enabled = False
			End If
		End If
	End Sub
	
	Private Sub tmrEncounterName_Tick(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles tmrEncounterName.Tick
		Static LastX, LastY As Short
		If Frozen = False Then
			If System.Math.Abs(HintX - LastX) < 8 And System.Math.Abs(HintY - LastY) < 8 Then
				MessageEncounterName(HintX, HintY)
			Else
				MessageShow("", 0)
				LastX = HintX
				LastY = HintY
			End If
		End If
	End Sub
	
	Private Sub tmrMoveMap_Tick(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles tmrMoveMap.Tick
		If Not Frozen And picGrid.Visible = False Then
			ScrollMap()
		End If
	End Sub
	
	Private Sub tmrMoveParty_Tick(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles tmrMoveParty.Tick
		If Not Frozen And picGrid.Visible = False Then
			MoveParty()
		End If
	End Sub
	
	'UPGRADE_ISSUE: PictureBox event picGrid.KeyDown was not upgraded. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="ABD9AF39-7E24-4AFF-AD8D-3675C1AA3054"'
	Private Sub picGrid_KeyDown(ByRef KeyCode As Short, ByRef Shift As Short)
		' KeyBoardShortCuts KeyCode, Shift
		ParseShortCuts(KeyCode, Shift)
	End Sub
	
	Private Sub picGrid_MouseDown(ByVal eventSender As System.Object, ByVal eventArgs As System.Windows.Forms.MouseEventArgs) Handles picGrid.MouseDown
		Dim Button As Short = eventArgs.Button \ &H100000
		Dim Shift As Short = System.Windows.Forms.Control.ModifierKeys \ &H10000
		Dim X As Single = eventArgs.X
		Dim Y As Single = eventArgs.Y
		ClickX = X : ClickY = Y
		If MenuNow = bdMenuTargetCreature Or MenuNow = bdMenuTargetAny Or MenuNow = bdMenuTargetParty Then
			If Target > -1 Then
				ConvoAction = Target + 1
			End If
		ElseIf Not Frozen Then 
			If Button = VB6.MouseButtonConstants.RightButton Then
				If Target > -1 Then
					CreatureWithTurn.CreatureTarget = CreatureTarget
					MenuSkillUse(True)
				End If
			Else
				CombatPC()
			End If
		End If
	End Sub
	
	Private Sub picGrid_MouseMove(ByVal eventSender As System.Object, ByVal eventArgs As System.Windows.Forms.MouseEventArgs) Handles picGrid.MouseMove
		Dim Button As Short = eventArgs.Button \ &H100000
		Dim Shift As Short = System.Windows.Forms.Control.ModifierKeys \ &H10000
		Dim X As Single = eventArgs.X
		Dim Y As Single = eventArgs.Y
		If Not Frozen Or (MenuNow = bdMenuTargetCreature Or MenuNow = bdMenuTargetAny Or MenuNow = bdMenuTargetParty) Then
			CombatMouseX = X : CombatMouseY = Y
			CombatRectangle()
		End If
	End Sub
	
	Private Function CombatAllDead() As Short
		Dim c As Short
		' If all enemies are dead, end combat
		CombatAllDead = True
		For c = 0 To CombatTurn
			If CreatureWithTurn.Friendly <> Turns(c).Ref.Friendly And Turns(c).Ref.HPNow > 0 And Turns(c).Ref.IsInanimate = False Then
				CombatAllDead = False
				Exit For
			End If
		Next c
	End Function
	
	Private Sub CombatPC()
		Dim OverSwing, c, NoFail, OldActionPoints As Short
		Dim ItemX As Item
		Dim Movement, Distance, MinRange, MaxRange As Short
		Dim Distance2, n, Distance1, Found As Short
		Dim CreatureX As Creature
		' If targeting something, then roll to hit
		If Target > -1 Then
			' If weapon in hand is a shoot and have no ammo, doesn't work and abort combat
			If CreatureWithTurn.ItemInHand.IsShooter = True Then
				If CreatureWithTurn.ItemAmmo.Index = CreatureWithTurn.ItemInHand.Index And CreatureWithTurn.ItemAmmo.Name = CreatureWithTurn.ItemInHand.Name Then
					CombatMessage(CreatureWithTurn.ItemInHand.Name & " out of ammo")
					Exit Sub
				End If
			End If
			' Compute (DistanceToTarget - Range - Movement)
			'UPGRADE_WARNING: Couldn't resolve default property of object CombatRange(). Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
			Distance = CombatRange(CreatureWithTurn, (CreatureTarget.Col), (CreatureTarget.Row))
			MinRange = Greatest((CreatureWithTurn.ItemMinRange), TargetCreatureRange)
			MaxRange = Greatest((CreatureWithTurn.ItemMaxRange), TargetCreatureRange)
			Movement = Greatest((CreatureWithTurn.MovementRange), 1)
			If Distance > MaxRange Then
				CombatMove(MOVEDIRECTION.bdMoveToward, MOVEMOTIVE.bdMoveTarget)
			ElseIf Distance < MinRange Then 
				CombatMove(MOVEDIRECTION.bdMoveAway, MOVEMOTIVE.bdMoveTarget)
			End If
			' Compute new distance after move
			'UPGRADE_WARNING: Couldn't resolve default property of object CombatRange(). Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
			Distance = CombatRange(CreatureWithTurn, (CreatureTarget.Col), (CreatureTarget.Row))
			' Compute Overswing penalty or message
			OverSwing = 0
			If (CreatureWithTurn.ActionPoints - CreatureWithTurn.ItemInHand.ActionPoints < 0) Or (CreatureWithTurn.ItemInHand.ActionPoints = 0 And CreatureWithTurn.ActionPoints < 10) Then
				' Compute over swing penalty or required action points
				If GlobalOverSwing = 1 Then
					If CreatureWithTurn.ItemInHand.ActionPoints = 0 Then
						OverSwing = CreatureWithTurn.ActionPoints - 10
					Else
						OverSwing = CreatureWithTurn.ActionPoints - CreatureWithTurn.ItemInHand.ActionPoints
					End If
				Else
					If CreatureWithTurn.ItemInHand.ActionPoints = 0 Then
						CombatMessage(CreatureWithTurn.ItemInHand.Name & " requires 10 Action Points")
					Else
						CombatMessage(CreatureWithTurn.ItemInHand.Name & " requires " & CreatureWithTurn.ItemInHand.ActionPoints & " Action Points")
					End If
					Distance = MaxRange + 1
				End If
			End If
			' If now in range, attack
			If Distance <= MaxRange And Distance >= MinRange And CreatureWithTurn.Friendly <> CreatureTarget.Friendly Then
				' Remove ActionPoints
				If CreatureWithTurn.ItemInHand.ActionPoints = 0 Then
					CreatureWithTurn.ActionPoints = CreatureWithTurn.ActionPoints - 10
				Else
					CreatureWithTurn.ActionPoints = CreatureWithTurn.ActionPoints - CreatureWithTurn.ItemInHand.ActionPoints
				End If
				CombatDraw()
				ItemNow = CreatureWithTurn.ItemInHand
				' Begin Attack
				CombatMessage(CreatureWithTurn.ItemInHand.Name & " Attack")
				PlaySFXItem(CreatureWithTurn)
				If CombatRollAttack(Greatest(CreatureWithTurn.ItemAmmo.DamageType, 0), ItemNow.AttackBonus + OverSwing) = True Then
					c = CombatRollDamage(bdContextDice, CreatureWithTurn.ItemAmmo.Damage - 1, Greatest(CreatureNow.ItemAmmo.DamageType, 0), True, ItemNow.DamageBonus + CreatureNow.ItemAmmo.DamageBonus)
				Else
					CombatMessage("Miss")
				End If
				CreatureWithTurn.UseAmmo()
				CombatDraw()
				' Fire Post-Attack Triggers
				CreatureNow = CreatureWithTurn
				ItemNow = CreatureWithTurn.ItemInHand
				NoFail = FireTriggers(CreatureWithTurn, bdPostAttack)
				For	Each ItemX In CreatureWithTurn.Items
					If ItemX.IsReady = True Then
						CreatureNow = CreatureWithTurn
						ItemNow = ItemX
						NoFail = FireTriggers(ItemX, bdPostAttack)
					End If
				Next ItemX
			End If
		Else
			CombatMove(MOVEDIRECTION.bdMoveClick, MOVEDIRECTION.bdMoveClick)
		End If
		' Check if all opponents are dead or not
		If CombatAllDead = True Then
			CombatNextTurn()
			Exit Sub
		End If
		' If out of ActionPoints, go to next turn
		If CreatureWithTurn.ActionPoints < 1 Then
			CombatNextTurn()
			Exit Sub
		ElseIf GlobalAutoEndTurn > 0 Then 
			' Determine if enemy exists within weapon range
			Found = False
			For c = 0 To CombatTurn
				CreatureX = Turns(c).Ref
				CombatFindPath(CreatureWithTurn, CreatureX.Col, CreatureX.Row)
				' Distance1 is from where CreatureWithTurn can move to
				If CreatureWithTurn.MovementRange < 1 Or CombatStepTop < 0 Then
					Distance1 = 999
				Else
					n = Greatest(CombatStepTop - CreatureWithTurn.MovementRange, 0)
					'UPGRADE_WARNING: Couldn't resolve default property of object CombatRange(). Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
					Distance1 = CombatRange(CreatureX, CombatStepCol(n), CombatStepRow(n))
				End If
				' Distance2 is from where CreatureWithTurn is standing
				'UPGRADE_WARNING: Couldn't resolve default property of object CombatRange(). Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
				Distance2 = CombatRange(CreatureWithTurn, (CreatureX.Col), (CreatureX.Row))
				' MaxRange is the ItemInHand.MaxRange
				MaxRange = CreatureWithTurn.ItemMaxRange
				If (Distance1 <= MaxRange Or Distance2 <= MaxRange) And CreatureWithTurn.Friendly <> CreatureX.Friendly Then
					Found = True
					Exit For
				End If
			Next c
			If Not Found Then
				CombatNextTurn()
				Exit Sub
			End If
		End If
	End Sub
	
	'UPGRADE_NOTE: Tome was upgraded to Tome_Renamed. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="A9E4979A-37FA-4718-9994-97DD76ED70A7"'
	Private Sub RealTimeTicks()
		Dim Tome_Renamed As Object
		Static c As Short
		GlobalTicks = LoopNumber(0, 9999, GlobalTicks, 1)
		c = c + 1
		If c > 9 Then
			c = 0
			'UPGRADE_WARNING: Couldn't resolve default property of object Tome.RealSeconds. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
			Tome.RealSeconds = Tome.RealSeconds + 1
			'UPGRADE_WARNING: Couldn't resolve default property of object Tome.RealSeconds. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
			If Tome.RealSeconds > 59 Then
				'UPGRADE_WARNING: Couldn't resolve default property of object Tome.RealSeconds. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
				Tome.RealSeconds = 1
				'UPGRADE_WARNING: Couldn't resolve default property of object Tome.RealMinutes. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
				Tome.RealMinutes = Tome.RealMinutes + 1
				'UPGRADE_WARNING: Couldn't resolve default property of object Tome.RealMinutes. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
				If Tome.RealMinutes > 59 Then
					'UPGRADE_WARNING: Couldn't resolve default property of object Tome.RealMinutes. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
					Tome.RealMinutes = 1
					'UPGRADE_WARNING: Couldn't resolve default property of object Tome.RealHours. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
					Tome.RealHours = Tome.RealHours + 1
					'UPGRADE_WARNING: Couldn't resolve default property of object Tome.RealHours. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
					If Tome.RealHours > 23 Then
						'UPGRADE_WARNING: Couldn't resolve default property of object Tome.RealHours. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
						Tome.RealHours = 1
						'UPGRADE_WARNING: Couldn't resolve default property of object Tome.RealDays. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
						Tome.RealDays = Tome.RealDays + 1
						'UPGRADE_WARNING: Couldn't resolve default property of object Tome.RealDays. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
						If Tome.RealDays > 10000 Then
							'UPGRADE_WARNING: Couldn't resolve default property of object Tome.RealDays. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
							Tome.RealDays = 1
						End If
					End If
				End If
			End If
		End If
	End Sub
	
	Private Sub RealTimeEvent()
		Dim NoFail As Short
		CreatureNow = CreatureWithTurn
		ItemNow = CreatureWithTurn.ItemInHand
		NoFail = FireTriggers(Tome, bdOnRealTime)
		CreatureNow = CreatureWithTurn
		ItemNow = CreatureWithTurn.ItemInHand
		NoFail = FireTriggers(EncounterNow, bdOnRealTime)
		CreatureNow = CreatureWithTurn
		ItemNow = CreatureWithTurn.ItemInHand
		NoFail = FireTriggers(CreatureWithTurn, bdOnRealTime)
	End Sub
	
	Private Sub tmrOnRealTime_Tick(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles tmrOnRealTime.Tick
		If Frozen = False Then
			RealTimeEvent()
			RealTimeTicks()
		End If
	End Sub
End Class