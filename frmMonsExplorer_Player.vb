Option Strict Off
Option Explicit On
Imports VB = Microsoft.VisualBasic
Imports Microsoft.VisualBasic.PowerPacks
Friend Class frmMonsExplorerPlayer
	Inherits System.Windows.Forms.Form
    Dim tome As Tome = Tome.getInstance()
	Dim CreatureX As Creature
	Dim PictureStyle As Short
	Dim TmpPictureFile As String
	Dim TmpFaceLeft As Short
	Dim TmpFaceTop As Short
	Dim PictureDir As String
	Dim TmpPortraitFile As String
	Dim PortraitDir As String
	Dim TmpWAV(3) As String
	Dim TmpWAVFlag(3) As Short
	'UPGRADE_NOTE: Size was upgraded to Size_Renamed. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="A9E4979A-37FA-4718-9994-97DD76ED70A7"'
	Dim Size_Renamed As Double
	'UPGRADE_NOTE: MouseDown was upgraded to MouseDown_Renamed. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="A9E4979A-37FA-4718-9994-97DD76ED70A7"'
	Dim MouseDown_Renamed As Short
	Dim strCurrCrea As String
	Dim fsoFile As Short
	Dim strDirectory As String
	
	Const bdGeneral As Short = 1
	Const bdType As Short = 2
	Const bdArmor As Short = 3
	Const bdItems As Short = 4
	Const bdConversations As Short = 5
	Const bdTriggers As Short = 6
	Const bdSounds As Short = 7
	Const bdWeakness As Short = 8
	
	Private Declare Function SetSysColors Lib "user32" (ByVal nChanges As Integer, ByRef lpSysColor As Integer, ByRef lpColorValues As Integer) As Integer
	Private Declare Function GetSysColor Lib "user32" (ByVal nIndex As Integer) As Integer
	Const COLOR_SCROLLBAR As Short = 0 'The Scrollbar colour
	Const COLOR_BACKGROUND As Short = 1 'Colour of the background with no wallpaper
	Const COLOR_ACTIVECAPTION As Short = 2 'Caption of Active Window
	Const COLOR_INACTIVECAPTION As Short = 3 'Caption of Inactive window
	Const COLOR_MENU As Short = 4 'Menu
	Const COLOR_WINDOW As Short = 5 'Windows background
	Const COLOR_WINDOWFRAME As Short = 6 'Window frame
	Const COLOR_MENUTEXT As Short = 7 'Window Text
	Const COLOR_WINDOWTEXT As Short = 8 '3D dark shadow (Win95)
	Const COLOR_CAPTIONTEXT As Short = 9 'Text in window caption
	Const COLOR_ACTIVEBORDER As Short = 10 'Border of active window
	Const COLOR_INACTIVEBORDER As Short = 11 'Border of inactive window
	Const COLOR_APPWORKSPACE As Short = 12 'Background of MDI desktop
	Const COLOR_HIGHLIGHT As Short = 13 'Selected item background
	Const COLOR_HIGHLIGHTTEXT As Short = 14 'Selected menu item
	Const COLOR_BTNFACE As Short = 15 'Button
	Const COLOR_BTNSHADOW As Short = 16 '3D shading of button
	Const COLOR_GRAYTEXT As Short = 17 'Grey text, of zero if dithering is used.
	Const COLOR_BTNTEXT As Short = 18 'Button text
	Const COLOR_INACTIVECAPTIONTEXT As Short = 19 'Text of inactive window
	Const COLOR_BTNHIGHLIGHT As Short = 20 '3D highlight of button
	Const COLOR_2NDACTIVECAPTION As Short = 27 'Win98 only: 2nd active window color
	Const COLOR_2NDINACTIVECAPTION As Short = 28 'Win98 only: 2nd inactive window color
	Dim dblBackGround, dblForeGround As Double
	
	Private Sub InitCreaExpl()

		' [Titi 2.4.9] save the user's background and foreground system colors
		dblBackGround = GetSysColor(COLOR_HIGHLIGHT)
		dblForeGround = GetSysColor(COLOR_HIGHLIGHTTEXT)
		' now, set background as black and text as light orange
		SetSysColors(1, COLOR_HIGHLIGHT, RGB(0, 0, 0))
		SetSysColors(1, COLOR_HIGHLIGHTTEXT, RGB(255, 192, 0))
		strDirectory = gAppPath & "\Roster\" & WorldNow.Name
		SizeAndShow()
	End Sub
	
	' Private Sub cmdClose_Click()
	' [Titi 2.4.9] Not in use any more
	' ********** EDIT EPHESTION
	'    If Me.Frame1.Visible = False Then
	'        Me.Frame1.Visible = True
	'        Me.Frame2.Visible = False
	'        cmdClose.Caption = "Skills and items"
	'    Else
	'        Me.Frame1.Visible = False
	'        Me.Frame2.Visible = True
	'        cmdClose.Caption = "Stats"
	'    End If
	' ********* END EDIT EPHESTION
	' End Sub
	
	Private Sub btnConvos_Click(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles btnConvos.Click
		Me.fraCreatureStats.Visible = False
		Me.fraCreatureSkills.Visible = False
		Me.fraCreatureConvos.Visible = True
	End Sub
	
	Private Sub btnSkills_Click(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles btnSkills.Click
		Me.fraCreatureStats.Visible = False
		Me.fraCreatureSkills.Visible = True
		Me.fraCreatureConvos.Visible = False
	End Sub
	
	Private Sub btnSounds_Click(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles btnSounds.Click
		Me.fraCreatureStats.Visible = False
		Me.fraCreatureSkills.Visible = False
		Me.fraCreatureConvos.Visible = True
	End Sub
	
	Private Sub btnStat_Click(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles btnStat.Click
		Me.fraCreatureStats.Visible = True
		Me.fraCreatureSkills.Visible = False
		Me.fraCreatureConvos.Visible = False
	End Sub
	
	Private Sub btnStats_Click(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles btnStats.Click
		Me.fraCreatureStats.Visible = True
		Me.fraCreatureSkills.Visible = False
		Me.fraCreatureConvos.Visible = False
	End Sub
	
	Private Sub btnTriggers_Click(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles btnTriggers.Click
		Me.fraCreatureStats.Visible = False
		Me.fraCreatureSkills.Visible = True
		Me.fraCreatureConvos.Visible = False
	End Sub
	
	Private Sub cmdPlaySound_Click(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles cmdPlaySound.Click
		Dim Index As Short = cmdPlaySound.GetIndex(eventSender)
		'UPGRADE_WARNING: Screen property Screen.MousePointer has a new behavior. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6BA9B8D2-2A32-4B6E-8D36-44949974A5B4"'
		System.Windows.Forms.Cursor.Current = System.Windows.Forms.Cursors.WaitCursor
		If txtSoundFile(Index).Text = "[Default]" Or txtSoundFile(Index).Text = "" Then
			Select Case Index
				Case 0 ' Move
                    Call PlaySoundFile("step1.wav", tome)
				Case 1 ' Attack does not have a default sound
				Case 2 ' Hit
                    Call PlaySoundFile("hit.wav", tome)
				Case 3 ' Die
                    Call PlaySoundFile("monsdie.wav", tome)
			End Select
		Else
            Call PlaySoundFile(txtSoundFile(Index).Text, tome)
		End If
		'UPGRADE_WARNING: Screen property Screen.MousePointer has a new behavior. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6BA9B8D2-2A32-4B6E-8D36-44949974A5B4"'
		System.Windows.Forms.Cursor.Current = System.Windows.Forms.Cursors.Default
	End Sub
	
	Private Sub frmMonsExplorerPlayer_Load(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles MyBase.Load
		InitCreaExpl()
		' [Titi 2.4.9]
		strDirectory = gAppPath & "\Roster\" & WorldNow.Name
	End Sub
	
	Private Sub LoadCreature()
		fsoFile = FreeFile
		FileOpen(fsoFile, strDirectory & "\" & CreatureNow.Name & ".rsc", OpenMode.Binary)
		CreatureX = New Creature
		CreatureX.ReadFromFile(fsoFile)
		' [Ephestion 2.4.8] browsing several times through the list caused an error (#67, too many files)
		FileClose(fsoFile)
		DisplayCreaInfo()
	End Sub
	
	Private Sub SizeAndShow()
		'    Frame1.Visible = True
		'    Frame2.Visible = False
		' [Titi 2.4.9] added a frame, and renamed the other two for clarity
		fraCreatureStats.Visible = True
		fraCreatureSkills.Visible = False
		fraCreatureConvos.Visible = False
		If TomeRosterDrag > -1 Then
			LoadCreature()
		End If
	End Sub
	
	Private Sub DisplayCreaInfo()
		Dim c As Short
		Dim iNumItems As Short
		Dim sTempWAV As String
		Dim strSize(4) As String
		' Describe creature's size
		strSize(4) = "Huge" : strSize(3) = "Large" : strSize(2) = "Normal" : strSize(1) = "Small" : strSize(0) = "Tiny"
		' Default size before load any picture
		Size_Renamed = 1
		' Default to normal picture style (not portrait)
		PictureStyle = 0
		Me.Text = "Creature [" & CreatureX.Name & "]"
		txtName.Text = CreatureX.Name
		' [Titi 2.4.9] added the name on the other two frames
		txtCreatureName.Text = CreatureX.Name
		txtNameCreature.Text = CreatureX.Name
		' yep, could have used a set of txtName(Index)
		txtComments.Text = CreatureX.Comments
		txtRace.Text = IIf(Len(CreatureX.Race) > 0, CreatureX.Race, "Undefined")
		txtSize.Text = "Undefined"
		txtSkillPoints.Text = CStr(CreatureX.SkillPoints)
		lblHome.Text = CreatureX.Home
		For c = 0 To 4
			If CreatureX.Size(c) < 0 Then txtSize.Text = strSize(c)
		Next 
		' Load and paint Picture from myPicture
		'    PictureDir = gAppPath & "\data\graphics\creatures"
		' [Titi 2.4.9]
		PictureDir = gDataPath & "\graphics\creatures"
		TmpPictureFile = CreatureX.PictureFile
		TmpFaceLeft = CreatureX.FaceLeft
		TmpFaceTop = CreatureX.FaceTop
		TmpPortraitFile = CreatureX.PortraitFile
		Select Case PictureStyle
			Case 0 ' Normal Picture
				If Len(TmpPictureFile) > 0 Then
					ShowMonsPicture(TmpPictureFile)
				End If
			Case 1 ' Portrait Picture
				If Len(TmpPortraitFile) > 0 Then
					ShowMonsPortrait(TmpPortraitFile)
				End If
		End Select
		chkIsInanimate.CheckState = CreatureX.IsInanimate * CShort(True)
		chkAgressive.CheckState = CreatureX.Agressive * CShort(True)
		chkGuard.CheckState = CreatureX.Guard * CShort(True)
		chkIsMale.CheckState = CreatureX.Male * CShort(True)
		chkFriend.CheckState = CreatureX.Friendly * CShort(True)
		chkDMControlled.CheckState = CreatureX.DMControlled * CShort(True)
		chkGuard.CheckState = CreatureX.Guard * CShort(True)
		chkRequiredInTome.CheckState = CreatureX.RequiredInTome * CShort(True)
		chkRequiredInTome.Refresh()
		txtLevel.Text = CStr(Val(CStr(CreatureX.Level)))
		txtSkillPoints.Text = CStr(Val(CStr(CreatureX.SkillPoints)))
		txtExperiencePoints.Text = CStr(Val(CStr(CreatureX.ExperiencePoints)))
		' family
		For c = 0 To 9
			chkFamily(c).CheckState = CreatureX.Family(c) * CShort(True)
		Next c
		' Combat Rank
		' txtCombatRank = CreatureX.CombatRank
		Select Case CreatureX.CombatRank
			Case 0
				txtCombat.Text = "Random"
			Case 1
				txtCombat.Text = "Back/Ranged"
			Case 2
				txtCombat.Text = "Middle"
			Case 3
				txtCombat.Text = "Front/Melee"
		End Select
		' Statistics
		txtStats(0).Text = CStr(Val(CStr(CreatureX.HPMax)))
		txtStats(1).Text = CStr(Val(CStr(CreatureX.Will))) 'don't know what will is, so I'm assuming intelligence
		txtStats(2).Text = CStr(Val(CStr(CreatureX.Strength)))
		txtStats(3).Text = CStr(Val(CStr(CreatureX.Agility)))
		txtStats(4).Text = CStr(Val(CStr(CreatureX.MovementCost)))
		txtStats(5).Text = CStr(Val(CStr(CreatureX.ActionPointsMax)))
		' Vices
		txtVices(0).Text = CStr(Val(CStr(CreatureX.Lunacy)))
		txtVices(1).Text = CStr(Val(CStr(CreatureX.Revelry)))
		txtVices(2).Text = CStr(Val(CStr(CreatureX.Wrath)))
		txtVices(3).Text = CStr(Val(CStr(CreatureX.Pride)))
		txtVices(4).Text = CStr(Val(CStr(CreatureX.Greed)))
		txtVices(5).Text = CStr(Val(CStr(CreatureX.Lust)))
		' Resistances
		'   CreatureX.BodyType(x)
		'   0 - Wing, 1 - Tail, 2 - Body, 3 - Head, 4 - Arm, 5 - Leg, 6 - Antenna
		'   7 - Tentacle, 8 - Abdomen, 9 - Back, 10 - Neck
		For c = 0 To 7
			Select Case CreatureX.BodyType(c)
				Case 0
					lblBody(c).Text = "Wing"
				Case 1
					lblBody(c).Text = "Tail"
				Case 2
					lblBody(c).Text = "Body"
				Case 3
					lblBody(c).Text = "Head"
				Case 4
					lblBody(c).Text = "Arm"
				Case 5
					lblBody(c).Text = "Leg"
				Case 6
					lblBody(c).Text = "Antenna"
				Case 7
					lblBody(c).Text = "Tentacle"
				Case 8
					lblBody(c).Text = "Abdomen"
				Case 9
					lblBody(c).Text = "Back"
				Case 10
					lblBody(c).Text = "Neck"
			End Select
			lblResPerc(c).Text = CreatureX.Resistance(c) & "%"
		Next 
		' Resistance Bonuses
		For c = 0 To 7
			Select Case CreatureX.ResistanceTypeBonus(c)
				Case 0
					lblBonusPerc(c).Text = "<None>"
				Case 1
					lblBonusPerc(c).Text = "10%"
				Case 2
					lblBonusPerc(c).Text = "20%"
				Case 3
					lblBonusPerc(c).Text = "30%"
				Case 4
					lblBonusPerc(c).Text = "40%"
				Case 5
					lblBonusPerc(c).Text = "50%"
				Case 6
					lblBonusPerc(c).Text = "60%"
				Case 7
					lblBonusPerc(c).Text = "70%"
				Case 8
					lblBonusPerc(c).Text = "80%"
				Case 9
					lblBonusPerc(c).Text = "90%"
				Case 10
					lblBonusPerc(c).Text = "Immune"
				Case 11
					lblBonusPerc(c).Text = "Double Damage"
				Case 12
					lblBonusPerc(c).Text = "Triple Damage"
			End Select
		Next 
		' Items
		lstItems.Items.Clear()
		iNumItems = CreatureX.Items.Count()
		For c = 1 To iNumItems
			'UPGRADE_WARNING: Couldn't resolve default property of object CreatureX.Items().Name. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
			lstItems.Items.Add(CreatureX.Items.Item(c).Name)
		Next 
		' Triggers
		lstTriggers.Items.Clear()
		iNumItems = CreatureX.Triggers.Count()
		For c = 1 To iNumItems
			'UPGRADE_WARNING: Couldn't resolve default property of object CreatureX.Triggers(c).Name. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
			'UPGRADE_WARNING: Couldn't resolve default property of object CreatureX.Triggers().TriggerType. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
			lstTriggers.Items.Add(TriggerName(CreatureX.Triggers.Item(c).TriggerType) & " " & CreatureX.Triggers.Item(c).Name)
		Next 
		' Conversations - lstConvos
		lstConvos.Items.Clear()
		iNumItems = CreatureX.Conversations.Count()
		For c = 1 To iNumItems
			'UPGRADE_WARNING: Couldn't resolve default property of object CreatureX.Conversations(c).Index. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
			If CreatureX.CurrentConvo = CreatureX.Conversations.Item(c).Index Then
				'UPGRADE_WARNING: Couldn't resolve default property of object CreatureX.Conversations().Name. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
				lstConvos.Items.Add("<Default> " & CreatureX.Conversations.Item(c).Name)
			Else
				'UPGRADE_WARNING: Couldn't resolve default property of object CreatureX.Conversations().Name. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
				lstConvos.Items.Add(CreatureX.Conversations.Item(c).Name)
			End If
		Next 
		' Sounds
		For c = 0 To 3
			Select Case c
				Case 0
					sTempWAV = CreatureX.MoveWAV
				Case 1
					sTempWAV = CreatureX.AttackWAV
				Case 2
					sTempWAV = CreatureX.HitWAV
				Case 3
					sTempWAV = CreatureX.DieWAV
			End Select
			If sTempWAV = "" Then
				sTempWAV = "[Default]"
			End If
			txtSoundFile(c).Text = sTempWAV
		Next 
		' reset comment displays
		If lstTriggers.Items.Count > 0 Then
			lstTriggers.SelectedIndex = 0
			lstTriggers_SelectedIndexChanged(lstTriggers, New System.EventArgs())
		End If
		If lstItems.Items.Count > 0 Then
			lstItems.SelectedIndex = 0
			lstItems_SelectedIndexChanged(lstItems, New System.EventArgs())
		End If
	End Sub
	
	'UPGRADE_NOTE: Tome was upgraded to Tome_Renamed. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="A9E4979A-37FA-4718-9994-97DD76ED70A7"'
	Private Sub ShowMonsPortrait(ByRef FileName As String)

		Dim X, Y As Short
		Dim NewWidth, NewHeight As Short
		'UPGRADE_WARNING: Arrays in structure bmMons may need to be initialized before they can be used. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="814DF224-76BD-4BB4-BFFB-EA359CB9FC48"'
		Dim bmMons As BITMAPINFO
		Dim hMemMons As Integer
		'UPGRADE_WARNING: Arrays in structure bmBlack may need to be initialized before they can be used. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="814DF224-76BD-4BB4-BFFB-EA359CB9FC48"'
		Dim bmBlack As BITMAPINFO
		'UPGRADE_WARNING: Arrays in structure bmMask may need to be initialized before they can be used. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="814DF224-76BD-4BB4-BFFB-EA359CB9FC48"'
		Dim bmMask As BITMAPINFO
		Dim rc, lpMem, TransparentRGB As Integer
		Dim PortraitFile As String
		'UPGRADE_WARNING: Screen property Screen.MousePointer has a new behavior. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6BA9B8D2-2A32-4B6E-8D36-44949974A5B4"'
		System.Windows.Forms.Cursor.Current = System.Windows.Forms.Cursors.WaitCursor ' Hourglass
		' Load Bitmap
		'UPGRADE_WARNING: Couldn't resolve default property of object Tome.LoadPath. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
		'UPGRADE_WARNING: Dir has a new behavior. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="9B7D5ADD-D8FE-4819-A36C-6DEDAF088CC7"'
        PortraitFile = Dir(tome.LoadPath & "\" & FileName)
		If PortraitFile = "" Then
			'        PortraitFile = gAppPath & "\data\graphics\portraits\" & Filename
			' [Titi 2.4.9]
			PortraitFile = gDataPath & "\graphics\portraits\" & FileName
		Else
			'UPGRADE_WARNING: Couldn't resolve default property of object Tome.LoadPath. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
            PortraitFile = tome.LoadPath & "\" & FileName
		End If
		PortraitDir = ClipPath(PortraitFile)
		' Load Creature bitmap
		ReadBitmapFile(PortraitFile, bmMons, hMemMons, TransparentRGB)
		' Make a copy of the current palette for the Portrait
		'UPGRADE_WARNING: Couldn't resolve default property of object bmBlack. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
		bmBlack = bmMons
		' Then change Pure Blue to Pure Black
		ChangeColor(bmBlack, TransparentRGB, 0, 0, 0)
		' Paint bitmap to Portrait box using converted palette
		lpMem = GlobalLock(hMemMons)
		picMons.Width = bmMons.bmiHeader.biWidth
		picMons.Height = bmMons.bmiHeader.biHeight
		'UPGRADE_ISSUE: PictureBox property picMons.hdc was not upgraded. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="CC4C7EC0-C903-48FC-ACCC-81861D12DA4A"'
        rc = SDIBits(picMons, 0, 0, bmMons.bmiHeader.biWidth, bmMons.bmiHeader.biHeight, 0, 0, bmMons.bmiHeader.biWidth, bmMons.bmiHeader.biHeight, lpMem, bmBlack, DIB_RGB_COLORS, SRCCOPY)
		picMons.Refresh()
		' Convert to Mask and copy to picMask (pure Blue is the mask color)
		MakeMask(bmMons, bmMask, TransparentRGB)
		picMask.Width = bmMons.bmiHeader.biWidth
		picMask.Height = bmMons.bmiHeader.biHeight
		'UPGRADE_ISSUE: PictureBox property picMask.hdc was not upgraded. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="CC4C7EC0-C903-48FC-ACCC-81861D12DA4A"'
        rc = SDIBits(picMask, 0, 0, bmMons.bmiHeader.biWidth, bmMons.bmiHeader.biHeight, 0, 0, bmMons.bmiHeader.biWidth, bmMons.bmiHeader.biHeight, lpMem, bmMask, DIB_RGB_COLORS, SRCCOPY)
		picMask.Refresh()
		' Release memory
		rc = GlobalUnlock(hMemMons)
		rc = GlobalFree(hMemMons)
		' Resize Creature to fit in space available
		Size_Renamed = 1
		If picMons.Width > 66 Then
			Size_Renamed = 66 / picMons.Width
		End If
		If picMons.Height * Size_Renamed > 76 Then
			Size_Renamed = 76 / picMons.Height
		End If
		' Center Creature in frame
		X = CShort(picCreature.ClientRectangle.Width - (picMons.Width * Size_Renamed)) / 2
		Y = CShort(picCreature.ClientRectangle.Height - (picMons.Height * Size_Renamed)) / 2
		NewHeight = CShort(picMons.Height * Size_Renamed)
		NewWidth = CShort(picMons.Width * Size_Renamed)
		' Paint Creature
		'UPGRADE_ISSUE: PictureBox method picCreature.Cls was not upgraded. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="CC4C7EC0-C903-48FC-ACCC-81861D12DA4A"'
        picCreature = Nothing
        picCreature.Invalidate()

		'UPGRADE_ISSUE: PictureBox property picCreature.hdc was not upgraded. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="CC4C7EC0-C903-48FC-ACCC-81861D12DA4A"'
        rc = SetSBMode(picCreature, 3)
		'UPGRADE_ISSUE: PictureBox property picMask.hdc was not upgraded. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="CC4C7EC0-C903-48FC-ACCC-81861D12DA4A"'
		'UPGRADE_ISSUE: PictureBox property picCreature.hdc was not upgraded. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="CC4C7EC0-C903-48FC-ACCC-81861D12DA4A"'
        rc = SBlt(picCreature, picMask, X, Y, NewWidth, NewHeight, 0, 0, picMask.Width, picMask.Height, SRCAND)
		'UPGRADE_ISSUE: PictureBox property picMons.hdc was not upgraded. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="CC4C7EC0-C903-48FC-ACCC-81861D12DA4A"'
		'UPGRADE_ISSUE: PictureBox property picCreature.hdc was not upgraded. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="CC4C7EC0-C903-48FC-ACCC-81861D12DA4A"'
        rc = SBlt(picCreature, picMons, X, Y, NewWidth, NewHeight, 0, 0, picMons.Width, picMons.Height, SRCPAINT)
		picCreature.Refresh()
		'UPGRADE_WARNING: Screen property Screen.MousePointer has a new behavior. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6BA9B8D2-2A32-4B6E-8D36-44949974A5B4"'
		System.Windows.Forms.Cursor.Current = System.Windows.Forms.Cursors.Default ' Default
	End Sub
	
	'UPGRADE_NOTE: Tome was upgraded to Tome_Renamed. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="A9E4979A-37FA-4718-9994-97DD76ED70A7"'
	Private Sub ShowMonsPicture(ByRef FileName As String)

		Dim X, Y As Short
		Dim NewWidth, NewHeight As Short
		'UPGRADE_WARNING: Arrays in structure bmMons may need to be initialized before they can be used. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="814DF224-76BD-4BB4-BFFB-EA359CB9FC48"'
		Dim bmMons As BITMAPINFO
		Dim hMemMons As Integer
		'UPGRADE_WARNING: Arrays in structure bmBlack may need to be initialized before they can be used. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="814DF224-76BD-4BB4-BFFB-EA359CB9FC48"'
		Dim bmBlack As BITMAPINFO
		'UPGRADE_WARNING: Arrays in structure bmMask may need to be initialized before they can be used. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="814DF224-76BD-4BB4-BFFB-EA359CB9FC48"'
		Dim bmMask As BITMAPINFO
		Dim rc, lpMem, TransparentRGB As Integer
		Dim PictureFile As String
		Dim oActivePic As System.Windows.Forms.PictureBox
		' /********EDIT EPHESTION ************* /
		'UPGRADE_ISSUE: picCreature was upgraded to a Panel, and cannot be coerced to a PictureBox. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="0FF1188E-27C0-4FED-842D-159C65894C9B"'
		oActivePic = picCreature
		' / ***********END EDIT EPHESTION
		'UPGRADE_WARNING: Screen property Screen.MousePointer has a new behavior. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6BA9B8D2-2A32-4B6E-8D36-44949974A5B4"'
		System.Windows.Forms.Cursor.Current = System.Windows.Forms.Cursors.WaitCursor ' Hourglass
		' Load Bitmap
		'UPGRADE_WARNING: Couldn't resolve default property of object Tome.LoadPath. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
		'UPGRADE_WARNING: Dir has a new behavior. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="9B7D5ADD-D8FE-4819-A36C-6DEDAF088CC7"'
        PictureFile = Dir(tome.LoadPath & "\" & FileName)
		If PictureFile = "" Then
			'        PictureFile = gAppPath & "\data\graphics\creatures\" & Filename
			' [Titi 2.4.9]
			PictureFile = gDataPath & "\graphics\creatures\" & FileName
		Else
			'UPGRADE_WARNING: Couldn't resolve default property of object Tome.LoadPath. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
            PictureFile = tome.LoadPath & "\" & FileName
		End If
		PictureDir = ClipPath(PictureFile)
		' Load Creature bitmap
		ReadBitmapFile(PictureFile, bmMons, hMemMons, TransparentRGB)
		' Make a copy of the current palette for the picture
		'UPGRADE_WARNING: Couldn't resolve default property of object bmBlack. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
		bmBlack = bmMons
		' Then change Pure Blue to Pure Black
		ChangeColor(bmBlack, TransparentRGB, 0, 0, 0)
		' Paint bitmap to picture box using converted palette
		lpMem = GlobalLock(hMemMons)
		picMons.Width = bmMons.bmiHeader.biWidth
		picMons.Height = bmMons.bmiHeader.biHeight
		'UPGRADE_ISSUE: PictureBox property picMons.hdc was not upgraded. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="CC4C7EC0-C903-48FC-ACCC-81861D12DA4A"'
        rc = SDIBits(picMons, X, Y, bmMons.bmiHeader.biWidth, bmMons.bmiHeader.biHeight, 0, 0, bmMons.bmiHeader.biWidth, bmMons.bmiHeader.biHeight, lpMem, bmBlack, DIB_RGB_COLORS, SRCCOPY)
		picMons.Refresh()
		' Convert to Mask and copy to picMask (pure Blue is the mask color)
		MakeMask(bmMons, bmMask, TransparentRGB)
		picMask.Width = bmMons.bmiHeader.biWidth
		picMask.Height = bmMons.bmiHeader.biHeight
		'UPGRADE_ISSUE: PictureBox property picMask.hdc was not upgraded. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="CC4C7EC0-C903-48FC-ACCC-81861D12DA4A"'
        rc = SDIBits(picMask, 0, 0, bmMons.bmiHeader.biWidth, bmMons.bmiHeader.biHeight, 0, 0, bmMons.bmiHeader.biWidth, bmMons.bmiHeader.biHeight, lpMem, bmMask, DIB_RGB_COLORS, SRCCOPY)
		picMask.Refresh()
		' Release memory
		rc = GlobalUnlock(hMemMons)
		rc = GlobalFree(hMemMons)
		' Resize Creature to fit in space available
		Size_Renamed = 1
		If picMons.Width > VB6.PixelsToTwipsX(oActivePic.ClientRectangle.Width) Then
			Size_Renamed = VB6.PixelsToTwipsX(oActivePic.ClientRectangle.Width) / picMons.Width
		End If
		If picMons.Height * Size_Renamed > VB6.PixelsToTwipsY(oActivePic.ClientRectangle.Height) Then
			Size_Renamed = VB6.PixelsToTwipsY(oActivePic.ClientRectangle.Height) / picMons.Height
		End If
		' Center Creature in frame
		X = CShort(VB6.PixelsToTwipsX(oActivePic.ClientRectangle.Width) - (picMons.Width * Size_Renamed)) / 2
		Y = CShort(VB6.PixelsToTwipsY(oActivePic.ClientRectangle.Height) - (picMons.Height * Size_Renamed)) / 2
		NewHeight = CShort(picMons.Height * Size_Renamed)
		NewWidth = CShort(picMons.Width * Size_Renamed)
		' Paint Creature
		'UPGRADE_ISSUE: PictureBox method oActivePic.Cls was not upgraded. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="CC4C7EC0-C903-48FC-ACCC-81861D12DA4A"'
        oActivePic = Nothing
        oActivePic.Invalidate()
		'UPGRADE_ISSUE: PictureBox property oActivePic.hdc was not upgraded. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="CC4C7EC0-C903-48FC-ACCC-81861D12DA4A"'
        rc = SetSBMode(oActivePic, 3)
		'UPGRADE_ISSUE: PictureBox property picMask.hdc was not upgraded. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="CC4C7EC0-C903-48FC-ACCC-81861D12DA4A"'
		'UPGRADE_ISSUE: PictureBox property oActivePic.hdc was not upgraded. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="CC4C7EC0-C903-48FC-ACCC-81861D12DA4A"'
        rc = SBlt(picMask, oActivePic, X, Y, NewWidth, NewHeight, 0, 0, picMask.Width, picMask.Height, SRCAND)
		'UPGRADE_ISSUE: PictureBox property picMons.hdc was not upgraded. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="CC4C7EC0-C903-48FC-ACCC-81861D12DA4A"'
		'UPGRADE_ISSUE: PictureBox property oActivePic.hdc was not upgraded. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="CC4C7EC0-C903-48FC-ACCC-81861D12DA4A"'
        rc = SBlt(oActivePic, picMons, X, Y, NewWidth, NewHeight, 0, 0, picMons.Width, picMons.Height, SRCPAINT)
		'UPGRADE_WARNING: Screen property Screen.MousePointer has a new behavior. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6BA9B8D2-2A32-4B6E-8D36-44949974A5B4"'
		System.Windows.Forms.Cursor.Current = System.Windows.Forms.Cursors.Default ' Default
	End Sub
	
	'UPGRADE_WARNING: Form event frmMonsExplorerPlayer.Deactivate has a new behavior. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6BA9B8D2-2A32-4B6E-8D36-44949974A5B4"'
	Private Sub frmMonsExplorerPlayer_Deactivate(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles MyBase.Deactivate
		Dim intG, intR, intB As Short
		' [Titi 2.4.9] revert to the user's background and foreground system colors
		intB = Val("&H" & VB.Left(Hex(dblBackGround), 2))
		intG = Val("&H" & Mid(Hex(dblBackGround), 3, 2))
		intR = Val("&H" & VB.Right(Hex(dblBackGround), 2))
		SetSysColors(1, COLOR_HIGHLIGHT, RGB(intR, intG, intB))
		intB = Val("&H" & VB.Left(Hex(dblForeGround), 2))
		intG = Val("&H" & Mid(Hex(dblForeGround), 3, 2))
		intR = Val("&H" & VB.Right(Hex(dblForeGround), 2))
		SetSysColors(1, COLOR_HIGHLIGHTTEXT, RGB(intR, intG, intB))
		Me.Close()
	End Sub
	
	'UPGRADE_WARNING: Event lstTriggers.SelectedIndexChanged may fire when form is initialized. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="88B12AE1-6DE0-48A0-86F1-60C0686C026A"'
	Private Sub lstTriggers_SelectedIndexChanged(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles lstTriggers.SelectedIndexChanged
		'UPGRADE_WARNING: Couldn't resolve default property of object CreatureX.Triggers().Comments. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
		txtTriggComm.Text = CreatureX.Triggers.Item(lstTriggers.SelectedIndex + 1).Comments
	End Sub
	
	'UPGRADE_WARNING: Event lstItems.SelectedIndexChanged may fire when form is initialized. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="88B12AE1-6DE0-48A0-86F1-60C0686C026A"'
	Private Sub lstItems_SelectedIndexChanged(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles lstItems.SelectedIndexChanged
		'UPGRADE_WARNING: Couldn't resolve default property of object CreatureX.Items().Comments. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
		txtItemComments.Text = CreatureX.Items.Item(lstItems.SelectedIndex + 1).Comments
	End Sub
	
	' -------------------------------------------------------------------
	' code below is to keep the user from editing displayed creature info
	Private Sub chkAgressive_MouseUp(ByVal eventSender As System.Object, ByVal eventArgs As System.Windows.Forms.MouseEventArgs) Handles chkAgressive.MouseUp
		Dim Button As Short = eventArgs.Button \ &H100000
		Dim Shift As Short = System.Windows.Forms.Control.ModifierKeys \ &H10000
		Dim X As Single = VB6.PixelsToTwipsX(eventArgs.X)
		Dim Y As Single = VB6.PixelsToTwipsY(eventArgs.Y)
		If chkAgressive.CheckState Then
			chkAgressive.CheckState = System.Windows.Forms.CheckState.Unchecked
		Else
			chkAgressive.CheckState = System.Windows.Forms.CheckState.Checked
		End If
		Beep()
	End Sub
	
	Private Sub chkDMControlled_MouseUp(ByVal eventSender As System.Object, ByVal eventArgs As System.Windows.Forms.MouseEventArgs) Handles chkDMControlled.MouseUp
		Dim Button As Short = eventArgs.Button \ &H100000
		Dim Shift As Short = System.Windows.Forms.Control.ModifierKeys \ &H10000
		Dim X As Single = VB6.PixelsToTwipsX(eventArgs.X)
		Dim Y As Single = VB6.PixelsToTwipsY(eventArgs.Y)
		If chkDMControlled.CheckState Then
			chkDMControlled.CheckState = System.Windows.Forms.CheckState.Unchecked
		Else
			chkDMControlled.CheckState = System.Windows.Forms.CheckState.Checked
		End If
		Beep()
	End Sub
	
	Private Sub chkFamily_MouseUp(ByVal eventSender As System.Object, ByVal eventArgs As System.Windows.Forms.MouseEventArgs) Handles chkFamily.MouseUp
		Dim Button As Short = eventArgs.Button \ &H100000
		Dim Shift As Short = System.Windows.Forms.Control.ModifierKeys \ &H10000
		Dim X As Single = VB6.PixelsToTwipsX(eventArgs.X)
		Dim Y As Single = VB6.PixelsToTwipsY(eventArgs.Y)
		Dim Index As Short = chkFamily.GetIndex(eventSender)
		If chkFamily(Index).CheckState Then
			chkFamily(Index).CheckState = System.Windows.Forms.CheckState.Unchecked
		Else
			chkFamily(Index).CheckState = System.Windows.Forms.CheckState.Checked
		End If
		Beep()
	End Sub
	
	Private Sub chkFriend_MouseUp(ByVal eventSender As System.Object, ByVal eventArgs As System.Windows.Forms.MouseEventArgs) Handles chkFriend.MouseUp
		Dim Button As Short = eventArgs.Button \ &H100000
		Dim Shift As Short = System.Windows.Forms.Control.ModifierKeys \ &H10000
		Dim X As Single = VB6.PixelsToTwipsX(eventArgs.X)
		Dim Y As Single = VB6.PixelsToTwipsY(eventArgs.Y)
		If chkFriend.CheckState Then
			chkFriend.CheckState = System.Windows.Forms.CheckState.Unchecked
		Else
			chkFriend.CheckState = System.Windows.Forms.CheckState.Checked
		End If
		Beep()
	End Sub
	
	
	Private Sub chkGuard_MouseUp(ByVal eventSender As System.Object, ByVal eventArgs As System.Windows.Forms.MouseEventArgs) Handles chkGuard.MouseUp
		Dim Button As Short = eventArgs.Button \ &H100000
		Dim Shift As Short = System.Windows.Forms.Control.ModifierKeys \ &H10000
		Dim X As Single = VB6.PixelsToTwipsX(eventArgs.X)
		Dim Y As Single = VB6.PixelsToTwipsY(eventArgs.Y)
		If chkGuard.CheckState Then
			chkGuard.CheckState = System.Windows.Forms.CheckState.Unchecked
		Else
			chkGuard.CheckState = System.Windows.Forms.CheckState.Checked
		End If
		Beep()
	End Sub
	
	Private Sub chkIsInanimate_MouseUp(ByVal eventSender As System.Object, ByVal eventArgs As System.Windows.Forms.MouseEventArgs) Handles chkIsInanimate.MouseUp
		Dim Button As Short = eventArgs.Button \ &H100000
		Dim Shift As Short = System.Windows.Forms.Control.ModifierKeys \ &H10000
		Dim X As Single = VB6.PixelsToTwipsX(eventArgs.X)
		Dim Y As Single = VB6.PixelsToTwipsY(eventArgs.Y)
		If chkIsInanimate.CheckState Then
			chkIsInanimate.CheckState = System.Windows.Forms.CheckState.Unchecked
		Else
			chkIsInanimate.CheckState = System.Windows.Forms.CheckState.Checked
		End If
		Beep()
	End Sub
	
	Private Sub chkIsMale_MouseUp(ByVal eventSender As System.Object, ByVal eventArgs As System.Windows.Forms.MouseEventArgs) Handles chkIsMale.MouseUp
		Dim Button As Short = eventArgs.Button \ &H100000
		Dim Shift As Short = System.Windows.Forms.Control.ModifierKeys \ &H10000
		Dim X As Single = VB6.PixelsToTwipsX(eventArgs.X)
		Dim Y As Single = VB6.PixelsToTwipsY(eventArgs.Y)
		If chkIsMale.CheckState Then
			chkIsMale.CheckState = System.Windows.Forms.CheckState.Unchecked
		Else
			chkIsMale.CheckState = System.Windows.Forms.CheckState.Checked
		End If
		Beep()
	End Sub
	
	Private Sub chkRequiredInTome_MouseUp(ByVal eventSender As System.Object, ByVal eventArgs As System.Windows.Forms.MouseEventArgs) Handles chkRequiredInTome.MouseUp
		Dim Button As Short = eventArgs.Button \ &H100000
		Dim Shift As Short = System.Windows.Forms.Control.ModifierKeys \ &H10000
		Dim X As Single = VB6.PixelsToTwipsX(eventArgs.X)
		Dim Y As Single = VB6.PixelsToTwipsY(eventArgs.Y)
		If chkRequiredInTome.CheckState Then
			chkRequiredInTome.CheckState = System.Windows.Forms.CheckState.Unchecked
		Else
			chkRequiredInTome.CheckState = System.Windows.Forms.CheckState.Checked
		End If
		Beep()
	End Sub
	
	Private Sub txtCombat_KeyPress(ByVal eventSender As System.Object, ByVal eventArgs As System.Windows.Forms.KeyPressEventArgs) Handles txtCombat.KeyPress
		Dim KeyAscii As Short = Asc(eventArgs.KeyChar)
		Beep()
		eventArgs.KeyChar = Chr(KeyAscii)
		If KeyAscii = 0 Then
			eventArgs.Handled = True
		End If
	End Sub
	
	Private Sub txtComments_KeyPress(ByVal eventSender As System.Object, ByVal eventArgs As System.Windows.Forms.KeyPressEventArgs) Handles txtComments.KeyPress
		Dim KeyAscii As Short = Asc(eventArgs.KeyChar)
		Beep()
		eventArgs.KeyChar = Chr(KeyAscii)
		If KeyAscii = 0 Then
			eventArgs.Handled = True
		End If
	End Sub
	
	Private Sub txtExperiencePoints_KeyPress(ByVal eventSender As System.Object, ByVal eventArgs As System.Windows.Forms.KeyPressEventArgs) Handles txtExperiencePoints.KeyPress
		Dim KeyAscii As Short = Asc(eventArgs.KeyChar)
		Beep()
		eventArgs.KeyChar = Chr(KeyAscii)
		If KeyAscii = 0 Then
			eventArgs.Handled = True
		End If
	End Sub
	
	Private Sub txtItemComments_KeyPress(ByVal eventSender As System.Object, ByVal eventArgs As System.Windows.Forms.KeyPressEventArgs) Handles txtItemComments.KeyPress
		Dim KeyAscii As Short = Asc(eventArgs.KeyChar)
		Beep()
		eventArgs.KeyChar = Chr(KeyAscii)
		If KeyAscii = 0 Then
			eventArgs.Handled = True
		End If
	End Sub
	
	Private Sub txtLevel_KeyPress(ByVal eventSender As System.Object, ByVal eventArgs As System.Windows.Forms.KeyPressEventArgs) Handles txtLevel.KeyPress
		Dim KeyAscii As Short = Asc(eventArgs.KeyChar)
		Beep()
		eventArgs.KeyChar = Chr(KeyAscii)
		If KeyAscii = 0 Then
			eventArgs.Handled = True
		End If
	End Sub
	
	Private Sub txtName_KeyPress(ByVal eventSender As System.Object, ByVal eventArgs As System.Windows.Forms.KeyPressEventArgs) Handles txtName.KeyPress
		Dim KeyAscii As Short = Asc(eventArgs.KeyChar)
		Beep()
		eventArgs.KeyChar = Chr(KeyAscii)
		If KeyAscii = 0 Then
			eventArgs.Handled = True
		End If
	End Sub
	
	Private Sub txtRace_KeyPress(ByVal eventSender As System.Object, ByVal eventArgs As System.Windows.Forms.KeyPressEventArgs) Handles txtRace.KeyPress
		Dim KeyAscii As Short = Asc(eventArgs.KeyChar)
		Beep()
		eventArgs.KeyChar = Chr(KeyAscii)
		If KeyAscii = 0 Then
			eventArgs.Handled = True
		End If
	End Sub
	
	Private Sub txtSize_KeyPress(ByVal eventSender As System.Object, ByVal eventArgs As System.Windows.Forms.KeyPressEventArgs) Handles txtSize.KeyPress
		Dim KeyAscii As Short = Asc(eventArgs.KeyChar)
		Beep()
		eventArgs.KeyChar = Chr(KeyAscii)
		If KeyAscii = 0 Then
			eventArgs.Handled = True
		End If
	End Sub
	
	Private Sub txtSkillPoints_KeyPress(ByVal eventSender As System.Object, ByVal eventArgs As System.Windows.Forms.KeyPressEventArgs) Handles txtSkillPoints.KeyPress
		Dim KeyAscii As Short = Asc(eventArgs.KeyChar)
		Beep()
		eventArgs.KeyChar = Chr(KeyAscii)
		If KeyAscii = 0 Then
			eventArgs.Handled = True
		End If
	End Sub
	
	Private Sub txtSoundFile_KeyPress(ByVal eventSender As System.Object, ByVal eventArgs As System.Windows.Forms.KeyPressEventArgs) Handles txtSoundFile.KeyPress
		Dim KeyAscii As Short = Asc(eventArgs.KeyChar)
		Dim Index As Short = txtSoundFile.GetIndex(eventSender)
		Beep()
		eventArgs.KeyChar = Chr(KeyAscii)
		If KeyAscii = 0 Then
			eventArgs.Handled = True
		End If
	End Sub
	
	Private Sub txtStats_KeyPress(ByVal eventSender As System.Object, ByVal eventArgs As System.Windows.Forms.KeyPressEventArgs) Handles txtStats.KeyPress
		Dim KeyAscii As Short = Asc(eventArgs.KeyChar)
		Dim Index As Short = txtStats.GetIndex(eventSender)
		Beep()
		eventArgs.KeyChar = Chr(KeyAscii)
		If KeyAscii = 0 Then
			eventArgs.Handled = True
		End If
	End Sub
	
	Private Sub txtTriggComm_KeyPress(ByVal eventSender As System.Object, ByVal eventArgs As System.Windows.Forms.KeyPressEventArgs) Handles txtTriggComm.KeyPress
		Dim KeyAscii As Short = Asc(eventArgs.KeyChar)
		Beep()
		eventArgs.KeyChar = Chr(KeyAscii)
		If KeyAscii = 0 Then
			eventArgs.Handled = True
		End If
	End Sub
	
	Private Sub txtVices_KeyPress(ByVal eventSender As System.Object, ByVal eventArgs As System.Windows.Forms.KeyPressEventArgs) Handles txtVices.KeyPress
		Dim KeyAscii As Short = Asc(eventArgs.KeyChar)
		Dim Index As Short = txtVices.GetIndex(eventSender)
		Beep()
		eventArgs.KeyChar = Chr(KeyAscii)
		If KeyAscii = 0 Then
			eventArgs.Handled = True
		End If
	End Sub
End Class