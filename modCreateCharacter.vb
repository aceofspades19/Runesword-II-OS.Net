Option Strict Off
Option Explicit On
Module modCreateCharacter
    Private tome As Tome = tome.getInstance()
	Public Sub CreatePCLoadWorlds(Optional ByRef WorldTemp As World = Nothing)
		'Dim c As Integer, i As Integer, k As Integer, n As Integer, rc As Integer, FileName As String
		Dim k, c, i, n As Short
		Dim Filename As String
		Dim WorldCnt As Short
		Dim WorldX As World
		Dim Text As String '* 8192
		Dim Kingdoms As Short
		Dim KingdomX As Kingdom
		Dim CreatureX As Object
		Dim PictureFile As String
		Dim PictureX As Factoid
		Dim TomeFileName As String
		Dim DirName As String
		'UPGRADE_NOTE: Default was upgraded to Default_Renamed. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="A9E4979A-37FA-4718-9994-97DD76ED70A7"'
		Dim Default_Renamed As Short
		Dim DefaultWorld As String
		Dim lResult As Integer
		'UPGRADE_WARNING: Screen property Screen.MousePointer has a new behavior. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6BA9B8D2-2A32-4B6E-8D36-44949974A5B4"'
		System.Windows.Forms.Cursor.Current = System.Windows.Forms.Cursors.WaitCursor
		' Fetch all Worlds *.ini files
		Worlds = New Collection
		WorldCnt = 0
		' [Titi 2.4.9] added to fake the current world
		If WorldTemp Is Nothing Then
			lResult = fReadValue(gAppPath & "\Settings.ini", "World", "Current", "S", "Eternia", Text)
		Else
			Text = WorldTemp.Name
		End If
		DefaultWorld = Text
		'UPGRADE_WARNING: Dir has a new behavior. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="9B7D5ADD-D8FE-4819-A36C-6DEDAF088CC7"'
		DirName = Dir(gAppPath & "\roster\*", FileAttribute.Directory)
		Do Until DirName = ""
			' If (DirName <> "." And DirName <> "..") And GetAttr(gAppPath & "\roster\" & DirName) = vbDirectory Then
			' [Titi 2.4.8] In some rare cases, the above led to an endless loop
			If (DirName <> "." And DirName <> "..") And oFileSys.CheckExists(gAppPath & "\roster\" & DirName, clsInOut.IOActionType.Folder, False) Then
				Filename = gAppPath & "\roster\" & DirName & "\" & DirName & ".ini"
				' Create a new World
				WorldX = New World
				WorldX.Path = gAppPath & "\roster\" & DirName & "\"
				WorldCnt = WorldCnt + 1
				WorldX.Index = WorldCnt
				' [borfaux] Added for 2.4.6
				' Which file to play for Worlds Intro Music
				lResult = fReadValue(Filename, "World", "Intro", "S", "", Text)
				WorldX.IntroMusic = Text
				' [Titi] Added for 2.4.7
				' Which folder to look into for Worlds Music
				lResult = fReadValue(Filename, "World", "MusicFolder", "S", "", Text)
				WorldX.MusicFolder = Text
				' Fetch the WorldName from the *.ini file
				lResult = fReadValue(Filename, "World", "Name", "S", "Untitled", Text)
				WorldX.Name = Text
				' [Titi] Added for 2.4.8
				' Get the World Currency from the *.ini file
				lResult = fReadValue(Filename, "World", "Currency", "S", "Gold piece|gold1.BMP", Text)
				WorldX.Money = Text
				' [Titi 2.4.9] Get the World HP and skillpoints defaults from the *.ini file
				lResult = fReadValue(Filename, "World", "SkillPtsPerLevel", "S", "10", Text)
				WorldX.SkillPtsPerLevel = CShort(Val(Text))
				lResult = fReadValue(Filename, "World", "HPPerLevel", "S", "10", Text)
				WorldX.HPPerLevel = CShort(Val(Text))
				lResult = fReadValue(Filename, "World", "RandomizeHP", "S", "0", Text)
				WorldX.RandomizeHP = CShort(Val(Text))
				' [Titi 2.4.9] Get the World Description from the *.ini file
				lResult = fReadValue(Filename, "World", "Description", "S", "No description available", Text)
				WorldX.Description = Text
				lResult = fReadValue(Filename, "World", "PicDesc", "S", gDataPath & "\Stock\blankmap.bmp|" & gDataPath & "\Stock\blankmap.bmp", Text)
				WorldX.PicDesc = Text
				' Add the World to the collection Worlds
				Worlds.Add(WorldX, "W" & WorldX.Index)
				If WorldX.Name = DefaultWorld Then
					Default_Renamed = WorldX.Index
				End If
				' Get the World's PictureFile name
				lResult = fReadValue(Filename, "World", "PictureFile", "S", "Untitled", Text)
				WorldX.PictureFile = Text
				' Get the World's Tome name
				lResult = fReadValue(Filename, "World", "Tome", "S", "Untitled", Text)
				WorldX.TomeName = Text
				' Read Tome
				WorldX.Tome = New Tome
			End If
			'UPGRADE_WARNING: Dir has a new behavior. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="9B7D5ADD-D8FE-4819-A36C-6DEDAF088CC7"'
			DirName = Dir()
		Loop 
		WorldNow = Worlds.Item(Default_Renamed)
		WorldNow.IsCurrent = True
		ScrollTop = 1
		'UPGRADE_WARNING: Screen property Screen.MousePointer has a new behavior. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6BA9B8D2-2A32-4B6E-8D36-44949974A5B4"'
		System.Windows.Forms.Cursor.Current = System.Windows.Forms.Cursors.Default
	End Sub
	
	Public Sub CreatePCClick(ByRef AtX As Short, ByRef AtY As Short, ByRef ButtonDown As Short)
		Dim n, c, k As Short
		Dim rc As Integer
		Dim TriggerX As Trigger
		Dim KingdomX As Kingdom
		' [Titi 2.4.7] what if more than 12 kingdoms?
		ReDim Preserve strKingdomNames(WorldNow.Kingdoms.Count() - 1)
		If PointIn(AtX, AtY, 380, 364, 90, 18) Then
			' Cancel/Back
			If ButtonDown Then
				If TomeAction = bdCreatePCKingdom Then
					frmMain.ShowButton((frmMain.picTomeNew), 380, 364, "Cancel", True)
				Else
					frmMain.ShowButton((frmMain.picTomeNew), 380, 364, "Back", True)
				End If
				frmMain.picTomeNew.Refresh()
			Else
				Call PlayClickSnd(modIOFunc.ClickType.ifClick)
				If TomeAction = bdCreatePCKingdom Then
					frmMain.ShowButton((frmMain.picTomeNew), 380, 364, "Cancel", False)
				Else
					frmMain.ShowButton((frmMain.picTomeNew), 380, 364, "Back", False)
				End If
				Select Case TomeAction
					Case bdCreatePCKingdom
						frmMain.picTomeNew.Refresh()
						' Either return to Gather or Main Menu
						If CreatePCReturn = True Then
							frmMain.TomeGatherList()
						Else
							frmMain.picTomeNew.Visible = False
						End If
					Case bdCreatePCSkills
						' Clear the Skills
						For	Each TriggerX In KingdomNow.Skills
							TriggerX.TempSkill = 0
						Next TriggerX
						CreatePCKingdom(KingdomIndex)
					Case bdCreatePCPicture
						CreatureNow.Male = True
						CreatePCSkills(1)
					Case bdCreatePCName
						frmMain.picCreateName.Visible = False
						CreatePCPicture(CreatePictureIndex)
				End Select
				frmMain.picTomeNew.Refresh()
			End If
		ElseIf PointIn(AtX, AtY, 476, 364, 90, 18) Then 
			' Next
			If ButtonDown Then
				If TomeAction = bdCreatePCName Then
					frmMain.ShowButton((frmMain.picTomeNew), 476, 364, "Save", True)
				Else
					frmMain.ShowButton((frmMain.picTomeNew), 476, 364, "Next", True)
				End If
				frmMain.picTomeNew.Refresh()
			Else
				If TomeAction = bdCreatePCName Then
					frmMain.ShowButton((frmMain.picTomeNew), 476, 364, "Save", False)
				Else
					frmMain.ShowButton((frmMain.picTomeNew), 476, 364, "Next", False)
				End If
				frmMain.picTomeNew.Refresh()
				Call PlayClickSnd(modIOFunc.ClickType.ifClick)
				Select Case TomeAction
					Case bdCreatePCKingdom
						ScrollTop = 1
						CreatePCSkills(1)
					Case bdCreatePCSkills
						CreatePCPicture(1)
					Case bdCreatePCPicture
						For c = 0 To WorldNow.Kingdoms.Count() - 1 ' [Titi 2.4.7] was 11
							strKingdomNames(c) = MakeCreatureName((KingdomNow.NameSet), (WorldNow.Name))
						Next c
						CreateNameNew = strKingdomNames(0)
						frmMain.picCreateName.Visible = True
						CreatePCName(0)
						CreatePCNameEnter(0)
					Case bdCreatePCName
						' Save Creature
						CreatePCSave()
						' Clear Skills
						For	Each TriggerX In KingdomNow.Skills
							TriggerX.TempSkill = 0
						Next TriggerX
						' Either return to Gather or Main Menu
						frmMain.picCreateName.Visible = False
						If CreatePCReturn = True Then
							' Reload the Roster
							frmMain.TomeLoadRoster()
							frmMain.TomeGatherList()
						Else
							frmMain.picTomeNew.Visible = False
						End If
				End Select
			End If
		ElseIf PointIn(AtX, AtY, 62, 331, 90, 18) And TomeAction = bdCreatePCKingdom Then 
			' Back Kingdoms
			If ButtonDown Then
				frmMain.ShowButton((frmMain.picTomeNew), 62, 331, "Back", True)
				frmMain.picTomeNew.Refresh()
			Else
				Call PlayClickSnd(modIOFunc.ClickType.ifClick)
				frmMain.ShowButton((frmMain.picTomeNew), 62, 331, "Back", False)
				frmMain.picTomeNew.Refresh()
				KingdomIndex = LoopNumber(1, WorldNow.Kingdoms.Count(), KingdomIndex, -1)
				KingdomX = WorldNow.Kingdoms.Item(KingdomIndex)
				CreatePCKingdom(KingdomIndex)
			End If
		ElseIf PointIn(AtX, AtY, 158, 331, 90, 18) And TomeAction = bdCreatePCKingdom Then 
			' More Kingdoms
			If ButtonDown Then
				frmMain.ShowButton((frmMain.picTomeNew), 158, 331, "More", True)
				frmMain.picTomeNew.Refresh()
			Else
				Call PlayClickSnd(modIOFunc.ClickType.ifClick)
				frmMain.ShowButton((frmMain.picTomeNew), 158, 331, "More", False)
				frmMain.picTomeNew.Refresh()
				KingdomIndex = LoopNumber(1, WorldNow.Kingdoms.Count(), KingdomIndex, 1)
				KingdomX = WorldNow.Kingdoms.Item(KingdomIndex)
				CreatePCKingdom(KingdomIndex)
			End If
		ElseIf PointIn(AtX, AtY, 52, 331, 82, 18) And TomeAction = bdCreatePCPicture And ButtonDown = True Then 
			' Male
			CreatureNow.Male = True
			CreatePCPicture(CreatePictureIndex)
			Call PlayClickSnd(modIOFunc.ClickType.ifClick)
		ElseIf PointIn(AtX, AtY, 166, 331, 98, 18) And TomeAction = bdCreatePCPicture And ButtonDown = True Then 
			' Female
			CreatureNow.Male = False
			CreatePCPicture(CreatePictureIndex)
			Call PlayClickSnd(modIOFunc.ClickType.ifClick)
		ElseIf PointIn(AtX, AtY, 25, 70, 272, 271) And TomeAction = bdCreatePCPicture And ButtonDown = True Then 
			' Choose Picture
			CreatePCPicture(Least(Int((AtX - 25) / 84) + Int((AtY - 70) / 86) * 3 + 1, KingdomNow.PictureFiles.Count()))
			Call PlayClickSnd(modIOFunc.ClickType.ifClick)
		ElseIf PointIn(AtX, AtY, 24, 68, 245, 260) And TomeAction = bdCreatePCName And ButtonDown = False Then 
			' Select Name from List
			c = Int((AtY - 68) / 22)
			' If IsBetween(c, 0, 11) Then '[Titi 2.4.7]
			If IsBetween(c, 0, WorldNow.Kingdoms.Count() - 1) Then
				CreateNameNew = strKingdomNames(c)
				CreatePCName(c)
				Call PlayClickSnd(modIOFunc.ClickType.ifClick)
			End If
		ElseIf PointIn(AtX, AtY, 276, 62, 18, 267) And TomeAction = bdCreatePCSkills Then 
			' ScrollBar Click
			'UPGRADE_ISSUE: frmMain.picTomeNew was upgraded to a Panel, and cannot be coerced to a PictureBox. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="0FF1188E-27C0-4FED-842D-159C65894C9B"'
			If frmMain.ScrollBarClick(AtX, AtY, ButtonDown, (frmMain.picTomeNew), 276, 62, 267, ScrollTop, KingdomNow.Skills.Count(), 9) = True Then
				CreatePCSkills(CreateSkillsIndex)
			End If
		ElseIf PointIn(AtX, AtY, 26, 73, 18, 250) And TomeAction = bdCreatePCSkills And ButtonDown = True Then 
			' Select Skills (check box)
			c = Int((AtY - 73) / 25) + ScrollTop
			If c <= KingdomNow.Skills.Count() Then
				'UPGRADE_WARNING: Couldn't resolve default property of object KingdomNow.Skills(c).TempSkill. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
				If KingdomNow.Skills.Item(c).TempSkill > 0 Then
					'UPGRADE_WARNING: Couldn't resolve default property of object KingdomNow.Skills().TempSkill. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
					KingdomNow.Skills.Item(c).TempSkill = 0
					Call PlayClickSnd(modIOFunc.ClickType.ifClick)
				Else
					' Count Total selected and total Skill Points Left
					For	Each TriggerX In KingdomNow.Skills
						If TriggerX.TempSkill > 0 Then
							k = k + TriggerX.TempSkill
						End If
					Next TriggerX
					' Check for limits
					'UPGRADE_WARNING: Couldn't resolve default property of object KingdomNow.Skills(c).Turns. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
					If CreatureNow.SkillPoints - k - KingdomNow.Skills.Item(c).Turns > -1 Then
						'UPGRADE_WARNING: Couldn't resolve default property of object KingdomNow.Skills(c).TempSkill. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
						'UPGRADE_WARNING: Couldn't resolve default property of object KingdomNow.Skills().Turns. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
						KingdomNow.Skills.Item(c).TempSkill = KingdomNow.Skills.Item(c).Turns
						Call PlayClickSnd(modIOFunc.ClickType.ifClick)
					Else
						Call PlayClickSnd(modIOFunc.ClickType.ifClickPass)
					End If
				End If
				CreatePCSkills(c)
			End If
		ElseIf PointIn(AtX, AtY, 53, 73, 276, 250) And TomeAction = bdCreatePCSkills And ButtonDown = True Then 
			' Describe Skills
			c = Int((AtY - 73) / 25) + ScrollTop
			If c <= KingdomNow.Skills.Count() Then
				CreatePCSkills(c)
				Call PlayClickSnd(modIOFunc.ClickType.ifClick)
			End If
		End If
	End Sub
	
	Public Sub CreatePCLoadMap(ByRef WorldX As World)
		frmMain.picTmp.Image = System.Drawing.Image.FromFile(gAppPath & "\roster\" & WorldX.Name & "\" & WorldX.PictureFile)
	End Sub
	
	Public Sub CreatePCDelete()
		Dim Filename As String
		Dim sMsg As String
		On Error GoTo Err_Handler
		DialogSetUp(modGameGeneral.DLGTYPE.bdDlgNoReply)
		DialogSetButton(1, "No")
		DialogSetButton(2, "Yes")
		'UPGRADE_WARNING: Couldn't resolve default property of object TomeRoster().Name. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
		DialogShow("DM", "Do you really wish to delete " & TomeRoster.Item(TomeRosterSelect).Name & "?")
		frmMain.picConvo.Visible = False
		' If Yes then Remove the Creature
		If ConvoAction = 0 Then
			'Kill gAppPath & "\roster\" & WorldNow.Name & "\" & TomeRoster(TomeRosterSelect).Name & ".rsc"
			'UPGRADE_WARNING: Couldn't resolve default property of object TomeRoster().Name. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
			If oFileSys.Delete(gAppPath & "\roster\" & WorldNow.Name & "\" & TomeRoster.Item(TomeRosterSelect).Name & ".rsc", clsInOut.IOActionType.File) Then
				'UPGRADE_WARNING: Couldn't resolve default property of object TomeRoster().Name. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
				sMsg = TomeRoster.Item(TomeRosterSelect).Name & " deleted."
			Else
				'UPGRADE_WARNING: Couldn't resolve default property of object TomeRoster().Name. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
				sMsg = "Unable to delete character: " & TomeRoster.Item(TomeRosterSelect).Name
			End If
			DialogSetUp(modGameGeneral.DLGTYPE.bdDlgNoReply)
			DialogSetButton(1, "Done")
			DialogShow("DM", sMsg) 'TomeRoster(TomeRosterSelect).Name & " deleted."
			TomeRoster.Remove(TomeRosterSelect)
			frmMain.picConvo.Visible = False
			frmMain.TomeGatherList()
		End If
Err_Handler: 
		oErr.logError("CreatePCDelete()")
		Resume Next
	End Sub
	
	Public Sub CreatePCKingdom(ByRef Index As Short)
		Dim c, n As Short
		Dim rc As Integer
		TomeAction = bdCreatePCKingdom
		' Show current Kingdom viewing
		KingdomNow = WorldNow.Kingdoms.Item(Index)
		CreatureNow = KingdomNow.Template
		KingdomIndex = Index
		' Show the Map
		'UPGRADE_ISSUE: PictureBox method picTomeNew.Cls was not upgraded. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="CC4C7EC0-C903-48FC-ACCC-81861D12DA4A"'
        frmMain.picTomeNew = Nothing
        frmMain.picTomeNew.Invalidate()
		'UPGRADE_ISSUE: PictureBox property picTmp.hdc was not upgraded. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="CC4C7EC0-C903-48FC-ACCC-81861D12DA4A"'
		'UPGRADE_ISSUE: PictureBox property picTomeNew.hdc was not upgraded. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="CC4C7EC0-C903-48FC-ACCC-81861D12DA4A"'
        rc = BBlt(frmMain.picTomeNew, frmMain.picTmp, 20, 63, 254, 265, Least((KingdomNow.Left_Renamed), frmMain.picTmp.ClientRectangle.Width - 254), Least((KingdomNow.Top), frmMain.picTmp.ClientRectangle.Height - 265), SRCCOPY)
		' Show the Name
		frmMain.ShowText((frmMain.picTomeNew), 30, 12, 519, 14, bdFontElixirWhite, "Where do you come from?", True, False)
		frmMain.ShowText((frmMain.picTomeNew), 19, 45, 273, 14, bdFontElixirWhite, (WorldNow.Name), True, False)
		frmMain.ShowText((frmMain.picTomeNew), 332, 48, 217, 15, bdFontElixirBlack, (KingdomNow.Name), True, False)
		frmMain.ShowText((frmMain.picTomeNew), 328, 108, 220, 190, bdFontNoxiousBlack, (CreatureNow.Comments), False, False)
		' Show basic stats
		frmMain.ShowText((frmMain.picTomeNew), 324, 84, 68, 14, bdFontNoxiousBlack, "Str " & CreatureNow.Strength, True, False)
		frmMain.ShowText((frmMain.picTomeNew), 409, 76, 68, 14, bdFontNoxiousBlack, "Int " & CreatureNow.Will, True, False)
		frmMain.ShowText((frmMain.picTomeNew), 486, 84, 68, 14, bdFontNoxiousBlack, "Agl " & CreatureNow.Agility, True, False)
		frmMain.ShowText((frmMain.picTomeNew), 328, 308, 220, 14, bdFontNoxiousBlack, "Health " & CreatureNow.HPNow & " (" & 100 * Int(Greatest((CreatureNow.HPNow), 0) / Greatest((CreatureNow.HPMax), 1)) & "%)", True, False)
		frmMain.ShowText((frmMain.picTomeNew), 328, 328, 220, 14, bdFontNoxiousBlack, "Defense " & CreatureNow.Defense, True, False)
		' Set up buttons
		frmMain.ShowButton((frmMain.picTomeNew), 62, 331, "Back", False)
		frmMain.ShowButton((frmMain.picTomeNew), 158, 331, "More", False)
		frmMain.ShowButton((frmMain.picTomeNew), 380, 364, "Cancel", False)
		frmMain.ShowButton((frmMain.picTomeNew), 476, 364, "Next", False)
		' Refresh the screen
		frmMain.picTomeNew.Top = (frmMain.ClientRectangle.Height - frmMain.picTomeNew.ClientRectangle.Height) / 2
		frmMain.picTomeNew.Left = (frmMain.ClientRectangle.Width - frmMain.picTomeNew.ClientRectangle.Width) / 2
		frmMain.picTomeNew.Refresh()
		frmMain.picTomeNew.BringToFront()
		frmMain.picTomeNew.Visible = True
		System.Windows.Forms.Application.DoEvents()
	End Sub
	
	'UPGRADE_NOTE: Tome was upgraded to Tome_Renamed. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="A9E4979A-37FA-4718-9994-97DD76ED70A7"'
	Private Sub CreatePCPicture(ByRef Index As Short)
		Dim Tome_Renamed As Object
		Dim c, n As Short
		Dim PictureFileX As Factoid
		Dim rc As Integer
		Dim Size As Double
		Dim Width, Height As Short
		Dim X, Y As Short
		'UPGRADE_WARNING: Arrays in structure bmMons may need to be initialized before they can be used. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="814DF224-76BD-4BB4-BFFB-EA359CB9FC48"'
		Dim bmMons As BITMAPINFO
		Dim hMemMons As Integer
		'UPGRADE_WARNING: Arrays in structure bmBlack may need to be initialized before they can be used. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="814DF224-76BD-4BB4-BFFB-EA359CB9FC48"'
		Dim bmBlack As BITMAPINFO
		'UPGRADE_WARNING: Arrays in structure bmMask may need to be initialized before they can be used. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="814DF224-76BD-4BB4-BFFB-EA359CB9FC48"'
		Dim bmMask As BITMAPINFO
		Dim lpMem, TransparentRGB As Integer
		'UPGRADE_WARNING: Screen property Screen.MousePointer has a new behavior. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6BA9B8D2-2A32-4B6E-8D36-44949974A5B4"'
		System.Windows.Forms.Cursor.Current = System.Windows.Forms.Cursors.WaitCursor
		TomeAction = bdCreatePCPicture
		CreatePictureIndex = Index
		'UPGRADE_ISSUE: PictureBox method picTomeNew.Cls was not upgraded. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="CC4C7EC0-C903-48FC-ACCC-81861D12DA4A"'
        frmMain.picTomeNew = Nothing
        frmMain.picTomeNew.Invalidate()
		frmMain.ShowText((frmMain.picTomeNew), 30, 12, 519, 14, bdFontElixirWhite, "What do you look like?", True, False)
		frmMain.ShowText((frmMain.picTomeNew), 19, 45, 273, 14, bdFontElixirWhite, "Available Portraits", True, False)
		' Show Kingdom Name
		frmMain.ShowText((frmMain.picTomeNew), 332, 48, 217, 15, bdFontElixirBlack, (CreatureNow.Name), True, False)
		' Show basic stats
		frmMain.ShowText((frmMain.picTomeNew), 324, 84, 68, 14, bdFontNoxiousBlack, "Str " & CreatureNow.Strength, True, False)
		frmMain.ShowText((frmMain.picTomeNew), 409, 76, 68, 14, bdFontNoxiousBlack, "Int " & CreatureNow.Will, True, False)
		frmMain.ShowText((frmMain.picTomeNew), 486, 84, 68, 14, bdFontNoxiousBlack, "Agl " & CreatureNow.Agility, True, False)
		frmMain.ShowText((frmMain.picTomeNew), 328, 308, 220, 14, bdFontNoxiousBlack, "Health " & CreatureNow.HPNow & " (" & 100 * Int(Greatest((CreatureNow.HPNow), 0) / Greatest((CreatureNow.HPMax), 1)) & "%)", True, False)
		frmMain.ShowText((frmMain.picTomeNew), 328, 328, 220, 14, bdFontNoxiousBlack, "Defense " & CreatureNow.Defense, True, False)
		' Show Male/Female
		If CreatureNow.Male = True Then
			'UPGRADE_ISSUE: PictureBox property picMisc.hdc was not upgraded. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="CC4C7EC0-C903-48FC-ACCC-81861D12DA4A"'
			'UPGRADE_ISSUE: PictureBox property picTomeNew.hdc was not upgraded. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="CC4C7EC0-C903-48FC-ACCC-81861D12DA4A"'
            rc = BBlt(frmMain.picTomeNew, frmMain.picMisc, 52, 331, 18, 18, 18, 18, SRCCOPY)
			'UPGRADE_ISSUE: PictureBox property picMisc.hdc was not upgraded. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="CC4C7EC0-C903-48FC-ACCC-81861D12DA4A"'
			'UPGRADE_ISSUE: PictureBox property picTomeNew.hdc was not upgraded. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="CC4C7EC0-C903-48FC-ACCC-81861D12DA4A"'
            rc = BBlt(frmMain.picTomeNew, frmMain.picMisc, 166, 331, 18, 18, 0, 18, SRCCOPY)
		Else
			'UPGRADE_ISSUE: PictureBox property picMisc.hdc was not upgraded. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="CC4C7EC0-C903-48FC-ACCC-81861D12DA4A"'
			'UPGRADE_ISSUE: PictureBox property picTomeNew.hdc was not upgraded. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="CC4C7EC0-C903-48FC-ACCC-81861D12DA4A"'
            rc = BBlt(frmMain.picTomeNew, frmMain.picMisc, 52, 331, 18, 18, 0, 18, SRCCOPY)
			'UPGRADE_ISSUE: PictureBox property picMisc.hdc was not upgraded. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="CC4C7EC0-C903-48FC-ACCC-81861D12DA4A"'
			'UPGRADE_ISSUE: PictureBox property picTomeNew.hdc was not upgraded. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="CC4C7EC0-C903-48FC-ACCC-81861D12DA4A"'
            rc = BBlt(frmMain.picTomeNew, frmMain.picMisc, 166, 331, 18, 18, 18, 18, SRCCOPY)
		End If
		frmMain.ShowText((frmMain.picTomeNew), 70, 331, 64, 14, bdFontElixirWhite, "Male", True, False)
		frmMain.ShowText((frmMain.picTomeNew), 186, 331, 80, 14, bdFontElixirWhite, "Female", True, False)
		' Show pictures
		For c = 1 To Least(KingdomNow.PictureFiles.Count(), 9)
			'UPGRADE_WARNING: Couldn't resolve default property of object KingdomNow.PictureFiles().Text. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
			CreatureNow.PictureFile = BreakText(KingdomNow.PictureFiles.Item(c).Text, 1)
			'UPGRADE_WARNING: Couldn't resolve default property of object KingdomNow.PictureFiles().Text. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
			If IsNumeric(BreakText(KingdomNow.PictureFiles.Item(c).Text, 2)) Then
				'UPGRADE_WARNING: Couldn't resolve default property of object KingdomNow.PictureFiles().Text. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
				CreatureNow.FaceLeft = CShort(BreakText(KingdomNow.PictureFiles.Item(c).Text, 2))
			End If
			'UPGRADE_WARNING: Couldn't resolve default property of object KingdomNow.PictureFiles().Text. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
			If IsNumeric(BreakText(KingdomNow.PictureFiles.Item(c).Text, 3)) Then
				'UPGRADE_WARNING: Couldn't resolve default property of object KingdomNow.PictureFiles().Text. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
				CreatureNow.FaceTop = CShort(BreakText(KingdomNow.PictureFiles.Item(c).Text, 3))
			End If
			frmMain.LoadCreaturePic(CreatureNow)
			' Load and show picture
			'UPGRADE_ISSUE: PictureBox property picMisc.hdc was not upgraded. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="CC4C7EC0-C903-48FC-ACCC-81861D12DA4A"'
			'UPGRADE_ISSUE: PictureBox property picTomeNew.hdc was not upgraded. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="CC4C7EC0-C903-48FC-ACCC-81861D12DA4A"'
            rc = BBlt(frmMain.picTomeNew, frmMain.picMisc, 25 + (n Mod 3) * 84, Int(n / 3) * 86 + 70, 71, 80, 0, 36, SRCCOPY)
			'UPGRADE_ISSUE: PictureBox property picFaces.hdc was not upgraded. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="CC4C7EC0-C903-48FC-ACCC-81861D12DA4A"'
			'UPGRADE_ISSUE: PictureBox property picTomeNew.hdc was not upgraded. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="CC4C7EC0-C903-48FC-ACCC-81861D12DA4A"'
            rc = BBlt(frmMain.picTomeNew, frmMain.picFaces, 28 + (n Mod 3) * 84, Int(n / 3) * 86 + 73, 66, 76, bdFaceMin + CreatureNow.Pic * 66, 0, SRCCOPY)
			' Show selector
			If c = CreatePictureIndex Then
				'UPGRADE_ISSUE: PictureBox property picFaces.hdc was not upgraded. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="CC4C7EC0-C903-48FC-ACCC-81861D12DA4A"'
				'UPGRADE_ISSUE: PictureBox property picTomeNew.hdc was not upgraded. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="CC4C7EC0-C903-48FC-ACCC-81861D12DA4A"'
                rc = BBlt(frmMain.picTomeNew, frmMain.picFaces, 28 + (n Mod 3) * 84, Int(n / 3) * 86 + 73, 66, 76, bdFaceSelect + 66, 0, SRCAND)
				'UPGRADE_ISSUE: PictureBox property picFaces.hdc was not upgraded. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="CC4C7EC0-C903-48FC-ACCC-81861D12DA4A"'
				'UPGRADE_ISSUE: PictureBox property picTomeNew.hdc was not upgraded. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="CC4C7EC0-C903-48FC-ACCC-81861D12DA4A"'
                rc = BBlt(frmMain.picTomeNew, frmMain.picFaces, 28 + (n Mod 3) * 84, Int(n / 3) * 86 + 73, 66, 76, bdFaceSelect, 0, SRCPAINT)
			End If
			n = n + 1
		Next c
		' Load Clean Creature Picture (full size miniature)
		'UPGRADE_WARNING: Couldn't resolve default property of object KingdomNow.PictureFiles().Text. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
		CreatureNow.PictureFile = BreakText(KingdomNow.PictureFiles.Item(Index).Text, 1)
		'UPGRADE_WARNING: Couldn't resolve default property of object KingdomNow.PictureFiles().Text. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
		If IsNumeric(BreakText(KingdomNow.PictureFiles.Item(Index).Text, 2)) Then
			'UPGRADE_WARNING: Couldn't resolve default property of object KingdomNow.PictureFiles().Text. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
			CreatureNow.FaceLeft = CShort(BreakText(KingdomNow.PictureFiles.Item(Index).Text, 2))
		End If
		'UPGRADE_WARNING: Couldn't resolve default property of object KingdomNow.PictureFiles().Text. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
		If IsNumeric(BreakText(KingdomNow.PictureFiles.Item(Index).Text, 3)) Then
			'UPGRADE_WARNING: Couldn't resolve default property of object KingdomNow.PictureFiles().Text. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
			CreatureNow.FaceTop = CShort(BreakText(KingdomNow.PictureFiles.Item(Index).Text, 3))
		End If
		Dim Filename As String
		'UPGRADE_WARNING: Couldn't resolve default property of object Tome.FullPath. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
		'UPGRADE_WARNING: Dir has a new behavior. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="9B7D5ADD-D8FE-4819-A36C-6DEDAF088CC7"'
        Filename = Dir(tome.FullPath & "\" & CreatureNow.PictureFile)
		If Filename = "" Then
			'UPGRADE_WARNING: Couldn't resolve default property of object Tome.FullPath. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
			'UPGRADE_WARNING: Dir has a new behavior. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="9B7D5ADD-D8FE-4819-A36C-6DEDAF088CC7"'
            Filename = Dir(tome.FullPath & "\creatures\" & CreatureNow.PictureFile)
			If Filename = "" Then
				'            FileName = gAppPath & "\data\graphics\creatures\" & CreatureNow.PictureFile
				Filename = gDataPath & "\graphics\creatures\" & CreatureNow.PictureFile
			Else
				'UPGRADE_WARNING: Couldn't resolve default property of object Tome.FullPath. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
                Filename = tome.FullPath & "\creatures\" & CreatureNow.PictureFile
			End If
		Else
			'UPGRADE_WARNING: Couldn't resolve default property of object Tome.FullPath. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
            Filename = tome.FullPath & "\" & CreatureNow.PictureFile
		End If
		ReadBitmapFile(Filename, bmMons, hMemMons, TransparentRGB)
		'ReadBitmapFile gAppPath & "\data\graphics\creatures\" & CreatureNow.PictureFile, bmMons, hMemMons, TransparentRGB
		' Make a copy of the current palette for the picture
		'UPGRADE_WARNING: Couldn't resolve default property of object bmBlack. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
		bmBlack = bmMons
		' Then change Pure Blue to Pure Black
		ChangeColor(bmBlack, TransparentRGB, 0, 0, 0)
		MakeMask(bmMons, bmMask, TransparentRGB)
		' Paint bitmap to picture box using converted palette
		lpMem = GlobalLock(hMemMons)
		Width = bmMons.bmiHeader.biWidth
		Height = bmMons.bmiHeader.biHeight
		Size = 1
		If Width > 175 Or Height > 175 Then
			Size = 175 / Width
			If Height * Size > 175 Then
				Size = 175 / Height
			End If
		End If
		X = 336 + (200 - Width * Size) / 2 : Y = 104 + (188 - Height * Size) / 2
		'UPGRADE_ISSUE: PictureBox property picTomeNew.hdc was not upgraded. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="CC4C7EC0-C903-48FC-ACCC-81861D12DA4A"'
        rc = SetSBMode(frmMain.picTomeNew, 3)
		'UPGRADE_ISSUE: PictureBox property picTomeNew.hdc was not upgraded. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="CC4C7EC0-C903-48FC-ACCC-81861D12DA4A"'
        rc = SDIBits(frmMain.picTomeNew, X, Y, Width * Size, Height * Size, 0, 0, Width, Height, lpMem, bmMask, DIB_RGB_COLORS, SRCAND)
		'UPGRADE_ISSUE: PictureBox property picTomeNew.hdc was not upgraded. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="CC4C7EC0-C903-48FC-ACCC-81861D12DA4A"'
        rc = SDIBits(frmMain.picTomeNew, X, Y, Width * Size, Height * Size, 0, 0, Width, Height, lpMem, bmBlack, DIB_RGB_COLORS, SRCPAINT)
		' Release memory
		rc = GlobalUnlock(hMemMons)
		rc = GlobalFree(hMemMons)
		' Set up buttons
		frmMain.ShowButton((frmMain.picTomeNew), 380, 364, "Back", False)
		frmMain.ShowButton((frmMain.picTomeNew), 476, 364, "Next", False)
		' Refresh the screen
		frmMain.picTomeNew.Refresh()
		frmMain.picTomeNew.BringToFront()
		frmMain.picTomeNew.Visible = True
		'UPGRADE_WARNING: Screen property Screen.MousePointer has a new behavior. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6BA9B8D2-2A32-4B6E-8D36-44949974A5B4"'
		System.Windows.Forms.Cursor.Current = System.Windows.Forms.Cursors.Default
	End Sub
	
	Private Sub CreatePCSkills(ByRef Index As Short)
		Dim n, c, SkillPointsUsed As Short
		Dim TriggerX As Trigger
		Dim Found As Short
		Dim rc As Integer
		TomeAction = bdCreatePCSkills
		CreateSkillsIndex = Index
		'UPGRADE_ISSUE: PictureBox method picTomeNew.Cls was not upgraded. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="CC4C7EC0-C903-48FC-ACCC-81861D12DA4A"'
        frmMain.picTomeNew = Nothing
        frmMain.picTomeNew.Invalidate()
		frmMain.ShowText((frmMain.picTomeNew), 30, 12, 519, 14, bdFontElixirWhite, "What skills do you have?", True, False)
		frmMain.ShowText((frmMain.picTomeNew), 32, 45, 156, 14, bdFontElixirWhite, KingdomNow.Name & " Skills", False, False)
		frmMain.ShowText((frmMain.picTomeNew), 217, 45, 50, 14, bdFontElixirWhite, "Cost", True, False)
		' Total up cost of Skills picked
		SkillPointsUsed = 0
		For c = 1 To KingdomNow.Skills.Count()
			'UPGRADE_WARNING: Couldn't resolve default property of object KingdomNow.Skills(c).TempSkill. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
			If KingdomNow.Skills.Item(c).TempSkill > 0 Then
				'UPGRADE_WARNING: Couldn't resolve default property of object KingdomNow.Skills().TempSkill. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
				SkillPointsUsed = SkillPointsUsed + KingdomNow.Skills.Item(c).TempSkill
			End If
		Next c
		' List available and selected skills
		n = 0 : Found = False
		For c = ScrollTop To Least(ScrollTop + 9, KingdomNow.Skills.Count())
			' Show clicked if chosen for Creature
			'UPGRADE_WARNING: Couldn't resolve default property of object KingdomNow.Skills(c).TempSkill. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
			If KingdomNow.Skills.Item(c).TempSkill > 0 Then
				'UPGRADE_ISSUE: PictureBox property picMisc.hdc was not upgraded. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="CC4C7EC0-C903-48FC-ACCC-81861D12DA4A"'
				'UPGRADE_ISSUE: PictureBox property picTomeNew.hdc was not upgraded. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="CC4C7EC0-C903-48FC-ACCC-81861D12DA4A"'
                rc = BBlt(frmMain.picTomeNew, frmMain.picMisc, 26, 73 + n * 25, 18, 18, 18, 18, SRCCOPY)
			Else
				'UPGRADE_ISSUE: PictureBox property picMisc.hdc was not upgraded. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="CC4C7EC0-C903-48FC-ACCC-81861D12DA4A"'
				'UPGRADE_ISSUE: PictureBox property picTomeNew.hdc was not upgraded. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="CC4C7EC0-C903-48FC-ACCC-81861D12DA4A"'
                rc = BBlt(frmMain.picTomeNew, frmMain.picMisc, 26, 73 + n * 25, 18, 18, 0, 18, SRCCOPY)
			End If
			' Show Skill text
			If c = Index Then
				'UPGRADE_WARNING: Couldn't resolve default property of object KingdomNow.Skills().Name. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
				frmMain.ShowText((frmMain.picTomeNew), 53, 76 + n * 25, 156, 14, bdFontNoxiousGold, KingdomNow.Skills.Item(c).Name, False, False)
				'UPGRADE_WARNING: Couldn't resolve default property of object KingdomNow.Skills().Turns. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
				frmMain.ShowText((frmMain.picTomeNew), 217, 76 + n * 25, 50, 14, bdFontNoxiousGold, KingdomNow.Skills.Item(c).Turns, True, False)
			Else
				'UPGRADE_WARNING: Couldn't resolve default property of object KingdomNow.Skills().Name. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
				frmMain.ShowText((frmMain.picTomeNew), 53, 76 + n * 25, 156, 14, bdFontNoxiousWhite, KingdomNow.Skills.Item(c).Name, False, False)
				'UPGRADE_WARNING: Couldn't resolve default property of object KingdomNow.Skills().Turns. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
				frmMain.ShowText((frmMain.picTomeNew), 217, 76 + n * 25, 50, 14, bdFontNoxiousWhite, KingdomNow.Skills.Item(c).Turns, True, False)
			End If
			n = n + 1
		Next c
		' Show clicked Skill description
		If Index > 0 And Index <= KingdomNow.Skills.Count() Then
			'UPGRADE_WARNING: Couldn't resolve default property of object KingdomNow.Skills().Name. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
			frmMain.ShowText((frmMain.picTomeNew), 332, 48, 217, 15, bdFontElixirBlack, KingdomNow.Skills.Item(Index).Name, True, False)
			'UPGRADE_WARNING: Couldn't resolve default property of object KingdomNow.Skills().Comments. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
			frmMain.ShowText((frmMain.picTomeNew), 328, 80, 220, 190, bdFontNoxiousBlack, KingdomNow.Skills.Item(Index).Comments, False, False)
		End If
		' Show Skill Points Left
		frmMain.ShowText((frmMain.picTomeNew), 32, 332, 232, 15, bdFontElixirWhite, "Skill Points " & CreatureNow.SkillPoints - SkillPointsUsed, True, False)
		' Set up buttons
		'UPGRADE_ISSUE: frmMain.picTomeNew was upgraded to a Panel, and cannot be coerced to a PictureBox. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="0FF1188E-27C0-4FED-842D-159C65894C9B"'
		frmMain.ScrollBarShow((frmMain.picTomeNew), 276, 62, 267, ScrollTop, KingdomNow.Skills.Count() - 9, 0)
		frmMain.ShowButton((frmMain.picTomeNew), 380, 364, "Back", False)
		frmMain.ShowButton((frmMain.picTomeNew), 476, 364, "Next", False)
		' Refresh the screen
		frmMain.picTomeNew.Refresh()
		frmMain.picTomeNew.BringToFront()
		frmMain.picTomeNew.Visible = True
	End Sub
	
	Public Sub CreatePCNameEnter(ByRef KeyAscii As Short)
		Select Case KeyAscii
			Case System.Windows.Forms.Keys.Back, System.Windows.Forms.Keys.Left ' BackSpace
				If Len(CreateNameNew) > 0 Then
					CreateNameNew = Left(CreateNameNew, Len(CreateNameNew) - 1)
				End If
			Case 32 ' Space
				If Len(CreateNameNew) < 20 Then
					CreateNameNew = CreateNameNew & " "
				End If
			Case 48 To 57, 65 To 90, 97 To 122
				If Len(CreateNameNew) < 20 Then
					CreateNameNew = CreateNameNew & Chr(KeyAscii)
				End If
		End Select
		'UPGRADE_ISSUE: PictureBox method picCreateName.Cls was not upgraded. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="CC4C7EC0-C903-48FC-ACCC-81861D12DA4A"'
        frmMain.picCreateName = Nothing
        frmMain.picCreateName.Invalidate()
		frmMain.ShowText((frmMain.picCreateName), 6, 7, frmMain.picCreateName.Width - 12, 14, bdFontElixirWhite, CreateNameNew & "\", False, False)
		frmMain.picCreateName.Focus()
	End Sub
	
	'UPGRADE_NOTE: Tome was upgraded to Tome_Renamed. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="A9E4979A-37FA-4718-9994-97DD76ED70A7"'
	Public Sub CreatePCName(ByRef Index As Short)
		Dim Tome_Renamed As Object
		Dim c, n As Short
		Dim rc As Integer
		Dim Size As Double
		Dim Width, Height As Short
		Dim X, Y As Short
		'UPGRADE_WARNING: Arrays in structure bmMons may need to be initialized before they can be used. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="814DF224-76BD-4BB4-BFFB-EA359CB9FC48"'
		Dim bmMons As BITMAPINFO
		Dim hMemMons As Integer
		'UPGRADE_WARNING: Arrays in structure bmBlack may need to be initialized before they can be used. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="814DF224-76BD-4BB4-BFFB-EA359CB9FC48"'
		Dim bmBlack As BITMAPINFO
		'UPGRADE_WARNING: Arrays in structure bmMask may need to be initialized before they can be used. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="814DF224-76BD-4BB4-BFFB-EA359CB9FC48"'
		Dim bmMask As BITMAPINFO
		Dim lpMem, TransparentRGB As Integer
		TomeAction = bdCreatePCName
		CreateNameIndex = Index
		'UPGRADE_ISSUE: PictureBox method picTomeNew.Cls was not upgraded. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="CC4C7EC0-C903-48FC-ACCC-81861D12DA4A"'
        frmMain.picTomeNew = Nothing
        frmMain.picTomeNew.Invalidate()
		frmMain.ShowText((frmMain.picTomeNew), 30, 12, 519, 14, bdFontElixirWhite, "Name and Save your character", True, False)
		frmMain.ShowText((frmMain.picTomeNew), 19, 45, 273, 14, bdFontElixirWhite, "Suggested Names", True, False)
		frmMain.ShowText((frmMain.picTomeNew), 332, 48, 217, 15, bdFontElixirBlack, (CreatureNow.Name), True, False)
		' List Suggested Names
		For c = 0 To WorldNow.Kingdoms.Count() - 1 ' [Titi 2.4.7] was 11
			If c = Index Then
				frmMain.ShowText((frmMain.picTomeNew), 24, 68 + c * 22, 245, 22, bdFontNoxiousGold, strKingdomNames(c), True, False)
			Else
				frmMain.ShowText((frmMain.picTomeNew), 24, 68 + c * 22, 245, 22, bdFontNoxiousWhite, strKingdomNames(c), True, False)
			End If
		Next c
		' Show basic stats
		frmMain.ShowText((frmMain.picTomeNew), 324, 84, 68, 14, bdFontNoxiousBlack, "Str " & CreatureNow.Strength, True, False)
		frmMain.ShowText((frmMain.picTomeNew), 409, 76, 68, 14, bdFontNoxiousBlack, "Int " & CreatureNow.Will, True, False)
		frmMain.ShowText((frmMain.picTomeNew), 486, 84, 68, 14, bdFontNoxiousBlack, "Agl " & CreatureNow.Agility, True, False)
		frmMain.ShowText((frmMain.picTomeNew), 328, 308, 220, 14, bdFontNoxiousBlack, "Health " & CreatureNow.HPNow & " (" & 100 * Int(Greatest((CreatureNow.HPNow), 0) / Greatest((CreatureNow.HPMax), 1)) & "%)", True, False)
		frmMain.ShowText((frmMain.picTomeNew), 328, 328, 220, 14, bdFontNoxiousBlack, "Defense " & CreatureNow.Defense, True, False)
		' Load Clean Creature Picture
		Dim Filename As String
		'UPGRADE_WARNING: Couldn't resolve default property of object Tome.FullPath. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
		'UPGRADE_WARNING: Dir has a new behavior. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="9B7D5ADD-D8FE-4819-A36C-6DEDAF088CC7"'
        Filename = Dir(tome.FullPath & "\" & CreatureNow.PictureFile)
		If Filename = "" Then
			'UPGRADE_WARNING: Couldn't resolve default property of object Tome.FullPath. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
			'UPGRADE_WARNING: Dir has a new behavior. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="9B7D5ADD-D8FE-4819-A36C-6DEDAF088CC7"'
            Filename = Dir(tome.FullPath & "\creatures\" & CreatureNow.PictureFile)
			If Filename = "" Then
				'            FileName = gAppPath & "\data\graphics\creatures\" & CreatureNow.PictureFile
				Filename = gDataPath & "\graphics\creatures\" & CreatureNow.PictureFile
			Else
				'UPGRADE_WARNING: Couldn't resolve default property of object Tome.FullPath. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
                Filename = tome.FullPath & "\creatures\" & CreatureNow.PictureFile
			End If
		Else
			'UPGRADE_WARNING: Couldn't resolve default property of object Tome.FullPath. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
            Filename = tome.FullPath & "\" & CreatureNow.PictureFile
		End If
		'ReadBitmapFile gAppPath & "\data\graphics\creatures\" & CreatureNow.PictureFile, bmMons, hMemMons, TransparentRGB
		ReadBitmapFile(Filename, bmMons, hMemMons, TransparentRGB)
		' Make a copy of the current palette for the picture
		'UPGRADE_WARNING: Couldn't resolve default property of object bmBlack. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
		bmBlack = bmMons
		' Then change Pure Blue to Pure Black
		ChangeColor(bmBlack, TransparentRGB, 0, 0, 0)
		MakeMask(bmMons, bmMask, TransparentRGB)
		' Paint bitmap to picture box using converted palette
		lpMem = GlobalLock(hMemMons)
		Width = bmMons.bmiHeader.biWidth
		Height = bmMons.bmiHeader.biHeight
		Size = 1
		If Width > 175 Or Height > 175 Then
			Size = 175 / Width
			If Height * Size > 175 Then
				Size = 175 / Height
			End If
		End If
		X = 336 + (200 - Width * Size) / 2 : Y = 104 + (188 - Height * Size) / 2
		'UPGRADE_ISSUE: PictureBox property picTomeNew.hdc was not upgraded. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="CC4C7EC0-C903-48FC-ACCC-81861D12DA4A"'
        rc = SetSBMode(frmMain.picTomeNew, 3)
		'UPGRADE_ISSUE: PictureBox property picTomeNew.hdc was not upgraded. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="CC4C7EC0-C903-48FC-ACCC-81861D12DA4A"'
        rc = SDIBits(frmMain.picTomeNew, X, Y, Width * Size, Height * Size, 0, 0, Width, Height, lpMem, bmMask, DIB_RGB_COLORS, SRCAND)
		'UPGRADE_ISSUE: PictureBox property picTomeNew.hdc was not upgraded. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="CC4C7EC0-C903-48FC-ACCC-81861D12DA4A"'
        rc = SDIBits(frmMain.picTomeNew, X, Y, Width * Size, Height * Size, 0, 0, Width, Height, lpMem, bmBlack, DIB_RGB_COLORS, SRCPAINT)
		' Release memory
		rc = GlobalUnlock(hMemMons)
		rc = GlobalFree(hMemMons)
		' Set up buttons
		frmMain.ShowButton((frmMain.picTomeNew), 380, 364, "Back", False)
		frmMain.ShowButton((frmMain.picTomeNew), 476, 364, "Save", False)
		' Refresh the screen
		CreatePCNameEnter(0)
		frmMain.picTomeNew.Refresh()
		frmMain.picTomeNew.BringToFront()
		frmMain.picTomeNew.Visible = True
	End Sub
	
	Public Sub CreatePCPickKingdom(ByRef AtX As Short, ByRef AtY As Short)
		Dim c As Short
		Dim rc As Integer
		Dim KingdomX As Kingdom
		For c = 1 To WorldNow.Kingdoms.Count()
			KingdomX = WorldNow.Kingdoms.Item(c)
			If PointIn(AtX, AtY, (KingdomX.Left_Renamed), (KingdomX.Top), (KingdomX.Width), (KingdomX.Height)) Then
				Exit For
			End If
		Next c
		If c <= WorldNow.Kingdoms.Count() Then
			Call PlayClickSnd(modIOFunc.ClickType.ifClick)
			CreatePCKingdom(c)
		End If
	End Sub
	
	Private Sub CreatePCSave()
		Dim Filename As String
		Dim c, NoFail As Short
		Dim CreatureX As Creature
		Dim TriggerX As Trigger
		Dim tmpfile As String
		On Error GoTo ErrorHandler
		'UPGRADE_WARNING: Screen property Screen.MousePointer has a new behavior. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6BA9B8D2-2A32-4B6E-8D36-44949974A5B4"'
		System.Windows.Forms.Cursor.Current = System.Windows.Forms.Cursors.WaitCursor
		CreatureX = New Creature
		CreatureX.Copy(CreatureNow)
		c = 0
		For	Each TriggerX In CreatureX.Triggers
			If TriggerX.TempSkill = 0 And TriggerX.TriggerType <> bdPostNewCharacter Then
				CreatureX.RemoveTrigger("T" & TriggerX.Index)
			Else
				c = c + TriggerX.TempSkill
			End If
		Next TriggerX
		' Fire Post-NewCharacter Triggers
		CreatureNow = CreatureX
		'NoFail = frmMain.FireTriggers(CreatureX, bdPostNewCharacter)
		NoFail = FireTriggers(CreatureX, bdPostNewCharacter)
		For	Each TriggerX In CreatureX.Triggers
			CreatureNow = CreatureX
			'NoFail = frmMain.FireTriggers(TriggerX, bdPostNewCharacter)
			NoFail = FireTriggers(TriggerX, bdPostNewCharacter)
		Next TriggerX
		CreatureX.SkillPoints = CreatureX.SkillPoints - c
		CreatureX.Name = Proper(CreateNameNew)
		CreatureX.FaceTop = CreatureNow.FaceTop
		CreatureX.FaceLeft = CreatureNow.FaceLeft
		On Error GoTo Err_Handler
		Filename = gAppPath & "\roster\" & WorldNow.Name & "\" & CreatureX.Name & ".rsc"
		If oFileSys.CheckExists(Filename, clsInOut.IOActionType.File) Then
			Call oFileSys.Delete(Filename, clsInOut.IOActionType.File, True)
		End If
		c = FreeFile
		FileOpen(c, Filename, OpenMode.Binary)
		CreatureX.SaveToFile(c)
		FileClose(c)
		'UPGRADE_WARNING: Screen property Screen.MousePointer has a new behavior. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6BA9B8D2-2A32-4B6E-8D36-44949974A5B4"'
		System.Windows.Forms.Cursor.Current = System.Windows.Forms.Cursors.Default
		Exit Sub
ErrorHandler: 
		Resume Next
Err_Handler: 
		oErr.logError("CreatePCSave()")
		Err.Clear()
	End Sub
End Module