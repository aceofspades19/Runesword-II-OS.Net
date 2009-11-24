Option Strict Off
Option Explicit On
Module modDungeonMaker
	
	Const bdMapAddWalls As Short = 1
	Const bdMapAddEnc As Short = 2
	
	Const bdTileNone As Short = 0
	Const bdTileLeftWall As Short = 1
	Const bdTileCornerWall As Short = 2
	Const bdTileFloor As Short = 3
	Const bdTileLeftArch As Short = 4
	Const bdTileLeftDoor As Short = 5
	Const bdTileExitUp As Short = 6
	Const bdTileExitDown As Short = 7
	Const bdTileFloorDeco As Short = 8
	Const bdTileWallDeco As Short = 9
	Const bdTileCornerArch As Short = 10
	
	Public Const bdEntryTown As Short = 0
	Public Const bdEntryWilderness As Short = 1
	Public Const bdEntryBuilding As Short = 2
	Public Const bdEntryDungeon As Short = 3
	Public Const bdEntryExit As Short = 0
	Public Const bdEntryExitUp As Short = 1
    Public Const bdEntryExitDown As Short = 2
    Private MapStack As ArrayList
    Private ThemeQue As ArrayList
	
	Private Sub SearchFolders(ByRef Extension As String, ByRef DirName As String, ByRef CollectionX As Collection)
		Dim Filename, LastFile As String
		Dim FactoidX As Factoid
		' Does a recursive search through a given directory for a given file extension.
		' Results are added to the given collection.
		'UPGRADE_WARNING: Dir has a new behavior. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="9B7D5ADD-D8FE-4819-A36C-6DEDAF088CC7"'
		Filename = Dir(DirName & "\*", FileAttribute.Directory)
		Do Until Filename = ""
			If Filename <> "." And Filename <> ".." Then
				If (GetAttr(DirName & "\" & Filename) And FileAttribute.Directory) = FileAttribute.Directory Then
					SearchFolders(Extension, DirName & "\" & Filename, CollectionX)
				End If
			End If
			LastFile = Filename
			'UPGRADE_WARNING: Dir has a new behavior. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="9B7D5ADD-D8FE-4819-A36C-6DEDAF088CC7"'
			Filename = Dir(DirName & "\*", FileAttribute.Directory)
			Do Until Filename = LastFile
				'UPGRADE_WARNING: Dir has a new behavior. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="9B7D5ADD-D8FE-4819-A36C-6DEDAF088CC7"'
				Filename = Dir()
			Loop 
			'UPGRADE_WARNING: Dir has a new behavior. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="9B7D5ADD-D8FE-4819-A36C-6DEDAF088CC7"'
			Filename = Dir()
		Loop 
		'UPGRADE_WARNING: Dir has a new behavior. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="9B7D5ADD-D8FE-4819-A36C-6DEDAF088CC7"'
		Filename = Dir(DirName & "\" & Extension, FileAttribute.Normal)
		Do Until Filename = ""
			FactoidX = New Factoid
			FactoidX.Text = DirName & "\" & Filename
			FactoidX.Index = CollectionX.Count() + 1
			CollectionX.Add(FactoidX, "F" & FactoidX.Index)
			'UPGRADE_WARNING: Dir has a new behavior. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="9B7D5ADD-D8FE-4819-A36C-6DEDAF088CC7"'
			Filename = Dir()
		Loop 
	End Sub
	
	Public Sub UberWizMainMap(ByRef UberWizX As UberWizard, ByRef Index As Short)
		' Pass in a MapSketch Index and it is made the Main Map
		Dim MapSketchX As MapSketch
		Dim EntryPointX As EntryPoint
		' Clear out all MapSketchs and EntryPoints
		For	Each MapSketchX In UberWizX.MapSketchs
			MapSketchX.IsUsed = False
			For	Each EntryPointX In MapSketchX.EntryPoints
				EntryPointX.IsUsed = False
			Next EntryPointX
		Next MapSketchX
		' Set the MainMap
		UberWizX.MainMapIndex = Index
		UberWizX.MainMap.IsUsed = True
		UberWizX.MainMap.IsSelected = True
		' Main Map is set to UberWiz's difficulty level
		UberWizX.MainMap.PartySize = UberWizX.TomePartySize
		UberWizX.MainMap.PartyAvgLevel = UberWizX.TomePartyAvgLevel
		' Find EntryPoint in MainMap for Tome
		For	Each EntryPointX In UberWizX.MainMap.EntryPoints
			If (EntryPointX.MapStyle = bdEntryTown Or EntryPointX.MapStyle = bdEntryWilderness) And (EntryPointX.Style = bdEntryExitUp Or EntryPointX.Style = bdEntryExit) Then
				UberWizX.AreaIndex = 1
				UberWizX.MapIndex = UberWizX.MainMap.Index
				UberWizX.EntryIndex = EntryPointX.Index
				EntryPointX.IsUsed = True
				Exit For
			End If
		Next EntryPointX
		' All other Maps from the same Tome are used by association
		If UberWizX.MainMap.TomeIndex > 0 Then
			For	Each MapSketchX In UberWizX.MapSketchs
				' If the Maps are in the same Tome
				If UberWizX.MainMap.TomeIndex = MapSketchX.TomeIndex Then
					MapSketchX.IsUsed = True
					' Every EntryPoint that is not Random is used by association
					For	Each EntryPointX In MapSketchX.EntryPoints
						If EntryPointX.AreaIndex > -1 Then
							EntryPointX.IsUsed = True
						End If
					Next EntryPointX
				End If
			Next MapSketchX
		End If
	End Sub
	
	Public Sub UberWizConnectMap(ByRef UberWizX As UberWizard, ByRef MapX As MapSketch, ByRef SizeLimit As Integer, ByRef Available As Short, ByRef Distance As Short)
		Dim ConnectQue As Object
		Dim MapQue As Object
		' This builds out from MapX until SizeLimit is reached or exhaust list of Entry Points
		' or pool of potential Maps
		Dim EntryPointX, EntryPointZ As EntryPoint
		Dim MapZ As MapSketch
		Dim ThemeSketchX As ThemeSketch
		Dim y, i, c, x, Found As Short
		' If reached SizeLimit then halt
		If SizeLimit < 1 Then
			Exit Sub
		End If
		' Que up unused EntryPoints
		Dim EntryQue(MapX.EntryPoints.Count() - 1) As Short
		c = 0
		For	Each EntryPointX In MapX.EntryPoints
			' If EntryPoint is not already connected
			If EntryPointX.IsUsed = False Then
				EntryQue(c) = EntryPointX.Index
				c = c + 1
			End If
		Next EntryPointX
		' Shuffle
		For i = 0 To c - 1
			x = Int(Rnd() * c)
			y = EntryQue(x)
			EntryQue(x) = EntryQue(i)
			EntryQue(i) = y
		Next i
		Available = Available + c + 1
		' Loop Through EntryPoints
		For i = 0 To c - 1
			EntryPointX = MapX.EntryPoints.Item("P" & EntryQue(i))
			Found = False
			' Que up potential connecting Maps
			'UPGRADE_ISSUE: As Integer was removed from ReDim MapQue(UberWizX.MapSketchs.Count - 1) statement. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="19AFCB41-AA8E-4E6B-A441-A3E802E5FD64"'
			ReDim MapQue(UberWizX.MapSketchs.Count() - 1)
			'UPGRADE_ISSUE: As Integer was removed from ReDim ConnectQue(UberWizX.MapSketchs.Count - 1) statement. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="19AFCB41-AA8E-4E6B-A441-A3E802E5FD64"'
			ReDim ConnectQue(UberWizX.MapSketchs.Count() - 1)
			x = 0
			For	Each MapZ In UberWizX.MapSketchs
				If MapZ.IsSelected = True And MapZ.IsUsed = False And EntryPointX.MapStyle = MapZ.MapStyle Then
					For	Each EntryPointZ In MapZ.EntryPoints
						If EntryPointZ.IsUsed = False And EntryPointX.ToStyle = EntryPointZ.Style Then
							'UPGRADE_WARNING: Couldn't resolve default property of object MapQue(). Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
							MapQue(x) = MapZ.Index
							'UPGRADE_WARNING: Couldn't resolve default property of object ConnectQue(). Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
							ConnectQue(x) = EntryPointZ.Index
							x = x + 1
							Exit For
						End If
					Next EntryPointZ
				End If
			Next MapZ
			' Connect to a random one
			If x > 0 Then
				x = Int(Rnd() * x)
				'UPGRADE_WARNING: Couldn't resolve default property of object MapQue(). Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
				MapZ = UberWizX.MapSketchs.Item("M" & MapQue(x))
				'UPGRADE_WARNING: Couldn't resolve default property of object ConnectQue(). Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
				EntryPointZ = MapZ.EntryPoints.Item("P" & ConnectQue(x))
				MapX.IsUsed = True
				EntryPointX.IsUsed = True
				EntryPointX.AreaIndex = 1
				EntryPointX.MapIndex = MapZ.Index
				EntryPointX.EntryIndex = EntryPointZ.Index
				' The connected map and entry point is used by association
				MapZ.IsUsed = True
				EntryPointZ.IsUsed = True
				EntryPointZ.AreaIndex = 1
				EntryPointZ.MapIndex = MapX.Index
				EntryPointZ.EntryIndex = EntryPointX.Index
				Available = Available - 1
				SizeLimit = SizeLimit - MapZ.TotalSize
				' Attached Map is raised in Difficulty
				MapZ.PartySize = MapX.PartySize
				MapZ.PartyAvgLevel = MapX.PartyAvgLevel
				If (Distance Mod 4) > Int(Rnd() * 4) Then
					MapZ.IncreaseDifficulty()
				End If
				' Adjust ThemeSketchs too
				For	Each ThemeSketchX In MapZ.ThemeSketchs
					ThemeSketchX.PartySize = MapZ.PartySize
					ThemeSketchX.PartyAvgLevel = MapZ.PartyAvgLevel
				Next ThemeSketchX
				' Go connect to new Map as well
				If Int(Rnd() * 100) > Available * 20 Then
					UberWizConnectMap(UberWizX, MapZ, SizeLimit, Available, Distance + 1)
				End If
			End If
		Next i
	End Sub
	
	Public Sub UberWizCopyThemes(ByRef UberWizX As UberWizard, ByRef MapSketchX As MapSketch)
		Dim ThemeSketchX As ThemeSketch
		For	Each ThemeSketchX In UberWizX.ThemeSketchs
			If ThemeSketchX.IsSelected = True Then
				MapSketchX.AddThemeSketch.Copy(ThemeSketchX)
				ThemeSketchX.IsSelected = False
			End If
		Next ThemeSketchX
		MapSketchX.NeedsThemes = False
	End Sub
	
	Public Sub UberWizScatterQuests(ByRef UberWizX As UberWizard)
		Dim i, MapCnt, c, x As Short
		Dim MapSketchX As MapSketch
		Dim ThemeX, ThemeZ As Theme
		Dim ThemeSketchX As ThemeSketch
		Dim EncounterX As Encounter
		' Que up list of randomizable Maps
		Dim MapQue(UberWizX.MapSketchs.Count()) As Short
		MapCnt = 0
		For	Each MapSketchX In UberWizX.MapSketchs
			If MapSketchX.IsSelected = True And MapSketchX.GenerateUponEntry = True Then
				MapQue(MapCnt) = MapSketchX.Index
				MapCnt = MapCnt + 1
			End If
		Next MapSketchX
		' Load the Major Theme
		If UberWizX.MajorThemeIndex > 0 And MapCnt > 0 Then
			ThemeX = New Theme
			c = FreeFile
			'UPGRADE_WARNING: Couldn't resolve default property of object UberWizX.ThemeSketchs(R & UberWizX.MajorThemeIndex).FullPath. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
			FileOpen(c, UberWizX.ThemeSketchs.Item("R" & UberWizX.MajorThemeIndex).FullPath, OpenMode.Binary)
			ThemeX.ReadFromFile(c)
			FileClose(c)
			' Scatter the Encounters from the Major Theme to new Themes on the Maps
			For	Each EncounterX In ThemeX.Encounters
				' Build an Encounter from the Major Theme and copy to a random MapSketch
				'UPGRADE_WARNING: Couldn't resolve default property of object UberWizX.MapSketchs().AddThemeQuest. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
				ThemeZ = UberWizX.MapSketchs.Item("M" & MapQue(Int(Rnd() * MapCnt))).AddThemeQuest
				ThemeZ.Name = ThemeX.Name
				ThemeZ.AddEncounter.Copy(EncounterX)
			Next EncounterX
		End If
		' Load the Quest Themes
		If MapCnt > 0 Then
			For	Each ThemeSketchX In UberWizX.ThemeSketchs
				If ThemeSketchX.IsSelected = True And ThemeSketchX.IsQuest = True Then
					ThemeX = New Theme
					c = FreeFile
					FileOpen(c, ThemeSketchX.FullPath, OpenMode.Binary)
					ThemeX.ReadFromFile(c)
					FileClose(c)
					' Scatter the Encounters from the Major Theme to new Themes on the Maps
					For	Each EncounterX In ThemeX.Encounters
						' Build an Encounter from the Major Theme and copy to a random MapSketch
						'UPGRADE_WARNING: Couldn't resolve default property of object UberWizX.MapSketchs().AddThemeQuest. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
						ThemeZ = UberWizX.MapSketchs.Item("M" & MapQue(Int(Rnd() * MapCnt))).AddThemeQuest
						ThemeZ.Name = ThemeX.Name
						ThemeZ.AddEncounter.Copy(EncounterX)
					Next EncounterX
				End If
			Next ThemeSketchX
		End If
	End Sub
	
	Public Sub UberWizFinish(ByRef UberWizX As UberWizard, ByRef UberIndex As Short, ByRef SetSize As Short)
		Dim x, c, i, y As Short
		Dim TomeSketchX As TomeSketch
		Dim MapSketchX As MapSketch
		Dim ThemeSketchX As ThemeSketch
		'UPGRADE_WARNING: Screen property Screen.MousePointer has a new behavior. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6BA9B8D2-2A32-4B6E-8D36-44949974A5B4"'
		System.Windows.Forms.Cursor.Current = System.Windows.Forms.Cursors.WaitCursor
		If UberIndex < 1 Then ' Load UberWiz
			InitUberWizMaps(UberWizX)
			InitUberWizThemes(UberWizX)
			InitUberWizCreatures(UberWizX)
			InitUberWizItems(UberWizX)
		End If
		' If no maps, then cannot continue
		If UberWizX.MapSketchs.Count() < 1 Then
			Exit Sub
		End If
		If UberIndex < 2 Then ' Choose Starting Tome
			' Count Potentials
			Dim MapStack(UberWizX.MapSketchs.Count()) As Short
			c = 0
			For	Each MapSketchX In UberWizX.MapSketchs
				If MapSketchX.IsMainMap = True Then
					MapStack(c) = MapSketchX.Index
					c = c + 1
				End If
			Next MapSketchX
			' Set up the MainMap
			modDungeonMaker.UberWizMainMap(UberWizX, MapStack(Int(Rnd() * c)))
		End If
		If UberIndex < 3 Then ' Edit Name of Tome
			If UberWizX.MainMap.TomeIndex > 0 Then
				TomeSketchX = UberWizX.TomeSketchs.Item("T" & UberWizX.MainMap.TomeIndex)
				UberWizX.TomeName = TomeSketchX.Name
				UberWizX.TomeComments = TomeSketchX.Comments
				If SetSize = True Then
					Select Case TomeSketchX.PartySize
						Case 0 To 1
							UberWizX.TomePartySize = 0
						Case 1 To 3
							UberWizX.TomePartySize = 1
						Case 4 To 6
							UberWizX.TomePartySize = 2
						Case Else
							UberWizX.TomePartySize = 3
					End Select
					Select Case TomeSketchX.PartyAvgLevel
						Case 0 To 3
							UberWizX.TomePartyAvgLevel = 0
						Case 4 To 6
							UberWizX.TomePartyAvgLevel = 1
						Case 7 To 10
							UberWizX.TomePartyAvgLevel = 2
						Case Else
							UberWizX.TomePartyAvgLevel = 3
					End Select
				End If
			Else
				UberWizX.TomeName = "Tome Wizard"
				UberWizX.TomeComments = "Generated by RuneSword Tome Wizard."
				If SetSize = True Then
					UberWizX.TomePartySize = 1
					UberWizX.TomePartyAvgLevel = 0
				End If
			End If
			' Set the Tome Size
			If SetSize = True Then
				Select Case Int(Rnd() * 7) + 1
					Case 1 ' Small / 8,000 Sq
						UberWizX.TotalSize = 10000 - Int(Rnd() * 4000)
					Case 2 To 4 ' Medium / 20,000 Sq
						UberWizX.TotalSize = 25000 - Int(Rnd() * 10000)
					Case 5 To 6 ' Large / 40,000 Sq
						UberWizX.TotalSize = 45000 - Int(Rnd() * 10000)
					Case 7 ' Huge / 60,000+
						UberWizX.TotalSize = 70000 - Int(Rnd() * 20000)
				End Select
			End If
		End If
		If UberIndex < 4 Then ' Choose Additional Maps
			' Count number of Maps not yet selected
			c = -1
			For	Each MapSketchX In UberWizX.MapSketchs
				If MapSketchX.IsSelected = False Then
					c = c + 1
				End If
			Next MapSketchX
			If c > -1 Then

				' Populate the stack
				i = 0
				For	Each MapSketchX In UberWizX.MapSketchs
					If MapSketchX.IsSelected = False Then
						MapStack(i) = MapSketchX.Index
						i = i + 1
					End If
				Next MapSketchX
				' Shuffle
				For i = 0 To c
					x = Int(Rnd() * (c + 1))
					y = MapStack(i)
					MapStack(i) = MapStack(x)
					MapStack(x) = y
				Next i
				' Select a few
				For i = 0 To Least(Int(UberWizX.TotalSize / 7000) * 5, c)
					'UPGRADE_WARNING: Couldn't resolve default property of object UberWizX.MapSketchs().IsSelected. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
					UberWizX.MapSketchs.Item("M" & MapStack(i)).IsSelected = True
				Next i
			End If
		End If
		If UberIndex < 5 Then ' Choose Size and Encounter Count
			For	Each MapSketchX In UberWizX.MapSketchs
				' If it's a Map Template, Selected and doesn't have Size set, then need to set it.
				If MapSketchX.IsSelected = True Then
					If MapSketchX.GenerateUponEntry = True And MapSketchX.Size = 0 Then
						Select Case Int(Rnd() * 75) + Int(UberWizX.TotalSize / 10000) * 5
							Case 0 To 14 ' Small
								MapSketchX.Size = 1
								MapSketchX.EncounterCount = 1
								MapSketchX.TotalSize = (Int(Rnd() * 16) + 16) * (Int(Rnd() * 16) + 16)
							Case 15 To 45 ' Medium
								MapSketchX.Size = 2
								MapSketchX.EncounterCount = Int(Rnd() * 2) + 2
								MapSketchX.TotalSize = (Int(Rnd() * 32) + 32) * (Int(Rnd() * 32) + 32)
							Case 46 To 85 ' Large
								MapSketchX.Size = 3
								MapSketchX.EncounterCount = Int(Rnd() * 2) + 3
								MapSketchX.TotalSize = (Int(Rnd() * 64) + 48) * (Int(Rnd() * 64) + 48)
							Case Else ' Huge
								MapSketchX.Size = 4
								MapSketchX.EncounterCount = 5
								MapSketchX.TotalSize = (Int(Rnd() * 96) + 64) * (Int(Rnd() * 96) + 64)
						End Select
					End If
				End If
			Next MapSketchX
		End If
		If UberIndex < 6 Then ' Choose Themes
			If UberWizX.ThemeSketchs.Count() > 0 Then
				Dim ThemeQue(UberWizX.ThemeSketchs.Count()) As Short
				' Que of Themes
				x = 0
				For	Each ThemeSketchX In UberWizX.ThemeSketchs
					If ThemeSketchX.IsMajorTheme = False And ThemeSketchX.IsQuest = False Then
						' And ((Abs(ThemeSketchX.Difficulty - UberWizX.Difficulty) < 4) Or (ThemeSketchX.Difficulty > 30 And UberWizX.Difficulty > 30)) Then
						ThemeQue(x) = ThemeSketchX.Index
						x = x + 1
					End If
				Next ThemeSketchX
				If x > 0 Then
					For	Each MapSketchX In UberWizX.MapSketchs
						' If found a Map that needs Themes, choose a few for it
						If MapSketchX.IsSelected = True And MapSketchX.NeedsThemes = True Then
							y = MapSketchX.EncounterCount
							For c = 0 To MapSketchX.Size - 1 + Int(Rnd() * MapSketchX.Size)
								i = ThemeQue(Int(Rnd() * x))
								'UPGRADE_WARNING: Couldn't resolve default property of object UberWizX.ThemeSketchs().Coverage. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
								y = y - UberWizX.ThemeSketchs.Item("R" & i).Coverage
								'UPGRADE_WARNING: Couldn't resolve default property of object UberWizX.ThemeSketchs().IsSelected. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
								UberWizX.ThemeSketchs.Item("R" & i).IsSelected = True
								' If Themes now cover entire Map, then exit
								If y < 1 Then
									Exit For
								End If
							Next c
							modDungeonMaker.UberWizCopyThemes(UberWizX, MapSketchX)
						End If
					Next MapSketchX
				End If
			End If
		End If
		If UberIndex < 7 Then ' Choose Major and Quest Themes
			If UberWizX.ThemeSketchs.Count() > 0 Then

				' Count Themes matching our difficulty
				x = 0
				For	Each ThemeSketchX In UberWizX.ThemeSketchs
					If ThemeSketchX.IsMajorTheme = True Then
						' And ((Abs(ThemeSketchX.Difficulty - UberWizX.Difficulty) < 4) Or (ThemeSketchX.Difficulty > 30 And UberWizX.Difficulty > 30)) Then
						ThemeQue(x) = ThemeSketchX.Index
						x = x + 1
					End If
				Next ThemeSketchX
				' Select a random Major Theme
				If x > 0 Then
					i = ThemeQue(Int(Rnd() * x))
					UberWizX.MajorThemeIndex = i
					'UPGRADE_WARNING: Couldn't resolve default property of object UberWizX.ThemeSketchs().IsSelected. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
					UberWizX.ThemeSketchs.Item("R" & i).IsSelected = True
				Else
					UberWizX.MajorThemeIndex = 0
				End If
				' Count Quest Themes matching our difficulty
				x = 0
				For	Each ThemeSketchX In UberWizX.ThemeSketchs
					If ThemeSketchX.IsQuest = True Then
						' And ((Abs(ThemeSketchX.Difficulty - UberWizX.Difficulty) < 4) Or (ThemeSketchX.Difficulty > 30 And UberWizX.Difficulty > 30)) Then
						ThemeQue(x) = ThemeSketchX.Index
						x = x + 1
					End If
				Next ThemeSketchX
				' Select Random number of Quest Themes
				If x > 0 Then
					For c = 0 To Int(Rnd() * (UberWizX.TotalSize / 2500))
						'UPGRADE_WARNING: Couldn't resolve default property of object UberWizX.ThemeSketchs().IsSelected. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
						UberWizX.ThemeSketchs.Item("R" & ThemeQue(Int(Rnd() * x))).IsSelected = True
					Next c
				End If
			End If
		End If
		If UberIndex < 8 And UberWizX.Creatures.Count() > 0 Then ' Choose additional Creatures
			' Pick some additional surprise Creatures based on EncounterPoints
			For c = 0 To Int(Rnd() * (UberWizX.TotalSize / 2500))
				x = Int(Rnd() * UberWizX.Creatures.Count()) + 1
				'UPGRADE_WARNING: Couldn't resolve default property of object UberWizX.Creatures().EncounterPoints. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
				If System.Math.Abs(UberWizX.Creatures.Item(x).EncounterPoints - UberWizX.MainMap.Difficulty) < 4 Then
					'UPGRADE_WARNING: Couldn't resolve default property of object UberWizX.Creatures().IsSelected. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
					UberWizX.Creatures.Item(x).IsSelected = True
				End If
			Next c
		End If
		If UberIndex < 9 And UberWizX.Items.Count() > 0 Then ' Choose additional Items
			' Pick some additional surprise Items based on EncounterPoints
			For c = 0 To Int(Rnd() * (UberWizX.TotalSize / 2500))
				x = Int(Rnd() * UberWizX.Items.Count()) + 1
				'UPGRADE_WARNING: Couldn't resolve default property of object UberWizX.Items().EncounterPoints. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
				If System.Math.Abs(UberWizX.Items.Item(x).EncounterPoints - UberWizX.MainMap.Difficulty) < 4 Then
					'UPGRADE_WARNING: Couldn't resolve default property of object UberWizX.Items().IsSelected. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
					UberWizX.Items.Item(x).IsSelected = True
				End If
			Next c
		End If
		'UPGRADE_WARNING: Screen property Screen.MousePointer has a new behavior. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6BA9B8D2-2A32-4B6E-8D36-44949974A5B4"'
		System.Windows.Forms.Cursor.Current = System.Windows.Forms.Cursors.Default
	End Sub
	
	Public Sub UberWizMakeMap(ByRef UberWizX As UberWizard, ByRef MapSketchX As MapSketch, ByRef AreaX As Area)
		Dim c, i As Short
		Dim PlotX As Plot
		Dim MapX As Map
		Dim ThemeX, ThemeZ As Theme
		Dim EntryPointX As EntryPoint
		Dim ThemeSketchX As ThemeSketch
		' Load up the Map
		If Right(MapSketchX.FullPath, 3) = "rsm" Then
			MapX = AreaX.Plot.AddMapWithIndex((MapSketchX.Index))
			i = MapX.Index
			c = FreeFile
			FileOpen(c, MapSketchX.FullPath, OpenMode.Binary)
			MapX.ReadFromFile(c)
			FileClose(c)
			MapX.Index = i
		Else
			PlotX = New Plot
			c = FreeFile
			FileOpen(c, MapSketchX.FullPath, OpenMode.Binary)
			PlotX.ReadFromFile(c)
			FileClose(c)
			MapX = AreaX.Plot.AddMapWithIndex((MapSketchX.Index))
			'UPGRADE_WARNING: Couldn't resolve default property of object PlotX.Maps(). Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
			MapX.Copy(PlotX.Maps.Item("M" & MapSketchX.MapIndexInArea))
		End If
		' Copy PartySize and PartyAvgLevel
		MapX.PartySize = MapSketchX.PartySize
		MapX.PartyAvgLevel = MapSketchX.PartyAvgLevel
		' Copy the Entry Points
		MapX.EntryPoints = MapSketchX.EntryPoints
		' Remove Entry Points that don't have a connection
		For	Each EntryPointX In MapSketchX.EntryPoints
			If EntryPointX.IsUsed = False Then
				' If sitting on an EntryPoint style tile, remove that first
				For i = 0 To 2
					' If the EntryPoint is on a valid spot on the Map
					If EntryPointX.MapX < MapX.Width And EntryPointX.MapY < MapX.Height Then
						c = MapX.Tile(i, EntryPointX.MapX, EntryPointX.MapY)
						' If the Map spot has a valid tile
						If c > 0 Then
							' And if the Tile's style is Exit Up or Exit Down then remove it
							'UPGRADE_WARNING: Couldn't resolve default property of object MapX.Tiles(L & c).Style. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
							If MapX.Tiles.Item("L" & c).Style = 6 Or MapX.Tiles.Item("L" & c).Style = 7 Then
								MapX.Tile(i, EntryPointX.MapX, EntryPointX.MapY) = 0
							End If
						End If
					End If
				Next i
				' Then remove EntryPoint itself
				MapX.RemoveEntryPoint("P" & EntryPointX.Index)
			End If
		Next EntryPointX
		' Add General Themes to Maps
		If MapSketchX.ThemeSketchs.Count() > 0 Then
			For	Each ThemeSketchX In MapSketchX.ThemeSketchs
				ThemeX = MapX.AddTheme
				i = ThemeX.Index
				c = FreeFile
				FileOpen(c, ThemeSketchX.FullPath, OpenMode.Binary)
				ThemeX.ReadFromFile(c)
				FileClose(c)
				ThemeX.Index = i
				ThemeX.PartySize = ThemeSketchX.PartySize
				ThemeX.PartyAvgLevel = ThemeSketchX.PartyAvgLevel
			Next ThemeSketchX
		End If
		' Add Quest Themes to Map
		If MapSketchX.ThemeQuests.Count() > 0 Then
			For	Each ThemeX In MapSketchX.ThemeQuests
				ThemeZ = MapX.AddTheme
				ThemeZ.Copy(ThemeX)
			Next ThemeX
		End If
		' Spin out Map
		If MapSketchX.GenerateUponEntry = True Then
			' Set up Size
			Select Case MapSketchX.Size
				Case 1 ' Small
					MapX.DefaultWidth = Int(Rnd() * 16) + 16 : MapX.DefaultHeight = Int(Rnd() * 16) + 16
				Case 2 ' Medium
					MapX.DefaultWidth = Int(Rnd() * 32) + 32 : MapX.DefaultHeight = Int(Rnd() * 32) + 32
				Case 3 ' Large
					MapX.DefaultWidth = Int(Rnd() * 64) + 48 : MapX.DefaultHeight = Int(Rnd() * 64) + 48
				Case 4 ' Huge
					MapX.DefaultWidth = Int(Rnd() * 96) + 64 : MapX.DefaultHeight = Int(Rnd() * 96) + 64
			End Select
			' Number of Encounters
			Select Case MapSketchX.EncounterCount
				Case 1 ' 1-8
					MapX.DefaultEncounters = Int(Rnd() * 5) + 5
				Case 2 ' 8-16
					MapX.DefaultEncounters = Int(Rnd() * 8) + 10
				Case 3 ' 16-32
					MapX.DefaultEncounters = Int(Rnd() * 16) + 16
				Case 4 ' 32-48
					MapX.DefaultEncounters = Int(Rnd() * 16) + 32
				Case 5 ' >48
					MapX.DefaultEncounters = Int(Rnd() * 16) + 48
			End Select
			' Make the Map
			modDungeonMaker.MakeMap(MapX, False)
		End If
	End Sub
	
	Public Sub InitUberWizMaps(ByRef UberWizX As UberWizard)
		Dim i, c, n As Short
		Dim CreatureX As Creature
		Dim TriggerX As Trigger
        Dim JournalX As Journal
        Dim FactoidX As Factoid
		Dim TomeX As Tome
		Dim AreaX As Area
		Dim MapX As Map
		Dim EncounterX As Encounter
		Dim TomeSketchX As TomeSketch
		Dim MapSketchX As MapSketch
		Dim EntryPointX As EntryPoint
		Dim MapSketchZ As MapSketch
		' Find all the Tome files available
		SearchFolders("*.tom", gAppPath & "\library", (UberWizX.TomeFiles))
		For c = 1 To UberWizX.TomeFiles.Count()
			' Read in the Tome
			TomeX = New Tome
			i = FreeFile
			'UPGRADE_WARNING: Couldn't resolve default property of object UberWizX.TomeFiles(c).Text. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
			FileOpen(i, UberWizX.TomeFiles.Item(c).Text, OpenMode.Binary)
			TomeX.ReadFromFile(i)
			FileClose(i)
			'UPGRADE_WARNING: Couldn't resolve default property of object UberWizX.TomeFiles().Text. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
			TomeX.FullPath = ClipPath(UberWizX.TomeFiles.Item(c).Text)
			' Build a TomeSketch for it
			TomeSketchX = UberWizX.AddTomeSketch
			TomeSketchX.Name = TomeX.Name
			TomeSketchX.FullPath = TomeX.FullPath
			TomeSketchX.Comments = TomeX.Comments
			TomeSketchX.PartySize = TomeX.PartySize
			TomeSketchX.PartyAvgLevel = TomeX.PartyAvgLevel
			For	Each CreatureX In TomeX.Creatures
				TomeSketchX.AddCreature.Copy(CreatureX)
			Next CreatureX
			For	Each TriggerX In TomeX.Triggers
				TomeSketchX.AddTrigger.Copy(TriggerX)
			Next TriggerX
			For	Each JournalX In TomeX.Journals
				'UPGRADE_WARNING: Couldn't resolve default property of object JournalX. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
				TomeSketchX.AddJournal.Copy(JournalX)
			Next JournalX
			For	Each FactoidX In TomeX.Factoids
				TomeSketchX.AddFactoid.Copy(FactoidX)
			Next FactoidX
			' Look at each Area in the Tome
			For	Each AreaX In TomeX.Areas
				' Read in all the Maps for the Area
				i = FreeFile
				AreaX.Plot = New Plot
				FileOpen(i, TomeX.FullPath & "\" & AreaX.Filename, OpenMode.Binary)
				AreaX.Plot.ReadFromFile(i)
				FileClose(i)
				' Add MapSketch to the UberWizard
				For	Each MapX In AreaX.Plot.Maps
					MapSketchX = UberWizX.AddMapSketch
					MapSketchX.MapName = MapX.Name
					MapSketchX.MapStyle = MapX.MapStyle
					MapSketchX.TotalSize = MapX.TotalSize
					MapSketchX.MapComments = MapX.Comments
					MapSketchX.FullPath = TomeX.FullPath & "\" & AreaX.Filename
					MapSketchX.MapIndexInArea = MapX.Index
					MapSketchX.AreaIndexInArea = AreaX.Index
					MapSketchX.TomeIndexInArea = TomeSketchX.Index
					MapSketchX.EntryPoints = MapX.EntryPoints
					MapSketchX.GenerateUponEntry = MapX.GenerateUponEntry
					MapSketchX.TomeIndex = TomeSketchX.Index
					' Determine if Map needs Themes and/or needs to Generate Themes
					MapSketchX.NeedsThemes = False
					MapSketchX.GenerateThemes = False
					If MapX.GenerateUponEntry = True Then
						MapSketchX.NeedsThemes = True
						MapSketchX.GenerateThemes = True
					Else
						For	Each EncounterX In MapX.Encounters
							If EncounterX.GenerateUponEntry = True Then
								MapSketchX.GenerateThemes = True
								If EncounterX.ParentTheme = 0 Then
									MapSketchX.NeedsThemes = True
								End If
							End If
						Next EncounterX
					End If
				Next MapX
			Next AreaX
			' Realign EntryPoints to new MapSketch indexes
			For	Each MapSketchX In UberWizX.MapSketchs
				' Only align Maps for the current Tome
				If MapSketchX.TomeIndexInArea = TomeSketchX.Index Then
					For	Each EntryPointX In MapSketchX.EntryPoints
						' If points to a valid Area
						If EntryPointX.AreaIndex > 0 Then
							' Find the Tome/Area/Map and set the new index
							For	Each MapSketchZ In UberWizX.MapSketchs
								' Only pick from Maps in the current Tome
								If MapSketchZ.TomeIndexInArea = TomeSketchX.Index Then
									If MapSketchZ.Index <> MapSketchX.Index And EntryPointX.AreaIndex = MapSketchZ.AreaIndexInArea And EntryPointX.MapIndex = MapSketchZ.MapIndexInArea Then
										EntryPointX.MapIndex = MapSketchZ.Index
										EntryPointX.IsUsed = True
									End If
								End If
							Next MapSketchZ
						End If
					Next EntryPointX
				End If
			Next MapSketchX
		Next c
		' Load up Map files (for Templates) (can reuse these multiple times)
		SearchFolders("*.rsm", gAppPath & "\library", (UberWizX.MapFiles))
		For c = 1 To UberWizX.MapFiles.Count()
			i = FreeFile
			MapX = New Map
			'UPGRADE_WARNING: Couldn't resolve default property of object UberWizX.MapFiles(c).Text. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
			FileOpen(i, UberWizX.MapFiles.Item(c).Text, OpenMode.Binary)
			MapX.ReadFromFile(i)
			FileClose(i)
			' Add to the MapSketch multiple times
			For n = 0 To Int(Rnd() * 6)
				' Add a MapSketch for it
				MapSketchX = UberWizX.AddMapSketch
				MapSketchX.MapName = MapX.Name
				MapSketchX.MapStyle = MapX.MapStyle
				MapSketchX.MapComments = MapX.Comments
				'UPGRADE_WARNING: Couldn't resolve default property of object UberWizX.MapFiles().Text. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
				MapSketchX.FullPath = UberWizX.MapFiles.Item(c).Text
				MapSketchX.MapIndex = 0
				For	Each EntryPointX In MapX.EntryPoints
					MapSketchX.AddEntryPoint.Copy(EntryPointX)
				Next EntryPointX
				MapSketchX.GenerateUponEntry = MapX.GenerateUponEntry
				' Determine if Map needs Themes and/or needs to Generate Themes
				MapSketchX.NeedsThemes = False
				MapSketchX.GenerateThemes = False
				If MapX.GenerateUponEntry = True Then
					MapSketchX.NeedsThemes = True
					MapSketchX.GenerateThemes = True
				Else
					For	Each EncounterX In MapX.Encounters
						If EncounterX.GenerateUponEntry = True Then
							MapSketchX.GenerateThemes = True
							If EncounterX.ParentTheme = 0 Then
								MapSketchX.NeedsThemes = True
							End If
						End If
					Next EncounterX
				End If
			Next n
		Next c
		' Determine which MapSketchs can be MainMaps (have EntryPoint to Main Menu)
		For	Each MapSketchX In UberWizX.MapSketchs
			For	Each EntryPointX In MapSketchX.EntryPoints
				If (EntryPointX.MapStyle = bdEntryTown Or EntryPointX.MapStyle = bdEntryWilderness) And (EntryPointX.Style = bdEntryExitUp Or EntryPointX.Style = bdEntryExit) And EntryPointX.AreaIndex = 0 Then
					MapSketchX.IsMainMap = True
					Exit For
				End If
			Next EntryPointX
		Next MapSketchX
	End Sub
	
	Public Sub InitUberWizThemes(ByRef UberWizX As UberWizard)
		Dim c, i As Short
		Dim ThemeX As Theme
		Dim ThemeSketchX As ThemeSketch
		' Load up Themes
		SearchFolders("*.rsr", gAppPath & "\library", (UberWizX.ThemeFiles))
		For c = 1 To UberWizX.ThemeFiles.Count()
			i = FreeFile
			ThemeX = New Theme
			'UPGRADE_WARNING: Couldn't resolve default property of object UberWizX.ThemeFiles(c).Text. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
			FileOpen(i, UberWizX.ThemeFiles.Item(c).Text, OpenMode.Binary)
			ThemeX.ReadFromFileHeader(i)
			FileClose(i)
			' Add ThemeSketch
			ThemeSketchX = UberWizX.AddThemeSketch
			ThemeSketchX.Name = ThemeX.Name
			ThemeSketchX.Comments = ThemeX.Comments
			'UPGRADE_WARNING: Couldn't resolve default property of object UberWizX.ThemeFiles().Text. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
			ThemeSketchX.FullPath = UberWizX.ThemeFiles.Item(c).Text
			ThemeSketchX.IsMajorTheme = ThemeX.IsMajorTheme
			ThemeSketchX.IsQuest = ThemeX.IsQuest
			ThemeSketchX.Style = ThemeX.Style
			ThemeSketchX.Coverage = ThemeX.Coverage
		Next c
	End Sub
	
	Public Sub InitUberWizCreatures(ByRef UberWizX As UberWizard)
		Dim c, i As Short
		Dim CreatureX, CreatureZ As Creature
		SearchFolders("*.rsc", gAppPath & "\library", (UberWizX.CreatureFiles))
		For c = 1 To UberWizX.CreatureFiles.Count()
			i = FreeFile
			CreatureX = New Creature
			'UPGRADE_WARNING: Couldn't resolve default property of object UberWizX.CreatureFiles(c).Text. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
			FileOpen(i, UberWizX.CreatureFiles.Item(c).Text, OpenMode.Binary)
			CreatureX.ReadFromFileHeader(i)
			FileClose(i)
			CreatureZ = UberWizX.AddCreature
			CreatureZ.Name = CreatureX.Name
			CreatureZ.Comments = CreatureX.Comments
			'UPGRADE_WARNING: Couldn't resolve default property of object UberWizX.CreatureFiles().Text. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
			CreatureZ.PictureFile = UberWizX.CreatureFiles.Item(c).Text
		Next c
	End Sub
	
	Public Sub InitUberWizItems(ByRef UberWizX As UberWizard)
		Dim c, i As Short
		Dim ItemX, ItemZ As Item
		SearchFolders("*.rsi", gAppPath & "\library", (UberWizX.ItemFiles))
		For c = 1 To UberWizX.ItemFiles.Count()
			i = FreeFile
			ItemX = New Item
			'UPGRADE_WARNING: Couldn't resolve default property of object UberWizX.ItemFiles(c).Text. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
			FileOpen(i, UberWizX.ItemFiles.Item(c).Text, OpenMode.Binary)
			ItemX.ReadFromFileHeader(i)
			FileClose(i)
			ItemZ = UberWizX.AddItem
			ItemZ.Name = ItemX.Name
			ItemZ.Comments = ItemX.Comments
			'UPGRADE_WARNING: Couldn't resolve default property of object UberWizX.ItemFiles().Text. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
			ItemZ.PictureFile = UberWizX.ItemFiles.Item(c).Text
		Next c
	End Sub
	
	Public Function RandomGem() As String
		Dim c As Short
		Randomize()
		c = Int(Rnd() * 37)
		Select Case c
			Case 0 : RandomGem = "Bloodstone"
			Case 1 : RandomGem = "Tigereye"
			Case 2 : RandomGem = "Quartz"
			Case 3 : RandomGem = "Jasper"
			Case 4 : RandomGem = "Moonstone"
			Case 5 : RandomGem = "Star Rose"
			Case 6 : RandomGem = "Crystal"
			Case 7 : RandomGem = "Onyx"
			Case 8 : RandomGem = "Sard"
			Case 9 : RandomGem = "Sardony"
			Case 10 : RandomGem = "Chalcedony"
			Case 11 : RandomGem = "Yellow"
			Case 12 : RandomGem = "Sunstone"
			Case 13 : RandomGem = "Jade"
			Case 14 : RandomGem = "Topaz"
			Case 15 : RandomGem = "White"
			Case 16 : RandomGem = "Garnet"
			Case 17 : RandomGem = "Amber"
			Case 18 : RandomGem = "Alexandrite"
			Case 19 : RandomGem = "Aquamarine"
			Case 20 : RandomGem = "Coral"
			Case 21 : RandomGem = "Peridot"
			Case 22 : RandomGem = "Spinel"
			Case 23 : RandomGem = "Crystal"
			Case 24 : RandomGem = "Gold"
			Case 25 : RandomGem = "Silver"
			Case 26 : RandomGem = "Platinum"
			Case 27 : RandomGem = "Copper"
			Case 28 : RandomGem = "Blue"
			Case 29 : RandomGem = "Red"
			Case 30 : RandomGem = "Yellow"
			Case 31 : RandomGem = "White"
			Case 32 : RandomGem = "Black"
			Case 33 : RandomGem = "Purple"
			Case 34 : RandomGem = "Green"
			Case 35 : RandomGem = "Crystal"
			Case 36 : RandomGem = "Glass"
		End Select
	End Function
	
	Public Function MakeCreatureName(ByRef Method As Short, ByRef WorldName As String) As String
		' Function to make random name for character
		Dim Letters, Vowel As Object
		Dim Filename As String
		Dim strNameSet, Name, Text As String
		Dim strNamesList(3) As String
		Dim i, c, hndFile As Short
		Dim lResult As Integer
		Dim blnFound As Boolean
		Randomize()
		' [Titi 2.4.7] Added possibility to choose names from a text file
		blnFound = False
		strNameSet = "[NameSet" & Right(Str(Method), Len(Str(Method)) - 1) & "]"
		If Method = 0 Then
			MakeCreatureName = ""
		Else
			Filename = gAppPath & "\Roster\" & WorldName & "\" & WorldName & ".txt"
			If oFileSys.CheckExists(Filename, clsInOut.IOActionType.File) Then
				hndFile = FreeFile
				FileOpen(hndFile, Filename, OpenMode.Input)
				Do Until EOF(hndFile)
					For i = 0 To 3
						strNamesList(i) = LineInput(hndFile)
					Next 
					If UCase(strNamesList(0)) = UCase(strNameSet) Then
						blnFound = True
						Exit Do ' found the correct nameset
					End If
				Loop 
				If blnFound Then
					For i = 1 To 3
						strNamesList(i) = Right(strNamesList(i), Len(strNamesList(i)) - InStr(strNamesList(i), "="))
					Next 
					strNamesList(0) = strNamesList(2)
					c = Int(Rnd() * 50) + 1
					' choose the c th name in the list
					While c > 0
						If InStr(strNamesList(0), ",") = 0 Then
							' last name processed: start again!
							strNamesList(0) = strNamesList(2)
						End If
						Text = Left(strNamesList(0), InStr(strNamesList(0), ",") - 1)
						strNamesList(0) = Right(strNamesList(0), Len(strNamesList(0)) - Len(Text) - 1)
						c = c - 1
					End While
					Name = Text
					' only 20% chance to have a title
					If Rnd() * 100 < 80 Then strNamesList(1) = "<none>"
					' only 20% chance to be famous
					If Rnd() * 100 < 80 Then strNamesList(3) = "<none>"
					' now, choose a title and a nickname (if any)
					c = Int(Rnd() * 50) + 1
					strNamesList(0) = strNamesList(1)
					If strNamesList(1) = "<none>" And strNamesList(3) = "<none>" Then
						' reuse the list of names to find a first name
						strNamesList(0) = strNamesList(2)
						c = Int(Rnd() * 50) + 1
						While c > 0
							If InStr(strNamesList(0), ",") = 0 Then
								' last name processed: start again!
								strNamesList(0) = strNamesList(2)
							End If
							Text = Left(strNamesList(0), InStr(strNamesList(0), ",") - 1)
							strNamesList(0) = Right(strNamesList(0), Len(strNamesList(0)) - Len(Text) - 1)
							c = c - 1
						End While
						Name = Text & " " & Name
					Else
						If strNamesList(1) <> "<none>" Then
							While c > 0
								If InStr(strNamesList(0), ",") = 0 Then
									' last name processed: start again!
									strNamesList(0) = strNamesList(1)
								End If
								Text = Left(strNamesList(0), InStr(strNamesList(0), ",") - 1)
								strNamesList(0) = Right(strNamesList(0), Len(strNamesList(0)) - Len(Text) - 1)
								c = c - 1
							End While
							Name = Text & " " & Name
						End If
						c = Int(Rnd() * 50) + 1
						strNamesList(0) = strNamesList(3)
						If strNamesList(3) <> "<none>" Then
							While c > 0
								If InStr(strNamesList(0), ",") = 0 Then
									' last name processed: start again!
									strNamesList(0) = strNamesList(3)
								End If
								Text = Left(strNamesList(0), InStr(strNamesList(0), ",") - 1)
								strNamesList(0) = Right(strNamesList(0), Len(strNamesList(0)) - Len(Text) - 1)
								c = c - 1
							End While
							Name = Name & " " & Text
						End If
					End If
				Else ' nameset not found --> random names will be created
					'UPGRADE_WARNING: Array has a new behavior. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="9B7D5ADD-D8FE-4819-A36C-6DEDAF088CC7"'
					'UPGRADE_WARNING: Couldn't resolve default property of object Letters. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
					Letters = New Object(){"ch", "th", "cr", "l", "br", "d", "f", "g", "j", "k", "m", "n", "p", "r", "s", "t", "v", "w", "z", "tr", "st", "sh"}
					'UPGRADE_WARNING: Array has a new behavior. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="9B7D5ADD-D8FE-4819-A36C-6DEDAF088CC7"'
					'UPGRADE_WARNING: Couldn't resolve default property of object Vowel. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
					Vowel = New Object(){"a", "e", "i", "o", "u", "y"}
					'UPGRADE_WARNING: Couldn't resolve default property of object Letters(Int(Rnd * 21)). Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
					'UPGRADE_WARNING: Couldn't resolve default property of object Vowel(Int(Rnd * 6)). Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
					'UPGRADE_WARNING: Couldn't resolve default property of object Letters(). Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
					Name = UCase(Mid(Letters(Int(Rnd() * 21)), 1, 1)) & Vowel(Int(Rnd() * 6)) & Letters(Int(Rnd() * 21))
					If Int(Rnd() * 100) > 80 Then
						'UPGRADE_WARNING: Couldn't resolve default property of object Letters(Int(Rnd * 21)). Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
						'UPGRADE_WARNING: Couldn't resolve default property of object Vowel(). Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
						Name = Name & Vowel(Int(Rnd() * 6)) & Letters(Int(Rnd() * 21))
					End If
					'UPGRADE_WARNING: Couldn't resolve default property of object Letters(Int(Rnd * 21)). Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
					'UPGRADE_WARNING: Couldn't resolve default property of object Vowel(Int(Rnd * 6)). Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
					'UPGRADE_WARNING: Couldn't resolve default property of object Letters(). Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
					Name = Name & " " & UCase(Mid(Letters(Int(Rnd() * 21)), 1, 1)) & Vowel(Int(Rnd() * 6)) & Letters(Int(Rnd() * 21))
					If Int(Rnd() * 100) > 80 Then
						'UPGRADE_WARNING: Couldn't resolve default property of object Letters(Int(Rnd * 21)). Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
						'UPGRADE_WARNING: Couldn't resolve default property of object Vowel(). Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
						Name = Name & Vowel(Int(Rnd() * 6)) & Letters(Int(Rnd() * 21))
					End If
					MakeCreatureName = Name
				End If
			Else ' name list not found --> random names will be created
				'    Select Case Method
				'        Case 0  ' None
				'            MakeCreatureName = ""
				'        Case 1  ' Numenorean
				'            Letters = Array("Petro", "Marina", "Parostannos", "Vasso", "Lotrina", "Theone", "Daphne", "Spindes", "Estrios", "Tannotos", "George", "Thalia", "Chanstontena", "Selednotos", "Lannos", "Chrastos", "Chrystena", "Phatros", "Mirios", "Spiro", "Thease", "Nichasios", "Pallis", "Meles", "Likonnotos", "Leorasios", "Nidatreas", "Genos", "Aphrysini", "Elistinos", "Petrina", "Alympia", "Geodisios", "Nalos", "Datre", "Arge", "Mikisios", "Comirios", "Asta", "Chrystine", "Christina", "Solon", "Changeoros", "Niroltannos", "Linda", "Endria", "Mikis", "Elaine", "Anstios", "Panstastina")
				'            MakeCreatureName = Letters(Int(Rnd * 50)) & " " & Letters(Int(Rnd * 50))
				'        Case 2  ' Harsh
				'            Letters = Array("Carorn", "Thorgrak", "Thalarrad", "Lorthak", "Loraldrarg", "Gormorn", "Shangrorn", "Orerall", "Bengrun", "Inodrand", "Ironarg", "Urkin", "Tortharg", "Tharen", "Urorthak", "Bongath", "Farthad", "Thoraruk", "Atheneon", "Madrarg", "Rerurmad", "Lengrarg", "Aranil", "Boretheon", "Fogirthane", "Farath", "Irurthorn", "Throom", "Idrararg", "Umanarg", "Kigrarg", "Sharangrim", "Toomuk", "Arirthim", "Fongrarg", "Imgalay", "Thuralarg", "Korarrand", "Lak", "Fathrorn", "Thothrim", "Glorthim", "Kerkin", "Asengrith", "Glereon", "Therthad", "Glorg", "Mororarg", "Igrararg", "Virarg")
				'            MakeCreatureName = Letters(Int(Rnd * 50))
				'        Case 3  ' Sweedish
				'            Letters = Array("Helegorm", "Onglir", "Helegur", "Daegolodh", "Redhel", "Derval", "Menerost", "Naeriel", "Nuril", "Lerchilthor", "Nedhelrin", "Aragnas", "Firuin", "Belegil", "Celebrimphor", "Baurad", "Maelian", "Denerdil", "Bretharmir", "Nildinael", "Urthar", "Hadring", "Mondel", "Aradan", "Elphril", "Falathorn", "Arunion", "Aradrion", "Endil", "Lanbor", "Delril", "Edhelorn", "Dolgil", "Fanir", "Eramoth", "Dinor", "Aglaras", "Berethel", "Lelfinir", "Heldan", "Rochonarth", "Maeglos", "Ningarthil", "Mangorel", "Ilthros", "Bendal", "Gelephor", "Bauros", "Galmindir", "Meneron")
				'            MakeCreatureName = Letters(Int(Rnd * 50))
				'        Case 4  ' Viking
				'            Letters = Array("Tarrid", "Knimvir", "Svognar", "Sigrir", "Eysti", "Herkja", "Jorokar", "Sladrir", "Sorli", "Gand", "Siredir", "Hrotti", "Stolfrun", "Hoki", "Smor", "Bur", "Urvid", "Horod", "Siling", "Hogni", "Tonning", "Hirgnir", "Hjotra", "Hrend", "Veld", "Aelleid", "Gjaldis", "Gollrek", "Gjalvir", "Dallvar", "Hoddvild", "Staki", "Hjottein", "Geirir", "Thurt", "Keddrosti", "Draerek", "Kjad", "Svidi", "Freigar", "Froki", "Hakar", "Armunrek", "Anhvid", "Nanmar", "Varming", "Rotti", "Frabbi", "Hofnheid", "Hrabbi")
				'            MakeCreatureName = Letters(Int(Rnd * 50)) & " " & Letters(Int(Rnd * 50))
				'        Case 5  ' Otherworld
				'            Letters = Array("Laindus", "Geriones", "Thenierant", "Ondas", "Asenna", "Hanon", "Astranan", "Mai", "Blon", "Dret", "Guerilence", "Aris", "Axilombe", "Meilyonan", "Alin", "Escudan", "Elamonis", "Marises", "Gredrant", "Tiadan", "Belius", "Hue", "Mangleins", "Jorsaine", "Rhonsience", "Vilake", "Bielilant", "Gai", "Guiomalant", "Bramebor", "Tonnade", "Tontrian", "Ale", "Vasant", "Praitt", "Vere", "Irlare", "Besade", "Lioladene", "Gorance", "Mourbin", "Gisone", "Lurognes", "Guerlirion", "Meinonains", "Gallain", "Ilant", "Milgrin", "Mancus", "Vyrius")
				'            MakeCreatureName = Letters(Int(Rnd * 50)) & " " & Letters(Int(Rnd * 50))
				'        Case 6  ' Mythology
				'            Letters = Array("Demas", "Gela", "Miel", "Biglian", "Allan", "Mancolan", "Dilanie", "Ranor", "Methlen", "Manda", "Lindunet", "Jan", "Joslette", "Golesa", "Clarlisa", "Jaynet", "Demian", "Gel", "Hendse", "Laura", "Cargela", "Chaline", "Rackim", "Bithron", "Surstia", "Daslorah", "Rayne", "Raig", "Ian", "Sean", "Chramy", "Kiriss", "Mielcan", "Julian", "Melinda", "Jamerah", "Hemin", "Morther", "Annah", "Kerry", "Gonia", "Mine", "Iam", "Jalan", "Rune", "Lesley", "Jem", "Darina", "Landa", "Nerder")
				'            MakeCreatureName = Letters(Int(Rnd * 50)) & " " & Letters(Int(Rnd * 50))
				'        Case 7  ' Greekish
				'            Letters = Array("Heriis", "Jonesmi", "Kaino", "Seira", "Ruuno", "Eliti", "Voin", "Linda", "Imma", "Lelias", "Maini", "Tanda", "Sellopi", "Tarkko", "Sarja", "Iiva", "Sirsti", "Aani", "Aunal", "Varia", "Killo", "Raili", "Unimi", "Alle", "Ina", "Terkko", "Suva", "Esmaa", "Lorkki", "Erpi", "Paulal", "Tiia", "Auna", "Tervo", "Sura", "Aija", "Sutsi", "Taaro", "Unarjo", "Iina", "Talku", "Lilepa", "Laina", "Auro", "Sirtti", "Ikko", "Jalli", "Anno", "Rinna", "Rarulmo")
				'            MakeCreatureName = Letters(Int(Rnd * 50)) & " " & Letters(Int(Rnd * 50))
				'        Case 8  ' Mythology 2
				'            Letters = Array("Line", "Bentaclus", "Prades", "Llan", "Khanton", "Belos", "Rataidnabithria", "Aulos", "Amonia", "Marajnis", "Gormios", "Heladnes", "Honis", "Danna", "Frivas", "Onca", "Nilanggit", "Daleitho", "Atethra", "Somors", "Shakti", "Litos", "Manja", "Apidwatha", "Pruantia", "Mutlomele", "Amis", "Cepseidfrabia", "Skinyes", "Cuan", "Clos", "Pretan", "Eittona", "Mincethrin", "Saule", "Anipnayen", "Deun", "Plana", "Lescan", "Eankos", "Molosvasbiclos", "Babos", "Luninnos", "Omcrusa", "Pellia", "Xinghurcan", "Invas", "Aidenis", "Holda", "Inia")
				'            Vowel = Array("Lord", "Baron", "Sir", "Duke")
				'            MakeCreatureName = Vowel(Int(Rnd * 3)) & " " & Letters(Int(Rnd * 50))
				'        Case 9  ' Irish
				'            Letters = Array("Searshary", "Bromer", "Brone", "Cocdoldser", "Aliper", "Daler", "Deard", "Hirt", "Morker", "Dorlerwin", "Janell", "Yeager", "Hond", "Alader", "Sminsmally", "Anagh", "Glonn", "Shett", "Hoome", "Boyncego", "Hell", "Coogan", "Affmott", "Bunnarser", "Bonck", "Stock", "Ranoseller", "Molin", "Scairker", "Hooder", "Phirtwis", "Worrier", "Rorshall", "Jorn", "Kirloat", "Woom", "Durson", "Clevend", "Hider", "Boy", "Riler", "Minell", "Teuarews", "Wirkendrup", "Alider", "Molds", "Hanosunnes", "Wencott", "Aldender", "Wallark")
				'            MakeCreatureName = Letters(Int(Rnd * 50)) & " " & Letters(Int(Rnd * 50))
				'        Case 10 ' Muse
				'            Letters = Array("Calliope", "Merpsichore", "Eutomne", "Terpsichore", "Perelymnia", "Calyhymnia", "Merpsichome", "Eutelpe", "Torpelymne", "Thalliope", "Meralymnia", "Urene", "Thallymere", "Tellymene", "Mellyhyme", "Polyhymnia", "Pelio", "Urane", "Thalperene", "Telpenere", "Urate", "Terpsichome", "Thania", "Erenia", "Clia", "Melliope", "Teliopere", "Merenelia", "Catolymnia", "Meliopere", "Menerallio", "Urelia", "Torenellia", "Telia", "Clio", "Eutolpe", "Melpenere", "Merellio", "Mellymene", "Toralia", "Callyhyme", "Pe", "Poratellia", "Calperate", "Telymnere", "Canerellio", "Clymene", "Telliope", "Uralio", "Callymene")
				'            Vowel = Array("the Dark", "the Wise", "the Elder", "the Grey", "the White", "the Black", "the Sage", "the True", "the Faithful", "the Kind")
				'            Select Case Int(Rnd * 100)
				'                Case 0 To 50
				'                    MakeCreatureName = Letters(Int(Rnd * 50)) & " " & Letters(Int(Rnd * 50))
				'                Case 51 To 85
				'                    MakeCreatureName = Letters(Int(Rnd * 50)) & " " & Vowel(Int(Rnd * 10))
				'                Case 86 To 99
				'                    MakeCreatureName = Letters(Int(Rnd * 50))
				'            End Select
				'        Case 11 ' Planets
				'            Letters = Array("Aloola", "Deyerar", "Chularir", "Aroot", "Onothal", "Aminan", "Anon", "Raurria", "La", "Khordurr", "Sason", "Boonor", "Tior", "Mindas", "Proll", "Julvar", "Bes-Vis", "Geloboo", "Zetedda", "Drorkr", "Carill", "Paulin", "Antabed", "Dabolvu", "Koneriu", "Gle", "Sorthe", "Vudruc-Krox", "Glaety", "Frauleme", "Eurona", "Chartrurre", "Dribarurdogin", "Phaesia", "Kemalla", "Onoon", "Muna", "Failo", "Ammis", "Miroll", "Harmeb", "Agredi", "Ponti", "Kontelvax-Ki", "Bigosa", "Smur", "Aneenegua", "Vandador", "Cegrak", "Yenynnetan")
				'            MakeCreatureName = Letters(Int(Rnd * 50)) & " " & Letters(Int(Rnd * 50))
				'        Case 12 ' Saints
				'            Letters = Array("Mariver", "Mapanus", "Ranard", "Eurasia", "Prace", "Ane", "Rasia", "Ase", "Jas", "Ailagia", "Elix", "Floncinus", "Arnard", "Sidric", "Jacis", "Chriremund", "Moria", "Ringil", "Balamas", "Admara", "Dobinus", "Vargil", "Amas", "Bilerascius", "Alose", "Aigela", "Minard", "Therius", "Basacancius", "Eracius", "Jelbeth", "Maura", "Elfric", "Atace", "Galene", "Brimbron", "Rine", "Golian", "Holas", "Daur", "Marencius", "Leonard", "Irace", "Phoeberno", "Anskar", "Aver", "Edmolas", "Theonard", "Fracus", "Giana")
				'            Vowel = Array("Bishop", "Saint", "Cleric", "Elder", "Friar")
				'            If Int(Rnd * 100) < 20 Then
				'                MakeCreatureName = Vowel(Int(Rnd * 5)) & " " & Letters(Int(Rnd * 50))
				'            Else
				'                MakeCreatureName = Letters(Int(Rnd * 50)) & " " & Letters(Int(Rnd * 50))
				'            End If
				'        Case 13 ' Welsh
				'            Letters = Array("Cleispad", "Cunweir", "Anwyll", "Einbyr", "Ceir", "Cyfyr", "Egrydwas", "Gonrhyr", "Goethyen", "Nwydedeg", "Daranwyr", "Giwg", "Ceirnyr", "Geslydd", "Har", "Beladar", "Enwydd", "Habwy", "Teithi", "Enwyll", "Gwarae", "Cem", "Cnoreid", "Efyr", "Cado", "Alasgyr", "Llwad", "Madd", "Gefan", "Dirmedoldeith", "Urnyn", "Ceiny", "Llydwen", "Cidydd", "Conwyr", "Gadan", "Coedarn", "Mindrwed", "Rhyduach", "Llafyr", "Cyndrwyn", "Gwid", "Bwerdwyen", "Alydd", "Lymys", "Achdras", "Nwadu", "Gasben", "Gwalig", "Tylwch")
				'            MakeCreatureName = Letters(Int(Rnd * 50))
				'        Case 14 ' Muslim
				'            Letters = Array("Idsaan", "Anwat", "Taal", "Kaaseelah", "Qiraar", "Naameesah", "Jafar", "Qaarah", "Areeb", "Rameer", "Khaabah", "Ratf", "Zailasah", "Adrah", "Deelamil", "Muneedah", "Saaneem", "Aafah", "Shihraaj", "Rahrabaa", "Baanee", "Halaal", "Samaahah", "Adsah", "Hassan", "Zaihaarah", "Khariyah", "Sabbeer", "Ims", "Shahaarah", "Taqee", "Haqq", "Aaaqi", "Ihaa", "Aki", "Aftah", "Imaamah", "Iltaan", "Aheem", "Habeenah", "Rulaabah", "Heesah", "Yaamatyab", "Shafee", "Yoosef", "Sughraa", "Zaadah", "Aanah", "Jah", "Khibaar")
				'            MakeCreatureName = Letters(Int(Rnd * 50)) & " " & Letters(Int(Rnd * 50))
				'        Case Else  ' Broken Strange
				'UPGRADE_WARNING: Array has a new behavior. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="9B7D5ADD-D8FE-4819-A36C-6DEDAF088CC7"'
				'UPGRADE_WARNING: Couldn't resolve default property of object Letters. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
				Letters = New Object(){"ch", "th", "cr", "l", "br", "d", "f", "g", "j", "k", "m", "n", "p", "r", "s", "t", "v", "w", "z", "tr", "st", "sh"}
				'UPGRADE_WARNING: Array has a new behavior. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="9B7D5ADD-D8FE-4819-A36C-6DEDAF088CC7"'
				'UPGRADE_WARNING: Couldn't resolve default property of object Vowel. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
				Vowel = New Object(){"a", "e", "i", "o", "u", "y"}
				'UPGRADE_WARNING: Couldn't resolve default property of object Letters(Int(Rnd * 21)). Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
				'UPGRADE_WARNING: Couldn't resolve default property of object Vowel(Int(Rnd * 6)). Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
				'UPGRADE_WARNING: Couldn't resolve default property of object Letters(). Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
				Name = UCase(Mid(Letters(Int(Rnd() * 21)), 1, 1)) & Vowel(Int(Rnd() * 6)) & Letters(Int(Rnd() * 21))
				If Int(Rnd() * 100) > 80 Then
					'UPGRADE_WARNING: Couldn't resolve default property of object Letters(Int(Rnd * 21)). Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
					'UPGRADE_WARNING: Couldn't resolve default property of object Vowel(). Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
					Name = Name & Vowel(Int(Rnd() * 6)) & Letters(Int(Rnd() * 21))
				End If
				'UPGRADE_WARNING: Couldn't resolve default property of object Letters(Int(Rnd * 21)). Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
				'UPGRADE_WARNING: Couldn't resolve default property of object Vowel(Int(Rnd * 6)). Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
				'UPGRADE_WARNING: Couldn't resolve default property of object Letters(). Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
				Name = Name & " " & UCase(Mid(Letters(Int(Rnd() * 21)), 1, 1)) & Vowel(Int(Rnd() * 6)) & Letters(Int(Rnd() * 21))
				If Int(Rnd() * 100) > 80 Then
					'UPGRADE_WARNING: Couldn't resolve default property of object Letters(Int(Rnd * 21)). Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
					'UPGRADE_WARNING: Couldn't resolve default property of object Vowel(). Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
					Name = Name & Vowel(Int(Rnd() * 6)) & Letters(Int(Rnd() * 21))
				End If
				'            MakeCreatureName = Name
			End If
			MakeCreatureName = Name
		End If
		'    End Select
	End Function
	
	Public Sub MakeMap(ByRef MapX As Map, ByRef Fill As Short)
		Dim y, x, c As Short
		Dim EncounterX As Encounter
		Randomize()
		' If appending, go mark the map as Halls
		If MapX.GenerateAppend = True Then
			' Set Hall Encounters
			For x = 0 To MapX.Width : For y = 0 To MapX.Height
					If MapX.BottomTile(x, y) > 0 And MapX.EncPointer(x, y) = 0 Then
						MapX.EncPointer(x, y) = -1
					End If
				Next y : Next x
		ElseIf MapX.DefaultStyle <> 5 Then  ' None (expected behavior is would not wipe map)
			' Wipe Map
			For x = 0 To MapX.Width : For y = 0 To MapX.Height
					MapX.BottomTile(x, y) = 0
					MapX.BottomFlip(x, y) = False
					MapX.MiddleTile(x, y) = 0
					MapX.MiddleFlip(x, y) = False
					MapX.TopTile(x, y) = 0
					MapX.TopFlip(x, y) = False
					MapX.Hidden(x, y) = False
					MapX.EncPointer(x, y) = 0
				Next y : Next x
			' Delete all encounters
			For	Each EncounterX In MapX.Encounters
				MapX.RemoveEncounter("E" & EncounterX.Index)
			Next EncounterX
		End If
		' Call routine based on DefaultStyle
		Select Case MapX.DefaultStyle
			Case 0 ' Tetris Spin
				TetrisSpin(MapX)
				WallMap(MapX)
			Case 1 ' Cavern Spin
				CavernSprawl(MapX)
				WallMap(MapX)
			Case 2 ' Bubble Tubes
				BubbleTubes(MapX)
				WallMap(MapX)
			Case 3 ' Rectangles
				Rectangles(MapX)
				WallMap(MapX)
			Case 4 ' Huge Cavern
				HugeCavern(MapX)
				WallMap(MapX)
			Case 5 ' None
				' Does not regen the Map
				If MapX.GenerateAppend = True Then
					WallMap(MapX)
				End If
		End Select
		MapX.GenerateUponEntry = False
		' Spin Themes for the Map
		If Fill = True Then
			FillMap(MapX)
		End If
	End Sub
	
	Private Sub WallMap(ByRef MapX As Map)
		Dim Found, y, x, c, i As Short
		Dim TileX As Tile
		Dim TileSetX As TileSet
		Dim TileSetPtr As Short
		Dim EntryX As EntryPoint
		Dim CornerWall, LeftWall, LeftArch As Short
		Dim RoomFloor, HallFloor, CornerArch As Short
		Dim ExitUp, ExitDown As Short
		Dim DoorFrom(500) As Short
		Dim DoorTo(500) As Short
		Dim DoorCnt, DoorListCnt As Short
		Dim DoorList(16) As Short
		Dim WallDecoCnt, FloorDecoCnt, Tick As Short
		Dim FloorDecoList(16) As Short
		Dim WallDecoList(16) As Short
		Dim ToX, ToY As Short
		' Find wall tiles
		LeftWall = 0 : CornerWall = 0 : LeftArch = 0 : RoomFloor = 0 : HallFloor = 0 : CornerArch = 0
		DoorListCnt = 0 : FloorDecoCnt = 0 : WallDecoCnt = 0
		For	Each TileX In MapX.Tiles
			Select Case TileX.Style
				Case bdTileLeftWall
					LeftWall = TileX.Index
				Case bdTileCornerWall
					CornerWall = TileX.Index
				Case bdTileLeftArch
					LeftArch = TileX.Index
				Case bdTileFloor
					If (RoomFloor > 0 And Int(Rnd() * 100) < 50) Or RoomFloor = 0 Then
						RoomFloor = TileX.Index
					End If
					If (HallFloor > 0 And Int(Rnd() * 100) < 50) Or HallFloor = 0 Then
						HallFloor = TileX.Index
					End If
				Case bdTileExitUp
					If (ExitUp > 0 And Int(Rnd() * 100) < 50) Or ExitUp = 0 Then
						ExitUp = TileX.Index
					End If
				Case bdTileExitDown
					If (ExitDown > 0 And Int(Rnd() * 100) < 50) Or ExitDown = 0 Then
						ExitDown = TileX.Index
					End If
				Case bdTileLeftDoor
					If DoorListCnt < 16 Then
						DoorList(DoorListCnt) = TileX.Index
						DoorListCnt = DoorListCnt + 1
					End If
				Case bdTileFloorDeco
					If FloorDecoCnt < 16 Then
						FloorDecoList(FloorDecoCnt) = TileX.Index
						FloorDecoCnt = FloorDecoCnt + 1
					End If
				Case bdTileWallDeco
					If WallDecoCnt < 16 Then
						WallDecoList(WallDecoCnt) = TileX.Index
						WallDecoCnt = WallDecoCnt + 1
					End If
				Case bdTileCornerArch
					CornerArch = TileX.Index
			End Select
		Next TileX
		' Paint map (build walls)
		For x = 0 To MapX.Width : For y = 0 To MapX.Height
				' Hide all spots
				MapX.Hidden(x, y) = True
				' Plot Floor Tile
				If MapX.BottomTile(x, y) = 0 Then
					If MapX.EncPointer(x, y) < 0 Then
						MapX.BottomTile(x, y) = HallFloor
						MapX.BottomFlip(x, y) = Int(Rnd() * 2) - 1
					ElseIf MapX.EncPointer(x, y) > 0 Then 
						MapX.BottomTile(x, y) = RoomFloor
						MapX.BottomFlip(x, y) = Int(Rnd() * 2) - 1
					End If
				End If
				' Plot Walls
				If MapX.MiddleTile(x, y) = 0 Then
					If y = 0 Then
						If MapX.EncPointer(x, y) <> 0 Then
							If x = MapX.Width Then
								MapX.MiddleTile(x, y) = CornerWall
							ElseIf MapX.EncPointer(x, y) <> MapX.EncPointer(x + 1, y) Then 
								MapX.MiddleTile(x, y) = CornerWall
							Else
								MapX.MiddleTile(x, y) = LeftWall
							End If
						ElseIf x < MapX.Width Then 
							If MapX.EncPointer(x, y) <> MapX.EncPointer(x + 1, y) Then
								MapX.MiddleTile(x, y) = LeftWall
								MapX.MiddleFlip(x, y) = True
							End If
						End If
					ElseIf x = MapX.Width Then 
						If MapX.EncPointer(x, y) <> 0 Then
							MapX.MiddleTile(x, y) = LeftWall
							MapX.MiddleFlip(x, y) = True
						End If
					ElseIf MapX.EncPointer(x, y) <> MapX.EncPointer(x, y - 1) And Not (MapX.EncPointer(x, y) < 0 And MapX.EncPointer(x, y - 1) < 0) Then 
						If MapX.EncPointer(x, y) <> MapX.EncPointer(x + 1, y) And Not (MapX.EncPointer(x, y) < 0 And MapX.EncPointer(x + 1, y) < 0) Then
							MapX.MiddleTile(x, y) = CornerWall
						Else
							MapX.MiddleTile(x, y) = LeftWall
						End If
					ElseIf MapX.EncPointer(x, y) <> MapX.EncPointer(x + 1, y) And Not (MapX.EncPointer(x, y) < 0 And MapX.EncPointer(x + 1, y) < 0) Then 
						MapX.MiddleTile(x, y) = LeftWall
						MapX.MiddleFlip(x, y) = True
					End If
				End If
			Next y : Next x
		' Build doors and arches
		DoorCnt = 0
		For x = 0 To MapX.Width - 1 : For y = 1 To MapX.Height
				' If in an area
				If MapX.EncPointer(x, y) <> 0 Then
					' And there is a wall
					If MapX.MiddleTile(x, y) > 0 Then
						' Determine the direction to look for a door
						If MapX.MiddleFlip(x, y) = False And MapX.MiddleTile(x, y) = LeftWall Then
							ToX = x : ToY = y - 1
						Else
							ToX = x + 1 : ToY = y
						End If
						If MapX.EncPointer(ToX, ToY) = 0 And MapX.MiddleTile(x, y) = CornerWall Then
							ToX = x : ToY = y - 1
						End If
						' If direction for door points to another area and don't have too many doors
						If MapX.EncPointer(ToX, ToY) <> 0 And DoorCnt < 500 Then
							' Find out if there already is a door between these two areas
							Found = False : i = 0
							For c = 0 To DoorCnt
								If DoorFrom(c) = MapX.EncPointer(x, y) And DoorTo(c) = MapX.EncPointer(ToX, ToY) Then
									Found = True
									Exit For
								End If
							Next c
							' If not, then create a door
							If Not Found Then
								If MapX.MiddleTile(x, y) = CornerWall Then
									MapX.TopTile(x, y) = LeftWall
									If ToY = y Then
										MapX.MiddleFlip(x, y) = True
									Else
										MapX.TopFlip(x, y) = True
									End If
								End If
								MapX.MiddleTile(x, y) = LeftArch
								DoorCnt = DoorCnt + 1
								DoorFrom(DoorCnt) = MapX.EncPointer(x, y)
								DoorTo(DoorCnt) = MapX.EncPointer(ToX, ToY)
								If MapX.TopTile(x, y) = 0 And Int(Rnd() * 100) > 5 Then
									MapX.TopTile(x, y) = DoorList(Int(Rnd() * DoorListCnt))
									MapX.TopFlip(x, y) = MapX.MiddleFlip(x, y)
								End If
							End If
						End If
					End If
				End If
			Next y : Next x
		' Paint stairs (build entry points)
		For	Each EntryX In MapX.EntryPoints
			Found = False
			Do Until Found = True
				x = Int(Rnd() * MapX.Width) + 1 : y = Int(Rnd() * MapX.Height) + 1
				If MapX.EncPointer(x, y) <> 0 And MapX.TopTile(x, y) = 0 Then
					Select Case EntryX.Style
						Case 0 ' Same level
							EntryX.MapX = x : EntryX.MapY = y
							Found = True
						Case 1 ' Exit Up
							If MapX.MiddleTile(x, y) = LeftWall Then
								EntryX.MapX = x : EntryX.MapY = y
								MapX.TopTile(x, y) = ExitUp
								MapX.TopFlip(x, y) = MapX.MiddleFlip(x, y)
								Found = True
							End If
						Case 2 ' Exit Down
							If MapX.MiddleTile(x, y) = 0 Then
								EntryX.MapX = x : EntryX.MapY = y
								MapX.TopTile(x, y) = ExitDown
								MapX.TopFlip(x, y) = Int(Rnd() * 2) - 1
								Found = True
							End If
						Case Else
							Found = True
					End Select
				End If
			Loop 
		Next EntryX
		' Scatter X Floor Decorations
		If FloorDecoCnt > 0 Then
			For c = 0 To CShort((CDbl(MapX.Width) * CDbl(MapX.Height)) / 50)
				Found = False : Tick = 0
				Do Until Found = True Or Tick > 100
					x = Int(Rnd() * MapX.Width) + 1 : y = Int(Rnd() * MapX.Height) + 1 : Tick = Tick + 1
					If MapX.EncPointer(x, y) <> 0 Then
						If MapX.TopTile(x, y) = 0 Then
							MapX.TopTile(x, y) = FloorDecoList(Int(Rnd() * FloorDecoCnt))
							MapX.TopFlip(x, y) = Int(Rnd() * 2) - 1
							Found = True
						ElseIf MapX.MiddleTile(x, y) = 0 Then 
							MapX.MiddleTile(x, y) = FloorDecoList(Int(Rnd() * FloorDecoCnt))
							MapX.MiddleFlip(x, y) = Int(Rnd() * 2) - 1
							Found = True
						End If
					End If
				Loop 
			Next c
		End If
		' Scatter X Wall Decorations
		If WallDecoCnt > 0 Then
			For c = 0 To Least(CShort((CDbl(MapX.Width) * CDbl(MapX.Height)) / 50), 50)
				Found = False : Tick = 0
				Do Until Found = True Or Tick > 100
					x = Int(Rnd() * MapX.Width) + 1 : y = Int(Rnd() * MapX.Height) + 1 : Tick = Tick + 1
					If MapX.EncPointer(x, y) > 0 And MapX.TopTile(x, y) = 0 Then
						If MapX.MiddleTile(x, y) = LeftWall Then
							MapX.TopTile(x, y) = WallDecoList(Int(Rnd() * WallDecoCnt))
							MapX.TopFlip(x, y) = MapX.MiddleFlip(x, y)
							Found = True
						ElseIf MapX.MiddleTile(x, y) = CornerWall Then 
							MapX.TopTile(x, y) = WallDecoList(Int(Rnd() * WallDecoCnt))
							MapX.TopFlip(x, y) = Int(Rnd() * 2) - 1
							Found = True
						End If
					End If
				Loop 
			Next c
		End If
	End Sub
	
	Private Sub TetrisSpin(ByRef MapX As Map)
		Dim h, c, w, i As Short
		Dim MaxH, MaxW, EncounterCnt As Short
		Dim Found, x, y, Plop As Short
		Dim nonCollisions, temp, Tick As Short
		Dim XDir, YDir As Short
		Dim EncounterX As Encounter
		Dim Geo(5, 5) As Short
		' Size Map (+/- 10% of Defaults)
		If MapX.GenerateAppend = False Then
			MapX.Width = Greatest(MapX.DefaultWidth * ((Int(Rnd() * 20) + 90) / 100), 24)
			MapX.Height = Greatest(MapX.DefaultHeight * ((Int(Rnd() * 20) + 90) / 100), 24)
			EncounterCnt = 1
		Else
			EncounterCnt = MapX.Encounters.Count()
		End If
		' Loop until all rooms are populated
		Plop = False : c = EncounterCnt - 1
		Do Until EncounterCnt > MapX.DefaultEncounters
			' Clear GeoMorph
			For w = 0 To 5 : For h = 0 To 5
					Geo(w, h) = 0
				Next h : Next w
			' Create new GeoMorph
			c = c + 1
			Select Case Int(Rnd() * 100)
				Case 0 To 30 ' Room
					EncounterX = MapX.AddEncounter
					EncounterX.GenerateUponEntry = True
					MaxW = Int(Rnd() * 4) + 2 : MaxH = Int(Rnd() * 4) + 2
					For w = 0 To MaxW : For h = 0 To MaxH
							Geo(w, h) = EncounterX.Index
						Next h : Next w
					EncounterCnt = EncounterCnt + 1
				Case 31 To 70 ' Hall
					Select Case Int(Rnd() * 2)
						Case 0 ' Left to Right
							MaxW = Int(Rnd() * 3) + 2
							For w = 0 To MaxW : h = 0
								Geo(w, h) = -c
							Next w
						Case 1 ' Top to Bottom
							MaxH = Int(Rnd() * 3) + 2
							w = 0 : For h = 0 To MaxH
								Geo(w, h) = -c
							Next h
					End Select
				Case 71 To 85 ' T-Intersection
					MaxW = Int(Rnd() * 3) + 2
					For w = 0 To MaxW : h = 0
						Geo(w, h) = -c
					Next w
					MaxH = Int(Rnd() * 3) + 2
					w = Int(Rnd() * MaxW) + 1 : For h = 0 To MaxH
						Geo(w, h) = -c
					Next h
				Case 86 To 99 ' L-Turn
					MaxW = Int(Rnd() * 3) + 2
					For w = 0 To MaxW : h = 0
						Geo(w, h) = -c
					Next w
					MaxH = Int(Rnd() * 3) + 2
					For h = 0 To MaxH : w = 0
						Geo(w, h) = -c
					Next h
			End Select
			' Apply 0 to 3 90 degree rotations to the geo
			For i = 0 To Int(Rnd() * 4) - 1
				For w = 0 To 2
					For h = 0 To 2
						temp = Geo(w, h)
						Geo(w, h) = Geo(h, 5 - w)
						Geo(h, 5 - w) = Geo(5 - w, 5 - h)
						Geo(5 - w, 5 - h) = Geo(5 - h, w)
						Geo(5 - h, w) = temp
					Next h
				Next w
			Next i
			
			' If this is the first GeoMorph, then just plop it down.
			' Else, spin it into the middle of the map.
			If Plop = False Then
				x = Int(Rnd() * 3) + Int(MapX.Width / 2)
				y = Int(Rnd() * 3) + Int(MapX.Height / 2)
				Plop = True
			Else
				Tick = 0
				Do 
					Tick = Tick + 1
					XDir = Int(Rnd() * 3) - 1 : YDir = 0
					If XDir = 0 Then
						YDir = 1 - 2 * Int(Rnd() * 2)
						x = Int(Rnd() * (MapX.Width - 6)) + 1
						y = 0
						If YDir < 1 Then
							y = MapX.Height - 6
						End If
					Else
						x = 1
						y = Int(Rnd() * (MapX.Height - 6))
						If XDir < 1 Then
							x = MapX.Width - 6
						End If
					End If
					Found = False
					nonCollisions = 0
					Do Until Found = True Or x + XDir > MapX.Width - 6 Or x + XDir < 0 Or y + YDir > MapX.Height - 6 Or y + YDir < 0
						For w = 0 To 5 : For h = 0 To 5
								If Geo(w, h) <> 0 Then
									If MapX.EncPointer(x + XDir + w, y + YDir + h) <> 0 Then
										Found = True
									End If
								End If
							Next h : Next w
						If Not Found Then
							nonCollisions = nonCollisions + 1
							x = x + XDir
							y = y + YDir
						End If
					Loop 
				Loop Until (Found = True And nonCollisions > 0) Or Tick > 1000
			End If
			For w = 0 To 5 : For h = 0 To 5
					If Geo(w, h) <> 0 Then
						MapX.EncPointer(x + w, y + h) = Geo(w, h)
					End If
				Next h : Next w
		Loop 
	End Sub
	
	Private Sub CavernSprawl(ByRef MapX As Map)
		Dim h, c, w, i As Short
		Dim MaxH, MaxW, EncounterCnt As Short
		Dim MinW, MinH As Short
		Dim Found, x, y, Plop As Short
		Dim nonCollisions, temp, Tick As Short
		Dim XDir, YDir As Short
		Dim EncounterX As Encounter
		Const geoDim As Short = 7
		Dim Morph(8, 7) As String
		Dim Geo(geoDim, geoDim) As Short
		' Set up Morphs
		Morph(0, 0) = "00011110"
		Morph(0, 1) = "00111110"
		Morph(0, 2) = "01111100"
		Morph(0, 3) = "01111100"
		Morph(0, 4) = "01111100"
		Morph(0, 5) = "01111100"
		Morph(0, 6) = "00111110"
		Morph(0, 7) = "00011110"
		
		Morph(1, 0) = "00011000"
		Morph(1, 1) = "00111100"
		Morph(1, 2) = "01111110"
		Morph(1, 3) = "11111111"
		Morph(1, 4) = "01111110"
		Morph(1, 5) = "00111100"
		Morph(1, 6) = "00011000"
		Morph(1, 7) = "00011000"
		
		Morph(2, 0) = "00111100"
		Morph(2, 1) = "01111110"
		Morph(2, 2) = "01100110"
		Morph(2, 3) = "11100111"
		Morph(2, 4) = "11100111"
		Morph(2, 5) = "01100110"
		Morph(2, 6) = "01111110"
		Morph(2, 7) = "00111100"
		
		Morph(3, 0) = "11100000"
		Morph(3, 1) = "00110000"
		Morph(3, 2) = "00111000"
		Morph(3, 3) = "00111000"
		Morph(3, 4) = "00011000"
		Morph(3, 5) = "00011000"
		Morph(3, 6) = "00001100"
		Morph(3, 7) = "00000111"
		
		Morph(4, 0) = "00001000"
		Morph(4, 1) = "00011000"
		Morph(4, 2) = "00010000"
		Morph(4, 3) = "11111111"
		Morph(4, 4) = "00001000"
		Morph(4, 5) = "00011000"
		Morph(4, 6) = "00010000"
		Morph(4, 7) = "00110000"
		
		Morph(5, 0) = "11111111"
		Morph(5, 1) = "11000000"
		Morph(5, 2) = "01100000"
		Morph(5, 3) = "00111000"
		Morph(5, 4) = "00011100"
		Morph(5, 5) = "00000110"
		Morph(5, 6) = "00000011"
		Morph(5, 7) = "11111111"
		
		Morph(6, 0) = "11001000"
		Morph(6, 1) = "11001100"
		Morph(6, 2) = "01100100"
		Morph(6, 3) = "00101111"
		Morph(6, 4) = "01111001"
		Morph(6, 5) = "01001111"
		Morph(6, 6) = "11001000"
		Morph(6, 7) = "01001111"
		
		Morph(7, 0) = "01111111"
		Morph(7, 1) = "00011000"
		Morph(7, 2) = "01110000"
		Morph(7, 3) = "01110000"
		Morph(7, 4) = "00011000"
		Morph(7, 5) = "00001100"
		Morph(7, 6) = "00011100"
		Morph(7, 7) = "00111110"
		
		Morph(8, 0) = "01111110"
		Morph(8, 1) = "11111111"
		Morph(8, 2) = "11111111"
		Morph(8, 3) = "11111111"
		Morph(8, 4) = "11111111"
		Morph(8, 5) = "11111111"
		Morph(8, 6) = "11111111"
		Morph(8, 7) = "01111110"
		
		' Size Map (+/- 10% of Defaults)
		MapX.Width = Greatest(MapX.DefaultWidth * ((Int(Rnd() * 20) + 90) / 100), 48)
		MapX.Height = Greatest(MapX.DefaultHeight * ((Int(Rnd() * 20) + 90) / 100), 48)
		' Loop until all rooms are populated
		Plop = False
		EncounterCnt = 1 : c = 0
		Do Until EncounterCnt > MapX.DefaultEncounters
			' Create new GeoMorph
			temp = 0
			c = c + 1
			Select Case Int(Rnd() * 100)
				Case 0 To 20 ' Bean, Cavern, Doughnut
					MinW = 0 : MaxW = Int(Rnd() * 3) + 3 : MinH = 0 : MaxH = Int(Rnd() * 3) + 3
					temp = Int(Rnd() * 3)
					EncounterX = MapX.AddEncounter
					EncounterX.GenerateUponEntry = True
					EncounterCnt = EncounterCnt + 1
				Case 97 To 99
					MinW = 0 : MaxW = Int(Rnd() * 2) + 6 : MinH = 0 : MaxH = Int(Rnd() * 2) + 6
					temp = 8
					EncounterX = MapX.AddEncounter
					EncounterX.GenerateUponEntry = True
					EncounterCnt = EncounterCnt + 1
				Case Else
					temp = Int(Rnd() * 5) + 3
					MinW = 0 : MaxW = 7
					MinH = 0 : MaxH = Int(Rnd() * 3) + 5
			End Select
			' Empty the Geo
			For w = 0 To 7 : For h = 0 To 7
					Geo(w, h) = 0
				Next h : Next w
			' Fill Geo with Morph
			For w = MinW To MaxW : For h = MinH To MaxH
					If Mid(Morph(temp, h), w + 1, 1) = "1" Then
						If temp < 3 Or temp = 8 Then ' Bean, Cavern, Doughnut
							Geo(w, h) = EncounterX.Index
						Else
							Geo(w, h) = -c
						End If
					End If
				Next h : Next w
			' Apply 0 to 3 90 degree rotations to the geo
			For i = 0 To Int(Rnd() * 4) - 1
				For w = 0 To 3
					For h = 0 To 3
						temp = Geo(w, h)
						Geo(w, h) = Geo(h, geoDim - w)
						Geo(h, geoDim - w) = Geo(geoDim - w, geoDim - h)
						Geo(geoDim - w, geoDim - h) = Geo(geoDim - h, w)
						Geo(geoDim - h, w) = temp
					Next h
				Next w
			Next i
			' If this is the first GeoMorph, then just plop it down.
			' Else, spin it into the middle of the map.
			If Plop = False Then
				x = Int(Rnd() * 3) + Int(MapX.Width / 2)
				y = Int(Rnd() * 3) + Int(MapX.Height / 2)
				Plop = True
			Else
				Tick = 0
				Do 
					Tick = Tick + 1
					XDir = Int(Rnd() * 3) - 1 : YDir = 0
					If XDir = 0 Then
						YDir = 1 - 2 * Int(Rnd() * 2)
						x = Int(Rnd() * (MapX.Width - 8)) + 1
						y = 0
						If YDir < 1 Then
							y = MapX.Height - 8
						End If
					Else
						x = 1
						y = Int(Rnd() * (MapX.Height - 8))
						If XDir < 1 Then
							x = MapX.Width - 8
						End If
					End If
					Found = False
					nonCollisions = 0
					Do Until Found = True Or x + XDir > MapX.Width - 8 Or x + XDir < 0 Or y + YDir > MapX.Height - 8 Or y + YDir < 0
						For w = 0 To 7 : For h = 0 To 7
								If Geo(w, h) <> 0 Then
									If MapX.EncPointer(x + XDir + w, y + YDir + h) <> 0 Then
										Found = True
									End If
								End If
							Next h : Next w
						If Not Found Then
							nonCollisions = nonCollisions + 1
							x = x + XDir
							y = y + YDir
						End If
					Loop 
				Loop Until (Found = True And nonCollisions > 0) Or Tick > 1000
			End If
			For w = 0 To 7 : For h = 0 To 7
					If Geo(w, h) <> 0 Then
						MapX.EncPointer(x + w, y + h) = Geo(w, h)
					End If
				Next h : Next w
		Loop 
	End Sub
	
	Private Sub BubbleTubes(ByRef MapX As Map)
		Dim h, c, w, i As Short
		Dim MaxH, MaxW, EncounterCnt As Short
		Dim MinW, MinH As Short
		Dim Found, x, y, Plop As Short
		Dim nonCollisions, temp, Tick As Short
		Dim XDir, YDir As Short
		Dim EncounterX As Encounter
		Dim Morph(5, 5) As String
		Dim Geo(5, 5) As Short
		' Set up Morphs
		Morph(0, 0) = "001100"
		Morph(0, 1) = "011110"
		Morph(0, 2) = "011110"
		Morph(0, 3) = "001100"
		Morph(1, 0) = "001100"
		Morph(1, 1) = "011110"
		Morph(1, 2) = "111111"
		Morph(1, 3) = "111111"
		Morph(1, 4) = "011110"
		Morph(1, 5) = "001100"
		Morph(2, 0) = "011110"
		Morph(2, 1) = "111111"
		Morph(2, 2) = "111111"
		Morph(2, 3) = "011110"
		Morph(3, 0) = "011100"
		Morph(3, 1) = "111110"
		Morph(3, 2) = "011100"
		Morph(4, 0) = "001000"
		Morph(4, 1) = "001000"
		Morph(4, 2) = "111000"
		Morph(4, 3) = "001111"
		Morph(4, 4) = "000100"
		Morph(4, 5) = "000100"
		
		' Size Map (+/- 10% of Defaults)
		MapX.Width = Greatest(MapX.DefaultWidth * ((Int(Rnd() * 20) + 90) / 100), 24)
		MapX.Height = Greatest(MapX.DefaultHeight * ((Int(Rnd() * 20) + 90) / 100), 24)
		' Loop until all rooms are populated
		Plop = False
		EncounterCnt = 1 : c = 0
		Do Until EncounterCnt > MapX.DefaultEncounters
			' Clear GeoMorph
			For w = 0 To 5 : For h = 0 To 5
					Geo(w, h) = 0
				Next h : Next w
			' Create new GeoMorph
			c = c + 1
			Select Case Int(Rnd() * 100)
				Case 0 To 15 ' Bubble
					EncounterX = MapX.AddEncounter
					EncounterX.GenerateUponEntry = True
					EncounterCnt = EncounterCnt + 1
					temp = Int(Rnd() * 4)
					For w = 0 To 5 : For h = 0 To 5
							If Mid(Morph(temp, h), w + 1, 1) = "1" Then
								Geo(w, h) = EncounterX.Index
							End If
						Next h : Next w
				Case 16 To 30 ' Spinner
					For w = 0 To 5 : For h = 0 To 5
							If Mid(Morph(4, h), w + 1, 1) = "1" Then
								Geo(w, h) = -c
							End If
						Next h : Next w
				Case Else ' Tube
					For h = 0 To Int(Rnd() * 4) + 2
						Geo(0, h) = -c
					Next h
			End Select
			' Fill Geo with Morph
			' Apply 0 to 3 90 degree rotations to the geo
			For i = 0 To Int(Rnd() * 4) - 1
				For w = 0 To 2
					For h = 0 To 2
						temp = Geo(w, h)
						Geo(w, h) = Geo(h, 5 - w)
						Geo(h, 5 - w) = Geo(5 - w, 5 - h)
						Geo(5 - w, 5 - h) = Geo(5 - h, w)
						Geo(5 - h, w) = temp
					Next h
				Next w
			Next i
			' If this is the first GeoMorph, then just plop it down.
			' Else, spin it into the middle of the map.
			If Plop = False Then
				x = Int(Rnd() * 3) + Int(MapX.Width / 2)
				y = Int(Rnd() * 3) + Int(MapX.Height / 2)
				Plop = True
			Else
				Tick = 0
				Do 
					Tick = Tick + 1
					XDir = Int(Rnd() * 3) - 1 : YDir = 0
					If XDir = 0 Then
						YDir = 1 - 2 * Int(Rnd() * 2)
						x = Int(Rnd() * (MapX.Width - 6)) + 1
						y = 0
						If YDir < 1 Then
							y = MapX.Height - 6
						End If
					Else
						x = 1
						y = Int(Rnd() * (MapX.Height - 6))
						If XDir < 1 Then
							x = MapX.Width - 6
						End If
					End If
					Found = False
					nonCollisions = 0
					Do Until Found = True Or x + XDir > MapX.Width - 6 Or x + XDir < 0 Or y + YDir > MapX.Height - 6 Or y + YDir < 0
						For w = 0 To 5 : For h = 0 To 5
								If Geo(w, h) <> 0 Then
									If MapX.EncPointer(x + XDir + w, y + YDir + h) <> 0 Then
										Found = True
									End If
								End If
							Next h : Next w
						If Not Found Then
							nonCollisions = nonCollisions + 1
							x = x + XDir
							y = y + YDir
						End If
					Loop 
				Loop Until (Found = True And nonCollisions > 0) Or Tick > 1000
			End If
			If Tick < 1000 Then
				For w = 0 To 5 : For h = 0 To 5
						If Geo(w, h) <> 0 Then
							MapX.EncPointer(x + w, y + h) = Geo(w, h)
						End If
					Next h : Next w
			End If
		Loop 
	End Sub
	
	
	Private Sub Rectangles(ByRef MapX As Map)
		Dim h, c, w, i As Short
		Dim MaxH, MaxW, EncounterCnt As Short
		Dim MinW, MinH As Short
		Dim Found, x, y, Plop As Short
		Dim nonCollisions, temp, Tick As Short
		Dim XDir, YDir As Short
		Dim EncounterX As Encounter
		Dim Geo(7, 7) As Short
		' Size Map (+/- 10% of Defaults)
		MapX.Width = Greatest(MapX.DefaultWidth * ((Int(Rnd() * 20) + 90) / 100), 24)
		MapX.Height = Greatest(MapX.DefaultHeight * ((Int(Rnd() * 20) + 90) / 100), 24)
		' Loop until all rooms are populated
		Plop = False
		EncounterCnt = 1 : c = 0
		Do Until EncounterCnt > MapX.DefaultEncounters
			' Clear GeoMorph
			For w = 0 To 7 : For h = 0 To 7
					Geo(w, h) = 0
				Next h : Next w
			' Create new GeoMorph
			c = c + 1
			Select Case Int(Rnd() * 100)
				Case 0 To 30 ' Room
					EncounterX = MapX.AddEncounter
					EncounterX.GenerateUponEntry = True
					EncounterCnt = EncounterCnt + 1
					MaxW = Int(Rnd() * 6) + 2 : MaxH = Int(Rnd() * 6) + 2
					For w = 0 To MaxW : For h = 0 To MaxH
							Geo(w, h) = EncounterX.Index
						Next h : Next w
				Case Else ' Empty Room
					MaxW = Greatest(Int(Rnd() * 4) - 2, 0) : MaxH = Int(Rnd() * 7) + 1
					For w = 0 To MaxW : For h = 0 To MaxH
							Geo(w, h) = -c
						Next h : Next w
			End Select
			' Apply 0 to 3 90 degree rotations to the geo
			For i = 0 To Int(Rnd() * 4) - 1
				For w = 0 To 3
					For h = 0 To 3
						temp = Geo(w, h)
						Geo(w, h) = Geo(h, 7 - w)
						Geo(h, 7 - w) = Geo(7 - w, 7 - h)
						Geo(7 - w, 7 - h) = Geo(7 - h, w)
						Geo(7 - h, w) = temp
					Next h
				Next w
			Next i
			' If this is the first GeoMorph, then just plop it down.
			' Else, spin it into the middle of the map.
			If Plop = False Then
				x = Int(Rnd() * 3) + Int(MapX.Width / 2)
				y = Int(Rnd() * 3) + Int(MapX.Height / 2)
				Plop = True
			Else
				Tick = 0
				Do 
					Tick = Tick + 1
					XDir = Int(Rnd() * 3) - 1 : YDir = 0
					If XDir = 0 Then
						YDir = 1 - 2 * Int(Rnd() * 2)
						x = Int(Rnd() * (MapX.Width - 8)) + 1
						y = 0
						If YDir < 1 Then
							y = MapX.Height - 8
						End If
					Else
						x = 1
						y = Int(Rnd() * (MapX.Height - 8))
						If XDir < 1 Then
							x = MapX.Width - 8
						End If
					End If
					Found = False
					nonCollisions = 0
					Do Until Found = True Or x + XDir > MapX.Width - 8 Or x + XDir < 0 Or y + YDir > MapX.Height - 8 Or y + YDir < 0
						For w = 0 To 7 : For h = 0 To 7
								If Geo(w, h) <> 0 Then
									If MapX.EncPointer(x + XDir + w, y + YDir + h) <> 0 Then
										Found = True
									End If
								End If
							Next h : Next w
						If Not Found Then
							nonCollisions = nonCollisions + 1
							x = x + XDir
							y = y + YDir
						End If
					Loop 
				Loop Until (Found = True And nonCollisions > 0) Or Tick > 1000
			End If
			If Tick < 1000 Then
				For w = 0 To 7 : For h = 0 To 7
						If Geo(w, h) <> 0 Then
							MapX.EncPointer(x + w, y + h) = Geo(w, h)
						End If
					Next h : Next w
			End If
		Loop 
	End Sub
	
	Private Sub HugeCavern(ByRef MapX As Map)
		Dim h, c, w, i As Short
		Dim MaxH, MaxW, EncounterCnt As Short
		Dim MinW, MinH As Short
		Dim Found, x, y, Plop As Short
		Dim nonCollisions, temp, Tick As Short
		Dim XDir, YDir As Short
		Dim EncounterX As Encounter
		Dim Geo(5, 5) As Short
		' Size Map (+/- 10% of Defaults)
		MapX.Width = Greatest(MapX.DefaultWidth * ((Int(Rnd() * 20) + 90) / 100), 32)
		MapX.Height = Greatest(MapX.DefaultHeight * ((Int(Rnd() * 20) + 90) / 100), 32)
		' Loop until all rooms are populated
		EncounterCnt = 1 : c = 1
		' For first room, create center piece Cavern (Spatter squares in an area)
		x = Int(Rnd() * MapX.Width / 8) + MapX.Width / 4 : MinW = Int(Rnd() * (MapX.Width - x)) + 1 : MaxW = MinW + x
		y = Int(Rnd() * MapX.Height / 8) + MapX.Height / 4 : MinH = Int(Rnd() * (MapX.Height - y)) + 1 : MaxH = MinH + y
		For i = 0 To ((MaxW - MinW) * (MaxH - MinH)) / 10
			x = Int(Rnd() * (MaxW - MinW)) + MinW : y = Int(Rnd() * (MaxH - MinH)) + MinH
			For XDir = 0 To 3 : For YDir = 0 To 3
					If x + XDir < MapX.Width And y + YDir < MapX.Height Then
						MapX.EncPointer(x + XDir, y + YDir) = -1
					End If
				Next YDir : Next XDir
		Next i
		' Spin out other areas
		Do Until EncounterCnt > Greatest((MapX.DefaultEncounters), 10)
			' Clear GeoMorph
			For w = 0 To 5 : For h = 0 To 5
					Geo(w, h) = 0
				Next h : Next w
			' Create new GeoMorph
			c = c + 1
			Select Case Int(Rnd() * 100)
				Case 0 To 25 ' Room
					EncounterX = MapX.AddEncounter
					EncounterX.GenerateUponEntry = True
					EncounterCnt = EncounterCnt + 1
					MaxW = Int(Rnd() * 4) + 2 : MaxH = Int(Rnd() * 3) + 3
					For w = 0 To MaxW : For h = 0 To MaxH
							Geo(w, h) = EncounterX.Index
						Next h : Next w
				Case Else ' Passaage
					MaxW = 0 : MaxH = Int(Rnd() * 3) + 1
					For w = 0 To MaxW : For h = 0 To MaxH
							Geo(w, h) = -c
						Next h : Next w
			End Select
			' Apply 0 to 3 90 degree rotations to the geo
			For i = 0 To Int(Rnd() * 4) - 1
				For w = 0 To 2
					For h = 0 To 2
						temp = Geo(w, h)
						Geo(w, h) = Geo(h, 5 - w)
						Geo(h, 5 - w) = Geo(5 - w, 5 - h)
						Geo(5 - w, 5 - h) = Geo(5 - h, w)
						Geo(5 - h, w) = temp
					Next h
				Next w
			Next i
			' Spin it into the middle of the map.
			Tick = 0
			Do 
				Tick = Tick + 1
				XDir = Int(Rnd() * 3) - 1 : YDir = 0
				If XDir = 0 Then
					YDir = 1 - 2 * Int(Rnd() * 2)
					x = Int(Rnd() * (MapX.Width - 6)) + 1
					y = 0
					If YDir < 1 Then
						y = MapX.Height - 6
					End If
				Else
					x = 1
					y = Int(Rnd() * (MapX.Height - 6))
					If XDir < 1 Then
						x = MapX.Width - 6
					End If
				End If
				Found = False
				nonCollisions = 0
				Do Until Found = True Or x + XDir > MapX.Width - 6 Or x + XDir < 0 Or y + YDir > MapX.Height - 6 Or y + YDir < 0
					For w = 0 To 5 : For h = 0 To 5
							If Geo(w, h) <> 0 Then
								If MapX.EncPointer(x + XDir + w, y + YDir + h) <> 0 Then
									Found = True
								End If
							End If
						Next h : Next w
					If Not Found Then
						nonCollisions = nonCollisions + 1
						x = x + XDir
						y = y + YDir
					End If
				Loop 
			Loop Until (Found = True And nonCollisions > 0) Or Tick > 1000
			If Tick < 1000 Then
				For w = 0 To 5 : For h = 0 To 5
						If Geo(w, h) <> 0 Then
							MapX.EncPointer(x + w, y + h) = Geo(w, h)
						End If
					Next h : Next w
			End If
		Loop 
	End Sub
	
	Public Sub FillMap(ByRef MapX As Map)
		Dim x, y As Short
		Dim ThemeX As Theme
		For	Each ThemeX In MapX.Themes
			ThemeMap(MapX, ThemeX)
		Next ThemeX
		' Clean up any other EncPointers which are invalid
		For x = 0 To MapX.Width : For y = 0 To MapX.Height
				If MapX.EncPointer(x, y) < 0 Then
					MapX.EncPointer(x, y) = 0
				End If
			Next y : Next x
	End Sub
	
	Public Sub ThemeMap(ByRef MapX As Map, ByRef ThemeX As Theme)
		Dim EncounterX, BlankEncounter As Encounter
		Dim EndX, StartX, StepX As Short
		Dim c, x, y, Tick As Short
		Dim Available, Mine As Short
		Dim CreatureX As Creature
		Dim ItemX As Item
		Dim ThemeZ As Theme
		' Clear the number of Encounters counted for each Theme
		For	Each ThemeZ In MapX.Themes
			ThemeZ.EncounterCount = 0
		Next ThemeZ
		' Empty existing Encounters with this Theme, Count Encounters available
		Dim MyEnc(MapX.Encounters.Count()) As Short
		Available = 0 : Mine = 0
		BlankEncounter = New Encounter
		BlankEncounter.ParentTheme = ThemeX.Index
		For	Each EncounterX In MapX.Encounters
			If EncounterX.ParentTheme = 0 And EncounterX.GenerateUponEntry = True Then
				Available = Available + 1
			ElseIf EncounterX.ParentTheme > 0 Then 
				If EncounterX.ParentTheme = ThemeX.Index Then
					EncounterX.Copy(BlankEncounter)
					MyEnc(Mine) = EncounterX.Index
					Mine = Mine + 1
				End If
				'UPGRADE_WARNING: Couldn't resolve default property of object MapX.Themes(R & EncounterX.ParentTheme).EncounterCount. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
				'UPGRADE_WARNING: Couldn't resolve default property of object MapX.Themes().EncounterCount. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
				MapX.Themes.Item("R" & EncounterX.ParentTheme).EncounterCount = MapX.Themes.Item("R" & EncounterX.ParentTheme).EncounterCount + 1
			End If
		Next EncounterX
		' Distribute available Encounters evenly among Themes
		Tick = 0
		Do Until Available < 1 Or Tick > 500
			For	Each ThemeZ In MapX.Themes
				If ThemeZ.EncounterCount < ThemeZ.Coverage Then
					ThemeZ.EncounterCount = ThemeZ.EncounterCount + 1
					Available = Available - 1
				End If
			Next ThemeZ
			Tick = Tick + 1
		Loop 
		' Search for Encounter to Theme
		If Int(Rnd() * 100) < 50 Then
			StartX = 0 : EndX = MapX.Width : StepX = 1
		Else
			StartX = MapX.Width : EndX = 0 : StepX = -1
		End If
		' Take hold of any Encounter without a Theme and that wants to be Generated (not to exceed coverage)
		For	Each EncounterX In MapX.Encounters
			If EncounterX.ParentTheme = 0 And EncounterX.GenerateUponEntry = True And Mine < ThemeX.EncounterCount Then
				EncounterX.ParentTheme = ThemeX.Index
				MyEnc(Mine) = EncounterX.Index
				Mine = Mine + 1
			End If
		Next EncounterX
		' Shuffle Encounter list on Theme to scatter randomly
		Dim Enc(ThemeX.Encounters.Count()) As Short
		For c = 0 To ThemeX.Encounters.Count() - 1
			'UPGRADE_WARNING: Couldn't resolve default property of object ThemeX.Encounters().Index. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
			Enc(c) = ThemeX.Encounters.Item(c + 1).Index
		Next c
		For c = 0 To ThemeX.Encounters.Count() - 1
			x = Int(Rnd() * ThemeX.Encounters.Count())
			y = Enc(c)
			Enc(c) = Enc(x)
			Enc(x) = y
		Next c
		' Shuffle Encounter list on Map (MyEnc)
		For c = 0 To ThemeX.EncounterCount - 1
			x = Int(Rnd() * ThemeX.EncounterCount)
			y = MyEnc(c)
			MyEnc(c) = MyEnc(x)
			MyEnc(x) = y
		Next c
		' Populate all Encounters for this Theme
		For c = 0 To Least(ThemeX.Coverage - 1, ThemeX.EncounterCount - 1)
			' If have Theme Encounters left to place, do so. Else, generate
			EncounterX = MapX.Encounters.Item("E" & MyEnc(c))
			If c < ThemeX.Encounters.Count() Then
				'UPGRADE_WARNING: Couldn't resolve default property of object ThemeX.Encounters(). Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
				EncounterX.Copy(ThemeX.Encounters.Item("E" & Enc(c)))
				EncounterX.ParentTheme = ThemeX.Index
				MakeEncounter(MapX, EncounterX)
			Else
				EncounterX.ReGenCreatures = True
				EncounterX.ReGenItems = True
				EncounterX.ReGenTriggers = True
				EncounterX.ReGenDescription = True
				MakeEncounter(MapX, EncounterX)
			End If
		Next c
	End Sub
	
	Public Sub MakeItem(ByRef ItemX As Item)
		Dim c As Short
		Dim ItemZ As Item
		' Only Make an Item if this is from the Random Family
		If ItemX.Family < 9 Or ItemX.Family > 13 Then
			Exit Sub
		End If
		Randomize()
		' Clear the Item
		c = ItemX.Family
		ItemX.Copy(New Item)
		ItemX.Family = c
		' If ItemX is of a random family, then make it
		Select Case ItemX.Family
			Case 9 ' Random Weapon
				ItemX.Family = 5 ' Weapon (Melee)
				ItemX.WearType = 10
				Select Case Int(Rnd() * 9)
					Case 0
						ItemX.Name = "Sword"
						ItemX.Value = Int(Rnd() * 5) + 10
						ItemX.Weight = 5 : ItemX.Bulk = 5
						ItemX.PictureFile = "sword.bmp"
						ItemX.DamageType = 1
						ItemX.Damage = 8 ' 3d6
						ItemX.ActionPoints = 10
					Case 1 : ItemX.Name = "Club"
						ItemX.Value = Int(Rnd() * 10) + 3
						ItemX.Weight = 5 : ItemX.Bulk = 5
						ItemX.PictureFile = "club.bmp"
						ItemX.DamageType = 2
						ItemX.Damage = 2 ' 2d4
						ItemX.ActionPoints = 9
					Case 2 : ItemX.Name = "Axe"
						ItemX.Value = Int(Rnd() * 7) + 8
						ItemX.Weight = 7 : ItemX.Bulk = 5
						ItemX.PictureFile = "hatchet.bmp"
						ItemX.DamageType = 1
						ItemX.Damage = 2 ' 3d4
						ItemX.ActionPoints = 11
					Case 3 : ItemX.Name = "Staff"
						ItemX.Value = Int(Rnd() * 5) + 5
						ItemX.Weight = 3 : ItemX.Bulk = 7
						If Int(Rnd() * 2) = 0 Then
							ItemX.PictureFile = "staff1.bmp"
						Else
							ItemX.PictureFile = "staff2.bmp"
						End If
						ItemX.DamageType = 2
						ItemX.Damage = 2 ' 2d4
						ItemX.ActionPoints = 8
					Case 4 : ItemX.Name = "Spear"
						ItemX.Value = Int(Rnd() * 10) + 5
						ItemX.Weight = 6 : ItemX.Bulk = 10
						ItemX.PictureFile = "spear.bmp"
						ItemX.DamageType = 1
						ItemX.Damage = 12 ' 2d8
						ItemX.ActionPoints = 12
					Case 5 : ItemX.Name = "Dagger"
						ItemX.Value = Int(Rnd() * 5) + 1
						ItemX.Weight = 2 : ItemX.Bulk = 2
						ItemX.PictureFile = "knife.bmp"
						ItemX.DamageType = 1
						ItemX.Damage = 7 ' 2d6
						ItemX.ActionPoints = 7
					Case 6 : ItemX.Name = "Mace"
						ItemX.Value = Int(Rnd() * 8) + 3
						ItemX.Weight = 5 : ItemX.Bulk = 5
						ItemX.PictureFile = "mace.bmp"
						ItemX.DamageType = 2
						ItemX.Damage = 3 ' 3d4
						ItemX.ActionPoints = 10
					Case 7 : ItemX.Name = "Hammer"
						ItemX.Value = Int(Rnd() * 10) + 5
						ItemX.Weight = 5 : ItemX.Bulk = 5
						ItemX.PictureFile = "hammer.bmp"
						ItemX.DamageType = 2
						ItemX.Damage = 7 ' 2d6
						ItemX.ActionPoints = 10
					Case 8 : ItemX.Name = "Rod"
						ItemX.Value = Int(Rnd() * 3) + 3
						ItemX.Weight = 4 : ItemX.Bulk = 4
						ItemX.PictureFile = "wand" & Int(Rnd() * 8) + 1 & ".bmp"
						ItemX.DamageType = 2
						ItemX.Damage = 11 ' 1d8
						ItemX.ActionPoints = 7
				End Select
				If Int(Rnd() * 20) > 10 Then
					Select Case Int(Rnd() * 13)
						Case 0 : ItemX.Name = "Long " & ItemX.Name
							ItemX.ActionPoints = ItemX.ActionPoints + 1
						Case 1 : ItemX.Name = "Short " & ItemX.Name
							If ItemX.Damage <> 11 Then
								ItemX.Damage = ItemX.Damage - 1
							Else
								ItemX.Damage = 6
							End If
						Case 2 : ItemX.Name = "War " & ItemX.Name
							ItemX.AttackBonus = 1
							ItemX.Value = ItemX.Value * 2
						Case 3 : ItemX.Name = "Spiked " & ItemX.Name
							If ItemX.Damage < 6 Then
								ItemX.Damage = ItemX.Damage + 1
							End If
						Case 4 : ItemX.Name = "Battle " & ItemX.Name
							ItemX.AttackBonus = 2
							ItemX.Damage = ItemX.Damage + 1
							ItemX.Value = ItemX.Value * 3
							ItemX.ActionPoints = ItemX.ActionPoints + 1
						Case 5 : ItemX.Name = "Iron " & ItemX.Name
						Case 6 : ItemX.Name = "Steel " & ItemX.Name
						Case 7 : ItemX.Name = "Studded " & ItemX.Name
						Case 8 : ItemX.Name = "Wooden " & ItemX.Name
							ItemX.DamageType = 2
							ItemX.Damage = Greatest(ItemX.Damage - 5, 0)
							ItemX.ActionPoints = ItemX.ActionPoints - 1
							ItemX.Value = Int(ItemX.Value / 2)
						Case 9 : ItemX.Name = "Silver " & ItemX.Name
							ItemX.DamageType = 3
							ItemX.ActionPoints = ItemX.ActionPoints - 1
						Case 10 : ItemX.Name = "Holy " & ItemX.Name
							ItemX.DamageType = 6
							ItemX.Value = ItemX.Value * 2
						Case 11 : ItemX.Name = "Unholy " & ItemX.Name
							ItemX.DamageType = 5
							ItemX.Value = ItemX.Value * 2
						Case 12 : ItemX.Name = "Ancient " & ItemX.Name
							ItemX.AttackBonus = Int(Rnd() * 5)
							ItemX.Damage = ItemX.Damage + Int(Rnd() * 3)
							ItemX.Value = ItemX.Value * (10 + Int(Rnd() * 5))
					End Select
				End If
			Case 10 ' Random Armor
				ItemX.Family = 4
				Select Case Int(Rnd() * 4)
					Case 0
						ItemX.Name = "Leather Armor"
						ItemX.WearType = 0
						ItemX.PictureFile = "armor2.bmp"
						ItemX.Resistance = 25
						ItemX.Defense = -1
						ItemX.ActionPoints = 6
						ItemX.Value = 50
						ItemX.Weight = 8
						ItemX.Bulk = 25
					Case 1
						ItemX.Name = "Chain Mail Armor"
						ItemX.WearType = 0
						ItemX.PictureFile = "armor1.bmp"
						ItemX.Resistance = 50
						ItemX.Defense = -2
						ItemX.ActionPoints = 8
						ItemX.Value = 100
						ItemX.Weight = 24
						ItemX.Bulk = 30
					Case 2
						ItemX.Name = "Wooden Shield"
						ItemX.WearType = 5
						ItemX.PictureFile = "shield2.bmp"
						ItemX.Resistance = 50
						ItemX.Defense = -2
						ItemX.ActionPoints = 8
						ItemX.Value = 100
						ItemX.Weight = 24
						ItemX.Bulk = 30
					Case 3
						ItemX.Name = "Iron Shield"
						ItemX.WearType = 5
						ItemX.PictureFile = "shield1.bmp"
						ItemX.Resistance = 50
						ItemX.Defense = -2
						ItemX.ActionPoints = 8
						ItemX.Value = 100
						ItemX.Weight = 24
						ItemX.Bulk = 30
				End Select
			Case 11 ' Random Gems
				' Random Gems are the word Gem with an adjective
				ItemX.Family = 0
				ItemX.Name = RandomGem & " Gem"
				Select Case Int(Rnd() * 6)
					Case 0
						ItemX.PictureFile = "gem1.bmp"
					Case 1
						ItemX.PictureFile = "gem2.bmp"
					Case 2
						ItemX.PictureFile = "gem3.bmp"
					Case 3
						ItemX.PictureFile = "gem4.bmp"
					Case 4
						ItemX.PictureFile = "gem5.bmp"
					Case 5
						ItemX.PictureFile = "gems.bmp"
				End Select
				ItemX.Value = 0
				c = 0
				Do 
					c = c + Int(Rnd() * 20)
				Loop Until Int(Rnd() * 20) < 10
				ItemX.Count = Int(Rnd() * 20) + 1
				ItemX.Value = (c + 5) * ItemX.Count
				ItemX.Weight = 0 : ItemX.Bulk = 0
				ItemX.CanCombine = True
			Case 12 ' Random Treasure
				Select Case Int(Rnd() * 3)
					Case 0 ' Random Gem
						ItemX.Family = 11
						MakeItem(ItemX)
					Case 1 ' Random Money
						ItemX.Family = 2
						' [Titi 2.4.8] check if gold coins are the world currency
						If Left(WorldNow.Money, InStr(WorldNow.Money, "|") - 1) = "Gold piece" Or Left(WorldNow.Money, InStr(WorldNow.Money, "|") - 1) = "Gold coin" Then
							Select Case Int(Rnd() * 4)
								Case 0
									ItemX.Name = "Gold Coin"
									ItemX.Value = 1
									ItemX.PictureFile = "gold1.bmp"
								Case 1
									ItemX.Name = "Silver Coin"
									ItemX.Value = 1
									ItemX.PictureFile = "silver1.bmp"
								Case 2
									ItemX.Name = "Copper Coin"
									ItemX.Value = 1
									ItemX.PictureFile = "gold1.bmp"
								Case 3
									ItemX.Name = "Platinum Coin"
									ItemX.Value = 3
									ItemX.PictureFile = "silver1.bmp"
							End Select
						Else
							' use the current money
							ItemX.Name = Left(WorldNow.Money, InStr(WorldNow.Money, "|") - 1)
							ItemX.Value = 1
							ItemX.PictureFile = Right(WorldNow.Money, Len(WorldNow.Money) - InStr(WorldNow.Money, "|"))
						End If
						c = 0
						Do 
							c = Int(Rnd() * 30)
							ItemX.Count = ItemX.Count + c
						Loop Until c < 20
						ItemX.Value = ItemX.Value * ItemX.Count
						ItemX.CanCombine = True
						If ItemX.NameText = "Copper Coin" Then
							ItemX.Value = ItemX.Value / 5
						ElseIf ItemX.NameText = "Silver Coin" Then 
							ItemX.Value = ItemX.Value / 2
						End If
					Case 2 ' Random Chest
						ItemX.Family = 0
						ItemX.Name = "Wooden Chest"
						ItemX.Weight = 10 : ItemX.Bulk = 25
						ItemX.Value = Int(Rnd() * 25) + 5
						ItemX.Capacity = 25
						ItemX.PictureFile = "chest1.bmp"
						c = 0
						Do Until Int(Rnd() * 100) < 25
							ItemZ = New Item
							ItemZ.Family = 9 + Int(Rnd() * 5)
							MakeItem(ItemZ)
							If ItemZ.Bulk + c < ItemX.Capacity Then
								ItemX.AddItem.Copy(ItemZ)
								c = c + ItemZ.Bulk
							Else
								Exit Do
							End If
						Loop 
				End Select
			Case 13 ' Random Junk
				ItemX.Family = 0
				Select Case Int(Rnd() * 10)
					Case 0
						ItemX.Name = "Rock"
						ItemX.Weight = 1 : ItemX.Bulk = 1
						ItemX.PictureFile = "rock.bmp"
					Case 1
						ItemX.Name = "Stone"
						ItemX.Weight = 2 : ItemX.Bulk = 3
						ItemX.PictureFile = "stone.bmp"
					Case 2
						ItemX.Name = "Bottle"
						ItemX.Weight = 1 : ItemX.Bulk = 2
						ItemX.ActionPoints = 8 : ItemX.Damage = 2
						ItemX.PictureFile = "flask4.bmp"
					Case 3
						ItemX.Name = "Worthless Gem"
						ItemX.Count = 1
						ItemX.Weight = 0 : ItemX.Bulk = 0
						ItemX.PictureFile = "gem1.bmp"
					Case 4
						ItemX.Name = "Metal Tray"
						ItemX.Weight = 1 : ItemX.Bulk = 1
						ItemX.PictureFile = "tray.bmp"
					Case 5
						ItemX.Name = "Blank Parchment"
						ItemX.Weight = 0 : ItemX.Bulk = 2
						ItemX.PictureFile = "scroll.bmp"
					Case 6
						ItemX.Name = "Book"
						ItemX.Weight = 1 : ItemX.Bulk = 1
						ItemX.PictureFile = "book.bmp"
					Case 7
						ItemX.Name = "Worthless Gem"
						ItemX.Count = Int(Rnd() * 10)
						ItemX.Weight = 0 : ItemX.Bulk = 0
						ItemX.PictureFile = "gem5.bmp"
					Case 8
						ItemX.Name = "Worthless Gem"
						ItemX.Count = Int(Rnd() * 10)
						ItemX.Weight = 0 : ItemX.Bulk = 0
						ItemX.PictureFile = "gem4.bmp"
					Case 9
						ItemX.Name = "Block of Wood"
						ItemX.Weight = 1 : ItemX.Bulk = 1
						ItemX.PictureFile = "wood.bmp"
				End Select
				ItemX.Value = Int(Rnd() * 5) + 1
		End Select
	End Sub
	
	Public Sub MakeCreature(ByRef CreatureX As Creature)
		CreatureX.HPNow = CreatureX.HPNow * (1.2 - (Int(Rnd() * 30) / 100))
	End Sub
	
	Public Sub MakeEncounter(ByRef MapX As Map, ByRef EncounterX As Encounter)
		Dim i, c, Cnt As Short
		Dim ThemeX As Theme
		Dim CreatureX As Creature
		Dim ItemX As Item
		Dim TriggerX As Trigger
		' Check and set ParentTheme
		If EncounterX.ParentTheme = 0 Then
			Exit Sub
		Else
			ThemeX = MapX.Themes.Item("R" & EncounterX.ParentTheme)
		End If
		' Generate Creatures
		If EncounterX.ReGenCreatures = True And ThemeX.Creatures.Count() > 0 Then
			' Count is based on Theme Party Size and Party Avg Level (Max of 12)
			c = 0 : i = Int(Rnd() * ThemeX.Creatures.Count()) + 1 : Cnt = 0
			Do Until c >= ThemeX.EncounterPoints Or Cnt > 12
				CreatureX = EncounterX.AddCreature
				'UPGRADE_WARNING: Couldn't resolve default property of object ThemeX.Creatures(). Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
				CreatureX.Copy(ThemeX.Creatures.Item(i))
				MakeCreature(CreatureX)
				c = c + CreatureX.EncounterPoints
				Cnt = Cnt + 1
			Loop 
		End If
		' Generate Items
		If EncounterX.ReGenItems = True And ThemeX.Items.Count() > 0 Then
			' Count is based on Theme Party Size and Party Avg Level (Max of 10)
			c = 0 : Cnt = 0
			Do Until c >= ThemeX.EncounterPoints Or Cnt > 10
				ItemX = EncounterX.AddItem
				'UPGRADE_WARNING: Couldn't resolve default property of object ThemeX.Items(). Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
				ItemX.Copy(ThemeX.Items.Item(Int(Rnd() * ThemeX.Items.Count()) + 1))
				MakeItem(ItemX)
				c = c + ItemX.EncounterPoints
				Cnt = Cnt + 1
			Loop 
		End If
		' Generate Triggers
		If EncounterX.ReGenTriggers = True And ThemeX.Triggers.Count() > 0 Then
			TriggerX = EncounterX.AddTrigger
			'UPGRADE_WARNING: Couldn't resolve default property of object ThemeX.Triggers(). Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
			TriggerX.Copy(ThemeX.Triggers.Item(Int(Rnd() * ThemeX.Triggers.Count()) + 1))
		End If
		' Position Creatures and Items
		For	Each CreatureX In EncounterX.Creatures
			CreatureX.MapX = 0 : CreatureX.MapY = 0
			PositionCreature(MapX, EncounterX, CreatureX)
		Next CreatureX
		For	Each ItemX In EncounterX.Items
			ItemX.MapX = 0 : ItemX.MapY = 0
			PositionItem(MapX, EncounterX, ItemX)
		Next ItemX
		' Generate Description
		If EncounterX.ReGenDescription = True Then
			If ThemeX.Factoids.Count() > 0 Then
			Else
				EncounterX.Name = ThemeX.Name & " #" & EncounterX.Index
				If EncounterX.Creatures.Count() > 1 Then
					'UPGRADE_WARNING: Couldn't resolve default property of object EncounterX.Creatures().Name. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
					EncounterX.FirstEntry = "You encounter some " & EncounterX.Creatures.Item(1).Name & "s."
				ElseIf EncounterX.Creatures.Count() = 1 Then 
					'UPGRADE_WARNING: Couldn't resolve default property of object EncounterX.Creatures().Name. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
					EncounterX.FirstEntry = "You encounter one " & EncounterX.Creatures.Item(1).Name & "."
				End If
			End If
		End If
		' Unset GenerateUponEntry
		EncounterX.GenerateUponEntry = False
	End Sub
	
	Public Sub PositionCreature(ByRef MapX As Map, ByRef EncounterX As Encounter, ByRef CreatureX As Creature)
		Dim y, x, c As Short
		Dim CreatureZ As Creature
		Dim AtX(1) As Short
		Dim AtY(1) As Short
		Dim Cnt, Tick As Short
		For y = 0 To MapX.Height : For x = 0 To MapX.Width
				If MapX.EncPointer(x, y) = EncounterX.Index Then
					Cnt = Cnt + 1
					ReDim Preserve AtX(Cnt)
					ReDim Preserve AtY(Cnt)
					AtX(Cnt) = x : AtY(Cnt) = y
				End If
			Next x : Next y
		Do Until Tick > 100
			c = Int(Rnd() * Cnt) + 1
			CreatureX.MapX = AtX(c) : CreatureX.MapY = AtY(c)
			If PositionCreatureSpot(EncounterX, CreatureX) = True Then
				Tick = 101
			Else
				Tick = Tick + 1
			End If
		Loop 
	End Sub
	
	Public Function PositionCreatureSpot(ByRef EncounterX As Encounter, ByRef CreatureX As Creature) As Short
		' This positions CreatureX in a valid TileSpot in its EncounterX or returns False
		Dim c As Short
		Dim Spots(8) As Short
		Dim Tick As Short
		Dim CreatureZ As Creature
		PositionCreatureSpot = False
		For	Each CreatureZ In EncounterX.Creatures
			If CreatureZ.MapX = CreatureX.MapX And CreatureZ.MapY = CreatureX.MapY And CreatureZ.Index <> CreatureX.Index Then
				Spots(CreatureZ.TileSpot) = True
			End If
		Next CreatureZ
		For c = 0 To 4
			If Spots(c) = False Then
				CreatureX.TileSpot = c
				PositionCreatureSpot = True
				Exit For
			End If
		Next c
	End Function
	
	Public Sub PositionItem(ByRef MapX As Map, ByRef EncounterX As Encounter, ByRef ItemX As Item)
		Dim y, x, c As Short
		Dim ItemZ As Item
		Dim AtX(1) As Short
		Dim AtY(1) As Short
		Dim Cnt, Tick As Short
		For y = 0 To MapX.Height : For x = 0 To MapX.Width
				If MapX.EncPointer(x, y) = EncounterX.Index Then
					Cnt = Cnt + 1
					ReDim Preserve AtX(Cnt)
					ReDim Preserve AtY(Cnt)
					AtX(Cnt) = x : AtY(Cnt) = y
				End If
			Next x : Next y
		Do Until Tick > 100
			c = Int(Rnd() * (Cnt + 1))
			ItemX.MapX = AtX(c) : ItemX.MapY = AtY(c)
			If PositionItemSpot(EncounterX, ItemX) = True Then
				Tick = 101
			Else
				Tick = Tick + 1
			End If
		Loop 
	End Sub
	
	Public Function PositionItemSpot(ByRef EncounterX As Encounter, ByRef ItemX As Item) As Short
		' This positions ItemX in a valid TileSpot in its EncounterX or returns False
		Dim c As Short
		Dim Spots(8) As Short
		Dim Tick As Short
		Dim ItemZ As Item
		PositionItemSpot = False
		For	Each ItemZ In EncounterX.Items
			If ItemZ.MapX = ItemX.MapX And ItemZ.MapY = ItemX.MapY And ItemZ.Index <> ItemX.Index Then
				Spots(ItemZ.TileSpot) = True
			End If
		Next ItemZ
		For c = 0 To 4
			If Spots(c) = False Then
				ItemX.TileSpot = c
				PositionItemSpot = True
				Exit For
			End If
		Next c
	End Function
End Module