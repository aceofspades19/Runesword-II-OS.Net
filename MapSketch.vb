Option Strict Off
Option Explicit On
Friend Class MapSketch
	
	' Used to control versions
	Private myVersion As Short
	' Unique number for this MapSketch inside an UberWizard
	Private myIndex As Short
	' Name of the Map it represents
	Private myMapName As String
	' Comments about the Map it represents
	Private myMapComments As String
	' FullPath of where the Map is located (Area file, *.rsa)
	Private myFullPath As String
	' Created for Party Size
	' Bit1-2 - 0 = Solo, 1 = 2-3, 2 = 4-6, 3 = 7-9
	' With Average Party Level
	' Bit3-4 - 0 = 1-3, 1 = 4-6, 2 = 7-10, 3 = 10+
	' Bit5-16   Not Used
	Dim myStyle As Short
	' Bit1  - On=IsUsed, Off=Not Used
	' Bit2  - On=GenerateUponEntry, Off=Not
	' Bit3  - On=NeedsThemes, Off=Not
	' Bit4  - On=IsMainMap, Off=Not
	' Bit5-8 - MapStyle
	' Bit9  - On=IsSelected to be used, Off=Not
	' Bit10 - On=GenerateThemes, Off=Not
	' Bit11-16 Not Used
	Private myFlags As Short
	' Number of times it can be used
	Private myUses As Short
	' Save settings for a random build of a Map
	Private mySize As Short
	' Map Set Size
	Private myTotalSize As Integer
	' If a Tome Map: This is the TomeIndex in UberWiz
	Private myTomeIndex As Short
	' Unknown use
	Private myEncounterCount As Short
	' Map Index in Area where this MapSketch is located
	Private myMapIndexInArea As Short
	' Map Index in Area where this MapSketch is located
	Private myAreaIndexInArea As Short
	' Area Index where MapSketch is placed when create Tome
	Private myTomeIndexInArea As Short
	' Tome Index where MapSketch is placed when create Tome
	Private myAreaIndex As Short
	' Map Index where MapSketch is placed in Area
	Private myMapIndex As Short
	' Collection of the EntryPoints in the Map. Connected up as we progress and build.
	Private myEntryPoints As New Collection
	' Collection of ThemeSketchs (if allowed in this Map)
	Private myThemeSketchs As New Collection
	' Collection of Themes used as Quest/Major Theme parts
	Private myThemeQuests As New Collection
	
	
	Public Property Version() As Short
		Get
			Version = myVersion
		End Get
		Set(ByVal Value As Short)
			myVersion = Value
		End Set
	End Property
	
	
	Public Property Index() As Short
		Get
			Index = myIndex
		End Get
		Set(ByVal Value As Short)
			myIndex = Value
		End Set
	End Property
	
	
	Public Property Uses() As Short
		Get
			Uses = myUses
		End Get
		Set(ByVal Value As Short)
			myUses = Value
		End Set
	End Property
	
	
	Public Property TomeIndex() As Short
		Get
			TomeIndex = myTomeIndex
		End Get
		Set(ByVal Value As Short)
			myTomeIndex = Value
		End Set
	End Property
	
	
	Public Property Size() As Short
		Get
			Size = mySize
		End Get
		Set(ByVal Value As Short)
			mySize = Value
		End Set
	End Property
	
	
	Public Property TotalSize() As Integer
		Get
			TotalSize = myTotalSize
		End Get
		Set(ByVal Value As Integer)
			myTotalSize = Value
		End Set
	End Property
	
	
	Public Property EncounterCount() As Short
		Get
			EncounterCount = myEncounterCount
		End Get
		Set(ByVal Value As Short)
			myEncounterCount = Value
		End Set
	End Property
	
	
	Public Property Flags() As Byte
		Get
			Flags = myFlags
		End Get
		Set(ByVal Value As Byte)
			myFlags = Value
		End Set
	End Property
	
	
	Public Property MapName() As String
		Get
			MapName = Trim(myMapName)
		End Get
		Set(ByVal Value As String)
			myMapName = Trim(Value)
		End Set
	End Property
	
	
	Public Property MapComments() As String
		Get
			MapComments = Trim(myMapComments)
		End Get
		Set(ByVal Value As String)
			myMapComments = Trim(Value)
		End Set
	End Property
	
	
	Public Property FullPath() As String
		Get
			FullPath = Trim(myFullPath)
		End Get
		Set(ByVal Value As String)
			myFullPath = Trim(Value)
		End Set
	End Property
	
	
	Public Property IsSelected() As Short
		Get
			IsSelected = (myFlags And &H100) > 0
		End Get
		Set(ByVal Value As Short)
			If Value Then
				myFlags = myFlags Or &H100
			Else
				myFlags = myFlags And (Not &H100)
			End If
		End Set
	End Property
	
	
	Public Property IsUsed() As Short
		Get
			IsUsed = (myFlags And &H1) > 0
		End Get
		Set(ByVal Value As Short)
			If Value Then
				myFlags = myFlags Or &H1
			Else
				myFlags = myFlags And (Not &H1)
			End If
		End Set
	End Property
	
	
	Public Property GenerateUponEntry() As Short
		Get
			GenerateUponEntry = (myFlags And &H2) > 0
		End Get
		Set(ByVal Value As Short)
			If Value Then
				myFlags = myFlags Or &H2
			Else
				myFlags = myFlags And (Not &H2)
			End If
		End Set
	End Property
	
	
	Public Property GenerateThemes() As Short
		Get
			GenerateThemes = (myFlags And &H200) > 0
		End Get
		Set(ByVal Value As Short)
			If Value Then
				myFlags = myFlags Or &H200
			Else
				myFlags = myFlags And (Not &H200)
			End If
		End Set
	End Property
	
	
	Public Property NeedsThemes() As Short
		Get
			NeedsThemes = (myFlags And &H4) > 0
		End Get
		Set(ByVal Value As Short)
			If Value Then
				myFlags = myFlags Or &H4
			Else
				myFlags = myFlags And (Not &H4)
			End If
		End Set
	End Property
	
	
	Public Property IsMainMap() As Short
		Get
			IsMainMap = (myFlags And &H8) > 0
		End Get
		Set(ByVal Value As Short)
			If Value Then
				myFlags = myFlags Or &H8
			Else
				myFlags = myFlags And (Not &H8)
			End If
		End Set
	End Property
	
	
	Public Property MapStyle() As Short
		Get
			MapStyle = CShort(CShort(myFlags And &HF0) / 16)
		End Get
		Set(ByVal Value As Short)
			myFlags = myFlags And &HFF0F
			myFlags = myFlags Or (Value * 16 And &HF0)
		End Set
	End Property
	
	
	Public Property PartySize() As Short
		Get
			PartySize = CShort(myStyle And &H3)
		End Get
		Set(ByVal Value As Short)
			Dim ThemeSketchX As ThemeSketch
			myStyle = myStyle And &HFFFC
			myStyle = myStyle Or (Value And &H3)
		End Set
	End Property
	
	
	Public Property PartyAvgLevel() As Short
		Get
			PartyAvgLevel = CShort(CShort(myStyle And &HC) / 4)
		End Get
		Set(ByVal Value As Short)
			Dim ThemeSketchX As ThemeSketch
			myStyle = myStyle And &HFFF3
			myStyle = myStyle Or (Value * 4 And &HC)
		End Set
	End Property
	
	Public ReadOnly Property Difficulty() As Short
		Get
			Dim c As Short
			Select Case Me.PartyAvgLevel
				Case 0 ' 1-3
					c = 1
				Case 1 ' 4-6
					c = 3
				Case 2 ' 7-10
					c = 6
				Case 3 ' 10+
					c = 12
			End Select
			Select Case Me.PartySize
				Case 0 ' Solo
					Difficulty = 1 * c
				Case 1 ' 2-3
					Difficulty = 3 * c
				Case 2 ' 4-6
					Difficulty = 5 * c
				Case 3 ' 7-9
					Difficulty = 10 * c
			End Select
		End Get
	End Property
	
	
	Public Property AreaIndex() As Short
		Get
			AreaIndex = myAreaIndex
		End Get
		Set(ByVal Value As Short)
			myAreaIndex = Value
		End Set
	End Property
	
	
	Public Property MapIndex() As Short
		Get
			MapIndex = myMapIndex
		End Get
		Set(ByVal Value As Short)
			myMapIndex = Value
		End Set
	End Property
	
	
	Public Property TomeIndexInArea() As Short
		Get
			TomeIndexInArea = myTomeIndexInArea
		End Get
		Set(ByVal Value As Short)
			myTomeIndexInArea = Value
		End Set
	End Property
	
	
	Public Property AreaIndexInArea() As Short
		Get
			AreaIndexInArea = myAreaIndexInArea
		End Get
		Set(ByVal Value As Short)
			myAreaIndexInArea = Value
		End Set
	End Property
	
	
	Public Property MapIndexInArea() As Short
		Get
			MapIndexInArea = myMapIndexInArea
		End Get
		Set(ByVal Value As Short)
			myMapIndexInArea = Value
		End Set
	End Property
	
	
	Public Property EntryPoints() As Collection
		Get
			EntryPoints = myEntryPoints
		End Get
		Set(ByVal Value As Collection)
			myEntryPoints = Value
		End Set
	End Property
	
	Public ReadOnly Property ThemeQuests() As Collection
		Get
			ThemeQuests = myThemeQuests
		End Get
	End Property
	
	
	Public Property ThemeSketchs() As Collection
		Get
			ThemeSketchs = myThemeSketchs
		End Get
		Set(ByVal Value As Collection)
			myThemeSketchs = Value
		End Set
	End Property
	
	Public Sub IncreaseDifficulty()
		' Max is 3 and 3 in both
		If Me.PartySize = 3 And Me.PartyAvgLevel = 3 Then
			Exit Sub
		End If
		' Increase PartySize and AvgLevel by 1 degree
		Select Case Me.PartySize * 4 + Me.PartyAvgLevel
			Case 3, 7, 11
				Me.PartySize = Me.PartySize + 1
				Me.PartyAvgLevel = 0
			Case Else
				Me.PartyAvgLevel = Me.PartyAvgLevel + 1
		End Select
	End Sub
	
	Public Function AddEntryPoint() As EntryPoint
		' Adds an EntryPoint and returns a reference to that EntryPoint
		Dim c As Short
		Dim EntryPointX As EntryPoint
		' Find new available unused index idenifier
		c = 1
		For	Each EntryPointX In myEntryPoints
			If EntryPointX.Index >= c Then
				c = EntryPointX.Index + 1
			End If
		Next EntryPointX
		' Set the index and add the door. Return the new door's index.
		EntryPointX = New EntryPoint
		EntryPointX.Index = c
		EntryPointX.Name = "EntryPoint" & c
		myEntryPoints.Add(EntryPointX, "P" & EntryPointX.Index)
		AddEntryPoint = EntryPointX
	End Function
	
	Public Sub RemoveEntryPoint(ByRef DeleteKey As String)
		' Remove EntryPoint
		myEntryPoints.Remove(DeleteKey)
	End Sub
	
	Public Function AddThemeQuest() As Theme
		' Adds an Theme and returns a reference to that Theme
		Dim c As Short
		Dim ThemeX As Theme
		' Find new available unused index idenifier
		c = 1
		For	Each ThemeX In myThemeQuests
			If ThemeX.Index <> c Then
				Exit For
			End If
			c = c + 1
		Next ThemeX
		' Set the index and add the Theme. Return the new Theme's index.
		ThemeX = New Theme
		ThemeX.Index = c
		ThemeX.Name = "Theme" & c
		myThemeQuests.Add(ThemeX, "R" & ThemeX.Index)
		AddThemeQuest = ThemeX
	End Function
	
	Public Sub RemoveThemeQuest(ByRef DeleteKey As String)
		myThemeQuests.Remove(DeleteKey)
	End Sub
	
	Public Function AddThemeSketch() As ThemeSketch
		' Adds an ThemeSketch and returns a reference to that ThemeSketch
		Dim c As Short
		Dim ThemeSketchX As ThemeSketch
		' Find new available unused index idenifier
		c = 1
		For	Each ThemeSketchX In myThemeSketchs
			If ThemeSketchX.Index >= c Then
				c = ThemeSketchX.Index + 1
			End If
		Next ThemeSketchX
		' Set the index and add the door. Return the new door's index.
		ThemeSketchX = New ThemeSketch
		ThemeSketchX.Index = c
		ThemeSketchX.Name = "ThemeSketch" & c
		myThemeSketchs.Add(ThemeSketchX, "R" & ThemeSketchX.Index)
		AddThemeSketch = ThemeSketchX
	End Function
	
	Public Sub RemoveThemeSketch(ByRef DeleteKey As String)
		' Remove ThemeSketch
		myThemeSketchs.Remove(DeleteKey)
	End Sub
	
	Public Sub Copy(ByRef FromMapSketch As MapSketch)
		Dim EntryPointX As EntryPoint
		Dim c As Short
		myVersion = FromMapSketch.Version
		myMapName = FromMapSketch.MapName
		myMapComments = FromMapSketch.MapComments
		myFullPath = FromMapSketch.FullPath
		myFlags = FromMapSketch.Flags
		myTomeIndex = FromMapSketch.TomeIndex
		myAreaIndex = FromMapSketch.AreaIndex
		myMapIndex = FromMapSketch.MapIndex
		' Copy Entry Points
		For c = 1 To FromMapSketch.EntryPoints.Count()
			EntryPointX = Me.AddEntryPoint
			'UPGRADE_WARNING: Couldn't resolve default property of object FromMapSketch.EntryPoints(). Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
			EntryPointX.Copy(FromMapSketch.EntryPoints.Item(c))
		Next c
	End Sub
	
	Public Sub ReadFromFile(ByRef FromFile As Short)
		On Error GoTo ErrorHandler
		Dim c As Short
		Dim Cnt As Integer
		Dim EntryPointX As EntryPoint
		'UPGRADE_WARNING: Get was upgraded to FileGet and has a new behavior. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="9B7D5ADD-D8FE-4819-A36C-6DEDAF088CC7"'
		FileGet(FromFile, myVersion)
		'UPGRADE_WARNING: Get was upgraded to FileGet and has a new behavior. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="9B7D5ADD-D8FE-4819-A36C-6DEDAF088CC7"'
		FileGet(FromFile, myIndex)
		'UPGRADE_WARNING: Get was upgraded to FileGet and has a new behavior. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="9B7D5ADD-D8FE-4819-A36C-6DEDAF088CC7"'
		FileGet(FromFile, Cnt)
		myMapName = ""
		For c = 1 To Cnt
			myMapName = myMapName & " "
		Next c
		'UPGRADE_WARNING: Get was upgraded to FileGet and has a new behavior. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="9B7D5ADD-D8FE-4819-A36C-6DEDAF088CC7"'
		FileGet(FromFile, myMapName)
		'UPGRADE_WARNING: Get was upgraded to FileGet and has a new behavior. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="9B7D5ADD-D8FE-4819-A36C-6DEDAF088CC7"'
		FileGet(FromFile, Cnt)
		myMapComments = ""
		For c = 1 To Cnt
			myMapComments = myMapComments & " "
		Next c
		'UPGRADE_WARNING: Get was upgraded to FileGet and has a new behavior. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="9B7D5ADD-D8FE-4819-A36C-6DEDAF088CC7"'
		FileGet(FromFile, myMapComments)
		'UPGRADE_WARNING: Get was upgraded to FileGet and has a new behavior. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="9B7D5ADD-D8FE-4819-A36C-6DEDAF088CC7"'
		FileGet(FromFile, Cnt)
		myFullPath = ""
		For c = 1 To Cnt
			myFullPath = myFullPath & " "
		Next c
		'UPGRADE_WARNING: Get was upgraded to FileGet and has a new behavior. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="9B7D5ADD-D8FE-4819-A36C-6DEDAF088CC7"'
		FileGet(FromFile, myFullPath)
		'UPGRADE_WARNING: Get was upgraded to FileGet and has a new behavior. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="9B7D5ADD-D8FE-4819-A36C-6DEDAF088CC7"'
		FileGet(FromFile, myFlags)
		'UPGRADE_WARNING: Get was upgraded to FileGet and has a new behavior. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="9B7D5ADD-D8FE-4819-A36C-6DEDAF088CC7"'
		FileGet(FromFile, myTomeIndex)
		'UPGRADE_WARNING: Get was upgraded to FileGet and has a new behavior. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="9B7D5ADD-D8FE-4819-A36C-6DEDAF088CC7"'
		FileGet(FromFile, myAreaIndex)
		'UPGRADE_WARNING: Get was upgraded to FileGet and has a new behavior. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="9B7D5ADD-D8FE-4819-A36C-6DEDAF088CC7"'
		FileGet(FromFile, myMapIndex)
		' Read EntryPoints
		'UPGRADE_WARNING: Get was upgraded to FileGet and has a new behavior. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="9B7D5ADD-D8FE-4819-A36C-6DEDAF088CC7"'
		FileGet(FromFile, Cnt)
		For c = 1 To Cnt
			EntryPointX = New EntryPoint
			EntryPointX.ReadFromFile(FromFile)
			myEntryPoints.Add(EntryPointX, "P" & EntryPointX.Index)
		Next c
		Exit Sub
ErrorHandler: 
		' [Titi 2.4.9]
		If Err.Number = 457 Then
			oErr.logError("MapSketch index conflict: " & EntryPointX.Index)
		Else
			oErr.logError("Cannot read map sketch, error#" & Err.Number & " (" & Err.Description & ")")
		End If
	End Sub
	
	Public Sub SaveToFile(ByRef ToFile As Short)
		Dim c As Short
		'UPGRADE_WARNING: Put was upgraded to FilePut and has a new behavior. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="9B7D5ADD-D8FE-4819-A36C-6DEDAF088CC7"'
		FilePut(ToFile, myVersion)
		'UPGRADE_WARNING: Put was upgraded to FilePut and has a new behavior. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="9B7D5ADD-D8FE-4819-A36C-6DEDAF088CC7"'
		FilePut(ToFile, myIndex)
		'UPGRADE_WARNING: Put was upgraded to FilePut and has a new behavior. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="9B7D5ADD-D8FE-4819-A36C-6DEDAF088CC7"'
		FilePut(ToFile, Len(myMapName))
		'UPGRADE_WARNING: Put was upgraded to FilePut and has a new behavior. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="9B7D5ADD-D8FE-4819-A36C-6DEDAF088CC7"'
		FilePut(ToFile, myMapName)
		'UPGRADE_WARNING: Put was upgraded to FilePut and has a new behavior. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="9B7D5ADD-D8FE-4819-A36C-6DEDAF088CC7"'
		FilePut(ToFile, Len(myMapComments))
		'UPGRADE_WARNING: Put was upgraded to FilePut and has a new behavior. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="9B7D5ADD-D8FE-4819-A36C-6DEDAF088CC7"'
		FilePut(ToFile, myMapComments)
		'UPGRADE_WARNING: Put was upgraded to FilePut and has a new behavior. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="9B7D5ADD-D8FE-4819-A36C-6DEDAF088CC7"'
		FilePut(ToFile, Len(myFullPath))
		'UPGRADE_WARNING: Put was upgraded to FilePut and has a new behavior. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="9B7D5ADD-D8FE-4819-A36C-6DEDAF088CC7"'
		FilePut(ToFile, myFullPath)
		'UPGRADE_WARNING: Put was upgraded to FilePut and has a new behavior. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="9B7D5ADD-D8FE-4819-A36C-6DEDAF088CC7"'
		FilePut(ToFile, myFlags)
		'UPGRADE_WARNING: Put was upgraded to FilePut and has a new behavior. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="9B7D5ADD-D8FE-4819-A36C-6DEDAF088CC7"'
		FilePut(ToFile, myTomeIndex)
		'UPGRADE_WARNING: Put was upgraded to FilePut and has a new behavior. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="9B7D5ADD-D8FE-4819-A36C-6DEDAF088CC7"'
		FilePut(ToFile, myAreaIndex)
		'UPGRADE_WARNING: Put was upgraded to FilePut and has a new behavior. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="9B7D5ADD-D8FE-4819-A36C-6DEDAF088CC7"'
		FilePut(ToFile, myMapIndex)
		' Save EntryPoints
		'UPGRADE_WARNING: Put was upgraded to FilePut and has a new behavior. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="9B7D5ADD-D8FE-4819-A36C-6DEDAF088CC7"'
		FilePut(ToFile, myEntryPoints.Count())
		For c = 1 To myEntryPoints.Count()
			'UPGRADE_WARNING: Couldn't resolve default property of object myEntryPoints().SaveToFile. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
			myEntryPoints.Item(c).SaveToFile(ToFile)
		Next c
	End Sub
End Class