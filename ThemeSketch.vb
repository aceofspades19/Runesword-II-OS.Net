Option Strict Off
Option Explicit On
Friend Class ThemeSketch
	
	' Version number of Class
	Dim myVersion As Short
	' Unique number for this Theme
	Dim myIndex As Short
	' Name of Theme
	Dim myName As String
	' Comments for Theme
	Dim myComments As String
	' Created for Party Size
	' Bit1-2 - 0 = Solo, 1 = 2-3, 2 = 4-6, 3 = 7-9
	' With Average Party Level
	' Bit3-4 - 0 = 1-3, 1 = 4-6, 2 = 7-10, 3 = 10+
	' Bit5-8 (MapStyle)
	'       0 - Town Map
	'       1 - Wilderness Map
	'       2 - Building Map
	'       3 - Dungeon Map
	' Bit9-16 - Not Used
	Dim myStyle As Short
	' Bit1      - IsQuest? (Can span multiple maps) On=Yes, Off=No
	' Bit2      - On=IsUsed Off=Not
	' Bit3      - On=IsSelected, Off=Not
	' Bit4      - IsMajorTheme? On=Yes, Off=No
	' Bit5-7    - Not Used
	Dim myFlags As Short
	' Amount of Encounters this should create
	Dim myCoverage As Short
	' Path where File is located
	Dim myFullPath As String
	
	
	Public Property Version() As Short
		Get
			Version = myVersion
		End Get
		Set(ByVal Value As Short)
			myVersion = Value
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
	
	
	Public Property Name() As String
		Get
			Name = Trim(myName)
		End Get
		Set(ByVal Value As String)
			myName = Trim(Value)
		End Set
	End Property
	
	
	Public Property Comments() As String
		Get
			Comments = Trim(myComments)
		End Get
		Set(ByVal Value As String)
			myComments = Trim(Value)
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
	
	
	Public Property Flags() As Short
		Get
			Flags = myFlags
		End Get
		Set(ByVal Value As Short)
			myFlags = Value
		End Set
	End Property
	
	
	Public Property IsQuest() As Short
		Get
			IsQuest = (myFlags And &H1) > 0
		End Get
		Set(ByVal Value As Short)
			If Value Then
				myFlags = myFlags Or &H1
			Else
				myFlags = myFlags And (Not &H1)
			End If
		End Set
	End Property
	
	
	Public Property IsMajorTheme() As Short
		Get
			IsMajorTheme = (myFlags And &H8) > 0
		End Get
		Set(ByVal Value As Short)
			If Value Then
				myFlags = myFlags Or &H8
			Else
				myFlags = myFlags And (Not &H8)
			End If
		End Set
	End Property
	
	
	Public Property IsUsed() As Short
		Get
			IsUsed = (myFlags And &H2) > 0
		End Get
		Set(ByVal Value As Short)
			If Value Then
				myFlags = myFlags Or &H2
			Else
				myFlags = myFlags And (Not &H2)
			End If
		End Set
	End Property
	
	
	Public Property IsSelected() As Short
		Get
			IsSelected = (myFlags And &H4) > 0
		End Get
		Set(ByVal Value As Short)
			If Value Then
				myFlags = myFlags Or &H4
			Else
				myFlags = myFlags And (Not &H4)
			End If
		End Set
	End Property
	
	
	Public Property PartySize() As Short
		Get
			PartySize = CShort(myStyle And &H3)
		End Get
		Set(ByVal Value As Short)
			myStyle = myStyle And &HFFFC
			myStyle = myStyle Or (Value And &H3)
		End Set
	End Property
	
	
	Public Property PartyAvgLevel() As Short
		Get
			PartyAvgLevel = CShort(CShort(myStyle And &HC) / 4)
		End Get
		Set(ByVal Value As Short)
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
	
	
	Public Property Coverage() As Short
		Get
			Coverage = myCoverage
		End Get
		Set(ByVal Value As Short)
			myCoverage = Value
		End Set
	End Property
	
	
	Public Property Style() As Short
		Get
			Style = myStyle
		End Get
		Set(ByVal Value As Short)
			myStyle = Value
		End Set
	End Property
	
	
	Public Property MapStyle() As Short
		Get
			MapStyle = CShort(CShort(myStyle And &HF0) / 16)
		End Get
		Set(ByVal Value As Short)
			myStyle = myStyle And &HFF0F
			myStyle = myStyle Or (Value * 16 And &HF0)
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
	
	Public Sub Copy(ByRef ThemeSketchX As ThemeSketch)
		myVersion = ThemeSketchX.Version
		myName = ThemeSketchX.Name
		myComments = ThemeSketchX.Comments
		myStyle = ThemeSketchX.Style
		myFlags = ThemeSketchX.Flags
		myFullPath = ThemeSketchX.FullPath
	End Sub
End Class