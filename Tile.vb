Option Strict Off
Option Explicit On
Friend Class Tile
	
	' Version number of Tile
	Private myVersion As Short
	' Name of the Tile
	Private myName As String
	' Picture Reference
	Private myPic As Short
	' Unique number for Tile
	Private myIndex As Short
	'   &H1  Bit 0  - On = Block Move Upper Left
	'   &H2  Bit 1  - On = Block Move Upper Right
	'   &H4  Bit 2  - On = Block Move Lower Left
	'   &H8  Bit 3  - On = Block Move Lower Right
	'   &H10 Bit 4  - On = Can't See Past Upper Left
	'   &H20 Bit 5  - On = Can't See Past Upper Right
	'   &H40 Bit 6  - On = Can't See Past Lower Left
	'   &H80 Bit 7  - On = Can't See Past Lower Right
	Private myBlocked As Byte
	' Used in Random Dungeon Maker
	' 0 - <None>, 1 - Left Wall, 2 - Top Wall, 3 - Floor
	' 4 - Left Arch, 5 - Left Door
	Private myStyle As Byte
	' Movement Cost per Tile
	Private myMovementCost As Byte
	' Terrain Byte Index: (to be defined)
	Private myTerrainType As Byte
	'   Bit 0   - On=Flip, Off=No Flip
	'   Bit 1-2 - 00 = Light, 01 = Dark, 10 = Very Dark, 11 = Pitch Black
	'   Bit 3-7 - Key bits; all 0's not locked. All 1's, no key will work.
	Private myKey As Byte
	' Number from 1-255 indicating how likely one click of the mouse will
	' cause the tile to switch pictures. This is only applicable if the
	' key is not set.
	Private myChance As Byte
	' Pointer to index of Tile Object in Map.Tiles Collection to Swap to.
	' 0 is no tile to swap to.
	Private mySwapTile As Short
	' Index of TileSet in parent Map Object. Groups tiles under a single
	' name and relates to a specific Combat Wallpaper BMP.
	Private myTileSet As Short
	' Not Used
	Private myNotUsed As Integer
	
	
	Public Property Index() As Short
		Get
			Index = myIndex
		End Get
		Set(ByVal Value As Short)
			myIndex = Value
		End Set
	End Property
	
	
	Public Property Pic() As Short
		Get
			Pic = myPic
		End Get
		Set(ByVal Value As Short)
			myPic = Value
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
	
	
	Public Property Chance() As Short
		Get
			Chance = myChance
		End Get
		Set(ByVal Value As Short)
			myChance = Value
		End Set
	End Property
	
	
	Public Property MovementCost() As Short
		Get
			MovementCost = myMovementCost
		End Get
		Set(ByVal Value As Short)
			myMovementCost = Value
		End Set
	End Property
	
	Public ReadOnly Property MovementCostTurns() As Short
		Get
			Select Case myMovementCost
				Case 0 To 9
					MovementCostTurns = myMovementCost + 1
				Case 10 To 13
					MovementCostTurns = (myMovementCost - 8) * 10
				Case 14 To 18
					MovementCostTurns = (myMovementCost - 13) * 100
				Case 19
					MovementCostTurns = 1000
				Case 20 To 24
					MovementCostTurns = (myMovementCost - 19) * 200
			End Select
		End Get
	End Property
	
	
	Public Property TerrainType() As Short
		Get
			TerrainType = myTerrainType
		End Get
		Set(ByVal Value As Short)
			myTerrainType = Value
		End Set
	End Property
	
	
	Public Property SwapTile() As Short
		Get
			SwapTile = mySwapTile
		End Get
		Set(ByVal Value As Short)
			mySwapTile = Value
		End Set
	End Property
	
	
	Public Property Key() As Short
		Get
			Key = myKey
		End Get
		Set(ByVal Value As Short)
			myKey = Value
		End Set
	End Property
	
	
	Public Property KeyBit(ByVal Index As Short) As Short
		Get
			KeyBit = (myKey And (2 ^ Index)) > 0
		End Get
		Set(ByVal Value As Short)
			If Value Then
				myKey = myKey Or (2 ^ Index)
			Else
				myKey = myKey And (Not (2 ^ Index))
			End If
		End Set
	End Property
	
	
	Public Property KeyBits() As Byte
		Get
			KeyBits = (myKey And &HF8)
		End Get
		Set(ByVal Value As Byte)
			' Wipe out the old key
			KeyBits = (myKey And &H7)
			' Set the new key
			KeyBits = (myKey Or Value)
		End Set
	End Property
	
	
	Public Property Style() As Short
		Get
			Style = CShort(myStyle)
		End Get
		Set(ByVal Value As Short)
			myStyle = CByte(Value)
			' Set other properties if you changed the Sytle
			Select Case myStyle
				Case 0 ' None
				Case 1 ' Left Wall
					Me.Blocked(0) = True
					Me.Blocked(1) = False
					Me.Blocked(2) = False
					Me.Blocked(3) = False
					Me.See(0) = True
					Me.See(1) = False
					Me.See(2) = False
					Me.See(3) = False
				Case 2 ' Top Wall
					Me.Blocked(0) = True
					Me.Blocked(1) = True
					Me.Blocked(2) = False
					Me.Blocked(3) = False
					Me.See(0) = True
					Me.See(1) = True
					Me.See(2) = False
					Me.See(3) = False
				Case 3 ' Floor
					myBlocked = 0
				Case 4 ' Left Arch
					myBlocked = 0
				Case 5 ' Left Door
					Me.Blocked(0) = True
					Me.Blocked(1) = False
					Me.Blocked(2) = False
					Me.Blocked(3) = False
					Me.See(0) = True
					Me.See(1) = False
					Me.See(2) = False
					Me.See(3) = False
					myChance = 100
				Case 10 ' Corner Arch
					Me.Blocked(0) = False
					Me.Blocked(1) = True
					Me.Blocked(2) = False
					Me.Blocked(3) = False
					Me.See(0) = False
					Me.See(1) = True
					Me.See(2) = False
					Me.See(3) = False
				Case 11 ' Block
					Me.Blocked(0) = True
					Me.Blocked(1) = True
					Me.Blocked(2) = True
					Me.Blocked(3) = True
					Me.See(0) = True
					Me.See(1) = True
					Me.See(2) = True
					Me.See(3) = True
			End Select
		End Set
	End Property
	
	
	Public Property See(ByVal Index As Short) As Short
		Get
			Select Case Index
				Case 0 ' Upper Left
					See = (myBlocked And 16) > 0
				Case 1 ' Upper Right
					See = (myBlocked And 32) > 0
				Case 2 ' Lower Left
					See = (myBlocked And 64) > 0
				Case 3 ' Lower Right
					See = (myBlocked And 128) > 0
				Case 4 ' Straight Up (Upper Left and Upper Right)
					See = (myBlocked And 16) > 0 And (myBlocked And 32) > 0
				Case 5 ' Straight Left (Upper Left and Lower Left)
					See = (myBlocked And 16) > 0 And (myBlocked And 64) > 0
				Case 6 ' Straight Right (Upper Right and Lower Right)
					See = (myBlocked And 32) > 0 And (myBlocked And 128) > 0
				Case 7 ' Straight Down (Lower Left and Lower Right)
					See = (myBlocked And 64) > 0 And (myBlocked And 128) > 0
			End Select
		End Get
		Set(ByVal Value As Short)
			If Value Then
				myBlocked = myBlocked Or 2 ^ (Index + 4)
			Else
				myBlocked = myBlocked And (Not 2 ^ (Index + 4))
			End If
		End Set
	End Property
	
	
	Public Property Blocked(ByVal Index As Short) As Short
		Get
			Select Case Index
				Case 0 ' Upper Left
					Blocked = (myBlocked And &H1) > 0
				Case 1 ' Upper Right
					Blocked = (myBlocked And &H2) > 0
				Case 2 ' Lower Left
					Blocked = (myBlocked And &H4) > 0
				Case 3 ' Lower Right
					Blocked = (myBlocked And &H8) > 0
				Case 4 ' Straight Up (Upper Left and Upper Right)
					Blocked = (myBlocked And &H1) > 0 And (myBlocked And &H2) > 0
				Case 5 ' Straight Left (Upper Left and Lower Left)
					Blocked = (myBlocked And &H1) > 0 And (myBlocked And &H4) > 0
				Case 6 ' Straight Right (Upper Right and Lower Right)
					Blocked = (myBlocked And &H2) > 0 And (myBlocked And &H8) > 0
				Case 7 ' Straight Down (Lower Left and Lower Right)
					Blocked = (myBlocked And &H4) > 0 And (myBlocked And &H8) > 0
			End Select
		End Get
		Set(ByVal Value As Short)
			If Value Then
				myBlocked = myBlocked Or 2 ^ Index
			Else
				myBlocked = myBlocked And (Not 2 ^ Index)
			End If
		End Set
	End Property
	
	
	Public Property BlockedAll() As Byte
		Get
			BlockedAll = myBlocked
		End Get
		Set(ByVal Value As Byte)
			myBlocked = Value
		End Set
	End Property
	
	
	Public Property Flip() As Short
		Get
			Flip = (myKey And &H1) > 0
		End Get
		Set(ByVal Value As Short)
			If Value Then
				myKey = myKey Or &H1
			Else
				myKey = myKey And (Not &H1)
			End If
		End Set
	End Property
	
	
	Public Property TileSet() As Short
		Get
			TileSet = myTileSet
		End Get
		Set(ByVal Value As Short)
			myTileSet = Value
		End Set
	End Property
	
	Public Sub Copy(ByRef FromTile As Tile)
        myName = FromTile.Name
        myPic = FromTile.Pic
        myBlocked = FromTile.BlockedAll
        myStyle = FromTile.Style
        myKey = FromTile.Key
        Me.Flip = FromTile.Flip
        myChance = FromTile.Chance
        mySwapTile = FromTile.SwapTile
        myTileSet = FromTile.TileSet
    End Sub
	
	Public Sub SaveToFile(ByRef ToFile As Short)
        'UPGRADE_WARNING: Put was upgraded to FilePut and has a new behavior. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="9B7D5ADD-D8FE-4819-A36C-6DEDAF088CC7"'
        FilePut(ToFile, myVersion)
        'UPGRADE_WARNING: Put was upgraded to FilePut and has a new behavior. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="9B7D5ADD-D8FE-4819-A36C-6DEDAF088CC7"'
        FilePut(ToFile, Len(myName))
        'UPGRADE_WARNING: Put was upgraded to FilePut and has a new behavior. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="9B7D5ADD-D8FE-4819-A36C-6DEDAF088CC7"'
        FilePut(ToFile, myName)
        'UPGRADE_WARNING: Put was upgraded to FilePut and has a new behavior. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="9B7D5ADD-D8FE-4819-A36C-6DEDAF088CC7"'
        FilePut(ToFile, myIndex)
        'UPGRADE_WARNING: Put was upgraded to FilePut and has a new behavior. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="9B7D5ADD-D8FE-4819-A36C-6DEDAF088CC7"'
        FilePut(ToFile, myPic)
        'UPGRADE_WARNING: Put was upgraded to FilePut and has a new behavior. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="9B7D5ADD-D8FE-4819-A36C-6DEDAF088CC7"'
        FilePut(ToFile, myBlocked)
        'UPGRADE_WARNING: Put was upgraded to FilePut and has a new behavior. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="9B7D5ADD-D8FE-4819-A36C-6DEDAF088CC7"'
        FilePut(ToFile, myStyle)
        'UPGRADE_WARNING: Put was upgraded to FilePut and has a new behavior. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="9B7D5ADD-D8FE-4819-A36C-6DEDAF088CC7"'
        FilePut(ToFile, myKey)
        'UPGRADE_WARNING: Put was upgraded to FilePut and has a new behavior. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="9B7D5ADD-D8FE-4819-A36C-6DEDAF088CC7"'
        FilePut(ToFile, myChance)
        'UPGRADE_WARNING: Put was upgraded to FilePut and has a new behavior. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="9B7D5ADD-D8FE-4819-A36C-6DEDAF088CC7"'
        FilePut(ToFile, mySwapTile)
        'UPGRADE_WARNING: Put was upgraded to FilePut and has a new behavior. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="9B7D5ADD-D8FE-4819-A36C-6DEDAF088CC7"'
        FilePut(ToFile, myTileSet)
        'UPGRADE_WARNING: Put was upgraded to FilePut and has a new behavior. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="9B7D5ADD-D8FE-4819-A36C-6DEDAF088CC7"'
        FilePut(ToFile, myMovementCost)
        'UPGRADE_WARNING: Put was upgraded to FilePut and has a new behavior. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="9B7D5ADD-D8FE-4819-A36C-6DEDAF088CC7"'
        FilePut(ToFile, myTerrainType)
        'UPGRADE_WARNING: Put was upgraded to FilePut and has a new behavior. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="9B7D5ADD-D8FE-4819-A36C-6DEDAF088CC7"'
        FilePut(ToFile, myNotUsed)
    End Sub
	
	Public Sub ReadFromFile(ByRef FromFile As Short)
		Dim Cnt As Integer
		Dim c As Short
		' Read Name
		'UPGRADE_WARNING: Get was upgraded to FileGet and has a new behavior. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="9B7D5ADD-D8FE-4819-A36C-6DEDAF088CC7"'
		FileGet(FromFile, myVersion)
		'UPGRADE_WARNING: Get was upgraded to FileGet and has a new behavior. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="9B7D5ADD-D8FE-4819-A36C-6DEDAF088CC7"'
		FileGet(FromFile, Cnt)
		myName = ""
		For c = 1 To Cnt
			myName = myName & " "
		Next c
		'UPGRADE_WARNING: Get was upgraded to FileGet and has a new behavior. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="9B7D5ADD-D8FE-4819-A36C-6DEDAF088CC7"'
		FileGet(FromFile, myName)
		' Read Index and others
		'UPGRADE_WARNING: Get was upgraded to FileGet and has a new behavior. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="9B7D5ADD-D8FE-4819-A36C-6DEDAF088CC7"'
		FileGet(FromFile, myIndex)
		'UPGRADE_WARNING: Get was upgraded to FileGet and has a new behavior. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="9B7D5ADD-D8FE-4819-A36C-6DEDAF088CC7"'
		FileGet(FromFile, myPic)
		'UPGRADE_WARNING: Get was upgraded to FileGet and has a new behavior. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="9B7D5ADD-D8FE-4819-A36C-6DEDAF088CC7"'
		FileGet(FromFile, myBlocked)
		'UPGRADE_WARNING: Get was upgraded to FileGet and has a new behavior. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="9B7D5ADD-D8FE-4819-A36C-6DEDAF088CC7"'
		FileGet(FromFile, myStyle)
		'UPGRADE_WARNING: Get was upgraded to FileGet and has a new behavior. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="9B7D5ADD-D8FE-4819-A36C-6DEDAF088CC7"'
		FileGet(FromFile, myKey)
		'UPGRADE_WARNING: Get was upgraded to FileGet and has a new behavior. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="9B7D5ADD-D8FE-4819-A36C-6DEDAF088CC7"'
		FileGet(FromFile, myChance)
		'UPGRADE_WARNING: Get was upgraded to FileGet and has a new behavior. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="9B7D5ADD-D8FE-4819-A36C-6DEDAF088CC7"'
		FileGet(FromFile, mySwapTile)
		'UPGRADE_WARNING: Get was upgraded to FileGet and has a new behavior. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="9B7D5ADD-D8FE-4819-A36C-6DEDAF088CC7"'
		FileGet(FromFile, myTileSet)
		'UPGRADE_WARNING: Get was upgraded to FileGet and has a new behavior. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="9B7D5ADD-D8FE-4819-A36C-6DEDAF088CC7"'
		FileGet(FromFile, myMovementCost)
		'UPGRADE_WARNING: Get was upgraded to FileGet and has a new behavior. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="9B7D5ADD-D8FE-4819-A36C-6DEDAF088CC7"'
		FileGet(FromFile, myTerrainType)
		'UPGRADE_WARNING: Get was upgraded to FileGet and has a new behavior. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="9B7D5ADD-D8FE-4819-A36C-6DEDAF088CC7"'
		FileGet(FromFile, myNotUsed)
	End Sub
	
	'UPGRADE_NOTE: Class_Initialize was upgraded to Class_Initialize_Renamed. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="A9E4979A-37FA-4718-9994-97DD76ED70A7"'
	Private Sub Class_Initialize_Renamed()
		myName = "Untitled"
	End Sub
	Public Sub New()
		MyBase.New()
		Class_Initialize_Renamed()
	End Sub
End Class