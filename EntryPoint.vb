Option Strict Off
Option Explicit On
Friend Class EntryPoint
	
	' Version number for Class
	Dim myVersion As Short
	' Unique number for this Entry Point
	Dim myIndex As Short
	' Name of Entry Point
	Dim myName As String
	' Description when enter/exit from this Entry Point
	Dim myDescription As String
	' X coordinate where entry point is located on this Map
	Dim myMapX As Short
	' Y coordinate where entry point is located on this Map
	Dim myMapY As Short
	' Bit1-4 (Style)
	'       0 - Exit
	'       1 - Exit Up
	'       2 - Exit Down
	'       3 - No Show Exit Sign
	' Bit5-8 (MapStyle)
	'       0 - To Town Map
	'       1 - To Wilderness Map
	'       2 - To Building Map
	'       3 - To Dungeon Map
	' Bit9  -   On=InUse with TomeWizard, Off=Not used
	' Bit10 -   On=No Show Exit Sign, Off=Show Exit Sign
	' Bit11-16   - Undefined
	Dim myFlags As Short
	
	' Index of Area (.rsa file) to jump to when hit this entry point. 0 is Main Menu. -1 is Random.
	Dim myAreaIndex As Short
	' Index of Map within above Area to jump to.
	Dim myMapIndex As Short
	' Index of EntryPoint within above Map.
	Dim myEntryIndex As Short
	
	
	Public Property Index() As Short
		Get
			Index = myIndex
		End Get
		Set(ByVal Value As Short)
			myIndex = Value
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
	
	
	Public Property Description() As String
		Get
			Description = Trim(myDescription)
		End Get
		Set(ByVal Value As String)
			myDescription = Trim(Value)
		End Set
	End Property
	
	
	Public Property MapX() As Short
		Get
			MapX = myMapX
		End Get
		Set(ByVal Value As Short)
			myMapX = Value
		End Set
	End Property
	
	
	Public Property MapY() As Short
		Get
			MapY = myMapY
		End Get
		Set(ByVal Value As Short)
			myMapY = Value
		End Set
	End Property
	
	
	Public Property Style() As Short
		Get
			Style = CShort(myFlags And &HF)
		End Get
		Set(ByVal Value As Short)
			myFlags = myFlags And &HFFF0
			myFlags = myFlags Or (Value And &HF)
		End Set
	End Property
	
	Public ReadOnly Property ToStyle() As Short
		Get
			Select Case Me.Style
				Case 0 ' Exit
					ToStyle = 0
				Case 1 ' Exit Up
					ToStyle = 2
				Case 2 ' Exit Down
					ToStyle = 1
			End Select
		End Get
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
	
	
	Public Property IsUsed() As Short
		Get
			IsUsed = (myFlags And &H100) > 0
		End Get
		Set(ByVal Value As Short)
			If Value Then
				myFlags = myFlags Or &H100
			Else
				myFlags = myFlags And (Not &H100)
			End If
		End Set
	End Property
	
	
	Public Property IsNoExitSign() As Short
		Get
			IsNoExitSign = (myFlags And &H200) > 0
		End Get
		Set(ByVal Value As Short)
			If Value Then
				myFlags = myFlags Or &H200
			Else
				myFlags = myFlags And (Not &H200)
			End If
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
	
	
	Public Property EntryIndex() As Short
		Get
			EntryIndex = myEntryIndex
		End Get
		Set(ByVal Value As Short)
			myEntryIndex = Value
		End Set
	End Property
	
	Public Sub Copy(ByRef FromEntryPoint As EntryPoint)
		myName = FromEntryPoint.Name
		myDescription = FromEntryPoint.Description
		myMapX = FromEntryPoint.MapX
		myMapY = FromEntryPoint.MapY
		myFlags = FromEntryPoint.Flags
		myAreaIndex = FromEntryPoint.AreaIndex
		myMapIndex = FromEntryPoint.MapIndex
		myEntryIndex = FromEntryPoint.EntryIndex
	End Sub
	
	Public Sub SaveToFile(ByRef ToFile As Short)
		'UPGRADE_WARNING: Put was upgraded to FilePut and has a new behavior. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="9B7D5ADD-D8FE-4819-A36C-6DEDAF088CC7"'
		FilePut(ToFile, myVersion)
		'UPGRADE_WARNING: Put was upgraded to FilePut and has a new behavior. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="9B7D5ADD-D8FE-4819-A36C-6DEDAF088CC7"'
		FilePut(ToFile, myIndex)
		'UPGRADE_WARNING: Put was upgraded to FilePut and has a new behavior. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="9B7D5ADD-D8FE-4819-A36C-6DEDAF088CC7"'
		FilePut(ToFile, Len(myName))
		'UPGRADE_WARNING: Put was upgraded to FilePut and has a new behavior. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="9B7D5ADD-D8FE-4819-A36C-6DEDAF088CC7"'
		FilePut(ToFile, myName)
		'UPGRADE_WARNING: Put was upgraded to FilePut and has a new behavior. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="9B7D5ADD-D8FE-4819-A36C-6DEDAF088CC7"'
		FilePut(ToFile, Len(myDescription))
		'UPGRADE_WARNING: Put was upgraded to FilePut and has a new behavior. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="9B7D5ADD-D8FE-4819-A36C-6DEDAF088CC7"'
		FilePut(ToFile, myDescription)
		'UPGRADE_WARNING: Put was upgraded to FilePut and has a new behavior. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="9B7D5ADD-D8FE-4819-A36C-6DEDAF088CC7"'
		FilePut(ToFile, myMapX)
		'UPGRADE_WARNING: Put was upgraded to FilePut and has a new behavior. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="9B7D5ADD-D8FE-4819-A36C-6DEDAF088CC7"'
		FilePut(ToFile, myMapY)
		'UPGRADE_WARNING: Put was upgraded to FilePut and has a new behavior. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="9B7D5ADD-D8FE-4819-A36C-6DEDAF088CC7"'
		FilePut(ToFile, myFlags)
		'UPGRADE_WARNING: Put was upgraded to FilePut and has a new behavior. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="9B7D5ADD-D8FE-4819-A36C-6DEDAF088CC7"'
		FilePut(ToFile, myAreaIndex)
		'UPGRADE_WARNING: Put was upgraded to FilePut and has a new behavior. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="9B7D5ADD-D8FE-4819-A36C-6DEDAF088CC7"'
		FilePut(ToFile, myMapIndex)
		'UPGRADE_WARNING: Put was upgraded to FilePut and has a new behavior. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="9B7D5ADD-D8FE-4819-A36C-6DEDAF088CC7"'
		FilePut(ToFile, myEntryIndex)
	End Sub
	
	Public Sub ReadFromFile(ByRef FromFile As Short)
		Dim Cnt As Integer
		Dim c As Short
		'UPGRADE_WARNING: Get was upgraded to FileGet and has a new behavior. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="9B7D5ADD-D8FE-4819-A36C-6DEDAF088CC7"'
		FileGet(FromFile, myVersion)
		'UPGRADE_WARNING: Get was upgraded to FileGet and has a new behavior. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="9B7D5ADD-D8FE-4819-A36C-6DEDAF088CC7"'
		FileGet(FromFile, myIndex)
		'UPGRADE_WARNING: Get was upgraded to FileGet and has a new behavior. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="9B7D5ADD-D8FE-4819-A36C-6DEDAF088CC7"'
		FileGet(FromFile, Cnt)
		myName = ""
		For c = 1 To Cnt
			myName = myName & " "
		Next c
		'UPGRADE_WARNING: Get was upgraded to FileGet and has a new behavior. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="9B7D5ADD-D8FE-4819-A36C-6DEDAF088CC7"'
		FileGet(FromFile, myName)
		'UPGRADE_WARNING: Get was upgraded to FileGet and has a new behavior. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="9B7D5ADD-D8FE-4819-A36C-6DEDAF088CC7"'
		FileGet(FromFile, Cnt)
		myDescription = ""
		For c = 1 To Cnt
			myDescription = myDescription & " "
		Next c
		'UPGRADE_WARNING: Get was upgraded to FileGet and has a new behavior. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="9B7D5ADD-D8FE-4819-A36C-6DEDAF088CC7"'
		FileGet(FromFile, myDescription)
		'UPGRADE_WARNING: Get was upgraded to FileGet and has a new behavior. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="9B7D5ADD-D8FE-4819-A36C-6DEDAF088CC7"'
		FileGet(FromFile, myMapX)
		'UPGRADE_WARNING: Get was upgraded to FileGet and has a new behavior. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="9B7D5ADD-D8FE-4819-A36C-6DEDAF088CC7"'
		FileGet(FromFile, myMapY)
		'UPGRADE_WARNING: Get was upgraded to FileGet and has a new behavior. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="9B7D5ADD-D8FE-4819-A36C-6DEDAF088CC7"'
		FileGet(FromFile, myFlags)
		'UPGRADE_WARNING: Get was upgraded to FileGet and has a new behavior. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="9B7D5ADD-D8FE-4819-A36C-6DEDAF088CC7"'
		FileGet(FromFile, myAreaIndex)
		'UPGRADE_WARNING: Get was upgraded to FileGet and has a new behavior. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="9B7D5ADD-D8FE-4819-A36C-6DEDAF088CC7"'
		FileGet(FromFile, myMapIndex)
		'UPGRADE_WARNING: Get was upgraded to FileGet and has a new behavior. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="9B7D5ADD-D8FE-4819-A36C-6DEDAF088CC7"'
		FileGet(FromFile, myEntryIndex)
	End Sub
End Class