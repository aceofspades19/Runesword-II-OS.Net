Option Strict Off
Option Explicit On
Friend Class Tome
	' This file is used to identify a group of Plots. It also serves as
	' the Party holder --- holding a collection of Creatures (CreatureWithTurns, Story
	' CreatureWithTurns and NCreatureWithTurns) which make up the Party.
	
	' Version number of this Class
	Private myVersion As Short
	' Index used in TomeList (not saved)
	Private myIndex As Short
	' The Name as it will appear in the Bookmarks.
	Private myName As String
	' Used only at Design Time.
	Private myComments As String
	' Tome's home directory
	Private myFullPath As String
	' Directory where Creator loaded from
	Private myLoadPath As String
	' FileName (not saved)
	Private myFileName As String
	
	' Name of Plot (.rsa file) to jump to when hit this entry point
	Dim myAreaIndex As Short
	' Name of Map within above Plot to jump to.
	Dim myMapIndex As Short
	' Name of EntryPoint within above Map.
	Dim myEntryIndex As Short
	
	' (X,Y) of Party on Plot.Map (not currently used)
	Private myMapX As Short
	Private myMapY As Short
	' This is run-time only data
	Private myMoveToX As Short
	Private myMoveToY As Short
	' Que of Steps from A* routine
	' Run-time only data
	Private myStepX() As Short
	Private myStepY() As Short
	Private myMaxStep As Short
	Private myNextStep As Short
	Private myLastEncounter As Short
	Private myCurrEncounter As Short
	Private myFleeToX As Short
	Private myFleeToY As Short
	
	' Current Turn, Cycle, Moon, Year
	Private myTurn As Byte ' 200 Turns per Cycle
	Private myCycle As Byte ' 50 Cycles per Moon
	Private myMoon As Byte ' 10 Moons per Year
	Private myYear As Short ' 10000 Years per Age
	' Real Seconds, Minutes, Hours, Days
	Private myRealSeconds As Short
	Private myRealMinutes As Short
	Private myRealHours As Short
	Private myRealDays As Short
	
	' Marked as True if this Tome needs to be saved (not saved)
	Private myDirty As Short
	' Used while loading up the main menu (not saved)
	Private myInWorld As Short
	
	' Bit0      -   On=OnAdventure, Off=Not OnAdventure
	' Bit1      -   On=IsReset, Off=Not Reset
	' Bit2      -   On=IsInPlay, Off=Not InPlay
	' Bit3-7    -   Not Used
	Private myOnAdventure As Byte
	
	' Max number of Creatures in the Party. 0 is no max.
	Private myPartySize As Byte
	' Max amount of Level in Party. 0 is no max.
	Private myPartyAvgLevel As Byte
	
	' Collection of Creature Objects which make up the Party. Upon arriving
	' at MainMenu, this is emptied of all CreatureWithTurns (they are saved back into
	' the Pool).
	Private myCreatures As New Collection
	' Triggers which follow with Party. Upon arriving at MainMenu, this
	' is cleared.
	Private myTriggers As New Collection
	' Journal Entries attached to this Tome
	Private myJournals As New Collection
	' This is the list of Factoids accquired by the Party.
	Private myFactoids As New Collection
	' Areas attached to this Tome
	Private myAreas As New Collection
	
	
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
	
	
	Public Property FullPath() As String
		Get
			FullPath = Trim(myFullPath)
		End Get
		Set(ByVal Value As String)
			myFullPath = Trim(Value)
		End Set
	End Property
	
	
	Public Property LoadPath() As String
		Get
			LoadPath = Trim(myLoadPath)
		End Get
		Set(ByVal Value As String)
			myLoadPath = Trim(Value)
		End Set
	End Property
	
	
	Public Property Filename() As String
		Get
			Filename = Trim(myFileName)
		End Get
		Set(ByVal Value As String)
			myFileName = Trim(Value)
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
	
	
	Public Property FleeToX() As Short
		Get
			FleeToX = myFleeToX
		End Get
		Set(ByVal Value As Short)
			myFleeToX = Value
		End Set
	End Property
	
	
	Public Property FleeToY() As Short
		Get
			FleeToY = myFleeToY
		End Get
		Set(ByVal Value As Short)
			myFleeToY = Value
		End Set
	End Property
	
	
	Public Property InWorld() As Short
		Get
			InWorld = myInWorld
		End Get
		Set(ByVal Value As Short)
			myInWorld = Value
		End Set
	End Property
	
	
	Public Property MoveToX() As Short
		Get
			MoveToX = myMoveToX
		End Get
		Set(ByVal Value As Short)
			myMaxStep = 0
			ReDim myStepX(0)
			ReDim myStepY(0)
			myMoveToX = Value
		End Set
	End Property
	
	
	Public Property MoveToY() As Short
		Get
			MoveToY = myMoveToY
		End Get
		Set(ByVal Value As Short)
			myMaxStep = 0
			ReDim myStepX(0)
			ReDim myStepY(0)
			myMoveToY = Value
		End Set
	End Property
	
	
	Public Property PartySize() As Short
		Get
			PartySize = myPartySize
		End Get
		Set(ByVal Value As Short)
			myPartySize = Value
		End Set
	End Property
	
	
	Public Property PartyAvgLevel() As Short
		Get
			PartyAvgLevel = myPartyAvgLevel
		End Get
		Set(ByVal Value As Short)
			myPartyAvgLevel = Value
		End Set
	End Property
	
	
	Public Property Turn() As Byte
		Get
			Turn = myTurn
		End Get
		Set(ByVal Value As Byte)
			myTurn = Value
		End Set
	End Property
	
	
	Public Property Cycle() As Byte
		Get
			Cycle = myCycle
		End Get
		Set(ByVal Value As Byte)
			myCycle = Value
		End Set
	End Property
	
	
	Public Property Moon() As Byte
		Get
			Moon = myMoon
		End Get
		Set(ByVal Value As Byte)
			myMoon = Value
		End Set
	End Property
	
	
	'UPGRADE_NOTE: Year was upgraded to Year_Renamed. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="A9E4979A-37FA-4718-9994-97DD76ED70A7"'
    Public Property Year() As Short
        Get
            Year = myYear
        End Get
        Set(ByVal Value As Short)
            myYear = Value
        End Set
    End Property
	
	
	Public Property RealSeconds() As Short
		Get
			RealSeconds = myRealSeconds
		End Get
		Set(ByVal Value As Short)
			myRealSeconds = Value
		End Set
	End Property
	
	
	Public Property RealMinutes() As Short
		Get
			RealMinutes = myRealMinutes
		End Get
		Set(ByVal Value As Short)
			myRealMinutes = Value
		End Set
	End Property
	
	
	Public Property RealHours() As Short
		Get
			RealHours = myRealHours
		End Get
		Set(ByVal Value As Short)
			myRealHours = Value
		End Set
	End Property
	
	
	Public Property RealDays() As Short
		Get
			RealDays = myRealDays
		End Get
		Set(ByVal Value As Short)
			myRealDays = Value
		End Set
	End Property
	
	
	Public Property Version() As Short
		Get
			Version = myVersion
		End Get
		Set(ByVal Value As Short)
			myVersion = Value
		End Set
	End Property
	
	
	Public Property OnAdventure() As Short
		Get
			OnAdventure = (myOnAdventure And &H1) > 0
		End Get
		Set(ByVal Value As Short)
			If Value Then
				myOnAdventure = myOnAdventure Or &H1
			Else
				myOnAdventure = myOnAdventure And (Not &H1)
			End If
		End Set
	End Property
	
	
	Public Property IsReset() As Short
		Get
			IsReset = (myOnAdventure And &H2) > 0
		End Get
		Set(ByVal Value As Short)
			If Value Then
				myOnAdventure = myOnAdventure Or &H2
			Else
				myOnAdventure = myOnAdventure And (Not &H2)
			End If
		End Set
	End Property
	
	
	Public Property IsInPlay() As Short
		Get
			IsInPlay = (myOnAdventure And &H4) > 0
		End Get
		Set(ByVal Value As Short)
			If Value Then
				myOnAdventure = myOnAdventure Or &H4
			Else
				myOnAdventure = myOnAdventure And (Not &H4)
			End If
		End Set
	End Property
	
	Public ReadOnly Property Triggers() As Collection
		Get
			Triggers = myTriggers
		End Get
	End Property
	
	Public ReadOnly Property Creatures() As Collection
		Get
			Creatures = myCreatures
		End Get
	End Property
	
	Public ReadOnly Property Factoids() As Collection
		Get
			Factoids = myFactoids
		End Get
	End Property
	
	Public ReadOnly Property Journals() As Collection
		Get
			Journals = myJournals
		End Get
	End Property
	
	Public ReadOnly Property Areas() As Collection
		Get
			Areas = myAreas
		End Get
	End Property
	
	Public ReadOnly Property Dirty() As Short
		Get
			Dirty = myDirty
		End Get
	End Property
	
	Public Sub AddStep(ByRef NewX As Short, ByRef NewY As Short)
		Dim c As Short
		' Preserve previous steps
		Dim KeepX(myMaxStep) As Short
		Dim KeepY(myMaxStep) As Short
		If myMaxStep > 0 Then
			For c = 0 To myMaxStep
				KeepX(c) = myStepX(c)
				KeepY(c) = myStepY(c)
			Next c
		End If
		' Resize for adding a step
		myMaxStep = myMaxStep + 1
		ReDim myStepX(myMaxStep)
		ReDim myStepY(myMaxStep)
		' Replace previous steps
		For c = 0 To myMaxStep - 1
			myStepX(c) = KeepX(c)
			myStepY(c) = KeepY(c)
		Next c
		' Add new step
		myStepX(myMaxStep) = NewX
		myStepY(myMaxStep) = NewY
	End Sub
	
	Public Sub NextStep(ByRef X As Short, ByRef Y As Short)
		Dim YDir, XDir, c As Short
		Dim CreatureX As Creature
		If myMaxStep > 0 Then
			c = 0
			XDir = myStepX(myMaxStep) - myMapX
			YDir = myStepY(myMaxStep) - myMapY
			For	Each CreatureX In myCreatures
				If XDir < 0 Then ' Moving Down Left
					Select Case CreatureX.TileSpot
						Case 0, 1, 3, 4, 6, 7
							CreatureX.TileSpot = CreatureX.TileSpot + 1
						Case Else
							CreatureX.TileSpot = CreatureX.TileSpot - 2
							CreatureX.MapX = myStepX(myMaxStep) : CreatureX.MapY = myStepY(myMaxStep)
					End Select
				ElseIf XDir > 0 Then  ' Moving Up Right
					Select Case CreatureX.TileSpot
						Case 1, 2, 4, 5, 7, 8
							CreatureX.TileSpot = CreatureX.TileSpot - 1
						Case Else
							CreatureX.TileSpot = CreatureX.TileSpot + 2
							CreatureX.MapX = myStepX(myMaxStep) : CreatureX.MapY = myStepY(myMaxStep)
					End Select
				ElseIf YDir < 0 Then  ' Moving Up Left
					Select Case CreatureX.TileSpot
						Case 3, 4, 5, 6, 7, 8
							CreatureX.TileSpot = CreatureX.TileSpot - 3
						Case Else
							CreatureX.TileSpot = CreatureX.TileSpot + 6
							CreatureX.MapX = myStepX(myMaxStep) : CreatureX.MapY = myStepY(myMaxStep)
					End Select
				ElseIf YDir > 0 Then  ' Moving Down Right
					Select Case CreatureX.TileSpot
						Case 0, 1, 2, 3, 4, 5
							CreatureX.TileSpot = CreatureX.TileSpot + 3
						Case Else
							CreatureX.TileSpot = CreatureX.TileSpot - 6
							CreatureX.MapX = myStepX(myMaxStep) : CreatureX.MapY = myStepY(myMaxStep)
					End Select
				End If
				If CreatureX.MapX = myStepX(myMaxStep) And CreatureX.MapY = myStepY(myMaxStep) Then
					c = c + 1
				End If
			Next CreatureX
			If c = myCreatures.Count() Then
				X = myStepX(myMaxStep) : Y = myStepY(myMaxStep)
				myMaxStep = myMaxStep - 1
			Else
				X = myMapX : Y = myMapY
			End If
		Else
			X = myMapX : Y = myMapY
		End If
	End Sub
	
	
	Public Sub NextMaxStep(ByRef X As Short, ByRef Y As Short)
		Dim YDir, XDir, c As Short
		Dim CreatureX As Creature
		If myMaxStep > 0 Then
			c = 0
			XDir = myStepX(myMaxStep) - myMapX
			YDir = myStepY(myMaxStep) - myMapY
			For	Each CreatureX In myCreatures
				CreatureX.MapX = myStepX(myMaxStep) : CreatureX.MapY = myStepY(myMaxStep)
			Next CreatureX
			X = myStepX(myMaxStep) : Y = myStepY(myMaxStep)
			myMaxStep = myMaxStep - 1
		Else
			X = myMapX : Y = myMapY
		End If
	End Sub
	
	Public Sub NextStepIs(ByRef X As Short, ByRef Y As Short)
		If myMaxStep > 0 Then
			X = myStepX(myMaxStep) : Y = myStepY(myMaxStep)
		Else
			X = myMapX : Y = myMapY
		End If
	End Sub
	
	Public Function AddTrigger() As Trigger
		' Adds a Trigger and returns a reference to that Trigger
		Dim c As Short
		Dim TriggerX As Trigger
		' Find new available unused index idenifier
		c = 1
		For	Each TriggerX In myTriggers
			If TriggerX.Index >= c Then
				c = TriggerX.Index + 1
			End If
		Next TriggerX
		' Create the Trigger, set the Index and add it.
		TriggerX = New Trigger
		TriggerX.Index = c
		TriggerX.Name = "Trigger" & c
		myTriggers.Add(TriggerX, "T" & TriggerX.Index)
		AddTrigger = TriggerX
	End Function
	
	Public Sub RemoveTrigger(ByRef DeleteKey As String)
		myTriggers.Remove(DeleteKey)
	End Sub
	
	Public Function AddCreature() As Creature
		' Adds a Creature and returns a reference to that Creature
		Dim c As Short
		Dim CreatureX As Creature
		' Find new available unused index identifier
		c = 1
		For	Each CreatureX In myCreatures
			If CreatureX.Index >= c Then
				c = CreatureX.Index + 1
			End If
		Next CreatureX
		' Create the Creature, set the Index and add it.
		CreatureX = New Creature
		CreatureX.Index = c
		CreatureX.Name = "Creature" & c
		myCreatures.Add(CreatureX, "X" & CreatureX.Index)
		AddCreature = CreatureX
	End Function
	
	Public Sub AppendCreature(ByRef CreatureX As Creature)
		' Adds a Creature and returns a reference to that Creature
		Dim c As Short
		Dim CreatureZ As Creature
		' Find new available unused index identifier
		c = 1
		For	Each CreatureZ In myCreatures
			If CreatureZ.Index >= c Then
				c = CreatureZ.Index + 1
			End If
		Next CreatureZ
		' Create the Creature, set the Index and add it.
		CreatureX.Index = c
		myCreatures.Add(CreatureX, "X" & CreatureX.Index)
	End Sub
	
	Public Sub RemoveCreature(ByRef DeleteKey As String)
		myCreatures.Remove(DeleteKey)
	End Sub
	
	Public Function AddFactoid() As Factoid
		' Adds a Factoid and returns a reference to that Factoid
		Dim c As Short
		Dim FactoidX As Factoid
		' Find new available unused index idenifier
		c = 1
		For	Each FactoidX In myFactoids
			If FactoidX.Index >= c Then
				c = FactoidX.Index + 1
			End If
		Next FactoidX
		' Create the Factoid, set the Index and add it.
		FactoidX = New Factoid
		FactoidX.Index = c
		myFactoids.Add(FactoidX, "F" & FactoidX.Index)
		AddFactoid = FactoidX
	End Function
	
	Public Sub RemoveFactoid(ByRef DeleteKey As String)
		myFactoids.Remove(DeleteKey)
	End Sub
	
	Public Function AddJournal() As Journal
		' Adds a Journal and returns a reference to that Journal
		Dim c As Short
		Dim JournalX As Journal
		' Find new available unused index idenifier
		c = 1
		For	Each JournalX In myJournals
			If JournalX.Index >= c Then
				c = JournalX.Index + 1
			End If
		Next JournalX
		' Create the Journal, set the Index and add it.
		JournalX = New Journal
		JournalX.Index = c
		myJournals.Add(JournalX, "J" & JournalX.Index)
		AddJournal = JournalX
	End Function
	
	Public Sub RemoveJournal(ByRef DeleteKey As String)
		'Public Sub RemoveJournal(Index As Integer)
		' [Titi 2.4.8] reverted to the "J" & index key
		myJournals.Remove(DeleteKey) 'Index
	End Sub
	
	Public Function AddArea() As Area
		' Adds a Area and returns a reference to that Area
		Dim c As Short
		Dim AreaX As Area
		' Find new available unused index idenifier
		c = 1
		For	Each AreaX In myAreas
			If AreaX.Index >= c Then
				c = AreaX.Index + 1
			End If
		Next AreaX
		' Create the Area, set the Index and add it.
		AreaX = New Area
		AreaX.Index = c
		AreaX.Name = "Area" & AreaX.Index
		myAreas.Add(AreaX, "A" & AreaX.Index)
		AddArea = AreaX
	End Function
	
	Public Sub RemoveArea(ByRef DeleteKey As String)
		myAreas.Remove(DeleteKey)
	End Sub
	
	Public Sub ReadFromFileHeader(ByRef FromFile As Short)
		Dim c As Short
		Dim Cnt As Integer
		' Read Item Name, Index and Comments
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
		' Read Comments
		'UPGRADE_WARNING: Get was upgraded to FileGet and has a new behavior. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="9B7D5ADD-D8FE-4819-A36C-6DEDAF088CC7"'
		FileGet(FromFile, Cnt)
		myComments = ""
		For c = 1 To Cnt
			myComments = myComments & " "
		Next c
		'UPGRADE_WARNING: Get was upgraded to FileGet and has a new behavior. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="9B7D5ADD-D8FE-4819-A36C-6DEDAF088CC7"'
		FileGet(FromFile, myComments)
		If Len(BreakText(myComments, 2)) > 0 Then
			myRealSeconds = CShort(BreakText(myComments, 2))
			myRealMinutes = CShort(BreakText(myComments, 3))
			myRealHours = CShort(BreakText(myComments, 4))
			myRealDays = CShort(BreakText(myComments, 5))
		Else
			myRealSeconds = 0
			myRealMinutes = 0
			myRealHours = 0
			myRealDays = 0
		End If
		myComments = BreakText(myComments, 1)
		' Read FullPath
		'UPGRADE_WARNING: Get was upgraded to FileGet and has a new behavior. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="9B7D5ADD-D8FE-4819-A36C-6DEDAF088CC7"'
		FileGet(FromFile, Cnt)
		myFullPath = ""
		For c = 1 To Cnt
			myFullPath = myFullPath & " "
		Next c
		'UPGRADE_WARNING: Get was upgraded to FileGet and has a new behavior. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="9B7D5ADD-D8FE-4819-A36C-6DEDAF088CC7"'
		FileGet(FromFile, myFullPath)
		' Read for Maps
		'UPGRADE_WARNING: Get was upgraded to FileGet and has a new behavior. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="9B7D5ADD-D8FE-4819-A36C-6DEDAF088CC7"'
		FileGet(FromFile, myAreaIndex)
		'UPGRADE_WARNING: Get was upgraded to FileGet and has a new behavior. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="9B7D5ADD-D8FE-4819-A36C-6DEDAF088CC7"'
		FileGet(FromFile, myMapIndex)
		'UPGRADE_WARNING: Get was upgraded to FileGet and has a new behavior. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="9B7D5ADD-D8FE-4819-A36C-6DEDAF088CC7"'
		FileGet(FromFile, myEntryIndex)
		'UPGRADE_WARNING: Get was upgraded to FileGet and has a new behavior. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="9B7D5ADD-D8FE-4819-A36C-6DEDAF088CC7"'
		FileGet(FromFile, myMapX)
		'UPGRADE_WARNING: Get was upgraded to FileGet and has a new behavior. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="9B7D5ADD-D8FE-4819-A36C-6DEDAF088CC7"'
		FileGet(FromFile, myMapY)
		'UPGRADE_WARNING: Get was upgraded to FileGet and has a new behavior. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="9B7D5ADD-D8FE-4819-A36C-6DEDAF088CC7"'
		FileGet(FromFile, myTurn)
		'UPGRADE_WARNING: Get was upgraded to FileGet and has a new behavior. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="9B7D5ADD-D8FE-4819-A36C-6DEDAF088CC7"'
		FileGet(FromFile, myCycle)
		'UPGRADE_WARNING: Get was upgraded to FileGet and has a new behavior. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="9B7D5ADD-D8FE-4819-A36C-6DEDAF088CC7"'
		FileGet(FromFile, myMoon)
		'UPGRADE_WARNING: Get was upgraded to FileGet and has a new behavior. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="9B7D5ADD-D8FE-4819-A36C-6DEDAF088CC7"'
		FileGet(FromFile, myYear)
		'UPGRADE_WARNING: Get was upgraded to FileGet and has a new behavior. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="9B7D5ADD-D8FE-4819-A36C-6DEDAF088CC7"'
		FileGet(FromFile, myPartySize)
		'UPGRADE_WARNING: Get was upgraded to FileGet and has a new behavior. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="9B7D5ADD-D8FE-4819-A36C-6DEDAF088CC7"'
		FileGet(FromFile, myPartyAvgLevel)
		'UPGRADE_WARNING: Get was upgraded to FileGet and has a new behavior. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="9B7D5ADD-D8FE-4819-A36C-6DEDAF088CC7"'
		FileGet(FromFile, myOnAdventure)
	End Sub
	
	Public Sub ReadFromFile(ByRef FromFile As Short)
		On Error GoTo ErrorHandler
		Dim c As Short
		Dim Cnt As Integer
		Dim myTrigger As Trigger
		Dim myCreature As Creature
		Dim myFactoid As Factoid
		Dim myJournal As Journal
		Dim myArea As Area
		Dim Reading As String
		' Read header information
		ReadFromFileHeader(FromFile)
		' Read Party Triggers for Tome
		Reading = "Trigger"
		'UPGRADE_WARNING: Get was upgraded to FileGet and has a new behavior. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="9B7D5ADD-D8FE-4819-A36C-6DEDAF088CC7"'
		FileGet(FromFile, Cnt)
		For c = 1 To Cnt
			myTrigger = New Trigger
			myTrigger.ReadFromFile(FromFile)
			myTriggers.Add(myTrigger, "T" & myTrigger.Index)
		Next c
		' Read Party Members for Tome
		Reading = "Creature"
		'UPGRADE_WARNING: Get was upgraded to FileGet and has a new behavior. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="9B7D5ADD-D8FE-4819-A36C-6DEDAF088CC7"'
		FileGet(FromFile, Cnt)
		For c = 1 To Cnt
			myCreature = New Creature
			myCreature.ReadFromFile(FromFile)
			myCreatures.Add(myCreature, "X" & myCreature.Index)
		Next c
		' Read Party Journals for Tome
		Reading = "Journal"
		'UPGRADE_WARNING: Get was upgraded to FileGet and has a new behavior. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="9B7D5ADD-D8FE-4819-A36C-6DEDAF088CC7"'
		FileGet(FromFile, Cnt)
		For c = 1 To Cnt
			myJournal = New Journal
			myJournal.ReadFromFile(FromFile)
			myJournals.Add(myJournal, "J" & myJournal.Index)
		Next c
		' Read Party Factoids for Tome
		Reading = "Factoid"
		'UPGRADE_WARNING: Get was upgraded to FileGet and has a new behavior. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="9B7D5ADD-D8FE-4819-A36C-6DEDAF088CC7"'
		FileGet(FromFile, Cnt)
		For c = 1 To Cnt
			myFactoid = New Factoid
			myFactoid.ReadFromFile(FromFile)
			myFactoids.Add(myFactoid, "F" & myFactoid.Index)
		Next c
		' Read Party Areas for Tome
		Reading = "Area"
		'UPGRADE_WARNING: Get was upgraded to FileGet and has a new behavior. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="9B7D5ADD-D8FE-4819-A36C-6DEDAF088CC7"'
		FileGet(FromFile, Cnt)
		For c = 1 To Cnt
			myArea = New Area
			myArea.ReadFromFile(FromFile)
			myAreas.Add(myArea, "A" & myArea.Index)
		Next c
		myDirty = False
		Exit Sub
ErrorHandler: 
		' [Titi 2.4.9] re-index components if necessary
		If Err.Number = 457 And Reading <> "" Then
			Select Case Reading
				Case "Creature"
					oErr.logError(Reading & " index conflict: " & myCreature.Index)
				Case "Trigger"
					oErr.logError(Reading & " index conflict: " & myTrigger.Index)
				Case "Journal"
					oErr.logError(Reading & " index conflict: " & myJournal.Index)
				Case "Factoid"
					oErr.logError(Reading & " index conflict: " & myFactoid.Index)
				Case "Area"
					oErr.logError(Reading & " index conflict: " & myArea.Index)
			End Select
		Else
			oErr.logError("Cannot read tome" & IIf(Reading <> vbNullString, "'s " & Reading, "") & ", error#" & Err.Number & " (" & Err.Description & ")")
		End If
	End Sub
	
	Public Sub SaveToFile(ByRef ToFile As Short)
		Dim c As Short
		Dim tmp As String
		' Save Item Name, Index and Comments
		'UPGRADE_WARNING: Put was upgraded to FilePut and has a new behavior. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="9B7D5ADD-D8FE-4819-A36C-6DEDAF088CC7"'
		FilePut(ToFile, myVersion)
		'UPGRADE_WARNING: Put was upgraded to FilePut and has a new behavior. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="9B7D5ADD-D8FE-4819-A36C-6DEDAF088CC7"'
		FilePut(ToFile, Len(myName))
		'UPGRADE_WARNING: Put was upgraded to FilePut and has a new behavior. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="9B7D5ADD-D8FE-4819-A36C-6DEDAF088CC7"'
		FilePut(ToFile, myName)
		If myRealSeconds > 0 Or myRealMinutes > 0 Or myRealHours > 0 Or myRealDays > 0 Then
			tmp = myComments & "|" & myRealSeconds & "|" & myRealMinutes & "|" & myRealHours & "|" & myRealDays
		Else
			tmp = myComments
		End If
		'UPGRADE_WARNING: Put was upgraded to FilePut and has a new behavior. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="9B7D5ADD-D8FE-4819-A36C-6DEDAF088CC7"'
		FilePut(ToFile, Len(tmp))
		'UPGRADE_WARNING: Put was upgraded to FilePut and has a new behavior. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="9B7D5ADD-D8FE-4819-A36C-6DEDAF088CC7"'
		FilePut(ToFile, tmp)
		'UPGRADE_WARNING: Put was upgraded to FilePut and has a new behavior. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="9B7D5ADD-D8FE-4819-A36C-6DEDAF088CC7"'
		FilePut(ToFile, Len(myFullPath))
		'UPGRADE_WARNING: Put was upgraded to FilePut and has a new behavior. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="9B7D5ADD-D8FE-4819-A36C-6DEDAF088CC7"'
		FilePut(ToFile, myFullPath)
		' Save PlotFile
		'UPGRADE_WARNING: Put was upgraded to FilePut and has a new behavior. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="9B7D5ADD-D8FE-4819-A36C-6DEDAF088CC7"'
		FilePut(ToFile, myAreaIndex)
		'UPGRADE_WARNING: Put was upgraded to FilePut and has a new behavior. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="9B7D5ADD-D8FE-4819-A36C-6DEDAF088CC7"'
		FilePut(ToFile, myMapIndex)
		'UPGRADE_WARNING: Put was upgraded to FilePut and has a new behavior. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="9B7D5ADD-D8FE-4819-A36C-6DEDAF088CC7"'
		FilePut(ToFile, myEntryIndex)
		'UPGRADE_WARNING: Put was upgraded to FilePut and has a new behavior. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="9B7D5ADD-D8FE-4819-A36C-6DEDAF088CC7"'
		FilePut(ToFile, myMapX)
		'UPGRADE_WARNING: Put was upgraded to FilePut and has a new behavior. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="9B7D5ADD-D8FE-4819-A36C-6DEDAF088CC7"'
		FilePut(ToFile, myMapY)
		'UPGRADE_WARNING: Put was upgraded to FilePut and has a new behavior. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="9B7D5ADD-D8FE-4819-A36C-6DEDAF088CC7"'
		FilePut(ToFile, myTurn)
		'UPGRADE_WARNING: Put was upgraded to FilePut and has a new behavior. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="9B7D5ADD-D8FE-4819-A36C-6DEDAF088CC7"'
		FilePut(ToFile, myCycle)
		'UPGRADE_WARNING: Put was upgraded to FilePut and has a new behavior. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="9B7D5ADD-D8FE-4819-A36C-6DEDAF088CC7"'
		FilePut(ToFile, myMoon)
		'UPGRADE_WARNING: Put was upgraded to FilePut and has a new behavior. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="9B7D5ADD-D8FE-4819-A36C-6DEDAF088CC7"'
		FilePut(ToFile, myYear)
		'UPGRADE_WARNING: Put was upgraded to FilePut and has a new behavior. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="9B7D5ADD-D8FE-4819-A36C-6DEDAF088CC7"'
		FilePut(ToFile, myPartySize)
		'UPGRADE_WARNING: Put was upgraded to FilePut and has a new behavior. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="9B7D5ADD-D8FE-4819-A36C-6DEDAF088CC7"'
		FilePut(ToFile, myPartyAvgLevel)
		'UPGRADE_WARNING: Put was upgraded to FilePut and has a new behavior. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="9B7D5ADD-D8FE-4819-A36C-6DEDAF088CC7"'
		FilePut(ToFile, myOnAdventure)
		' Save Triggers
		'UPGRADE_WARNING: Put was upgraded to FilePut and has a new behavior. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="9B7D5ADD-D8FE-4819-A36C-6DEDAF088CC7"'
		FilePut(ToFile, myTriggers.Count())
		For c = 1 To myTriggers.Count()
			'UPGRADE_WARNING: Couldn't resolve default property of object myTriggers().SaveToFile. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
			myTriggers.Item(c).SaveToFile(ToFile)
		Next c
		' Save Party
		'UPGRADE_WARNING: Put was upgraded to FilePut and has a new behavior. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="9B7D5ADD-D8FE-4819-A36C-6DEDAF088CC7"'
		FilePut(ToFile, myCreatures.Count())
		For c = 1 To myCreatures.Count()
			'UPGRADE_WARNING: Couldn't resolve default property of object myCreatures().SaveToFile. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
			myCreatures.Item(c).SaveToFile(ToFile)
		Next c
		' Save Journals
		'UPGRADE_WARNING: Put was upgraded to FilePut and has a new behavior. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="9B7D5ADD-D8FE-4819-A36C-6DEDAF088CC7"'
		FilePut(ToFile, myJournals.Count())
		For c = 1 To myJournals.Count()
			'UPGRADE_WARNING: Couldn't resolve default property of object myJournals().SaveToFile. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
			myJournals.Item(c).SaveToFile(ToFile)
		Next c
		' Save Factoids
		'UPGRADE_WARNING: Put was upgraded to FilePut and has a new behavior. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="9B7D5ADD-D8FE-4819-A36C-6DEDAF088CC7"'
		FilePut(ToFile, myFactoids.Count())
		For c = 1 To myFactoids.Count()
			'UPGRADE_WARNING: Couldn't resolve default property of object myFactoids().SaveToFile. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
			myFactoids.Item(c).SaveToFile(ToFile)
		Next c
		' Save Areas
		'UPGRADE_WARNING: Put was upgraded to FilePut and has a new behavior. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="9B7D5ADD-D8FE-4819-A36C-6DEDAF088CC7"'
		FilePut(ToFile, myAreas.Count())
		For c = 1 To myAreas.Count()
			'UPGRADE_WARNING: Couldn't resolve default property of object myAreas().SaveToFile. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
			myAreas.Item(c).SaveToFile(ToFile)
		Next c
	End Sub
	
	'UPGRADE_NOTE: Class_Initialize was upgraded to Class_Initialize_Renamed. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="A9E4979A-37FA-4718-9994-97DD76ED70A7"'
	Private Sub Class_Initialize_Renamed()
		myName = "Untitled"
		myLoadPath = "Not Saved"
		myTurn = 100
		myCycle = 25
		myMoon = 5
		myYear = 5000
		myPartyAvgLevel = 1
		myPartySize = 5
	End Sub
	Public Sub New()
		MyBase.New()
		Class_Initialize_Renamed()
	End Sub
End Class