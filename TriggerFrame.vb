Option Strict Off
Option Explicit On
Friend Class TriggerFrame
	' Unique index for frame in Collection
	Dim myIndex As Short
	' Trigger reference in frame
	Dim myTrigger As Trigger
	' Local Byte variables
	Dim myByteB, myByteA, myByteC As Byte
	' Local Integer variables
	Dim myIntegerB, myIntegerA, myIntegerC As Short
	' Local Random variable
	Dim myRandom As Short
	' Local String variable
	Dim myTextB, myTextA, myTextC As String
	' Local Object variables
	Dim myParentX As Object
	Dim myTileX(4) As Short
	Dim myTileY(4) As Short
	Dim myEncounterX(4) As Encounter
	Dim myCreatureX(4) As Creature
	Dim myItemX(4) As Item
	Dim myTriggerX(4) As Trigger
	' Current statement processing
	Dim myStatementNow, myProcess As Short
	' Total number of Statements processed
	Dim myTicks As Short
	' Abort and StopExit flag
	Dim myAbort, myStopExit As Short
	' Failure Flag for Statements being processed
	Dim myFail As Short
	' FoundIt Flag for the Find Statement
	Dim myFoundIt As Short
	' Stacks for If, Loop and Case statements
	Dim myIfStack(15) As Short
	Dim myLoopStack(15) As Short
	Dim myEachStack(15) As Short
	Dim myCaseStack(15) As Short
	Dim myIfStackTop As Short
	Dim myLoopStackTop As Short
	
	
	Public Property Index() As Short
		Get
			Index = myIndex
		End Get
		Set(ByVal Value As Short)
			myIndex = Value
		End Set
	End Property
	
	
	Public Property ByteA() As Byte
		Get
			ByteA = myByteA
		End Get
		Set(ByVal Value As Byte)
			myByteA = Value
		End Set
	End Property
	
	
	Public Property ByteB() As Byte
		Get
			ByteB = myByteB
		End Get
		Set(ByVal Value As Byte)
			myByteB = Value
		End Set
	End Property
	
	
	Public Property ByteC() As Byte
		Get
			ByteC = myByteC
		End Get
		Set(ByVal Value As Byte)
			myByteC = Value
		End Set
	End Property
	
	
	Public Property IntegerA() As Short
		Get
			IntegerA = myIntegerA
		End Get
		Set(ByVal Value As Short)
			myIntegerA = Value
		End Set
	End Property
	
	
	Public Property IntegerB() As Short
		Get
			IntegerB = myIntegerB
		End Get
		Set(ByVal Value As Short)
			myIntegerB = Value
		End Set
	End Property
	
	
	Public Property IntegerC() As Short
		Get
			IntegerC = myIntegerC
		End Get
		Set(ByVal Value As Short)
			myIntegerC = Value
		End Set
	End Property
	
	
	Public Property Random() As Short
		Get
			Random = Int(myRandom * Rnd()) + 1
		End Get
		Set(ByVal Value As Short)
			myRandom = Value
		End Set
	End Property
	
	
	Public Property TextA() As String
		Get
			TextA = myTextA
		End Get
		Set(ByVal Value As String)
			myTextA = Value
		End Set
	End Property
	
	
	Public Property TextB() As String
		Get
			TextB = myTextB
		End Get
		Set(ByVal Value As String)
			myTextB = Value
		End Set
	End Property
	
	
	Public Property TextC() As String
		Get
			TextC = myTextC
		End Get
		Set(ByVal Value As String)
			myTextC = Value
		End Set
	End Property
	
	
	Public Property ParentX() As Object
		Get
			ParentX = myParentX
		End Get
		Set(ByVal Value As Object)
			myParentX = Value
		End Set
	End Property
	
	
	Public Property EncounterX(ByVal Index As Short) As Encounter
		Get
			EncounterX = myEncounterX(Index)
		End Get
		Set(ByVal Value As Encounter)
			myEncounterX(Index) = Value
		End Set
	End Property
	
	
	Public Property CreatureX(ByVal Index As Short) As Creature
		Get
			CreatureX = myCreatureX(Index)
		End Get
		Set(ByVal Value As Creature)
			myCreatureX(Index) = Value
		End Set
	End Property
	
	
	Public Property ItemX(ByVal Index As Short) As Item
		Get
			ItemX = myItemX(Index)
		End Get
		Set(ByVal Value As Item)
			myItemX(Index) = Value
		End Set
	End Property
	
	
	Public Property TriggerX(ByVal Index As Short) As Trigger
		Get
			TriggerX = myTriggerX(Index)
		End Get
		Set(ByVal Value As Trigger)
			myTriggerX(Index) = Value
		End Set
	End Property
	
	
	Public Property TileX(ByVal Index As Short) As Short
		Get
			TileX = myTileX(Index)
		End Get
		Set(ByVal Value As Short)
			myTileX(Index) = Value
		End Set
	End Property
	
	
	Public Property TileY(ByVal Index As Short) As Short
		Get
			TileY = myTileY(Index)
		End Get
		Set(ByVal Value As Short)
			myTileY(Index) = Value
		End Set
	End Property
	
	
	Public Property StopExit() As Short
		Get
			StopExit = myStopExit
		End Get
		Set(ByVal Value As Short)
			myStopExit = Value
		End Set
	End Property
	
	
	Public Property Ticks() As Short
		Get
			Ticks = myTicks
		End Get
		Set(ByVal Value As Short)
			myTicks = Value
		End Set
	End Property
	
	
	Public Property IfStackTop() As Short
		Get
			IfStackTop = myIfStackTop
		End Get
		Set(ByVal Value As Short)
			myIfStackTop = Value
		End Set
	End Property
	
	
	Public Property IfStack(ByVal Index As Short) As Short
		Get
			IfStack = myIfStack(Index)
		End Get
		Set(ByVal Value As Short)
			myIfStack(Index) = Value
		End Set
	End Property
	
	
	Public Property Abort() As Short
		Get
			Abort = myAbort
		End Get
		Set(ByVal Value As Short)
			myAbort = Value
		End Set
	End Property
	
	
	Public Property Fail() As Short
		Get
			Fail = myFail
		End Get
		Set(ByVal Value As Short)
			myFail = Value
		End Set
	End Property
	
	
	Public Property FoundIt() As Short
		Get
			FoundIt = myFoundIt
		End Get
		Set(ByVal Value As Short)
			myFoundIt = Value
		End Set
	End Property
	
	
	Public Property Process() As Short
		Get
			Process = myProcess
		End Get
		Set(ByVal Value As Short)
			myProcess = Value
		End Set
	End Property
	
	
	Public Property LoopStackTop() As Short
		Get
			LoopStackTop = myLoopStackTop
		End Get
		Set(ByVal Value As Short)
			myLoopStackTop = Value
		End Set
	End Property
	
	
	Public Property LoopStack(ByVal Index As Short) As Short
		Get
			LoopStack = myLoopStack(Index)
		End Get
		Set(ByVal Value As Short)
			myLoopStack(Index) = Value
		End Set
	End Property
	
	
	Public Property CaseStack(ByVal Index As Short) As Short
		Get
			CaseStack = myCaseStack(Index)
		End Get
		Set(ByVal Value As Short)
			myCaseStack(Index) = Value
		End Set
	End Property
	
	
	Public Property EachStack(ByVal Index As Short) As Short
		Get
			EachStack = myEachStack(Index)
		End Get
		Set(ByVal Value As Short)
			myEachStack(Index) = Value
		End Set
	End Property
	
	
	Public Property StatementNow() As Short
		Get
			StatementNow = myStatementNow
		End Get
		Set(ByVal Value As Short)
			myStatementNow = Value
		End Set
	End Property
	
	Public Property Trigger() As Trigger
		Get
			Trigger = myTrigger
		End Get
		Set(ByVal Value As Trigger)
			myTrigger = Value
		End Set
	End Property
	
	'UPGRADE_NOTE: Class_Initialize was upgraded to Class_Initialize_Renamed. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="A9E4979A-37FA-4718-9994-97DD76ED70A7"'
	Private Sub Class_Initialize_Renamed()
		Dim c As Short
		myByteA = 0 : myByteB = 0 : myByteC = 0
		myIntegerA = 0 : myIntegerB = 0 : myIntegerC = 0
		myTextA = "" : myTextB = "" : myTextC = ""
		myStatementNow = 1
		For c = 1 To 3
			myCreatureX(c) = New Creature
			myEncounterX(c) = New Encounter
			myItemX(c) = New Item
			myTriggerX(c) = New Trigger
		Next c
		myAbort = False
		myStopExit = False
		myFail = False
	End Sub
	Public Sub New()
		MyBase.New()
		Class_Initialize_Renamed()
	End Sub
End Class