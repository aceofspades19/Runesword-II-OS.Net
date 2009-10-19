Option Strict Off
Option Explicit On
Friend Class TomeSketch
	
	' Unique Index
	Private myIndex As Short
	' Name of Tome
	Private myName As String
	' Comments for Tome
	Private myComments As String
	' Location of Tome File
	Private myFullPath As String
	' If a Tome Map: Max number of Creatures in the Party. 0 is no max.
	Private myPartySize As Byte
	' If a Tome Map: Max amount of Level in Party. 0 is no max.
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
	
	
	Public Property Comments() As String
		Get
			Comments = Trim(myComments)
		End Get
		Set(ByVal Value As String)
			myComments = Trim(Value)
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
		' Adds a Trigger and returns a reference to that Trigger
		Dim c As Short
		Dim CreatureX As Creature
		' Find new available unused index idenifier
		c = 1
		For	Each CreatureX In myCreatures
			If CreatureX.Index >= c Then
				c = CreatureX.Index + 1
			End If
		Next CreatureX
		' Create the Trigger, set the Index and add it.
		CreatureX = New Creature
		CreatureX.Index = c
		CreatureX.Name = "Creature" & c
		myCreatures.Add(CreatureX, "X" & CreatureX.Index)
		AddCreature = CreatureX
	End Function
	
	Public Sub AppendCreature(ByRef CreatureX As Creature)
		' Adds a Trigger and returns a reference to that Trigger
		Dim c As Short
		Dim CreatureZ As Creature
		' Find new available unused index idenifier
		c = 1
		For	Each CreatureZ In myCreatures
			If CreatureZ.Index >= c Then
				c = CreatureZ.Index + 1
			End If
		Next CreatureZ
		' Create the Trigger, set the Index and add it.
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
		myJournals.Remove(DeleteKey)
	End Sub
End Class