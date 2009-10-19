Option Strict Off
Option Explicit On
Friend Class Theme
	
	' Version number of Class
	Dim myVersion As Short
	' Unique number for this Theme
	Dim myIndex As Short
	' Name of Theme
	Dim myName As String
	' Comments for Theme
	Dim myComments As String
	' Maximum numer of Encounters covered by the Theme
	Dim myCoverage As Short
	' Number of Encounters already covered by Theme
	Dim myEncounterCount As Short
	' Bit1      - IsQuest? (Can span multiple maps) On=Yes, Off=No
	' Bit2-3    - Not Used
	' Bit4      - IsMajorStory? (Awards Skill Points) On=Yes, Off=No
	' Bit5-7    - Not Used
	Dim myFlags As Short
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
	' Total SkillPoints to value total Encounters built by this Theme
	Dim mySkillPoints As Short
	' List of Encounters for Theme
	Dim myEncounters As New Collection
	' List of Creatures for Theme.
	Dim myCreatures As New Collection
	' List of Items for Theme.
	Dim myItems As New Collection
	' List of Triggers (usually Traps because they can appear anywhere).
	Dim myTriggers As New Collection
	' List of descriptions for this Theme. Applies to Encounter
	' FirstEntry description only.
	Dim myFactoids As New Collection
	
	
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
	
	
	Public Property EncounterCount() As Short
		Get
			EncounterCount = myEncounterCount
		End Get
		Set(ByVal Value As Short)
			myEncounterCount = Value
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
	
	Public ReadOnly Property EncounterPoints() As Short
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
					c = 1 * c
				Case 1 ' 2-3
					c = 3 * c
				Case 2 ' 4-6
					c = 5 * c
				Case 3 ' 7-9
					c = 10 * c
			End Select
			' Randomize (EncounterPoints vary each time you ask for them)
			Select Case Int(Rnd() * 100)
				Case 0
					EncounterPoints = Int(c / 4)
				Case 1 To 5
					EncounterPoints = Int(c / 3)
				Case 6 To 15
					EncounterPoints = Int(c / 2)
				Case 16 To 25
					EncounterPoints = Int(c / 1.2)
				Case 26 To 75
					EncounterPoints = c
				Case 76 To 85
					EncounterPoints = Int(c * 1.2)
				Case 86 To 94
					EncounterPoints = Int(c * 2)
				Case 95 To 98
					EncounterPoints = Int(c * 3)
				Case 99
					EncounterPoints = Int(c * 4)
			End Select
		End Get
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
	
	
	Public Property Coverage() As Short
		Get
			Coverage = myCoverage
		End Get
		Set(ByVal Value As Short)
			myCoverage = Value
		End Set
	End Property
	
	
	Public Property SkillPoints() As Short
		Get
			SkillPoints = mySkillPoints
		End Get
		Set(ByVal Value As Short)
			mySkillPoints = Value
		End Set
	End Property
	
	Public ReadOnly Property Encounters() As Collection
		Get
			Encounters = myEncounters
		End Get
	End Property
	
	Public ReadOnly Property Creatures() As Collection
		Get
			Creatures = myCreatures
		End Get
	End Property
	
	Public ReadOnly Property Items() As Collection
		Get
			Items = myItems
		End Get
	End Property
	
	Public ReadOnly Property Triggers() As Collection
		Get
			Triggers = myTriggers
		End Get
	End Property
	
	Public ReadOnly Property Factoids() As Collection
		Get
			Factoids = myFactoids
		End Get
	End Property
	
	Public Function AddEncounter() As Encounter
		' Adds an Encounter and returns a reference to that Encounter
		Dim c As Short
		Dim EncounterX As Encounter
		' Find new available unused index identifier
		c = 1
		For	Each EncounterX In myEncounters
			If EncounterX.Index >= c Then
				c = EncounterX.Index + 1
			End If
		Next EncounterX
		' Set the index and add the encounter. Return the new encounter's index.
		EncounterX = New Encounter
		EncounterX.Index = c
		EncounterX.Name = "Encounter" & c
		EncounterX.ParentTheme = myIndex
		myEncounters.Add(EncounterX, "E" & EncounterX.Index)
		AddEncounter = EncounterX
	End Function
	
	Public Sub RemoveEncounter(ByRef DeleteKey As String)
		myEncounters.Remove(DeleteKey)
	End Sub
	
	Public Function AddCreature() As Creature
		' Adds an Creature and returns a reference to that Creature
		Dim c As Short
		Dim CreatureX As Creature
		' Find new available unused index identifier
		c = 1
		For	Each CreatureX In myCreatures
			If CreatureX.Index >= c Then
				c = CreatureX.Index + 1
			End If
		Next CreatureX
		' Set the index and add the Creature. Return the new Creature's index.
		CreatureX = New Creature
		CreatureX.Index = c
		CreatureX.Name = "Creature" & c
		myCreatures.Add(CreatureX, "X" & CreatureX.Index)
		AddCreature = CreatureX
	End Function
	
	Public Sub RemoveCreature(ByRef DeleteKey As String)
		myCreatures.Remove(DeleteKey)
	End Sub
	
	Public Function AddItem() As Item
		' Adds an Item and returns a reference to that Item
		Dim c As Short
		Dim ItemX As Item
		' Find new available unused index identifier
		c = 1
		For	Each ItemX In myItems
			If ItemX.Index >= c Then
				c = ItemX.Index + 1
			End If
		Next ItemX
		' Set the index and add the Item. Return the new Item's index.
		ItemX = New Item
		ItemX.Index = c
		ItemX.Name = "Item" & c
		myItems.Add(ItemX, "I" & ItemX.Index)
		AddItem = ItemX
	End Function
	
	Public Sub RemoveItem(ByRef DeleteKey As String)
		myItems.Remove(DeleteKey)
	End Sub
	
	Public Function AddTrigger() As Trigger
		' Adds an Trigger and returns a reference to that Trigger
		Dim c As Short
		Dim TriggerX As Trigger
		' Find new available unused index identifier
		c = 1
		For	Each TriggerX In myTriggers
			If TriggerX.Index >= c Then
				c = TriggerX.Index + 1
			End If
		Next TriggerX
		' Set the index and add the Trigger. Return the new Trigger's index.
		TriggerX = New Trigger
		TriggerX.Index = c
		TriggerX.Name = "Trigger" & c
		myTriggers.Add(TriggerX, "T" & TriggerX.Index)
		AddTrigger = TriggerX
	End Function
	
	Public Sub RemoveTrigger(ByRef DeleteKey As String)
		myTriggers.Remove(DeleteKey)
	End Sub
	
	Public Function AddFactoid() As Factoid
		' Adds an Description and returns a reference to that Description
		Dim c As Short
		Dim FactoidX As Factoid
		' Find new available unused index identifier
		c = 1
		For	Each FactoidX In myFactoids
			If FactoidX.Index >= c Then
				c = FactoidX.Index + 1
			End If
		Next FactoidX
		' Set the index and add the Description. Return the new Description's index.
		FactoidX = New Factoid
		FactoidX.Index = c
		myFactoids.Add(FactoidX, "F" & FactoidX.Index)
		AddFactoid = FactoidX
	End Function
	
	Public Sub RemoveFactoid(ByRef DeleteKey As String)
		myFactoids.Remove(DeleteKey)
	End Sub
	
	Public Sub Copy(ByRef FromTheme As Theme)
		Dim c As Short
		Dim myEncounter As Encounter
		Dim myItem As Item
		Dim myCreature As Creature
		Dim myTrigger As Trigger
		Dim myFactoid As Factoid
		' Copy Theme Name
		myName = FromTheme.Name
		myComments = FromTheme.Comments
		myCoverage = FromTheme.Coverage
		mySkillPoints = FromTheme.SkillPoints
		myStyle = FromTheme.Style
		myFlags = FromTheme.Flags
		' Copy Encounters for Theme
		myEncounters = New Collection
		For c = 1 To FromTheme.Encounters.Count()
			myEncounter = Me.AddEncounter
			'UPGRADE_WARNING: Couldn't resolve default property of object FromTheme.Encounters(). Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
			myEncounter.Copy(FromTheme.Encounters.Item(c))
		Next c
		' Copy Creatures for Theme
		myCreatures = New Collection
		For c = 1 To FromTheme.Creatures.Count()
			myCreature = Me.AddCreature
			'UPGRADE_WARNING: Couldn't resolve default property of object FromTheme.Creatures(). Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
			myCreature.Copy(FromTheme.Creatures.Item(c))
		Next c
		' Copy Items for Theme
		myItems = New Collection
		For c = 1 To FromTheme.Items.Count()
			myItem = Me.AddItem
			'UPGRADE_WARNING: Couldn't resolve default property of object FromTheme.Items(). Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
			myItem.Copy(FromTheme.Items.Item(c))
		Next c
		' Copy Triggers for Theme
		myTriggers = New Collection
		For c = 1 To FromTheme.Triggers.Count()
			myTrigger = Me.AddTrigger
			'UPGRADE_WARNING: Couldn't resolve default property of object FromTheme.Triggers(). Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
			myTrigger.Copy(FromTheme.Triggers.Item(c))
		Next c
		' Copy Descriptions for Theme
		myFactoids = New Collection
		For c = 1 To FromTheme.Factoids.Count()
			myFactoid = Me.AddFactoid
			'UPGRADE_WARNING: Couldn't resolve default property of object FromTheme.Factoids(). Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
			myFactoid.Copy(FromTheme.Factoids.Item(c))
		Next c
	End Sub
	
	Public Sub ReadFromFileHeader(ByRef FromFile As Short)
		Dim c As Short
		Dim Cnt As Integer
		' Read Theme Name
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
		' Read Theme Comments
		'UPGRADE_WARNING: Get was upgraded to FileGet and has a new behavior. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="9B7D5ADD-D8FE-4819-A36C-6DEDAF088CC7"'
		FileGet(FromFile, Cnt)
		myComments = ""
		For c = 1 To Cnt
			myComments = myComments & " "
		Next c
		'UPGRADE_WARNING: Get was upgraded to FileGet and has a new behavior. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="9B7D5ADD-D8FE-4819-A36C-6DEDAF088CC7"'
		FileGet(FromFile, myComments)
		'UPGRADE_WARNING: Get was upgraded to FileGet and has a new behavior. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="9B7D5ADD-D8FE-4819-A36C-6DEDAF088CC7"'
		FileGet(FromFile, myCoverage)
		'UPGRADE_WARNING: Get was upgraded to FileGet and has a new behavior. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="9B7D5ADD-D8FE-4819-A36C-6DEDAF088CC7"'
		FileGet(FromFile, myIndex)
		'UPGRADE_WARNING: Get was upgraded to FileGet and has a new behavior. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="9B7D5ADD-D8FE-4819-A36C-6DEDAF088CC7"'
		FileGet(FromFile, mySkillPoints)
		'UPGRADE_WARNING: Get was upgraded to FileGet and has a new behavior. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="9B7D5ADD-D8FE-4819-A36C-6DEDAF088CC7"'
		FileGet(FromFile, myStyle)
		'UPGRADE_WARNING: Get was upgraded to FileGet and has a new behavior. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="9B7D5ADD-D8FE-4819-A36C-6DEDAF088CC7"'
		FileGet(FromFile, myFlags)
	End Sub
	
	Public Sub ReadFromFile(ByRef FromFile As Short)
		On Error GoTo ErrorHandler
		Dim c As Short
		Dim Cnt As Integer
		Dim myEncounter As Encounter
		Dim myItem As Item
		Dim myCreature As Creature
		Dim myTrigger As Trigger
		Dim myFactoid As Factoid
		Dim Reading As String
		ReadFromFileHeader(FromFile)
		' Read Encounters for Theme
		Reading = "Encounter"
		'UPGRADE_WARNING: Get was upgraded to FileGet and has a new behavior. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="9B7D5ADD-D8FE-4819-A36C-6DEDAF088CC7"'
		FileGet(FromFile, Cnt)
		For c = 1 To Cnt
			myEncounter = New Encounter
			myEncounter.ReadFromFile(FromFile)
			myEncounters.Add(myEncounter, "E" & myEncounter.Index)
		Next c
		' Read Creatures for Theme
		Reading = "Creature"
		'UPGRADE_WARNING: Get was upgraded to FileGet and has a new behavior. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="9B7D5ADD-D8FE-4819-A36C-6DEDAF088CC7"'
		FileGet(FromFile, Cnt)
		For c = 1 To Cnt
			myCreature = New Creature
			myCreature.ReadFromFile(FromFile)
			myCreatures.Add(myCreature, "X" & myCreature.Index)
		Next c
		' Read Items for Theme
		Reading = "Item"
		'UPGRADE_WARNING: Get was upgraded to FileGet and has a new behavior. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="9B7D5ADD-D8FE-4819-A36C-6DEDAF088CC7"'
		FileGet(FromFile, Cnt)
		For c = 1 To Cnt
			myItem = New Item
			myItem.ReadFromFile(FromFile)
			myItems.Add(myItem, "I" & myItem.Index)
		Next c
		' Read Triggers for Theme
		Reading = "Trigger"
		'UPGRADE_WARNING: Get was upgraded to FileGet and has a new behavior. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="9B7D5ADD-D8FE-4819-A36C-6DEDAF088CC7"'
		FileGet(FromFile, Cnt)
		For c = 1 To Cnt
			myTrigger = New Trigger
			myTrigger.ReadFromFile(FromFile)
			myTriggers.Add(myTrigger, "T" & myTrigger.Index)
		Next c
		' Read Descriptions for Theme
		Reading = "Description"
		'UPGRADE_WARNING: Get was upgraded to FileGet and has a new behavior. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="9B7D5ADD-D8FE-4819-A36C-6DEDAF088CC7"'
		FileGet(FromFile, Cnt)
		For c = 1 To Cnt
			myFactoid = New Factoid
			myFactoid.ReadFromFile(FromFile)
			myFactoids.Add(myFactoid, "F" & myFactoid.Index)
		Next c
		Exit Sub
ErrorHandler: 
		' [Titi 2.4.9]
		If Err.Number = 457 And Reading <> "" Then
			Select Case Reading
				Case "Encounter"
					oErr.logError(Reading & " index conflict: " & myEncounter.Index)
				Case "Creature"
					oErr.logError(Reading & " index conflict: " & myCreature.Index)
				Case "Item"
					oErr.logError(Reading & " index conflict: " & myItem.Index)
				Case "Trigger"
					oErr.logError(Reading & " index conflict: " & myTrigger.Index)
				Case "Description"
					oErr.logError(Reading & " index conflict: " & myFactoid.Index)
			End Select
		Else
			oErr.logError("Cannot read theme" & IIf(Reading <> vbNullString, "'s " & Reading, "") & ", error#" & Err.Number & " (" & Err.Description & ")")
		End If
	End Sub
	
	Public Sub SaveToFile(ByRef ToFile As Short)
		Dim c As Short
		' Save Theme Name
		'UPGRADE_WARNING: Put was upgraded to FilePut and has a new behavior. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="9B7D5ADD-D8FE-4819-A36C-6DEDAF088CC7"'
		FilePut(ToFile, myVersion)
		'UPGRADE_WARNING: Put was upgraded to FilePut and has a new behavior. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="9B7D5ADD-D8FE-4819-A36C-6DEDAF088CC7"'
		FilePut(ToFile, Len(myName))
		'UPGRADE_WARNING: Put was upgraded to FilePut and has a new behavior. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="9B7D5ADD-D8FE-4819-A36C-6DEDAF088CC7"'
		FilePut(ToFile, myName)
		'UPGRADE_WARNING: Put was upgraded to FilePut and has a new behavior. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="9B7D5ADD-D8FE-4819-A36C-6DEDAF088CC7"'
		FilePut(ToFile, Len(myComments))
		'UPGRADE_WARNING: Put was upgraded to FilePut and has a new behavior. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="9B7D5ADD-D8FE-4819-A36C-6DEDAF088CC7"'
		FilePut(ToFile, myComments)
		'UPGRADE_WARNING: Put was upgraded to FilePut and has a new behavior. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="9B7D5ADD-D8FE-4819-A36C-6DEDAF088CC7"'
		FilePut(ToFile, myCoverage)
		'UPGRADE_WARNING: Put was upgraded to FilePut and has a new behavior. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="9B7D5ADD-D8FE-4819-A36C-6DEDAF088CC7"'
		FilePut(ToFile, myIndex)
		'UPGRADE_WARNING: Put was upgraded to FilePut and has a new behavior. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="9B7D5ADD-D8FE-4819-A36C-6DEDAF088CC7"'
		FilePut(ToFile, mySkillPoints)
		'UPGRADE_WARNING: Put was upgraded to FilePut and has a new behavior. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="9B7D5ADD-D8FE-4819-A36C-6DEDAF088CC7"'
		FilePut(ToFile, myStyle)
		'UPGRADE_WARNING: Put was upgraded to FilePut and has a new behavior. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="9B7D5ADD-D8FE-4819-A36C-6DEDAF088CC7"'
		FilePut(ToFile, myFlags)
		' Save Encounters for Theme
		'UPGRADE_WARNING: Put was upgraded to FilePut and has a new behavior. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="9B7D5ADD-D8FE-4819-A36C-6DEDAF088CC7"'
		FilePut(ToFile, myEncounters.Count())
		For c = 1 To myEncounters.Count()
			'UPGRADE_WARNING: Couldn't resolve default property of object myEncounters().SaveToFile. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
			myEncounters.Item(c).SaveToFile(ToFile)
		Next c
		' Save Creatures for Theme
		'UPGRADE_WARNING: Put was upgraded to FilePut and has a new behavior. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="9B7D5ADD-D8FE-4819-A36C-6DEDAF088CC7"'
		FilePut(ToFile, myCreatures.Count())
		For c = 1 To myCreatures.Count()
			'UPGRADE_WARNING: Couldn't resolve default property of object myCreatures().SaveToFile. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
			myCreatures.Item(c).SaveToFile(ToFile)
		Next c
		' Save Items for Theme
		'UPGRADE_WARNING: Put was upgraded to FilePut and has a new behavior. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="9B7D5ADD-D8FE-4819-A36C-6DEDAF088CC7"'
		FilePut(ToFile, myItems.Count())
		For c = 1 To myItems.Count()
			'UPGRADE_WARNING: Couldn't resolve default property of object myItems().SaveToFile. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
			myItems.Item(c).SaveToFile(ToFile)
		Next c
		' Save Triggers for Theme
		'UPGRADE_WARNING: Put was upgraded to FilePut and has a new behavior. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="9B7D5ADD-D8FE-4819-A36C-6DEDAF088CC7"'
		FilePut(ToFile, myTriggers.Count())
		For c = 1 To myTriggers.Count()
			'UPGRADE_WARNING: Couldn't resolve default property of object myTriggers().SaveToFile. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
			myTriggers.Item(c).SaveToFile(ToFile)
		Next c
		' Save Descriptions for Theme
		'UPGRADE_WARNING: Put was upgraded to FilePut and has a new behavior. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="9B7D5ADD-D8FE-4819-A36C-6DEDAF088CC7"'
		FilePut(ToFile, myFactoids.Count())
		For c = 1 To myFactoids.Count()
			'UPGRADE_WARNING: Couldn't resolve default property of object myFactoids().SaveToFile. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
			myFactoids.Item(c).SaveToFile(ToFile)
		Next c
	End Sub
	
	'UPGRADE_NOTE: Class_Initialize was upgraded to Class_Initialize_Renamed. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="A9E4979A-37FA-4718-9994-97DD76ED70A7"'
	Private Sub Class_Initialize_Renamed()
		myName = "Unnamed"
		Me.PartySize = 2
		Me.PartyAvgLevel = 0
		Me.Coverage = 1
	End Sub
	Public Sub New()
		MyBase.New()
		Class_Initialize_Renamed()
	End Sub
End Class