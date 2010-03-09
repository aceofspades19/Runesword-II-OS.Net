Option Strict Off
Option Explicit On
Friend Class Trigger
	
	' Version number of Class
	Private myVersion As Short
	' Index for Trigger
	Private myIndex As Short
	' Name for Trigger
	Private myName As String
	' Comments for this Trigger
	Private myComments As String
	' Trigger On (for Text of this, see Function)
	Private myTriggerType As Byte
	' Turns / Turns (how long it lasts) or Cost for a Skill.
	Private myTurns As Byte
	' Variables
	Private myB, myA, myC As Byte
	
	' Can have SkillPoints associated with it. The Party gets these points
	' if they deactivate the Trigger. Or it can be used for other reasons
	' inside the Trigger.Statements. This is Total spent for a Skill.
	Private mySkillPoints As Short
	' Run-Time Only: Store temporary SkillPoints spent. Allows resetting of spend in game.
	Private myTempSkill As Short
	
	' Trigger Style (Cast is special meaning this Trigger is a spell)
	' Bit 0     - On=Skill
	' Bit 1     - On=Trap
	' Bit 2     - On=Curse
	' Bit 3     - On=Poison
	' Bit 4     - On=Evil
	' Bit 5     - On=Fear
	' Bit 6     - On=Magic
	' Bit 7     - On=Timed
	' Bit 8     - On=Lunacy
	' Bit 9     - On=Revelry
	' Bit 10    - On=Wrath
	' Bit 11    - On=Pride
	' Bit 12    - On=Greed
	' Bit 13    - On=Lust
	' Bit 14    - On=Fish
	' Bit 15    - On=Not Used
	Private myStyle As Short
	' Collection of Statements in Package
	Private myStatements As New Collection
	' Collection of Things (Creature, Item or Trigger) in Trigger
	Private myCreatures As New Collection
	Private myItems As New Collection
	Private myTriggers As New Collection
	
	
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
	
	
	Public Property TriggerType() As Byte
		Get
			TriggerType = myTriggerType
		End Get
		Set(ByVal Value As Byte)
			myTriggerType = Value
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
	
	
	Public Property TempSkill() As Short
		Get
			TempSkill = myTempSkill
		End Get
		Set(ByVal Value As Short)
			myTempSkill = Value
		End Set
	End Property
	
	
	Public Property Statements() As Collection
		Get
			Statements = myStatements
		End Get
		Set(ByVal Value As Collection)
			myStatements = Value
		End Set
	End Property
	
	
	Public Property Turns() As Byte
		Get
			Turns = myTurns
		End Get
		Set(ByVal Value As Byte)
			myTurns = Value
		End Set
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
	
	
	Public Property Style(ByVal Index As Short) As Short
		Get
			Style = (myStyle And (2 ^ Index)) > 0
		End Get
		Set(ByVal Value As Short)
			If Value Then
				myStyle = myStyle Or (2 ^ Index)
			Else
				myStyle = myStyle And (Not 2 ^ Index)
			End If
		End Set
	End Property
	
	
	Public Property Styles() As Short
		Get
			Styles = myStyle
		End Get
		Set(ByVal Value As Short)
			myStyle = Value
		End Set
	End Property
	
	
	Public Property IsSkill() As Short
		Get
			IsSkill = (myStyle And &H1) > 0
		End Get
		Set(ByVal Value As Short)
			If Value Then
				myStyle = myStyle Or &H1
			Else
				myStyle = myStyle And (Not &H1)
			End If
		End Set
	End Property
	
	
	Public Property IsTrap() As Short
		Get
			IsTrap = (myStyle And &H2) > 0
		End Get
		Set(ByVal Value As Short)
			If Value Then
				myStyle = myStyle Or &H2
			Else
				myStyle = myStyle And (Not &H2)
			End If
		End Set
	End Property
	
	
	Public Property IsPrimarySkill() As Short
		Get
			IsPrimarySkill = (myStyle And &H2) > 0
		End Get
		Set(ByVal Value As Short)
			If Value Then
				myStyle = myStyle Or &H2
			Else
				myStyle = myStyle And (Not &H2)
			End If
		End Set
	End Property
	
	
	Public Property IsCurse() As Short
		Get
			IsCurse = (myStyle And &H4) > 0
		End Get
		Set(ByVal Value As Short)
			If Value Then
				myStyle = myStyle Or &H4
			Else
				myStyle = myStyle And (Not &H4)
			End If
		End Set
	End Property
	
	
	Public Property IsPoison() As Short
		Get
			IsPoison = (myStyle And &H8) > 0
		End Get
		Set(ByVal Value As Short)
			If Value Then
				myStyle = myStyle Or &H8
			Else
				myStyle = myStyle And (Not &H8)
			End If
		End Set
	End Property
	
	
	Public Property IsEvil() As Short
		Get
			IsEvil = (myStyle And &H10) > 0
		End Get
		Set(ByVal Value As Short)
			If Value Then
				myStyle = myStyle Or &H10
			Else
				myStyle = myStyle And (Not &H10)
			End If
		End Set
	End Property
	
	
	Public Property IsFear() As Short
		Get
			IsFear = (myStyle And &H20) > 0
		End Get
		Set(ByVal Value As Short)
			If Value Then
				myStyle = myStyle Or &H20
			Else
				myStyle = myStyle And (Not &H20)
			End If
		End Set
	End Property
	
	
	Public Property IsMagic() As Short
		Get
			IsMagic = (myStyle And &H40) > 0
		End Get
		Set(ByVal Value As Short)
			If Value Then
				myStyle = myStyle Or &H40
			Else
				myStyle = myStyle And (Not &H40)
			End If
		End Set
	End Property
	
	
	Public Property IsTimed() As Short
		Get
			IsTimed = (myStyle And &H80) > 0
		End Get
		Set(ByVal Value As Short)
			If Value Then
				myStyle = myStyle Or &H80
			Else
				myStyle = myStyle And (Not &H80)
			End If
		End Set
	End Property
	
	
	Public Property IsLunacy() As Short
		Get
			IsLunacy = (myStyle And &H100) > 0
		End Get
		Set(ByVal Value As Short)
			If Value Then
				myStyle = myStyle Or &H100
			Else
				myStyle = myStyle And (Not &H100)
			End If
		End Set
	End Property
	
	
	Public Property IsRevelry() As Short
		Get
			IsRevelry = (myStyle And &H200) > 0
		End Get
		Set(ByVal Value As Short)
			If Value Then
				myStyle = myStyle Or &H200
			Else
				myStyle = myStyle And (Not &H200)
			End If
		End Set
	End Property
	
	
	Public Property IsWrath() As Short
		Get
			IsWrath = (myStyle And &H400) > 0
		End Get
		Set(ByVal Value As Short)
			If Value Then
				myStyle = myStyle Or &H400
			Else
				myStyle = myStyle And (Not &H400)
			End If
		End Set
	End Property
	
	
	Public Property IsPride() As Short
		Get
			IsPride = (myStyle And &H800) > 0
		End Get
		Set(ByVal Value As Short)
			If Value Then
				myStyle = myStyle Or &H800
			Else
				myStyle = myStyle And (Not &H800)
			End If
		End Set
	End Property
	
	
	Public Property IsGreed() As Short
		Get
			IsGreed = (myStyle And &H1000) > 0
		End Get
		Set(ByVal Value As Short)
			If Value Then
				myStyle = myStyle Or &H1000
			Else
				myStyle = myStyle And (Not &H1000)
			End If
		End Set
	End Property
	
	
	Public Property IsLust() As Short
		Get
			IsLust = (myStyle And &H2000) > 0
		End Get
		Set(ByVal Value As Short)
			If Value Then
				myStyle = myStyle Or &H2000
			Else
				myStyle = myStyle And (Not &H2000)
			End If
		End Set
	End Property
	
	
	Public Property IsFish() As Short
		Get
			IsFish = (myStyle And &H4000) > 0
		End Get
		Set(ByVal Value As Short)
			If Value Then
				myStyle = myStyle Or &H4000
			Else
				myStyle = myStyle And (Not &H4000)
			End If
		End Set
	End Property
	
	
	Public Property VarA() As Byte
		Get
			VarA = myA
		End Get
		Set(ByVal Value As Byte)
			myA = Value
		End Set
	End Property
	
	
	Public Property VarB() As Byte
		Get
			VarB = myB
		End Get
		Set(ByVal Value As Byte)
			myB = Value
		End Set
	End Property
	
	
	Public Property VarC() As Byte
		Get
			VarC = myC
		End Get
		Set(ByVal Value As Byte)
			myC = Value
		End Set
	End Property
	
	Public Function AddCreature() As Creature
		' Adds an Creature and returns a reference to that Creature
		Dim c As Short
		Dim CreatureX As Creature
		' Find new available unused index idenifier
		c = 1
		For	Each CreatureX In myCreatures
			If CreatureX.Index >= c Then
				c = CreatureX.Index + 1
			End If
		Next CreatureX
		' Set the index and add the door. Return the new door's index.
		CreatureX = New Creature
		CreatureX.Index = c
		CreatureX.Name = "Creature" & c
		myCreatures.Add(CreatureX, "X" & CreatureX.Index)
		AddCreature = CreatureX
	End Function
	
	Public Function AddCreatureAsIndex(ByRef Index As Short) As Creature
		' Adds an Creature and returns a reference to that Creature

		Dim CreatureX As Creature
		' Set the index and add the door. Return the new door's index.
		CreatureX = New Creature
		CreatureX.Index = Index
		CreatureX.Name = "Creature" & Index
		myCreatures.Add(CreatureX, "X" & CreatureX.Index)
		AddCreatureAsIndex = CreatureX
	End Function
	
	Public Sub RemoveCreature(ByRef DeleteKey As String)
		myCreatures.Remove(DeleteKey)
	End Sub
	
	Public Function AddItem() As Item
		' Adds an Item and returns a reference to that Item
		Dim c As Short
		Dim ItemX As Item
		' Find new available unused index idenifier
		c = 1
		For	Each ItemX In myItems
			If ItemX.Index >= c Then
				c = ItemX.Index + 1
			End If
		Next ItemX
		' Set the index and add the door. Return the new door's index.
		ItemX = New Item
		ItemX.Index = c
		ItemX.Name = "Item" & c
		myItems.Add(ItemX, "I" & ItemX.Index)
		AddItem = ItemX
	End Function
	
	Public Function AddItemAsIndex(ByRef Index As Short) As Item
		' Adds an Item and returns a reference to that Item

		Dim ItemX As Item
		' Set the index and add the door. Return the new door's index.
		ItemX = New Item
		ItemX.Index = Index
		ItemX.Name = "Item" & Index
		myItems.Add(ItemX, "I" & ItemX.Index)
		AddItemAsIndex = ItemX
	End Function
	
	Public Sub RemoveItem(ByRef DeleteKey As String)
		myItems.Remove(DeleteKey)
	End Sub
	
	Public Function AddTrigger() As Trigger
		' Adds an Item and returns a reference to that Item
		Dim c As Short
		Dim TriggerX As Trigger
		' Find new available unused index idenifier
		c = 1
		For	Each TriggerX In myTriggers
			If TriggerX.Index >= c Then
				c = TriggerX.Index + 1
			End If
		Next TriggerX
		' Set the index and add the door. Return the new door's index.
		TriggerX = New Trigger
		TriggerX.Index = c
		TriggerX.Name = "Trigger" & c
		myTriggers.Add(TriggerX, "T" & TriggerX.Index)
		AddTrigger = TriggerX
	End Function
	
	Public Function AddTriggerAsIndex(ByRef Index As Short) As Trigger
		' Adds an Item and returns a reference to that Item

		Dim TriggerX As Trigger
		' Set the index and add the door. Return the new door's index.
		TriggerX = New Trigger
		TriggerX.Index = Index
		TriggerX.Name = "Trigger" & Index
		myTriggers.Add(TriggerX, "T" & TriggerX.Index)
		AddTriggerAsIndex = TriggerX
	End Function
	
	Public Sub RemoveTrigger(ByRef DeleteKey As String)
		myTriggers.Remove(DeleteKey)
	End Sub
	
	Public Function AddStatement(Optional ByRef AfterIndex As Object = Nothing) As Statement
		' Adds an Statement and returns a reference to that Statement
        Dim c As Short
		Dim StatementX As Statement
		' Find new available unused index idenifier
		c = 1
		For	Each StatementX In myStatements
			If StatementX.Index >= c Then
				c = StatementX.Index + 1
			End If
		Next StatementX
		' Create the Statement and set the index
		StatementX = New Statement
		StatementX.Index = c
		' Either add to end of Statements or Before a Statement
		'UPGRADE_NOTE: IsMissing() was changed to IsNothing(). Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="8AE1CB93-37AB-439A-A4FF-BE3B6760BB23"'
		If IsNothing(AfterIndex) Then
			myStatements.Add(StatementX, "S" & StatementX.Index)
		Else
			'UPGRADE_WARNING: Couldn't resolve default property of object AfterIndex. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
			myStatements.Add(StatementX, "S" & StatementX.Index,  , "S" & AfterIndex)
		End If
		AddStatement = StatementX
	End Function
	
	Public Function AddStatementAsIndex(ByRef Index As Short) As Statement
		' Adds an Statement and returns a reference to that Statement

		Dim StatementX As Statement
		' Create the Statement and set the index
		StatementX = New Statement
		StatementX.Index = Index
		myStatements.Add(StatementX, "S" & StatementX.Index)
		AddStatementAsIndex = StatementX
	End Function
	
	Public Sub Copy(ByRef FromTrigger As Trigger)

		Dim StatementX As Statement
		Dim CreatureX As Creature
		Dim ItemX As Item
		Dim TriggerX As Trigger
		Dim myStatement As Statement
		Dim myCreature As Creature
		Dim myItem As Item
		Dim myTrigger As Trigger
		' Copy Name
		myName = FromTrigger.Name
		' Copy Comments
		myComments = FromTrigger.Comments
		' Copy Type, Turns and A, B and C
		myTriggerType = FromTrigger.TriggerType
		myTurns = FromTrigger.Turns
		myA = FromTrigger.VarA
		myB = FromTrigger.VarB
		myC = FromTrigger.VarC
		mySkillPoints = FromTrigger.SkillPoints
		myStyle = FromTrigger.Styles
		myTempSkill = FromTrigger.TempSkill
		' Copy Statements to Trigger
		myStatements = New Collection
		For	Each StatementX In FromTrigger.Statements
			myStatement = Me.AddStatementAsIndex((StatementX.Index))
			myStatement.Copy(StatementX)
		Next StatementX
		' Copy Creatures to Trigger
		myCreatures = New Collection
		For	Each CreatureX In FromTrigger.Creatures
			myCreature = Me.AddCreatureAsIndex((CreatureX.Index))
			myCreature.Copy(CreatureX)
		Next CreatureX
		' Copy Items to Trigger
		myItems = New Collection
		For	Each ItemX In FromTrigger.Items
			myItem = Me.AddItemAsIndex((ItemX.Index))
			myItem.Copy(ItemX)
		Next ItemX
		' Copy Triggers to Trigger
		myTriggers = New Collection
		For	Each TriggerX In FromTrigger.Triggers
			myTrigger = Me.AddTriggerAsIndex((TriggerX.Index))
			myTrigger.Copy(TriggerX)
		Next TriggerX
	End Sub
	
	Public Sub ReadFromFile(ByRef FromFile As Short)
		On Error GoTo ErrorHandler
		Dim c As Short
		Dim Cnt As Integer
		Dim myStatement As Statement
		Dim myCreature As Creature
		Dim myItem As Item
		Dim myTrigger As Trigger
		Dim Reading As String
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
		' Read Trigger Index
		'UPGRADE_WARNING: Get was upgraded to FileGet and has a new behavior. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="9B7D5ADD-D8FE-4819-A36C-6DEDAF088CC7"'
		FileGet(FromFile, myIndex)
		' Read Comments
		'UPGRADE_WARNING: Get was upgraded to FileGet and has a new behavior. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="9B7D5ADD-D8FE-4819-A36C-6DEDAF088CC7"'
		FileGet(FromFile, Cnt)
		myComments = ""
		For c = 1 To Cnt
			myComments = myComments & " "
		Next c
		'UPGRADE_WARNING: Get was upgraded to FileGet and has a new behavior. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="9B7D5ADD-D8FE-4819-A36C-6DEDAF088CC7"'
		FileGet(FromFile, myComments)
		' Read Type, Turns and A, B and C
		'UPGRADE_WARNING: Get was upgraded to FileGet and has a new behavior. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="9B7D5ADD-D8FE-4819-A36C-6DEDAF088CC7"'
		FileGet(FromFile, myTriggerType)
		'UPGRADE_WARNING: Get was upgraded to FileGet and has a new behavior. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="9B7D5ADD-D8FE-4819-A36C-6DEDAF088CC7"'
		FileGet(FromFile, myTurns)
		'UPGRADE_WARNING: Get was upgraded to FileGet and has a new behavior. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="9B7D5ADD-D8FE-4819-A36C-6DEDAF088CC7"'
		FileGet(FromFile, myA)
		'UPGRADE_WARNING: Get was upgraded to FileGet and has a new behavior. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="9B7D5ADD-D8FE-4819-A36C-6DEDAF088CC7"'
		FileGet(FromFile, myB)
		'UPGRADE_WARNING: Get was upgraded to FileGet and has a new behavior. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="9B7D5ADD-D8FE-4819-A36C-6DEDAF088CC7"'
		FileGet(FromFile, myC)
		'UPGRADE_WARNING: Get was upgraded to FileGet and has a new behavior. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="9B7D5ADD-D8FE-4819-A36C-6DEDAF088CC7"'
		FileGet(FromFile, mySkillPoints)
		'UPGRADE_WARNING: Get was upgraded to FileGet and has a new behavior. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="9B7D5ADD-D8FE-4819-A36C-6DEDAF088CC7"'
		FileGet(FromFile, myStyle)
		' Read Statements for Trigger
		myStatements = New Collection
		Reading = "Statement"
		'UPGRADE_WARNING: Get was upgraded to FileGet and has a new behavior. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="9B7D5ADD-D8FE-4819-A36C-6DEDAF088CC7"'
		FileGet(FromFile, Cnt)
		For c = 1 To Cnt
			myStatement = New Statement
			myStatement.ReadFromFile(FromFile)
			myStatements.Add(myStatement, "S" & myStatement.Index)
		Next c
		' Read Creatures for Trigger
		Reading = "Creature"
		'UPGRADE_WARNING: Get was upgraded to FileGet and has a new behavior. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="9B7D5ADD-D8FE-4819-A36C-6DEDAF088CC7"'
		FileGet(FromFile, Cnt)
		For c = 1 To Cnt
			myCreature = New Creature
			myCreature.ReadFromFile(FromFile)
			myCreatures.Add(myCreature, "X" & myCreature.Index)
		Next c
		' Read Items for Trigger
		Reading = "Item"
		'UPGRADE_WARNING: Get was upgraded to FileGet and has a new behavior. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="9B7D5ADD-D8FE-4819-A36C-6DEDAF088CC7"'
		FileGet(FromFile, Cnt)
		For c = 1 To Cnt
			myItem = New Item
			myItem.ReadFromFile(FromFile)
			myItems.Add(myItem, "I" & myItem.Index)
		Next c
		' Read Triggers for Trigger
		Reading = "Sub-trigger"
		'UPGRADE_WARNING: Get was upgraded to FileGet and has a new behavior. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="9B7D5ADD-D8FE-4819-A36C-6DEDAF088CC7"'
		FileGet(FromFile, Cnt)
		For c = 1 To Cnt
			myTrigger = New Trigger
			myTrigger.ReadFromFile(FromFile)
			myTriggers.Add(myTrigger, "T" & myTrigger.Index)
		Next c
		Exit Sub
ErrorHandler: 
		' [Titi 2.4.9]
		If Err.Number = 457 And Reading <> "" Then
			Select Case Reading
				Case "Statement"
					oErr.logError(Reading & " index conflict: " & myStatement.Index)
				Case "Creature"
					oErr.logError(Reading & " index conflict: " & myCreature.Index)
				Case "Item"
					oErr.logError(Reading & " index conflict: " & myItem.Index)
				Case "Sub-trigger"
					oErr.logError(Reading & " index conflict: " & myTrigger.Index)
			End Select
		Else
			oErr.logError("Cannot read trigger" & IIf(Reading <> vbNullString, "'s " & Reading, "") & ", error#" & Err.Number & " (" & Err.Description & ")")
		End If
	End Sub
	
	Public Sub SaveToFile(ByRef ToFile As Short)
		Dim c As Short
		' Save Trigger Index and Comments
		'UPGRADE_WARNING: Put was upgraded to FilePut and has a new behavior. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="9B7D5ADD-D8FE-4819-A36C-6DEDAF088CC7"'
		FilePut(ToFile, myVersion)
		'UPGRADE_WARNING: Put was upgraded to FilePut and has a new behavior. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="9B7D5ADD-D8FE-4819-A36C-6DEDAF088CC7"'
		FilePut(ToFile, Len(myName))
		'UPGRADE_WARNING: Put was upgraded to FilePut and has a new behavior. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="9B7D5ADD-D8FE-4819-A36C-6DEDAF088CC7"'
		FilePut(ToFile, myName)
		'UPGRADE_WARNING: Put was upgraded to FilePut and has a new behavior. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="9B7D5ADD-D8FE-4819-A36C-6DEDAF088CC7"'
		FilePut(ToFile, myIndex)
		'UPGRADE_WARNING: Put was upgraded to FilePut and has a new behavior. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="9B7D5ADD-D8FE-4819-A36C-6DEDAF088CC7"'
		FilePut(ToFile, Len(myComments))
		'UPGRADE_WARNING: Put was upgraded to FilePut and has a new behavior. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="9B7D5ADD-D8FE-4819-A36C-6DEDAF088CC7"'
		FilePut(ToFile, myComments)
		' Save Type, Turns and A, B and C
		'UPGRADE_WARNING: Put was upgraded to FilePut and has a new behavior. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="9B7D5ADD-D8FE-4819-A36C-6DEDAF088CC7"'
		FilePut(ToFile, myTriggerType)
		'UPGRADE_WARNING: Put was upgraded to FilePut and has a new behavior. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="9B7D5ADD-D8FE-4819-A36C-6DEDAF088CC7"'
		FilePut(ToFile, myTurns)
		'UPGRADE_WARNING: Put was upgraded to FilePut and has a new behavior. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="9B7D5ADD-D8FE-4819-A36C-6DEDAF088CC7"'
		FilePut(ToFile, myA)
		'UPGRADE_WARNING: Put was upgraded to FilePut and has a new behavior. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="9B7D5ADD-D8FE-4819-A36C-6DEDAF088CC7"'
		FilePut(ToFile, myB)
		'UPGRADE_WARNING: Put was upgraded to FilePut and has a new behavior. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="9B7D5ADD-D8FE-4819-A36C-6DEDAF088CC7"'
		FilePut(ToFile, myC)
		'UPGRADE_WARNING: Put was upgraded to FilePut and has a new behavior. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="9B7D5ADD-D8FE-4819-A36C-6DEDAF088CC7"'
		FilePut(ToFile, mySkillPoints)
		'UPGRADE_WARNING: Put was upgraded to FilePut and has a new behavior. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="9B7D5ADD-D8FE-4819-A36C-6DEDAF088CC7"'
		FilePut(ToFile, myStyle)
		' Save Statements for Trigger
		'UPGRADE_WARNING: Put was upgraded to FilePut and has a new behavior. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="9B7D5ADD-D8FE-4819-A36C-6DEDAF088CC7"'
		FilePut(ToFile, myStatements.Count())
		For c = 1 To myStatements.Count()
			'UPGRADE_WARNING: Couldn't resolve default property of object myStatements().SaveToFile. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
			myStatements.Item(c).SaveToFile(ToFile)
		Next c
		' Save Creatures for Trigger
		'UPGRADE_WARNING: Put was upgraded to FilePut and has a new behavior. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="9B7D5ADD-D8FE-4819-A36C-6DEDAF088CC7"'
		FilePut(ToFile, myCreatures.Count())
		For c = 1 To myCreatures.Count()
			'UPGRADE_WARNING: Couldn't resolve default property of object myCreatures().SaveToFile. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
			myCreatures.Item(c).SaveToFile(ToFile)
		Next c
		' Save Items for Trigger
		'UPGRADE_WARNING: Put was upgraded to FilePut and has a new behavior. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="9B7D5ADD-D8FE-4819-A36C-6DEDAF088CC7"'
		FilePut(ToFile, myItems.Count())
		For c = 1 To myItems.Count()
			'UPGRADE_WARNING: Couldn't resolve default property of object myItems().SaveToFile. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
			myItems.Item(c).SaveToFile(ToFile)
		Next c
		' Save Triggers for Trigger
		'UPGRADE_WARNING: Put was upgraded to FilePut and has a new behavior. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="9B7D5ADD-D8FE-4819-A36C-6DEDAF088CC7"'
		FilePut(ToFile, myTriggers.Count())
		For c = 1 To myTriggers.Count()
			'UPGRADE_WARNING: Couldn't resolve default property of object myTriggers().SaveToFile. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
			myTriggers.Item(c).SaveToFile(ToFile)
		Next c
	End Sub
	
	'UPGRADE_NOTE: Class_Initialize was upgraded to Class_Initialize_Renamed. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="A9E4979A-37FA-4718-9994-97DD76ED70A7"'
	Private Sub Class_Initialize_Renamed()
		myTriggerType = 0 ' None
		myA = 0
		myB = 0
		myC = 0
	End Sub
	Public Sub New()
		MyBase.New()
		Class_Initialize_Renamed()
	End Sub
End Class