Option Strict Off
Option Explicit On
Friend Class Encounter
	' Encounter is zero or more Events, zero or more Creatures and zero
	' or more Items
	
	' Version number for Class
	Private myVersion As Short
	' myName is used for reference in Plot only (show in TreeView)
	Private myName As String
	' myIndex is a unique number for this Encounter in the Map
	Private myIndex As Short
	' myFirstEntry is text displayed when you first enter the Encounter
	Private myFirstEntry As String
	' myFirstEntry is text displayed when you enter second and subsequent times
	Private mySecondEntry As String
	' Flags:
	' Bit 0     &H1     On=Use as Hint?, Off=Don't use as hint
	' Bit 1     &H2     On=Have already entered, Off=Have not entered
	' Bit 2     &H4     On=Cannot fight here, Off=Can fight here.
	' Bit 3     &H8     On=Can Talk here, Off=Cannot talk here
	' Bit 4     &H10    On=Can Flee this encounter, Off=Cannot flee
	' Bit 5     &H20    On=ReGen on Entry, Off=No auto ReGen
	' Bit 6     &H40    On=Dark, Off=Well Lit
	' Bit 7     Not used
	Private myFlags As Byte
	' WallPaper PictureFile (if used)
	Private myWallpaper As String
	' Extended Encounter Settings (if used)
	Private myCombatGrid(bdCombatWidth, bdCombatHeight) As Byte
	
	' Chance to sucessfully flee (if allowed)
	Private myChanceToFlee As Byte
	
	' Belongs to Theme (0 is not assigned to any theme)
	Private myParentTheme As Short
	' Classificattion (inherits MapStyle from Theme or Map)
	' 0 - None, 1 - Safe Area, 2 - Dangerous Area, 3 - Inn/Pub, 4 - Merchant,
	' 5 - Meeting Area, 6 - Guard Area, 7 - Temple Area, 8 - Dwelling
	Private myClassification As Byte
	' ReGen flags are:
	' Bit 0     &H1     On=ReGen Creatures, Off=Do not regen them
	' Bit 1     &H2     On=ReGen Items, Off=Do not regen them
	' Bit 2     &H4     On=ReGen Triggers, Off=Do not regen them
	' Bit 3     &H8     On=ReGen Descriptions, Off=Do not regen them
	' Bit 4     &H10    On=ReGen Locked (do not change Map layout of Encounter), Off=You can rearange the shape
	' Bit 5     &H20    On=Not Active, Off=IsActive
	' Bit 6-7           Not used
	Private myReGenFlags As Byte
	
	' myTriggers is one or more Triggers in this Encounter
	Private myTriggers As New Collection
	' myCreatures is one or more Creatures in this Encounter. You are in combat
	' with them iff one or more live Creatures exist. You encounter them by
	' moving the party into any tile of the Encounter, although each Creature
	' has its own location within the Encounter.
	Private myCreatures As New Collection
	' myItems is one or more Items in this Encounter. They are found by
	' searching anywhere in the Encounter iff there are no live and
	' unfriendly Creatures.
	Private myItems As New Collection
	
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
	
	
	Public Property Index() As Short
		Get
			Index = myIndex
		End Get
		Set(ByVal Value As Short)
			myIndex = Value
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
	
	
	Public Property Name() As String
		Get
			Name = Trim(myName)
		End Get
		Set(ByVal Value As String)
			myName = Trim(Value)
		End Set
	End Property
	
	
	Public Property FirstEntry() As String
		Get
			FirstEntry = Trim(myFirstEntry)
		End Get
		Set(ByVal Value As String)
			myFirstEntry = Trim(Value)
		End Set
	End Property
	
	
	Public Property SecondEntry() As String
		Get
			SecondEntry = Trim(mySecondEntry)
		End Get
		Set(ByVal Value As String)
			mySecondEntry = Trim(Value)
		End Set
	End Property
	
	
	Public Property Wallpaper() As String
		Get
			Wallpaper = Trim(myWallpaper)
		End Get
		Set(ByVal Value As String)
			myWallpaper = Trim(Value)
		End Set
	End Property
	
	
	Public Property CombatGrid(ByVal X As Short, ByVal Y As Short) As Short
		Get
			CombatGrid = myCombatGrid(Least(Greatest(X, 0), bdCombatWidth), Least(Greatest(Y, 0), bdCombatHeight))
		End Get
		Set(ByVal Value As Short)
			' 0 = Open, 1 = Blocked, 2 = Party Starting, 3 = Monster Starting
			myCombatGrid(Least(Greatest(X, 0), bdCombatWidth), Least(Greatest(Y, 0), bdCombatHeight)) = Value
		End Set
	End Property
	
	
	Public Property UseHint() As Short
		Get
			UseHint = (myFlags And &H1) > 0
		End Get
		Set(ByVal Value As Short)
			If Value Then
				myFlags = myFlags Or &H1
			Else
				myFlags = myFlags And (Not &H1)
			End If
		End Set
	End Property
	
	
	Public Property HaveEntered() As Short
		Get
			HaveEntered = (myFlags And &H2) > 0
		End Get
		Set(ByVal Value As Short)
			If Value Then
				myFlags = myFlags Or &H2
			Else
				myFlags = myFlags And (Not &H2)
			End If
		End Set
	End Property
	
	
	Public Property CanTalk() As Short
		Get
			CanTalk = (myFlags And &H8) > 0
		End Get
		Set(ByVal Value As Short)
			If Value Then
				myFlags = myFlags Or &H8
			Else
				myFlags = myFlags And (Not &H8)
			End If
		End Set
	End Property
	
	
	Public Property CanFight() As Short
		Get
			CanFight = (myFlags And &H4) > 0
		End Get
		Set(ByVal Value As Short)
			If Value Then
				myFlags = myFlags Or &H4
			Else
				myFlags = myFlags And (Not &H4)
			End If
		End Set
	End Property
	
	
	Public Property CanFlee() As Short
		Get
			CanFlee = (myFlags And &H10) > 0
		End Get
		Set(ByVal Value As Short)
			If Value Then
				myFlags = myFlags Or &H10
			Else
				myFlags = myFlags And (Not &H10)
			End If
		End Set
	End Property
	
	
	Public Property IsDark() As Short
		Get
			IsDark = (myFlags And &H40) > 0
		End Get
		Set(ByVal Value As Short)
			If Value Then
				myFlags = myFlags Or &H40
			Else
				myFlags = myFlags And (Not &H40)
			End If
		End Set
	End Property
	
	Public ReadOnly Property EncounterPoints() As Short
		Get
			Dim CreatureX As Creature
			Dim ItemX As Item
			For	Each CreatureX In myCreatures
				EncounterPoints = EncounterPoints + CreatureX.EncounterPoints
			Next CreatureX
			For	Each ItemX In myItems
				EncounterPoints = EncounterPoints + ItemX.EncounterPoints
			Next ItemX
		End Get
	End Property
	
	
	Public Property GenerateUponEntry() As Short
		Get
			GenerateUponEntry = (myFlags And &H20) > 0
		End Get
		Set(ByVal Value As Short)
			If Value Then
				myFlags = myFlags Or &H20
			Else
				myFlags = myFlags And (Not &H20)
			End If
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
	
	
	Public Property ChanceToFlee() As Short
		Get
			ChanceToFlee = myChanceToFlee
		End Get
		Set(ByVal Value As Short)
			myChanceToFlee = Value
		End Set
	End Property
	
	
	Public Property ParentTheme() As Short
		Get
			ParentTheme = myParentTheme
		End Get
		Set(ByVal Value As Short)
			myParentTheme = Value
		End Set
	End Property
	
	
	Public Property Classification() As Short
		Get
			Classification = myClassification
		End Get
		Set(ByVal Value As Short)
			myClassification = Value
		End Set
	End Property
	
	
	Public Property ReGenFlags() As Short
		Get
			ReGenFlags = myReGenFlags
		End Get
		Set(ByVal Value As Short)
			myReGenFlags = Value
		End Set
	End Property
	
	
	Public Property ReGenCreatures() As Short
		Get
			ReGenCreatures = (myReGenFlags And &H1) > 0
		End Get
		Set(ByVal Value As Short)
			If Value Then
				myReGenFlags = myReGenFlags Or &H1
			Else
				myReGenFlags = myReGenFlags And (Not &H1)
			End If
		End Set
	End Property
	
	
	Public Property ReGenItems() As Short
		Get
			ReGenItems = (myReGenFlags And &H2) > 0
		End Get
		Set(ByVal Value As Short)
			If Value Then
				myReGenFlags = myReGenFlags Or &H2
			Else
				myReGenFlags = myReGenFlags And (Not &H2)
			End If
		End Set
	End Property
	
	
	Public Property ReGenTriggers() As Short
		Get
			ReGenTriggers = (myReGenFlags And &H4) > 0
		End Get
		Set(ByVal Value As Short)
			If Value Then
				myReGenFlags = myReGenFlags Or &H4
			Else
				myReGenFlags = myReGenFlags And (Not &H4)
			End If
		End Set
	End Property
	
	
	Public Property ReGenDescription() As Short
		Get
			ReGenDescription = (myReGenFlags And &H8) > 0
		End Get
		Set(ByVal Value As Short)
			If Value Then
				myReGenFlags = myReGenFlags Or &H8
			Else
				myReGenFlags = myReGenFlags And (Not &H8)
			End If
		End Set
	End Property
	
	
	Public Property ReGenLocked() As Short
		Get
			ReGenLocked = (myReGenFlags And &H10) > 0
		End Get
		Set(ByVal Value As Short)
			If Value Then
				myReGenFlags = myReGenFlags Or &H10
			Else
				myReGenFlags = myReGenFlags And (Not &H10)
			End If
		End Set
	End Property
	
	
	Public Property IsActive() As Short
		Get
			IsActive = (myReGenFlags And &H20) = 0
		End Get
		Set(ByVal Value As Short)
			If Value Then
				myReGenFlags = myReGenFlags And (Not &H20)
			Else
				myReGenFlags = myReGenFlags Or &H20
			End If
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
	
	Public Sub RemoveTrigger(ByRef DeleteKey As String)
		myTriggers.Remove(DeleteKey)
	End Sub
	
	Public Sub Copy(ByRef FromEncounter As Encounter)
		Dim c As Short
		Dim myTrigger As Trigger
		Dim myItem As Item
		Dim myCreature As Creature
		' Copy Encounter Name
		myName = FromEncounter.Name
		myFirstEntry = FromEncounter.FirstEntry
		mySecondEntry = FromEncounter.SecondEntry
		myWallpaper = FromEncounter.Wallpaper
		myFlags = FromEncounter.Flags
		myChanceToFlee = FromEncounter.ChanceToFlee
		myParentTheme = FromEncounter.ParentTheme
		myClassification = FromEncounter.Classification
		myReGenFlags = FromEncounter.ReGenFlags
		' Copy Triggers for Encounter
		myTriggers = New Collection
		For c = 1 To FromEncounter.Triggers.Count()
			myTrigger = Me.AddTrigger
			'UPGRADE_WARNING: Couldn't resolve default property of object FromEncounter.Triggers(). Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
			myTrigger.Copy(FromEncounter.Triggers.Item(c))
		Next c
		' Copy Creatures for Encounter
		myCreatures = New Collection
		For c = 1 To FromEncounter.Creatures.Count()
			myCreature = Me.AddCreature
			'UPGRADE_WARNING: Couldn't resolve default property of object FromEncounter.Creatures(). Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
			myCreature.Copy(FromEncounter.Creatures.Item(c))
		Next c
		' Copy Items for Encounter
		myItems = New Collection
		For c = 1 To FromEncounter.Items.Count()
			myItem = Me.AddItem
			'UPGRADE_WARNING: Couldn't resolve default property of object FromEncounter.Items(). Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
			myItem.Copy(FromEncounter.Items.Item(c))
		Next c
	End Sub
	
	Public Sub SaveToFile(ByRef ToFile As Short)
		Dim X, c, Y As Short
		Dim tmp As String
		' Save Encounter Name
		'UPGRADE_WARNING: Put was upgraded to FilePut and has a new behavior. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="9B7D5ADD-D8FE-4819-A36C-6DEDAF088CC7"'
		FilePut(ToFile, myVersion)
		'UPGRADE_WARNING: Put was upgraded to FilePut and has a new behavior. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="9B7D5ADD-D8FE-4819-A36C-6DEDAF088CC7"'
		FilePut(ToFile, Len(myName))
		'UPGRADE_WARNING: Put was upgraded to FilePut and has a new behavior. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="9B7D5ADD-D8FE-4819-A36C-6DEDAF088CC7"'
		FilePut(ToFile, myName)
		'UPGRADE_WARNING: Put was upgraded to FilePut and has a new behavior. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="9B7D5ADD-D8FE-4819-A36C-6DEDAF088CC7"'
		FilePut(ToFile, myIndex)
		'UPGRADE_WARNING: Put was upgraded to FilePut and has a new behavior. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="9B7D5ADD-D8FE-4819-A36C-6DEDAF088CC7"'
		FilePut(ToFile, Len(myFirstEntry))
		'UPGRADE_WARNING: Put was upgraded to FilePut and has a new behavior. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="9B7D5ADD-D8FE-4819-A36C-6DEDAF088CC7"'
		FilePut(ToFile, myFirstEntry)
		'UPGRADE_WARNING: Put was upgraded to FilePut and has a new behavior. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="9B7D5ADD-D8FE-4819-A36C-6DEDAF088CC7"'
		FilePut(ToFile, Len(mySecondEntry))
		'UPGRADE_WARNING: Put was upgraded to FilePut and has a new behavior. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="9B7D5ADD-D8FE-4819-A36C-6DEDAF088CC7"'
		FilePut(ToFile, mySecondEntry)
		If Len(myWallpaper) > 0 Then
			tmp = myWallpaper & "|"
			For X = 0 To bdCombatWidth : For Y = 0 To bdCombatHeight
					tmp = tmp & VB6.Format(myCombatGrid(X, Y))
				Next Y : Next X
			'UPGRADE_WARNING: Put was upgraded to FilePut and has a new behavior. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="9B7D5ADD-D8FE-4819-A36C-6DEDAF088CC7"'
			FilePut(ToFile, Len(tmp))
			'UPGRADE_WARNING: Put was upgraded to FilePut and has a new behavior. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="9B7D5ADD-D8FE-4819-A36C-6DEDAF088CC7"'
			FilePut(ToFile, tmp)
		Else
			'UPGRADE_WARNING: Put was upgraded to FilePut and has a new behavior. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="9B7D5ADD-D8FE-4819-A36C-6DEDAF088CC7"'
			FilePut(ToFile, Len(myWallpaper))
			'UPGRADE_WARNING: Put was upgraded to FilePut and has a new behavior. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="9B7D5ADD-D8FE-4819-A36C-6DEDAF088CC7"'
			FilePut(ToFile, myWallpaper)
		End If
		'UPGRADE_WARNING: Put was upgraded to FilePut and has a new behavior. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="9B7D5ADD-D8FE-4819-A36C-6DEDAF088CC7"'
		FilePut(ToFile, myFlags)
		'UPGRADE_WARNING: Put was upgraded to FilePut and has a new behavior. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="9B7D5ADD-D8FE-4819-A36C-6DEDAF088CC7"'
		FilePut(ToFile, myChanceToFlee)
		'UPGRADE_WARNING: Put was upgraded to FilePut and has a new behavior. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="9B7D5ADD-D8FE-4819-A36C-6DEDAF088CC7"'
		FilePut(ToFile, myParentTheme)
		'UPGRADE_WARNING: Put was upgraded to FilePut and has a new behavior. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="9B7D5ADD-D8FE-4819-A36C-6DEDAF088CC7"'
		FilePut(ToFile, myClassification)
		'UPGRADE_WARNING: Put was upgraded to FilePut and has a new behavior. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="9B7D5ADD-D8FE-4819-A36C-6DEDAF088CC7"'
		FilePut(ToFile, myReGenFlags)
		' Save Triggers for Encounter
		'UPGRADE_WARNING: Put was upgraded to FilePut and has a new behavior. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="9B7D5ADD-D8FE-4819-A36C-6DEDAF088CC7"'
		FilePut(ToFile, myTriggers.Count())
		For c = 1 To myTriggers.Count()
			'UPGRADE_WARNING: Couldn't resolve default property of object myTriggers().SaveToFile. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
			myTriggers.Item(c).SaveToFile(ToFile)
		Next c
		' Save Creatures for Encounter
		'UPGRADE_WARNING: Put was upgraded to FilePut and has a new behavior. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="9B7D5ADD-D8FE-4819-A36C-6DEDAF088CC7"'
		FilePut(ToFile, myCreatures.Count())
		For c = 1 To myCreatures.Count()
			'UPGRADE_WARNING: Couldn't resolve default property of object myCreatures().SaveToFile. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
			myCreatures.Item(c).SaveToFile(ToFile)
		Next c
		' Save Items for Encounter
		'UPGRADE_WARNING: Put was upgraded to FilePut and has a new behavior. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="9B7D5ADD-D8FE-4819-A36C-6DEDAF088CC7"'
		FilePut(ToFile, myItems.Count())
		For c = 1 To myItems.Count()
			'UPGRADE_WARNING: Couldn't resolve default property of object myItems().SaveToFile. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
			myItems.Item(c).SaveToFile(ToFile)
		Next c
	End Sub
	
	Public Sub ReadFromFile(ByRef FromFile As Short)
		On Error GoTo ErrorHandler
		Dim X, c, Y As Short
		Dim Cnt As Integer
		Dim myTrigger As Trigger
		Dim myCreature As Creature
		Dim tmp, Reading As String
		Dim myItem As Item
		' Read Encounter Name and Index
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
		'UPGRADE_WARNING: Get was upgraded to FileGet and has a new behavior. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="9B7D5ADD-D8FE-4819-A36C-6DEDAF088CC7"'
		FileGet(FromFile, myIndex)
		' Read Entry Text
		'UPGRADE_WARNING: Get was upgraded to FileGet and has a new behavior. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="9B7D5ADD-D8FE-4819-A36C-6DEDAF088CC7"'
		FileGet(FromFile, Cnt)
		myFirstEntry = ""
		For c = 1 To Cnt
			myFirstEntry = myFirstEntry & " "
		Next c
		'UPGRADE_WARNING: Get was upgraded to FileGet and has a new behavior. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="9B7D5ADD-D8FE-4819-A36C-6DEDAF088CC7"'
		FileGet(FromFile, myFirstEntry)
		' Read Second Entry Text (if any)
		'UPGRADE_WARNING: Get was upgraded to FileGet and has a new behavior. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="9B7D5ADD-D8FE-4819-A36C-6DEDAF088CC7"'
		FileGet(FromFile, Cnt)
		mySecondEntry = ""
		For c = 1 To Cnt
			mySecondEntry = mySecondEntry & " "
		Next c
		'UPGRADE_WARNING: Get was upgraded to FileGet and has a new behavior. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="9B7D5ADD-D8FE-4819-A36C-6DEDAF088CC7"'
		FileGet(FromFile, mySecondEntry)
		' Read Wallpaper PictureFile (if any)
		'UPGRADE_WARNING: Get was upgraded to FileGet and has a new behavior. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="9B7D5ADD-D8FE-4819-A36C-6DEDAF088CC7"'
		FileGet(FromFile, Cnt)
		myWallpaper = ""
		For c = 1 To Cnt
			myWallpaper = myWallpaper & " "
		Next c
		'UPGRADE_WARNING: Get was upgraded to FileGet and has a new behavior. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="9B7D5ADD-D8FE-4819-A36C-6DEDAF088CC7"'
		FileGet(FromFile, myWallpaper)
		' Check for extended Encounter combat settings
		tmp = modBD.BreakText(myWallpaper, 2)
		If Len(tmp) > 0 Then
			c = 1
			For X = 0 To bdCombatWidth : For Y = 0 To bdCombatHeight
					If IsNumeric(Mid(tmp, c, 1)) = True Then
						myCombatGrid(X, Y) = CByte(Mid(tmp, c, 1))
					Else
						myCombatGrid(X, Y) = 0
					End If
					c = c + 1
				Next Y : Next X
			myWallpaper = modBD.BreakText(myWallpaper, 1)
		End If
		'UPGRADE_WARNING: Get was upgraded to FileGet and has a new behavior. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="9B7D5ADD-D8FE-4819-A36C-6DEDAF088CC7"'
		FileGet(FromFile, myFlags)
		'UPGRADE_WARNING: Get was upgraded to FileGet and has a new behavior. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="9B7D5ADD-D8FE-4819-A36C-6DEDAF088CC7"'
		FileGet(FromFile, myChanceToFlee)
		'UPGRADE_WARNING: Get was upgraded to FileGet and has a new behavior. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="9B7D5ADD-D8FE-4819-A36C-6DEDAF088CC7"'
		FileGet(FromFile, myParentTheme)
		'UPGRADE_WARNING: Get was upgraded to FileGet and has a new behavior. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="9B7D5ADD-D8FE-4819-A36C-6DEDAF088CC7"'
		FileGet(FromFile, myClassification)
		'UPGRADE_WARNING: Get was upgraded to FileGet and has a new behavior. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="9B7D5ADD-D8FE-4819-A36C-6DEDAF088CC7"'
		FileGet(FromFile, myReGenFlags)
		' Read Triggers for Encounter
		Reading = "Trigger"
		'UPGRADE_WARNING: Get was upgraded to FileGet and has a new behavior. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="9B7D5ADD-D8FE-4819-A36C-6DEDAF088CC7"'
		FileGet(FromFile, Cnt)
		For c = 1 To Cnt
			myTrigger = New Trigger
			myTrigger.ReadFromFile(FromFile)
			myTriggers.Add(myTrigger, "T" & myTrigger.Index)
		Next c
		' Read Creatures for Encounter
		Reading = "Creature"
		'UPGRADE_WARNING: Get was upgraded to FileGet and has a new behavior. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="9B7D5ADD-D8FE-4819-A36C-6DEDAF088CC7"'
		FileGet(FromFile, Cnt)
		For c = 1 To Cnt
			myCreature = New Creature
			myCreature.ReadFromFile(FromFile)
			myCreatures.Add(myCreature, "X" & myCreature.Index)
		Next c
		' Read Items for Encounter
		Reading = "Item"
		'UPGRADE_WARNING: Get was upgraded to FileGet and has a new behavior. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="9B7D5ADD-D8FE-4819-A36C-6DEDAF088CC7"'
		FileGet(FromFile, Cnt)
		For c = 1 To Cnt
			myItem = New Item
			myItem.ReadFromFile(FromFile)
			myItems.Add(myItem, "I" & myItem.Index)
		Next c
		Exit Sub
ErrorHandler: 
		' [Titi 2.4.9] re-index components if necessary
		If Err.Number = 457 And Reading <> "" Then
			Select Case Reading
				Case "Trigger"
					oErr.logError(Reading & " index conflict: " & myTrigger.Index)
				Case "Creature"
					oErr.logError(Reading & " index conflict: " & myCreature.Index)
				Case "Item"
					oErr.logError(Reading & " index conflict: " & myItem.Index)
			End Select
		Else
			oErr.logError("Cannot read encounter" & IIf(Reading <> vbNullString, "'s " & Reading, "") & ", error#" & Err.Number & " (" & Err.Description & ")")
		End If
	End Sub
	
	'UPGRADE_NOTE: Class_Initialize was upgraded to Class_Initialize_Renamed. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="A9E4979A-37FA-4718-9994-97DD76ED70A7"'
	Private Sub Class_Initialize_Renamed()
		Me.CanFight = True
		Me.CanFlee = True
		Me.ChanceToFlee = 100
		Me.Name = "BlankEncounter"
	End Sub
	Public Sub New()
		MyBase.New()
		Class_Initialize_Renamed()
	End Sub
End Class