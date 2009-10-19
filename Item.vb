Option Strict Off
Option Explicit On
Friend Class Item
	
	' Version number for Item
	Private myVersion As Short
	' Unique index for this Item
	Private myIndex As Short
	' Name of this Item
	Private myName As String
	' Comments for this Item (can be used as Description)
	Private myComments As String
	' See SetUpItemFamily in modBD.
	Private myFamily As Byte
	'   &H1 Bit 1   - On=Use Comments as Description, Off=Don't
	'   &H2 Bit 2   - On=Soft Bulk (as Bulky as Bulkiest Item within), Off=Hard Bulk
	'   &H4 Bit 3   - On=Can Combine when match name, Off=Cannot combine
	'   Bit 4-8     - Key Type (32 varieties, used only with Key Family or Container Family items).
	Private myFlags As Byte
	' myCount is the number of Items this Item represents. Like Items
	' (with the EXACT same name) can be combined. myCount is normally
	' just 1.
	Private myCount As Short
	' Value in BP of this Item
	Private myValue As Short
	' myWeight is the weight of this Item NOT including Contents. To
	' get the "total" weight, you must add up the Items weight plus
	' myWeight. Characters can carry Strength * 15 in weight units, Strength 10 carry 150.
	Private myWeight As Short
	' myBulk is the bulk of this Item NOT including Contents. If flagged
	' as "soft capacity," the total bulk this Item represents is the
	' bulk of all items inside this one.
	' Characters can carry Agility * 8 in bulk units.
	Private myBulk As Short
	' Capacity (Total Capacity - 0 is no capacity, cannot be opened; 255 is infinite capacity)
	Private myCapacity As Short
	' Used only in UberWizard and while Dragging
	Private mySelected As Short
	'   Bit 1-2     - 0 - Short, 1 - Medium, 2- Long Range, 3 - Undefined
	'   Bit 3-4     - 0 - Melee, 1 - Ranged, 2 - Ammo, 3 - Thrown
	'   Bit 5-7     - Ammo Type - Arrow (Short), Arrow (Long), Sling (Light), Sling (Heavy),
	'                             Gun (Small), Gun (Large), Energy (Low), Energy (High)
	'   Bit 8       - On = IsSelected (UberWiz and Inventory)
	Private myShootType As Byte
	
	' Defines coordinate in Inventory screen. (0 = Not Placed, 1-255 = Inventory/Container)
	Private myInvSpot As Byte
	
	Private myNotUsed As Short
	
	' Weapon ActionPoints is how much ActionPoints it takes to use
	Private myActionPoints As Short
	
	' Run-time only
	Private myPic As Byte
	' Name of the image file for this item (in /item directory)
	Private myPictureFile As String
	' X and Y on Map for Encounter. Not used on Creatures.
	Private myMapX As Short
	Private myMapY As Short
	' A number from 0 to 4 indicating the spot on the Tile.
	Private myTileSpot As Byte
	
	' Bits 1-4     - Type of Resistance (Armor) or Damage (Weapons)
	'       0 - None, 1 - Sharp, 2 - Blunt, 3 - Cold, 4 - Fire, 5 - Evil
	'       6 - Holy, 7 - Magic, 8 - Mind
	' Bit 5-8       - Amount of Resistance (0-15)
	Private myResistanceType As Byte
	
	'   Bit 1-4    - WearType (Armor Only)
	'       0 - Body, 1 - Helm, 2 - Glove, 3 - Bracelet, 4 - Backpack, 5 - Shield,
	'       6 - Boots, 7 - Necklace, 8 - Belt, 9 - Ring, 10 - One Hand, 11 - Two Handed,
	'       12-15 - Undefined
	'   Bit 5-14    - Not Used
	'   Bit 15      - On = IsReady
	Private myWearType As Short
	' Amount of dice of damage will do if you in Combat
	Private myDamage As Short
	' Amount added to d20 when attempting to hit
	Private myAttackBonus As Short
	' Amount of Resistance provided by Item when wearing
	Private myResistance As Short
	' Amount of Defense this item provides if equipped in hand.
	Private myDefense As Short
	' Amount of additional damage from the Item
	Private myDamageBonus As Short
	
	' Items contained in this Item
	Private myItems As New Collection
	' Triggers attached to this Item
	Private myTriggers As New Collection
	
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
	
	Public Function AddItemBefore(ByRef BeforeKey As String) As Item
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
		myItems.Add(ItemX, "I" & ItemX.Index, BeforeKey)
		AddItem = ItemX
	End Function
	
	Public Sub RemoveItem(ByRef DeleteKey As String)
		myItems.Remove(DeleteKey)
	End Sub
	
	
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
			Dim i, intCopyNumber As Short
			If myCount > 1 Then
				' [Titi 2.4.8] put the copy number at the very end (after the final 's'!)
				For i = Len(myName) To 1 Step -1
					If Not IsNumeric(Mid(myName, i)) Then
						Exit For
					Else
						intCopyNumber = Val(Mid(myName, i))
					End If
				Next 
				myName = Left(myName, i)
				Name = VB6.Format(myCount) & " " & Trim(myName) & "s"
				If intCopyNumber > 0 Then
					myName = myName & Str(intCopyNumber)
					Name = Name & Str(intCopyNumber)
				End If
			Else
				Name = Trim(myName)
			End If
		End Get
		Set(ByVal Value As String)
			myName = Trim(Value)
		End Set
	End Property
	
	Public ReadOnly Property NameText() As String
		Get
			NameText = Trim(myName)
		End Get
	End Property
	
	
	Public Property PictureFile() As String
		Get
			PictureFile = Trim(myPictureFile)
		End Get
		Set(ByVal Value As String)
			myPictureFile = Trim(Value)
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
	
	Public ReadOnly Property Hint() As String
		Get
			If Me.WearType < 10 Then ' Potential Armor
				Hint = " * Protect " & Me.Resistance & "%"
			ElseIf Me.Damage > 0 Then  ' Potential Weapon
				Hint = " * Damage " & (Me.Damage - 1) Mod 5 + 1 & "d" & Int(((Me.Damage - 1) Mod 25) / 5) * 2 + 4
				If Me.Damage - 1 > 24 Then
					Hint = Hint & "+" & Int((Me.Damage - 1) / 25)
				End If
			End If
			Hint = Me.Name & " * Bulk " & Me.Bulk & "  Wgt " & Me.Weight & Hint
		End Get
	End Property
	
	
	Public Property Family() As Short
		Get
			Family = myFamily
		End Get
		Set(ByVal Value As Short)
			myFamily = Value
		End Set
	End Property
	
	
	Public Property UseDescription() As Short
		Get
			UseDescription = (myFlags And &H1) > 0
		End Get
		Set(ByVal Value As Short)
			If Value Then
				myFlags = myFlags Or &H1
			Else
				myFlags = myFlags And (Not &H1)
			End If
		End Set
	End Property
	
	
	Public Property CanCombine() As Short
		Get
			CanCombine = (myFlags And &H4) > 0
		End Get
		Set(ByVal Value As Short)
			If Value Then
				myFlags = myFlags Or &H4
			Else
				myFlags = myFlags And (Not &H4)
			End If
		End Set
	End Property
	
	
	Public Property IsSelected() As Short
		Get
			IsSelected = (myShootType And &H80) > 0
		End Get
		Set(ByVal Value As Short)
			If Value Then
				myShootType = myShootType Or &H80
			Else
				myShootType = myShootType And (Not &H80)
			End If
		End Set
	End Property
	
	
	Public Property SoftCapacity() As Short
		Get
			SoftCapacity = (myFlags And &H2) > 0
		End Get
		Set(ByVal Value As Short)
			If Value Then
				myFlags = myFlags Or &H2
			Else
				myFlags = myFlags And (Not &H2)
			End If
		End Set
	End Property
	
	
	Public Property Count() As Short
		Get
			Count = myCount
		End Get
		Set(ByVal Value As Short)
			myCount = Value
			If myCount < 1 Then
				myCount = 1
			End If
		End Set
	End Property
	
	
	Public Property Value() As Short
		Get
			Value = myValue
		End Get
		Set(ByVal Value As Short)
			myValue = Value
		End Set
	End Property
	
	Public ReadOnly Property EncounterPoints() As Short
		Get
			Dim TriggerX As Trigger
			EncounterPoints = myValue
			For	Each TriggerX In myTriggers
				EncounterPoints = EncounterPoints + 5 * TriggerX.Statements.Count()
			Next TriggerX
		End Get
	End Property
	
	Public ReadOnly Property WeightEmpty() As Short
		Get
			WeightEmpty = myWeight
		End Get
	End Property
	
	
	Public Property Weight() As Short
		Get
			Dim ItemX As Item
			Dim c As Short
			c = myWeight
			' If Item weighs something, then add what's inside as well
			If myWeight > 0 Then
				For	Each ItemX In myItems
					c = c + ItemX.Weight
				Next ItemX
			End If
			Weight = c
		End Get
		Set(ByVal Value As Short)
			myWeight = Value
		End Set
	End Property
	
	Public ReadOnly Property Full() As Short
		Get
			Dim ItemX As Item
			Full = 0
			For	Each ItemX In myItems
				Full = Full + ItemX.Bulk
			Next ItemX
		End Get
	End Property
	
	Public ReadOnly Property Money() As Integer
		Get
			Dim ItemX As Item
			Money = 0
			If Me.Family = 2 Then ' Money Family
				Money = Me.Value
			End If
			For	Each ItemX In Me.Items
				Money = Money + ItemX.Money
			Next ItemX
		End Get
	End Property
	
	
	Public Property Bulk() As Short
		Get
			Dim ItemX As Item
			' By default, bulk is the bulk of the item.
			Bulk = myBulk
			' If soft capacity, then Bulk is the biggest thing inside the item.
			If Me.SoftCapacity = True Then
				For	Each ItemX In myItems
					If ItemX.Bulk > Bulk Then
						Bulk = ItemX.Bulk
					End If
				Next ItemX
			End If
			If Bulk < 0 Then
				Bulk = 0
			End If
		End Get
		Set(ByVal Value As Short)
			myBulk = Value
		End Set
	End Property
	
	Public ReadOnly Property BulkEmpty() As Short
		Get
			BulkEmpty = myBulk
		End Get
	End Property
	
	
	Public Property Capacity() As Short
		Get
			Capacity = myCapacity
		End Get
		Set(ByVal Value As Short)
			myCapacity = Value
		End Set
	End Property
	
	
	Public Property Selected() As Short
		Get
			Selected = mySelected
		End Get
		Set(ByVal Value As Short)
			mySelected = Value
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
	
	
	Public Property InvSpot() As Short
		Get
			InvSpot = myInvSpot
		End Get
		Set(ByVal Value As Short)
			myInvSpot = Value
		End Set
	End Property
	
	
	Public Property TileSpot() As Short
		Get
			TileSpot = myTileSpot
		End Get
		Set(ByVal Value As Short)
			myTileSpot = Value
		End Set
	End Property
	
	Public ReadOnly Property TileSpotX() As Short
		Get
			Select Case myTileSpot
				Case 0 : TileSpotX = 32
				Case 1 : TileSpotX = 8
				Case 2 : TileSpotX = 32
				Case 3 : TileSpotX = 56
				Case 4 : TileSpotX = 32
			End Select
		End Get
	End Property
	
	Public ReadOnly Property TileSpotY() As Short
		Get
			Select Case myTileSpot
				Case 0 : TileSpotY = 40
				Case 1 : TileSpotY = 52
				Case 2 : TileSpotY = 52
				Case 3 : TileSpotY = 52
				Case 4 : TileSpotY = 64
			End Select
		End Get
	End Property
	
	
	Public Property Pic() As Short
		Get
			Pic = myPic
		End Get
		Set(ByVal Value As Short)
			If Value > 255 Then Value = 255
			If Value < 0 Then Value = 0
			myPic = Value
		End Set
	End Property
	
	
	Public Property Resistance() As Short
		Get
			Resistance = myResistance
		End Get
		Set(ByVal Value As Short)
			myResistance = Value
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
	
	
	Public Property KeyBit(ByVal Index As Short) As Short
		Get
			KeyBit = (myFlags And (2 ^ Index)) > 0
		End Get
		Set(ByVal Value As Short)
			If Value Then
				myFlags = myFlags Or (2 ^ Index)
			Else
				myFlags = myFlags And (Not (2 ^ Index))
			End If
		End Set
	End Property
	
	
	Public Property KeyBits() As Byte
		Get
			KeyBits = (myFlags And &HF8)
		End Get
		Set(ByVal Value As Byte)
			' Wipe out the old key
			myFlags = (myFlags And &H7)
			' Set the new key
			myFlags = (myFlags Or Value)
		End Set
	End Property
	
	
	Public Property ShootType() As Short
		Get
			ShootType = System.Math.Abs(CInt((myShootType And &H10) > 0)) + System.Math.Abs(CInt((myShootType And &H20) > 0)) * 2 + System.Math.Abs(CInt((myShootType And &H40) > 0)) * 4
		End Get
		Set(ByVal Value As Short)
			' Clear previous value
			myShootType = myShootType And (Not &H70)
			Select Case Value
				Case 1, 3, 5, 7
					' 001
					myShootType = myShootType Or &H10
			End Select
			Select Case Value
				Case 2, 3, 6, 7
					' 010
					myShootType = myShootType Or &H20
			End Select
			Select Case Value
				Case 4, 5, 6, 7
					' 100
					myShootType = myShootType Or &H40
			End Select
		End Set
	End Property
	
	
	Public Property RangeType() As Short
		Get
			RangeType = myShootType And &H3
			If RangeType > 2 Then
				RangeType = 2
			End If
		End Get
		Set(ByVal Value As Short)
			' Clear out previous value
			myShootType = myShootType And (Not &H3)
			Select Case Value
				Case 1
					myShootType = myShootType Or &H1
				Case 2
					myShootType = myShootType Or &H2
			End Select
		End Set
	End Property
	
	
	Public Property ShootTypes() As Short
		Get
			ShootTypes = myShootType
		End Get
		Set(ByVal Value As Short)
			myShootType = Value
		End Set
	End Property
	
	
	Public Property ResistanceBonus() As Short
		Get
			ResistanceBonus = System.Math.Abs(CInt((myResistanceType And &H10) > 0)) + System.Math.Abs(CInt((myResistanceType And &H20) > 0)) * 2 + System.Math.Abs(CInt((myResistanceType And &H40) > 0)) * 4 + System.Math.Abs(CInt((myResistanceType And &H80) > 0)) * 8
		End Get
		Set(ByVal Value As Short)
			' Clear current value
			myResistanceType = myResistanceType And (Not &HF0)
			Select Case Value
				Case 1, 3, 5, 7, 9, 11, 13, 15
					' 0001
					myResistanceType = myResistanceType Or &H10
			End Select
			Select Case Value
				Case 2, 3, 6, 7, 10, 11, 14, 15
					' 0010
					myResistanceType = myResistanceType Or &H20
			End Select
			Select Case Value
				Case 4, 5, 6, 7, 12, 13, 14, 15
					' 0100
					myResistanceType = myResistanceType Or &H40
			End Select
			Select Case Value
				Case 8, 9, 10, 11, 12, 13, 14, 15
					' 1000
					myResistanceType = myResistanceType Or &H80
			End Select
		End Set
	End Property
	
	Public ReadOnly Property MaxRange() As Short
		Get
			' This is the range in grid squares of this item
			Select Case (myShootType And &H3)
				Case 0 ' Short
					MaxRange = 1
				Case 1 ' Medium
					MaxRange = 2
				Case 2 ' Long
					MaxRange = 50
				Case Else
					MaxRange = 1
			End Select
		End Get
	End Property
	
	Public ReadOnly Property MinRange() As Short
		Get
			' This is the range in grid squares of this item
			Select Case (myShootType And &H3)
				Case 0 ' Short
					MinRange = 1
				Case 1 ' Medium
					MinRange = 1
				Case 2 ' Long
					MinRange = 3
				Case Else
					MinRange = 1
			End Select
		End Get
	End Property
	
	
	Public Property IsReady() As Short
		Get
			IsReady = (myWearType And &H4000) > 0
		End Get
		Set(ByVal Value As Short)
			If Value Then
				myWearType = myWearType Or &H4000
			Else
				myWearType = myWearType And (Not &H4000)
			End If
		End Set
	End Property
	
	
	Public Property IsMoney() As Short
		Get
			IsMoney = (myFamily = 2)
		End Get
		Set(ByVal Value As Short)
			If Value = True Then
				myFamily = 2
			ElseIf myFamily = 2 Then 
				myFamily = 0
			End If
		End Set
	End Property
	
	Public ReadOnly Property InHand() As Short
		Get
			InHand = ((Me.WearType = 10 Or Me.WearType = 11) And Me.IsReady = True)
		End Get
	End Property
	
	Public ReadOnly Property IsAmmo() As Short
		Get
			IsAmmo = (Me.Family = 6)
		End Get
	End Property
	
	Public ReadOnly Property IsShooter() As Short
		Get
			IsShooter = (Me.Family = 7)
		End Get
	End Property
	
	Public ReadOnly Property WearTypes() As Short
		Get
			WearTypes = myWearType
		End Get
	End Property
	
	
	Public Property WearType() As Short
		Get
			If myFamily > 3 And myFamily < 9 Then ' Armor or Weapon
				WearType = myWearType And &HF
				If WearType > 11 Then
					WearType = 10
				End If
			Else
				WearType = 10
			End If
		End Get
		Set(ByVal Value As Short)
			' Clear current WearType
			myWearType = myWearType And (Not &HF)
			' Set new WearType
			myWearType = (myWearType Or (Value And &HF))
		End Set
	End Property
	
	
	Public Property Damage() As Short
		Get
			Damage = myDamage
		End Get
		Set(ByVal Value As Short)
			myDamage = Value
		End Set
	End Property
	
	
	Public Property AttackBonus() As Short
		Get
			AttackBonus = myAttackBonus
		End Get
		Set(ByVal Value As Short)
			myAttackBonus = Value
		End Set
	End Property
	
	
	Public Property DamageBonus() As Short
		Get
			DamageBonus = myDamageBonus
		End Get
		Set(ByVal Value As Short)
			myDamageBonus = Value
		End Set
	End Property
	
	
	Public Property DamageType() As Short
		Get
			If myFamily > 4 And myFamily < 9 Then
				DamageType = myResistanceType And &HF
			Else
				DamageType = 0
			End If
		End Get
		Set(ByVal Value As Short)
			' Clear current value
			myResistanceType = myResistanceType And (Not &HF)
			' Set new value
			If myFamily > 4 And myFamily < 9 Then
				myResistanceType = myResistanceType Or (Value And &HF)
			End If
		End Set
	End Property
	
	
	Public Property ResistanceTypeReal() As Byte
		Get
			ResistanceTypeReal = myResistanceType
		End Get
		Set(ByVal Value As Byte)
			myResistanceType = Value
		End Set
	End Property
	
	
	Public Property ResistanceType() As Short
		Get
			If myFamily = 4 Then
				ResistanceType = myResistanceType And &HF
			Else
				ResistanceType = 0
			End If
		End Get
		Set(ByVal Value As Short)
			' Clear current value
			myResistanceType = myResistanceType And (Not &HF)
			' Set new value
			If myFamily = 4 Then
				myResistanceType = myResistanceType Or (Value And &HF)
			End If
		End Set
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
	
	
	Public Property Defense() As Short
		Get
			Defense = myDefense
		End Get
		Set(ByVal Value As Short)
			myDefense = Value
		End Set
	End Property
	
	
	Public Property ActionPoints() As Short
		Get
			If myActionPoints < 0 Then
				ActionPoints = 0
			Else
				ActionPoints = myActionPoints
			End If
		End Get
		Set(ByVal Value As Short)
			myActionPoints = Value
		End Set
	End Property
	
	Public Function RemoveMoney(ByRef AmountToRemove As Short) As Short
		Dim Take As Double
		Dim ItemX As Item
		' Remove from Contents
		For	Each ItemX In Me.Items
			If ItemX.RemoveMoney(AmountToRemove) = True Then
				Me.RemoveItem("I" & ItemX.Index)
			End If
		Next ItemX
		' Remove Me.Value if asking for more
		RemoveMoney = False
		If Me.IsMoney = True And AmountToRemove > 0 Then
			If myValue > 0 Then
				Take = Int(AmountToRemove / (myValue / myCount))
				If myBulk > 0 Then
					myBulk = myBulk - Take * Int(myBulk / myCount)
				End If
				If myWeight > 0 Then
					myWeight = myWeight - Take * Int(myWeight / myCount)
				End If
				If myCount > 0 Then
					myCount = myCount - Take
				End If
				myValue = myValue - AmountToRemove
				AmountToRemove = AmountToRemove - myValue
			End If
			If myValue < 1 Then
				RemoveMoney = True
			End If
		End If
	End Function
	
	Public Sub Copy(ByRef FromItem As Item)
		Dim c As Short
		Dim myTrigger As Trigger
		Dim myItem As Item
		' Copy Item Name,and Comments
		myName = FromItem.NameText
		myComments = FromItem.Comments
		' Copy Attributes
		myFamily = FromItem.Family
		myFlags = FromItem.Flags
		myCount = FromItem.Count
		myValue = FromItem.Value
		myWeight = FromItem.WeightEmpty
		myBulk = FromItem.BulkEmpty
		myCapacity = FromItem.Capacity
		mySelected = FromItem.Selected
		myResistanceType = FromItem.ResistanceTypeReal
		myResistance = FromItem.Resistance
		myDefense = FromItem.Defense
		myShootType = FromItem.ShootTypes
		myActionPoints = FromItem.ActionPoints
		myWearType = FromItem.WearTypes
		myDamage = FromItem.Damage
		myAttackBonus = FromItem.AttackBonus
		myDamageBonus = FromItem.DamageBonus
		myDefense = FromItem.Defense
		myPictureFile = FromItem.PictureFile
		myPic = FromItem.Pic
		' Copy Triggers for Item
		myTriggers = New Collection
		For c = 1 To FromItem.Triggers.Count()
			myTrigger = Me.AddTrigger
			'UPGRADE_WARNING: Couldn't resolve default property of object FromItem.Triggers(). Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
			myTrigger.Copy(FromItem.Triggers.Item(c))
		Next c
		' Copy Items for Item
		myItems = New Collection
		For c = 1 To FromItem.Items.Count()
			myItem = Me.AddItem
			'UPGRADE_WARNING: Couldn't resolve default property of object FromItem.Items(). Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
			myItem.Copy(FromItem.Items.Item(c))
		Next c
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
	End Sub
	
	Public Sub ReadFromFile(ByRef FromFile As Short)
		On Error GoTo ErrorHandler
		Dim c As Short
		Dim Cnt As Integer
		Dim myTrigger As Trigger
		Dim myItem As Item
		Dim Reading As String
		ReadFromFileHeader(FromFile)
		' Read Attributes
		'UPGRADE_WARNING: Get was upgraded to FileGet and has a new behavior. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="9B7D5ADD-D8FE-4819-A36C-6DEDAF088CC7"'
		FileGet(FromFile, myFamily)
		'UPGRADE_WARNING: Get was upgraded to FileGet and has a new behavior. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="9B7D5ADD-D8FE-4819-A36C-6DEDAF088CC7"'
		FileGet(FromFile, myFlags)
		'UPGRADE_WARNING: Get was upgraded to FileGet and has a new behavior. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="9B7D5ADD-D8FE-4819-A36C-6DEDAF088CC7"'
		FileGet(FromFile, myCount)
		'UPGRADE_WARNING: Get was upgraded to FileGet and has a new behavior. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="9B7D5ADD-D8FE-4819-A36C-6DEDAF088CC7"'
		FileGet(FromFile, myValue)
		'UPGRADE_WARNING: Get was upgraded to FileGet and has a new behavior. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="9B7D5ADD-D8FE-4819-A36C-6DEDAF088CC7"'
		FileGet(FromFile, myWeight)
		'UPGRADE_WARNING: Get was upgraded to FileGet and has a new behavior. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="9B7D5ADD-D8FE-4819-A36C-6DEDAF088CC7"'
		FileGet(FromFile, myBulk)
		'UPGRADE_WARNING: Get was upgraded to FileGet and has a new behavior. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="9B7D5ADD-D8FE-4819-A36C-6DEDAF088CC7"'
		FileGet(FromFile, myCapacity)
		'UPGRADE_WARNING: Get was upgraded to FileGet and has a new behavior. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="9B7D5ADD-D8FE-4819-A36C-6DEDAF088CC7"'
		FileGet(FromFile, myDamageBonus)
		'UPGRADE_WARNING: Get was upgraded to FileGet and has a new behavior. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="9B7D5ADD-D8FE-4819-A36C-6DEDAF088CC7"'
		FileGet(FromFile, myNotUsed)
		'UPGRADE_WARNING: Get was upgraded to FileGet and has a new behavior. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="9B7D5ADD-D8FE-4819-A36C-6DEDAF088CC7"'
		FileGet(FromFile, myShootType)
		'UPGRADE_WARNING: Get was upgraded to FileGet and has a new behavior. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="9B7D5ADD-D8FE-4819-A36C-6DEDAF088CC7"'
		FileGet(FromFile, myResistanceType)
		'UPGRADE_WARNING: Get was upgraded to FileGet and has a new behavior. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="9B7D5ADD-D8FE-4819-A36C-6DEDAF088CC7"'
		FileGet(FromFile, myInvSpot)
		'UPGRADE_WARNING: Get was upgraded to FileGet and has a new behavior. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="9B7D5ADD-D8FE-4819-A36C-6DEDAF088CC7"'
		FileGet(FromFile, myWearType)
		'UPGRADE_WARNING: Get was upgraded to FileGet and has a new behavior. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="9B7D5ADD-D8FE-4819-A36C-6DEDAF088CC7"'
		FileGet(FromFile, myDamage)
		'UPGRADE_WARNING: Get was upgraded to FileGet and has a new behavior. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="9B7D5ADD-D8FE-4819-A36C-6DEDAF088CC7"'
		FileGet(FromFile, myAttackBonus)
		'UPGRADE_WARNING: Get was upgraded to FileGet and has a new behavior. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="9B7D5ADD-D8FE-4819-A36C-6DEDAF088CC7"'
		FileGet(FromFile, myDefense)
		'UPGRADE_WARNING: Get was upgraded to FileGet and has a new behavior. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="9B7D5ADD-D8FE-4819-A36C-6DEDAF088CC7"'
		FileGet(FromFile, myMapX)
		'UPGRADE_WARNING: Get was upgraded to FileGet and has a new behavior. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="9B7D5ADD-D8FE-4819-A36C-6DEDAF088CC7"'
		FileGet(FromFile, myMapY)
		'UPGRADE_WARNING: Get was upgraded to FileGet and has a new behavior. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="9B7D5ADD-D8FE-4819-A36C-6DEDAF088CC7"'
		FileGet(FromFile, myActionPoints)
		'UPGRADE_WARNING: Get was upgraded to FileGet and has a new behavior. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="9B7D5ADD-D8FE-4819-A36C-6DEDAF088CC7"'
		FileGet(FromFile, myTileSpot)
		'UPGRADE_WARNING: Get was upgraded to FileGet and has a new behavior. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="9B7D5ADD-D8FE-4819-A36C-6DEDAF088CC7"'
		FileGet(FromFile, myResistance)
		'UPGRADE_WARNING: Get was upgraded to FileGet and has a new behavior. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="9B7D5ADD-D8FE-4819-A36C-6DEDAF088CC7"'
		FileGet(FromFile, myDefense)
		'UPGRADE_WARNING: Get was upgraded to FileGet and has a new behavior. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="9B7D5ADD-D8FE-4819-A36C-6DEDAF088CC7"'
		FileGet(FromFile, Cnt)
		myPictureFile = ""
		For c = 1 To Cnt
			myPictureFile = myPictureFile & " "
		Next c
		'UPGRADE_WARNING: Get was upgraded to FileGet and has a new behavior. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="9B7D5ADD-D8FE-4819-A36C-6DEDAF088CC7"'
		FileGet(FromFile, myPictureFile)
		' Read Triggers for Item
		'UPGRADE_WARNING: Get was upgraded to FileGet and has a new behavior. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="9B7D5ADD-D8FE-4819-A36C-6DEDAF088CC7"'
		FileGet(FromFile, Cnt)
		Reading = "Trigger"
		For c = 1 To Cnt
			myTrigger = New Trigger
			myTrigger.ReadFromFile(FromFile)
			myTriggers.Add(myTrigger, "T" & myTrigger.Index)
		Next c
		' Read Items for Item
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
		' [Titi 2.4.9]
		If Err.Number = 457 And Reading <> "" Then
			Select Case Reading
				Case "Trigger"
					oErr.logError(Reading & " index conflict: " & myTrigger.Index)
				Case "Item"
					oErr.logError(Reading & " index conflict: " & myItem.Index)
			End Select
		Else
			oErr.logError("Cannot read item" & IIf(Reading <> vbNullString, "'s " & Reading, "") & ", error#" & Err.Number & " (" & Err.Description & ")")
		End If
	End Sub
	
	Public Sub SaveToFile(ByRef ToFile As Short)
		Dim c As Short
		' Save Item Name, Index and Comments
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
		' Save Attributes
		'UPGRADE_WARNING: Put was upgraded to FilePut and has a new behavior. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="9B7D5ADD-D8FE-4819-A36C-6DEDAF088CC7"'
		FilePut(ToFile, myFamily)
		'UPGRADE_WARNING: Put was upgraded to FilePut and has a new behavior. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="9B7D5ADD-D8FE-4819-A36C-6DEDAF088CC7"'
		FilePut(ToFile, myFlags)
		'UPGRADE_WARNING: Put was upgraded to FilePut and has a new behavior. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="9B7D5ADD-D8FE-4819-A36C-6DEDAF088CC7"'
		FilePut(ToFile, myCount)
		'UPGRADE_WARNING: Put was upgraded to FilePut and has a new behavior. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="9B7D5ADD-D8FE-4819-A36C-6DEDAF088CC7"'
		FilePut(ToFile, myValue)
		'UPGRADE_WARNING: Put was upgraded to FilePut and has a new behavior. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="9B7D5ADD-D8FE-4819-A36C-6DEDAF088CC7"'
		FilePut(ToFile, myWeight)
		'UPGRADE_WARNING: Put was upgraded to FilePut and has a new behavior. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="9B7D5ADD-D8FE-4819-A36C-6DEDAF088CC7"'
		FilePut(ToFile, myBulk)
		'UPGRADE_WARNING: Put was upgraded to FilePut and has a new behavior. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="9B7D5ADD-D8FE-4819-A36C-6DEDAF088CC7"'
		FilePut(ToFile, myCapacity)
		'UPGRADE_WARNING: Put was upgraded to FilePut and has a new behavior. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="9B7D5ADD-D8FE-4819-A36C-6DEDAF088CC7"'
		FilePut(ToFile, myDamageBonus)
		'UPGRADE_WARNING: Put was upgraded to FilePut and has a new behavior. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="9B7D5ADD-D8FE-4819-A36C-6DEDAF088CC7"'
		FilePut(ToFile, myNotUsed)
		'UPGRADE_WARNING: Put was upgraded to FilePut and has a new behavior. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="9B7D5ADD-D8FE-4819-A36C-6DEDAF088CC7"'
		FilePut(ToFile, myShootType)
		'UPGRADE_WARNING: Put was upgraded to FilePut and has a new behavior. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="9B7D5ADD-D8FE-4819-A36C-6DEDAF088CC7"'
		FilePut(ToFile, myResistanceType)
		'UPGRADE_WARNING: Put was upgraded to FilePut and has a new behavior. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="9B7D5ADD-D8FE-4819-A36C-6DEDAF088CC7"'
		FilePut(ToFile, myInvSpot)
		'UPGRADE_WARNING: Put was upgraded to FilePut and has a new behavior. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="9B7D5ADD-D8FE-4819-A36C-6DEDAF088CC7"'
		FilePut(ToFile, myWearType)
		'UPGRADE_WARNING: Put was upgraded to FilePut and has a new behavior. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="9B7D5ADD-D8FE-4819-A36C-6DEDAF088CC7"'
		FilePut(ToFile, myDamage)
		'UPGRADE_WARNING: Put was upgraded to FilePut and has a new behavior. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="9B7D5ADD-D8FE-4819-A36C-6DEDAF088CC7"'
		FilePut(ToFile, myAttackBonus)
		'UPGRADE_WARNING: Put was upgraded to FilePut and has a new behavior. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="9B7D5ADD-D8FE-4819-A36C-6DEDAF088CC7"'
		FilePut(ToFile, myDefense)
		'UPGRADE_WARNING: Put was upgraded to FilePut and has a new behavior. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="9B7D5ADD-D8FE-4819-A36C-6DEDAF088CC7"'
		FilePut(ToFile, myMapX)
		'UPGRADE_WARNING: Put was upgraded to FilePut and has a new behavior. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="9B7D5ADD-D8FE-4819-A36C-6DEDAF088CC7"'
		FilePut(ToFile, myMapY)
		'UPGRADE_WARNING: Put was upgraded to FilePut and has a new behavior. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="9B7D5ADD-D8FE-4819-A36C-6DEDAF088CC7"'
		FilePut(ToFile, myActionPoints)
		'UPGRADE_WARNING: Put was upgraded to FilePut and has a new behavior. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="9B7D5ADD-D8FE-4819-A36C-6DEDAF088CC7"'
		FilePut(ToFile, myTileSpot)
		'UPGRADE_WARNING: Put was upgraded to FilePut and has a new behavior. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="9B7D5ADD-D8FE-4819-A36C-6DEDAF088CC7"'
		FilePut(ToFile, myResistance)
		'UPGRADE_WARNING: Put was upgraded to FilePut and has a new behavior. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="9B7D5ADD-D8FE-4819-A36C-6DEDAF088CC7"'
		FilePut(ToFile, myDefense)
		'UPGRADE_WARNING: Put was upgraded to FilePut and has a new behavior. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="9B7D5ADD-D8FE-4819-A36C-6DEDAF088CC7"'
		FilePut(ToFile, Len(myPictureFile))
		'UPGRADE_WARNING: Put was upgraded to FilePut and has a new behavior. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="9B7D5ADD-D8FE-4819-A36C-6DEDAF088CC7"'
		FilePut(ToFile, myPictureFile)
		' Save Triggers for Item
		'UPGRADE_WARNING: Put was upgraded to FilePut and has a new behavior. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="9B7D5ADD-D8FE-4819-A36C-6DEDAF088CC7"'
		FilePut(ToFile, myTriggers.Count())
		For c = 1 To myTriggers.Count()
			'UPGRADE_WARNING: Couldn't resolve default property of object myTriggers().SaveToFile. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
			myTriggers.Item(c).SaveToFile(ToFile)
		Next c
		' Save Items for Item
		'UPGRADE_WARNING: Put was upgraded to FilePut and has a new behavior. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="9B7D5ADD-D8FE-4819-A36C-6DEDAF088CC7"'
		FilePut(ToFile, myItems.Count())
		For c = 1 To myItems.Count()
			'UPGRADE_WARNING: Couldn't resolve default property of object myItems().SaveToFile. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
			myItems.Item(c).SaveToFile(ToFile)
		Next c
	End Sub
	
	'UPGRADE_NOTE: Class_Initialize was upgraded to Class_Initialize_Renamed. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="A9E4979A-37FA-4718-9994-97DD76ED70A7"'
	Private Sub Class_Initialize_Renamed()
		myName = "Item"
		myCount = 1
		myActionPoints = 10
	End Sub
	Public Sub New()
		MyBase.New()
		Class_Initialize_Renamed()
	End Sub
End Class