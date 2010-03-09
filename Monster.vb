Option Strict Off
Option Explicit On
Friend Class Creature
	
	' Version of Class
	Private myVersion As Short
	' Unique index for this Creature
	Private myIndex As Short
	' Name of this Creature
	Private myName As String
	' Comments for this Creature (can be used as Description)
	Private myComments As String
	' Race string for Creature (separated by anything)
	Private myRace As String
	' [Titi 2.4.9] home tome
	Private myHome As String
	
	' Flags for small checks
	'   Bit0    On=Agressive (Can't Ignore when Encounter)
	'   Bit1    On=Friendly
	'   Bit2    On=DMControlled
	'   Bit3    On=Guard (Can't Take Treasture)
	'   Bit4    On=HadTurn
	'   Bit5    On=ScarletLetter
	'   Bit6    On=Male, Off=Female
	'   Bit7    On=Required in Tome
	Private myFlags As Byte
	
	' Current status of Creature
	'   Bit0    On=IsFrozen
	'   Bit1-2  CombatRank 0=Random, 1=Back, 2=Middle, 3=Front
	'   Bit3    On=IsSelected (UberWizard use only)
	'   Bit4    On=IsInanimate
	'   Bit5    Not Used
	'   Bit6    On=IsAfraid (Morale Failure/Flees Combat)
	'   Bit7    On=IsUnconscious
	Private myStatus As Byte
	
	' Current Eternium of Creature (only applicable for Spell Casters)
	Private myEternium As Byte
	
	' File name of Picture File (BMP) in /creature directory PLUS Portrait File Name (if applicable)
	Private myPictureFile As String
	' Top and Left of Face
	Private myFaceTop As Short
	Private myFaceLeft As Short
	' Pointer to PictureBox (used only at RUN time, not saved)
	Private myPic As Short
	
	' X and Y on Map for Encounter (where this Creature is NOW)
	Private myMapX As Short
	Private myMapY As Short
	' Number from 0 to 4 for spot on tile (where this Creature is NOW)
	Private myTileSpot As Byte
	' X and Y on Grid for Combat (again, where at NOW)
	Private myCol As Short
	Private myRow As Short
	' This is run-time only data
	Private myMoveToX As Short
	Private myMoveToY As Short
	' Que of Steps from A* routine
	' Run-time only data
	Private myStepX() As Short
	Private myStepY() As Short
	Private myMaxStep As Short
	Private myNextStep As Short
	Private mySteps As Short
	' Run-time only: current qued runes
	Private myRunes(5) As Short
	Private myRuneTop As Short
	Private myRuneQueLimit As Short
	' Run-time only: Initiative of Creature
	Private myInitiative As Short
	' Run-time only: Facing of Creature Picture in Combat
	Private myFacing As Short
	' Run-time only: current Action Points of Creature
	Private myActionPoints As Short
	' Run-time only: Bonus to General Resistance (0) and ResistanceTypes (1-8).
	Private myResistanceBonus(8) As Short
	' Run-Time only: CreatureTarget (for me)
	Private myCreatureTarget As Creature
	' Run-Time only: Msg Que
	Private myMsgQue(7) As String
	Private myMsgQueTop As Short
	
	'   0 - Wing, 1 - Tail, 2 - Body, 3 - Head, 4 - Arm, 5 - Leg, 6 - Antenna
	'   7 - Tentacle, 8 - Abdomen, 9 - Back, 10 - Neck
	Private myBodyType(7) As Byte
	' Armor Rank for each Armor chit (Percent Resistance)
	Private myResistance(7) As Short
	
	' Resistance Type Bonus Values (Range from 0 to 100% by 10% increments)
	Private myResistance1 As Byte
	Private myResistance2 As Byte
	Private myResistance3 As Byte
	Private myResistance4 As Byte
	
	' Default CombatAction (Right MouseButton)
	Private myCombatAction1 As Short
	Private myCombatAction2 As Short
	
	Private myNotUsed2 As Byte
	
	' Creature stats
	Private myStrength As Byte
	Private myWill As Byte
	Private myAgility As Byte
	Private myExperiencePoints As Integer
	Private mySkillPoints As Short
	' Run-time only: Creature Bonus Stats
	Private myStrengthBonus As Short
	Private myWillBonus As Short
	Private myAgilityBonus As Short
	Private myAttackBonus As Short
	Private myDamageBonus As Short
	Private myDefenseBonus As Short
	Private myMovementCostBonus As Short
	' Run-time only: Current Combat Attitude
	'   Bit0    On=Allowed Opportunity Attack, Off=Not Allowed
	'   Bit1-4  Not Used
	'   Bit5-7  Current Combat Attitude 0 thru 15
	Private myCombatAttitude As Byte
	
	Private myTempSkillSet As Short
	Private myTempSkillSpend As Short
	
	Private myLevel As Byte
	Private myHPNow As Short
	Private myHPMax As Short
	Private myLunacy As Byte
	Private myRevelry As Byte
	Private myWrath As Byte
	Private myPride As Byte
	Private myGreed As Byte
	Private myLust As Byte
	
	' More Flags
	'   Bit1    On=MoveWAV one time, Off=Many
	'   Bit2    On=DieWAV one time, Off=Many
	'   Bit3    On=AttackWAV one time, Off=Many
	'   Bit4    On=HitWAV one time, Off=Many
	'   Bit5    On=On Adventure, Off=Available
	'   Bit6-8  Not Used
	Private myFlagsWAV As Byte
	' Sound files
	Private myMoveWAV As String
	Private myDieWAV As String
	Private myAttackWAV As String
	Private myHitWAV As String
	
	' Family type
	'Private Const bdSentient = &H1
	'Private Const bdReptile = &H2
	'Private Const bdInsect = &H4
	'Private Const bdBlob = &H8
	'Private Const bdAnimal = &HF
	'Private Const bdVeggie = &H10
	'Private Const bdBird = &H20
	'Private Const bdUndead = &H40
	'Private Const bdAquatic = &H80
	'Private Const bdMagical = &HF0
	Private myFamily As Short
	
	' Size of Creature
	' Bit 0     -   On = Tiny
	' Bit 1     -   On = Small
	' Bit 2     -   On = Medium
	' Bit 3     -   On = Large
	' Bit 4     -   On = Huge
	' Bit 5-7   -   Not Used
	Private mySize As Byte
	
	' Items held by this Creature
	Private myItems As New Collection
	' Triggers attached to this Creature
	Private myTriggers As New Collection
	' Conversations attached to this Creature
	Private myConversations As New Collection
	' Current Conversation Index for the Creature (0 is nothing current)
	Private myCurrentConvo As Short
	
	
	Public Property Index() As Short
		Get
			Index = myIndex
		End Get
		Set(ByVal Value As Short)
			myIndex = Value
		End Set
	End Property
	
	
	Public Property Initiative() As Short
		Get
			Initiative = myInitiative
		End Get
		Set(ByVal Value As Short)
			myInitiative = Value
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
	
	
	Public Property Home() As String
		Get
			Home = myHome
		End Get
		Set(ByVal Value As String)
			' NewValue contains the name of the tome folder followed by the file name of the home tome
			Dim StartAt As Short
			' \ is used as carriage return. Change it to /
			StartAt = InStr(Value, "\")
			Do Until StartAt = 0
				Value = Left(Value, StartAt - 1) & "/" & Mid(Value, StartAt + 1)
				StartAt = InStr(Value, "\")
			Loop 
			myHome = Value
		End Set
	End Property
	
	
	Public Property Race() As String
		Get
			Race = Trim(myRace)
		End Get
		Set(ByVal Value As String)
			myRace = Trim(Value)
		End Set
	End Property
	
	
	Public Property MoveWAV() As String
		Get
			MoveWAV = Trim(myMoveWAV)
		End Get
		Set(ByVal Value As String)
			myMoveWAV = Trim(Value)
		End Set
	End Property
	
	
	Public Property DieWAV() As String
		Get
			DieWAV = Trim(myDieWAV)
		End Get
		Set(ByVal Value As String)
			myDieWAV = Trim(Value)
		End Set
	End Property
	
	
	Public Property HitWAV() As String
		Get
			HitWAV = Trim(myHitWAV)
		End Get
		Set(ByVal Value As String)
			myHitWAV = Trim(Value)
		End Set
	End Property
	
	
	Public Property AttackWAV() As String
		Get
			AttackWAV = Trim(myAttackWAV)
		End Get
		Set(ByVal Value As String)
			myAttackWAV = Trim(Value)
		End Set
	End Property
	
	
	Public Property MoveWAVOneTime() As Short
		Get
			MoveWAVOneTime = (myFlagsWAV And &H1) > 0
		End Get
		Set(ByVal Value As Short)
			If Value Then
				myFlagsWAV = myFlagsWAV Or &H1
			Else
				myFlagsWAV = myFlagsWAV And (Not &H1)
			End If
		End Set
	End Property
	
	
	Public Property DieWAVOneTime() As Short
		Get
			DieWAVOneTime = (myFlagsWAV And &H2) > 0
		End Get
		Set(ByVal Value As Short)
			If Value Then
				myFlagsWAV = myFlagsWAV Or &H2
			Else
				myFlagsWAV = myFlagsWAV And (Not &H2)
			End If
		End Set
	End Property
	
	
	Public Property AttackWAVOneTime() As Short
		Get
			AttackWAVOneTime = (myFlagsWAV And &H4) > 0
		End Get
		Set(ByVal Value As Short)
			If Value Then
				myFlagsWAV = myFlagsWAV Or &H4
			Else
				myFlagsWAV = myFlagsWAV And (Not &H4)
			End If
		End Set
	End Property
	
	
	Public Property HitWAVOneTime() As Short
		Get
			HitWAVOneTime = (myFlagsWAV And &H8) > 0
		End Get
		Set(ByVal Value As Short)
			If Value Then
				myFlagsWAV = myFlagsWAV Or &H8
			Else
				myFlagsWAV = myFlagsWAV And (Not &H8)
			End If
		End Set
	End Property
	
	
	Public Property CurrentConvo() As Short
		Get
			CurrentConvo = myCurrentConvo
		End Get
		Set(ByVal Value As Short)
			myCurrentConvo = Value
		End Set
	End Property
	
	
	Public Property MsgQueTop() As Short
		Get
			MsgQueTop = myMsgQueTop
		End Get
		Set(ByVal Value As Short)
			myMsgQueTop = Value
		End Set
	End Property
	
	
	Public Property MsgQue(ByVal Index As Short) As String
		Get
			MsgQue = myMsgQue(Index)
		End Get
		Set(ByVal Value As String)
			myMsgQue(Index) = Value
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
	
	Public ReadOnly Property ItemMaxRange() As Short
		Get
			ItemMaxRange = ItemInHand.MaxRange
		End Get
	End Property
	
	Public ReadOnly Property ItemMinRange() As Short
		Get
			ItemMinRange = ItemInHand.MinRange
		End Get
	End Property
	
	Public ReadOnly Property ItemAmmo() As Item
		Get
            Dim ItemX, ItemH, ItemZ As Item
			ItemH = Me.ItemInHand
			If ItemH.IsShooter = True Then
				' Search through equipped first
				For	Each ItemX In myItems
					If ItemX.IsReady = True Then
						' See if item itself is ammo
						If ItemX.IsAmmo = True And ItemX.ShootType = ItemH.ShootType Then
							ItemAmmo = ItemX
							Exit Property
						Else
							' Search inside the item for ammo
							For	Each ItemZ In ItemX.Items
								If ItemZ.IsAmmo = True And ItemZ.ShootType = ItemH.ShootType Then
									ItemAmmo = ItemZ
									Exit Property
								End If
							Next ItemZ
						End If
					End If
				Next ItemX
				' Unequipped next
				For	Each ItemX In myItems
					If ItemX.IsReady = False Then
						' See if item itself is ammo
						If ItemX.IsAmmo = True And ItemX.ShootType = ItemH.ShootType Then
							ItemAmmo = ItemX
							Exit Property
						Else
							' Search inside the item for ammo
							For	Each ItemZ In ItemX.Items
								If ItemZ.IsAmmo = True And ItemZ.ShootType = ItemH.ShootType Then
									ItemAmmo = ItemZ
									Exit Property
								End If
							Next ItemZ
						End If
					End If
				Next ItemX
			End If
			ItemAmmo = ItemH
		End Get
	End Property
	
	Public ReadOnly Property ItemInHand() As Item
		Get
            Dim ItemX As Item
			For	Each ItemX In myItems
				If ItemX.IsReady And ItemX.InHand Then
					ItemInHand = ItemX
					Exit Property
				End If
			Next ItemX
			ItemInHand = New Item
			ItemInHand.Name = "Hand"
			ItemInHand.Damage = 1
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
	
	Public ReadOnly Property Conversations() As Collection
		Get
			Conversations = myConversations
		End Get
	End Property
	
	Public ReadOnly Property MovementRange() As Short
		Get
			MovementRange = Int(Me.ActionPoints / Me.MovementCost)
		End Get
	End Property
	
	Public ReadOnly Property MovementCost() As Short
		Get
			Select Case Me.Agility
				Case 0 To 4
					MovementCost = 6
				Case 5 To 9
					MovementCost = 5
				Case 10 To 14
					MovementCost = 4
				Case 15 To 19
					MovementCost = 3
				Case 20 To 24
					MovementCost = 2
				Case Else
					MovementCost = 1
			End Select
			' Penalty if total weight is near or over max
			' [Titi 2.4.8] increased penalties - and now also accounts for bulkiness!
			If Me.Weight > Me.MaxWeight * 0.75 Or Me.Bulk > 0.75 * Me.Agility * 8 Then
				MovementCost = MovementCost + 2
			End If
			If Me.Weight > Me.MaxWeight * 0.9 Or Me.Bulk > 0.9 * Me.Agility * 8 Then
				MovementCost = Int(MovementCost * 1.5 + 1)
			End If
			' Add in MovementCostBonus
			MovementCost = MovementCost - Me.MovementCostBonus
			If MovementCost < 1 Then
				MovementCost = 1
			End If
		End Get
	End Property
	
	
	Public Property MovementCostBonus() As Short
		Get
			MovementCostBonus = myMovementCostBonus
		End Get
		Set(ByVal Value As Short)
			myMovementCostBonus = Value
		End Set
	End Property
	
	
	Public Property Friendly() As Short
		Get
			Friendly = (myFlags And &H2) > 0
		End Get
		Set(ByVal Value As Short)
			If Value Then
				myFlags = myFlags Or &H2
			Else
				myFlags = myFlags And (Not &H2)
			End If
		End Set
	End Property
	
	
	Public Property DMControlled() As Short
		Get
			DMControlled = (myFlags And &H4) > 0
		End Get
		Set(ByVal Value As Short)
			If Value Then
				myFlags = myFlags Or &H4
			Else
				myFlags = myFlags And (Not &H4)
			End If
		End Set
	End Property
	
	
	Public Property Guard() As Short
		Get
			Guard = (myFlags And &H8) > 0
		End Get
		Set(ByVal Value As Short)
			If Value Then
				myFlags = myFlags Or &H8
			Else
				myFlags = myFlags And (Not &H8)
			End If
		End Set
	End Property
	
	
	Public Property IsSelected() As Short
		Get
			IsSelected = (myStatus And &H8) > 0
		End Get
		Set(ByVal Value As Short)
			If Value Then
				myStatus = myStatus Or &H8
			Else
				myStatus = myStatus And (Not &H8)
			End If
		End Set
	End Property
	
	
	Public Property Agressive() As Short
		Get
			Agressive = (myFlags And &H1) > 0
		End Get
		Set(ByVal Value As Short)
			If Value Then
				myFlags = myFlags Or &H1
			Else
				myFlags = myFlags And (Not &H1)
			End If
		End Set
	End Property
	
	
	Public Property OpportunityAttack() As Short
		Get
			OpportunityAttack = (myCombatAttitude And &H10) > 0
		End Get
		Set(ByVal Value As Short)
			If Value Then
				myCombatAttitude = myCombatAttitude Or &H10
			Else
				myCombatAttitude = myCombatAttitude And (Not &H10)
			End If
		End Set
	End Property
	
	
	Public Property CombatAttitude() As Short
		Get
			CombatAttitude = (myCombatAttitude And &HF)
		End Get
		Set(ByVal Value As Short)
			myCombatAttitude = myCombatAttitude And (Not &HF)
			myCombatAttitude = myCombatAttitude Or Value
		End Set
	End Property
	
	
	Public Property CombatAction1() As Short
		Get
			CombatAction1 = myCombatAction1
		End Get
		Set(ByVal Value As Short)
			myCombatAction1 = Value
		End Set
	End Property
	
	
	Public Property CreatureTarget() As Creature
		Get
			CreatureTarget = myCreatureTarget
		End Get
		Set(ByVal Value As Creature)
			myCreatureTarget = Value
		End Set
	End Property
	
	
	Public Property CombatRank() As Short
		Get
			CombatRank = CShort(myStatus And &H6) / 2
		End Get
		Set(ByVal Value As Short)
			myStatus = myStatus And (Not &H6)
			myStatus = myStatus Or (Value * 2)
		End Set
	End Property
	
	
	Public Property Frozen() As Short
		Get
			Frozen = (myStatus And &H1) > 0
		End Get
		Set(ByVal Value As Short)
			If Value Then
				myStatus = myStatus Or &H1
			Else
				myStatus = myStatus And (Not &H1)
			End If
		End Set
	End Property
	
	
	Public Property Afraid() As Short
		Get
			Afraid = (myStatus And &H40) > 0
		End Get
		Set(ByVal Value As Short)
			If Value Then
				myStatus = myStatus Or &H40
			Else
				myStatus = myStatus And (Not &H40)
			End If
		End Set
	End Property
	
	Public ReadOnly Property Pronoun(ByVal Index As Short) As String
		Get
			Select Case Index
				Case 1
					If Me.Male = True Then
						Pronoun = "he"
					Else
						Pronoun = "she"
					End If
				Case 2
					If Me.Male = True Then
						Pronoun = "him"
					Else
						Pronoun = "her"
					End If
				Case 3
					If Me.Male = True Then
						Pronoun = "his"
					Else
						Pronoun = "her"
					End If
			End Select
		End Get
	End Property
	
	
	Public Property HadTurn() As Short
		Get
			HadTurn = (myFlags And &H10) > 0
		End Get
		Set(ByVal Value As Short)
			If Value Then
				myFlags = myFlags Or &H10
			Else
				myFlags = myFlags And (Not &H10)
			End If
		End Set
	End Property
	
	
	Public Property IsInanimate() As Short
		Get
			IsInanimate = (myStatus And &H10) > 0
		End Get
		Set(ByVal Value As Short)
			If Value Then
				myStatus = myStatus Or &H10
			Else
				myStatus = myStatus And (Not &H10)
			End If
		End Set
	End Property
	
	
	Public Property RequiredInTome() As Short
		Get
			RequiredInTome = (myFlags And &H80) > 0
		End Get
		Set(ByVal Value As Short)
			If Value Then
				myFlags = myFlags Or &H80
			Else
				myFlags = myFlags And (Not &H80)
			End If
		End Set
	End Property
	
	
	Public Property Unconscious() As Short
		Get
			Unconscious = (myStatus And &H80) > 0
		End Get
		Set(ByVal Value As Short)
			If Value Then
				myStatus = myStatus Or &H80
			Else
				myStatus = myStatus And (Not &H80)
			End If
		End Set
	End Property
	
	
	Public Property ScarletLetter() As Short
		Get
			ScarletLetter = (myFlags And &H20) > 0
		End Get
		Set(ByVal Value As Short)
			If Value Then
				myFlags = myFlags Or &H20
			Else
				myFlags = myFlags And (Not &H20)
			End If
		End Set
	End Property
	
	
	Public Property Male() As Short
		Get
			Male = (myFlags And &H40) > 0
		End Get
		Set(ByVal Value As Short)
			If Value Then
				myFlags = myFlags Or &H40
			Else
				myFlags = myFlags And (Not &H40)
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
	
	
	Public Property Status() As Short
		Get
			Status = myStatus
		End Get
		Set(ByVal Value As Short)
			myStatus = Value
		End Set
	End Property
	
	Public ReadOnly Property StatusText() As String
		Get
			If myHPNow < 1 Then
				StatusText = "Dead"
			ElseIf Me.Unconscious = True Then 
				StatusText = "Unconscious"
			ElseIf Me.Frozen = True Then 
				StatusText = "Frozen"
			ElseIf myActionPoints < 1 Then 
				StatusText = "Exhausted"
			Else
				StatusText = "Ok"
			End If
		End Get
	End Property
	
	
	Public Property Eternium() As Short
		Get
			Eternium = myEternium
		End Get
		Set(ByVal Value As Short)
			Dim c As Short
			c = Me.EterniumMax
			If Value < 1 Then
				myEternium = 0
			ElseIf Value <= c Then 
				myEternium = Value
			Else
				myEternium = c
			End If
		End Set
	End Property
	
	Public ReadOnly Property EterniumMax() As Short
		Get
			EterniumMax = myWill * 4
			If EterniumMax > 250 Then
				EterniumMax = 250
			End If
		End Get
	End Property
	
	
	Public Property Runes(ByVal Index As Short) As Short
		Get
			Runes = myRunes(Index)
		End Get
		Set(ByVal Value As Short)
			myRunes(Index) = Value
		End Set
	End Property
	
	
	Public Property RuneTop() As Short
		Get
			RuneTop = myRuneTop
		End Get
		Set(ByVal Value As Short)
			myRuneTop = Value
		End Set
	End Property
	
	
	Public Property RuneQueLimit() As Short
		Get
			RuneQueLimit = myRuneQueLimit
		End Get
		Set(ByVal Value As Short)
			myRuneQueLimit = Value
		End Set
	End Property
	
	
	Public Property PictureFile() As String
		Get
			PictureFile = Trim(BreakText(myPictureFile, 1))
		End Get
		Set(ByVal Value As String)
			If Len(BreakText(myPictureFile, 2)) > 0 Then
				myPictureFile = Trim(Value) & "|" & BreakText(myPictureFile, 2)
			Else
				myPictureFile = Trim(Value)
			End If
		End Set
	End Property
	
	
	Public Property PortraitFile() As String
		Get
			PortraitFile = Trim(BreakText(myPictureFile, 2))
		End Get
		Set(ByVal Value As String)
			myPictureFile = BreakText(myPictureFile, 1) & "|" & Value
		End Set
	End Property
	
	
	Public Property OnAdventure() As Short
		Get
			OnAdventure = (myFlagsWAV And &H20) > 0
		End Get
		Set(ByVal Value As Short)
			If Value Then
				myFlagsWAV = myFlagsWAV Or &H20
			Else
				myFlagsWAV = myFlagsWAV And (Not &H20)
			End If
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
	
	
	Public Property FaceLeft() As Short
		Get
			FaceLeft = myFaceLeft
		End Get
		Set(ByVal Value As Short)
			myFaceLeft = Value
		End Set
	End Property
	
	
	Public Property FaceTop() As Short
		Get
			FaceTop = myFaceTop
		End Get
		Set(ByVal Value As Short)
			myFaceTop = Value
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
	
	
	Public Property Steps() As Short
		Get
			Steps = mySteps
		End Get
		Set(ByVal Value As Short)
			mySteps = Value
		End Set
	End Property
	
	Public ReadOnly Property TileSpotX() As Short
		Get
			Select Case myTileSpot
				Case 0 : TileSpotX = 32
				Case 1 : TileSpotX = 20
				Case 2 : TileSpotX = 8
				Case 3 : TileSpotX = 44
				Case 4 : TileSpotX = 32
				Case 5 : TileSpotX = 20
				Case 7 : TileSpotX = 44
				Case 6 : TileSpotX = 56
				Case 8 : TileSpotX = 32
			End Select
		End Get
	End Property
	
	Public ReadOnly Property TileSpotY() As Short
		Get
			Select Case myTileSpot
				Case 0 : TileSpotY = 40
				Case 1 : TileSpotY = 46
				Case 2 : TileSpotY = 52
				Case 3 : TileSpotY = 46
				Case 4 : TileSpotY = 52
				Case 5 : TileSpotY = 58
				Case 6 : TileSpotY = 52
				Case 7 : TileSpotY = 58
				Case 8 : TileSpotY = 64
			End Select
		End Get
	End Property
	
	
	Public Property TileSpot() As Short
		Get
			TileSpot = myTileSpot
		End Get
		Set(ByVal Value As Short)
			myTileSpot = Value
		End Set
	End Property
	
	
	Public Property Col() As Short
		Get
			Col = myCol
		End Get
		Set(ByVal Value As Short)
			myCol = Value
		End Set
	End Property
	
	
	Public Property Row() As Short
		Get
			Row = myRow
		End Get
		Set(ByVal Value As Short)
			myRow = Value
		End Set
	End Property
	
	
	Public Property Resistance(ByVal Index As Short) As Short
		Get
			' Normal resistance: Index is location of hit (Body, Head, etc.)
			Resistance = myResistance(Index) + Me.ResistanceBonus(0)
		End Get
		Set(ByVal Value As Short)
			myResistance(Index) = Value
		End Set
	End Property
	
	
	Public Property ResistanceBonus(ByVal Index As Short) As Short
		Get
			' Bonus to resistance of various types: Index is 0 = General, 1 = Sharp, 2 = Blunt, etc.
			ResistanceBonus = myResistanceBonus(Index)
		End Get
		Set(ByVal Value As Short)
			myResistanceBonus(Index) = Value
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
	
	
	Public Property DefenseBonus() As Short
		Get
			Dim c As Short
			' Modifier is based on the Creature's size
			' Bit 0     -   On = Tiny (+4)
			' Bit 1     -   On = Small (+2)
			' Bit 2     -   On = Medium (0)
			' Bit 3     -   On = Large (-2)
			' Bit 4     -   On = Huge (-4)
			DefenseBonus = myDefenseBonus
			For c = 0 To 4
				If Me.Size(c) = True Then
					DefenseBonus = DefenseBonus + (2 - c) * 2
				End If
			Next c
		End Get
		Set(ByVal Value As Short)
			myDefenseBonus = Value
		End Set
	End Property
	
	
	Public Property AttackBonus() As Short
		Get
			AttackBonus = myAttackBonus + Fix((Me.Agility - 10) / 3)
		End Get
		Set(ByVal Value As Short)
			myAttackBonus = Value
		End Set
	End Property
	
	
	Public Property WillBonus() As Short
		Get
			WillBonus = myWillBonus
		End Get
		Set(ByVal Value As Short)
			myWillBonus = Value
		End Set
	End Property
	
	
	Public Property AgilityBonus() As Short
		Get
			AgilityBonus = myAgilityBonus
		End Get
		Set(ByVal Value As Short)
			myAgilityBonus = Value
		End Set
	End Property
	
	
	Public Property StrengthBonus() As Short
		Get
			StrengthBonus = myStrengthBonus
		End Get
		Set(ByVal Value As Short)
			myStrengthBonus = Value
		End Set
	End Property
	
	
	Public Property Lunacy() As Short
		Get
			Lunacy = myLunacy
		End Get
		Set(ByVal Value As Short)
			myLunacy = Value
		End Set
	End Property
	
	
	Public Property Revelry() As Short
		Get
			Revelry = myRevelry
		End Get
		Set(ByVal Value As Short)
			myRevelry = Value
		End Set
	End Property
	
	
	Public Property Wrath() As Short
		Get
			Wrath = myWrath
		End Get
		Set(ByVal Value As Short)
			myWrath = Value
		End Set
	End Property
	
	
	Public Property Pride() As Short
		Get
			Pride = myPride
		End Get
		Set(ByVal Value As Short)
			myPride = Value
		End Set
	End Property
	
	
	Public Property Greed() As Short
		Get
			Greed = myGreed
		End Get
		Set(ByVal Value As Short)
			myGreed = Value
		End Set
	End Property
	
	
	Public Property Lust() As Short
		Get
			Lust = myLust
		End Get
		Set(ByVal Value As Short)
			myLust = Value
		End Set
	End Property
	
	
	Public Property CombatAction2() As Short
		Get
			CombatAction2 = myCombatAction2
		End Get
		Set(ByVal Value As Short)
			myCombatAction2 = Value
		End Set
	End Property
	
	Public ReadOnly Property ActionPointsMax() As Short
		Get
			Dim ItemX As Item
			' [Titi 2.4.9] RT6 when (myWill + myAgility + myStrength) is greater than 255
			'    ActionPointsMax = Fix((myWill + myAgility + myStrength) / 3) + Me.Level
			ActionPointsMax = Fix(myWill / 3 + myAgility / 3 + myStrength / 3) + Me.Level
			For	Each ItemX In myItems
				If ItemX.IsReady = True And ItemX.InHand = False Then
					ActionPointsMax = ActionPointsMax - ItemX.ActionPoints
				End If
			Next ItemX
		End Get
	End Property
	
	
	Public Property ResistanceTypeBonus(ByVal Index As Short) As Short
		Get
			Dim b1 As Byte
			' Special resistance to certain attacks: Index is 0 = Sharp, 1 = Blunt, etc.
			' Returns an index:  0 = None, 1 = 10%, 2 = 20% ... 10 = 100%, 11 = Double Damage, 12 = Triple Damage
			b1 = 2 ^ Index
			ResistanceTypeBonus = System.Math.Abs(CInt((myResistance1 And b1) > 0)) + System.Math.Abs(CInt((myResistance2 And b1) > 0)) * 2 + System.Math.Abs(CInt((myResistance3 And b1) > 0)) * 4 + System.Math.Abs(CInt((myResistance4 And b1) > 0)) * 8
		End Get
		Set(ByVal Value As Short)
			Dim b1 As Byte
			' Index is 0 = Sharp, 1 = Blunt, etc.
			b1 = 2 ^ Index
			' Clear the current value
			myResistance1 = myResistance1 And (Not b1)
			myResistance2 = myResistance2 And (Not b1)
			myResistance3 = myResistance3 And (Not b1)
			myResistance4 = myResistance4 And (Not b1)
			' Set the new value
			Select Case Value
				Case 1, 3, 5, 7, 9, 11, 13, 15
					' 0001
					myResistance1 = myResistance1 Or b1
			End Select
			Select Case Value
				Case 2, 3, 6, 7, 10, 11, 14, 15
					' 0010
					myResistance2 = myResistance2 Or b1
			End Select
			Select Case Value
				Case 4, 5, 6, 7, 12, 13, 14, 15
					' 0100
					myResistance3 = myResistance3 Or b1
			End Select
			Select Case Value
				Case 8, 9, 10, 11, 12, 13, 14, 15
					' 1000
					myResistance4 = myResistance4 Or b1
			End Select
		End Set
	End Property
	
	Public ReadOnly Property ResistanceWithArmor(ByVal Index As Short) As Short
		Get
			Dim ItemX As Item
			' Index is location of hit: Body, Head, etc.
			ResistanceWithArmor = Me.Resistance(Index) + Me.ResistanceBonus(0)
			' Add up Armor Resistance
			For	Each ItemX In Me.Items
				If ItemX.IsReady = True Then
					Select Case Index
						Case 0 To 3 ' Body
							If ItemX.WearType = 0 Then ' Body
								ResistanceWithArmor = ResistanceWithArmor + ItemX.Resistance
							End If
						Case 4 To 5 ' Shield
							If ItemX.WearType = 5 Then ' Shield
								ResistanceWithArmor = ResistanceWithArmor + ItemX.Resistance
							End If
						Case 6 ' Head
							If ItemX.WearType = 1 Then ' Helm
								ResistanceWithArmor = ResistanceWithArmor + ItemX.Resistance
							End If
						Case 7 ' Neck (No Armor)
					End Select
				End If
			Next ItemX
		End Get
	End Property
	
	Public ReadOnly Property ResistanceTypeWithArmor(ByVal DamageType As Short) As Short
		Get
			Dim ItemX As Item

			' Damage type is:
			'       0 - None, 1 - Sharp, 2 - Blunt, 3 - Cold, 4 - Heat, 5 - Evil
			'       6 - Good, 7 - Energy, 8 - Mind
			If DamageType > 0 Then
				Select Case Me.ResistanceTypeBonus(DamageType - 1)
					Case 0 To 10
						ResistanceTypeWithArmor = Me.ResistanceTypeBonus(DamageType - 1) * 10
					Case 11
						ResistanceTypeWithArmor = -100
					Case 12
						ResistanceTypeWithArmor = -200
				End Select
				' Add Bonus (if any)
				ResistanceTypeWithArmor = ResistanceTypeWithArmor + Me.ResistanceBonus(DamageType)
				' Add up Armor Resistance
				For	Each ItemX In Me.Items
					If ItemX.IsReady = True And ItemX.ResistanceType = DamageType Then
						Select Case ItemX.ResistanceBonus
							Case 0 To 10
								ResistanceTypeWithArmor = ResistanceTypeWithArmor + ItemX.ResistanceBonus * 10
							Case 11
								ResistanceTypeWithArmor = ResistanceTypeWithArmor - 100
							Case 12
								ResistanceTypeWithArmor = ResistanceTypeWithArmor - 200
						End Select
					End If
				Next ItemX
			End If
		End Get
	End Property
	
	
	Public Property BodyType(ByVal Index As Short) As Short
		Get
			BodyType = myBodyType(Index)
		End Get
		Set(ByVal Value As Short)
			myBodyType(Index) = Value
		End Set
	End Property
	
	
	Public Property Strength() As Short
		Get
			Strength = myStrength + myStrengthBonus
		End Get
		Set(ByVal Value As Short)
			myStrength = Value
		End Set
	End Property
	
	
	Public Property Will() As Short
		Get
			Will = myWill + myWillBonus
		End Get
		Set(ByVal Value As Short)
			myWill = Value
		End Set
	End Property
	
	
	Public Property Agility() As Short
		Get
			Agility = myAgility + myAgilityBonus
		End Get
		Set(ByVal Value As Short)
			myAgility = Value
		End Set
	End Property
	
	
	Public Property Facing() As Short
		Get
			Facing = myFacing
		End Get
		Set(ByVal Value As Short)
			myFacing = Value
		End Set
	End Property
	
	Public ReadOnly Property AllowedTurn() As Short
		Get
			If myHPNow > 0 And Me.Unconscious = False And Me.Frozen = False And Me.Afraid = False And Me.IsInanimate = False Then
				AllowedTurn = True
			Else
				AllowedTurn = False
			End If
		End Get
	End Property
	
	Public ReadOnly Property IsSpellCaster() As Short
		Get
			Dim TriggerX As Trigger
			Dim StatementX As Statement
			IsSpellCaster = False
			For	Each TriggerX In myTriggers
				If TriggerX.TriggerType = 39 Then ' OnSkillUse Trigger
					For	Each StatementX In TriggerX.Statements
						If StatementX.Statement = 47 Then ' Sorcery Statement
							IsSpellCaster = True
							Exit For
						End If
					Next StatementX
				End If
			Next TriggerX
		End Get
	End Property
	
	
	Public Property Money() As Integer
		Get
			Dim ItemX As Item
			Money = 0
			For	Each ItemX In Me.Items
				Money = Money + ItemX.Money
			Next ItemX
		End Get
		Set(ByVal Value As Integer)
			Const bdMoneyType As Short = 2
			Dim ItemX As Item
			Dim c As Integer
			
			If Value > Me.Money Then
				' Add Money
				c = Value - Me.Money
				For	Each ItemX In Me.Items
					If ItemX.Family = bdMoneyType And (ItemX.Value / ItemX.Count) = 1 Then
						ItemX.Value = ItemX.Value + c
						ItemX.Count = ItemX.Count + c
						Exit Property
					End If
				Next ItemX
				' Didn't find Money, so create some
				ItemX = Me.AddItem
				' [Titi 2.4.8] let's use the currency of the world!
				ItemX.Name = Left(WorldNow.Money, InStr(WorldNow.Money, "|") - 1)
				'        ItemX.Name = "Gold piece"
				ItemX.Family = bdMoneyType
				ItemX.Value = c
				ItemX.Count = c
				ItemX.Weight = c
				ItemX.Bulk = Greatest(CShort(c / 10), 1)
				ItemX.CanCombine = True
				ItemX.PictureFile = Right(WorldNow.Money, Len(WorldNow.Money) - InStr(WorldNow.Money, "|"))
				'        ItemX.PictureFile = "gold1.bmp"
			ElseIf Value < Me.Money Then 
				' Remove Money
				c = Me.Money - Value
				For	Each ItemX In Me.Items
					If ItemX.RemoveMoney(CShort(c)) = True Then
						Me.RemoveItem("I" & ItemX.Index)
					End If
				Next ItemX
			End If
		End Set
	End Property
	
	
	Public Property ActionPoints() As Short
		Get
			ActionPoints = myActionPoints
		End Get
		Set(ByVal Value As Short)
			If Value < 1 Then
				myActionPoints = 0
			ElseIf Value > 255 Then 
				myActionPoints = 255
			Else
				myActionPoints = Value
			End If
		End Set
	End Property
	
	
	Public Property ExperiencePoints() As Integer
		Get
			ExperiencePoints = myExperiencePoints
		End Get
		Set(ByVal Value As Integer)
			myExperiencePoints = Value
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
	
	
	Public Property TempSkillSet() As Short
		Get
			TempSkillSet = myTempSkillSet
		End Get
		Set(ByVal Value As Short)
			myTempSkillSet = Value
		End Set
	End Property
	
	
	Public Property TempSkillSpend() As Short
		Get
			TempSkillSpend = myTempSkillSpend
		End Get
		Set(ByVal Value As Short)
			myTempSkillSpend = Value
		End Set
	End Property
	
	
	Public Property Level() As Short
		Get
			Level = myLevel
		End Get
		Set(ByVal Value As Short)
			myLevel = Value
			If myLevel < 1 Then
				myLevel = 1
			End If
		End Set
	End Property
	
	Public ReadOnly Property EncounterPoints() As Short
		Get
			Select Case myLevel
				Case 0 To 1
					EncounterPoints = 1
				Case 2
					EncounterPoints = 2
				Case 3
					EncounterPoints = 3
				Case 4
					EncounterPoints = 4
				Case 5
					EncounterPoints = 5
				Case 6
					EncounterPoints = 10
				Case 7
					EncounterPoints = 15
				Case 8
					EncounterPoints = 20
				Case 9
					EncounterPoints = 30
				Case Else
					EncounterPoints = 75
			End Select
		End Get
	End Property
	
	
	Public Property HPNow() As Short
		Get
			HPNow = myHPNow
		End Get
		Set(ByVal Value As Short)
			myHPNow = Value
			If myHPNow > myHPMax Then
				myHPNow = myHPMax
			End If
		End Set
	End Property
	
	
	Public Property HPMax() As Short
		Get
			HPMax = myHPMax
		End Get
		Set(ByVal Value As Short)
			myHPMax = Value
		End Set
	End Property
	
	Public ReadOnly Property Weight() As Short
		Get
			Dim ItemX As Item
			Dim c As Short
			c = 0
			For	Each ItemX In myItems
				c = c + ItemX.Weight
			Next ItemX
			Weight = c
		End Get
	End Property
	
	Public ReadOnly Property MaxWeight() As Short
		Get
			MaxWeight = myStrength * 15
		End Get
	End Property
	
	Public ReadOnly Property Bulk() As Short
		Get
			Dim ItemX As Item
			Dim c As Double
			c = 0
			For	Each ItemX In myItems
				If ItemX.IsReady = False Then
					c = c + ItemX.Bulk
				End If
			Next ItemX
			Bulk = c
		End Get
	End Property
	
	Public ReadOnly Property Defense() As Short
		Get
			Dim ItemX As Item
			Defense = 13 + Me.DefenseBonus
			' Modify based on current Health
			If Me.HPNow > 0 And Me.HPNow < Me.HPMax Then
				Defense = Defense - 3 + System.Math.Round((Me.HPNow / Me.HPMax) * 3)
			End If
			' Modify based on what you're wearing
			For	Each ItemX In Me.Items
				' If wearing and not an InHand type item
				If ItemX.IsReady = True Then
					Defense = Defense + ItemX.Defense
				End If
			Next ItemX
			If Defense < 0 Then
				Defense = 1
			End If
		End Get
	End Property
	
	
	Public Property Family(ByVal Check As Short) As Short
		Get
			Family = (myFamily And (2 ^ Check)) > 0
		End Get
		Set(ByVal Value As Short)
			If Value Then
				myFamily = myFamily Or (2 ^ Check)
			Else
				myFamily = myFamily And (Not (2 ^ Check))
			End If
		End Set
	End Property
	
	
	Public Property Size(ByVal Check As Short) As Short
		Get
			Size = (mySize And (2 ^ Check)) > 0
		End Get
		Set(ByVal Value As Short)
			If Value Then
				mySize = mySize Or (2 ^ Check)
			Else
				mySize = mySize And (Not (2 ^ Check))
			End If
		End Set
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
		If myMaxStep > 0 Then
			X = myStepX(myMaxStep) : Y = myStepY(myMaxStep)
			myMaxStep = myMaxStep - 1
		Else
			X = myMapX : Y = myMapY
		End If
	End Sub
	
	Public Sub UseAmmo()
		Dim ItemX, ItemAmmo, ItemZ As Item

		ItemAmmo = Me.ItemAmmo
		If ItemAmmo.IsAmmo = True Then
			If ItemAmmo.Count = 1 Then
				' Search through equipped first
				For	Each ItemX In myItems
					If ItemX.Name = ItemAmmo.Name And ItemX.Index = ItemAmmo.Index Then
						Me.RemoveItem("I" & ItemX.Index)
						Exit Sub
					End If
					For	Each ItemZ In ItemX.Items
						If ItemZ.Name = ItemAmmo.Name And ItemZ.Index = ItemAmmo.Index Then
							ItemX.RemoveItem("I" & ItemZ.Index)
							Exit Sub
						End If
					Next ItemZ
				Next ItemX
			Else
				ItemAmmo.Bulk = ItemAmmo.Bulk - Int(ItemAmmo.Bulk / ItemAmmo.Count)
				ItemAmmo.Weight = ItemAmmo.Weight - Int(ItemAmmo.Weight / ItemAmmo.Count)
				ItemAmmo.Value = ItemAmmo.Value - Int(ItemAmmo.Value / ItemAmmo.Count)
				ItemAmmo.Count = ItemAmmo.Count - 1
			End If
		ElseIf ItemAmmo.Family = 8 Then  ' Thrown item
			' Weapon (Thrown) self distructs
			If ItemAmmo.Count = 1 Then
				Me.RemoveItem("I" & ItemAmmo.Index)
			Else
				ItemAmmo.Count = ItemAmmo.Count - 1
			End If
		End If
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
		AddItemBefore = ItemX
	End Function
	
	Public Sub RemoveItem(ByRef DeleteKey As String)
		myItems.Remove(DeleteKey)
	End Sub
	
	Public Function AddTrigger() As Trigger
		' Adds an Item and returns a reference to that Item
		Dim c As Short
		Dim TriggerX As Trigger
		' Find new available unused index identifier
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
	
	Public Function AddConversation() As Conversation
		' Adds a Conversation and returns a reference to that Conversation
		Dim c As Short
		Dim ConversationX As Conversation
		' Find new available unused index idenifier
		c = 1
		For	Each ConversationX In myConversations
			If ConversationX.Index >= c Then
				c = ConversationX.Index + 1
			End If
		Next ConversationX
		' Create the Conversation, set the Index and add it.
		ConversationX = New Conversation
		ConversationX.Index = c
		ConversationX.Name = "Convo" & c
		myConversations.Add(ConversationX, "C" & ConversationX.Index)
		AddConversation = ConversationX
	End Function
	
	Public Sub RemoveConversation(ByRef DeleteKey As String)
		'UPGRADE_WARNING: Couldn't resolve default property of object myConversations(DeleteKey).Index. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
		If myConversations.Item(DeleteKey).Index = myCurrentConvo Then
			myCurrentConvo = 0
		End If
		myConversations.Remove(DeleteKey)
	End Sub
	
	Public Sub Copy(ByRef FromCreature As Creature)
        Dim c As Short
		Dim myTrigger As Trigger
		Dim myItem As Item
		Dim ConversationX As Conversation
		myName = FromCreature.Name
		myComments = FromCreature.Comments
		myRace = FromCreature.Race
		myHome = FromCreature.Home ' [Titi 2.4.9]
		myFlags = FromCreature.Flags
		myStatus = FromCreature.Status
		myEternium = FromCreature.Eternium
		myPictureFile = FromCreature.PictureFile
		myPic = FromCreature.Pic
		myFaceTop = FromCreature.FaceTop
		myFaceLeft = FromCreature.FaceLeft
		For c = 0 To 7
			myResistance(c) = FromCreature.Resistance(c)
			myBodyType(c) = FromCreature.BodyType(c)
		Next c
		For c = 0 To 7
			Me.ResistanceTypeBonus(c) = FromCreature.ResistanceTypeBonus(c)
		Next c
		myStrength = FromCreature.Strength
		myWill = FromCreature.Will
		myAgility = FromCreature.Agility
		myActionPoints = FromCreature.ActionPoints
		myExperiencePoints = FromCreature.ExperiencePoints
		mySkillPoints = FromCreature.SkillPoints
		myLevel = FromCreature.Level
		myHPNow = FromCreature.HPNow
		myHPMax = FromCreature.HPMax
		myLunacy = FromCreature.Lunacy
		myRevelry = FromCreature.Revelry
		myWrath = FromCreature.Wrath
		myPride = FromCreature.Pride
		myGreed = FromCreature.Greed
		myLust = FromCreature.Lust
		myCombatAction2 = FromCreature.CombatAction2
		myMoveWAV = FromCreature.MoveWAV
		myDieWAV = FromCreature.DieWAV
		myAttackWAV = FromCreature.AttackWAV
		myHitWAV = FromCreature.HitWAV
		Me.MoveWAVOneTime = FromCreature.MoveWAVOneTime
		Me.DieWAVOneTime = FromCreature.DieWAVOneTime
		Me.AttackWAVOneTime = FromCreature.AttackWAVOneTime
		Me.HitWAVOneTime = FromCreature.HitWAVOneTime
		For c = 0 To 9
			Me.Family(c) = FromCreature.Family(c)
		Next c
		For c = 0 To 4
			Me.Size(c) = FromCreature.Size(c)
		Next c
		' Copy Items for Creature
		myItems = New Collection
		For c = 1 To FromCreature.Items.Count()
			myItem = Me.AddItem
			'UPGRADE_WARNING: Couldn't resolve default property of object FromCreature.Items(). Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
			myItem.Copy(FromCreature.Items.Item(c))
		Next c
		' Copy Triggers for Creature
		myCurrentConvo = FromCreature.CurrentConvo
		myTriggers = New Collection
		For c = 1 To FromCreature.Triggers.Count()
			myTrigger = Me.AddTrigger
			'UPGRADE_WARNING: Couldn't resolve default property of object FromCreature.Triggers(). Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
			myTrigger.Copy(FromCreature.Triggers.Item(c))
		Next c
		' Copy Conversation for Creature
		myConversations = New Collection
		For c = 1 To FromCreature.Conversations.Count()
			ConversationX = Me.AddConversation
			'UPGRADE_WARNING: Couldn't resolve default property of object FromCreature.Conversations(). Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
			ConversationX.Copy(FromCreature.Conversations.Item(c))
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
		'UPGRADE_WARNING: Get was upgraded to FileGet and has a new behavior. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="9B7D5ADD-D8FE-4819-A36C-6DEDAF088CC7"'
		FileGet(FromFile, Cnt)
		myRace = ""
		For c = 1 To Cnt
			myRace = myRace & " "
		Next c
		'UPGRADE_WARNING: Get was upgraded to FileGet and has a new behavior. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="9B7D5ADD-D8FE-4819-A36C-6DEDAF088CC7"'
		FileGet(FromFile, myRace)
		' [Titi 2.4.9] added the home tome
		If myVersion = 249 Then
			' for compatibility with previous versions
			'UPGRADE_WARNING: Get was upgraded to FileGet and has a new behavior. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="9B7D5ADD-D8FE-4819-A36C-6DEDAF088CC7"'
			FileGet(FromFile, Cnt)
			myHome = ""
			For c = 1 To Cnt
				myHome = myHome & " "
			Next c
			'UPGRADE_WARNING: Get was upgraded to FileGet and has a new behavior. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="9B7D5ADD-D8FE-4819-A36C-6DEDAF088CC7"'
			FileGet(FromFile, myHome)
		Else
			myVersion = 249
			myHome = WorldNow.Name & "\" & "Homeless.tom"
		End If
		'UPGRADE_WARNING: Get was upgraded to FileGet and has a new behavior. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="9B7D5ADD-D8FE-4819-A36C-6DEDAF088CC7"'
		FileGet(FromFile, myFlags)
		'UPGRADE_WARNING: Get was upgraded to FileGet and has a new behavior. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="9B7D5ADD-D8FE-4819-A36C-6DEDAF088CC7"'
		FileGet(FromFile, myStatus)
		'UPGRADE_WARNING: Get was upgraded to FileGet and has a new behavior. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="9B7D5ADD-D8FE-4819-A36C-6DEDAF088CC7"'
		FileGet(FromFile, myEternium)
		'UPGRADE_WARNING: Get was upgraded to FileGet and has a new behavior. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="9B7D5ADD-D8FE-4819-A36C-6DEDAF088CC7"'
		FileGet(FromFile, Cnt)
		myPictureFile = ""
		For c = 1 To Cnt
			myPictureFile = myPictureFile & " "
		Next c
		'UPGRADE_WARNING: Get was upgraded to FileGet and has a new behavior. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="9B7D5ADD-D8FE-4819-A36C-6DEDAF088CC7"'
		FileGet(FromFile, myPictureFile)
		'UPGRADE_WARNING: Get was upgraded to FileGet and has a new behavior. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="9B7D5ADD-D8FE-4819-A36C-6DEDAF088CC7"'
		FileGet(FromFile, myFaceLeft)
		'UPGRADE_WARNING: Get was upgraded to FileGet and has a new behavior. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="9B7D5ADD-D8FE-4819-A36C-6DEDAF088CC7"'
		FileGet(FromFile, myFaceTop)
		'UPGRADE_WARNING: Get was upgraded to FileGet and has a new behavior. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="9B7D5ADD-D8FE-4819-A36C-6DEDAF088CC7"'
		FileGet(FromFile, myMapX)
		'UPGRADE_WARNING: Get was upgraded to FileGet and has a new behavior. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="9B7D5ADD-D8FE-4819-A36C-6DEDAF088CC7"'
		FileGet(FromFile, myMapY)
		'UPGRADE_WARNING: Get was upgraded to FileGet and has a new behavior. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="9B7D5ADD-D8FE-4819-A36C-6DEDAF088CC7"'
		FileGet(FromFile, myTileSpot)
		For c = 0 To 7
			'UPGRADE_WARNING: Get was upgraded to FileGet and has a new behavior. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="9B7D5ADD-D8FE-4819-A36C-6DEDAF088CC7"'
			FileGet(FromFile, myResistance(c))
			'UPGRADE_WARNING: Get was upgraded to FileGet and has a new behavior. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="9B7D5ADD-D8FE-4819-A36C-6DEDAF088CC7"'
			FileGet(FromFile, myBodyType(c))
		Next c
		'UPGRADE_WARNING: Get was upgraded to FileGet and has a new behavior. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="9B7D5ADD-D8FE-4819-A36C-6DEDAF088CC7"'
		FileGet(FromFile, myResistance1)
		'UPGRADE_WARNING: Get was upgraded to FileGet and has a new behavior. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="9B7D5ADD-D8FE-4819-A36C-6DEDAF088CC7"'
		FileGet(FromFile, myResistance2)
		'UPGRADE_WARNING: Get was upgraded to FileGet and has a new behavior. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="9B7D5ADD-D8FE-4819-A36C-6DEDAF088CC7"'
		FileGet(FromFile, myResistance3)
		'UPGRADE_WARNING: Get was upgraded to FileGet and has a new behavior. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="9B7D5ADD-D8FE-4819-A36C-6DEDAF088CC7"'
		FileGet(FromFile, myResistance4)
		'UPGRADE_WARNING: Get was upgraded to FileGet and has a new behavior. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="9B7D5ADD-D8FE-4819-A36C-6DEDAF088CC7"'
		FileGet(FromFile, myStrength)
		'UPGRADE_WARNING: Get was upgraded to FileGet and has a new behavior. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="9B7D5ADD-D8FE-4819-A36C-6DEDAF088CC7"'
		FileGet(FromFile, myWill)
		'UPGRADE_WARNING: Get was upgraded to FileGet and has a new behavior. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="9B7D5ADD-D8FE-4819-A36C-6DEDAF088CC7"'
		FileGet(FromFile, myAgility)
		'UPGRADE_WARNING: Get was upgraded to FileGet and has a new behavior. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="9B7D5ADD-D8FE-4819-A36C-6DEDAF088CC7"'
		FileGet(FromFile, myNotUsed2)
		'UPGRADE_WARNING: Get was upgraded to FileGet and has a new behavior. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="9B7D5ADD-D8FE-4819-A36C-6DEDAF088CC7"'
		FileGet(FromFile, myExperiencePoints)
		'UPGRADE_WARNING: Get was upgraded to FileGet and has a new behavior. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="9B7D5ADD-D8FE-4819-A36C-6DEDAF088CC7"'
		FileGet(FromFile, mySkillPoints)
		'UPGRADE_WARNING: Get was upgraded to FileGet and has a new behavior. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="9B7D5ADD-D8FE-4819-A36C-6DEDAF088CC7"'
		FileGet(FromFile, myLevel)
		'UPGRADE_WARNING: Get was upgraded to FileGet and has a new behavior. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="9B7D5ADD-D8FE-4819-A36C-6DEDAF088CC7"'
		FileGet(FromFile, myHPNow)
		'UPGRADE_WARNING: Get was upgraded to FileGet and has a new behavior. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="9B7D5ADD-D8FE-4819-A36C-6DEDAF088CC7"'
		FileGet(FromFile, myHPMax)
	End Sub
	
	Public Sub ReadFromFile(ByRef FromFile As Short)
		On Error GoTo ErrorHandler
		Dim c As Short
		Dim Cnt As Integer
		Dim myItem As Item
		Dim myTrigger As Trigger
		Dim ConversationX As Conversation
		Dim Reading As String
		ReadFromFileHeader(FromFile)
		' Read other stats
		'UPGRADE_WARNING: Get was upgraded to FileGet and has a new behavior. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="9B7D5ADD-D8FE-4819-A36C-6DEDAF088CC7"'
		FileGet(FromFile, myLunacy)
		'UPGRADE_WARNING: Get was upgraded to FileGet and has a new behavior. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="9B7D5ADD-D8FE-4819-A36C-6DEDAF088CC7"'
		FileGet(FromFile, myRevelry)
		'UPGRADE_WARNING: Get was upgraded to FileGet and has a new behavior. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="9B7D5ADD-D8FE-4819-A36C-6DEDAF088CC7"'
		FileGet(FromFile, myWrath)
		'UPGRADE_WARNING: Get was upgraded to FileGet and has a new behavior. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="9B7D5ADD-D8FE-4819-A36C-6DEDAF088CC7"'
		FileGet(FromFile, myPride)
		'UPGRADE_WARNING: Get was upgraded to FileGet and has a new behavior. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="9B7D5ADD-D8FE-4819-A36C-6DEDAF088CC7"'
		FileGet(FromFile, myGreed)
		'UPGRADE_WARNING: Get was upgraded to FileGet and has a new behavior. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="9B7D5ADD-D8FE-4819-A36C-6DEDAF088CC7"'
		FileGet(FromFile, myLust)
		'UPGRADE_WARNING: Get was upgraded to FileGet and has a new behavior. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="9B7D5ADD-D8FE-4819-A36C-6DEDAF088CC7"'
		FileGet(FromFile, myCombatAction1)
		'UPGRADE_WARNING: Get was upgraded to FileGet and has a new behavior. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="9B7D5ADD-D8FE-4819-A36C-6DEDAF088CC7"'
		FileGet(FromFile, myCombatAction2)
		'UPGRADE_WARNING: Get was upgraded to FileGet and has a new behavior. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="9B7D5ADD-D8FE-4819-A36C-6DEDAF088CC7"'
		FileGet(FromFile, myFamily)
		'UPGRADE_WARNING: Get was upgraded to FileGet and has a new behavior. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="9B7D5ADD-D8FE-4819-A36C-6DEDAF088CC7"'
		FileGet(FromFile, mySize)
		' Read WAV file names
		'UPGRADE_WARNING: Get was upgraded to FileGet and has a new behavior. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="9B7D5ADD-D8FE-4819-A36C-6DEDAF088CC7"'
		FileGet(FromFile, Cnt)
		myMoveWAV = ""
		For c = 1 To Cnt
			myMoveWAV = myMoveWAV & " "
		Next c
		'UPGRADE_WARNING: Get was upgraded to FileGet and has a new behavior. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="9B7D5ADD-D8FE-4819-A36C-6DEDAF088CC7"'
		FileGet(FromFile, myMoveWAV)
		'UPGRADE_WARNING: Get was upgraded to FileGet and has a new behavior. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="9B7D5ADD-D8FE-4819-A36C-6DEDAF088CC7"'
		FileGet(FromFile, Cnt)
		myDieWAV = ""
		For c = 1 To Cnt
			myDieWAV = myDieWAV & " "
		Next c
		'UPGRADE_WARNING: Get was upgraded to FileGet and has a new behavior. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="9B7D5ADD-D8FE-4819-A36C-6DEDAF088CC7"'
		FileGet(FromFile, myDieWAV)
		'UPGRADE_WARNING: Get was upgraded to FileGet and has a new behavior. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="9B7D5ADD-D8FE-4819-A36C-6DEDAF088CC7"'
		FileGet(FromFile, Cnt)
		myAttackWAV = ""
		For c = 1 To Cnt
			myAttackWAV = myAttackWAV & " "
		Next c
		'UPGRADE_WARNING: Get was upgraded to FileGet and has a new behavior. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="9B7D5ADD-D8FE-4819-A36C-6DEDAF088CC7"'
		FileGet(FromFile, myAttackWAV)
		'UPGRADE_WARNING: Get was upgraded to FileGet and has a new behavior. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="9B7D5ADD-D8FE-4819-A36C-6DEDAF088CC7"'
		FileGet(FromFile, Cnt)
		myHitWAV = ""
		For c = 1 To Cnt
			myHitWAV = myHitWAV & " "
		Next c
		'UPGRADE_WARNING: Get was upgraded to FileGet and has a new behavior. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="9B7D5ADD-D8FE-4819-A36C-6DEDAF088CC7"'
		FileGet(FromFile, myHitWAV)
		'UPGRADE_WARNING: Get was upgraded to FileGet and has a new behavior. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="9B7D5ADD-D8FE-4819-A36C-6DEDAF088CC7"'
		FileGet(FromFile, myFlagsWAV)
		' Read Items
		Reading = "Item"
		'UPGRADE_WARNING: Get was upgraded to FileGet and has a new behavior. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="9B7D5ADD-D8FE-4819-A36C-6DEDAF088CC7"'
		FileGet(FromFile, Cnt)
		For c = 1 To Cnt
			myItem = New Item
			myItem.ReadFromFile(FromFile)
			myItems.Add(myItem, "I" & myItem.Index)
		Next c
		' Read Triggers
		Reading = "Trigger"
		'UPGRADE_WARNING: Get was upgraded to FileGet and has a new behavior. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="9B7D5ADD-D8FE-4819-A36C-6DEDAF088CC7"'
		FileGet(FromFile, Cnt)
		For c = 1 To Cnt
			myTrigger = New Trigger
			myTrigger.ReadFromFile(FromFile)
			myTriggers.Add(myTrigger, "T" & myTrigger.Index)
		Next c
		' Read Conversations
		Reading = "Conversation"
		'UPGRADE_WARNING: Get was upgraded to FileGet and has a new behavior. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="9B7D5ADD-D8FE-4819-A36C-6DEDAF088CC7"'
		FileGet(FromFile, myCurrentConvo)
		'UPGRADE_WARNING: Get was upgraded to FileGet and has a new behavior. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="9B7D5ADD-D8FE-4819-A36C-6DEDAF088CC7"'
		FileGet(FromFile, Cnt)
		For c = 1 To Cnt
			ConversationX = New Conversation
			ConversationX.ReadFromFile(FromFile)
			myConversations.Add(ConversationX, "C" & ConversationX.Index)
		Next c
		Exit Sub
ErrorHandler: 
		' [Titi 2.4.9] re-index components if necessary
		If Err.Number = 457 And Reading <> "" Then
			Select Case Reading
				Case "Item"
					oErr.logError(Reading & " index conflict: " & myItem.Index)
				Case "Trigger"
					oErr.logError(Reading & " index conflict: " & myTrigger.Index)
				Case "Conversation"
					oErr.logError(Reading & " index conflict: " & ConversationX.Index)
			End Select
		Else
			oErr.logError("Cannot read creature" & IIf(Reading <> vbNullString, "'s " & Reading, "") & ", error#" & Err.Number & " (" & Err.Description & ")")
		End If
	End Sub
	
	Public Sub SaveToFile(ByRef ToFile As Short)
		Dim c As Short
		myVersion = 249 ' [Titi 2.4.9]
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
		'UPGRADE_WARNING: Put was upgraded to FilePut and has a new behavior. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="9B7D5ADD-D8FE-4819-A36C-6DEDAF088CC7"'
		FilePut(ToFile, Len(myRace))
		'UPGRADE_WARNING: Put was upgraded to FilePut and has a new behavior. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="9B7D5ADD-D8FE-4819-A36C-6DEDAF088CC7"'
		FilePut(ToFile, myRace)
		'UPGRADE_WARNING: Put was upgraded to FilePut and has a new behavior. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="9B7D5ADD-D8FE-4819-A36C-6DEDAF088CC7"'
		FilePut(ToFile, Len(myHome)) ' [Titi 2.4.9]
		'UPGRADE_WARNING: Put was upgraded to FilePut and has a new behavior. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="9B7D5ADD-D8FE-4819-A36C-6DEDAF088CC7"'
		FilePut(ToFile, myHome) ' [Titi 2.4.9]
		'UPGRADE_WARNING: Put was upgraded to FilePut and has a new behavior. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="9B7D5ADD-D8FE-4819-A36C-6DEDAF088CC7"'
		FilePut(ToFile, myFlags)
		'UPGRADE_WARNING: Put was upgraded to FilePut and has a new behavior. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="9B7D5ADD-D8FE-4819-A36C-6DEDAF088CC7"'
		FilePut(ToFile, myStatus)
		'UPGRADE_WARNING: Put was upgraded to FilePut and has a new behavior. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="9B7D5ADD-D8FE-4819-A36C-6DEDAF088CC7"'
		FilePut(ToFile, myEternium)
		'UPGRADE_WARNING: Put was upgraded to FilePut and has a new behavior. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="9B7D5ADD-D8FE-4819-A36C-6DEDAF088CC7"'
		FilePut(ToFile, Len(myPictureFile))
		'UPGRADE_WARNING: Put was upgraded to FilePut and has a new behavior. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="9B7D5ADD-D8FE-4819-A36C-6DEDAF088CC7"'
		FilePut(ToFile, myPictureFile)
		'UPGRADE_WARNING: Put was upgraded to FilePut and has a new behavior. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="9B7D5ADD-D8FE-4819-A36C-6DEDAF088CC7"'
		FilePut(ToFile, myFaceLeft)
		'UPGRADE_WARNING: Put was upgraded to FilePut and has a new behavior. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="9B7D5ADD-D8FE-4819-A36C-6DEDAF088CC7"'
		FilePut(ToFile, myFaceTop)
		'UPGRADE_WARNING: Put was upgraded to FilePut and has a new behavior. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="9B7D5ADD-D8FE-4819-A36C-6DEDAF088CC7"'
		FilePut(ToFile, myMapX)
		'UPGRADE_WARNING: Put was upgraded to FilePut and has a new behavior. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="9B7D5ADD-D8FE-4819-A36C-6DEDAF088CC7"'
		FilePut(ToFile, myMapY)
		'UPGRADE_WARNING: Put was upgraded to FilePut and has a new behavior. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="9B7D5ADD-D8FE-4819-A36C-6DEDAF088CC7"'
		FilePut(ToFile, myTileSpot)
		For c = 0 To 7
			'UPGRADE_WARNING: Put was upgraded to FilePut and has a new behavior. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="9B7D5ADD-D8FE-4819-A36C-6DEDAF088CC7"'
			FilePut(ToFile, myResistance(c))
			'UPGRADE_WARNING: Put was upgraded to FilePut and has a new behavior. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="9B7D5ADD-D8FE-4819-A36C-6DEDAF088CC7"'
			FilePut(ToFile, myBodyType(c))
		Next c
		'UPGRADE_WARNING: Put was upgraded to FilePut and has a new behavior. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="9B7D5ADD-D8FE-4819-A36C-6DEDAF088CC7"'
		FilePut(ToFile, myResistance1)
		'UPGRADE_WARNING: Put was upgraded to FilePut and has a new behavior. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="9B7D5ADD-D8FE-4819-A36C-6DEDAF088CC7"'
		FilePut(ToFile, myResistance2)
		'UPGRADE_WARNING: Put was upgraded to FilePut and has a new behavior. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="9B7D5ADD-D8FE-4819-A36C-6DEDAF088CC7"'
		FilePut(ToFile, myResistance3)
		'UPGRADE_WARNING: Put was upgraded to FilePut and has a new behavior. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="9B7D5ADD-D8FE-4819-A36C-6DEDAF088CC7"'
		FilePut(ToFile, myResistance4)
		'UPGRADE_WARNING: Put was upgraded to FilePut and has a new behavior. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="9B7D5ADD-D8FE-4819-A36C-6DEDAF088CC7"'
		FilePut(ToFile, myStrength)
		'UPGRADE_WARNING: Put was upgraded to FilePut and has a new behavior. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="9B7D5ADD-D8FE-4819-A36C-6DEDAF088CC7"'
		FilePut(ToFile, myWill)
		'UPGRADE_WARNING: Put was upgraded to FilePut and has a new behavior. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="9B7D5ADD-D8FE-4819-A36C-6DEDAF088CC7"'
		FilePut(ToFile, myAgility)
		'UPGRADE_WARNING: Put was upgraded to FilePut and has a new behavior. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="9B7D5ADD-D8FE-4819-A36C-6DEDAF088CC7"'
		FilePut(ToFile, myNotUsed2)
		'UPGRADE_WARNING: Put was upgraded to FilePut and has a new behavior. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="9B7D5ADD-D8FE-4819-A36C-6DEDAF088CC7"'
		FilePut(ToFile, myExperiencePoints)
		'UPGRADE_WARNING: Put was upgraded to FilePut and has a new behavior. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="9B7D5ADD-D8FE-4819-A36C-6DEDAF088CC7"'
		FilePut(ToFile, mySkillPoints)
		'UPGRADE_WARNING: Put was upgraded to FilePut and has a new behavior. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="9B7D5ADD-D8FE-4819-A36C-6DEDAF088CC7"'
		FilePut(ToFile, myLevel)
		'UPGRADE_WARNING: Put was upgraded to FilePut and has a new behavior. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="9B7D5ADD-D8FE-4819-A36C-6DEDAF088CC7"'
		FilePut(ToFile, myHPNow)
		'UPGRADE_WARNING: Put was upgraded to FilePut and has a new behavior. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="9B7D5ADD-D8FE-4819-A36C-6DEDAF088CC7"'
		FilePut(ToFile, myHPMax)
		'UPGRADE_WARNING: Put was upgraded to FilePut and has a new behavior. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="9B7D5ADD-D8FE-4819-A36C-6DEDAF088CC7"'
		FilePut(ToFile, myLunacy)
		'UPGRADE_WARNING: Put was upgraded to FilePut and has a new behavior. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="9B7D5ADD-D8FE-4819-A36C-6DEDAF088CC7"'
		FilePut(ToFile, myRevelry)
		'UPGRADE_WARNING: Put was upgraded to FilePut and has a new behavior. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="9B7D5ADD-D8FE-4819-A36C-6DEDAF088CC7"'
		FilePut(ToFile, myWrath)
		'UPGRADE_WARNING: Put was upgraded to FilePut and has a new behavior. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="9B7D5ADD-D8FE-4819-A36C-6DEDAF088CC7"'
		FilePut(ToFile, myPride)
		'UPGRADE_WARNING: Put was upgraded to FilePut and has a new behavior. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="9B7D5ADD-D8FE-4819-A36C-6DEDAF088CC7"'
		FilePut(ToFile, myGreed)
		'UPGRADE_WARNING: Put was upgraded to FilePut and has a new behavior. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="9B7D5ADD-D8FE-4819-A36C-6DEDAF088CC7"'
		FilePut(ToFile, myLust)
		'UPGRADE_WARNING: Put was upgraded to FilePut and has a new behavior. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="9B7D5ADD-D8FE-4819-A36C-6DEDAF088CC7"'
		FilePut(ToFile, myCombatAction1)
		'UPGRADE_WARNING: Put was upgraded to FilePut and has a new behavior. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="9B7D5ADD-D8FE-4819-A36C-6DEDAF088CC7"'
		FilePut(ToFile, myCombatAction2)
		'UPGRADE_WARNING: Put was upgraded to FilePut and has a new behavior. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="9B7D5ADD-D8FE-4819-A36C-6DEDAF088CC7"'
		FilePut(ToFile, myFamily)
		'UPGRADE_WARNING: Put was upgraded to FilePut and has a new behavior. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="9B7D5ADD-D8FE-4819-A36C-6DEDAF088CC7"'
		FilePut(ToFile, mySize)
		' Save WAV File names
		'UPGRADE_WARNING: Put was upgraded to FilePut and has a new behavior. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="9B7D5ADD-D8FE-4819-A36C-6DEDAF088CC7"'
		FilePut(ToFile, Len(myMoveWAV))
		'UPGRADE_WARNING: Put was upgraded to FilePut and has a new behavior. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="9B7D5ADD-D8FE-4819-A36C-6DEDAF088CC7"'
		FilePut(ToFile, myMoveWAV)
		'UPGRADE_WARNING: Put was upgraded to FilePut and has a new behavior. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="9B7D5ADD-D8FE-4819-A36C-6DEDAF088CC7"'
		FilePut(ToFile, Len(myDieWAV))
		'UPGRADE_WARNING: Put was upgraded to FilePut and has a new behavior. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="9B7D5ADD-D8FE-4819-A36C-6DEDAF088CC7"'
		FilePut(ToFile, myDieWAV)
		'UPGRADE_WARNING: Put was upgraded to FilePut and has a new behavior. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="9B7D5ADD-D8FE-4819-A36C-6DEDAF088CC7"'
		FilePut(ToFile, Len(myAttackWAV))
		'UPGRADE_WARNING: Put was upgraded to FilePut and has a new behavior. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="9B7D5ADD-D8FE-4819-A36C-6DEDAF088CC7"'
		FilePut(ToFile, myAttackWAV)
		'UPGRADE_WARNING: Put was upgraded to FilePut and has a new behavior. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="9B7D5ADD-D8FE-4819-A36C-6DEDAF088CC7"'
		FilePut(ToFile, Len(myHitWAV))
		'UPGRADE_WARNING: Put was upgraded to FilePut and has a new behavior. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="9B7D5ADD-D8FE-4819-A36C-6DEDAF088CC7"'
		FilePut(ToFile, myHitWAV)
		'UPGRADE_WARNING: Put was upgraded to FilePut and has a new behavior. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="9B7D5ADD-D8FE-4819-A36C-6DEDAF088CC7"'
		FilePut(ToFile, myFlagsWAV)
		' Save Items
		'UPGRADE_WARNING: Put was upgraded to FilePut and has a new behavior. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="9B7D5ADD-D8FE-4819-A36C-6DEDAF088CC7"'
		FilePut(ToFile, myItems.Count())
		For c = 1 To myItems.Count()
			'UPGRADE_WARNING: Couldn't resolve default property of object myItems().SaveToFile. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
			myItems.Item(c).SaveToFile(ToFile)
		Next c
		' Save Triggers
		'UPGRADE_WARNING: Put was upgraded to FilePut and has a new behavior. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="9B7D5ADD-D8FE-4819-A36C-6DEDAF088CC7"'
		FilePut(ToFile, myTriggers.Count())
		For c = 1 To myTriggers.Count()
			'UPGRADE_WARNING: Couldn't resolve default property of object myTriggers().SaveToFile. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
			myTriggers.Item(c).SaveToFile(ToFile)
		Next c
		' Save Conversations
		'UPGRADE_WARNING: Put was upgraded to FilePut and has a new behavior. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="9B7D5ADD-D8FE-4819-A36C-6DEDAF088CC7"'
		FilePut(ToFile, myCurrentConvo)
		'UPGRADE_WARNING: Put was upgraded to FilePut and has a new behavior. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="9B7D5ADD-D8FE-4819-A36C-6DEDAF088CC7"'
		FilePut(ToFile, myConversations.Count())
		For c = 1 To myConversations.Count()
			'UPGRADE_WARNING: Couldn't resolve default property of object myConversations().SaveToFile. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
			myConversations.Item(c).SaveToFile(ToFile)
		Next c
	End Sub
	
	'UPGRADE_NOTE: Class_Initialize was upgraded to Class_Initialize_Renamed. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="A9E4979A-37FA-4718-9994-97DD76ED70A7"'
	Private Sub Class_Initialize_Renamed()
		myName = "Creature"
		Me.DMControlled = True
		Me.Agressive = True
		Me.Guard = True
		Me.Friendly = False
		Me.Male = True
		Me.Family(0) = True
		myHPNow = 25 : myHPMax = 25
		myStrength = 10 : myWill = 10 : myAgility = 10
		myLunacy = 18 : myRevelry = 18 : myWrath = 18 : myPride = 18 : myGreed = 18 : myLust = 18
		myLevel = 1
		myMaxStep = 0
		myNextStep = 0
		myBodyType(0) = 2
		myBodyType(1) = 2
		myBodyType(2) = 2
		myBodyType(3) = 2
		myBodyType(4) = 4
		myBodyType(5) = 4
		myBodyType(6) = 3
		myBodyType(7) = 10
	End Sub
	Public Sub New()
		MyBase.New()
		Class_Initialize_Renamed()
	End Sub
End Class