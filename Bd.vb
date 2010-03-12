Option Strict Off
Option Explicit On
Imports VB = Microsoft.VisualBasic
Module modBD
	' API function declarations
	Declare Sub Sleep Lib "kernel32" (ByVal dwMilliseconds As Integer)
	
	Private Declare Function ShellExecute Lib "shell32.dll"  Alias "ShellExecuteA"(ByVal hWnd As Integer, ByVal lpOperation As String, ByVal lpFile As String, ByVal lpParameters As String, ByVal lpDirectory As String, ByVal nShowCmd As Integer) As Integer
	
	Public gAppPath As String
    Public gDataPath As String = "Data" ' [Titi 2.4.9]
	Public gDebug As Boolean
	' [Titi 2.4.7] added for world settings creation option
	Public WorldNow As World
	Public strKingdomNames() As String
	Public strKingdomSet() As String
	Public strKingdomTemplate() As String
	Public strKingdomLocation() As String
	Public strKingdomPictures() As String
	' [Titi 2.4.8] added to allow changing the names of the runes
	Public strRunesList As String
	Public intNbRunes As Short
	
	' Reference to current Tome
	Public Tome As New Tome
	' Reference to current Plot
	Public Area As New Area
	
	' Arrays for StatementToText processing
	Public List(5, 255) As String
	Public ListIndex(255) As Short
	Public Const bdStatementCount As Short = 74
	Public Const bdContextCount As Short = 32
	
	' Width and Height of combat window
	Public Const bdCombatWidth As Short = 14 ' One less than total width
	Public Const bdCombatHeight As Short = 13 ' One less than total height
	
	' Bits for flag checks
	Public Bit(19) As Single
	Public Const bdInt As Short = 0
	Public Const bdByte As Short = 1
	Public Const bdLong As Short = 2
	
	' Game Options
	Public GlobalGameSpeed As Short
	Public GlobalMusicState As Short
	Public GlobalMusicRandom As Short
	Public GlobalMusicName As String
	Public GlobalDebugMode As Short
	Public GlobalRightClick As Short
	Public GlobalDiceRolling As Short
	Public GlobalOverSwing As Short
	Public GlobalMouseClick As Short
	Public GlobalAutoEndTurn As Short
	Public GlobalFastMove As Short
	Public GlobalScreenX As Short
	Public GlobalScreenY As Short
	Public GlobalScreenColor As Short
    Public GlobalInterfaceName As String = "Standard"
	Public GlobalDiceName As String
	Public GlobalLicNumber1 As String
	Public GlobalLicNumber2 As String
	Public GlobalLicNumber3 As String
	Public GlobalCredits As String
	
	' Option value for Playing WAV Files
	Public GlobalWAVState As Short
	
	' Constants for event processing
	Public Const bdNbTriggers As Short = 78
	Public Const bdNone As Short = 0
	Public Const bdPostTurnCycle As Short = 1
	Public Const bdPrePickLock As Short = 2
	Public Const bdPreUnlock As Short = 3
	Public Const bdPreOpen As Short = 4
	Public Const bdPostEnterEncounter As Short = 5
	Public Const bdPreSearch As Short = 6
	Public Const bdPostSearch As Short = 7
	Public Const bdPreTake As Short = 8
	Public Const bdPostTake As Short = 9
	Public Const bdPreExitTome As Short = 10
	Public Const bdPostEnterTome As Short = 11
	Public Const bdPreSplit As Short = 12
	Public Const bdPostSplit As Short = 13
	Public Const bdPreDropItem As Short = 14
	Public Const bdPostDropItem As Short = 15
	Public Const bdPreExamine As Short = 16
	Public Const bdPostExamine As Short = 17
	Public Const bdPreCombine As Short = 18
	Public Const bdOnCombine As Short = 19
	Public Const bdPostCombine As Short = 20
	Public Const bdPreSearchTraps As Short = 21
	Public Const bdPostSearchTraps As Short = 22
	Public Const bdPreReady As Short = 23
	Public Const bdPostReady As Short = 24
	Public Const bdPreUnReady As Short = 25
	Public Const bdPostUnReady As Short = 26
	Public Const bdPreUseOnCreature As Short = 27
	Public Const bdOnUseOnCreature As Short = 28
	Public Const bdPostUseOnCreature As Short = 29
	Public Const bdPreUseOnItem As Short = 30
	Public Const bdOnUseOnItem As Short = 31
	Public Const bdPostUseOnItem As Short = 32
	Public Const bdPreUseInEncounter As Short = 33
	Public Const bdOnUseInEncounter As Short = 34
	Public Const bdPostUseInEncounter As Short = 35
	Public Const bdPreTopic As Short = 36
	Public Const bdPostTopic As Short = 37
	Public Const bdOnAttack As Short = 38
	Public Const bdOnSkillUse As Short = 39
	Public Const bdOnIgnore As Short = 40
	Public Const bdOnCast As Short = 41
	Public Const bdOnListen As Short = 42
	Public Const bdPreAttacked As Short = 43
	Public Const bdPostAttacked As Short = 44
	Public Const bdPreRollAttack As Short = 45
	Public Const bdPostRollAttack As Short = 46
	Public Const bdPreRollArmorChit As Short = 47
	Public Const bdPostRollArmorChit As Short = 48
	Public Const bdPreRollDamage As Short = 49
	Public Const bdPostRollDamage As Short = 50
	Public Const bdPreApplyDamage As Short = 51
	Public Const bdPostApplyDamage As Short = 52
	Public Const bdPreDeath As Short = 53
	Public Const bdPostDeath As Short = 54
	Public Const bdPreCombatMove As Short = 55
	Public Const bdPostCombatMove As Short = 56
	Public Const bdPostStepEncounter As Short = 57
	Public Const bdPostCombat As Short = 58
	Public Const bdPreTurn As Short = 59
	Public Const bdPostTurn As Short = 60
	Public Const bdPreEnterEncounter As Short = 61
	Public Const bdPreRaiseSkill As Short = 62
	Public Const bdPostRaiseSkill As Short = 63
	Public Const bdPreSellItem As Short = 64
	Public Const bdPostSellItem As Short = 65
	Public Const bdPreBuyItem As Short = 66
	Public Const bdPostBuyItem As Short = 67
	Public Const bdPreCombat As Short = 68
	Public Const bdPostNewCharacter As Short = 69
	Public Const bdPostAttack As Short = 70
	Public Const bdPreCombatStep As Short = 71
	Public Const bdPostCombatStep As Short = 72
	Public Const bdOnSkillTarget As Short = 73
	Public Const bdOnCastTarget As Short = 74
	Public Const bdPreCast As Short = 75
	Public Const bdOnStatus As Short = 76
	Public Const bdOnRealTime As Short = 77
	Public Const bdOnEnterHome As Short = 78 ' [Titi 2.4.9]
	
	'[borfaux] Added to 2.4.6
	' Allows the user to change which data directory to use.
	' Example: /DataPath C:\Runesword
	'      Or: /DataPath C:\Runesword\Source
	Public Sub ParseCommandLineArgs()
		Dim CommandArgs() As String
		Dim ArgText As Object
		gAppPath = My.Application.Info.DirectoryPath
		GlobalDebugMode = False
		CommandArgs = Split(VB.Command(), "/")
		If InStr(1, VB.Command(), "/DataPath ") Then
			For	Each ArgText In CommandArgs
				'UPGRADE_WARNING: Couldn't resolve default property of object ArgText. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
				If Left(ArgText, 9) = "DataPath " And Len(ArgText) > 9 Then
					'UPGRADE_WARNING: Couldn't resolve default property of object ArgText. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
					gAppPath = Mid(ArgText, 10)
					Exit For
				End If
			Next ArgText
		End If
		If InStr(1, VB.Command(), "/DebugOut") Then
			GlobalDebugMode = True
		End If
	End Sub
	
	Public Function HyperJump(ByVal URL As String) As Integer
		HyperJump = ShellExecute(0, vbNullString, URL, vbNullString, vbNullString, AppWinStyle.NormalFocus)
	End Function
	
	Public Function TriggerName(ByVal Index As Short) As String
		Select Case Index
			Case bdNone
				TriggerName = "<None>"
			Case bdPostTurnCycle
				TriggerName = "Post-TurnCycle"
			Case bdPrePickLock
				TriggerName = "Pre-PickLock"
			Case bdPreUnlock
				TriggerName = "Pre-Unlock"
			Case bdPreOpen
				TriggerName = "Pre-Open"
			Case bdPostEnterEncounter
				TriggerName = "Post-EnterEncounter"
			Case bdPreSearch
				TriggerName = "Pre-Search"
			Case bdPostSearch
				TriggerName = "Post-Search"
			Case bdPreTake
				TriggerName = "Pre-Take"
			Case bdPostTake
				TriggerName = "Post-Take"
			Case bdPreExitTome
				TriggerName = "Pre-ExitTome"
			Case bdPostEnterTome
				TriggerName = "Post-EnterTome"
			Case bdPreSplit
				TriggerName = "Pre-Split"
			Case bdPostSplit
				TriggerName = "Post-Split"
			Case bdPreDropItem
				TriggerName = "Pre-DropItem"
			Case bdPostDropItem
				TriggerName = "Post-DropItem"
			Case bdPreExamine
				TriggerName = "Pre-Examine"
			Case bdPostExamine
				TriggerName = "Post-Examine"
			Case bdPreCombine
				TriggerName = "Pre-Combine"
			Case bdOnCombine
				TriggerName = "On-Combine"
			Case bdPostCombine
				TriggerName = "Post-Combine"
			Case bdPreSearchTraps
				TriggerName = "Pre-SearchTraps"
			Case bdPostSearchTraps
				TriggerName = "Post-SearchTraps"
			Case bdPreReady
				TriggerName = "Pre-Equip"
			Case bdPostReady
				TriggerName = "Post-Equip"
			Case bdPreUnReady
				TriggerName = "Pre-UnEquip"
			Case bdPostUnReady
				TriggerName = "Post-UnEquip"
			Case bdPreUseOnCreature
				TriggerName = "Pre-TargetOfUseOnCreature"
			Case bdOnUseOnCreature
				TriggerName = "On-UseOnCreature"
			Case bdPostUseOnCreature
				TriggerName = "Post-TargetOfUseOnCreature"
			Case bdPreUseOnItem
				TriggerName = "Pre-TargetOfUseOnItem"
			Case bdOnUseOnItem
				TriggerName = "On-UseOnItem"
			Case bdPostUseOnItem
				TriggerName = "Post-TargetOfUseOnItem"
			Case bdPreUseInEncounter
				TriggerName = "Pre-TargetOfUseInEncounter"
			Case bdOnUseInEncounter
				TriggerName = "On-UseInEncounter"
			Case bdPostUseInEncounter
				TriggerName = "Post-TargetOfUseInEncounter"
			Case bdPreTopic
				TriggerName = "Pre-Topic"
			Case bdPostTopic
				TriggerName = "Post-Topic"
			Case bdOnAttack
				TriggerName = "On-Attack"
			Case bdOnSkillUse
				TriggerName = "On-SkillUse"
			Case bdOnIgnore
				TriggerName = "On-Ignore"
			Case bdOnCast
				TriggerName = "On-Cast"
			Case bdOnListen
				TriggerName = "On-Listen"
			Case bdPreAttacked
				TriggerName = "Pre-TargetOfAttack"
			Case bdPostAttacked
				TriggerName = "Post-TargetOfAttack"
			Case bdPreRollAttack
				TriggerName = "Pre-RollAttack"
			Case bdPostRollAttack
				TriggerName = "Post-RollAttack"
			Case bdPreRollArmorChit
				TriggerName = "Pre-RollArmorChit"
			Case bdPostRollArmorChit
				TriggerName = "Post-RollArmorChit"
			Case bdPreRollDamage
				TriggerName = "Pre-RollDamage"
			Case bdPostRollDamage
				TriggerName = "Post-RollDamage"
			Case bdPreApplyDamage
				TriggerName = "Pre-TargetOfApplyDamage"
			Case bdPostApplyDamage
				TriggerName = "Post-TargetOfApplyDamage"
			Case bdPreDeath
				TriggerName = "Pre-Death"
			Case bdPostDeath
				TriggerName = "Post-Death"
			Case bdPreCombatMove
				TriggerName = "Pre-CombatMove"
			Case bdPostCombatMove
				TriggerName = "Post-CombatMove"
			Case bdPostStepEncounter
				TriggerName = "Post-StepEncounter"
			Case bdPreCombat
				TriggerName = "Pre-Combat"
			Case bdPostCombat
				TriggerName = "Post-Combat"
			Case bdPreTurn
				TriggerName = "Pre-Turn"
			Case bdPostTurn
				TriggerName = "Post-Turn"
			Case bdPreEnterEncounter
				TriggerName = "Pre-EnterEncounter"
			Case bdPreRaiseSkill
				TriggerName = "Pre-RaiseSkill"
			Case bdPostRaiseSkill
				TriggerName = "Post-RaiseSkill"
			Case bdPreSellItem
				TriggerName = "Pre-SellItem"
			Case bdPostSellItem
				TriggerName = "Post-SellItem"
			Case bdPreBuyItem
				TriggerName = "Pre-BuyItem"
			Case bdPostBuyItem
				TriggerName = "Post-BuyItem"
			Case bdPreCombat
				TriggerName = "Pre-Combat"
			Case bdPostNewCharacter
				TriggerName = "Post-NewCharacter"
			Case bdPostAttack
				TriggerName = "Post-Attack"
			Case bdPreCombatStep
				TriggerName = "Pre-CombatStep"
			Case bdPostCombatStep
				TriggerName = "Post-CombatStep"
			Case bdOnSkillTarget
				TriggerName = "On-SkillTarget"
			Case bdOnCastTarget
				TriggerName = "On-CastTarget"
			Case bdPreCast
				TriggerName = "Pre-Cast"
			Case bdOnStatus
				TriggerName = "On-Status"
			Case bdOnRealTime
				TriggerName = "On-RealTime"
			Case bdOnEnterHome
				TriggerName = "On-EnterHome"
		End Select
	End Function
	
	Public Function OpToText(ByRef Op As Short) As String
		Select Case Op
			Case 0 ' =
				OpToText = "="
			Case 1 ' +
				OpToText = "+"
			Case 2 ' -
				OpToText = "-"
			Case 3 ' *
				OpToText = "*"
			Case 4 ' /
				OpToText = "/"
			Case 5 ' >
				OpToText = ">"
			Case 6 ' <
				OpToText = "<"
			Case 7 ' >=
				OpToText = ">="
			Case 8 ' <=
				OpToText = "<="
			Case 9 ' Or
				OpToText = "Or"
			Case 10 ' And
				OpToText = "And"
			Case 11 ' Xor
				OpToText = "XOr"
			Case 12 ' Like
				OpToText = "Like"
			Case 13 ' <>
				OpToText = "<>"
		End Select
	End Function
	
    Public Sub StatementToList(ByRef TomeX As Tome, ByRef AreaX As Area, ByRef TrigX As Trigger, ByRef StmtX As Statement, ByRef ListIndex() As Short, ByRef List(,) As String)
        Dim c As Short
        Dim Text As String
        Dim AreaZ As Area
        Dim MapX As Map
        Dim CreatureX As Creature
        Dim ItemX As Item
        Dim TriggerX As Trigger
        Select Case StmtX.Statement
            Case 0 'None
            Case 1 'Label [Context.Var]
            Case 2 'If [Context.Var] [Op] [Context.Var]
            Case 3 'Else
            Case 4 'ElseIf [Context.Var] [Op] [Context.Var]
            Case 5 'And [Context.Var] [Op] [Context.Var]
            Case 6 'Or [Context.Var] [Op] [Context.Var]
            Case 7 'EndIf
            Case 8 'While [Context.Var] [Op] [Context.Var]
            Case 9, 26, 70 'ForEach [List] In [List], Destroy [List] In [List], Find [List] Named 'Text' In [List]
                List(0, 0) = "CreatureA"
                List(0, 1) = "CreatureB"
                List(0, 2) = "CreatureC"
                List(0, 3) = "ItemA"
                List(0, 4) = "ItemB"
                List(0, 5) = "ItemC"
                List(0, 6) = "TriggerA"
                List(0, 7) = "TriggerB"
                List(0, 8) = "TriggerC"
                ListIndex(0) = 8
                If StmtX.Statement = 26 Then
                    List(0, 9) = "CreatureNow"
                    List(0, 10) = "CreatureTarget"
                    List(0, 11) = "ItemNow"
                    List(0, 12) = "ItemTarget"
                    List(0, 13) = "TriggerNow"
                    List(0, 14) = "TriggerTarget"
                    ListIndex(0) = 14
                End If
                List(1, 0) = "EncounterNow"
                List(1, 1) = "EncounterA"
                List(1, 2) = "EncounterB"
                List(1, 3) = "EncounterC"
                List(1, 4) = "EncounterTarget"
                List(1, 5) = ""
                List(1, 6) = ""
                List(1, 7) = ""
                List(1, 8) = ""
                List(1, 9) = ""
                List(1, 10) = ""
                List(1, 11) = "TriggerNow"
                List(1, 12) = "TriggerA"
                List(1, 13) = "TriggerB"
                List(1, 14) = "TriggerC"
                List(1, 15) = "TriggerTarget"
                ListIndex(1) = 15
                If StmtX.B(0) > 2 Then
                    List(1, 5) = "ItemNow"
                    List(1, 6) = "ItemA"
                    List(1, 7) = "ItemB"
                    List(1, 8) = "ItemC"
                    List(1, 9) = "ItemTarget"
                    List(1, 16) = "CreatureNow"
                    List(1, 17) = "CreatureA"
                    List(1, 18) = "CreatureB"
                    List(1, 19) = "CreatureC"
                    List(1, 20) = "CreatureTarget"
                    ListIndex(1) = 20
                End If
                If StmtX.B(0) < 3 Or StmtX.B(0) > 5 Then
                    List(1, 10) = "Party"
                End If
            Case 10 'Next
            Case 11 'Branch [Context.Var]
            Case 12 'Put [Context.Var] [Op] [Context.Var] Into [Context.Var]
            Case 13 'Set [List] = [List]
                List(0, 0) = "CreatureNow"
                List(0, 1) = "CreatureA"
                List(0, 2) = "CreatureB"
                List(0, 3) = "CreatureC"
                List(0, 4) = "CreatureTarget"
                List(0, 5) = "EncounterNow"
                List(0, 6) = "EncounterA"
                List(0, 7) = "EncounterB"
                List(0, 8) = "EncounterC"
                List(0, 9) = "EncounterTarget"
                List(0, 10) = "ItemNow"
                List(0, 11) = "ItemA"
                List(0, 12) = "ItemB"
                List(0, 13) = "ItemC"
                List(0, 14) = "ItemTarget"
                List(0, 15) = "TileNow"
                List(0, 16) = "TileA"
                List(0, 17) = "TileB"
                List(0, 18) = "TileC"
                List(0, 19) = "TileTarget"
                List(0, 20) = "TriggerNow"
                List(0, 21) = "TriggerA"
                List(0, 22) = "TriggerB"
                List(0, 23) = "TriggerC"
                List(0, 24) = "TriggerTarget"
                ListIndex(0) = 24
                Select Case StmtX.B(0)
                    Case 0 To 4 ' Creatures
                        List(1, 0) = "CreatureNow"
                        List(1, 1) = "CreatureA"
                        List(1, 2) = "CreatureB"
                        List(1, 3) = "CreatureC"
                        List(1, 4) = "CreatureTarget"
                    Case 5 To 9 ' Encounters
                        List(1, 0) = "EncounterNow"
                        List(1, 1) = "EncounterA"
                        List(1, 2) = "EncounterB"
                        List(1, 3) = "EncounterC"
                        List(1, 4) = "EncounterTarget"
                    Case 10 To 14 ' Items
                        List(1, 0) = "ItemNow"
                        List(1, 1) = "ItemA"
                        List(1, 2) = "ItemB"
                        List(1, 3) = "ItemC"
                        List(1, 4) = "ItemTarget"
                    Case 15 To 19 ' Tiles
                        List(1, 0) = "TileNow"
                        List(1, 1) = "TileA"
                        List(1, 2) = "TileB"
                        List(1, 3) = "TileC"
                        List(1, 4) = "TileTarget"
                    Case 20 To 24 ' Triggers
                        List(1, 0) = "TriggerNow"
                        List(1, 1) = "TriggerA"
                        List(1, 2) = "TriggerB"
                        List(1, 3) = "TriggerC"
                        List(1, 4) = "TriggerTarget"
                End Select
                ListIndex(1) = 4
            Case 14 'Select [Context.Var]
            Case 15 'Case [Context.Var]
            Case 16 'EndSelect
            Case 17 'Exit [List]
                List(0, 0) = "Trigger"
                List(0, 1) = "Loop"
                List(0, 2) = "Abort"
                ListIndex(0) = 2
            Case 18 'CopyCreature [List] Into [List]
                For c = 0 To 255
                    List(0, c) = ""
                Next c
                ListIndex(0) = 0
                List(0, 0) = Chr(34) & "None" & Chr(34)
                For Each CreatureX In TrigX.Creatures
                    If CreatureX.Index < 256 Then
                        List(0, CreatureX.Index) = Chr(34) & CreatureX.Name & Chr(34)
                        If CreatureX.Index > ListIndex(0) Then
                            ListIndex(0) = CreatureX.Index
                        End If
                    End If
                Next CreatureX
                List(1, 0) = "EncounterNow"
                List(1, 1) = "EncounterA"
                List(1, 2) = "EncounterB"
                List(1, 3) = "EncounterC"
                List(1, 4) = "EncounterTarget"
                List(1, 5) = "Party"
                List(1, 6) = "TriggerNow"
                List(1, 7) = "TriggerA"
                List(1, 8) = "TriggerB"
                List(1, 9) = "TriggerC"
                List(1, 10) = "TriggerTarget"
                ListIndex(1) = 10
            Case 19 'CopyItem [List] Into [List]
                For c = 0 To 255
                    List(0, c) = ""
                Next c
                ListIndex(0) = 0
                List(0, 0) = Chr(34) & "None" & Chr(34)
                For Each ItemX In TrigX.Items
                    If ItemX.Index < 256 Then
                        List(0, ItemX.Index) = Chr(34) & ItemX.Name & Chr(34)
                        If ItemX.Index > ListIndex(0) Then
                            ListIndex(0) = ItemX.Index
                        End If
                    End If
                Next ItemX
                List(1, 0) = "CreatureNow"
                List(1, 1) = "CreatureA"
                List(1, 2) = "CreatureB"
                List(1, 3) = "CreatureC"
                List(1, 4) = "CreatureTarget"
                List(1, 5) = "EncounterNow"
                List(1, 6) = "EncounterA"
                List(1, 7) = "EncounterB"
                List(1, 8) = "EncounterC"
                List(1, 9) = "EncounterTarget"
                List(1, 10) = "ItemNow"
                List(1, 11) = "ItemA"
                List(1, 12) = "ItemB"
                List(1, 13) = "ItemC"
                List(1, 14) = "ItemTarget"
                List(1, 15) = "TriggerNow"
                List(1, 16) = "TriggerA"
                List(1, 17) = "TriggerB"
                List(1, 18) = "TriggerC"
                List(1, 19) = "TriggerTarget"
                ListIndex(1) = 19
            Case 20 'CopyTrigger [List] Into [List]
                For c = 0 To 255
                    List(0, c) = ""
                Next c
                ListIndex(0) = 0
                List(0, 0) = Chr(34) & "None" & Chr(34)
                For Each TriggerX In TrigX.Triggers
                    If TriggerX.Index < 256 Then
                        List(0, TriggerX.Index) = Chr(34) & TriggerX.Name & Chr(34)
                        If TriggerX.Index > ListIndex(0) Then
                            ListIndex(0) = TriggerX.Index
                        End If
                    End If
                Next TriggerX
                List(1, 0) = "CreatureNow"
                List(1, 1) = "CreatureA"
                List(1, 2) = "CreatureB"
                List(1, 3) = "CreatureC"
                List(1, 4) = "CreatureTarget"
                List(1, 5) = "EncounterNow"
                List(1, 6) = "EncounterA"
                List(1, 7) = "EncounterB"
                List(1, 8) = "EncounterC"
                List(1, 9) = "EncounterTarget"
                List(1, 10) = "ItemNow"
                List(1, 11) = "ItemA"
                List(1, 12) = "ItemB"
                List(1, 13) = "ItemC"
                List(1, 14) = "ItemTarget"
                List(1, 15) = "Party"
                List(1, 16) = "TriggerNow"
                List(1, 17) = "TriggerA"
                List(1, 18) = "TriggerB"
                List(1, 19) = "TriggerC"
                List(1, 20) = "TriggerTarget"
                ListIndex(1) = 20
            Case 21 'CopyTile 'Text' To ([Context.Var],[Context.Var],[List])
                List(0, 0) = "Bottom"
                List(0, 1) = "BottomFlip"
                List(0, 2) = "Middle"
                List(0, 3) = "MiddleFlip"
                List(0, 4) = "Top"
                List(0, 5) = "TopFlip"
                ListIndex(0) = 5
            Case 22 'CopyText 'Text' Into [Context.Var]
            Case 23, 42 'Runes ([List],[List],[List],[List],[List],[List]) Fail Save, SorceryQueRune ([List],[List],[List],[List],[List],[List]) Check
                List(0, 0) = ""
                ' [Titi 2.4.8] get the runes names
                Text = Right(strRunesList, Len(strRunesList) - 5) ' get rid of "List="
                For c = 1 To intNbRunes - 1
                    List(0, c) = Left(Text, InStr(Text, ",") - 1)
                    Text = Right(Text, Len(Text) - Len(List(0, c)) - 1)
                Next
                List(0, c) = Text
                ListIndex(0) = intNbRunes
                '            List(0, 1) = "Blood"
                '            List(0, 2) = "Bile"
                '            List(0, 3) = "Oil"
                '            List(0, 4) = "Nectar"
                '            List(0, 5) = "Fire"
                '            List(0, 6) = "Earth"
                '            List(0, 7) = "Water"
                '            List(0, 8) = "Air"
                '            List(0, 9) = "Time"
                '            List(0, 10) = "Moons"
                '            List(0, 11) = "Suns"
                '            List(0, 12) = "Space"
                '            List(0, 13) = "Insect"
                '            List(0, 14) = "Man"
                '            List(0, 15) = "Fish"
                '            List(0, 16) = "Animal"
                '            List(0, 17) = "Twilight"
                '            List(0, 18) = "Abyss"
                '            List(0, 19) = "Dreams"
                '            List(0, 20) = "Eternium"
                '            ListIndex(0) = 20
                If StmtX.Statement = 23 Then
                    List(1, 0) = ""
                    List(1, 1) = "Fail"
                    List(1, 2) = "Save"
                    List(1, 3) = "FailSave"
                    ListIndex(1) = 3
                End If
            Case 24 'AddFactoid 'Text'
            Case 25 'AddJournalEntry 'Text'
            Case 27 'CombatApplyDamage [Context.Var] To [List]
                List(0, 0) = "CreatureNow"
                List(0, 1) = "CreatureA"
                List(0, 2) = "CreatureB"
                List(0, 3) = "CreatureC"
                List(0, 4) = "CreatureTarget"
                ListIndex(0) = 4
            Case 28, 29 'CombatRollAttack [List] Bonus [Context.Var], CombatRollDamage [Context.Var] As [List] Check
                List(0, 0) = "Normal"
                List(0, 1) = "Sharp"
                List(0, 2) = "Blunt"
                List(0, 3) = "Cold"
                List(0, 4) = "Fire"
                List(0, 5) = "Evil"
                List(0, 6) = "Holy"
                List(0, 7) = "Magic"
                List(0, 8) = "Mind"
                ListIndex(0) = 8
            Case 30 'DestroyFactoid 'Text'
            Case 31 'TargetCreature [List] As [List] Within [Context.Var] Dead
                List(0, 0) = "CreatureA"
                List(0, 1) = "CreatureB"
                List(0, 2) = "CreatureC"
                List(0, 3) = "CreatureTarget"
                ListIndex(0) = 3
                List(1, 0) = "Any"
                List(1, 1) = "Party"
                List(1, 2) = "EncounterNow"
                ListIndex(1) = 2
            Case 32 'TargetItem [List]
                List(0, 0) = "ItemA"
                List(0, 1) = "ItemB"
                List(0, 2) = "ItemC"
                ListIndex(0) = 2
            Case 33 'PlaySound 'Text' Pause
            Case 34 'PlayMusic 'Text' Pause
            Case 35 'MoveParty To [List] At ([Context.Var],[Context.Var])
                For c = 0 To 255
                    List(0, c) = ""
                Next c
                ListIndex(0) = 0
                List(0, 0) = Chr(34) & "None" & Chr(34)
                For Each MapX In AreaX.Plot.Maps
                    If MapX.Index < 256 Then
                        List(0, MapX.Index) = Chr(34) & MapX.Name & Chr(34)
                        If MapX.Index > ListIndex(0) Then
                            ListIndex(0) = MapX.Index
                        End If
                    End If
                Next MapX
            Case 36 'DialogShow 'Text' Says 'Text' As [List]
                List(0, 0) = "Normal"
                List(0, 1) = "BriefBox"
                List(0, 2) = "ReplyPick"
                List(0, 3) = "ReplyText"
                List(0, 4) = "BriefLine" ' Deprecated, but still supported
                '           List(0, 5) = "BriefLine"
                ListIndex(0) = 4 '5
            Case 37 'DialogReply 'Text'
            Case 38 'DialogAccept [Context.Var]
            Case 39 'DialogDice 'Text' Says 'Text' Rolls [List] Into [Context.Var]
                For c = 0 To 24
                    List(0, c) = Trim(Str(c Mod 5 + 1)) & "d" & Trim(Str(Int((c Mod 25) / 5) * 2 + 4))
                Next c
                List(0, 25) = "1d20"
                ListIndex(0) = 25
            Case 40 'DialogHide
            Case 41 'CutScene [List] Display 'Text' Picture 'Text'
                List(0, 0) = "PicLeftTextRight"
                List(0, 1) = "PicRightTextLeft"
                List(0, 2) = "PicTopTextBottom"
                List(0, 3) = "PicBottomTextTop"
                List(0, 4) = "PicCenter"
                List(0, 5) = "TextCenter"
                ListIndex(0) = 5
            Case 43 'CombatMove [List] [List]
                List(0, 0) = "Toward"
                List(0, 1) = "AwayFrom"
                ListIndex(0) = 1
                List(1, 0) = "Closest"
                List(1, 1) = "Farthest"
                List(1, 2) = "Strongest"
                List(1, 3) = "Weakest"
                List(1, 4) = "Random"
                List(1, 5) = "CreatureTarget"
                ListIndex(1) = 5
            Case 44 'Reserved
            Case 45 'CombatTarget [List]
                List(0, 0) = "Closest"
                List(0, 1) = "Farthest"
                List(0, 2) = "Strongest"
                List(0, 3) = "Weakest"
                List(0, 4) = "Random"
                ListIndex(0) = 4
            Case 46 'PlaySFX 'Text' As ([List],[List],[List],Number)
                List(0, 0) = "Head"
                List(0, 1) = "Center"
                ListIndex(0) = 1
                List(1, 0) = "Fling"
                List(1, 1) = "Stream"
                List(1, 2) = "BurstHere"
                List(1, 3) = "BurstThere"
                List(1, 4) = "Down"
                ListIndex(1) = 4
                List(2, 0) = "Fast"
                List(2, 1) = "Medium"
                List(2, 2) = "Slow"
                ListIndex(2) = 2
                For c = 0 To 64
                    List(3, c) = VB6.Format(c, "00")
                Next c
                ListIndex(3) = 64
            Case 47 'Sorcery [Context.Var]
            Case 48 'DialogBuySell [List] At [Context.Var]
                List(0, 0) = "CreatureNow"
                List(0, 1) = "EncounterNow"
                List(0, 2) = "TriggerNow"
                ListIndex(0) = 2
            Case 49 'Reserved
            Case 50 'ExecTrigger [List]
                For c = 0 To 255
                    List(0, c) = ""
                Next c
                ListIndex(0) = 0
                List(0, 0) = Chr(34) & "None" & Chr(34)
                For Each TriggerX In TrigX.Triggers
                    If TriggerX.Index < 256 Then
                        List(0, TriggerX.Index) = Chr(34) & TriggerX.Name & Chr(34)
                        If TriggerX.Index > ListIndex(0) Then
                            ListIndex(0) = TriggerX.Index
                        End If
                    End If
                Next TriggerX
            Case 51 'DialogAcceptText [Context.Var]
            Case 52 'CombatStart
            Case 53 'Let [Context.Var] = [Context.Var]
            Case 54 '' Text
            Case 55 'RandomizeEncounter 'Text'
            Case 56 'RandomTheme 'Text'
            Case 57 'AwardExperience [Context.Var] To [List]
                List(0, 0) = "Party"
                List(0, 1) = "CreatureNow"
                List(0, 2) = "CreatureA"
                List(0, 3) = "CreatureB"
                List(0, 4) = "CreatureC"
                List(0, 5) = "CreatureTarget"
                ListIndex(0) = 5
            Case 58 'MoveItem [List] From [List] To [List] Copy
                List(0, 0) = "ItemNow"
                List(0, 1) = "ItemA"
                List(0, 2) = "ItemB"
                List(0, 3) = "ItemC"
                List(0, 4) = "ItemTarget"
                ListIndex(0) = 4
                For c = 1 To 2
                    List(c, 0) = "CreatureNow"
                    List(c, 1) = "CreatureA"
                    List(c, 2) = "CreatureB"
                    List(c, 3) = "CreatureC"
                    List(c, 4) = "CreatureTarget"
                    List(c, 5) = "EncounterNow"
                    List(c, 6) = "EncounterA"
                    List(c, 7) = "EncounterB"
                    List(c, 8) = "EncounterC"
                    List(c, 9) = "EncounterTarget"
                    List(c, 10) = "ItemNow"
                    List(c, 11) = "ItemA"
                    List(c, 12) = "ItemB"
                    List(c, 13) = "ItemC"
                    List(c, 14) = "ItemTarget"
                    List(c, 15) = "TriggerNow"
                    List(c, 16) = "TriggerA"
                    List(c, 17) = "TriggerB"
                    List(c, 18) = "TriggerC"
                    List(c, 19) = "TriggerTarget"
                    ListIndex(c) = 19
                Next c
            Case 59 'AddQuest 'Text' As 'Text'
            Case 60 'RemoveQuest 'Text'
            Case 61 'RemoveTopic 'Text'
            Case 62 'CombatAttackWithWeapon
            Case 63 'CombatAttackWithSpecial 'Text' As [Context.Var]
            Case 64 'MovePartyMapName ([Context.Var],[Context.Var],[List],'Text')
                For c = 0 To 255
                    List(0, c) = ""
                Next c
                ListIndex(0) = 0
                List(0, 0) = Chr(34) & "None" & Chr(34)
                For Each AreaZ In TomeX.Areas
                    If AreaZ.Index < 256 Then
                        List(0, AreaZ.Index) = Chr(34) & AreaZ.Name & Chr(34)
                        If AreaZ.Index > ListIndex(0) Then
                            ListIndex(0) = AreaZ.Index
                        End If
                    End If
                Next AreaZ
            Case 65 'MoveCreature [List] From [List] To [List] Copy
                List(0, 0) = "CreatureNow"
                List(0, 1) = "CreatureA"
                List(0, 2) = "CreatureB"
                List(0, 3) = "CreatureC"
                List(0, 4) = "CreatureTarget"
                ListIndex(0) = 4
                For c = 1 To 2
                    List(c, 0) = "CreatureNow"
                    List(c, 1) = "CreatureA"
                    List(c, 2) = "CreatureB"
                    List(c, 3) = "CreatureC"
                    List(c, 4) = "CreatureTarget"
                    List(c, 5) = "EncounterNow"
                    List(c, 6) = "EncounterA"
                    List(c, 7) = "EncounterB"
                    List(c, 8) = "EncounterC"
                    List(c, 9) = "EncounterTarget"
                    List(c, 10) = "ItemNow"
                    List(c, 11) = "ItemA"
                    List(c, 12) = "ItemB"
                    List(c, 13) = "ItemC"
                    List(c, 14) = "ItemTarget"
                    List(c, 15) = "TriggerNow"
                    List(c, 16) = "TriggerA"
                    List(c, 17) = "TriggerB"
                    List(c, 18) = "TriggerC"
                    List(c, 19) = "TriggerTarget"
                    List(c, 20) = "Party"
                    ListIndex(c) = 20
                Next c
            Case 66 'IfText [Context.Var] [Op] 'Text'
            Case 67 'TargetEncounter 'Text'
            Case 68 'TargetTile 'Text' At ([Context.Var],[Context.Var],[List])
                List(0, 0) = "Bottom"
                List(0, 1) = "Middle"
                List(0, 2) = "Top"
                ListIndex(0) = 2
            Case 69 'PlayVideo 'Text' Pause
            Case 71 'CallTrigger [List] In [List]
                List(0, 0) = "TriggerNow"
                List(0, 1) = "TriggerA"
                List(0, 2) = "TriggerB"
                List(0, 3) = "TriggerC"
                List(0, 4) = "TriggerTarget"
                ListIndex(0) = 4
                List(1, 0) = "CreatureNow"
                List(1, 1) = "CreatureA"
                List(1, 2) = "CreatureB"
                List(1, 3) = "CreatureC"
                List(1, 4) = "CreatureTarget"
                List(1, 5) = "EncounterNow"
                List(1, 6) = "EncounterA"
                List(1, 7) = "EncounterB"
                List(1, 8) = "EncounterC"
                List(1, 9) = "EncounterTarget"
                List(1, 10) = "ItemNow"
                List(1, 11) = "ItemA"
                List(1, 12) = "ItemB"
                List(1, 13) = "ItemC"
                List(1, 14) = "ItemTarget"
                List(1, 15) = "Party"
                List(1, 16) = "TriggerNow"
                List(1, 17) = "TriggerA"
                List(1, 18) = "TriggerB"
                List(1, 19) = "TriggerC"
                List(1, 20) = "TriggerTarget"
                ListIndex(1) = 20
            Case 72 'CombatAnimation 'Text' For [List] Frames [List]
                List(0, 0) = "CreatureNow"
                List(0, 1) = "CreatureA"
                List(0, 2) = "CreatureB"
                List(0, 3) = "CreatureC"
                List(0, 4) = "CreatureTarget"
                ListIndex(0) = 4
                For c = 2 To 32
                    List(1, c - 2) = VB6.Format(c, "00")
                Next c
                ListIndex(1) = 30
            Case 73 'MapAnimation 'Text' Frames [List] Level [List] At (Context.Var, Context.Var)
                For c = 2 To 32
                    List(0, c - 2) = VB6.Format(c, "00")
                Next c
                ListIndex(0) = 30
                List(1, 0) = "Bottom"
                List(1, 1) = "BottomFlip"
                List(1, 2) = "Middle"
                List(1, 3) = "MiddleFlip"
                List(1, 4) = "Top"
                List(1, 5) = "TopFlip"
                ListIndex(1) = 5
        End Select
    End Sub
	
	Public Function TextQuote(ByVal Text As String) As String
		' [Titi] 2.4.6 - made the function public
		Dim c As Short
		c = InStr(Text, Chr(34))
		Do While c > 0
			Mid(Text, c, 1) = "'"
			c = InStr(Text, Chr(34))
		Loop 
		TextQuote = Text
	End Function
	
	Public Function StatementToText(ByRef TomeX As Tome, ByRef AreaX As Area, ByRef TrigX As Trigger, ByRef StmtX As Statement, ByRef Text As String) As Short
		Dim c As Short
		StatementToText = True
		StatementToList(TomeX, AreaX, TrigX, StmtX, ListIndex, List)
		Select Case StmtX.Statement
			Case 0 'None
				Text = ""
			Case 1 'Label [Context.Var]
				Text = "Label " & ContextToText(StmtX.B(0)) & "." & VarToText(StmtX.B(0), StmtX.B(1))
			Case 2 'If [Context.Var] [Op] [Context.Var]
				Text = "If " & ContextToText(StmtX.B(0)) & "." & VarToText(StmtX.B(0), StmtX.B(1)) & " " & OpToText(StmtX.B(2)) & " " & ContextToText(StmtX.B(3)) & "." & VarToText(StmtX.B(3), StmtX.B(4))
			Case 3 'Else
				Text = "Else"
			Case 4 'ElseIf [Context.Var] [Op] [Context.Var]
				Text = "ElseIf " & ContextToText(StmtX.B(0)) & "." & VarToText(StmtX.B(0), StmtX.B(1)) & " " & OpToText(StmtX.B(2)) & " " & ContextToText(StmtX.B(3)) & "." & VarToText(StmtX.B(3), StmtX.B(4))
			Case 5 'And [Context.Var] [Op] [Context.Var]
				Text = "And " & ContextToText(StmtX.B(0)) & "." & VarToText(StmtX.B(0), StmtX.B(1)) & " " & OpToText(StmtX.B(2)) & " " & ContextToText(StmtX.B(3)) & "." & VarToText(StmtX.B(3), StmtX.B(4))
			Case 6 'Or [Context.Var] [Op] [Context.Var]
				Text = "Or " & ContextToText(StmtX.B(0)) & "." & VarToText(StmtX.B(0), StmtX.B(1)) & " " & OpToText(StmtX.B(2)) & " " & ContextToText(StmtX.B(3)) & "." & VarToText(StmtX.B(3), StmtX.B(4))
			Case 7 'EndIf
				Text = "EndIf"
			Case 8 'While [Context.Var] [Op] [Context.Var]
				Text = "While " & ContextToText(StmtX.B(0)) & "." & VarToText(StmtX.B(0), StmtX.B(1)) & " " & OpToText(StmtX.B(2)) & " " & ContextToText(StmtX.B(3)) & "." & VarToText(StmtX.B(3), StmtX.B(4))
			Case 9 'ForEach [List] In [List]
				Text = "ForEach " & List(0, StmtX.B(0)) & " In " & List(1, StmtX.B(1))
			Case 10 'Next
				Text = "Next"
			Case 11 'Branch [Context.Var]
				Text = "Branch " & ContextToText(StmtX.B(0)) & "." & VarToText(StmtX.B(0), StmtX.B(1))
			Case 12 'Put [Context.Var] [Op] [Context.Var] Into [Context.Var]
				Text = "Put " & ContextToText(StmtX.B(0)) & "." & VarToText(StmtX.B(0), StmtX.B(1)) & " " & OpToText(StmtX.B(2)) & " " & ContextToText(StmtX.B(3)) & "." & VarToText(StmtX.B(3), StmtX.B(4)) & " Into " & ContextToText(StmtX.B(5)) & "." & VarToText(StmtX.B(5), StmtX.B(6))
			Case 13 'Set [List] = [List]
				Text = "Set " & List(0, StmtX.B(0)) & " = " & List(1, StmtX.B(1))
			Case 14 'Select [Context.Var]
				Text = "Select " & ContextToText(StmtX.B(0)) & "." & VarToText(StmtX.B(0), StmtX.B(1))
			Case 15 'Case [Context.Var]
				Text = "Case " & ContextToText(StmtX.B(0)) & "." & VarToText(StmtX.B(0), StmtX.B(1))
			Case 16 'EndSelect
				Text = "EndSelect"
			Case 17 'Exit [List]
				Text = "Exit " & List(0, StmtX.B(0))
			Case 18 'CopyCreature [List] Into [List]
				Text = "CopyCreature " & List(0, StmtX.B(0)) & " Into " & List(1, StmtX.B(1))
			Case 19 'CopyItem [List] Into [List]
				Text = "CopyItem " & List(0, StmtX.B(0)) & " Into " & List(1, StmtX.B(1))
			Case 20 'CopyTrigger [List] Into [List]
				Text = "CopyTrigger " & List(0, StmtX.B(0)) & " Into " & List(1, StmtX.B(1))
			Case 21 'CopyTile 'Text' Level [List] At ([Context.Var],[Context.Var])
				Text = "CopyTile " & Chr(34) & TextQuote(StmtX.Text) & Chr(34) & " Level " & List(0, StmtX.B(2)) & " At (" & ContextToText(StmtX.B(0)) & "." & VarToText(StmtX.B(0), StmtX.B(1)) & "," & ContextToText(StmtX.B(3)) & "." & VarToText(StmtX.B(3), StmtX.B(4)) & ")"
			Case 22 'CopyText 'Text' Into [Context.Var]
				Text = "CopyText " & Chr(34) & TextQuote(StmtX.Text) & Chr(34) & " Into " & ContextToText(StmtX.B(0)) & "." & VarToText(StmtX.B(0), StmtX.B(1))
			Case 23 'Runes ([List],[List],[List],[List],[List],[List]) Fail Save
				Text = "Runes ("
				For c = 0 To 5
					If StmtX.B(c) > 0 Then
						If c = 0 Then
							Text = Text & List(0, StmtX.B(c))
						Else
							Text = Text & ", " & List(0, StmtX.B(c))
						End If
					End If
				Next c
				Text = Text & ")"
				If StmtX.B(6) > 0 Then
					Text = Text & " " & List(1, StmtX.B(6))
				End If
			Case 24 'AddFactoid 'Text'
				Text = "AddFactoid " & Chr(34) & TextQuote(StmtX.Text) & Chr(34)
			Case 25 'AddJournalEntry 'Text'
				Text = "AddJournalEntry " & Chr(34) & TextQuote(StmtX.Text) & Chr(34)
			Case 26 'Destroy [List] In [List]
				Text = "Destroy " & List(0, StmtX.B(0)) & " In " & List(1, StmtX.B(1))
			Case 27 'CombatApplyDamage [Context.Var] To [List]
				Text = "CombatApplyDamage " & ContextToText(StmtX.B(3)) & "." & VarToText(StmtX.B(3), StmtX.B(4)) & " To " & List(0, StmtX.B(0))
			Case 28 'CombatRollAttack [List] Bonus [Context.Var]
				Text = "CombatRollAttack " & List(0, StmtX.B(5)) & " Bonus " & ContextToText(StmtX.B(0)) & "." & VarToText(StmtX.B(0), StmtX.B(1))
			Case 29 'CombatRollDamage [List] Damage [Context.Var] Check
				Text = "CombatRollDamage " & List(0, StmtX.B(5)) & " Damaged " & ContextToText(StmtX.B(0)) & "." & VarToText(StmtX.B(0), StmtX.B(1))
				If StmtX.B(2) > 0 Then
					Text = Text & " Check"
				End If
			Case 30 'DestroyFactoid 'Text'
				Text = "DestroyFactoid " & Chr(34) & TextQuote(StmtX.Text) & Chr(34)
			Case 31 'TargetCreature [List] As [List] Within [Context.Var] Dead
				Text = "TargetCreature " & List(0, StmtX.B(0)) & " In " & List(1, StmtX.B(1)) & " Within " & ContextToText(StmtX.B(3)) & "." & VarToText(StmtX.B(3), StmtX.B(4))
				If StmtX.B(2) > 0 Then
					Text = Text & " Dead"
				End If
			Case 32 'TargetItem [List]
				Text = "TargetItem " & List(0, StmtX.B(0))
			Case 33 'PlaySound 'Text' Pause
				Text = "PlaySound " & Chr(34) & TextQuote(StmtX.Text) & Chr(34)
				If StmtX.B(0) > 0 Then
					Text = Text & " Pause"
				End If
			Case 34 'PlayMusic 'Text' Pause
				Text = "PlayMusic " & Chr(34) & TextQuote(StmtX.Text) & Chr(34)
				If StmtX.B(0) > 0 Then
					Text = Text & " Pause"
				End If
			Case 35 'MoveParty To [List] At ([Context.Var],[Context.Var])
				Text = "MoveParty To " & List(0, StmtX.B(5)) & " At (" & ContextToText(StmtX.B(0)) & "." & VarToText(StmtX.B(0), StmtX.B(1)) & ", " & ContextToText(StmtX.B(3)) & "." & VarToText(StmtX.B(3), StmtX.B(4)) & ")"
			Case 36 'DialogShow [List] 'Text' Says 'Text'
				' [Titi 2.4.6] remove the second "Briefline", but only for display (to keep compatibility with existing tomes)
				c = StmtX.B(0) : If c = 5 Then c = 4
				Text = "DialogShow " & Chr(34) & BreakText(TextQuote(StmtX.Text), 1) & Chr(34) & " " & List(0, c) & " Says " & Chr(34) & BreakText(TextQuote(StmtX.Text), 2) & Chr(34)
			Case 37 'DialogReply 'Text'
				Text = "DialogReply " & Chr(34) & BreakText(TextQuote(StmtX.Text), 1) & Chr(34)
			Case 38 'DialogAccept [Context.Var]
				Text = "DialogAccept " & ContextToText(StmtX.B(0)) & "." & VarToText(StmtX.B(0), StmtX.B(1))
			Case 39 'DialogDice 'Text' [List] Says 'Text' Into [Context.Var]
				Text = "DialogDice " & Chr(34) & BreakText(TextQuote(StmtX.Text), 1) & Chr(34) & " " & List(0, StmtX.B(2)) & " Says " & Chr(34) & BreakText(TextQuote(StmtX.Text), 2) & Chr(34) & " Into " & ContextToText(StmtX.B(0)) & "." & VarToText(StmtX.B(0), StmtX.B(1))
			Case 40 'DialogHide
				Text = "DialogHide"
			Case 41 'CutScene 'Text' Picture 'Text' As [List]
				Text = "CutScene " & Chr(34) & BreakText(TextQuote(StmtX.Text), 1) & Chr(34) & " Picture " & Chr(34) & BreakText(TextQuote(StmtX.Text), 2) & Chr(34) & " As " & List(0, StmtX.B(0))
			Case 42 'QueRunes ([List],[List],[List],[List],[List],[List]) Check
				Text = "QueRunes ("
				For c = 0 To 5
					If StmtX.B(c) > 0 Then
						If c = 0 Then
							Text = Text & List(0, StmtX.B(c))
						Else
							Text = Text & ", " & List(0, StmtX.B(c))
						End If
					End If
				Next c
				Text = Text & ")"
				If StmtX.B(6) > 0 Then
					Text = Text & " Check"
				End If
			Case 43 'CombatMove [List] [List]
				Text = "CombatMove " & List(0, StmtX.B(0)) & " " & List(1, StmtX.B(1))
			Case 44 'Reserved
			Case 45 'CombatTarget [List]
				Text = "CombatTarget " & List(0, StmtX.B(0))
			Case 46 'PlaySFX 'Text' As ([List],[List],[List],Number)
				Text = "PlaySFX " & Chr(34) & TextQuote(StmtX.Text) & Chr(34) & " As " & List(0, StmtX.B(0)) & " " & List(1, StmtX.B(1)) & " " & List(2, StmtX.B(2)) & " Frames " & List(3, StmtX.B(3))
			Case 47 'Sorcery [Context.Var]
				Text = "Sorcery " & ContextToText(StmtX.B(0)) & "." & VarToText(StmtX.B(0), StmtX.B(1))
			Case 48 'DialogBuySell [List] At [Context.Var]
				Text = "DialogBuySell " & List(0, StmtX.B(2)) & " Rate " & ContextToText(StmtX.B(0)) & "." & VarToText(StmtX.B(0), StmtX.B(1))
			Case 49 'Reserved
			Case 50 'ExecTrigger [List]
				Text = "ExecTrigger " & List(0, StmtX.B(0))
			Case 51 'DialogAcceptText [Context.Var]
				Text = "DialogAcceptText " & ContextToText(StmtX.B(0)) & "." & VarToText(StmtX.B(0), StmtX.B(1))
			Case 52 'CombatStart
				Text = "CombatStart"
			Case 53 'Let [Context.Var] = [Context.Var]
				Text = "Let " & ContextToText(StmtX.B(0)) & "." & VarToText(StmtX.B(0), StmtX.B(1)) & " = " & ContextToText(StmtX.B(3)) & "." & VarToText(StmtX.B(3), StmtX.B(4))
			Case 54 '' Text
				Text = "' " & TextQuote(StmtX.Text)
			Case 55 'RandomizeEncounter 'Text'
				Text = "RandomizeEncounter " & Chr(34) & TextQuote(StmtX.Text) & Chr(34)
			Case 56 'RandomTheme 'Text'
				Text = "RandomTheme " & Chr(34) & TextQuote(StmtX.Text) & Chr(34)
			Case 57 'AwardExperience [Context.Var] To [List]
				Text = "AwardExperience " & ContextToText(StmtX.B(5)) & "." & VarToText(StmtX.B(5), StmtX.B(6)) & " To " & List(0, StmtX.B(0))
			Case 58 'MoveItem [List] From [List] To [List] Copy
				Text = "MoveItem " & List(0, StmtX.B(0)) & " From " & List(1, StmtX.B(1)) & " To " & List(2, StmtX.B(3))
				If StmtX.B(2) > 0 Then
					Text = Text & " Copy"
				End If
			Case 59 'AddQuest 'Text' As 'Text'
				Text = "AddQuest " & Chr(34) & BreakText(TextQuote(StmtX.Text), 1) & Chr(34) & " As " & Chr(34) & BreakText(TextQuote(StmtX.Text), 2) & Chr(34)
			Case 60 'RemoveQuest 'Text'
				Text = "RemoveQuest " & Chr(34) & TextQuote(StmtX.Text) & Chr(34)
			Case 61 'RemoveTopic 'Text'
				' [Titi 2.4.9]
				Text = "RemoveTopic " & Chr(34) & TextQuote(StmtX.Text) & Chr(34)
			Case 62 'CombatAttackWithWeapon
				Text = "CombatAttackWithWeapon"
			Case 63 'CombatAttackWithSpecial 'Text' As [Context.Var]
				Text = "CombatAttackWithSpecial " & Chr(34) & TextQuote(StmtX.Text) & Chr(34) & " As " & ContextToText(StmtX.B(5)) & "." & VarToText(StmtX.B(5), StmtX.B(6))
			Case 64 'MovePartyMapName ('[List]' ,'Text', [Context.Var], [Context.Var])
				Text = "MovePartyMapName To " & Chr(34) & TextQuote(StmtX.Text) & Chr(34) & " " & List(0, StmtX.B(2)) & " At (" & ContextToText(StmtX.B(0)) & "." & VarToText(StmtX.B(0), StmtX.B(1)) & ", " & ContextToText(StmtX.B(3)) & "." & VarToText(StmtX.B(3), StmtX.B(4)) & ")"
			Case 65 'MoveCreature [List] From [List] To [List] Copy
				Text = "MoveCreature " & List(0, StmtX.B(0)) & " From " & List(1, StmtX.B(1)) & " To " & List(2, StmtX.B(3))
				If StmtX.B(2) > 0 Then
					Text = Text & " Copy"
				End If
			Case 66 'IfText [Context.Var] [Op] 'Text'
				Text = "IfText " & ContextToText(StmtX.B(0)) & "." & VarToText(StmtX.B(0), StmtX.B(1)) & " " & OpToText(StmtX.B(2)) & " " & Chr(34) & TextQuote(StmtX.Text) & Chr(34)
			Case 67 'TargetEncounter 'Text'
				Text = "TargetEncounter " & Chr(34) & TextQuote(StmtX.Text) & Chr(34)
			Case 68 'TargetTile 'Text' At ([Context.Var],[Context.Var],[List])
				Text = "TargetTile At (" & ContextToText(StmtX.B(0)) & "." & VarToText(StmtX.B(0), StmtX.B(1)) & "," & ContextToText(StmtX.B(3)) & "." & VarToText(StmtX.B(3), StmtX.B(4)) & ")" ' Level " & List(0, StmtX.B(2))
			Case 69 'PlayVideo 'Text' Pause
				Text = "PlayVideo " & Chr(34) & TextQuote(StmtX.Text) & Chr(34)
				If StmtX.B(0) > 0 Then
					Text = Text & " Pause"
				End If
			Case 70 'Find [List] Named 'Text' In [List]
				Text = "Find " & List(0, StmtX.B(0)) & " In " & List(1, StmtX.B(1)) & " Named " & Chr(34) & TextQuote(StmtX.Text) & Chr(34)
			Case 71 'CallTrigger [List] In [List]
				Text = "CallTrigger " & List(0, StmtX.B(0)) & " In " & List(1, StmtX.B(1))
			Case 72 ' CombatAnimation 'Text' For [List] Frames [List]
				Text = "CombatAnimation " & Chr(34) & BreakText(TextQuote(StmtX.Text), 1) & Chr(34) & " For " & List(0, StmtX.B(0)) & " Frames " & List(1, StmtX.B(1))
			Case 73 ' MapAnimation 'Text' Frames [List] Level [List] At (Context.Var, Context.Var)
				Text = "MapAnimation " & Chr(34) & BreakText(TextQuote(StmtX.Text), 1) & Chr(34) & " Frames " & List(0, StmtX.B(2)) & " Level " & List(1, StmtX.B(5)) & " At (" & ContextToText(StmtX.B(0)) & "." & VarToText(StmtX.B(0), StmtX.B(1)) & "," & ContextToText(StmtX.B(3)) & "." & VarToText(StmtX.B(3), StmtX.B(4)) & ")"
			Case 74 ' MapRefresh
				Text = "MapRefresh"
		End Select
	End Function
	
	Public Function ContextToText(ByRef Context As Short) As String
		ContextToText = ""
		Select Case Context
			Case 0 : ContextToText = "Tome"
			Case 1 : ContextToText = "Map"
			Case 2 : ContextToText = "TileNow"
			Case 3 : ContextToText = "TileA"
			Case 4 : ContextToText = "TileB"
			Case 5 : ContextToText = "TileC"
			Case 6 : ContextToText = "TileTarget"
			Case 7 : ContextToText = "EncounterNow"
			Case 8 : ContextToText = "EncounterA"
			Case 9 : ContextToText = "EncounterB"
			Case 10 : ContextToText = "EncounterC"
			Case 11 : ContextToText = "EncounterTarget"
			Case 12 : ContextToText = "CreatureNow"
			Case 13 : ContextToText = "CreatureA"
			Case 14 : ContextToText = "CreatureB"
			Case 15 : ContextToText = "CreatureC"
			Case 16 : ContextToText = "CreatureTarget"
			Case 17 : ContextToText = "ItemNow"
			Case 18 : ContextToText = "ItemA"
			Case 19 : ContextToText = "ItemB"
			Case 20 : ContextToText = "ItemC"
			Case 21 : ContextToText = "ItemTarget"
			Case 22 : ContextToText = "TriggerNow"
			Case 23 : ContextToText = "TriggerA"
			Case 24 : ContextToText = "TriggerB"
			Case 25 : ContextToText = "TriggerC"
			Case 26 : ContextToText = "TriggerTarget"
			Case 27 : ContextToText = "Local"
			Case 28 : ContextToText = "Global"
			Case 29 : ContextToText = "Pos"
			Case 30 : ContextToText = "Neg"
			Case 31 : ContextToText = "Dice"
			Case 32 : ContextToText = "Random"
		End Select
	End Function
	
	Public Function VarToText(ByRef Context As Short, ByRef Var As Short) As String
		' [Titi 2.4.8] get the runes names
		Dim sRune() As String
		Dim Text As String
		Dim c As Short
		Text = Right(strRunesList, Len(strRunesList) - 5) ' get rid of "List="
		ReDim sRune(intNbRunes)
		For c = 1 To intNbRunes - 1
			sRune(c) = Left(Text, InStr(Text, ",") - 1)
			Text = Right(Text, Len(Text) - Len(sRune(c)) - 1)
		Next 
		sRune(c) = Text
		VarToText = ""
		Select Case Context
			Case 0 ' Tome
				Select Case Var
					Case 0 : VarToText = "PartyCount"
					Case 1 : VarToText = "Comments"
					Case 2 : VarToText = "Factoids"
					Case 3 : VarToText = "<Reserved>"
					Case 4 : VarToText = "MapX"
					Case 5 : VarToText = "MapY"
					Case 6 : VarToText = "MoveToX"
					Case 7 : VarToText = "MoveToY"
					Case 8 : VarToText = "Name"
					Case 9 : VarToText = "WorldCurrency" '"<Reserved>"
					Case 10 : VarToText = "TimeDay"
					Case 11 : VarToText = "TimeMoon"
					Case 12 : VarToText = "TimeTurn"
					Case 13 : VarToText = "TimeYear"
					Case 14 : VarToText = "RealSeconds"
					Case 15 : VarToText = "RealMinutes"
					Case 16 : VarToText = "RealHours"
					Case 17 : VarToText = "RealDays"
				End Select
			Case 1 ' Map
				Select Case Var
					Case 0 : VarToText = "Comments"
					Case 1 : VarToText = "RandEncounterCount"
					Case 2 : VarToText = "RandMapHeight"
					Case 3 : VarToText = "RandMapStyle"
					Case 4 : VarToText = "RandMapWidth"
					Case 5 : VarToText = "RandDifficulty"
					Case 6 : VarToText = "RandExperiencePoints"
					Case 7 : VarToText = "Height"
					Case 8 : VarToText = "Left"
					Case 9 : VarToText = "Name"
					Case 10 : VarToText = "RandGenerateUponEntry"
					Case 11 To 30 : VarToText = "Rune" & sRune(Var - 10)
						'                Case 11: VarToText = "RuneBlood"
						'                Case 12: VarToText = "RuneBile"
						'                Case 13: VarToText = "RuneOil"
						'                Case 14: VarToText = "RuneNectar"
						'                Case 15: VarToText = "RuneFire"
						'                Case 16: VarToText = "RuneEarth"
						'                Case 17: VarToText = "RuneWater"
						'                Case 18: VarToText = "RuneAir"
						'                Case 19: VarToText = "RuneTime"
						'                Case 20: VarToText = "RuneMoon"
						'                Case 21: VarToText = "RuneSun"
						'                Case 22: VarToText = "RuneSpace"
						'                Case 23: VarToText = "RuneInsect"
						'                Case 24: VarToText = "RuneMan"
						'                Case 25: VarToText = "RuneAnimal"
						'                Case 26: VarToText = "RuneFish"
						'                Case 27: VarToText = "RuneTwilight"
						'                Case 28: VarToText = "RuneAbyss"
						'                Case 29: VarToText = "RuneEternium"
						'                Case 30: VarToText = "RuneDreams"
					Case 31 : VarToText = "Top"
					Case 32 : VarToText = "Width"
					Case 33 : VarToText = "IsOutsideMap"
					Case 34 : VarToText = "MapStyle"
				End Select
			Case 2 To 6 ' Tile
				Select Case Var
					Case 0 : VarToText = "CanSeeDown"
					Case 1 : VarToText = "CanSeeLeft"
					Case 2 : VarToText = "CanSeeRight"
					Case 3 : VarToText = "CanSeeUp"
					Case 4 : VarToText = "ChanceToOpen"
					Case 5 : VarToText = "<Reserved>"
					Case 6 : VarToText = "IsBlockedDown"
					Case 7 : VarToText = "IsBlockedLeft"
					Case 8 : VarToText = "IsBlockedRight"
					Case 9 : VarToText = "IsBlockedUp"
					Case 10 : VarToText = "Key"
					Case 11 To 15 : VarToText = "<Reserved>"
					Case 16 : VarToText = "Name"
					Case 17 : VarToText = "Style"
					Case 18 : VarToText = "DoorTileName"
					Case 19 : VarToText = "MovementCost"
					Case 20 : VarToText = "TerrainType"
				End Select
			Case 7 To 11 ' Encounter
				Select Case Var
					Case 0 : VarToText = "CanFight"
					Case 1 : VarToText = "CanFlee"
					Case 2 : VarToText = "CanTalk"
					Case 3 : VarToText = "ChanceToFlee"
					Case 4 : VarToText = "Classification"
					Case 5 : VarToText = "FirstEntry"
					Case 6 : VarToText = "HaveEntered"
					Case 7 : VarToText = "<Reserved>"
					Case 8 : VarToText = "IsDark"
					Case 9 : VarToText = "Name"
					Case 10 : VarToText = "RandThemeName"
					Case 11 : VarToText = "RandAddCreatures"
					Case 12 : VarToText = "RandAddItems"
					Case 13 : VarToText = "RandAddTriggers"
					Case 14 : VarToText = "RandDescription"
					Case 15 : VarToText = "RandGenerateUponEntry"
					Case 16 : VarToText = "SecondEntry"
					Case 17 : VarToText = "ShowHint"
					Case 18 : VarToText = "CreatureCount"
					Case 19 : VarToText = "ItemCount"
					Case 20 : VarToText = "IsActive"
				End Select
			Case 12 To 16 ' Creature
				Select Case Var
					Case 0 : VarToText = "BodyPart1"
					Case 1 : VarToText = "BodyPart2"
					Case 2 : VarToText = "BodyPart3"
					Case 3 : VarToText = "BodyPart4"
					Case 4 : VarToText = "BodyPart5"
					Case 5 : VarToText = "BodyPart6"
					Case 6 : VarToText = "BodyPart7"
					Case 7 : VarToText = "BodyPart8"
					Case 8 : VarToText = "Protection1%"
					Case 9 : VarToText = "Protection2%"
					Case 10 : VarToText = "Protection3%"
					Case 11 : VarToText = "Protection4%"
					Case 12 : VarToText = "Protection5%"
					Case 13 : VarToText = "Protection6%"
					Case 14 : VarToText = "Protection7%"
					Case 15 : VarToText = "Protection8%"
					Case 16 : VarToText = "Col"
					Case 17 : VarToText = "Comments"
					Case 18 : VarToText = "ActionPoints"
					Case 19 : VarToText = "ExperiencePoints"
					Case 20 : VarToText = "ExperienceLevel"
						'                Case 21: VarToText = "<Reserved>"
						' [Titi 2.4.9]
					Case 21 : VarToText = "Home"
					Case 22 : VarToText = "Greed"
					Case 23 : VarToText = "Health"
					Case 24 : VarToText = "HealthMax"
					Case 25 : VarToText = "Lunacy"
					Case 26 : VarToText = "Lust"
					Case 27 : VarToText = "<Reserved>"
					Case 28 : VarToText = "IsAgressive"
					Case 29 : VarToText = "IsComputerControlled"
					Case 30 : VarToText = "IsFriendly"
					Case 31 : VarToText = "IsGuarding"
					Case 32 : VarToText = "ProtectSharp%"
					Case 33 : VarToText = "ProtectBlunt%"
					Case 34 : VarToText = "ProtectCold%"
					Case 35 : VarToText = "ProtectFire%"
					Case 36 : VarToText = "ProtectEvil%"
					Case 37 : VarToText = "ProtectHoly%"
					Case 38 : VarToText = "ProtectMagic%"
					Case 39 : VarToText = "ProtectMind%"
					Case 40 To 45 : VarToText = "<Reserved>"
					Case 46 : VarToText = "IsTypeAnimal"
					Case 47 : VarToText = "IsTypeBird"
					Case 48 : VarToText = "IsTypeBlob"
					Case 49 : VarToText = "IsTypeAquatic"
					Case 50 : VarToText = "IsTypeHuge"
					Case 51 : VarToText = "IsTypeSentient"
					Case 52 : VarToText = "IsTypeInsect"
					Case 53 : VarToText = "IsTypeLarge"
					Case 54 : VarToText = "IsTypeMagical"
					Case 55 : VarToText = "IsTypeMedium"
					Case 56 : VarToText = "IsTypeReptile"
					Case 57 : VarToText = "IsTypeSmall"
					Case 58 : VarToText = "IsTypeTiny"
					Case 59 : VarToText = "IsTypeVeggie"
					Case 60 : VarToText = "IsTypeUndead"
					Case 61 To 74 : VarToText = "<Reserved>"
					Case 75 : VarToText = "Eternium"
					Case 76 : VarToText = "MapX"
					Case 77 : VarToText = "MapY"
					Case 78 : VarToText = "Name"
					Case 79 : VarToText = "Strength"
					Case 80 : VarToText = "Pride"
					Case 81 : VarToText = "Race"
					Case 82 : VarToText = "Revelry"
					Case 83 : VarToText = "Row"
					Case 84 : VarToText = "RuneQue1"
					Case 85 : VarToText = "RuneQue2"
					Case 86 : VarToText = "RuneQue3"
					Case 87 : VarToText = "RuneQue4"
					Case 88 : VarToText = "RuneQue5"
					Case 89 : VarToText = "RuneQue6"
					Case 90 : VarToText = "ScarletLetter"
					Case 91 : VarToText = "Agility"
					Case 92 : VarToText = "SkillPoints"
					Case 93 : VarToText = "Status"
					Case 94 : VarToText = "<Reserved>"
					Case 95 : VarToText = "Defense"
					Case 96 : VarToText = "Intelligence"
					Case 97 : VarToText = "Wrath"
					Case 98 : VarToText = "IsMale"
					Case 99 : VarToText = "IsUnconscious"
					Case 100 : VarToText = "IsFrozen"
					Case 101 : VarToText = "RangeToTarget"
					Case 102 : VarToText = "<Reserved>"
					Case 103 : VarToText = "ProtectSharpBonus%"
					Case 104 : VarToText = "ProtectBluntBonus%"
					Case 105 : VarToText = "ProtectColdBonus%"
					Case 106 : VarToText = "ProtectFireBonus%"
					Case 107 : VarToText = "ProtectEvilBonus%"
					Case 108 : VarToText = "ProtectHolyBonus%"
					Case 109 : VarToText = "ProtectMagicBonus%"
					Case 110 : VarToText = "ProtectMindBonus%"
					Case 111 : VarToText = "ProtectBonus%"
					Case 112 To 132 : VarToText = "<Reserved>"
					Case 133 : VarToText = "EterniumMax"
					Case 134 : VarToText = "WeightMax"
					Case 135 : VarToText = "Weight"
					Case 136 : VarToText = "Bulk"
					Case 137 : VarToText = "StrengthBonus"
					Case 138 : VarToText = "AgilityBonus"
					Case 139 : VarToText = "IntelligenceBonus"
					Case 140 : VarToText = "AttackBonus"
					Case 141 : VarToText = "DamageBonus"
					Case 142 : VarToText = "DefenseBonus"
					Case 143 : VarToText = "ArmorBonus"
					Case 144 : VarToText = "IsSpellCaster"
					Case 145 : VarToText = "Money"
					Case 146 : VarToText = "ActionPointsMax"
					Case 147 : VarToText = "IsAfraid"
					Case 148 : VarToText = "IsInanimate"
					Case 149 : VarToText = "<Reserved>"
					Case 150 : VarToText = "CombatRank"
					Case 151 : VarToText = "PronounHeShe"
					Case 152 : VarToText = "PronounHimHer"
					Case 153 : VarToText = "PronounHisHer"
					Case 154 : VarToText = "MovementCost"
					Case 155 : VarToText = "MovementCostBonus"
					Case 156 : VarToText = "Index"
					Case 157 : VarToText = "PictureFile"
					Case 158 : VarToText = "FaceLeft"
					Case 159 : VarToText = "FaceTop"
					Case 160 : VarToText = "Initiative"
					Case 161 : VarToText = "Facing"
				End Select
			Case 17 To 21 ' Item
				Select Case Var
					Case 0 : VarToText = "Protection%"
					Case 1 : VarToText = "Bulk"
					Case 2 : VarToText = "CanCombine"
					Case 3 : VarToText = "ArmorType"
					Case 4 To 15 : VarToText = "<Reserved>"
					Case 16 : VarToText = "Capacity"
					Case 17 : VarToText = "<Reserved>"
					Case 18 : VarToText = "Comments"
					Case 19 : VarToText = "Count"
					Case 20 : VarToText = "DamageDice"
					Case 21 : VarToText = "DamageBonus"
					Case 22 : VarToText = "IsWeaponMelee"
					Case 23 : VarToText = "IsWeaponRanged"
					Case 24 : VarToText = "IsEquipped"
					Case 25 : VarToText = "IsWeaponThrown"
					Case 26 : VarToText = "Key"
					Case 27 : VarToText = "ProtectionBonusType"
					Case 28 : VarToText = "ProtectionBonus%"
					Case 29 : VarToText = "IsWeaponAmmo"
					Case 30 : VarToText = "<Reserved>"
					Case 31 : VarToText = "IsRangeLong"
					Case 32 : VarToText = "IsRangeMedium"
					Case 33 : VarToText = "IsRangeShort"
					Case 34 To 35 : VarToText = "<Reserved>"
					Case 36 : VarToText = "IsSoftBulk"
					Case 37 : VarToText = "IsDamageTypeSharp"
					Case 38 : VarToText = "IsDamageTypeBlunt"
					Case 39 : VarToText = "IsDamageTypeCold"
					Case 40 : VarToText = "IsDamageTypeFire"
					Case 41 : VarToText = "IsDamageTypeHoly"
					Case 42 : VarToText = "IsDamageTypeEvil"
					Case 43 : VarToText = "IsDamageTypeMagic"
					Case 44 : VarToText = "IsDamageTypeMind"
					Case 45 To 49 : VarToText = "<Reserved>"
					Case 50 : VarToText = "MapX"
					Case 51 : VarToText = "MapY"
					Case 52 : VarToText = "Name"
					Case 53 : VarToText = "DefenseBonus"
					Case 54 : VarToText = "<Reserved>"
					Case 55 : VarToText = "AttackBonus"
					Case 56 : VarToText = "UseAsDescription"
					Case 57 : VarToText = "Value"
					Case 58 : VarToText = "Weight"
					Case 59 : VarToText = "RequiresTwoHands"
					Case 60 : VarToText = "IsInHand"
					Case 61 : VarToText = "AmmoType"
					Case 62 To 63 : VarToText = "<Reserved>"
					Case 64 : VarToText = "Family"
					Case 65 : VarToText = "ActionPoints"
				End Select
			Case 22 To 26 ' Trigger
				Select Case Var
					Case 0 : VarToText = "ByteA"
					Case 1 : VarToText = "ByteB"
					Case 2 : VarToText = "ByteC"
					Case 3 : VarToText = "Comments"
					Case 4 : VarToText = "<Reserved>"
					Case 5 : VarToText = "IsTimed"
					Case 6 : VarToText = "IsCurse"
					Case 7 : VarToText = "IsEvil"
					Case 8 : VarToText = "IsFear"
					Case 9 : VarToText = "IsFish"
					Case 10 : VarToText = "IsGreed"
					Case 11 : VarToText = "IsLunacy"
					Case 12 : VarToText = "IsLust"
					Case 13 : VarToText = "IsPoison"
					Case 14 : VarToText = "IsPride"
					Case 15 : VarToText = "IsMagic"
					Case 16 : VarToText = "IsSkill"
					Case 17 : VarToText = "IsRevelry"
					Case 18 : VarToText = "IsTrap"
					Case 19 : VarToText = "IsWrath"
					Case 20 : VarToText = "Name"
					Case 21 : VarToText = "SkillPoints"
					Case 22 : VarToText = "TriggerType"
					Case 23 : VarToText = "Turns"
					Case 24 : VarToText = "SkillLevel"
				End Select
			Case 27 ' Local
				Select Case Var
					Case 0 : VarToText = "Abort"
					Case 1 : VarToText = "ByteA"
					Case 2 : VarToText = "ByteB"
					Case 3 : VarToText = "ByteC"
					Case 4 : VarToText = "Fail"
					Case 5 : VarToText = "IntegerA"
					Case 6 : VarToText = "IntegerB"
					Case 7 : VarToText = "IntegerC"
					Case 8 : VarToText = "StopExit"
					Case 9 : VarToText = "TextA"
					Case 10 : VarToText = "TextB"
					Case 11 : VarToText = "TextC"
					Case 12 : VarToText = "Random"
					Case 13 : VarToText = "FoundIt"
				End Select
			Case 28 ' Global
				Select Case Var
					Case 0 : VarToText = "False"
					Case 1 : VarToText = "True"
					Case 2 : VarToText = "InCombat"
					Case 3 : VarToText = "DieRollType"
					Case 4 : VarToText = "DieRollCount"
					Case 5 : VarToText = "ArmorRoll"
					Case 6 : VarToText = "AttackRoll"
					Case 7 : VarToText = "DamageRoll"
					Case 8 : VarToText = "Offer"
					Case 9 To 28 : VarToText = "Rune" & sRune(Var - 8)
						'                Case 9: VarToText = "RuneBlood"
						'                Case 10: VarToText = "RuneBile"
						'                Case 11: VarToText = "RuneOil"
						'                Case 12: VarToText = "RuneNectar"
						'                Case 13: VarToText = "RuneFire"
						'                Case 14: VarToText = "RuneEarth"
						'                Case 15: VarToText = "RuneWater"
						'                Case 16: VarToText = "RuneAir"
						'                Case 17: VarToText = "RuneTime"
						'                Case 18: VarToText = "RuneMoon"
						'                Case 19: VarToText = "RuneSuns"
						'                Case 20: VarToText = "RuneSpace"
						'                Case 21: VarToText = "RuneInsect"
						'                Case 22: VarToText = "RuneMan"
						'                Case 23: VarToText = "RuneFish"
						'                Case 24: VarToText = "RuneAnimal"
						'                Case 25: VarToText = "RuneTwilight"
						'                Case 26: VarToText = "RuneAbyss"
						'                Case 27: VarToText = "RuneDreams"
						'                Case 28: VarToText = "RuneEternium"
					Case 29 : VarToText = "SkillLevel"
					Case 30 : VarToText = "IsAttackAir"
					Case 31 : VarToText = "IsAttackBlunt"
					Case 32 : VarToText = "IsAttackEarth"
					Case 33 : VarToText = "IsAttackFire"
					Case 34 : VarToText = "IsAttackHoly"
					Case 35 : VarToText = "IsAttackDisease"
					Case 36 : VarToText = "IsAttackMagic"
					Case 37 : VarToText = "IsAttackSharp"
					Case 38 : VarToText = "IsAttackEvil"
					Case 39 : VarToText = "IsAttackCold"
					Case 40 : VarToText = "IsAttackIllusion"
					Case 41 : VarToText = "IsAttackWater"
					Case 42 : VarToText = "IsAttackFear"
					Case 43 To 57 : VarToText = "<Reserved>"
					Case 58 : VarToText = "IntegerA"
					Case 59 : VarToText = "IntegerB"
					Case 60 : VarToText = "IntegerC"
					Case 61 : VarToText = "TextA"
					Case 62 : VarToText = "TextB"
					Case 63 : VarToText = "TextC"
					Case 64 : VarToText = "TimeDayName"
					Case 65 : VarToText = "TimeMoonName"
					Case 66 : VarToText = "TimeYearName"
					Case 67 : VarToText = "TimeTurnName"
					Case 68 : VarToText = "PickLockChance"
					Case 69 : VarToText = "<Reserved>"
					Case 70 : VarToText = "RemoveTrapChance"
					Case 71 : VarToText = "HitLocation"
					Case 72 : VarToText = "SpellFizzleChance"
					Case 73 : VarToText = "Ticks"
				End Select
			Case 29, 30 ' Positive and Negative Numbers
				If Var > -1 And Var < 256 Then
					VarToText = VB6.Format(Var)
				End If
			Case 31 ' Dice
				If Var > -1 And Var < 256 Then
					VarToText = Trim(Str(Var Mod 5 + 1)) & "d" & Trim(Str(Int((Var Mod 25) / 5) * 2 + 4))
					If Var > 24 Then
						VarToText = VarToText & "+" & Trim(Str(Int(Var / 25)))
					End If
				End If
			Case 32 ' Random number
				If Var > -1 And Var < 256 Then
					VarToText = Trim(VB6.Format(Var + 1))
				End If
		End Select
	End Function
	
	Public Function Proper(ByVal StringX As String) As String
		Dim c As Short
		For c = 1 To Len(StringX)
			If c = 1 Then
				Proper = UCase(Mid(StringX, 1, 1)) & Mid(StringX, 2)
			ElseIf Asc(Mid(StringX, c - 1, 1)) = 32 Then 
				Proper = Left(Proper, c - 1) & UCase(Mid(StringX, c, 1)) & Mid(StringX, c + 1)
			End If
		Next c
	End Function
	
	Public Function AOrAn(ByRef InString As String) As String
		Select Case UCase(Mid(InString, 1, 1))
			Case "A", "E", "I", "O", "U", "Y"
				AOrAn = "an "
			Case Else
				AOrAn = "a "
		End Select
	End Function
	
    Public Function Limit(ByRef ChkText As System.Windows.Forms.TextBox, ByRef KeyAscii As Short, ByRef Mode As Short) As Short
        Dim NewText As String
        ' Limit to numbers and backspace
        If KeyAscii > 0 Then
            Select Case KeyAscii
                Case System.Windows.Forms.Keys.Back, System.Windows.Forms.Keys.Delete, 45, 48 To 57
                    ' Replace ChkText.Text with text instead of validating
                    'UPGRADE_WARNING: Couldn't resolve default property of object ChkText.SelLength. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
                    If Not ChkText.Size.Equals(0) Then
                        Select Case KeyAscii
                            Case System.Windows.Forms.Keys.Back, System.Windows.Forms.Keys.Delete
                                'UPGRADE_WARNING: Couldn't resolve default property of object ChkText.SelLength. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
                                'UPGRADE_WARNING: Couldn't resolve default property of object ChkText.SelStart. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
                                NewText = Mid(ChkText.Text, 1, ChkText.SelectionStart & Mid(ChkText.Text, ChkText.SelectionStart + ChkText.SelectionLength + 1))
                            Case 45, 48 To 57
                                'UPGRADE_WARNING: Couldn't resolve default property of object ChkText.SelLength. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
                                'UPGRADE_WARNING: Couldn't resolve default property of object ChkText.SelStart. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
                                NewText = Mid(ChkText.Text, 1, ChkText.SelectionStart) & Chr(KeyAscii) & Mid(ChkText.Text, ChkText.SelectionStart + ChkText.SelectionLength + 1)
                        End Select
                    Else
                        NewText = ChkText.Text & Chr(KeyAscii)
                    End If
                    ' Check for valid limit of Integer or Byte
                    Select Case Mode
                        Case bdInt
                            If Val(NewText) > 32000 Then
                                Limit = 0
                            Else
                                Limit = KeyAscii
                            End If
                        Case bdByte
                            If Val(NewText) < 0 Or Val(NewText) > 255 Then
                                Limit = 0
                            Else
                                Limit = KeyAscii
                            End If
                        Case bdLong
                            If Val(NewText) < -2000000000 Or Val(NewText) > 2000000000 Then
                                Limit = 0
                            Else
                                Limit = KeyAscii
                            End If
                    End Select
                Case Else
                    Limit = 0
            End Select
        End If
    End Function
	
	Sub FlickText(ByRef InLabel As System.Windows.Forms.Label)
		Dim c As Short

		'    Call PlaySoundFile(gAppPath & "\data\stock\click.wav", Tome, True)
		Call PlaySoundFile(gDataPath & "\stock\click.wav", Tome, True)
		For c = 0 To 1
			InLabel.ForeColor = System.Drawing.ColorTranslator.FromOle(&HC0C0C0)
			InLabel.Refresh()
			Sleep(50)
			InLabel.ForeColor = System.Drawing.ColorTranslator.FromOle(RGB(255, 255, 0))
			InLabel.Refresh()
			Sleep(50)
		Next c
	End Sub
	
	Public Function LoopNumber(ByVal FromNum As Short, ByVal ToNum As Short, ByVal AtNum As Short, ByVal AddNum As Short) As Short
		' Note: Might be good idea to create a loop function which given a range of
		' number like 0-64 at 13, add/subtract a number and return position in that
		' circle of numbers.
		' So if pass in 0,127,125,+10 returns 135-127-0-1 or 7.
		' If pass in 0,255,0,-1 returns -1 + 255 or 254
		' If pass in 5,15,7,-10 return -3+15-5+1 or 8
		AtNum = AtNum + AddNum
		If AtNum > ToNum Then
			AtNum = AtNum - ToNum + FromNum - 1
		ElseIf AtNum < FromNum Then 
			AtNum = AtNum + ToNum - FromNum + 1
		End If
		LoopNumber = AtNum
	End Function
	
	Function Greatest(ByRef XX As Short, ByRef YY As Short) As Short
		Greatest = XX
		If YY > XX Then
			Greatest = YY
		End If
	End Function
	
	Function IsBetween(ByVal aa As Short, ByVal bb As Short, ByVal cc As Short) As Short
		' Simple function to see if A is between B and C
		If aa >= bb And aa <= cc Then
			IsBetween = True
		Else
			IsBetween = False
		End If
		
	End Function
	
	Public Sub ConvertDice(ByVal InValue As Short, ByRef RollCnt As Short, ByRef Dice As Short, ByRef Damage As Short)
		RollCnt = (InValue - 1) Mod 5 + 1
		Dice = Int(((InValue - 1) Mod 25) / 5) * 2 + 4
		Damage = Int((InValue - 1) / 25)
	End Sub
	
	Public Function Least(ByRef XX As Short, ByRef YY As Short) As Short
		Least = XX
		If YY < XX Then
			Least = YY
		End If
	End Function
	
	Public Function LenStr(ByRef InString As String) As Short
		Dim c As Short
		For c = Len(InString) To 1 Step -1
			If Asc(Mid(InString, c, 1)) <> 0 And Asc(Mid(InString, c, 1)) <> 32 Then
				Exit For
			End If
		Next c
		LenStr = c
	End Function
	
	Function PointIn(ByRef XX As Short, ByRef YY As Short, ByRef LeftX As Short, ByRef TopY As Short, ByRef WidthX As Short, ByRef HeightY As Short) As Short
		' Determines if a values is inside a rectangle
		If XX >= LeftX And XX <= LeftX + WidthX And YY >= TopY And YY <= TopY + HeightY Then
			PointIn = True
		Else
			PointIn = False
		End If
	End Function
	
	Function OverLap(ByRef X1 As Short, ByRef Y1 As Short, ByRef Width1 As Short, ByRef Height1 As Short, ByRef X2 As Short, ByRef Y2 As Short, ByRef Width2 As Short, ByRef Height2 As Short) As Short
		' Checks to see if two rectangles overlap each other. If so, return TRUE
		OverLap = False
		' First, check if X1 is on X2
		If IsBetween(X1, X2, X2 + Width2) Or IsBetween(X1 + Width1, X2, X2 + Width2) Then
			If IsBetween(Y1, Y2, Y2 + Height2) Or IsBetween(Y1 + Height1, Y2, Y2 + Height2) Then
				OverLap = True
			End If
		End If
		' Then check if X2 is on X1
		If IsBetween(X2, X1, X1 + Width1) Or IsBetween(X2 + Width2, X1, X1 + Width1) Then
			If IsBetween(Y2, Y1, Y1 + Height1) Or IsBetween(Y2 + Height2, Y1, Y1 + Height1) Then
				OverLap = True
			End If
		End If
	End Function
	
	Public Sub SetUpBits()
		' Set up bits for all bitwise checking
		Bit(0) = 1
		Bit(1) = 2
		Bit(2) = 4
		Bit(3) = 8
		Bit(4) = 16
		Bit(5) = 32
		Bit(6) = 64
		Bit(7) = 128
		Bit(8) = 256
		Bit(9) = 512
		Bit(10) = 1024
		Bit(11) = 2048
		Bit(12) = 4096
		Bit(13) = 8192
		Bit(14) = 16384
		Bit(15) = 32768
		Bit(16) = 65536
		Bit(17) = 131072
		Bit(18) = 262144
		Bit(19) = 524288
	End Sub
	
	Public Function BreakText(ByVal InString As String, ByVal Pos As Short) As String
		Dim i, c, LastItem As Short
		c = 0 : i = InStr(InString, "|") : LastItem = 0
		Do Until c = Pos Or i < 1
			BreakText = Mid(InString, LastItem + 1, i - LastItem - 1)
			LastItem = i
			i = InStr(LastItem + 1, InString, "|")
			c = c + 1
		Loop 
		If c = 0 And Pos > 1 Then
			BreakText = ""
		ElseIf i < 1 And c < Pos Then 
			BreakText = Mid(InString, LastItem + 1)
		End If
	End Function
	
	Public Sub SetUpStatus(ByRef Status() As String)
		' Fill Status
		Status(0) = "Ok"
		Status(1) = "Asleep"
		Status(2) = "Sluggish"
		Status(3) = "Poison 10%"
		Status(4) = "Poison 20%"
		Status(5) = "Poison 30%"
		Status(6) = "Poison 40%"
		Status(7) = "Poison 50%"
		Status(8) = "Poison 60%"
		Status(9) = "Poison 70%"
		Status(10) = "Poison 80%"
		Status(11) = "Poison 90%"
		Status(12) = "Poison 100%"
		Status(13) = "Berserk"
		Status(14) = "Illusion"
		Status(15) = "Hungry"
		Status(16) = "Unconscious"
		Status(17) = "Dead"
	End Sub
	
	Public Function SetupTriggerSchool(ByRef Index As Short) As String
		Select Case Index
			Case 1
				SetupTriggerSchool = "Unknown"
			Case 2
				SetupTriggerSchool = "Curse"
			Case 3
				SetupTriggerSchool = "Poison"
			Case 4
				SetupTriggerSchool = "Evil"
			Case 5
				SetupTriggerSchool = "Fear"
			Case 6
				SetupTriggerSchool = "Magic"
			Case 7
				SetupTriggerSchool = "Timed"
			Case 8
				SetupTriggerSchool = "Lunacy"
			Case 9
				SetupTriggerSchool = "Revelry"
			Case 10
				SetupTriggerSchool = "Wrath"
			Case 11
				SetupTriggerSchool = "Pride"
			Case 12
				SetupTriggerSchool = "Greed"
			Case 13
				SetupTriggerSchool = "Lust"
			Case 14
				SetupTriggerSchool = "Dreamer"
			Case Else
				SetupTriggerSchool = "Unknown"
		End Select
	End Function
	
	Public Function SetUpTileStyle(ByRef Index As Short) As String
		Select Case Index
			Case 0
				SetUpTileStyle = "<None>"
			Case 1
				SetUpTileStyle = "Left Wall"
			Case 2
				SetUpTileStyle = "Corner Wall"
			Case 3
				SetUpTileStyle = "Floor"
			Case 4
				SetUpTileStyle = "Left Arch"
			Case 5
				SetUpTileStyle = "Left Door"
			Case 6
				SetUpTileStyle = "Exit Up"
			Case 7
				SetUpTileStyle = "Exit Down"
			Case 8
				SetUpTileStyle = "Floor Decoration"
			Case 9
				SetUpTileStyle = "Wall Decoration"
			Case 10
				SetUpTileStyle = "Corner Arch"
			Case 11
				SetUpTileStyle = "Block"
			Case Else
				SetUpTileStyle = ""
		End Select
	End Function
	
	Public Function SetUpArmorType(ByRef Index As Short) As String
		Select Case Index
			Case 0
				SetUpArmorType = "Wing"
			Case 1
				SetUpArmorType = "Tail"
			Case 2
				SetUpArmorType = "Body"
			Case 3
				SetUpArmorType = "Head"
			Case 4
				SetUpArmorType = "Arm"
			Case 5
				SetUpArmorType = "Leg"
			Case 6
				SetUpArmorType = "Antenna"
			Case 7
				SetUpArmorType = "Tentacle"
			Case 8
				SetUpArmorType = "Abdomen"
			Case 9
				SetUpArmorType = "Back"
			Case 10
				SetUpArmorType = "Neck"
			Case Else
				SetUpArmorType = "<Undefined>"
		End Select
	End Function
	
	Public Function SetUpWearType(ByRef Index As Short) As String
		Select Case Index
			Case 0
				SetUpWearType = "Body"
			Case 1
				SetUpWearType = "Helm"
			Case 2
				SetUpWearType = "Glove"
			Case 3
				SetUpWearType = "Bracelet"
			Case 4
				SetUpWearType = "Backpack"
			Case 5
				SetUpWearType = "Shield"
			Case 6
				SetUpWearType = "Boots"
			Case 7
				SetUpWearType = "Necklace"
			Case 8
				SetUpWearType = "Belt"
			Case 9
				SetUpWearType = "Ring"
			Case 10
				SetUpWearType = "OneHanded"
			Case 11
				SetUpWearType = "TwoHanded"
			Case Else
				SetUpWearType = "<Undefined>"
		End Select
	End Function
	
	Public Function SetUpTileTerrain(ByRef Index As Short) As String
		Select Case Index
			Case 0
				SetUpTileTerrain = "<None>"
			Case 1
				SetUpTileTerrain = "Building"
			Case 2
				SetUpTileTerrain = "Dirt Road"
			Case 3
				SetUpTileTerrain = "Paved Road"
			Case 4
				SetUpTileTerrain = "Desert"
			Case 5
				SetUpTileTerrain = "Rocky Plains"
			Case 6
				SetUpTileTerrain = "Grassy Plains"
			Case 7
				SetUpTileTerrain = "Rolling Plains"
			Case 8
				SetUpTileTerrain = "Light Woods"
			Case 9
				SetUpTileTerrain = "Dark Woods"
			Case 10
				SetUpTileTerrain = "Foot Hills"
			Case 11
				SetUpTileTerrain = "Low Mountains"
			Case 12
				SetUpTileTerrain = "High Mountains"
			Case 13
				SetUpTileTerrain = "Swamp"
			Case 14
				SetUpTileTerrain = "River"
			Case 15
				SetUpTileTerrain = "Shoreline"
			Case 16
				SetUpTileTerrain = "Shallow Water"
			Case 17
				SetUpTileTerrain = "Deep Water"
			Case Else
				SetUpTileTerrain = ""
		End Select
	End Function
	
	Public Function SetUpEncClass(ByRef Index As Short) As String
		Select Case Index
			Case 0
				SetUpEncClass = "<None>"
			Case 1
				SetUpEncClass = "Safe Area"
			Case 2
				SetUpEncClass = "Dangerous Area"
			Case 3
				SetUpEncClass = "Inn/Pub"
			Case 4
				SetUpEncClass = "Merchant"
			Case 5
				SetUpEncClass = "Meeting Area"
			Case 6
				SetUpEncClass = "Guard Area"
			Case 7
				SetUpEncClass = "Temple Area"
			Case 8
				SetUpEncClass = "Dwelling Area"
			Case Else
				SetUpEncClass = ""
		End Select
	End Function
	
	Public Function SetUpResistanceType(ByRef Index As Short) As String
		Select Case Index
			Case 0
				SetUpResistanceType = "<None>"
			Case 1
				SetUpResistanceType = "Sharp"
			Case 2
				SetUpResistanceType = "Blunt"
			Case 3
				SetUpResistanceType = "Cold"
			Case 4
				SetUpResistanceType = "Fire"
			Case 5
				SetUpResistanceType = "Evil"
			Case 6
				SetUpResistanceType = "Good"
			Case 7
				SetUpResistanceType = "Magic"
			Case 8
				SetUpResistanceType = "Mind"
		End Select
	End Function
	
	Public Function SetUpItemFamily(ByRef Index As Short) As String
		Select Case Index
			Case 0
				SetUpItemFamily = "<None>"
			Case 1
				SetUpItemFamily = "Key"
			Case 2
				SetUpItemFamily = "Money"
			Case 3
				SetUpItemFamily = "Food"
			Case 4
				SetUpItemFamily = "Armor"
			Case 5
				SetUpItemFamily = "Weapon (Melee)"
			Case 6
				SetUpItemFamily = "Weapon (Ammo)"
			Case 7
				SetUpItemFamily = "Weapon (Ranged)"
			Case 8
				SetUpItemFamily = "Weapon (Thrown)"
			Case 9
				SetUpItemFamily = "Random Weapon"
			Case 10
				SetUpItemFamily = "Random Armor"
			Case 11
				SetUpItemFamily = "Random Gems"
			Case 12
				SetUpItemFamily = "Random Treasure"
			Case 13
				SetUpItemFamily = "Random Junk"
			Case 14
				SetUpItemFamily = "Fragile"
			Case 15
				SetUpItemFamily = "Locked"
			Case 16
				SetUpItemFamily = "Jammed Shut"
			Case 17
				SetUpItemFamily = "Durable"
			Case 18 ' [Titi] added for 2.4.8
				SetUpItemFamily = "Worthless currency"
			Case Else
				SetUpItemFamily = ""
		End Select
	End Function
	
    Public Sub InitializeRunes(ByRef strWorldName As String, ByRef spath As String)
        ' [Titi 2.4.8] added to use the runes of the current world
        Dim Filename, Text As String
        Dim hndFile As Short
        Filename = "Roster\" & strWorldName & "\" & strWorldName & ".txt"
        If Not (System.IO.File.Exists(Filename)) Then
            ' if not found, use Eternia's settings
            Filename = gAppPath & "\Roster\Eternia\Eternia.txt"
        End If
        hndFile = FreeFile()
        FileOpen(hndFile, Filename, OpenMode.Input)
        Text = LineInput(hndFile) ' first line = runes section
        Text = LineInput(hndFile) ' second line = runes count
        intNbRunes = CShort(Val(Right(Text, Len(Text) - 6))) ' get rid of "Count="
        If intNbRunes > 20 Then intNbRunes = 20 ' max 20
        strRunesList = LineInput(hndFile) ' third line = list of runes names
        Text = LineInput(hndFile) ' fourth line = runes pictures set
        FileClose(hndFile)
        Text = Right(Text, Len(Text) - 11) ' get rid of "PictureSet="

    End Sub
End Module