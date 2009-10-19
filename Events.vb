Option Strict Off
Option Explicit On
Imports VB = Microsoft.VisualBasic
Module modEvents
	Private FrameNow As TriggerFrame
	Const bdIfSkip As Short = -1
	Const bdIfProcess As Short = 0
	Const bdIfEnd As Short = 1
	
	Private Function FindMatch(ByRef StartLine As Short, ByRef TrigToCheck As Trigger) As Short
		' Function by Phule, which will hopefully find the "matching" statements
		' for ForEach/Next, While/Next pairs
		Dim i As Short
		Dim Depth As Short
		Dim Direction As Short
		i = StartLine ' used for scanning the trigger
		Depth = 0 ' tied with the start line
		Direction = 0 ' -1 = up (for a Next), 1 = down (for a ForEach or While)
		'UPGRADE_WARNING: Couldn't resolve default property of object TrigToCheck.Statements(StartLine).Statement. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
		If TrigToCheck.Statements.Item(StartLine).Statement = 10 Then ' Next
			Direction = -1
			'UPGRADE_WARNING: Couldn't resolve default property of object TrigToCheck.Statements(StartLine).Statement. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
		ElseIf TrigToCheck.Statements.Item(StartLine).Statement = 8 Then  ' While
			Direction = 1
			'UPGRADE_WARNING: Couldn't resolve default property of object TrigToCheck.Statements(StartLine).Statement. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
		ElseIf TrigToCheck.Statements.Item(StartLine).Statement = 9 Then  ' ForEach
			Direction = 1
		End If
		' check boundaries while we are at it... - phule
		If ((Direction = 0) Or (i <= 0) Or (i >= TrigToCheck.Statements.Count())) Then
			' quit if we didn't get one of the ones we are looking for
			FindMatch = i
			Exit Function
		End If
		' otherwise, scan up or down, counting as we go, to find the "matching"
		' statement
		Depth = Direction
		With TrigToCheck
			' scan up or down, counting "depth" until we hit 0 - which is
			' the statement we are after
			Do 
				i = i + Direction
				'UPGRADE_WARNING: Couldn't resolve default property of object TrigToCheck.Statements(i).Statement. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
				If .Statements.Item(i).Statement = 10 Then ' Next
					Depth = Depth - 1
					'UPGRADE_WARNING: Couldn't resolve default property of object TrigToCheck.Statements(i).Statement. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
				ElseIf .Statements.Item(i).Statement = 8 Then  ' While
					Depth = Depth + 1
					'UPGRADE_WARNING: Couldn't resolve default property of object TrigToCheck.Statements(i).Statement. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
				ElseIf .Statements.Item(i).Statement = 9 Then  ' ForEach
					Depth = Depth + 1
				End If
				' i = i + Direction
				' Whoops - need to check boundaries! - Phule
			Loop Until ((Depth = 0) Or (i <= 0) Or (i > .Statements.Count()))
			FindMatch = i - 1
		End With
	End Function
	
	'UPGRADE_NOTE: Map was upgraded to Map_Renamed. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="A9E4979A-37FA-4718-9994-97DD76ED70A7"'
	'UPGRADE_NOTE: Tome was upgraded to Tome_Renamed. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="A9E4979A-37FA-4718-9994-97DD76ED70A7"'
	'UPGRADE_NOTE: Area was upgraded to Area_Renamed. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="A9E4979A-37FA-4718-9994-97DD76ED70A7"'
	Public Function FireTrigger(ByRef ParentX As Object, ByRef TrigToFire As Trigger) As Short
		Dim Area_Renamed As Object
		Dim Tome_Renamed As Object
		Dim Map_Renamed As Object
		Static PgmStack As New Collection
		Dim CreatureX, CreatureY As Creature
		Dim ItemX, ItemY As Item
		Dim EncounterX As Encounter
		Dim ThemeX As Theme
		Dim TriggerX As Trigger
		Dim TileX As Tile
		Dim AreaX As Area
		Dim MapX As Map
		Dim TopicX As Topic
		Dim ConvoX As Conversation
		Dim X, i, c, RefreshMap, Y As Short
		Dim ReplyTxt As String
		Dim sText As String
		Dim Found As Short
		Dim FrameX As TriggerFrame
		Dim StmtX As Statement
		Dim FactoidX As Factoid
		Dim JournalX As Journal
		Dim ObjectX, ObjectY As Object
		Dim t1, t2 As Single
		Dim Height, Width, Cells As Short
		Dim rc As Short
		' Add TriggerName to Debug
		If GlobalDebugMode = 1 Then
			'frmMain.DebugHeader = modBD.TriggerName(TrigToFire.TriggerType) & " * " & TrigToFire.Name
			DebugHeader = modBD.TriggerName(TrigToFire.TriggerType) & " * " & TrigToFire.Name
		End If
		' If no statements, then return
		If TrigToFire.Statements.Count() = 0 Then
			FireTrigger = True
			Exit Function
		End If
		' Put new Trigger on stack
		c = 1
		For	Each FrameX In PgmStack
			If FrameX.Index >= c Then
				c = FrameX.Index + 1
			End If
		Next FrameX
		FrameX = New TriggerFrame
		FrameX.Trigger = TrigToFire
		FrameX.Index = c
		PgmStack.Add(FrameX, "T" & FrameX.Index)
		' Expose to other functions in this module
		FrameNow = FrameX
		' Set default values in Frame
		FrameX.ParentX = ParentX
		FrameX.EncounterX(0) = EncounterNow
		'If frmMain.Map.EncPointer(frmMain.Tome.MoveToX, frmMain.Tome.MoveToY) > 0 Then
		'If frmMain.Map.EncPointer(Tome.MoveToX, Tome.MoveToY) > 0 Then
		'UPGRADE_WARNING: Couldn't resolve default property of object Tome.MoveToY. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
		'UPGRADE_WARNING: Couldn't resolve default property of object Tome.MoveToX. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
		'UPGRADE_WARNING: Couldn't resolve default property of object Map.EncPointer. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
		If Map.EncPointer(Tome.MoveToX, Tome.MoveToY) > 0 Then
			'Set FrameX.EncounterX(4) = frmMain.Map.Encounters("E" & frmMain.Map.EncPointer(frmMain.Tome.MoveToX, frmMain.Tome.MoveToY))
			'UPGRADE_WARNING: Couldn't resolve default property of object Tome.MoveToY. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
			'UPGRADE_WARNING: Couldn't resolve default property of object Tome.MoveToX. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
			'UPGRADE_WARNING: Couldn't resolve default property of object Map.EncPointer. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
			'UPGRADE_WARNING: Couldn't resolve default property of object Map.Encounters. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
			FrameX.EncounterX(4) = Map.Encounters("E" & Map.EncPointer(Tome.MoveToX, Tome.MoveToY))
		Else
			FrameX.EncounterX(4) = EncounterNow
		End If
		FrameX.CreatureX(0) = CreatureNow
		FrameX.CreatureX(4) = CreatureTarget
		FrameX.ItemX(0) = ItemNow
		FrameX.ItemX(4) = ItemTarget
		FrameX.TriggerX(0) = FrameX.Trigger
		FrameX.TriggerX(4) = FrameX.Trigger
		'FrameX.TileX(0) = frmMain.Tome.MapX
		'UPGRADE_WARNING: Couldn't resolve default property of object Tome.MapX. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
		FrameX.TileX(0) = Tome.MapX
		'FrameX.TileY(0) = frmMain.Tome.MapY
		'UPGRADE_WARNING: Couldn't resolve default property of object Tome.MapY. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
		FrameX.TileY(0) = Tome.MapY
		'FrameX.TileX(4) = frmMain.Tome.MoveToX
		'UPGRADE_WARNING: Couldn't resolve default property of object Tome.MoveToX. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
		FrameX.TileX(4) = Tome.MoveToX
		'FrameX.TileY(4) = frmMain.Tome.MoveToY
		'UPGRADE_WARNING: Couldn't resolve default property of object Tome.MoveToY. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
		FrameX.TileY(4) = Tome.MoveToY
		' Set GlobalSkillLevel (if this is a Skill Trigger)
		If FrameX.Trigger.IsSkill = True Then
			If FrameX.Trigger.SkillPoints > 0 And FrameX.Trigger.Turns > 0 Then
				'frmMain.GlobalSkillLevel = Int(FrameX.Trigger.SkillPoints / FrameX.Trigger.Turns)
				GlobalSkillLevel = Int(FrameX.Trigger.SkillPoints / FrameX.Trigger.Turns)
			Else
				'frmMain.GlobalSkillLevel = 1
				GlobalSkillLevel = 1
			End If
		End If
		' Process statements
		Dim Runes(5) As Byte
		With FrameX
			Do Until .StopExit
				.Ticks = .Ticks + 1
				StmtX = .Trigger.Statements(.StatementNow)
				'UPGRADE_WARNING: Couldn't resolve default property of object ReplaceText(). Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
				sText = ReplaceText((StmtX.Text))
				' Reset FrameNow just in case some other Trigger fired
				FrameNow = FrameX
				' Pre-Process Statements
				.Process = True
				' PROBLEM: Nested If Statements. They SHOULD NOT process anything until it
				' hits an Else, And, Or or EndIf of the same LEVEL OF DEPTH. Currently, it
				' considers statements (and processes them) even if the current IF failed.
				' It has something to do with a Select statement "turning on" the logic
				' to process statements again.
				Select Case StmtX.Statement
					Case 2, 14, 66 ' If, Select, IfText
						If .IfStackTop < 15 Then
							.IfStackTop = .IfStackTop + 1
							.IfStack(.IfStackTop) = bdIfProcess
						Else
							.Process = False
							.Abort = True
							.StopExit = True
						End If
					Case 15 ' Case
						If .IfStackTop > 0 Then
							Select Case .IfStack(.IfStackTop)
								Case bdIfProcess
									.IfStack(.IfStackTop) = bdIfEnd
								Case bdIfSkip
									.IfStack(.IfStackTop) = bdIfProcess
							End Select
						End If
					Case 3, 4, 6 ' Else, ElseIf, Or
						If .IfStackTop > 0 Then
							Select Case .IfStack(.IfStackTop)
								Case bdIfProcess
									.IfStack(.IfStackTop) = bdIfEnd
									.Process = False
								Case bdIfSkip
									.IfStack(.IfStackTop) = bdIfProcess
								Case bdIfEnd
									.Process = False
							End Select
						End If
					Case 7, 16 ' EndIf, EndSelect
						If .IfStackTop > 0 Then
							.IfStackTop = .IfStackTop - 1
						End If
					Case 10 ' Next
						If .LoopStackTop > 0 Then
							' And we have a statement to go back to
							If .LoopStack(.LoopStackTop) > 0 Then
								' Go to the statement pointed to by the stack
								.StatementNow = .LoopStack(.LoopStackTop)
								StmtX = .Trigger.Statements(.StatementNow)
							Else
								.LoopStackTop = .LoopStackTop - 1
							End If
						End If
					Case Else
						' If inside an If loop and failed check, then skip ahead
						If .LoopStackTop > 0 Then
							.Process = (.LoopStack(.LoopStackTop) > 0)
						End If
						If .IfStackTop > 0 Then
							For c = .IfStackTop To 1 Step -1
								If .IfStack(c) <> bdIfProcess Then
									.Process = False
									Exit For
								End If
							Next c
						End If
				End Select
				' Process Statements
				If .Process Then
					' Add Statement to DebugQue
					If GlobalDebugMode = 1 Then
						'frmMain.DebugAdd .Trigger, StmtX
						DebugAdd(.Trigger, StmtX)
					End If
					Select Case StmtX.Statement
						Case 0 ' <None>
							' Do Nothing
						Case 1 ' Label
							' Do Nothing
						Case 2, 4, 5, 6 ' If, ElseIf, And, Or
							'UPGRADE_WARNING: Couldn't resolve default property of object CompareVars(StmtX.B(0), StmtX.B(1), StmtX.B(2), StmtX.B(3), StmtX.B(4)). Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
							If CompareVars(StmtX.B(0), StmtX.B(1), StmtX.B(2), StmtX.B(3), StmtX.B(4)) > 0 Then
								.IfStack(.IfStackTop) = bdIfProcess
							Else
								.IfStack(.IfStackTop) = bdIfSkip
							End If
						Case 3 ' Else
							' Do Nothing
						Case 7 ' End If
							' Do Nothing
						Case 8 ' While
							If .LoopStack(.LoopStackTop) <> .StatementNow Then
								If .LoopStackTop < 15 Then
									.LoopStackTop = .LoopStackTop + 1
									.LoopStack(.LoopStackTop) = .StatementNow
								Else
									.Abort = True
									.StopExit = True
								End If
							End If
							'UPGRADE_WARNING: Couldn't resolve default property of object CompareVars(StmtX.B(0), StmtX.B(1), StmtX.B(2), StmtX.B(3), StmtX.B(4)). Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
							If CompareVars(StmtX.B(0), StmtX.B(1), StmtX.B(2), StmtX.B(3), StmtX.B(4)) < 1 Then
								' All right... lets see if having the statement
								' jump from here does the trick.... - Phule
								.LoopStack(.LoopStackTop) = 0
								.StatementNow = FindMatch(.StatementNow, TrigToFire)
								StmtX = .Trigger.Statements(.StatementNow)
							End If
						Case 9 ' ForEach Object In Object.Collection
							' If this ForEach not on the Stack, put it there
							If .LoopStack(.LoopStackTop) <> .StatementNow Then
								If .LoopStackTop < 15 Then
									.LoopStackTop = .LoopStackTop + 1
									.LoopStack(.LoopStackTop) = .StatementNow
									.EachStack(.LoopStackTop) = 0
								Else
									' We're out of stack space, so have to abort Trigger
									.Abort = True
									.StopExit = True
								End If
							End If
							' Find the Collection looping through
							Select Case StmtX.B(1)
								Case 0 To 4 ' In Encounter
									ObjectX = .EncounterX(StmtX.B(1))
								Case 5 To 9 ' In Item
									ObjectX = .ItemX(StmtX.B(1) - 5)
								Case 10 ' In Tome
									'Set ObjectX = frmMain.Tome
									ObjectX = Tome
								Case 11 To 15 ' In Trigger
									ObjectX = .TriggerX(StmtX.B(1) - 11)
								Case 16 To 20 ' In Creature
									ObjectX = .CreatureX(StmtX.B(1) - 16)
							End Select
							' If first time through, set match
							If .EachStack(.LoopStackTop) = 0 Then
								Select Case StmtX.B(0)
									Case 0 To 2 ' Creatures
										'UPGRADE_WARNING: Couldn't resolve default property of object ObjectX.Creatures. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
										If ObjectX.Creatures.Count > 0 Then
											'UPGRADE_WARNING: Couldn't resolve default property of object ObjectX.Creatures. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
											.EachStack(.LoopStackTop) = ObjectX.Creatures(1).Index
										End If
									Case 3 To 5 ' Items
										'UPGRADE_WARNING: Couldn't resolve default property of object ObjectX.Items. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
										If ObjectX.Items.Count > 0 Then
											'UPGRADE_WARNING: Couldn't resolve default property of object ObjectX.Items. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
											.EachStack(.LoopStackTop) = ObjectX.Items(1).Index
										End If
									Case 6 To 8 ' Triggers
										'UPGRADE_WARNING: Couldn't resolve default property of object ObjectX.Triggers. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
										If ObjectX.Triggers.Count > 0 Then
											'UPGRADE_WARNING: Couldn't resolve default property of object ObjectX.Triggers. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
											.EachStack(.LoopStackTop) = ObjectX.Triggers(1).Index
										End If
								End Select
							End If
							' Find the matching object
							Select Case StmtX.B(0)
								Case 0 To 2 ' Creatures
									'UPGRADE_WARNING: Couldn't resolve default property of object ObjectX.Creatures. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
									For c = 1 To ObjectX.Creatures.Count
										'UPGRADE_WARNING: Couldn't resolve default property of object ObjectX.Creatures. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
										If ObjectX.Creatures(c).Index = .EachStack(.LoopStackTop) Then
											'UPGRADE_WARNING: Couldn't resolve default property of object ObjectX.Creatures. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
											.CreatureX(StmtX.B(0) + 1) = ObjectX.Creatures(c)
											'UPGRADE_WARNING: Couldn't resolve default property of object ObjectX.Creatures. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
											If c < ObjectX.Creatures.Count Then
												'UPGRADE_WARNING: Couldn't resolve default property of object ObjectX.Creatures. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
												.EachStack(.LoopStackTop) = ObjectX.Creatures(c + 1).Index
											Else
												.EachStack(.LoopStackTop) = -1
											End If
											Exit For
										End If
									Next c
									'UPGRADE_WARNING: Couldn't resolve default property of object ObjectX.Creatures. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
									If c > ObjectX.Creatures.Count Then
										' All right... lets see if having the statement
										' jump from here does the trick.... - Phule
										.LoopStack(.LoopStackTop) = 0
										.StatementNow = FindMatch(.StatementNow, TrigToFire)
										StmtX = .Trigger.Statements(.StatementNow)
									End If
								Case 3 To 5 ' Items
									'UPGRADE_WARNING: Couldn't resolve default property of object ObjectX.Items. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
									For c = 1 To ObjectX.Items.Count
										'UPGRADE_WARNING: Couldn't resolve default property of object ObjectX.Items. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
										If ObjectX.Items(c).Index = .EachStack(.LoopStackTop) Then
											'UPGRADE_WARNING: Couldn't resolve default property of object ObjectX.Items. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
											.ItemX(StmtX.B(0) - 2) = ObjectX.Items(c)
											'UPGRADE_WARNING: Couldn't resolve default property of object ObjectX.Items. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
											If c < ObjectX.Items.Count Then
												'UPGRADE_WARNING: Couldn't resolve default property of object ObjectX.Items. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
												.EachStack(.LoopStackTop) = ObjectX.Items(c + 1).Index
											Else
												.EachStack(.LoopStackTop) = -1
											End If
											Exit For
										End If
									Next c
									'UPGRADE_WARNING: Couldn't resolve default property of object ObjectX.Items. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
									If c > ObjectX.Items.Count Then
										' All right... lets see if having the statement
										' jump from here does the trick.... - Phule
										.LoopStack(.LoopStackTop) = 0
										.StatementNow = FindMatch(.StatementNow, TrigToFire)
										StmtX = .Trigger.Statements(.StatementNow)
									End If
								Case 6 To 8 ' Triggers
									'UPGRADE_WARNING: Couldn't resolve default property of object ObjectX.Triggers. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
									For c = 1 To ObjectX.Triggers.Count
										'UPGRADE_WARNING: Couldn't resolve default property of object ObjectX.Triggers. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
										If ObjectX.Triggers(c).Index = .EachStack(.LoopStackTop) Then
											'UPGRADE_WARNING: Couldn't resolve default property of object ObjectX.Triggers. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
											.TriggerX(StmtX.B(0) - 5) = ObjectX.Triggers(c)
											'UPGRADE_WARNING: Couldn't resolve default property of object ObjectX.Triggers. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
											If c < ObjectX.Triggers.Count Then
												'UPGRADE_WARNING: Couldn't resolve default property of object ObjectX.Triggers. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
												.EachStack(.LoopStackTop) = ObjectX.Triggers(c + 1).Index
											Else
												.EachStack(.LoopStackTop) = -1
											End If
											Exit For
										End If
									Next c
									'UPGRADE_WARNING: Couldn't resolve default property of object ObjectX.Triggers. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
									If c > ObjectX.Triggers.Count Then
										' All right... lets see if having the statement
										' jump from here does the trick.... - Phule
										.LoopStack(.LoopStackTop) = 0
										.StatementNow = FindMatch(.StatementNow, TrigToFire)
										StmtX = .Trigger.Statements(.StatementNow)
									End If
							End Select
						Case 10 ' Next
							' Do Nothing
						Case 11 ' Branch
							' Loop and find the matching Label
							For c = 1 To .Trigger.Statements.Count()
								' Find a Label statement
								'UPGRADE_WARNING: Couldn't resolve default property of object FrameX.Trigger.Statements(c).Statement. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
								If .Trigger.Statements.Item(c).Statement = 1 Then
									' If match, then branch there
									'UPGRADE_WARNING: Couldn't resolve default property of object FrameX.Trigger.Statements().B. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
									'UPGRADE_WARNING: Couldn't resolve default property of object CompareVars(FrameX.Trigger.Statements(c).B(0), FrameX.Trigger.Statements(c).B(1), 0, StmtX.B(0), StmtX.B(1)). Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
									If CompareVars(.Trigger.Statements.Item(c).B(0), .Trigger.Statements.Item(c).B(1), 0, StmtX.B(0), StmtX.B(1)) > 0 Then
										.StatementNow = c - 1
										Exit For
									End If
								End If
							Next c
						Case 12 ' Put
							PutVarContext(StmtX.B(5), StmtX.B(6), CompareVars(StmtX.B(0), StmtX.B(1), StmtX.B(2), StmtX.B(3), StmtX.B(4)))
						Case 13 ' Set
							Select Case StmtX.B(0)
								Case 0 ' CreatureNow
									.CreatureX(StmtX.B(0)) = .CreatureX(StmtX.B(1))
									CreatureNow = .CreatureX(StmtX.B(1))
								Case 1 To 3 ' CreatureA, B, C
									.CreatureX(StmtX.B(0)) = .CreatureX(StmtX.B(1))
								Case 4 ' CreatureTarget
									.CreatureX(StmtX.B(0)) = .CreatureX(StmtX.B(1))
									CreatureTarget = .CreatureX(StmtX.B(1))
								Case 5 To 9 ' EncounterNow, A, B, C, Target
									.EncounterX(StmtX.B(0) - 5) = .EncounterX(StmtX.B(1))
								Case 10 ' ItemTarget
									.ItemX(StmtX.B(0) - 10) = .ItemX(StmtX.B(1))
									ItemTarget = .ItemX(StmtX.B(1))
								Case 11 To 13 ' ItemA, B, C
									.ItemX(StmtX.B(0) - 10) = .ItemX(StmtX.B(1))
								Case 14 ' ItemNow
									.ItemX(StmtX.B(0) - 10) = .ItemX(StmtX.B(1))
									ItemNow = .ItemX(StmtX.B(1))
								Case 15 To 19 ' TileNow, TileA, B, C, TileTarget
									.TileX(StmtX.B(0) - 15) = .TileX(StmtX.B(1))
									.TileY(StmtX.B(0) - 15) = .TileY(StmtX.B(1))
								Case 20 To 24 ' TriggerNow, A, B, C, TriggerTarget
									.TriggerX(StmtX.B(0) - 20) = .TriggerX(StmtX.B(1))
							End Select
						Case 14 ' Select
							If IsNumeric(GetVarContext(StmtX.B(0), StmtX.B(1))) Then
								'UPGRADE_WARNING: Couldn't resolve default property of object GetVarContext(). Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
								.CaseStack(.IfStackTop) = GetVarContext(StmtX.B(0), StmtX.B(1))
							Else
								.CaseStack(.IfStackTop) = 0
							End If
						Case 15 ' Case
							'UPGRADE_WARNING: Couldn't resolve default property of object GetVarContext(StmtX.B(0), StmtX.B(1)). Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
							If GetVarContext(StmtX.B(0), StmtX.B(1)) = .CaseStack(.IfStackTop) Then
								.IfStack(.IfStackTop) = bdIfProcess
							Else
								.IfStack(.IfStackTop) = bdIfSkip
							End If
						Case 16 ' EndSelect
							' Do Nothing
						Case 17 ' Exit
							Select Case StmtX.B(0)
								Case 0 ' Trigger
									.StopExit = True
								Case 1 ' Loop
									.LoopStack(.LoopStackTop) = 0
								Case 2 ' Abort
									.StopExit = True
									.Abort = True
							End Select
						Case 18 ' CopyCreature
							' Indicates if copied to Encounter or Party
							Found = False
							' Copy To
							Select Case StmtX.B(1)
								Case 0 To 4 ' Encounter
									ObjectX = .EncounterX(StmtX.B(1)).AddCreature
									Found = True
								Case 5 ' Party
									'Set ObjectX = frmMain.Tome.AddCreature
									'UPGRADE_WARNING: Couldn't resolve default property of object Tome.AddCreature. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
									ObjectX = Tome.AddCreature
									Found = True
								Case 6 To 10 ' Trigger
									ObjectX = .TriggerX(StmtX.B(1) - 6).AddCreature
							End Select
							For	Each CreatureX In .Trigger.Creatures
								If CreatureX.Index = StmtX.B(0) Then
									'UPGRADE_WARNING: Couldn't resolve default property of object ObjectX.Copy. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
									ObjectX.Copy(.Trigger.Creatures.Item("X" & StmtX.B(0)))
									'UPGRADE_WARNING: Couldn't resolve default property of object ObjectX. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
									frmMain.LoadCreaturePic(ObjectX)
									Exit For
								End If
							Next CreatureX
							If Found = True Then
								If frmMain.picGrid.Visible = True Then
									frmMain.CombatArrays()
									frmMain.CombatDraw()
								Else
									Select Case StmtX.B(1)
										Case 0 To 4
											'modDungeonMaker.PositionCreature frmMain.Map, .EncounterX(StmtX.B(1)), ObjectX
											'UPGRADE_WARNING: Couldn't resolve default property of object ObjectX. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
											'UPGRADE_WARNING: Couldn't resolve default property of object Map. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
											modDungeonMaker.PositionCreature(Map, .EncounterX(StmtX.B(1)), ObjectX)
										Case 5
											' [Titi 2.4.9] set the coordinates of the new creatures
											'UPGRADE_WARNING: Couldn't resolve default property of object ObjectX.MapX. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
											'UPGRADE_WARNING: Couldn't resolve default property of object Tome.MapX. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
											ObjectX.MapX = Tome.MapX
											'UPGRADE_WARNING: Couldn't resolve default property of object ObjectX.MapY. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
											'UPGRADE_WARNING: Couldn't resolve default property of object Tome.MapY. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
											ObjectX.MapY = Tome.MapY
											frmMain.MenuDrawParty()
									End Select
								End If
							End If
						Case 19 ' CopyItem
							Select Case StmtX.B(1)
								Case 0 To 4
									ObjectX = .CreatureX(StmtX.B(1)).AddItem
								Case 5 To 9
									ObjectX = .EncounterX(StmtX.B(1) - 5).AddItem
								Case 10 To 14
									ObjectX = .ItemX(StmtX.B(1) - 10).AddItem
								Case 15 To 19
									ObjectX = .TriggerX(StmtX.B(1) - 15).AddItem
							End Select
							For	Each ItemX In .Trigger.Items
								If ItemX.Index = StmtX.B(0) Then
									'UPGRADE_WARNING: Couldn't resolve default property of object ObjectX.Copy. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
									ObjectX.Copy(.Trigger.Items.Item("I" & StmtX.B(0)))
									'UPGRADE_WARNING: Couldn't resolve default property of object ObjectX.IsReady. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
									ObjectX.IsReady = False
									'UPGRADE_WARNING: Couldn't resolve default property of object ObjectX. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
									frmMain.LoadItemPic(ObjectX)
								End If
							Next ItemX
						Case 20 ' CopyTrigger
							Select Case StmtX.B(1)
								Case 0 To 4
									ObjectX = .CreatureX(StmtX.B(1)).AddTrigger
								Case 5 To 9
									ObjectX = .EncounterX(StmtX.B(1) - 5).AddTrigger
								Case 10 To 14
									ObjectX = .ItemX(StmtX.B(1) - 10).AddTrigger
								Case 15
									'Set ObjectX = frmMain.Tome.AddTrigger
									'UPGRADE_WARNING: Couldn't resolve default property of object Tome.AddTrigger. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
									ObjectX = Tome.AddTrigger
								Case 16 To 20
									ObjectX = .TriggerX(StmtX.B(1) - 16).AddTrigger
							End Select
							For	Each TriggerX In .Trigger.Triggers
								If TriggerX.Index = StmtX.B(0) Then
									'UPGRADE_WARNING: Couldn't resolve default property of object ObjectX.Copy. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
									ObjectX.Copy(.Trigger.Triggers.Item("T" & StmtX.B(0)))
									Exit For
								End If
							Next TriggerX
						Case 21 ' CopyTile
							' Validate X, Y coordinate
							'UPGRADE_WARNING: Couldn't resolve default property of object GetVarContext(). Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
							X = GetVarContext(StmtX.B(0), StmtX.B(1))
							'UPGRADE_WARNING: Couldn't resolve default property of object GetVarContext(). Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
							Y = GetVarContext(StmtX.B(3), StmtX.B(4))
							'UPGRADE_WARNING: Couldn't resolve default property of object Map.Height. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
							'UPGRADE_WARNING: Couldn't resolve default property of object Map.Width. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
							If PointIn(X, Y, 0, 0, Map.Width, Map.Height) Then
								' Locate Tile Name in Current Map
								'UPGRADE_WARNING: Couldn't resolve default property of object Map.Tiles. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
								For	Each TileX In Map.Tiles
									If TileX.Name = sText Then
										Select Case StmtX.B(2)
											Case 0 ' Bottom
												'UPGRADE_WARNING: Couldn't resolve default property of object Map.BottomTile. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
												Map.BottomTile(X, Y) = TileX.Index
												'UPGRADE_WARNING: Couldn't resolve default property of object Map.BottomFlip. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
												Map.BottomFlip(X, Y) = False
											Case 1 ' Bottom (Flip)
												'UPGRADE_WARNING: Couldn't resolve default property of object Map.BottomTile. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
												Map.BottomTile(X, Y) = TileX.Index
												'UPGRADE_WARNING: Couldn't resolve default property of object Map.BottomFlip. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
												Map.BottomFlip(X, Y) = True
											Case 2 ' Middle
												'UPGRADE_WARNING: Couldn't resolve default property of object Map.MiddleTile. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
												Map.MiddleTile(X, Y) = TileX.Index
												'UPGRADE_WARNING: Couldn't resolve default property of object Map.MiddleFlip. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
												Map.MiddleFlip(X, Y) = False
											Case 3 ' Middle (Flip)
												'UPGRADE_WARNING: Couldn't resolve default property of object Map.MiddleTile. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
												Map.MiddleTile(X, Y) = TileX.Index
												'UPGRADE_WARNING: Couldn't resolve default property of object Map.MiddleFlip. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
												Map.MiddleFlip(X, Y) = True
											Case 4 ' Top
												'UPGRADE_WARNING: Couldn't resolve default property of object Map.TopTile. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
												Map.TopTile(X, Y) = TileX.Index
												'UPGRADE_WARNING: Couldn't resolve default property of object Map.TopFlip. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
												Map.TopFlip(X, Y) = False
											Case 5 ' Top (Flip)
												'UPGRADE_WARNING: Couldn't resolve default property of object Map.TopTile. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
												Map.TopTile(X, Y) = TileX.Index
												'UPGRADE_WARNING: Couldn't resolve default property of object Map.TopFlip. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
												Map.TopFlip(X, Y) = True
										End Select
										RefreshMap = True
										Exit For
									End If
								Next TileX
							End If
						Case 22 ' CopyText
							PutVarContext(StmtX.B(0), StmtX.B(1), sText)
						Case 23 ' Runes
							' Do Nothing
						Case 24 ' AddFactoid
							'UPGRADE_WARNING: Couldn't resolve default property of object Tome.AddFactoid. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
							FactoidX = Tome.AddFactoid
							FactoidX.Text = sText
						Case 25 ' AddJournalEntry
							'UPGRADE_WARNING: Couldn't resolve default property of object Tome.AddJournal. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
							Tome.AddJournal.Text = sText
							If frmMain.picJournal.Visible = True And JournalMode = 0 Then
								JournalEntryLoad()
								JournalShow()
							End If
						Case 26 ' Destroy
							' Indicates if destroying in Party or Encounters
							Found = False
							' Object contained in
							Select Case StmtX.B(1)
								Case 0 To 4
									ObjectX = .EncounterX(StmtX.B(1))
									Found = True
								Case 5 To 9
									ObjectX = .ItemX(StmtX.B(1) - 5)
								Case 10
									ObjectX = Tome
									Found = True
								Case 11 To 15
									ObjectX = .TriggerX(StmtX.B(1) - 11)
								Case 16 To 20
									ObjectX = .CreatureX(StmtX.B(1) - 16)
							End Select
							' Object to destroy
							Select Case StmtX.B(0)
								Case 0 To 2, 9, 10 ' Creature
									If StmtX.B(0) = 9 Then
										c = 0
									ElseIf StmtX.B(0) = 10 Then 
										c = 4
									Else
										c = StmtX.B(0) + 1
									End If
									'UPGRADE_WARNING: Couldn't resolve default property of object ObjectX.Creatures. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
									For	Each CreatureX In ObjectX.Creatures
										If CreatureX.Index = .CreatureX(c).Index Then
											'UPGRADE_WARNING: Couldn't resolve default property of object ObjectX.RemoveCreature. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
											ObjectX.RemoveCreature("X" & .CreatureX(c).Index)
											Exit For
										End If
									Next CreatureX
									If Found = True Then
										' If combat is active, reset the combat creatures
										If frmMain.picGrid.Visible = True Then
											frmMain.CombatArrays()
											frmMain.CombatDraw()
										ElseIf StmtX.B(1) = 10 Then 
											' If destroying from Party, redraw the Party
											frmMain.MenuDrawParty()
										End If
									End If
								Case 3 To 5, 11, 12 ' Item
									If StmtX.B(0) = 11 Then
										c = 0
									ElseIf StmtX.B(0) = 12 Then 
										c = 4
									Else
										c = StmtX.B(0) - 2
									End If
									'UPGRADE_WARNING: Couldn't resolve default property of object ObjectX.Items. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
									For	Each ItemX In ObjectX.Items
										If ItemX.Index = .ItemX(c).Index Then
											'UPGRADE_WARNING: Couldn't resolve default property of object ObjectX.RemoveItem. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
											ObjectX.RemoveItem("I" & .ItemX(c).Index)
											Exit For
										End If
									Next ItemX
								Case 6 To 8, 13, 14 ' Trigger
									If StmtX.B(0) = 13 Then
										c = 0
									ElseIf StmtX.B(0) = 14 Then 
										c = 4
									Else
										c = StmtX.B(0) - 5
									End If
									'UPGRADE_WARNING: Couldn't resolve default property of object ObjectX.Triggers. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
									For	Each TriggerX In ObjectX.Triggers
										If TriggerX.Index = .TriggerX(c).Index Then
											'UPGRADE_WARNING: Couldn't resolve default property of object ObjectX.RemoveTrigger. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
											ObjectX.RemoveTrigger("T" & .TriggerX(c).Index)
											Exit For
										End If
									Next TriggerX
							End Select
						Case 27 ' CombatApplyDamage
							'UPGRADE_WARNING: Couldn't resolve default property of object GetVarContext(StmtX.B(3), StmtX.B(4)). Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
							frmMain.CombatApplyDamage(.CreatureX(StmtX.B(0)), GetVarContext(StmtX.B(3), StmtX.B(4)))
						Case 28 ' CombatRollAttack
							'UPGRADE_WARNING: Couldn't resolve default property of object GetVarContext(StmtX.B(0), StmtX.B(1)). Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
							c = frmMain.CombatRollAttack(StmtX.B(5), GetVarContext(StmtX.B(0), StmtX.B(1)))
							If c = False Then
								.Fail = 1
							End If
						Case 29 ' CombatRollDamage
							c = frmMain.CombatRollDamage(StmtX.B(0), StmtX.B(1), StmtX.B(5), StmtX.B(2) > 0, 0)
							.Fail = c
						Case 30 ' DestroyFactoid
							'UPGRADE_WARNING: Couldn't resolve default property of object Tome.Factoids. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
							For	Each FactoidX In Tome.Factoids
								If InStr(UCase(FactoidX.Text), UCase(sText)) > 0 Then
									'UPGRADE_WARNING: Couldn't resolve default property of object Tome.RemoveFactoid. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
									Tome.RemoveFactoid("F" & FactoidX.Index)
								End If
							Next FactoidX
						Case 31 ' TargetCreature
							' Target Creature into CreatureX (within Range)
							CreatureX = New Creature
							If IsNumeric(GetVarContext(StmtX.B(3), StmtX.B(4))) = True Then
								'UPGRADE_WARNING: Couldn't resolve default property of object GetVarContext(). Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
								If frmMain.TargetCreature(CreatureX, GetVarContext(StmtX.B(3), StmtX.B(4)), StmtX.B(1), StmtX.B(2)) = False Then
									.StopExit = True
									.Abort = True
								End If
							Else
								If frmMain.TargetCreature(CreatureX, 11, StmtX.B(1), StmtX.B(2)) = False Then
									.StopExit = True
									.Abort = True
								End If
							End If
							.CreatureX(StmtX.B(0) + 1) = CreatureX
							If StmtX.B(0) = 3 Then ' If setting CreatureTarget, then set frmMain
								CreatureTarget = CreatureX
							End If
						Case 32 ' TargetItem
							' Target Item into ItemX Using CreatureWithTurn Inventory or Encounter
							If frmMain.TargetItem(ItemX) = False Then
								.StopExit = True
								.Abort = True
							End If
							FrameNow = FrameX
							.ItemX(StmtX.B(0) + 1) = ItemX
						Case 33 ' PlaySound
							'UPGRADE_WARNING: Couldn't resolve default property of object Tome. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
							Call PlaySoundFile(StmtX.Text, Tome, True, IIf(StmtX.B(0) > 0, 5, 0))
						Case 34 ' PlayMusic
							'UPGRADE_WARNING: Couldn't resolve default property of object Tome.LoadPath. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
							Call PlayMusic(StmtX.Text, frmMain, Tome.LoadPath)
							GlobalMusicName = StmtX.Text
						Case 35 ' MoveParty
							'For Each MapX In frmMain.Area.Plot.Maps
							For	Each MapX In Area.Plot.Maps
								If MapX.Index = StmtX.B(5) Then
									'UPGRADE_WARNING: Couldn't resolve default property of object GetVarContext(StmtX.B(3), StmtX.B(4)). Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
									'UPGRADE_WARNING: Couldn't resolve default property of object GetVarContext(). Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
									'UPGRADE_WARNING: Couldn't resolve default property of object Area.Index. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
									frmMain.TomeStartArea(Area.Index, (MapX.Index), -1, GetVarContext(StmtX.B(0), StmtX.B(1)), GetVarContext(StmtX.B(3), StmtX.B(4)))
									Exit For
								End If
							Next MapX
						Case 36 ' DialogShow
							Select Case StmtX.B(0)
								Case 0 ' Normal
									DialogSetUp(modGameGeneral.DLGTYPE.bdDlgNoReply)
									DialogShow(BreakText(sText, 1), BreakText(sText, 2))
								Case 1 ' Brief TextBox
									DialogBrief(BreakText(sText, 1), BreakText(sText, 2))
								Case 2 ' Reply WithPick
									DialogSetUp(modGameGeneral.DLGTYPE.bdDlgWithReply)
									ReplyTxt = sText
								Case 3 ' Reply WithText
									DialogSetUp(modGameGeneral.DLGTYPE.bdDlgReplyText)
									ReplyTxt = sText
								Case 4, 5 ' Brief TextLine
									frmMain.MessageShow(BreakText(sText, 2), 0)
							End Select
						Case 37 ' DialogReply
							DialogReplyAdd(sText)
						Case 38 ' DialogAccept
							c = DialogShow(BreakText(ReplyTxt, 1), BreakText(ReplyTxt, 2)) + 1
							PutVarContext(StmtX.B(0), StmtX.B(1), c)
						Case 39 ' DialogDice
							DialogDice(BreakText(StmtX.Text, 1), BreakText(sText, 2), StmtX.B(2), c)
							PutVarContext(StmtX.B(0), StmtX.B(1), c)
						Case 40 ' DialogHide
							DialogHide()
						Case 41 ' CutScene
							frmMain.CutScene(StmtX.B(0), BreakText(sText, 1), BreakText(sText, 2))
							RefreshMap = True
						Case 42 ' SorceryQueRunes
							For c = 0 To 5
								Runes(c) = StmtX.B(c)
							Next c
							c = frmMain.SorceryMatchRunes(.CreatureX(0), Runes, StmtX.B(6) = 0)
							If c <> 0 And .Fail = 0 Then
								.Fail = 1
							End If
						Case 43 ' CombatMove
							If frmMain.picGrid.Visible = True Then
								frmMain.CombatMove(StmtX.B(0), StmtX.B(1))
							Else
								.StopExit = True
								.Abort = True
							End If
						Case 44 ' CombatAttack
							frmMain.CombatAttack(StmtX.B(0), sText, StmtX.B(3), StmtX.B(4))
						Case 45 ' CombatTarget
							If frmMain.picGrid.Visible = True Then
								FrameNow.CreatureX(4) = frmMain.CombatTarget(FrameNow.CreatureX(0), StmtX.B(0))
								CreatureTarget = FrameNow.CreatureX(4)
							Else
								.StopExit = True
								.Abort = True
							End If
						Case 46 ' PlaySFX
							frmMain.PlaySFX((StmtX.Text), StmtX.B(0), StmtX.B(1), StmtX.B(2), StmtX.B(3))
						Case 47 ' Sorcery
							If IsNumeric(GetVarContext(StmtX.B(0), StmtX.B(1))) Then
								'UPGRADE_WARNING: Couldn't resolve default property of object GetVarContext(). Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
								GlobalSpellFizzleChance = GetVarContext(StmtX.B(0), StmtX.B(1))
							Else
								GlobalSpellFizzleChance = 0
							End If
						Case 48 ' DialogBuySell
							If IsNumeric(GetVarContext(StmtX.B(0), StmtX.B(1))) Then
								'UPGRADE_WARNING: Couldn't resolve default property of object GetVarContext(). Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
								c = GetVarContext(StmtX.B(0), StmtX.B(1))
							Else
								c = 0
							End If
							'frmMain.DialogBuySell .Trigger, StmtX.B(2), c
							DialogBuySell(.Trigger, StmtX.B(2), c)
						Case 49 ' Not Used
						Case 50 ' ExecTrigger
							If StmtX.B(0) > 0 Then
								For	Each TriggerX In .Trigger.Triggers
									If TriggerX.Index = StmtX.B(0) Then
										.Fail = Not FireTrigger(.Trigger, .Trigger.Triggers.Item("T" & StmtX.B(0)))
										Exit For
									End If
								Next TriggerX
							End If
							' Reset pointers back to current Trigger
							FrameNow = FrameX
						Case 51 ' DialogAcceptText
							c = DialogShow(BreakText(ReplyTxt, 1), BreakText(ReplyTxt, 2))
							' If no text is entered, then set Local.Fail to True
							If LenStr(DialogText) < 1 Then
								.Fail = 1
							End If
							PutVarContext(StmtX.B(0), StmtX.B(1), DialogText)
						Case 52 ' CombatStart
							frmMain.CombatStart()
						Case 53 ' Let
							PutVarContext(StmtX.B(0), StmtX.B(1), GetVarContext(StmtX.B(3), StmtX.B(4)))
						Case 54 ' Comment
							' Do nothing
						Case 55 ' RandomizeEncounter
							' Find Encounter matching Name on current Map
							'For Each EncounterX In frmMain.Map.Encounters
							'UPGRADE_WARNING: Couldn't resolve default property of object Map.Encounters. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
							For	Each EncounterX In Map.Encounters
								If InStr(UCase(EncounterX.Name), UCase(sText)) > 0 Then
									'UPGRADE_WARNING: Couldn't resolve default property of object Map. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
									MakeEncounter(Map, EncounterX)
								End If
							Next EncounterX
						Case 56 ' RandomizeTheme
							'UPGRADE_WARNING: Couldn't resolve default property of object Map.Themes. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
							For	Each ThemeX In Map.Themes
								If InStr(UCase(ThemeX.Name), UCase(sText)) > 0 Then
									'UPGRADE_WARNING: Couldn't resolve default property of object Map. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
									ThemeMap(Map, ThemeX)
								End If
							Next ThemeX
						Case 57 ' AwardExperience
							If StmtX.B(0) = 0 Then
								' Award to the Party
								'UPGRADE_WARNING: Couldn't resolve default property of object GetVarContext(). Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
								frmMain.AwardExperienceToParty(GetVarContext(StmtX.B(5), StmtX.B(6)))
							Else
								' Award to a single Creature
								'UPGRADE_WARNING: Couldn't resolve default property of object GetVarContext(StmtX.B(5), StmtX.B(6)). Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
								frmMain.AwardExperienceToCreature(.CreatureX(StmtX.B(0) - 1), GetVarContext(StmtX.B(5), StmtX.B(6)))
							End If
						Case 58 ' MoveItem
							' From Object
							Select Case StmtX.B(1)
								Case 0 To 4 ' CreatureX
									ObjectX = .CreatureX(StmtX.B(1))
								Case 5 To 9 ' EncounterX
									ObjectX = .EncounterX(StmtX.B(1) - 5)
								Case 10 To 14 ' ItemX
									ObjectX = .ItemX(StmtX.B(1) - 10)
								Case 15 To 19 ' TriggerX
									ObjectX = .TriggerX(StmtX.B(1) - 15)
							End Select
							' To Object
							Select Case StmtX.B(3)
								Case 0 To 4 ' CreatureX
									ObjectY = .CreatureX(StmtX.B(3))
								Case 5 To 9 ' EncounterX
									ObjectY = .EncounterX(StmtX.B(3) - 5)
								Case 10 To 14 ' ItemX
									ObjectY = .ItemX(StmtX.B(3) - 10)
								Case 15 To 19 ' TriggerX
									ObjectY = .TriggerX(StmtX.B(3) - 15)
							End Select
							' Locate ItemX In ObjectX
							'UPGRADE_WARNING: Couldn't resolve default property of object ObjectX.Items. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
							For	Each ItemX In ObjectX.Items
								If ItemX.Name = .ItemX(StmtX.B(0)).Name And ItemX.Index = .ItemX(StmtX.B(0)).Index Then
									'UPGRADE_WARNING: Couldn't resolve default property of object ObjectY.AddItem. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
									ItemY = ObjectY.AddItem
									ItemY.Copy(ItemX)
									ItemY.IsReady = False
									If StmtX.B(2) = 0 Then
										'UPGRADE_WARNING: Couldn't resolve default property of object ObjectX.RemoveItem. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
										ObjectX.RemoveItem("I" & ItemX.Index)
									End If
									Exit For
								End If
							Next ItemX
						Case 59 ' AddQuest
							'UPGRADE_WARNING: Couldn't resolve default property of object Tome.AddJournal. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
							Tome.AddJournal.Text = "Quest|" & sText
							If JournalMode = 1 And frmMain.picJournal.Visible = True Then
								JournalQuestsLoad()
								JournalShow()
							End If
						Case 60 ' RemoveQuest
							' With the new Add / Delete buttons in the journal, indices are likely to change.
							' In that case, JournalX.Index is not the actual index
							'               c = 0
							'UPGRADE_WARNING: Couldn't resolve default property of object Tome.Journals. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
							For	Each JournalX In Tome.Journals
								'                   c = c + 1
								If InStr(UCase(JournalX.Text), UCase(BreakText(sText, 1))) > 0 Then
									'UPGRADE_WARNING: Couldn't resolve default property of object Tome.RemoveJournal. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
									Tome.RemoveJournal("J" & JournalX.Index) 'c ' JournalX.Index ' <-- [Titi] 2.4.6 ; with JournalX.Index, probable Run-Time error 5
									' [Titi 2.4.8] reverted to the original numbering "J" & index to allow removal in creator too
								End If
							Next JournalX
							If JournalMode = 1 And frmMain.picJournal.Visible = True Then
								JournalQuestsLoad()
								JournalShow()
							End If
						Case 61 ' RemoveTopic [Titi 2.4.9]
							For	Each ConvoX In CreatureNow.Conversations
								For	Each TopicX In ConvoX
									If InStr(UCase(TopicX.Say), UCase(BreakText(sText, 1))) > 0 Then
										'UPGRADE_WARNING: Couldn't resolve default property of object CreatureNow.Conversations().RemoveTopic. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
										CreatureNow.Conversations.Item(ConvoX.Index).RemoveTopic("Q" & TopicX.Index)
									End If
								Next TopicX
							Next ConvoX
						Case 62 ' CombatAttackWithWeapon
							frmMain.CombatAttack(62, "", 0, 0)
						Case 63 ' CombatAttackWithSpecial
							frmMain.CombatAttack(63, sText, StmtX.B(5), StmtX.B(6))
						Case 64 ' MovePartyMapName
							' Find Area in Tome
							'UPGRADE_WARNING: Couldn't resolve default property of object Tome.Areas. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
							For	Each AreaX In Tome.Areas
								If AreaX.Index = StmtX.B(2) Then
									'UPGRADE_WARNING: Couldn't resolve default property of object Area.Index. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
									If AreaX.Index <> Area.Index Then
										'UPGRADE_WARNING: Couldn't resolve default property of object GetVarContext(StmtX.B(3), StmtX.B(4)). Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
										'UPGRADE_WARNING: Couldn't resolve default property of object GetVarContext(). Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
										frmMain.TomeStartArea((AreaX.Index), 0, -1, GetVarContext(StmtX.B(0), StmtX.B(1)), GetVarContext(StmtX.B(3), StmtX.B(4)), StmtX.Text)
									Else
										' Find Map Name in Area
										'UPGRADE_WARNING: Couldn't resolve default property of object Area.Plot. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
										For	Each MapX In Area.Plot.Maps
											If MapX.Name = StmtX.Text Then
												'UPGRADE_WARNING: Couldn't resolve default property of object GetVarContext(StmtX.B(3), StmtX.B(4)). Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
												'UPGRADE_WARNING: Couldn't resolve default property of object GetVarContext(). Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
												frmMain.TomeStartArea((AreaX.Index), (MapX.Index), -1, GetVarContext(StmtX.B(0), StmtX.B(1)), GetVarContext(StmtX.B(3), StmtX.B(4)))
											End If
										Next MapX
										Exit For
									End If
								End If
							Next AreaX
						Case 65 ' MoveCreature
							' Determines if need to refresh EncounterNow or Party
							Found = False
							' From Object
							Select Case StmtX.B(1)
								Case 0 To 4 ' CreatureX
									ObjectX = .CreatureX(StmtX.B(1))
								Case 5 To 9 ' EncounterX
									ObjectX = .EncounterX(StmtX.B(1) - 5)
									Found = True
								Case 10 To 14 ' ItemX
									ObjectX = .ItemX(StmtX.B(1) - 10)
								Case 15 To 19 ' TriggerX
									ObjectX = .TriggerX(StmtX.B(1) - 15)
								Case 20
									ObjectX = Tome
									Found = True
							End Select
							' To Object
							Select Case StmtX.B(3)
								Case 0 To 4 ' CreatureX
									ObjectY = .CreatureX(StmtX.B(3))
								Case 5 To 9 ' EncounterX
									ObjectY = .EncounterX(StmtX.B(3) - 5)
									Found = True
								Case 10 To 14 ' ItemX
									ObjectY = .ItemX(StmtX.B(3) - 10)
								Case 15 To 19 ' TriggerX
									ObjectY = .TriggerX(StmtX.B(3) - 15)
								Case 20
									ObjectY = Tome
									Found = True
							End Select
							' Locate CreatureX In ObjectX
							'UPGRADE_WARNING: Couldn't resolve default property of object ObjectX.Creatures. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
							For	Each CreatureX In ObjectX.Creatures
								If CreatureX.Name = .CreatureX(StmtX.B(0)).Name And CreatureX.Index = .CreatureX(StmtX.B(0)).Index Then
									'UPGRADE_WARNING: Couldn't resolve default property of object ObjectY.AddCreature. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
									CreatureY = ObjectY.AddCreature
									CreatureY.Copy(CreatureX)
									CreatureY.Initiative = CreatureX.Initiative
									CreatureY.Row = CreatureX.Row
									CreatureY.Col = CreatureX.Col
									frmMain.LoadCreaturePic(CreatureY)
									If StmtX.B(2) = 0 Then
										'UPGRADE_WARNING: Couldn't resolve default property of object ObjectX.RemoveCreature. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
										ObjectX.RemoveCreature("X" & CreatureX.Index)
									End If
									Exit For
								End If
							Next CreatureX
							' Reset combat screen if move to/from Encounter or Party
							If Found = True Then
								' If combat is active, reset the combat creatures
								If frmMain.picGrid.Visible = True Then
									frmMain.CombatArrays()
									frmMain.CombatDraw()
								Else
									' If moving to Encounters, position the Creature
									Select Case StmtX.B(3)
										Case 5 To 9 ' Encounters
											'UPGRADE_WARNING: Couldn't resolve default property of object ObjectY. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
											'UPGRADE_WARNING: Couldn't resolve default property of object Map. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
											modDungeonMaker.PositionCreature(Map, ObjectY, CreatureY)
									End Select
									' If moving to/from Party, redraw the Party
									If StmtX.B(3) = 20 Or StmtX.B(1) = 20 Then
										frmMain.MenuDrawParty()
									End If
								End If
							End If
						Case 66 ' IfText
							'UPGRADE_WARNING: Couldn't resolve default property of object CompareVars(StmtX.B(0), StmtX.B(1), StmtX.B(2), StmtX.B(3), StmtX.B(4), sText). Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
							If CompareVars(StmtX.B(0), StmtX.B(1), StmtX.B(2), StmtX.B(3), StmtX.B(4), sText) > 0 Then
								.IfStack(.IfStackTop) = bdIfProcess
							Else
								.IfStack(.IfStackTop) = bdIfSkip
							End If
						Case 67 ' TargetEncounter
							c = False
							'UPGRADE_WARNING: Couldn't resolve default property of object Area.Plot. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
							For	Each MapX In Area.Plot.Maps
								For	Each EncounterX In MapX.Encounters
									If EncounterX.Name Like sText Then
										.EncounterX(4) = EncounterX
										c = True
										Exit For
									End If
								Next EncounterX
								If c = True Then
									Exit For
								End If
							Next MapX
							If c = False Then
								.Fail = 1
							End If
						Case 68 ' TargetTile
							' Validate X, Y coordinate
							'UPGRADE_WARNING: Couldn't resolve default property of object GetVarContext(). Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
							X = GetVarContext(StmtX.B(0), StmtX.B(1))
							'UPGRADE_WARNING: Couldn't resolve default property of object GetVarContext(). Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
							Y = GetVarContext(StmtX.B(3), StmtX.B(4))
							.Fail = 0
							'UPGRADE_WARNING: Couldn't resolve default property of object Map.Height. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
							'UPGRADE_WARNING: Couldn't resolve default property of object Map.Width. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
							If PointIn(X, Y, 0, 0, Map.Width, Map.Height) Then
								' [Titi 2.4.6] - now set TileTarget
								.TileX(4) = X : .TileY(4) = Y
							Else
								'UPGRADE_WARNING: Couldn't resolve default property of object Tome.MapX. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
								.TileX(4) = Tome.MapX
								'UPGRADE_WARNING: Couldn't resolve default property of object Tome.MapY. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
								.TileY(4) = Tome.MapY
								.Fail = 1
							End If
						Case 69 ' PlayVideo
							'UPGRADE_WARNING: Couldn't resolve default property of object Tome.LoadPath. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
							frmMain.CutSceneVideo((StmtX.Text), Tome.LoadPath)
						Case 70 ' Find
							' Find the Collection looping through
							Select Case StmtX.B(1)
								Case 0 To 4 ' In Encounter
									ObjectX = .EncounterX(StmtX.B(1))
								Case 5 To 9 ' In Item
									ObjectX = .ItemX(StmtX.B(1) - 5)
								Case 10 ' In Tome
									'Set ObjectX = frmMain.Tome
									ObjectX = Tome
								Case 11 To 15 ' In Trigger
									ObjectX = .TriggerX(StmtX.B(1) - 11)
								Case 16 To 20 ' In Creature
									ObjectX = .CreatureX(StmtX.B(1) - 16)
							End Select
							' Loop through and find the Thing named
							.FoundIt = 0
							Select Case StmtX.B(0)
								Case 0 To 2 ' Creatures
									'UPGRADE_WARNING: Couldn't resolve default property of object ObjectX.Creatures. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
									For	Each CreatureX In ObjectX.Creatures
										If InStr(UCase(CreatureX.Name), UCase(sText)) > 0 Or InStr(UCase(sText), UCase(CreatureX.Name)) > 0 Then
											.CreatureX(StmtX.B(0) + 1) = CreatureX
											.FoundIt = 1
											Exit For
										End If
									Next CreatureX
								Case 3 To 5 ' Items
									'UPGRADE_WARNING: Couldn't resolve default property of object ObjectX.Items. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
									For	Each ItemX In ObjectX.Items
										If InStr(UCase(ItemX.Name), UCase(sText)) > 0 Or InStr(UCase(sText), UCase(ItemX.Name)) > 0 Then
											.ItemX(StmtX.B(0) - 2) = ItemX
											.FoundIt = 1
											Exit For
										End If
									Next ItemX
								Case 6 To 8 ' Triggers
									'UPGRADE_WARNING: Couldn't resolve default property of object ObjectX.Triggers. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
									For	Each TriggerX In ObjectX.Triggers
										If InStr(UCase(TriggerX.Name), UCase(sText)) > 0 Or InStr(UCase(sText), UCase(TriggerX.Name)) > 0 Then
											.TriggerX(StmtX.B(0) - 5) = TriggerX
											.FoundIt = 1
											Exit For
										End If
									Next TriggerX
							End Select
						Case 71 ' CallTrigger
							Select Case StmtX.B(1)
								Case 0 To 4 ' Creatures
									ObjectX = .CreatureX(StmtX.B(1))
								Case 5 To 9 ' Encounters
									ObjectX = .EncounterX(StmtX.B(1) - 5)
								Case 10 To 14 ' Items
									ObjectX = .ItemX(StmtX.B(1) - 10)
								Case 15 ' Party
									'Set ObjectX = frmMain.Tome
									ObjectX = Tome
								Case 16 To 20 ' Triggers
									ObjectX = .TriggerX(StmtX.B(1) - 16)
							End Select
							.Fail = Not FireTrigger(ObjectX, .TriggerX(StmtX.B(0)))
							' Reset pointers back to current Trigger
							FrameNow = FrameX
						Case 72 ' CombatAnimation 'Text' For [B0] Frames [B1]
							c = frmMain.LoadSFXPic((StmtX.Text))
							' If loaded and in combat...
							If c > -1 And frmMain.picGrid.Visible = True Then
								' Point to the correct Creature
								ObjectX = .CreatureX(StmtX.B(0))
								Cells = StmtX.B(1) + 2 : X = 0
								Width = VB6.PixelsToTwipsX(frmMain.picSFXPic(c).Width) / (Cells * 2)
								Height = VB6.PixelsToTwipsY(frmMain.picSFXPic(c).Height) / 2
								'UPGRADE_WARNING: Couldn't resolve default property of object ObjectX.Pic. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
								frmMain.picCPic(ObjectX.Pic).Width = VB6.TwipsToPixelsX(Width * 2)
								'UPGRADE_WARNING: Couldn't resolve default property of object ObjectX.Pic. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
								frmMain.picCPic(ObjectX.Pic).Height = VB6.TwipsToPixelsY(Height * 4)
								Do Until X = Cells
									' Timer set here so it takes into account drawing time
									t1 = VB.Timer()
									' Draw Normal and Flip
									'UPGRADE_WARNING: Couldn't resolve default property of object ObjectX.Pic. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
									'UPGRADE_ISSUE: PictureBox method picCPic.Cls was not upgraded. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="CC4C7EC0-C903-48FC-ACCC-81861D12DA4A"'
									frmMain.picCPic(ObjectX.Pic).Cls()
									'UPGRADE_ISSUE: PictureBox property picSFXPic.hdc was not upgraded. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="CC4C7EC0-C903-48FC-ACCC-81861D12DA4A"'
									'UPGRADE_WARNING: Couldn't resolve default property of object ObjectX.Pic. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
									'UPGRADE_ISSUE: PictureBox property picCPic.hdc was not upgraded. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="CC4C7EC0-C903-48FC-ACCC-81861D12DA4A"'
									rc = BitBlt(frmMain.picCPic(ObjectX.Pic).hdc, 0, 0, Width, Height, frmMain.picSFXPic(c).hdc, X * Width, 0, SRCCOPY)
									'UPGRADE_ISSUE: PictureBox property picSFXPic.hdc was not upgraded. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="CC4C7EC0-C903-48FC-ACCC-81861D12DA4A"'
									'UPGRADE_WARNING: Couldn't resolve default property of object ObjectX.Pic. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
									'UPGRADE_ISSUE: PictureBox property picCPic.hdc was not upgraded. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="CC4C7EC0-C903-48FC-ACCC-81861D12DA4A"'
									rc = BitBlt(frmMain.picCPic(ObjectX.Pic).hdc, Width, 0, Width, Height, frmMain.picSFXPic(c).hdc, (VB6.PixelsToTwipsX(frmMain.picSFXPic(c).Width) / 2) + (Cells - X - 1) * Width, 0, SRCCOPY)
									' Paint Mask Normal and Flip
									'UPGRADE_ISSUE: PictureBox property picSFXPic.hdc was not upgraded. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="CC4C7EC0-C903-48FC-ACCC-81861D12DA4A"'
									'UPGRADE_WARNING: Couldn't resolve default property of object ObjectX.Pic. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
									'UPGRADE_ISSUE: PictureBox property picCPic.hdc was not upgraded. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="CC4C7EC0-C903-48FC-ACCC-81861D12DA4A"'
									rc = BitBlt(frmMain.picCPic(ObjectX.Pic).hdc, 0, Height, Width, Height, frmMain.picSFXPic(c).hdc, X * Width, Height, SRCCOPY)
									'UPGRADE_ISSUE: PictureBox property picSFXPic.hdc was not upgraded. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="CC4C7EC0-C903-48FC-ACCC-81861D12DA4A"'
									'UPGRADE_WARNING: Couldn't resolve default property of object ObjectX.Pic. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
									'UPGRADE_ISSUE: PictureBox property picCPic.hdc was not upgraded. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="CC4C7EC0-C903-48FC-ACCC-81861D12DA4A"'
									rc = BitBlt(frmMain.picCPic(ObjectX.Pic).hdc, Width, Height, Width, Height, frmMain.picSFXPic(c).hdc, (VB6.PixelsToTwipsX(frmMain.picSFXPic(c).Width)) / 2 + (Cells - X - 1) * Width, Height, SRCCOPY)
									' Draw Normal and Flip (Yellow)
									'UPGRADE_ISSUE: PictureBox property picSFXPic.hdc was not upgraded. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="CC4C7EC0-C903-48FC-ACCC-81861D12DA4A"'
									'UPGRADE_WARNING: Couldn't resolve default property of object ObjectX.Pic. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
									'UPGRADE_ISSUE: PictureBox property picCPic.hdc was not upgraded. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="CC4C7EC0-C903-48FC-ACCC-81861D12DA4A"'
									rc = BitBlt(frmMain.picCPic(ObjectX.Pic).hdc, 0, Height * 2, Width, Height, frmMain.picSFXPic(c).hdc, X * Width, 0, SRCCOPY)
									'UPGRADE_ISSUE: PictureBox property picSFXPic.hdc was not upgraded. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="CC4C7EC0-C903-48FC-ACCC-81861D12DA4A"'
									'UPGRADE_WARNING: Couldn't resolve default property of object ObjectX.Pic. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
									'UPGRADE_ISSUE: PictureBox property picCPic.hdc was not upgraded. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="CC4C7EC0-C903-48FC-ACCC-81861D12DA4A"'
									rc = BitBlt(frmMain.picCPic(ObjectX.Pic).hdc, Width, Height * 2, Width, Height, frmMain.picSFXPic(c).hdc, (VB6.PixelsToTwipsX(frmMain.picSFXPic(c).Width) / 2) + (Cells - X - 1) * Width, 0, SRCCOPY)
									' Draw Normal and Flip (Red)
									'UPGRADE_ISSUE: PictureBox property picSFXPic.hdc was not upgraded. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="CC4C7EC0-C903-48FC-ACCC-81861D12DA4A"'
									'UPGRADE_WARNING: Couldn't resolve default property of object ObjectX.Pic. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
									'UPGRADE_ISSUE: PictureBox property picCPic.hdc was not upgraded. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="CC4C7EC0-C903-48FC-ACCC-81861D12DA4A"'
									rc = BitBlt(frmMain.picCPic(ObjectX.Pic).hdc, 0, Height * 3, Width, Height, frmMain.picSFXPic(c).hdc, X * Width, 0, SRCCOPY)
									'UPGRADE_ISSUE: PictureBox property picSFXPic.hdc was not upgraded. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="CC4C7EC0-C903-48FC-ACCC-81861D12DA4A"'
									'UPGRADE_WARNING: Couldn't resolve default property of object ObjectX.Pic. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
									'UPGRADE_ISSUE: PictureBox property picCPic.hdc was not upgraded. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="CC4C7EC0-C903-48FC-ACCC-81861D12DA4A"'
									rc = BitBlt(frmMain.picCPic(ObjectX.Pic).hdc, Width, Height * 3, Width, Height, frmMain.picSFXPic(c).hdc, (VB6.PixelsToTwipsX(frmMain.picSFXPic(c).Width) / 2) + (Cells - X - 1) * Width, 0, SRCCOPY)
									'UPGRADE_WARNING: Couldn't resolve default property of object ObjectX.Pic. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
									frmMain.picCPic(ObjectX.Pic).Refresh()
									frmMain.CombatDraw()
									Do Until VB.Timer() - t1 > 0.05
									Loop 
									X = X + 1
								Loop 
								'UPGRADE_WARNING: Couldn't resolve default property of object ObjectX. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
								frmMain.LoadCreaturePicForce(ObjectX)
								'UPGRADE_WARNING: Couldn't resolve default property of object ObjectX. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
								frmMain.LoadCreaturePic(ObjectX)
							End If
						Case 73 ' MapAnimation 'Text' Frames B2 Level B5 At (B0.B1, B3.B4)
							' Use LoadSFXPic to load picture
							' Alter PlotTile so that it uses an SFX picture instead of copy
							' from picTPic. Can't use negatives. Need somewhat to designate
							' that this tile is running an animation.
							' Then run DrawMapRegion (this will be *very* slow)
						Case 74 ' MapRefresh
							frmMain.DrawMapAll()
					End Select
				End If
				.StatementNow = .StatementNow + 1
				If .StatementNow > .Trigger.Statements.Count() Then
					.StopExit = True
				End If
				If .Ticks > 5000 Then
					.StopExit = True
					.Abort = True
				End If
			Loop 
			' Remove charge from Trigger if not 0, not a Skill Trigger and not a Timed Trigger
			If .Trigger.Turns > 0 And .Trigger.IsSkill = False And .Trigger.IsTimed = False Then
				.Trigger.Turns = .Trigger.Turns - 1
				' If out of turns then delete trigger
				If .Trigger.Turns < 1 Then
					'UPGRADE_WARNING: Couldn't resolve default property of object FrameX.ParentX.RemoveTrigger. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
					.ParentX.RemoveTrigger("T" & .Trigger.Index)
				End If
			End If
		End With
		' Pop stack
		PgmStack.Remove("T" & FrameX.Index)
		' Close Dialog if open
		If frmMain.picConvo.Visible = True Then
			DialogHide()
		End If
		' Close CutScene if open
		If RefreshMap = True Then
			frmMain.DrawMapAll()
		End If
		' Abort now
		If FrameX.Abort = True Then
			FireTrigger = False
			Exit Function
		End If
		FireTrigger = (FrameX.Fail = 0)
		DebugShow()
	End Function
	
	Public Function ReplaceText(ByRef InString As String) As Object
		Dim EndAt, StartAt, c As Short
		Dim Var, Context, NewText As String
		StartAt = InStr(InString, "[")
		Do Until StartAt = 0
			EndAt = InStr(StartAt, InString, "]")
			If EndAt > 0 Then
				Context = Mid(InString, StartAt + 1, EndAt - StartAt - 1)
				c = InStr(Context, ".")
				If c > 0 Then
					Var = Mid(Context, c + 1)
					Context = Left(Context, c - 1)
				Else
					Var = "Name"
				End If
				'UPGRADE_WARNING: Couldn't resolve default property of object ConvertVarContext(). Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
				NewText = ConvertVarContext(Context, Var)
				InString = Left(InString, StartAt - 1) & NewText & Mid(InString, EndAt + 1)
				StartAt = InStr(InString, "[")
			Else
				StartAt = 0
			End If
		Loop 
		' Substitute the \ notation for carriage return (\n) and (\)
		StartAt = InStr(InString, "\n")
		Do Until StartAt = 0
			InString = Left(InString, StartAt - 1) & Chr(13) & Chr(10) & Mid(InString, StartAt + 2)
			StartAt = InStr(InString, "\n")
		Loop 
		StartAt = InStr(InString, "\")
		Do Until StartAt = 0
			InString = Left(InString, StartAt - 1) & Mid(InString, StartAt + 1)
			StartAt = InStr(InString, "\")
		Loop 
		'UPGRADE_WARNING: Couldn't resolve default property of object ReplaceText. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
		ReplaceText = InString
	End Function
	
	Public Function ConvertVarContext(ByRef InContext As String, ByRef InVar As String) As Object
		Dim c, i As Short
		Dim Text As String
		'UPGRADE_WARNING: Couldn't resolve default property of object ConvertVarContext. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
		ConvertVarContext = ""
		For c = 0 To 32
			If InContext = ContextToText(c) Then
				For i = 0 To 255
					If InVar = VarToText(c, i) Then
						'UPGRADE_WARNING: Couldn't resolve default property of object GetVarContext(). Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
						ConvertVarContext = GetVarContext(c, i)
						Exit For
					End If
				Next i
				Exit For
			End If
		Next c
	End Function
	
	Private Function CompareVars(ByRef InContext1 As Short, ByRef InVar1 As Short, ByRef Op As Short, ByRef InContext2 As Short, ByRef InVar2 As Short, Optional ByRef InText As Object = Nothing) As Object
		Dim aa, bb As Object
		' Function which fetches and compares two BD vars. Returns 0 if fails, 1 if succeed.
		'UPGRADE_WARNING: Couldn't resolve default property of object GetVarContext(). Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
		'UPGRADE_WARNING: Couldn't resolve default property of object aa. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
		aa = GetVarContext(InContext1, InVar1)
		'UPGRADE_NOTE: IsMissing() was changed to IsNothing(). Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="8AE1CB93-37AB-439A-A4FF-BE3B6760BB23"'
		If IsNothing(InText) Then
			'UPGRADE_WARNING: Couldn't resolve default property of object GetVarContext(). Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
			'UPGRADE_WARNING: Couldn't resolve default property of object bb. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
			bb = GetVarContext(InContext2, InVar2)
		Else
			'UPGRADE_WARNING: Couldn't resolve default property of object InText. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
			'UPGRADE_WARNING: Couldn't resolve default property of object bb. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
			bb = InText
		End If
		'UPGRADE_WARNING: Couldn't resolve default property of object CompareVars. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
		CompareVars = 0 ' Default to failure
		Select Case Op
			Case 0 ' =
				If IsNumeric(aa) And IsNumeric(bb) Then
					'UPGRADE_WARNING: Couldn't resolve default property of object bb. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
					'UPGRADE_WARNING: Couldn't resolve default property of object aa. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
					If aa = bb Then
						'UPGRADE_WARNING: Couldn't resolve default property of object CompareVars. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
						CompareVars = 1
					End If
				Else
					'UPGRADE_WARNING: Couldn't resolve default property of object bb. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
					'UPGRADE_WARNING: Couldn't resolve default property of object aa. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
					If UCase(VB6.Format(aa)) = UCase(VB6.Format(bb)) Then
						'UPGRADE_WARNING: Couldn't resolve default property of object CompareVars. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
						CompareVars = 1
					End If
				End If
			Case 1 ' +
				If IsNumeric(aa) And IsNumeric(bb) Then
					'UPGRADE_WARNING: Couldn't resolve default property of object bb. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
					'UPGRADE_WARNING: Couldn't resolve default property of object aa. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
					'UPGRADE_WARNING: Couldn't resolve default property of object CompareVars. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
					CompareVars = aa + bb
				Else
					'UPGRADE_WARNING: Couldn't resolve default property of object bb. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
					'UPGRADE_WARNING: Couldn't resolve default property of object aa. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
					'UPGRADE_WARNING: Couldn't resolve default property of object CompareVars. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
					CompareVars = VB6.Format(aa) & " " & VB6.Format(bb)
				End If
			Case 2 ' -
				'UPGRADE_WARNING: Couldn't resolve default property of object bb. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
				'UPGRADE_WARNING: Couldn't resolve default property of object aa. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
				'UPGRADE_WARNING: Couldn't resolve default property of object CompareVars. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
				CompareVars = aa - bb
			Case 3 ' *
				'UPGRADE_WARNING: Couldn't resolve default property of object bb. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
				'UPGRADE_WARNING: Couldn't resolve default property of object aa. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
				'UPGRADE_WARNING: Couldn't resolve default property of object CompareVars. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
				CompareVars = aa * bb
			Case 4 ' /
				'UPGRADE_WARNING: Couldn't resolve default property of object bb. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
				If bb <> 0 Then
					'UPGRADE_WARNING: Couldn't resolve default property of object bb. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
					'UPGRADE_WARNING: Couldn't resolve default property of object aa. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
					'UPGRADE_WARNING: Couldn't resolve default property of object CompareVars. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
					CompareVars = aa / bb
				End If
			Case 5 ' >
				'UPGRADE_WARNING: Couldn't resolve default property of object bb. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
				'UPGRADE_WARNING: Couldn't resolve default property of object aa. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
				If aa > bb Then
					'UPGRADE_WARNING: Couldn't resolve default property of object CompareVars. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
					CompareVars = 1
				End If
			Case 6 ' <
				'UPGRADE_WARNING: Couldn't resolve default property of object bb. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
				'UPGRADE_WARNING: Couldn't resolve default property of object aa. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
				If aa < bb Then
					'UPGRADE_WARNING: Couldn't resolve default property of object CompareVars. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
					CompareVars = 1
				End If
			Case 7 ' >=
				'UPGRADE_WARNING: Couldn't resolve default property of object bb. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
				'UPGRADE_WARNING: Couldn't resolve default property of object aa. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
				If aa >= bb Then
					'UPGRADE_WARNING: Couldn't resolve default property of object CompareVars. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
					CompareVars = 1
				End If
			Case 8 ' <=
				'UPGRADE_WARNING: Couldn't resolve default property of object bb. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
				'UPGRADE_WARNING: Couldn't resolve default property of object aa. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
				If aa <= bb Then
					'UPGRADE_WARNING: Couldn't resolve default property of object CompareVars. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
					CompareVars = 1
				End If
			Case 9 ' Or
				If (aa Or bb) > 0 Then
					'UPGRADE_WARNING: Couldn't resolve default property of object CompareVars. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
					CompareVars = 1
				End If
			Case 10 ' And
				If (aa And bb) > 0 Then
					'UPGRADE_WARNING: Couldn't resolve default property of object CompareVars. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
					CompareVars = 1
				End If
			Case 11 ' Xor
				If (aa Xor bb) > 0 Then
					'UPGRADE_WARNING: Couldn't resolve default property of object CompareVars. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
					CompareVars = 1
				End If
			Case 12 ' Like
				If Len(aa) > 0 And Len(bb) > 0 Then
					'UPGRADE_WARNING: Couldn't resolve default property of object aa. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
					'UPGRADE_WARNING: Couldn't resolve default property of object bb. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
					If InStr(UCase(aa), UCase(bb)) > 0 Or InStr(UCase(bb), UCase(aa)) > 0 Then
						'UPGRADE_WARNING: Couldn't resolve default property of object CompareVars. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
						CompareVars = 1
					End If
				End If
			Case 13 ' <> (Not Equal)
				'UPGRADE_WARNING: Couldn't resolve default property of object bb. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
				'UPGRADE_WARNING: Couldn't resolve default property of object aa. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
				If aa <> bb Then
					'UPGRADE_WARNING: Couldn't resolve default property of object CompareVars. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
					CompareVars = 1
				End If
		End Select
	End Function
	
	'UPGRADE_NOTE: Tome was upgraded to Tome_Renamed. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="A9E4979A-37FA-4718-9994-97DD76ED70A7"'
	'UPGRADE_NOTE: Map was upgraded to Map_Renamed. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="A9E4979A-37FA-4718-9994-97DD76ED70A7"'
	Public Function GetVarContext(ByRef AsContext As Short, ByRef AsVar As Short) As Object
		Dim Map_Renamed As Object
		Dim Tome_Renamed As Object
		Dim ObjectX As Object
		Dim TriggerX As Trigger
		Dim Found As Short
		Dim X, NoFail, c, Y As Short
		' Get value
		Select Case AsContext
			Case 0 'Tome
				Select Case AsVar
					Case 0
						' PartyCount
						'UPGRADE_WARNING: Couldn't resolve default property of object Tome.Creatures. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
						'UPGRADE_WARNING: Couldn't resolve default property of object GetVarContext. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
						GetVarContext = Tome.Creatures.Count
					Case 1 ' Comments
						'UPGRADE_WARNING: Couldn't resolve default property of object Tome.Comments. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
						'UPGRADE_WARNING: Couldn't resolve default property of object GetVarContext. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
						GetVarContext = Tome.Comments
					Case 2 'Factoids
						'UPGRADE_WARNING: Couldn't resolve default property of object GetVarContext. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
						GetVarContext = ""
						'UPGRADE_WARNING: Couldn't resolve default property of object Tome.Factoids. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
						For	Each ObjectX In Tome.Factoids
							'UPGRADE_WARNING: Couldn't resolve default property of object ObjectX.Text. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
							'UPGRADE_WARNING: Couldn't resolve default property of object GetVarContext. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
							GetVarContext = GetVarContext & "|" & ObjectX.Text
						Next ObjectX
					Case 3
						' Not Used
					Case 4 'MapX
						'UPGRADE_WARNING: Couldn't resolve default property of object Tome.MapX. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
						'UPGRADE_WARNING: Couldn't resolve default property of object GetVarContext. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
						GetVarContext = Tome.MapX
					Case 5 'MapY
						'UPGRADE_WARNING: Couldn't resolve default property of object Tome.MapY. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
						'UPGRADE_WARNING: Couldn't resolve default property of object GetVarContext. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
						GetVarContext = Tome.MapY
					Case 6 'MoveToX
						'UPGRADE_WARNING: Couldn't resolve default property of object Tome.MoveToX. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
						'UPGRADE_WARNING: Couldn't resolve default property of object GetVarContext. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
						GetVarContext = Tome.MoveToX
					Case 7 'MoveToY
						'UPGRADE_WARNING: Couldn't resolve default property of object Tome.MoveToY. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
						'UPGRADE_WARNING: Couldn't resolve default property of object GetVarContext. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
						GetVarContext = Tome.MoveToY
					Case 8 'Name
						'UPGRADE_WARNING: Couldn't resolve default property of object Tome.Name. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
						'UPGRADE_WARNING: Couldn't resolve default property of object GetVarContext. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
						GetVarContext = Tome.Name
					Case 9 : GetVarContext = Left(WorldNow.Money, InStr(WorldNow.Money, "|") - 1)
						' Not Used
					Case 10 'TimeCycle
						'UPGRADE_WARNING: Couldn't resolve default property of object Tome.Cycle. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
						'UPGRADE_WARNING: Couldn't resolve default property of object GetVarContext. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
						GetVarContext = Tome.Cycle
					Case 11 'TimeMoon
						'UPGRADE_WARNING: Couldn't resolve default property of object Tome.Moon. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
						'UPGRADE_WARNING: Couldn't resolve default property of object GetVarContext. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
						GetVarContext = Tome.Moon
					Case 12 'TimeTurn
						'UPGRADE_WARNING: Couldn't resolve default property of object Tome.Turn. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
						'UPGRADE_WARNING: Couldn't resolve default property of object GetVarContext. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
						GetVarContext = Tome.Turn
					Case 13 'TimeYear
						'UPGRADE_WARNING: Couldn't resolve default property of object Tome.Year. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
						'UPGRADE_WARNING: Couldn't resolve default property of object GetVarContext. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
						GetVarContext = Tome.Year
					Case 14 'RealSeconds
						'UPGRADE_WARNING: Couldn't resolve default property of object Tome.RealSeconds. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
						'UPGRADE_WARNING: Couldn't resolve default property of object GetVarContext. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
						GetVarContext = Tome.RealSeconds
					Case 15 'RealMinutes
						'UPGRADE_WARNING: Couldn't resolve default property of object Tome.RealMinutes. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
						'UPGRADE_WARNING: Couldn't resolve default property of object GetVarContext. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
						GetVarContext = Tome.RealMinutes
					Case 16 'RealHours
						'UPGRADE_WARNING: Couldn't resolve default property of object Tome.RealHours. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
						'UPGRADE_WARNING: Couldn't resolve default property of object GetVarContext. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
						GetVarContext = Tome.RealHours
					Case 17 'RealDays
						'UPGRADE_WARNING: Couldn't resolve default property of object Tome.RealDays. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
						'UPGRADE_WARNING: Couldn't resolve default property of object GetVarContext. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
						GetVarContext = Tome.RealDays
				End Select
			Case 1 'Map
				Select Case AsVar
					Case 0 'Comments
						'UPGRADE_WARNING: Couldn't resolve default property of object Map.Comments. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
						'UPGRADE_WARNING: Couldn't resolve default property of object GetVarContext. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
						GetVarContext = Map.Comments
					Case 1 'DefaultEncounters
						'UPGRADE_WARNING: Couldn't resolve default property of object Map.DefaultEncounters. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
						'UPGRADE_WARNING: Couldn't resolve default property of object GetVarContext. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
						GetVarContext = Map.DefaultEncounters
					Case 2 'DefaultHeight
						'UPGRADE_WARNING: Couldn't resolve default property of object Map.DefaultHeight. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
						'UPGRADE_WARNING: Couldn't resolve default property of object GetVarContext. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
						GetVarContext = Map.DefaultHeight
					Case 3 'DefaultStyle
						'UPGRADE_WARNING: Couldn't resolve default property of object Map.DefaultStyle. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
						'UPGRADE_WARNING: Couldn't resolve default property of object GetVarContext. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
						GetVarContext = Map.DefaultStyle
					Case 4 'DefaultWidth
						'UPGRADE_WARNING: Couldn't resolve default property of object Map.DefaultWidth. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
						'UPGRADE_WARNING: Couldn't resolve default property of object GetVarContext. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
						GetVarContext = Map.DefaultWidth
					Case 5 'Difficulty
					Case 6 'ExperiencePoints
						'UPGRADE_WARNING: Couldn't resolve default property of object Map.ExperiencePoints. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
						'UPGRADE_WARNING: Couldn't resolve default property of object GetVarContext. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
						GetVarContext = Map.ExperiencePoints
					Case 7 'Height
						'UPGRADE_WARNING: Couldn't resolve default property of object Map.Height. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
						'UPGRADE_WARNING: Couldn't resolve default property of object GetVarContext. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
						GetVarContext = Map.Height
					Case 8 'Left
						'UPGRADE_WARNING: Couldn't resolve default property of object Map.Left. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
						'UPGRADE_WARNING: Couldn't resolve default property of object GetVarContext. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
						GetVarContext = Map.Left
					Case 9 'Name
						'UPGRADE_WARNING: Couldn't resolve default property of object Map.Name. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
						'UPGRADE_WARNING: Couldn't resolve default property of object GetVarContext. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
						GetVarContext = Map.Name
					Case 10 'ReGenUponEntry
						'UPGRADE_WARNING: Couldn't resolve default property of object Map.GenerateUponEntry. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
						'UPGRADE_WARNING: Couldn't resolve default property of object GetVarContext. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
						GetVarContext = Map.GenerateUponEntry
					Case 11 'Rune: Blood
						'UPGRADE_WARNING: Couldn't resolve default property of object Map.Runes. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
						'UPGRADE_WARNING: Couldn't resolve default property of object GetVarContext. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
						GetVarContext = Map.Runes(0)
					Case 12 'Rune: Bile
						'UPGRADE_WARNING: Couldn't resolve default property of object Map.Runes. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
						'UPGRADE_WARNING: Couldn't resolve default property of object GetVarContext. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
						GetVarContext = Map.Runes(1)
					Case 13 'Rune: Oil
						'UPGRADE_WARNING: Couldn't resolve default property of object Map.Runes. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
						'UPGRADE_WARNING: Couldn't resolve default property of object GetVarContext. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
						GetVarContext = Map.Runes(2)
					Case 14 'Rune: Nectar
						'UPGRADE_WARNING: Couldn't resolve default property of object Map.Runes. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
						'UPGRADE_WARNING: Couldn't resolve default property of object GetVarContext. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
						GetVarContext = Map.Runes(3)
					Case 15 'Rune: Fire
						'UPGRADE_WARNING: Couldn't resolve default property of object Map.Runes. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
						'UPGRADE_WARNING: Couldn't resolve default property of object GetVarContext. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
						GetVarContext = Map.Runes(4)
					Case 16 'Rune: Earth
						'UPGRADE_WARNING: Couldn't resolve default property of object Map.Runes. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
						'UPGRADE_WARNING: Couldn't resolve default property of object GetVarContext. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
						GetVarContext = Map.Runes(5)
					Case 17 'Rune: Water
						'UPGRADE_WARNING: Couldn't resolve default property of object Map.Runes. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
						'UPGRADE_WARNING: Couldn't resolve default property of object GetVarContext. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
						GetVarContext = Map.Runes(6)
					Case 18 'Rune: Air
						'UPGRADE_WARNING: Couldn't resolve default property of object Map.Runes. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
						'UPGRADE_WARNING: Couldn't resolve default property of object GetVarContext. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
						GetVarContext = Map.Runes(7)
					Case 19 'Rune: Time
						'UPGRADE_WARNING: Couldn't resolve default property of object Map.Runes. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
						'UPGRADE_WARNING: Couldn't resolve default property of object GetVarContext. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
						GetVarContext = Map.Runes(8)
					Case 20 'Rune: Moon
						'UPGRADE_WARNING: Couldn't resolve default property of object Map.Runes. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
						'UPGRADE_WARNING: Couldn't resolve default property of object GetVarContext. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
						GetVarContext = Map.Runes(9)
					Case 21 'Rune: Sun
						'UPGRADE_WARNING: Couldn't resolve default property of object Map.Runes. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
						'UPGRADE_WARNING: Couldn't resolve default property of object GetVarContext. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
						GetVarContext = Map.Runes(10)
					Case 22 'Rune: Space
						'UPGRADE_WARNING: Couldn't resolve default property of object Map.Runes. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
						'UPGRADE_WARNING: Couldn't resolve default property of object GetVarContext. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
						GetVarContext = Map.Runes(11)
					Case 23 'Rune: Insect
						'UPGRADE_WARNING: Couldn't resolve default property of object Map.Runes. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
						'UPGRADE_WARNING: Couldn't resolve default property of object GetVarContext. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
						GetVarContext = Map.Runes(12)
					Case 24 'Rune: Man
						'UPGRADE_WARNING: Couldn't resolve default property of object Map.Runes. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
						'UPGRADE_WARNING: Couldn't resolve default property of object GetVarContext. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
						GetVarContext = Map.Runes(13)
					Case 25 'Rune: Animal
						'UPGRADE_WARNING: Couldn't resolve default property of object Map.Runes. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
						'UPGRADE_WARNING: Couldn't resolve default property of object GetVarContext. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
						GetVarContext = Map.Runes(14)
					Case 26 'Rune: Fish
						'UPGRADE_WARNING: Couldn't resolve default property of object Map.Runes. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
						'UPGRADE_WARNING: Couldn't resolve default property of object GetVarContext. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
						GetVarContext = Map.Runes(15)
					Case 27 'Rune: Twilight
						'UPGRADE_WARNING: Couldn't resolve default property of object Map.Runes. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
						'UPGRADE_WARNING: Couldn't resolve default property of object GetVarContext. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
						GetVarContext = Map.Runes(16)
					Case 28 'Rune: Abyss
						'UPGRADE_WARNING: Couldn't resolve default property of object Map.Runes. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
						'UPGRADE_WARNING: Couldn't resolve default property of object GetVarContext. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
						GetVarContext = Map.Runes(17)
					Case 29 'Rune: Eternium
						'UPGRADE_WARNING: Couldn't resolve default property of object Map.Runes. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
						'UPGRADE_WARNING: Couldn't resolve default property of object GetVarContext. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
						GetVarContext = Map.Runes(18)
					Case 30 'Rune: Dreams
						'UPGRADE_WARNING: Couldn't resolve default property of object Map.Runes. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
						'UPGRADE_WARNING: Couldn't resolve default property of object GetVarContext. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
						GetVarContext = Map.Runes(19)
					Case 31 'Top
						'UPGRADE_WARNING: Couldn't resolve default property of object Map.Top. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
						'UPGRADE_WARNING: Couldn't resolve default property of object GetVarContext. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
						GetVarContext = Map.Top
					Case 32 'Width
						'UPGRADE_WARNING: Couldn't resolve default property of object Map.Width. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
						'UPGRADE_WARNING: Couldn't resolve default property of object GetVarContext. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
						GetVarContext = Map.Width
					Case 33 ' IsOutside
						'UPGRADE_WARNING: Couldn't resolve default property of object Map.IsOutside. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
						GetVarContext = System.Math.Abs(Map.IsOutside)
					Case 34 ' MapStyle
						'UPGRADE_WARNING: Couldn't resolve default property of object Map.MapStyle. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
						Select Case Map.MapStyle
							Case 0 ' Town
								'UPGRADE_WARNING: Couldn't resolve default property of object GetVarContext. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
								GetVarContext = "Town"
							Case 1 ' Wilderness
								'UPGRADE_WARNING: Couldn't resolve default property of object GetVarContext. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
								GetVarContext = "Wilderness"
							Case 2 ' Building
								'UPGRADE_WARNING: Couldn't resolve default property of object GetVarContext. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
								GetVarContext = "Building"
							Case 3 ' Dungeon
								'UPGRADE_WARNING: Couldn't resolve default property of object GetVarContext. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
								GetVarContext = "Dungeon"
						End Select
				End Select
			Case 2, 3, 4, 5, 6 'TileNow, TileA, TileB, TileC, TileTarget
				X = FrameNow.TileX(AsContext - 2) : Y = FrameNow.TileY(AsContext - 2)
				Select Case AsVar
					Case 0 ' CanSeeDown
						'UPGRADE_WARNING: Couldn't resolve default property of object Map.See. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
						GetVarContext = System.Math.Abs(Map.See(X, Y, 3))
					Case 1 ' CanSeeLeft
						'UPGRADE_WARNING: Couldn't resolve default property of object Map.See. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
						GetVarContext = System.Math.Abs(Map.See(X, Y, 2))
					Case 2 ' CanSeeRight
						'UPGRADE_WARNING: Couldn't resolve default property of object Map.See. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
						GetVarContext = System.Math.Abs(Map.See(X, Y, 1))
					Case 3 ' CanSeeUp
						'UPGRADE_WARNING: Couldn't resolve default property of object Map.See. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
						GetVarContext = System.Math.Abs(Map.See(X, Y, 0))
					Case 4 ' ChanceToOpen
						Found = False
						'UPGRADE_WARNING: Couldn't resolve default property of object GetVarContext. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
						GetVarContext = 0
						'UPGRADE_WARNING: Couldn't resolve default property of object Map.TopTile. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
						If Map.TopTile(X, Y) > 0 Then
							'UPGRADE_WARNING: Couldn't resolve default property of object Map.TopTile. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
							'UPGRADE_WARNING: Couldn't resolve default property of object Map.Tiles. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
							If Map.Tiles("L" & Map.TopTile(X, Y)).SwapTile > 0 Then
								'UPGRADE_WARNING: Couldn't resolve default property of object Map.TopTile. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
								'UPGRADE_WARNING: Couldn't resolve default property of object Map.Tiles. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
								'UPGRADE_WARNING: Couldn't resolve default property of object GetVarContext. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
								GetVarContext = Map.Tiles("L" & Map.TopTile(X, Y)).Chance
								Found = True
							End If
						End If
						'UPGRADE_WARNING: Couldn't resolve default property of object Map.MiddleTile. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
						If Map.MiddleTile(X, Y) > 0 And Found = False Then
							'UPGRADE_WARNING: Couldn't resolve default property of object Map.MiddleTile. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
							'UPGRADE_WARNING: Couldn't resolve default property of object Map.Tiles. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
							If Map.Tiles("L" & Map.MiddleTile(X, Y)).SwapTile > 0 Then
								'UPGRADE_WARNING: Couldn't resolve default property of object Map.MiddleTile. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
								'UPGRADE_WARNING: Couldn't resolve default property of object Map.Tiles. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
								'UPGRADE_WARNING: Couldn't resolve default property of object GetVarContext. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
								GetVarContext = Map.Tiles("L" & Map.MiddleTile(X, Y)).Chance
								Found = True
							End If
						End If
						'UPGRADE_WARNING: Couldn't resolve default property of object Map.BottomTile. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
						If Map.BottomTile(X, Y) > 0 And Found = False Then
							'UPGRADE_WARNING: Couldn't resolve default property of object Map.BottomTile. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
							'UPGRADE_WARNING: Couldn't resolve default property of object Map.Tiles. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
							If Map.Tiles("L" & Map.BottomTile(X, Y)).SwapTile > 0 Then
								'UPGRADE_WARNING: Couldn't resolve default property of object Map.BottomTile. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
								'UPGRADE_WARNING: Couldn't resolve default property of object Map.Tiles. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
								'UPGRADE_WARNING: Couldn't resolve default property of object GetVarContext. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
								GetVarContext = Map.Tiles("L" & Map.BottomTile(X, Y)).Chance
								Found = True
							End If
						End If
					Case 5 ' Not Used
					Case 6 ' IsBlockedDown
						'UPGRADE_WARNING: Couldn't resolve default property of object Map.Blocked. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
						GetVarContext = System.Math.Abs(Map.Blocked(X, Y, 3))
					Case 7 ' IsBlockedLeft
						'UPGRADE_WARNING: Couldn't resolve default property of object Map.Blocked. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
						GetVarContext = System.Math.Abs(Map.Blocked(X, Y, 2))
					Case 8 ' IsBlockedRight
						'UPGRADE_WARNING: Couldn't resolve default property of object Map.Blocked. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
						GetVarContext = System.Math.Abs(Map.Blocked(X, Y, 1))
					Case 9 ' IsBlockedUp
						'UPGRADE_WARNING: Couldn't resolve default property of object Map.Blocked. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
						GetVarContext = System.Math.Abs(Map.Blocked(X, Y, 0))
					Case 10 ' Key
						Found = False
						'UPGRADE_WARNING: Couldn't resolve default property of object GetVarContext. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
						GetVarContext = 0
						'UPGRADE_WARNING: Couldn't resolve default property of object Map.TopTile. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
						If Map.TopTile(X, Y) > 0 Then
							'UPGRADE_WARNING: Couldn't resolve default property of object Map.TopTile. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
							'UPGRADE_WARNING: Couldn't resolve default property of object Map.Tiles. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
							If Map.Tiles("L" & Map.TopTile(X, Y)).SwapTile > 0 Then
								'UPGRADE_WARNING: Couldn't resolve default property of object Map.TopTile. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
								'UPGRADE_WARNING: Couldn't resolve default property of object Map.Tiles. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
								'UPGRADE_WARNING: Couldn't resolve default property of object GetVarContext. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
								GetVarContext = Map.Tiles("L" & Map.TopTile(X, Y)).KeyBits
								Found = True
							End If
						End If
						'UPGRADE_WARNING: Couldn't resolve default property of object Map.MiddleTile. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
						If Map.MiddleTile(X, Y) > 0 And Found = False Then
							'UPGRADE_WARNING: Couldn't resolve default property of object Map.MiddleTile. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
							'UPGRADE_WARNING: Couldn't resolve default property of object Map.Tiles. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
							If Map.Tiles("L" & Map.MiddleTile(X, Y)).SwapTile > 0 Then
								'UPGRADE_WARNING: Couldn't resolve default property of object Map.MiddleTile. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
								'UPGRADE_WARNING: Couldn't resolve default property of object Map.Tiles. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
								'UPGRADE_WARNING: Couldn't resolve default property of object GetVarContext. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
								GetVarContext = Map.Tiles("L" & Map.MiddleTile(X, Y)).KeyBits
								Found = True
							End If
						End If
						'UPGRADE_WARNING: Couldn't resolve default property of object Map.BottomTile. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
						If Map.BottomTile(X, Y) > 0 And Found = False Then
							'UPGRADE_WARNING: Couldn't resolve default property of object Map.BottomTile. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
							'UPGRADE_WARNING: Couldn't resolve default property of object Map.Tiles. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
							If Map.Tiles("L" & Map.BottomTile(X, Y)).SwapTile > 0 Then
								'UPGRADE_WARNING: Couldn't resolve default property of object Map.BottomTile. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
								'UPGRADE_WARNING: Couldn't resolve default property of object Map.Tiles. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
								'UPGRADE_WARNING: Couldn't resolve default property of object GetVarContext. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
								GetVarContext = Map.Tiles("L" & Map.BottomTile(X, Y)).KeyBits
								Found = True
							End If
						End If
					Case 11 To 15 ' Not Used
					Case 16 ' Name
						'UPGRADE_WARNING: Couldn't resolve default property of object GetVarContext. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
						GetVarContext = ""
						'UPGRADE_WARNING: Couldn't resolve default property of object Map.BottomTile. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
						If Map.BottomTile(X, Y) > 0 Then
							'UPGRADE_WARNING: Couldn't resolve default property of object Map.BottomTile. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
							'UPGRADE_WARNING: Couldn't resolve default property of object Map.Tiles. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
							'UPGRADE_WARNING: Couldn't resolve default property of object GetVarContext. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
							GetVarContext = GetVarContext & "|" & Map.Tiles("L" & Map.BottomTile(X, Y)).Name
						End If
						'UPGRADE_WARNING: Couldn't resolve default property of object Map.MiddleTile. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
						If Map.MiddleTile(X, Y) > 0 Then
							'UPGRADE_WARNING: Couldn't resolve default property of object Map.MiddleTile. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
							'UPGRADE_WARNING: Couldn't resolve default property of object Map.Tiles. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
							'UPGRADE_WARNING: Couldn't resolve default property of object GetVarContext. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
							GetVarContext = GetVarContext & "|" & Map.Tiles("L" & Map.MiddleTile(X, Y)).Name
						End If
						'UPGRADE_WARNING: Couldn't resolve default property of object Map.TopTile. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
						If Map.TopTile(X, Y) > 0 Then
							'UPGRADE_WARNING: Couldn't resolve default property of object Map.TopTile. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
							'UPGRADE_WARNING: Couldn't resolve default property of object Map.Tiles. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
							'UPGRADE_WARNING: Couldn't resolve default property of object GetVarContext. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
							GetVarContext = GetVarContext & "|" & Map.Tiles("L" & Map.TopTile(X, Y)).Name
						End If
					Case 17 ' Style
						'UPGRADE_WARNING: Couldn't resolve default property of object GetVarContext. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
						GetVarContext = ""
						'UPGRADE_WARNING: Couldn't resolve default property of object Map.BottomTile. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
						If Map.BottomTile(X, Y) > 0 Then
							'UPGRADE_WARNING: Couldn't resolve default property of object Map.BottomTile. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
							'UPGRADE_WARNING: Couldn't resolve default property of object Map.Tiles. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
							'UPGRADE_WARNING: Couldn't resolve default property of object GetVarContext. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
							GetVarContext = GetVarContext & "|" & modBD.SetUpTileStyle(Map.Tiles("L" & Map.BottomTile(X, Y)).Style)
						End If
						'UPGRADE_WARNING: Couldn't resolve default property of object Map.MiddleTile. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
						If Map.MiddleTile(X, Y) > 0 Then
							'UPGRADE_WARNING: Couldn't resolve default property of object Map.MiddleTile. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
							'UPGRADE_WARNING: Couldn't resolve default property of object Map.Tiles. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
							'UPGRADE_WARNING: Couldn't resolve default property of object GetVarContext. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
							GetVarContext = GetVarContext & "|" & modBD.SetUpTileStyle(Map.Tiles("L" & Map.MiddleTile(X, Y)).Style)
						End If
						'UPGRADE_WARNING: Couldn't resolve default property of object Map.TopTile. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
						If Map.TopTile(X, Y) > 0 Then
							'UPGRADE_WARNING: Couldn't resolve default property of object Map.TopTile. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
							'UPGRADE_WARNING: Couldn't resolve default property of object Map.Tiles. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
							'UPGRADE_WARNING: Couldn't resolve default property of object GetVarContext. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
							GetVarContext = GetVarContext & "|" & modBD.SetUpTileStyle(Map.Tiles("L" & Map.TopTile(X, Y)).Style)
						End If
					Case 18 ' DoorName
						Found = False
						'UPGRADE_WARNING: Couldn't resolve default property of object GetVarContext. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
						GetVarContext = ""
						'UPGRADE_WARNING: Couldn't resolve default property of object Map.TopTile. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
						If Map.TopTile(X, Y) > 0 Then
							'UPGRADE_WARNING: Couldn't resolve default property of object Map.TopTile. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
							'UPGRADE_WARNING: Couldn't resolve default property of object Map.Tiles. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
							If Map.Tiles("L" & Map.TopTile(X, Y)).SwapTile > 0 Then
								'UPGRADE_WARNING: Couldn't resolve default property of object Map.TopTile. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
								'UPGRADE_WARNING: Couldn't resolve default property of object Map.Tiles. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
								'UPGRADE_WARNING: Couldn't resolve default property of object GetVarContext. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
								GetVarContext = Map.Tiles("L" & Map.Tiles("L" & Map.TopTile(X, Y)).SwapTile)
								Found = True
							End If
						End If
						'UPGRADE_WARNING: Couldn't resolve default property of object Map.MiddleTile. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
						If Map.MiddleTile(X, Y) > 0 Then
							'UPGRADE_WARNING: Couldn't resolve default property of object Map.MiddleTile. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
							'UPGRADE_WARNING: Couldn't resolve default property of object Map.Tiles. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
							If Map.Tiles("L" & Map.MiddleTile(X, Y)).SwapTile > 0 Then
								'UPGRADE_WARNING: Couldn't resolve default property of object Map.MiddleTile. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
								'UPGRADE_WARNING: Couldn't resolve default property of object Map.Tiles. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
								'UPGRADE_WARNING: Couldn't resolve default property of object GetVarContext. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
								GetVarContext = Map.Tiles("L" & Map.Tiles("L" & Map.MiddleTile(X, Y)).SwapTile)
								Found = True
							End If
						End If
						'UPGRADE_WARNING: Couldn't resolve default property of object Map.BottomTile. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
						If Map.BottomTile(X, Y) > 0 Then
							'UPGRADE_WARNING: Couldn't resolve default property of object Map.BottomTile. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
							'UPGRADE_WARNING: Couldn't resolve default property of object Map.Tiles. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
							If Map.Tiles("L" & Map.BottomTile(X, Y)).SwapTile > 0 Then
								'UPGRADE_WARNING: Couldn't resolve default property of object Map.BottomTile. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
								'UPGRADE_WARNING: Couldn't resolve default property of object Map.Tiles. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
								'UPGRADE_WARNING: Couldn't resolve default property of object GetVarContext. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
								GetVarContext = Map.Tiles("L" & Map.Tiles("L" & Map.BottomTile(X, Y)).SwapTile)
								Found = True
							End If
						End If
						If Not Found Then
							'UPGRADE_WARNING: Couldn't resolve default property of object GetVarContext. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
							GetVarContext = "<None>"
						End If
					Case 19 ' MovementCost
						'UPGRADE_WARNING: Couldn't resolve default property of object Map.MovementCost. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
						'UPGRADE_WARNING: Couldn't resolve default property of object GetVarContext. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
						GetVarContext = Map.MovementCost(X, Y)
					Case 20 ' TerrainType
						'UPGRADE_WARNING: Couldn't resolve default property of object GetVarContext. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
						GetVarContext = ""
						'UPGRADE_WARNING: Couldn't resolve default property of object Map.BottomTile. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
						If Map.BottomTile(X, Y) > 0 Then
							'UPGRADE_WARNING: Couldn't resolve default property of object Map.BottomTile. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
							'UPGRADE_WARNING: Couldn't resolve default property of object Map.Tiles. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
							'UPGRADE_WARNING: Couldn't resolve default property of object GetVarContext. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
							GetVarContext = GetVarContext & "|" & modBD.SetUpTileTerrain(Map.Tiles("L" & Map.BottomTile(X, Y)).TerrainType)
						End If
						'UPGRADE_WARNING: Couldn't resolve default property of object Map.MiddleTile. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
						If Map.MiddleTile(X, Y) > 0 Then
							'UPGRADE_WARNING: Couldn't resolve default property of object Map.MiddleTile. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
							'UPGRADE_WARNING: Couldn't resolve default property of object Map.Tiles. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
							'UPGRADE_WARNING: Couldn't resolve default property of object GetVarContext. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
							GetVarContext = GetVarContext & "|" & modBD.SetUpTileTerrain(Map.Tiles("L" & Map.MiddleTile(X, Y)).TerrainType)
						End If
						'UPGRADE_WARNING: Couldn't resolve default property of object Map.TopTile. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
						If Map.TopTile(X, Y) > 0 Then
							'UPGRADE_WARNING: Couldn't resolve default property of object Map.TopTile. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
							'UPGRADE_WARNING: Couldn't resolve default property of object Map.Tiles. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
							'UPGRADE_WARNING: Couldn't resolve default property of object GetVarContext. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
							GetVarContext = GetVarContext & "|" & modBD.SetUpTileTerrain(Map.Tiles("L" & Map.TopTile(X, Y)).TerrainType)
						End If
				End Select
			Case 7, 8, 9, 10, 11 'EncounterNow, EncounterA, EncounterB, EncounterC, EncounterTarget
				ObjectX = FrameNow.EncounterX(AsContext - 7)
				Select Case AsVar
					Case 0 ' CanFight
						'UPGRADE_WARNING: Couldn't resolve default property of object ObjectX.CanFight. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
						GetVarContext = System.Math.Abs(ObjectX.CanFight)
					Case 1 ' CanFlee
						'UPGRADE_WARNING: Couldn't resolve default property of object ObjectX.CanFlee. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
						GetVarContext = System.Math.Abs(ObjectX.CanFlee)
					Case 2 ' CanTalk
						'UPGRADE_WARNING: Couldn't resolve default property of object ObjectX.CanTalk. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
						GetVarContext = System.Math.Abs(ObjectX.CanTalk)
					Case 3 ' ChanceToFlee
						'UPGRADE_WARNING: Couldn't resolve default property of object ObjectX.ChanceToFlee. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
						'UPGRADE_WARNING: Couldn't resolve default property of object GetVarContext. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
						GetVarContext = ObjectX.ChanceToFlee
					Case 4 ' Randomize: Classification
						'UPGRADE_WARNING: Couldn't resolve default property of object ObjectX.Classification. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
						'UPGRADE_WARNING: Couldn't resolve default property of object GetVarContext. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
						GetVarContext = modBD.SetUpEncClass(ObjectX.Classification)
					Case 5 ' FirstEntry
						'   [Titi 2.4.7]     GetVarContext = Abs(ObjectX.FirstEntry)  <-- RT 13 here, FirstEntry is a string!
						'UPGRADE_WARNING: Couldn't resolve default property of object ObjectX.FirstEntry. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
						'UPGRADE_WARNING: Couldn't resolve default property of object GetVarContext. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
						GetVarContext = ObjectX.FirstEntry
					Case 6 ' HaveEntered
						'UPGRADE_WARNING: Couldn't resolve default property of object ObjectX.HaveEntered. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
						GetVarContext = System.Math.Abs(ObjectX.HaveEntered)
					Case 8 ' IsDark
						'UPGRADE_WARNING: Couldn't resolve default property of object ObjectX.IsDark. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
						GetVarContext = System.Math.Abs(ObjectX.IsDark)
					Case 9 ' Name
						'UPGRADE_WARNING: Couldn't resolve default property of object ObjectX.Name. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
						'UPGRADE_WARNING: Couldn't resolve default property of object GetVarContext. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
						GetVarContext = ObjectX.Name
					Case 10 ' Randomize: ThemeName
						'UPGRADE_WARNING: Couldn't resolve default property of object ObjectX.ParentTheme. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
						If ObjectX.ParentTheme > 0 Then
							'UPGRADE_WARNING: Couldn't resolve default property of object ObjectX.ParentTheme. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
							'UPGRADE_WARNING: Couldn't resolve default property of object Map.Themes. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
							'UPGRADE_WARNING: Couldn't resolve default property of object GetVarContext. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
							GetVarContext = Map.Themes("R" & ObjectX.ParentTheme)
						Else
							'UPGRADE_WARNING: Couldn't resolve default property of object GetVarContext. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
							GetVarContext = "<None>"
						End If
					Case 11 ' Randomize: AddCreatures
						'UPGRADE_WARNING: Couldn't resolve default property of object ObjectX.ReGenCreatures. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
						GetVarContext = System.Math.Abs(ObjectX.ReGenCreatures)
					Case 12 ' Randomize: AddItems
						'UPGRADE_WARNING: Couldn't resolve default property of object ObjectX.ReGenItems. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
						GetVarContext = System.Math.Abs(ObjectX.ReGenItems)
					Case 13 ' Randomize: AddTriggers
						'UPGRADE_WARNING: Couldn't resolve default property of object ObjectX.ReGenTriggers. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
						GetVarContext = System.Math.Abs(ObjectX.ReGenTriggers)
					Case 14 ' Randomize: Description
						'UPGRADE_WARNING: Couldn't resolve default property of object ObjectX.ReGenNewDescriptions. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
						GetVarContext = System.Math.Abs(ObjectX.ReGenNewDescriptions)
					Case 15 ' Ranzomize: GenerateUponEntry
						'UPGRADE_WARNING: Couldn't resolve default property of object ObjectX.GenerateUponEntry. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
						GetVarContext = System.Math.Abs(ObjectX.GenerateUponEntry)
					Case 16 ' SecondEntry
						'   [Titi 2.4.7]     GetVarContext = Abs(ObjectX.HaveEntered)  <-- should be SecondEntry!
						'UPGRADE_WARNING: Couldn't resolve default property of object ObjectX.SecondEntry. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
						'UPGRADE_WARNING: Couldn't resolve default property of object GetVarContext. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
						GetVarContext = ObjectX.SecondEntry
					Case 17 ' ShowHint
						'UPGRADE_WARNING: Couldn't resolve default property of object ObjectX.UseHint. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
						GetVarContext = System.Math.Abs(ObjectX.UseHint)
					Case 18 ' CreatureCount
						'UPGRADE_WARNING: Couldn't resolve default property of object ObjectX.Creatures. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
						GetVarContext = System.Math.Abs(ObjectX.Creatures.Count)
					Case 19 ' ItemCount
						'UPGRADE_WARNING: Couldn't resolve default property of object ObjectX.Items. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
						GetVarContext = System.Math.Abs(ObjectX.Items.Count)
				End Select
			Case 12, 13, 14, 15, 16 'CreatureNow, CreatureA, CreatureB, CreatureC, CreatureTarget
				ObjectX = FrameNow.CreatureX(AsContext - 12)
				Select Case AsVar
					Case 0 To 7 ' BodyPart1-8
						'UPGRADE_WARNING: Couldn't resolve default property of object ObjectX.BodyType. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
						'UPGRADE_WARNING: Couldn't resolve default property of object GetVarContext. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
						GetVarContext = ObjectX.BodyType(AsVar)
					Case 8 To 15 ' Resistance1-8
						'UPGRADE_WARNING: Couldn't resolve default property of object ObjectX.ResistanceWithArmor. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
						'UPGRADE_WARNING: Couldn't resolve default property of object GetVarContext. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
						GetVarContext = ObjectX.ResistanceWithArmor(AsVar - 8)
					Case 16 ' Col
						'UPGRADE_WARNING: Couldn't resolve default property of object ObjectX.Col. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
						'UPGRADE_WARNING: Couldn't resolve default property of object GetVarContext. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
						GetVarContext = ObjectX.Col
					Case 17 ' Comments
						'UPGRADE_WARNING: Couldn't resolve default property of object ObjectX.Comments. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
						'UPGRADE_WARNING: Couldn't resolve default property of object GetVarContext. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
						GetVarContext = ObjectX.Comments
					Case 18 ' ActionPoints
						'UPGRADE_WARNING: Couldn't resolve default property of object ObjectX.ActionPoints. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
						'UPGRADE_WARNING: Couldn't resolve default property of object GetVarContext. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
						GetVarContext = ObjectX.ActionPoints
					Case 19 ' ExperiencePoints
						'UPGRADE_WARNING: Couldn't resolve default property of object ObjectX.ExperiencePoints. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
						'UPGRADE_WARNING: Couldn't resolve default property of object GetVarContext. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
						GetVarContext = ObjectX.ExperiencePoints
					Case 20 ' ExperienceLevel
						'UPGRADE_WARNING: Couldn't resolve default property of object ObjectX.Level. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
						'UPGRADE_WARNING: Couldn't resolve default property of object GetVarContext. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
						GetVarContext = ObjectX.Level
					Case 21 ' Not Used --> Now Home
						' [Titi 2.4.9]
						'UPGRADE_WARNING: Couldn't resolve default property of object ObjectX.Home. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
						'UPGRADE_WARNING: Couldn't resolve default property of object GetVarContext. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
						GetVarContext = ObjectX.Home
					Case 22 ' Greed
						'UPGRADE_WARNING: Couldn't resolve default property of object ObjectX.Greed. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
						'UPGRADE_WARNING: Couldn't resolve default property of object GetVarContext. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
						GetVarContext = ObjectX.Greed
					Case 23 ' HealthNow
						'UPGRADE_WARNING: Couldn't resolve default property of object ObjectX.HPNow. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
						'UPGRADE_WARNING: Couldn't resolve default property of object GetVarContext. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
						GetVarContext = ObjectX.HPNow
					Case 24 ' HealthMax
						'UPGRADE_WARNING: Couldn't resolve default property of object ObjectX.HPMax. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
						'UPGRADE_WARNING: Couldn't resolve default property of object GetVarContext. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
						GetVarContext = ObjectX.HPMax
					Case 25 ' Lunacy
						'UPGRADE_WARNING: Couldn't resolve default property of object ObjectX.Lunacy. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
						'UPGRADE_WARNING: Couldn't resolve default property of object GetVarContext. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
						GetVarContext = ObjectX.Lunacy
					Case 26 ' Lust
						'UPGRADE_WARNING: Couldn't resolve default property of object ObjectX.Lust. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
						'UPGRADE_WARNING: Couldn't resolve default property of object GetVarContext. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
						GetVarContext = ObjectX.Lust
					Case 27 ' Not Used
					Case 28 ' IsAgressive
						'UPGRADE_WARNING: Couldn't resolve default property of object ObjectX.Agressive. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
						GetVarContext = System.Math.Abs(ObjectX.Agressive)
					Case 29 ' IsDMControlled
						'UPGRADE_WARNING: Couldn't resolve default property of object ObjectX.DMControlled. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
						GetVarContext = System.Math.Abs(ObjectX.DMControlled)
					Case 30 ' IsFriendly
						'UPGRADE_WARNING: Couldn't resolve default property of object ObjectX.Friendly. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
						GetVarContext = System.Math.Abs(ObjectX.Friendly)
					Case 31 ' IsGuarding
						'UPGRADE_WARNING: Couldn't resolve default property of object ObjectX.Guard. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
						GetVarContext = System.Math.Abs(ObjectX.Guard)
					Case 32 To 39 ' ProtectionSharp, Blunt, Cold, Fire, Evil, Holy, Magic, Mind
						'UPGRADE_WARNING: Couldn't resolve default property of object ObjectX.ResistanceTypeWithArmor. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
						'UPGRADE_WARNING: Couldn't resolve default property of object GetVarContext. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
						GetVarContext = ObjectX.ResistanceTypeWithArmor(AsVar - 31)
					Case 46 ' IsTypeAnimal
						'UPGRADE_WARNING: Couldn't resolve default property of object ObjectX.Family. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
						GetVarContext = System.Math.Abs(ObjectX.Family(4))
					Case 47 ' IsTypeBird
						'UPGRADE_WARNING: Couldn't resolve default property of object ObjectX.Family. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
						GetVarContext = System.Math.Abs(ObjectX.Family(6))
					Case 48 ' IsTypeBlob
						'UPGRADE_WARNING: Couldn't resolve default property of object ObjectX.Family. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
						GetVarContext = System.Math.Abs(ObjectX.Family(3))
					Case 49 ' IsTypeFamily
						'UPGRADE_WARNING: Couldn't resolve default property of object ObjectX.Family. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
						GetVarContext = System.Math.Abs(ObjectX.Family(8))
					Case 50 ' IsTypeHuge
						'UPGRADE_WARNING: Couldn't resolve default property of object ObjectX.Size. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
						GetVarContext = System.Math.Abs(ObjectX.Size(4))
					Case 51 ' IsTypeHuman
						'UPGRADE_WARNING: Couldn't resolve default property of object ObjectX.Family. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
						GetVarContext = System.Math.Abs(ObjectX.Family(0))
					Case 52 ' IsTypeInsect
						'UPGRADE_WARNING: Couldn't resolve default property of object ObjectX.Family. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
						GetVarContext = System.Math.Abs(ObjectX.Family(2))
					Case 53 ' IsTypeLarge
						'UPGRADE_WARNING: Couldn't resolve default property of object ObjectX.Size. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
						GetVarContext = System.Math.Abs(ObjectX.Size(3))
					Case 54 ' IsTypeMagical
						'UPGRADE_WARNING: Couldn't resolve default property of object ObjectX.Family. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
						GetVarContext = System.Math.Abs(ObjectX.Family(9))
					Case 55 ' IsTypeMedium
						'UPGRADE_WARNING: Couldn't resolve default property of object ObjectX.Size. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
						GetVarContext = System.Math.Abs(ObjectX.Size(2))
					Case 56 ' IsTypeReptile
						'UPGRADE_WARNING: Couldn't resolve default property of object ObjectX.Family. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
						GetVarContext = System.Math.Abs(ObjectX.Family(1))
					Case 57 ' IsTypeSmall
						'UPGRADE_WARNING: Couldn't resolve default property of object ObjectX.Size. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
						GetVarContext = System.Math.Abs(ObjectX.Size(1))
					Case 58 ' IsTypeTiny
						'UPGRADE_WARNING: Couldn't resolve default property of object ObjectX.Size. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
						GetVarContext = System.Math.Abs(ObjectX.Size(0))
					Case 59 ' IsTypeVeggie
						'UPGRADE_WARNING: Couldn't resolve default property of object ObjectX.Family. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
						GetVarContext = System.Math.Abs(ObjectX.Family(5))
					Case 60 ' IsTypeUndead
						'UPGRADE_WARNING: Couldn't resolve default property of object ObjectX.Family. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
						GetVarContext = System.Math.Abs(ObjectX.Family(7))
					Case 75 ' Eternium
						'UPGRADE_WARNING: Couldn't resolve default property of object ObjectX.Eternium. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
						'UPGRADE_WARNING: Couldn't resolve default property of object GetVarContext. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
						GetVarContext = ObjectX.Eternium
					Case 76 ' MapX
						'UPGRADE_WARNING: Couldn't resolve default property of object ObjectX.MapX. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
						'UPGRADE_WARNING: Couldn't resolve default property of object GetVarContext. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
						GetVarContext = ObjectX.MapX
					Case 77 ' MapY
						'UPGRADE_WARNING: Couldn't resolve default property of object ObjectX.MapY. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
						'UPGRADE_WARNING: Couldn't resolve default property of object GetVarContext. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
						GetVarContext = ObjectX.MapY
					Case 78 ' Name
						'UPGRADE_WARNING: Couldn't resolve default property of object ObjectX.Name. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
						'UPGRADE_WARNING: Couldn't resolve default property of object GetVarContext. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
						GetVarContext = ObjectX.Name
					Case 79 ' Strength
						'UPGRADE_WARNING: Couldn't resolve default property of object ObjectX.Strength. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
						'UPGRADE_WARNING: Couldn't resolve default property of object GetVarContext. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
						GetVarContext = ObjectX.Strength
					Case 80 ' Pride
						'UPGRADE_WARNING: Couldn't resolve default property of object ObjectX.Pride. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
						'UPGRADE_WARNING: Couldn't resolve default property of object GetVarContext. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
						GetVarContext = ObjectX.Pride
					Case 81 ' Race
						'UPGRADE_WARNING: Couldn't resolve default property of object ObjectX.Race. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
						'UPGRADE_WARNING: Couldn't resolve default property of object GetVarContext. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
						GetVarContext = ObjectX.Race
					Case 82 ' Revelry
						'UPGRADE_WARNING: Couldn't resolve default property of object ObjectX.Revelry. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
						'UPGRADE_WARNING: Couldn't resolve default property of object GetVarContext. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
						GetVarContext = ObjectX.Revelry
					Case 83 ' Row
						'UPGRADE_WARNING: Couldn't resolve default property of object ObjectX.Row. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
						'UPGRADE_WARNING: Couldn't resolve default property of object GetVarContext. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
						GetVarContext = ObjectX.Row
					Case 84 ' RuneQue1
						'UPGRADE_WARNING: Couldn't resolve default property of object ObjectX.Runes. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
						'UPGRADE_WARNING: Couldn't resolve default property of object GetVarContext. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
						GetVarContext = ObjectX.Runes(0)
					Case 85 ' RuneQue2
						'UPGRADE_WARNING: Couldn't resolve default property of object ObjectX.Runes. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
						'UPGRADE_WARNING: Couldn't resolve default property of object GetVarContext. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
						GetVarContext = ObjectX.Runes(1)
					Case 86 ' RuneQue3
						'UPGRADE_WARNING: Couldn't resolve default property of object ObjectX.Runes. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
						'UPGRADE_WARNING: Couldn't resolve default property of object GetVarContext. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
						GetVarContext = ObjectX.Runes(2)
					Case 87 ' RuneQue4
						'UPGRADE_WARNING: Couldn't resolve default property of object ObjectX.Runes. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
						'UPGRADE_WARNING: Couldn't resolve default property of object GetVarContext. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
						GetVarContext = ObjectX.Runes(3)
					Case 88 ' RuneQue5
						'UPGRADE_WARNING: Couldn't resolve default property of object ObjectX.Runes. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
						'UPGRADE_WARNING: Couldn't resolve default property of object GetVarContext. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
						GetVarContext = ObjectX.Runes(4)
					Case 89 ' RuneQue6
						'UPGRADE_WARNING: Couldn't resolve default property of object ObjectX.Runes. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
						'UPGRADE_WARNING: Couldn't resolve default property of object GetVarContext. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
						GetVarContext = ObjectX.Runes(5)
					Case 90 ' ScarletLetter
						'UPGRADE_WARNING: Couldn't resolve default property of object ObjectX.ScarletLetter. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
						'UPGRADE_WARNING: Couldn't resolve default property of object GetVarContext. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
						GetVarContext = ObjectX.ScarletLetter
					Case 91 ' Agility
						'UPGRADE_WARNING: Couldn't resolve default property of object ObjectX.Agility. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
						'UPGRADE_WARNING: Couldn't resolve default property of object GetVarContext. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
						GetVarContext = ObjectX.Agility
					Case 92 ' SkillPoints
						'UPGRADE_WARNING: Couldn't resolve default property of object ObjectX.SkillPoints. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
						'UPGRADE_WARNING: Couldn't resolve default property of object GetVarContext. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
						GetVarContext = ObjectX.SkillPoints
					Case 93 ' Status
						'UPGRADE_WARNING: Couldn't resolve default property of object ObjectX.StatusText. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
						'UPGRADE_WARNING: Couldn't resolve default property of object GetVarContext. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
						GetVarContext = ObjectX.StatusText
					Case 94 ' Not Used
						' Not Used
					Case 95 ' Defense
						'UPGRADE_WARNING: Couldn't resolve default property of object ObjectX.Defense. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
						'UPGRADE_WARNING: Couldn't resolve default property of object GetVarContext. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
						GetVarContext = ObjectX.Defense
					Case 96 ' Will
						'UPGRADE_WARNING: Couldn't resolve default property of object ObjectX.Will. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
						'UPGRADE_WARNING: Couldn't resolve default property of object GetVarContext. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
						GetVarContext = ObjectX.Will
					Case 97 ' Wrath
						'UPGRADE_WARNING: Couldn't resolve default property of object ObjectX.Wrath. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
						'UPGRADE_WARNING: Couldn't resolve default property of object GetVarContext. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
						GetVarContext = ObjectX.Wrath
					Case 98 ' IsMale
						'UPGRADE_WARNING: Couldn't resolve default property of object ObjectX.Male. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
						'UPGRADE_WARNING: Couldn't resolve default property of object GetVarContext. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
						GetVarContext = ObjectX.Male
					Case 99 ' IsUnconscious
						'UPGRADE_WARNING: Couldn't resolve default property of object ObjectX.Unconscious. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
						'UPGRADE_WARNING: Couldn't resolve default property of object GetVarContext. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
						GetVarContext = ObjectX.Unconscious
					Case 100 ' IsFrozen
						'UPGRADE_WARNING: Couldn't resolve default property of object ObjectX.Frozen. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
						'UPGRADE_WARNING: Couldn't resolve default property of object GetVarContext. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
						GetVarContext = ObjectX.Frozen
					Case 101 ' RangeToTarget
						'UPGRADE_WARNING: Couldn't resolve default property of object ObjectX.Row. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
						'UPGRADE_WARNING: Couldn't resolve default property of object ObjectX.Col. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
						GetVarContext = Int(System.Math.Sqrt((ObjectX.Col - FrameNow.CreatureX(4).Col) ^ 2 + (ObjectX.Row - FrameNow.CreatureX(4).Row) ^ 2))
					Case 102 ' ActionPointsMax
						'UPGRADE_WARNING: Couldn't resolve default property of object ObjectX.ActionPointsMax. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
						'UPGRADE_WARNING: Couldn't resolve default property of object GetVarContext. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
						GetVarContext = ObjectX.ActionPointsMax
					Case 103 To 110 ' ProtectionBonusSharp, Blunt, Cold, Fire, Evil, Holy, Magic, Mind
						'UPGRADE_WARNING: Couldn't resolve default property of object ObjectX.ResistanceBonus. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
						'UPGRADE_WARNING: Couldn't resolve default property of object GetVarContext. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
						GetVarContext = ObjectX.ResistanceBonus(AsVar - 102)
					Case 111 ' ResistanceBonus (everthing but the specials above)
						'UPGRADE_WARNING: Couldn't resolve default property of object ObjectX.ResistanceBonus. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
						'UPGRADE_WARNING: Couldn't resolve default property of object GetVarContext. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
						GetVarContext = ObjectX.ResistanceBonus(0)
					Case 133 ' EterniumMax
						'UPGRADE_WARNING: Couldn't resolve default property of object ObjectX.EterniumMax. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
						'UPGRADE_WARNING: Couldn't resolve default property of object GetVarContext. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
						GetVarContext = ObjectX.EterniumMax
					Case 134 ' WeightMax
						'UPGRADE_WARNING: Couldn't resolve default property of object ObjectX.MaxWeight. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
						'UPGRADE_WARNING: Couldn't resolve default property of object GetVarContext. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
						GetVarContext = ObjectX.MaxWeight
					Case 135 ' Weight
						'UPGRADE_WARNING: Couldn't resolve default property of object ObjectX.Weight. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
						'UPGRADE_WARNING: Couldn't resolve default property of object GetVarContext. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
						GetVarContext = ObjectX.Weight
					Case 136 ' Bulk
						'UPGRADE_WARNING: Couldn't resolve default property of object ObjectX.Bulk. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
						'UPGRADE_WARNING: Couldn't resolve default property of object GetVarContext. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
						GetVarContext = ObjectX.Bulk
					Case 137 ' StrengthBonus
						'UPGRADE_WARNING: Couldn't resolve default property of object ObjectX.StrengthBonus. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
						'UPGRADE_WARNING: Couldn't resolve default property of object GetVarContext. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
						GetVarContext = ObjectX.StrengthBonus
					Case 138 ' AgilityBonus
						'UPGRADE_WARNING: Couldn't resolve default property of object ObjectX.AgilityBonus. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
						'UPGRADE_WARNING: Couldn't resolve default property of object GetVarContext. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
						GetVarContext = ObjectX.AgilityBonus
					Case 139 ' WillBonus
						'UPGRADE_WARNING: Couldn't resolve default property of object ObjectX.WillBonus. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
						'UPGRADE_WARNING: Couldn't resolve default property of object GetVarContext. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
						GetVarContext = ObjectX.WillBonus
					Case 140 ' AttackBonus
						'UPGRADE_WARNING: Couldn't resolve default property of object ObjectX.AttackBonus. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
						'UPGRADE_WARNING: Couldn't resolve default property of object GetVarContext. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
						GetVarContext = ObjectX.AttackBonus
					Case 141 ' DamageBonus
						'UPGRADE_WARNING: Couldn't resolve default property of object ObjectX.DamageBonus. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
						'UPGRADE_WARNING: Couldn't resolve default property of object GetVarContext. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
						GetVarContext = ObjectX.DamageBonus
					Case 142 ' DefenseBonus
						'UPGRADE_WARNING: Couldn't resolve default property of object ObjectX.DefenseBonus. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
						'UPGRADE_WARNING: Couldn't resolve default property of object GetVarContext. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
						GetVarContext = ObjectX.DefenseBonus
					Case 143 ' ArmorBonus
						'UPGRADE_WARNING: Couldn't resolve default property of object ObjectX.ArmorBonus. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
						'UPGRADE_WARNING: Couldn't resolve default property of object GetVarContext. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
						GetVarContext = ObjectX.ArmorBonus
					Case 144 ' IsSpellCaster
						'UPGRADE_WARNING: Couldn't resolve default property of object ObjectX.IsSpellCaster. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
						GetVarContext = System.Math.Abs(ObjectX.IsSpellCaster)
					Case 145 ' Money
						'UPGRADE_WARNING: Couldn't resolve default property of object ObjectX.Money. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
						'UPGRADE_WARNING: Couldn't resolve default property of object GetVarContext. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
						GetVarContext = ObjectX.Money
					Case 146 ' ActionPointsMax
						'UPGRADE_WARNING: Couldn't resolve default property of object ObjectX.ActionPointsMax. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
						'UPGRADE_WARNING: Couldn't resolve default property of object GetVarContext. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
						GetVarContext = ObjectX.ActionPointsMax
					Case 147 ' IsAfraid
						'UPGRADE_WARNING: Couldn't resolve default property of object ObjectX.Afraid. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
						GetVarContext = System.Math.Abs(ObjectX.Afraid)
					Case 148 ' IsInanimate
						'UPGRADE_WARNING: Couldn't resolve default property of object ObjectX.IsInanimate. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
						GetVarContext = System.Math.Abs(ObjectX.IsInanimate)
					Case 149 ' Not Used
					Case 150 ' CombatRank
						'UPGRADE_WARNING: Couldn't resolve default property of object ObjectX.CombatRank. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
						'UPGRADE_WARNING: Couldn't resolve default property of object GetVarContext. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
						GetVarContext = ObjectX.CombatRank
					Case 151 ' Pronoun_HeShe
						'UPGRADE_WARNING: Couldn't resolve default property of object ObjectX.Pronoun. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
						'UPGRADE_WARNING: Couldn't resolve default property of object GetVarContext. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
						GetVarContext = ObjectX.Pronoun(1)
					Case 152 ' Pronoun_HimHer
						'UPGRADE_WARNING: Couldn't resolve default property of object ObjectX.Pronoun. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
						'UPGRADE_WARNING: Couldn't resolve default property of object GetVarContext. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
						GetVarContext = ObjectX.Pronoun(2)
					Case 153 ' Pronoun_HisHer
						'UPGRADE_WARNING: Couldn't resolve default property of object ObjectX.Pronoun. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
						'UPGRADE_WARNING: Couldn't resolve default property of object GetVarContext. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
						GetVarContext = ObjectX.Pronoun(3)
					Case 154 ' MovementCost
						'UPGRADE_WARNING: Couldn't resolve default property of object ObjectX.MovementCost. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
						'UPGRADE_WARNING: Couldn't resolve default property of object GetVarContext. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
						GetVarContext = ObjectX.MovementCost
					Case 155 ' MovementCostBonus
						'UPGRADE_WARNING: Couldn't resolve default property of object ObjectX.MovementCostBonus. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
						'UPGRADE_WARNING: Couldn't resolve default property of object GetVarContext. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
						GetVarContext = ObjectX.MovementCostBonus
					Case 156 ' Index
						'UPGRADE_WARNING: Couldn't resolve default property of object ObjectX.Index. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
						'UPGRADE_WARNING: Couldn't resolve default property of object GetVarContext. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
						GetVarContext = ObjectX.Index
					Case 157 ' PictureFile
						'UPGRADE_WARNING: Couldn't resolve default property of object ObjectX.PictureFile. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
						'UPGRADE_WARNING: Couldn't resolve default property of object GetVarContext. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
						GetVarContext = ObjectX.PictureFile
					Case 158 ' FaceLeft
						'UPGRADE_WARNING: Couldn't resolve default property of object ObjectX.FaceLeft. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
						'UPGRADE_WARNING: Couldn't resolve default property of object GetVarContext. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
						GetVarContext = ObjectX.FaceLeft
					Case 159 ' FaceTop
						'UPGRADE_WARNING: Couldn't resolve default property of object ObjectX.FaceTop. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
						'UPGRADE_WARNING: Couldn't resolve default property of object GetVarContext. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
						GetVarContext = ObjectX.FaceTop
					Case 160 ' Initiative
						'UPGRADE_WARNING: Couldn't resolve default property of object ObjectX.Initiative. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
						'UPGRADE_WARNING: Couldn't resolve default property of object GetVarContext. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
						GetVarContext = ObjectX.Initiative
					Case 161 ' Facing
						'UPGRADE_WARNING: Couldn't resolve default property of object ObjectX.Facing. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
						'UPGRADE_WARNING: Couldn't resolve default property of object GetVarContext. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
						GetVarContext = ObjectX.Facing
				End Select
			Case 17, 18, 19, 20, 21 'ItemNow, ItemA, ItemB, ItemC, ItemTarget
				ObjectX = FrameNow.ItemX(AsContext - 17)
				Select Case AsVar
					Case 0 ' Protection%
						'UPGRADE_WARNING: Couldn't resolve default property of object ObjectX.Resistance. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
						'UPGRADE_WARNING: Couldn't resolve default property of object GetVarContext. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
						GetVarContext = ObjectX.Resistance
					Case 1 ' Bulk
						'UPGRADE_WARNING: Couldn't resolve default property of object ObjectX.Bulk. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
						'UPGRADE_WARNING: Couldn't resolve default property of object GetVarContext. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
						GetVarContext = ObjectX.Bulk
					Case 2 ' CanCombine
						'UPGRADE_WARNING: Couldn't resolve default property of object ObjectX.CanCombine. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
						GetVarContext = System.Math.Abs(ObjectX.CanCombine)
					Case 3 ' ArmorType
						'UPGRADE_WARNING: Couldn't resolve default property of object ObjectX.WearType. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
						'UPGRADE_WARNING: Couldn't resolve default property of object GetVarContext. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
						GetVarContext = modBD.SetUpWearType(ObjectX.WearType)
					Case 16 ' Capacity
						'UPGRADE_WARNING: Couldn't resolve default property of object ObjectX.Capacity. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
						'UPGRADE_WARNING: Couldn't resolve default property of object GetVarContext. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
						GetVarContext = ObjectX.Capacity
					Case 18 ' Comments
						'UPGRADE_WARNING: Couldn't resolve default property of object ObjectX.Comments. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
						'UPGRADE_WARNING: Couldn't resolve default property of object GetVarContext. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
						GetVarContext = ObjectX.Comments
					Case 19 ' Count
						'UPGRADE_WARNING: Couldn't resolve default property of object ObjectX.Count. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
						'UPGRADE_WARNING: Couldn't resolve default property of object GetVarContext. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
						GetVarContext = ObjectX.Count
					Case 20 ' DamageDice
						'UPGRADE_WARNING: Couldn't resolve default property of object ObjectX.Damage. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
						'UPGRADE_WARNING: Couldn't resolve default property of object GetVarContext. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
						GetVarContext = ObjectX.Damage
					Case 21 ' DamageBonus
						'UPGRADE_WARNING: Couldn't resolve default property of object ObjectX.DamageBonus. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
						'UPGRADE_WARNING: Couldn't resolve default property of object GetVarContext. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
						GetVarContext = ObjectX.DamageBonus
					Case 22 ' IsWeaponMelee
						'UPGRADE_WARNING: Couldn't resolve default property of object ObjectX.Family. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
						GetVarContext = System.Math.Abs(CInt(ObjectX.Family = 5))
					Case 23 ' IsWeaponRanged
						'UPGRADE_WARNING: Couldn't resolve default property of object ObjectX.Family. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
						GetVarContext = System.Math.Abs(CInt(ObjectX.Family = 7))
					Case 24 ' IsEquipped
						'UPGRADE_WARNING: Couldn't resolve default property of object ObjectX.IsReady. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
						GetVarContext = System.Math.Abs(ObjectX.IsReady)
					Case 25 ' IsWeaponThrown
						'UPGRADE_WARNING: Couldn't resolve default property of object ObjectX.Family. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
						GetVarContext = System.Math.Abs(CInt(ObjectX.Family = 8))
					Case 26 ' Key
						'UPGRADE_WARNING: Couldn't resolve default property of object ObjectX.KeyBits. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
						GetVarContext = System.Math.Abs(ObjectX.KeyBits)
					Case 27 ' ProtectionBonusType
						'UPGRADE_WARNING: Couldn't resolve default property of object ObjectX.ResistanceType. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
						'UPGRADE_WARNING: Couldn't resolve default property of object GetVarContext. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
						GetVarContext = ObjectX.ResistanceType
					Case 28 ' ProtectionBonus
						'UPGRADE_WARNING: Couldn't resolve default property of object ObjectX.ResistanceBonus. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
						'UPGRADE_WARNING: Couldn't resolve default property of object GetVarContext. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
						GetVarContext = ObjectX.ResistanceBonus
					Case 29 ' IsWeaponAmmo
						'UPGRADE_WARNING: Couldn't resolve default property of object ObjectX.Family. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
						GetVarContext = System.Math.Abs(CInt(ObjectX.Family = 6))
					Case 31 ' IsRangeLong
						'UPGRADE_WARNING: Couldn't resolve default property of object ObjectX.RangeType. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
						GetVarContext = System.Math.Abs(CInt(ObjectX.RangeType = 2))
					Case 32 ' IsRangeMedium
						'UPGRADE_WARNING: Couldn't resolve default property of object ObjectX.RangeType. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
						GetVarContext = System.Math.Abs(CInt(ObjectX.RangeType = 1))
					Case 33 ' IsRangeShort
						'UPGRADE_WARNING: Couldn't resolve default property of object ObjectX.RangeType. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
						GetVarContext = System.Math.Abs(CInt(ObjectX.RangeType = 0))
					Case 36 ' IsSoftBulk
						'UPGRADE_WARNING: Couldn't resolve default property of object ObjectX.SoftCapacity. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
						GetVarContext = System.Math.Abs(ObjectX.SoftCapacity)
					Case 37 ' IsDamageTypeSharp
						'UPGRADE_WARNING: Couldn't resolve default property of object ObjectX.DamageType. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
						GetVarContext = System.Math.Abs(CInt(ObjectX.DamageType = 1))
					Case 38 ' IsDamageTypeBlunt
						'UPGRADE_WARNING: Couldn't resolve default property of object ObjectX.DamageType. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
						GetVarContext = System.Math.Abs(CInt(ObjectX.DamageType = 2))
					Case 39 ' IsDamageTypeCold
						'UPGRADE_WARNING: Couldn't resolve default property of object ObjectX.DamageType. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
						GetVarContext = System.Math.Abs(CInt(ObjectX.DamageType = 3))
					Case 40 ' IsDamageTypeFire
						'UPGRADE_WARNING: Couldn't resolve default property of object ObjectX.DamageType. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
						GetVarContext = System.Math.Abs(CInt(ObjectX.DamageType = 4))
					Case 41 ' IsDamageTypeHoly
						'UPGRADE_WARNING: Couldn't resolve default property of object ObjectX.DamageType. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
						GetVarContext = System.Math.Abs(CInt(ObjectX.DamageType = 6))
					Case 42 ' IsDamageTypeEvil
						'UPGRADE_WARNING: Couldn't resolve default property of object ObjectX.DamageType. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
						GetVarContext = System.Math.Abs(CInt(ObjectX.DamageType = 5))
					Case 43 ' IsDamageTypeMagic
						'UPGRADE_WARNING: Couldn't resolve default property of object ObjectX.DamageType. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
						GetVarContext = System.Math.Abs(CInt(ObjectX.DamageType = 7))
					Case 44 ' IsDamageTypeMind
						'UPGRADE_WARNING: Couldn't resolve default property of object ObjectX.DamageType. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
						GetVarContext = System.Math.Abs(CInt(ObjectX.DamageType = 8))
					Case 50 ' MapX
						'UPGRADE_WARNING: Couldn't resolve default property of object ObjectX.MapX. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
						'UPGRADE_WARNING: Couldn't resolve default property of object GetVarContext. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
						GetVarContext = ObjectX.MapX
					Case 51 ' MapY
						'UPGRADE_WARNING: Couldn't resolve default property of object ObjectX.MapY. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
						'UPGRADE_WARNING: Couldn't resolve default property of object GetVarContext. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
						GetVarContext = ObjectX.MapY
					Case 52 ' Name
						'UPGRADE_WARNING: Couldn't resolve default property of object ObjectX.NameText. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
						'UPGRADE_WARNING: Couldn't resolve default property of object GetVarContext. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
						GetVarContext = ObjectX.NameText
					Case 53 ' DefenseBonus
						'UPGRADE_WARNING: Couldn't resolve default property of object ObjectX.Defense. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
						'UPGRADE_WARNING: Couldn't resolve default property of object GetVarContext. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
						GetVarContext = ObjectX.Defense
					Case 55 ' AttackBonus
						'UPGRADE_WARNING: Couldn't resolve default property of object ObjectX.AttackBonus. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
						'UPGRADE_WARNING: Couldn't resolve default property of object GetVarContext. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
						GetVarContext = ObjectX.AttackBonus
					Case 56 ' UseAsDescription
						'UPGRADE_WARNING: Couldn't resolve default property of object ObjectX.UseDescription. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
						GetVarContext = System.Math.Abs(ObjectX.UseDescription)
					Case 57 ' Value
						'UPGRADE_WARNING: Couldn't resolve default property of object ObjectX.Value. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
						'UPGRADE_WARNING: Couldn't resolve default property of object GetVarContext. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
						GetVarContext = ObjectX.Value
					Case 58 ' Weight
						'UPGRADE_WARNING: Couldn't resolve default property of object ObjectX.Weight. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
						'UPGRADE_WARNING: Couldn't resolve default property of object GetVarContext. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
						GetVarContext = ObjectX.Weight
					Case 59 ' RequiresTwoHands
						'UPGRADE_WARNING: Couldn't resolve default property of object ObjectX.WearType. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
						GetVarContext = System.Math.Abs(CInt(ObjectX.WearType = 11))
					Case 60 ' IsInHand
						'UPGRADE_WARNING: Couldn't resolve default property of object ObjectX.InHand. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
						GetVarContext = System.Math.Abs(ObjectX.InHand)
					Case 61 ' AmmoType
						'UPGRADE_WARNING: Couldn't resolve default property of object ObjectX.ShootType. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
						'UPGRADE_WARNING: Couldn't resolve default property of object GetVarContext. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
						GetVarContext = ObjectX.ShootType
					Case 64 ' Family
						'UPGRADE_WARNING: Couldn't resolve default property of object ObjectX.Family. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
						'UPGRADE_WARNING: Couldn't resolve default property of object GetVarContext. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
						GetVarContext = modBD.SetUpItemFamily(ObjectX.Family)
					Case 65 ' ActionPoints
						'UPGRADE_WARNING: Couldn't resolve default property of object ObjectX.ActionPoints. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
						'UPGRADE_WARNING: Couldn't resolve default property of object GetVarContext. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
						GetVarContext = ObjectX.ActionPoints
				End Select
			Case 22, 23, 24, 25, 26 'TriggerNow, TriggerA, TriggerB, TriggerC, TriggerTarget
				ObjectX = FrameNow.TriggerX(AsContext - 22)
				Select Case AsVar
					Case 0 'ByteA
						'UPGRADE_WARNING: Couldn't resolve default property of object ObjectX.VarA. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
						'UPGRADE_WARNING: Couldn't resolve default property of object GetVarContext. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
						GetVarContext = ObjectX.VarA
					Case 1 'ByteB
						'UPGRADE_WARNING: Couldn't resolve default property of object ObjectX.VarB. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
						'UPGRADE_WARNING: Couldn't resolve default property of object GetVarContext. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
						GetVarContext = ObjectX.VarB
					Case 2 'ByteC
						'UPGRADE_WARNING: Couldn't resolve default property of object ObjectX.VarC. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
						'UPGRADE_WARNING: Couldn't resolve default property of object GetVarContext. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
						GetVarContext = ObjectX.VarC
					Case 3 'Comments
						'UPGRADE_WARNING: Couldn't resolve default property of object ObjectX.Comments. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
						'UPGRADE_WARNING: Couldn't resolve default property of object GetVarContext. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
						GetVarContext = ObjectX.Comments
					Case 4 ' Not Used
					Case 5 'IsTimed
						'UPGRADE_WARNING: Couldn't resolve default property of object ObjectX.IsTimed. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
						GetVarContext = System.Math.Abs(ObjectX.IsTimed)
					Case 6 'IsCurse
						'UPGRADE_WARNING: Couldn't resolve default property of object ObjectX.IsCurse. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
						GetVarContext = System.Math.Abs(ObjectX.IsCurse)
					Case 7 'IsEvil
						'UPGRADE_WARNING: Couldn't resolve default property of object ObjectX.IsEvil. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
						GetVarContext = System.Math.Abs(ObjectX.IsEvil)
					Case 8 'IsFear
						'UPGRADE_WARNING: Couldn't resolve default property of object ObjectX.IsFear. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
						GetVarContext = System.Math.Abs(ObjectX.IsFear)
					Case 9 'IsFish
						'UPGRADE_WARNING: Couldn't resolve default property of object ObjectX.IsFish. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
						GetVarContext = System.Math.Abs(ObjectX.IsFish)
					Case 10 'IsGreed
						'UPGRADE_WARNING: Couldn't resolve default property of object ObjectX.IsGreed. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
						GetVarContext = System.Math.Abs(ObjectX.IsGreed)
					Case 11 'IsLunacy
						'UPGRADE_WARNING: Couldn't resolve default property of object ObjectX.IsLunacy. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
						GetVarContext = System.Math.Abs(ObjectX.IsLunacy)
					Case 12 'IsLust
						'UPGRADE_WARNING: Couldn't resolve default property of object ObjectX.IsLust. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
						GetVarContext = System.Math.Abs(ObjectX.IsLust)
					Case 13 'IsPoison
						'UPGRADE_WARNING: Couldn't resolve default property of object ObjectX.IsPoison. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
						GetVarContext = System.Math.Abs(ObjectX.IsPoison)
					Case 14 'IsPride
						'UPGRADE_WARNING: Couldn't resolve default property of object ObjectX.IsPride. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
						GetVarContext = System.Math.Abs(ObjectX.IsPride)
					Case 15 'IsMagic
						'UPGRADE_WARNING: Couldn't resolve default property of object ObjectX.IsMagic. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
						GetVarContext = System.Math.Abs(ObjectX.IsMagic)
					Case 16 'IsSkill
						'UPGRADE_WARNING: Couldn't resolve default property of object ObjectX.IsSkill. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
						GetVarContext = System.Math.Abs(ObjectX.IsSkill)
					Case 17 'IsRevelry
						'UPGRADE_WARNING: Couldn't resolve default property of object ObjectX.IsRevelry. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
						GetVarContext = System.Math.Abs(ObjectX.IsRevelry)
					Case 18 'IsTrap
						'UPGRADE_WARNING: Couldn't resolve default property of object ObjectX.IsTrap. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
						GetVarContext = System.Math.Abs(ObjectX.IsTrap)
					Case 19 'IsWrath
						'UPGRADE_WARNING: Couldn't resolve default property of object ObjectX.IsWrath. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
						GetVarContext = System.Math.Abs(ObjectX.IsWrath)
					Case 20 'Name
						'UPGRADE_WARNING: Couldn't resolve default property of object ObjectX.Name. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
						'UPGRADE_WARNING: Couldn't resolve default property of object GetVarContext. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
						GetVarContext = ObjectX.Name
					Case 21 ' SkillPoints
						'UPGRADE_WARNING: Couldn't resolve default property of object ObjectX.SkillPoints. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
						'UPGRADE_WARNING: Couldn't resolve default property of object GetVarContext. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
						GetVarContext = ObjectX.SkillPoints
					Case 22 'TriggerType
						'UPGRADE_WARNING: Couldn't resolve default property of object ObjectX.TriggerType. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
						GetVarContext = TriggerName(ObjectX.TriggerType)
					Case 23 'Turns
						'UPGRADE_WARNING: Couldn't resolve default property of object ObjectX.Turns. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
						'UPGRADE_WARNING: Couldn't resolve default property of object GetVarContext. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
						GetVarContext = ObjectX.Turns
					Case 24 'SkillLevel
						'UPGRADE_WARNING: Couldn't resolve default property of object ObjectX.Turns. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
						'UPGRADE_WARNING: Couldn't resolve default property of object ObjectX.SkillPoints. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
						If ObjectX.SkillPoints > 0 And ObjectX.Turns > 0 Then
							'UPGRADE_WARNING: Couldn't resolve default property of object ObjectX.Turns. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
							'UPGRADE_WARNING: Couldn't resolve default property of object ObjectX.SkillPoints. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
							GetVarContext = Int(ObjectX.SkillPoints / ObjectX.Turns)
						Else
							'UPGRADE_WARNING: Couldn't resolve default property of object GetVarContext. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
							GetVarContext = 0
						End If
				End Select
			Case 27 'Local
				Select Case AsVar
					Case 0 'Abort
						'UPGRADE_WARNING: Couldn't resolve default property of object GetVarContext. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
						GetVarContext = FrameNow.Abort
					Case 1 'ByteA
						'UPGRADE_WARNING: Couldn't resolve default property of object GetVarContext. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
						GetVarContext = FrameNow.ByteA
					Case 2 'ByteB
						'UPGRADE_WARNING: Couldn't resolve default property of object GetVarContext. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
						GetVarContext = FrameNow.ByteB
					Case 3 'ByteC
						'UPGRADE_WARNING: Couldn't resolve default property of object GetVarContext. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
						GetVarContext = FrameNow.ByteC
					Case 4 'Fail
						'UPGRADE_WARNING: Couldn't resolve default property of object GetVarContext. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
						GetVarContext = FrameNow.Fail
					Case 5 'IntegerA
						'UPGRADE_WARNING: Couldn't resolve default property of object GetVarContext. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
						GetVarContext = FrameNow.IntegerA
					Case 6 'IntegerB
						'UPGRADE_WARNING: Couldn't resolve default property of object GetVarContext. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
						GetVarContext = FrameNow.IntegerB
					Case 7 'IntegerC
						'UPGRADE_WARNING: Couldn't resolve default property of object GetVarContext. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
						GetVarContext = FrameNow.IntegerC
					Case 8 'StopExit
						'UPGRADE_WARNING: Couldn't resolve default property of object GetVarContext. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
						GetVarContext = FrameNow.StopExit
					Case 9 'TextA
						'UPGRADE_WARNING: Couldn't resolve default property of object GetVarContext. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
						GetVarContext = FrameNow.TextA
					Case 10 'TextB
						'UPGRADE_WARNING: Couldn't resolve default property of object GetVarContext. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
						GetVarContext = FrameNow.TextB
					Case 11 'TextC
						'UPGRADE_WARNING: Couldn't resolve default property of object GetVarContext. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
						GetVarContext = FrameNow.TextC
					Case 12 'Random
						'UPGRADE_WARNING: Couldn't resolve default property of object GetVarContext. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
						GetVarContext = FrameNow.Random
					Case 13 ' FoundIt
						'UPGRADE_WARNING: Couldn't resolve default property of object GetVarContext. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
						GetVarContext = FrameNow.FoundIt
				End Select
			Case 28 'Global
				Select Case AsVar
					Case 0 ' False
						'UPGRADE_WARNING: Couldn't resolve default property of object GetVarContext. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
						GetVarContext = 0
					Case 1 ' True
						'UPGRADE_WARNING: Couldn't resolve default property of object GetVarContext. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
						GetVarContext = 1
					Case 2 ' InCombat
						GetVarContext = System.Math.Abs(CInt(frmMain.picGrid.Visible = True))
					Case 3 ' DieTypeRoll
						'UPGRADE_WARNING: Couldn't resolve default property of object GetVarContext. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
						GetVarContext = GlobalDieTypeRoll
					Case 4 ' DieCountRoll
						'UPGRADE_WARNING: Couldn't resolve default property of object GetVarContext. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
						GetVarContext = GlobalDieCountRoll
					Case 5 ' ArmorRoll
						'UPGRADE_WARNING: Couldn't resolve default property of object GetVarContext. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
						GetVarContext = GlobalArmorRoll
					Case 6 ' HitRoll
						'UPGRADE_WARNING: Couldn't resolve default property of object GetVarContext. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
						GetVarContext = GlobalAttackRoll
					Case 7 ' DamageRoll
						'UPGRADE_WARNING: Couldn't resolve default property of object GetVarContext. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
						GetVarContext = GlobalDamageRoll
					Case 8 ' Offer
						'UPGRADE_WARNING: Couldn't resolve default property of object GetVarContext. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
						GetVarContext = GlobalOffer
					Case 9 To 28 ' RunePool
						'UPGRADE_WARNING: Couldn't resolve default property of object GetVarContext. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
						GetVarContext = frmMain.GetRunePool(AsVar - 9)
					Case 29 ' SkillLevel
						'UPGRADE_WARNING: Couldn't resolve default property of object GetVarContext. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
						GetVarContext = GlobalSkillLevel
					Case 30 'IsAttackStyleAir
						GetVarContext = System.Math.Abs(CInt((GlobalDamageStyle And &H200) > 0))
					Case 31 'IsAttackStyleBlunt
						GetVarContext = System.Math.Abs(CInt((GlobalDamageStyle And &H800) > 0))
					Case 32 'IsAttackStyleEarth
						GetVarContext = System.Math.Abs(CInt((GlobalDamageStyle And &H100) > 0))
					Case 33 'IsAttackStyleFire
						GetVarContext = System.Math.Abs(CInt((GlobalDamageStyle And &H80) > 0))
					Case 34 'IsAttackStyleHoly
						GetVarContext = System.Math.Abs(CInt((GlobalDamageStyle And &H8) > 0))
					Case 35 'IsAttackStyleDisease
						GetVarContext = System.Math.Abs(CInt((GlobalDamageStyle And &H2) > 0))
					Case 36 'IsAttackStyleMagic
						GetVarContext = System.Math.Abs(CInt((GlobalDamageStyle And &H20) > 0))
					Case 37 'IsAttackStyleSharp
						GetVarContext = System.Math.Abs(CInt((GlobalDamageStyle And &H400) > 0))
					Case 38 'IsAttackStyleEvil
						GetVarContext = System.Math.Abs(CInt((GlobalDamageStyle And &H4) > 0))
					Case 39 'IsAttackStyleCold
						GetVarContext = System.Math.Abs(CInt((GlobalDamageStyle And &H1) > 0))
					Case 40 'IsAttackStyleIllusion
						GetVarContext = System.Math.Abs(CInt((GlobalDamageStyle And &H20) > 0))
					Case 41 'IsAttackStyleWater
						GetVarContext = System.Math.Abs(CInt((GlobalDamageStyle And &H80) > 0))
					Case 42 'IsAttackStyleFear
						GetVarContext = System.Math.Abs(CInt((GlobalDamageStyle And &H10) > 0))
					Case 43 To 57
						' Not Used
					Case 58 ' IntegerA
						'UPGRADE_WARNING: Couldn't resolve default property of object GetVarContext. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
						GetVarContext = GlobalIntegerA
					Case 59 ' IntegerB
						'UPGRADE_WARNING: Couldn't resolve default property of object GetVarContext. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
						GetVarContext = GlobalIntegerB
					Case 60 ' IntegerC
						'UPGRADE_WARNING: Couldn't resolve default property of object GetVarContext. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
						GetVarContext = GlobalIntegerC
					Case 61 ' TextA
						'UPGRADE_WARNING: Couldn't resolve default property of object GetVarContext. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
						GetVarContext = GlobalTextA
					Case 62 ' TextB
						'UPGRADE_WARNING: Couldn't resolve default property of object GetVarContext. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
						GetVarContext = GlobalTextB
					Case 63 ' TextC
						'UPGRADE_WARNING: Couldn't resolve default property of object GetVarContext. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
						GetVarContext = GlobalTextC
					Case 64 ' DayName
						'UPGRADE_WARNING: Couldn't resolve default property of object GetVarContext. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
						GetVarContext = GlobalDayName
					Case 65 ' MoonName
						'UPGRADE_WARNING: Couldn't resolve default property of object GetVarContext. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
						GetVarContext = GlobalMoonName
					Case 66 ' YearName
						'UPGRADE_WARNING: Couldn't resolve default property of object GetVarContext. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
						GetVarContext = GlobalYearName
					Case 67 ' TurnName
						'UPGRADE_WARNING: Couldn't resolve default property of object GetVarContext. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
						GetVarContext = GlobalTurnName
					Case 68 ' PickLockChance
						'UPGRADE_WARNING: Couldn't resolve default property of object GetVarContext. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
						GetVarContext = GlobalPickLockChance
					Case 69 ' Not Used
					Case 70 ' RemoveTrapChance
						'UPGRADE_WARNING: Couldn't resolve default property of object GetVarContext. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
						GetVarContext = GlobalRemoveTrapChance
					Case 71 ' HitLocation
						'UPGRADE_WARNING: Couldn't resolve default property of object GetVarContext. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
						GetVarContext = GlobalHitLocation
					Case 72 ' SpellFizzleChance
						'UPGRADE_WARNING: Couldn't resolve default property of object GetVarContext. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
						GetVarContext = GlobalSpellFizzleChance
					Case 73 ' Ticks
						'UPGRADE_WARNING: Couldn't resolve default property of object GetVarContext. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
						GetVarContext = GlobalTicks
				End Select
			Case 29 'Pos
				'UPGRADE_WARNING: Couldn't resolve default property of object GetVarContext. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
				GetVarContext = AsVar
			Case 30 'Neg
				'UPGRADE_WARNING: Couldn't resolve default property of object GetVarContext. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
				GetVarContext = -AsVar
			Case 31 'Dice
				'UPGRADE_WARNING: Couldn't resolve default property of object GetVarContext. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
				GetVarContext = 0
				For c = 1 To AsVar Mod 5 + 1
					'UPGRADE_WARNING: Couldn't resolve default property of object GetVarContext. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
					GetVarContext = GetVarContext + Int(Rnd() * (Int((AsVar Mod 25) / 5) * 2 + 4) + 1)
				Next c
				If AsVar > 24 Then
					'UPGRADE_WARNING: Couldn't resolve default property of object GetVarContext. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
					GetVarContext = GetVarContext + Int(AsVar / 25)
				End If
			Case 32 'Random
				'UPGRADE_WARNING: Couldn't resolve default property of object GetVarContext. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
				GetVarContext = Int(Rnd() * (AsVar + 1)) + 1
			Case Else
				'UPGRADE_WARNING: Couldn't resolve default property of object GetVarContext. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
				GetVarContext = 0
		End Select
	End Function
	
	'UPGRADE_NOTE: Tome was upgraded to Tome_Renamed. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="A9E4979A-37FA-4718-9994-97DD76ED70A7"'
	'UPGRADE_NOTE: Map was upgraded to Map_Renamed. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="A9E4979A-37FA-4718-9994-97DD76ED70A7"'
	Private Sub PutVarContext(ByRef IntoContext As Short, ByRef IntoVar As Short, ByRef AsValue As Object)
		Dim Map_Renamed As Object
		Dim Tome_Renamed As Object
		Dim ObjectX As Object
		Dim FactoidX As Factoid
		Dim Y, X, Found As Short
		'UPGRADE_NOTE: sByte was upgraded to sByte_Renamed. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="A9E4979A-37FA-4718-9994-97DD76ED70A7"'
		Dim sByte_Renamed As Byte
		Dim sInteger As Short
		Dim sText As String
		Dim ThemeX As Theme
		' Smudge the AsValue into Byte, Integer and Text
		If IsNumeric(AsValue) Then
			'UPGRADE_WARNING: Couldn't resolve default property of object AsValue. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
			If AsValue > 255 Then
				sByte_Renamed = 255
				'UPGRADE_WARNING: Couldn't resolve default property of object AsValue. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
			ElseIf AsValue < 0 Then 
				sByte_Renamed = 0
			Else
				'UPGRADE_WARNING: Couldn't resolve default property of object AsValue. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
				sByte_Renamed = CByte(AsValue)
			End If
			'UPGRADE_WARNING: Couldn't resolve default property of object AsValue. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
			If AsValue > 32000 Then
				sInteger = 32000
				'UPGRADE_WARNING: Couldn't resolve default property of object AsValue. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
			ElseIf AsValue < -32000 Then 
				sInteger = -32000
			Else
				'UPGRADE_WARNING: Couldn't resolve default property of object AsValue. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
				sInteger = CShort(AsValue)
			End If
		Else
			sByte_Renamed = 0 : sInteger = 0
		End If
		'UPGRADE_WARNING: Couldn't resolve default property of object AsValue. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
		sText = VB6.Format(AsValue)
		Select Case IntoContext
			Case 0 ' Tome
				Select Case IntoVar
					Case 0 ' PartyCount
						' Read Only
					Case 1 ' Comments
						'UPGRADE_WARNING: Couldn't resolve default property of object Tome.Comments. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
						Tome.Comments = sText
					Case 2 'Factoids
						'UPGRADE_WARNING: Couldn't resolve default property of object Tome.AddFactoid. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
						FactoidX = Tome.AddFactoid
						FactoidX.Text = sText
					Case 3
						' Not Used
					Case 4 'MapX
						' Read Only
					Case 5 'MapY
						' Read Only
					Case 6 'MoveToX
						'UPGRADE_WARNING: Couldn't resolve default property of object Tome.MoveToX. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
						Tome.MoveToX = sByte_Renamed
					Case 7 'MoveToY
						'UPGRADE_WARNING: Couldn't resolve default property of object Tome.MoveToY. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
						Tome.MoveToY = sByte_Renamed
					Case 8 'Name
						'UPGRADE_WARNING: Couldn't resolve default property of object Tome.Name. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
						Tome.Name = sText
					Case 9
						' Not Used
					Case 10 'TimeCycle
						'UPGRADE_WARNING: Couldn't resolve default property of object Tome.Cycle. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
						Tome.Cycle = sByte_Renamed
					Case 11 'TimeMoon
						'UPGRADE_WARNING: Couldn't resolve default property of object Tome.Moon. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
						Tome.Moon = sByte_Renamed
					Case 12 'TimeTurn
						'UPGRADE_WARNING: Couldn't resolve default property of object Tome.Turn. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
						Tome.Turn = sByte_Renamed
					Case 13 'TimeYear
						'UPGRADE_WARNING: Couldn't resolve default property of object Tome.Year. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
						Tome.Year = sInteger
					Case 14 'RealSeconds
						' Read Only
					Case 15 'RealMinutes
						' Read Only
					Case 16 'RealHours
						' Read Only
					Case 17 'RealDays
						' Read Only
				End Select
			Case 1 'Map
				Select Case IntoVar
					Case 0 'Comments
						'UPGRADE_WARNING: Couldn't resolve default property of object Map.Comments. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
						Map.Comments = sText
					Case 1 'Randomize: Encounter Count
						'UPGRADE_WARNING: Couldn't resolve default property of object Map.DefaultEncounters. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
						Map.DefaultEncounters = sInteger
					Case 2 'Randomize: Height
						'UPGRADE_WARNING: Couldn't resolve default property of object Map.DefaultHeight. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
						Map.DefaultHeight = sInteger
					Case 3 'Randomize: Map Style
						'UPGRADE_WARNING: Couldn't resolve default property of object Map.DefaultStyle. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
						Map.DefaultStyle = sInteger
					Case 4 'Randomize: Map Width
						'UPGRADE_WARNING: Couldn't resolve default property of object Map.DefaultWidth. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
						Map.DefaultWidth = sInteger
					Case 5 'Randomize: Difficulty
					Case 6 'Randomize: ExperiencePoints
						'UPGRADE_WARNING: Couldn't resolve default property of object Map.ExperiencePoints. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
						Map.ExperiencePoints = sInteger
					Case 7 'Height
						' Read Only
					Case 8 'Left
						' Read Only
					Case 9 'Name
						'UPGRADE_WARNING: Couldn't resolve default property of object Map.Name. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
						Map.Name = sText
					Case 10 'ReGenUponEntry
						'UPGRADE_WARNING: Couldn't resolve default property of object Map.GenerateUponEntry. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
						Map.GenerateUponEntry = (sInteger <> 0)
					Case 11 'Rune: Blood
						'UPGRADE_WARNING: Couldn't resolve default property of object Map.Runes. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
						Map.Runes(0) = sInteger
					Case 12 'Rune: Bile
						'UPGRADE_WARNING: Couldn't resolve default property of object Map.Runes. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
						Map.Runes(1) = sInteger
					Case 13 'Rune: Oil
						'UPGRADE_WARNING: Couldn't resolve default property of object Map.Runes. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
						Map.Runes(2) = sInteger
					Case 14 'Rune: Nectar
						'UPGRADE_WARNING: Couldn't resolve default property of object Map.Runes. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
						Map.Runes(3) = sInteger
					Case 15 'Rune: Fire
						'UPGRADE_WARNING: Couldn't resolve default property of object Map.Runes. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
						Map.Runes(4) = sInteger
					Case 16 'Rune: Earth
						'UPGRADE_WARNING: Couldn't resolve default property of object Map.Runes. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
						Map.Runes(5) = sInteger
					Case 17 'Rune: Water
						'UPGRADE_WARNING: Couldn't resolve default property of object Map.Runes. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
						Map.Runes(6) = sInteger
					Case 18 'Rune: Air
						'UPGRADE_WARNING: Couldn't resolve default property of object Map.Runes. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
						Map.Runes(7) = sInteger
					Case 19 'Rune: Time
						'UPGRADE_WARNING: Couldn't resolve default property of object Map.Runes. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
						Map.Runes(8) = sInteger
					Case 20 'Rune: Moon
						'UPGRADE_WARNING: Couldn't resolve default property of object Map.Runes. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
						Map.Runes(9) = sInteger
					Case 21 'Rune: Sun
						'UPGRADE_WARNING: Couldn't resolve default property of object Map.Runes. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
						Map.Runes(10) = sInteger
					Case 22 'Rune: Space
						'UPGRADE_WARNING: Couldn't resolve default property of object Map.Runes. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
						Map.Runes(11) = sInteger
					Case 23 'Rune: Insect
						'UPGRADE_WARNING: Couldn't resolve default property of object Map.Runes. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
						Map.Runes(12) = sInteger
					Case 24 'Rune: Man
						'UPGRADE_WARNING: Couldn't resolve default property of object Map.Runes. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
						Map.Runes(13) = sInteger
					Case 25 'Rune: Animal
						'UPGRADE_WARNING: Couldn't resolve default property of object Map.Runes. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
						Map.Runes(14) = sInteger
					Case 26 'Rune: Fish
						'UPGRADE_WARNING: Couldn't resolve default property of object Map.Runes. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
						Map.Runes(15) = sInteger
					Case 27 'Rune: Twilight
						'UPGRADE_WARNING: Couldn't resolve default property of object Map.Runes. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
						Map.Runes(16) = sInteger
					Case 28 'Rune: Abyss
						'UPGRADE_WARNING: Couldn't resolve default property of object Map.Runes. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
						Map.Runes(17) = sInteger
					Case 29 'Rune: Eternium
						'UPGRADE_WARNING: Couldn't resolve default property of object Map.Runes. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
						Map.Runes(18) = sInteger
					Case 30 'Rune: Dreams
						'UPGRADE_WARNING: Couldn't resolve default property of object Map.Runes. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
						Map.Runes(19) = sInteger
					Case 31 'Top
						' Read Only
					Case 32 'Width
						' Read Only
					Case 33 'IsOutsideMap
						' Read Only
					Case 34 'MapStyle
						' Read Only
				End Select
			Case 2, 3, 4, 5, 6 'TileNow, TileA, TileB, TileC, TileTarget
				X = FrameNow.TileX(IntoContext - 2) : Y = FrameNow.TileY(IntoContext - 2)
				Select Case IntoVar
					Case 0 ' CanSeeDown
						' Read Only
					Case 1 ' CanSeeLeft
						' Read Only
					Case 2 ' CanSeeRight
						' Read Only
					Case 3 ' CanSeeUp
						' Read Only
					Case 4 ' ChanceToOpen
						Found = False
						'UPGRADE_WARNING: Couldn't resolve default property of object Map.TopTile. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
						If Map.TopTile(X, Y) > 0 Then
							'UPGRADE_WARNING: Couldn't resolve default property of object Map.TopTile. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
							'UPGRADE_WARNING: Couldn't resolve default property of object Map.Tiles. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
							If Map.Tiles("L" & Map.TopTile(X, Y)).SwapTile > 0 Then
								'UPGRADE_WARNING: Couldn't resolve default property of object Map.TopTile. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
								'UPGRADE_WARNING: Couldn't resolve default property of object Map.Tiles. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
								Map.Tiles("L" & Map.TopTile(X, Y)).Chance = sByte_Renamed
								Found = True
							End If
						End If
						'UPGRADE_WARNING: Couldn't resolve default property of object Map.MiddleTile. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
						If Map.MiddleTile(X, Y) > 0 And Found = False Then
							'UPGRADE_WARNING: Couldn't resolve default property of object Map.MiddleTile. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
							'UPGRADE_WARNING: Couldn't resolve default property of object Map.Tiles. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
							If Map.Tiles("L" & Map.MiddleTile(X, Y)).SwapTile > 0 Then
								'UPGRADE_WARNING: Couldn't resolve default property of object Map.MiddleTile. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
								'UPGRADE_WARNING: Couldn't resolve default property of object Map.Tiles. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
								Map.Tiles("L" & Map.MiddleTile(X, Y)).Chance = sByte_Renamed
								Found = True
							End If
						End If
						'UPGRADE_WARNING: Couldn't resolve default property of object Map.BottomTile. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
						If Map.BottomTile(X, Y) > 0 And Found = False Then
							'UPGRADE_WARNING: Couldn't resolve default property of object Map.BottomTile. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
							'UPGRADE_WARNING: Couldn't resolve default property of object Map.Tiles. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
							If Map.Tiles("L" & Map.BottomTile(X, Y)).SwapTile > 0 Then
								'UPGRADE_WARNING: Couldn't resolve default property of object Map.BottomTile. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
								'UPGRADE_WARNING: Couldn't resolve default property of object Map.Tiles. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
								Map.Tiles("L" & Map.BottomTile(X, Y)).Chance = sByte_Renamed
								Found = True
							End If
						End If
					Case 5 ' Not Used
					Case 6 ' IsBlockedDown
						' Read Only
					Case 7 ' IsBlockedLeft
						' Read Only
					Case 8 ' IsBlockedRight
						' Read Only
					Case 9 ' IsBlockedUp
						' Read Only
					Case 10 ' Key
						Found = False
						' [Titi 2.4.8] TileTarget.Key was not set
						If X = FrameNow.TileX(4) And Y = FrameNow.TileY(4) Then
							' TileTarget found
							'UPGRADE_WARNING: Couldn't resolve default property of object Map.TopTile. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
							If Map.TopTile(X, Y) > 0 Then
								'UPGRADE_WARNING: Couldn't resolve default property of object Map.TopTile. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
								'UPGRADE_WARNING: Couldn't resolve default property of object Map.Tiles. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
								If Map.Tiles("L" & Map.TopTile(X, Y)).Key > 0 Then
									' tile has a key pattern, therefore it's a door tile!
									'UPGRADE_WARNING: Couldn't resolve default property of object Map.TopTile. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
									'UPGRADE_WARNING: Couldn't resolve default property of object Map.Tiles. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
									Map.Tiles("L" & Map.TopTile(X, Y)).Key = sByte_Renamed
								End If
							End If
							'UPGRADE_WARNING: Couldn't resolve default property of object Map.MiddleTile. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
							If Map.MiddleTile(X, Y) > 0 Then
								'UPGRADE_WARNING: Couldn't resolve default property of object Map.MiddleTile. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
								'UPGRADE_WARNING: Couldn't resolve default property of object Map.Tiles. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
								If Map.Tiles("L" & Map.MiddleTile(X, Y)).Key > 0 Then
									' tile has a key pattern, therefore it's a door tile!
									'UPGRADE_WARNING: Couldn't resolve default property of object Map.MiddleTile. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
									'UPGRADE_WARNING: Couldn't resolve default property of object Map.Tiles. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
									Map.Tiles("L" & Map.MiddleTile(X, Y)).Key = sByte_Renamed
								End If
							End If
							'UPGRADE_WARNING: Couldn't resolve default property of object Map.BottomTile. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
							If Map.BottomTile(X, Y) > 0 Then
								'UPGRADE_WARNING: Couldn't resolve default property of object Map.BottomTile. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
								'UPGRADE_WARNING: Couldn't resolve default property of object Map.Tiles. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
								If Map.Tiles("L" & Map.BottomTile(X, Y)).Key > 0 Then
									' tile has a key pattern, therefore it's a door tile!
									'UPGRADE_WARNING: Couldn't resolve default property of object Map.BottomTile. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
									'UPGRADE_WARNING: Couldn't resolve default property of object Map.Tiles. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
									Map.Tiles("L" & Map.BottomTile(X, Y)).Key = sByte_Renamed
								End If
							End If
							Found = True
						End If
						'UPGRADE_WARNING: Couldn't resolve default property of object Map.TopTile. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
						If Map.TopTile(X, Y) > 0 And Found = False Then
							'UPGRADE_WARNING: Couldn't resolve default property of object Map.TopTile. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
							'UPGRADE_WARNING: Couldn't resolve default property of object Map.Tiles. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
							If Map.Tiles("L" & Map.TopTile(X, Y)).SwapTile > 0 Then
								'UPGRADE_WARNING: Couldn't resolve default property of object Map.TopTile. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
								'UPGRADE_WARNING: Couldn't resolve default property of object Map.Tiles. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
								Map.Tiles("L" & Map.TopTile(X, Y)).Key = sByte_Renamed
								Found = True
							End If
						End If
						'UPGRADE_WARNING: Couldn't resolve default property of object Map.MiddleTile. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
						If Map.MiddleTile(X, Y) > 0 And Found = False Then
							'UPGRADE_WARNING: Couldn't resolve default property of object Map.MiddleTile. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
							'UPGRADE_WARNING: Couldn't resolve default property of object Map.Tiles. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
							If Map.Tiles("L" & Map.MiddleTile(X, Y)).SwapTile > 0 Then
								'UPGRADE_WARNING: Couldn't resolve default property of object Map.MiddleTile. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
								'UPGRADE_WARNING: Couldn't resolve default property of object Map.Tiles. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
								Map.Tiles("L" & Map.MiddleTile(X, Y)).Key = sByte_Renamed
								Found = True
							End If
						End If
						'UPGRADE_WARNING: Couldn't resolve default property of object Map.BottomTile. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
						If Map.BottomTile(X, Y) > 0 And Found = False Then
							'UPGRADE_WARNING: Couldn't resolve default property of object Map.BottomTile. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
							'UPGRADE_WARNING: Couldn't resolve default property of object Map.Tiles. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
							If Map.Tiles("L" & Map.BottomTile(X, Y)).SwapTile > 0 Then
								'UPGRADE_WARNING: Couldn't resolve default property of object Map.BottomTile. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
								'UPGRADE_WARNING: Couldn't resolve default property of object Map.Tiles. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
								Map.Tiles("L" & Map.BottomTile(X, Y)).Key = sByte_Renamed
								Found = True
							End If
						End If
					Case 16 ' Name
						' Read Only
					Case 17 ' Style
						' Read Only
					Case 18 ' DoorTileName
						' Read Only
					Case 19 ' MovementCost
						' Read Only
					Case 20 ' TerrainType
						' Read Only
				End Select
			Case 7, 8, 9, 10, 11 'EncounterNow, EncounterA, EncounterB, EncounterC, EncounterTarget
				ObjectX = FrameNow.EncounterX(IntoContext - 7)
				Select Case IntoVar
					Case 0 ' CanFight
						'UPGRADE_WARNING: Couldn't resolve default property of object ObjectX.CanFight. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
						ObjectX.CanFight = (sInteger <> 0)
					Case 1 ' CanFlee
						'UPGRADE_WARNING: Couldn't resolve default property of object ObjectX.CanFlee. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
						ObjectX.CanFlee = (sInteger <> 0)
					Case 2 ' CanTalk
						'UPGRADE_WARNING: Couldn't resolve default property of object ObjectX.CanTalk. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
						ObjectX.CanTalk = (sInteger <> 0)
					Case 3 ' ChanceToFlee
						'UPGRADE_WARNING: Couldn't resolve default property of object ObjectX.ChanceToFlee. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
						ObjectX.ChanceToFlee = sInteger
					Case 4 ' Classification
						' Read Only
					Case 5 ' FirstEntry
						'                    ObjectX.FirstEntry = (sInteger <> 0)  <-- [Titi 2.4.7] this should be a string!
						'UPGRADE_WARNING: Couldn't resolve default property of object ObjectX.FirstEntry. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
						ObjectX.FirstEntry = sText
					Case 6 ' HaveEntered
						'UPGRADE_WARNING: Couldn't resolve default property of object ObjectX.HaveEntered. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
						ObjectX.HaveEntered = (sInteger <> 0)
					Case 7 ' Not Used
					Case 8 ' IsDark
						'UPGRADE_WARNING: Couldn't resolve default property of object ObjectX.IsDark. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
						ObjectX.IsDark = (sInteger <> 0)
					Case 9 ' Name
						'UPGRADE_WARNING: Couldn't resolve default property of object ObjectX.Name. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
						ObjectX.Name = sText
					Case 10 ' Randomize: ThemeName
						'UPGRADE_WARNING: Couldn't resolve default property of object Map.Themes. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
						For	Each ThemeX In Map.Themes
							If InStr(UCase(ThemeX.Name), UCase(sText)) > 0 Then
								'UPGRADE_WARNING: Couldn't resolve default property of object ObjectX.ParentTheme. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
								ObjectX.ParentTheme = ThemeX.Index
							End If
						Next ThemeX
					Case 11 ' Randomize: AddCreatures
						'UPGRADE_WARNING: Couldn't resolve default property of object ObjectX.ReGenCreatures. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
						ObjectX.ReGenCreatures = (sInteger <> 0)
					Case 12 ' Randomize: AddItems
						'UPGRADE_WARNING: Couldn't resolve default property of object ObjectX.ReGenItems. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
						ObjectX.ReGenItems = (sInteger <> 0)
					Case 13 ' Randomize: AddTriggers
						'UPGRADE_WARNING: Couldn't resolve default property of object ObjectX.ReGenTriggers. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
						ObjectX.ReGenTriggers = (sInteger <> 0)
					Case 14 ' Randomize: Description
						'UPGRADE_WARNING: Couldn't resolve default property of object ObjectX.ReGenNewDescriptions. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
						ObjectX.ReGenNewDescriptions = (sInteger <> 0)
					Case 15 ' Randomize: GenerateUponEntry
						'UPGRADE_WARNING: Couldn't resolve default property of object ObjectX.GenerateUponEntry. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
						ObjectX.GenerateUponEntry = (sInteger <> 0)
					Case 16 ' SecondEntry
						'                    ObjectX.HaveEntered = (sInteger <> 0)  <-- [Titi 2.4.7] this should be SecondEntry!
						'UPGRADE_WARNING: Couldn't resolve default property of object ObjectX.SecondEntry. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
						ObjectX.SecondEntry = sText
					Case 17 ' ShowHint
						'UPGRADE_WARNING: Couldn't resolve default property of object ObjectX.UseHint. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
						ObjectX.UseHint = (sInteger <> 0)
					Case 18 ' CreatureCount
						' Read Only
					Case 19 ' ItemCount
						' Read Only
					Case 20 ' IsActive
						' If changing the status, then redraw the map
						'UPGRADE_WARNING: Couldn't resolve default property of object ObjectX.IsActive. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
						If ObjectX.IsActive <> (sInteger <> 0) Then
							'UPGRADE_WARNING: Couldn't resolve default property of object ObjectX.IsActive. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
							ObjectX.IsActive = (sInteger <> 0)
							frmMain.DrawMapAll()
							' If turned on where the Party stands, enter it!
							'UPGRADE_WARNING: Couldn't resolve default property of object Tome.MapY. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
							'UPGRADE_WARNING: Couldn't resolve default property of object Tome.MapX. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
							'UPGRADE_WARNING: Couldn't resolve default property of object Map.EncPointer. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
							'UPGRADE_WARNING: Couldn't resolve default property of object ObjectX.Index. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
							If ObjectX.Index = Map.EncPointer(Tome.MapX, Tome.MapY) Then
								'UPGRADE_WARNING: Couldn't resolve default property of object Tome.MapY. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
								'UPGRADE_WARNING: Couldn't resolve default property of object Tome.MapX. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
								'UPGRADE_WARNING: Couldn't resolve default property of object Map.EncPointer. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
								'UPGRADE_WARNING: Couldn't resolve default property of object ObjectX.IsActive. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
								If ObjectX.IsActive = True And Map.EncPointer(Tome.MapX, Tome.MapY) > 0 Then
									'UPGRADE_WARNING: Couldn't resolve default property of object Tome.MapY. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
									'UPGRADE_WARNING: Couldn't resolve default property of object Tome.MapX. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
									'UPGRADE_WARNING: Couldn't resolve default property of object Map.EncPointer. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
									'UPGRADE_WARNING: Couldn't resolve default property of object Map.Encounters. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
									EncounterNow = Map.Encounters("E" & Map.EncPointer(Tome.MapX, Tome.MapY))
									frmMain.EncounterEnter()
								Else
									EncounterNow = New Encounter
								End If
							End If
						End If
				End Select
			Case 12, 13, 14, 15, 16 'CreatureNow, CreatureA, CreatureB, CreatureC, CreatureTarget
				ObjectX = FrameNow.CreatureX(IntoContext - 12)
				Select Case IntoVar
					Case 0 To 7 ' BodyType 1 to 8
						' Read Only
					Case 8 To 15 ' Resistance 1 to 8
						' Read Only
					Case 16 ' Col
						'UPGRADE_WARNING: Couldn't resolve default property of object ObjectX.Col. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
						ObjectX.Col = Least(CShort(sByte_Renamed), bdCombatWidth)
					Case 17 ' Comments
						'UPGRADE_WARNING: Couldn't resolve default property of object ObjectX.Comments. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
						ObjectX.Comments = sText
					Case 18 ' Action Points
						'UPGRADE_WARNING: Couldn't resolve default property of object ObjectX.ActionPoints. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
						ObjectX.ActionPoints = sInteger
					Case 19 ' ExperiencePoints
						' Read Only
					Case 20 ' ExperienceLevel
						' Read Only
					Case 21 ' Not Used --> now Home
						' [Titi 2.4.9]
						'UPGRADE_WARNING: Couldn't resolve default property of object ObjectX.Home. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
						ObjectX.Home = sText
					Case 22 ' Greed
						'UPGRADE_WARNING: Couldn't resolve default property of object ObjectX.Greed. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
						ObjectX.Greed = sByte_Renamed
					Case 23 ' HealthNow
						'UPGRADE_WARNING: Couldn't resolve default property of object ObjectX.HPNow. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
						ObjectX.HPNow = sInteger
						frmMain.MenuDrawParty()
					Case 24 ' HealthMax
						'UPGRADE_WARNING: Couldn't resolve default property of object ObjectX.HPMax. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
						ObjectX.HPMax = sInteger
						frmMain.MenuDrawParty()
					Case 25 ' Lunacy
						'UPGRADE_WARNING: Couldn't resolve default property of object ObjectX.Lunacy. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
						ObjectX.Lunacy = sByte_Renamed
					Case 26 ' Lust
						'UPGRADE_WARNING: Couldn't resolve default property of object ObjectX.Lust. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
						ObjectX.Lust = sByte_Renamed
					Case 27 ' Not Used
					Case 28 ' IsAgressive
						'UPGRADE_WARNING: Couldn't resolve default property of object ObjectX.Agressive. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
						ObjectX.Agressive = (sInteger <> 0)
					Case 29 ' IsDMControlled
						'UPGRADE_WARNING: Couldn't resolve default property of object ObjectX.DMControlled. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
						ObjectX.DMControlled = (sInteger <> 0)
					Case 30 ' IsFriendly
						'UPGRADE_WARNING: Couldn't resolve default property of object ObjectX.Friendly. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
						ObjectX.Friendly = (sInteger <> 0)
					Case 31 ' IsGuarding
						'UPGRADE_WARNING: Couldn't resolve default property of object ObjectX.Guard. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
						ObjectX.Guard = (sInteger <> 0)
					Case 32 To 39 ' ResistanceSharp, Blunt, Cold, Fire, Evil, Holy, Magic, Mind
						' Read Only
					Case 46 ' IsTypeAnimal
						'UPGRADE_WARNING: Couldn't resolve default property of object ObjectX.Family. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
						ObjectX.Family(4) = (sInteger <> 0)
					Case 47 ' IsTypeBird
						'UPGRADE_WARNING: Couldn't resolve default property of object ObjectX.Family. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
						ObjectX.Family(6) = (sInteger <> 0)
					Case 48 ' IsTypeBlob
						'UPGRADE_WARNING: Couldn't resolve default property of object ObjectX.Family. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
						ObjectX.Family(3) = (sInteger <> 0)
					Case 49 ' IsTypeFamily
						'UPGRADE_WARNING: Couldn't resolve default property of object ObjectX.Family. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
						ObjectX.Family(8) = (sInteger <> 0)
					Case 50 ' IsTypeHuge
						'UPGRADE_WARNING: Couldn't resolve default property of object ObjectX.Size. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
						ObjectX.Size(4) = (sInteger <> 0)
					Case 51 ' IsTypeHuman
						'UPGRADE_WARNING: Couldn't resolve default property of object ObjectX.Family. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
						ObjectX.Family(0) = (sInteger <> 0)
					Case 52 ' IsTypeInsect
						'UPGRADE_WARNING: Couldn't resolve default property of object ObjectX.Family. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
						ObjectX.Family(2) = (sInteger <> 0)
					Case 53 ' IsTypeLarge
						'UPGRADE_WARNING: Couldn't resolve default property of object ObjectX.Size. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
						ObjectX.Size(3) = (sInteger <> 0)
					Case 54 ' IsTypeMagical
						'UPGRADE_WARNING: Couldn't resolve default property of object ObjectX.Family. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
						ObjectX.Family(9) = (sInteger <> 0)
					Case 55 ' IsTypeMedium
						'UPGRADE_WARNING: Couldn't resolve default property of object ObjectX.Size. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
						ObjectX.Size(2) = (sInteger <> 0)
					Case 56 ' IsTypeReptile
						'UPGRADE_WARNING: Couldn't resolve default property of object ObjectX.Family. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
						ObjectX.Family(1) = (sInteger <> 0)
					Case 57 ' IsTypeSmall
						'UPGRADE_WARNING: Couldn't resolve default property of object ObjectX.Size. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
						ObjectX.Size(1) = (sInteger <> 0)
					Case 58 ' IsTypeTiny
						'UPGRADE_WARNING: Couldn't resolve default property of object ObjectX.Size. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
						ObjectX.Size(0) = (sInteger <> 0)
					Case 59 ' IsTypeVeggie
						'UPGRADE_WARNING: Couldn't resolve default property of object ObjectX.Family. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
						ObjectX.Family(5) = (sInteger <> 0)
					Case 60 ' IsTypeUndead
						'UPGRADE_WARNING: Couldn't resolve default property of object ObjectX.Family. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
						ObjectX.Family(7) = (sInteger <> 0)
					Case 75 ' Eternium
						'UPGRADE_WARNING: Couldn't resolve default property of object ObjectX.Eternium. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
						ObjectX.Eternium = sByte_Renamed
						frmMain.MenuDrawParty()
					Case 76 ' MapX
						' Read Only
					Case 77 ' MapY
						' Read Only
					Case 78 ' Name
						'UPGRADE_WARNING: Couldn't resolve default property of object ObjectX.Name. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
						ObjectX.Name = sText
					Case 79 ' Strength
						'UPGRADE_WARNING: Couldn't resolve default property of object ObjectX.Strength. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
						ObjectX.Strength = sByte_Renamed
					Case 80 ' Pride
						'UPGRADE_WARNING: Couldn't resolve default property of object ObjectX.Pride. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
						ObjectX.Pride = sByte_Renamed
					Case 81 ' Race
						'UPGRADE_WARNING: Couldn't resolve default property of object ObjectX.Race. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
						ObjectX.Race = sText
					Case 82 ' Revelry
						'UPGRADE_WARNING: Couldn't resolve default property of object ObjectX.Revelry. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
						ObjectX.Revelry = sByte_Renamed
					Case 83 ' Row
						'UPGRADE_WARNING: Couldn't resolve default property of object ObjectX.Row. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
						ObjectX.Row = Least(CShort(sByte_Renamed), bdCombatHeight)
					Case 84 ' RuneQue1
						'UPGRADE_WARNING: Couldn't resolve default property of object ObjectX.Runes. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
						ObjectX.Runes(0) = sByte_Renamed
					Case 85 ' RuneQue2
						'UPGRADE_WARNING: Couldn't resolve default property of object ObjectX.Runes. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
						ObjectX.Runes(1) = sByte_Renamed
					Case 86 ' RuneQue3
						'UPGRADE_WARNING: Couldn't resolve default property of object ObjectX.Runes. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
						ObjectX.Runes(2) = sByte_Renamed
					Case 87 ' RuneQue4
						'UPGRADE_WARNING: Couldn't resolve default property of object ObjectX.Runes. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
						ObjectX.Runes(3) = sByte_Renamed
					Case 88 ' RuneQue5
						'UPGRADE_WARNING: Couldn't resolve default property of object ObjectX.Runes. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
						ObjectX.Runes(4) = sByte_Renamed
					Case 89 ' RuneQue6
						'UPGRADE_WARNING: Couldn't resolve default property of object ObjectX.Runes. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
						ObjectX.Runes(5) = sByte_Renamed
					Case 90 ' ScarletLetter
					Case 91 ' Agility
						'UPGRADE_WARNING: Couldn't resolve default property of object ObjectX.Agility. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
						ObjectX.Agility = sByte_Renamed
					Case 92 ' SkillPoints
						'UPGRADE_WARNING: Couldn't resolve default property of object ObjectX.SkillPoints. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
						ObjectX.SkillPoints = sInteger
					Case 93 ' Status
						' Read Only
					Case 94 ' Not Used
						' Not Used
					Case 95 ' Attack
						' Read Only
					Case 96 ' Will
						'UPGRADE_WARNING: Couldn't resolve default property of object ObjectX.Will. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
						ObjectX.Will = sByte_Renamed
					Case 97 ' Wrath
						'UPGRADE_WARNING: Couldn't resolve default property of object ObjectX.Wrath. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
						ObjectX.Wrath = sByte_Renamed
					Case 98 ' Male
						'UPGRADE_WARNING: Couldn't resolve default property of object ObjectX.Male. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
						ObjectX.Male = (sInteger <> 0)
					Case 99 ' Unconscious
						'UPGRADE_WARNING: Couldn't resolve default property of object ObjectX.Unconscious. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
						ObjectX.Unconscious = (sInteger <> 0)
					Case 100 ' Frozen
						'UPGRADE_WARNING: Couldn't resolve default property of object ObjectX.Frozen. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
						ObjectX.Frozen = (sInteger <> 0)
					Case 101 ' RangeToTarget
						' Read Only
					Case 102 ' ActionPointsMax
						' Read Only
					Case 103 To 110 ' ResistanceBonusSharp, Blunt, Cold, Fire, Evil, Holy, Magic, Mind
						'UPGRADE_WARNING: Couldn't resolve default property of object ObjectX.ResistanceBonus. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
						ObjectX.ResistanceBonus(IntoVar - 102) = sInteger
					Case 111 ' ResistanceBonus (everthing but the specials above)
						'UPGRADE_WARNING: Couldn't resolve default property of object ObjectX.ResistanceBonus. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
						ObjectX.ResistanceBonus(0) = sInteger
					Case 133 ' EterniumMax
						' Read Only
					Case 134 ' WeightMax
						' Read Only
					Case 135 ' Weight
						' Read Only
					Case 136 ' Bulk
						' Read Only
					Case 137 ' StrengthBonus
						'UPGRADE_WARNING: Couldn't resolve default property of object ObjectX.StrengthBonus. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
						ObjectX.StrengthBonus = sInteger
					Case 138 ' AgilityBonus
						'UPGRADE_WARNING: Couldn't resolve default property of object ObjectX.AgilityBonus. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
						ObjectX.AgilityBonus = sInteger
					Case 139 ' WillBonus
						'UPGRADE_WARNING: Couldn't resolve default property of object ObjectX.WillBonus. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
						ObjectX.WillBonus = sInteger
					Case 140 ' AttackBonus
						'UPGRADE_WARNING: Couldn't resolve default property of object ObjectX.AttackBonus. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
						ObjectX.AttackBonus = sInteger
					Case 141 ' DamageBonus
						'UPGRADE_WARNING: Couldn't resolve default property of object ObjectX.DamageBonus. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
						ObjectX.DamageBonus = sInteger
					Case 142 ' DefenseBonus
						'UPGRADE_WARNING: Couldn't resolve default property of object ObjectX.DefenseBonus. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
						ObjectX.DefenseBonus = sInteger
					Case 143 ' ArmorBonus
						'UPGRADE_WARNING: Couldn't resolve default property of object ObjectX.ArmorBonus. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
						ObjectX.ArmorBonus = sInteger
					Case 144 ' IsSpellCaster
						' Read Only
					Case 145 ' Money
						'UPGRADE_WARNING: Couldn't resolve default property of object ObjectX.Money. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
						ObjectX.Money = CInt(sInteger)
					Case 146 ' ActionPointsMax
						' Read Only
					Case 147 ' IsAfraid
						'UPGRADE_WARNING: Couldn't resolve default property of object ObjectX.Afraid. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
						ObjectX.Afraid = (sInteger <> 0)
					Case 148 ' IsInanimate
						'UPGRADE_WARNING: Couldn't resolve default property of object ObjectX.IsInanimate. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
						ObjectX.IsInanimate = (sInteger <> 0)
					Case 149 ' Not used
					Case 150 ' CombatRank
						'UPGRADE_WARNING: Couldn't resolve default property of object ObjectX.CombatRank. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
						ObjectX.CombatRank = Greatest(Least(sInteger, 3), 0)
					Case 151 ' Pronoun_HeShe
						' Read Only
					Case 152 ' Pronoun_HimHer
						' Read Only
					Case 153 ' Pronoun_HisHer
						' Read Only
					Case 154 ' MovementCost
						' Read Only
					Case 155 ' MovementCostBonus
						'UPGRADE_WARNING: Couldn't resolve default property of object ObjectX.MovementCostBonus. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
						ObjectX.MovementCostBonus = sInteger
					Case 156 ' Index
						' Read Only
					Case 157 ' PictureFile
						' Validate Picture
						Found = True
						'UPGRADE_WARNING: Couldn't resolve default property of object Tome.LoadPath. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
						'UPGRADE_WARNING: Dir has a new behavior. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="9B7D5ADD-D8FE-4819-A36C-6DEDAF088CC7"'
						If Dir(Tome.LoadPath & "\" & sText) = "" Then
							'UPGRADE_WARNING: Couldn't resolve default property of object Tome.LoadPath. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
							'UPGRADE_WARNING: Dir has a new behavior. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="9B7D5ADD-D8FE-4819-A36C-6DEDAF088CC7"'
							If Dir(Tome.LoadPath & "\creatures\" & sText) = "" Then
								'                            If Dir(gAppPath & "\data\graphics\creatures\" & sText) = "" Then
								'UPGRADE_WARNING: Dir has a new behavior. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="9B7D5ADD-D8FE-4819-A36C-6DEDAF088CC7"'
								If Dir(gDataPath & "\graphics\creatures\" & sText) = "" Then
									Found = False
								End If
							End If
						End If
						If Found = True Then
							'UPGRADE_WARNING: Couldn't resolve default property of object ObjectX.PictureFile. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
							ObjectX.PictureFile = sText
							'UPGRADE_WARNING: Couldn't resolve default property of object ObjectX. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
							frmMain.LoadCreaturePic(ObjectX)
							If frmMain.picGrid.Visible = True Then
								frmMain.CombatDraw()
							Else
								frmMain.MenuDrawParty()
								frmMain.DrawMapAll()
							End If
						End If
					Case 158 ' FaceLeft
						'UPGRADE_WARNING: Couldn't resolve default property of object ObjectX.FaceLeft. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
						ObjectX.FaceLeft = sInteger
					Case 159 ' FaceTop
						'UPGRADE_WARNING: Couldn't resolve default property of object ObjectX.FaceTop. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
						ObjectX.FaceTop = sInteger
					Case 160 ' Initiative
						'UPGRADE_WARNING: Couldn't resolve default property of object ObjectX.Initiative. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
						ObjectX.Initiative = sInteger
					Case 161 ' Facing
						'UPGRADE_WARNING: Couldn't resolve default property of object ObjectX.Facing. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
						ObjectX.Facing = System.Math.Abs(System.Math.Sign(sInteger > 0))
				End Select
			Case 17, 18, 19, 20, 21 'ItemNow, ItemA, ItemB, ItemC, ItemTarget
				ObjectX = FrameNow.ItemX(IntoContext - 17)
				Select Case IntoVar
					Case 0 ' Resistance
						'UPGRADE_WARNING: Couldn't resolve default property of object ObjectX.Resistance. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
						ObjectX.Resistance = sInteger
					Case 1 ' Bulk
						'UPGRADE_WARNING: Couldn't resolve default property of object ObjectX.Bulk. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
						ObjectX.Bulk = sInteger
					Case 2 ' CanCombine
						'UPGRADE_WARNING: Couldn't resolve default property of object ObjectX.CanCombine. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
						ObjectX.CanCombine = (sInteger <> 0)
					Case 3 ' ArmorType
						' Read Only
					Case 16 ' Capacity
						'UPGRADE_WARNING: Couldn't resolve default property of object ObjectX.Capacity. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
						ObjectX.Capacity = sInteger
					Case 18 ' Comments
						'UPGRADE_WARNING: Couldn't resolve default property of object ObjectX.Comments. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
						ObjectX.Comments = sText
					Case 19 ' Count
						'UPGRADE_WARNING: Couldn't resolve default property of object ObjectX.Count. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
						ObjectX.Count = sInteger
					Case 20 ' DamageDice
						'UPGRADE_WARNING: Couldn't resolve default property of object ObjectX.Damage. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
						ObjectX.Damage = sInteger
					Case 21 ' DamageBonus
						'UPGRADE_WARNING: Couldn't resolve default property of object ObjectX.DamageBonus. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
						ObjectX.DamageBonus = sInteger
					Case 22 ' IsWeaponMelee
						' Read Only
					Case 23 ' IsWeaponRanged
						' Read Only
					Case 24 ' IsEquipped
						'UPGRADE_WARNING: Couldn't resolve default property of object ObjectX.IsReady. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
						ObjectX.IsReady = (sInteger <> 0)
						If frmMain.picInventory.Visible = True Then
							frmMain.InventoryShow(0) ' Redisplay Items in Inventory
						End If
					Case 25 ' IsWeaponThrown
						' Read Only
					Case 26 ' IsKey
						'UPGRADE_WARNING: Couldn't resolve default property of object ObjectX.KeyBits. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
						ObjectX.KeyBits = sByte_Renamed
					Case 27 ' ProtectionBonusType
						'UPGRADE_WARNING: Couldn't resolve default property of object ObjectX.ResistanceType. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
						ObjectX.ResistanceType = sByte_Renamed
					Case 28 ' ProtectionBonus
						'UPGRADE_WARNING: Couldn't resolve default property of object ObjectX.ResistanceBonus. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
						ObjectX.ResistanceBonus = sInteger
					Case 31 ' IsRangeLong
						'UPGRADE_WARNING: Couldn't resolve default property of object ObjectX.RangeType. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
						ObjectX.RangeType = 2
					Case 32 ' IsRangeMedium
						'UPGRADE_WARNING: Couldn't resolve default property of object ObjectX.RangeType. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
						ObjectX.RangeType = 1
					Case 33 ' IsRangeShort
						'UPGRADE_WARNING: Couldn't resolve default property of object ObjectX.RangeType. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
						ObjectX.RangeType = 0
					Case 36 ' IsSoftBulk
						'UPGRADE_WARNING: Couldn't resolve default property of object ObjectX.SoftCapacity. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
						ObjectX.SoftCapacity = (sInteger <> 0)
					Case 37 ' IsDamageTypeSharp
						'UPGRADE_WARNING: Couldn't resolve default property of object ObjectX.DamageType. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
						ObjectX.DamageType = 1
					Case 38 ' IsDamageTypeBlunt
						'UPGRADE_WARNING: Couldn't resolve default property of object ObjectX.DamageType. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
						ObjectX.DamageType = 2
					Case 39 ' IsDamageTypeCold
						'UPGRADE_WARNING: Couldn't resolve default property of object ObjectX.DamageType. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
						ObjectX.DamageType = 3
					Case 40 ' IsDamageTypeFire
						'UPGRADE_WARNING: Couldn't resolve default property of object ObjectX.DamageType. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
						ObjectX.DamageType = 4
					Case 41 ' IsDamageTypeHoly
						'UPGRADE_WARNING: Couldn't resolve default property of object ObjectX.DamageType. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
						ObjectX.DamageType = 6
					Case 42 ' IsDamageTypeEvil
						'UPGRADE_WARNING: Couldn't resolve default property of object ObjectX.DamageType. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
						ObjectX.DamageType = 5
					Case 43 ' IsDamageTypeMagic
						'UPGRADE_WARNING: Couldn't resolve default property of object ObjectX.DamageType. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
						ObjectX.DamageType = 7
					Case 44 ' IsDamageTypeMind
						'UPGRADE_WARNING: Couldn't resolve default property of object ObjectX.DamageType. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
						ObjectX.DamageType = 8
					Case 50 ' MapX
						' Read Only
					Case 51 ' MapY
						' Read Only
					Case 52 ' Name
						'UPGRADE_WARNING: Couldn't resolve default property of object ObjectX.Name. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
						ObjectX.Name = sText
					Case 53 ' Defense
						'UPGRADE_WARNING: Couldn't resolve default property of object ObjectX.Defense. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
						ObjectX.Defense = sInteger
					Case 54 ' Not Used
					Case 55 ' AttackBonus
						'UPGRADE_WARNING: Couldn't resolve default property of object ObjectX.AttackBonus. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
						ObjectX.AttackBonus = sInteger
					Case 56 ' UseAsDescription
						'UPGRADE_WARNING: Couldn't resolve default property of object ObjectX.UseDescription. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
						ObjectX.UseDescription = (sInteger <> 0)
					Case 57 ' Value
						'UPGRADE_WARNING: Couldn't resolve default property of object ObjectX.Value. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
						ObjectX.Value = sInteger
					Case 58 ' Weight
						'UPGRADE_WARNING: Couldn't resolve default property of object ObjectX.Weight. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
						ObjectX.Weight = sInteger
					Case 59 ' CanWearTwoHanded
						'UPGRADE_WARNING: Couldn't resolve default property of object ObjectX.WearType. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
						ObjectX.WearType(13) = (sInteger <> 0)
					Case 60 ' IsInHand
						' Read only
					Case 61 ' IsAmmoArrow
						'UPGRADE_WARNING: Couldn't resolve default property of object ObjectX.ShootType. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
						ObjectX.ShootType = sByte_Renamed
					Case 64 ' Family
						If Len(sText) > 1 Then
							X = 0
							Do Until X = 19 ' modBD.SetUpItemFamily(x) <> ""
								If InStr(UCase(modBD.SetUpItemFamily(X)), UCase(sText)) > 0 Then
									'UPGRADE_WARNING: Couldn't resolve default property of object ObjectX.Family. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
									ObjectX.Family = X
									Exit Do
								End If
								X = X + 1 ' missing? Added in 2.4.8 [Titi]
							Loop 
						End If
					Case 65 ' ActionPoints
						'UPGRADE_WARNING: Couldn't resolve default property of object ObjectX.ActionPoints. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
						ObjectX.ActionPoints = sInteger
				End Select
			Case 22, 23, 24, 25, 26 'TriggerNow, TriggerA, TriggerB, TriggerC, TriggerTarget
				ObjectX = FrameNow.TriggerX(IntoContext - 22)
				Select Case IntoVar
					Case 0 'ByteA
						'UPGRADE_WARNING: Couldn't resolve default property of object ObjectX.VarA. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
						ObjectX.VarA = sByte_Renamed
					Case 1 'ByteB
						'UPGRADE_WARNING: Couldn't resolve default property of object ObjectX.VarB. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
						ObjectX.VarB = sByte_Renamed
					Case 2 'ByteC
						'UPGRADE_WARNING: Couldn't resolve default property of object ObjectX.VarC. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
						ObjectX.VarC = sByte_Renamed
					Case 3 'Comments
						'UPGRADE_WARNING: Couldn't resolve default property of object ObjectX.Comments. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
						ObjectX.Comments = sText
					Case 4 'Not Used
					Case 5 'IsTimed
						'UPGRADE_WARNING: Couldn't resolve default property of object ObjectX.IsTimed. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
						ObjectX.IsTimed = (sInteger <> 0)
					Case 6 'IsCurse
						'UPGRADE_WARNING: Couldn't resolve default property of object ObjectX.IsCurse. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
						ObjectX.IsCurse = (sInteger <> 0)
					Case 7 'IsEvil
						'UPGRADE_WARNING: Couldn't resolve default property of object ObjectX.IsEvil. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
						ObjectX.IsEvil = (sInteger <> 0)
					Case 8 'IsFear
						'UPGRADE_WARNING: Couldn't resolve default property of object ObjectX.IsFear. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
						ObjectX.IsFear = (sInteger <> 0)
					Case 9 'IsFish
						'UPGRADE_WARNING: Couldn't resolve default property of object ObjectX.IsFish. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
						ObjectX.IsFish = (sInteger <> 0)
					Case 10 'IsGreed
						'UPGRADE_WARNING: Couldn't resolve default property of object ObjectX.IsGreed. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
						ObjectX.IsGreed = (sInteger <> 0)
					Case 11 'IsLunacy
						'UPGRADE_WARNING: Couldn't resolve default property of object ObjectX.IsLunacy. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
						ObjectX.IsLunacy = (sInteger <> 0)
					Case 12 'IsLust
						'UPGRADE_WARNING: Couldn't resolve default property of object ObjectX.IsLust. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
						ObjectX.IsLust = (sInteger <> 0)
					Case 13 'IsPoison
						'UPGRADE_WARNING: Couldn't resolve default property of object ObjectX.IsPoison. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
						ObjectX.IsPoison = (sInteger <> 0)
					Case 14 'IsPride
						'UPGRADE_WARNING: Couldn't resolve default property of object ObjectX.IsPride. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
						ObjectX.IsPride = (sInteger <> 0)
					Case 15 'IsMagic
						'UPGRADE_WARNING: Couldn't resolve default property of object ObjectX.IsMagic. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
						ObjectX.IsMagic = (sInteger <> 0)
					Case 16 'IsSkill
						'UPGRADE_WARNING: Couldn't resolve default property of object ObjectX.IsSkill. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
						ObjectX.IsSkill = (sInteger <> 0)
					Case 17 'IsRevelry
						'UPGRADE_WARNING: Couldn't resolve default property of object ObjectX.IsRevelry. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
						ObjectX.IsRevelry = (sInteger <> 0)
					Case 18 'IsTrap
						'UPGRADE_WARNING: Couldn't resolve default property of object ObjectX.IsTrap. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
						ObjectX.IsTrap = (sInteger <> 0)
					Case 19 'IsWrath
						'UPGRADE_WARNING: Couldn't resolve default property of object ObjectX.IsWrath. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
						ObjectX.IsWrath = (sInteger <> 0)
					Case 20 'Name
						'UPGRADE_WARNING: Couldn't resolve default property of object ObjectX.Name. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
						ObjectX.Name = sText
					Case 21 'SkillPoints
						'UPGRADE_WARNING: Couldn't resolve default property of object ObjectX.SkillPoints. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
						ObjectX.SkillPoints = sInteger
					Case 22 'TriggerType
						' Read Only
					Case 23 ' Turns
						'UPGRADE_WARNING: Couldn't resolve default property of object ObjectX.Turns. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
						ObjectX.Turns = sByte_Renamed
					Case 24 ' SkillLevel
						' ReadOnly
				End Select
			Case 27 'Local
				Select Case IntoVar
					Case 0 'Abort
						FrameNow.Abort = sInteger
					Case 1 'ByteA
						FrameNow.ByteA = sByte_Renamed
					Case 2 'ByteB
						FrameNow.ByteB = sByte_Renamed
					Case 3 'ByteC
						FrameNow.ByteC = sByte_Renamed
					Case 4 'Fail
						FrameNow.Fail = sInteger
					Case 5 'IntegerA
						FrameNow.IntegerA = sInteger
					Case 6 'IntegerB
						FrameNow.IntegerB = sInteger
					Case 7 'IntegerC
						FrameNow.IntegerC = sInteger
					Case 8 'StopExit
						FrameNow.StopExit = sInteger
					Case 9 'TextA
						FrameNow.TextA = sText
					Case 10 'TextB
						FrameNow.TextB = sText
					Case 11 'TextC
						FrameNow.TextC = sText
					Case 12 'Random
						FrameNow.Random = sInteger
					Case 13 'FoundIt
						FrameNow.FoundIt = sInteger
				End Select
			Case 28 'Global
				Select Case IntoVar
					Case 0 ' False
						' Read Only
					Case 1 ' True
						' Read Only
					Case 2 ' InCombat
						' Read Only
					Case 3 ' DieTypeRoll
						' Read Only
					Case 4 ' DieCountRoll
						' Read Only
					Case 5 ' ArmorRoll
						GlobalArmorRoll = Least(Greatest(sInteger, 1), 8)
					Case 6 ' HitRoll
						GlobalAttackRoll = sInteger
					Case 7 ' DamageRoll
						sInteger = sInteger
						GlobalDamageRoll = sInteger
					Case 8 ' Offer
						' Read Only
					Case 9 To 28 ' RunePool
						frmMain.LetRunePool(IntoVar - 9, sInteger)
					Case 29 ' SkillLevel
						GlobalSkillLevel = sInteger
					Case 58 ' IntegerA
						GlobalIntegerA = sInteger
					Case 59 ' IntegerB
						GlobalIntegerB = sInteger
					Case 60 ' IntegerC
						GlobalIntegerC = sInteger
					Case 61 ' TextA
						GlobalTextA = sText
					Case 62 ' TextB
						GlobalTextB = sText
					Case 63 ' TextC
						GlobalTextC = sText
					Case 64 ' DayName
						GlobalDayName = sText
					Case 65 ' MoonName
						GlobalMoonName = sText
					Case 66 ' YearName
						GlobalYearName = sText
					Case 67 ' TurnName
						GlobalTurnName = sText
					Case 68 ' PickLockChance
						GlobalPickLockChance = sInteger
					Case 69 ' Not Used
					Case 70 ' RemoveTrapChance
						GlobalRemoveTrapChance = sInteger
					Case 71 ' HitLocation
						' Read Only
					Case 71 ' HitLocation
						GlobalHitLocation = CStr(sInteger)
					Case 72 ' SpellFizzleChance
						GlobalSpellFizzleChance = sInteger
					Case 73 ' Ticks
						GlobalTicks = sInteger
				End Select
			Case 29 'Pos
				' Read only at run-time
			Case 30 'Neg
				' Read only at run-time
			Case 31 'Dice
				' Read only at run-time
			Case 32 'Random
				' Read only at run-time
		End Select
	End Sub
End Module