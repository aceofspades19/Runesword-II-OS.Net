Option Strict Off
Option Explicit On
Module modDebug
	
	' Debug Variables
	Public DebugQue As New Collection
	Public DebugHeader As String
	Dim DebugTop As Short
	Dim DebugThumbY As Short
	
	
	Public Sub DebugAdd(ByRef TrigX As Trigger, ByRef StmtX As Statement)
		Dim c As Short
		Dim FactoidX As Factoid
		Dim Text, Text2 As String
		' Add Statement to Que
		c = 1
		For	Each FactoidX In DebugQue
			If FactoidX.Index >= c Then
				c = FactoidX.Index + 1
			End If
		Next FactoidX
		' Set the index and add the door. Return the new door's index.
		FactoidX = New Factoid
		FactoidX.Index = c
		DebugQue.Add(FactoidX, "F" & FactoidX.Index)
		' Show Statement
		modBD.StatementToText(Tome, Area, TrigX, StmtX, Text)
		FactoidX.Text = Text
		' Decrypt Statement
		'    Select Case StmtX.Statement
		'        Case 0, 3, 7, 10, 23 ' <None>, Else, EndIf, Next, Runes
		'        Case 1, 11, 14, 15 ' Label, Branch, Select
		'            Text = GetVarContext(StmtX.B(0), StmtX.B(1))
		'        Case 2, 4, 5, 6, 8, 12 ' If, ElseIf, And, Or, While, Put
		'            Select Case StmtX.B(2)
		'                Case 0
		'                    Text = "Equals"
		'                Case 1
		'                    Text = "+"
		'                Case 2
		'                    Text = "-"
		'                Case 3
		'                    Text = "*"
		'                Case 4
		'                    Text = "Divided By"
		'                Case 5
		'                    Text = ">"
		'                Case 6
		'                    Text = "<"
		'                Case 7
		'                    Text = "> Or Equal"
		'                Case 8
		'                    Text = "< Or Equal"
		'                Case 9
		'                    Text = "Or"
		'                Case 10
		'                    Text = "And"
		'                Case 11
		'                    Text = "Xor"
		'                Case 12
		'                    Text = "Like"
		'            End Select
		'            Text = GetVarContext(StmtX.B(0), StmtX.B(1)) & " " & Text & " " & GetVarContext(StmtX.B(3), StmtX.B(4))
		'            If StmtX.Statement = 12 Then
		'                Text = Text & " Into " & ContextToText(StmtX.B(5)) & "." & VarToText(StmtX.B(5), StmtX.B(6))
		'            End If
	End Sub
	
	Public Sub DebugShow()
		Dim Found, OldMouse As Short
		If GlobalDebugMode = 1 Then
			' Setup with List
			'DialogSetUp bdDlgDebug
			DialogSetUp(modGameGeneral.DLGTYPE.bdDlgDebug)
			DialogSetButton(1, "Done")
			DialogSetButton(2, "Halt")
			DebugTop = 1
			DebugList()
			DialogShow("DM", "")
			DialogHide()
			If ConvoAction = 0 Then ' Halt, Turn Off Debug
				GlobalDebugMode = 0
			End If
			DebugQue = New Collection
		End If
	End Sub
	
	Public Sub DebugList()
		Dim n, c As Short
		Dim rc As Integer
		Dim FactoidX As Factoid
		n = 0
		'UPGRADE_ISSUE: PictureBox method picConvoList.Cls was not upgraded. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="CC4C7EC0-C903-48FC-ACCC-81861D12DA4A"'
		frmMain.picConvoList.Cls()
		For c = 0 To Least(18, DebugQue.Count() - DebugTop)
			' Show Description
			'ShowText frmMain.picConvoList, 8, 6 + 14 * n, 354, 10, bdFontSmallWhite, DebugQue(DebugTop + c).Text, False, False
			'UPGRADE_WARNING: Couldn't resolve default property of object DebugQue().Text. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
			frmMain.ShowText((frmMain.picConvoList), 8, 6 + 14 * n, 354, 10, bdFontSmallWhite, DebugQue.Item(DebugTop + c).Text, False, False)
			n = n + 1
		Next c
		'ScrollBarShow frmMain.picConvoList, 331, 2, 269, DebugTop, DebugQue.Count - 18, 0
		frmMain.ScrollBarShow((frmMain.picConvoList), 331, 2, 269, DebugTop, DebugQue.Count() - 18, 0)
	End Sub
	
	Public Sub DebugClick(ByRef AtX As Short, ByRef AtY As Short, ByRef ButtonDown As Short)
		Dim c As Short
		Dim rc As Integer
		If PointIn(AtX, AtY, 331, 2, 18, 269) Then
			' ScrollBar Click
			'If ScrollBarClick(AtX, AtY, ButtonDown, frmMain.picConvoList, 331, 2, 269, DebugTop, DebugQue.Count, 18) = True Then
			If frmMain.ScrollBarClick(AtX, AtY, ButtonDown, (frmMain.picConvoList), 331, 2, 269, DebugTop, DebugQue.Count(), 18) = True Then
				DebugList()
			End If
		End If
	End Sub
End Module