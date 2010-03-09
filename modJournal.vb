Option Strict Off
Option Explicit On
Module modJournal
	
	Public JournalX, JournalDragDrop, JournalY As Short
	Public JournalMode As Short
	Public MicroMapLeft, MicroMapTop As Short
	
	Public Const bdMicroSize As Double = 0.25
	Public Const bdMicroSizeX As Short = 52
	Public Const bdMicroSizeY As Short = 76
	Public Const bdMicroSizeWidth As Short = 295
	Public Const bdMicroSizeHeight As Short = 228
	Public Const bdTileWidth As Short = 24
	Public Const bdTileHeight As Short = 18
	Public Const bdTileBlack As Short = -1
	Public Const bdTileGrey As Short = -2
	Public Const bdTileDark As Short = -3
	Public Const bdMapPartySize As Double = 0.3
	
	Public Const bdSideLeft As Short = 1
	Public Const bdSideTop As Short = 2
	Public Const bdSideRight As Short = 4
	Public Const bdSideFirstBottom As Short = 8
	Public Const bdSideSecondBottom As Short = 16
	
	Private bAddMode As Boolean
    Private blnDeleteButton As Boolean
    Private tome As Tome = Tome.getInstance()
	
	'UPGRADE_NOTE: Tome was upgraded to Tome_Renamed. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="A9E4979A-37FA-4718-9994-97DD76ED70A7"'
	Public Sub JournalClick(ByRef AtX As Short, ByRef AtY As Short, ByRef ButtonDown As Short)
	
		Dim JournalX As Journal

		If bAddMode Then
			If PointIn(AtX, AtY, 268, 350, 90, 18) Then
				' Accept
				frmMain.ShowButton((frmMain.picJournal), 268, 350, "Accept", ButtonDown)
				frmMain.picJournal.Refresh()
				If ButtonDown = False Then
					Call PlayClickSnd(modIOFunc.ClickType.ifClick)
					bAddMode = False
					'UPGRADE_WARNING: Couldn't resolve default property of object JournalList(JournalTop).Text. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
					If JournalList.Item(JournalTop).Text <> vbNullString Then
						'UPGRADE_WARNING: Couldn't resolve default property of object Tome.AddJournal. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
						'UPGRADE_WARNING: Couldn't resolve default property of object JournalList().Text. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
                        tome.AddJournal.Text = JournalList.Item(JournalTop).Text
					Else
						If JournalList.Count() = 1 Then
							'UPGRADE_WARNING: Couldn't resolve default property of object JournalList().Text. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
							JournalList.Item(JournalTop).Text = "No Journal Entries."
						Else
							JournalList.Remove(JournalTop)
							JournalTop = JournalTop - 1
						End If
					End If
					JournalShow()
				End If
			ElseIf PointIn(AtX, AtY, 168, 350, 90, 18) Then 
				' Cancel
				frmMain.ShowButton((frmMain.picJournal), 268, 350, "Cancel", ButtonDown)
				frmMain.picJournal.Refresh()
				If Not ButtonDown Then
					Call PlayClickSnd(modIOFunc.ClickType.ifClick)
					bAddMode = False
					If JournalList.Count() = 1 Then ' And JournalList(JournalTop).Text = "" Then
						' [Titi 2.4.7] to prevent RT5 if cancelling first entry
						'UPGRADE_WARNING: Couldn't resolve default property of object JournalList().Text. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
						JournalList.Item(JournalTop).Text = "No Journal Entries."
					Else
						JournalList.Remove(JournalTop)
						JournalTop = JournalTop - 1
					End If
					JournalShow()
				End If
			End If
		Else ' If not in add new journal entry mode
			If PointIn(AtX, AtY, 268, 350, 90, 18) Then
				' Done
				frmMain.ShowButton((frmMain.picJournal), 268, 350, "Done", ButtonDown)
				frmMain.picJournal.Refresh()
				If ButtonDown = False Then
					Call PlayClickSnd(modIOFunc.ClickType.ifClick)
                    frmMain.TomeSaveArea(TomeSavePathName)
					frmMain.picJournal.Visible = False
					frmMain.picMap.Focus()
				End If
			ElseIf PointIn(AtX, AtY, 168, 350, 90, 18) Then 
				' Add Journal Entry
				frmMain.ShowButton((frmMain.picJournal), 168, 350, "Add", ButtonDown)
				frmMain.picJournal.Refresh()
				If ButtonDown = False Then
					Call PlayClickSnd(modIOFunc.ClickType.ifClick)
					bAddMode = True
					'UPGRADE_WARNING: Couldn't resolve default property of object JournalList(JournalTop).Text. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
					If JournalList.Item(JournalTop).Text <> "No Journal Entries." Then
						JournalX = New Journal
						JournalX.Text = ""
						JournalList.Add(JournalX)
						JournalTop = JournalList.Count()
					Else
						'UPGRADE_WARNING: Couldn't resolve default property of object JournalList().Text. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
						JournalList.Item(JournalTop).Text = ""
					End If
					JournalShow()
				End If
			ElseIf PointIn(AtX, AtY, 68, 350, 90, 18) Then 
				' Delete Journal Entry
				If blnDeleteButton Then
					' will run only if the button exists!
					frmMain.ShowButton((frmMain.picJournal), 268, 350, "Delete", ButtonDown)
					frmMain.picJournal.Refresh()
					If Not ButtonDown Then
						Call PlayClickSnd(modIOFunc.ClickType.ifClick)
						' [Titi] 2.4.6
						'            intJournal = 0
						' with the new add/delete functionalities, JournalX.Index may not be the current index
                        For Each JournalX In tome.Journals
                            '                intJournal = intJournal + 1
                            'UPGRADE_WARNING: Couldn't resolve default property of object JournalList(JournalTop).Text. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
                            If JournalX.Text = JournalList.Item(JournalTop).Text Then Exit For
                        Next JournalX
						'UPGRADE_WARNING: Couldn't resolve default property of object Tome.RemoveJournal. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
                        tome.RemoveJournal("J" & JournalX.Index) ' intJournal
						' [Titi 2.4.8] reverted to the "J" & index tagging for compatibility with creator
						'UPGRADE_WARNING: Couldn't resolve default property of object Tome.Journals. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
                        If tome.Journals.Count = CountQuests() Then ' all entries are quests
                            'UPGRADE_WARNING: Couldn't resolve default property of object JournalList().Text. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
                            JournalList.Item(JournalTop).Text = "No Journal Entries."
                        Else
                            JournalList.Remove(JournalTop)
                            JournalTop = JournalTop - 1
                            If JournalTop = 0 Then JournalTop = 1
                        End If
						JournalShow()
					End If
				End If ' Delete Button
			ElseIf PointIn(AtX, AtY, 36, 42, 102, 18) And ButtonDown = True Then 
				' Journal
				JournalMode = 0
				bAddMode = False
				Call PlayClickSnd(modIOFunc.ClickType.ifClick)
				JournalEntryLoad()
				JournalShow()
			ElseIf PointIn(AtX, AtY, 138, 42, 102, 18) And ButtonDown = True Then 
				' Quests
				JournalMode = 1
				Call PlayClickSnd(modIOFunc.ClickType.ifClick)
				JournalQuestsLoad()
				JournalShow()
			ElseIf PointIn(AtX, AtY, 240, 42, 102, 18) And ButtonDown = True Then 
				' Map
				JournalMode = 2
				Call PlayClickSnd(modIOFunc.ClickType.ifClick)
				JournalShow()
			ElseIf PointIn(AtX, AtY, 317, 81, 18, 218) And Not bAddMode Then 
				'UPGRADE_ISSUE: frmMain.picJournal was upgraded to a Panel, and cannot be coerced to a PictureBox. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="0FF1188E-27C0-4FED-842D-159C65894C9B"'
				If frmMain.ScrollBarClick(AtX, AtY, ButtonDown, (frmMain.picJournal), 317, 81, 218, JournalTop, JournalList.Count() + 1, 1) = True Then
					JournalShow()
				End If
			End If
		End If
	End Sub
	
	Public Sub JournalShow()
		Dim rc As Integer
		Dim c As Short

		blnDeleteButton = False
		'UPGRADE_ISSUE: PictureBox method picJournal.Cls was not upgraded. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="CC4C7EC0-C903-48FC-ACCC-81861D12DA4A"'
        frmMain.picJournal = Nothing
        frmMain.picJournal.Invalidate()
		frmMain.ShowText((frmMain.picJournal), 0, 12, frmMain.picJournal.Width, 14, bdFontElixirWhite, "Journal Entries", True, False)
		If bAddMode Then
			frmMain.ShowText((frmMain.picJournal), 0, 45, frmMain.picJournal.Width, 14, bdFontNoxiousBlack, "New Journal Entry", -1, False)
			'UPGRADE_WARNING: Couldn't resolve default property of object JournalList().Text. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
			frmMain.ShowText((frmMain.picJournal), 36, 81, 276, 220, bdFontNoxiousBlack, JournalList.Item(JournalTop).Text & "\", False, False)
			frmMain.ShowButton((frmMain.picJournal), 168, 350, "Cancel", False)
			frmMain.ShowButton((frmMain.picJournal), 268, 350, "Accept", False)
			'UPGRADE_ISSUE: frmMain.picJournal was upgraded to a Panel, and cannot be coerced to a PictureBox. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="0FF1188E-27C0-4FED-842D-159C65894C9B"'
			frmMain.ScrollBarShow((frmMain.picJournal), 317, 81, 218, JournalTop, JournalList.Count(), 0)
		Else
			For c = 0 To 2
				If c = JournalMode Then
					'UPGRADE_ISSUE: PictureBox property picMisc.hdc was not upgraded. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="CC4C7EC0-C903-48FC-ACCC-81861D12DA4A"'
					'UPGRADE_ISSUE: PictureBox property picJournal.hdc was not upgraded. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="CC4C7EC0-C903-48FC-ACCC-81861D12DA4A"'
                    rc = BBlt(frmMain.picJournal, frmMain.picMisc, 48 + c * 102, 42, 18, 18, 18, 18, SRCCOPY)
				Else
					'UPGRADE_ISSUE: PictureBox property picMisc.hdc was not upgraded. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="CC4C7EC0-C903-48FC-ACCC-81861D12DA4A"'
					'UPGRADE_ISSUE: PictureBox property picJournal.hdc was not upgraded. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="CC4C7EC0-C903-48FC-ACCC-81861D12DA4A"'
                    rc = BBlt(frmMain.picJournal, frmMain.picMisc, 48 + c * 102, 42, 18, 18, 0, 18, SRCCOPY)
				End If
			Next c
			
			frmMain.ShowText((frmMain.picJournal), 72, 45, 102, 14, bdFontNoxiousBlack, "Journal", False, False)
			frmMain.ShowText((frmMain.picJournal), 174, 45, 102, 14, bdFontNoxiousBlack, "Quests", False, False)
			frmMain.ShowText((frmMain.picJournal), 274, 45, 102, 14, bdFontNoxiousBlack, "Map", False, False)
			
			frmMain.ShowButton((frmMain.picJournal), 268, 350, "Done", False)
			
			' Show stuff
			Select Case JournalMode
				Case 1 ' Quests
					frmMain.picMicroBox.Visible = False
					'UPGRADE_WARNING: Couldn't resolve default property of object JournalList().Text. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
					frmMain.ShowText((frmMain.picJournal), 36, 81, 276, 28, bdFontElixirBlack, BreakText(JournalList.Item(JournalTop).Text, 2), False, False)
					'UPGRADE_WARNING: Couldn't resolve default property of object JournalList().Text. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
					frmMain.ShowText((frmMain.picJournal), 36, 118, 276, 202, bdFontNoxiousBlack, BreakText(JournalList.Item(JournalTop).Text, 3), False, False)
					frmMain.ShowText((frmMain.picJournal), 0, 318, frmMain.picJournal.Width, 18, bdFontNoxiousBlack, "Quest " & JournalTop & " of " & JournalList.Count(), True, False)
					'UPGRADE_ISSUE: frmMain.picJournal was upgraded to a Panel, and cannot be coerced to a PictureBox. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="0FF1188E-27C0-4FED-842D-159C65894C9B"'
					frmMain.ScrollBarShow((frmMain.picJournal), 317, 81, 218, JournalTop, JournalList.Count(), 0)
				Case 2 ' Map
					frmMain.picMicroBox.Visible = True
					JournalShowMap()
				Case Else ' Journal
					frmMain.picMicroBox.Visible = False
					If ShowDeleteButton Then
						frmMain.ShowButton((frmMain.picJournal), 68, 350, "Delete", False)
						blnDeleteButton = True
					End If
					frmMain.ShowButton((frmMain.picJournal), 168, 350, "Add", False)
					'UPGRADE_WARNING: Couldn't resolve default property of object JournalList().Text. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
					frmMain.ShowText((frmMain.picJournal), 36, 81, 276, 220, bdFontNoxiousBlack, JournalList.Item(JournalTop).Text, False, False)
					frmMain.ShowText((frmMain.picJournal), 0, 318, frmMain.picJournal.Width, 18, bdFontNoxiousBlack, "Entry " & JournalTop & " of " & JournalList.Count(), True, False)
					'UPGRADE_ISSUE: frmMain.picJournal was upgraded to a Panel, and cannot be coerced to a PictureBox. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="0FF1188E-27C0-4FED-842D-159C65894C9B"'
					frmMain.ScrollBarShow((frmMain.picJournal), 317, 81, 218, JournalTop, JournalList.Count(), 0)
			End Select
		End If
		frmMain.picJournal.BringToFront()
		frmMain.picJournal.Visible = True
		frmMain.picJournal.Refresh()
		'check for character input if adding a journal entry
		If bAddMode Then
			
		End If
	End Sub
	
	Public Function ShowDeleteButton() As Boolean
		' [Titi] 2.4.6
		'UPGRADE_WARNING: Couldn't resolve default property of object JournalList(JournalTop).Text. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
		If CountQuests = JournalList.Count() And JournalList.Item(JournalTop).Text = "No Journal Entries." Then
			' all entries are quests
			ShowDeleteButton = False
			'UPGRADE_WARNING: Couldn't resolve default property of object JournalList(JournalTop).Text. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
		ElseIf JournalList.Item(JournalTop).Text = "No Journal Entries." Then 
			' case of the first entry - we know there's no quest, but do we have a journal entry?
			ShowDeleteButton = False
		Else
			ShowDeleteButton = True
		End If
	End Function
	
	Public Function CountQuests() As Short
		' [Titi] 2.4.6
		Dim intQuests As Short
		Dim JournalX As Journal
		intQuests = 0
        For Each JournalX In tome.Journals
            If Left(BreakText(JournalX.Text, 1), 5) = "Quest" Then intQuests = intQuests + 1
        Next JournalX
		CountQuests = intQuests
	End Function
	
	Public Sub AddJournalText(ByRef KeyAscii As Short)
		If Not bAddMode Then Exit Sub
		Select Case KeyAscii
			Case System.Windows.Forms.Keys.Return
				'UPGRADE_WARNING: Couldn't resolve default property of object JournalList(JournalTop).Text. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
				'UPGRADE_WARNING: Couldn't resolve default property of object JournalList().Text. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
				JournalList.Item(JournalTop).Text = JournalList.Item(JournalTop).Text & vbCr
			Case System.Windows.Forms.Keys.Back, System.Windows.Forms.Keys.Left ' BackSpace
				'UPGRADE_WARNING: Couldn't resolve default property of object JournalList().Text. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
				If Len(JournalList.Item(JournalTop).Text) > 0 Then
					'UPGRADE_WARNING: Couldn't resolve default property of object JournalList(JournalTop).Text. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
					'UPGRADE_WARNING: Couldn't resolve default property of object JournalList().Text. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
					JournalList.Item(JournalTop).Text = Left(JournalList.Item(JournalTop).Text, Len(JournalList.Item(JournalTop).Text) - 1)
				End If
			Case System.Windows.Forms.Keys.Space ' Space
				'UPGRADE_WARNING: Couldn't resolve default property of object JournalList(JournalTop).Text. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
				'UPGRADE_WARNING: Couldn't resolve default property of object JournalList().Text. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
				JournalList.Item(JournalTop).Text = JournalList.Item(JournalTop).Text & vbLf
				'Case 48 To 57, 65 To 90, 97 To 122
			Case 33 To 126 '57, 65 To 90, 97 To 122
				'If Len(JournalList(JournalTop).Text) < 20 Then
				'UPGRADE_WARNING: Couldn't resolve default property of object JournalList(JournalTop).Text. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
				'UPGRADE_WARNING: Couldn't resolve default property of object JournalList().Text. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
				JournalList.Item(JournalTop).Text = JournalList.Item(JournalTop).Text & Chr(KeyAscii)
				'End If
		End Select
		'UPGRADE_ISSUE: PictureBox method picJournal.Cls was not upgraded. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="CC4C7EC0-C903-48FC-ACCC-81861D12DA4A"'
        frmMain.picJournal = Nothing
        frmMain.picJournal.Invalidate()
		'UPGRADE_WARNING: Couldn't resolve default property of object JournalList().Text. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
		frmMain.ShowText((frmMain.picJournal), 36, 81, 276, 220, bdFontNoxiousBlack, JournalList.Item(JournalTop).Text & "\", False, False)
		frmMain.picJournal.Focus()
		JournalShow()
	End Sub
	
	Public Sub JournalShowMap()
        Dim Side, X, Y As Short

		'UPGRADE_NOTE: my was upgraded to my_Renamed. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="A9E4979A-37FA-4718-9994-97DD76ED70A7"'
        Dim XMap, YMap As Short

		Dim FromY, ToY As Short
		Dim FromX, ToX As Short
		' Limit of MicroMap Picture
		SetMicroMapCursor(0, 0, MicroMapLeft, MicroMapTop, FromX, FromY, XMap, YMap)
		SetMicroMapCursor(bdMicroSizeWidth + 24, bdMicroSizeHeight, MicroMapLeft, MicroMapTop, ToX, ToY, XMap, YMap)
		FromX = Int(FromX / bdTileWidth) : FromY = Int(FromY / (bdTileHeight / 3)) + (XMap + YMap) Mod 2
		ToX = Int(ToX / bdTileWidth) + 1 : ToY = Int(ToY / (bdTileHeight / 3))
		' Draw Top of Region
		'UPGRADE_ISSUE: PictureBox method picMicroMap.Cls was not upgraded. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="CC4C7EC0-C903-48FC-ACCC-81861D12DA4A"'
        frmMain.picMicroMap = Nothing
        frmMain.picMicroMap.Invalidate()
		Y = FromY
		For X = FromX To ToX - 1
			Side = bdSideTop
			If X = FromX Then
				Side = Side Or bdSideLeft
			End If
			If X = ToX Then
				Side = Side Or bdSideRight
			End If
			frmMain.DrawTileMicro(X, Y, Side)
		Next X
		' Draw Middle of Region
		For Y = FromY + 1 To ToY
			For X = FromX To ToX - (Y Mod 2)
				Side = 0
				If X = FromX And (Y Mod 2) = 0 Then
					Side = Side Or bdSideLeft
				End If
				If X = ToX And (Y Mod 2) = 0 Then
					Side = Side Or bdSideRight
				End If
				frmMain.DrawTileMicro(X, Y, Side)
			Next X
		Next Y
		' Draw Bottom of Region
		For Y = ToY + 1 To ToY + 2
			For X = FromX To ToX - (Y Mod 2)
				If Y = ToY + 1 Then
					Side = bdSideFirstBottom
				Else
					Side = bdSideSecondBottom
				End If
				If X = FromX And (Y Mod 2) = 0 Then
					Side = Side Or bdSideLeft
				End If
				If X = ToX And (Y Mod 2) = 0 Then
					Side = Side Or bdSideRight
				End If
				frmMain.DrawTileMicro(X, Y, Side)
			Next X
		Next Y
		frmMain.picMicroMap.Refresh()
	End Sub
	
	Public Sub JournalMoveMap()

		MicroMapLeft = MicroMapLeft - Int(frmMain.picMicroMap.Left / 24)
		MicroMapTop = MicroMapTop - Int(frmMain.picMicroMap.Left / 24)
		MicroMapLeft = MicroMapLeft + Int(frmMain.picMicroMap.Top / 18)
		MicroMapTop = MicroMapTop - Int(frmMain.picMicroMap.Top / 18)
		frmMain.picMicroMap.Visible = False
		frmMain.picMicroMap.Top = 0 : frmMain.picMicroMap.Left = 0
		frmMain.picMicroMap.Visible = True
		JournalShowMap()
	End Sub
	
	Public Sub JournalQuestsLoad()
        Dim Found As Short
		Dim JournalX As Journal
		JournalList = New Collection
		Found = False
        For Each JournalX In tome.Journals
            If Left(BreakText(JournalX.Text, 1), 5) = "Quest" Then
                JournalList.Add(JournalX)
                Found = True
            End If
        Next JournalX
		If Found = False Then
			JournalX = New Journal
			JournalX.Text = "Quest|No Quests.|"
			JournalList.Add(JournalX)
		End If
		JournalTop = 1
	End Sub
	
	Public Sub JournalEntryLoad()
        Dim Found As Short
		Dim JournalX As Journal
		JournalList = New Collection
		Found = False
        For Each JournalX In tome.Journals
            If Left(BreakText(JournalX.Text, 1), 5) <> "Quest" And Left(BreakText(JournalX.Text, 1), 3) <> "Map" Then
                JournalList.Add(JournalX)
                Found = True
            End If
        Next JournalX
		If Found = False Then
			JournalX = New Journal
			JournalX.Text = "No Journal Entries."
			JournalList.Add(JournalX)
		End If
		JournalTop = 1
	End Sub
	
	Public Sub SetMicroMapCursor(ByVal ScreenX As Short, ByVal ScreenY As Short, ByRef MapLeft As Short, ByRef MapTop As Short, ByRef ToScreenX As Short, ByRef ToScreenY As Short, ByRef ToMapX As Short, ByRef ToMapY As Short)
		' Converts any ScreenX, ScreenY to true blue MapX and MapY
		Dim TileY24, TileY8, TileY48 As Short
		Dim TileX24, TileX48 As Short
		TileY8 = bdTileHeight / 9 : TileY24 = TileY8 * 3 : TileY48 = TileY8 * 6
		TileX24 = bdTileWidth / 4 : TileX48 = TileX24 * 2
		If Int(ScreenY / (TileY8 * 2)) Mod 3 = 1 Or (Int(ScreenY / TileY8) Mod 3 = 1 And Int((ScreenX - TileX24) / TileX48) Mod 2 = 0) Then
			ToMapX = MapLeft + Int(ScreenX / bdTileWidth) - Int(ScreenY / TileY48)
			ToMapY = MapTop + Int(ScreenX / bdTileWidth) + Int(ScreenY / TileY48)
			ToScreenX = Int(ScreenX / bdTileWidth) * bdTileWidth
			ToScreenY = Int(ScreenY / TileY48) * TileY48
		Else
			ToMapX = MapLeft + Int((ScreenX - TileX48) / bdTileWidth) - Int((ScreenY - TileY24) / TileY48)
			ToMapY = MapTop + Int((ScreenX + TileX48) / bdTileWidth) + Int((ScreenY - TileY24) / TileY48)
			ToScreenX = Int((ScreenX - TileX48) / bdTileWidth) * bdTileWidth + TileX48
			ToScreenY = Int((ScreenY - TileY24) / TileY48) * TileY48 + TileY24
		End If
	End Sub
End Module