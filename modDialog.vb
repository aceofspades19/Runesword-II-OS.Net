Option Strict Off
Option Explicit On
Module modDialog

    Private tome As Tome = Tome.getInstance()
	' Variables for dialog box
	Public ConvoAction As Short
	Public SorceryAction As Short
	Public ReplyIndex(16) As Short
	Public ReplyHeight(16) As Short
	Public ReplySelect As Short
	Public ReplyTop As Short
	Public ReplyList(16) As String
	Public ButtonList(4) As String
	Public DialogStyle As modGameGeneral.DLGTYPE
	Public DialogText As String
	Public DialogReplySelect As Short
	Public DialogReplyTop As Short
	Public DialogReplyList(16) As String
	Public DialogReplyHeight(16) As Short
	
	Public bDialog As Boolean
	
	' List Box for Buy/Sell
	Dim SellThumbY As Short
	Dim SellTop As Short
	Dim SellSelect As Short
	Dim BuyRate As Double
	Dim BuyStyle As Short
	Dim BuyFrom As Object
	Dim BuyThumbY As Short
	Dim BuyTop As Short
	Dim BuySelect As Short
	' NPC is always SellList, PC is always BuyList
	Dim SellList As Collection
	Dim SellList2Index As Collection
	Dim BuyList As Collection
	Dim BuyList2Index As Collection
	
	Public Const bdMaxItemPics As Short = 48
	Public ItemPicFile(bdMaxItemPics) As String
	Public ItemPicTime(bdMaxItemPics) As Double
	Public ItemPicWidth(bdMaxItemPics) As Byte
	Public ItemPicHeight(bdMaxItemPics) As Byte
	Public Const bdMaxMonsPics As Short = 12
	Public PicFile(bdMaxMonsPics) As String
	Public PicPortrait(bdMaxMonsPics) As String
	Public PicFileTime(bdMaxMonsPics) As Double
	
	Public Sub DialogSetUp(ByRef Style As modGameGeneral.DLGTYPE)
		Dim c As Short
		Dim rc As Integer
		' This makes for a cleaner resize of the dialog box
		With frmMain
			'UPGRADE_ISSUE: PictureBox method picConvo.Cls was not upgraded. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="CC4C7EC0-C903-48FC-ACCC-81861D12DA4A"'
            .picConvo = Nothing
            .picConvo.Invalidate()
			'UPGRADE_ISSUE: PictureBox method picConvoList.Cls was not upgraded. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="CC4C7EC0-C903-48FC-ACCC-81861D12DA4A"'
            .picConvoList.Invalidate() : .picConvoList.Visible = False
			.picConvoEnter.Visible = False
			Select Case Style
				Case modGameGeneral.DLGTYPE.bdDlgNoReply, modGameGeneral.DLGTYPE.bdDlgItem
					.picConvo.Height = 146
				Case modGameGeneral.DLGTYPE.bdDlgWithReply
					DialogReplyTop = 0
				Case modGameGeneral.DLGTYPE.bdDlgItemList, modGameGeneral.DLGTYPE.bdDlgCreatureList, modGameGeneral.DLGTYPE.bdDlgDebug
					.picConvoList.Visible = True
					.picConvo.Height = 348
				Case modGameGeneral.DLGTYPE.bdDlgReplyText
					.picConvoEnter.Visible = True
				Case modGameGeneral.DLGTYPE.bdDlgSave
					.picConvo.Height = 290
					.picConvoEnter.Left = 106 : .picConvoEnter.Top = 224
					.picConvoEnter.Visible = True
				Case modGameGeneral.DLGTYPE.bdDlgCredits
					.picConvo.Height = 498
			End Select
			DialogStyle = Style
			' Set up default buttons
			ButtonList(0) = "Done"
			For c = 1 To 4
				ButtonList(c) = ""
			Next c
		End With
	End Sub
	
	Public Sub DialogClick(ByRef AtX As Short, ByRef AtY As Short, ByRef ButtonDown As Short)
		Dim ButtonCnt, c, PosY As Short
		Dim rc As Integer
		' Count the available buttons (compatabiity issues here)
		For c = 0 To 4
			If ButtonList(c) <> "" Then
				ButtonCnt = c + 1
			End If
		Next c
		With frmMain
			' Check for button clicks
			For c = 1 To ButtonCnt
				If PointIn(AtX, AtY, .picConvo.Width - 6 - 96 * c, .picConvo.Height - 30, 90, 18) Then
					.ShowButton(.picConvo, .picConvo.Width - 6 - 96 * c, .picConvo.Height - 30, ButtonList(c - 1), ButtonDown)
					.picConvo.Refresh()
					If ButtonDown = False Then
						Call PlayClickSnd(modIOFunc.ClickType.ifClick)
						ConvoAction = ButtonCnt - c
					End If
				End If
			Next c
			' Check for DialogReply Clicks
			If DialogStyle = modGameGeneral.DLGTYPE.bdDlgWithReply Then
				If PointIn(AtX, AtY, 88, 34, 354, .picConvo.Height - 70) And ButtonDown = False Then
					' Choose a Reply
					PosY = DialogReplyHeight(0) - 4
					For c = 1 To DialogReplyTop
						If PointIn(AtX, AtY, 88, PosY, 354, DialogReplyHeight(c)) Then
							Call PlayClickSnd(modIOFunc.ClickType.ifClick)
							ConvoAction = c - 1
							Exit For
						End If
						PosY = PosY + DialogReplyHeight(c) + 4
					Next c
				End If
			End If
		End With
	End Sub
	
	Public Function DialogShow(ByRef CreatureX As Object, ByRef Text As String) As Short
        Dim c, PosY As Short
        Dim OldPointer As Cursor
		Dim rc As Integer
        Dim OldFrozen As Cursor
		'UPGRADE_WARNING: Screen property Screen.MousePointer has a new behavior. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6BA9B8D2-2A32-4B6E-8D36-44949974A5B4"'
		OldPointer = System.Windows.Forms.Cursor.Current

		'UPGRADE_WARNING: Screen property Screen.MousePointer has a new behavior. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6BA9B8D2-2A32-4B6E-8D36-44949974A5B4"'
		System.Windows.Forms.Cursor.Current = System.Windows.Forms.Cursors.Default
		Frozen = True
		With frmMain
			' Display Character shown by Character Name
			DialogShowFace(CreatureX)
			Select Case DialogStyle
				Case modGameGeneral.DLGTYPE.bdDlgNoReply, modGameGeneral.DLGTYPE.bdDlgItem
					.ShowText(.picConvo, 0, 12, .picConvo.Width, 14, bdFontElixirWhite, "Game Master", True, False)
					PosY = .ShowText(.picConvo, 88, 34, 354, 430, bdFontNoxiousWhite, Text, False, True)
					.picConvo.Height = Greatest(PosY + 70, 146)
				Case modGameGeneral.DLGTYPE.bdDlgWithReply
					.ShowText(.picConvo, 0, 12, .picConvo.Width, 14, bdFontElixirWhite, "Choose a Reply", True, False)
					DialogReplyList(0) = Text
					DialogReplyShow()
				Case modGameGeneral.DLGTYPE.bdDlgWithDice
					.ShowText(.picConvo, 0, 12, .picConvo.Width, 14, bdFontElixirWhite, "Roll Dice", True, False)
				Case modGameGeneral.DLGTYPE.bdDlgItemList
					.ShowText(.picConvo, 0, 12, .picConvo.Width, 14, bdFontElixirWhite, Text, True, False)
				Case modGameGeneral.DLGTYPE.bdDlgCreatureList
					.ShowText(.picConvo, 0, 12, .picConvo.Width, 14, bdFontElixirWhite, "Target Creature", True, False)
				Case modGameGeneral.DLGTYPE.bdDlgDebug
					.ShowText(.picConvo, 0, 12, .picConvo.Width, 14, bdFontSmallWhite, DebugHeader, True, False)
				Case modGameGeneral.DLGTYPE.bdDlgReplyText
					' [Titi 2.4.9] change title for shortcuts
					If Left(Text & Space(8), 8) = "Shortcut" Then
						.ShowText(.picConvo, 0, 12, .picConvo.Width, 14, bdFontElixirWhite, "Game Options", True, False)
					Else
						.ShowText(.picConvo, 0, 12, .picConvo.Width, 14, bdFontElixirWhite, "Enter a Reply", True, False)
					End If
					PosY = .ShowText(.picConvo, 88, 34, 354, 430, bdFontNoxiousWhite, Text, False, True)
					.picConvoEnter.Top = PosY + 40
					.picConvo.Height = Greatest(.picConvoEnter.Height + .picConvoEnter.Top + 40, 146)
					.picConvo.Visible = True
					DialogEnter("")
			End Select
			' Paint Buttons
			For c = 1 To 5
				If ButtonList(c - 1) <> "" Then
					.ShowButton(.picConvo, .picConvo.Width - 6 - 96 * c, .picConvo.Height - 30, ButtonList(c - 1), False)
				End If
			Next c
			' Disable Talk (if visible)
			If .picTalk.Visible = True Then
				.picTalk.Enabled = False
			End If
			' Paint bottom of picConvo
			'UPGRADE_ISSUE: PictureBox property picConvoBottom.hdc was not upgraded. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="CC4C7EC0-C903-48FC-ACCC-81861D12DA4A"'
			'UPGRADE_ISSUE: PictureBox property picConvo.hdc was not upgraded. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="CC4C7EC0-C903-48FC-ACCC-81861D12DA4A"'
            rc = BBlt(.picConvo, .picConvoBottom, 0, .picConvo.Height - 6, .picConvo.Width, 6, 0, 0, SRCCOPY)
			.picConvo.Top = Greatest((.picBox.ClientRectangle.Height - .picConvo.ClientRectangle.Height) / 2, 0)
			.picConvo.Refresh()
			.picConvo.Visible = True
			.picConvo.BringToFront()
			.picConvo.Focus()
			If DialogStyle = modGameGeneral.DLGTYPE.bdDlgReplyText Then
				.picConvoEnter.Focus()
			End If
			ConvoAction = -1
			Frozen = True
			bDialog = True
			'UPGRADE_WARNING: Screen property Screen.MousePointer has a new behavior. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6BA9B8D2-2A32-4B6E-8D36-44949974A5B4"'
			System.Windows.Forms.Cursor.Current = System.Windows.Forms.Cursors.Default
			Do Until ConvoAction > -1
				System.Windows.Forms.Application.DoEvents()
				'[borf] 2.4.6a Freeze issue - Phule's example #1
				' If Talk form was activated over a Dialog needs to
				' become visible again.
				If .picConvo.Visible = False And bDialog Then
					.picConvo.Visible = True
				End If
			Loop 
			bDialog = False
			DialogShow = ConvoAction
			If .picTalk.Visible = True Then
				.picTalk.Enabled = True
			End If
		End With
        'Frozen = OldFrozen
		'UPGRADE_ISSUE: Screen property Screen.MousePointer does not support custom mousepointers. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="45116EAB-7060-405E-8ABE-9DBB40DC2E86"'
		'UPGRADE_WARNING: Screen property Screen.MousePointer has a new behavior. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6BA9B8D2-2A32-4B6E-8D36-44949974A5B4"'
		System.Windows.Forms.Cursor.Current = OldPointer
	End Function
	
	'UPGRADE_NOTE: Tome was upgraded to Tome_Renamed. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="A9E4979A-37FA-4718-9994-97DD76ED70A7"'
	Public Sub DialogShowFace(ByRef CreatureX As Object)
		Dim Tome_Renamed As Object
		Dim rc As Integer
		Dim c, Found As Short
		Dim X, Y As Short
		X = 14 : Y = 37
		With frmMain
			' Paint frame
			'UPGRADE_ISSUE: PictureBox property picMisc.hdc was not upgraded. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="CC4C7EC0-C903-48FC-ACCC-81861D12DA4A"'
			'UPGRADE_ISSUE: PictureBox property picConvo.hdc was not upgraded. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="CC4C7EC0-C903-48FC-ACCC-81861D12DA4A"'
            rc = BBlt(.picConvo, .picMisc, X - 4, Y - 4, 74, 84, 0, 36, SRCCOPY)
			' Paint Creature portrait
			'UPGRADE_WARNING: IsObject has a new behavior. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="9B7D5ADD-D8FE-4819-A36C-6DEDAF088CC7"'
			If IsReference(CreatureX) Then
				'UPGRADE_WARNING: Couldn't resolve default property of object CreatureX.Pic. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
				'UPGRADE_ISSUE: PictureBox property picFaces.hdc was not upgraded. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="CC4C7EC0-C903-48FC-ACCC-81861D12DA4A"'
				'UPGRADE_ISSUE: PictureBox property picConvo.hdc was not upgraded. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="CC4C7EC0-C903-48FC-ACCC-81861D12DA4A"'
                rc = BBlt(.picConvo, .picFaces, X, Y, 66, 76, bdFaceMin + CreatureX.Pic * 66, 0, SRCCOPY)
			Else
				'UPGRADE_WARNING: Couldn't resolve default property of object CreatureX. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
				If CreatureX = "CreatureWithTurn" Then
					'UPGRADE_ISSUE: PictureBox property picFaces.hdc was not upgraded. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="CC4C7EC0-C903-48FC-ACCC-81861D12DA4A"'
					'UPGRADE_ISSUE: PictureBox property picConvo.hdc was not upgraded. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="CC4C7EC0-C903-48FC-ACCC-81861D12DA4A"'
                    rc = BBlt(.picConvo, .picFaces, X, Y, 66, 76, bdFaceMin + CreatureWithTurn.Pic * 66, 0, SRCCOPY)
					'UPGRADE_WARNING: Couldn't resolve default property of object CreatureX. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
				ElseIf CreatureX = "CreatureNow" Then 
					'UPGRADE_ISSUE: PictureBox property picFaces.hdc was not upgraded. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="CC4C7EC0-C903-48FC-ACCC-81861D12DA4A"'
					'UPGRADE_ISSUE: PictureBox property picConvo.hdc was not upgraded. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="CC4C7EC0-C903-48FC-ACCC-81861D12DA4A"'
                    rc = BBlt(.picConvo, .picFaces, X, Y, 66, 76, bdFaceMin + CreatureWithTurn.Pic * 66, 0, SRCCOPY)
					'UPGRADE_WARNING: Couldn't resolve default property of object CreatureX. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
				ElseIf CreatureX = "CreatureTarget" Then 
					'UPGRADE_ISSUE: PictureBox property picFaces.hdc was not upgraded. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="CC4C7EC0-C903-48FC-ACCC-81861D12DA4A"'
					'UPGRADE_ISSUE: PictureBox property picConvo.hdc was not upgraded. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="CC4C7EC0-C903-48FC-ACCC-81861D12DA4A"'
                    rc = BBlt(.picConvo, .picFaces, X, Y, 66, 76, bdFaceMin + CreatureTarget.Pic * 66, 0, SRCCOPY)
				Else
					Found = False
                    For c = 1 To tome.Creatures.Count()
                        'UPGRADE_WARNING: Couldn't resolve default property of object CreatureX. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
                        'UPGRADE_WARNING: Couldn't resolve default property of object Tome.Creatures. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
                        If InStr(tome.Creatures(c).Name, CreatureX) > 0 Then
                            'UPGRADE_WARNING: Couldn't resolve default property of object Tome.Creatures. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
                            'UPGRADE_ISSUE: PictureBox property picFaces.hdc was not upgraded. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="CC4C7EC0-C903-48FC-ACCC-81861D12DA4A"'
                            'UPGRADE_ISSUE: PictureBox property picConvo.hdc was not upgraded. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="CC4C7EC0-C903-48FC-ACCC-81861D12DA4A"'
                            rc = BBlt(.picConvo, .picFaces, X, Y, 66, 76, bdFaceMin + tome.Creatures(c).Pic * 66, 0, SRCCOPY)
                            Found = True
                            Exit For
                        End If
                    Next c
					If Not Found Then
						For c = 1 To EncounterNow.Creatures.Count()
							'UPGRADE_WARNING: Couldn't resolve default property of object CreatureX. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
							'UPGRADE_WARNING: Couldn't resolve default property of object EncounterNow.Creatures().Name. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
							If InStr(EncounterNow.Creatures.Item(c).Name, CreatureX) > 0 Then
								'UPGRADE_WARNING: Couldn't resolve default property of object EncounterNow.Creatures().Pic. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
								'UPGRADE_ISSUE: PictureBox property picFaces.hdc was not upgraded. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="CC4C7EC0-C903-48FC-ACCC-81861D12DA4A"'
								'UPGRADE_ISSUE: PictureBox property picConvo.hdc was not upgraded. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="CC4C7EC0-C903-48FC-ACCC-81861D12DA4A"'
                                rc = BBlt(.picConvo, .picFaces, X, Y, 66, 76, bdFaceMin + EncounterNow.Creatures.Item(c).Pic * 66, 0, SRCCOPY)
								Found = True
								Exit For
							End If
						Next c
						If Not Found Then
							'UPGRADE_ISSUE: PictureBox property picFaces.hdc was not upgraded. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="CC4C7EC0-C903-48FC-ACCC-81861D12DA4A"'
							'UPGRADE_ISSUE: PictureBox property picConvo.hdc was not upgraded. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="CC4C7EC0-C903-48FC-ACCC-81861D12DA4A"'
                            rc = BBlt(.picConvo, .picFaces, X, Y, 66, 76, bdFaceDM, 0, SRCCOPY)
							Found = True
						End If
					End If
				End If
			End If
			.picConvo.Refresh()
		End With
	End Sub
	
	'Private Sub DialogShowFaceBrief(CreatureX As Creature)
	Public Sub DialogShowFaceBrief(ByRef CreatureX As Creature)
		Dim rc As Integer
		Dim X, Y As Short
		X = 14 : Y = 37
		With frmMain
			'UPGRADE_ISSUE: PictureBox property picMisc.hdc was not upgraded. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="CC4C7EC0-C903-48FC-ACCC-81861D12DA4A"'
			'UPGRADE_ISSUE: PictureBox property picDialogBrief.hdc was not upgraded. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="CC4C7EC0-C903-48FC-ACCC-81861D12DA4A"'
            rc = BBlt(.picDialogBrief, .picMisc, X - 4, Y - 4, 74, 84, 0, 36, SRCCOPY)
			If CreatureX.Name = "DM" Then
				.ShowText(.picDialogBrief, 0, 12, .picDialogBrief.Width, 14, bdFontElixirWhite, "Game Master", True, False)
				'UPGRADE_ISSUE: PictureBox property picFaces.hdc was not upgraded. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="CC4C7EC0-C903-48FC-ACCC-81861D12DA4A"'
				'UPGRADE_ISSUE: PictureBox property picDialogBrief.hdc was not upgraded. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="CC4C7EC0-C903-48FC-ACCC-81861D12DA4A"'
                rc = BBlt(.picDialogBrief, .picFaces, X, Y, 66, 76, bdFaceDM, 0, SRCCOPY)
			ElseIf CreatureX.Name = "CreatureWithTurn" Then 
				.ShowText(.picDialogBrief, 0, 12, .picDialogBrief.ClientRectangle.Width, 14, bdFontElixirWhite, (CreatureWithTurn.Name), True, False)
				'UPGRADE_ISSUE: PictureBox property picFaces.hdc was not upgraded. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="CC4C7EC0-C903-48FC-ACCC-81861D12DA4A"'
				'UPGRADE_ISSUE: PictureBox property picDialogBrief.hdc was not upgraded. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="CC4C7EC0-C903-48FC-ACCC-81861D12DA4A"'
                rc = BBlt(.picDialogBrief, .picFaces, X, Y, 66, 76, bdFaceMin + CreatureWithTurn.Pic * 66, 0, SRCCOPY)
			Else
				.ShowText(.picDialogBrief, 0, 12, .picDialogBrief.Width, 14, bdFontElixirWhite, (CreatureX.Name), True, False)
				'UPGRADE_ISSUE: PictureBox property picFaces.hdc was not upgraded. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="CC4C7EC0-C903-48FC-ACCC-81861D12DA4A"'
				'UPGRADE_ISSUE: PictureBox property picDialogBrief.hdc was not upgraded. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="CC4C7EC0-C903-48FC-ACCC-81861D12DA4A"'
                rc = BBlt(.picDialogBrief, .picFaces, X, Y, 66, 76, bdFaceMin + CreatureX.Pic * 66, 0, SRCCOPY)
			End If
			.picDialogBrief.Refresh()
		End With
	End Sub
	
	Public Sub DialogReplyMove(ByRef AtX As Short, ByRef AtY As Short)
		Dim c, PosY As Short
		PosY = DialogReplyHeight(0) - 4
		For c = 1 To DialogReplyTop
			If PointIn(AtX, AtY, 88, PosY, 354, DialogReplyHeight(c)) And c <> DialogReplySelect Then
				DialogReplySelect = c
				DialogReplyShow()
				Exit For
			End If
			PosY = PosY + DialogReplyHeight(c) + 4
		Next c
	End Sub
	
	Public Sub DialogReplyAdd(ByRef Text As String)
		If DialogReplyTop < 16 Then
			DialogReplyTop = DialogReplyTop + 1
			DialogReplyList(DialogReplyTop) = Text
		End If
	End Sub
	
	Public Sub DialogReplyShow()
		Dim c, PosY As Short
		' List Reply or First/Second Talk
		PosY = 42 + 16 + frmMain.ShowText((frmMain.picConvo), 88, 34, 354, 430, bdFontNoxiousWhite, DialogReplyList(0), False, True)
		DialogReplyHeight(0) = PosY
		' List Topics
		For c = 1 To DialogReplyTop
			If DialogReplySelect = c Then
				DialogReplyHeight(c) = frmMain.ShowText((frmMain.picConvo), 88, PosY, 354, 152, bdFontNoxiousGold, DialogReplyList(c), False, True)
			Else
				DialogReplyHeight(c) = frmMain.ShowText((frmMain.picConvo), 88, PosY, 354, 152, bdFontNoxiousGrey, DialogReplyList(c), False, True)
			End If
			PosY = PosY + DialogReplyHeight(c) + 4
		Next c
		frmMain.picConvo.Height = PosY + 36
		frmMain.picConvo.Refresh()
	End Sub
	
	Public Sub DialogEnter(ByRef Text As String)
		With frmMain
			'UPGRADE_ISSUE: PictureBox method picConvoEnter.Cls was not upgraded. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="CC4C7EC0-C903-48FC-ACCC-81861D12DA4A"'
            .picConvoEnter = Nothing
            .picConvoEnter.Invalidate()
			.ShowText(.picConvoEnter, 6, 7, .picConvoEnter.Width - 12, 14, bdFontNoxiousWhite, Text & "\", False, False)
			DialogText = Text
			.picConvoEnter.Focus()
		End With
	End Sub
	
	Public Sub DialogDM(ByRef Text As String)
        Dim OldPointer As Cursor
		Dim rc As Integer
		'UPGRADE_WARNING: Screen property Screen.MousePointer has a new behavior. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6BA9B8D2-2A32-4B6E-8D36-44949974A5B4"'
		OldPointer = System.Windows.Forms.Cursor.Current
		'UPGRADE_WARNING: Screen property Screen.MousePointer has a new behavior. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6BA9B8D2-2A32-4B6E-8D36-44949974A5B4"'
		System.Windows.Forms.Cursor.Current = System.Windows.Forms.Cursors.Default
		DialogSetUp(modGameGeneral.DLGTYPE.bdDlgNoReply)
		DialogSetButton(1, "Done")
		DialogShow("DM", Text)
		DialogHide()
		'UPGRADE_ISSUE: Screen property Screen.MousePointer does not support custom mousepointers. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="45116EAB-7060-405E-8ABE-9DBB40DC2E86"'
		'UPGRADE_WARNING: Screen property Screen.MousePointer has a new behavior. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6BA9B8D2-2A32-4B6E-8D36-44949974A5B4"'
		System.Windows.Forms.Cursor.Current = OldPointer
	End Sub
	
	'UPGRADE_NOTE: Tome was upgraded to Tome_Renamed. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="A9E4979A-37FA-4718-9994-97DD76ED70A7"'
	Public Sub DialogBrief(ByRef CreatureX As Object, ByRef Text As String)
		Dim Tome_Renamed As Object
		Dim CreatureZ As Creature
		Dim Found, c As Short
		Dim DialogX As DialogStack
		Found = False
		'UPGRADE_WARNING: IsObject has a new behavior. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="9B7D5ADD-D8FE-4819-A36C-6DEDAF088CC7"'
		If IsReference(CreatureX) Then
			Found = True
			CreatureZ = CreatureX ' [Titi 2.4.8] added
		Else
			'UPGRADE_WARNING: Couldn't resolve default property of object CreatureX. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
			If CreatureX = "CreatureWithTurn" Then
				CreatureZ = CreatureWithTurn
				Found = True
				'UPGRADE_WARNING: Couldn't resolve default property of object CreatureX. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
			ElseIf CreatureX = "CreatureWithTurn" Then 
				CreatureZ = CreatureWithTurn
				Found = True
				'UPGRADE_WARNING: Couldn't resolve default property of object CreatureX. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
			ElseIf CreatureX = "CreatureTarget" Then 
				CreatureZ = CreatureTarget
				Found = True
			Else
                For c = 1 To tome.Creatures.Count()
                    'UPGRADE_WARNING: Couldn't resolve default property of object CreatureX. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
                    'UPGRADE_WARNING: Couldn't resolve default property of object Tome.Creatures. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
                    If InStr(tome.Creatures(c).Name, CreatureX) > 0 Then
                        'UPGRADE_WARNING: Couldn't resolve default property of object Tome.Creatures. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
                        CreatureZ = tome.Creatures(c)
                        Found = True
                        Exit For
                    End If
                Next c
				If Not Found Then
					For c = 1 To EncounterNow.Creatures.Count()
						'UPGRADE_WARNING: Couldn't resolve default property of object CreatureX. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
						'UPGRADE_WARNING: Couldn't resolve default property of object EncounterNow.Creatures().Name. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
						If InStr(EncounterNow.Creatures.Item(c).Name, CreatureX) > 0 Then
							CreatureZ = EncounterNow.Creatures.Item(c)
							Found = True
							Exit For
						End If
					Next c
				End If
			End If
		End If
		' If not found then it's the DM Orb
		If Not Found Then
			CreatureZ = New Creature
			CreatureZ.Name = "DM"
		End If
		' Find new available unused index identifier
		c = 1
		For	Each DialogX In DialogBriefSet
			If DialogX.Index >= c Then
				c = DialogX.Index + 1
			End If
		Next DialogX
		' Set the index and add the Topic.
		DialogX = New DialogStack
		DialogX.Index = c
		DialogX.Text = Text
		DialogX.CreatureTalking = CreatureZ
		DialogBriefSet.Add(DialogX, "D" & DialogX.Index)
		frmMain.tmrDialogBrief.Enabled = True
	End Sub
	
	Public Sub DialogHide()
		With frmMain
			.picConvo.Visible = False
			.picMap.Focus()
			.Refresh()
		End With
	End Sub
	
	Public Sub DialogSetButton(ByRef ButtonNumber As Short, ByRef Text As String)
		Dim c As Short
		' Paint the button
		ButtonList(ButtonNumber - 1) = Text
	End Sub
	
	'Public Sub DialogAcceptText()
	'    DialogSetUp bdDlgReplyText
	'End Sub
	
	'UPGRADE_NOTE: Rate was upgraded to Rate_Renamed. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="A9E4979A-37FA-4718-9994-97DD76ED70A7"'
	Public Sub DialogBuySell(ByRef TriggerX As Trigger, ByRef Style As Short, ByRef Rate_Renamed As Short)
		' CreatureTarget Buys and Sells ItemNow for a fixed amount = ItemNow.Value +/- (ItemNow.Value * (Rate/100))
		Dim ItemX, ItemZ As Item
		Dim DialogTalkVisible As Short
		' Setup CreatureSelling as CreatureTarget
		CreatureSelling = CreatureTarget
		' Setup list of Items to Sell
		Select Case Style
			Case 0 ' Creature
				BuyFrom = CreatureTarget
			Case 1 ' Encounter
				BuyFrom = EncounterNow
			Case 2 ' Trigger
				BuyFrom = TriggerX
		End Select
		' Set up buy/sell variables Transactions
		BuyList = New Collection
		'UPGRADE_WARNING: Couldn't resolve default property of object BuyFrom.Items. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
		For	Each ItemX In BuyFrom.Items
			If ItemX.IsMoney = False Then
				BuyList.Add(ItemX)
			End If
		Next ItemX
		BuyTop = Least(1, BuyList.Count()) : BuySelect = BuyTop
		BuyRate = Rate_Renamed / 100 : BuyStyle = Style
		' Set up CreatureWithTurn items
		DialogBuySellPC()
		' Hide Talk box if talking
		DialogTalkVisible = frmMain.picTalk.Visible
		frmMain.picTalk.Visible = False
		' Do some Transations
		DialogBuySellShow()
		' Wait here until done shopping
		Do Until frmMain.picBuySell.Visible = False
			System.Windows.Forms.Application.DoEvents()
		Loop 
		' Show Talk box if was talking
		frmMain.picTalk.Visible = DialogTalkVisible
	End Sub
	
	Public Sub DialogBuySellClick(ByRef AtX As Short, ByRef AtY As Short, ByRef ButtonDown As Short)
		Dim c As Short
		If PointIn(AtX, AtY, frmMain.picBuySell.ClientRectangle.Width - 102, frmMain.picBuySell.ClientRectangle.Height - 30, 90, 18) Then
			' Done
			frmMain.ShowButton((frmMain.picBuySell), frmMain.picBuySell.ClientRectangle.Width - 102, frmMain.picBuySell.ClientRectangle.Height - 30, "Done", ButtonDown)
			frmMain.picBuySell.Refresh()
			If ButtonDown = False Then
				Call PlayClickSnd(modIOFunc.ClickType.ifClick)
				frmMain.picBuySell.Visible = False
			End If
		ElseIf PointIn(AtX, AtY, 41, frmMain.picBuySell.ClientRectangle.Height - 30, 90, 18) And BuySelect > 0 Then 
			' Buy
			frmMain.ShowButton((frmMain.picBuySell), 41, frmMain.picBuySell.ClientRectangle.Height - 30, "Buy", ButtonDown)
			frmMain.picBuySell.Refresh()
			If ButtonDown = False Then
				Call PlayClickSnd(modIOFunc.ClickType.ifClick)
				DialogBuySellTransaction(True)
			End If
		ElseIf PointIn(AtX, AtY, 351, frmMain.picBuySell.ClientRectangle.Height - 30, 90, 18) And SellSelect > 0 Then 
			' Sell
			frmMain.ShowButton((frmMain.picBuySell), 351, frmMain.picBuySell.ClientRectangle.Height - 30, "Sell", ButtonDown)
			frmMain.picBuySell.Refresh()
			If ButtonDown = False Then
				Call PlayClickSnd(modIOFunc.ClickType.ifClick)
				DialogBuySellTransaction(False)
			End If
		ElseIf PointIn(AtX, AtY, 644, 44, 18, 312) And SellList.Count() > 0 Then 
			' Sell ScrollBar
			If frmMain.ScrollBarClick(AtX, AtY, ButtonDown, (frmMain.picBuySell), 644, 44, 312, SellTop, SellList.Count() + 1, 3) = True Then
				DialogBuySellShow()
			End If
		ElseIf PointIn(AtX, AtY, 357, 50, 278, 306) And ButtonDown = True And SellList.Count() > 0 Then 
			' Select Item to Sell
			c = Int((AtY - 50) / 102)
			If IsBetween(c, 0, 2) Then
				SellSelect = SellTop + c
				DialogBuySellShow()
				Call PlayClickSnd(modIOFunc.ClickType.ifClick)
			End If
		ElseIf PointIn(AtX, AtY, 23, 44, 18, 312) And BuyList.Count() > 0 Then 
			' Buy ScrollBar
			If frmMain.ScrollBarClick(AtX, AtY, ButtonDown, (frmMain.picBuySell), 23, 44, 312, BuyTop, BuyList.Count() + 1, 3) = True Then
				DialogBuySellShow()
			End If
		ElseIf PointIn(AtX, AtY, 43, 50, 278, 306) And ButtonDown = False And BuyList.Count() > 0 Then 
			' Select to Buy
			c = Int((AtY - 50) / 102)
			If IsBetween(c, 0, 2) Then
				BuySelect = BuyTop + c
				DialogBuySellShow()
				Call PlayClickSnd(modIOFunc.ClickType.ifClick)
			End If
		End If
	End Sub
	
	Private Sub DialogBuySellTransaction(ByRef IsBuying As Short)
		Dim ItemX, ItemZ As Item
		Dim CreatureFrom, CreatureTo As Creature
		Dim NoFail As Short
		If IsBuying = True Then
			ItemX = BuyList.Item(BuySelect)
			If ItemX.IsMoney = True Then
				GlobalOffer = CInt(ItemX.Value)
			Else
				GlobalOffer = CInt(ItemX.Value + ItemX.Value * BuyRate)
			End If
			CreatureFrom = CreatureSelling
			CreatureTo = CreatureWithTurn
		Else
			ItemX = SellList.Item(SellSelect)
			If ItemX.IsMoney = True Then
				GlobalOffer = CInt(ItemX.Value)
			Else
				GlobalOffer = CInt(ItemX.Value - ItemX.Value * BuyRate)
			End If
			CreatureFrom = CreatureWithTurn
			CreatureTo = CreatureSelling
		End If
		' Fire PreSell Triggers
		CreatureNow = CreatureFrom
		ItemNow = ItemX
		NoFail = FireTriggers(CreatureNow, bdPreSellItem)
		If NoFail = True Then
			CreatureNow = CreatureFrom
			ItemNow = ItemX
			NoFail = FireTriggers(ItemX, bdPreSellItem)
			' Fire PreBuy Triggers
			If NoFail = True Then
				CreatureNow = CreatureTo
				ItemNow = ItemX
				NoFail = FireTriggers(CreatureNow, bdPreBuyItem)
				If NoFail = True Then
					CreatureNow = CreatureTo
					ItemNow = ItemX
					NoFail = FireTriggers(ItemX, bdPreBuyItem)
				End If
			End If
		End If
		If ItemX.Capacity > 0 And ItemX.Items.Count() > 0 Then
			If IsBuying = True Then
				DialogDM("You cannot buy the " & ItemX.Name & " while it contains items.")
			Else
				DialogDM("You cannot sell the " & ItemX.Name & " while it contains items.")
			End If
			NoFail = False
		End If
		If NoFail = True Then
			' Check if Buyer has the funds (NPCs automatically have unlimited funds)
			If IsBuying = True Then
				' Check if have enough money to buy
				If CreatureWithTurn.Money < GlobalOffer Then
					DialogDM(CreatureWithTurn.Name & " doesn't have enough money to buy the " & ItemX.Name & ".")
					Exit Sub
				End If
				' Remove the Funds
				CreatureWithTurn.Money = CreatureWithTurn.Money - GlobalOffer
			Else
				' Check if PC can UnReady the Item selling
				If frmMain.UnReadyItem(CreatureWithTurn, ItemX) = False Then
					Exit Sub
				End If
				' Grant Funds
				ItemZ = New Item
				ItemZ.IsMoney = True
				' [Titi 2.4.8] let's use the currency of the world!
				ItemZ.Name = Left(WorldNow.Money, InStr(WorldNow.Money, "|") - 1)
				ItemZ.Value = GlobalOffer
				ItemZ.Count = GlobalOffer
				ItemZ.CanCombine = True
				ItemZ.PictureFile = Right(WorldNow.Money, Len(WorldNow.Money) - InStr(WorldNow.Money, "|"))
				If frmMain.CombineWithAnything(CreatureWithTurn, ItemZ) = False Then
					CreatureWithTurn.AddItem.Copy(ItemZ)
				End If
			End If
			' Move Item
			If IsBuying = True Then
				'UPGRADE_WARNING: Couldn't resolve default property of object BuyFrom.RemoveItem. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
				BuyFrom.RemoveItem("I" & ItemX.Index)
				ItemZ = CreatureWithTurn.AddItem
				BuyList.Remove(BuySelect)
				SellList.Add(ItemZ)
				SellTop = Greatest(SellTop, 1)
				BuySelect = BuySelect - 1
				If BuyTop + 2 > BuyList.Count() Then
					BuyTop = Greatest(BuyTop - 1, BuyList.Count())
				End If
			Else
				CreatureWithTurn.RemoveItem("I" & ItemX.Index)
				'UPGRADE_WARNING: Couldn't resolve default property of object BuyFrom.AddItem. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
				ItemZ = BuyFrom.AddItem
				SellList.Remove(SellSelect)
				BuyList.Add(ItemZ)
				BuyTop = Greatest(BuyTop, 1)
				SellSelect = SellSelect - 1
				If SellTop + 2 > SellList.Count() Then
					SellTop = Greatest(SellTop - 1, SellList.Count())
				End If
			End If
			ItemZ.Copy(ItemX)
			' Fire PostBuy Triggers
			CreatureNow = CreatureTo
			ItemNow = ItemZ
			NoFail = FireTriggers(CreatureTo, bdPostBuyItem)
			If NoFail = True Then
				CreatureNow = CreatureTo
				ItemNow = ItemZ
				NoFail = FireTriggers(ItemX, bdPostBuyItem)
				' Fire PostSell Triggers
				If NoFail = True Then
					CreatureNow = CreatureFrom
					ItemNow = ItemZ
					NoFail = FireTriggers(CreatureNow, bdPostSellItem)
					If NoFail = True Then
						CreatureNow = CreatureFrom
						ItemNow = ItemZ
						NoFail = FireTriggers(ItemX, bdPostSellItem)
					End If
				End If
			End If
			' Redisplay
			DialogBuySellShow()
		End If
	End Sub
	
	Public Sub DialogBuySellShow()
		Dim rc As Integer
		Dim n, c, Font As Short
		Dim ItemX As Item
		Dim SellPrice As Integer
		Dim Width, Height As Short
		Dim Text As String
		With frmMain
			'UPGRADE_ISSUE: PictureBox method picBuySell.Cls was not upgraded. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="CC4C7EC0-C903-48FC-ACCC-81861D12DA4A"'
            .picBuySell = Nothing
            .picBuySell.Invalidate()
			' Show Names
			.ShowText(.picBuySell, 47, 12, 284, 14, bdFontElixirWhite, (CreatureSelling.Name), 0, False)
			'    .ShowText .picBuySell, 351, 12, 284, 14, bdFontElixirWhite, CreatureWithTurn.Name, 0, False
			' Show CreatureNow's amount of Money
			' [Titi 2.4.8] let's use the currency of the world!
			Text = Left(WorldNow.Money, InStr(WorldNow.Money, "|") - 1)
			If CreatureWithTurn.Money >= 2 Then Text = Text & "s"
			' [Titi 2.4.9] prevent overlap of name and money
			c = Len(VB6.Format(CreatureWithTurn.Money, "###,###,##0") & " " & Text)
			.ShowText(.picBuySell, 351, 12, 284, 14, bdFontElixirWhite, Left(CreatureWithTurn.Name & Space(28 - c), 28 - c), 0, False)
			.ShowText(.picBuySell, 351, 12, 284, 14, bdFontElixirWhite, VB6.Format(CreatureWithTurn.Money, "###,###,##0") & " " & Text, 1, False)
			' List PC Buying Items
			If BuyList.Count() > 0 And BuyTop > 0 Then
				n = 0
				For c = BuyTop To Least(BuyTop + 2, BuyList.Count())
					ItemX = BuyList.Item(c)
					' Show Picture
					.LoadItemPic(ItemX)
					Width = ItemPicWidth(ItemX.Pic)
					Height = ItemPicHeight(ItemX.Pic)
					'UPGRADE_ISSUE: PictureBox property picItem.hdc was not upgraded. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="CC4C7EC0-C903-48FC-ACCC-81861D12DA4A"'
					'UPGRADE_ISSUE: PictureBox property picBuySell.hdc was not upgraded. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="CC4C7EC0-C903-48FC-ACCC-81861D12DA4A"'
                    rc = BBlt(.picBuySell, frmMain.picItem, 43 + (64 - Width) / 2, 48 + 102 * n + (96 - Height) / 2, Width, Height, 64 * ItemX.Pic - 64, 96, SRCAND)
					'UPGRADE_ISSUE: PictureBox property picItem.hdc was not upgraded. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="CC4C7EC0-C903-48FC-ACCC-81861D12DA4A"'
					'UPGRADE_ISSUE: PictureBox property picBuySell.hdc was not upgraded. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="CC4C7EC0-C903-48FC-ACCC-81861D12DA4A"'
                    rc = BBlt(.picBuySell, frmMain.picItem, 43 + (64 - Width) / 2, 48 + 102 * n + (96 - Height) / 2, Width, Height, 64 * ItemX.Pic - 64, 0, SRCPAINT)
					If ItemX.Capacity > 0 Then
						.ShowText(.picBuySell, 43 + (64 - Width) / 2, 48 + 102 * n + (96 - Height) / 2, Width, 10, bdFontSmallWhite, "(" & ItemX.Items.Count() & ")", 1, False)
					End If
					If ItemX.Count > 1 Then
						.ShowText(.picBuySell, 43 + (64 - Width) / 2, 48 + 102 * n + Height - 10 + (96 - Height) / 2, Width, 10, bdFontSmallWhite, CStr(ItemX.Count), True, False)
					End If
					' Show Description and Value
					If BuySelect = c Then
						Font = bdFontNoxiousGold
					Else
						Font = bdFontNoxiousWhite
					End If
					.ShowText(.picBuySell, 115, 56 + 102 * n, 214, 14, Font, (ItemX.NameText), True, False)
					.ShowText(.picBuySell, 115, 74 + 102 * n, 214, 42, Font, "Bulk " & ItemX.Bulk & "  Wgt " & ItemX.Weight, True, False)
					If ItemX.WearType < 10 Then ' Potential Armor
						Text = "Protect " & ItemX.Resistance & "%"
					ElseIf ItemX.Damage > 0 Then  ' Potential Weapon
						Text = "Damage " & (ItemX.Damage - 1) Mod 5 + 1 & "d" & Int(((ItemX.Damage - 1) Mod 25) / 5) * 2 + 4
						If ItemX.Damage - 1 > 24 Then
							Text = Text & "+" & Int((ItemX.Damage - 1) / 25)
						End If
					Else
						Text = ""
					End If
					' [Titi 2.4.8] let's use the currency of the world!
					Text = Left(WorldNow.Money, InStr(WorldNow.Money, "|") - 1)
					If ItemX.Value >= 2 Then Text = Text & "s"
					.ShowText(.picBuySell, 115, 92 + 102 * n, 214, 14, Font, Text, True, False)
					.ShowText(.picBuySell, 115, 128 + 102 * n, 214, 14, Font, VB6.Format(CInt(ItemX.Value + ItemX.Value * BuyRate), "###,###,##0") & " " & Text, True, False)
					n = n + 1
				Next c
			End If
			.ScrollBarShow(.picBuySell, 23, 44, 312, BuyTop, BuyList.Count() - 2, 0)
			' List PC Selling Items
			If SellList.Count() > 0 And SellTop > 0 Then
				n = 0
				For c = SellTop To Least(SellTop + 2, SellList.Count())
					ItemX = SellList.Item(c)
					' Show Picture
					.LoadItemPic(ItemX)
					Width = ItemPicWidth(ItemX.Pic)
					Height = ItemPicHeight(ItemX.Pic)
					'UPGRADE_ISSUE: PictureBox property picItem.hdc was not upgraded. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="CC4C7EC0-C903-48FC-ACCC-81861D12DA4A"'
					'UPGRADE_ISSUE: PictureBox property picBuySell.hdc was not upgraded. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="CC4C7EC0-C903-48FC-ACCC-81861D12DA4A"'
                    rc = BBlt(.picBuySell, frmMain.picItem, 578 + (64 - Width) / 2, 48 + 102 * n + (96 - Height) / 2, Width, Height, 64 * ItemX.Pic - 64, 96, SRCAND)
					'UPGRADE_ISSUE: PictureBox property picItem.hdc was not upgraded. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="CC4C7EC0-C903-48FC-ACCC-81861D12DA4A"'
					'UPGRADE_ISSUE: PictureBox property picBuySell.hdc was not upgraded. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="CC4C7EC0-C903-48FC-ACCC-81861D12DA4A"'
                    rc = BBlt(.picBuySell, frmMain.picItem, 578 + (64 - Width) / 2, 48 + 102 * n + (96 - Height) / 2, Width, Height, 64 * ItemX.Pic - 64, 0, SRCPAINT)
					If ItemX.Capacity > 0 Then
						.ShowText(.picBuySell, 578 + (64 - Width) / 2, 48 + 102 * n + (96 - Height) / 2, Width, 10, bdFontSmallWhite, "(" & ItemX.Items.Count() & ")", 1, False)
					End If
					If ItemX.Count > 1 Then
						.ShowText(.picBuySell, 578 + (64 - Width) / 2, 48 + 102 * n + Height - 10 + (96 - Height) / 2, Width, 10, bdFontSmallWhite, CStr(ItemX.Count), True, False)
					End If
					' Show Description and Value
					If SellSelect = c Then
						Font = bdFontNoxiousGold
					Else
						Font = bdFontNoxiousWhite
					End If
					' Show if wearing
					If ItemX.IsReady = True Then
						.ShowText(.picBuySell, 357, 56 + 102 * n, 214, 14, Font, ItemX.NameText & " (Worn)", True, False)
					Else
						.ShowText(.picBuySell, 357, 56 + 102 * n, 214, 14, Font, (ItemX.NameText), True, False)
					End If
					.ShowText(.picBuySell, 357, 74 + 102 * n, 214, 42, Font, "Bulk " & ItemX.Bulk & "  Wgt " & ItemX.Weight, True, False)
					If ItemX.WearType < 10 Then ' Potential Armor
						Text = "Protect " & ItemX.Resistance & "%"
					ElseIf ItemX.Damage > 0 Then  ' Potential Weapon
						Text = "Damage " & (ItemX.Damage - 1) Mod 5 + 1 & "d" & Int(((ItemX.Damage - 1) Mod 25) / 5) * 2 + 4
						If ItemX.Damage - 1 > 24 Then
							Text = Text & "+" & Int((ItemX.Damage - 1) / 25)
						End If
					Else
						Text = ""
					End If
					' [Titi 2.4.8] let's use the currency of the world!
					Text = Left(WorldNow.Money, InStr(WorldNow.Money, "|") - 1)
					If ItemX.Value >= 2 Then Text = Text & "s"
					.ShowText(.picBuySell, 357, 92 + 102 * n, 214, 14, Font, Text, True, False)
					.ShowText(.picBuySell, 357, 128 + 102 * n, 214, 14, Font, VB6.Format(CInt(ItemX.Value - ItemX.Value * BuyRate), "###,###,##0") & " " & Text, True, False)
					n = n + 1
				Next c
			End If
			.ScrollBarShow(.picBuySell, 644, 44, 312, SellTop, SellList.Count() - 2, 0)
			' Show Buttons
			.ShowButton(.picBuySell, 41, .picBuySell.ClientRectangle.Height - 30, "Buy", False)
			.ShowButton(.picBuySell, 351, .picBuySell.ClientRectangle.Height - 30, "Sell", False)
			.ShowButton(.picBuySell, .picBuySell.ClientRectangle.Width - 102, .picBuySell.ClientRectangle.Height - 30, "Done", False)
			.picBuySell.BringToFront()
			.picBuySell.Visible = True
		End With
	End Sub
	
	Public Sub DialogBuySellPC()
		' Set up CreatureWithTurn to sell stuff
		Dim ItemX, ItemZ As Item
		SellList = New Collection
		For	Each ItemX In CreatureWithTurn.Items
			If ItemX.IsMoney = False Then
				SellList.Add(ItemX)
			End If
		Next ItemX
		SellTop = Least(1, SellList.Count()) : SellSelect = SellTop
	End Sub
	
	Public Sub DialogDice(ByRef CreatureX As Object, ByRef Text As String, ByRef AsDice As Short, ByRef TotalDie As Short)
        Dim PosY As Short
        Dim OldPointer As Cursor
		Dim rc As Integer
		With frmMain
			'UPGRADE_WARNING: Screen property Screen.MousePointer has a new behavior. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6BA9B8D2-2A32-4B6E-8D36-44949974A5B4"'
			OldPointer = System.Windows.Forms.Cursor.Current
			'UPGRADE_WARNING: Screen property Screen.MousePointer has a new behavior. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6BA9B8D2-2A32-4B6E-8D36-44949974A5B4"'
			System.Windows.Forms.Cursor.Current = System.Windows.Forms.Cursors.Default
			.picConvoList.Visible = False
			' Display Character portrait
			'UPGRADE_ISSUE: PictureBox method picConvo.Cls was not upgraded. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="CC4C7EC0-C903-48FC-ACCC-81861D12DA4A"'
            .picConvo = Nothing
            .picConvo.Invalidate()
			DialogShowFace(CreatureX)
			.ShowText(.picConvo, 0, 12, .picConvo.Width, 14, bdFontElixirWhite, "Roll Dice", True, False)
			PosY = .ShowText(.picConvo, 88, 34, 354, 430, bdFontNoxiousWhite, Text, False, True)
			.picConvo.Height = Greatest(PosY + 90, 146)
			' Paint Buttons
			DialogSetButton(1, "Roll")
			.ShowButton(.picConvo, .picConvo.Width - 102, .picConvo.Height - 30, ButtonList(0), False)
			' Paint bottom of picConvo
			'UPGRADE_ISSUE: PictureBox property picConvoBottom.hdc was not upgraded. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="CC4C7EC0-C903-48FC-ACCC-81861D12DA4A"'
			'UPGRADE_ISSUE: PictureBox property picConvo.hdc was not upgraded. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="CC4C7EC0-C903-48FC-ACCC-81861D12DA4A"'
            rc = BBlt(.picConvo, .picConvoBottom, 0, .picConvo.Height - 6, .picConvo.Width, 6, 0, 0, SRCCOPY)
			.picConvo.Refresh()
			' Center picConvo
			.picConvo.Top = (.picBox.ClientRectangle.Height - .picConvo.ClientRectangle.Height) / 2
			.picConvo.Left = (.ClientRectangle.Width - .picConvo.ClientRectangle.Width) / 2
			ConvoAction = -1
			.picConvo.Visible = True
			.picConvo.BringToFront()
			.picConvo.Focus()
			Frozen = True
			Do Until ConvoAction > -1
				Do Until ConvoAction > -1
					System.Windows.Forms.Application.DoEvents()
				Loop 
				If ButtonList(0) = "Roll" Then
					'UPGRADE_WARNING: Screen property Screen.MousePointer has a new behavior. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6BA9B8D2-2A32-4B6E-8D36-44949974A5B4"'
					System.Windows.Forms.Cursor.Current = System.Windows.Forms.Cursors.WaitCursor
					If AsDice < 24 Then
						TotalDie = .DiceRoll(AsDice Mod 5 + 1, Int(AsDice / 5) * 2 + 4, 0, False, False)
					Else
						TotalDie = .DiceRoll(1, 20, 0, False, False)
					End If
					'UPGRADE_WARNING: Screen property Screen.MousePointer has a new behavior. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6BA9B8D2-2A32-4B6E-8D36-44949974A5B4"'
					System.Windows.Forms.Cursor.Current = System.Windows.Forms.Cursors.Default
					DialogSetButton(1, "Ok")
					.ShowButton(.picConvo, .picConvo.Width - 102, .picConvo.Height - 30, ButtonList(0), False)
					.picConvo.Refresh()
					ConvoAction = -1
				End If
			Loop 
			Frozen = False
			'UPGRADE_ISSUE: Screen property Screen.MousePointer does not support custom mousepointers. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="45116EAB-7060-405E-8ABE-9DBB40DC2E86"'
			'UPGRADE_WARNING: Screen property Screen.MousePointer has a new behavior. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6BA9B8D2-2A32-4B6E-8D36-44949974A5B4"'
			System.Windows.Forms.Cursor.Current = OldPointer
		End With
	End Sub
	
	Public Function DialogItem(ByRef ItemX As Item, ByRef Text As String) As Short
        Dim OldPointer As Cursor
		Dim rc As Integer
		With frmMain
			' Show Item and description
			'UPGRADE_WARNING: Screen property Screen.MousePointer has a new behavior. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6BA9B8D2-2A32-4B6E-8D36-44949974A5B4"'
			OldPointer = System.Windows.Forms.Cursor.Current
			'UPGRADE_WARNING: Screen property Screen.MousePointer has a new behavior. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6BA9B8D2-2A32-4B6E-8D36-44949974A5B4"'
			System.Windows.Forms.Cursor.Current = System.Windows.Forms.Cursors.Default
			DialogSetUp(modGameGeneral.DLGTYPE.bdDlgItem)
			.ShowText(.picConvo, 0, 12, .picConvo.Width, 14, bdFontElixirWhite, (ItemX.Name), True, False)
			.ShowText(.picConvo, 56, 34, 380, 102, bdFontNoxiousGold, "Bulk " & ItemX.Bulk & " Weight " & ItemX.Weight, False, False)
			.ShowText(.picConvo, 56, 50, 380, 102, bdFontNoxiousWhite, Text, False, False)
			.picConvo.Top = (.picBox.ClientRectangle.Height - .picConvo.ClientRectangle.Height) / 2
			.picConvo.Left = (.ClientRectangle.Width - .picConvo.ClientRectangle.Width) / 2
			' Show ItemX's picture
			.LoadItemPic(ItemX)
			'UPGRADE_ISSUE: PictureBox property picItem.hdc was not upgraded. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="CC4C7EC0-C903-48FC-ACCC-81861D12DA4A"'
			'UPGRADE_ISSUE: PictureBox property picConvo.hdc was not upgraded. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="CC4C7EC0-C903-48FC-ACCC-81861D12DA4A"'
            rc = BBlt(.picConvo, .picItem, 20, 34, ItemPicWidth(ItemX.Pic) / 3, ItemPicHeight(ItemX.Pic) / 3, 64 * ItemX.Pic - 32, 96 * 2, SRCAND)
			'UPGRADE_ISSUE: PictureBox property picItem.hdc was not upgraded. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="CC4C7EC0-C903-48FC-ACCC-81861D12DA4A"'
			'UPGRADE_ISSUE: PictureBox property picConvo.hdc was not upgraded. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="CC4C7EC0-C903-48FC-ACCC-81861D12DA4A"'
            rc = BBlt(.picConvo, .picItem, 20, 34, ItemPicWidth(ItemX.Pic) / 3, ItemPicHeight(ItemX.Pic) / 3, 64 * ItemX.Pic - 64, 96 * 2, SRCPAINT)
			' Paint Buttons
			DialogSetButton(1, "Done")
			.ShowButton(.picConvo, .picConvo.Width - 6 - 96, .picConvo.Height - 30, "Done", False)
			If .picInventory.Visible = True Then
				If ItemX.Count > 1 Then
					DialogSetButton(2, "Split")
					.ShowButton(.picConvo, .picConvo.Width - 6 - 96 * 2, .picConvo.Height - 30, "Split", False)
				Else
					DialogSetButton(2, "Use")
					.ShowButton(.picConvo, .picConvo.Width - 6 - 96 * 2, .picConvo.Height - 30, "Use", False)
				End If
			End If
			' Paint bottom of picConvo
			'UPGRADE_ISSUE: PictureBox property picTomeNew.hdc was not upgraded. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="CC4C7EC0-C903-48FC-ACCC-81861D12DA4A"'
			'UPGRADE_ISSUE: PictureBox property picConvo.hdc was not upgraded. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="CC4C7EC0-C903-48FC-ACCC-81861D12DA4A"'
            rc = BBlt(.picConvo, .picTomeNew, 0, .picConvo.Height - 6, .picConvo.Width, 6, 0, .picTomeNew.Height - 6, SRCCOPY)
			.picConvo.Refresh()
			.picConvo.Visible = True
			.picConvo.BringToFront()
			.picConvo.Focus()
			ConvoAction = -1
			Frozen = True
			'UPGRADE_WARNING: Screen property Screen.MousePointer has a new behavior. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6BA9B8D2-2A32-4B6E-8D36-44949974A5B4"'
			System.Windows.Forms.Cursor.Current = System.Windows.Forms.Cursors.Default
			Do Until ConvoAction > -1
				System.Windows.Forms.Application.DoEvents()
			Loop 
			' Use Item?
			If ConvoAction = 0 And .picInventory.Visible = True Then
				.picConvo.Visible = False
				.InventoryClose()
				.UseShow(ItemX)
			End If
			DialogItem = ConvoAction
			'UPGRADE_ISSUE: Screen property Screen.MousePointer does not support custom mousepointers. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="45116EAB-7060-405E-8ABE-9DBB40DC2E86"'
			'UPGRADE_WARNING: Screen property Screen.MousePointer has a new behavior. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6BA9B8D2-2A32-4B6E-8D36-44949974A5B4"'
			System.Windows.Forms.Cursor.Current = OldPointer
		End With
	End Function
End Module