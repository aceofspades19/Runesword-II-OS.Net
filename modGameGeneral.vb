Option Strict Off
Option Explicit On
Module modGameGeneral
	
	' Nasty Nasty variable be vawie vawie carefull with this one
	Public Frozen As Short
	
	' Reference to current Tome
	'Public Tome As New Tome
	' Reference to current Plot
	'Public Area As New Area
	' Reference to current Map in Plot
	Public Map As New Map
	
	' Reference to current Encounter in Map
	Public EncounterNow As New Encounter
	' Always points to the current Creature taking a turn
	Public CreatureWithTurn As New Creature
	' Points to Creature Selling objects
	Public CreatureSelling As Creature
	' Below are all used by Triggers ONLY
	Public CreatureNow As New Creature
	Public CreatureTarget As New Creature
	Public ItemNow As New Item
	Public ItemTarget As New Item
	
	' Creatures now in Menu
	Public MenuParty(4) As Creature
	Public MenuPartyIndex As Short
	' Stack of Text for DialogBrief
	Public DialogBriefSet As New Collection
	
	' Font related variables
	Public aX(3, 89) As Short
	Public aY(3, 89) As Short
	Public aW(3, 89) As Short
	Public Const bdFontSmallWhite As Short = 0
	Public Const bdFontNoxiousWhite As Short = 1
	Public Const bdFontNoxiousGold As Short = 2
	Public Const bdFontNoxiousGrey As Short = 3
	Public Const bdFontNoxiousBlack As Short = 4
	Public Const bdFontElixirWhite As Short = 5
	Public Const bdFontElixirGold As Short = 6
	Public Const bdFontElixirGrey As Short = 7
	Public Const bdFontElixirBlack As Short = 8
	Public Const bdFontLargeWhite As Short = 9
	
	' Variables to store Create Character information
	Public TomeList As Collection
	Public TomeRoster As Collection
	Public TomeRosterSelect As Short
	Public TomeMenu As Short
	Public TomeAction As Short
	Public TomeIndex As Short
	Public TomeRosterTop As Short
	Public TomeRosterDragX, TomeRosterDrag, TomeRosterDragY As Short
	Public TomeListTop As Short
	Public Worlds As Collection
	' Public WorldNow As World
	' moved to modBD for compatibility with world settings creation
	Public KingdomNow As Kingdom
	Public KingdomIndex As Short
	' Public KingdomNames(11) As String
	' moved to modBD for compatibility with world settings creation
	Public CreateNameNew As String
	Public CreateSkillsIndex As Short
	Public CreateNameIndex As Short
	Public CreatePictureIndex As Short
	Public CreatePCReturn As Short
	Public TomeSavePos As Short
	Public TomeSavePathName As String
	Public CombatAction1, CombatAction2 As Short
	
	Public Const bdTomeNew As Short = 0
	Public Const bdTomeGather As Short = 1
	Public Const bdCreatePCKingdom As Short = 2
	Public Const bdCreatePCPicture As Short = 3
	Public Const bdCreatePCSkills As Short = 4
	Public Const bdCreatePCName As Short = 5
	Public Const bdTomeSaves As Short = 6
	Public Const bdTomeOptions As Short = 7
	Public Const bdTomeSaveAs As Short = 8
	Public Const bdMenuSkill As Short = 9
	
	' Game Options
	'Public Const bdOptionsCount As Integer = 13
	Public Const bdOptionsCount As Short = 17
	Public GlobalInterfaceList As Collection
	Public GlobalInterfaceIndex As Short
	Public GlobalDiceList As Collection
	Public GlobalDiceIndex As Short
	' [Titi 2.4.9] ASCII code of shortcuts
	Public bdKey(25) As Short
	' [Titi 2.4.9] Automatic centering of the map on the party?
	Public GlobalAutoCenter As Short
	
	' Constants for Misc Interface Pieces
	Public Const bdIntWidth As Short = 90
	Public Const bdIntHeight As Short = 16
	
	' Variables for Tome Wizard
	Public TomeWizardParty As Short
	Public TomeWizardLevel As Short
	Public TomeWizardStory As Short
	Public TomeWizardSize As Short
	Public UberWizMaps As UberWizard
	Public BuildMessage As String
	
	' Scratch variables
	Public Darker As Short
	Public DarkDir As Short
	
	' Scroll Box Variables
	Public ScrollThumbY As Short
	Public ScrollTop As Short
	Public ScrollSelect As Short
	Public ScrollList As Collection
	Public ScrollList2 As Collection
	Public JournalThumbY As Short
	Public JournalTop As Short
	Public JournalList As Collection
	
	' Sorcery variables
	Public GlobalSaveStyle As Short
	
	' Create Character variables
	Public Const bdCPCShowStats As Short = 0
	Public Const bdCPCShowSkills As Short = 1
	Public Const bdCPCShowNoName As Short = 2
	Public Const bdFaceDM As Short = 198
	Public Const bdFaceBlank As Short = 132
	Public Const bdFaceScroll As Short = 264
	Public Const bdFaceSelect As Short = 330
	Public Const bdFaceMin As Short = 462
	
	' Inventory variables
	Public InvX(128) As Short
	Public InvY(128) As Short
	Public InvZ(128) As Short
	Public InvC(80) As Short
	Public InvDragMode As Short
	Public InvDragItem As New Item
	Public InvDragPos As Short
	Public InvDragIndex As Short
	Public InvContainer As Item
	Public InvNowShow As Short
	Public InvTopIndex As Short
	Public InvTooMany As Short
	
	Public Const bdInvItems As Short = 0
	Public Const bdInvStatus As Short = 1
	Public Const bdInvObjects As Short = 0
	Public Const bdInvWear As Short = 1
	Public Const bdInvContainer As Short = 2
	Public Const bdInvEncounter As Short = 3
	Public Const bdInvParty As Short = 4
	
	'Public Const bdDlgNoReply As Integer = 0
	'Public Const bdDlgWithReply As Integer = 1
	'Public Const bdDlgItemList As Integer = 2
	'Public Const bdDlgWithDice As Integer = 3
	'Public Const bdDlgReplyText As Integer = 4
	'Public Const bdDlgItem As Integer = 6
	'Public Const bdDlgSave As Integer = 7
	'Public Const bdDlgCreatureList As Integer = 8
	'Public Const bdDlgDebug As Integer = 9
	'Public Const bdDlgCredits As Integer = 10
	Public Enum DLGTYPE
		bdDlgNoReply = 0
		bdDlgWithReply = 1
		bdDlgItemList = 2
		bdDlgWithDice = 3
		bdDlgReplyText = 4
		bdDlgItem = 6
		bdDlgSave = 7
		bdDlgCreatureList = 8
		bdDlgDebug = 9
		bdDlgCredits = 10
	End Enum
	
	Public Const bdSellStyleInv As Short = 0
	Public Const bdSellStyleEnc As Short = 1
	
	Public Function OpName(ByVal Index As Short) As String
		Select Case Index
			Case 0 ' =
				OpName = "EqualTo"
			Case 1 ' +
				OpName = "Plus"
			Case 2 ' -
				OpName = "Minus"
			Case 3 ' *
				OpName = "X"
			Case 4 ' /
				OpName = "DividedBy"
			Case 5 ' >
				OpName = "GreaterThan"
			Case 6 ' <
				OpName = "LessThan"
			Case 7 ' >=
				OpName = "GreaterOrEqual"
			Case 8 ' <=
				OpName = "LessOrEqual"
			Case 9 ' Or
				OpName = "Or"
			Case 10 ' And
				OpName = "And"
			Case 11 ' Xor
				OpName = "XOr"
			Case 12 ' Like
				OpName = "Like"
		End Select
	End Function
	
	Public Sub OptionsLoad(ByRef Refresh As Short)
		' Revised for 2-4-6 by Count0 to handle no options file present and outdated options file present cases
		Dim GlobalFlags, FromFile, c As Short
		Dim Cnt As Integer
		'Dim FromFile As Integer, GlobalFlags As Integer, Cnt As Integer, c As Integer
		Dim sPath As String
		Dim sVersion As String
		On Error GoTo Err_Handler
		'    If Not oFileSys.CheckExists(gAppPath & "\Data", Folder) Then
		If Not oFileSys.CheckExists(gDataPath, clsInOut.IOActionType.Folder) Then
			Err.Raise(1000, "Engine.OptionsLoad", "No Data Folder found.")
		End If
		'    sPath = gAppPath & "\data\stock\"
		sPath = gDataPath & "\stock\"
		If Not oFileSys.CheckExists(sPath & "options.dat", clsInOut.IOActionType.File, False) Then OptionsFileDefault() : Exit Sub
		' check for options file version for this build
		sVersion = My.Application.Info.Version.Major & "." & My.Application.Info.Version.Minor & "." & My.Application.Info.Version.Revision
		FromFile = FreeFile
		' Now we open it for real and read from it.
		FileOpen(FromFile, sPath & "options.dat", OpenMode.Binary)
		'UPGRADE_WARNING: Get was upgraded to FileGet and has a new behavior. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="9B7D5ADD-D8FE-4819-A36C-6DEDAF088CC7"'
		FileGet(FromFile, Cnt)
		If Cnt > 1 Then
			GlobalLicNumber1 = ""
			For c = 1 To Cnt
				GlobalLicNumber1 = GlobalLicNumber1 & " "
			Next c
			GlobalMusicRandom = 1 ' default to ON
		Else
			GlobalMusicRandom = Cnt
		End If
		'UPGRADE_WARNING: Get was upgraded to FileGet and has a new behavior. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="9B7D5ADD-D8FE-4819-A36C-6DEDAF088CC7"'
		FileGet(FromFile, GlobalLicNumber1)
		'UPGRADE_WARNING: Get was upgraded to FileGet and has a new behavior. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="9B7D5ADD-D8FE-4819-A36C-6DEDAF088CC7"'
		FileGet(FromFile, GlobalMusicState)
		'UPGRADE_WARNING: Get was upgraded to FileGet and has a new behavior. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="9B7D5ADD-D8FE-4819-A36C-6DEDAF088CC7"'
		FileGet(FromFile, GlobalRightClick)
		'UPGRADE_WARNING: Get was upgraded to FileGet and has a new behavior. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="9B7D5ADD-D8FE-4819-A36C-6DEDAF088CC7"'
		FileGet(FromFile, GlobalGameSpeed)
		'UPGRADE_WARNING: Get was upgraded to FileGet and has a new behavior. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="9B7D5ADD-D8FE-4819-A36C-6DEDAF088CC7"'
		FileGet(FromFile, GlobalDiceRolling)
		'UPGRADE_WARNING: Get was upgraded to FileGet and has a new behavior. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="9B7D5ADD-D8FE-4819-A36C-6DEDAF088CC7"'
		FileGet(FromFile, GlobalWAVState)
		'UPGRADE_WARNING: Get was upgraded to FileGet and has a new behavior. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="9B7D5ADD-D8FE-4819-A36C-6DEDAF088CC7"'
		FileGet(FromFile, GlobalDebugMode)
		'UPGRADE_WARNING: Get was upgraded to FileGet and has a new behavior. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="9B7D5ADD-D8FE-4819-A36C-6DEDAF088CC7"'
		FileGet(FromFile, GlobalOverSwing)
		'UPGRADE_WARNING: Get was upgraded to FileGet and has a new behavior. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="9B7D5ADD-D8FE-4819-A36C-6DEDAF088CC7"'
		FileGet(FromFile, GlobalFlags)
		'UPGRADE_WARNING: Get was upgraded to FileGet and has a new behavior. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="9B7D5ADD-D8FE-4819-A36C-6DEDAF088CC7"'
		FileGet(FromFile, Cnt)
		GlobalLicNumber2 = ""
		For c = 1 To Cnt
			GlobalLicNumber2 = GlobalLicNumber2 & " "
		Next c
		'UPGRADE_WARNING: Get was upgraded to FileGet and has a new behavior. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="9B7D5ADD-D8FE-4819-A36C-6DEDAF088CC7"'
		FileGet(FromFile, GlobalLicNumber2)
		' Read GlobalInterfaceName
		'UPGRADE_WARNING: Get was upgraded to FileGet and has a new behavior. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="9B7D5ADD-D8FE-4819-A36C-6DEDAF088CC7"'
		FileGet(FromFile, Cnt)
		GlobalInterfaceName = ""
		For c = 1 To Cnt
			GlobalInterfaceName = GlobalInterfaceName & " "
		Next c
		'UPGRADE_WARNING: Get was upgraded to FileGet and has a new behavior. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="9B7D5ADD-D8FE-4819-A36C-6DEDAF088CC7"'
		FileGet(FromFile, GlobalInterfaceName)
		' Read GlobalDiceName
		'UPGRADE_WARNING: Get was upgraded to FileGet and has a new behavior. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="9B7D5ADD-D8FE-4819-A36C-6DEDAF088CC7"'
		FileGet(FromFile, Cnt)
		GlobalDiceName = ""
		For c = 1 To Cnt
			GlobalDiceName = GlobalDiceName & " "
		Next c
		'UPGRADE_WARNING: Get was upgraded to FileGet and has a new behavior. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="9B7D5ADD-D8FE-4819-A36C-6DEDAF088CC7"'
		FileGet(FromFile, GlobalDiceName)
		'UPGRADE_WARNING: Get was upgraded to FileGet and has a new behavior. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="9B7D5ADD-D8FE-4819-A36C-6DEDAF088CC7"'
		FileGet(FromFile, Cnt)
		GlobalLicNumber3 = ""
		For c = 1 To Cnt + 1 ' last character will be GlobalCredits status
			GlobalLicNumber3 = GlobalLicNumber3 & " "
		Next c
		'UPGRADE_WARNING: Get was upgraded to FileGet and has a new behavior. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="9B7D5ADD-D8FE-4819-A36C-6DEDAF088CC7"'
		FileGet(FromFile, GlobalLicNumber3)
		GlobalCredits = Right(GlobalLicNumber3, 1) ' now retrieve GlobalCredits value
		GlobalLicNumber3 = Left(GlobalLicNumber3, Cnt) ' and set GlobalLicNumber3 to its correct value
		' [Titi 2.4.9] keyboard shortcuts
		For c = 1 To 25
			'UPGRADE_WARNING: Get was upgraded to FileGet and has a new behavior. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="9B7D5ADD-D8FE-4819-A36C-6DEDAF088CC7"'
			FileGet(FromFile, bdKey(c))
		Next 
		FileClose(FromFile)
		' Decompose GlobalMouseClick and GlobalAutoEndTurn
		GlobalMouseClick = System.Math.Abs(CInt((GlobalFlags And &H1) > 0))
		GlobalAutoEndTurn = System.Math.Abs(CInt((GlobalFlags And &H2) > 0))
		GlobalFastMove = System.Math.Abs(CInt((GlobalFlags And &H4) > 0))
		GlobalAutoCenter = System.Math.Abs(CInt((GlobalFlags And &H8) > 0)) ' [Titi 2.4.9]
		' Double check GlobalInterfaceName
		'    If Not oFileSys.CheckExists(gAppPath & "\Data\Interface\" & GlobalInterfaceName, Folder) Then GlobalInterfaceName = "WoodWeave"
		If Not oFileSys.CheckExists(gDataPath & "\Interface\" & GlobalInterfaceName, clsInOut.IOActionType.Folder) Then GlobalInterfaceName = "WoodWeave"
		' Double check GlobalDiceName
		'    If Not oFileSys.CheckExists(gAppPath & "\Data\Interface\Dice\" & GlobalDiceName, File) Then GlobalDiceName = "Bone White.bmp"
		If Not oFileSys.CheckExists(gDataPath & "\Interface\Dice\" & GlobalDiceName, clsInOut.IOActionType.File) Then GlobalDiceName = "Bone White.bmp"
Exit_Sub: 
		Exit Sub
Err_Handler: 
		oErr.logError("You may not have the required data files installed.  Check for a Data folder.", CErrorHandler.ErrorLevel.ERR_CRITICAL)
		'Stop
		Resume Exit_Sub
	End Sub
	
	Private Sub OptionsFileDefault()
		' for 2-4-6 revision, by Count0
		' set what would be in contents of options file when not available
		Call oErr.LogText("setting game default game options since no options to load")
		GlobalWAVState = 1
		GlobalGameSpeed = 12
		GlobalMusicState = 1
		GlobalDebugMode = 0
		GlobalRightClick = 1
		GlobalDiceRolling = 1
		GlobalOverSwing = 1
		GlobalMouseClick = 1
		GlobalAutoEndTurn = 0
		GlobalFastMove = 0
		GlobalAutoCenter = 0 ' [Titi 2.4.9]
		' the following are already set by gameInit()
		' GlobalScreenX
		' GlobalScreenY
		' GlobalScreenColor
		GlobalInterfaceName = "WoodWeave"
		GlobalDiceName = "Bone White.bmp"
		GlobalLicNumber1 = "0000"
		' GlobalLicNumber2
		' GlobalLicNumber3
		GlobalCredits = "0"
		GlobalMusicRandom = 1
		InitKeyboardShortCuts()
	End Sub
	
	Public Sub OptionsSave()
		' Requires revision
		Dim ToFile, GlobalFlags As Short
		Dim c As Short
		' Compose GlobalFlags
		If GlobalMouseClick = 1 Then
			GlobalFlags = GlobalFlags Or &H1
		End If
		If GlobalAutoEndTurn = 1 Then
			GlobalFlags = GlobalFlags Or &H2
		End If
		If GlobalFastMove = 1 Then
			GlobalFlags = GlobalFlags Or &H4
		End If
		If GlobalAutoCenter = 1 Then ' [Titi 2.4.9]
			GlobalFlags = GlobalFlags Or &H8
		End If
		' Save options file
		ToFile = FreeFile
		'    Open gAppPath & "\data\stock\options.dat" For Binary As ToFile
		FileOpen(ToFile, gDataPath & "\stock\options.dat", OpenMode.Binary)
		'Put ToFile, , Len(GlobalLicNumber1)
		'UPGRADE_WARNING: Put was upgraded to FilePut and has a new behavior. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="9B7D5ADD-D8FE-4819-A36C-6DEDAF088CC7"'
		FilePut(ToFile, CInt(GlobalMusicRandom))
		'UPGRADE_WARNING: Put was upgraded to FilePut and has a new behavior. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="9B7D5ADD-D8FE-4819-A36C-6DEDAF088CC7"'
		FilePut(ToFile, "") 'GlobalLicNumber1
		'UPGRADE_WARNING: Put was upgraded to FilePut and has a new behavior. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="9B7D5ADD-D8FE-4819-A36C-6DEDAF088CC7"'
		FilePut(ToFile, GlobalMusicState)
		'UPGRADE_WARNING: Put was upgraded to FilePut and has a new behavior. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="9B7D5ADD-D8FE-4819-A36C-6DEDAF088CC7"'
		FilePut(ToFile, GlobalRightClick)
		'UPGRADE_WARNING: Put was upgraded to FilePut and has a new behavior. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="9B7D5ADD-D8FE-4819-A36C-6DEDAF088CC7"'
		FilePut(ToFile, GlobalGameSpeed)
		'UPGRADE_WARNING: Put was upgraded to FilePut and has a new behavior. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="9B7D5ADD-D8FE-4819-A36C-6DEDAF088CC7"'
		FilePut(ToFile, GlobalDiceRolling)
		'UPGRADE_WARNING: Put was upgraded to FilePut and has a new behavior. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="9B7D5ADD-D8FE-4819-A36C-6DEDAF088CC7"'
		FilePut(ToFile, GlobalWAVState)
		'UPGRADE_WARNING: Put was upgraded to FilePut and has a new behavior. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="9B7D5ADD-D8FE-4819-A36C-6DEDAF088CC7"'
		FilePut(ToFile, GlobalDebugMode)
		'UPGRADE_WARNING: Put was upgraded to FilePut and has a new behavior. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="9B7D5ADD-D8FE-4819-A36C-6DEDAF088CC7"'
		FilePut(ToFile, GlobalOverSwing)
		'UPGRADE_WARNING: Put was upgraded to FilePut and has a new behavior. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="9B7D5ADD-D8FE-4819-A36C-6DEDAF088CC7"'
		FilePut(ToFile, GlobalFlags)
		'UPGRADE_WARNING: Put was upgraded to FilePut and has a new behavior. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="9B7D5ADD-D8FE-4819-A36C-6DEDAF088CC7"'
		FilePut(ToFile, Len(GlobalLicNumber2))
		'UPGRADE_WARNING: Put was upgraded to FilePut and has a new behavior. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="9B7D5ADD-D8FE-4819-A36C-6DEDAF088CC7"'
		FilePut(ToFile, GlobalLicNumber2)
		' Write GlobalInterfaceName
		'UPGRADE_WARNING: Put was upgraded to FilePut and has a new behavior. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="9B7D5ADD-D8FE-4819-A36C-6DEDAF088CC7"'
		FilePut(ToFile, Len(GlobalInterfaceName))
		'UPGRADE_WARNING: Put was upgraded to FilePut and has a new behavior. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="9B7D5ADD-D8FE-4819-A36C-6DEDAF088CC7"'
		FilePut(ToFile, GlobalInterfaceName)
		' Write GlobalDiceName
		'UPGRADE_WARNING: Put was upgraded to FilePut and has a new behavior. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="9B7D5ADD-D8FE-4819-A36C-6DEDAF088CC7"'
		FilePut(ToFile, Len(GlobalDiceName))
		'UPGRADE_WARNING: Put was upgraded to FilePut and has a new behavior. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="9B7D5ADD-D8FE-4819-A36C-6DEDAF088CC7"'
		FilePut(ToFile, GlobalDiceName)
		'UPGRADE_WARNING: Put was upgraded to FilePut and has a new behavior. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="9B7D5ADD-D8FE-4819-A36C-6DEDAF088CC7"'
		FilePut(ToFile, Len(GlobalLicNumber3))
		'UPGRADE_WARNING: Put was upgraded to FilePut and has a new behavior. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="9B7D5ADD-D8FE-4819-A36C-6DEDAF088CC7"'
		FilePut(ToFile, GlobalLicNumber3)
		'UPGRADE_WARNING: Put was upgraded to FilePut and has a new behavior. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="9B7D5ADD-D8FE-4819-A36C-6DEDAF088CC7"'
		FilePut(ToFile, GlobalCredits)
		' [Titi 2.4.9] keyboard shortcuts
		For c = 1 To 25
			'UPGRADE_WARNING: Put was upgraded to FilePut and has a new behavior. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="9B7D5ADD-D8FE-4819-A36C-6DEDAF088CC7"'
			FilePut(ToFile, bdKey(c))
		Next 
		FileClose(ToFile)
	End Sub
	
	Public Sub InitKeyboardShortCuts()
		' [Titi 2.4.9] init keyboard shortcuts
		' 200 means CTRL, 100 means SHIFT
		bdKey(1) = 200 + System.Windows.Forms.Keys.D ' Debug
		bdKey(2) = 200 + System.Windows.Forms.Keys.Q ' Quit
		bdKey(3) = 200 + System.Windows.Forms.Keys.X ' Exit
		bdKey(4) = 200 + System.Windows.Forms.Keys.Z ' Speed
		bdKey(5) = 200 + System.Windows.Forms.Keys.L ' Load
		bdKey(6) = 200 + System.Windows.Forms.Keys.S ' Save
		bdKey(7) = 200 + System.Windows.Forms.Keys.M ' Music
		bdKey(8) = 200 + System.Windows.Forms.Keys.W ' Sound
		bdKey(9) = 100 + System.Windows.Forms.Keys.Left ' Scroll map left
		bdKey(10) = 100 + System.Windows.Forms.Keys.Right ' Scroll map right
		bdKey(11) = 100 + System.Windows.Forms.Keys.Up ' Scroll map up
		bdKey(12) = 100 + System.Windows.Forms.Keys.Down ' Scroll map down
		bdKey(13) = System.Windows.Forms.Keys.Left ' Move party left
		bdKey(14) = System.Windows.Forms.Keys.Right ' Move party right
		bdKey(15) = System.Windows.Forms.Keys.Up ' Move party up
		bdKey(16) = System.Windows.Forms.Keys.Down ' Move party down
		bdKey(17) = System.Windows.Forms.Keys.J ' Journals and Quests
		bdKey(18) = System.Windows.Forms.Keys.L ' Listen
		bdKey(19) = System.Windows.Forms.Keys.T ' Talk
		bdKey(20) = System.Windows.Forms.Keys.K ' Skills
		bdKey(21) = System.Windows.Forms.Keys.Z ' Status
		bdKey(22) = System.Windows.Forms.Keys.S ' Search
		bdKey(23) = System.Windows.Forms.Keys.O ' Open
		bdKey(24) = System.Windows.Forms.Keys.I ' Inventory
		bdKey(25) = System.Windows.Forms.Keys.E ' Equip
	End Sub
	
	Public Function NotTwice(ByRef intKey As Short, ByRef code As Short) As Boolean
		Dim i As Short
		NotTwice = True
		For i = 1 To 25
			If bdKey(i) = code And intKey <> i Then
				DialogSetUp(DLGTYPE.bdDlgNoReply)
				DialogShow("", "Shortcut already used!")
				DialogHide()
				NotTwice = False
			End If
		Next 
	End Function
End Module