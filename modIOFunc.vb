Option Strict Off
Option Explicit On
Imports VB = Microsoft.VisualBasic
Module modIOFunc

    Private tome As Tome = Tome.getInstance()
	Private Declare Function sndPlaySound Lib "winmm.dll"  Alias "sndPlaySoundA"(ByVal lpszSoundName As String, ByVal uFlags As Integer) As Integer
	
	Public Enum ClickType
		ifClickMenu = 0
		ifClick = 1
		ifClickDrop = 2
		ifClickLow = 3
		ifClickPass = 4
		ifClickCast = 5
	End Enum
	
	Public Enum RNDMUSICSTYLE
		Adventure = 0
		Combat = 1
	End Enum
	
	Public oFileSys As clsInOut
	Public oErr As CErrorHandler
	Public oGameMusic As IMCI
	
	'Public Sub PlayVideo(ByVal VideoName As String, Optional AlternatePath As Variant)
	'
	'End Sub
	
	'UPGRADE_NOTE: Tome was upgraded to Tome_Renamed. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="A9E4979A-37FA-4718-9994-97DD76ED70A7"'
	Public Sub PlayMusic(ByVal SongName As String, ByRef frm As System.Windows.Forms.Form, Optional ByRef AlternatePath As Object = Nothing)
		Dim Tome_Renamed As Object
		Dim c As Short
		Dim FileName As String
		On Error GoTo Err_Handler
		If GlobalMusicState = 0 Then Exit Sub
		If SongName = vbNullString Then Exit Sub
		' Find location of Song
		' music selection will be in the following priority order:
		' 1- from the tome files
		' 2- from the world dedicated music folder
		' 3- from the default RS music
		' tome music
		'UPGRADE_WARNING: Couldn't resolve default property of object Tome.FullPath. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
        FileName = tome.FullPath & "\" & SongName
		If oFileSys.CheckExists(FileName, clsInOut.IOActionType.File) Then GoTo PlaySong
		'UPGRADE_WARNING: Couldn't resolve default property of object Tome.FullPath. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
        FileName = tome.FullPath & "\Combat\" & SongName
		If oFileSys.CheckExists(FileName, clsInOut.IOActionType.File) Then GoTo PlaySong
		'UPGRADE_WARNING: Couldn't resolve default property of object Tome.FullPath. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
        FileName = tome.FullPath & "\Adventure\" & SongName
		If oFileSys.CheckExists(FileName, clsInOut.IOActionType.File) Then GoTo PlaySong
		'UPGRADE_NOTE: IsMissing() was changed to IsNothing(). Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="8AE1CB93-37AB-439A-A4FF-BE3B6760BB23"'
		If Not IsNothing(AlternatePath) Then
			' structure of world specific music folder is supposed to be the same as the default
			'UPGRADE_WARNING: Couldn't resolve default property of object AlternatePath. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
			FileName = AlternatePath + "\" + SongName
			If oFileSys.CheckExists(FileName, clsInOut.IOActionType.File) Then GoTo PlaySong
			'UPGRADE_WARNING: Couldn't resolve default property of object AlternatePath. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
			FileName = AlternatePath + "\Combat\" + SongName
			If oFileSys.CheckExists(FileName, clsInOut.IOActionType.File) Then GoTo PlaySong
			'UPGRADE_WARNING: Couldn't resolve default property of object AlternatePath. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
			FileName = AlternatePath + "\Adventure\" + SongName
			If oFileSys.CheckExists(FileName, clsInOut.IOActionType.File) Then GoTo PlaySong
		End If
		' default music
		'    FileName = gAppPath & "\data\music\" & SongName
		FileName = gDataPath & "\music\" & SongName
		If oFileSys.CheckExists(FileName, clsInOut.IOActionType.File) Then GoTo PlaySong
		'    FileName = gAppPath & "\data\music\combat\" & SongName
		FileName = gDataPath & "\music\combat\" & SongName
		If oFileSys.CheckExists(FileName, clsInOut.IOActionType.File) Then GoTo PlaySong
		'    FileName = gAppPath & "\data\music\adventure\" & SongName
		FileName = gDataPath & "\music\adventure\" & SongName
		If oFileSys.CheckExists(FileName, clsInOut.IOActionType.File) Then GoTo PlaySong
		Exit Sub
PlaySong: 
		' Play MID with Windows Media Player
		'frm.MediaPlayerMusic.Filename = Filename 'SongName
		If oGameMusic.InitAudio(FileName, "GlobalMusicState") Then
			GlobalMusicName = SongName
			'frm.MediaPlayerMusic.Play
			oGameMusic.Play()
		End If
		Exit Sub
Err_Handler: 
		oErr.logError("PlayMusic: " & FileName & " not found")
		Err.Clear()
		Exit Sub
	End Sub
	
	Sub PlaySoundFile(ByVal SoundName As String, Optional ByRef Tome As Tome = Nothing, Optional ByRef ForcedPath As Boolean = False, Optional ByRef Wait As Short = 0, Optional ByRef uFlag As Integer = &H1)
		Dim rc As Integer
		Dim PauseUntil As Single
		Dim FileName As String
		On Error GoTo Err_Handler
		If GlobalWAVState = 0 Then Exit Sub ' Are sound effects ON
		If SoundName = vbNullString Then Exit Sub
		' If not ending in .wav, then do that now
		If Right(LCase(SoundName), 4) <> ".wav" Then
			SoundName = SoundName & ".wav"
		End If
		If ForcedPath Then
			FileName = SoundName
			If oFileSys.CheckExists(FileName, clsInOut.IOActionType.File) Then GoTo PlaySound
		End If
		'    FileName = gAppPath & "\data\sounds\" & SoundName
		FileName = gDataPath & "\sounds\" & SoundName
		If oFileSys.CheckExists(FileName, clsInOut.IOActionType.File) Then GoTo PlaySound
		If Not Tome Is Nothing Then
            FileName = Tome.FullPath & "\" & SoundName
			If oFileSys.CheckExists(FileName, clsInOut.IOActionType.File) Then GoTo PlaySound
		End If
		Exit Sub
PlaySound: 
		' Play the Sound file
		'SND_SYNC      = &H00 Program loses control until the sound ends
		'SND_ASYN      = &H01 Program gets back control.
		'SND_NODEFAULT = &H02 Don't beep if can't find sound.
		'SND_LOOP      = &H08 Play until called again with WaveFile = ""
		'SND_NOSTOP    = &H10 If sound is already playing, return false and don't play.
		If Wait > 0 Then
			PauseUntil = VB.Timer() + Wait '5
			Do Until rc = 1 Or VB.Timer() > PauseUntil
				rc = sndPlaySound(FileName, &H10)
			Loop 
		Else
			rc = sndPlaySound(FileName, uFlag)
		End If
		
		Exit Sub
Err_Handler: 
		oErr.logError("PlaySound: " & FileName & " not found.")
		Err.Clear()
		Resume Next
	End Sub
	
	' following are only required in  the game engine
	'UPGRADE_NOTE: Tome was upgraded to Tome_Renamed. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="A9E4979A-37FA-4718-9994-97DD76ED70A7"'
	Public Sub PlayMusicRnd(ByRef Style As RNDMUSICSTYLE, ByRef frm As System.Windows.Forms.Form)

		Dim FileName As String
		Dim c, Pick As Short
		Dim sPath As String
		If GlobalMusicState = 1 And GlobalMusicRandom = 1 Then
			' [Titi 2.4.7] music can now belong to a world
			' music selection will then be in the following priority order:
			' 1- from the tome files
			' 2- from the world dedicated music folder
			' 3- from the default RS music
			'UPGRADE_WARNING: Couldn't resolve default property of object Tome.LoadPath. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
            sPath = tome.LoadPath & "\" & IIf(Style, "Combat", "Adventure")
			If Not oFileSys.CheckExists(sPath, clsInOut.IOActionType.Folder) Then sPath = WorldNow.MusicFolder & "\" & IIf(Style, "Combat", "Adventure")
			' default music
			'        If Not oFileSys.CheckExists(sPath, Folder) Then sPath = gAppPath & "\data\music\" & IIf(Style, "Combat", "Adventure")
			If Not oFileSys.CheckExists(sPath, clsInOut.IOActionType.Folder) Then sPath = gDataPath & "\music\" & IIf(Style, "Combat", "Adventure")
			' Count available files
			c = 0
			'UPGRADE_WARNING: Dir has a new behavior. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="9B7D5ADD-D8FE-4819-A36C-6DEDAF088CC7"'
			FileName = Dir(sPath & "\*.mp3")
			Do Until FileName = ""
				c = c + 1
				'UPGRADE_WARNING: Dir has a new behavior. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="9B7D5ADD-D8FE-4819-A36C-6DEDAF088CC7"'
				FileName = Dir()
			Loop 
			'UPGRADE_WARNING: Dir has a new behavior. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="9B7D5ADD-D8FE-4819-A36C-6DEDAF088CC7"'
			FileName = Dir(sPath & "\*.mid")
			Do Until FileName = ""
				c = c + 1
				'UPGRADE_WARNING: Dir has a new behavior. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="9B7D5ADD-D8FE-4819-A36C-6DEDAF088CC7"'
				FileName = Dir()
			Loop 
			' If have at least 1 then
			If c > 0 Then
				' Find that file
				Pick = Int(Rnd() * c) ' was c+1 --> bug if only one file available: would never be found [Titi 2.4.7]
				c = 0
				'UPGRADE_WARNING: Dir has a new behavior. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="9B7D5ADD-D8FE-4819-A36C-6DEDAF088CC7"'
				FileName = Dir(sPath & "\*.mp3")
				Do Until FileName = "" Or c = Pick
					c = c + 1
					'UPGRADE_WARNING: Dir has a new behavior. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="9B7D5ADD-D8FE-4819-A36C-6DEDAF088CC7"'
					FileName = Dir()
				Loop 
				If c <> Pick Then
					c = c + 1
					'UPGRADE_WARNING: Dir has a new behavior. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="9B7D5ADD-D8FE-4819-A36C-6DEDAF088CC7"'
					FileName = Dir(sPath & "\*.mid")
					Do Until FileName = "" Or c = Pick
						c = c + 1
						'UPGRADE_WARNING: Dir has a new behavior. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="9B7D5ADD-D8FE-4819-A36C-6DEDAF088CC7"'
						FileName = Dir()
					Loop 
				End If
				' [Titi 2.4.7] added sPath to allow world-specific music folder
				'            Call PlayMusic(Filename, frm)
				Call PlayMusic(FileName, frm, sPath)
				GlobalMusicName = FileName
			End If
		End If
		Exit Sub
	End Sub
	
	Public Sub PlayClickSnd(ByRef ClickType As ClickType)
		Dim sSound As String
		If GlobalMouseClick = 1 And ClickType < ClickType.ifClickLow Then
			Select Case ClickType
				Case ClickType.ifClickMenu
					sSound = "click2.wav"
				Case ClickType.ifClick
					sSound = "click1.wav"
				Case ClickType.ifClickDrop
					sSound = "click4.wav"
				Case ClickType.ifClickLow
					sSound = "click3.wav"
			End Select
			If sSound <> vbNullString Then
				'            Call PlaySoundFile(gAppPath & "\data\interface\" & GlobalInterfaceName & "\" & sSound, Tome, True)
                Call PlaySoundFile(gDataPath & "\interface\" & GlobalInterfaceName & "\" & sSound, tome, True)
			End If
		Else
			Select Case ClickType
				Case ClickType.ifClickPass
					sSound = "pass.wav"
				Case ClickType.ifClickCast
					sSound = "cast.wav"
			End Select
			If sSound <> vbNullString Then
				'            Call PlaySoundFile(gAppPath & "\data\stock\" & sSound, Tome, True, 5, &H10)
                Call PlaySoundFile(gDataPath & "\stock\" & sSound, tome, True, 5, &H10)
			End If
		End If
	End Sub
	
	Public Function ClipPath(ByRef FullPath As String) As String
		' Takes in a full path (including the file name) and returns just the path
		Dim c As Short
		For c = Len(FullPath) To 1 Step -1
			If Mid(FullPath, c, 1) = "/" Or Mid(FullPath, c, 1) = "\" Then
				ClipPath = Mid(FullPath, 1, c - 1)
				Exit Function
			End If
		Next c
		ClipPath = ""
	End Function
	
	Public Function fCreateObject(ByRef SID As String) As Object
		On Error GoTo Err_Handler
		fCreateObject = CreateObject(SID)
		Exit Function
Err_Handler: 
		oErr.logError("fCreateObject()")
		Err.Raise(Err.Number, "fCreateObject", Err.Description & ": '" & SID & "'")
	End Function
End Module