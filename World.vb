Option Strict Off
Option Explicit On
Friend Class World
	
	' Name of the World
	Private myName As String
	' Index of World
	Private myIndex As Short
	' File Name of Map for World
	Private myPictureFile As String
	' Tome Name of Tome (containing Creature templates) for World
	Private myTomeName As String
	' [borfaux] Added for 2.4.6
	' Worlds Intro Music
	Private myIntro As String
	' [Titi] Added for 2.4.7
	' Worlds Music folder
	Private myMusicFolder As String
	' [Titi] Added for 2.4.8
	' Worlds currency name
	Private myCurrency As String
	' [Titi] Added for 2.4.9
	' Worlds description
	Private myDescription As String
	Private myPicDesc As String
	Private mySkillPtsPerLevel As Short
	Private myHPPerLevel As Short
	Private myRandomizeHP As Boolean
	' Actual Tome for World
	Private myTome As Tome
	' Current Creature in Tome used as Template
	Private myTemplate As Creature
	' CreatureWithTurn being built from CreatureTmplt
	Private myCreatureWithTurn As Creature
	' Stat adjusting
	Private myAdjustStat As Short
	
	Private myKingdoms As New Collection
	
	Private bIsCurrent As Boolean
	Private bIsLoaded As Boolean
	
	Private sPath As String
	
	
	Public Property MusicFolder() As String
		Get
			MusicFolder = myMusicFolder
		End Get
		Set(ByVal Value As String)
			myMusicFolder = Value
		End Set
	End Property
	
	
	Public Property Description() As String
		Get
			Description = myDescription
		End Get
		Set(ByVal Value As String)
			myDescription = Value
		End Set
	End Property
	
	
	Public Property Money() As String
		Get
			Money = myCurrency
		End Get
		Set(ByVal Value As String)
			myCurrency = Value
		End Set
	End Property
	
	
	Public Property SkillPtsPerLevel() As Short
		Get
			SkillPtsPerLevel = mySkillPtsPerLevel
		End Get
		Set(ByVal Value As Short)
			mySkillPtsPerLevel = Value
		End Set
	End Property
	
	
	Public Property HPPerLevel() As Short
		Get
			HPPerLevel = myHPPerLevel
		End Get
		Set(ByVal Value As Short)
			myHPPerLevel = Value
		End Set
	End Property
	
	
	Public Property RandomizeHP() As Boolean
		Get
			RandomizeHP = myRandomizeHP
		End Get
		Set(ByVal Value As Boolean)
			myRandomizeHP = Value
		End Set
	End Property
	
	
	Public Property IntroMusic() As String
		Get
			IntroMusic = myIntro
		End Get
		Set(ByVal Value As String)
			myIntro = Value
		End Set
	End Property
	
	
	Public Property Path() As String
		Get
			Path = sPath
		End Get
		Set(ByVal Value As String)
			sPath = Value
		End Set
	End Property
	
	
	Public Property IsCurrent() As Boolean
		Get
			IsCurrent = bIsCurrent
			'UPGRADE_NOTE: Object myTome may not be destroyed until it is garbage collected. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6E35BFF6-CD74-4B09-9689-3E1A43DF8969"'
			myTome = Nothing
		End Get
		Set(ByVal Value As Boolean)
			Dim c, Kingdoms As Short
			Dim n, i, k As Short
			Dim FileName As String
			Dim Text As String
			Dim KingdomX As Kingdom
			Dim CreatureX As Creature
			Dim PictureX As Factoid
			Dim lResult As Integer
			bIsCurrent = Value
			If Value And Not bIsLoaded Then
				c = FreeFile
				FileOpen(c, sPath & TomeName, OpenMode.Binary)
				myTome.ReadFromFile(c)
				FileClose(c)
				'Get Kingdoms in the World
				FileName = Path & Name & ".ini"
				lResult = fReadValue(FileName, "World", "Kingdoms", "S", 0, Kingdoms)
				For c = 1 To Kingdoms
					' Create a new Kingdom
					KingdomX = AddKingdom
					' Get the Kingdom Name
					lResult = fReadValue(FileName, "Kingdom" & c, "Name", "S", "Untitled", Text)
					KingdomX.Name = Text
					' Map coordinates
					lResult = fReadValue(FileName, "Kingdom" & c, "Location", "S", "0", Text)
					Text = Text & ","
					i = InStr(1, Text, ",") : k = 0 : n = 0
					Do While i > 0
						Select Case n
							Case 0 ' Left
								KingdomX.Left_Renamed = Mid(Text, k + 1, i - k - 1)
							Case 1 ' Top
								KingdomX.Top = CShort(Mid(Text, k + 1, i - k - 1))
							Case 2 ' Width
								KingdomX.Width = CShort(Mid(Text, k + 1, i - k - 1))
							Case 3 ' Height
								KingdomX.Height = CShort(Mid(Text, k + 1, i - k - 1))
						End Select
						k = i
						i = InStr(k + 1, Text, ",")
						n = n + 1
					Loop 
					' Get Kingdom Creature Template Name
					lResult = fReadValue(FileName, "Kingdom" & c, "Template", "S", "0", Text, 1024)
					' Find the template in the Tome
					KingdomX.Template = New Creature
					For	Each CreatureX In myTome.Creatures
						If CreatureX.Name = Text Then
							KingdomX.Template = CreatureX
							Exit For
						End If
					Next CreatureX
					' Get PictureFile names (parse up into text)
					lResult = fReadValue(FileName, "Kingdom" & c, "PictureFiles", "S", "0", Text, 8192)
					Text = Text & ","
					i = InStr(1, Text, ",")
					k = 0
					Do While i > 0
						PictureX = KingdomX.AddPictureFile
						PictureX.Text = Mid(Text, k + 1, i - k - 1)
						k = i
						i = InStr(k + 1, Text, ",")
					Loop 
					' Get NameSet Index
					lResult = fReadValue(FileName, "Kingdom" & c, "NameSet", "S", "0", Text)
					KingdomX.NameSet = CShort(Text)
				Next c
				bIsLoaded = True
			End If
		End Set
	End Property
	
	
	Public Property Name() As String
		Get
			Name = myName
		End Get
		Set(ByVal Value As String)
			myName = Value
		End Set
	End Property
	
	
	Public Property Index() As Short
		Get
			Index = myIndex
		End Get
		Set(ByVal Value As Short)
			myIndex = Value
		End Set
	End Property
	
	Public ReadOnly Property Kingdoms() As Collection
		Get
			Kingdoms = myKingdoms
		End Get
	End Property
	
	
	Public Property PictureFile() As String
		Get
			PictureFile = myPictureFile
		End Get
		Set(ByVal Value As String)
			myPictureFile = Value
		End Set
	End Property
	
	
	Public Property PicDesc() As String
		Get
			PicDesc = myPicDesc
		End Get
		Set(ByVal Value As String)
			myPicDesc = Value
		End Set
	End Property
	
	
	Public Property TomeName() As String
		Get
			TomeName = myTomeName
		End Get
		Set(ByVal Value As String)
			myTomeName = Value
		End Set
	End Property
	
	
	Public Property Tome() As Tome
		Get
			Tome = myTome
		End Get
		Set(ByVal Value As Tome)
			myTome = Value
		End Set
	End Property
	
	
	Public Property Template() As Creature
		Get
			Template = myTemplate
		End Get
		Set(ByVal Value As Creature)
			myTemplate = Value
		End Set
	End Property
	
	
	Public Property CreatureWithTurn() As Creature
		Get
			CreatureWithTurn = myCreatureWithTurn
		End Get
		Set(ByVal Value As Creature)
			myCreatureWithTurn = Value
		End Set
	End Property
	
	
	Public Property AdjustStat() As Short
		Get
			AdjustStat = myAdjustStat
		End Get
		Set(ByVal Value As Short)
			myAdjustStat = Value
		End Set
	End Property
	
	Public Function AddKingdom() As Kingdom
		' Adds an Kingdom and returns a reference to that Kingdom
		Dim c As Short
		Dim KingdomX As Kingdom
		' Find new available unused index idenifier
		c = 1
		For	Each KingdomX In myKingdoms
			If KingdomX.Index >= c Then
				c = KingdomX.Index + 1
			End If
		Next KingdomX
		' Set the index and add the door. Return the new door's index.
		KingdomX = New Kingdom
		KingdomX.Index = c
		myKingdoms.Add(KingdomX, "K" & KingdomX.Index)
		AddKingdom = KingdomX
	End Function
	
	'UPGRADE_NOTE: Class_Initialize was upgraded to Class_Initialize_Renamed. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="A9E4979A-37FA-4718-9994-97DD76ED70A7"'
	Private Sub Class_Initialize_Renamed()
		'UPGRADE_NOTE: Object myTome may not be destroyed until it is garbage collected. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6E35BFF6-CD74-4B09-9689-3E1A43DF8969"'
		myTome = Nothing
		bIsLoaded = False
	End Sub
	Public Sub New()
		MyBase.New()
		Class_Initialize_Renamed()
	End Sub
End Class