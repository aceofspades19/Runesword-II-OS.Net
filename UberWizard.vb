Option Strict Off
Option Explicit On
Friend Class UberWizard
	
	Private myVersion As Short
	Private myMainMapIndex As Short
	Private myMakingMapIndex As Short
	
	' Name of Plot (.rsa file) to jump to when hit this entry point
	Private myAreaIndex As Short
	' Name of Map within above Plot to jump to.
	Private myMapIndex As Short
	' Name of EntryPoint within above Map.
	Private myEntryIndex As Short
	
	' Index of Major Theme (<1 if not selected)
	Private myMajorThemeIndex As Short
	
	' Tome Information
	Private myTomeName As String
	Private myTomeComments As String
	Private myTomePartySize As Byte
	Private myTomePartyAvgLevel As Byte
	Private myTotalSize As Integer
	
	' Various collections
	Private myExcludeDirs As New Collection
	Private myTomeSketchs As New Collection
	Private myMapSketchs As New Collection
	Private myThemeSketchs As New Collection
	Private myCreatures As New Collection
	Private myItems As New Collection
	Private myTomeFiles As New Collection
	Private myMapFiles As New Collection
	Private myThemeFiles As New Collection
	Private myCreatureFiles As New Collection
	Private myItemFiles As New Collection
	
	
	Public Property TomeName() As String
		Get
			TomeName = Trim(myTomeName)
		End Get
		Set(ByVal Value As String)
			myTomeName = Trim(Value)
		End Set
	End Property
	
	
	Public Property TomeComments() As String
		Get
			TomeComments = Trim(myTomeComments)
		End Get
		Set(ByVal Value As String)
			myTomeComments = Trim(Value)
		End Set
	End Property
	
	
	Public Property TomePartySize() As Short
		Get
			TomePartySize = myTomePartySize
		End Get
		Set(ByVal Value As Short)
			myTomePartySize = Value
		End Set
	End Property
	
	
	Public Property TomePartyAvgLevel() As Short
		Get
			TomePartyAvgLevel = myTomePartyAvgLevel
		End Get
		Set(ByVal Value As Short)
			myTomePartyAvgLevel = Value
		End Set
	End Property
	
	
	Public Property TotalSize() As Integer
		Get
			TotalSize = myTotalSize
		End Get
		Set(ByVal Value As Integer)
			myTotalSize = Value
		End Set
	End Property
	
	Public ReadOnly Property Difficulty() As Short
		Get
			Dim c As Short
			Select Case Me.TomePartyAvgLevel
				Case 0 ' 1-3
					c = 1
				Case 1 ' 4-6
					c = 3
				Case 2 ' 7-10
					c = 6
				Case 3 ' 10+
					c = 12
			End Select
			Select Case Me.TomePartySize
				Case 0 ' Solo
					Difficulty = 1 * c
				Case 1 ' 2-3
					Difficulty = 3 * c
				Case 2 ' 4-6
					Difficulty = 5 * c
				Case 3 ' 7-9
					Difficulty = 10 * c
			End Select
		End Get
	End Property
	
	
	Public Property Version() As Short
		Get
			Version = myVersion
		End Get
		Set(ByVal Value As Short)
			myVersion = Value
		End Set
	End Property
	
	
	Public Property AreaIndex() As Short
		Get
			AreaIndex = myAreaIndex
		End Get
		Set(ByVal Value As Short)
			myAreaIndex = Value
		End Set
	End Property
	
	
	Public Property MapIndex() As Short
		Get
			MapIndex = myMapIndex
		End Get
		Set(ByVal Value As Short)
			myMapIndex = Value
		End Set
	End Property
	
	
	Public Property EntryIndex() As Short
		Get
			EntryIndex = myEntryIndex
		End Get
		Set(ByVal Value As Short)
			myEntryIndex = Value
		End Set
	End Property
	
	Public ReadOnly Property TomeSketchs() As Collection
		Get
			TomeSketchs = myTomeSketchs
		End Get
	End Property
	
	
	Public Property MapSketchs() As Collection
		Get
			MapSketchs = myMapSketchs
		End Get
		Set(ByVal Value As Collection)
			myMapSketchs = Value
		End Set
	End Property
	
	
	Public Property ThemeSketchs() As Collection
		Get
			ThemeSketchs = myThemeSketchs
		End Get
		Set(ByVal Value As Collection)
			myThemeSketchs = Value
		End Set
	End Property
	
	
	Public Property Creatures() As Collection
		Get
			Creatures = myCreatures
		End Get
		Set(ByVal Value As Collection)
			myCreatures = Value
		End Set
	End Property
	
	
	Public Property Items() As Collection
		Get
			Items = myItems
		End Get
		Set(ByVal Value As Collection)
			myItems = Value
		End Set
	End Property
	
	Public ReadOnly Property ExcludeDirs() As Collection
		Get
			ExcludeDirs = myExcludeDirs
		End Get
	End Property
	
	Public ReadOnly Property TomeFiles() As Collection
		Get
			TomeFiles = myTomeFiles
		End Get
	End Property
	
	Public ReadOnly Property MapFiles() As Collection
		Get
			MapFiles = myMapFiles
		End Get
	End Property
	
	Public ReadOnly Property CreatureFiles() As Collection
		Get
			CreatureFiles = myCreatureFiles
		End Get
	End Property
	
	Public ReadOnly Property ThemeFiles() As Collection
		Get
			ThemeFiles = myThemeFiles
		End Get
	End Property
	
	Public ReadOnly Property ItemFiles() As Collection
		Get
			ItemFiles = myItemFiles
		End Get
	End Property
	
	Public ReadOnly Property MainMap() As MapSketch
		Get
			MainMap = myMapSketchs.Item("M" & myMainMapIndex)
		End Get
	End Property
	
	
	Public Property MainMapIndex() As Short
		Get
			MainMapIndex = myMainMapIndex
		End Get
		Set(ByVal Value As Short)
			myMainMapIndex = Value
		End Set
	End Property
	
	
	Public Property MajorThemeIndex() As Short
		Get
			MajorThemeIndex = myMajorThemeIndex
		End Get
		Set(ByVal Value As Short)
			myMajorThemeIndex = Value
		End Set
	End Property
	
	Public ReadOnly Property MakingMap() As MapSketch
		Get
			MakingMap = myMapSketchs.Item("M" & myMakingMapIndex)
		End Get
	End Property
	
	
	Public Property MakingMapIndex() As Short
		Get
			MakingMapIndex = myMakingMapIndex
		End Get
		Set(ByVal Value As Short)
			myMakingMapIndex = Value
		End Set
	End Property
	
	Public Function AddTomeSketch() As TomeSketch
		' Adds an TomeSketch and returns a reference to that TomeSketch
		Dim c As Short
		Dim TomeSketchX As TomeSketch
		' Find new available unused index idenifier
		c = 1
		For	Each TomeSketchX In myTomeSketchs
			If TomeSketchX.Index >= c Then
				c = TomeSketchX.Index + 1
			End If
		Next TomeSketchX
		' Set the index and add the TomeSketch. Return the new TomeSketchs index.
		TomeSketchX = New TomeSketch
		TomeSketchX.Index = c
		myTomeSketchs.Add(TomeSketchX, "T" & TomeSketchX.Index)
		AddTomeSketch = TomeSketchX
	End Function
	
	Public Sub RemoveTomeSketch(ByRef DeleteKey As String)
		' Remove TomeSketch
		myTomeSketchs.Remove(DeleteKey)
	End Sub
	
	Public Function AddMapSketch() As MapSketch
		' Adds an MapSketch and returns a reference to that MapSketch
		Dim c As Short
		Dim MapSketchX As MapSketch
		' Find new available unused index idenifier
		c = 1
		For	Each MapSketchX In myMapSketchs
			If MapSketchX.Index >= c Then
				c = MapSketchX.Index + 1
			End If
		Next MapSketchX
		' Set the index and add the MapSketch. Return the new MapSketchs index.
		MapSketchX = New MapSketch
		MapSketchX.Index = c
		myMapSketchs.Add(MapSketchX, "M" & MapSketchX.Index)
		AddMapSketch = MapSketchX
	End Function
	
	Public Sub RemoveMapSketch(ByRef DeleteKey As String)
		' Remove MapSketch
		myMapSketchs.Remove(DeleteKey)
	End Sub
	
	
	Public Function AddThemeSketch() As ThemeSketch
		' Adds an ThemeSketch and returns a reference to that ThemeSketch
		Dim c As Short
		Dim ThemeSketchX As ThemeSketch
		' Find new available unused index idenifier
		c = 1
		For	Each ThemeSketchX In myThemeSketchs
			If ThemeSketchX.Index >= c Then
				c = ThemeSketchX.Index + 1
			End If
		Next ThemeSketchX
		' Set the index and add the ThemeSketch. Return the new ThemeSketchs index.
		ThemeSketchX = New ThemeSketch
		ThemeSketchX.Index = c
		myThemeSketchs.Add(ThemeSketchX, "R" & ThemeSketchX.Index)
		AddThemeSketch = ThemeSketchX
	End Function
	
	Public Sub RemoveThemeSketch(ByRef DeleteKey As String)
		' Remove ThemeSketch
		myThemeSketchs.Remove(DeleteKey)
	End Sub
	
	Public Function AddCreature() As Creature
		' Adds an Creature and returns a reference to that Creature
		Dim c As Short
		Dim CreatureX As Creature
		' Find new available unused index idenifier
		c = 1
		For	Each CreatureX In myCreatures
			If CreatureX.Index >= c Then
				c = CreatureX.Index + 1
			End If
		Next CreatureX
		' Set the index and add the Creature. Return the new Creatures index.
		CreatureX = New Creature
		CreatureX.Index = c
		myCreatures.Add(CreatureX, "X" & CreatureX.Index)
		AddCreature = CreatureX
	End Function
	
	Public Sub RemoveCreature(ByRef DeleteKey As String)
		' Remove Creature
		myCreatures.Remove(DeleteKey)
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
		' Set the index and add the Item. Return the new Items index.
		ItemX = New Item
		ItemX.Index = c
		myItems.Add(ItemX, "I" & ItemX.Index)
		AddItem = ItemX
	End Function
	
	Public Sub RemoveItem(ByRef DeleteKey As String)
		' Remove Item
		myItems.Remove(DeleteKey)
	End Sub
End Class