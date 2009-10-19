Option Strict Off
Option Explicit On
Friend Class Kingdom
	
	' Name of the Kingdom
	Private myName As String
	' Index of Kingdom
	Private myIndex As Short
	' Coordinates of Kingdom on World Map
	Private myLeft As Short
	Private myTop As Short
	Private myWidth As Short
	Private myHeight As Short
	' NameSet index
	Private myNameSet As Short
	' List of PictureFiles for CreatureWithTurns
	Private myPictureFiles As New Collection
	' Points to the Creature used as a template
	Private myTemplate As Creature
	' List of Skills for Template
	Private mySkills As Collection
	
	
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
	
	
	Public Property NameSet() As Short
		Get
			NameSet = myNameSet
		End Get
		Set(ByVal Value As Short)
			myNameSet = Value
		End Set
	End Property
	
	
	Public Property Template() As Creature
		Get
			Template = myTemplate
		End Get
		Set(ByVal Value As Creature)
			Dim TriggerX As Trigger
			Dim c As Short
			myTemplate = Value
			' Set up skills list
			mySkills = New Collection
			For	Each TriggerX In Value.Triggers
				If TriggerX.IsSkill = True Then
					For c = 1 To mySkills.Count()
						'UPGRADE_WARNING: Couldn't resolve default property of object mySkills().Name. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
						If UCase(mySkills.Item(c).Name) > UCase(TriggerX.Name) Then
							mySkills.Add(TriggerX, "T" & TriggerX.Index, c)
							Exit For
						End If
					Next c
					If c > mySkills.Count() Then
						mySkills.Add(TriggerX, "T" & TriggerX.Index)
					End If
				End If
			Next TriggerX
		End Set
	End Property
	
	Public ReadOnly Property Skills() As Collection
		Get
			Skills = mySkills
		End Get
	End Property
	
	Public ReadOnly Property PictureFiles() As Collection
		Get
			PictureFiles = myPictureFiles
		End Get
	End Property
	
	
	'UPGRADE_NOTE: Left was upgraded to Left_Renamed. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="A9E4979A-37FA-4718-9994-97DD76ED70A7"'
	Public Property Left_Renamed() As Short
		Get
			Left_Renamed = myLeft
		End Get
		Set(ByVal Value As Short)
			myLeft = Value
		End Set
	End Property
	
	
	Public Property Top() As Short
		Get
			Top = myTop
		End Get
		Set(ByVal Value As Short)
			myTop = Value
		End Set
	End Property
	
	
	Public Property Width() As Short
		Get
			Width = myWidth
		End Get
		Set(ByVal Value As Short)
			myWidth = Value
		End Set
	End Property
	
	
	Public Property Height() As Short
		Get
			Height = myHeight
		End Get
		Set(ByVal Value As Short)
			myHeight = Value
		End Set
	End Property
	
	Public Function AddPictureFile() As Factoid
		' Adds an PictureFile and returns a reference to that PictureFile
		Dim c As Short
		Dim PictureFileX As Factoid
		' Find new available unused index idenifier
		c = 1
		For	Each PictureFileX In myPictureFiles
			If PictureFileX.Index >= c Then
				c = PictureFileX.Index + 1
			End If
		Next PictureFileX
		' Set the index and add the door. Return the new door's index.
		PictureFileX = New Factoid
		PictureFileX.Text = "Untitled" & c
		PictureFileX.Index = c
		myPictureFiles.Add(PictureFileX, "P" & PictureFileX.Index)
		AddPictureFile = PictureFileX
	End Function
End Class