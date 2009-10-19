Option Strict Off
Option Explicit On
Friend Class Plot
	
	' Version number of Class
	Private myVersion As Short
	' Password protection
	Private myPassword As String
	' Plot is an Object which holds one or more EncounterMaps
	Private myMaps As New Collection
	
	Public Function AddMapWithIndex(ByRef Index As Short) As Map
		'Adds a Map and returns a reference to that Map
		Dim MapX As Map
		' Set the index and add the map. Return the new map's index.
		MapX = New Map
		MapX.Index = Index
		MapX.Name = "Map" & Index
		myMaps.Add(MapX, "M" & MapX.Index)
		AddMapWithIndex = MapX
	End Function
	
	Public Function AddMap() As Map
		'Adds a Map and returns a reference to that Map
		Dim c As Short
		Dim MapX As Map
		' Find next available index for map adding
		c = 1
		For	Each MapX In myMaps
			If MapX.Index >= c Then
				c = MapX.Index + 1
			End If
		Next MapX
		' Set the index and add the map. Return the new map's index.
		MapX = New Map
		MapX.Index = c
		MapX.Name = "Map" & c
		myMaps.Add(MapX, "M" & MapX.Index)
		AddMap = MapX
	End Function
	
	Public Sub RemoveMap(ByRef DeleteKey As String)
		myMaps.Remove(DeleteKey)
	End Sub
	
	Public ReadOnly Property TotalSize() As Integer
		Get
			Dim MapX As Map
			For	Each MapX In myMaps
				TotalSize = TotalSize + MapX.TotalSize
			Next MapX
		End Get
	End Property
	
	Public ReadOnly Property TotalMaps() As Short
		Get
			TotalMaps = myMaps.Count()
		End Get
	End Property
	
	Public ReadOnly Property Maps() As Collection
		Get
			Maps = myMaps
		End Get
	End Property
	
	Public Sub SaveToFile(ByRef ToFile As Short)
		Dim c As Short
		'UPGRADE_WARNING: Put was upgraded to FilePut and has a new behavior. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="9B7D5ADD-D8FE-4819-A36C-6DEDAF088CC7"'
		FilePut(ToFile, myVersion)
		'UPGRADE_WARNING: Put was upgraded to FilePut and has a new behavior. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="9B7D5ADD-D8FE-4819-A36C-6DEDAF088CC7"'
		FilePut(ToFile, Len(myPassword))
		'UPGRADE_WARNING: Put was upgraded to FilePut and has a new behavior. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="9B7D5ADD-D8FE-4819-A36C-6DEDAF088CC7"'
		FilePut(ToFile, myPassword)
		'UPGRADE_WARNING: Put was upgraded to FilePut and has a new behavior. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="9B7D5ADD-D8FE-4819-A36C-6DEDAF088CC7"'
		FilePut(ToFile, myMaps.Count())
		For c = 1 To myMaps.Count()
			'UPGRADE_WARNING: Couldn't resolve default property of object myMaps().SaveToFile. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
			myMaps.Item(c).SaveToFile(ToFile)
		Next c
	End Sub
	
	Public Sub ReadFromFile(ByRef FromFile As Short)
		On Error GoTo ErrorHandler
		Dim c As Short
		Dim Cnt As Integer
		Dim myMap As Map
		' Read Password
		'UPGRADE_WARNING: Get was upgraded to FileGet and has a new behavior. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="9B7D5ADD-D8FE-4819-A36C-6DEDAF088CC7"'
		FileGet(FromFile, myVersion)
		'UPGRADE_WARNING: Get was upgraded to FileGet and has a new behavior. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="9B7D5ADD-D8FE-4819-A36C-6DEDAF088CC7"'
		FileGet(FromFile, Cnt)
		myPassword = ""
		For c = 1 To Cnt
			myPassword = myPassword & " "
		Next c
		'UPGRADE_WARNING: Get was upgraded to FileGet and has a new behavior. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="9B7D5ADD-D8FE-4819-A36C-6DEDAF088CC7"'
		FileGet(FromFile, myPassword)
		' Load Maps
		'UPGRADE_WARNING: Get was upgraded to FileGet and has a new behavior. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="9B7D5ADD-D8FE-4819-A36C-6DEDAF088CC7"'
		FileGet(FromFile, Cnt)
		For c = 1 To Cnt
			myMap = New Map
			myMap.ReadFromFile(FromFile)
			myMaps.Add(myMap, "M" & myMap.Index)
		Next c
		Exit Sub
ErrorHandler: 
		' [Titi 2.4.9]
		If Err.Number = 457 Then
			oErr.logError("Map index conflict: " & myMap.Index)
		Else
			oErr.logError("Cannot read plot, error#" & Err.Number & " (" & Err.Description & ")")
		End If
	End Sub
End Class