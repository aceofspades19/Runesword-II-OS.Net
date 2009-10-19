Option Strict Off
Option Explicit On
Friend Class Container
	
	' The Container that is open
	Private myItem As Item
	' Index to first Item showing inside Container
	Private myTopIndex As Short
	' The index to the picContainer() that represents this Container
	Private myIndex As Short
	
	
	Public Property Index() As Short
		Get
			Index = myIndex
		End Get
		Set(ByVal Value As Short)
			myIndex = Value
		End Set
	End Property
	
	
	Public Property ItemX() As Item
		Get
			ItemX = myItem
		End Get
		Set(ByVal Value As Item)
			myItem = Value
		End Set
	End Property
	
	
	Public Property TopIndex() As Short
		Get
			TopIndex = myTopIndex
		End Get
		Set(ByVal Value As Short)
			myTopIndex = Value
		End Set
	End Property
End Class