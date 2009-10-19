Option Strict Off
Option Explicit On
Friend Class DialogStack
	
	' Unique index for this Factoid
	Private myIndex As Short
	' Things said by Creature
	Private myText As String
	' Creature talking
	Private myCreature As Creature
	
	
	Public Property Index() As Short
		Get
			Index = myIndex
		End Get
		Set(ByVal Value As Short)
			myIndex = Value
		End Set
	End Property
	
	
	Public Property Text() As String
		Get
			Text = Trim(myText)
		End Get
		Set(ByVal Value As String)
			myText = Trim(Value)
		End Set
	End Property
	
	
	Public Property CreatureTalking() As Creature
		Get
			CreatureTalking = myCreature
		End Get
		Set(ByVal Value As Creature)
			myCreature = Value
		End Set
	End Property
End Class