Option Strict Off
Option Explicit On
Module modCombat
	' Grid data type is only used in the combat grid.
	Structure Grid
		Dim Ref As Creature
		Dim Target As Creature
	End Structure
	
	'Public Turns(48) As Grid
	Public Turns(48) As Grid
End Module