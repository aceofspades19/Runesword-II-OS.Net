Option Strict Off
Option Explicit On
Module modTriggers
	
	' Global Variables (accessible via Triggers)
	Public GlobalAttackRoll As Short
	Public GlobalArmorRoll As Short
	Public GlobalHitLocation As String
	Public GlobalDamageRoll As Short
	Public GlobalDieTypeRoll As Short
	Public GlobalDieCountRoll As Short
	Public GlobalOffer As Short
	Public GlobalSkillLevel As Short
	Public GlobalDamageStyle As Short
	Public GlobalPickLockChance As Short
	Public GlobalRemoveTrapChance As Short
	Public GlobalIntegerA As Short
	Public GlobalIntegerB As Short
	Public GlobalIntegerC As Short
	Public GlobalTextA As String
	Public GlobalTextB As String
	Public GlobalTextC As String
	Public GlobalDayName As String
	Public GlobalMoonName As String
	Public GlobalYearName As String
	Public GlobalTurnName As String
	Public GlobalSpellFizzleChance As Short
	Public GlobalTicks As Short
	
	Public Function FireTriggers(ByRef ObjectX As Object, ByRef AsTriggerType As Short) As Short
		Dim rc As Short
		Dim TriggerX As Trigger
		FireTriggers = True
		'UPGRADE_WARNING: Couldn't resolve default property of object ObjectX.Triggers. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
		For	Each TriggerX In ObjectX.Triggers
			If TriggerX.TriggerType = AsTriggerType Then
				rc = FireTrigger(ObjectX, TriggerX)
				If rc = False Then
					FireTriggers = False
				End If
			End If
		Next TriggerX
	End Function
End Module