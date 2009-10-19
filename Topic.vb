Option Strict Off
Option Explicit On
Friend Class Topic
	' changed first
	
	' Version number of Class
	Private myVersion As Short
	' Unique number for the Topic
	Private myIndex As Short
	' Text of what CreatureWithTurn says
	Private mySay As String
	' Text of reply
	Private myReply As String
	' Action to perform after Reply
	' 0 - Do nothing
	' 1 - Set Package named in ActionTopic as CurrentConvo for NCreatureWithTurn
	' 2 - Remove this Topic
	' 3 - Add Topic in ActionTopic
	' 4 - Remove this Topic and add Topic in ActionTopic
	' 5 - Close Convo
	Private myAction As Byte
	' Reference to Index of Topic in same Convo or to different Convo
	Private myActionRef As Short
	' True if shows up as a default, else must be Action que'd.
	Private myDefault As Short
	' List of Factoids *required* in Tome for Topic to list in
	' Conversation. Must have *all* Factoids for Topic to list.
	Private myFactoids As New Collection
	' Triggers for this Topic: Post-Reply fires just after Reply is
	' displayed; Pre-Reply fires just after click on Topic, before Reply
	' is displayed (allows changing this Topic's text based on Statements);
	' Pre-ShowTopic fires just before list Topic. Fail indicates do not
	' list it (allows checking things with Statements before Topic will
	' be available).
	Private myTriggers As New Collection
	
	
	Public Property Index() As Short
		Get
			Index = myIndex
		End Get
		Set(ByVal Value As Short)
			myIndex = Value
		End Set
	End Property
	
	
	Public Property Say() As String
		Get
			Say = Trim(mySay)
		End Get
		Set(ByVal Value As String)
			mySay = Trim(Value)
		End Set
	End Property
	
	
	Public Property Reply() As String
		Get
			Reply = Trim(myReply)
		End Get
		Set(ByVal Value As String)
			myReply = Trim(Value)
		End Set
	End Property
	
	
	Public Property Action() As Byte
		Get
			Action = myAction
		End Get
		Set(ByVal Value As Byte)
			myAction = Value
		End Set
	End Property
	
	
	Public Property ActionRef() As Short
		Get
			ActionRef = CShort(Trim(CStr(myActionRef)))
		End Get
		Set(ByVal Value As Short)
			myActionRef = CShort(Trim(CStr(Value)))
		End Set
	End Property
	
	
	'UPGRADE_NOTE: Default was upgraded to Default_Renamed. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="A9E4979A-37FA-4718-9994-97DD76ED70A7"'
	Public Property Default_Renamed() As Short
		Get
			Default_Renamed = myDefault
		End Get
		Set(ByVal Value As Short)
			myDefault = Value
		End Set
	End Property
	
	Public ReadOnly Property Factoids() As Collection
		Get
			Factoids = myFactoids
		End Get
	End Property
	
	Public ReadOnly Property Triggers() As Collection
		Get
			Triggers = myTriggers
		End Get
	End Property
	
	Public Function AddFactoid() As Factoid
		' Adds an Item and returns a reference to that Item
		Dim c As Short
		Dim FactoidX As Factoid
		' Find new available unused index idenifier
		c = 1
		For	Each FactoidX In myFactoids
			If FactoidX.Index >= c Then
				c = FactoidX.Index + 1
			End If
		Next FactoidX
		' Set the index and add the factoid.
		FactoidX = New Factoid
		FactoidX.Index = c
		myFactoids.Add(FactoidX, "F" & FactoidX.Index)
		AddFactoid = FactoidX
	End Function
	
	Public Sub RemoveFactoid(ByRef DeleteKey As String)
		myFactoids.Remove(DeleteKey)
	End Sub
	
	Public Function AddTrigger() As Trigger
		' Adds an Item and returns a reference to that Item
		Dim c As Short
		Dim TriggerX As Trigger
		' Find new available unused index idenifier
		c = 1
		For	Each TriggerX In myTriggers
			If TriggerX.Index >= c Then
				c = TriggerX.Index + 1
			End If
		Next TriggerX
		' Set the index and add the door. Return the new door's index.
		TriggerX = New Trigger
		TriggerX.Index = c
		TriggerX.Name = "Trigger" & c
		myTriggers.Add(TriggerX, "T" & TriggerX.Index)
		AddTrigger = TriggerX
	End Function
	
	Public Sub RemoveTrigger(ByRef DeleteKey As String)
		myTriggers.Remove(DeleteKey)
	End Sub
	
	Public Sub Copy(ByRef FromTopic As Topic)
		Dim c As Short
		Dim FactoidX As Factoid
		Dim TriggerX As Trigger
		' Copy properties
		mySay = FromTopic.Say
		myReply = FromTopic.Reply
		myAction = FromTopic.Action
		myActionRef = FromTopic.ActionRef
		myDefault = FromTopic.Default_Renamed
		' Copy Factoids
		myFactoids = New Collection
		For c = 1 To FromTopic.Factoids.Count() 'myFactoids.Count [Titi 2.4.8]
			FactoidX = Me.AddFactoid
			'UPGRADE_WARNING: Couldn't resolve default property of object FromTopic.Factoids(). Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
			FactoidX.Copy(FromTopic.Factoids.Item(c))
		Next c
		' Copy Triggers
		myTriggers = New Collection
		For c = 1 To FromTopic.Triggers.Count()
			TriggerX = Me.AddTrigger
			'UPGRADE_WARNING: Couldn't resolve default property of object FromTopic.Triggers(). Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
			TriggerX.Copy(FromTopic.Triggers.Item(c))
		Next c
	End Sub
	
	Public Sub ReadFromFile(ByRef FromFile As Short)
		On Error GoTo ErrorHandler
		Dim c As Short
		Dim Cnt As Integer
		Dim FactoidX As Factoid
		Dim TriggerX As Trigger
		Dim Reading As String
		' Read Topic Index, Text, etc
		'UPGRADE_WARNING: Get was upgraded to FileGet and has a new behavior. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="9B7D5ADD-D8FE-4819-A36C-6DEDAF088CC7"'
		FileGet(FromFile, myVersion)
		'UPGRADE_WARNING: Get was upgraded to FileGet and has a new behavior. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="9B7D5ADD-D8FE-4819-A36C-6DEDAF088CC7"'
		FileGet(FromFile, myIndex)
		mySay = ""
		'UPGRADE_WARNING: Get was upgraded to FileGet and has a new behavior. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="9B7D5ADD-D8FE-4819-A36C-6DEDAF088CC7"'
		FileGet(FromFile, Cnt)
		For c = 1 To Cnt
			mySay = mySay & " "
		Next c
		'UPGRADE_WARNING: Get was upgraded to FileGet and has a new behavior. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="9B7D5ADD-D8FE-4819-A36C-6DEDAF088CC7"'
		FileGet(FromFile, mySay)
		myReply = ""
		'UPGRADE_WARNING: Get was upgraded to FileGet and has a new behavior. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="9B7D5ADD-D8FE-4819-A36C-6DEDAF088CC7"'
		FileGet(FromFile, Cnt)
		For c = 1 To Cnt
			myReply = myReply & " "
		Next c
		'UPGRADE_WARNING: Get was upgraded to FileGet and has a new behavior. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="9B7D5ADD-D8FE-4819-A36C-6DEDAF088CC7"'
		FileGet(FromFile, myReply)
		'UPGRADE_WARNING: Get was upgraded to FileGet and has a new behavior. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="9B7D5ADD-D8FE-4819-A36C-6DEDAF088CC7"'
		FileGet(FromFile, myAction)
		'UPGRADE_WARNING: Get was upgraded to FileGet and has a new behavior. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="9B7D5ADD-D8FE-4819-A36C-6DEDAF088CC7"'
		FileGet(FromFile, myActionRef)
		'UPGRADE_WARNING: Get was upgraded to FileGet and has a new behavior. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="9B7D5ADD-D8FE-4819-A36C-6DEDAF088CC7"'
		FileGet(FromFile, myDefault)
		' Read Factoids for Topic
		'UPGRADE_WARNING: Get was upgraded to FileGet and has a new behavior. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="9B7D5ADD-D8FE-4819-A36C-6DEDAF088CC7"'
		FileGet(FromFile, Cnt)
		Reading = "Factoid"
		For c = 1 To Cnt
			FactoidX = New Factoid
			FactoidX.ReadFromFile(FromFile)
			myFactoids.Add(FactoidX, "F" & FactoidX.Index)
		Next c
		' Read Triggers for Topic
		Reading = "Trigger"
		'UPGRADE_WARNING: Get was upgraded to FileGet and has a new behavior. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="9B7D5ADD-D8FE-4819-A36C-6DEDAF088CC7"'
		FileGet(FromFile, Cnt)
		For c = 1 To Cnt
			TriggerX = New Trigger
			TriggerX.ReadFromFile(FromFile)
			myTriggers.Add(TriggerX, "T" & TriggerX.Index)
		Next c
		Exit Sub
ErrorHandler: 
		' [Titi 2.4.9]
		If Err.Number = 457 And Reading <> "" Then
			Select Case Reading
				Case "Factoid"
					oErr.logError(Reading & " index conflict: " & FactoidX.Index)
				Case "Trigger"
					oErr.logError(Reading & " index conflict: " & TriggerX.Index)
			End Select
		Else
			oErr.logError("Cannot read topic" & IIf(Reading <> vbNullString, "'s " & Reading, "") & ", error#" & Err.Number & " (" & Err.Description & ")")
		End If
	End Sub
	
	Public Sub SaveToFile(ByRef ToFile As Short)
		Dim c As Short
		' Save Topic Name, Index, etc
		'UPGRADE_WARNING: Put was upgraded to FilePut and has a new behavior. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="9B7D5ADD-D8FE-4819-A36C-6DEDAF088CC7"'
		FilePut(ToFile, myVersion)
		'UPGRADE_WARNING: Put was upgraded to FilePut and has a new behavior. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="9B7D5ADD-D8FE-4819-A36C-6DEDAF088CC7"'
		FilePut(ToFile, myIndex)
		'UPGRADE_WARNING: Put was upgraded to FilePut and has a new behavior. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="9B7D5ADD-D8FE-4819-A36C-6DEDAF088CC7"'
		FilePut(ToFile, Len(mySay))
		'UPGRADE_WARNING: Put was upgraded to FilePut and has a new behavior. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="9B7D5ADD-D8FE-4819-A36C-6DEDAF088CC7"'
		FilePut(ToFile, mySay)
		'UPGRADE_WARNING: Put was upgraded to FilePut and has a new behavior. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="9B7D5ADD-D8FE-4819-A36C-6DEDAF088CC7"'
		FilePut(ToFile, Len(myReply))
		'UPGRADE_WARNING: Put was upgraded to FilePut and has a new behavior. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="9B7D5ADD-D8FE-4819-A36C-6DEDAF088CC7"'
		FilePut(ToFile, myReply)
		'UPGRADE_WARNING: Put was upgraded to FilePut and has a new behavior. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="9B7D5ADD-D8FE-4819-A36C-6DEDAF088CC7"'
		FilePut(ToFile, myAction)
		'UPGRADE_WARNING: Put was upgraded to FilePut and has a new behavior. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="9B7D5ADD-D8FE-4819-A36C-6DEDAF088CC7"'
		FilePut(ToFile, myActionRef)
		'UPGRADE_WARNING: Put was upgraded to FilePut and has a new behavior. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="9B7D5ADD-D8FE-4819-A36C-6DEDAF088CC7"'
		FilePut(ToFile, myDefault)
		' Save Factoids for Topic
		'UPGRADE_WARNING: Put was upgraded to FilePut and has a new behavior. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="9B7D5ADD-D8FE-4819-A36C-6DEDAF088CC7"'
		FilePut(ToFile, myFactoids.Count())
		For c = 1 To myFactoids.Count()
			'UPGRADE_WARNING: Couldn't resolve default property of object myFactoids().SaveToFile. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
			myFactoids.Item(c).SaveToFile(ToFile)
		Next c
		' Save Triggers for Topic
		'UPGRADE_WARNING: Put was upgraded to FilePut and has a new behavior. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="9B7D5ADD-D8FE-4819-A36C-6DEDAF088CC7"'
		FilePut(ToFile, myTriggers.Count())
		For c = 1 To myTriggers.Count()
			'UPGRADE_WARNING: Couldn't resolve default property of object myTriggers().SaveToFile. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
			myTriggers.Item(c).SaveToFile(ToFile)
		Next c
	End Sub
	
	'UPGRADE_NOTE: Class_Initialize was upgraded to Class_Initialize_Renamed. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="A9E4979A-37FA-4718-9994-97DD76ED70A7"'
	Private Sub Class_Initialize_Renamed()
		mySay = "You say...."
		myDefault = True
		myReply = "He/She replies...."
		myAction = 0 ' Do Nothing
	End Sub
	Public Sub New()
		MyBase.New()
		Class_Initialize_Renamed()
	End Sub
End Class