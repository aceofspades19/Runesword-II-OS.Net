Option Strict Off
Option Explicit On
Friend Class Conversation
	
	' Version number of Class
	Private myVersion As Short
	' Unique index of the Package
	Private myIndex As Short
	' Name of the Convo Package
	Private myName As String
	' Description when first approach
	Private myFirstTalk As String
	' Description for any approach after the first (blank if NA)
	Private mySecondTalk As String
	' Pointer to Reply (not saved)
	Private myReply As Short
	' &H1   - On = Have Talked, Off = Have not
	' &H2   - Not Used
	Private myFlags As Byte
	' Collection of Topics in this Package
	Private myTopics As New Collection
	
	
	Public Property Index() As Short
		Get
			Index = myIndex
		End Get
		Set(ByVal Value As Short)
			myIndex = Value
		End Set
	End Property
	
	
	Public Property Reply() As Short
		Get
			Reply = myReply
		End Get
		Set(ByVal Value As Short)
			myReply = Value
		End Set
	End Property
	
	
	Public Property Flags() As Short
		Get
			Flags = myFlags
		End Get
		Set(ByVal Value As Short)
			myFlags = Value
		End Set
	End Property
	
	
	Public Property Name() As String
		Get
			Name = Trim(myName)
		End Get
		Set(ByVal Value As String)
			myName = Trim(Value)
		End Set
	End Property
	
	
	Public Property FirstTalk() As String
		Get
			FirstTalk = Trim(myFirstTalk)
		End Get
		Set(ByVal Value As String)
			myFirstTalk = Trim(Value)
		End Set
	End Property
	
	
	Public Property SecondTalk() As String
		Get
			SecondTalk = Trim(mySecondTalk)
		End Get
		Set(ByVal Value As String)
			mySecondTalk = Trim(Value)
		End Set
	End Property
	
	
	Public Property HaveTalked() As Short
		Get
			HaveTalked = (myFlags And &H1) > 0
		End Get
		Set(ByVal Value As Short)
			If Value Then
				myFlags = myFlags Or &H1
			Else
				myFlags = myFlags And (Not &H1)
			End If
		End Set
	End Property
	
	Public ReadOnly Property Topics() As Collection
		Get
			Topics = myTopics
		End Get
	End Property
	
	Public Function AddTopic() As Topic
		' Adds a Topic and returns a reference to that Topic
		Dim c As Short
		Dim TopicX As Topic
		' Find new available unused index idenifier
		c = 1
		For	Each TopicX In myTopics
			If TopicX.Index >= c Then
				c = TopicX.Index + 1
			End If
		Next TopicX
		' Set the index and add the Topic.
		TopicX = New Topic
		TopicX.Index = c
		TopicX.Say = "Topic" & c
		myTopics.Add(TopicX, "Q" & TopicX.Index)
		AddTopic = TopicX
	End Function
	
	Public Sub RemoveTopic(ByRef DeleteKey As String)
		myTopics.Remove(DeleteKey)
	End Sub
	
	Public Sub Copy(ByRef FromConversation As Conversation)
		Dim c As Short
		Dim TopicX As Topic
		' Copy Conversation Name
		myName = FromConversation.Name
		myFirstTalk = FromConversation.FirstTalk
		mySecondTalk = FromConversation.SecondTalk
		myFlags = FromConversation.Flags
		' Copy Topics for Conversation
		myTopics = New Collection
		For c = 1 To FromConversation.Topics.Count()
			TopicX = Me.AddTopic
			'UPGRADE_WARNING: Couldn't resolve default property of object FromConversation.Topics(). Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
			TopicX.Copy(FromConversation.Topics.Item(c))
		Next c
	End Sub
	
	Public Sub ReadFromFile(ByRef FromFile As Short)
		On Error GoTo ErrorHandler
		Dim c As Short
		Dim Cnt As Integer
		Dim TopicX As Topic
		' Read Conversation Name and Index
		'UPGRADE_WARNING: Get was upgraded to FileGet and has a new behavior. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="9B7D5ADD-D8FE-4819-A36C-6DEDAF088CC7"'
		FileGet(FromFile, myVersion)
		'UPGRADE_WARNING: Get was upgraded to FileGet and has a new behavior. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="9B7D5ADD-D8FE-4819-A36C-6DEDAF088CC7"'
		FileGet(FromFile, myIndex)
		'UPGRADE_WARNING: Get was upgraded to FileGet and has a new behavior. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="9B7D5ADD-D8FE-4819-A36C-6DEDAF088CC7"'
		FileGet(FromFile, Cnt)
		myName = ""
		For c = 1 To Cnt
			myName = myName & " "
		Next c
		'UPGRADE_WARNING: Get was upgraded to FileGet and has a new behavior. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="9B7D5ADD-D8FE-4819-A36C-6DEDAF088CC7"'
		FileGet(FromFile, myName)
		' Read Talk Text
		'UPGRADE_WARNING: Get was upgraded to FileGet and has a new behavior. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="9B7D5ADD-D8FE-4819-A36C-6DEDAF088CC7"'
		FileGet(FromFile, Cnt)
		myFirstTalk = ""
		For c = 1 To Cnt
			myFirstTalk = myFirstTalk & " "
		Next c
		'UPGRADE_WARNING: Get was upgraded to FileGet and has a new behavior. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="9B7D5ADD-D8FE-4819-A36C-6DEDAF088CC7"'
		FileGet(FromFile, myFirstTalk)
		'UPGRADE_WARNING: Get was upgraded to FileGet and has a new behavior. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="9B7D5ADD-D8FE-4819-A36C-6DEDAF088CC7"'
		FileGet(FromFile, Cnt)
		mySecondTalk = ""
		For c = 1 To Cnt
			mySecondTalk = mySecondTalk & " "
		Next c
		'UPGRADE_WARNING: Get was upgraded to FileGet and has a new behavior. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="9B7D5ADD-D8FE-4819-A36C-6DEDAF088CC7"'
		FileGet(FromFile, mySecondTalk)
		'UPGRADE_WARNING: Get was upgraded to FileGet and has a new behavior. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="9B7D5ADD-D8FE-4819-A36C-6DEDAF088CC7"'
		FileGet(FromFile, myFlags)
		' Read Topics for Conversation
		'UPGRADE_WARNING: Get was upgraded to FileGet and has a new behavior. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="9B7D5ADD-D8FE-4819-A36C-6DEDAF088CC7"'
		FileGet(FromFile, Cnt)
		For c = 1 To Cnt
			TopicX = New Topic
			TopicX.ReadFromFile(FromFile)
			myTopics.Add(TopicX, "Q" & TopicX.Index)
		Next c
		Exit Sub
ErrorHandler: 
		' [Titi 2.4.9]
		If Err.Number = 457 Then
			oErr.logError("Topic index conflict: " & TopicX.Index)
		Else
			oErr.logError("Cannot read conversation, error#" & Err.Number & " (" & Err.Description & ")")
		End If
	End Sub
	
	Public Sub SaveToFile(ByRef ToFile As Short)
		Dim c As Short
		' Save Conversation
		'UPGRADE_WARNING: Put was upgraded to FilePut and has a new behavior. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="9B7D5ADD-D8FE-4819-A36C-6DEDAF088CC7"'
		FilePut(ToFile, myVersion)
		'UPGRADE_WARNING: Put was upgraded to FilePut and has a new behavior. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="9B7D5ADD-D8FE-4819-A36C-6DEDAF088CC7"'
		FilePut(ToFile, myIndex)
		'UPGRADE_WARNING: Put was upgraded to FilePut and has a new behavior. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="9B7D5ADD-D8FE-4819-A36C-6DEDAF088CC7"'
		FilePut(ToFile, Len(myName))
		'UPGRADE_WARNING: Put was upgraded to FilePut and has a new behavior. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="9B7D5ADD-D8FE-4819-A36C-6DEDAF088CC7"'
		FilePut(ToFile, myName)
		'UPGRADE_WARNING: Put was upgraded to FilePut and has a new behavior. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="9B7D5ADD-D8FE-4819-A36C-6DEDAF088CC7"'
		FilePut(ToFile, Len(myFirstTalk))
		'UPGRADE_WARNING: Put was upgraded to FilePut and has a new behavior. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="9B7D5ADD-D8FE-4819-A36C-6DEDAF088CC7"'
		FilePut(ToFile, myFirstTalk)
		'UPGRADE_WARNING: Put was upgraded to FilePut and has a new behavior. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="9B7D5ADD-D8FE-4819-A36C-6DEDAF088CC7"'
		FilePut(ToFile, Len(mySecondTalk))
		'UPGRADE_WARNING: Put was upgraded to FilePut and has a new behavior. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="9B7D5ADD-D8FE-4819-A36C-6DEDAF088CC7"'
		FilePut(ToFile, mySecondTalk)
		'UPGRADE_WARNING: Put was upgraded to FilePut and has a new behavior. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="9B7D5ADD-D8FE-4819-A36C-6DEDAF088CC7"'
		FilePut(ToFile, myFlags)
		' Save Topics for Conversation
		'UPGRADE_WARNING: Put was upgraded to FilePut and has a new behavior. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="9B7D5ADD-D8FE-4819-A36C-6DEDAF088CC7"'
		FilePut(ToFile, myTopics.Count())
		For c = 1 To myTopics.Count()
			'UPGRADE_WARNING: Couldn't resolve default property of object myTopics().SaveToFile. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
			myTopics.Item(c).SaveToFile(ToFile)
		Next c
	End Sub
	
	'UPGRADE_NOTE: Class_Initialize was upgraded to Class_Initialize_Renamed. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="A9E4979A-37FA-4718-9994-97DD76ED70A7"'
	Private Sub Class_Initialize_Renamed()
		myName = "Undefined"
	End Sub
	Public Sub New()
		MyBase.New()
		Class_Initialize_Renamed()
	End Sub
End Class