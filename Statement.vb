Option Strict Off
Option Explicit On
Friend Class Statement
	
	' Index while contained in a Collection
	Private myIndex As Short
	' Type of Statement (see Routine which returns text)
	Private myStatement As Byte
	' Each Statement Type has certain attributes stored here
	Private myB(6) As Byte
	' Can have some text associated with it
	Private myText As String
	
	
	Public Property Index() As Short
		Get
			Index = myIndex
		End Get
		Set(ByVal Value As Short)
			myIndex = Value
		End Set
	End Property
	
	
	Public Property Statement() As Byte
		Get
			Statement = myStatement
		End Get
		Set(ByVal Value As Byte)
			Dim c As Short
			myStatement = Value
			' Wipe out myB()
			For c = 0 To 6
				myB(c) = 0
			Next c
			' Wipe out myText
			myText = ""
		End Set
	End Property
	
	
	Public Property B(ByVal Index As Short) As Short
		Get
			B = CShort(myB(Index))
		End Get
		Set(ByVal Value As Short)
			myB(Index) = CByte(Value)
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
	
	Public Sub Copy(ByRef FromStatement As Statement)
		Dim c As Short
		' Copy Type and Vars
		myStatement = FromStatement.Statement
		For c = 0 To 6
			myB(c) = FromStatement.B(c)
		Next c
		' Copy Text
		myText = FromStatement.Text
	End Sub
	
	Public Sub ReadFromFile(ByRef FromFile As Short)
		Dim c As Short
		Dim Cnt As Integer
		' Read Index and Vars
		'UPGRADE_WARNING: Get was upgraded to FileGet and has a new behavior. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="9B7D5ADD-D8FE-4819-A36C-6DEDAF088CC7"'
		FileGet(FromFile, myIndex)
		'UPGRADE_WARNING: Get was upgraded to FileGet and has a new behavior. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="9B7D5ADD-D8FE-4819-A36C-6DEDAF088CC7"'
		FileGet(FromFile, myStatement)
		For c = 0 To 6
			'UPGRADE_WARNING: Get was upgraded to FileGet and has a new behavior. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="9B7D5ADD-D8FE-4819-A36C-6DEDAF088CC7"'
			FileGet(FromFile, myB(c))
		Next c
		' Read Text
		'UPGRADE_WARNING: Get was upgraded to FileGet and has a new behavior. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="9B7D5ADD-D8FE-4819-A36C-6DEDAF088CC7"'
		FileGet(FromFile, Cnt)
		myText = ""
		For c = 1 To Cnt
			myText = myText & " "
		Next c
		'UPGRADE_WARNING: Get was upgraded to FileGet and has a new behavior. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="9B7D5ADD-D8FE-4819-A36C-6DEDAF088CC7"'
		FileGet(FromFile, myText)
	End Sub
	
	Public Sub SaveToFile(ByRef ToFile As Short)
		Dim c As Short
		' Save Index and Vars
		'UPGRADE_WARNING: Put was upgraded to FilePut and has a new behavior. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="9B7D5ADD-D8FE-4819-A36C-6DEDAF088CC7"'
		FilePut(ToFile, myIndex)
		'UPGRADE_WARNING: Put was upgraded to FilePut and has a new behavior. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="9B7D5ADD-D8FE-4819-A36C-6DEDAF088CC7"'
		FilePut(ToFile, myStatement)
		For c = 0 To 6
			'UPGRADE_WARNING: Put was upgraded to FilePut and has a new behavior. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="9B7D5ADD-D8FE-4819-A36C-6DEDAF088CC7"'
			FilePut(ToFile, myB(c))
		Next c
		' Save Text
		'UPGRADE_WARNING: Put was upgraded to FilePut and has a new behavior. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="9B7D5ADD-D8FE-4819-A36C-6DEDAF088CC7"'
		FilePut(ToFile, Len(myText))
		'UPGRADE_WARNING: Put was upgraded to FilePut and has a new behavior. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="9B7D5ADD-D8FE-4819-A36C-6DEDAF088CC7"'
		FilePut(ToFile, myText)
	End Sub
	
	'UPGRADE_NOTE: Class_Initialize was upgraded to Class_Initialize_Renamed. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="A9E4979A-37FA-4718-9994-97DD76ED70A7"'
	Private Sub Class_Initialize_Renamed()
		myStatement = 54 ' TriggerComment
		myText = ""
	End Sub
	Public Sub New()
		MyBase.New()
		Class_Initialize_Renamed()
	End Sub
End Class