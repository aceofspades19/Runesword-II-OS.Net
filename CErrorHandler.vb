Option Strict Off
Option Explicit On
Friend Class CErrorHandler
	
	Public Enum ErrorOutput
		LOG_NONE = 0
		LOG_TOFILE = 2 ' To file
		LOG_DEBUG = 4 ' To debug.print
		'    LOG_LEVEL3 = 8
	End Enum
	
	Public Enum ErrorLevel
		ERR_NONE = 0
		ERR_GENERAL = 1 ' To file
		ERR_DEBUG = 2 ' To file
		ERR_CRITICAL = 3 ' To debug.print
		'    LOG_LEVEL3 = 8
	End Enum
	
	Private oFs As Object
	Private oFile As Object
	Private m_Filename As String
	'Private m_Show As Integer
	Private m_ErrorOutput As ErrorOutput
	Private m_ErrorLevel As ErrorLevel
	
	Public Sub Initialize(ByRef Output As ErrorOutput, ByRef errLevel As ErrorLevel, Optional ByRef FileName As String = "")
		On Error GoTo Err_Handler
		m_ErrorOutput = Output
		m_ErrorLevel = errLevel
		If m_ErrorOutput = ErrorOutput.LOG_NONE Then Exit Sub
		If oFs Is Nothing Then InitFileSystem()
		'm_Show = Show
		If FileName = "" Then FileName = "ErrLog.txt"
		m_Filename = My.Application.Info.DirectoryPath & "\" & FileName
		Exit Sub
Err_Handler: 
		'Debug.Print "Good gosh can you believe it an error in the error handler. DOH"
	End Sub
	
	Public Sub LogText(ByRef Comments As String)
		If m_ErrorOutput = ErrorOutput.LOG_NONE Then Exit Sub
		logError(Comments)
		'On Error GoTo Err_Handler
		'    If oFs Is Nothing Then InitFileSystem
		'    If oFile Is Nothing Then OpenFile
		'    oFile.WriteLine (Comments)
		'    If m_Show Then Debug.Print Comments
		'    Exit Sub
		'Err_Handler:
		'    'Debug.Print "Good gosh can you believe it an error in the error handler. DOH"
	End Sub
	
	'Public Sub logError(Comments As String, Optional errLevel As ErrorLevel = ERR_GENERAL)
	'    If m_ErrorOutput = LOG_NONE Then Exit Sub
	'    If errLevel < m_ErrorLevel Then Exit Sub
	'    Dim sMsg As String
	'    Dim ErrOut As ErrorOutput
	'    Dim timestamp As String
	'
	'    timestamp = Format$(Now, "yyyy mm dd hh:mm:ss")
	''On Error GoTo Err_Handler
	'    'Comments = "[" & timestamp & "] " & Comments '& vbCrLf
	'    sMsg = "[" & timestamp & "] " & vbTab & "Source:" & Err.Source & vbCrLf & vbTab & "Description:" & Err.Description & "  (" & Err.Number & ")"
	'    If m_ErrorOutput And LOG_TOFILE Then
	'        If oFs Is Nothing Then InitFileSystem
	'        If oFile Is Nothing Then OpenFile
	'        If Err.Number <> 0 Then
	'            oFile.WriteLine (sMsg)
	'        End If
	'        oFile.WriteLine ("[" & timestamp & "] " & Comments)
	'    End If
	'    'If m_Show Then Debug.Print Comments & vbCrLf & sMsg
	'    If m_ErrorOutput And LOG_DEBUG Then
	'        If Err.Number <> 0 Then
	'            Debug.Print sMsg
	'        End If
	'        Debug.Print "[" & timestamp & "] " & Comments
	'    End If
	'    If Err.Number > 0 And errLevel = ERR_CRITICAL Then
	'        Comments = Comments
	'        sMsg = "Description: " & Err.Description & "(#" & Err.Number & ")" & vbCrLf & _
	''               "     Source: " & Err.Source & vbCrLf & _
	''               "    Comment: " & Comments
	'
	'        MsgBox sMsg, vbCritical, "System Error"
	'    End If
	'    Exit Sub
	''Err_Handler:
	'    'Debug.Print "Good gosh can you believe it an error in the error handler. DOH"
	'End Sub
	Public Sub logError(ByRef Comments As String, Optional ByRef errLevel As ErrorLevel = ErrorLevel.ERR_GENERAL)
		'    If errLevel < m_ErrorLevel Then GoTo ERR_CHECK_CRITICAL
		If m_ErrorOutput = ErrorOutput.LOG_NONE Then GoTo ERR_CHECK_CRITICAL
		Dim sMsg As String
		Dim ErrOut As ErrorOutput
		Dim timestamp As String
		Dim intErrNumber As Short
		timestamp = VB6.Format(Now, "yyyy mm dd hh:mm:ss")
		'On Error GoTo Err_Handler
		'Comments = "[" & timestamp & "] " & Comments '& vbCrLf
		sMsg = "[" & timestamp & "] " & vbCrLf & vbTab & "Source:" & Err.Source & vbCrLf & vbTab & "Description:" & Err.Description & "  (" & Err.Number & ")"
		If m_ErrorOutput And ErrorOutput.LOG_TOFILE Then
			intErrNumber = Err.Number
			If oFs Is Nothing Then InitFileSystem() ' [Titi] due to the error management in the procedure, err.number will be reset
			If oFile Is Nothing Then OpenFile() ' same comment here
			If intErrNumber <> 0 Then
				'UPGRADE_WARNING: Couldn't resolve default property of object oFile.WriteLine. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
				oFile.WriteLine(sMsg)
			End If
			'UPGRADE_WARNING: Couldn't resolve default property of object oFile.WriteLine. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
			oFile.WriteLine("[" & timestamp & "] " & Comments)
		End If
		'If m_Show Then Debug.Print Comments & vbCrLf & sMsg
		If m_ErrorOutput And ErrorOutput.LOG_DEBUG Then
			If Err.Number <> 0 Then
				Debug.Print(sMsg)
			End If
			Debug.Print("[" & timestamp & "] " & Comments)
		End If
ERR_CHECK_CRITICAL: 
		If Err.Number > 0 And errLevel = ErrorLevel.ERR_CRITICAL Then
			'Comments = Comments
			sMsg = "Description: " & Err.Description & "(#" & Err.Number & ")" & vbCrLf & "     Source: " & Err.Source & vbCrLf & "    Comment: " & Comments
			
			MsgBox(sMsg, MsgBoxStyle.Critical, "System Error")
		ElseIf errLevel = ErrorLevel.ERR_CRITICAL Then 
		End If
		Exit Sub
		'Err_Handler:
		'Debug.Print "Good gosh can you believe it an error in the error handler. DOH"
	End Sub
	
	'UPGRADE_NOTE: Reset was upgraded to Reset_Renamed. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="A9E4979A-37FA-4718-9994-97DD76ED70A7"'
	Public Sub Reset_Renamed()
		If m_ErrorOutput = ErrorOutput.LOG_NONE Then Exit Sub
		On Error GoTo Err_Handler
		If oFs Is Nothing Then InitFileSystem()
		If oFile Is Nothing Then Initialize(m_ErrorOutput, m_ErrorLevel)
		'UPGRADE_WARNING: Couldn't resolve default property of object oFile.Close. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
		oFile.Close()
		'UPGRADE_WARNING: Couldn't resolve default property of object oFs.CreateTextFile. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
		oFile = oFs.CreateTextFile(m_Filename, True)
		Exit Sub
Err_Handler: 
		Debug.Print("Good gosh can you believe it an error in the error handler. DOH")
	End Sub
	
	Private Sub OpenFile()
		If m_ErrorOutput = ErrorOutput.LOG_NONE Then Exit Sub
		On Error GoTo Err_Handler
		'UPGRADE_WARNING: Couldn't resolve default property of object oFs.CreateTextFile. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
		oFile = oFs.CreateTextFile(m_Filename, True)
		Exit Sub
Err_Handler: 
		'UPGRADE_NOTE: Object oFile may not be destroyed until it is garbage collected. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6E35BFF6-CD74-4B09-9689-3E1A43DF8969"'
		oFile = Nothing
	End Sub
	
	Private Sub InitFileSystem()
		If m_ErrorOutput = ErrorOutput.LOG_NONE Then Exit Sub
		If Not oFs Is Nothing Then Exit Sub
		On Error GoTo Err_Handler
		oFs = fCreateObject("Scripting.FileSystemObject")
		Exit Sub
Err_Handler: 
		Err.Clear()
		'UPGRADE_NOTE: Object oFs may not be destroyed until it is garbage collected. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6E35BFF6-CD74-4B09-9689-3E1A43DF8969"'
		oFs = Nothing
	End Sub
	
	'UPGRADE_NOTE: Class_Initialize was upgraded to Class_Initialize_Renamed. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="A9E4979A-37FA-4718-9994-97DD76ED70A7"'
	Private Sub Class_Initialize_Renamed()
		'    Call InitFileSystem
	End Sub
	Public Sub New()
		MyBase.New()
		Class_Initialize_Renamed()
	End Sub
	
	'UPGRADE_NOTE: Class_Terminate was upgraded to Class_Terminate_Renamed. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="A9E4979A-37FA-4718-9994-97DD76ED70A7"'
	Private Sub Class_Terminate_Renamed()
		If m_ErrorOutput = ErrorOutput.LOG_NONE Then Exit Sub
		'UPGRADE_WARNING: Couldn't resolve default property of object oFile.Close. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
		If Not oFile Is Nothing Then oFile.Close()
		'UPGRADE_NOTE: Object oFile may not be destroyed until it is garbage collected. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6E35BFF6-CD74-4B09-9689-3E1A43DF8969"'
		oFile = Nothing
		'UPGRADE_NOTE: Object oFs may not be destroyed until it is garbage collected. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6E35BFF6-CD74-4B09-9689-3E1A43DF8969"'
		oFs = Nothing
	End Sub
	Protected Overrides Sub Finalize()
		Class_Terminate_Renamed()
		MyBase.Finalize()
	End Sub
	
	Private Function fCreateObject(ByRef SID As String) As Object
		On Error GoTo Err_Handler
		fCreateObject = CreateObject(SID)
		Exit Function
Err_Handler: 
		Debug.Print(Err.Description & ": '" & SID & "'" & " (" & Err.Number & ")")
	End Function
End Class