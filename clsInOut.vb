Option Strict Off
Option Explicit On
Friend Class clsInOut
	
	Private Declare Function GetShortPathName Lib "kernel32"  Alias "GetShortPathNameA"(ByVal lpszLongPath As String, ByVal lpszShortPath As String, ByVal lBuffer As Integer) As Integer
	Dim oFs As Object
	
	Public Enum IOActionType
		File = 0
		Folder = 1
	End Enum
	
	Public Function CheckExists(ByRef FilePathName As String, ByRef CheckType As IOActionType, Optional ByRef Create As Boolean = False) As Boolean
		Dim a As Object
		CheckExists = False
		If CheckType = IOActionType.File Then
			'UPGRADE_WARNING: Couldn't resolve default property of object oFs.fileExists. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
			If oFs.fileExists(FilePathName) Then
				CheckExists = True
			Else
				If Create Then
					'UPGRADE_WARNING: Couldn't resolve default property of object oFs.CreateTextFile. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
					a = oFs.CreateTextFile(FilePathName, True)
					'UPGRADE_WARNING: Couldn't resolve default property of object a.Close. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
					a.Close()
					CheckExists = True
				End If
			End If
		ElseIf CheckType = IOActionType.Folder Then 
			'UPGRADE_WARNING: Couldn't resolve default property of object oFs.FolderExists. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
			If oFs.FolderExists(FilePathName) Then
				CheckExists = True
			Else
				If Create Then
					'UPGRADE_WARNING: Couldn't resolve default property of object oFs.CreateFolder. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
					oFs.CreateFolder(FilePathName)
				End If
			End If
		End If
	End Function
	
	Public Function GetFileType(ByRef Filename As String) As String
		Dim f As Object
		
		On Error GoTo Err_Handler
		'UPGRADE_WARNING: Couldn't resolve default property of object oFs.GetFile. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
		f = oFs.GetFile(Filename)
		'UPGRADE_WARNING: Couldn't resolve default property of object f.Type. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
		GetFileType = f.Type
		Exit Function
		
Err_Handler: 
		Err.Clear()
		GetFileType = ""
	End Function
	
	Public Function Delete(ByRef ioName As String, ByRef ioType As IOActionType, Optional ByRef Force As Boolean = False) As Boolean
		
		On Error GoTo Err_Handler
		If ioType = IOActionType.File Then
			'UPGRADE_WARNING: Couldn't resolve default property of object oFs.DeleteFile. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
			Call oFs.DeleteFile(ioName, Force)
		Else
			'UPGRADE_WARNING: Couldn't resolve default property of object oFs.DeleteFolder. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
			Call oFs.DeleteFolder(ioName, Force)
		End If
		
Exit_Function: 
		Delete = True
		Exit Function
Err_Handler: 
		Err.Clear()
		Delete = False
	End Function
	
	Public Function Copy(ByRef Src As String, ByRef Dest As String, ByRef ioType As IOActionType, Optional ByRef Forced As Boolean = True) As Object
		
		On Error GoTo Err_Handler
		If ioType = IOActionType.File Then
			'UPGRADE_WARNING: Couldn't resolve default property of object oFs.CopyFile. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
			Call oFs.CopyFile(Src, Dest, Forced)
		Else
			'UPGRADE_WARNING: Couldn't resolve default property of object oFs.DeleteFolder. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
			Call oFs.DeleteFolder(Src, Dest, Forced)
		End If
		
Exit_Function: 
		'UPGRADE_WARNING: Couldn't resolve default property of object Copy. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
		Copy = True
		Exit Function
Err_Handler: 
		Err.Clear()
		'UPGRADE_WARNING: Couldn't resolve default property of object Copy. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
		Copy = False
	End Function
	
	Public Function AssignTmpName(ByRef ioType As IOActionType) As String
		'UPGRADE_WARNING: Couldn't resolve default property of object oFs.GetTempName. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
		AssignTmpName = oFs.GetTempName
		If ioType = IOActionType.Folder Then
			AssignTmpName = Left(AssignTmpName, InStr(1, AssignTmpName, ".") - 1)
		End If
	End Function
	
	Public Function ShortPath(ByRef Path As String) As String
		Dim strCmdLine As String
		Dim retval As Integer
		' convert path to its short path form
		strCmdLine = New String(Chr(0), 255)
		retval = GetShortPathName(Path, strCmdLine, 254)
		ShortPath = Left(strCmdLine, retval)
	End Function
	
	Public Function Move(ByRef ioName As String, ByRef ioType As IOActionType, Optional ByRef NewName As String = "", Optional ByRef CreateTmp As Boolean = False) As String
		Dim f As Object
		
		On Error GoTo Err_Handler
		If CreateTmp Then
			'UPGRADE_WARNING: Couldn't resolve default property of object oFs.GetTempName. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
			NewName = ClipPath(ioName) & "\" & oFs.GetTempName
		End If
		
		If ioType = IOActionType.File Then
			'UPGRADE_WARNING: Couldn't resolve default property of object oFs.MoveFile. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
			oFs.MoveFile(ioName, NewName)
		Else
			'UPGRADE_WARNING: Couldn't resolve default property of object oFs.MoveFolder. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
			oFs.MoveFolder(ioName, NewName)
		End If
		Move = NewName
		Exit Function
		
Err_Handler: 
		Err.Clear()
		Move = ""
	End Function
	
	Private Sub InitFileSystem()
		If Not oFs Is Nothing Then Exit Sub
		On Error GoTo Err_Handler
		oFs = fCreateObject("Scripting.FileSystemObject")
		Exit Sub
Err_Handler: 
		'UPGRADE_NOTE: Object oFs may not be destroyed until it is garbage collected. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6E35BFF6-CD74-4B09-9689-3E1A43DF8969"'
		oFs = Nothing
	End Sub
	
	'UPGRADE_NOTE: Class_Initialize was upgraded to Class_Initialize_Renamed. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="A9E4979A-37FA-4718-9994-97DD76ED70A7"'
	Private Sub Class_Initialize_Renamed()
		Call InitFileSystem()
	End Sub
	Public Sub New()
		MyBase.New()
		Class_Initialize_Renamed()
	End Sub
	
	'UPGRADE_NOTE: Class_Terminate was upgraded to Class_Terminate_Renamed. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="A9E4979A-37FA-4718-9994-97DD76ED70A7"'
	Private Sub Class_Terminate_Renamed()
		'UPGRADE_NOTE: Object oFs may not be destroyed until it is garbage collected. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6E35BFF6-CD74-4B09-9689-3E1A43DF8969"'
		oFs = Nothing
	End Sub
	Protected Overrides Sub Finalize()
		Class_Terminate_Renamed()
		MyBase.Finalize()
	End Sub
End Class