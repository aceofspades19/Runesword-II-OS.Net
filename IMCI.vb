Option Strict Off
Option Explicit On
Friend Class IMCI
	
	Private Const MCIERR_BASE As Short = 256
	Private Const MCIERR_BAD_CONSTANT As Integer = (MCIERR_BASE + 34)
	Private Const MCIERR_BAD_INTEGER As Integer = (MCIERR_BASE + 14)
	Private Const MCIERR_BAD_TIME_FORMAT As Integer = (MCIERR_BASE + 37)
	Private Const MCIERR_CANNOT_LOAD_DRIVER As Integer = (MCIERR_BASE + 10)
	Private Const MCIERR_CANNOT_USE_ALL As Integer = (MCIERR_BASE + 23)
	Private Const MCIERR_CREATEWINDOW As Integer = (MCIERR_BASE + 91)
	Private Const MCIERR_CUSTOM_DRIVER_BASE As Integer = (MCIERR_BASE + 256)
	Private Const MCIERR_DEVICE_LENGTH As Integer = (MCIERR_BASE + 54)
	Private Const MCIERR_DEVICE_LOCKED As Integer = (MCIERR_BASE + 32)
	Private Const MCIERR_DEVICE_NOT_INSTALLED As Integer = (MCIERR_BASE + 50)
	Private Const MCIERR_DEVICE_NOT_READY As Integer = (MCIERR_BASE + 20)
	Private Const MCIERR_DEVICE_OPEN As Integer = (MCIERR_BASE + 9)
	Private Const MCIERR_DEVICE_ORD_LENGTH As Integer = (MCIERR_BASE + 55)
	Private Const MCIERR_DEVICE_TYPE_REQUIRED As Integer = (MCIERR_BASE + 31)
	Private Const MCIERR_DRIVER As Integer = (MCIERR_BASE + 22)
	Private Const MCIERR_DRIVER_INTERNAL As Integer = (MCIERR_BASE + 16)
	Private Const MCIERR_DUPLICATE_ALIAS As Integer = (MCIERR_BASE + 33)
	Private Const MCIERR_DUPLICATE_FLAGS As Integer = (MCIERR_BASE + 39)
	Private Const MCIERR_EXTENSION_NOT_FOUND As Integer = (MCIERR_BASE + 25)
	Private Const MCIERR_EXTRA_CHARACTERS As Integer = (MCIERR_BASE + 49)
	Private Const MCIERR_FILE_NOT_FOUND As Integer = (MCIERR_BASE + 19)
	Private Const MCIERR_FILE_NOT_SAVED As Integer = (MCIERR_BASE + 30)
	Private Const MCIERR_FILE_READ As Integer = (MCIERR_BASE + 92)
	Private Const MCIERR_FILE_WRITE As Integer = (MCIERR_BASE + 93)
	Private Const MCIERR_FILENAME_REQUIRED As Integer = (MCIERR_BASE + 48)
	Private Const MCIERR_FLAGS_NOT_COMPATIBLE As Integer = (MCIERR_BASE + 28)
	Private Const MCIERR_GET_CD As Integer = (MCIERR_BASE + 51)
	Private Const MCIERR_HARDWARE As Integer = (MCIERR_BASE + 6)
	Private Const MCIERR_ILLEGAL_FOR_AUTO_OPEN As Integer = (MCIERR_BASE + 47)
	Private Const MCIERR_INTERNAL As Integer = (MCIERR_BASE + 21)
	Private Const MCIERR_INVALID_DEVICE_ID As Integer = (MCIERR_BASE + 1)
	Private Const MCIERR_INVALID_DEVICE_NAME As Integer = (MCIERR_BASE + 7)
	Private Const MCIERR_INVALID_FILE As Integer = (MCIERR_BASE + 40)
	Private Const MCIERR_MISSING_COMMAND_STRING As Integer = (MCIERR_BASE + 11)
	Private Const MCIERR_MISSING_DEVICE_NAME As Integer = (MCIERR_BASE + 36)
	Private Const MCIERR_MISSING_PARAMETER As Integer = (MCIERR_BASE + 17)
	Private Const MCIERR_MISSING_STRING_ARGUMENT As Integer = (MCIERR_BASE + 13)
	Private Const MCIERR_MULTIPLE As Integer = (MCIERR_BASE + 24)
	Private Const MCIERR_MUST_USE_SHAREABLE As Integer = (MCIERR_BASE + 35)
	Private Const MCIERR_NEW_REQUIRES_ALIAS As Integer = (MCIERR_BASE + 43)
	Private Const MCIERR_NO_CLOSING_QUOTE As Integer = (MCIERR_BASE + 38)
	Private Const MCIERR_NO_ELEMENT_ALLOWED As Integer = (MCIERR_BASE + 45)
	Private Const MCIERR_NO_INTEGER As Integer = (MCIERR_BASE + 56)
	Private Const MCIERR_NO_WINDOW As Integer = (MCIERR_BASE + 90)
	Private Const MCIERR_NONAPPLICABLE_FUNCTION As Integer = (MCIERR_BASE + 46)
	Private Const MCIERR_NOTIFY_ON_AUTO_OPEN As Integer = (MCIERR_BASE + 44)
	Private Const MCIERR_NULL_PARAMETER_BLOCK As Integer = (MCIERR_BASE + 41)
	Private Const MCIERR_OUT_OF_MEMORY As Integer = (MCIERR_BASE + 8)
	Private Const MCIERR_OUTOFRANGE As Integer = (MCIERR_BASE + 26)
	Private Const MCIERR_PARAM_OVERFLOW As Integer = (MCIERR_BASE + 12)
	Private Const MCIERR_PARSER_INTERNAL As Integer = (MCIERR_BASE + 15)
	Private Const MCIERR_SEQ_DIV_INCOMPATIBLE As Integer = (MCIERR_BASE + 80)
	Private Const MCIERR_SEQ_NOMIDIPRESENT As Integer = (MCIERR_BASE + 87)
	Private Const MCIERR_SEQ_PORT_INUSE As Integer = (MCIERR_BASE + 81)
	Private Const MCIERR_SEQ_PORT_MAPNODEVICE As Integer = (MCIERR_BASE + 83)
	Private Const MCIERR_SEQ_PORT_MISCERROR As Integer = (MCIERR_BASE + 84)
	Private Const MCIERR_SEQ_PORT_NONEXISTENT As Integer = (MCIERR_BASE + 82)
	Private Const MCIERR_SEQ_PORTUNSPECIFIED As Integer = (MCIERR_BASE + 86)
	Private Const MCIERR_SEQ_TIMER As Integer = (MCIERR_BASE + 85)
	Private Const MCIERR_SET_CD As Integer = (MCIERR_BASE + 52)
	Private Const MCIERR_SET_DRIVE As Integer = (MCIERR_BASE + 53)
	Private Const MCIERR_UNNAMED_RESOURCE As Integer = (MCIERR_BASE + 42)
	Private Const MCIERR_UNRECOGNIZED_COMMAND As Integer = (MCIERR_BASE + 5)
	Private Const MCIERR_UNRECOGNIZED_KEYWORD As Integer = (MCIERR_BASE + 3)
	Private Const MCIERR_UNSUPPORTED_FUNCTION As Integer = (MCIERR_BASE + 18)
	Private Const MCIERR_WAVE_INPUTSINUSE As Integer = (MCIERR_BASE + 66)
	Private Const MCIERR_WAVE_INPUTSUNSUITABLE As Integer = (MCIERR_BASE + 72)
	Private Const MCIERR_WAVE_INPUTUNSPECIFIED As Integer = (MCIERR_BASE + 69)
	Private Const MCIERR_WAVE_OUTPUTSINUSE As Integer = (MCIERR_BASE + 64)
	Private Const MCIERR_WAVE_OUTPUTSUNSUITABLE As Integer = (MCIERR_BASE + 70)
	Private Const MCIERR_WAVE_OUTPUTUNSPECIFIED As Integer = (MCIERR_BASE + 68)
	Private Const MCIERR_WAVE_SETINPUTINUSE As Integer = (MCIERR_BASE + 67)
	Private Const MCIERR_WAVE_SETINPUTUNSUITABLE As Integer = (MCIERR_BASE + 73)
	Private Const MCIERR_WAVE_SETOUTPUTINUSE As Integer = (MCIERR_BASE + 65)
	Private Const MCIERR_WAVE_SETOUTPUTUNSUITABLE As Integer = (MCIERR_BASE + 71)
	Private Const MCIERR_VIDEO_NOCODEX As Integer = (MCIERR_BASE + 257)
	
	'UPGRADE_ISSUE: Declaring a parameter 'As Any' is not supported. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="FAE78A8D-8978-4FD4-8208-5B7324A8F795"'
    Private Declare Function mciSendString Lib "winmm.dll" Alias "mciSendStringA" (ByVal lpstrCommand As String, ByVal lpstrReturnString As Object, ByVal uReturnLength As Integer, ByVal hwndCallback As Integer) As Integer
	
	Public Enum VIDEOSTATE
		vsNOT_READY = 0
		vsPAUSED = 1
		vsPLAYING = 2
		vsSTOPPED = 3
		vsERROR = 4
	End Enum
	
	Private Enum DEVICETYPE
		vsUnknown = 0
		vsAudtio = 1
		vsVideo = 2
	End Enum
	
	Dim m_DeviceType As DEVICETYPE
	Dim m_Error As Integer
	
	Dim m_Filename As String
	Dim m_Wait As Boolean
	Dim m_hWnd As Integer
	
	Dim m_Left As Integer
	Dim m_Top As Integer
	Dim m_Width As Integer
	Dim m_Height As Integer
	Dim m_IsOpen As Boolean
	Dim m_Alias As String

    Public Shared Function Spaces(ByVal numSpace As Integer)
        Dim i As Integer
        Dim s As String
        For i = 0 To numSpace
            s += " "
        Next
        Return s
    End Function
    'UPGRADE_NOTE: Class_Initialize was upgraded to Class_Initialize_Renamed. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="A9E4979A-37FA-4718-9994-97DD76ED70A7"'
    Private Sub Class_Initialize_Renamed()
        m_IsOpen = False
        m_Error = 0
        m_hWnd = 0
        m_Filename = ""
        m_Wait = False
        m_DeviceType = DEVICETYPE.vsUnknown
        m_Alias = ""

        m_Left = 0
        m_Top = 0
        m_Width = 0
        m_Height = 0
    End Sub
    Public Sub New()
        MyBase.New()
        Class_Initialize_Renamed()
    End Sub

    Public Function Status() As VIDEOSTATE
        'UPGRADE_NOTE: Left was upgraded to Left_Renamed. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="A9E4979A-37FA-4718-9994-97DD76ED70A7"'

        'UPGRADE_NOTE: Left was upgraded to Left_Renamed. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="A9E4979A-37FA-4718-9994-97DD76ED70A7"'
        'UPGRADE_NOTE: Left was upgraded to Left_Renamed. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="A9E4979A-37FA-4718-9994-97DD76ED70A7"'
        'UPGRADE_NOTE: Left was upgraded to Left_Renamed. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="A9E4979A-37FA-4718-9994-97DD76ED70A7"'
        Dim sParse As String
      

        m_Error = mciSendString("status " & m_Alias & " mode", sParse, Len(sParse) - 1, 0)
        If m_Error Then GoTo Err_Handler
        If sParse = "not ready" Then
            Status = VIDEOSTATE.vsNOT_READY
        ElseIf sParse = "paused" Then
            Status = VIDEOSTATE.vsPAUSED
        ElseIf sParse = "playing" Then
            Status = VIDEOSTATE.vsPLAYING
        ElseIf sParse = "stopped" Then
            Status = VIDEOSTATE.vsSTOPPED
        Else
            Status = VIDEOSTATE.vsERROR
        End If
        'Debug.Print sParse
        Exit Function
Err_Handler:
        Status = VIDEOSTATE.vsERROR
    End Function

    Public Function InitAudio(ByRef Filename As String, ByRef Name As String) As Boolean
        Dim sCmd As String
        If m_IsOpen Then
            Call mciClose()
            m_IsOpen = False
            m_Error = 0
            m_Filename = ""
            m_Wait = False
            m_DeviceType = DEVICETYPE.vsUnknown
            m_Alias = ""
        End If

        m_Alias = Name
        m_Filename = Filename

        InitAudio = True
        sCmd = "Open " & Chr(34) & m_Filename & Chr(34) & " Alias " & m_Alias
        'Debug.Print sCmd
        m_Error = mciSendString(sCmd, 0, 0, 0)
        If m_Error Then GoTo Err_Handler
        m_DeviceType = DEVICETYPE.vsAudtio
        m_IsOpen = True

        Exit Function
Err_Handler:
        Call mciClose()
        InitAudio = False
    End Function

    Public Function InitAVI(ByRef Filename As String, ByRef hWnd As Integer, ByRef Name As String, Optional ByRef Align As System.Windows.Forms.DockStyle = 0) As Boolean
        'UPGRADE_NOTE: Left was upgraded to Left_Renamed. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="A9E4979A-37FA-4718-9994-97DD76ED70A7"'

        Dim sParse As New VB6.FixedLengthString(128)
        Dim lPos As Integer
        Dim lStart As Integer
        Dim sCmd As String

        InitAVI = True

        m_Alias = Name
        m_Filename = Filename
        m_hWnd = hWnd

        'sCmd = "Open " & Chr$(34) & m_Filename & Chr$(34) & " Type avivideo Alias " & m_Alias & " Parent " & m_hWnd & " Style " & &H40000000
        sCmd = "Open " & Chr(34) & m_Filename & Chr(34) & " Alias " & m_Alias & " Parent " & m_hWnd & " Style " & &H40000000
        'Debug.Print sCmd
        m_Error = mciSendString(sCmd, 0, 0, 0)
        If m_Error Then GoTo Err_Handler
        m_IsOpen = True
        m_DeviceType = DEVICETYPE.vsVideo
        ' [borf] check to see if the video was decodeable
        m_Error = mciSendString("Status " & m_Alias & " video", sParse.Value, Len(sParse.Value) - 1, 0)
        If m_Error Then GoTo Err_Handler
        If sParse Is "off" Then
            m_Error = MCIERR_VIDEO_NOCODEX
            GoTo Err_Handler
        End If
        ' Get Width & Height of the Video
        sParse = Spaces(128)
        m_Error = mciSendString("Where " & m_Alias & " Destination", sParse.Value, Len(sParse.Value) - 1, 0)
        If m_Error Then GoTo Err_Handler

        lStart = InStr(1, sParse.Value, " ") 'pos of top
        lPos = InStr(lStart + 1, sParse.Value, " ") 'pos of left
        lStart = InStr(lPos + 1, sParse.Value, " ") 'pos width
        m_Width = CInt(Mid(sParse.Value, lPos, lStart - lPos))
        m_Height = CInt(Mid(sParse.Value, lStart + 1))
        If m_Error Then GoTo Err_Handler

        Exit Function
Err_Handler:
        Call mciClose()
        InitAVI = False
    End Function

    Public Sub Play()
        If Not m_IsOpen Then GoTo Err_Handler

        If m_DeviceType = DEVICETYPE.vsVideo Then
            m_Error = mciSendString("put " & m_Alias & " window at " & m_Left & " " & m_Top & " " & m_Width & " " & m_Height, 0, 0, 0)
            If m_Error Then GoTo Err_Handler

            m_Error = mciSendString("play " & m_Alias & " from 0" & IIf(m_Wait, " WAIT", ""), 0, 0, 0)
            If m_Error Then GoTo Err_Handler
        Else
            m_Error = mciSendString("play " & m_Alias & " from 0", 0, 0, 0)
        End If

        Exit Sub
Err_Handler:

    End Sub

    Public Sub PausePlay()
        If Not m_IsOpen Then GoTo Err_Handler
        m_Error = mciSendString("pause " & m_Alias, 0, 0, 0)
        Exit Sub
Err_Handler:
    End Sub

    Public Sub StopPlay()
        If Not m_IsOpen Then GoTo Err_Handler
        m_Error = mciSendString("stop " & m_Alias, 0, 0, 0)
        Exit Sub
Err_Handler:
    End Sub

    Public Sub ResumePlay()
        If Not m_IsOpen Then GoTo Err_Handler
        m_Error = mciSendString("resume " & m_Alias, 0, 0, 0)
        Exit Sub
Err_Handler:
    End Sub

    Public Sub mciClose(Optional ByRef CloseAll As Boolean = False)
        If Not m_IsOpen Then GoTo Err_Handler
        If CloseAll Then
            m_Error = mciSendString("close all", 0, 0, 0)
        Else
            m_Error = mciSendString("close " & m_Alias, 0, 0, 0)
        End If
        m_IsOpen = False
        Exit Sub
Err_Handler:
    End Sub

    Public ReadOnly Property GetError() As Integer
        Get
            GetError = m_Error
        End Get
    End Property


    Public Property Height() As Integer
        Get
            Height = m_Height
        End Get
        Set(ByVal Value As Integer)
            m_Height = Value
        End Set
    End Property


    Public Property Top() As Integer
        Get
            Top = m_Top
        End Get
        Set(ByVal Value As Integer)
            m_Top = Value
        End Set
    End Property


    'UPGRADE_NOTE: Left was upgraded to Left_Renamed. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="A9E4979A-37FA-4718-9994-97DD76ED70A7"'
    Public Property Left_Renamed() As Integer
        Get
            Left_Renamed = m_Left
        End Get
        Set(ByVal Value As Integer)
            m_Left = Value
        End Set
    End Property


    Public Property Width() As Integer
        Get
            Width = m_Width
        End Get
        Set(ByVal Value As Integer)
            m_Width = Value
        End Set
    End Property


    Public Property Wnd() As Integer
        Get
            Wnd = m_hWnd
        End Get
        Set(ByVal Value As Integer)
            m_hWnd = Value
        End Set
    End Property


    Public Property Wait() As Boolean
        Get
            Wait = m_Wait
        End Get
        Set(ByVal Value As Boolean)
            m_Wait = Value
        End Set
    End Property
End Class