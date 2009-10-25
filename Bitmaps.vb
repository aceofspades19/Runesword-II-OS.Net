Option Strict Off
Option Explicit On
Imports System.Drawing.Imaging
Imports System.Drawing


Module modBitmaps
	
	' Bitmap functions
    Declare Auto Function BitBlt Lib "GDI32.DLL" ( _
        ByVal hdcDest As IntPtr, _
        ByVal nXDest As Integer, _
        ByVal nYDest As Integer, _
        ByVal nWidth As Integer, _
        ByVal nHeight As Integer, _
        ByVal hdcSrc As IntPtr, _
        ByVal nXSrc As Integer, _
        ByVal nYSrc As Integer, _
        ByVal dwRop As Int32) As Boolean
	'Declare Function GetDIBits Lib "gdi32" (ByVal aHDC As Long, ByVal hBitmap As Long, ByVal nStartScan As Long, ByVal nNumScans As Long, lpBits As Any, lpBI As BITMAPINFO, ByVal wUsage As Long) As Long
    Declare Function StretchDIBits Lib "gdi32" ( _
        ByVal hdc As Long, _
        ByVal DestX As Integer, _
        ByVal DestY As Integer, _
        ByVal wDestWidth As Integer, _
        ByVal wDestHeight As Integer, _
        ByVal SrcX As Integer, _
        ByVal SrcY As Integer, _
        ByVal wSrcWidth As Integer, _
        ByVal wSrcHeight As Integer, _
        ByVal lpBits As Long, _
        ByVal BitsInfo As BITMAPINFO, _
        ByVal wUsage As Integer, _
        ByVal dwRop As Long) As Integer
	Declare Function StretchBlt Lib "gdi32" (ByVal hdc As Integer, ByVal X As Integer, ByVal Y As Integer, ByVal nWidth As Integer, ByVal nHeight As Integer, ByVal hSrcDC As Integer, ByVal xSrc As Integer, ByVal ySrc As Integer, ByVal nSrcWidth As Integer, ByVal nSrcHeight As Integer, ByVal dwRop As Integer) As Integer
	Declare Function SetStretchBltMode Lib "gdi32" (ByVal hdc As Integer, ByVal nStretchMode As Integer) As Integer
	
	Public Const DIB_RGB_COLORS As Short = 0
	
	Public Const SRCCOPY As Integer = &HCC0020
	Public Const SRCAND As Integer = &H8800C6
	Public Const SRCPAINT As Integer = &HEE0086
	
	' Pixel functions
	Declare Function GetPixel Lib "gdi32" (ByVal hdc As Integer, ByVal X As Integer, ByVal Y As Integer) As Integer
	Declare Function SetPixel Lib "gdi32" (ByVal hdc As Integer, ByVal X As Integer, ByVal Y As Integer, ByVal crColor As Integer) As Integer
	
	' Global Heap functions
	Declare Function GlobalAlloc Lib "kernel32" (ByVal wFlags As Integer, ByVal dwBytes As Integer) As Integer
	Declare Function GlobalLock Lib "kernel32" (ByVal hMem As Integer) As Integer
	Declare Function GlobalUnlock Lib "kernel32" (ByVal hMem As Integer) As Integer
	Declare Function GlobalFree Lib "kernel32" (ByVal hMem As Integer) As Integer
	'Declare Function MemoryRead Lib "toolhelp" (wSel As Long, dwOffset As Long, lpvBuf As Long, dwcb As Long) As Long
	
	Public Const GMEM_MOVEABLE As Integer = &H2
	Public Const GMEM_ZEROINIT As Integer = &H40
	
	' File functions

    Declare Function OpenFile Lib "kernel32" (ByVal lpFileName As String, ByVal lpReOpenBuff As OFSTRUCT, ByVal wStyle As Long) As Long
	'UPGRADE_ISSUE: Declaring a parameter 'As Any' is not supported. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="FAE78A8D-8978-4FD4-8208-5B7324A8F795"'
    Declare Function hread Lib "kernel32" Alias "_hread" (ByVal hFile As Integer, ByRef lpBuffer As Object, ByVal lBytes As Integer) As Integer
	Declare Function lclose Lib "kernel32"  Alias "_lclose"(ByVal hFile As Integer) As Integer
	Declare Function llseek Lib "kernel32"  Alias "_llseek"(ByVal hFile As Integer, ByVal lOffset As Integer, ByVal iOrigin As Integer) As Integer
	
	Public Const OF_READ As Integer = &H0
	Public Const OF_SHARE_DENY_NONE As Integer = &H40
	
	' Cursor functions
	'UPGRADE_WARNING: Structure POINTAPI may require marshalling attributes to be passed as an argument in this Declare statement. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="C429C3A5-5D47-4CD9-8F51-74A1616405DC"'
	Declare Function GetCursorPos Lib "user32" (ByRef lpPoint As POINTAPI) As Integer
	
	' Pen and Brush functions
	'UPGRADE_NOTE: GetObject was upgraded to GetObject_Renamed. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="A9E4979A-37FA-4718-9994-97DD76ED70A7"'
	'UPGRADE_ISSUE: Declaring a parameter 'As Any' is not supported. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="FAE78A8D-8978-4FD4-8208-5B7324A8F795"'
	Declare Function GetObject_Renamed Lib "gdi32"  Alias "GetObjectA"(ByVal hObject As Integer, ByVal nCount As Integer, ByRef lpObject As Object) As Integer
	'Declare Function CreatePen Lib "gdi32" (ByVal nPenStyle As Long, ByVal nWidth As Long, ByVal crColor As Long) As Long
	'Declare Function SelectObject Lib "gdi32" (ByVal hdc As Long, ByVal hObject As Long) As Long
	'Declare Function Ellipse Lib "gdi32" (ByVal hdc As Long, ByVal X1 As Long, ByVal Y1 As Long, ByVal X2 As Long, ByVal Y2 As Long) As Long
	'Declare Function CreateSolidBrush Lib "gdi32" (ByVal crColor As Long) As Long
	'Declare Function DeleteObject Lib "gdi32" (ByVal hObject As Long) As Long
	
	' Device Mode functions (Display Settings)
	'UPGRADE_ISSUE: Declaring a parameter 'As Any' is not supported. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="FAE78A8D-8978-4FD4-8208-5B7324A8F795"'
	Declare Function EnumDisplaySettings Lib "user32"  Alias "EnumDisplaySettingsA"(ByVal lpszDeviceName As Integer, ByVal iModeNum As Integer, ByRef lpDevMode As Object) As Boolean
	'UPGRADE_ISSUE: Declaring a parameter 'As Any' is not supported. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="FAE78A8D-8978-4FD4-8208-5B7324A8F795"'
	Declare Function ChangeDisplaySettings Lib "user32"  Alias "ChangeDisplaySettingsA"(ByRef lpDevMode As Object, ByVal dwFlags As Integer) As Integer
	
	' Variables for determining screen resolution
	Declare Function GetSystemMetrics Lib "user32" (ByVal nIndex As Integer) As Integer
	
	Public Const SM_CXSCREEN As Short = 0
	Public Const SM_CYSCREEN As Short = 1
	
	' Variables for Device Mode (to change system color depth and resolution)
	Public Const CCDEVICENAME As Short = 32
	Public Const CCFORMNAME As Short = 32
	Public Const DISP_CHANGE_SUCCESSFUL As Short = 0
	Public Const DISP_CHANGE_RESTART As Short = 1
	Public Const DISP_CHANGE_FAILED As Short = -1
	Public Const DISP_CHANGE_BADMODE As Short = -2
	Public Const DISP_CHANGE_NOTUPDATED As Short = -3
	Public Const DISP_CHANGE_BADFLAGS As Short = -4
	Public Const DISP_CHANGE_BADPARAM As Short = -5
	Public Const CDS_UPDATEREGISTRY As Integer = &H1
	Public Const CDS_TEST As Integer = &H2
	Public Const DM_BITSPERPEL As Integer = &H40000
	Public Const DM_PELSWIDTH As Integer = &H80000
	Public Const DM_PELSHEIGHT As Integer = &H100000
	
	Public Structure DEVMODE
		'UPGRADE_WARNING: Fixed-length string size must fit in the buffer. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="3C1E4426-0B80-443E-B943-0627CD55D48B"'
		<VBFixedString(CCDEVICENAME),System.Runtime.InteropServices.MarshalAs(System.Runtime.InteropServices.UnmanagedType.ByValArray,SizeConst:=CCDEVICENAME)> Public dmDeviceName() As Char
		Dim dmSpecVersion As Short
		Dim dmDriverVersion As Short
		Dim dmSize As Short
		Dim dmDriverExtra As Short
		Dim dmFields As Integer
		Dim dmOrientation As Short
		Dim dmPaperSize As Short
		Dim dmPaperLength As Short
		Dim dmPaperWidth As Short
		Dim dmScale As Short
		Dim dmCopies As Short
		Dim dmDefaultSource As Short
		Dim dmPrintQuality As Short
		Dim dmColor As Short
		Dim dmDuplex As Short
		Dim dmYResolution As Short
		Dim dmTTOption As Short
		Dim dmCollate As Short
		'UPGRADE_WARNING: Fixed-length string size must fit in the buffer. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="3C1E4426-0B80-443E-B943-0627CD55D48B"'
		<VBFixedString(CCFORMNAME),System.Runtime.InteropServices.MarshalAs(System.Runtime.InteropServices.UnmanagedType.ByValArray,SizeConst:=CCFORMNAME)> Public dmFormName() As Char
		Dim dmUnusedPadding As Short
		Dim dmBitsPerPel As Short
		Dim dmPelsWidth As Integer
		Dim dmPelsHeight As Integer
		Dim dmDisplayFlags As Integer
		Dim dmDisplayFrequency As Integer
	End Structure
	
	' Types for various Bitmap functions: read a file, determine Palette, etc.
	Structure POINTAPI
		Dim X As Integer
		Dim Y As Integer
	End Structure
	
	Structure BITMAP
		Dim bmType As Integer
		Dim bmWidth As Integer
		Dim bmHeight As Integer
		Dim bmWidthBytes As Integer
		Dim bmPlanes As Short
		Dim bmBitsPixel As Short
		Dim bmBits As Integer
	End Structure
	
	Structure BITMAPFILEHEADER
		Dim bfType As Short
		Dim bfSize As Integer
		Dim bfReserved1 As Short
		Dim bfReserved2 As Short
		Dim bfOffBits As Integer
	End Structure
	
	Structure BITMAPINFOHEADER
		Dim biSize As Integer
		Dim biWidth As Integer
		Dim biHeight As Integer
		Dim biPlanes As Short
		Dim biBitCount As Short
		Dim biCompression As Integer
		Dim biSizeImage As Integer
		Dim biXPelsPerMeter As Integer
		Dim biYPelsPerMeter As Integer
		Dim biClrUsed As Integer
		Dim biClrImportant As Integer
	End Structure
	
	Structure RGBQUAD
		Dim rgbBlue As Byte
		Dim rgbGreen As Byte
		Dim rgbRed As Byte
		Dim rgbReserved As Byte
	End Structure
	
	Structure BITMAPINFO
		Dim bmiHeader As BITMAPINFOHEADER
		<VBFixedArray(255)> Dim bmiColors() As RGBQUAD
		
		'UPGRADE_TODO: "Initialize" must be called to initialize instances of this structure. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="B4BFF9E0-8631-45CF-910E-62AB3970F27B"'
		Public Sub Initialize()
			ReDim bmiColors(255)
		End Sub
	End Structure
	
	Structure OFSTRUCT
		Dim cBytes As Byte
		Dim fFixedDisk As Byte
		Dim nErrCode As Short
		Dim Reserved1 As Short
		Dim Reserved2 As Short
		<VBFixedArray(127)> Dim szPathName() As Byte
		
		'UPGRADE_TODO: "Initialize" must be called to initialize instances of this structure. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="B4BFF9E0-8631-45CF-910E-62AB3970F27B"'
		Public Sub Initialize()
			ReDim szPathName(127)
		End Sub
    End Structure
    Public Function SDIBits(ByVal pic As PictureBox, ByVal DestX As Integer, ByVal DestY As Integer, ByVal wDestWidth As Integer, _
                            ByVal wDestHeight As Integer, ByVal SrcX As Integer, ByVal srcY As Integer, ByVal wSrvWidth As Integer, _
                            ByVal wSrcHeight As Integer, ByVal lpBits As Long, ByVal BitsInfo As BITMAPINFO, ByVal wUsage As Integer, _
                            ByVal dwRop As Long) As Integer
        Dim src As Graphics = pic.CreateGraphics
        Dim hdc As IntPtr = src.GetHdc
        StretchDIBits(hdc, DestX, DestY, wDestWidth, wDestHeight, SrcX, srcY, wSrvWidth, wSrcHeight, _
                        lpBits, BitsInfo, wUsage, dwRop)
        src.ReleaseHdc(hdc)
    End Function
    'wrapper for win api function setstrechbitmode
    Public Function SetSBMode(ByVal pic As PictureBox, ByVal mode As Integer) As Integer
        Dim src As Graphics = pic.CreateGraphics
        Dim hdc As IntPtr = src.GetHdc
        SetStretchBltMode(hdc, mode)
        src.ReleaseHdc(hdc)
    End Function

    'Wrapper for win api function bitblt
    Public Function CopyRect(ByVal src As PictureBox, _
       ByVal rect As RectangleF, ByVal dwRop As Int32) As System.Drawing.Bitmap
        'Get a Graphics Object from the form

        Dim srcPic As Graphics = src.CreateGraphics
        'Create a EMPTY bitmap from that graphics

        Dim srcBmp As New System.Drawing.Bitmap(src.Width, src.Height, srcPic)
        'Create a Graphics object in memory from that bitmap

        Dim srcMem As Graphics = Graphics.FromImage(srcBmp)

        'get the IntPtr's of the graphics

        Dim HDC1 As IntPtr = srcPic.GetHdc
        'get the IntPtr's of the graphics

        Dim HDC2 As IntPtr = srcMem.GetHdc

        'get the picture 

        BitBlt(HDC2, 0, 0, rect.Width, _
          rect.Height, HDC1, rect.X, rect.Y, dwRop)

        'Clone the bitmap so we can dispose this one 

        CopyRect = srcBmp.Clone()

        'Clean Up 

        srcPic.ReleaseHdc(HDC1)
        srcMem.ReleaseHdc(HDC2)
        srcPic.Dispose()
        srcMem.Dispose()
        srcMem.Dispose()
    End Function

    Public Function ChangeScreenSettings(ByRef lWidth As Short, ByRef lHeight As Short, ByRef lColors As Short) As Short
        Dim tDevMode As DEVMODE
        Dim lTemp, lIndex As Integer
        lIndex = 0
        Do
            'UPGRADE_WARNING: Couldn't resolve default property of object tDevMode. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
            lTemp = EnumDisplaySettings(0, lIndex, tDevMode)
            If lTemp = 0 Then Exit Do
            lIndex = lIndex + 1
            With tDevMode
                If .dmPelsWidth = lWidth And .dmPelsHeight = lHeight And .dmBitsPerPel = lColors Then
                    '                lTemp = ChangeDisplaySettings(tDevMode, CDS_UPDATEREGISTRY)
                    'UPGRADE_WARNING: Couldn't resolve default property of object tDevMode. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
                    lTemp = ChangeDisplaySettings(tDevMode, 0)
                    Exit Do
                End If
            End With
        Loop
        Select Case lTemp
            Case DISP_CHANGE_SUCCESSFUL
                '            MsgBox "The display settings changed successfully", vbInformation
            Case DISP_CHANGE_RESTART
                MsgBox("The computer must be restarted in order for the graphics mode to work", MsgBoxStyle.Question)
            Case DISP_CHANGE_FAILED
                MsgBox("The display driver failed the specified graphics mode", MsgBoxStyle.Critical)
            Case DISP_CHANGE_BADMODE
                MsgBox("The graphics mode is not supported", MsgBoxStyle.Critical)
            Case DISP_CHANGE_NOTUPDATED
                MsgBox("Unable to write settings to the registry", MsgBoxStyle.Critical)
            Case DISP_CHANGE_BADFLAGS
                MsgBox("You Passed invalid data", MsgBoxStyle.Critical)
        End Select
        ChangeScreenSettings = lTemp
    End Function

    Public Sub ReadBitmapFile(ByVal FileName As String, ByRef bmInfo As BITMAPINFO, ByRef hMem As Integer, ByRef TransparentRGB As Integer)
        ' This does indeed read the header and info of a BMP file.
        ' Including (most important here) the Palette information into
        ' the array bmInfo.bmiColors(). The final hread loads the actual
        ' bits to lpMem (a section of memory on the heap).
        Dim rc As Integer
        'UPGRADE_WARNING: Arrays in structure fileStruct may need to be initialized before they can be used. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="814DF224-76BD-4BB4-BFFB-EA359CB9FC48"'
        Dim fileStruct As OFSTRUCT
        Dim bmFileHeader As BITMAPFILEHEADER
        Dim hFile, lpMem As Integer
        Dim BytesRead As Integer
        Dim OneByte As Byte
        ' Open the bitmap file
        hFile = OpenFile(FileName, fileStruct, OF_READ Or OF_SHARE_DENY_NONE)
        ' Read the bitmap's file header info
        'UPGRADE_WARNING: Couldn't resolve default property of object bmFileHeader. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
        BytesRead = hread(hFile, bmFileHeader, Len(bmFileHeader))
        ' Read the bitmap's info (including the palette entries)
        'UPGRADE_WARNING: Couldn't resolve default property of object bmInfo. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
        BytesRead = hread(hFile, bmInfo, Len(bmInfo))
        If bmInfo.bmiHeader.biClrUsed > 0 And bmInfo.bmiHeader.biClrUsed < 256 Then
            '-- Fix for non 256-color BMP's [borfaux]
            rc = llseek(hFile, -((256 * 4) - bmInfo.bmiHeader.biClrUsed * 4), 1)
        End If
        ' Read the next byte and the back up a byte
        BytesRead = hread(hFile, OneByte, 1)
        TransparentRGB = RGB(bmInfo.bmiColors(OneByte).rgbRed, bmInfo.bmiColors(OneByte).rgbGreen, bmInfo.bmiColors(OneByte).rgbBlue)
        rc = llseek(hFile, -1, 1)
        ' Free any existing pointed to memory
        rc = GlobalFree(hMem)
        ' Allocate chunk of memory based on size of bitmap
        '    hMem = GlobalAlloc(GMEM_MOVEABLE Or GMEM_ZEROINIT, _
        ''        (CLng(bmInfo.bmiHeader.biWidth / 4) * 4 + 4) * bmInfo.bmiHeader.biHeight)
        hMem = GlobalAlloc(GMEM_MOVEABLE Or GMEM_ZEROINIT, (CInt(bmInfo.bmiHeader.biWidth / 4) * 4 + 4) * bmInfo.bmiHeader.biHeight)
        ' Lock that chunk of memory
        lpMem = GlobalLock(hMem)
        ' Read the bitmap
        BytesRead = hread(hFile, lpMem, (CInt(bmInfo.bmiHeader.biWidth / 4) * 4 + 4) * bmInfo.bmiHeader.biHeight)
        ' Close the file
        rc = lclose(hFile)
        'Unlock resources
        rc = GlobalUnlock(hMem)
    End Sub

    Public Sub ChangeColor(ByRef bmInfo As BITMAPINFO, ByRef FromColor As Integer, ByRef ToRed As Byte, ByRef ToGreen As Byte, ByRef ToBlue As Byte)
        Dim c As Short
        For c = 0 To 255
            If RGB(bmInfo.bmiColors(c).rgbRed, bmInfo.bmiColors(c).rgbGreen, bmInfo.bmiColors(c).rgbBlue) = FromColor Then
                bmInfo.bmiColors(c).rgbRed = ToRed
                bmInfo.bmiColors(c).rgbGreen = ToGreen
                bmInfo.bmiColors(c).rgbBlue = ToBlue
            End If
        Next c
    End Sub

    Public Sub MakeMask(ByRef bmInfo As BITMAPINFO, ByRef Mask As BITMAPINFO, ByVal TransparentColor As Integer)
        Dim c As Short
        'UPGRADE_WARNING: Couldn't resolve default property of object Mask. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
        Mask = bmInfo
        For c = 0 To 255
            ' Change entry matching TransparentColor to Black in bmInfo
            If RGB(bmInfo.bmiColors(c).rgbRed, bmInfo.bmiColors(c).rgbGreen, bmInfo.bmiColors(c).rgbBlue) = TransparentColor Then
                bmInfo.bmiColors(c).rgbRed = 0
                bmInfo.bmiColors(c).rgbGreen = 0
                bmInfo.bmiColors(c).rgbBlue = 0
            End If
            ' Create mask based on TransparentColor
            If RGB(Mask.bmiColors(c).rgbRed, Mask.bmiColors(c).rgbGreen, Mask.bmiColors(c).rgbBlue) = TransparentColor Then
                Mask.bmiColors(c).rgbRed = 255
                Mask.bmiColors(c).rgbGreen = 255
                Mask.bmiColors(c).rgbBlue = 255
            Else
                Mask.bmiColors(c).rgbRed = 0
                Mask.bmiColors(c).rgbGreen = 0
                Mask.bmiColors(c).rgbBlue = 0
            End If
        Next c
    End Sub

    Public Sub MakeFadeToWhite(ByRef bmInfo As BITMAPINFO, ByRef bmFadeToWhite() As BITMAPINFO)
        Dim c, i As Short
        For c = 0 To UBound(bmFadeToWhite)
            'UPGRADE_WARNING: Couldn't resolve default property of object bmFadeToWhite(c). Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
            bmFadeToWhite(c) = bmInfo
            For i = 0 To 255
                bmFadeToWhite(c).bmiColors(i).rgbRed = bmInfo.bmiColors(i).rgbRed + (255 - bmInfo.bmiColors(i).rgbRed) * (c / UBound(bmFadeToWhite))
                bmFadeToWhite(c).bmiColors(i).rgbGreen = bmInfo.bmiColors(i).rgbGreen + (255 - bmInfo.bmiColors(i).rgbGreen) * (c / UBound(bmFadeToWhite))
                bmFadeToWhite(c).bmiColors(i).rgbBlue = bmInfo.bmiColors(i).rgbBlue + (255 - bmInfo.bmiColors(i).rgbBlue) * (c / UBound(bmFadeToWhite))
            Next i
        Next c
    End Sub

    Public Sub MakeGray(ByRef bmInfo As BITMAPINFO, ByRef bmHalfTone As BITMAPINFO, ByVal Percent As Short)
        Dim c As Short
        Dim i As Byte
        'UPGRADE_WARNING: Couldn't resolve default property of object bmHalfTone. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
        bmHalfTone = bmInfo
        For c = 0 To 255
            i = CByte((CShort(bmInfo.bmiColors(c).rgbRed) + CShort(bmInfo.bmiColors(c).rgbGreen) + CShort(bmInfo.bmiColors(c).rgbBlue)) / 3)
            If i * (Percent / 100) < 255 Then
                i = i * (Percent / 100)
            Else
                i = 255
            End If
            bmHalfTone.bmiColors(c).rgbRed = i
            bmHalfTone.bmiColors(c).rgbGreen = i
            bmHalfTone.bmiColors(c).rgbBlue = i
        Next c
    End Sub

    Public Sub MakeDark(ByRef bmInfo As BITMAPINFO, ByRef bmDarkTone As BITMAPINFO, ByVal Percent As Short)
        Dim c As Short
        'UPGRADE_WARNING: Couldn't resolve default property of object bmDarkTone. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
        bmDarkTone = bmInfo
        For c = 0 To 255
            bmDarkTone.bmiColors(c).rgbRed = bmDarkTone.bmiColors(c).rgbRed * (100 - Percent) / 100
            bmDarkTone.bmiColors(c).rgbGreen = bmDarkTone.bmiColors(c).rgbGreen * (100 - Percent) / 100
            bmDarkTone.bmiColors(c).rgbBlue = bmDarkTone.bmiColors(c).rgbBlue * (100 - Percent) / 100
            If Percent > 30 And bmDarkTone.bmiColors(c).rgbBlue <> 0 Then
                If bmDarkTone.bmiColors(c).rgbBlue + 25 > 255 Then
                    bmDarkTone.bmiColors(c).rgbBlue = 255
                Else
                    bmDarkTone.bmiColors(c).rgbBlue = bmDarkTone.bmiColors(c).rgbBlue + 25
                End If
            End If
        Next c
    End Sub

    Public Sub MakeFadeToBlack(ByRef bmInfo As BITMAPINFO, ByRef bmFadeToBlack() As BITMAPINFO)
        Dim c, i As Short
        For c = UBound(bmFadeToBlack) To 0 Step -1
            'UPGRADE_WARNING: Couldn't resolve default property of object bmFadeToBlack(c). Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
            bmFadeToBlack(c) = bmInfo
            For i = 0 To 255
                bmFadeToBlack(c).bmiColors(i).rgbRed = bmInfo.bmiColors(i).rgbRed * (c / UBound(bmFadeToBlack))
                bmFadeToBlack(c).bmiColors(i).rgbGreen = bmInfo.bmiColors(i).rgbGreen * (c / UBound(bmFadeToBlack))
                bmFadeToBlack(c).bmiColors(i).rgbBlue = bmInfo.bmiColors(i).rgbBlue * (c / UBound(bmFadeToBlack))
            Next i
        Next c
    End Sub
End Module