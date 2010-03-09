Option Strict Off
Option Explicit On
Imports System.Collections.ArrayList
Friend Class Map
    ' Map is an Object which holds one set of MapLayers, zero or
    ' more Doors, zero or more Encounters, etc.
    'Oh god this is a singeleton ;(
    Private Shared instance As Map
    ' Version number of Class
    Private myVersion As Short
    ' myName is used for reference in Plot only (show in TreeView)
    Private myName As String
    ' myIndex is a unique number for this Map in the Plot
    Private myIndex As Short
    ' myComments are text comments used at Design Time only
    Private myComments As String

    ' PictureFile is the file name of the graphic. Looks in /tome first,
    ' then in /data.
    Private myPictureFile As String
    ' Pointer to PictureBox (used only at RUN time, not saved)
    Private myPic As Short

    ' ReGen: Style to use in making map
    ' 0 - Tetris Spin
    Private myDefaultStyle As Byte
    ' RegEn: Total number of Skill points to use in creating this map
    Private myExperiencePoints As Short
    ' Created for Party Size
    ' Bit1-2 - 0 = Solo, 1 = 2-3, 2 = 4-6, 3 = 7-9
    ' With Average Party Level
    ' Bit3-4 - 0 = 1-3, 1 = 4-6, 2 = 7-10, 3 = 10+
    ' Bit5  - Off=Normal TileSet, On=No Tile Set
    ' Bit6  - Off=Use Runes, On=Don't Use Runes
    ' Bit7-8   Not Used
    Private myStyle As Byte
    ' Flags:
    ' Bit1  - On=Generate upon Entry Off=Don't.
    ' Bit2  - On=Append to Existing Map, Off=Don't.
    ' Bit3-4  Not Used
    ' Bit5-8 (MapStyle)
    '       0 - Town Map
    '       1 - Wilderness Map
    '       2 - Building Map
    '       3 - Dungeon Map
    ' Bit9-16 Not Used
    Private myGenerateUponEntry As Short
    ' ReGen: Default Size of Map (Max: 256x256)
    Private myDefaultWidth As Short
    Private myDefaultHeight As Short
    ' ReGen: Default number of Encounters in Map
    Private myDefaultEncounters As Short

    Private myRunes(19) As Short

    ' myWidth and myHeight are in Tiles. Max is a reasonable number.
    Private myWidth As Short
    Private myHeight As Short
    ' myLeft and myTop are the last displayed Left and Top of the Map.
    Private myLeft As Short
    Private myTop As Short
    ' Layers myBottomTile and myTopTile point to tile in Tile graphic (bmp)
    Private myTopTile(,) As Short
    Private myMiddleTile(,) As Short
    Private myBottomTile(,) As Short
    ' Layer which points to an Encounter.Index. This means a single tile
    ' points to zero or one Encounter.
    Private myEncPointer(,) As Short
    ' myVisible is set of 8 bit flags per coordinate:
    '   &H1  Bit 0  - On = Is Hidden, Off = Is Not Hidden
    '   &H2  Bit 1  - Not Used
    '   &H4  Bit 2  - Not Used
    '   &H8  Bit 3  - On = Flip MiddleTile, Off = Don't flip
    '   &H10 Bit 4  - On = Flip BottomTile, Off = Don't flip
    '   &H20 Bit 5  - On = Flip TopTile, Off = Don't flip
    '   &H40 Bit 6  - Not Used
    '   &H80 Bit 7  - Not Used
    Private myVisible(,) As Byte

    ' List of EntryPoints (x,y) coordinates, Plot, Map and EntryPoint Ref)
    Private myEntryPoints As New Collection
    ' Collection of Tiles (picture, swap tile, triggers) for this Map
    Private myTiles As New Collection
    ' Set of TileSet names for the TileSet for this map. Just used as
    ' a collection point for tiles.
    Private myTileSets As New Collection
    ' myEncounters holds Encounters: Creatures, Items and Triggers.
    Private myEncounters As New Collection
    ' myThemes holds Encounters, PlotItems, Creatures, Items, Triggers and
    ' Descriptions for random placement
    Private myThemes As New Collection

    Public Shared Function GetInstance() As Map
        If Instance Is Nothing Then
            instance = New Map
        End If
        Return instance
    End Function
    Public ReadOnly Property TotalSize() As Integer
        Get
            TotalSize = myWidth * myHeight
        End Get
    End Property


    Public Property Pic() As Short
        Get
            Pic = myPic
        End Get
        Set(ByVal Value As Short)
            myPic = Value
        End Set
    End Property


    Public Property PictureFile() As String
        Get
            PictureFile = Trim(myPictureFile)
        End Get
        Set(ByVal Value As String)
            myPictureFile = Trim(Value)
        End Set
    End Property


    Public Property DefaultStyle() As Byte
        Get
            DefaultStyle = myDefaultStyle
        End Get
        Set(ByVal Value As Byte)
            myDefaultStyle = Value
        End Set
    End Property


    Public Property ExperiencePoints() As Short
        Get
            ExperiencePoints = myExperiencePoints
        End Get
        Set(ByVal Value As Short)
            myExperiencePoints = Value
        End Set
    End Property


    Public Property GenerateUponEntry() As Short
        Get
            GenerateUponEntry = (myGenerateUponEntry And &H1) > 0
        End Get
        Set(ByVal Value As Short)
            If Value Then
                myGenerateUponEntry = myGenerateUponEntry Or &H1
            Else
                myGenerateUponEntry = myGenerateUponEntry And (Not &H1)
            End If
        End Set
    End Property


    Public Property GenerateAppend() As Short
        Get
            GenerateAppend = (myGenerateUponEntry And &H2) > 0
        End Get
        Set(ByVal Value As Short)
            If Value Then
                myGenerateUponEntry = myGenerateUponEntry Or &H2
            Else
                myGenerateUponEntry = myGenerateUponEntry And (Not &H2)
            End If
        End Set
    End Property

    Public ReadOnly Property IsOutside() As Short
        Get
            Select Case Me.MapStyle
                Case 0 ' Town
                    IsOutside = True
                Case 1 ' Wilderness
                    IsOutside = True
                Case Else
                    IsOutside = False
            End Select
        End Get
    End Property


    Public Property MapStyle() As Short
        Get
            MapStyle = CShort(CShort(myGenerateUponEntry And &HF0) / 16)
        End Get
        Set(ByVal Value As Short)
            myGenerateUponEntry = myGenerateUponEntry And &HFF0F
            myGenerateUponEntry = myGenerateUponEntry Or (Value * 16 And &HF0)
        End Set
    End Property


    Public Property PartySize() As Short
        Get
            PartySize = CShort(myStyle And &H3)
        End Get
        Set(ByVal Value As Short)
            myStyle = myStyle And &HFFFC
            myStyle = myStyle Or (Value And &H3)
        End Set
    End Property


    Public Property PartyAvgLevel() As Short
        Get
            PartyAvgLevel = CShort(CShort(myStyle And &HC) / 4)
        End Get
        Set(ByVal Value As Short)
            myStyle = myStyle And &HFFF3
            myStyle = myStyle Or (Value * 4 And &HC)
        End Set
    End Property


    Public Property IsNoTiles() As Short
        Get
            IsNoTiles = (myStyle And &H20) > 0
        End Get
        Set(ByVal Value As Short)
            If Value Then
                myStyle = myStyle Or &H20
            Else
                myStyle = myStyle And (Not &H20)
            End If
        End Set
    End Property


    Public Property IsNoRunes() As Short
        Get
            IsNoRunes = (myStyle And &H40) > 0
        End Get
        Set(ByVal Value As Short)
            If Value Then
                myStyle = myStyle Or &H40
            Else
                myStyle = myStyle And (Not &H40)
            End If
        End Set
    End Property

    Public ReadOnly Property EncounterPoints() As Short
        Get
            Dim EncounterX As Encounter
            For Each EncounterX In myEncounters
                EncounterPoints = EncounterPoints + EncounterX.EncounterPoints
            Next EncounterX
        End Get
    End Property


    Public Property DefaultWidth() As Short
        Get
            DefaultWidth = myDefaultWidth
        End Get
        Set(ByVal Value As Short)
            myDefaultWidth = Value
        End Set
    End Property


    Public Property DefaultHeight() As Short
        Get
            DefaultHeight = myDefaultHeight
        End Get
        Set(ByVal Value As Short)
            myDefaultHeight = Value
        End Set
    End Property


    Public Property DefaultEncounters() As Short
        Get
            DefaultEncounters = myDefaultEncounters
        End Get
        Set(ByVal Value As Short)
            myDefaultEncounters = Value
        End Set
    End Property


    Public Property Runes(ByVal Index As Short) As Short
        Get
            Runes = myRunes(Index)
        End Get
        Set(ByVal Value As Short)
            myRunes(Index) = Value
        End Set
    End Property

    Public ReadOnly Property Encounters() As Collection
        Get
            Encounters = myEncounters
        End Get
    End Property


    Public Property EntryPoints() As Collection
        Get
            EntryPoints = myEntryPoints
        End Get
        Set(ByVal Value As Collection)
            myEntryPoints = Value
        End Set
    End Property

    Public ReadOnly Property Tiles() As Collection
        Get
            Tiles = myTiles
        End Get
    End Property


    Public Property Index() As Short
        Get
            Index = myIndex
        End Get
        Set(ByVal Value As Short)
            myIndex = Value
        End Set
    End Property

    Public ReadOnly Property TileSets() As Collection
        Get
            TileSets = myTileSets
        End Get
    End Property

    Public ReadOnly Property Themes() As Collection
        Get
            Themes = myThemes
        End Get
    End Property


    Public Property Width() As Short
        Get
            Width = myWidth
        End Get
        Set(ByVal Value As Short)
            Dim TempB(myWidth, myHeight) As Byte
            Dim TempM(myWidth, myHeight) As Byte
            Dim TempT(myWidth, myHeight) As Byte
            Dim TempV(myWidth, myHeight) As Byte
            Dim TempE(myWidth, myHeight) As Short
            Dim c, i As Short
            ' Only change width if truly different
            If Value <> myWidth Then
                If Value < 0 Then
                    Value = 0
                End If
                ' Preserve Arrays
                For c = 0 To myWidth : For i = 0 To myHeight
                        TempB(c, i) = myBottomTile(c, i)
                        TempM(c, i) = myMiddleTile(c, i)
                        TempT(c, i) = myTopTile(c, i)
                        TempV(c, i) = myVisible(c, i)
                        TempE(c, i) = myEncPointer(c, i)
                    Next i : Next c
                ' Resize
                ReDim myBottomTile(Value, myHeight)
                ReDim myMiddleTile(Value, myHeight)
                ReDim myTopTile(Value, myHeight)
                ReDim myVisible(Value, myHeight)
                ReDim myEncPointer(Value, myHeight)
                ' Copy back
                For c = 0 To Least(myWidth, Value) : For i = 0 To myHeight
                        myBottomTile(c, i) = TempB(c, i)
                        myMiddleTile(c, i) = TempM(c, i)
                        myTopTile(c, i) = TempT(c, i)
                        myVisible(c, i) = TempV(c, i)
                        myEncPointer(c, i) = TempE(c, i)
                    Next i : Next c
                myWidth = Value
            End If
        End Set
    End Property


    Public Property Height() As Short
        Get
            Height = myHeight
        End Get
        Set(ByVal Value As Short)
            Dim TempB(myWidth, myHeight) As Byte
            Dim TempM(myWidth, myHeight) As Byte
            Dim TempT(myWidth, myHeight) As Byte
            Dim TempV(myWidth, myHeight) As Byte
            Dim TempE(myWidth, myHeight) As Short
            Dim c, i As Short
            ' Only change height if truly different
            If Value <> myHeight Then
                If Value < 0 Then
                    Value = 0
                End If
                ' Preserve Arrays
                For c = 0 To myWidth : For i = 0 To myHeight
                        TempB(c, i) = myBottomTile(c, i)
                        TempM(c, i) = myMiddleTile(c, i)
                        TempT(c, i) = myTopTile(c, i)
                        TempV(c, i) = myVisible(c, i)
                        TempE(c, i) = myEncPointer(c, i)
                    Next i : Next c
                ' Resize
                ReDim myBottomTile(myWidth, Value)
                ReDim myMiddleTile(myWidth, Value)
                ReDim myTopTile(myWidth, Value)
                ReDim myVisible(myWidth, Value)
                ReDim myEncPointer(myWidth, Value)
                ' Copy back
                For c = 0 To myWidth : For i = 0 To Least(myHeight, Value)
                        myBottomTile(c, i) = TempB(c, i)
                        myMiddleTile(c, i) = TempM(c, i)
                        myTopTile(c, i) = TempT(c, i)
                        myVisible(c, i) = TempV(c, i)
                        myEncPointer(c, i) = TempE(c, i)
                    Next i : Next c
                myHeight = Value
            End If
        End Set
    End Property



    Public Property Left() As Short
        Get
            Left = myLeft
        End Get
        Set(ByVal Value As Short)
            myLeft = Value
        End Set
    End Property


    Public Property Top() As Short
        Get
            Top = myTop
        End Get
        Set(ByVal Value As Short)
            myTop = Value
        End Set
    End Property


    Public Property Hidden(ByVal X As Short, ByVal Y As Short) As Short
        Get
            If X < 0 Or Y < 0 Or X > myWidth Or Y > myHeight Then
                Hidden = False
            Else
                Hidden = (myVisible(X, Y) And &H1) > 0
            End If
        End Get
        Set(ByVal Value As Short)
            If X < 0 Or Y < 0 Or X > myWidth Or Y > myHeight Then
                Exit Property
            Else
                If Value Then
                    myVisible(X, Y) = myVisible(X, Y) Or &H1
                Else
                    myVisible(X, Y) = myVisible(X, Y) And (Not &H1)
                End If
            End If
        End Set
    End Property


    Public Property BottomFlip(ByVal X As Short, ByVal Y As Short) As Short
        Get
            ' TRUE if flipped at (x,y), FALSE if not
            BottomFlip = (myVisible(X, Y) And &H10) > 0
        End Get
        Set(ByVal Value As Short)
            If Value Then
                myVisible(X, Y) = myVisible(X, Y) Or &H10
            Else
                myVisible(X, Y) = myVisible(X, Y) And (Not &H10)
            End If
        End Set
    End Property


    Public Property TopFlip(ByVal X As Short, ByVal Y As Short) As Short
        Get
            ' TRUE if flipped at (x,y), FALSE if not
            TopFlip = (myVisible(X, Y) And &H20) > 0
        End Get
        Set(ByVal Value As Short)
            If Value Then
                myVisible(X, Y) = myVisible(X, Y) Or &H20
            Else
                myVisible(X, Y) = myVisible(X, Y) And (Not &H20)
            End If
        End Set
    End Property


    Public Property Tile(ByVal Index As Short, ByVal X As Short, ByVal Y As Short) As Short
        Get
            Select Case Index
                Case 0
                    Tile = myBottomTile(X, Y)
                Case 1
                    Tile = myMiddleTile(X, Y)
                Case 2
                    Tile = myTopTile(X, Y)
            End Select
        End Get
        Set(ByVal Value As Short)
            Select Case Index
                Case 0
                    myBottomTile(X, Y) = Value
                Case 1
                    myMiddleTile(X, Y) = Value
                Case 2
                    myTopTile(X, Y) = Value
            End Select
        End Set
    End Property


    Public Property BottomTile(ByVal X As Short, ByVal Y As Short) As Short
        Get
            BottomTile = myBottomTile(X, Y)
        End Get
        Set(ByVal Value As Short)
            myBottomTile(X, Y) = Value
        End Set
    End Property


    Public Property TopTile(ByVal X As Short, ByVal Y As Short) As Short
        Get
            TopTile = myTopTile(X, Y)
        End Get
        Set(ByVal Value As Short)
            myTopTile(X, Y) = Value
        End Set
    End Property


    Public Property MiddleTile(ByVal X As Short, ByVal Y As Short) As Short
        Get
            MiddleTile = myMiddleTile(X, Y)
        End Get
        Set(ByVal Value As Short)
            myMiddleTile(X, Y) = Value
        End Set
    End Property


    Public Property MiddleFlip(ByVal X As Short, ByVal Y As Short) As Short
        Get
            ' TRUE if flipped at (x,y), FALSE if not
            MiddleFlip = (myVisible(X, Y) And &H8) > 0
        End Get
        Set(ByVal Value As Short)
            If Value Then
                myVisible(X, Y) = myVisible(X, Y) Or &H8
            Else
                myVisible(X, Y) = myVisible(X, Y) And (Not &H8)
            End If
        End Set
    End Property


    Public Property EncPointer(ByVal X As Short, ByVal Y As Short) As Short
        Get
            EncPointer = myEncPointer(X, Y)
        End Get
        Set(ByVal Value As Short)
            myEncPointer(X, Y) = Value
        End Set
    End Property


    Public Property Visible(ByVal X As Short, ByVal Y As Short) As Short
        Get
            Visible = myVisible(X, Y)
        End Get
        Set(ByVal Value As Short)
            myVisible(X, Y) = Value
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


    Public Property Comments() As String
        Get
            Comments = Trim(myComments)
        End Get
        Set(ByVal Value As String)
            myComments = Trim(Value)
        End Set
    End Property

    Public Function AddEntryPoint() As EntryPoint
        ' Adds an EntryPoint and returns a reference to that EntryPoint
        Dim c As Short
        Dim EntryPointX As EntryPoint
        ' Find new available unused index identifier
        c = 1
        For Each EntryPointX In myEntryPoints
            If EntryPointX.Index >= c Then
                c = EntryPointX.Index + 1
            End If
        Next EntryPointX
        ' Set the index and add the entrypoint. Return the new entrypoint's index.
        EntryPointX = New EntryPoint
        EntryPointX.Index = c
        EntryPointX.Name = "EntryPoint" & c
        myEntryPoints.Add(EntryPointX, "P" & EntryPointX.Index)
        AddEntryPoint = EntryPointX
    End Function

    Public Function AddEntryPointAsIndex(ByRef Index As Short) As EntryPoint
        ' Adds an EntryPoint and returns a reference to that EntryPoint

        Dim EntryPointX As EntryPoint
        ' Set the index and add the entrypoint. Return the new entrypoint's index.
        EntryPointX = New EntryPoint
        EntryPointX.Index = Index
        EntryPointX.Name = "EntryPoint" & Index
        myEntryPoints.Add(EntryPointX, "P" & EntryPointX.Index)
        AddEntryPointAsIndex = EntryPointX
    End Function

    Public Sub RemoveEntryPoint(ByRef DeleteKey As String)
        ' Remove Encounter
        myEntryPoints.Remove(DeleteKey)
    End Sub

    Public Function AddEncounter() As Encounter
        ' Adds an Encounter and returns a reference to that Encounter
        Dim c As Short
        Dim EncounterX As Encounter
        ' Find new available unused index identifier
        c = 1
        For Each EncounterX In myEncounters
            If EncounterX.Index >= c Then
                c = EncounterX.Index + 1
            End If
        Next EncounterX
        ' Set the index and add the encounter. Return the new encounter's index.
        EncounterX = New Encounter
        EncounterX.Index = c
        EncounterX.Name = "Encounter" & c
        myEncounters.Add(EncounterX, "E" & EncounterX.Index)
        AddEncounter = EncounterX
    End Function

    Public Function AddEncounterAsIndex(ByRef Index As Short) As Encounter
        ' Adds an Encounter and returns a reference to that Encounter

        Dim EncounterX As Encounter
        ' Set the index and add the encounter. Return the new encounter's index.
        EncounterX = New Encounter
        EncounterX.Index = Index
        EncounterX.Name = "Encounter" & Index
        myEncounters.Add(EncounterX, "E" & EncounterX.Index)
        AddEncounterAsIndex = EncounterX
    End Function

    Public Function AddTheme() As Theme
        ' Adds an Theme and returns a reference to that Theme
        Dim c As Short
        Dim ThemeX As Theme
        ' Find new available unused index identifier
        c = 1
        For Each ThemeX In myThemes
            If ThemeX.Index >= c Then
                c = ThemeX.Index + 1
            End If
        Next ThemeX
        ' Set the index and add the Theme. Return the new Theme's index.
        ThemeX = New Theme
        ThemeX.Index = c
        ThemeX.Name = "Theme" & c
        myThemes.Add(ThemeX, "R" & ThemeX.Index)
        AddTheme = ThemeX
    End Function

    Public Function AddThemeAsIndex(ByRef Index As Short) As Theme
        ' Adds an Theme and returns a reference to that Theme

        Dim ThemeX As Theme
        ' Set the index and add the Theme. Return the new Theme's index.
        ThemeX = New Theme
        ThemeX.Index = Index
        ThemeX.Name = "Theme" & Index
        myThemes.Add(ThemeX, "R" & ThemeX.Index)
        AddThemeAsIndex = ThemeX
    End Function

    Public Function AddTile() As Tile
        ' Adds a Tile and returns a reference to that Tile
        Dim c As Short
        Dim TileX As Tile
        ' Find new available unused index identifier
        c = 1
        For Each TileX In myTiles
            If TileX.Index >= c Then
                c = TileX.Index + 1
            End If
        Next TileX
        ' Create the Tile and set the Index
        TileX = New Tile
        TileX.Index = c
        TileX.Name = "Tile" & c
        myTiles.Add(TileX, "L" & TileX.Index)
        AddTile = TileX
    End Function

    Public Function AddTileAsIndex(ByRef Index As Short) As Tile
        ' Adds a Tile and returns a reference to that Tile

        Dim TileX As Tile
        ' Create the Tile and set the Index
        TileX = New Tile
        TileX.Index = Index
        TileX.Name = "Tile" & Index
        myTiles.Add(TileX, "L" & TileX.Index)
        AddTileAsIndex = TileX
    End Function


    Public Sub RemoveEncounter(ByRef DeleteKey As String)
        Dim X, c, Y As Short
        ' Remove all references to this Encounter
        'UPGRADE_WARNING: Couldn't resolve default property of object myEncounters().Index. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
        c = myEncounters.Item(DeleteKey).Index
        For X = 0 To myWidth : For Y = 0 To myHeight
                If myEncPointer(X, Y) = c Then
                    myEncPointer(X, Y) = 0
                End If
            Next Y : Next X
        ' Remove Encounter
        myEncounters.Remove(DeleteKey)
    End Sub

    Public Sub RemoveTile(ByRef DeleteKey As String)
        Dim X, Index, Y As Short
        Dim TileX As Tile
        ' Remove references to this tile
        'UPGRADE_WARNING: Couldn't resolve default property of object myTiles().Index. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
        Index = myTiles.Item(DeleteKey).Index
        For X = 0 To myWidth
            For Y = 0 To myHeight
                If myTopTile(X, Y) = Index Then
                    myTopTile(X, Y) = 0
                End If
                If myMiddleTile(X, Y) = Index Then
                    myMiddleTile(X, Y) = 0
                End If
                If myBottomTile(X, Y) = Index Then
                    myBottomTile(X, Y) = 0
                End If
            Next Y
        Next X
        For Each TileX In myTiles
            If TileX.SwapTile = Index Then
                TileX.SwapTile = 0
            End If
        Next TileX
        ' Remove the tile
        myTiles.Remove(DeleteKey)
    End Sub

    Public Sub RemoveTheme(ByRef DeleteKey As String)
        myThemes.Remove(DeleteKey)
    End Sub

    Public Function AddTileSet() As TileSet
        ' Adds a TileSet and returns a reference to that TileSet
        Dim c As Short
        Dim TileSetX As TileSet
        ' Find new available unused index identifier
        c = 1
        For Each TileSetX In myTileSets
            If TileSetX.Index <> c Then
                Exit For
            End If
            c = c + 1
        Next TileSetX
        ' Create the TileSet and set the Index
        TileSetX = New TileSet
        TileSetX.Index = c
        TileSetX.Name = "TileSet" & c
        myTileSets.Add(TileSetX, "W" & TileSetX.Index)
        AddTileSet = TileSetX
    End Function

    Public Sub RemoveTileSet(ByRef DeleteKey As String)
        Dim c As Short
        Dim TileX As Tile
        ' Remove pointer from all Tiles
        'UPGRADE_WARNING: Couldn't resolve default property of object myTileSets().Index. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
        c = myTileSets.Item(DeleteKey).Index
        For Each TileX In myTiles
            If TileX.TileSet = c Then
                TileX.TileSet = 0
            End If
        Next TileX
        ' Remove TileSet
        myTileSets.Remove(DeleteKey)
    End Sub

    Public Sub ReadFromFileHeader(ByRef FromFile As Short)
        Dim c As Short
        Dim Cnt As Integer
        ' Read Map Name and descriptors
        'UPGRADE_WARNING: Get was upgraded to FileGet and has a new behavior. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="9B7D5ADD-D8FE-4819-A36C-6DEDAF088CC7"'
        FileGet(FromFile, myVersion)
        'UPGRADE_WARNING: Get was upgraded to FileGet and has a new behavior. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="9B7D5ADD-D8FE-4819-A36C-6DEDAF088CC7"'
        FileGet(FromFile, Cnt)
        myName = ""
        For c = 1 To Cnt
            myName = myName & " "
        Next c
        'UPGRADE_WARNING: Get was upgraded to FileGet and has a new behavior. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="9B7D5ADD-D8FE-4819-A36C-6DEDAF088CC7"'
        FileGet(FromFile, myName)
        'UPGRADE_WARNING: Get was upgraded to FileGet and has a new behavior. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="9B7D5ADD-D8FE-4819-A36C-6DEDAF088CC7"'
        FileGet(FromFile, myIndex)
    End Sub

    Public Sub ReadFromFile(ByRef FromFile As Short)
        On Error GoTo ErrorHandler
        Dim X, c, Y As Short
        Dim Cnt As Integer
        Dim TileX As Tile
        Dim EncounterX As Encounter
        Dim TileSetX As TileSet
        Dim EntryPointX As EntryPoint
        Dim ThemeX As Theme
        Dim Reading As String
        ' Read File Header
        ReadFromFileHeader(FromFile)
        ' Read other parameters
        'UPGRADE_WARNING: Get was upgraded to FileGet and has a new behavior. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="9B7D5ADD-D8FE-4819-A36C-6DEDAF088CC7"'
        FileGet(FromFile, Cnt)
        myComments = ""
        For c = 1 To Cnt
            myComments = myComments & " "
        Next c
        'UPGRADE_WARNING: Get was upgraded to FileGet and has a new behavior. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="9B7D5ADD-D8FE-4819-A36C-6DEDAF088CC7"'
        FileGet(FromFile, myComments)
        ' Read PictureFile for Map
        'UPGRADE_WARNING: Get was upgraded to FileGet and has a new behavior. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="9B7D5ADD-D8FE-4819-A36C-6DEDAF088CC7"'
        FileGet(FromFile, Cnt)
        myPictureFile = ""
        For c = 1 To Cnt
            myPictureFile = myPictureFile & " "
        Next c
        'UPGRADE_WARNING: Get was upgraded to FileGet and has a new behavior. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="9B7D5ADD-D8FE-4819-A36C-6DEDAF088CC7"'
        FileGet(FromFile, myPictureFile)
        ' Read Map Layers
        'UPGRADE_WARNING: Get was upgraded to FileGet and has a new behavior. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="9B7D5ADD-D8FE-4819-A36C-6DEDAF088CC7"'
        FileGet(FromFile, myWidth)
        'UPGRADE_WARNING: Get was upgraded to FileGet and has a new behavior. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="9B7D5ADD-D8FE-4819-A36C-6DEDAF088CC7"'
        FileGet(FromFile, myHeight)
        'UPGRADE_WARNING: Get was upgraded to FileGet and has a new behavior. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="9B7D5ADD-D8FE-4819-A36C-6DEDAF088CC7"'
        FileGet(FromFile, myLeft)
        'UPGRADE_WARNING: Get was upgraded to FileGet and has a new behavior. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="9B7D5ADD-D8FE-4819-A36C-6DEDAF088CC7"'
        FileGet(FromFile, myTop)
        ReDim myBottomTile(myWidth, myHeight)
        ReDim myMiddleTile(myWidth, myHeight)
        ReDim myTopTile(myWidth, myHeight)
        ReDim myVisible(myWidth, myHeight)
        ReDim myEncPointer(myWidth, myHeight)
        For X = 0 To myWidth : For Y = 0 To myHeight
                'UPGRADE_WARNING: Get was upgraded to FileGet and has a new behavior. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="9B7D5ADD-D8FE-4819-A36C-6DEDAF088CC7"'
                FileGet(FromFile, myBottomTile(X, Y))
                'UPGRADE_WARNING: Get was upgraded to FileGet and has a new behavior. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="9B7D5ADD-D8FE-4819-A36C-6DEDAF088CC7"'
                FileGet(FromFile, myMiddleTile(X, Y))
                'UPGRADE_WARNING: Get was upgraded to FileGet and has a new behavior. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="9B7D5ADD-D8FE-4819-A36C-6DEDAF088CC7"'
                FileGet(FromFile, myTopTile(X, Y))
                'UPGRADE_WARNING: Get was upgraded to FileGet and has a new behavior. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="9B7D5ADD-D8FE-4819-A36C-6DEDAF088CC7"'
                FileGet(FromFile, myVisible(X, Y))
                'UPGRADE_WARNING: Get was upgraded to FileGet and has a new behavior. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="9B7D5ADD-D8FE-4819-A36C-6DEDAF088CC7"'
                FileGet(FromFile, myEncPointer(X, Y))
            Next Y : Next X
        ' Read ReGen Parameters
        'UPGRADE_WARNING: Get was upgraded to FileGet and has a new behavior. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="9B7D5ADD-D8FE-4819-A36C-6DEDAF088CC7"'
        FileGet(FromFile, myDefaultStyle)
        'UPGRADE_WARNING: Get was upgraded to FileGet and has a new behavior. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="9B7D5ADD-D8FE-4819-A36C-6DEDAF088CC7"'
        FileGet(FromFile, myExperiencePoints)
        'UPGRADE_WARNING: Get was upgraded to FileGet and has a new behavior. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="9B7D5ADD-D8FE-4819-A36C-6DEDAF088CC7"'
        FileGet(FromFile, myStyle)
        'UPGRADE_WARNING: Get was upgraded to FileGet and has a new behavior. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="9B7D5ADD-D8FE-4819-A36C-6DEDAF088CC7"'
        FileGet(FromFile, myGenerateUponEntry)
        'UPGRADE_WARNING: Get was upgraded to FileGet and has a new behavior. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="9B7D5ADD-D8FE-4819-A36C-6DEDAF088CC7"'
        FileGet(FromFile, myDefaultWidth)
        'UPGRADE_WARNING: Get was upgraded to FileGet and has a new behavior. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="9B7D5ADD-D8FE-4819-A36C-6DEDAF088CC7"'
        FileGet(FromFile, myDefaultHeight)
        'UPGRADE_WARNING: Get was upgraded to FileGet and has a new behavior. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="9B7D5ADD-D8FE-4819-A36C-6DEDAF088CC7"'
        FileGet(FromFile, myDefaultEncounters)
        ' Read Runes
        For c = 0 To 19
            'UPGRADE_WARNING: Get was upgraded to FileGet and has a new behavior. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="9B7D5ADD-D8FE-4819-A36C-6DEDAF088CC7"'
            FileGet(FromFile, myRunes(c))
        Next c
        ' Read EntryPoints
        Reading = "EntryPoint"
        'UPGRADE_WARNING: Get was upgraded to FileGet and has a new behavior. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="9B7D5ADD-D8FE-4819-A36C-6DEDAF088CC7"'
        FileGet(FromFile, Cnt)
        For c = 1 To Cnt
            EntryPointX = New EntryPoint
            EntryPointX.ReadFromFile(FromFile)
            myEntryPoints.Add(EntryPointX, "P" & EntryPointX.Index)
        Next c
        ' Read TileSets
        Reading = "TileSet"
        'UPGRADE_WARNING: Get was upgraded to FileGet and has a new behavior. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="9B7D5ADD-D8FE-4819-A36C-6DEDAF088CC7"'
        FileGet(FromFile, Cnt)
        For c = 1 To Cnt
            TileSetX = New TileSet
            TileSetX.ReadFromFile(FromFile)
            myTileSets.Add(TileSetX, "W" & TileSetX.Index)
        Next c
        ' Read Tiles
        Reading = "Tile"
        'UPGRADE_WARNING: Get was upgraded to FileGet and has a new behavior. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="9B7D5ADD-D8FE-4819-A36C-6DEDAF088CC7"'
        FileGet(FromFile, Cnt)
        For c = 1 To Cnt
            TileX = New Tile
            TileX.ReadFromFile(FromFile)
            myTiles.Add(TileX, "L" & TileX.Index)
        Next c
        ' Read Themes
        Reading = "Theme"
        'UPGRADE_WARNING: Get was upgraded to FileGet and has a new behavior. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="9B7D5ADD-D8FE-4819-A36C-6DEDAF088CC7"'
        FileGet(FromFile, Cnt)
        For c = 1 To Cnt
            ThemeX = New Theme
            ThemeX.ReadFromFile(FromFile)
            myThemes.Add(ThemeX, "R" & ThemeX.Index)
        Next c
        ' Read Encounters
        Reading = "Encounter"
        'UPGRADE_WARNING: Get was upgraded to FileGet and has a new behavior. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="9B7D5ADD-D8FE-4819-A36C-6DEDAF088CC7"'
        FileGet(FromFile, Cnt)
        For c = 1 To Cnt
            EncounterX = New Encounter
            EncounterX.ReadFromFile(FromFile)
            myEncounters.Add(EncounterX, "E" & EncounterX.Index)
        Next
        Exit Sub
ErrorHandler:
        ' [Titi 2.4.9]
        If Err.Number = 457 And Reading <> "" Then
            Select Case Reading
                Case "EntryPoint"
                    oErr.logError(Reading & " index conflict: " & EntryPointX.Index)
                Case "TileSet"
                    oErr.logError(Reading & " index conflict: " & TileSetX.Index)
                Case "Tile"
                    oErr.logError(Reading & " index conflict: " & TileX.Index)
                Case "Theme"
                    oErr.logError(Reading & " index conflict: " & ThemeX.Index)
                Case "Encounter"
                    oErr.logError(Reading & " index conflict: " & EncounterX.Index)
            End Select
        Else
            oErr.logError("Cannot read map" & IIf(Reading <> vbNullString, "'s " & Reading, "") & ", error#" & Err.Number & " (" & Err.Description & ")")
        End If
    End Sub

    Public Sub Copy(ByRef FromMap As Map)
        Dim X, c, Y As Short
        Dim TileX As Tile
        Dim EncounterX As Encounter
        Dim TileSetX As TileSet
        Dim ThemeX As Theme
        Dim EntryPointX As EntryPoint
        ' Copy Map Name and descriptors
        myName = FromMap.Name
        myComments = FromMap.Comments
        ' Copy PictureFile for Map
        myPictureFile = FromMap.PictureFile
        ' Copy Map Layers
        Me.Width = FromMap.Width
        Me.Height = FromMap.Height
        myLeft = FromMap.Left
        myTop = FromMap.Top
        For X = 0 To myWidth : For Y = 0 To myHeight
                myBottomTile(X, Y) = FromMap.BottomTile(X, Y)
                myMiddleTile(X, Y) = FromMap.MiddleTile(X, Y)
                myTopTile(X, Y) = FromMap.TopTile(X, Y)
                myVisible(X, Y) = FromMap.Visible(X, Y)
                myEncPointer(X, Y) = FromMap.EncPointer(X, Y)
            Next Y : Next X
        ' Copy ReGen Parameters
        myDefaultStyle = FromMap.DefaultStyle
        myExperiencePoints = FromMap.ExperiencePoints
        Me.GenerateUponEntry = FromMap.GenerateUponEntry
        Me.MapStyle = FromMap.MapStyle
        myDefaultWidth = FromMap.DefaultWidth
        myDefaultHeight = FromMap.DefaultHeight
        myDefaultEncounters = FromMap.DefaultEncounters
        ' Copy Runes
        For c = 0 To 19
            Me.Runes(c) = FromMap.Runes(c)
        Next c
        ' Copy TileSets
        myTileSets = New Collection
        For c = 1 To FromMap.TileSets.Count()
            TileSetX = Me.AddTileSet
            'UPGRADE_WARNING: Couldn't resolve default property of object FromMap.TileSets(). Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
            TileSetX.Copy(FromMap.TileSets.Item(c))
        Next c
        ' Copy Tiles
        myTiles = New Collection
        For c = 1 To FromMap.Tiles.Count()
            'UPGRADE_WARNING: Couldn't resolve default property of object FromMap.Tiles().Index. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
            TileX = Me.AddTileAsIndex(FromMap.Tiles.Item(c).Index)
            'UPGRADE_WARNING: Couldn't resolve default property of object FromMap.Tiles(). Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
            TileX.Copy(FromMap.Tiles.Item(c))
        Next c
        ' Copy Themes
        myThemes = New Collection
        For c = 1 To FromMap.Themes.Count()
            'UPGRADE_WARNING: Couldn't resolve default property of object FromMap.Themes().Index. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
            ThemeX = Me.AddThemeAsIndex(FromMap.Themes.Item(c).Index)
            'UPGRADE_WARNING: Couldn't resolve default property of object FromMap.Themes(). Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
            ThemeX.Copy(FromMap.Themes.Item(c))
        Next c
        ' Copy Encounters
        myEncounters = New Collection
        For c = 1 To FromMap.Encounters.Count()
            'UPGRADE_WARNING: Couldn't resolve default property of object FromMap.Encounters().Index. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
            EncounterX = Me.AddEncounterAsIndex(FromMap.Encounters.Item(c).Index)
            'UPGRADE_WARNING: Couldn't resolve default property of object FromMap.Encounters(). Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
            EncounterX.Copy(FromMap.Encounters.Item(c))
        Next c
        ' Copy EntryPoints
        myEntryPoints = New Collection
        For c = 1 To FromMap.EntryPoints.Count()
            'UPGRADE_WARNING: Couldn't resolve default property of object FromMap.EntryPoints().Index. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
            EntryPointX = Me.AddEntryPointAsIndex(FromMap.EntryPoints.Item(c).Index)
            'UPGRADE_WARNING: Couldn't resolve default property of object FromMap.EntryPoints(). Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
            EntryPointX.Copy(FromMap.EntryPoints.Item(c))
        Next c
    End Sub

    Public Sub SaveToFile(ByRef ToFile As Short)
        Dim X, c, Y As Short
        ' Save Map Name and descriptors
        'UPGRADE_WARNING: Put was upgraded to FilePut and has a new behavior. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="9B7D5ADD-D8FE-4819-A36C-6DEDAF088CC7"'
        FilePut(ToFile, myVersion)
        'UPGRADE_WARNING: Put was upgraded to FilePut and has a new behavior. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="9B7D5ADD-D8FE-4819-A36C-6DEDAF088CC7"'
        FilePut(ToFile, Len(myName))
        'UPGRADE_WARNING: Put was upgraded to FilePut and has a new behavior. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="9B7D5ADD-D8FE-4819-A36C-6DEDAF088CC7"'
        FilePut(ToFile, myName)
        'UPGRADE_WARNING: Put was upgraded to FilePut and has a new behavior. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="9B7D5ADD-D8FE-4819-A36C-6DEDAF088CC7"'
        FilePut(ToFile, myIndex)
        'UPGRADE_WARNING: Put was upgraded to FilePut and has a new behavior. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="9B7D5ADD-D8FE-4819-A36C-6DEDAF088CC7"'
        FilePut(ToFile, Len(myComments))
        'UPGRADE_WARNING: Put was upgraded to FilePut and has a new behavior. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="9B7D5ADD-D8FE-4819-A36C-6DEDAF088CC7"'
        FilePut(ToFile, myComments)
        ' Save PictureFile for Map
        'UPGRADE_WARNING: Put was upgraded to FilePut and has a new behavior. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="9B7D5ADD-D8FE-4819-A36C-6DEDAF088CC7"'
        FilePut(ToFile, Len(myPictureFile))
        'UPGRADE_WARNING: Put was upgraded to FilePut and has a new behavior. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="9B7D5ADD-D8FE-4819-A36C-6DEDAF088CC7"'
        FilePut(ToFile, myPictureFile)
        ' Save Map Layers
        'UPGRADE_WARNING: Put was upgraded to FilePut and has a new behavior. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="9B7D5ADD-D8FE-4819-A36C-6DEDAF088CC7"'
        FilePut(ToFile, myWidth)
        'UPGRADE_WARNING: Put was upgraded to FilePut and has a new behavior. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="9B7D5ADD-D8FE-4819-A36C-6DEDAF088CC7"'
        FilePut(ToFile, myHeight)
        'UPGRADE_WARNING: Put was upgraded to FilePut and has a new behavior. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="9B7D5ADD-D8FE-4819-A36C-6DEDAF088CC7"'
        FilePut(ToFile, myLeft)
        'UPGRADE_WARNING: Put was upgraded to FilePut and has a new behavior. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="9B7D5ADD-D8FE-4819-A36C-6DEDAF088CC7"'
        FilePut(ToFile, myTop)
        For X = 0 To myWidth : For Y = 0 To myHeight
                'UPGRADE_WARNING: Put was upgraded to FilePut and has a new behavior. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="9B7D5ADD-D8FE-4819-A36C-6DEDAF088CC7"'
                FilePut(ToFile, myBottomTile(X, Y))
                'UPGRADE_WARNING: Put was upgraded to FilePut and has a new behavior. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="9B7D5ADD-D8FE-4819-A36C-6DEDAF088CC7"'
                FilePut(ToFile, myMiddleTile(X, Y))
                'UPGRADE_WARNING: Put was upgraded to FilePut and has a new behavior. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="9B7D5ADD-D8FE-4819-A36C-6DEDAF088CC7"'
                FilePut(ToFile, myTopTile(X, Y))
                'UPGRADE_WARNING: Put was upgraded to FilePut and has a new behavior. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="9B7D5ADD-D8FE-4819-A36C-6DEDAF088CC7"'
                FilePut(ToFile, myVisible(X, Y))
                'UPGRADE_WARNING: Put was upgraded to FilePut and has a new behavior. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="9B7D5ADD-D8FE-4819-A36C-6DEDAF088CC7"'
                FilePut(ToFile, myEncPointer(X, Y))
            Next Y : Next X
        ' Save ReGen Parameters
        'UPGRADE_WARNING: Put was upgraded to FilePut and has a new behavior. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="9B7D5ADD-D8FE-4819-A36C-6DEDAF088CC7"'
        FilePut(ToFile, myDefaultStyle)
        'UPGRADE_WARNING: Put was upgraded to FilePut and has a new behavior. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="9B7D5ADD-D8FE-4819-A36C-6DEDAF088CC7"'
        FilePut(ToFile, myExperiencePoints)
        'UPGRADE_WARNING: Put was upgraded to FilePut and has a new behavior. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="9B7D5ADD-D8FE-4819-A36C-6DEDAF088CC7"'
        FilePut(ToFile, myStyle)
        'UPGRADE_WARNING: Put was upgraded to FilePut and has a new behavior. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="9B7D5ADD-D8FE-4819-A36C-6DEDAF088CC7"'
        FilePut(ToFile, myGenerateUponEntry)
        'UPGRADE_WARNING: Put was upgraded to FilePut and has a new behavior. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="9B7D5ADD-D8FE-4819-A36C-6DEDAF088CC7"'
        FilePut(ToFile, myDefaultWidth)
        'UPGRADE_WARNING: Put was upgraded to FilePut and has a new behavior. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="9B7D5ADD-D8FE-4819-A36C-6DEDAF088CC7"'
        FilePut(ToFile, myDefaultHeight)
        'UPGRADE_WARNING: Put was upgraded to FilePut and has a new behavior. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="9B7D5ADD-D8FE-4819-A36C-6DEDAF088CC7"'
        FilePut(ToFile, myDefaultEncounters)
        ' Save Runes
        For c = 0 To 19
            'UPGRADE_WARNING: Put was upgraded to FilePut and has a new behavior. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="9B7D5ADD-D8FE-4819-A36C-6DEDAF088CC7"'
            FilePut(ToFile, myRunes(c))
        Next c
        ' Save EntryPoints
        'UPGRADE_WARNING: Put was upgraded to FilePut and has a new behavior. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="9B7D5ADD-D8FE-4819-A36C-6DEDAF088CC7"'
        FilePut(ToFile, myEntryPoints.Count())
        For c = 1 To myEntryPoints.Count()
            'UPGRADE_WARNING: Couldn't resolve default property of object myEntryPoints().SaveToFile. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
            myEntryPoints.Item(c).SaveToFile(ToFile)
        Next c
        ' Save TileSets
        'UPGRADE_WARNING: Put was upgraded to FilePut and has a new behavior. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="9B7D5ADD-D8FE-4819-A36C-6DEDAF088CC7"'
        FilePut(ToFile, myTileSets.Count())
        For c = 1 To myTileSets.Count()
            'UPGRADE_WARNING: Couldn't resolve default property of object myTileSets().SaveToFile. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
            myTileSets.Item(c).SaveToFile(ToFile)
        Next c
        ' Save Tiles
        'UPGRADE_WARNING: Put was upgraded to FilePut and has a new behavior. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="9B7D5ADD-D8FE-4819-A36C-6DEDAF088CC7"'
        FilePut(ToFile, myTiles.Count())
        For c = 1 To myTiles.Count()
            'UPGRADE_WARNING: Couldn't resolve default property of object myTiles().SaveToFile. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
            myTiles.Item(c).SaveToFile(ToFile)
        Next c
        ' Save Themes
        'UPGRADE_WARNING: Put was upgraded to FilePut and has a new behavior. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="9B7D5ADD-D8FE-4819-A36C-6DEDAF088CC7"'
        FilePut(ToFile, myThemes.Count())
        For c = 1 To myThemes.Count()
            'UPGRADE_WARNING: Couldn't resolve default property of object myThemes().SaveToFile. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
            myThemes.Item(c).SaveToFile(ToFile)
        Next
        ' Save Encounters
        'UPGRADE_WARNING: Put was upgraded to FilePut and has a new behavior. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="9B7D5ADD-D8FE-4819-A36C-6DEDAF088CC7"'
        FilePut(ToFile, myEncounters.Count())
        For c = 1 To myEncounters.Count()
            'UPGRADE_WARNING: Couldn't resolve default property of object myEncounters().SaveToFile. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
            myEncounters.Item(c).SaveToFile(ToFile)
        Next
    End Sub

    Public Function MovementCost(ByVal X As Short, ByVal Y As Short) As Short
        Dim c, i As Short
        ' Calculates the average cost in 1/10 turn units of moving through a given tile
        ' Ranges from 1 to 10000.
        If X < 0 Or X > myWidth Or Y < 0 Or Y > myHeight Then
            c = 0
        Else
            c = 0 : i = 0
            If myTopTile(X, Y) > 0 Then
                'UPGRADE_WARNING: Couldn't resolve default property of object myTiles().MovementCostTurns. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
                c = c + myTiles.Item("L" & myTopTile(X, Y)).MovementCostTurns
                i = i + 1
            End If
            If myMiddleTile(X, Y) > 0 Then
                'UPGRADE_WARNING: Couldn't resolve default property of object myTiles().MovementCostTurns. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
                c = c + myTiles.Item("L" & myMiddleTile(X, Y)).MovementCostTurns
                i = i + 1
            End If
            If myBottomTile(X, Y) > 0 Then
                'UPGRADE_WARNING: Couldn't resolve default property of object myTiles().MovementCostTurns. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
                c = c + myTiles.Item("L" & myBottomTile(X, Y)).MovementCostTurns
                i = i + 1
            End If
            If i > 0 Then
                c = Int(c / i)
            End If
        End If
        MovementCost = c
    End Function

    Public Function Blocked(ByVal X As Short, ByVal Y As Short, ByVal Side As Object) As Short
        Dim c As Short
        ' Tile is blocked by any in Top, Middle or Bottom on either side
        c = 0
        If myTopTile(X, Y) > 0 Then
            'UPGRADE_WARNING: Couldn't resolve default property of object Side. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
            'UPGRADE_WARNING: Mod has a new behavior. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="9B7D5ADD-D8FE-4819-A36C-6DEDAF088CC7"'
            'UPGRADE_WARNING: Couldn't resolve default property of object myTiles().Blocked. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
            c = c + myTiles.Item("L" & myTopTile(X, Y)).Blocked(Side - (1 - 2 * (Side Mod 2)) * Me.TopFlip(X, Y))
        End If
        If myMiddleTile(X, Y) > 0 Then
            'UPGRADE_WARNING: Couldn't resolve default property of object Side. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
            'UPGRADE_WARNING: Mod has a new behavior. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="9B7D5ADD-D8FE-4819-A36C-6DEDAF088CC7"'
            'UPGRADE_WARNING: Couldn't resolve default property of object myTiles().Blocked. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
            c = c + myTiles.Item("L" & myMiddleTile(X, Y)).Blocked(Side - (1 - 2 * (Side Mod 2)) * Me.MiddleFlip(X, Y))
        End If
        If myBottomTile(X, Y) > 0 Then
            'UPGRADE_WARNING: Couldn't resolve default property of object Side. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
            'UPGRADE_WARNING: Mod has a new behavior. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="9B7D5ADD-D8FE-4819-A36C-6DEDAF088CC7"'
            'UPGRADE_WARNING: Couldn't resolve default property of object myTiles().Blocked. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
            c = c + myTiles.Item("L" & myBottomTile(X, Y)).Blocked(Side - (1 - 2 * (Side Mod 2)) * Me.BottomFlip(X, Y))
        End If
        ' Check surrounding tile
        Select Case Side
            Case 0 ' Top Left
                Y = Y - 1
                'UPGRADE_WARNING: Couldn't resolve default property of object Side. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
                Side = 3
            Case 1 ' Top Right
                X = X + 1
                'UPGRADE_WARNING: Couldn't resolve default property of object Side. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
                Side = 2
            Case 2 ' Bottom Left
                X = X - 1
                'UPGRADE_WARNING: Couldn't resolve default property of object Side. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
                Side = 1
            Case 3 ' Bottom Right
                Y = Y + 1
                'UPGRADE_WARNING: Couldn't resolve default property of object Side. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
                Side = 0
            Case 4 ' Up
                Y = Y - 1 : X = X + 1
                'UPGRADE_WARNING: Couldn't resolve default property of object Side. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
                Side = 7
            Case 5 ' Left
                Y = Y - 1 : X = X - 1
                'UPGRADE_WARNING: Couldn't resolve default property of object Side. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
                Side = 6
            Case 6 ' Right
                Y = Y + 1 : X = X + 1
                'UPGRADE_WARNING: Couldn't resolve default property of object Side. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
                Side = 5
            Case 7 ' Down
                Y = Y + 1 : X = X - 1
                'UPGRADE_WARNING: Couldn't resolve default property of object Side. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
                Side = 4
        End Select
        If X < 0 Or X > myWidth Or Y < 0 Or Y > myHeight Then
            c = -1
        ElseIf myTopTile(X, Y) + myMiddleTile(X, Y) + myBottomTile(X, Y) = 0 Then
            c = -1
        Else
            If myTopTile(X, Y) > 0 Then
                'UPGRADE_WARNING: Couldn't resolve default property of object Side. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
                'UPGRADE_WARNING: Mod has a new behavior. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="9B7D5ADD-D8FE-4819-A36C-6DEDAF088CC7"'
                'UPGRADE_WARNING: Couldn't resolve default property of object myTiles().Blocked. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
                c = c + myTiles.Item("L" & myTopTile(X, Y)).Blocked(Side - (1 - 2 * (Side Mod 2)) * Me.TopFlip(X, Y))
            End If
            If myMiddleTile(X, Y) > 0 Then
                'UPGRADE_WARNING: Couldn't resolve default property of object Side. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
                'UPGRADE_WARNING: Mod has a new behavior. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="9B7D5ADD-D8FE-4819-A36C-6DEDAF088CC7"'
                'UPGRADE_WARNING: Couldn't resolve default property of object myTiles().Blocked. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
                c = c + myTiles.Item("L" & myMiddleTile(X, Y)).Blocked(Side - (1 - 2 * (Side Mod 2)) * Me.MiddleFlip(X, Y))
            End If
            If myBottomTile(X, Y) > 0 Then
                'UPGRADE_WARNING: Couldn't resolve default property of object Side. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
                'UPGRADE_WARNING: Mod has a new behavior. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="9B7D5ADD-D8FE-4819-A36C-6DEDAF088CC7"'
                'UPGRADE_WARNING: Couldn't resolve default property of object myTiles().Blocked. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
                c = c + myTiles.Item("L" & myBottomTile(X, Y)).Blocked(Side - (1 - 2 * (Side Mod 2)) * Me.BottomFlip(X, Y))
            End If
        End If
        Blocked = (c < 0)
    End Function

    Public Function See(ByVal X As Short, ByVal Y As Short, ByVal Side As Object) As Short
        Dim c As Short
        ' Tile is See by any in Top, Middle or Bottom on either side
        c = 0
        If myTopTile(X, Y) > 0 Then
            'UPGRADE_WARNING: Couldn't resolve default property of object Side. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
            'UPGRADE_WARNING: Mod has a new behavior. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="9B7D5ADD-D8FE-4819-A36C-6DEDAF088CC7"'
            'UPGRADE_WARNING: Couldn't resolve default property of object myTiles().See. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
            c = c + myTiles.Item("L" & myTopTile(X, Y)).See(Side - (1 - 2 * (Side Mod 2)) * Me.TopFlip(X, Y))
        End If
        If myMiddleTile(X, Y) > 0 Then
            'UPGRADE_WARNING: Couldn't resolve default property of object Side. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
            'UPGRADE_WARNING: Mod has a new behavior. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="9B7D5ADD-D8FE-4819-A36C-6DEDAF088CC7"'
            'UPGRADE_WARNING: Couldn't resolve default property of object myTiles().See. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
            c = c + myTiles.Item("L" & myMiddleTile(X, Y)).See(Side - (1 - 2 * (Side Mod 2)) * Me.MiddleFlip(X, Y))
        End If
        If myBottomTile(X, Y) > 0 Then
            'UPGRADE_WARNING: Couldn't resolve default property of object Side. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
            'UPGRADE_WARNING: Mod has a new behavior. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="9B7D5ADD-D8FE-4819-A36C-6DEDAF088CC7"'
            'UPGRADE_WARNING: Couldn't resolve default property of object myTiles().See. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
            c = c + myTiles.Item("L" & myBottomTile(X, Y)).See(Side - (1 - 2 * (Side Mod 2)) * Me.BottomFlip(X, Y))
        End If
        ' Check surrounding tile
        Select Case Side
            Case 0 ' Top Left
                Y = Y - 1
                'UPGRADE_WARNING: Couldn't resolve default property of object Side. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
                Side = 3
            Case 1 ' Top Right
                X = X + 1
                'UPGRADE_WARNING: Couldn't resolve default property of object Side. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
                Side = 2
            Case 2 ' Bottom Left
                X = X - 1
                'UPGRADE_WARNING: Couldn't resolve default property of object Side. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
                Side = 1
            Case 3 ' Bottom Right
                Y = Y + 1
                'UPGRADE_WARNING: Couldn't resolve default property of object Side. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
                Side = 0
            Case 4 ' Up
                Y = Y - 1 : X = X + 1
                'UPGRADE_WARNING: Couldn't resolve default property of object Side. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
                Side = 7
            Case 5 ' Left
                Y = Y - 1 : X = X - 1
                'UPGRADE_WARNING: Couldn't resolve default property of object Side. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
                Side = 6
            Case 6 ' Right
                Y = Y + 1 : X = X + 1
                'UPGRADE_WARNING: Couldn't resolve default property of object Side. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
                Side = 5
            Case 7 ' Down
                Y = Y + 1 : X = X - 1
                'UPGRADE_WARNING: Couldn't resolve default property of object Side. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
                Side = 4
        End Select
        If X < 0 Or X > myWidth Or Y < 0 Or Y > myHeight Then
            c = 0
        Else
            If myTopTile(X, Y) > 0 Then
                'UPGRADE_WARNING: Couldn't resolve default property of object Side. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
                'UPGRADE_WARNING: Mod has a new behavior. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="9B7D5ADD-D8FE-4819-A36C-6DEDAF088CC7"'
                'UPGRADE_WARNING: Couldn't resolve default property of object myTiles().See. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
                c = c + myTiles.Item("L" & myTopTile(X, Y)).See(Side - (1 - 2 * (Side Mod 2)) * Me.TopFlip(X, Y))
            End If
            If myMiddleTile(X, Y) > 0 Then
                'UPGRADE_WARNING: Couldn't resolve default property of object Side. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
                'UPGRADE_WARNING: Mod has a new behavior. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="9B7D5ADD-D8FE-4819-A36C-6DEDAF088CC7"'
                'UPGRADE_WARNING: Couldn't resolve default property of object myTiles().See. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
                c = c + myTiles.Item("L" & myMiddleTile(X, Y)).See(Side - (1 - 2 * (Side Mod 2)) * Me.MiddleFlip(X, Y))
            End If
            If myBottomTile(X, Y) > 0 Then
                'UPGRADE_WARNING: Couldn't resolve default property of object Side. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
                'UPGRADE_WARNING: Mod has a new behavior. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="9B7D5ADD-D8FE-4819-A36C-6DEDAF088CC7"'
                'UPGRADE_WARNING: Couldn't resolve default property of object myTiles().See. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
                c = c + myTiles.Item("L" & myBottomTile(X, Y)).See(Side - (1 - 2 * (Side Mod 2)) * Me.BottomFlip(X, Y))
            End If
        End If
        See = Not (c > -1)
    End Function

    'UPGRADE_NOTE: Class_Initialize was upgraded to Class_Initialize_Renamed. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="A9E4979A-37FA-4718-9994-97DD76ED70A7"'
    Private Sub Class_Initialize_Renamed()
        Dim c, i As Short
        ReDim myBottomTile(31, 31)
        ReDim myMiddleTile(31, 31)
        ReDim myTopTile(31, 31)
        ReDim myVisible(31, 31)
        ReDim myEncPointer(31, 31)
        myWidth = 31 : myHeight = 31
        For c = 0 To 31 : For i = 0 To 31 : Me.Hidden(c, i) = True : Next i : Next c
        myDefaultWidth = myWidth : myDefaultHeight = myHeight
        myDefaultEncounters = 30
        myLeft = Int(myWidth / 16) - 1
        myTop = -Int(myWidth / 16) - 2
        myName = "Untitled"
        For c = 0 To 19
            ' Set number of runes
            Select Case c
                Case 0, 4, 8, 12, 16
                    myRunes(c) = 1
                Case 1, 2, 5, 6, 9, 10, 13, 14, 17, 18
                    myRunes(c) = 2
                Case 3, 7, 11, 15, 19
                    myRunes(c) = 3
            End Select
        Next c
        ' Default bmp is bdout2.bmp with 128 tiles
        myPic = -1
        myPictureFile = "bdout2.bmp"
    End Sub
    Public Sub New()
        MyBase.New()
        Class_Initialize_Renamed()
    End Sub
End Class

