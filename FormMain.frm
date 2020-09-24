VERSION 5.00
Begin VB.Form FormMain 
   Caption         =   "Bomberman X"
   ClientHeight    =   7200
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   7200
   BeginProperty Font 
      Name            =   "MS Sans Serif"
      Size            =   36
      Charset         =   0
      Weight          =   700
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "FormMain.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MDIChild        =   -1  'True
   ScaleHeight     =   480
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   480
   Begin VB.PictureBox pctScreen 
      BackColor       =   &H00000000&
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   7200
      Left            =   0
      ScaleHeight     =   476
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   476
      TabIndex        =   0
      Top             =   0
      Width           =   7200
      Begin VB.Timer tAnimator 
         Interval        =   100
         Left            =   120
         Top             =   120
      End
   End
End
Attribute VB_Name = "FormMain"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub Form_Load()
    ChDir App.Path
    Randomize
    Me.Show
    
    Call StartNewGame
    
    Dim ddsdPrimary As DDSURFACEDESC2
    Dim ddsdBack As DDSURFACEDESC2
    '----------------------------- Initialize DirectDraw object.
    Set DD = DX.DirectDrawCreate("")
    FormMain.Show
    DD.SetCooperativeLevel hWnd, DDSCL_NORMAL
    '----------------------------- Create Clipper.
    Set DDClip = DD.CreateClipper(0)
    DDClip.SetHWnd pctScreen.hWnd
    '----------------------------- Create Primary Surface.
    ddsdPrimary.lFlags = DDSD_CAPS
    ddsdPrimary.ddsCaps.lCaps = DDSCAPS_PRIMARYSURFACE
    Set DDSPrimary = DD.CreateSurface(ddsdPrimary)
    Call DDSPrimary.SetClipper(DDClip)
    '----------------------------- Create 1 BackBuffer Surface.
    ddsdBack.lFlags = DDSD_CAPS Or DDSD_WIDTH Or DDSD_HEIGHT
    ddsdBack.ddsCaps.lCaps = DDSCAPS_OFFSCREENPLAIN Or DDSCAPS_SYSTEMMEMORY
    ddsdBack.lWidth = SCR_WIDTH
    ddsdBack.lHeight = SCR_HEIGHT
    Set DDSBack = DD.CreateSurface(ddsdBack)
    DDSBack.SetFont FormMain.Font
    '----------------------------- Define BackBuffer RECT.
    rcBackBuffer.Right = ddsdBack.lWidth
    rcBackBuffer.Bottom = ddsdBack.lHeight
    '----------------------------- Initialize DDSurfaces we need in this application.
    Call InitSurfaces
    '----------------------------- Render Loop.
    bExecuting = True
    Do While bExecuting
        RenderFrame
        DoEvents
    Loop
End Sub

Private Sub Form_Resize()
pctScreen.Width = FormMain.ScaleWidth
   pctScreen.Height = FormMain.ScaleHeight
End Sub

Private Sub Form_Unload(Cancel As Integer)
    bExecuting = False
    End
End Sub

Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
    Dim i%
    Select Case Screen_Focus
        Case SCR_PLAYING
            With Player_1
            If .dying Then Exit Sub
            Select Case KeyCode
                Case vbKeyUp
                    .moving_down = False
                    .moving_up = True: .nFace = FACE_BACK
                Case vbKeyDown
                    .moving_up = False
                    .moving_down = True: .nFace = FACE_FRONT
                Case vbKeyLeft
                    .moving_right = False
                    .moving_left = True: .nFace = FACE_LEFT
                Case vbKeyRight
                    .moving_left = False
                    .moving_right = True: .nFace = FACE_RIGHT
            End Select
            End With
        Case Else
    End Select

End Sub

Private Sub Form_KeyUp(KeyCode As Integer, Shift As Integer)
Dim i%
    Select Case Screen_Focus
        Case SCR_PLAYING
            With Player_1
            Select Case KeyCode
                Case vbKeyUp: .moving_up = False
                Case vbKeyDown: .moving_down = False
                Case vbKeyLeft: .moving_left = False
                Case vbKeyRight: .moving_right = False
                Case vbKeySpace
                    For i = 1 To 4
                        If .Bomb(i).nColumn = 0 And .nDropped_Bomb < .nMax_Bomb Then
                            .Bomb(i).nColumn = .nColumn
                            .Bomb(i).nRow = .nRow
                            .Bomb(i).nFrame = 0
                            .nDropped_Bomb = .nDropped_Bomb + 1
                            nPlayer(.Bomb(i).nColumn, .Bomb(i).nRow) = OBSTACLE_BOMB
                            Exit For
                        End If
                    Next i
            End Select
            End With
        Case SCR_GAME_OVER: Call StartNewGame
        Case SCR_STAGE_CLEAR:
            If KeyCode = vbKeyEscape Then Call StartNewGame
        Case Else
    End Select
End Sub

Sub InitSurfaces()
    Dim CK As DDCOLORKEY
    CK.low = RGB(255, 0, 255)
    CK.high = CK.low
    Dim ddsd As DDSURFACEDESC2
    ddsd.lFlags = DDSD_CAPS Or DDSD_WIDTH Or DDSD_HEIGHT
    ddsd.ddsCaps.lCaps = DDSCAPS_OFFSCREENPLAIN Or DDSCAPS_SYSTEMMEMORY
    ddsd.lWidth = SCR_WIDTH
    ddsd.lHeight = SCR_HEIGHT
    Set ddsGameBkg = DD.CreateSurfaceFromFile("GameBkg3.bmp", ddsd)
    ddsd.lWidth = 128: ddsd.lHeight = 192
    Set ddsEnemy = DD.CreateSurfaceFromFile("Enemy.bmp", ddsd)
    ddsEnemy.SetColorKey DDCKEY_SRCBLT, CK
    ddsd.lWidth = 240: ddsd.lHeight = 168
    Set ddsExplode = DD.CreateSurfaceFromFile("Explode.bmp", ddsd)
    ddsExplode.SetColorKey DDCKEY_SRCBLT, CK
    ddsd.lWidth = 160: ddsd.lHeight = 192
    Set ddsBomb = DD.CreateSurfaceFromFile("Bomb.bmp", ddsd)
    ddsBomb.SetColorKey DDCKEY_SRCBLT, CK
    ddsd.lWidth = 160: ddsd.lHeight = 192
    Set ddsGameObj = DD.CreateSurfaceFromFile("Obstacle.bmp", ddsd)
    ddsGameObj.SetColorKey DDCKEY_SRCBLT, CK
    ddsd.lWidth = 128: ddsd.lHeight = 240
    Set ddsBombman = DD.CreateSurfaceFromFile("Bombman.bmp", ddsd)
    ddsBombman.SetColorKey DDCKEY_SRCBLT, CK
    
End Sub

Sub StartNewGame()
    Screen_Focus = SCR_PLAYING
    Call InitStage(1)
    Call InitNewPlayer
End Sub

Sub ClearStage()
Dim i%, j%
    For i = 1 To MAX_COLUMN
        For j = 1 To MAX_ROW
            nStage(i, j) = 0
            BoomBlock(i, j) = False
            FrameBlock(i, j) = False
            nPlayer(i, j) = 0
            nTreasure(i, j) = 0
        Next j
    Next i
End Sub

Sub InitStage(StageID As Integer)
Dim i%, j%, k%, Ctr%
Dim Clear_Player As MYPLAYER
    Call ClearStage
    Select Case StageID
        Case 1
            For i = 1 To MAX_COLUMN
                For j = 1 To MAX_ROW Step 2
                    '------------------------- Fill in Undestructable Object.
                    If i Mod 2 = 0 Then nStage(i, j) = 1
                Next j
            Next i
            For i = 1 To MAX_COLUMN
                For j = 1 To MAX_ROW
                    '------------------------- Fill in Destructable Object.
                    If nStage(i, j) = 0 And Int(Rnd * 10 + 1) >= 7 And Ctr < MAX_OBSTACLE Then
                        nStage(i, j) = OBSTACLE_DESTRUCTABLE
                        Ctr = Ctr + 1
                    End If
                Next j
            Next i
            
            '------------------------- Clear Start Point.
            nStage(MAX_COLUMN, MAX_ROW) = 0
            nStage(11, MAX_ROW) = 0
            nStage(10, MAX_ROW) = 0
            nStage(11, 13) = 0
            
            '------------------------- Random Door Location.
ReLocateDoor:
            i = Int(Rnd * 10 + 1)
            j = Int(Rnd * 12 + 1)
            If nStage(i, j) = OBSTACLE_DESTRUCTABLE Then
                nTreasure(i, j) = SPECIAL_DOOR
            Else: GoTo ReLocateDoor
            End If
            '------------------------- Put Enemy into Stage.
            For k = 1 To MAX_ENEMY
                Enemy(k) = Clear_Player
ReRandomEnemy:
                i = Int(Rnd * 10 + 1)
                j = Int(Rnd * 12 + 1)
                If nStage(i, j) = 0 And nPlayer(i, j) = 0 Then
                    With Enemy(k)
                    .nColumn = i: .nRow = j
                    .nDestColumn = .nColumn: .nDestRow = .nRow
                    .nFrame = 0
                    .nFace = FACE_FRONT
                    .nDestX = 48 + (.nColumn - 1) * TILE_SIZE
                    .nDestY = ((.nRow - 1) * TILE_SIZE)
                    .nMove_speed = 0.4
    
                    nPlayer(.nColumn, .nRow) = k
                    End With
                Else: GoTo ReRandomEnemy
                End If
            Next k
        Case 2
            
        Case Else
    End Select
End Sub

Sub InitNewPlayer()
Dim i%
Dim Clear_Player  As MYPLAYER
    Player_1 = Clear_Player
    With Player_1
        .nColumn = MAX_COLUMN: .nRow = MAX_ROW
        .nFrame = 1: .nFace = FACE_FRONT
        .nDestX = 48 + (.nColumn - 1) * TILE_SIZE
        .nDestY = ((.nRow - 1) * TILE_SIZE)
        .nMove_speed = 2
        .nMax_Bomb = 1
    End With
End Sub

Sub UpdatePlayer()
    Dim i%, j%
    With Player_1
        '------------------------- Finding the current position on map.
        .nColumn = 1 + (Int((.nDestX - 32) / TILE_SIZE))
        .nRow = 1 + (Int((.nDestY + 16) / TILE_SIZE))
        If StageClear Then Screen_Focus = SCR_STAGE_CLEAR: Exit Sub
        For i = 1 To MAX_ENEMY
            If .nColumn = Enemy(i).nColumn And .nRow = Enemy(i).nRow Then TestDie (ID_PLAYER): Exit For
        Next i
        '------------------------- Do moving upnorth.
        If .moving_up Then
            .nFace = FACE_BACK
            If .nDestY > (.nRow - 1) * TILE_SIZE Then
                .nDestY = .nDestY - .nMove_speed
                If .nDestX > 48 + (.nColumn - 1) * TILE_SIZE Then
                    .nDestX = .nDestX - .nMove_speed
                ElseIf .nDestX < 48 + (.nColumn - 1) * TILE_SIZE Then
                    .nDestX = .nDestX + .nMove_speed
                End If
            ElseIf .nRow > 1 Then
                If nStage(.nColumn, .nRow - 1) = 0 Then
                    .nDestY = .nDestY - .nMove_speed
                    If .nDestX > 48 + (.nColumn - 1) * TILE_SIZE Then
                        .nDestX = .nDestX - .nMove_speed
                    ElseIf .nDestX < 48 + (.nColumn - 1) * TILE_SIZE Then
                        .nDestX = .nDestX + .nMove_speed
                    End If
                Else
                    If .nColumn > 1 Then
                        If nStage(.nColumn - 1, .nRow - 1) = 0 And .nDestX + 8 < 48 + (.nColumn - 1) * TILE_SIZE Then .nDestX = .nDestX - .nMove_speed
                    End If
                    If .nColumn < MAX_COLUMN Then
                        If nStage(.nColumn + 1, .nRow - 1) = 0 And .nDestX - 8 > 48 + (.nColumn - 1) * TILE_SIZE Then .nDestX = .nDestX + .nMove_speed
                    End If
                End If
            End If
        End If
        '------------------------- Do moving downward.
        If .moving_down Then
            .nFace = FACE_FRONT
            If .nDestY < (.nRow - 1) * TILE_SIZE Then
                .nDestY = .nDestY + .nMove_speed
                If .nDestX > 48 + (.nColumn - 1) * TILE_SIZE Then
                    .nDestX = .nDestX - .nMove_speed
                ElseIf .nDestX < 48 + (.nColumn - 1) * TILE_SIZE Then
                    .nDestX = .nDestX + .nMove_speed
                End If
            ElseIf .nRow < MAX_ROW Then
                If nStage(.nColumn, .nRow + 1) = 0 Then
                    .nDestY = .nDestY + .nMove_speed
                    If .nDestX > 48 + (.nColumn - 1) * TILE_SIZE Then
                        .nDestX = .nDestX - .nMove_speed
                    ElseIf .nDestX < 48 + (.nColumn - 1) * TILE_SIZE Then
                        .nDestX = .nDestX + .nMove_speed
                    End If
                Else
                    If .nColumn > 1 Then
                        If nStage(.nColumn - 1, .nRow + 1) = 0 And .nDestX + 8 < 48 + (.nColumn - 1) * TILE_SIZE Then .nDestX = .nDestX - .nMove_speed
                    End If
                    If .nColumn < MAX_COLUMN Then
                        If nStage(.nColumn + 1, .nRow + 1) = 0 And .nDestX - 8 > 48 + (.nColumn - 1) * TILE_SIZE Then .nDestX = .nDestX + .nMove_speed
                    End If
                End If
            End If
        End If
        '------------------------- Do moving leftward.
        If .moving_left Then
            .nFace = FACE_LEFT
            If .nDestX > 48 + (.nColumn - 1) * TILE_SIZE Then
                .nDestX = .nDestX - .nMove_speed
                    If .nDestY > (.nRow - 1) * TILE_SIZE Then
                        .nDestY = .nDestY - .nMove_speed
                    ElseIf .nDestY < (.nRow - 1) * TILE_SIZE Then
                        .nDestY = .nDestY + .nMove_speed
                    End If
            ElseIf .nColumn > 1 Then
                If nStage(.nColumn - 1, .nRow) = 0 Then
                    .nDestX = .nDestX - .nMove_speed
                    If .nDestY > (.nRow - 1) * TILE_SIZE Then
                        .nDestY = .nDestY - .nMove_speed
                    ElseIf .nDestY < (.nRow - 1) * TILE_SIZE Then
                        .nDestY = .nDestY + .nMove_speed
                    End If
                Else
                    If .nRow > 1 Then
                        If nStage(.nColumn - 1, .nRow - 1) = 0 And .nDestY + 8 < (.nRow - 1) * TILE_SIZE Then .nDestY = .nDestY - .nMove_speed
                    End If
                    If .nRow < MAX_ROW Then
                        If nStage(.nColumn - 1, .nRow + 1) = 0 And .nDestY - 8 > (.nRow - 1) * TILE_SIZE Then .nDestY = .nDestY + .nMove_speed
                    End If
                End If
            End If
        End If
        '------------------------- Do moving rightward.
        If .moving_right Then
            .nFace = FACE_RIGHT
            If .nDestX < 48 + (.nColumn - 1) * TILE_SIZE Then
                .nDestX = .nDestX + .nMove_speed
                    If .nDestY > (.nRow - 1) * TILE_SIZE Then
                        .nDestY = .nDestY - .nMove_speed
                    ElseIf .nDestY < (.nRow - 1) * TILE_SIZE Then
                        .nDestY = .nDestY + .nMove_speed
                    End If
            ElseIf .nColumn < MAX_COLUMN Then
                If nStage(.nColumn + 1, .nRow) = 0 Then
                    .nDestX = .nDestX + .nMove_speed
                    If .nDestY > (.nRow - 1) * TILE_SIZE Then
                        .nDestY = .nDestY - .nMove_speed
                    ElseIf .nDestY < (.nRow - 1) * TILE_SIZE Then
                        .nDestY = .nDestY + .nMove_speed
                    End If
                Else
                    If .nRow > 1 Then
                        If nStage(.nColumn + 1, .nRow - 1) = 0 And .nDestY + 8 < (.nRow - 1) * TILE_SIZE Then .nDestY = .nDestY - .nMove_speed
                    End If
                    If .nRow < MAX_ROW Then
                        If nStage(.nColumn + 1, .nRow + 1) = 0 And .nDestY - 8 > (.nRow - 1) * TILE_SIZE Then .nDestY = .nDestY + .nMove_speed
                    End If
                End If
            End If
        End If
        If Not .moving_up And Not .moving_down And Not .moving_left And Not .moving_right And Not .dying Then .nFrame = 1
    End With
End Sub

Sub UpdateEnemy()
Dim i%, j%
For i = 1 To MAX_ENEMY
    With Enemy(i)
    If .nColumn > 0 And Not .dying Then
        Select Case .nFace
            Case FACE_FRONT
                If .nDestY < (.nDestRow - 1) * TILE_SIZE Then
                    .nDestY = .nDestY + .nMove_speed
                Else
                    nPlayer(.nColumn, .nRow) = 0
                    '------------------------- Finding the current position on map.
                    .nColumn = 1 + (Int((.nDestX - 32) / TILE_SIZE))
                    .nRow = 1 + (Int((.nDestY + 16) / TILE_SIZE))
                    
                    If .nRow < MAX_ROW Then
                        nPlayer(.nColumn, .nRow) = i
                        If nStage(.nColumn, .nRow + 1) = 0 And nPlayer(.nColumn, .nRow + 1) = 0 Then
                            .nDestRow = .nRow + 1
                        Else
DoNextMove:
                            Select Case TryNewMove(i)
                                Case FACE_FRONT:
                                    .nFace = FACE_FRONT
                                    .nDestRow = .nDestRow + 1
                                Case FACE_BACK:
                                    .nFace = FACE_BACK
                                    .nDestRow = .nDestRow - 1
                                    
                                Case FACE_LEFT:
                                    .nFace = FACE_LEFT
                                    .nDestColumn = .nDestColumn - 1
                                Case FACE_RIGHT:
                                    .nFace = FACE_RIGHT
                                    .nDestColumn = .nDestColumn + 1
                                Case Else
                                    .nDestColumn = .nColumn
                                    .nDestRow = .nDestRow
                            End Select
                        End If
                    Else: GoTo DoNextMove
                    End If
                End If
            Case FACE_BACK
                If .nDestY > (.nDestRow - 1) * TILE_SIZE Then
                    .nDestY = .nDestY - .nMove_speed
                Else
                    nPlayer(.nColumn, .nRow) = 0
                    '------------------------- Finding the current position on map.
                    .nColumn = 1 + (Int((.nDestX - 32) / TILE_SIZE))
                    .nRow = 1 + (Int((.nDestY + 16) / TILE_SIZE))
                    
                    If .nRow > 1 Then
                        nPlayer(.nColumn, .nRow) = i
                        If nStage(.nColumn, .nRow - 1) = 0 And nPlayer(.nColumn, .nRow - 1) = 0 Then
                            .nDestRow = .nDestRow - 1
                        Else: GoTo DoNextMove
                        End If
                    Else: GoTo DoNextMove
                    End If
                End If
            Case FACE_LEFT
                If .nDestX > 48 + (.nDestColumn - 1) * TILE_SIZE Then
                    .nDestX = .nDestX - .nMove_speed
                Else
                    nPlayer(.nColumn, .nRow) = 0
                    '------------------------- Finding the current position on map.
                    .nColumn = 1 + (Int((.nDestX - 32) / TILE_SIZE))
                    .nRow = 1 + (Int((.nDestY + 16) / TILE_SIZE))
                    
                    If .nColumn > 1 Then
                        nPlayer(.nColumn, .nRow) = i
                        If nStage(.nColumn - 1, .nRow) = 0 And nPlayer(.nColumn - 1, .nRow) = 0 Then
                            .nDestColumn = .nDestColumn - 1
                        Else: GoTo DoNextMove
                        End If
                    Else: GoTo DoNextMove
                    End If
                End If
            Case FACE_RIGHT
                If .nDestX < 48 + (.nDestColumn - 1) * TILE_SIZE Then
                    .nDestX = .nDestX + .nMove_speed
                Else
                    nPlayer(.nColumn, .nRow) = 0
                    '------------------------- Finding the current position on map.
                    .nColumn = 1 + (Int((.nDestX - 32) / TILE_SIZE))
                    .nRow = 1 + (Int((.nDestY + 16) / TILE_SIZE))
                    
                    If .nColumn < MAX_COLUMN Then
                        nPlayer(.nColumn, .nRow) = i
                        If nStage(.nColumn + 1, .nRow) = 0 And nPlayer(.nColumn + 1, .nRow) = 0 Then
                            .nDestColumn = .nDestColumn + 1
                        Else: GoTo DoNextMove
                        End If
                    Else: GoTo DoNextMove
                    End If
                End If
        End Select
    End If
    End With
Next i
End Sub

Function TryNewMove(ID As Integer) As Byte
With Enemy(ID)
    Select Case Int(Rnd * 4)
        Case FACE_FRONT
            If .nRow < MAX_ROW Then
                If nStage(.nColumn, .nRow + 1) = 0 Then
                    TryNewMove = FACE_FRONT
                Else: TryNewMove = 9
                End If
            Else: TryNewMove = 9
            End If
        Case FACE_BACK
            If .nRow > 1 Then
                If nStage(.nColumn, .nRow - 1) = 0 Then
                    TryNewMove = FACE_BACK
                Else: TryNewMove = 9
                End If
            Else: TryNewMove = 9
            End If
        Case FACE_LEFT
            If .nColumn > 1 Then
                If nStage(.nColumn - 1, .nRow) = 0 Then
                    TryNewMove = FACE_LEFT
                Else: TryNewMove = 9
                End If
            Else: TryNewMove = 9
            End If
        Case FACE_RIGHT
             If .nColumn < MAX_COLUMN Then
                If nStage(.nColumn + 1, .nRow) = 0 Then
                    TryNewMove = FACE_RIGHT
                Else: TryNewMove = 9
                End If
            Else: TryNewMove = 9
            End If
    End Select
End With
End Function

Sub AnimateObstacle()
Dim i%, j%
    For i = 1 To MAX_COLUMN
        For j = 1 To MAX_ROW
            If FrameBlock(i, j) < 4 Then
                FrameBlock(i, j) = FrameBlock(i, j) + 1
            Else: FrameBlock(i, j) = 0
            End If
        Next j
    Next i
End Sub

Sub AnimatePlayer()
Dim i%
    '--------------------- Animated Enemy.
    For i = 1 To MAX_ENEMY
        With Enemy(i)
        If .nColumn > 0 Then
            If Not .dying Then
                If .nFrame < 3 Then
                    .nFrame = .nFrame + 1
                Else: .nFrame = 0
                End If
            Else
                If .nFrame < 4 Then
                    .nFrame = .nFrame + 1
                Else
                    .nFrame = 0
                    nPlayer(.nColumn, .nRow) = 0
                    .nColumn = 0: .nRow = 0
                    .dying = False
                End If
            End If
        End If
        End With
    Next i
    '--------------------- Animated Player.
    With Player_1
    If Not .dying Then
        If .moving_up Or .moving_down Or .moving_left Or .moving_right Then
            If .nFrame < 3 Then
                .nFrame = .nFrame + 1
            Else: .nFrame = 0
            End If
        End If
    Else
        If Not .died Then
            If .nFrame < 4 Then
                .nFrame = .nFrame + 1
            Else:
                .nFrame = 0
                .died = True
            End If
        Else
            If .nFrame < 3 Then
                .nFrame = .nFrame + 1
            Else:
                .nFrame = 0
                Screen_Focus = SCR_GAME_OVER
            End If
        End If
    End If
    End With
End Sub

Sub AnimateBomb()
Dim i%, j%, k%
With Player_1
    For i = 1 To 4
        If .Bomb(i).nColumn > 0 Then
            If .Bomb(i).nFrame < 4 Then
                .Bomb(i).nFrame = .Bomb(i).nFrame + 1
                If .Bomb(i).is_booming Then
                    '--------------------- If there is Enemy at Center then BOOM!
                    j = nPlayer(.Bomb(i).nColumn, .Bomb(i).nRow)
                    Call TestDie(j)
                    If .Bomb(i).nColumn = .nColumn And .Bomb(i).nRow = .nRow Then Call TestDie(ID_PLAYER)
                    '--------------------- If there is Enemy at Above then BOOM!
                    If .Bomb(i).nRow > 1 Then
                        j = nPlayer(.Bomb(i).nColumn, .Bomb(i).nRow - 1)
                        Call TestDie(j)
                        If .Bomb(i).nColumn = .nColumn And .Bomb(i).nRow - 1 = .nRow Then Call TestDie(ID_PLAYER)
                    End If
                    '--------------------- If there is Enemy at Under then BOOM!
                    If .Bomb(i).nRow < MAX_ROW Then
                        j = nPlayer(.Bomb(i).nColumn, .Bomb(i).nRow + 1)
                        Call TestDie(j)
                        If .Bomb(i).nColumn = .nColumn And .Bomb(i).nRow + 1 = .nRow Then Call TestDie(ID_PLAYER)
                    End If
                    '--------------------- If there is Enemy at Left then BOOM!
                    If .Bomb(i).nColumn > 1 Then
                        j = nPlayer(.Bomb(i).nColumn - 1, .Bomb(i).nRow)
                        Call TestDie(j)
                        If .Bomb(i).nColumn - 1 = .nColumn And .Bomb(i).nRow = .nRow Then Call TestDie(ID_PLAYER)
                    End If
                    '--------------------- If there is Enemy at Right then BOOM!
                    If .Bomb(i).nColumn < MAX_COLUMN Then
                        j = nPlayer(.Bomb(i).nColumn + 1, .Bomb(i).nRow)
                        Call TestDie(j)
                        If .Bomb(i).nColumn + 1 = .nColumn And .Bomb(i).nRow = .nRow Then Call TestDie(ID_PLAYER)
                    End If
                End If
            Else:
                .Bomb(i).nFrame = 0
                If .Bomb(i).is_booming Then
                    '--------------------- If obstacle on above then disappear.
                    If .Bomb(i).nRow > 1 Then
                        If nStage(.Bomb(i).nColumn, .Bomb(i).nRow - 1) = OBSTACLE_DESTRUCTABLE Then
                            BoomBlock(.Bomb(i).nColumn, .Bomb(i).nRow - 1) = False
                            nStage(.Bomb(i).nColumn, .Bomb(i).nRow - 1) = 0
                        End If
                    End If
                    '--------------------- If obstacle at under then disappear
                    If .Bomb(i).nRow < MAX_ROW Then
                        If nStage(.Bomb(i).nColumn, .Bomb(i).nRow + 1) = OBSTACLE_DESTRUCTABLE Then
                            BoomBlock(.Bomb(i).nColumn, .Bomb(i).nRow + 1) = False
                            nStage(.Bomb(i).nColumn, .Bomb(i).nRow + 1) = 0
                        End If
                    End If
                    '--------------------- If obstacle on the left then disappear.
                    If .Bomb(i).nColumn > 1 Then
                        If nStage(.Bomb(i).nColumn - 1, .Bomb(i).nRow) = OBSTACLE_DESTRUCTABLE Then
                            BoomBlock(.Bomb(i).nColumn - 1, .Bomb(i).nRow) = False
                            nStage(.Bomb(i).nColumn - 1, .Bomb(i).nRow) = 0
                        End If
                    End If
                    '--------------------- If obstacle on the left then disappear.
                    If .Bomb(i).nColumn < MAX_COLUMN Then
                        If nStage(.Bomb(i).nColumn + 1, .Bomb(i).nRow) = OBSTACLE_DESTRUCTABLE Then
                            BoomBlock(.Bomb(i).nColumn + 1, .Bomb(i).nRow) = False
                            nStage(.Bomb(i).nColumn + 1, .Bomb(i).nRow) = 0
                        End If
                    End If
                    .Bomb(i).nColumn = 0
                    .Bomb(i).nRow = 0
                    .Bomb(i).is_booming = False
                    .nDropped_Bomb = .nDropped_Bomb - 1
                End If
                              
            End If
            
            
            If Not .Bomb(i).is_booming Then
                .Bomb(i).nCtr = .Bomb(i).nCtr + 1
                If .Bomb(i).nCtr = 10 Then
                    .Bomb(i).is_booming = True
                    .Bomb(i).nFrame = 0
                    .Bomb(i).nCtr = 0
                    nPlayer(.Bomb(i).nColumn, .Bomb(i).nRow) = 0
                    '--------------------- Destroy nearby Obstacle on above.
                    If .Bomb(i).nRow > 1 Then
                        If nStage(.Bomb(i).nColumn, .Bomb(i).nRow - 1) = OBSTACLE_DESTRUCTABLE Then
                            FrameBlock(.Bomb(i).nColumn, .Bomb(i).nRow - 1) = 0
                            BoomBlock(.Bomb(i).nColumn, .Bomb(i).nRow - 1) = True
                        End If
                    End If
                    '--------------------- Destroy nearby Obstacle at under.
                    If .Bomb(i).nRow < MAX_ROW Then
                        If nStage(.Bomb(i).nColumn, .Bomb(i).nRow + 1) = OBSTACLE_DESTRUCTABLE Then
                            FrameBlock(.Bomb(i).nColumn, .Bomb(i).nRow + 1) = 0
                            BoomBlock(.Bomb(i).nColumn, .Bomb(i).nRow + 1) = True
                        End If
                    End If
                    '--------------------- Destroy nearby Obstacle on the left.
                    If .Bomb(i).nColumn > 1 Then
                        If nStage(.Bomb(i).nColumn - 1, .Bomb(i).nRow) = OBSTACLE_DESTRUCTABLE Then
                            FrameBlock(.Bomb(i).nColumn - 1, .Bomb(i).nRow) = 0
                            BoomBlock(.Bomb(i).nColumn - 1, .Bomb(i).nRow) = True
                        End If
                    End If
                    '--------------------- Destroy nearby Obstacle on the right.
                    If .Bomb(i).nColumn < MAX_COLUMN Then
                        If nStage(.Bomb(i).nColumn + 1, .Bomb(i).nRow) = OBSTACLE_DESTRUCTABLE Then
                            FrameBlock(.Bomb(i).nColumn + 1, .Bomb(i).nRow) = 0
                            BoomBlock(.Bomb(i).nColumn + 1, .Bomb(i).nRow) = True
                        End If
                    End If
                End If      '.Bomb(i).nCtr = 10
            End If      'Not .Bomb(i).is_booming
            
        End If      '.Bomb(i).nColumn > 0
    Next i
End With
End Sub

Sub TestDie(ID As Integer)
    Select Case ID
        Case 0
        Case 1 To MAX_ENEMY
            If Not Enemy(ID).dying Then
                Enemy(ID).nFrame = 0
                Enemy(ID).dying = True
            End If
        Case ID_PLAYER
            If Not Player_1.dying Then
                Player_1.nFrame = 0
                Player_1.dying = True
            End If
    End Select
End Sub

Function StageClear() As Boolean
Dim i%, j%, k%, Ctr%
    For i = 1 To MAX_COLUMN
        For j = 1 To MAX_ROW
            If nTreasure(i, j) = SPECIAL_DOOR Then
                If i = Player_1.nColumn And j = Player_1.nRow Then
                    For k = 1 To MAX_ENEMY
                        If Enemy(k).nColumn = 0 Then Ctr = Ctr + 1
                    Next k
                    If Ctr = 4 Then StageClear = True
                End If
                Exit Function
            End If
        Next j
    Next i
End Function

Private Sub tAnimator_Timer()
    Call AnimateObstacle
    Call AnimatePlayer
    Call AnimateBomb
End Sub

Sub RenderFrame()
    Dim bRestore As Boolean
    '--------------------- this will keep us from trying to blt in case we lose the surfaces (another fullscreen app takes over)
    bRestore = False
    Do Until DD.TestCooperativeLevel = DD_OK
        DoEvents
        bRestore = True
    Loop
    '--------------------- if we lost and got back the surfaces, then restore them
    DoEvents
    If bRestore Then
        bRestore = False
        DD.RestoreAllSurfaces
        Call InitSurfaces
    End If
    
    Dim i%, j%
    Dim rcDest As RECT
    Dim rcSrc As RECT
    Select Case Screen_Focus
    
        Case SCR_PLAYING
            rcDest.Left = 0: rcDest.Top = 0:
            rcDest.Right = SCR_WIDTH
            rcDest.Bottom = SCR_HEIGHT
            rcSrc = rcDest
            '------------------------- Blt Game Background.
            Call DDSBack.Blt(rcDest, ddsGameBkg, rcSrc, DDBLT_WAIT)
            
            For j = 1 To MAX_ROW
                For i = 1 To MAX_COLUMN
                    rcDest.Left = 48 + (i - 1) * 32
                    rcDest.Top = 4 + (j - 1) * 32
                    rcDest.Right = rcDest.Left + TILE_SIZE
                    rcDest.Bottom = rcDest.Top + 48
                    
                    '------------------------- Blt Special.
                    If nTreasure(i, j) = SPECIAL_DOOR And nStage(i, j) = 0 Then
                        rcSrc.Left = FrameBlock(i, j) * TILE_SIZE
                        rcSrc.Top = 144
                        rcSrc.Right = rcSrc.Left + TILE_SIZE
                        rcSrc.Bottom = rcSrc.Top + 48
                        Call DDSBack.Blt(rcDest, ddsGameObj, rcSrc, DDBLT_WAIT Or DDBLT_KEYSRC)
                    End If
                
                    If nStage(i, j) > 0 Then
                        rcSrc.Left = FrameBlock(i, j) * TILE_SIZE
                        '------------------------- Blt Obstacle on the map.
                        If nStage(i, j) = OBSTACLE_DESTRUCTABLE And BoomBlock(i, j) Then
                            rcSrc.Top = 96
                        Else: rcSrc.Top = (nStage(i, j) - 1) * 48
                        End If
                        rcSrc.Right = rcSrc.Left + TILE_SIZE
                        rcSrc.Bottom = rcSrc.Top + 48
                        Call DDSBack.Blt(rcDest, ddsGameObj, rcSrc, DDBLT_WAIT Or DDBLT_KEYSRC)
                    End If
                    
                Next i
            Next j
            
            With Player_1
            '------------------------------- Blt Bomb.
            For i = 1 To .nMax_Bomb
                rcDest.Left = 48 + (.Bomb(i).nColumn - 1) * 32
                rcDest.Top = 16 + (.Bomb(i).nRow - 1) * 32
                rcDest.Right = rcDest.Left + TILE_SIZE
                rcDest.Bottom = rcDest.Top + TILE_SIZE
                If Not .Bomb(i).is_booming Then
                    rcSrc.Left = .Bomb(i).nFrame * TILE_SIZE
                    rcSrc.Top = 0
                    rcSrc.Right = rcSrc.Left + TILE_SIZE
                    rcSrc.Bottom = rcSrc.Top + TILE_SIZE
                    Call DDSBack.Blt(rcDest, ddsBomb, rcSrc, DDBLT_WAIT Or DDBLT_KEYSRC)
                Else
                    '----------------- Blt Center Explosion.
                    rcSrc.Left = .Bomb(i).nFrame * TILE_SIZE
                    rcSrc.Top = TILE_SIZE
                    rcSrc.Right = rcSrc.Left + TILE_SIZE
                    rcSrc.Bottom = rcSrc.Top + TILE_SIZE
                    Call DDSBack.Blt(rcDest, ddsBomb, rcSrc, DDBLT_WAIT Or DDBLT_KEYSRC)
                    '----------------- Blt Upward Explosion.
                    If .Bomb(i).nRow > 1 Then
                        If nStage(.Bomb(i).nColumn, .Bomb(i).nRow - 1) = 0 Then
                            rcDest.Left = 48 + (.Bomb(i).nColumn - 1) * 32
                            rcDest.Top = 16 + (.Bomb(i).nRow - 1 - 1) * 32
                            rcDest.Right = rcDest.Left + TILE_SIZE
                            rcDest.Bottom = rcDest.Top + TILE_SIZE
                            rcSrc.Left = .Bomb(i).nFrame * TILE_SIZE
                            rcSrc.Top = TILE_SIZE * 2
                            rcSrc.Right = rcSrc.Left + TILE_SIZE
                            rcSrc.Bottom = rcSrc.Top + TILE_SIZE
                            Call DDSBack.Blt(rcDest, ddsBomb, rcSrc, DDBLT_WAIT Or DDBLT_KEYSRC)
                        End If
                    End If
                    '----------------- Blt Downward Explosion.
                    If .Bomb(i).nRow < MAX_ROW Then
                        If nStage(.Bomb(i).nColumn, .Bomb(i).nRow + 1) = 0 Then
                            rcDest.Left = 48 + (.Bomb(i).nColumn - 1) * 32
                            rcDest.Top = 16 + (.Bomb(i).nRow + 1 - 1) * 32
                            rcDest.Right = rcDest.Left + TILE_SIZE
                            rcDest.Bottom = rcDest.Top + TILE_SIZE
                            rcSrc.Left = .Bomb(i).nFrame * TILE_SIZE
                            rcSrc.Top = TILE_SIZE * 3
                            rcSrc.Right = rcSrc.Left + TILE_SIZE
                            rcSrc.Bottom = rcSrc.Top + TILE_SIZE
                            Call DDSBack.Blt(rcDest, ddsBomb, rcSrc, DDBLT_WAIT Or DDBLT_KEYSRC)
                        End If
                    End If
                    '----------------- Blt Leftward Explosion.
                    If .Bomb(i).nColumn > 1 Then
                        If nStage(.Bomb(i).nColumn - 1, .Bomb(i).nRow) = 0 Then
                            rcDest.Left = 48 + (.Bomb(i).nColumn - 1 - 1) * 32
                            rcDest.Top = 16 + (.Bomb(i).nRow - 1) * 32
                            rcDest.Right = rcDest.Left + TILE_SIZE
                            rcDest.Bottom = rcDest.Top + TILE_SIZE
                            rcSrc.Left = .Bomb(i).nFrame * TILE_SIZE
                            rcSrc.Top = TILE_SIZE * 4
                            rcSrc.Right = rcSrc.Left + TILE_SIZE
                            rcSrc.Bottom = rcSrc.Top + TILE_SIZE
                            Call DDSBack.Blt(rcDest, ddsBomb, rcSrc, DDBLT_WAIT Or DDBLT_KEYSRC)
                        End If
                    End If
                    '----------------- Blt Leftward Explosion.
                    If .Bomb(i).nColumn < MAX_COLUMN Then
                        If nStage(.Bomb(i).nColumn + 1, .Bomb(i).nRow) = 0 Then
                            rcDest.Left = 48 + (.Bomb(i).nColumn + 1 - 1) * 32
                            rcDest.Top = 16 + (.Bomb(i).nRow - 1) * 32
                            rcDest.Right = rcDest.Left + TILE_SIZE
                            rcDest.Bottom = rcDest.Top + TILE_SIZE
                            rcSrc.Left = .Bomb(i).nFrame * TILE_SIZE
                            rcSrc.Top = TILE_SIZE * 5
                            rcSrc.Right = rcSrc.Left + TILE_SIZE
                            rcSrc.Bottom = rcSrc.Top + TILE_SIZE
                            Call DDSBack.Blt(rcDest, ddsBomb, rcSrc, DDBLT_WAIT Or DDBLT_KEYSRC)
                        End If
                    End If
                End If
            Next i
            
            '------------------------------- Blt Player 1.
            Call UpdatePlayer
            If Not .dying Then
                rcDest.Left = .nDestX:  rcDest.Top = .nDestY:
                rcDest.Right = rcDest.Left + TILE_SIZE
                rcDest.Bottom = rcDest.Top + 48
                rcSrc.Left = .nFrame * TILE_SIZE:  rcSrc.Top = .nFace * 48
                rcSrc.Right = rcSrc.Left + TILE_SIZE
                rcSrc.Bottom = rcSrc.Top + 48
                Call DDSBack.Blt(rcDest, ddsBombman, rcSrc, DDBLT_WAIT Or DDBLT_KEYSRC)
            Else
                If Not .died Then
                    rcDest.Left = .nDestX - 8: rcDest.Top = .nDestY - 48
                    rcDest.Right = rcDest.Left + 48
                    rcDest.Bottom = rcDest.Top + 96
                    rcSrc.Left = .nFrame * 48:  rcSrc.Top = 72
                    rcSrc.Right = rcSrc.Left + 48
                    rcSrc.Bottom = rcSrc.Top + 96
                    Call DDSBack.Blt(rcDest, ddsExplode, rcSrc, DDBLT_WAIT Or DDBLT_KEYSRC)
                Else
                    rcDest.Left = .nDestX:  rcDest.Top = .nDestY:
                    rcDest.Right = rcDest.Left + TILE_SIZE
                    rcDest.Bottom = rcDest.Top + 48
                    rcSrc.Left = .nFrame * TILE_SIZE:  rcSrc.Top = 192
                    rcSrc.Right = rcSrc.Left + TILE_SIZE
                    rcSrc.Bottom = rcSrc.Top + 48
                    Call DDSBack.Blt(rcDest, ddsBombman, rcSrc, DDBLT_WAIT Or DDBLT_KEYSRC)
                End If
            End If
            End With
            '------------------------------- Blt Enemy.
            For i = 1 To MAX_ENEMY
            Call UpdateEnemy
                With Enemy(i)
                If .nColumn > 0 Then
                    If Not .dying Then
                        rcDest.Left = .nDestX:  rcDest.Top = .nDestY:
                        rcDest.Right = rcDest.Left + TILE_SIZE
                        rcDest.Bottom = rcDest.Top + 48
                        rcSrc.Left = .nFrame * TILE_SIZE
                        rcSrc.Top = .nFace * 48
                        rcSrc.Right = rcSrc.Left + TILE_SIZE
                        rcSrc.Bottom = rcSrc.Top + 48
                        Call DDSBack.Blt(rcDest, ddsEnemy, rcSrc, DDBLT_WAIT Or DDBLT_KEYSRC)
                    Else
                        rcDest.Left = .nDestX - 8: rcDest.Top = .nDestY - 24
                        rcDest.Right = rcDest.Left + 48
                        rcDest.Bottom = rcDest.Top + 72
                        rcSrc.Left = .nFrame * 48: rcSrc.Top = 0
                        rcSrc.Right = rcSrc.Left + 48
                        rcSrc.Bottom = rcSrc.Top + 72
                        Call DDSBack.Blt(rcDest, ddsExplode, rcSrc, DDBLT_WAIT Or DDBLT_KEYSRC)
                    End If
                End If
                End With
            Next i
        Case SCR_GAME_OVER
            DDSBack.BltColorFill rcBackBuffer, vbBlack
            '------------------------------- Blt Died Player.
            With Player_1
            rcDest.Left = .nDestX:  rcDest.Top = .nDestY:
            rcDest.Right = rcDest.Left + TILE_SIZE
            rcDest.Bottom = rcDest.Top + 48
            rcSrc.Left = .nFrame * TILE_SIZE:  rcSrc.Top = 192
            rcSrc.Right = rcSrc.Left + TILE_SIZE
            rcSrc.Bottom = rcSrc.Top + 48
            Call DDSBack.Blt(rcDest, ddsBombman, rcSrc, DDBLT_WAIT Or DDBLT_KEYSRC)
            End With
            
            DDSBack.SetForeColor &H808080
            DDSBack.DrawText 100, 204, "GAME OVER", False
            DDSBack.SetForeColor vbWhite
            DDSBack.DrawText 96, 200, "GAME OVER", False
        Case SCR_STAGE_CLEAR
            DDSBack.BltColorFill rcBackBuffer, vbBlack
            '------------------------------- Blt Died Player.
            With Player_1
            rcDest.Left = .nDestX:  rcDest.Top = .nDestY:
            rcDest.Right = rcDest.Left + TILE_SIZE
            rcDest.Bottom = rcDest.Top + 48
            rcSrc.Left = TILE_SIZE:  rcSrc.Top = 0
            rcSrc.Right = rcSrc.Left + TILE_SIZE
            rcSrc.Bottom = rcSrc.Top + 48
            Call DDSBack.Blt(rcDest, ddsBombman, rcSrc, DDBLT_WAIT Or DDBLT_KEYSRC)
            End With
            
            DDSBack.SetForeColor &H808080
            DDSBack.DrawText 84, 204, "STAGE CLEAR", False
            DDSBack.SetForeColor vbWhite
            DDSBack.DrawText 80, 200, "STAGE CLEAR", False
        Case Else
        
    End Select
    '--------------------- Blit BackBuffer to Primary Screen.
    Call DX.GetWindowRect(FormMain.pctScreen.hWnd, rcPrimary)
    Call DDSPrimary.Blt(rcPrimary, DDSBack, rcBackBuffer, DDBLT_WAIT)
End Sub

