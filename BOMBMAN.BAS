Attribute VB_Name = "MBomber"
Option Explicit
Public Declare Function BitBlt Lib "gdi32" (ByVal hDestDC As Long, ByVal X As Long, ByVal Y As Long, ByVal nWidth As Long, ByVal nHeight As Long, ByVal hSrcDC As Long, ByVal xSrc As Long, ByVal ySrc As Long, ByVal dwRop As Long) As Long

Public Const SCR_WIDTH = 480
Public Const SCR_HEIGHT = 480
Public Const TILE_SIZE = 32

Public DX As New DirectX7
Public DD As DirectDraw7
Public DDClip As DirectDrawClipper
Public DDSPrimary As DirectDrawSurface7
Public DDSBack As DirectDrawSurface7
Public rcPrimary As RECT
Public rcBackBuffer As RECT

Public bExecuting As Boolean

Public Const SCR_PLAYING = 1
Public Const SCR_GAME_OVER = 2
Public Const SCR_STAGE_CLEAR = 3

Public Const FACE_FRONT = 0
Public Const FACE_BACK = 1
Public Const FACE_LEFT = 2
Public Const FACE_RIGHT = 3

Public Const ID_PLAYER = 5
Public Const OBSTACLE_BOMB = 3
Public Const OBSTACLE_DESTRUCTABLE = 2

Public Const SPECIAL_DOOR = 1

Public Const MAX_COLUMN = 12
Public Const MAX_ROW = 14
Public Const MAX_BOMB = 4
Public Const MAX_OBSTACLE = 60
Public Const MAX_ENEMY = 4


Public Type MYBOMB
    nRow As Byte
    nColumn As Byte
    nFrame As Byte
    nFace As Byte
    nCtr As Byte
    
    nDestX As Integer
    nDestY As Integer
    
    is_booming As Boolean
End Type

Public Type MYPLAYER
    nRow As Byte
    nColumn As Byte
    nDestColumn As Byte
    nDestRow As Byte
    nFrame As Byte
    nFace As Byte

    dying As Boolean
    died As Boolean
    moving_up As Boolean
    moving_down As Boolean
    moving_left As Boolean
    moving_right As Boolean
    nMove_speed As Single

    nDestX As Single
    nDestY As Single
    
    Bomb(1 To 4) As MYBOMB
    nMax_Bomb As Byte
    nDropped_Bomb As Byte
End Type

Public ddsGameBkg As DirectDrawSurface7
Public ddsGameObj As DirectDrawSurface7
Public ddsBombman As DirectDrawSurface7
Public ddsBomb As DirectDrawSurface7
Public ddsEnemy As DirectDrawSurface7
Public ddsExplode As DirectDrawSurface7

Public Player_1 As MYPLAYER
Public Enemy(1 To MAX_ENEMY) As MYPLAYER

Public nStage(1 To MAX_COLUMN, 1 To MAX_ROW) As Integer
Public nTreasure(1 To MAX_COLUMN, 1 To MAX_ROW) As Integer
Public BoomBlock(1 To MAX_COLUMN, 1 To MAX_ROW) As Boolean
Public FrameBlock(1 To MAX_COLUMN, 1 To MAX_ROW) As Integer
Public nPlayer(1 To MAX_COLUMN, 1 To MAX_ROW) As Integer

Public Screen_Focus%











