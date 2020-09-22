Attribute VB_Name = "modSetup"
Option Explicit
'The following API calls are for:

'blitting
Public Declare Function BitBlt Lib "gdi32" (ByVal hDestDC As Long, ByVal X As Long, ByVal Y As Long, ByVal nWidth As Long, ByVal nHeight As Long, ByVal hSrcDC As Long, ByVal xSrc As Long, ByVal ySrc As Long, ByVal dwRop As Long) As Long

'keyboard
Public Declare Function GetAsyncKeyState Lib "user32" (ByVal vKey As Long) As Integer

'creating buffers / loading sprites
Public Declare Function CreateCompatibleBitmap Lib "gdi32" (ByVal hdc As Long, ByVal nWidth As Long, ByVal nHeight As Long) As Long
Public Declare Function CreateCompatibleDC Lib "gdi32" (ByVal hdc As Long) As Long
Public Declare Function GetDC Lib "user32" (ByVal hwnd As Long) As Long

'loading sprites
Public Declare Function SelectObject Lib "gdi32" (ByVal hdc As Long, ByVal hObject As Long) As Long

'cleanup
Public Declare Function DeleteObject Lib "gdi32" (ByVal hObject As Long) As Long
Public Declare Function DeleteDC Lib "gdi32" (ByVal hdc As Long) As Long

'sound
Declare Function sndPlaySound Lib "winmm.dll" Alias "sndPlaySoundA" (ByVal lpszSoundName As String, ByVal uFlags As Long) As Long

'our Buffer's DC
Public myBackBuffer As Long
Public myBufferBMP As Long

'sound options
Public Enum SoundOps
       SND_SYNC = &H0
      SND_ASYNC = &H1
  SND_NODEFAULT = &H2
       SND_LOOP = &H8
     SND_NOSTOP = &H10
      SND_PURGE = &H40
     SND_NOWAIT = &H2000
End Enum

Public Type UDTLevel
 junction As Boolean 'is this a junction? In other words is it a block where the enemy or Pacman can move in more than 3 directions
 up As Boolean 'can one move to the block above ?
 down As Boolean 'can one move to the block below ?
 left As Boolean 'can one move to the block to the left ?
 right As Boolean 'can one move to the block to the right ?
 block As String * 1 'what is in this block (pill,superpill,wall or blank)
End Type

Public Type UDTGame
 lives As Integer
 score As Integer
 extralives As Integer
End Type

'variables to find path
Public StageMap As Map
Public Path As Path

Public pacman As New clsPacMan
Public Enemy() As New clsEnemy
Public Level(31, 31) As UDTLevel
Public Game As UDTGame
Public Rev(3) As Integer ' Reverse direction number
Public XD(3) As Integer ' Holds the X directions
Public YD(3) As Integer ' Holds the Y directions
Public i As Long
Public j As Long
Public k As Long

Public Sub AddToScore(points As Integer)
 Game.score = Game.score + points
 frmGame.lblScore.Caption = "Points: " + CStr(Game.score)
End Sub

Public Sub AddToLives(lives As Integer)
 Game.lives = Game.lives + lives
 frmGame.lblLives.Caption = "X " + CStr(Game.lives)
End Sub

Public Sub ExtraLife()
 If Game.score \ 10000 > Game.extralives Then
  Game.extralives = Game.extralives + 1
  AddToLives (1)
 End If
End Sub
