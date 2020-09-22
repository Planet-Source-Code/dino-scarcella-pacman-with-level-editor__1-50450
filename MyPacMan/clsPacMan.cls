VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsPacMan"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit
Public HitWall As Boolean
Public Dead As Boolean
Public Direction As Integer
Public OldXpos As Long
Public OldYpos As Long
Public xPos As Long
Public yPos As Long
Public Mouthstate As Integer
Public Pillsleft As Integer
Private increment As Integer

Public Sub Move()
 'if pacman is perfectly within a block, certain decisions need to be made
 If (xPos Mod 20 = 1 And yPos Mod 20 = 1) Then
  'we want to move up
  If GetAsyncKeyState(vbKeyUp) And (Level(xPos \ 20 + 1, yPos \ 20 + 1).up = True) Then
   HitWall = False
   Direction = 0
  End If
 
  'we want to move down
  If GetAsyncKeyState(vbKeyDown) And (Level(xPos \ 20 + 1, yPos \ 20 + 1).down = True) Then
   HitWall = False
   Direction = 1
  End If
 
  'we want to move left
  If GetAsyncKeyState(vbKeyLeft) And (Level(xPos \ 20 + 1, yPos \ 20 + 1).left = True) Then
   HitWall = False
   Direction = 2
  End If
 
  'we want to move right
  If GetAsyncKeyState(vbKeyRight) And (Level(xPos \ 20 + 1, yPos \ 20 + 1).right = True) Then
   HitWall = False
   Direction = 3
  End If
  
  'check if we've hit a wall
  If ((Direction = 0) And (Level(xPos \ 20 + 1, yPos \ 20 + 1).up = False)) Or ((Direction = 1) And (Level(xPos \ 20 + 1, yPos \ 20 + 1).down = False)) Or ((Direction = 2) And (Level(xPos \ 20 + 1, yPos \ 20 + 1).left = False)) Or ((Direction = 3) And (Level(xPos \ 20 + 1, yPos \ 20 + 1).right = False)) Then
   HitWall = True
   'if we've hit a wall,just check if he is dead, don't move his co-ordinates
   GoTo checkifpacmanisdead
  End If
 End If
 
 'if we've reached this point, we are sure we have not hit a wall
 HitWall = False
 
 'store old positions before we increment positions
 OldXpos = xPos
 OldYpos = yPos
 
 'increment xPos and Ypos according to direction
 Select Case Direction
  Case 0
   yPos = yPos - 1
  Case 1
   yPos = yPos + 1
  Case 2
   xPos = xPos - 1
  Case 3
   xPos = xPos + 1
 End Select
 
 'deal with exit points on left and right
 If xPos > 600 Then
  xPos = -19
 ElseIf xPos < -19 Then
  xPos = 600
 End If
 
 'change how wide Pac Man's mouth is
 Mouthstate = Mouthstate + increment
 If Mouthstate = 3 Then increment = -1
 If Mouthstate = 0 Then increment = 1
 
checkifpacmanisdead:
 'check if pacman is dead
 Call IsPacManDead
End Sub

Public Sub initializePacMan()
 Randomize
 Mouthstate = 1
 increment = 1
 
 Do
  xPos = (Rnd * 27) + 2
  yPos = (Rnd * 27) + 2
 Loop Until Level(xPos, yPos).block = " " And (xPos < 13 Or xPos > 18 And yPos < 13 Or yPos > 18)
 
 xPos = ((xPos - 1) * 20) + 1
 yPos = ((yPos - 1) * 20) + 1
 Direction = 3
 Dead = False
 HitWall = False 'hitwall may have been true causing pac not to be displayed
 BitBlt frmGame.pctBuffer.hdc, 0, 0, 19, 19, frmGame.pctBlank.hdc, 0, 0, vbSrcCopy
End Sub

Private Sub IsPacManDead()
 For i = 1 To UBound(Enemy)
  If Abs(Enemy(i).xPos - xPos) < 9 And Abs(Enemy(i).yPos - yPos) < 9 And Enemy(i).Eyesonly = False Then
   'make pacman dead
   If Enemy(i).EatenModeTimer = 0 Then
    'this means pacman has died, in this event make dead equal true and play the relevent sound. I put a delay so that the player can see where he died and to let the sound play before we restart the game or go on to the next life
    Dead = True
    sndPlay "killed", SoundOps.SND_ASYNC
    delay (1200)
   End If
   'if pacman has hit enemy and enemy is "eatable", make enemy eyes only
   If Enemy(i).EatenModeTimer > 0 Then
    Enemy(i).Eyesonly = True
    Enemy(i).PathCounter = 0 ' gives the signal to start planning route to get back to box
    AddToScore (200)
    Call ExtraLife
    sndPlay "ghosteat", SoundOps.SND_ASYNC
   End If
  End If
 Next i
End Sub
Public Sub showPacManBlit()
If HitWall = False Then
 frmGame.pctBase.PaintPicture frmGame.pctPac.Picture, xPos, yPos, 19, 19, (76 + (Direction * 19)), (Mouthstate * 19), 19, 19, vbSrcAnd
 frmGame.pctBase.PaintPicture frmGame.pctPac.Picture, xPos, yPos, 19, 19, (Direction * 19), (Mouthstate * 19), 19, 19, vbSrcPaint
End If
End Sub

Public Sub hidePacManBlit()
 ' restore the background from under the sprite
 If HitWall = False Then
  BitBlt frmGame.pctBase.hdc, OldXpos, OldYpos, 19, 19, frmGame.pctBuffer.hdc, 0, 0, vbSrcCopy
 End If
 
 ' check if we've eaten a pill or not,if we have, hide the pill by painting over the buffer with a white block,decrement pillsleft and change the level array
 If (OldXpos Mod 20 = 1 And OldYpos Mod 20 = 1) And (Level(OldXpos \ 20 + 1, OldYpos \ 20 + 1).block = "O" Or Level(OldXpos \ 20 + 1, OldYpos \ 20 + 1).block = "o") And (HitWall = False) Then
  ' check if it's a superpill and make the enemy "edible"
  If (Level(OldXpos \ 20 + 1, OldYpos \ 20 + 1).block = "O") Then
   For i = 1 To UBound(Enemy)
    ' if they were following a path ignore this, because when they
    ' become "not edible" they will follow the rest of the old path
    If Enemy(i).Eyesonly = False Then Enemy(i).PathCounter = -1
    Enemy(i).EatenModeTimer = 1000
   Next i
   sndPlay "fruiteat", SoundOps.SND_ASYNC
  'if its a small pill make a different sound to a super pill
  Else
   sndPlay "eatpill", SoundOps.SND_ASYNC
  End If
  ' pacman has eaten a pill so we need to decrement pillsleft
  pacman.Pillsleft = pacman.Pillsleft - 1
  ' add to the score and then check if we need to add lives to PacMan
  AddToScore (10)
  Call ExtraLife
  Level(xPos \ 20 + 1, yPos \ 20 + 1).block = " "
  BitBlt frmGame.pctBase.hdc, OldXpos, OldYpos, 19, 19, frmGame.pctBlank.hdc, 0, 0, vbSrcCopy
 End If
 
End Sub

Public Sub bufferPacManBackground()
 'copy info behind the sprite into a buffer
 If HitWall = False Then
  BitBlt frmGame.pctBuffer.hdc, 0, 0, 19, 19, frmGame.pctBase.hdc, xPos, yPos, vbSrcCopy
 End If
End Sub

