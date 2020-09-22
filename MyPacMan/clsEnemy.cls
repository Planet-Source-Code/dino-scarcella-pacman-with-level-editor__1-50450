VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsEnemy"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit
Public InGame As Boolean ' if enemy is in the game, in other words, looking to kill Pacman
Public Eyesonly As Boolean ' if enemy is eyes only (true)
Public HitWall As Boolean ' has the enemy hitten the wall
Public xPos As Integer ' x position of the enemy
Public yPos As Integer ' y position of the enemy
Private OldXpos As Integer 'old x position of the enemy
Private OldYpos As Integer 'old y position of the enemy
Private OldpX As Integer
Private OldpY As Integer
Public Direction As Integer ' direction the enemy is currently moving
Private PathToTake() As Integer ' the predefined path that an enemy must take when in eyesonly mode
Public PathCounter As Integer ' holds where we currently are in the pathtotake() array
Public Bounced As Integer ' counts how many times enemy has bounced up and down in the box
Public EatenModeTimer As Integer ' how long the enemy stays in eaten mode for
Public Index As Integer 'which enemy is this
Public DelayTime As Integer ' this is to slow the enemy down in eaten mode by half
Public Speed As Integer ' speed the enemy moves at
Public Flash As Integer ' integer which determines if the enemy in eatenmode must flash or not when showing the sprite

Public Sub Move(ByRef Enemy() As clsEnemy, pacman As clsPacMan)
  Dim XD2 As Integer
  Dim YD2 As Integer
  Dim Xs As Integer
  Dim Ys As Integer
  Dim X As Long
  Dim Y As Long
  Dim pX As Long
  Dim pY As Long
  Dim back As Integer
  Dim counter As Integer
  
  ' Memorize the ghost position
  OldXpos = xPos
  OldYpos = yPos
  
  HitWall = False
  
  X = xPos \ 20 + 1 ' ghost x pos
  Y = yPos \ 20 + 1 ' ghost y pos
  pX = pacman.xPos \ 20 + 1 ' pacman x pos
  pY = pacman.yPos \ 20 + 1 ' pacman y pos
       
  If InGame = False Then
   Select Case yPos
    Case Is = 261
     Bounced = Bounced + 1
     Direction = 1
    Case Is = 321
     Bounced = Bounced + 1
     Direction = 0
   End Select
   If Bounced >= (Index * 3 - 3) Then InGame = True
  End If
  
  If InGame = True Then
   If (xPos Mod 20 = 1 And yPos Mod 20 = 1) Then
   
    'if the enemy is in the box and is not in there for renewal,get it out of the box
    If InBox(xPos, yPos) And Eyesonly = False Then
     Select Case xPos
      Case Is = 261
       Direction = 3
      Case Is = 321
       Direction = 2
      Case Is = 281
       'in case the box is blocked
       If Level(X, Y - 1).block <> "B" Then
        Direction = 0
       Else
        Direction = 3
       End If
      Case Is = 301
       'in case the box is blocked
       If Level(X, Y - 1).block <> "B" Then
        Direction = 0
       Else
        Direction = 2
       End If
     End Select
    End If
    
    'if the enemy is in the box and is there for renewal,if it's at the bottom of the box,return it to normal
    If InBox(xPos, yPos) And Eyesonly = True Then
     If yPos = 321 Then
      Direction = 0
      EatenModeTimer = 0
      Eyesonly = False
      Speed = 1
     End If
     GoTo incrementposition
    End If
    
    If InBox(xPos, yPos) = False Then
     
     'Check if Ghost has returned to the box for renewal after been eaten by Pacman
     If (xPos = 281 Or xPos = 301) And yPos = 221 And Eyesonly Then
      InGame = False  ' ghost is not in the game anymore
      Direction = 1   ' head down into the box for renewal
     Else
     
      If Abs(pX - X) + Abs(pY - Y) > 1 + Index Then ' every pacman looks a bit further to where pacman is with the increase of Index
        pX = pX + XD(pacman.Direction) * Index * 2
        pY = pY + YD(pacman.Direction) * Index * 2
        pX = Split(ClosestBlockTo(pX, pY))(0)
        pY = Split(ClosestBlockTo(pX, pY))(1)
      End If
      
      ' if ghost is eyesonly then make it head towards the ghostbox
      ' instead of pacman
      If Eyesonly = True Then
repeat_random:
       pX = 15 + Rnd
       pY = 12
       'fixes the problem if one of these is a block
       If Level(pX, pY).block = "B" Then GoTo repeat_random
       'checks if it's the first time that enemy is eyesonly and flush in block
       If PathCounter = 0 Then
        Set Path = bAStar(StageMap, X, Y, pX, pY, 0, 0, 31, 31)
        getPathToTake
        Speed = 2
       End If
       Direction = PathToTake(PathCounter)
       PathCounter = PathCounter + 1
       If PathCounter > UBound(PathToTake) Then PathCounter = -1
       GoTo incrementposition
      End If
      
      ' make ghost hitwall for now, so it needs to look to see if it can move
      ' in another direction first
      HitWall = True
      
      If EatenModeTimer = 0 Then
        If PathCounter <> -1 Then
         ' if we come across the current px,py or get closer to
         ' current px,py than old px,py while getting to old px,py
         ' then theres no more use of following the path, is there ?
         If (X = pX And Y = pY) Or (Abs(X - pX) + Abs(Y - pY) < Abs(X - OldpX) + Abs(X - OldpX)) Then
          PathCounter = -1
          GoTo simplersearch
         End If
         Direction = PathToTake(PathCounter): HitWall = False
         PathCounter = PathCounter + 1
         If PathCounter > UBound(PathToTake) Then PathCounter = -1
         GoTo incrementposition
        Else
         ' if we are not in the block we want to be in then
         If Not (X = pX And Y = pY) Then
          'hold the x and y coords where we want to go to so that we do
          'not lose them in the future
          OldpX = pX
          OldpY = pY
          Set Path = bAStar(StageMap, X, Y, pX, pY, 0, 0, 31, 31)
          getPathToTake
          'if no path is found do the simpler path searching
          If UBound(PathToTake) = 0 Then
           PathCounter = -1
           GoTo simplersearch
          End If
          'if the path needs pac to go in reverse forget it
          If UBound(PathToTake) < 10 Or PathToTake(PathCounter) = Rev(Direction) Then
           PathCounter = -1
           GoTo simplersearch
          End If
          Direction = PathToTake(PathCounter): HitWall = False
          PathCounter = PathCounter + 1
          If PathCounter > UBound(PathToTake) Then PathCounter = -1
          GoTo incrementposition
         End If
        End If
      End If
       
simplersearch:
      ' if ghost has come across a junction ( T , L or Crossroad intersection)
      If Level(X, Y).junction = True Then
       
       XD2 = Abs(pX - X) ' distance from pac or ghostbox
       YD2 = Abs(pY - Y)
       Xs = Sgn(pX - X) ' direction from pac or ghostbox
       Ys = Sgn(pY - Y)
       
       ' if ghost is in being eaten mode, then reverse it's dirs so
       ' that it runs away instead of chasing!
       If EatenModeTimer > 0 Then
        Xs = -Xs
        Ys = -Ys
       End If
            
       ' back is the opp dir of where the ghost is heading, it's there so that
       ' the ghost doesn't back track on itself
       back = Rev(Direction)
               
       ' first line of looking, basically if pac is above move up, etc.
       If Ys < 0 And Level(X, Y).up And back <> 0 Then Direction = 0: HitWall = False
       If Ys > 0 And Level(X, Y).down And back <> 1 Then Direction = 1: HitWall = False
       If Xs < 0 And Level(X, Y).left And back <> 2 Then Direction = 2: HitWall = False
       If Xs > 0 And Level(X, Y).right And back <> 3 Then Direction = 3: HitWall = False
          
       ' if a wall is where ghost wants to go, look to the sides
       If HitWall = True Then
        If Ys = 0 Then ' if pac is on the same Y as ghost
         If Level(X, Y).up And back <> 0 Then Direction = 0: HitWall = False
         If Level(X, Y).down And back <> 1 Then Direction = 1: HitWall = False
        End If
        If Xs = 0 Then ' if pac is on the same X as ghost
         If Level(X, Y).left And back <> 2 Then Direction = 2: HitWall = False
         If Level(X, Y).right And back <> 3 Then Direction = 3: HitWall = False
        End If
       End If
       ' if a wall is still in front of where the ghost wishes to go
       ' then look for any other way to go besides back again
       If HitWall = True Then
         If Level(X, Y).up And back <> 0 Then Direction = 0: HitWall = False
         If Level(X, Y).down And back <> 1 Then Direction = 1: HitWall = False
         If Level(X, Y).left And back <> 2 Then Direction = 2: HitWall = False
         If Level(X, Y).right And back <> 3 Then Direction = 3: HitWall = False
       End If
       
      End If
      
      If Level(X, Y).junction = False Then
       
       'if this is a straight or dead end,determine if it's a straight or dead end and then act according to this
       counter = 0
       If Level(X, Y).up Then counter = counter + 1
       If Level(X, Y).down Then counter = counter + 1
       If Level(X, Y).left Then counter = counter + 1
       If Level(X, Y).right Then counter = counter + 1
       'if dead end go backwards
       If counter = 1 Then
        Direction = Rev(Direction)
        HitWall = False
       End If
       'if a straight then carry on going in current direction
       If counter = 2 Then
        HitWall = False
       End If
      End If
     End If
    End If
   End If
  End If
  
incrementposition:
 If HitWall = False Then
  Select Case Direction
   Case 0
    yPos = yPos - Speed
   Case 1
    yPos = yPos + Speed
   Case 2
    xPos = xPos - Speed
   Case 3
    xPos = xPos + Speed
  End Select
  
  'enemy can move out of exit point and come back on the other side
  If xPos > 600 Then
   xPos = -19
  ElseIf xPos < -19 Then
   xPos = 600
  End If
  
  'decrement the amount of time enemy is in eatenmode
  If EatenModeTimer > 0 Then EatenModeTimer = EatenModeTimer - 1
 End If
End Sub

Public Sub Initialize(j As Integer)
  Index = j
  On Error Resume Next 'in case this object is already loaded
  Load frmGame.pctEnemyBuffer(Index) 'load a new picture box
  BitBlt frmGame.pctEnemyBuffer(Index).hdc, 0, 0, 19, 19, frmGame.pctBlank.hdc, 0, 0, vbSrcCopy 'start the box having a black block
  
  If Index = 1 Then
   InGame = True
  Else
   InGame = False
  End If
  
  xPos = 261 + ((Index - 1) Mod 4) * 20
  If Index Mod 2 = 0 Then
   yPos = 321
   Direction = 0
  Else
   yPos = 261
   Direction = 1
  End If
  
  Eyesonly = False
  Bounced = 0
  EatenModeTimer = 0
  PathCounter = -1
  Speed = 1
End Sub

Public Function pathFinding(DestinationX_simple_coordinate As Long, DestinationY_simple_coordinate As Long, CurrentX As Long, CurrentY As Long, Current_Direction As Integer) As Boolean
 Dim X As Integer
 Dim Y As Integer
 Dim Xs As Integer
 Dim Ys As Integer
 Dim back As Integer
 Dim HitWall2 As Boolean
 Dim direction2 As Integer
 Dim BlocksVisited() As String
 
 pathFinding = True
 
 X = CurrentX
 Y = CurrentY
 direction2 = Current_Direction
 
 'lets get the x and y co-ords of the block the enemy is going to first
 While (X Mod 20 <> 1 Or Y Mod 20 <> 1)
  If Current_Direction = 0 Then Y = Y - 1
  If Current_Direction = 1 Then Y = Y + 1
  If Current_Direction = 2 Then X = X - 1
  If Current_Direction = 3 Then X = X + 1
 Wend
 
 'simplify the co-ordinates
 X = X \ 20 + 1
 Y = Y \ 20 + 1
 
ReDim BlocksVisited(0)
ReDim PathToTake(0)
While ClosestBlockTo(DestinationX_simple_coordinate, DestinationY_simple_coordinate) <> BlocksVisited(UBound(BlocksVisited))
 ' direction from pac or ghostbox
 Xs = Sgn(DestinationX_simple_coordinate - X)
 Ys = Sgn(DestinationY_simple_coordinate - Y)
 ' make ghost hitwall for now, so it needs to look to see if it can move
 ' in another direction first
 HitWall2 = True
            
 ' back is the opp dir of where the ghost is heading, it's there so that
 ' the ghost doesn't back track on itself
 back = Rev(direction2)
               
 ' first line of looking, basically if pac is above move up, etc.
 If Ys < 0 And Level(X, Y).up And Not inArray(BlocksVisited, CStr(X) + " " + CStr(Y - 1)) And back <> 0 Then direction2 = 0: HitWall2 = False
 If Ys > 0 And Level(X, Y).down And Not inArray(BlocksVisited, CStr(X) + " " + CStr(Y + 1)) And back <> 1 Then direction2 = 1: HitWall2 = False
 If Xs < 0 And Level(X, Y).left And Not inArray(BlocksVisited, CStr(X - 1) + " " + CStr(Y)) And back <> 2 Then direction2 = 2: HitWall2 = False
 If Xs > 0 And Level(X, Y).right And Not inArray(BlocksVisited, CStr(X + 1) + " " + CStr(Y)) And back <> 3 Then direction2 = 3: HitWall2 = False
          
 ' if a wall is where ghost wants to go, look to the sides
 If HitWall2 = True Then
  If Ys = 0 Then ' if pac is on the same Y as ghost
   If Level(X, Y).up And Not inArray(BlocksVisited, CStr(X) + " " + CStr(Y - 1)) And back <> 0 Then direction2 = 0: HitWall2 = False
   If Level(X, Y).down And Not inArray(BlocksVisited, CStr(X) + " " + CStr(Y + 1)) And back <> 1 Then direction2 = 1: HitWall2 = False
  End If
  If Xs = 0 Then ' if pac is on the same X as ghost
   If Level(X, Y).left And Not inArray(BlocksVisited, CStr(X - 1) + " " + CStr(Y)) And back <> 2 Then direction2 = 2: HitWall2 = False
   If Level(X, Y).right And Not inArray(BlocksVisited, CStr(X + 1) + " " + CStr(Y)) And back <> 3 Then direction2 = 3: HitWall2 = False
  End If
 End If
 ' if a wall is still in front of where the ghost wishes to go
 ' then look for any other way to go besides back again
 If HitWall2 = True Then
  If Level(X, Y).up And Not inArray(BlocksVisited, CStr(X) + " " + CStr(Y - 1)) And back <> 0 Then direction2 = 0: HitWall2 = False
  If Level(X, Y).down And Not inArray(BlocksVisited, CStr(X) + " " + CStr(Y + 1)) And back <> 1 Then direction2 = 1: HitWall2 = False
  If Level(X, Y).left And Not inArray(BlocksVisited, CStr(X - 1) + " " + CStr(Y)) And back <> 2 Then direction2 = 2: HitWall2 = False
  If Level(X, Y).right And Not inArray(BlocksVisited, CStr(X + 1) + " " + CStr(Y)) And back <> 3 Then direction2 = 3: HitWall2 = False
 End If
 
 If HitWall2 = True Then
  pathFinding = False
  Exit Function
 End If
 
 If direction2 = 0 Then Y = Y - 1
 If direction2 = 1 Then Y = Y + 1
 If direction2 = 2 Then
  X = X - 1
  If X = -1 Then X = 30
 End If
 If direction2 = 3 Then
  X = X + 1
  If X = 31 Then X = 1
 End If
 
 ReDim Preserve PathToTake(UBound(PathToTake) + 1)
 ReDim Preserve BlocksVisited(UBound(BlocksVisited) + 1)
 PathToTake(UBound(PathToTake)) = direction2
 BlocksVisited(UBound(BlocksVisited)) = CStr(X) + " " + CStr(Y)
Wend
PathCounter = 1
End Function

Private Sub getPathToTake()
Dim CoOrd As CoOrdinate
Dim LastCoOrd As CoOrdinate

 ReDim PathToTake(0)
 PathCounter = 0
 If Not (Path Is Nothing) Then
    Set LastCoOrd = Nothing
    For Each CoOrd In Path
        If Not (LastCoOrd Is Nothing) Then
            PathCounter = PathCounter + 1
            ReDim Preserve PathToTake(PathCounter)
            If CoOrd.Y < LastCoOrd.Y Then PathToTake(PathCounter) = 0
            If CoOrd.Y > LastCoOrd.Y Then PathToTake(PathCounter) = 1
            If CoOrd.X < LastCoOrd.X Then PathToTake(PathCounter) = 2
            If CoOrd.X > LastCoOrd.X Then PathToTake(PathCounter) = 3
        End If
        Set LastCoOrd = CoOrd
    Next
 End If
 PathCounter = 1
End Sub

Private Function ClosestBlockTo(X As Long, Y As Long) As String
 Dim minimum_distance As Single
 Dim newX As Integer
 Dim newY As Integer
 Dim counteri As Integer
 Dim counterj As Integer
 
 If X < 1 Or X > 30 Or Y < 1 Or Y > 30 Or (X >= 13 And X <= 18 And Y >= 13 And Y <= 18) Then
  minimum_distance = 2000
  'first find x and y within range of stage
  For counteri = 2 To 29
   For counterj = 2 To 29
    If Not (counteri >= 13 And counteri <= 18 And counterj >= 13 And counterj <= 18) Then
     If Sqr(((X - counterj) ^ 2) + ((Y - counteri) ^ 2)) < minimum_distance Then
      minimum_distance = Sqr(((X - counterj) ^ 2) + ((Y - counteri) ^ 2))
      newX = counterj
      newY = counteri
     End If
    End If
   Next counterj
  Next counteri
 Else
  newX = X
  newY = Y
 End If
 
 If Level(newX, newY).block <> "B" Then
  ClosestBlockTo = CStr(newX) + " " + CStr(newY)
  Exit Function
 End If
 
 minimum_distance = 2000
 For counteri = 2 To 29
   For counterj = 2 To 29
    If Not (counteri >= 13 And counteri <= 18 And counterj >= 13 And counterj <= 18) Then
     If Level(counterj, counteri).block <> "B" Then
      If Sqr(((newX - counterj) ^ 2) + ((newY - counteri) ^ 2)) < minimum_distance Then
       minimum_distance = Sqr(((newX - counterj) ^ 2) + ((newY - counteri) ^ 2))
       ClosestBlockTo = CStr(counterj) + " " + CStr(counteri)
      End If
     End If
    End If
   Next counterj
  Next counteri
End Function

Private Function inArray(arr() As String, Key As String) As Boolean
 inArray = False
 For i = 0 To UBound(arr)
  If arr(i) = Key Then
   inArray = True
   Exit Function
  End If
 Next i
End Function

Private Function InBox(X As Integer, Y As Integer) 'is enemy in the box
 If X >= 261 And X <= 321 And Y >= 241 And Y <= 321 Then
  InBox = True
 Else
  InBox = False
 End If
End Function

Public Sub showEnemyBlit()
 'flash ranges from 0 to 15
 Flash = (Flash + 1) Mod 16
 'show the enemy sprites
 'if enemy is in normal mode
 If EatenModeTimer = 0 And Eyesonly = False Then
  frmGame.pctBase.PaintPicture frmGame.pctEnemy.Picture, xPos, yPos, 19, 19, 114, Direction * 19, 19, 19, vbSrcAnd
  frmGame.pctBase.PaintPicture frmGame.pctEnemy.Picture, xPos, yPos, 19, 19, (((Index - 1) Mod 4) * 19), Direction * 19, 19, 19, vbSrcPaint
 End If
 'if enemy is "edible"
 If EatenModeTimer > 0 And Eyesonly = False Then
  frmGame.pctBase.PaintPicture frmGame.pctEnemy.Picture, xPos, yPos, 19, 19, 114, Direction * 19, 19, 19, vbSrcAnd
  'if "edible" mode is ending, flash the enemy to warn player
  If EatenModeTimer <= 100 Then
   If Flash <= 7 Then
    frmGame.pctBase.PaintPicture frmGame.pctEnemy.Picture, xPos, yPos, 19, 19, 76, Direction * 19, 19, 19, vbSrcPaint
   Else
    frmGame.pctBase.PaintPicture frmGame.pctEnemy.Picture, xPos, yPos, 19, 19, (((Index - 1) Mod 4) * 19), Direction * 19, 19, 19, vbSrcPaint
   End If
   Exit Sub
  End If
  frmGame.pctBase.PaintPicture frmGame.pctEnemy.Picture, xPos, yPos, 19, 19, 76, Direction * 19, 19, 19, vbSrcPaint
 End If
 'if enemy has been eaten, make enemy use the eyesonly sprite
 If Eyesonly = True Then
  frmGame.pctBase.PaintPicture frmGame.pctEnemy.Picture, xPos, yPos, 19, 19, 133, Direction * 19, 19, 19, vbSrcAnd
  frmGame.pctBase.PaintPicture frmGame.pctEnemy.Picture, xPos, yPos, 19, 19, 95, Direction * 19, 19, 19, vbSrcPaint
 End If
End Sub

Public Sub hideEnemyBlit()
 'put here because when we close frmGame an error occurs saying that pctEnemyBuffer(1) does not exist
 'this occurs because every element after 0 is loaded in code and is not put on form (done in initialize method for enemies)
 On Error Resume Next
 'restore the background from under the sprite
  BitBlt frmGame.pctBase.hdc, OldXpos, OldYpos, 19, 19, frmGame.pctEnemyBuffer(Index).hdc, 0, 0, vbSrcCopy
End Sub

Public Sub bufferEnemyBackground()
 'put here because when we close frmGame an error occurs saying that pctEnemyBuffer(1) does not exist
 'this occurs because every element after 0 is loaded in code and is not put on form (done in initialize method for enemies)
 On Error Resume Next
 'copy info behind the sprite into a buffer
 BitBlt frmGame.pctEnemyBuffer(Index).hdc, 0, 0, 19, 19, frmGame.pctBase.hdc, xPos, yPos, vbSrcCopy
End Sub
