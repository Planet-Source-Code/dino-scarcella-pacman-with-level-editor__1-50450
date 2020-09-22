Attribute VB_Name = "modLevel"
Option Explicit

Public Sub LoadLevel(filename As String)
 Dim ju As Integer
 
 pacman.Pillsleft = 0
 CreateBaseLevel
 Open filename For Random Access Read As #1 Len = LenB(Level(0, 0).block)
 
  For i = 1 To 30
   For j = 1 To 30
    If Not (i >= 13 And i <= 18 And j >= 13 And j <= 18) Then
     Get #1, (((i - 1) * 30) + j), Level(j, i).block
     If Level(j, i).block = "O" Or Level(j, i).block = "o" Then pacman.Pillsleft = pacman.Pillsleft + 1
    End If
   Next j
  Next i
  
  Close #1
   
  'load whether there is a junction, if top is free,bottom is free,left is free or right is free
  For j = 1 To 30
    For i = 1 To 30
    
      ju = 0
      'initialize moves to left,right,up and down as false (this is if we load a new level)
      Level(i, j).up = False
      Level(i, j).down = False
      Level(i, j).left = False
      Level(i, j).right = False
      
      With Level(i, j)
        
        If .block <> "B" Then
         If Level(i, j - 1).block <> "B" Then ju = ju Or 1: .up = True
         If Level(i, j + 1).block <> "B" Then ju = ju Or 2: .down = True
         If Level(i - 1, j).block <> "B" Then ju = ju Or 4: .left = True
         If Level(i + 1, j).block <> "B" Then ju = ju Or 8: .right = True
          
         If ju < 5 Or ju = 8 Or ju = 12 Then
            .junction = False ' delete straights & dead ends
         Else
            .junction = True
         End If
         
        End If
      End With
    Next
  Next
  
  'load level into StageMap for AStar Search
  Set StageMap = Nothing
  Set StageMap = New Map
  Set Path = Nothing
  For i = 0 To 31
   For j = 0 To 31
    If Level(i, j).block = "B" Then
     StageMap.Add "R" & i & "C" & j, i, j, 10, 0, "R" & i & "C" & j
    Else
     StageMap.Add "R" & i & "C" & j, i, j, 1, 1, "R" & i & "C" & j
    End If
   Next j
  Next i
End Sub

Private Sub CreateBaseLevel()
 For i = 1 To 30
  Level(i, 1).block = "B"
  Level(i, 30).block = "B"
  If (i <= 14) Or (i >= 17) Then
   Level(1, i).block = "B"
   Level(30, i).block = "B"
  End If
 Next i
 
 For i = 13 To 18
  Level(13, i).block = "B"
  Level(18, i).block = "B"
  If (i <= 14) Or (i >= 17) Then
   Level(i, 13).block = "B"
  End If
  Level(i, 18).block = "B"
 Next i
End Sub
