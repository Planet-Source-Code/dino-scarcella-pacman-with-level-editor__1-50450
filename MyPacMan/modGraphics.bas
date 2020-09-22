Attribute VB_Name = "modGraphics"
Option Explicit

Public Sub RefreshLevelandSprites()
 frmGame.pctBase.Cls
 
 For i = 2 To 29
   For j = 2 To 29
    If Not (i >= 13 And i <= 18 And j >= 13 And j <= 18) Then
     Select Case Level(j, i).block
      Case "B"
       frmGame.pctBase.PaintPicture frmGame.pctBlock.Picture, (j - 1) * 20 + 1, (i - 1) * 20 + 1, 19, 19
      Case "o"
       frmGame.pctBase.PaintPicture frmGame.pctPill.Picture, (j - 1) * 20 + 1, (i - 1) * 20 + 1, 19, 19
      Case "O"
       frmGame.pctBase.PaintPicture frmGame.pctSuperpill.Picture, (j - 1) * 20 + 1, (i - 1) * 20 + 1, 19, 19
      Case " "
       frmGame.pctBase.PaintPicture frmGame.pctBlank.Picture, (j - 1) * 20 + 1, (i - 1) * 20 + 1, 19, 19
     End Select
    End If
   Next j
  Next i
 
  pacman.bufferPacManBackground
  pacman.showPacManBlit
  
  For i = 1 To UBound(Enemy)
   Enemy(i).bufferEnemyBackground
  Next i
  For i = 1 To UBound(Enemy)
   Enemy(i).showEnemyBlit
  Next i
End Sub

Public Sub playIntro() 'plays the ready! sign

For i = 1 To 4
 frmGame.lblReady.Visible = True
 delay (500)
 frmGame.lblReady.Visible = False
 delay (500)
Next i

End Sub
