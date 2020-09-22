VERSION 5.00
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "comdlg32.ocx"
Begin VB.Form frmMenu 
   BackColor       =   &H00000000&
   Caption         =   "PacMan Menu"
   ClientHeight    =   4260
   ClientLeft      =   5685
   ClientTop       =   2265
   ClientWidth     =   4440
   LinkTopic       =   "Form1"
   ScaleHeight     =   4260
   ScaleWidth      =   4440
   Begin VB.TextBox lblLevel 
      BackColor       =   &H00000000&
      BorderStyle     =   0  'None
      ForeColor       =   &H0000FFFF&
      Height          =   315
      Left            =   270
      Locked          =   -1  'True
      TabIndex        =   11
      Text            =   "Level :"
      Top             =   2985
      Width           =   3885
   End
   Begin VB.PictureBox pctEnemyBig 
      AutoRedraw      =   -1  'True
      AutoSize        =   -1  'True
      Height          =   1740
      Left            =   2280
      Picture         =   "frmMenu.frx":0000
      ScaleHeight     =   1680
      ScaleWidth      =   2100
      TabIndex        =   8
      Top             =   3345
      Visible         =   0   'False
      Width           =   2160
   End
   Begin VB.PictureBox pctAnim 
      AutoRedraw      =   -1  'True
      BorderStyle     =   0  'None
      Height          =   2580
      Left            =   0
      Picture         =   "frmMenu.frx":B802
      ScaleHeight     =   172
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   300
      TabIndex        =   7
      Top             =   15
      Width           =   4500
      Begin VB.Label lblPacman 
         BackStyle       =   0  'Transparent
         Caption         =   "PACMAN"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   24
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H0000FFFF&
         Height          =   525
         Left            =   1200
         TabIndex        =   10
         Top             =   1080
         Visible         =   0   'False
         Width           =   2100
      End
      Begin VB.Label lblNickname 
         Alignment       =   2  'Center
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "BLINKY"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   15.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H0000FFFF&
         Height          =   360
         Left            =   1695
         TabIndex        =   9
         Top             =   930
         Visible         =   0   'False
         Width           =   1110
      End
   End
   Begin VB.PictureBox pctBigPac 
      AutoRedraw      =   -1  'True
      AutoSize        =   -1  'True
      Height          =   1620
      Left            =   2280
      Picture         =   "frmMenu.frx":37764
      ScaleHeight     =   1560
      ScaleWidth      =   780
      TabIndex        =   6
      Top             =   3345
      Visible         =   0   'False
      Width           =   840
   End
   Begin VB.PictureBox pctBigBlank 
      AutoRedraw      =   -1  'True
      AutoSize        =   -1  'True
      Height          =   480
      Left            =   3975
      Picture         =   "frmMenu.frx":3B706
      ScaleHeight     =   420
      ScaleWidth      =   420
      TabIndex        =   5
      Top             =   3345
      Visible         =   0   'False
      Width           =   480
   End
   Begin VB.CommandButton cmdExit 
      DownPicture     =   "frmMenu.frx":3C078
      Height          =   870
      Left            =   2955
      Picture         =   "frmMenu.frx":3FBA2
      Style           =   1  'Graphical
      TabIndex        =   3
      Top             =   3330
      Width           =   1320
   End
   Begin VB.CommandButton cmdChangeLevel 
      DownPicture     =   "frmMenu.frx":436CC
      Height          =   870
      Left            =   1515
      Picture         =   "frmMenu.frx":471F6
      Style           =   1  'Graphical
      TabIndex        =   2
      Top             =   3330
      Width           =   1320
   End
   Begin VB.CommandButton cmdStart 
      DownPicture     =   "frmMenu.frx":4AD20
      Height          =   870
      Left            =   90
      Picture         =   "frmMenu.frx":4E84A
      Style           =   1  'Graphical
      TabIndex        =   1
      Top             =   3330
      Width           =   1320
   End
   Begin VB.ComboBox cmbEnemies 
      Appearance      =   0  'Flat
      BackColor       =   &H00000000&
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H0000FFFF&
      Height          =   330
      ItemData        =   "frmMenu.frx":52374
      Left            =   2730
      List            =   "frmMenu.frx":52390
      Style           =   2  'Dropdown List
      TabIndex        =   0
      Top             =   2610
      Width           =   1410
   End
   Begin MSComDlg.CommonDialog dlgOpen 
      Left            =   3960
      Top             =   3450
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
   End
   Begin VB.Label Label1 
      BackColor       =   &H00000000&
      Caption         =   "Amount of enemies"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H0000FFFF&
      Height          =   285
      Left            =   285
      TabIndex        =   4
      Top             =   2655
      Width           =   1410
   End
End
Attribute VB_Name = "frmMenu"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private AnimOn As Boolean

Private Sub cmdChangeLevel_Click()
 dlgOpen.ShowOpen
 If dlgOpen.filename <> "" Then
  lblLevel.Text = "Level : " + dlgOpen.filename
  lblLevel.ToolTipText = "Level : " + dlgOpen.filename
 End If
End Sub

Private Sub cmdExit_Click()
 Unload Me
 Unload frmGame
End Sub

Private Sub cmdStart_Click()
 pctAnim.Cls
 frmMenu.Hide
 frmGame.Show
 'initialize the score,extralives and lives
 Game.lives = 3
 Game.extralives = 0
 Game.score = 0
 'we use these methods just to display for now
 AddToLives (0)
 AddToScore (0)
 sndPlay "startmusic", SoundOps.SND_ASYNC
 'redimension array according to choice made in combobox
 ReDim Enemy(1 To CInt(cmbEnemies.List(cmbEnemies.ListIndex)))
 
startnewlevel:
 LoadLevel (Mid(lblLevel.Text, 9, Len(lblLevel.Text) - 8))
 
startnewlife:
 'initialize pacman and all enemies involved
 pacman.initializePacMan
 
 For i = 1 To UBound(Enemy)
  Enemy(i).Initialize (i)
 Next i
 
 'show initial positions of sprites on screen
 RefreshLevelandSprites
 
 'play the ready intro
 playIntro
 
 'game loop
 While pacman.Pillsleft > 0 And pacman.Dead = False
  delay (1)
  
  'move co-ordinates to appropriate positions
  pacman.Move
  For i = 1 To UBound(Enemy)
   Enemy(i).Move Enemy, pacman
  Next i
  
  'hide all sprites involved
  For i = 1 To UBound(Enemy)
   Enemy(i).hideEnemyBlit
  Next i
  pacman.hidePacManBlit
  
  'buffer all sprites involved
  pacman.bufferPacManBackground
  For i = 1 To UBound(Enemy)
   Enemy(i).bufferEnemyBackground
  Next i
  
  'show all sprites involved
  pacman.showPacManBlit
  For i = 1 To UBound(Enemy)
   Enemy(i).showEnemyBlit
  Next i
 Wend
 
 If pacman.Dead = True And Game.lives > 0 Then
  AddToLives (-1)
  GoTo startnewlife
 ElseIf pacman.Pillsleft <= 0 Then
  GoTo startnewlevel
 Else
  frmGame.Hide
  frmMenu.Show
 End If
 
End Sub

Private Sub Form_Load()
 lblLevel.Text = "Level : " + App.Path + "\maps\level1.pmm"
 lblLevel.ToolTipText = "Level : " + App.Path + "\maps\level1.pmm"
 dlgOpen.Filter = "PacMan Map files (*.pmm) | *.pmm"
 cmbEnemies.ListIndex = 0
 
  ' Memorize the Directions
  YD(0) = -1
  YD(1) = 1
  XD(2) = -1
  XD(3) = 1
  
  ' Memorize the Reverse Direction to the directions
  Rev(0) = 1
  Rev(1) = 0
  Rev(2) = 3
  Rev(3) = 2
End Sub

Private Sub Form_Activate()
 Dim pacstate As Integer
 Dim n As Integer
 Dim m As Integer
 Dim i As Integer
 Dim ind As Integer
 AnimOn = True
 
While AnimOn = True
 sndPlay "interm", SoundOps.SND_ASYNC
 For n = -47 To 300
  showBigPac n, 100, pacstate
  showSmallEnemy n + 26, 107, 4, 3
  delay (1)
  hideBig n, 100
  hideSmall n + 26, 107
  pacstate = pacstate + 1
  If pacstate = 4 Then pacstate = 0
 Next n
 
 sndPlay "interm", SoundOps.SND_ASYNC
 For n = 346 To -47 Step -1
  showBigEnemy n, 98, 0, 2
  showSmallPac n - 20, 107, pacstate
  delay (1)
  hideBig n, 98
  hideSmall n - 20, 107
  pacstate = pacstate + 1
  If pacstate = 4 Then pacstate = 0
 Next n
 
 For m = 0 To 3
  If m Mod 2 = 0 Then
   For n = -28 To 136
    showBigEnemy n, 98, m, 3
    delay (1)
    hideBig n, 98
   Next n
  End If
  If m Mod 2 = 1 Then
   For n = 328 To 136 Step -1
    showBigEnemy n, 98, m, 2
    delay (1)
    hideBig n, 98
   Next n
  End If
  
  Select Case m
   Case 0
    lblNickname.ForeColor = RGB(248, 123, 0)
    lblNickname.Caption = "BLINKY"
   Case 1
    lblNickname.ForeColor = RGB(250, 0, 0)
    lblNickname.Caption = "PINKY"
   Case 2
    lblNickname.ForeColor = RGB(0, 254, 254)
    lblNickname.Caption = "INKY"
   Case 3
    lblNickname.ForeColor = RGB(250, 187, 218)
    lblNickname.Caption = "CLYDE"
  End Select
  
  lblNickname.Visible = True
  delay (4000)
  lblNickname.Visible = False

  If m Mod 2 = 1 Then
   For n = 137 To 328
    showBigEnemy n, 98, m, 3
    delay (1)
    hideBig n, 98
   Next n
  End If
  If m Mod 2 = 0 Then
   For n = 135 To -28 Step -1
    showBigEnemy n, 98, m, 2
    delay (1)
    hideBig n, 98
   Next n
  End If
 Next m
 
 For n = 1 To 10
  lblPacman.Visible = True
  delay (500)
  lblPacman.Visible = False
  delay (500)
 Next n
Wend
End Sub

Private Sub Form_Deactivate()
 AnimOn = False
End Sub

Public Sub hideBig(X As Integer, Y As Integer)
 BitBlt pctAnim.hdc, X, Y, 28, 28, pctBigBlank.hdc, 0, 0, vbSrcCopy
End Sub

Public Sub hideSmall(X As Integer, Y As Integer)
 BitBlt pctAnim.hdc, X, Y, 19, 19, frmGame.pctBuffer.hdc, 0, 0, vbSrcCopy
End Sub

Public Sub showSmallEnemy(X As Integer, Y As Integer, Index As Integer, Direction As Integer)
 pctAnim.PaintPicture frmGame.pctEnemy.Picture, X, Y, 19, 19, 114, Direction * 19, 19, 19, vbSrcAnd
 pctAnim.PaintPicture frmGame.pctEnemy.Picture, X, Y, 19, 19, Index * 19, Direction * 19, 19, 19, vbSrcPaint
End Sub

Public Sub showBigEnemy(X As Integer, Y As Integer, Index As Integer, Direction As Integer)
 pctAnim.PaintPicture pctEnemyBig.Picture, X, Y, 28, 28, 112, Direction * 28, 28, 28, vbSrcAnd
 pctAnim.PaintPicture pctEnemyBig.Picture, X, Y, 28, 28, Index * 28, Direction * 28, 28, 28, vbSrcPaint
End Sub

Public Sub showBigPac(X As Integer, Y As Integer, state As Integer)
 pctAnim.PaintPicture pctBigPac.Picture, X, Y, 26, 26, 26, state * 26, 26, 26, vbSrcAnd
 pctAnim.PaintPicture pctBigPac.Picture, X, Y, 26, 26, 0, state * 26, 26, 26, vbSrcPaint
End Sub

Public Sub showSmallPac(X As Integer, Y As Integer, state As Integer)
 pctAnim.PaintPicture frmGame.pctPac.Picture, X, Y, 19, 19, 114, state * 19, 19, 19, vbSrcAnd
 pctAnim.PaintPicture frmGame.pctPac.Picture, X, Y, 19, 19, 38, state * 19, 19, 19, vbSrcPaint
End Sub

Private Sub Form_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
 cmdStart.Picture = LoadPicture(App.Path + "/images/startgame1.bmp")
 cmdChangeLevel.Picture = LoadPicture(App.Path + "/images/changelevel1.bmp")
 cmdExit.Picture = LoadPicture(App.Path + "/images/exit1.bmp")
End Sub

Private Sub Label1_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
 cmdStart.Picture = LoadPicture(App.Path + "/images/startgame1.bmp")
 cmdChangeLevel.Picture = LoadPicture(App.Path + "/images/changelevel1.bmp")
 cmdExit.Picture = LoadPicture(App.Path + "/images/exit1.bmp")
End Sub

Private Sub lblLevel_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
 cmdStart.Picture = LoadPicture(App.Path + "/images/startgame1.bmp")
 cmdChangeLevel.Picture = LoadPicture(App.Path + "/images/changelevel1.bmp")
 cmdExit.Picture = LoadPicture(App.Path + "/images/exit1.bmp")
End Sub

Private Sub pctAnim_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
 cmdStart.Picture = LoadPicture(App.Path + "/images/startgame1.bmp")
 cmdChangeLevel.Picture = LoadPicture(App.Path + "/images/changelevel1.bmp")
 cmdExit.Picture = LoadPicture(App.Path + "/images/exit1.bmp")
End Sub

Private Sub cmdStart_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
 cmdStart.Picture = LoadPicture(App.Path + "/images/startgame2.bmp")
 cmdChangeLevel.Picture = LoadPicture(App.Path + "/images/changelevel1.bmp")
 cmdExit.Picture = LoadPicture(App.Path + "/images/exit1.bmp")
End Sub

Private Sub cmdExit_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
 cmdStart.Picture = LoadPicture(App.Path + "/images/startgame1.bmp")
 cmdChangeLevel.Picture = LoadPicture(App.Path + "/images/changelevel1.bmp")
 cmdExit.Picture = LoadPicture(App.Path + "/images/exit2.bmp")
End Sub

Private Sub cmdChangeLevel_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
 cmdChangeLevel.Picture = LoadPicture(App.Path + "/images/changelevel2.bmp")
 cmdStart.Picture = LoadPicture(App.Path + "/images/startgame1.bmp")
 cmdExit.Picture = LoadPicture(App.Path + "/images/exit1.bmp")
End Sub
