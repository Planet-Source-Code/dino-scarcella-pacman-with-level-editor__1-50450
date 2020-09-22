VERSION 5.00
Begin VB.Form frmGame 
   AutoRedraw      =   -1  'True
   BackColor       =   &H00000000&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Pac Man Game"
   ClientHeight    =   9510
   ClientLeft      =   3075
   ClientTop       =   705
   ClientWidth     =   9030
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   634
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   602
   Begin VB.PictureBox pctLives 
      AutoSize        =   -1  'True
      BackColor       =   &H00000000&
      BorderStyle     =   0  'None
      Height          =   390
      Left            =   7875
      Picture         =   "frmGame.frx":0000
      ScaleHeight     =   390
      ScaleWidth      =   375
      TabIndex        =   10
      Top             =   9060
      Width           =   375
   End
   Begin VB.PictureBox pctEnemyBuffer 
      AutoRedraw      =   -1  'True
      Height          =   345
      Index           =   0
      Left            =   5760
      ScaleHeight     =   19
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   19
      TabIndex        =   8
      Top             =   9135
      Visible         =   0   'False
      Width           =   345
   End
   Begin VB.PictureBox pctEnemy 
      AutoRedraw      =   -1  'True
      AutoSize        =   -1  'True
      BorderStyle     =   0  'None
      Height          =   1140
      Left            =   2850
      Picture         =   "frmGame.frx":07FA
      ScaleHeight     =   76
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   152
      TabIndex        =   7
      Top             =   9105
      Visible         =   0   'False
      Width           =   2280
   End
   Begin VB.PictureBox pctBlank 
      AutoSize        =   -1  'True
      Height          =   345
      Left            =   7125
      Picture         =   "frmGame.frx":8F9C
      ScaleHeight     =   19
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   19
      TabIndex        =   6
      Top             =   9135
      Visible         =   0   'False
      Width           =   345
   End
   Begin VB.PictureBox pctSuperpill 
      AutoSize        =   -1  'True
      Height          =   345
      Left            =   6435
      Picture         =   "frmGame.frx":9452
      ScaleHeight     =   19
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   19
      TabIndex        =   5
      Top             =   9135
      Visible         =   0   'False
      Width           =   345
   End
   Begin VB.PictureBox pctPill 
      AutoSize        =   -1  'True
      Height          =   345
      Left            =   6105
      Picture         =   "frmGame.frx":9908
      ScaleHeight     =   19
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   19
      TabIndex        =   4
      Top             =   9135
      Visible         =   0   'False
      Width           =   345
   End
   Begin VB.PictureBox pctBlock 
      AutoSize        =   -1  'True
      Height          =   345
      Left            =   6780
      Picture         =   "frmGame.frx":9DBE
      ScaleHeight     =   19
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   19
      TabIndex        =   3
      Top             =   9135
      Visible         =   0   'False
      Width           =   345
   End
   Begin VB.PictureBox pctBase 
      AutoRedraw      =   -1  'True
      AutoSize        =   -1  'True
      BorderStyle     =   0  'None
      Height          =   9015
      Left            =   0
      Picture         =   "frmGame.frx":A274
      ScaleHeight     =   601
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   601
      TabIndex        =   2
      Top             =   0
      Width           =   9015
      Begin VB.Label lblReady 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "READY!"
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
         Left            =   3930
         TabIndex        =   12
         Top             =   4335
         Width           =   1155
      End
   End
   Begin VB.PictureBox pctPac 
      AutoRedraw      =   -1  'True
      AutoSize        =   -1  'True
      BorderStyle     =   0  'None
      Height          =   1140
      Left            =   3090
      Picture         =   "frmGame.frx":112DE2
      ScaleHeight     =   76
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   152
      TabIndex        =   1
      Top             =   9120
      Visible         =   0   'False
      Width           =   2280
   End
   Begin VB.PictureBox pctBuffer 
      AutoRedraw      =   -1  'True
      Height          =   345
      Left            =   5415
      Picture         =   "frmGame.frx":11B584
      ScaleHeight     =   19
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   19
      TabIndex        =   0
      Top             =   9135
      Visible         =   0   'False
      Width           =   345
   End
   Begin VB.Label lblScore 
      BackColor       =   &H00000000&
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   14.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   345
      Left            =   75
      TabIndex        =   11
      Top             =   9090
      Width           =   2670
   End
   Begin VB.Label lblLives 
      Alignment       =   1  'Right Justify
      BackColor       =   &H00000000&
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   14.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   345
      Left            =   7995
      TabIndex        =   9
      Top             =   9090
      Width           =   960
   End
End
Attribute VB_Name = "frmGame"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub Form_Unload(Cancel As Integer)
 End
End Sub
