VERSION 5.00
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "comdlg32.ocx"
Begin VB.Form frmLevelEditor 
   AutoRedraw      =   -1  'True
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Level Editor"
   ClientHeight    =   9015
   ClientLeft      =   3075
   ClientTop       =   990
   ClientWidth     =   9450
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   PaletteMode     =   1  'UseZOrder
   ScaleHeight     =   601
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   630
   Begin VB.PictureBox pctBase 
      AutoRedraw      =   -1  'True
      AutoSize        =   -1  'True
      BorderStyle     =   0  'None
      Height          =   9015
      Left            =   0
      Picture         =   "frmLevelEditor.frx":0000
      ScaleHeight     =   601
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   601
      TabIndex        =   5
      Top             =   0
      Width           =   9015
   End
   Begin VB.PictureBox pctGroup 
      Appearance      =   0  'Flat
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   1680
      Left            =   9045
      ScaleHeight     =   1680
      ScaleWidth      =   405
      TabIndex        =   0
      Top             =   0
      Width           =   405
      Begin VB.OptionButton optPill 
         Height          =   390
         Left            =   15
         Picture         =   "frmLevelEditor.frx":108B6E
         Style           =   1  'Graphical
         TabIndex        =   4
         Top             =   420
         Width           =   390
      End
      Begin VB.OptionButton optSuperpill 
         Height          =   390
         Left            =   15
         Picture         =   "frmLevelEditor.frx":109024
         Style           =   1  'Graphical
         TabIndex        =   3
         Top             =   840
         Width           =   390
      End
      Begin VB.OptionButton optBlank 
         Height          =   390
         Left            =   15
         Picture         =   "frmLevelEditor.frx":1094DA
         Style           =   1  'Graphical
         TabIndex        =   2
         Top             =   1260
         Width           =   390
      End
      Begin VB.OptionButton optBlock 
         Height          =   390
         Left            =   15
         Picture         =   "frmLevelEditor.frx":109990
         Style           =   1  'Graphical
         TabIndex        =   1
         Top             =   0
         Width           =   390
      End
   End
   Begin MSComDlg.CommonDialog dlgFile 
      Left            =   9000
      Top             =   1665
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
   End
   Begin VB.Menu mnuFile 
      Caption         =   "File"
      Begin VB.Menu mnuNew 
         Caption         =   "New"
      End
      Begin VB.Menu mnuOpen 
         Caption         =   "Open"
      End
      Begin VB.Menu mnuSave 
         Caption         =   "Save"
      End
      Begin VB.Menu mnuExit 
         Caption         =   "Exit"
      End
   End
End
Attribute VB_Name = "frmLevelEditor"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Dim defaultStage(30, 30) As String * 1
Dim stage(30, 30) As String * 1
Public filename As String
Public i, j As Integer

Private Sub Form_Load()
 Call ResetArrays
 Call ResetDisplay
 filename = "Untitled"
 frmLevelEditor.Caption = "PacMan Level Editor - Untitled"
 dlgFile.Filter = "PacMan Map files (*.pmm) | *.pmm"
End Sub

Private Sub pctBase_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
 If Button = vbLeftButton And Me.Tag <> "" Then
  Call InsertIntoArrayAndDisplay(X, Y)
 End If
End Sub

Private Sub pctBase_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
 If Button = vbLeftButton And Me.Tag <> "" Then
  Call InsertIntoArrayAndDisplay(X, Y)
 End If
End Sub

Private Sub mnuExit_Click()
 If Not arraysequal Then
  i = MsgBox("Do you want to save your work before exiting?", vbYesNo, "SAVE")
  If i = vbYes Then
   Call saveFile
   If Not IsPillPresent Then Exit Sub
  End If
 End If
 End
End Sub

Private Sub mnuNew_Click()
 
 If Not arraysequal Then
  i = MsgBox("Do you want to save your work before opening a new file?", vbYesNo, "SAVE")
  
  If i = vbYes Then saveFile
 End If
 
 Call ResetDisplay
 Call ResetArrays
 filename = "Untitled"
 frmLevelEditor.Caption = "PacMan Level Editor - " + filename
End Sub

Private Sub mnuOpen_Click()
 dlgFile.ShowOpen
 
 If (dlgFile.FileTitle <> "") Then
  filename = dlgFile.filename
  Call ResetDisplay
  Open filename For Random Access Read As #1 Len = LenB(stage(0, 0))
   
  For i = 1 To 28
   For j = 1 To 28
    If Not (i >= 12 And i <= 17 And j >= 12 And j <= 17) Then
     Get #1, ((i * 30) + (j + 1)), stage(j, i)
     defaultStage(j, i) = stage(j, i)
     Select Case stage(j, i)
      Case "B"
       pctBase.PaintPicture optBlock.Picture, j * 20 + 1, i * 20 + 1, 19, 19, 0, 0, 19, 19, vbSrcCopy
      Case "o"
       pctBase.PaintPicture optPill.Picture, j * 20 + 1, i * 20 + 1, 19, 19, 0, 0, 19, 19, vbSrcCopy
      Case "O"
       pctBase.PaintPicture optSuperpill.Picture, j * 20 + 1, i * 20 + 1, 19, 19, 0, 0, 19, 19, vbSrcCopy
      Case " "
       pctBase.PaintPicture optBlank.Picture, j * 20 + 1, i * 20 + 1, 19, 19, 0, 0, 19, 19, vbSrcPaint
     End Select
    End If
   Next j
  Next i
   
  Close #1
 End If
 frmLevelEditor.Caption = "PacMan Level Editor - " + filename
End Sub

Private Sub mnuSave_Click()
 Call saveFile
End Sub

Public Sub InsertIntoArrayAndDisplay(X As Single, Y As Single)
 i = X Mod 20
 j = Y Mod 20
 'check on border lines and ignore anything on borders
 If i = 0 Or j = 0 Then Exit Sub
 
 i = X \ 20
 j = Y \ 20
 'check if block chosen is part of cage
 If (i = 0) Or (j = 0) Or (i = 29) Or (j = 29) Or (i >= 12 And i <= 17 And j >= 12 And j <= 17) Then
  MsgBox "This block is not editable.", vbInformation, "Editing Error"
  Exit Sub
 End If
 
 Select Case Me.Tag
  Case "B"
   pctBase.PaintPicture optBlank.Picture, i * 20 + 1, j * 20 + 1, 19, 19, 0, 0, 19, 19, vbSrcAnd
   pctBase.PaintPicture optBlock.Picture, i * 20 + 1, j * 20 + 1, 19, 19, 0, 0, 19, 19, vbSrcPaint
  Case "o"
   pctBase.PaintPicture optBlank.Picture, i * 20 + 1, j * 20 + 1, 19, 19, 0, 0, 19, 19, vbSrcAnd
   pctBase.PaintPicture optPill.Picture, i * 20 + 1, j * 20 + 1, 19, 19, 0, 0, 19, 19, vbSrcPaint
  Case "O"
   pctBase.PaintPicture optBlank.Picture, i * 20 + 1, j * 20 + 1, 19, 19, 0, 0, 19, 19, vbSrcAnd
   pctBase.PaintPicture optSuperpill.Picture, i * 20 + 1, j * 20 + 1, 19, 19, 0, 0, 19, 19, vbSrcPaint
  Case " "
   pctBase.PaintPicture optBlank.Picture, i * 20 + 1, j * 20 + 1, 19, 19, 0, 0, 19, 19, vbSrcAnd
 End Select
 stage(i, j) = Me.Tag
End Sub

Public Sub ResetArrays()
 'do border
 For i = 0 To 29
  defaultStage(i, 0) = "B"
  defaultStage(i, 29) = "B"
  If (i <= 13) Or (i >= 16) Then
   defaultStage(0, i) = "B"
   defaultStage(29, i) = "B"
  Else
   defaultStage(0, i) = " "
   defaultStage(29, i) = " "
  End If
 Next i
 
 'do border of box
 For i = 12 To 17
  defaultStage(12, i) = "B"
  defaultStage(17, i) = "B"
  If (i <= 13) Or (i >= 16) Then
   defaultStage(i, 12) = "B"
  Else
   defaultStage(i, 12) = " "
  End If
  defaultStage(i, 17) = "B"
 Next i
 
 'do inside box
 For i = 13 To 16
  For j = 13 To 16
   defaultStage(i, j) = " "
  Next j
 Next i
 
 'do everything else
 For i = 1 To 28
  For j = 1 To 28
   If (i <= 13 Or i >= 16) And (j <= 13 Or j >= 16) Then defaultStage(i, j) = " "
  Next j
 Next i
 
 For i = 0 To 29
  For j = 0 To 29
   stage(i, j) = defaultStage(i, j)
  Next j
 Next i
End Sub

Public Function arraysequal() As Boolean
 
 arraysequal = True
 For i = 0 To 29
  For j = 0 To 29
   If stage(j, i) <> defaultStage(j, i) Then
    arraysequal = False
    Exit Function
   End If
  Next j
 Next i
 
End Function

Public Function IsPillPresent() As Boolean 'we have to see if a pill is present in the stage, because a pacman stage without pills has no meaning (there's no objective to the game)
 
 IsPillPresent = False
 For i = 0 To 29
  For j = 0 To 29
   If (stage(i, j) = "O") Or (stage(i, j) = "o") Then
    IsPillPresent = True
    Exit Function
   End If
  Next j
 Next i
 
End Function

Private Sub optBlank_Click()
 Me.Tag = " "
End Sub

Private Sub optBlock_Click()
 Me.Tag = "B"
End Sub

Private Sub optPill_Click()
 Me.Tag = "o"
End Sub

Private Sub optSuperpill_Click()
 Me.Tag = "O"
End Sub

Private Sub ResetDisplay()
 pctBase.Cls
End Sub

Private Sub saveFile()
 Dim counter As Long
 
    If Not IsPillPresent Then
     MsgBox "One can not make a PacMan stage without pills, make sure a pill is included in your design.", vbExclamation, "Stage Saving Error"
     Exit Sub
    End If
    
    If (filename <> "Untitled") Then
     Open filename For Random Access Write As #1 Len = LenB(stage(0, 0))
   
     For i = 0 To 29
      For j = 0 To 29
       counter = counter + 1
       Put #1, counter, stage(j, i)
      Next j
     Next i
   
     Close #1
    End If
 
    If filename = "Untitled" Then
     dlgFile.ShowSave
  
     If (dlgFile.FileTitle <> "") Then
      filename = dlgFile.filename
      Open filename For Random Access Write As #1 Len = LenB(stage(0, 0))
   
      For i = 0 To 29
       For j = 0 To 29
        counter = counter + 1
        Put #1, counter, stage(j, i)
       Next j
      Next i
   
      Close #1
     End If
    End If
    
    frmLevelEditor.Caption = "PacMan Level Editor - " + filename
End Sub
