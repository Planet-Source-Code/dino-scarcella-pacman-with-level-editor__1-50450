Attribute VB_Name = "modSound"
Option Explicit

Public Sub sndPlay(strName As String, sndType As Long)

 ' procedure to play a sound
 sndPlaySound App.Path & "/sfx/" & strName & ".wav", sndType

End Sub
