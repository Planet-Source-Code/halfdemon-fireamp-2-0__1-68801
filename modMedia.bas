Attribute VB_Name = "modMedia"
Option Explicit

'
' module containing media related functions
'

Dim FireAMP As QuartzTypeLib.FilgraphManager       ' FireAMP!
Public FireAMP_Pos As QuartzTypeLib.IMediaPosition ' player position
Public FireAMP_Vol As QuartzTypeLib.IBasicAudio    ' player volume
Public FireAMP_VideoWin As QuartzTypeLib.IVideoWindow ' video window
Public currentVolume As Long

Public isPlaying As Boolean
Public curFile As String
Public playingIndex As Integer

Public Function PlayClip(fileToPlay As String, Optional testFile As Boolean = False) As Boolean
 Dim e As ErrStruct

On Error GoTo errHandle
 Set FireAMP = New FilgraphManager
  
 FireAMP.RenderFile (fileToPlay)
 
 
 
If Not testFile Then
 Set FireAMP_Pos = FireAMP
 Set FireAMP_VideoWin = FireAMP
 Set FireAMP_Vol = FireAMP
 

 frmFireMain.fraVideo.Visible = False
  frmFireMain.picVisual.Visible = False
  
  
      
  
 Select Case LCase(getFileExtensionFromPath(fileToPlay))
 Case "mpg", "mpeg", "dat", "mov", "wmv", "rm", "rmvb", "mov", "avi"
 If frmFullScreen.Visible Then
 FireAMP_VideoWin.HideCursor True
 refreshVideo frmFullScreen.Frame1
Else
FireAMP_VideoWin.HideCursor False
   refreshVideo frmFireMain.fraVideo
  frmFireMain.fraVideo.Visible = True
End If
 Case Else
  frmFireMain.picVisual.Visible = True
 End Select
FireAMP.Run
isPlaying = True
End If

PlayClip = True

Exit Function

errHandle:
  PlayClip = False
If Not testFile Then
 e.errNum = Err.Number
 e.errShortDesc = "FireAMP media error"
 e.errLongDesc = Err.Description
 
  logError e
  End If
  
Err.Clear
End Function

Public Sub StopClip()
If FireAMP Is Nothing Then Exit Sub
 FireAMP.Stop
 Set FireAMP = Nothing ' release object
 Set FireAMP_Pos = Nothing
 Set FireAMP_VideoWin = Nothing
 Set FireAMP_Vol = Nothing
isPlaying = False
End Sub

Public Sub PauseClip()
  FireAMP.Pause
End Sub

Public Sub ResumeClip()
  FireAMP.Run
End Sub
Public Sub refreshVideo(objOwner As Frame)

FireAMP_VideoWin.Width = objOwner.Width
FireAMP_VideoWin.Height = objOwner.Height

FireAMP_VideoWin.Left = 0
FireAMP_VideoWin.Top = 0

FireAMP_VideoWin.WindowStyle = CLng(&H6000000)  ' window style: no border
FireAMP_VideoWin.Owner = objOwner.hwnd          ' assign window region

End Sub
