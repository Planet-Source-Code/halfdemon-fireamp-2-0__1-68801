Attribute VB_Name = "modCommon"
Option Explicit

'
' module containing commonly used functions and subroutines
'

' last updated: 2006 May 06

Global Fsys As New FileSystemObject         'global FileSystemObject

Public Declare Sub Sleep Lib "kernel32" (ByVal dwMilliseconds As Long)

Declare Function SendMessage Lib "user32" Alias "SendMessageA" _
        (ByVal hwnd As Long, ByVal wMsg As Long, ByVal wParam As Long, _
        lParam As Any) As Long

Public Declare Function TextOut Lib "gdi32" Alias "TextOutA" (ByVal hdc As Long, ByVal X As Long, ByVal y As Long, ByVal lpString As String, ByVal nCount As Long) As Long
Public Declare Function ShellExecute Lib "shell32" Alias "ShellExecuteA" (ByVal hwnd As Long, ByVal lpOperation As String, ByVal lpFile As String, ByVal lpParameters As String, ByVal lpDirectory As String, ByVal nShowCmd As Long) As Long
Public Declare Function SetCapture Lib "user32" (ByVal hwnd As Long) As Long
Public Declare Function GetDC Lib "user32" (ByVal hwnd As Long) As Long
Public Declare Function InitCommonControls Lib "comctl32" () As Long

Public Type RECT
  Left As Long
  Top As Long
  Right As Long
  Bottom As Long
End Type
Public Declare Function SystemParametersInfo Lib "user32" Alias "SystemParametersInfoA" (ByVal uAction As Long, ByVal uParam As Long, lpvParam As Any, ByVal fuWinIni As Long) As Long

Public Const SPI_GETWORKAREA = 48
Const CB_FINDSTRING = &H14C
Const CB_FINDSTRINGEXACT = &H158
Const LB_FINDSTRING = &H18F
Const LB_FINDSTRINGEXACT = &H1A2

Public infoStr As String, Pos As Single, infoStr1 As String
Public InfoParts() As String, currentPart As Integer

Public sSize As Single, tSize As Single
'flag to stop scanning for media
Public stopScan As Boolean

' type to encapsulate errors
Public Type ErrStruct
  errNum As Long
  errShortDesc As String
  errLongDesc As String
End Type


Public Type FireAMPoptions
' general
enableVisualizations As Byte

'start up
showSplashScreen As Byte
loadDefaultSkin As Byte
checkAssociationsAtStartUp As Byte
defaultSepChar As Byte
' file types
MIDI As Byte
WAV As Byte
MP3 As Byte
MPG As Byte
WMA As Byte

End Type
Public theOptions As FireAMPoptions
Public sepChar As Byte, selectedIndex As Integer

Public sleepFactor As Byte
Public oldPath As String

Public Const HTTPGET = "GET {PATH} HTTP/1.1" & vbCrLf & _
                       "Host: {HOST}" & vbCrLf & _
                       "Accept: */*" & vbCrLf & _
                       "User-Agent: FireAMP Update System" & vbCrLf & _
                       vbCrLf
Public Const HTTPGETRANGE = "GET {PATH} HTTP/1.1" & vbCrLf & _
                            "Range: bytes={RANGE}" & vbCrLf & _
                            "Host: {HOST}" & vbCrLf & _
                            "Accept: */*" & vbCrLf & _
                            "User-Agent: FireAMP Update System" & vbCrLf & _
                            vbCrLf


' subroutine to log errors
Public Sub logError(theError As ErrStruct)
 
Dim Fout As textStream
 With frmFireTrap
  .lblError = theError.errShortDesc
  .lblReason = theError.errLongDesc
  .lblNum = "Error #" & theError.errNum
 End With
frmFireTrap.Show vbModal

' log error to file
 Set Fout = Fsys.OpenTextFile(App.path & "\FireAMP.Errors.Log", ForAppending, True)
 If Fsys.GetFile(App.path & "\FireAMP.Errors.Log").Size > 10& * 1024& Then ' greater than 10kb
  Fout.Close
  Kill App.path & "\FireAMP.Errors.Log"
  Set Fout = Fsys.OpenTextFile(App.path & "\FireAMP.Errors.Log", ForAppending, True)
 End If
Fout.WriteLine "FireAMP error #" & theError.errNum
Fout.WriteLine "Error occured on: " & Now
Fout.WriteLine "Short Desc: " & theError.errShortDesc
Fout.WriteLine "Long Desc: " & theError.errLongDesc
Fout.WriteLine String(40, "-")
Fout.Close

End Sub

' function to check if the given char is valid or not
Private Function isAllowedChar(testStr As String) As Boolean
 isAllowedChar = InStr(1, "ABCDEFGHIJKLMNOPQRSTUVWXYZ abcdefghijklmnopqrstuvwxyz1234567890(){}[]!;:'"",.$%*+#|\/~&", testStr)
End Function

' function to remove unwnated chars
Public Function toStdString(theString As String) As String
Dim retStr As String, i As Integer, j As Integer
retStr = Space(Len(theString)) ' fill up return string
Let j = 1
For i = 1 To Len(theString)

 If isAllowedChar(Mid(theString, i, 1)) Then
  ' mid is more faster and efficient than '&'
  Mid(retStr, j, 1) = Mid(theString, i, 1)
  j = j + 1
 End If
Next i

toStdString = Trim(retStr)
 
End Function

' bunch of useful functions
Public Function getDirNameFromPath(dirPath As String) As String
    
    getDirNameFromPath = Right(dirPath, Len(dirPath) - InStrRev(dirPath, "\"))
End Function

Public Function getFileTitleFromPath(FilePath As String) As String

Dim temp As String
temp = Mid(FilePath, Len(Left(FilePath, InStrRev(FilePath, "\"))) + 1)
getFileTitleFromPath = Left(temp, InStrRev(temp, ".") - 1)
End Function

Public Function getFileExtensionFromPath(FilePath As String) As String

getFileExtensionFromPath = Trim(LCase(Right(FilePath, Len(FilePath) - InStrRev(FilePath, "."))))
End Function

Public Function getFolderTitleFromPath(folderPath As String) As String
  
getFolderTitleFromPath = Mid(folderPath, InStrRev(folderPath, "\") + 1)
End Function

' function to convert seconds to HH:MM:SS format
Public Function convertToStdTime(ByVal Seconds As Long) As String
On Error Resume Next
'Format input value to "00:00:00"
Dim HH As Long                   'Hours
Dim MM As Long                   'Minutes
Dim SS As Long                   'Seconds
Dim Tmp As String                'Temporary value

 'Old values time is made of
 HH = Seconds \ 3600
 MM = Seconds \ 60 Mod 60
 SS = Seconds Mod 60
 
 'If there is hour
 If HH > 0 Then Tmp = Format$(HH, "00:")
 'Format input
 convertToStdTime = Tmp & Format$(MM, "00:") & Format$(SS, "00")
End Function

' function to translate the bar position into a position value
Public Function getBarPosition(picBar As PictureBox, picBarBack As PictureBox, iMax As Single) As Single
getBarPosition = ((picBar.Left) * iMax) / (picBarBack.ScaleWidth - picBar.ScaleWidth)

End Function

'updates the seek bar
Public Sub updateBar(picBar As PictureBox, picBarBack As PictureBox, ByVal iMax As Double, ByVal iPos As Double)
On Error Resume Next
    'Set the position bar to player position
    picBar.Move ((picBarBack.ScaleWidth - picBar.ScaleWidth) * ((iPos) / iMax)) ' bar position depends on the maximum value
    picBarBack.CurrentX = picBarBack.ScaleWidth / 2
    picBarBack.CurrentY = 1
    picBar.CurrentX = picBar.ScaleWidth / 2
    picBar.CurrentY = 1
DoEvents
End Sub


Sub Main()

Dim f As Integer
On Error GoTo e
f = FreeFile
If Fsys.FileExists(App.path & "\FireAMP.Options") Then
Open App.path & "\FireAMP.Options" For Binary Access Read As f
Get #f, , theOptions
Close #f
Else
InitCommonControls
Load frmFirstRun
  frmFirstRun.Show
Exit Sub
End If
sepChar = theOptions.defaultSepChar
If sepChar = 0 Then sepChar = CByte(Asc("/"))
DoEvents
frmSplash.Show
DoEvents

If App.PrevInstance Then End

If Command$ <> "" Then
Select Case getFileExtensionFromPath(Command$)
Case "fpl"
 openPlayList frmFirePL.lstPl, frmFirePL.lstPaths, Command$
Case "mp3", "wma", "wmv", "mid", "rmi", "mpg", "mpeg", "mpe"

 curFile = Replace(Command, """", "")
 frmFirePL.lstPl.ListItems.Add , , getFileTitleFromPath(Command)
 frmFirePL.lstPaths.AddItem Command$
 frmFireMain.picBtn_Click (0)

End Select
End If

' load visualization settings
frmSplash.lblStatus.Caption = "Loading Visualization Settings ..."
Dim t As textStream
If Fsys.FileExists(App.path & "\Sodas.list") Then
Set t = Fsys.OpenTextFile(App.path & "\Sodas.list", ForReading, True)
While Not t.AtEndOfStream
 modPi.importXPIList t.readLine
Wend
t.Close
Set t = Nothing
End If
frmFireMain.Show
Exit Sub
e:
MsgBox "error in Sub Main()"
End
End Sub

' general purpose sub to scan a path for media files, can display progress
' and status in picture boxes and labels

Sub scanFolder(FolderSpec As String, lstPaths As ListBox, lstPl As ListView, Optional lblBar As PictureBox = Nothing, Optional lblBarBack As PictureBox = Nothing, Optional lblStatus As Label = Nothing, Optional lblData As Label = Nothing)
On Error GoTo e
DoEvents
Dim i As Integer

Dim thisFolder As Folder
Dim sFolders As Folders
Dim fileItem As file, folderItem As Folder
Dim allFiles As Files

Set thisFolder = Fsys.GetFolder(FolderSpec)
Set sFolders = thisFolder.SubFolders
Set allFiles = thisFolder.Files

If stopScan Then Exit Sub

For Each folderItem In sFolders
DoEvents
If Not lblData Is Nothing Then lblData.Caption = "Looking in:" & vbNewLine & vbNewLine & folderItem.path
scanFolder folderItem.path, lstPaths, lstPl, lblBar, lblBarBack, lblStatus, lblData

Next

For Each fileItem In allFiles
sSize = sSize + fileItem.Size
If isMediaFile(fileItem.path) Then
lstPaths.AddItem fileItem.path
lstPl.ListItems.Add , , getFileTitleFromPath(fileItem.path)

 End If
Next
DoEvents
If Not lblBar Is Nothing And Not lblBarBack Is Nothing Then updateBar lblBar, lblBarBack, tSize, sSize
    If Not lblStatus Is Nothing Then lblStatus.Caption = "Scanned " & Round(sSize / (1024! * 1024!)) & "MB of " & Round(tSize / (1024! * 1024!)) & "MB so far."
  Exit Sub
e:

End Sub

' function used by ScanFolder to determine if the file is a media file
Public Function isMediaFile(FilePath As String) As Boolean
Dim ext As String

ext = getFileExtensionFromPath(FilePath)
isMediaFile = CBool(InStr("mp3 wma wmv mpg mpeg mpe rm rmvb mid rmi avi mov", ext))
End Function

Public Function isVideoFile(FilePath As String)
Dim ext As String

ext = getFileExtensionFromPath(FilePath)
isVideoFile = CBool(InStr("wmv mpg mpeg mpe rm rmvb avi dat mov", ext))
End Function

