Attribute VB_Name = "modsFireSkinParser"
'
' module containing skin parser function
'

'almost perfected

Option Explicit
Public plColor As Long
Public plCurrent As Long

Public SkinName As String
Public SkinAuthor As String
Public SkinNotes As String, titleBackColor As Long

Public Wx As Integer, Wy As Integer, cWx As Integer, cWy As Integer
Public isTitleColorFilled As Boolean, isAlbumColorFilled As Boolean
Private isVisBoxResized As Boolean

'the one and only function to layout a skin
Public Sub renderSkin(SkinName As String)
'Unload frmFireMain

On Error GoTo errHandle
'On Error Resume Next

Dim skinUnPacker As New FireSkinLibrary.FireSkinner, i As Integer

' check if temporary directory exists
If Fsys.FolderExists(App.path & "\Temp") Then
'yes? delete all files

Else
'no? well create one...
Fsys.CreateFolder App.path & "\Temp"
End If
 
' delete any previous files
If Dir(App.path & "\temp\*.*") <> "" Then Kill App.path & "\temp\*.*"
'unpack the skin

skinUnPacker.decodeFireSkin SkinName, App.path & "\temp\"

' we donot need the skin unpacker anymore
Set skinUnPacker = Nothing

isTitleColorFilled = False
isAlbumColorFilled = False
isVisBoxResized = False

'start to parse the files
Dim Fin As textStream
'locate a .fss file in the temp directory
If Fsys.FileExists(App.path & "\Temp\" & Dir(App.path & "\Temp\*.fss")) Then
  Set Fin = Fsys.OpenTextFile(App.path & "\Temp\" & getFileTitleFromPath(SkinName) & ".fss")
 Else
  Dim e As ErrStruct
  e.errNum = 7
  e.errShortDesc = "Corrupt skin! The default skin will be loaded"
  e.errLongDesc = "The skin specification was not found in the archive"
  logError e
 Exit Sub
End If


Dim Line$, windowRegion As Long, ln As Integer
Dim parts() As String, rgn() As String
'clear all

Dim Control As Object
For Each Control In frmFireMain
If TypeOf Control Is PictureBox Then Control.Cls
Next

plCurrent = -1

While Not Fin.AtEndOfStream
Line = Fin.readLine


' load files
If LCase(Line) = "files" Then
Fin.readLine

While Line <> "%>"

Line = Fin.readLine
  ln = ln + 1
  parts = Split(Line, ":")
  If UBound(parts) > 0 Then
   parts(1) = Trim(parts(1))
   Select Case parts(0)
   'load main picture
   Case "main":
   Set frmFireMain.picSkin.Picture = LoadPicture(App.path & "\temp\" & parts(1))
   frmFireMain.Height = frmFireMain.picSkin.Height
   frmFireMain.Width = frmFireMain.picSkin.Width
    windowRegion = MakeRegion(frmFireMain.picSkin)
     SetWindowRgn frmFireMain.hwnd, windowRegion, True
   'load playlist picture
   Case "playlist":
   Set frmFirePL.picSkin.Picture = LoadPicture(App.path & "\temp\" & parts(1))
   frmFirePL.Width = frmFirePL.picSkin.Width
   frmFirePL.Height = frmFirePL.picSkin.Height
    windowRegion = MakeRegion(frmFirePL.picSkin)
     SetWindowRgn frmFirePL.hwnd, windowRegion, True
   
   'load other pictures
   Case "buttons": frmFireMain.picBtnSrc.Picture = LoadPicture(App.path & "\temp\" & parts(1))
   Case "controls": frmFireMain.picCtrlSrc.Picture = LoadPicture(App.path & "\temp\" & parts(1))
   
   Case "seek-bar": frmFireMain.picBarBack.Picture = LoadPicture(App.path & "\temp\" & parts(1))
   Case "pl-bar": frmFirePL.picBack.Picture = LoadPicture(App.path & "\temp\" & parts(1))
   Case "seek-bar-front": frmFireMain.picBarFront.Picture = LoadPicture(App.path & "\temp\" & parts(1))
   Case "pl-bar-front": frmFirePL.picBar.Picture = LoadPicture(App.path & "\temp\" & parts(1))
   
  
   End Select
  End If
  Wend

End If

'load data
If LCase(Line) = "data" Then

Fin.readLine
While Line <> "%>"
 Line = Fin.readLine

parts = Split(Line, ":")
  If UBound(parts) > 0 Then
     checkData parts(0), parts(1)
  End If
Wend

End If

' arrange elements
If LCase(Line) = "arrange" Then

Fin.readLine
While Line <> "%>"
 Line = Fin.readLine
 ln = ln + 1
parts = Split(Line, ":")
  If UBound(parts) > 0 Then
     Arrange parts(0), parts(1)
   
  End If
Wend
End If

' change fonts
If LCase(Line) = "fonts" Then

Fin.readLine
While Line <> "%>"
 Line = Fin.readLine
parts = Split(Line, ":")
  If UBound(parts) > 0 Then
   changeFont parts(0), parts(1)
  End If
Wend
End If

' change colors
If LCase(Line) = "colors" Then

Fin.readLine
While Line <> "%>"
 Line = Fin.readLine
parts = Split(Line, ":")
  If UBound(parts) > 0 Then
   changeColor parts(0), parts(1)
  End If
Wend
End If

Wend


With frmFireMain
.picBtnSrc.Refresh
.picCtrlSrc.Refresh

Wx = .picBtnSrc.ScaleWidth / 2
Wy = .picBtnSrc.ScaleHeight / 4

cWx = .picCtrlSrc.ScaleWidth / 2
cWy = .picCtrlSrc.ScaleHeight / 2

For i = 0 To 2
.picBtn(i).Width = Wx
.picBtn(i).Height = Wy
Next

For i = 0 To 1
.picCtrl(i).Width = cWx
.picCtrl(i).Height = cWy
Next

BitBlt .picBtn(1).hdc, 0, 0, Wx, Wy, .picSkin.hdc, .picBtn(1).Left, .picBtn(1).Top, vbSrcCopy
BitBlt .picBtn(0).hdc, 0, 0, Wx, Wy, .picSkin.hdc, .picBtn(0).Left, .picBtn(0).Top, vbSrcCopy
BitBlt .picBtn(2).hdc, 0, 0, Wx, Wy, .picSkin.hdc, .picBtn(2).Left, .picBtn(2).Top, vbSrcCopy

BitBlt .picCtrl(0).hdc, 0, 0, cWx, cWy, .picSkin.hdc, .picCtrl(0).Left, .picCtrl(0).Top, vbSrcCopy
BitBlt .picCtrl(1).hdc, 0, 0, cWx, cWy, .picSkin.hdc, .picCtrl(1).Left, .picCtrl(1).Top, vbSrcCopy

.picBtn(0).Refresh
.picBtn(1).Refresh
.picBtn(2).Refresh

.picCtrl(0).Refresh
.picCtrl(1).Refresh


TransparentBlt .picBtn(0).hdc, 0, 0, Wx, Wy, .picBtnSrc.hdc, 0, 0, Wx, Wy, vbGreen
TransparentBlt .picBtn(1).hdc, 0, 0, Wx, Wy, .picBtnSrc.hdc, 0, Wy, Wx, Wy, vbGreen 'stop
TransparentBlt .picBtn(2).hdc, 0, 0, Wx, Wy, .picBtnSrc.hdc, 0, Wy * 3, Wx, Wy, vbGreen 'open

TransparentBlt .picCtrl(0).hdc, 0, 0, cWx, cWy, .picCtrlSrc.hdc, 0, 0, cWx, cWy, vbGreen
TransparentBlt .picCtrl(1).hdc, 0, 0, cWx, cWy, .picCtrlSrc.hdc, 0, cWy, cWx, cWy, vbGreen

.picBtn(0).Refresh
.picBtn(1).Refresh
.picBtn(2).Refresh

.picCtrl(0).Refresh
.picCtrl(1).Refresh

BitBlt .picScroll.hdc, 0, 0, .picScroll.ScaleWidth, .picScroll.ScaleWidth, .picSkin.hdc, .picScroll.Left, .picScroll.Top, vbSrcCopy
.picScroll.Refresh


If Not isTitleColorFilled Then
BitBlt .picTitle.hdc, 0, 0, .picTitle.ScaleWidth, .picTitle.ScaleWidth, .picSkin.hdc, .picTitle.Left, .picTitle.Top, vbSrcCopy
.picTitle.Refresh

End If



If Not isAlbumColorFilled Then
BitBlt .picAlbum.hdc, 0, 0, .picAlbum.ScaleWidth, .picAlbum.ScaleWidth, .picSkin.hdc, .picAlbum.Left, .picAlbum.Top, vbSrcCopy
.picAlbum.Refresh
End If

 If plCurrent = -1 Then plCurrent = plColor

Dim Count As Integer
For Count = 1 To frmFirePL.lstPl.ListItems.Count
If frmFirePL.lstPl.ListItems.Item(Count).Bold Then
frmFirePL.lstPl.ListItems.Item(Count).ForeColor = plCurrent
Else
frmFirePL.lstPl.ListItems.Item(Count).ForeColor = plColor
End If

Next Count


End With

SaveSetting App.EXEName, "Settings", "Skin", SkinName


Set Fin = Nothing
Exit Sub

errHandle:
   
   e.errShortDesc = "Unexpected error in skin parser"
   e.errLongDesc = Err.Description
   e.errNum = Err.Number
    logError e
    Err.Clear
End Sub

Function parseRGB(RGBString As String) As Long

parseRGB = RGB(Val("&h" & Mid(RGBString, 1, 2)), Val("&h" & Mid(RGBString, 3, 2)), Val("&h" & Mid(RGBString, 5, 2)))
End Function

Sub checkData(Line$, Data$)
If Line Like "name" Then
SkinName = Data
ElseIf Line Like "author" Then
SkinAuthor = Data
ElseIf Line Like "notes" Then
SkinNotes = Data
End If
End Sub
Sub Arrange(Line$, Data$)
If Line Like "main-caption" Then
moveObject frmFireMain.lblCaption, Data

ElseIf Line Like "main-seek-bar" Then
moveObject frmFireMain.picBarBack, Data

ElseIf Line Like "main-play-button" Then
moveObject frmFireMain.picBtn(0), Data

ElseIf Line Like "main-stop-button" Then
moveObject frmFireMain.picBtn(1), Data

ElseIf Line Like "main-open-button" Then
moveObject frmFireMain.picBtn(2), Data

ElseIf Line Like "main-close-button" Then
moveObject frmFireMain.picCtrl(0), Data

ElseIf Line Like "main-min-button" Then
moveObject frmFireMain.picCtrl(1), Data

ElseIf Line Like "main-time" Then
moveObject frmFireMain.lblStatus, Data

ElseIf Line Like "main-info" Then
moveObject frmFireMain.picScroll, Data

ElseIf Line Like "pl-caption" Then
moveObject frmFirePL.lblCaption, Data

ElseIf Line Like "pl-list" Then
moveObject frmFirePL.lstPl, Data

ElseIf Line Like "pl-bar" Then
moveObject frmFirePL.picBack, Data

ElseIf Line Like "song-title" Then
moveObject frmFireMain.picTitle, Data
moveObject frmFireMain.picScrollBuffer, Data

ElseIf Line Like "song-album" Then
moveObject frmFireMain.picAlbum, Data

ElseIf Line Like "video" Then
moveObject frmFireMain.fraVideo, Data
If Not isVisBoxResized Then moveObject frmFireMain.fraDisplay, Data


ElseIf Line Like "vis-box" Then
moveObject frmFireMain.fraDisplay, Data
isVisBoxResized = True

ElseIf Line Like "vis" Then
 moveObject frmFireMain.picVisual, Data
End If
End Sub

Sub moveObject(theObject As Object, ByVal Data As String)
Dim Regions() As String
Regions = Split(Data, ",")
theObject.Move Val(Regions(0)), Val(Regions(1))
If Regions(2) <> "" Then theObject.Height = Val(Regions(2))
If Regions(3) <> "" Then theObject.Width = Val(Regions(3))
End Sub

Sub changeFont(Line$, Data$)
If Line Like "main-caption" Then
makeFontChange frmFireMain.lblCaption, Data

ElseIf Line Like "main-title" Then
makeFontChange frmFireMain.lblTitle, Data
makeFontChange frmFireMain.picTitle, Data
makeFontChange frmFireMain.picScrollBuffer, Data

ElseIf Line Like "main-album" Then
makeFontChange frmFireMain.lblAlbum, Data

ElseIf Line Like "main-time" Then
makeFontChange frmFireMain.lblStatus, Data

ElseIf Line Like "main-info" Then
makeFontChange frmFireMain.lblInfo, Data

ElseIf Line Like "pl-caption" Then
makeFontChange frmFirePL.lblCaption, Data

ElseIf Line Like "pl-list" Then
makeFontChange frmFirePL.lstPl, Data

End If
End Sub

Sub makeFontChange(theObject As Object, Data As String)
Dim Regions() As String
Regions = Split(Data, ",")

theObject.Font.Name = Regions(0)
theObject.Font.Size = Val(Regions(1))

Select Case LCase(Regions(2))
 Case "b"
 theObject.Font.Bold = True
 Case "i"
  theObject.Font.Italic = True
 Case "u"
  theObject.Font.Underline = True
 Case "s"
 theObject.Font.Strikethrough = True
  Case "n"
 theObject.Font.Bold = False
 theObject.Font.Italic = False
 theObject.Font.Underline = False
 theObject.Font.Strikethrough = False
 End Select
 
End Sub

Sub changeColor(Line$, Data$)
If Line Like "main-caption" Then
makeColorChange frmFireMain.lblCaption, Data

ElseIf Line Like "main-title" Then
makeColorChange frmFireMain.lblTitle, Data
makeColorChange frmFireMain.picTitle, Data
makeColorChange frmFireMain.picScrollBuffer, Data

ElseIf Line Like "main-album" Then
makeColorChange frmFireMain.lblAlbum, Data


ElseIf Line Like "main-time" Then
makeColorChange frmFireMain.lblStatus, Data

ElseIf Line Like "main-info" Then
makeColorChange frmFireMain.lblInfo, Data

ElseIf Line Like "main-album-back" Then
frmFireMain.picAlbum.BackColor = parseRGB(Data)
isAlbumColorFilled = True

ElseIf Line Like "main-title-back" Then
frmFireMain.picTitle.BackColor = parseRGB(Data)
titleBackColor = parseRGB(Data)
isTitleColorFilled = True

ElseIf Line Like "pl-caption" Then
makeColorChange frmFirePL.lblCaption, Data

ElseIf Line Like "pl-list" Then
makeColorChange frmFirePL.lstPl, Data
plColor = parseRGB(Data)

ElseIf Line Like "pl-current" Then plCurrent = parseRGB(Data)

ElseIf Line Like "pl-back-list" Then
frmFirePL.lstPl.BackColor = parseRGB(Data)

ElseIf Line Like "vis-box-fore" Then
makeColorChange frmFireMain.ScopeBuff, Data
End If
End Sub

Sub makeColorChange(theObject As Object, Data As String)
theObject.ForeColor = parseRGB(Data)
End Sub

Sub readSkinData(skinFile As String, ByRef SkinName As String, ByRef SkinAuthor As String, ByRef SkinNotes As String)
Dim Fin As textStream, readLine As String, parts() As String
Dim skinUnPacker As New FireSkinLibrary.FireSkinner, i As Integer
If Fsys.FolderExists(App.path & "\temp1") Then Fsys.DeleteFolder App.path & "\temp1"
Fsys.CreateFolder App.path & "\temp1"
skinUnPacker.decodeFireSkin skinFile, App.path & "\temp1\"
Set skinUnPacker = Nothing

Set Fin = Fsys.OpenTextFile((App.path & "\temp1\" & getFileTitleFromPath(skinFile) & ".fss"))


While Not Fin.AtEndOfStream
readLine = Fin.readLine
If LCase(readLine) = "data" Then
   While readLine <> "%>"
     readLine = Fin.readLine
     parts = Split(readLine, ":")
        If UBound(parts) > 0 Then
            If LCase(parts(0)) = "name" Then
                SkinName = parts(1)
            ElseIf LCase(parts(0)) = "author" Then
                SkinAuthor = parts(1)
            ElseIf LCase(parts(0)) = "notes" Then
                SkinNotes = parts(1)
            End If
        End If
    Wend
End If
Wend
Fin.Close
Set Fin = Nothing
Fsys.DeleteFolder App.path & "\temp1"
End Sub
