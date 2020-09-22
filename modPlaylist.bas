Attribute VB_Name = "modPlaylist"
Option Explicit

'
' module containing Playlist related sub-routines
'

Public Declare Function GetFullPathName Lib "kernel32" Alias "GetFullPathNameA" (ByVal lpFileName As String, ByVal nBufferLength As Long, ByVal lpBuffer As String, ByVal lpFilePart As String) As Long
Public doRepeatPlaylist As Boolean, doRandomPlayback As Boolean
Public lib As New Dictionary, invLib As New Dictionary

Public Sub savePlayList(lstPath As ListBox, file As String)
Dim OutStream As textStream, i As Integer
i = 0
If Trim(file) = "" Then GoTo e
Set OutStream = Fsys.CreateTextFile(file, True, False)

'fpl headers
OutStream.WriteLine ("<?fpl version=" & Chr(&H22) & "1.0" & Chr(&H22) & "?>") 'write XML header
OutStream.WriteLine ("<!-- Created on " & Date & "; WARNING !!! This File is Machine Generated. Do NOT Edit. -->") ' write date
OutStream.WriteLine ("<playlist generator=" & Chr(&H22) & "FireAMP"" " & "version=""" & App.Major & "." & App.Minor & "." & App.Revision & Chr(&H22) & ">") ' write main tag

While i < lstPath.ListCount
OutStream.WriteLine "    <path> " & lstPath.List(i) & " </path>" 'write path
OutStream.WriteLine "    <name> " & getFileTitleFromPath(lstPath.List(i)) & " </name>" 'write song name
i = i + 1
Wend

OutStream.WriteLine ("</playlist>")
Set OutStream = Nothing ' destroy object

e:
End Sub

' Loads playlist

Public Sub openPlayList(lstPlaylist As ListView, lstPath As ListBox, file As String)
Dim InStream As textStream

Dim a As Boolean, str As String
Dim LST As ListItem
Dim i As Integer, ext As String

Let a = False
If Trim(file) = "" Then GoTo e
Set InStream = Fsys.OpenTextFile(Trim(file), ForReading, False, TristateFalse)


If Not StrComp(Replace(InStream.Read(5), "<?", " "), "fpl") Then
 Dim e As ErrStruct
 e.errNum = 5
 e.errShortDesc = "This does not appear to be a FireAMP! Playlist"
 e.errLongDesc = "The playlist recently opened did not have the FireAMP! playlist header in it. The File is either corrupt or invalid"
 logError e
Exit Sub
End If

InStream.SkipLine ' Skip header
InStream.SkipLine ' Skip date
InStream.SkipLine ' Skip main tag

lstPath.Clear
lstPlaylist.ListItems.Clear

While InStream.AtEndOfStream = False
str = InStream.readLine
If a = True Then
Set LST = lstPlaylist.ListItems.Add(, , parseString(str, 7, 7))   ' load name
Else
lstPath.AddItem (parseString(str, 7, 7))   ' load path
End If
a = Not a

Wend
lstPath.RemoveItem (lstPath.ListCount - 1)
Set InStream = Nothing ' destroy object
e:
End Sub

Public Function parseString(Src As String, Start As Integer, Finish As Integer) As String
On Error Resume Next
Dim str As String, str1 As String
Src = Trim(Src)
str = Left(Src, Len(Src) - Start)
str1 = Right(str, Len(str) - Finish)
parseString = str1
End Function

Public Sub readM3UFile(M3UFile As String, lstPl As ListView, lstPaths As ListBox)

Dim Fsys As New Scripting.FileSystemObject
Dim Fin As textStream
Dim readLine As String, temp() As String, parts() As String, isPath As Boolean

Set Fin = Fsys.OpenTextFile(M3UFile, ForReading, False)
isPath = False

lstPaths.Clear
lstPl.ListItems.Clear

If Fin.readLine <> "#EXTM3U" Then

MsgBox "Not an m3u file"
  Fin.Close
  Set Fin = Nothing
  Set Fsys = Nothing

  Exit Sub
End If

While Not Fin.AtEndOfStream
readLine = Fin.readLine

If isPath Then
   If Mid(readLine, 1, 1) = "\" Then
    Dim path As String
    path = Space(255)
GetFullPathName readLine, Len(path), path, readLine
 lstPaths.AddItem path
Else
    lstPaths.AddItem readLine
End If

' not ispath
Else
 parts = Split(readLine, ":")
    If parts(0) = "#EXTINF" Then
       temp = Split(parts(1), ",")
           Dim LST As ListItem
   
        Set LST = lstPl.ListItems.Add(, , temp(1))
        LST.SubItems(1) = convertToStdTime(temp(0))
   
        Set LST = Nothing
     Else
      lstPl.ListItems.Add , , "Unknown"
     End If
End If

isPath = Not isPath
Wend

Fin.Close
Set Fin = Nothing
Set Fsys = Nothing

End Sub

'sub to read and load a PLS playlist
Public Sub loadPLS(PLSFile As String, lstPl As ListView, lstPaths As ListBox)
Dim Fsys As New Scripting.FileSystemObject
Dim Fin As textStream
Dim readLine As String, temp() As String, parts$()

lstPaths.Clear
lstPl.ListItems.Clear

Set Fin = Fsys.OpenTextFile(PLSFile, ForReading, False)

If Fin.readLine <> "[playlist]" Then
MsgBox "Not a PLS playlist"
Fin.Close
Set Fin = Nothing
Set Fsys = Nothing
Exit Sub
End If

While Not Fin.AtEndOfStream
readLine = Fin.readLine
parts() = Split(readLine, "=")

 If parts(0) Like "File*" Then
    Dim path As String
    path = Space(255)
        GetFullPathName parts(1), Len(path), path, readLine

        lstPaths.AddItem path
 ElseIf parts(0) Like "Title*" Then
         readLine = Fin.readLine
         temp() = Split(readLine, "=")
           Dim LST As ListItem
   
        Set LST = lstPl.ListItems.Add(, , parts(1))
        LST.SubItems(1) = convertToStdTime(temp(1))
 ElseIf parts(0) Like "NumberOfEntries" Or parts(0) Like "Version" Then
 End If
 Wend


Fin.Close
Set Fin = Nothing
Set Fsys = Nothing

End Sub

Sub toHTMLPlaylist()

Dim textStream As textStream
Set textStream = Fsys.CreateTextFile(App.path & "\Playlist.html")
Dim i As Integer
textStream.WriteLine "<html>"
textStream.WriteLine "<head><title>FireAMP! Playlist</title></head>"
textStream.WriteLine "<body style=""background:indianred;color:ghostwhite;font-family:Arial"">"
textStream.WriteLine "<h1 style=""border-bottom:5px double RGB(255,128,128)"">FireAMP! Playlist</h1>"
For i = 0 To frmFirePL.lstPaths.ListCount - 1
 textStream.WriteLine "<li>" & frmFirePL.lstPl.ListItems(i + 1).Text

Next
textStream.WriteLine "</body>"
textStream.WriteLine "</html>"
textStream.Close
Set textStream = Nothing
ShellExecute frmFireMain.hwnd, "open", App.path & "\Playlist.html", "", "", 0
End Sub
