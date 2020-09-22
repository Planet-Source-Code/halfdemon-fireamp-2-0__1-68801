VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "FireSkinner"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
' FireSkin class file

Option Explicit

Public Type FireSkin

FireSkinHeader As String      ' FireSkin Header: FireSkin
fileNames() As String         ' array of filenames
Locations() As String         ' locations of file in archive
nFiles As Integer             ' total number of files in archive

End Type


Private theFireSkin As FireSkin ' global FireSkin object


Private Sub setUpArchive(fileNameArray() As String) ' set up the FireSkin structure
On Error Resume Next
theFireSkin.nFiles = UBound(fileNameArray)

theFireSkin.FireSkinHeader = "FireSkin"

ReDim theFireSkin.fileNames(0 To theFireSkin.nFiles)
ReDim theFireSkin.Locations(0 To theFireSkin.nFiles)
Dim i As Integer
 For i = 0 To theFireSkin.nFiles
  theFireSkin.fileNames(i) = fileNameArray(i)
  theFireSkin.Locations(i) = FileLen(fileNameArray(i))
  
 Next i

End Sub

'
' sub routine to "make" a FireSkin
'

Public Sub makeFireSkin(fileNameArray() As String, DestArchive As String)


 Dim f As Integer, i As Integer, size As Integer, g As Integer
 Dim strFiles As String
 size = 0
 f = FreeFile
   setUpArchive fileNameArray
On Error Resume Next

 Dim buff() As Byte

Open DestArchive For Binary Access Read Write Lock Write As #f
Put #f, , theFireSkin
  For i = 0 To theFireSkin.nFiles
 ReDim buff(0 To FileLen(theFireSkin.fileNames(i)))
 g = FreeFile
 Open theFireSkin.fileNames(i) For Binary Access Read As #g
 Get #g, , buff
 Close #g
 
 Put #f, , buff
 Next i
  
   
Close #f
 Exit Sub

 
End Sub

'
' sub routine to "decode" a FireSkin
'

Public Sub decodeFireSkin(FireSkinArchive As String, destArchivePath As String)
On Error Resume Next

Dim f As Integer, g As Integer
f = FreeFile
Dim buff() As Byte
Dim m As FireSkin

Open FireSkinArchive For Binary Access Read As #f

Get #f, , m

If m.FireSkinHeader <> "FireSkin" Then
 MsgBox "Could not read FireSkin header! Corrupt file", vbOKOnly + vbInformation, "Invalid FireSkin"
Close #f
Exit Sub
End If
Dim i

For i = 0 To UBound(m.fileNames)
g = FreeFile
Open destArchivePath & getFileTitle(m.fileNames(i)) For Binary Access Read Write Lock Write As #g
 ReDim buff(0 To m.Locations(i)) As Byte
 Get #f, , buff
 Put #g, , buff
Close #g
Next
Close #f
Exit Sub

End Sub


Private Function getFileTitle(FilePath As String)
getFileTitle = Right(FilePath, Len(FilePath) - InStrRev(FilePath, "\"))
End Function

Public Function getFireSkinVersion() As String
getFireSkinVersion = App.Major & "." & App.Minor & "." & App.Revision
End Function



