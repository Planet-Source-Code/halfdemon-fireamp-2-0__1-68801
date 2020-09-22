Attribute VB_Name = "modPi"
' module containing plug-in loader
Option Explicit

Public Type PiType
 PiObject As String
 isConfigurable As Boolean
End Type

Public oPlugIn As Object

'function to read pi properties from file
Public Function getPlugInObject(PiFile As String) As PiType

Dim Fsys As New FileSystemObject
Dim Fin As TextStream, readLine As String
Dim p As PiType
   
Set Fin = Fsys.OpenTextFile(PiFile)

While Not Fin.AtEndOfStream
readLine = Fin.readLine
' found an object
  If InStr(1, readLine, "OBJECT") Then
   p.PiObject = Trim(Split(readLine, ":")(1))
  End If
' found a config setting
 If InStr(1, readLine, "CONFIG") Then
    p.isConfigurable = CBool(Split(readLine, ":")(1))
  End If
 
Wend
Fin.Close
Set Fin = Nothing
Set Fsys = Nothing

getPlugInObject = p

End Function


Public Sub importXPIList(File As String)
Dim InStream As TextStream

Dim a As Boolean, str As String
Dim i As Integer, ext As String, c As Integer

Let a = True
If Trim(File) = "" Or Dir(File) = "" Then GoTo e
Set InStream = Fsys.OpenTextFile(Trim(File), ForReading, False)


If Not StrComp(Replace(InStream.Read(5), "<?", " "), "Soda") Then
 Dim e As ErrStruct
 e.errNum = 10
 e.errShortDesc = "This does not appear to be a FireAMP! Orange Soda list"
 e.errLongDesc = "The file recently opened did not have the FireAMP! Orange Soda header in it. The File is either corrupt or invalid"
 logError e
Exit Sub
End If

InStream.SkipLine ' Skip header
InStream.SkipLine ' Skip comment
InStream.SkipLine ' Skip main tag
i = frmDummy.mnuObjectName.Count - 1

While InStream.AtEndOfStream = False

str = InStream.readLine
If str = "</list>" Then GoTo JMP

If a = True Then
On Error Resume Next
 Load frmDummy.mnuObjectName(i)

i = i + 1
frmDummy.mnuObjectName(i - 1).Caption = Split(parseString(str, 9, 9), ",")(1) ' load name
frmDummy.mnuObjectName(i - 1).tag = Split(parseString(str, 9, 9), ",")(0) ' load object
Else
frmDummy.mnuObjectName(i - 1).tag = frmDummy.mnuObjectName(i - 1).tag & "," & Val(parseString(str, 8, 8))
End If
a = Not a
Wend

JMP:
For i = 0 To frmDummy.mnuObjectName.Count - 1
 frmDummy.mnuObjectName(i).Checked = False
 Next i
Set InStream = Nothing ' destroy object
e:
End Sub

