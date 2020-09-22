Attribute VB_Name = "modID3_1x"
Option Explicit

' module containing ID3 related stuff

' ID3 1.x tag structure
Public Type tagID3_1x
   
  tag As String * 3           ' byte 003
  Title As String             ' byte 033
  Artist As String            ' byte 063
  Album As String             ' byte 093
  Year As String              ' byte 097
  Comment As String           ' byte 125
  Filler As Byte              ' byte 126
  Track As Byte               ' byte 127
  Genre As Byte               ' byte 128
End Type

' mp3 header data structure
Public Type mp3HeaderData
ID As String
Layer As String
ProtectionBitSet As Boolean
BitRate As Integer
Frequency As Long
Padded As Boolean
PrivateBitSet As Boolean
Mode As String
ModeExt As Long
CopyRighted As Boolean
Original As Boolean
Emphasis As Boolean
End Type


Public Function readID3_1x(mp3File As String) As tagID3_1x

Dim fNum As Integer
Dim theTag As tagID3_1x

Dim Title As String * 30
Dim Artist As String * 30
Dim Album As String * 30
Dim Comment As String * 28
Dim Track As Byte
Dim Year As String * 4

fNum = FreeFile

' fill the tag with dummy info
theTag.Album = ""
theTag.Artist = ""
theTag.Comment = ""
theTag.Genre = 255
theTag.Track = 0
theTag.Title = ""
theTag.Year = ""

' open mp3 file

On Error GoTo errHandle
Reset
If Not Fsys.FileExists(mp3File) Then Exit Function
' read tag
Open mp3File For Binary As #fNum
   If LOF(fNum) > 128 Then
      Get #fNum, LOF(fNum) - 127, theTag.tag
              If Not theTag.tag = "TAG" Then
              readID3_1x = theTag
              Exit Function
         Else
            
            Get #fNum, , Title
            theTag.Title = toStdString(Title)
            
            
            Get #fNum, , Artist
            theTag.Artist = toStdString(Artist)
            
            Get #fNum, , Album
            theTag.Album = toStdString(Album)
            
            Get #fNum, , Year
            theTag.Year = toStdString(theTag.Year)
            
            Get #fNum, , Comment
            theTag.Comment = toStdString(Comment)
            
            Get #fNum, , theTag.Filler
            Get #fNum, , theTag.Track
                                    
            
            Get #fNum, , theTag.Genre
                        
         End If
      End If
   
   Close #fNum
readID3_1x = theTag
  
Exit Function
errHandle:
 
Err.Clear ' must always clear curent error
End Function

Public Sub writeID3_1x(theTag As tagID3_1x, mp3File As String)

Dim fNum As Integer, offSet As Integer

Dim Title As String * 30
Dim Artist As String * 30
Dim Album As String * 30

Dim Year As String * 4
Dim Comment As String * 28

Reset
fNum = FreeFile
removeTags mp3File


Title = ""
Artist = ""
Album = ""
Comment = ""
Year = ""

On Error GoTo errHandle
If Not Fsys.FileExists(mp3File) Then Exit Sub
    LSet Title = theTag.Title
    LSet Artist = theTag.Artist
    LSet Album = theTag.Album
    LSet Year = theTag.Year
    LSet Comment = theTag.Comment

   Open mp3File For Binary Access Read Write Lock Write As #fNum
                    
                Seek #fNum, LOF(fNum)
                
                theTag.tag = "TAG"
                Put #fNum, , theTag.tag
                Put #fNum, , Title
                Put #fNum, , Artist
                Put #fNum, , Album
                Put #fNum, , Year
                Put #fNum, , Comment
                Put #fNum, , theTag.Filler
                Put #fNum, , theTag.Track
                Put #fNum, , theTag.Genre

Close #fNum
Exit Sub
errHandle:

Close #fNum
Err.Clear

End Sub

Public Sub removeTags(mp3File As String)


Dim inFile As Integer, outFile As Integer, offSet As Integer, tag As String * 3
Dim tagsFound As Integer
inFile = FreeFile
tagsFound = 1

offSet = 127
On Error GoTo EN
Reset
If Not Fsys.FileExists(mp3File) Then Exit Sub

Open mp3File For Binary Access Read Write Lock Read As #inFile

Do
Seek #inFile, LOF(inFile) - offSet
  Get #inFile, , tag
 offSet = offSet + 127
 tagsFound = tagsFound + 1
Loop Until tag <> "TAG"


Close #inFile

Dim buff() As Byte

ReDim buff(0 To FileLen(mp3File) - 127 * tagsFound) As Byte

inFile = FreeFile

Open mp3File For Binary Access Read Write Lock Read As #inFile

outFile = FreeFile

Open App.path & "\temp.cpy" For Binary Access Read Write Lock Write As #outFile

Get #inFile, , buff
Put #outFile, , buff

Close #inFile
Close #outFile

FileCopy App.path & "\temp.cpy", mp3File
Kill App.path & "\temp.cpy"

EN:

End Sub

Public Function isTagCompletelyFilled(tag As tagID3_1x) As Boolean

'isTagCompletelyFilled = CBool(tag.Album <> "" & tag.Artist <> "" & tag.Title <> "" & tag.Track <> 0)
isTagCompletelyFilled = True
End Function

' from PSC.. thank you unknown author :)
Public Function ReadMP3Header(FileName As String) As mp3HeaderData
On Error GoTo fault

Dim bTMP1 As Byte, bTMP2 As Byte
Dim filenum As Integer
filenum = FreeFile
Dim i As Long, ValidHeader As Boolean
Dim StartByte As Long
Open FileName For Binary Access Read As #filenum
    
' FINDING FIRST FRAME HEADER

    For i = 1 To LOF(filenum) - 1
        Get #filenum, i, bTMP1
        If bTMP1 = &HFF Then
            Get #filenum, i + 1, bTMP2
            If bTMP2 And &HE0 = &HE0 Then
                ValidHeader = True
                StartByte = i + 1
                Exit For
            End If
        End If
    Next
    
    If Not ValidHeader Then Exit Function
    
    Dim b1 As Byte, b2 As Byte, b3 As Byte, b4 As Byte, b5 As Byte
    
' GETTING BYTES THAT CONTAIN HEADER INFORMATIONS

    Get #filenum, StartByte, b1
    Get #filenum, StartByte + 1, b2
    Get #filenum, StartByte + 2, b3
    Get #filenum, StartByte + 3, b4
    Get #filenum, StartByte + 4, b5
    
Close #filenum

' READING MPEG MODE
Select Case CInt(b1 And &H18) / 8
    Case 0
        ReadMP3Header.ID = "Mpeg 2.5"
    Case 1
        ReadMP3Header.ID = "Not defined"
    Case 2
        ReadMP3Header.ID = "Mpeg 2"
    Case 3
        ReadMP3Header.ID = "Mpeg 1"
End Select
        
    
'READING LAYER INFO
Select Case (b1 And &H6)
    Case &H0
        ReadMP3Header.Layer = "Not defined"
    Case &H2
        ReadMP3Header.Layer = "Layer III"
    Case &H4
        ReadMP3Header.Layer = "Layer II"
    Case &H6
        ReadMP3Header.Layer = "Layer I"
End Select

' READING PROTECTION BIT, AND PROTECTION CHECKSUM IF THE BIT IS NOT SET
If (b1 And &H1) = &H1 Then
    ReadMP3Header.ProtectionBitSet = True
Else
    ReadMP3Header.ProtectionBitSet = False
    'ProtectionChecksum = Hex(b4) & " " & Hex(b5)
End If

' READING BITRATE INFO
Dim arg1 As Integer, arg2 As Integer, arg3 As Integer
arg1 = CInt(b1 And &H8) / 8
arg2 = CInt(b1 And &H6) / 2
arg3 = CInt(b2 And &HF0) / 16
ReadMP3Header.BitRate = fnGetBitrate(arg1, arg2, arg3)

' READING FREQUENCY (SAMPLERATE)
Dim k As Long
Select Case CInt(b1 And &H18) / 8
    Case 0
        k = 1
    Case 1
        k = 0
    Case 2
        k = 2
    Case 3
        k = 4
End Select
Select Case CInt(b2 And &HC) / 4
    Case 0
        ReadMP3Header.Frequency = k * 11025
    Case 1
        ReadMP3Header.Frequency = k * 12000
    Case 2
        ReadMP3Header.Frequency = k * 8000
End Select

' READING PADDING BIT
If (b2 And &H2) = &H2 Then ReadMP3Header.Padded = True

' READING PRIVATE BIT
If (b2 And &H1) = &H1 Then ReadMP3Header.PrivateBitSet = True

' READING CHANNEL MODE
Select Case CInt(b3 And &HC0) / CInt(&H40)
    Case 0
        ReadMP3Header.Mode = "Stereo"
    Case 1
        ReadMP3Header.Mode = "Joint Stereo"
    Case 2
        ReadMP3Header.Mode = "Dual channel"
    Case 3
        ReadMP3Header.Mode = "Mono"
End Select

' READING MODE EXTENSION (I DON'T KNOW WHAT IT IS)
ReadMP3Header.ModeExt = CInt(b3 And &H30) / CInt(&H10)

' READING COPYRIGHT BIT
If (b3 And &H8) = &H8 Then ReadMP3Header.CopyRighted = True

' READING ORIGINAL HOME BIT
If (b3 And &H4) = &H4 Then ReadMP3Header.Original = True

' READING EMPHASIS INFO
Select Case b3 And &H3
    Case 0
        ReadMP3Header.Emphasis = "None"
    Case 2
        ReadMP3Header.Emphasis = "Not defined"
    Case 1
        ReadMP3Header.Emphasis = "50/15 ms"
    Case 3
        ReadMP3Header.Emphasis = "CCITT j. 17"
End Select
    
GoTo Finish

fault:
Close filenum

Finish:
End Function


' FUNCTION FOR GETTING BITRATE INFO
Private Function fnGetBitrate(arg1 As Integer, arg2 As Integer, arg3 As Integer) As Integer
Dim i As Integer
Dim a(1, 3, 15) As Integer
For i = 1 To 14
    a(1, 3, i) = i * 32
    If i < 5 Then
        a(1, 2, i) = 8 * (i + 4)
        a(1, 1, i) = 8 * (i + 3)
    Else
        a(1, 2, i) = a(1, 2, i - 4) * 2
        a(1, 1, i) = a(1, 1, i - 4) * 2
    End If
    If i < 9 Then
        a(0, 1, i) = i * 8
    Else
        a(0, 1, i) = (i - 4) * 16
    End If
    a(0, 2, i) = a(0, 1, i)
Next
a(1, 2, 1) = 32
a(0, 3, 1) = 32
a(0, 3, 2) = 48
a(0, 3, 3) = 56
a(0, 3, 4) = 64
a(0, 3, 5) = 80
a(0, 3, 6) = 96
a(0, 3, 7) = 112
a(0, 3, 8) = 128
a(0, 3, 9) = 144
a(0, 3, 10) = 160
a(0, 3, 11) = 176
a(0, 3, 12) = 192
a(0, 3, 13) = 224
a(0, 3, 14) = 256

fnGetBitrate = a(arg1, arg2, arg3)
If arg3 = 15 Then fnGetBitrate = 1
If arg3 = 0 Then fnGetBitrate = 0
End Function


' function to get a genre corresponding to a genre number
Public Function getGenre(Genre As Byte)
Dim sName As String
   Select Case Genre
   'A
   Case 34: sName = "Acid"
   Case 74: sName = "Acid Jazz"
   Case 73: sName = "Acid Punk"
   Case 99: sName = "Acoustic"
   Case 40: sName = "Alt.Rock"
   Case 20: sName = "Alternative"
   Case 26: sName = "Ambient"
   Case 145: sName = "Anime"
   Case 90: sName = "Avant Garde"
   
   'B
   Case 116: sName = "Ballad"
   Case 41: sName = "Bass"
   Case 135: sName = "Beat"
   Case 85: sName = "Bebob"
   Case 96: sName = "Big Band"
   Case 138: sName = "Black Metal"
   Case 89: sName = "Blue Grass"
   Case 0: sName = "Blues"
   Case 107: sName = "Booty Bass"
   Case 132: sName = "Brit Pop"
   
   'C
   Case 65: sName = "Cabaret"
   Case 88: sName = "Celtic"
   Case 104: sName = "Chamber Music"
   Case 102: sName = "Chanson"
   Case 97: sName = "Chorus"
   Case 136: sName = "Christian Gangsta Rap"
   Case 61: sName = "Christian Rap"
   Case 141: sName = "Christian Rock"
   Case 1: sName = "Classic Rock"
   Case 32: sName = "Classical"
   Case 112: sName = "Club"
   Case 128: sName = "Club - House"
   Case 57: sName = "Comedy"
   Case 140: sName = "Contemporary Christian"
   Case 2: sName = "Country"
   Case 139: sName = "Crossover"
   Case 58: sName = "Cult"
   
   'D
   Case 3: sName = "Dance"
   Case 125: sName = "Dance Hall"
   Case 50: sName = "Darkwave"
   Case 22: sName = "Death Metal"
   Case 4: sName = "Disco"
   Case 55: sName = "Dream"
   Case 127: sName = "Drum & Bass"
   Case 122: sName = "Drum Solo"
   Case 120: sName = "Duet"
   
   'E
   Case 98: sName = "Easy Listening"
   Case 52: sName = "Electronic"
   Case 48: sName = "Ethnic"
   Case 54: sName = "Eurodance"
   Case 124: sName = "Euro - House"
   Case 25: sName = "Euro - Techno"
   
   'F
   Case 84: sName = "Fast Fusion"
   Case 80: sName = "Folk"
   Case 81: sName = "Folk / Rock"
   Case 115: sName = "Folklore"
   Case 119: sName = "Freestyle"
   Case 5: sName = "Funk"
   Case 30: sName = "Fusion"
   
   'G
   Case 36: sName = "Game"
   Case 59: sName = "Gangsta Rap"
   Case 126: sName = "Goa"
   Case 38: sName = "Gospel"
   Case 49: sName = "Gothic"
   Case 91: sName = "Gothic Rock"
   Case 6: sName = "Grunge"
   
   'H
   Case 79: sName = "Hard Rock"
   Case 129: sName = "Hardcore"
   Case 137: sName = "Heavy Metal"
   Case 7: sName = "Hip Hop"
   Case 35: sName = "House"
   Case 100: sName = "Humour"
   
   'I
   Case 131: sName = "Indie"
   Case 19: sName = "Industrial"
   Case 33: sName = "Instrumental"
   Case 46: sName = "Instrumental Pop"
   Case 47: sName = "Instrumental Rock"
   
   'J
   Case 8: sName = "Jazz"
   Case 29: sName = "Jazz - Funk"
   Case 146: sName = "JPop"
   Case 63: sName = "Jungle"
   
   'L
   Case 86: sName = "Latin"
   Case 71: sName = "Lo - fi"
   
   'M
   Case 45: sName = "Meditative"
   Case 142: sName = "Merengue"
   Case 9: sName = "Metal"
   Case 77: sName = "Musical"
   Case 82: sName = "National Folk"

   'N
   Case 64: sName = "Native American"
   Case 133: sName = "Negerpunk"
   Case 10: sName = "New Age"
   Case 66: sName = "New Wave"
   Case 39: sName = "Noise"
   
   'O
   Case 11: sName = "Oldies"
   Case 103: sName = "Opera"
   Case 12: sName = "Other"
   
   'P
   Case 75: sName = "Polka"
   Case 134: sName = "Polsk Punk"
   Case 13: sName = "Pop"
   Case 62: sName = "Pop / Funk"
   Case 53: sName = "Pop / Folk"
   Case 109: sName = "Pr0n Groove"
   Case 117: sName = "Power Ballad"
   Case 23: sName = "Pranks"
   Case 108: sName = "Primus"
   Case 92: sName = "Progressive Rock"
   Case 67: sName = "Psychedelic"
   Case 93: sName = "Psychedelic Rock"
   Case 43: sName = "Punk"
   Case 121: sName = "Punk Rock"
   
   'R
   Case 14: sName = "R&B"
   Case 15: sName = "Rap"
   Case 68: sName = "Rave"
   Case 16: sName = "Reggae"
   Case 76: sName = "Retro"
   Case 87: sName = "Revival"
   Case 118: sName = "Rhythmic Soul"
   Case 17: sName = "Rock"
   Case 78: sName = "Rock 'n'Roll"
   
   'S
   Case 143: sName = "Salsa"
   Case 114: sName = "Samba"
   Case 110: sName = "Satire"
   Case 69: sName = "Showtunes"
   Case 21: sName = "Ska"
   Case 111: sName = "Slow Jam"
   Case 95: sName = "Slow Rock"
   Case 105: sName = "Sonata"
   Case 42: sName = "Soul"
   Case 37: sName = "Sound Clip"
   Case 24: sName = "Soundtrack"
   Case 56: sName = "Southern Rock"
   Case 44: sName = "Space"
   Case 101: sName = "Speech"
   Case 83: sName = "Swing"
   Case 94: sName = "Symphonic Rock"
   Case 106: sName = "Symphony"
   Case 147: sName = "Synth Pop"

   'T
   Case 113: sName = "Tango"
   Case 18: sName = "Techno"
   Case 51: sName = "Techno - Industrial"
   Case 130: sName = "Terror"
   Case 144: sName = "Thrash Metal"
   Case 60: sName = "Top 40"
   Case 70: sName = "Trailer"
   Case 31: sName = "Trance"
   Case 72: sName = "Tribal"
   Case 27: sName = "Trip Hop"
   
   'V
   Case 28: sName = "Vocal"
   
   End Select
   getGenre = sName
End Function

