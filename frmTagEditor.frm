VERSION 5.00
Begin VB.Form frmTagEditor 
   AutoRedraw      =   -1  'True
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   " ID3 Editor"
   ClientHeight    =   6855
   ClientLeft      =   150
   ClientTop       =   540
   ClientWidth     =   8970
   Icon            =   "frmTagEditor.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   6855
   ScaleWidth      =   8970
   StartUpPosition =   3  'Windows Default
   Begin VB.Frame fraNotSupported 
      Height          =   5895
      Left            =   3600
      TabIndex        =   40
      Top             =   240
      Width           =   5175
      Begin VB.Label Label13 
         AutoSize        =   -1  'True
         Caption         =   "Tag Editing not Supported"
         Height          =   195
         Left            =   1800
         TabIndex        =   41
         Top             =   2880
         Width           =   1860
      End
   End
   Begin VB.Frame fraGroupOptions 
      Caption         =   "Group Options"
      Height          =   2295
      Left            =   3600
      TabIndex        =   1
      Top             =   720
      Width           =   5175
      Begin VB.CheckBox chkOverWrite 
         Caption         =   "Overwrite using File Mask"
         Height          =   255
         Left            =   360
         TabIndex        =   39
         Top             =   1200
         Width           =   2655
      End
      Begin VB.TextBox txtFileMask 
         Height          =   285
         Left            =   1800
         TabIndex        =   37
         Text            =   "%TRACK% %TITLE%"
         Top             =   840
         Width           =   2895
      End
      Begin VB.CheckBox chkAutoTrack 
         Caption         =   "Auto Track"
         Height          =   255
         Left            =   360
         TabIndex        =   32
         Top             =   1560
         Width           =   1095
      End
      Begin VB.CheckBox chkStripFiles 
         Caption         =   "Strip Tags from file(s)"
         Height          =   255
         Left            =   360
         TabIndex        =   30
         Top             =   1920
         Width           =   1935
      End
      Begin VB.PictureBox picHolder2 
         BorderStyle     =   0  'None
         Height          =   2775
         Left            =   240
         ScaleHeight     =   2775
         ScaleWidth      =   4815
         TabIndex        =   26
         Top             =   2520
         Width           =   4815
         Begin VB.TextBox txtRenameMask 
            Height          =   285
            Left            =   120
            TabIndex        =   31
            Text            =   "%TITLE%"
            Top             =   1080
            Width           =   4695
         End
         Begin VB.OptionButton optNoRename 
            Caption         =   "No Renaming"
            Height          =   195
            Left            =   120
            TabIndex        =   28
            Top             =   360
            Value           =   -1  'True
            Width           =   2775
         End
         Begin VB.OptionButton optRename 
            Caption         =   "Rename Files using File Mask"
            Height          =   255
            Left            =   120
            TabIndex        =   27
            Top             =   720
            Width           =   2655
         End
         Begin VB.Label Label11 
            AutoSize        =   -1  'True
            Caption         =   "To restore files to their original names, run the ""UnName.bat"" file which will be created at the file's location."
            Height          =   390
            Left            =   120
            TabIndex        =   29
            Top             =   2160
            Width           =   4665
            WordWrap        =   -1  'True
         End
      End
      Begin VB.PictureBox picHolder1 
         BorderStyle     =   0  'None
         Height          =   975
         Left            =   2760
         ScaleHeight     =   975
         ScaleWidth      =   2295
         TabIndex        =   24
         Top             =   1200
         Width           =   2295
         Begin VB.CommandButton cmdParse 
            Caption         =   "Parse"
            Height          =   375
            Left            =   720
            TabIndex        =   38
            Top             =   0
            Width           =   1215
         End
         Begin VB.CommandButton cmdMoreOptions 
            Caption         =   "More Options"
            Height          =   375
            Left            =   240
            TabIndex        =   25
            Top             =   600
            Width           =   1815
         End
      End
      Begin VB.ComboBox cboDirName 
         Height          =   315
         Left            =   1800
         Style           =   2  'Dropdown List
         TabIndex        =   3
         Top             =   360
         Width           =   2895
      End
      Begin VB.Label Label6 
         AutoSize        =   -1  'True
         Caption         =   "File name mask:"
         Height          =   195
         Left            =   360
         TabIndex        =   9
         Top             =   840
         Width           =   1140
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "Directory name is:"
         Height          =   195
         Left            =   360
         TabIndex        =   2
         Top             =   360
         Width           =   1260
      End
   End
   Begin VB.CommandButton cmdSaveAll 
      Caption         =   "&Save All"
      Height          =   375
      Left            =   5040
      TabIndex        =   23
      Top             =   6360
      Width           =   1335
   End
   Begin VB.CheckBox chkEnableGroupOptions 
      Caption         =   "Enable Group Options"
      Height          =   255
      Left            =   3720
      TabIndex        =   22
      Top             =   240
      Value           =   1  'Checked
      Width           =   1815
   End
   Begin VB.CommandButton cmdExit 
      Caption         =   "&Close"
      Height          =   375
      Left            =   7320
      TabIndex        =   21
      Top             =   6360
      Width           =   1335
   End
   Begin VB.CommandButton cmdSave 
      Caption         =   "&Save"
      Height          =   375
      Left            =   3600
      TabIndex        =   20
      Top             =   6360
      Width           =   1335
   End
   Begin VB.Frame fraID3 
      Caption         =   "ID3 v1.x"
      Height          =   2895
      Left            =   3600
      TabIndex        =   4
      Top             =   3240
      Width           =   5175
      Begin VB.TextBox txtTrack 
         Height          =   285
         Left            =   3960
         MaxLength       =   3
         TabIndex        =   19
         Top             =   1560
         Width           =   855
      End
      Begin VB.ComboBox cboGenre 
         Height          =   315
         Left            =   1200
         Style           =   2  'Dropdown List
         TabIndex        =   17
         Top             =   2280
         Width           =   3615
      End
      Begin VB.TextBox txtComment 
         Height          =   285
         Left            =   1200
         MaxLength       =   28
         TabIndex        =   16
         Top             =   1920
         Width           =   3615
      End
      Begin VB.TextBox txtYear 
         Height          =   285
         Left            =   1200
         MaxLength       =   4
         TabIndex        =   15
         Top             =   1560
         Width           =   1095
      End
      Begin VB.TextBox txtAlbum 
         Height          =   285
         Left            =   1200
         MaxLength       =   30
         TabIndex        =   14
         Top             =   1200
         Width           =   3615
      End
      Begin VB.TextBox txtArtist 
         Height          =   285
         Left            =   1200
         MaxLength       =   30
         TabIndex        =   13
         Top             =   840
         Width           =   3615
      End
      Begin VB.TextBox txtTitle 
         Height          =   285
         Left            =   1200
         MaxLength       =   30
         TabIndex        =   12
         Top             =   480
         Width           =   3615
      End
      Begin VB.Label Label10 
         AutoSize        =   -1  'True
         Caption         =   "Track#"
         Height          =   195
         Left            =   3360
         TabIndex        =   18
         Top             =   1560
         Width           =   525
      End
      Begin VB.Label Label8 
         AutoSize        =   -1  'True
         Caption         =   "Comments:"
         Height          =   195
         Left            =   360
         TabIndex        =   11
         Top             =   1920
         Width           =   780
      End
      Begin VB.Label Label7 
         AutoSize        =   -1  'True
         Caption         =   "Year:"
         Height          =   195
         Left            =   360
         TabIndex        =   10
         Top             =   1560
         Width           =   375
      End
      Begin VB.Label Label5 
         AutoSize        =   -1  'True
         Caption         =   "Genre:"
         Height          =   195
         Left            =   360
         TabIndex        =   8
         Top             =   2280
         Width           =   480
      End
      Begin VB.Label Label4 
         AutoSize        =   -1  'True
         Caption         =   "Album:"
         Height          =   195
         Left            =   360
         TabIndex        =   7
         Top             =   1200
         Width           =   480
      End
      Begin VB.Label Label3 
         AutoSize        =   -1  'True
         Caption         =   "Artist:"
         Height          =   195
         Left            =   360
         TabIndex        =   6
         Top             =   840
         Width           =   390
      End
      Begin VB.Label Label2 
         AutoSize        =   -1  'True
         Caption         =   "Title:"
         Height          =   195
         Left            =   360
         TabIndex        =   5
         Top             =   480
         Width           =   345
      End
   End
   Begin VB.Frame fraList 
      Caption         =   "Files to Edit"
      Height          =   6495
      Left            =   120
      TabIndex        =   0
      Top             =   240
      Width           =   3375
      Begin VB.ListBox lstPaths 
         Height          =   1230
         Left            =   480
         TabIndex        =   34
         Top             =   1680
         Visible         =   0   'False
         Width           =   1215
      End
      Begin VB.ListBox lstFiles 
         Height          =   5520
         Left            =   120
         TabIndex        =   33
         Top             =   240
         Width           =   3135
      End
      Begin VB.Label Label12 
         AutoSize        =   -1  'True
         Caption         =   "Alt + Down: Move Track Down"
         Height          =   195
         Left            =   120
         TabIndex        =   36
         Top             =   6120
         Width           =   2205
      End
      Begin VB.Label Label9 
         AutoSize        =   -1  'True
         Caption         =   "Alt + Up: Move Track Up"
         Height          =   195
         Left            =   120
         TabIndex        =   35
         Top             =   5880
         Width           =   1785
      End
   End
End
Attribute VB_Name = "frmTagEditor"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private moreOptionsFlag As Boolean
Private boldIndex As Integer

Private Sub chkEnableGroupOptions_Click()
fraGroupOptions.Enabled = chkEnableGroupOptions.Value
End Sub

Private Sub cmdExit_Click()

' refresh playlist
Dim i As Integer
With frmFirePL
 
 .lstPaths.Clear
 .lstPl.ListItems.Clear
  
  For i = 0 To lstPaths.ListCount - 1
   .lstPaths.AddItem lstPaths.List(i)
   .lstPl.ListItems.Add , , getFileTitleFromPath(lstPaths.List(i))
  Next
.lstPl.ListItems(selectedIndex).Selected = True
.lstPl.ListItems(boldIndex).Bold = True
End With

Unload Me
End Sub

Private Sub cmdMoreOptions_Click()
moreOptionsFlag = Not moreOptionsFlag

If moreOptionsFlag Then
   
 cmdMoreOptions.Caption = "Less Options"
 fraGroupOptions.Height = fraGroupOptions.Height + 3125
Else
  cmdMoreOptions.Caption = "More Options"
  fraGroupOptions.Height = fraGroupOptions.Height - 3125
End If

End Sub

Private Sub cmdParse_Click()
parseFileName getFileTitleFromPath(lstPaths.List(lstFiles.ListIndex)), txtFileMask.Text
End Sub

Private Sub cmdSave_Click()

Dim tag As tagID3_1x, file As String, i As Integer
i = lstFiles.ListIndex

file = lstPaths.List(i)

If chkStripFiles.Value Then

removeTags file
Exit Sub
End If

If chkOverWrite Then cmdParse_Click

tag.Title = txtTitle.Text

If cboGenre.ListIndex > 0 Then tag.Genre = CByte(cboGenre.ListIndex - 1)

If Trim(txtArtist.Text) = "" And cboDirName.Text = "artist" And fraGroupOptions.Enabled Then
    tag.Artist = getFolderTitleFromPath(Fsys.GetParentFolderName(file))
Else
    tag.Artist = txtArtist.Text
End If

If Trim(txtAlbum.Text) = "" And cboDirName.Text = "Album" And fraGroupOptions.Enabled Then
    tag.Album = getFolderTitleFromPath(Fsys.GetParentFolderName(file))
Else
    tag.Album = txtAlbum.Text
End If


tag.Comment = txtComment.Text
tag.Year = txtYear.Text
If chkAutoTrack.Value Then
    tag.Track = CByte(lstFiles.ListIndex + 1)
Else
    tag.Track = CByte(Val(txtTrack.Text))
End If



' write ID3v1.x Tag to file
writeID3_1x tag, file

Dim newfile As String
' rename
If isTagCompletelyFilled(tag) And optRename.Value Then
Dim Mask As String
Mask = txtRenameMask.Text

Mask = Replace(Mask, "%TRACK%", Format(tag.Track, "00"))
Mask = Replace(Mask, "%ARTIST%", tag.Artist)
Mask = Replace(Mask, "%ALBUM%", tag.Album)
Mask = Replace(Mask, "%TITLE%", tag.Title)
newfile = Mask

newfile = Fsys.GetParentFolderName(file) & "\" & newfile & ".mp3"
If Not optNoRename.Value Then
DoEvents
       'FSys.MoveFile file, newfile
       Name file As newfile
       lstFiles.RemoveItem i
       lstPaths.RemoveItem i
      
       lstFiles.AddItem getFileTitleFromPath(newfile), i
       lstPaths.AddItem newfile, i
       
Dim Fout As textStream
Set Fout = Fsys.OpenTextFile(Fsys.GetParentFolderName(file) & "\UnName.bat", ForAppending, True)
Fout.WriteLine "rename """ & getFileTitleFromPath(newfile) & ".mp3" & """ """ & getFileTitleFromPath(file) & ".mp3""" & vbNewLine
Fout.Close
Set Fout = Nothing

End If
End If

End Sub

Private Sub cmdSaveAll_Click()

Dim i As Integer
For i = 0 To lstFiles.ListCount - 1
lstFiles.ListIndex = i
cmdSave_Click
Next
    
End Sub

Private Sub Form_Load()

Dim i As Byte
cboDirName.AddItem "Album"
cboDirName.AddItem "Artist"

cboGenre.AddItem "Not Set"
For i = 0 To 147
cboGenre.AddItem getGenre(i)
Next

End Sub

Public Sub lstFiles_Click()

If getFileExtensionFromPath(lstPaths.List(lstFiles.ListIndex)) <> "mp3" Then
fraNotSupported.Visible = True

Else
fraNotSupported.Visible = False
Dim i As Integer, tag As tagID3_1x

tag = readID3_1x(lstPaths.List(lstFiles.ListIndex))

txtTitle.Text = tag.Title
txtAlbum.Text = tag.Album
txtArtist.Text = tag.Artist
txtYear.Text = tag.Year
txtTrack.Text = tag.Track
txtComment.Text = tag.Comment
Debug.Print tag.Genre

If tag.Genre < 147 Then
    cboGenre.ListIndex = tag.Genre + 1
Else
    cboGenre.ListIndex = 0
End If

If chkOverWrite Then cmdParse_Click
End If

End Sub

Private Sub lstFiles_KeyDown(KeyCode As Integer, Shift As Integer)

Dim i As Integer, Title As String, path As String
i = lstFiles.ListIndex
If KeyCode = vbKeyUp And Shift = 4 And i > 0 Then
 
 Title = lstFiles.List(i)
 path = lstPaths.List(i)
 
 lstFiles.RemoveItem i
 lstPaths.RemoveItem i
 
 lstFiles.AddItem Title, i - 1
 lstPaths.AddItem path, i - 1
lstFiles.ListIndex = i - 1
End If

If KeyCode = vbKeyDown And Shift = 4 And i < lstFiles.ListCount - 1 Then
 
 Title = lstFiles.List(i)
 path = lstPaths.List(i)
 
 lstFiles.RemoveItem i
 lstPaths.RemoveItem i
 
 lstFiles.AddItem Title, i + 1
 lstPaths.AddItem path, i + 1
lstFiles.ListIndex = i + 1
End If


End Sub

Public Sub importCurrentPlaylist()
 Dim i As Integer
 Dim Item As String
 
 For i = 0 To frmFirePL.lstPaths.ListCount - 1
 If frmFirePL.lstPl.ListItems(i + 1).Bold = True Then boldIndex = i + 1
 Item = frmFirePL.lstPaths.List(i)
  
    lstFiles.AddItem getFileTitleFromPath(Item)
  lstPaths.AddItem Item
  Next
 
lstPaths.Refresh
lstFiles.Refresh
End Sub

' parses ID3 Info from filename
Private Sub parseFileName(FLName As String, Mask As String)

Dim maskParts() As String, i As Integer
Dim sepChar As String
Dim parts() As String

maskParts = Split(Mask, " ")
FLName = Replace(FLName, "[", "")
FLName = Replace(FLName, "(", "")
FLName = Replace(FLName, "{", "")


For i = 0 To UBound(maskParts)
maskParts(i) = Trim(maskParts(i))

sepChar = " "
sepChar = Right(maskParts(i), 1)


If sepChar = "%" Then sepChar = " "

Debug.Print """" & sepChar & """"
Debug.Print FLName

 If InStr(1, maskParts(i), "%TRACK%") Then
        txtTrack.Text = Trim(Split(FLName, sepChar, UBound(maskParts) + 1 - i)(0))
         FLName = Replace(FLName, txtTrack.Text, "")
        
 ElseIf InStr(1, maskParts(i), "%ALBUM%") Then
        txtAlbum.Text = Trim(Split(FLName, sepChar, UBound(maskParts) + 1 - i)(0))
         FLName = Replace(FLName, txtAlbum.Text, "")
        
 ElseIf InStr(1, maskParts(i), "%TITLE%") Then
        txtTitle.Text = Trim(Split(FLName, sepChar, UBound(maskParts) + 1 - i)(0))
         FLName = Replace(FLName, txtTitle.Text, "")
         
 ElseIf InStr(1, maskParts(i), "%ARTIST%") Then
        txtArtist.Text = Trim(Split(FLName, sepChar, UBound(maskParts) + 1 - i)(0))
         FLName = Replace(FLName, txtArtist.Text, "")
         
        
 End If
 FLName = Trim(IIf(sepChar <> " ", Replace(FLName, sepChar, ""), FLName))
 
Next

txtArtist.Text = Trim(Replace(txtArtist.Text, "_", " "))
txtTitle.Text = Trim(Replace(txtTitle.Text, "_", " "))
txtAlbum.Text = Trim(Replace(txtAlbum.Text, "_", " "))

FLName = txtTitle.Text
Mid(FLName, 1, 1) = UCase(Mid(FLName, 1, 1))
txtTitle.Text = FLName


End Sub

